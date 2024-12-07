import os
import sys
import sqlite3
import subprocess
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QProgressBar, QMessageBox, QInputDialog, QDialogButtonBox, QDialog
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
from PyQt5.QtWinExtras import QtWin
import time
from dois.p import executar_codigo

LICENSE_FILE = "a.txt"
SECRET_PASSWORD = "Password1u$"
MASTER_PASSWORD = "ChangePass80@"
AUTHORIZED_SERIAL = "BRG319FDZP"

def obter_serial_placa_mae():
    try:
        result = subprocess.run(['wmic', 'bios', 'get', 'serialnumber'], capture_output=True, text=True)
        serial_line = result.stdout.strip().split('\n')[-1].strip()
        return serial_line
    except Exception as e:
        print(f"Contate o fornecedor: {e}")
        return None

def verificar_sm():
    serial_atual = obter_serial_placa_mae()
    if serial_atual != AUTHORIZED_SERIAL:
        print(f"Acesso negado. O aplicativo só pode ser executado na máquina autorizada.")
        sys.exit(1)

verificar_sm()

class ChangePasswordDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Alterar Senha")
        
        self.master_password_label = QLabel("Senha Mestre:", self)
        self.new_secret_password_label = QLabel("Nova Senha de Execução:", self)

        self.master_password_input = QLineEdit(self)
        self.master_password_input.setEchoMode(QLineEdit.Password)
        
        self.new_secret_password_input = QLineEdit(self)
        self.new_secret_password_input.setEchoMode(QLineEdit.Password)
        
        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)

        layout = QVBoxLayout()
        layout.addWidget(self.master_password_label)
        layout.addWidget(self.master_password_input)
        layout.addWidget(self.new_secret_password_label)
        layout.addWidget(self.new_secret_password_input)
        layout.addWidget(self.buttons)

        self.setLayout(layout)

    def get_passwords(self):
        return self.master_password_input.text(), self.new_secret_password_input.text()

class Interface(QWidget):
    def __init__(self):
        super().__init__()
        icon_path = r'C:\Users\lucas\Downloads\Projeto IM DICAP\vitalicio\app\dist\256X256ICOdistoutro.ico'
        self.setWindowIcon(QIcon(icon_path))
        
        if sys.platform.startswith('win') and QtWin:
            QtWin.setCurrentProcessExplicitAppUserModelID('mycompany.myproduct.subproduct.version')
            self.setWindowIcon(QIcon(icon_path))
        
        self.conn = sqlite3.connect('licenses.db')
        self.create_table_if_not_exists()

        self.max_execucoes = None
        self.execucoes_restantes = None
        
        self.initUI()

        if not self.load_license() or (self.execucoes_restantes is not None and self.execucoes_restantes <= 0):
            self.request_license()

    def create_table_if_not_exists(self):
        cursor = self.conn.cursor()
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS licenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            max_execucoes INTEGER,
            execucoes_restantes INTEGER
        )
        ''')
        self.conn.commit()

    def initUI(self):
        self.label = QLabel('Insira o caminho da pasta que contém as planilhas DEFINICAO e CADASTRO:', self)
        self.label.setAlignment(Qt.AlignHCenter)
        self.label.setStyleSheet("font-size: 18px;")

        self.input = QLineEdit(self)
        self.input.setFixedHeight(self.input.sizeHint().height() + 10)
        self.input.setStyleSheet("font-size: 18px;")
        
        self.button = QPushButton('EXECUTAR', self)
        self.button.setStyleSheet("font-size: 18px;")
        self.button.setFixedHeight(self.input.sizeHint().height() + 10)
        self.button.clicked.connect(self.on_executar)
        
        self.progress = QProgressBar(self)
        self.progress.setFixedHeight(self.input.sizeHint().height() + 10)
        self.progress.setValue(0)
        self.progress.setTextVisible(True)
        self.progress.setAlignment(Qt.AlignCenter)
        self.progress.setStyleSheet("font-size: 18px;")
        
        self.warning_label = QLabel('Revenda não autorizada ou pirataria é contra a lei, para mais informações entre em contato: imdicap@gmail.com', self)
        self.warning_label.setAlignment(Qt.AlignHCenter)
        self.warning_label.setStyleSheet("font-size: 13px; font-weight: bold; color: darkred;")

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.input)
        layout.addWidget(self.button)
        layout.addWidget(self.progress)
        layout.addWidget(self.warning_label)
        layout.setSpacing(5)
        self.setWindowFlags(Qt.WindowCloseButtonHint | Qt.WindowMinimizeButtonHint)
        self.setLayout(layout)
        
        self.setWindowTitle('IM & DICAP - Inteligência Manual & Definição de Impostos para Cadastro Automatizado de Produtos')
        self.setFixedSize(750, 200)
        self.show()

    def request_license(self):
        while True:
            password, ok = QInputDialog.getText(self, 'Verificação de Senha', 'Digite a senha de execução:', QLineEdit.Password)
            if ok and password == SECRET_PASSWORD:
                executions, ok = QInputDialog.getText(self, 'Definição de Licença', 'Chave de licença:')
                if ok and (executions.isdigit() or executions == "VITAL"):
                    self.max_execucoes = None if executions == "VITAL" else int(executions)
                    self.execucoes_restantes = self.max_execucoes
                    self.save_license()
                    return
                else:
                    self.show_message('Entrada Inválida', 'Definição de licença inválida')
            else:
                change_password = QMessageBox.question(self, 'Opção', "Deseja alterar a senha de execução?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if change_password == QMessageBox.Yes:
                    if self.handle_password_change():
                        continue
                sys.exit()

    def handle_password_change(self):
        dialog = ChangePasswordDialog()
        while dialog.exec_():
            master_password, new_secret_password = dialog.get_passwords()
            if master_password == MASTER_PASSWORD and new_secret_password:
                global SECRET_PASSWORD
                SECRET_PASSWORD = new_secret_password
                QMessageBox.information(self, "Sucesso", "Senha de execução alterada com sucesso!")
                return True
            else:
                QMessageBox.warning(self, "Falha", "Senha mestre incorreta ou nova senha inválida.")
        return False 

    def save_license(self):
        cursor = self.conn.cursor()
        cursor.execute('INSERT INTO licenses (max_execucoes, execucoes_restantes) VALUES (?, ?)', (self.max_execucoes, self.execucoes_restantes))
        self.conn.commit()

    def load_license(self):
        cursor = self.conn.cursor()
        cursor.execute('SELECT max_execucoes, execucoes_restantes FROM licenses ORDER BY id DESC LIMIT 1')
        row = cursor.fetchone()
        
        if row:
            self.max_execucoes = row[0]
            self.execucoes_restantes = row[1]
            return True
        return False

    def show_message(self, title, message):
        msg_box = QMessageBox(QMessageBox.Information, title, message)
        msg_box.setWindowFlag(Qt.WindowStaysOnTopHint)
        msg_box.exec_()

    def on_executar(self):
        caminho = self.input.text()

        definicao_path = os.path.join(caminho, 'DEFINICAO.xlsx')
        cadastro_path = os.path.join(caminho, 'CADASTRO.xlsx')

        if not os.path.exists(definicao_path) or not os.path.exists(cadastro_path):
            QMessageBox.warning(self, 'Erro', (
                'Os arquivos não foram encontrados.\n\nCertifique-se de que:\n'
                '1- Foi mencionado o caminho da pasta que contém as duas planilhas ao invés de mencionar o caminho de uma única planilha.\n'
                '2- O caminho mencionado existe e tente acessá-lo manualmente.\n'
                '3- O caminho da pasta contém duas planilhas nomeadas como DEFINICAO e CADASTRO.\n'
                '4- O caminho da pasta não possui restrição de acesso.\n'
                '5- O procedimento pode ser testado usando outro caminho, basta mudar os arquivos de pasta.'
            ))
            return

        try:
            start_time = time.time()
            executar_codigo(caminho, self.update_progress)
            end_time = time.time()

            elapsed_time = end_time - start_time

            success_msg = QMessageBox(
                QMessageBox.Information,
                'Processo Concluído',
                f'Os dados da planilha CADASTRO foram atualizados com base nos dados da planilha DEFINICAO.\n\nTempo gasto: {elapsed_time:.2f} segundos'
            )
            success_msg.setWindowFlag(Qt.WindowStaysOnTopHint)
            success_msg.exec_()

            if self.execucoes_restantes is not None and self.execucoes_restantes > 0:
                self.execucoes_restantes -= 1
                self.save_license()

        except Exception as e:
            QMessageBox.critical(self, 'Erro durante a Execução', f'Ocorreu um erro durante a execução: {str(e)}')

        if self.execucoes_restantes is not None and self.execucoes_restantes <= 0:
            QMessageBox.critical(self, 'Licença Expirada', 'Entre em contato com o fornecedor')
            sys.exit()

    def update_progress(self, percentual):
        self.progress.setValue(percentual)
        self.progress.setFormat(f'{percentual}%')

def iniciar_interface():
    app = QApplication(sys.argv)
    ex = Interface()
    sys.exit(app.exec_())