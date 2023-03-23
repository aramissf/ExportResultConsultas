# importando bibliotecas
import sys
from design import *
from PyQt5.QtWidgets import QMainWindow, QApplication, QMessageBox
import pyodbc
from openpyxl import Workbook
from tkinter import messagebox

class Exporta(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)
        self.btn_salva_conf.clicked.connect(self.salvaConfig)
    
        
        
        

    def salvaConfig(self):
        #Lendo os parametro de configuração preenchidos
        servidor = self.tx_servidor.text()
        banco = self.tx_banco.text()
        usuario = self.tx_user.text()
        senha = self.tx_senha.text() 
        nome = self.tx_nome_arquivo.text()
        localsalvar = self.tx_local_salvar.text()
        
        #Criando arquivo de configuracao
        with open("config.txt", 'w') as config:
            config.write('serivor : ' + servidor +'\n')
            config.write('banco : ' + banco+'\n')
            config.write('usuario : ' + usuario+'\n')
            config.write('senha : ' + senha+'\n')
            config.write('nome : ' + nome+'\n')
            config.write('caminho : ' + localsalvar+'\n')
        config.close()

        #Informando salvamento dos parametros
        messagebox.showinfo("Messagem", "Salvo com Sucesso!")
        
            

if __name__ == '__main__':
    qt = QApplication(sys.argv)
    exporta = Exporta()
    exporta.show()
    qt.exec_()








