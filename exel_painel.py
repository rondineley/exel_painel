from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QWidget, QFileDialog, QMessageBox, QScrollArea
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor
import sys
import os
import csv



class ArtStationStyleApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Galeria de Arquivos Excel")
        self.setGeometry(100, 100, 1200, 800)
        self.setStyleSheet("background-color: white; color: black;")

        container = QWidget(self)
        self.setCentralWidget(container)
        self.layout_principal = QVBoxLayout(container)

        self.label_titulo = QLabel("Galeria de Arquivos")
        self.label_titulo.setAlignment(Qt.AlignCenter)
        self.label_titulo.setStyleSheet("font-size: 24px; font-weight: bold; color: #00ADEF;")
        self.layout_principal.addWidget(self.label_titulo)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("background-color: #f4f4f4; border: none;")
        self.layout_principal.addWidget(self.scroll_area)

        self.file_container = QWidget()
        self.file_layout = QVBoxLayout(self.file_container)
        self.scroll_area.setWidget(self.file_container)

        self.botao_container = QWidget()
        self.botao_layout = QHBoxLayout(self.botao_container)  # Layout horizontal para alinhar os botões à direita
        self.botao_layout.setAlignment(Qt.AlignRight)  # Alinha à direita
        self.layout_principal.addWidget(self.botao_container)

        self.botao_selecionar = self.criar_botao("Selecionar Diretório", self.selecionar_diretorio)
        self.botao_layout.addWidget(self.botao_selecionar)

        self.botao_exportar = self.criar_botao("Exportar Lista", self.exportar_lista)
        self.botao_layout.addWidget(self.botao_exportar)

        self.botao_classificar_nome = self.criar_botao("Classificar por Nome", self.classificar_por_nome)
        self.botao_layout.addWidget(self.botao_classificar_nome)

        self.botao_classificar_data = self.criar_botao("Classificar por Data de Modificação", self.classificar_por_data)
        self.botao_layout.addWidget(self.botao_classificar_data)

        self.botao_sair = self.criar_botao("Sair", self.close, estilo="background-color: #D9534F;")
        self.botao_layout.addWidget(self.botao_sair)

        self.diretorio_atual = ""

    
    
    
    def criar_botao(self, texto, funcao, estilo=None):
        botao = QPushButton(texto)
        botao.setStyleSheet(f"""
            QPushButton {{
                background-color: #A8D08D;  /* Verde Claro */
                border: none;
                padding: 12px;
                font-size: 18px;
                color: black;
                border-radius: 15px;
                width: 300px; /* Largura fixada para o botão */
            }}
            QPushButton:hover {{
                background-color: #7EA67E;  /* Verde escuro ao passar o mouse */
            }}
        """ if not estilo else estilo)
        botao.clicked.connect(funcao)
        return botao

   
   
    def selecionar_diretorio(self):
        diretorio = QFileDialog.getExistingDirectory(self, "Selecione o Diretório")
        if diretorio:
            self.diretorio_atual = diretorio
            self.listar_arquivos_excel(diretorio)

    
    
    
    def listar_arquivos_excel(self, diretorio):
        for i in reversed(range(self.file_layout.count())):
            widget = self.file_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)

        self.arquivos = []
        for raiz, _, arquivos in os.walk(diretorio):
            for arquivo in arquivos:
                if arquivo.endswith(('.xls', '.xlsx')):
                    caminho = os.path.join(raiz, arquivo)
                    self.arquivos.append(caminho)
                    self.adicionar_item_lista(caminho)

   
   
    def adicionar_item_lista(self, caminho_arquivo):
        item_container = QWidget()
        item_layout = QVBoxLayout(item_container)  # Usamos QVBoxLayout para ocupar toda a linha
        item_container.setStyleSheet("""
            QWidget {
                background-color: #D9EED6;  /* Verde claro (Excel) */
                border: 2px solid #A8C686;
                border-radius: 5px;
                padding: 5px;
            }
        """)

        label_nome = QLabel(os.path.basename(caminho_arquivo))
        label_nome.setStyleSheet("font-size: 16px; font-weight: bold; color: black; padding: 10px;")
        item_layout.addWidget(label_nome)

        botao_abrir = self.criar_botao("Abrir", lambda: self.abrir_arquivo_excel(caminho_arquivo))
        item_layout.addWidget(botao_abrir)
        self.file_layout.addWidget(item_container)

   
   
   
   
    def abrir_arquivo_excel(self, caminho_arquivo):
        try:
            os.startfile(caminho_arquivo)  # Abre o arquivo com o programa padrão (Excel)
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Não foi possível abrir o arquivo:\n{str(e)}")

   
   
   
    def exportar_lista(self):
        if not self.diretorio_atual:
            QMessageBox.warning(self, "Aviso", "Selecione um diretório primeiro!")
            return

        try:
            with open("lista_arquivos.csv", "w", newline="") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["Nome do Arquivo", "Caminho"])
                for caminho_arquivo in self.arquivos:
                    writer.writerow([os.path.basename(caminho_arquivo), caminho_arquivo])
            QMessageBox.information(self, "Exportar Lista", "A lista de arquivos foi exportada com sucesso!")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao exportar lista: {str(e)}")

   
   
   
    def classificar_por_nome(self):
        if not self.arquivos:
            QMessageBox.warning(self, "Aviso", "Não há arquivos para classificar!")
            return

        # Classifica por nome
        self.arquivos.sort(key=lambda x: os.path.basename(x).lower())
        self.atualizar_lista()

  
  
  
    def classificar_por_data(self):
        if not self.arquivos:
            QMessageBox.warning(self, "Aviso", "Não há arquivos para classificar!")
            return
        self.arquivos.sort(key=lambda x: os.path.getmtime(x))
        self.atualizar_lista()

    
    
    
    def atualizar_lista(self):
        # Limpa a lista atual
        for i in reversed(range(self.file_layout.count())):
            widget = self.file_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)
        for caminho_arquivo in self.arquivos:
            self.adicionar_item_lista(caminho_arquivo)



def main():
    app = QApplication(sys.argv)
    janela = ArtStationStyleApp()
    janela.show()
    sys.exit(app.exec_())







if __name__ == "__main__":
    main()
