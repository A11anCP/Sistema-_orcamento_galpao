import sys
import os
from PyQt5.QtWidgets import QApplication
from ui_main import OrcamentoApp

def main():
    # Verifica se o arquivo Excel existe
    excel_file = "modelo_orçamento_prémoldado.xlsx"
    if not os.path.exists(excel_file):
        print(f"AVISO: Arquivo {excel_file} não encontrado.")
        print("O sistema funcionará com dados de exemplo.")
    
    # Cria aplicação
    app = QApplication(sys.argv)
    
    # Configura estilo
    app.setStyle('Fusion')
    
    # Cria e exibe janela principal
    window = OrcamentoApp()
    window.show()
    
    # Executa aplicação
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()