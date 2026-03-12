import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QTabWidget, QGroupBox, QLabel, 
                             QLineEdit, QPushButton, QTextEdit, QComboBox,
                             QGridLayout, QScrollArea, QMessageBox, QFrame,
                             QTableWidget, QTableWidgetItem, QHeaderView)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

class OrcamentoApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.calculator = None
        self.initUI()
        
    def format_currency(self, value):
        """Formata valor para o padrão brasileiro R$ 10.000,00"""
        try:
            return f"R$ {float(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except (ValueError, TypeError):
            return "R$ 0,00"
        
    def initUI(self):
        self.setWindowTitle('Sistema de Orçamento - Galpões Pré-Moldados')
        self.setGeometry(100, 100, 1400, 900)
        
        # Widget central com abas
        self.tabs = QTabWidget()
        
        # Cria as abas
        self.tab_entrada = self.create_entrada_tab()
        self.tab_banco_dados = self.create_banco_dados_tab()
        self.tab_resultados = self.create_resultados_tab()
        self.tab_orcamento = self.create_orcamento_tab()
        
        self.tabs.addTab(self.tab_entrada, "Dados de Entrada")
        self.tabs.addTab(self.tab_banco_dados, "Banco de Dados")
        self.tabs.addTab(self.tab_resultados, "Resultados Detalhados")
        self.tabs.addTab(self.tab_orcamento, "Orçamento Final")
        
        self.setCentralWidget(self.tabs)
        
    def create_entrada_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        # Área de scroll para muitos campos
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        
        # Grupo: Dados Básicos
        grupo_basico = QGroupBox("Dados Básicos do Galpão")
        layout_basico = QGridLayout()
        
        self.entries = {}
        
        # Campos de entrada - AGORA VAZIOS
        campos = [
            ('Frente (m):', 'frente', ''),
            ('Lateral (m):', 'lateral', ''),
            ('Altura Livre (m):', 'altura', ''),
            ('Nome do Cliente:', 'cliente_nome', ''),
            ('Telefone:', 'telefone', ''),
            ('Cidade:', 'cidade', 'IRINEÓPOLIS'),
        ]
        
        for row, (label, key, default) in enumerate(campos):
            layout_basico.addWidget(QLabel(label), row, 0)
            if key == 'cidade':
                combo = QComboBox()
                combo.addItems([
                    'IRINEÓPOLIS', 'PORTO UNIÃO', 'UNIÃO DA VITÓRIA', 'CANOINHAS',
                    'TRÊS BARRAS', 'SÃO MATEUS DO SUL', 'BELA VISTA DO TOLDO',
                    'PAULA FREITAS', 'PAULO FRONTIN', 'RIO AZUL', 'MALLET',
                    'CRUZ MACHADO', 'PAPANDUVA', 'MATOS COSTA', 'TIMBÓ GRANDE',
                    'MAFRA', 'CALMON', 'MAJOR VIEIRA', 'SANTA CECÍLIA', 'BITURUNA'
                ])
                combo.setCurrentText(default)
                self.entries[key] = combo
                layout_basico.addWidget(combo, row, 1)
            else:
                entry = QLineEdit()
                if key == 'frente':
                    entry.setPlaceholderText('Ex: 10')
                elif key == 'lateral':
                    entry.setPlaceholderText('Ex: 20')
                elif key == 'altura':
                    entry.setPlaceholderText('Ex: 5')
                elif key == 'telefone':
                    entry.setPlaceholderText('(47) 99999-9999')
                self.entries[key] = entry
                layout_basico.addWidget(entry, row, 1)
        
        grupo_basico.setLayout(layout_basico)
        scroll_layout.addWidget(grupo_basico)
        
        # Grupo: Especificações Técnicas
        grupo_espec = QGroupBox("Especificações Técnicas")
        layout_espec = QGridLayout()
        
        # Tipo de telha
        layout_espec.addWidget(QLabel("Tipo de Telha:"), 0, 0)
        self.combo_telha = QComboBox()
        self.combo_telha.addItems([
            'TELHA SIMPLES',
            'TELHA ISOPOR E MANTA', 
            'TELHA BANDEJA',
            'TELHA SANDUICHE',
            'TELHA SIMPLES TP 25 Pintada',
            'TELHA COM ISOPOR',
            'TELHA+MANTA',
            'TELHA+EPS+PINTADA'
        ])
        layout_espec.addWidget(self.combo_telha, 0, 1)
        
        # Tipo de cobertura
        layout_espec.addWidget(QLabel("Tipo de Cobertura:"), 1, 0)
        self.combo_cobertura = QComboBox()
        self.combo_cobertura.addItems(['UMA ÁGUA', 'DUAS ÁGUAS'])
        self.combo_cobertura.setCurrentText('DUAS ÁGUAS')
        layout_espec.addWidget(self.combo_cobertura, 1, 1)
        
        # Opcionais
        layout_espec.addWidget(QLabel("Fechamento:"), 2, 0)
        self.combo_fechamento = QComboBox()
        self.combo_fechamento.addItems(['NÃO', 'SIM'])
        layout_espec.addWidget(self.combo_fechamento, 2, 1)
        
        layout_espec.addWidget(QLabel("Porta:"), 3, 0)
        self.combo_porta = QComboBox()
        self.combo_porta.addItems(['NÃO', 'SIM'])
        layout_espec.addWidget(self.combo_porta, 3, 1)
        
        layout_espec.addWidget(QLabel("Janela:"), 4, 0)
        self.combo_janela = QComboBox()
        self.combo_janela.addItems(['NÃO', 'SIM'])
        layout_espec.addWidget(self.combo_janela, 4, 1)
        
        layout_espec.addWidget(QLabel("Placas:"), 5, 0)
        self.combo_placas = QComboBox()
        self.combo_placas.addItems(['NÃO', 'SIM'])
        layout_espec.addWidget(self.combo_placas, 5, 1)
        
        layout_espec.addWidget(QLabel("Platibanda:"), 6, 0)
        self.combo_platibanda = QComboBox()
        self.combo_platibanda.addItems(['NÃO', 'SIM'])
        layout_espec.addWidget(self.combo_platibanda, 6, 1)
        
        layout_espec.addWidget(QLabel("Projeto:"), 7, 0)
        self.combo_projeto = QComboBox()
        self.combo_projeto.addItems(['NÃO', 'SIM'])
        self.combo_projeto.setCurrentText('SIM')
        layout_espec.addWidget(self.combo_projeto, 7, 1)
        
        layout_espec.addWidget(QLabel("Laje:"), 8, 0)
        self.combo_laje = QComboBox()
        self.combo_laje.addItems(['NÃO', 'SIM'])
        layout_espec.addWidget(self.combo_laje, 8, 1)
        
        layout_espec.addWidget(QLabel("Vigas:"), 9, 0)
        self.combo_vigas = QComboBox()
        self.combo_vigas.addItems(['NÃO', 'SIM'])
        layout_espec.addWidget(self.combo_vigas, 9, 1)
        
        layout_espec.addWidget(QLabel("Portão:"), 10, 0)
        self.combo_portao = QComboBox()
        self.combo_portao.addItems(['NÃO', 'SIM'])
        layout_espec.addWidget(self.combo_portao, 10, 1)
        
        # Com Nota Fiscal
        layout_espec.addWidget(QLabel("Com Nota Fiscal:"), 11, 0)
        self.combo_nota = QComboBox()
        self.combo_nota.addItems(['NÃO', 'SIM'])
        layout_espec.addWidget(self.combo_nota, 11, 1)
        
        grupo_espec.setLayout(layout_espec)
        scroll_layout.addWidget(grupo_espec)
        
        # Grupo: Dimensões do Portão (agora com quantidade)
        self.grupo_portao = QGroupBox("Dimensões do Portão")
        layout_portao = QGridLayout()
        
        layout_portao.addWidget(QLabel("Largura (m):"), 0, 0)
        self.entry_portao_largura = QLineEdit("5")
        layout_portao.addWidget(self.entry_portao_largura, 0, 1)
        
        layout_portao.addWidget(QLabel("Altura (m):"), 1, 0)
        self.entry_portao_altura = QLineEdit("5")
        layout_portao.addWidget(self.entry_portao_altura, 1, 1)
        
        # NOVO: Campo para quantidade de portões
        layout_portao.addWidget(QLabel("Quantidade:"), 2, 0)
        self.entry_portao_quantidade = QLineEdit("1")
        self.entry_portao_quantidade.setPlaceholderText("1")
        layout_portao.addWidget(self.entry_portao_quantidade, 2, 1)
        
        self.grupo_portao.setLayout(layout_portao)
        self.grupo_portao.setVisible(False)
        scroll_layout.addWidget(self.grupo_portao)
        
        # Conecta o sinal
        self.combo_portao.currentTextChanged.connect(self.toggle_grupo_portao)
        
        # Botões de ação
        grupo_acoes = QGroupBox("Ações")
        layout_acoes = QHBoxLayout()
        
        btn_calcular = QPushButton("Calcular Orçamento")
        btn_calcular.clicked.connect(self.calcular_orcamento)
        btn_calcular.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; font-weight: bold; }")
        
        btn_limpar = QPushButton("Limpar Campos")
        btn_limpar.clicked.connect(self.limpar_campos)
        btn_limpar.setStyleSheet("QPushButton { background-color: #f44336; color: white; }")
        
        layout_acoes.addWidget(btn_calcular)
        layout_acoes.addWidget(btn_limpar)
        grupo_acoes.setLayout(layout_acoes)
        scroll_layout.addWidget(grupo_acoes)
        
        scroll_layout.addStretch()
        scroll.setWidget(scroll_content)
        layout.addWidget(scroll)
        
        widget.setLayout(layout)
        return widget
        
    def toggle_grupo_portao(self, texto):
        """Mostra ou oculta o grupo de dimensões do portão"""
        self.grupo_portao.setVisible(texto == 'SIM')
        
    def create_banco_dados_tab(self):
        """Cria aba para editar o banco de dados de materiais"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        # Título
        title = QLabel("BANCO DE DADOS - PREÇOS DE MATERIAIS")
        title.setStyleSheet("font-weight: bold; font-size: 16px; color: #2E86AB;")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        # Instruções
        instrucoes = QLabel("Edite os valores dos materiais abaixo. Os valores serão usados nos cálculos do orçamento.")
        instrucoes.setStyleSheet("color: #666; font-style: italic;")
        layout.addWidget(instrucoes)
        
        # Informação do custo do concreto
        self.label_custo_concreto = QLabel("Custo do concreto: R$ 0,00/m³")
        self.label_custo_concreto.setStyleSheet("font-weight: bold; color: #2E86AB;")
        layout.addWidget(self.label_custo_concreto)
        
        # Tabela de materiais
        self.tabela_materiais = QTableWidget()
        self.tabela_materiais.setColumnCount(3)
        self.tabela_materiais.setHorizontalHeaderLabels(["Material", "Valor Unitário (R$)", "Medida de Compra"])
        
        # Configurar tabela
        header = self.tabela_materiais.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        
        layout.addWidget(self.tabela_materiais)
        
        # Botões de ação
        grupo_botoes = QGroupBox("Ações")
        layout_botoes = QHBoxLayout()
        
        btn_carregar = QPushButton("Carregar Valores Atuais")
        btn_carregar.clicked.connect(self.carregar_banco_dados)
        
        btn_salvar = QPushButton("Salvar Alterações")
        btn_salvar.clicked.connect(self.salvar_banco_dados)
        btn_salvar.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; }")
        
        btn_restaurar = QPushButton("Restaurar Valores Padrão")
        btn_restaurar.clicked.connect(self.restaurar_banco_dados)
        btn_restaurar.setStyleSheet("QPushButton { background-color: #FF9800; color: white; }")
        
        layout_botoes.addWidget(btn_carregar)
        layout_botoes.addWidget(btn_salvar)
        layout_botoes.addWidget(btn_restaurar)
        grupo_botoes.setLayout(layout_botoes)
        layout.addWidget(grupo_botoes)
        
        # Carrega dados iniciais
        self.carregar_banco_dados()
        
        widget.setLayout(layout)
        return widget
        
    def carregar_banco_dados(self):
        """Carrega os dados do banco de dados para a tabela com 3 colunas"""
        if not self.calculator:
            from calculator import OrcamentoCalculator
            self.calculator = OrcamentoCalculator()
            
        banco_dados = self.calculator.get_banco_dados()
        
        # Atualiza informação do custo do concreto
        custo_concreto = self.calculator.get_material_value('CONCRETO_M3')
        self.label_custo_concreto.setText(f"Custo do concreto: {self.format_currency(custo_concreto)}/m³")
        
        # Ordena materiais por categoria
        categorias = {
            'Telhas': [k for k in banco_dados.keys() if 'TELHA' in k or 'CHAPA FRIZADA' in k],
            'Perfis': [k for k in banco_dados.keys() if 'PERFIL' in k],
            'Terças': [k for k in banco_dados.keys() if 'TERÇA' in k],
            'Ferros': [k for k in banco_dados.keys() if 'FERRO' in k],
            'Acessórios': [k for k in banco_dados.keys() if k in ['BARRA RED 1/4 VERGA', 'BARRA RED 5/16 VERGA', 'CUMEEIRA', 'TIRANTE', 'CHAPA', 'ESPAÇADORES', 'TELA_4.2']],
            'Concreto e Agregados': [k for k in banco_dados.keys() if k in ['CIMENTO', 'AREIA', 'BRITA/PÓ', 'ADITIVO 730 CAA SUPERPLASTIFICANTE', 'ACELERADOR (SECANTE)', 'DESMOLDANTE', 'ÁGUA', 'CONCRETO_M3']],
            'Outros': [k for k in banco_dados.keys() if k in ['PAREDE_METALICA']]
        }
        
        # Conta total de itens
        total_itens = sum(len(itens) for itens in categorias.values())
        self.tabela_materiais.setRowCount(total_itens)
        
        row = 0
        for categoria, itens in categorias.items():
            # Adiciona linha de categoria
            if itens:
                item_categoria = QTableWidgetItem(f"--- {categoria} ---")
                item_categoria.setBackground(Qt.lightGray)
                item_categoria.setFlags(Qt.ItemIsEnabled)
                self.tabela_materiais.setItem(row, 0, item_categoria)
                self.tabela_materiais.setItem(row, 1, QTableWidgetItem(""))
                self.tabela_materiais.setItem(row, 2, QTableWidgetItem(""))
                row += 1
                
                # Adiciona itens da categoria
                for material in sorted(itens):
                    item_data = banco_dados[material]
                    valor = item_data['valor']
                    medida = item_data['medida']
                    
                    item_material = QTableWidgetItem(material)
                    self.tabela_materiais.setItem(row, 0, item_material)
                    
                    item_valor = QTableWidgetItem(f"{valor:.2f}")
                    self.tabela_materiais.setItem(row, 1, item_valor)
                    
                    item_medida = QTableWidgetItem(medida)
                    self.tabela_materiais.setItem(row, 2, item_medida)
                    
                    row += 1
        
    def salvar_banco_dados(self):
        """Salva as alterações do banco de dados com a nova estrutura"""
        try:
            updates = {}
            
            for row in range(self.tabela_materiais.rowCount()):
                material_item = self.tabela_materiais.item(row, 0)
                valor_item = self.tabela_materiais.item(row, 1)
                
                if material_item and valor_item and material_item.text() and not material_item.text().startswith('---'):
                    material = material_item.text()
                    try:
                        valor = float(valor_item.text())
                        updates[material] = valor
                    except ValueError:
                        QMessageBox.warning(self, "Aviso", f"Valor inválido para {material}")
            
            if updates:
                self.calculator.update_banco_dados(updates)
                
                # Atualiza informação do custo do concreto
                custo_concreto = self.calculator.get_material_value('CONCRETO_M3')
                self.label_custo_concreto.setText(f"Custo do concreto: {self.format_currency(custo_concreto)}/m³")
                
                QMessageBox.information(self, "Sucesso", 
                                      f"Banco de dados atualizado com sucesso!\n"
                                      f"Novo custo do concreto: {self.format_currency(custo_concreto)}/m³")
                
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao salvar banco de dados: {str(e)}")
            
    def restaurar_banco_dados(self):
        """Restaura os valores padrão do banco de dados"""
        resposta = QMessageBox.question(self, "Confirmar", 
                                      "Deseja restaurar os valores padrão do banco de dados?",
                                      QMessageBox.Yes | QMessageBox.No)
        
        if resposta == QMessageBox.Yes:
            from calculator import OrcamentoCalculator
            self.calculator = OrcamentoCalculator()
            self.carregar_banco_dados()
            QMessageBox.information(self, "Sucesso", "Valores padrão restaurados!")
        
    def create_resultados_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        self.texto_resultados = QTextEdit()
        self.texto_resultados.setReadOnly(True)
        self.texto_resultados.setFont(QFont("Courier", 10))
        
        layout.addWidget(QLabel("Resultados Detalhados dos Cálculos:"))
        layout.addWidget(self.texto_resultados)
        
        widget.setLayout(layout)
        return widget
        
    def create_orcamento_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        # Grupo do orçamento final
        grupo_orcamento = QGroupBox("ORÇAMENTO FINAL")
        grupo_orcamento.setStyleSheet("QGroupBox { font-weight: bold; color: #2E86AB; }")
        layout_orcamento = QVBoxLayout()
        
        self.labels_orcamento = {}
        
        # Campos do orçamento
        campos_orcamento = [
            ('Área do Galpão:', 'area'),
            ('Custo Cobertura:', 'custo_cobertura'),
            ('Custo Estrutura:', 'custo_estrutura'), 
            ('Custo Complementos:', 'custo_complementos'),
            ('Custo Projeto:', 'custo_projeto'),
            ('Frete:', 'frete'),
            ('Custo Total + Frete:', 'custo_total_com_frete'),
            ('VALOR TOTAL DO ORÇAMENTO:', 'valor_total')
        ]
        
        for label_text, key in campos_orcamento:
            frame = QFrame()
            frame_layout = QHBoxLayout()
            label_widget = QLabel(label_text)
            if 'TOTAL' in label_text:
                label_widget.setStyleSheet("font-weight: bold; color: #2E86AB;")
            frame_layout.addWidget(label_widget)
            
            value_label = QLabel("R$ 0,00")
            if 'TOTAL' in label_text:
                value_label.setStyleSheet("font-weight: bold; font-size: 14px; color: #2E86AB;")
            value_label.setAlignment(Qt.AlignRight)
            frame_layout.addWidget(value_label)
            frame_layout.addStretch()
            frame.setLayout(frame_layout)
            layout_orcamento.addWidget(frame)
            
            self.labels_orcamento[key] = value_label
        
        grupo_orcamento.setLayout(layout_orcamento)
        layout.addWidget(grupo_orcamento)
        
        # Informações do frete
        grupo_frete = QGroupBox("Informações de Frete")
        layout_frete = QVBoxLayout()
        
        self.label_distancia = QLabel("Distância: 0 km")
        self.label_viagens = QLabel("Total de viagens: 0")
        self.label_valor_frete = QLabel("Valor do frete: R$ 0,00")
        
        layout_frete.addWidget(self.label_distancia)
        layout_frete.addWidget(self.label_viagens)
        layout_frete.addWidget(self.label_valor_frete)
        grupo_frete.setLayout(layout_frete)
        layout.addWidget(grupo_frete)
        
        # Observação sobre margem e nota
        grupo_observacao = QGroupBox("Observações")
        layout_observacao = QVBoxLayout()
        self.label_observacao = QLabel("• Margem de lucro aplicada: 100% sobre o custo total + frete\n• Nota fiscal: NÃO")
        self.label_observacao.setStyleSheet("color: #666; font-style: italic;")
        layout_observacao.addWidget(self.label_observacao)
        grupo_observacao.setLayout(layout_observacao)
        layout.addWidget(grupo_observacao)
        
        # Botões de exportação
        grupo_export = QGroupBox("Exportar Orçamento")
        layout_export = QHBoxLayout()
        
        btn_export_excel = QPushButton("Exportar para Excel")
        btn_export_excel.clicked.connect(self.exportar_excel)
        
        btn_export_pdf = QPushButton("Exportar para PDF") 
        btn_export_pdf.clicked.connect(self.exportar_pdf)
        
        layout_export.addWidget(btn_export_excel)
        layout_export.addWidget(btn_export_pdf)
        grupo_export.setLayout(layout_export)
        layout.addWidget(grupo_export)
        
        # Condições comerciais
        grupo_condicoes = QGroupBox("Condições Comerciais")
        layout_condicoes = QVBoxLayout()
        
        condicoes_texto = """
        • PAGAMENTO:
        • 50% DE ENTRADA NA ASSINATURA DO CONTRATO;
        • 25% NO INICIO DA OBRA;
        • 25% NA CONCLUSÃO DA OBRA;
        
        • PRAZO:
        • 90 dias úteis após a assinatura do contrato;
        • Em caso de projeto, 90 dias após a entrega do projeto;
        
        Não inclui:
        • Sem piso, sem calha, sem fechamento e sem esquadrias
        • Agregados e cimento para fundação por conta do contratante
        • Se a fundação exceder 5 metros fica a custo do cliente
        """
        
        label_condicoes = QLabel(condicoes_texto)
        layout_condicoes.addWidget(label_condicoes)
        grupo_condicoes.setLayout(layout_condicoes)
        layout.addWidget(grupo_condicoes)
        
        widget.setLayout(layout)
        return widget
        
    def calcular_orcamento(self):
        try:
            # Inicializa calculadora se não existir
            if not self.calculator:
                from calculator import OrcamentoCalculator
                self.calculator = OrcamentoCalculator()
            
            # Coleta dados de entrada - com tratamento para campos vazios
            def get_float(entry, default=0):
                try:
                    return float(entry.text()) if entry.text().strip() else default
                except ValueError:
                    return default
            
            dados_entrada = {
                'D4': get_float(self.entries['frente'], 0),
                'D5': get_float(self.entries['lateral'], 0),
                'D6': get_float(self.entries['altura'], 5),  # Altura padrão 5 se vazio
                'C9': self.entries['cliente_nome'].text(),
                'C10': self.entries['cidade'].currentText(),
                'TELEFONE': self.entries['telefone'].text(),
                'C15': self.combo_telha.currentText(),
                'C16': self.combo_cobertura.currentText(),
                'C11': self.combo_fechamento.currentText(),
                'C12': self.combo_porta.currentText(),
                'C13': self.combo_janela.currentText(),
                'C14': self.combo_placas.currentText(),
                'C17': self.combo_platibanda.currentText(),
                'C18': self.combo_projeto.currentText(),
                'C19': self.combo_laje.currentText(),
                'C20': self.combo_vigas.currentText(),
                'C21': self.combo_portao.currentText(),
                'COM_NOTA': self.combo_nota.currentText(),
                'PORTAO_LARGURA': get_float(self.entry_portao_largura, 5),
                'PORTAO_ALTURA': get_float(self.entry_portao_altura, 5),
                'PORTAO_QUANTIDADE': get_float(self.entry_portao_quantidade, 1),
            }
            
            # Valida campos obrigatórios
            if dados_entrada['D4'] <= 0 or dados_entrada['D5'] <= 0:
                QMessageBox.warning(self, "Aviso", "Preencha a frente e lateral do galpão!")
                return
            
            # Define valores na calculadora
            for cell, value in dados_entrada.items():
                self.calculator.set_input_value('ENTRADA', cell, value)
            
            # Executa cálculos
            self.calculator.calculate_all()
            
            # Atualiza resultados
            self.atualizar_resultados()
            self.atualizar_orcamento_final()
            self.atualizar_info_frete()
            
            # Muda para aba de resultados
            self.tabs.setCurrentIndex(2)
            
            self.statusBar().showMessage('Orçamento calculado com sucesso!')
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao calcular orçamento: {str(e)}")
            
    def atualizar_resultados(self):
        if not self.calculator:
            return
            
        resultado_texto = "=== RESULTADOS DETALHADOS ===\n\n"
        
        # Dados básicos
        entrada = self.calculator.results['ENTRADA']
        resultado_texto += f"DADOS BÁSICOS:\n"
        resultado_texto += f"Cliente: {entrada.get('C9', '')}\n"
        resultado_texto += f"Telefone: {entrada.get('TELEFONE', '')}\n"
        resultado_texto += f"Cidade: {entrada.get('C10', '')}\n"
        resultado_texto += f"Medida: {entrada.get('D2', '')}\n"
        resultado_texto += f"Área: {entrada.get('D3', 0):.2f} m²\n"
        resultado_texto += f"Telha: {entrada.get('C15', '')}\n"
        resultado_texto += f"Cobertura: {entrada.get('C16', '')}\n"
        resultado_texto += f"Com Nota: {entrada.get('COM_NOTA', 'NÃO')}\n\n"
        
        # Cobertura
        cobertura = self.calculator.results['COBERTURA']
        resultado_texto += f"COBERTURA:\n"
        resultado_texto += f"Quantidade de Tesouras: {cobertura.get('C24', 0)}\n"
        resultado_texto += f"Custo Tesouras: {self.format_currency(cobertura.get('C25', 0))}\n"
        resultado_texto += f"Metragem de Telha: {cobertura.get('C30', 0):.1f} m²\n"
        resultado_texto += f"Custo Telhas: {self.format_currency(cobertura.get('C31', 0))}\n"
        resultado_texto += f"Custo Terças: {self.format_currency(cobertura.get('C40', 0))}\n"
        resultado_texto += f"Custo Eitão: {self.format_currency(cobertura.get('C47', 0))}\n"
        resultado_texto += f"Custo Cumeeira: {self.format_currency(cobertura.get('C52', 0))}\n"
        
        # Estrutura
        pilares = self.calculator.results['PILARES']
        vigas = self.calculator.results['VIGAS']
        calices = self.calculator.results['CALICES']
        resultado_texto += f"ESTRUTURA:\n"
        resultado_texto += f"Quantidade de Pilares: {pilares.get('C11', 0)}\n"
        resultado_texto += f"Altura dos Pilares: {pilares.get('C12', 0):.1f} m\n"
        resultado_texto += f"Volume de Concreto Pilares: {pilares.get('volume_concreto', 0):.2f} m³\n"
        resultado_texto += f"Custo Pilares: {self.format_currency(pilares.get('C22', 0))}\n"
        resultado_texto += f"Custo Cálices: {self.format_currency(calices.get('C11', 0))}\n"
        resultado_texto += f"Metragem de Vigas: {vigas.get('C12', 0):.1f} m\n"
        resultado_texto += f"Custo Vigas: {self.format_currency(vigas.get('C22', 0))}\n\n"
        
        # Complementos
        fechamentos = self.calculator.results['FECHAMENTOS']
        laje = self.calculator.results['LAJE']
        portao = self.calculator.results['PORTAO']
        resultado_texto += f"COMPLEMENTOS:\n"
        resultado_texto += f"Fechamento: {entrada.get('C11', 'NÃO')} - {self.format_currency(fechamentos.get('C9', 0))}\n"
        resultado_texto += f"Platibanda: {entrada.get('C17', 'NÃO')} - {self.format_currency(fechamentos.get('PLATIBANDA', 0))}\n"
        resultado_texto += f"Laje: {entrada.get('C19', 'NÃO')} - {self.format_currency(laje.get('B5', 0))}\n"
        resultado_texto += f"Portão: {entrada.get('C21', 'NÃO')} - {self.format_currency(portao.get('C11', 0))}\n\n"
        
        # Projeto
        orcamento_final = self.calculator.get_orcamento_final()
        resultado_texto += f"PROJETO:\n"
        resultado_texto += f"Incluído: {entrada.get('C18', 'SIM')} - {self.format_currency(orcamento_final.get('custo_projeto', 0))}\n\n"
        
        # Frete
        frete = self.calculator.results['FRETE']
        resultado_texto += f"FRETE:\n"
        resultado_texto += f"Distância: {frete.get('distancia', 0)} km\n"
        resultado_texto += f"Total de viagens: {frete.get('total_viagens', 0)}\n"
        resultado_texto += f"Valor do frete: {self.format_currency(frete.get('valor_final', 0))}\n\n"
        
        # Custo do concreto
        custo_concreto_info = self.calculator.get_custo_concreto_info()
        resultado_texto += f"CUSTO DO CONCRETO:\n"
        resultado_texto += f"Valor por m³: {self.format_currency(custo_concreto_info['custo_por_m3'])}\n\n"
        
        # Margem de lucro
        resultado_texto += f"RESUMO FINAL:\n"
        resultado_texto += f"Custo Total Materiais: {self.format_currency(orcamento_final.get('custo_total_materiais', 0))}\n"
        resultado_texto += f"Frete: {self.format_currency(orcamento_final.get('frete', 0))}\n"
        resultado_texto += f"Custo Total + Frete: {self.format_currency(orcamento_final.get('custo_total_com_frete', 0))}\n"
        resultado_texto += f"Margem aplicada: 100%\n"
        if orcamento_final.get('com_nota', False):
            resultado_texto += f"Acréscimo Nota Fiscal: 8%\n"
        resultado_texto += f"VALOR FINAL: {self.format_currency(orcamento_final.get('valor_venda', 0))}\n"
        
        self.texto_resultados.setText(resultado_texto)
        
    def atualizar_orcamento_final(self):
        """Atualiza os labels do orçamento final"""
        if not self.calculator:
            return
            
        orcamento = self.calculator.get_orcamento_final()
        entrada = self.calculator.results['ENTRADA']
        
        self.labels_orcamento['area'].setText(f"{entrada.get('D3', 0):.2f} m²")
        self.labels_orcamento['custo_cobertura'].setText(self.format_currency(orcamento.get('custo_cobertura', 0)))
        self.labels_orcamento['custo_estrutura'].setText(self.format_currency(orcamento.get('custo_estrutura', 0)))
        self.labels_orcamento['custo_complementos'].setText(self.format_currency(orcamento.get('custo_complementos', 0)))
        self.labels_orcamento['custo_projeto'].setText(self.format_currency(orcamento.get('custo_projeto', 0)))
        self.labels_orcamento['frete'].setText(self.format_currency(orcamento.get('frete', 0)))
        self.labels_orcamento['custo_total_com_frete'].setText(self.format_currency(orcamento.get('custo_com_frete', 0)))
        self.labels_orcamento['valor_total'].setText(self.format_currency(orcamento.get('valor_venda', 0)))
        
        nota_texto = "SIM" if orcamento.get('com_nota', False) else "NÃO"
        self.label_observacao.setText(
            f"• Margem de lucro aplicada: 100% sobre o custo total + frete\n"
            f"• Projeto: calculado separadamente e somado após a margem\n"
            f"• Nota fiscal: {nota_texto}"
    )
    def atualizar_info_frete(self):
        if not self.calculator:
            return
            
        frete = self.calculator.results['FRETE']
        self.label_distancia.setText(f"Distância: {frete.get('distancia', 0)} km")
        self.label_viagens.setText(f"Total de viagens: {frete.get('total_viagens', 0)}")
        self.label_valor_frete.setText(f"Valor do frete: {self.format_currency(frete.get('valor_final', 0))}")
            
    def exportar_excel(self):
        try:
            from export_manager import ExportManager
            exporter = ExportManager(self.calculator)
            success = exporter.export_to_excel("orcamento_exportado.xlsx")
            if success:
                QMessageBox.information(self, "Sucesso", "Orçamento exportado para Excel com sucesso!")
            else:
                QMessageBox.warning(self, "Aviso", "Erro ao exportar para Excel. Verifique se o arquivo não está aberto.")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao exportar para Excel: {str(e)}")
            
    def exportar_pdf(self):
        try:
            from export_manager import ExportManager
            exporter = ExportManager(self.calculator)
            success = exporter.export_to_pdf("orcamento_exportado.pdf")
            if success:
                QMessageBox.information(self, "Sucesso", "Orçamento exportado para PDF com sucesso!")
            else:
                QMessageBox.warning(self, "Aviso", "Erro ao exportar para PDF.")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao exportar para PDF: {str(e)}")
            
    def limpar_campos(self):
        """Limpa todos os campos de entrada"""
        try:
            for key, entry in self.entries.items():
                if isinstance(entry, QLineEdit):
                    entry.clear()
                elif isinstance(entry, QComboBox) and key == 'cidade':
                    entry.setCurrentText('IRINEÓPOLIS')
            
            self.combo_telha.setCurrentIndex(0)
            self.combo_cobertura.setCurrentIndex(1)
            self.combo_fechamento.setCurrentIndex(0)
            self.combo_porta.setCurrentIndex(0)
            self.combo_janela.setCurrentIndex(0)
            self.combo_placas.setCurrentIndex(0)
            self.combo_platibanda.setCurrentIndex(0)
            self.combo_projeto.setCurrentIndex(1)
            self.combo_laje.setCurrentIndex(0)
            self.combo_vigas.setCurrentIndex(0)
            self.combo_portao.setCurrentIndex(0)
            self.combo_nota.setCurrentIndex(0)
            
            self.entry_portao_largura.setText("5")
            self.entry_portao_altura.setText("5")
            self.grupo_portao.setVisible(False)
            
            self.texto_resultados.clear()
            for label in self.labels_orcamento.values():
                label.setText("R$ 0,00")
                
            self.label_distancia.setText("Distância: 0 km")
            self.label_viagens.setText("Total de viagens: 0")
            self.label_valor_frete.setText("Valor do frete: R$ 0,00")
            
            self.label_observacao.setText("• Margem de lucro aplicada: 100% sobre o custo total + frete\n• Nota fiscal: NÃO")
                
            self.statusBar().showMessage('Campos limpos com sucesso!')
            
        except Exception as e:
            QMessageBox.warning(self, "Aviso", f"Erro ao limpar campos: {str(e)}")