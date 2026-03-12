import pandas as pd
import re
import math

class OrcamentoCalculator:
    def __init__(self):
        self.data = {
            'ENTRADA': {},
            'BANCO_DE_DADOS': {},
            'COBERTURA': {},
            'PILARES': {},
            'CALICES': {},
            'FECHAMENTOS': {},
            'VIGAS': {},
            'LAJE': {},
            'PORTAO': {},
            'FRETE': {},
            'ORCAMENTO_FINAL': {}
        }
        self.results = {}
        
        # Tabela de distâncias (cidades e km)
        self.distancias = {
            'IRINEÓPOLIS': 0,
            'PORTO UNIÃO': 40,
            'UNIÃO DA VITÓRIA': 42,
            'CANOINHAS': 46,
            'TRÊS BARRAS': 65,
            'SÃO MATEUS DO SUL': 72,
            'BELA VISTA DO TOLDO': 41,
            'PAULA FREITAS': 5,
            'PAULO FRONTIN': 32,
            'RIO AZUL': 74,
            'MALLET': 53,
            'CRUZ MACHADO': 93,
            'PAPANDUVA': 90,
            'MATOS COSTA': 67,
            'TIMBÓ GRANDE': 61,
            'MAFRA': 112,
            'CALMON': 85,
            'MAJOR VIEIRA': 70,
            'SANTA CECÍLIA': 126,
            'BITURUNA': 117
        }
        
        # Inicializa banco de dados com valores padrão do Excel
        self._initialize_banco_dados()
        
    def _initialize_banco_dados(self):
        """Inicializa o banco de dados com valores padrão incluindo medidas de compra"""
        banco = {
            # Telhas (valores por m²)
            'TELHA ISOPOR E MANTA': {'valor': 75.0, 'medida': 'm²'},
            'TELHA BANDEJA': {'valor': 105.8, 'medida': 'm²'},
            'TELHA SANDUICHE': {'valor': 86.8, 'medida': 'm²'},
            'TELHA SIMPLES TP 25 Pintada': {'valor': 48.9, 'medida': 'm²'},
            'TELHA SIMPLES': {'valor': 36.9, 'medida': 'm²'},
            'TELHA COM ISOPOR': {'valor': 51.9, 'medida': 'm²'},
            'TELHA+MANTA': {'valor': 46.9, 'medida': 'm²'},
            'TELHA+EPS+PINTADA': {'valor': 69.0, 'medida': 'm²'},
            'CHAPA FRIZADA': {'valor': 89.9, 'medida': 'm²'},
            
            # Perfis (valores por barra de 6m)
            'PERFIL U 75mm 2,00': {'valor': 106.2, 'medida': 'barra 6m'},
            'PERFIL U 75mm 2,25': {'valor': 117.8, 'medida': 'barra 6m'},
            'PERFIL U 75mm 2,65': {'valor': 136.8, 'medida': 'barra 6m'},
            'PERFIL U 75mm 3': {'valor': 155.8, 'medida': 'barra 6m'},
            'PERFIL U 68mm 2,00': {'valor': 85.9, 'medida': 'barra 6m'},
            'PERFIL U 68mm 2,25': {'valor': 97.3, 'medida': 'barra 6m'},
            'PERFIL U 68mm 3': {'valor': 114.5, 'medida': 'barra 6m'},
            'PERFIL U 92 mm 2,00': {'valor': 103.36, 'medida': 'barra 6m'},
            'PERFIL U 92 mm 2,25': {'valor': 115.5, 'medida': 'barra 6m'},
            'PERFIL U 92 mm 3': {'valor': 148.5, 'medida': 'barra 6m'},
            'PERFIL U 100mm 2,00': {'valor': 123.12, 'medida': 'barra 6m'},
            'PERFIL U 100mm 2,25': {'valor': 139.08, 'medida': 'barra 6m'},
            'PERFIL U 100mm 2,65': {'valor': 162.33, 'medida': 'barra 6m'},
            'PERFIL U 100mm 3': {'valor': 182.85, 'medida': 'barra 6m'},
            
            # Terças (valores por barra de 6m)
            'TERÇA ENR 75mm 2,00': {'valor': 123.12, 'medida': 'barra 6m'},
            'TERÇA ENR 75mm 2,25': {'valor': 137.56, 'medida': 'barra 6m'},
            'TERÇA ENR 75mm 2,65': {'valor': 162.64, 'medida': 'barra 6m'},
            
            # Ferros e acessórios
            'BARRA RED 1/4 VERGA': {'valor': 22.8, 'medida': 'barra 12m'},
            'BARRA RED 5/16 VERGA': {'valor': 35.0, 'medida': 'barra 12m'},
            'CUMEEIRA': {'valor': 37.51, 'medida': 'metro'},
            'TIRANTE': {'valor': 45.0, 'medida': 'unidade'},
            
            # Concreto e agregados (valores por unidade de compra) - COMPLETO
            'CIMENTO': {'valor': 34.0, 'medida': 'saco 40kg'},
            'AREIA': {'valor': 64.0, 'medida': 'm³'},
            'BRITA/PÓ': {'valor': 75.0, 'medida': 'm³'},
            'ADITIVO 730 CAA SUPERPLASTIFICANTE': {'valor': 2111.0, 'medida': '200L'},
            'ACELERADOR (SECANTE)': {'valor': 1400.0, 'medida': '200L'},
            'DESMOLDANTE': {'valor': 2180.0, 'medida': '200L'},
            'ÁGUA': {'valor': 0.0, 'medida': 'L'},
            
            # Ferro por diâmetro (valores por barra de 12m)
            'FERRO_5': {'valor': 15.9, 'medida': 'barra 12m'},
            'FERRO_6.3': {'valor': 26.7, 'medida': 'barra 12m'},
            'FERRO_8': {'valor': 40.61, 'medida': 'barra 12m'},
            'FERRO_10': {'valor': 61.76, 'medida': 'barra 12m'},
            'FERRO_12.5': {'valor': 94.07, 'medida': 'barra 12m'},
            'FERRO_16': {'valor': 162.13, 'medida': 'barra 12m'},
            'FERRO_20': {'valor': 259.59, 'medida': 'barra 12m'},
            
            # Outros
            'CHAPA': {'valor': 27.0, 'medida': 'm²'},
            'ESPAÇADORES': {'valor': 140.0, 'medida': '500 unidades'},
            'TELA_4.2': {'valor': 1686.0, 'medida': '120m'},
            'PAREDE_METALICA': {'valor': 120.0, 'medida': 'm²'},
            
            # Concreto (valor fixo do Excel)
            'CONCRETO_M3': {'valor': 300.0, 'medida': 'm³'},
        }
        
        self.data['BANCO_DE_DADOS'] = banco
        
    def get_material_value(self, material_name):
        """Obtém o valor de um material do banco de dados"""
        banco = self.data['BANCO_DE_DADOS']
        if material_name in banco:
            return banco[material_name]['valor']
        return 0
    
    def get_material_measure(self, material_name):
        """Obtém a medida de um material do banco de dados"""
        banco = self.data['BANCO_DE_DADOS']
        if material_name in banco:
            return banco[material_name]['medida']
        return ''
    
    def get_banco_dados(self):
        """Retorna o banco de dados completo"""
        return self.data['BANCO_DE_DADOS']
    
    def get_banco_dados_completo(self):
        """Retorna o banco de dados completo com estrutura de dicionário"""
        return self.data['BANCO_DE_DADOS']
        
    def update_banco_dados(self, updates):
        """Atualiza valores do banco de dados"""
        for material, valor in updates.items():
            if material in self.data['BANCO_DE_DADOS']:
                self.data['BANCO_DE_DADOS'][material]['valor'] = valor
        
    def format_currency(self, value):
        """Formata valor para o padrão brasileiro R$ 10.000,00"""
        try:
            return f"R$ {float(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except (ValueError, TypeError):
            return "R$ 0,00"
        
    def set_input_value(self, sheet, cell, value):
        """Define um valor de entrada"""
        if sheet not in self.data:
            self.data[sheet] = {}
        self.data[sheet][cell] = value
        
    def calculate_all(self):
        """Executa todos os cálculos baseados nas entradas"""
        self._calculate_entrada()
        self._calculate_banco_dados()
        self._calculate_cobertura()
        self._calculate_pilares()
        self._calculate_calices()
        self._calculate_fechamentos()
        self._calculate_vigas()
        self._calculate_laje()
        self._calculate_portao()
        self._calculate_frete()
        self._calculate_orcamento_final()
        
    def _calculate_entrada(self):
        """Calcula valores da planilha ENTRADA"""
        entrada = self.data['ENTRADA']
        
        # Dados básicos
        frente = entrada.get('D4', 9)
        lateral = entrada.get('D5', 60)
        altura_livre = entrada.get('D6', 5)
        
        # Cálculos básicos (conforme Excel)
        entrada['D2'] = f"{frente}x{lateral}"
        entrada['D3'] = frente * lateral
        
        # Configurações
        entrada['C9'] = entrada.get('C9', '')  # Nome do cliente
        entrada['C10'] = entrada.get('C10', 'IRINEÓPOLIS')  # Cidade padrão
        entrada['C15'] = entrada.get('C15', 'TELHA SIMPLES')
        entrada['C16'] = entrada.get('C16', 'DUAS ÁGUAS')
        entrada['TELEFONE'] = entrada.get('TELEFONE', '')
        entrada['COM_NOTA'] = entrada.get('COM_NOTA', 'NÃO')
        entrada['PORTAO_LARGURA'] = entrada.get('PORTAO_LARGURA', 5)
        entrada['PORTAO_ALTURA'] = entrada.get('PORTAO_ALTURA', 5)
        
        # Opcionais conforme Excel
        entrada['C11'] = entrada.get('C11', 'NÃO')  # Fechamento
        entrada['C12'] = entrada.get('C12', 'NÃO')  # Porta
        entrada['C13'] = entrada.get('C13', 'NÃO')  # Janela
        entrada['C14'] = entrada.get('C14', 'NÃO')  # Placas
        entrada['C17'] = entrada.get('C17', 'NÃO')  # Platibanda
        entrada['C18'] = entrada.get('C18', 'SIM')  # Projeto
        entrada['C19'] = entrada.get('C19', 'NÃO')  # Laje
        entrada['C20'] = entrada.get('C20', 'NÃO')  # Viga
        entrada['C21'] = entrada.get('C21', 'NÃO')  # Portão
        
        self.results['ENTRADA'] = entrada
        
    def _calculate_banco_dados(self):
        """Calcula valores do banco de dados"""
        self.results['BANCO_DE_DADOS'] = self.data['BANCO_DE_DADOS']
        
    def _calculate_cobertura(self):
        """Calcula valores da cobertura conforme Excel"""
        cobertura = {}
        entrada = self.results['ENTRADA']
        banco = self.results['BANCO_DE_DADOS']
        
        frente = entrada.get('D4', 9)
        lateral = entrada.get('D5', 60)
        
        # Dimensões da cobertura (conforme Excel)
        cobertura['B5'] = frente + 1  # Frente + 1
        cobertura['D5'] = lateral + 1  # Lado + 1
        
        # Cálculos de inclinação (para duas águas) - conforme Excel
        altura = cobertura['B5'] / 10
        lado = cobertura['B5'] / 2
        medida_inclinacao = math.sqrt((altura ** 2) + (lado ** 2))
        
        cobertura['C9'] = altura
        cobertura['C10'] = lado
        cobertura['C11'] = medida_inclinacao
        
        # CÁLCULO DE TESOURAS
        # Perímetro da tesoura (conforme Excel C14)
        perimetro_tesoura = cobertura['B5'] + cobertura['C11'] * 2
        
        # Quantidade de perfis por tesoura (conforme Excel C15)
        qntd_perfil_por_tesoura = perimetro_tesoura / 6
        
        # Quantidade de perfis externos (conforme Excel C16)
        qntd_perfil_externo = math.ceil(qntd_perfil_por_tesoura)
        
        # Perfis conforme Excel
        perfil_externo = 'PERFIL U 75mm 2,00'
        perfil_interno = 'PERFIL U 68mm 2,00'
        
        # Valor dos perfis (conforme Excel C18 e C21)
        valor_perfil_externo = self.get_material_value(perfil_externo)
        valor_perfil_interno = self.get_material_value(perfil_interno)
        
        # Quantidade de perfis internos (conforme Excel C19)
        qntd_perfil_interno = qntd_perfil_externo + 1
        
        # Custo por tesoura (conforme Excel C23)
        custo_por_tesoura = (qntd_perfil_externo * valor_perfil_externo) + (qntd_perfil_interno * valor_perfil_interno)
        
        # Quantidade de tesouras (conforme Excel C24)
        qntd_tesouras = math.floor((cobertura['D5'] / 5) + 1)
        
        # Valor total das tesouras (conforme Excel C25)
        valor_total_tesouras = custo_por_tesoura * qntd_tesouras
        
        cobertura['C24'] = qntd_tesouras
        cobertura['C25'] = valor_total_tesouras
        
        # Cálculo de telhas (conforme Excel)
        modelo_telha = entrada.get('C15', 'TELHA SIMPLES')
        valor_telha = self.get_material_value(modelo_telha)
        
        # Metragem de telha (conforme Excel C30)
        metragem_telha = math.ceil((cobertura['C11'] * cobertura['D5']) * 2)
        valor_total_telha = valor_telha * metragem_telha
        
        cobertura['C30'] = metragem_telha
        cobertura['C31'] = valor_total_telha
        
        # Cálculo de terças
        # Valor por metro da terça (conforme Excel C34)
        valor_terca_metro = self.get_material_value('TERÇA ENR 75mm 2,00') / 6
        
        # Metragem da terça (conforme Excel C35)
        metragem_terca = cobertura['D5']
        
        # Quantidade de terças por lado (conforme Excel C36)
        qntd_terca_por_lado = cobertura['C11'] / 1.2
        
        # Quantidade total de terças (conforme Excel C37)
        qntd_total_terca = math.ceil(qntd_terca_por_lado * 2)
        
        # Total de metros de terça (conforme Excel C38)
        total_metros_terca = metragem_terca * qntd_total_terca
        
        # Quantidade de terças (conforme Excel C39)
        qntd_tercas = total_metros_terca / 6
        
        # Valor total das terças (conforme Excel C40)
        valor_total_terca = total_metros_terca * valor_terca_metro
        
        cobertura['C40'] = valor_total_terca
        
        # Cálculo de eitão (conforme Excel)
        area_eitao = cobertura['C9'] * cobertura['C10']  # Altura * Lado
        valor_eitao_m2 = 80  # Valor fixo conforme Excel
        custo_eitao = area_eitao * valor_eitao_m2
        qntd_eitao = 2
        valor_total_eitao = custo_eitao * qntd_eitao
        
        cobertura['C47'] = valor_total_eitao
        
        # Cálculo de cumeeira (conforme Excel)
        metragem_cumeeira = cobertura['D5']
        valor_cumeeira_metro = self.get_material_value('CUMEEIRA')
        valor_total_cumeeira = metragem_cumeeira * valor_cumeeira_metro
        
        cobertura['C52'] = valor_total_cumeeira
        
        self.results['COBERTURA'] = cobertura
        
    def _calculate_pilares(self):
        """Calcula valores dos pilares conforme Excel"""
        pilares = {}
        entrada = self.results['ENTRADA']
        
        numero_pilares = 26
        altura_pilar = entrada.get('D6', 5) + 1  # 6m
        
        # Dimensões do pilar
        largura = 0.26
        altura_secao = 0.30

        # 1. CONCRETO
        volume_por_pilar = 0.26 * 0.30 * altura_pilar  # 0.468 m³
        volume_total = volume_por_pilar * numero_pilares  # 12.168 m³
        # Valor do concreto por m³ (calculado ou fixo)
        valor_concreto_m3 = 426.00  # Extraído do Excel
        custo_concreto = volume_total * valor_concreto_m3

        # 2. AÇO LONGITUDINAL
        # 2 barras de 10mm
        valor_barra_10 = self.get_material_value('FERRO_10') / 2  # Meia barra (6m)
        custo_ferro_10 = 2 * valor_barra_10 * numero_pilares
        
        # 4 barras de 12.5mm
        valor_barra_12_5 = self.get_material_value('FERRO_12.5') / 2  # Meia barra (6m)
        custo_ferro_12_5 = 4 * valor_barra_12_5 * numero_pilares
        
        custo_ferro_long = custo_ferro_10 + custo_ferro_12_5

        # 3. ESTRIBOS
        qtd_estribos = math.ceil((altura_pilar / 0.15) + 1)  # 41
        comprimento_estribo = (largura + altura_secao) * 2 + 0.10  # 1.22m
        comprimento_total = qtd_estribos * comprimento_estribo * numero_pilares
        valor_ferro_5_metro = self.get_material_value('FERRO_5') / 12
        custo_estribos = comprimento_total * valor_ferro_5_metro

        # 4. CHAPA
        custo_chapa = numero_pilares * (self.get_material_value('CHAPA') / 2)

        # 5. GANCHOS
        comprimento_ganchos = 2 * 1.0 * numero_pilares  # 52m
        valor_ferro_10_metro = self.get_material_value('FERRO_10') / 12
        custo_ganchos = comprimento_ganchos * valor_ferro_10_metro

        # CUSTO TOTAL
        custo_total = custo_concreto + custo_ferro_long + custo_estribos + custo_chapa + custo_ganchos

        pilares['C22'] = custo_total
        pilares['C11'] = numero_pilares
        pilares['C12'] = altura_pilar
        
        self.results['PILARES'] = pilares
        return pilares


    def _calculate_calices(self):
        """Calcula valores dos cálices conforme Excel"""
        calices = {}
        numero_pilares = self.results['PILARES'].get('C11', 26)

        # VALOR POR CÁLICE (extraído do Excel)
        valor_por_calice = 200.00
        custo_total = valor_por_calice * numero_pilares

        calices['C11'] = custo_total
        self.results['CALICES'] = calices
        return calices
        
    def _calculate_fechamentos(self):
        """Calcula valores de fechamentos conforme Excel"""
        fechamentos = {}
        entrada = self.results['ENTRADA']
        
        tem_fechamento = entrada.get('C11', 'NÃO') == 'SIM'
        tem_platibanda = entrada.get('C17', 'NÃO') == 'SIM'
        
        if tem_fechamento:
            frente = entrada.get('D4', 9)
            lateral = entrada.get('D5', 60)
            perimetro = (frente + lateral) * 2
            area_fechamento = perimetro * 5.5
            
            custo_fechamento = area_fechamento * self.get_material_value('PAREDE_METALICA')
            fechamentos['C9'] = custo_fechamento
        else:
            fechamentos['C9'] = 0
            
        if tem_platibanda:
            perimetro_platibanda = (entrada.get('D4', 9) + entrada.get('D5', 60)) * 2
            area_platibanda = perimetro_platibanda * 1  # 1m de altura
            custo_platibanda = area_platibanda * 80  # R$ 80/m² estimado
            fechamentos['PLATIBANDA'] = custo_platibanda
        else:
            fechamentos['PLATIBANDA'] = 0
            
        self.results['FECHAMENTOS'] = fechamentos
        
    def _calculate_vigas(self):
        """Calcula valores das vigas conforme Excel"""
        vigas = {}
        entrada = self.results['ENTRADA']
        
        tem_vigas = entrada.get('C20', 'NÃO') == 'SIM'
        
        if tem_vigas:
            numero_vigas = 10
            comprimento_viga = 5
            
            # Concreto (50x22 cm)
            area_viga = 0.5 * 0.22
            volume_viga = area_viga * comprimento_viga
            custo_concreto = volume_viga * 300 * numero_vigas  # Usar 300 fixo
            
            # Aço longitudinal (8 barras de 12.5mm)
            comprimento_ferro = comprimento_viga
            valor_ferro_metro = self.get_material_value('FERRO_12.5') / 12
            custo_ferro = 8 * comprimento_ferro * valor_ferro_metro * numero_vigas
            
            # Estribos
            numero_estribos = math.ceil((comprimento_viga / 0.15) + 1)
            comprimento_estribo = (0.5 + 0.22) * 2 + 0.1
            valor_ferro_estribo = self.get_material_value('FERRO_5') / 12
            custo_estribos = numero_estribos * comprimento_estribo * valor_ferro_estribo * numero_vigas
            
            custo_total_vigas = custo_concreto + custo_ferro + custo_estribos
            vigas['C12'] = numero_vigas * comprimento_viga
            vigas['C22'] = custo_total_vigas
        else:
            vigas['C12'] = 0
            vigas['C22'] = 0
            
        self.results['VIGAS'] = vigas
        
    def _calculate_laje(self):
        """Calcula valores da laje conforme Excel"""
        laje = {}
        entrada = self.results['ENTRADA']
        
        tem_laje = entrada.get('C19', 'NÃO') == 'SIM'
        
        if tem_laje:
            area = entrada.get('D3', 540)
            custo_laje = area * 200  # R$ 200/m² para laje pré-moldada
            laje['B5'] = custo_laje
        else:
            laje['B5'] = 0
            
        self.results['LAJE'] = laje
        
    def _calculate_portao(self):
        """Calcula valores do portão: R$ 350,00 por m² × quantidade"""
        portao = {}
        entrada = self.results['ENTRADA']
        
        tem_portao = entrada.get('C21', 'NÃO') == 'SIM'
        
        if tem_portao:
            largura = float(entrada.get('PORTAO_LARGURA', 5))
            altura = float(entrada.get('PORTAO_ALTURA', 5))
            quantidade = float(entrada.get('PORTAO_QUANTIDADE', 1))
            
            area_por_portao = largura * altura  # m² por portão
            area_total = area_por_portao * quantidade
            valor_por_m2 = 350  # R$ 350,00 por metro quadrado
            
            # Cálculo correto: área total × valor por m²
            custo_portao = area_total * valor_por_m2
            
            portao['C11'] = custo_portao
            portao['area_por_portao'] = area_por_portao
            portao['area_total'] = area_total
            portao['quantidade'] = quantidade
            portao['largura'] = largura
            portao['altura'] = altura
            portao['valor_por_m2'] = valor_por_m2
        else:
            portao['C11'] = 0
            portao['area_por_portao'] = 0
            portao['area_total'] = 0
            portao['quantidade'] = 0
            
        self.results['PORTAO'] = portao
        return portao
        
    def _calculate_frete(self):
        """Calcula o valor do frete conforme Excel"""
        frete = {}
        entrada = self.results['ENTRADA']
        
        cidade = entrada.get('C10', 'IRINEÓPOLIS')
        distancia = self.distancias.get(cidade.upper(), 0)
        
        valor_por_km = 2
        
        viagens_fundacao = 4
        viagens_pilar = 3
        viagens_tesouras = 2
        viagens_placas = 3
        viagens_lajes = 0
        viagens_vigas = 0
        
        total_viagens = viagens_fundacao + viagens_pilar + viagens_tesouras + viagens_placas + viagens_lajes + viagens_vigas
        
        valor_ida = distancia * valor_por_km * total_viagens
        valor_ida_volta = valor_ida * 2
        
        frete['distancia'] = distancia
        frete['total_viagens'] = total_viagens
        frete['valor_ida'] = valor_ida
        frete['valor_ida_volta'] = valor_ida_volta
        frete['valor_final'] = valor_ida_volta
        
        self.results['FRETE'] = frete
        
    def _calculate_orcamento_final(self):
        """Calcula o orçamento final conforme Excel - CORREÇÃO DA ORDEM"""
        orcamento = {}
        entrada = self.results['ENTRADA']
        
        # Soma todos os custos de materiais (SEM o projeto ainda)
        cobertura = self.results['COBERTURA']
        custo_cobertura = cobertura.get('C25', 0) + \
                        cobertura.get('C31', 0) + \
                        cobertura.get('C40', 0) + \
                        cobertura.get('C47', 0) + \
                        cobertura.get('C52', 0) + \
                        cobertura.get('CONTRAVENTAMENTO', 0)
                        
        custo_estrutura = self.results['PILARES'].get('C22', 0) + \
                        self.results['CALICES'].get('C11', 0) + \
                        self.results['VIGAS'].get('C22', 0)
                        
        custo_complementos = self.results['FECHAMENTOS'].get('C9', 0) + \
                        self.results['FECHAMENTOS'].get('PLATIBANDA', 0) + \
                        self.results['LAJE'].get('B5', 0) + \
                        self.results['PORTAO'].get('C11', 0)
        
        # CUSTO TOTAL DOS MATERIAIS (sem projeto)
        custo_materiais = custo_cobertura + custo_estrutura + custo_complementos
        
        # CÁLCULO DO PROJETO (se incluído)
        tem_projeto = entrada.get('C18', 'SIM') == 'SIM'
        custo_projeto = 0
        if tem_projeto:
            area = entrada.get('D3', 0)
            custo_projeto = area * 15  # R$ 15/m²
        
        # FRETE
        frete = self.results['FRETE'].get('valor_final', 0)
        
        # ORDEM CORRETA DOS CÁLCULOS:
        # 1. Primeiro: (custo_materiais + frete) * 2
        # 2. Depois: + custo_projeto
        custo_com_frete = custo_materiais + frete
        valor_com_margem = custo_com_frete * 2  # Margem de 100%
        valor_venda = valor_com_margem + custo_projeto  # Soma o projeto DEPOIS da margem
        
        # Se tiver nota fiscal, aplica sobre o valor total
        com_nota = entrada.get('COM_NOTA', 'NÃO') == 'SIM'
        if com_nota:
            valor_venda = valor_venda * 1.08
        
        # Salva todos os valores para referência
        orcamento['custo_materiais'] = custo_materiais
        orcamento['custo_cobertura'] = custo_cobertura
        orcamento['custo_estrutura'] = custo_estrutura
        orcamento['custo_complementos'] = custo_complementos
        orcamento['custo_projeto'] = custo_projeto
        orcamento['frete'] = frete
        orcamento['custo_com_frete'] = custo_com_frete
        orcamento['valor_com_margem'] = valor_com_margem
        orcamento['valor_venda'] = valor_venda
        orcamento['com_nota'] = com_nota
        
        self.results['ORCAMENTO_FINAL'] = orcamento
        return orcamento
        
    def get_result(self, sheet, cell):
        """Obtém um resultado específico"""
        if sheet in self.results:
            return self.results[sheet].get(cell, 0)
        return 0
        
    def get_orcamento_final(self):
        """Retorna o orçamento final"""
        return self.results.get('ORCAMENTO_FINAL', {})

    def get_custo_concreto_info(self):
        """Retorna informações detalhadas sobre o custo do concreto"""
        info = {
            'custo_por_m3': 300.0,  # Valor fixo do Excel
            'componentes': {
                'CIMENTO': self.get_material_value('CIMENTO'),
                'AREIA': self.get_material_value('AREIA'),
                'BRITA/PÓ': self.get_material_value('BRITA/PÓ'),
                'ADITIVO': self.get_material_value('ADITIVO 730 CAA SUPERPLASTIFICANTE'),
                'ACELERADOR': self.get_material_value('ACELERADOR (SECANTE)'),
                'DESMOLDANTE': self.get_material_value('DESMOLDANTE')
            }
        }
        return info
    
    def _calcular_custo_concreto_exato(self):
        """Calcula o custo exato do concreto baseado nos componentes do Excel"""
        # Componentes por m³ de concreto (do Excel)
        cimento_kg = 320  # kg
        areia_kg = 655    # kg
        brita_kg = 300    # kg
        aditivo_l = 1.2   # L
        acelerador_l = 1.0 # L
        desmoldante_l = 0.05 # L
        
        # Valores dos materiais
        valor_cimento_saco = self.get_material_value('CIMENTO')  # 34.00 / 40kg
        valor_cimento_kg = valor_cimento_saco / 40
        
        valor_areia_m3 = self.get_material_value('AREIA')  # 64.00 / m³
        # Densidade da areia ~1600 kg/m³
        valor_areia_kg = valor_areia_m3 / 1600
        
        valor_brita_m3 = self.get_material_value('BRITA/PÓ')  # 75.00 / m³
        valor_brita_kg = valor_brita_m3 / 1600
        
        valor_aditivo_200l = self.get_material_value('ADITIVO 730 CAA SUPERPLASTIFICANTE')
        valor_aditivo_l = valor_aditivo_200l / 200
        
        valor_acelerador_200l = self.get_material_value('ACELERADOR (SECANTE)')
        valor_acelerador_l = valor_acelerador_200l / 200
        
        valor_desmoldante_200l = self.get_material_value('DESMOLDANTE')
        valor_desmoldante_l = valor_desmoldante_200l / 200
        
        # Cálculo do custo por m³
        custo_m3 = (cimento_kg * valor_cimento_kg +
                    areia_kg * valor_areia_kg +
                    brita_kg * valor_brita_kg +
                    aditivo_l * valor_aditivo_l +
                    acelerador_l * valor_acelerador_l +
                    desmoldante_l * valor_desmoldante_l)
        
        return custo_m3