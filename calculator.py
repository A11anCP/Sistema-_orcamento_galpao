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
            
            # Materiais adicionais do concreto conforme Excel
            'CONCRETO_M3': {'valor': 300.0, 'medida': 'm³'},  # Valor fixo do Excel
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
        """Calcula valores da cobertura conforme Excel - CORREÇÃO DA TESOURA"""
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
        
        # CÁLCULO DE TESOURAS - CORRIGIDO CONFORME EXCEL
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
        
        # Cálculo de terças - CORREÇÃO CONFORME EXCEL
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
        
        # Contraventamento (estimativa)
        valor_contraventamento = metragem_telha * 5  # Estimativa R$ 5 por m² de telha
        
        cobertura['CONTRAVENTAMENTO'] = valor_contraventamento
        
        self.results['COBERTURA'] = cobertura
        
    def _calculate_pilares(self):
        """Calcula valores dos pilares EXATAMENTE como no Excel"""
        pilares = {}
        entrada = self.results['ENTRADA']
        
        # CONFIGURAÇÃO CONFORME EXCEL
        numero_pilares = 26
        altura_pilar = entrada.get('D6', 5) + 1  # Altura livre + fundação
        
        # DIMENSÕES DO PILAR (26x30 cm) - CONFORME EXCEL
        largura = 0.26  # m
        altura_secao = 0.30   # m
        
        # CÁLCULO DE CONCRETO - EXATO DO EXCEL
        area_base = largura * altura_secao  # 0.078 m²
        volume_pilar = area_base * altura_pilar  # m³ por pilar
        volume_total_concreto = volume_pilar * numero_pilares
        
        # CUSTO CONCRETO - USANDO VALOR FIXO DO EXCEL
        custo_concreto_por_m3 = 300  # R$ 300/m³ conforme Excel
        custo_concreto = volume_total_concreto * custo_concreto_por_m3
        
        # AÇO LONGITUDINAL - EXATO DO EXCEL
        # 4 barras de 12.5mm por pilar
        numero_barras_longitudinais = 4
        comprimento_barra_longitudinal = altura_pilar
        
        # LÓGICA DO EXCEL PARA VALOR DA BARRA
        valor_ferro_12_5_barra = self.get_material_value('FERRO_12.5')  # 94.07
        if comprimento_barra_longitudinal <= 6:
            valor_por_barra_longitudinal = valor_ferro_12_5_barra / 2  # 47.035 para barra de 6m
        else:
            valor_por_barra_longitudinal = valor_ferro_12_5_barra  # 94.07 para barra de 12m
        
        custo_ferro_longitudinal = numero_barras_longitudinais * valor_por_barra_longitudinal * numero_pilares
        
        # ESTRIBOS - EXATO DO EXCEL
        numero_estribos_por_pilar = math.ceil((altura_pilar / 0.15) + 1)
        comprimento_estribo = (largura + altura_secao) * 2 + 0.10  # Perímetro + 10cm dobra
        
        valor_ferro_5_barra = self.get_material_value('FERRO_5')  # 15.9
        valor_ferro_5_metro = valor_ferro_5_barra / 12  # 1.325 por metro
        
        comprimento_total_estribos = numero_estribos_por_pilar * comprimento_estribo * numero_pilares
        custo_estribos = comprimento_total_estribos * valor_ferro_5_metro
        
        # CHAPA - EXATO DO EXCEL
        # No Excel: =C11*('BANCO DE DADOS'!J14/2) = 26 * (27/2) = 26 * 13.5 = 351
        custo_chapa = numero_pilares * (self.get_material_value('CHAPA') / 2)
        
        # GANCHOS - EXATO DO EXCEL
        # No Excel: 2 ganchos por pilar de 1m cada, usando FERRO_10
        numero_ganchos_por_pilar = 2
        comprimento_gancho = 1.0  # 1m cada
        
        valor_ferro_10_barra = self.get_material_value('FERRO_10')  # 61.76
        valor_ferro_10_metro = valor_ferro_10_barra / 12  # 5.1467 por metro
        
        comprimento_total_ganchos = numero_ganchos_por_pilar * comprimento_gancho * numero_pilares
        custo_ganchos = comprimento_total_ganchos * valor_ferro_10_metro
        
        # CUSTO TOTAL - SOMANDO TODOS CONFORME EXCEL
        custo_total_pilar = (custo_concreto + 
                            custo_ferro_longitudinal + 
                            custo_estribos + 
                            custo_chapa + 
                            custo_ganchos)
        
        pilares['C11'] = numero_pilares
        pilares['C12'] = altura_pilar
        pilares['volume_concreto'] = volume_total_concreto
        pilares['comprimento_ferro'] = comprimento_barra_longitudinal * numero_barras_longitudinais * numero_pilares
        pilares['comprimento_estribos'] = comprimento_total_estribos
        pilares['C22'] = custo_total_pilar
        
        self.results['PILARES'] = pilares
        
    def _calculate_calices(self):
        """Calcula valores dos cálices EXATAMENTE como no Excel"""
        calices = {}
        pilares = self.results['PILARES']
        
        numero_pilares = pilares.get('C11', 26)
        
        # CÁLCULO DE CONCRETO - EXATO DO EXCEL
        # Volume por cálice: =(((0.08*1)*0.63)+(0.08*1*0.7))*2
        volume_por_calice = (((0.08 * 1) * 0.63) + (0.08 * 1 * 0.7)) * 2  # 0.2128 m³
        volume_total_concreto = volume_por_calice * numero_pilares
        
        # CUSTO CONCRETO - USANDO VALOR FIXO DO EXCEL
        custo_concreto_por_m3 = 300  # R$ 300/m³
        custo_concreto = volume_total_concreto * custo_concreto_por_m3
        
        # FERRO LONGITUDINAL - EXATO DO EXCEL
        # 4 barras de 1m cada por cálice, usando FERRO_10
        qntd_barras_ferro = 4
        comprimento_barra_ferro = 1.0
        
        valor_ferro_10_barra = self.get_material_value('FERRO_10')  # 61.76
        valor_ferro_10_metro = valor_ferro_10_barra / 12  # 5.1467 por metro
        
        comprimento_total_ferro = qntd_barras_ferro * comprimento_barra_ferro * numero_pilares
        custo_ferro = comprimento_total_ferro * valor_ferro_10_metro
        
        # ESTRIBOS - EXATO DO EXCEL
        # 8 estribos por cálice, comprimento = 0.7+0.6+0.7+0.6 = 2.6m
        qntd_estribos_por_calice = 8
        comprimento_estribo = 0.7 + 0.6 + 0.7 + 0.6  # 2.6m
        
        valor_ferro_5_barra = self.get_material_value('FERRO_5')  # 15.9
        valor_ferro_5_metro = valor_ferro_5_barra / 12  # 1.325 por metro
        
        comprimento_total_estribos = qntd_estribos_por_calice * comprimento_estribo * numero_pilares
        custo_estribos = comprimento_total_estribos * valor_ferro_5_metro
        
        # MALHA - EXATO DO EXCEL
        # Comprimento da malha = 0.7+0.6+0.7+0.6 = 2.6m por cálice
        comprimento_malha_por_calice = 0.7 + 0.6 + 0.7 + 0.6  # 2.6m
        comprimento_total_malha = comprimento_malha_por_calice * numero_pilares
        
        # TELA_4.2 = 1686 por 120m -> 14.05 por metro
        valor_tela_4_2 = self.get_material_value('TELA_4.2')  # 1686
        custo_malha_por_metro = valor_tela_4_2 / 120  # 14.05 por metro
        custo_malha = comprimento_total_malha * custo_malha_por_metro
        
        # CUSTO TOTAL - SOMANDO CONFORME EXCEL
        custo_total_calices = custo_concreto + custo_ferro + custo_estribos + custo_malha
        
        calices['volume_concreto'] = volume_total_concreto
        calices['comprimento_ferro'] = comprimento_total_ferro
        calices['comprimento_estribos'] = comprimento_total_estribos
        calices['comprimento_malha'] = comprimento_total_malha
        calices['C11'] = custo_total_calices
        
        self.results['CALICES'] = calices
        
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
        """Calcula valores do portão conforme Excel"""
        portao = {}
        entrada = self.results['ENTRADA']
        
        tem_portao = entrada.get('C21', 'NÃO') == 'SIM'
        
        if tem_portao:
            largura = entrada.get('PORTAO_LARGURA', 5)
            altura = entrada.get('PORTAO_ALTURA', 5)
            area_portao = largura * altura
            custo_por_m2 = 350  # Conforme Excel
            quantidade = 2  # Dois portões conforme prática
            
            portao['C11'] = area_portao * custo_por_m2 * quantidade
        else:
            portao['C11'] = 0
            
        self.results['PORTAO'] = portao
        
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
        """Calcula o orçamento final conforme Excel"""
        orcamento = {}
        entrada = self.results['ENTRADA']
        
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
        
        tem_projeto = entrada.get('C18', 'SIM') == 'SIM'
        custo_projeto = 0
        if tem_projeto:
            area = entrada.get('D3', 0)
            custo_projeto = area * 15
        
        custo_total_materiais = custo_cobertura + custo_estrutura + custo_complementos + custo_projeto
        
        frete = self.results['FRETE'].get('valor_final', 0)
        custo_total_com_frete = custo_total_materiais + frete
        
        valor_venda = custo_total_com_frete * 2
        
        com_nota = entrada.get('COM_NOTA', 'NÃO') == 'SIM'
        if com_nota:
            valor_venda = valor_venda * 1.08
        
        orcamento['custo_total_materiais'] = custo_total_materiais
        orcamento['custo_cobertura'] = custo_cobertura
        orcamento['custo_estrutura'] = custo_estrutura
        orcamento['custo_complementos'] = custo_complementos
        orcamento['custo_projeto'] = custo_projeto
        orcamento['frete'] = frete
        orcamento['custo_total_com_frete'] = custo_total_com_frete
        orcamento['valor_venda'] = valor_venda
        orcamento['com_nota'] = com_nota
        
        self.results['ORCAMENTO_FINAL'] = orcamento
        
    def get_result(self, sheet, cell):
        """Obtém um resultado específico"""
        if sheet in self.results:
            return self.results[sheet].get(cell, 0)
        return 0
        
    def get_orcamento_final(self):
        """Retorna o orçamento final"""
        return self.results.get('ORCAMENTO_FINAL', {})
        
    def get_banco_dados(self):
        """Retorna o banco de dados completo"""
        return self.data['BANCO_DE_DADOS']
    
    def get_banco_dados_completo(self):
        """Retorna o banco de dados completo com estrutura de dicionário"""
        return self.data['BANCO_DE_DADOS']

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