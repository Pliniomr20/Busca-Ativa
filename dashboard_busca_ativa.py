import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from pathlib import Path
from io import BytesIO
import numpy as np
import datetime as dt

# Imports para geração de PDF (ReportLab)
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor, black, white
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# --- 1. Camada de Configuração e Constantes ---
class Config:
    def __init__(self):
        self.base_dir = Path(__file__).resolve().parent
        # Nota: Assegure-se de que estes arquivos existam no diretório de execução
        self.logo_path = self.base_dir / "LOGO_M_F_1_-removebg-preview.png"
        self.font_path = self.base_dir / 'ARIALN.TTF'
        self.font_bold_path = self.base_dir / 'ARIALNB.TTF'
        self.excel_file = "BUSCA ATIVA.xlsx"
        self.excel_sheet = "Sheet1"
        self.coluna_colaborador = 'NOME_AGENTE'
        
        # Paleta de cores padronizada e profissional (tons de verde/azul)
        self.palette = {
            "PRIMARY": "#2E7D32",         # Verde Escuro (Ação Principal)
            "ACCENT": "#4CAF50",          # Verde Médio (Ênfase)
            "SECONDARY_ACCENT": "#1B5E20",# Verde Muito Escuro (Texto/Títulos)
            "BACKGROUND_LIGHT": "#F5F5F5",# Fundo Leve
            "TEXT_DEFAULT": "#212529",    # Texto Padrão
            "GREY_LIGHT": "#E0E0E0",      # Borda/Separador
            "GREY_DARK": "#757575",       # Texto Secundário 
            "WHITE": "#FFFFFF",
            "SHADOW_LIGHT": "rgba(0,0,0,0.08)",
            "SUCCESS": "#2E7D32",         # Indicadores Positivos
            "WARNING": "#FFB300",         # Indicadores de Atenção (Amarelo)
            "DANGER": "#C62828",          # Indicadores Críticos (Vermelho)
            "CHART_BAR_1": "#4CAF50",     # Cor para gráficos
            "CHART_BAR_2": "#1B5E20"      # Cor secundária para gráficos
        }
        
        # Metas e Serviços
        self.metas_regionais = {'NORTE': 4764, 'NORDESTE': 2418, 'SUL': 4547}
        self.servicos = {
            'executados': ['CONCLUIDO OK', 'DESCARREGADO COM IMPEDIMENTO', 'DESCARREGADO SEM IMPEDIMENTO', 'IMPROCEDENTE'],
            'em_campo': ['ALVO EM CAMPO'],
            'a_atribuir': ['ALVO NAO ATRIBUIDO'],
            'pendentes': ['ALVO ENVIADO - NAO RECEBIDO'],
            'produtivos': ['CONCLUIDO OK'],
            'improdutivos': ['DESCARREGADO COM IMPEDIMENTO', 'IMPROCEDENTE', 'DESCARREGADO SEM IMPEDIMENTO']
        }

        # Lista de Colaboradores (Mantida a lista original)
        self.colaboradores_list = [
            'ANTONIO SALIM GARCIA', 'AUGUSTO CESAR DE OLIVEIRA', 'CARLOS COSMO ALVES RIBEIRO', 'CARLOS DANIEL CUSTODIO DA SILVA',
            'CARLOS EDUARDO CARDOSO DE ARAUJO', 'CLEBER PEREIRA CARDOSO', 'CLEMILSON RODRIGUES DA TRINDADE', 'CRISTIANO DE JESUS MONTEIRO',
            'DAMIAO PEREIRA DE MENESES', 'DIEGO DIEFLEI ARAUJO DA COSTA', 'DIEGO FRANCISCO PESSOA DE MORAES', 'DORISMAR DUARTE SANTOS',
            'EDNELSON MACEDO TORRES', 'ELVIS DO NASCIMENTO RIBEIRO', 'FERNANDO FERREIRA DE LIMA', 'FILLIPE RODRIGUES DE SOUZA',
            'FLAVIO FERREIRA BORGES', 'FRANCISCO DAS CHAGAS SOUSA', 'GUILHERME SCOTT BASILIO ONOFRE', 'HAYNOANN DOUGLAS DOS SANTOS GOMES SEVERINO',
            'HELLEN CRISTINA VALADARES FERREIRA', 'HENRIQUE VINICIUS JACOB DE PAULO', 'HIGOR VINICIUS DE CASTRO', 'HYGOR MATEUS BATISTA RIBEIRO DA SILVA',
            'IGOR SILVA SANTOS', 'IRISLAN SANTINNI TORRES DE SOUSA', 'IDAMAR VIEIRA DE OLIVEIRA FILHO', 'JEFFERSON PEREIRA DE MAGALHAES',
            'JOAO NETO ROCHA DA SILVA', 'JONATAN RODRIGO BATISTA FELIX', 'JUAN SOUZA AMARAL', 'KEVEN LUIZ SOUSA DE FREITAS',
            'KLEBER FERNANDES DE AZEVEDO', 'LAZARO BRAZ DE SOUSA', 'MARCELO MENDES RAMOS', 'MARCIO WAGNER JOSE LOPES SANCHES',
            'MATEUS LIMA MENDONCA', 'MATHEUS HENRIQUE DOS SANTOS SILVA', 'MAURICIO JOSE PEREIRA VAZ', 'MAYCON EDUARDO FIGUEREDO',
            'NELSON NERES SOARES', 'ODEILDO DA COSTA SANTANA', 'OTAVIO RODRIGUES OLIMPIO', 'PABLO NUNES DOS PRAZERES',
            'ODAIR ALVES DOS SANTOS', 'PAULO VINICIOS HABERMANN DA ROCHA PINTO', 'RAFAEL DUARTE MARQUES', 'PEDRO HENRIQUE DA CRUZ',
            'RICARDO DE AMORIM CARNEIRO', 'RONILSON DAS CHAGAS OLIVEIRA', 'TIAGO LUCIO FERNANDES SOUSA', 'WANDERSON MENDES DE MOURA',
            'VALDEMAR DE ALMEIDA FILHO', 'WANDERSON MORAES SOEIRO', 'WENDER SOARES DA SILVA', 'JOAO CLEVISTON DANTAS',
            'WEVERSON DA SILVA', 'JEFFERSON DOUGLAS DE SOUSA MAIA', 'BRUNO ALVES FERREIRA', 'KEVERSON ANTONIO DE SOUZA SIQUEIRA',
            'WENDER DE CASTRO VIEIRA', 'ALAN ALVES AURELIANO', 'BRENNO PEREIRA CAMPOS DE OLIVEIRA', 'BRUNO HENRIQUE DE MARINS CABRAL',
            'CLEITON ARAUJO DE OLIVEIRA', 'DAITON DIEGO DA SILVA ROMEIRO', 'DHYOGO VIEIRA DE MOURA', 'DIEGO GADELHA DE LIMA',
            'DOUGLAS KAIQUE DOS SANTOS REIS', 'HENRIQUE BARBOSA NUNES', 'JOAO VITOR VIEIRA DOS SANTOS', 'LUCIANO SANTOS DE MIRANDA',
            'MARCIO PAULO SILVA', 'MARK ETIENNE RODRIGUES DA COSTA', 'MATHEUS DE JESUS SILVA', 'MURILLO GABRIEL DA SILVA LOBO',
            'MURILO MATHEUS BORGES RODRIGUES', 'NATANIEL VIANA DA SILVA', 'RONAN DA PENHA DE MORAIS', 'RICARDO DA SILVA PEREIRA',
            'SAMUEL ALVES DIAS', 'SANDRO SANTOS ARAUJO', 'WELBESON RODRIGUES DA COSTA', 'VALMIR LOURENCO BORGES',
            'BRENNER OLIVEIRA DE MELO', 'ADRIANO RIBEIRO SANTOS', 'ALEX SILVA OLIVEIRA', 'CAIO GUSTAVO DANTAS SILVA',
            'DJALMA MACIEL MARTINS', 'EDIVALDO MOURA DE OLIVEIRA', 'FLAVIO DOURADO DE SOUZA', 'HYGOR DOS SANTOS SOUSA',
            'JONATHAN LIMA DA ROCHA MACHADO', 'JOSE DOURADO DE OLIVEIRA FILHO', 'JOSE WILLAME DA SILVA MOTA', 'MARLLON BRUNNO ALEM ALVES',
            'PEDRO HENRIQUE CIRINO DE MELO', 'ROSIMAR PEREIRA LEITE', 'JOENDERSON DE JESUS AVELINO', 'MARCOS ANTONIO RODRIGUES DA SILVA',
            'RODRIGGO WAGNER CAMPOS DA SILVA', 'ALVARO DA SILVA ROCHA', 'DANILO MIGUEL DE OLIVEIRA', 'DIEGO FONSECA DOS SANTOS',
            'MATEUS DIAS DOS SANTOS', 'DANIEL LUIZ CORREIA PANTA', 'RAFAEL PEREIRA DE OLIVEIRA', 'RONIS MARCIO CANDIDO FERREIRA',
            'CLAUDINEIA MIRANDA SOUZA', 'LHORRAN FHILLYPHE TAVARES NOGUEIRA', 'KLEVER PEREIRA DOS SANTOS', 'INGRITH LORENA PEREIRA DE OLIVEIRA',
            'ADNEY HENRIQUE NOGUEIRA LOPES', 'CARLOS HENRIQUE GONÇALVES MELO'
        ]

config = Config()

# --- Funções de Formatação e Utilitários ---
def formatar_inteiro(valor: float | int) -> str:
    if pd.isna(valor) or valor is None: return "0"
    try: valor = int(valor)
    except (ValueError, TypeError): return "Inválido"
    # Formatação com ponto como separador de milhar
    return f"{valor:,}".replace(",", "TEMP").replace(".", ",").replace("TEMP", ".")

def get_status_kpi_color(kpi_value, threshold, inverse=False):
    """Retorna a cor da paleta baseado no valor e no threshold."""
    if inverse:
        return config.palette["DANGER"] if kpi_value > threshold else config.palette["SUCCESS"]
    else:
        return config.palette["SUCCESS"] if kpi_value >= threshold else config.palette["DANGER"]

# --- 2. Camada de Acesso e Processamento de Dados ---
@st.cache_data(ttl=3600, show_spinner="Carregando e processando dados de Busca Ativa...")
def carregar_e_processar_dados(caminho_arquivo: Path) -> pd.DataFrame:
    if not caminho_arquivo.exists():
        st.error(f"Erro Crítico: Arquivo de dados não encontrado em '{caminho_arquivo}'.")
        st.stop()
    
    try:
        df = pd.read_excel(caminho_arquivo, sheet_name=config.excel_sheet)
        
        # Limpeza e Padronização de Colunas
        df.columns = df.columns.str.strip().str.upper().str.replace(' ', '_').str.replace('[^A-Z0-9_]', '', regex=True)
        
        required_cols = ['REGIONAL', 'MUNICIPIO', 'NOME_FASE', 'ALVO_CONDICAO_OBJETIVA', config.coluna_colaborador, 'DATA_DEVOLUCAO']
        if not all(col in df.columns for col in required_cols):
            missing_cols = [col for col in required_cols if col not in df.columns]
            st.error(f"Colunas essenciais faltando na planilha: {missing_cols}")
            st.stop()
            
        if df.empty:
            st.warning("Nenhum dado foi encontrado. Verifique a planilha.")
            st.stop()
            
        # Limpeza e Padronização de Dados
        cols_to_upper = ['NOME_FASE', 'REGIONAL', 'MUNICIPIO', config.coluna_colaborador]
        for col in cols_to_upper:
            df[col] = df[col].astype(str).str.upper().str.strip()
            
        df['DATA_DEVOLUCAO'] = pd.to_datetime(df['DATA_DEVOLUCAO'], errors='coerce', dayfirst=True)
        
        # COLUNA AUXILIAR PARA FILTRAGEM: OBTÉM APENAS A PARTE DA DATA
        df['DATA_FILTRO_AUX'] = df['DATA_DEVOLUCAO'].dt.date
        
        # Filtragem inicial apenas para regionais válidas
        regionais_validas = list(config.metas_regionais.keys())
        df = df[df['REGIONAL'].isin(regionais_validas)].copy()

        return df
    
    except Exception as e:
        st.error(f"Erro fatal ao carregar o arquivo Excel: {e}")
        st.exception(e)
        st.stop()

# Carregamento da base de dados
try:
    df_principal = carregar_e_processar_dados(config.base_dir / config.excel_file)
except:
    # Caso o arquivo não exista ou dê erro de I/O, cria um DF vazio para o diagnóstico
    df_principal = pd.DataFrame(columns=['REGIONAL', 'MUNICIPIO', 'NOME_FASE', 'ALVO_CONDICAO_OBJETIVA', config.coluna_colaborador, 'DATA_DEVOLUCAO', 'DATA_FILTRO_AUX'])


# --- 3. Camada de Lógica de Negócio e Agregação ---
def calcular_indicadores_totais(df_base_total: pd.DataFrame, df_para_analise: pd.DataFrame, colaboradores_list: list) -> dict:
    if df_base_total.empty:
        return {
            "total": 0, "executados_totais": 0, "executados_produtivos": 0, "executados_improdutivos": 0,
            "em_campo": 0, "a_atribuir": 0, "pendentes": 0, "colaboradores_nao_encontrados": [],
            "executados_mf_produtivos": 0
        }
    
    # KPIs da Base Geral
    qtd_executados_totais = df_base_total['NOME_FASE'].isin(config.servicos['executados']).sum()
    qtd_executados_produtivos = df_base_total['NOME_FASE'].isin(config.servicos['produtivos']).sum()
    qtd_executados_improdutivos = df_base_total['NOME_FASE'].isin(config.servicos['improdutivos']).sum()
    qtd_em_campo = df_base_total['NOME_FASE'].isin(config.servicos['em_campo']).sum()
    qtd_pendentes = df_base_total['NOME_FASE'].isin(config.servicos['pendentes']).sum()

    # KPI de "A Atribuir" (considerando ALVO_CONDICAO_OBJETIVA = 'SIM')
    df_filtrado_sim = df_base_total[df_base_total['ALVO_CONDICAO_OBJETIVA'].str.upper().str.strip() == 'SIM'].copy()
    qtd_a_atribuir_base = df_filtrado_sim['NOME_FASE'].isin(config.servicos['a_atribuir']).sum()
    qtd_total_servicos_base = len(df_filtrado_sim)
    
    # KPI Específico para Colaboradores do Gestor (MF)
    qtd_executados_mf_produtivos = df_para_analise['NOME_FASE'].isin(config.servicos['produtivos']).sum()

    colaboradores_na_base = set(df_base_total[config.coluna_colaborador].unique())
    colaboradores_nao_encontrados = [c for c in colaboradores_list if c.upper().strip() not in colaboradores_na_base]
    
    return {
        "total": qtd_total_servicos_base,
        "executados_totais": qtd_executados_totais,
        "executados_produtivos": qtd_executados_produtivos,
        "executados_improdutivos": qtd_executados_improdutivos,
        "em_campo": qtd_em_campo,
        "a_atribuir": qtd_a_atribuir_base,
        "pendentes": qtd_pendentes,
        "colaboradores_nao_encontrados": colaboradores_nao_encontrados,
        "executados_mf_produtivos": qtd_executados_mf_produtivos
    }

def calcular_metas_por_regional(df_base: pd.DataFrame, metas_regionais: dict, selecao_regional: list) -> list:
    """Calcula a performance em relação à meta (apenas CONCLUIDO OK) para as regionais selecionadas."""
    metas_kpis = []
    
    # Filtrar a base apenas para os dados de meta (base total, não a filtrada por data/município no app)
    df_base_filtrada = df_base[df_base['REGIONAL'].isin(selecao_regional)].copy()
    
    for regional in sorted(selecao_regional):
        df_regional = df_base_filtrada[df_base_filtrada['REGIONAL'] == regional]
        meta = metas_regionais.get(regional, 0)
        
        # Considera apenas 'CONCLUIDO OK' (Produtivo) para a meta
        executados_para_meta = df_regional['NOME_FASE'].str.upper().str.strip().eq('CONCLUIDO OK').sum()
        
        restante = meta - executados_para_meta
        percentual = (executados_para_meta / meta) * 100 if meta > 0 else 0
        
        metas_kpis.append({
            'regional': regional,
            'meta': meta,
            'executados': executados_para_meta,
            'restante': restante,
            'percentual': percentual
        })
    return metas_kpis

def agregar_por_dimensao(df: pd.DataFrame, coluna_agregacao: str, servico_type: str) -> pd.DataFrame:
    """Agrega dados por dimensão (Regional/Município) e tipo de serviço."""
    if df.empty or coluna_agregacao not in df.columns:
        return pd.DataFrame(columns=['Dimensão', 'Métrica'])
    
    is_a_atribuir = servico_type == 'a_atribuir'
    
    # Tratamento especial para 'A Atribuir' que depende de 'ALVO_CONDICAO_OBJETIVA'
    df_temp = df[df['ALVO_CONDICAO_OBJETIVA'].str.upper().str.strip() == 'SIM'].copy() if is_a_atribuir else df.copy()

    if servico_type in config.servicos:
        df_agregado = df_temp.groupby(coluna_agregacao)['NOME_FASE'].apply(
            lambda x: x.isin(config.servicos[servico_type]).sum()
        ).reset_index()
    else:
        # Fallback para contagem total se o tipo de serviço não for especificado
        df_agregado = df_temp.groupby(coluna_agregacao)['NOME_FASE'].count().reset_index()
        
    df_agregado.columns = ['Dimensão', 'Métrica']
    return df_agregado.sort_values(by='Métrica', ascending=False)

def agregar_desempenho_colaborador(df: pd.DataFrame, colaboradores_list: list) -> pd.DataFrame:
    """Calcula o desempenho individual dos colaboradores na lista."""
    colaboradores_upper = [c.upper().strip() for c in colaboradores_list]
    df_filtrado = df[df[config.coluna_colaborador].isin(colaboradores_upper)].copy()
    
    if df_filtrado.empty:
        return pd.DataFrame()
        
    df_agregado = df_filtrado.groupby(config.coluna_colaborador).agg(
        Qtd_Executados=('NOME_FASE', lambda x: x.isin(config.servicos['executados']).sum()),
        Qtd_Produtivos=('NOME_FASE', lambda x: x.isin(config.servicos['produtivos']).sum()),
        Qtd_Improdutivos=('NOME_FASE', lambda x: x.isin(config.servicos['improdutivos']).sum()),
        Qtd_Em_Campo=('NOME_FASE', lambda x: x.isin(config.servicos['em_campo']).sum()),
        Qtd_Pendentes=('NOME_FASE', lambda x: x.isin(config.servicos['pendentes']).sum()),
        Qtd_Alocados=('NOME_FASE', 'count'),
    ).reset_index()

    # Calcular taxa de produtividade
    df_agregado['Taxa_Produtividade'] = np.where(
        df_agregado['Qtd_Executados'] > 0, 
        df_agregado['Qtd_Produtivos'] / df_agregado['Qtd_Executados'], 
        0
    )
    
    return df_agregado.sort_values(by='Qtd_Produtivos', ascending=False)


# --- Funções de Visualização ---
def plot_bar_chart_v2(df_data, x_col, y_col, title, x_label, y_label, color_col=None, color_discrete_map=None, color_discrete_sequence=None, orientation='h', height=None):
    """Função centralizada para criar gráficos de barra consistentes e profissionais."""
    if df_data.empty:
        fig = px.bar(title=f"<b>{title}</b>")
        fig.update_layout(
            xaxis_title_text=x_label, yaxis_title_text=y_label,
            xaxis_visible=False, yaxis_visible=False,
            annotations=[dict(text="Nenhum dado disponível.", xref="paper", yref="paper", showarrow=False, font_size=16, font_color=config.palette['GREY_DARK'])]
        )
        return fig
    
    # Determinar eixos
    x_plot = y_col if orientation == 'h' else x_col
    y_plot = x_col if orientation == 'h' else y_col
    
    # Definir altura dinâmica para gráficos horizontais
    if height is None and orientation == 'h':
        height = min(750, len(df_data[x_col].unique()) * 30 + 150)
    elif height is None:
        height = 550

    # Usando o argumento color_discrete_sequence se ele foi passado
    color_seq = color_discrete_sequence if color_discrete_sequence else [config.palette["CHART_BAR_1"]]

    fig = px.bar(
        df_data,
        x=x_plot,
        y=y_plot,
        color=color_col,
        title=f"<b>{title}</b>",
        labels={x_col: x_label, y_col: y_col},
        color_discrete_map=color_discrete_map,
        color_discrete_sequence=color_seq,
        template='plotly_white',
        orientation=orientation,
        height=height
    )

    # Configurações de Traços (Barras)
    fig.update_traces(
        text=df_data[x_plot].apply(formatar_inteiro) if orientation == 'h' else df_data[y_plot].apply(formatar_inteiro),
        texttemplate='%{x}' if orientation == 'h' else '%{y}',
        textposition='outside',
        hovertemplate="<b>%{y}</b><br>Quantidade: %{x}<extra></extra>" if orientation == 'h' else "<b>%{x}</b><br>Quantidade: %{y}<extra></extra>",
        marker_line_width=1,
        marker_line_color=config.palette['WHITE']
    )
    
    # Definindo a ordenação e a grade de forma limpa
    layout_updates = {
        'xaxis_title_font': dict(size=13, color=config.palette['SECONDARY_ACCENT']),
        'yaxis_title_font': dict(size=13, color=config.palette['SECONDARY_ACCENT']),
        'title_font': dict(size=18, color=config.palette['SECONDARY_ACCENT']),
        'font_color': config.palette['TEXT_DEFAULT'],
        'showlegend': True if color_col else False,
        'legend_title_text': color_col,
        'margin': dict(l=20, r=60, t=50, b=20),
        'title_x': 0.05,
        'title_y': 0.95,
        'plot_bgcolor': config.palette['WHITE'],
        'paper_bgcolor': config.palette['WHITE'],
    }
    
    # Adicionando a configuração de eixos de forma condicional
    if orientation == 'h':
        layout_updates['yaxis'] = {'categoryorder': 'total ascending', 'showgrid': False}
        layout_updates['xaxis'] = {'categoryorder': 'total descending', 'showgrid': True, 'gridcolor': config.palette['GREY_LIGHT']}
    else:
        layout_updates['xaxis'] = {'categoryorder': 'total descending', 'showgrid': True, 'gridcolor': config.palette['GREY_LIGHT']}
        layout_updates['yaxis'] = {'categoryorder': 'total ascending', 'showgrid': True, 'gridcolor': config.palette['GREY_LIGHT']}


    fig.update_layout(**layout_updates)
    
    return fig

# Função para exibir um KPI com estilização aprimorada
def display_kpi_metric(label, value, icon, sub_label=None, sub_value=None):
    """Exibe um KPI profissional com ícone e sub-detalhes."""
    
    sub_detail_html = ""
    if sub_label and sub_value is not None:
        # Simplificando a estrutura do sub-detalhe para evitar erros de injeção de HTML
        sub_detail_html = f"""
            <div style="margin-top: 15px; border-top: 1px solid {config.palette['GREY_LIGHT']}; padding-top: 10px;">
                <div style="display: flex; justify-content: space-between; font-size: 14px; color: {config.palette['GREY_DARK']};">
                    <span>{sub_label}</span>
                    <span style="font-weight: bold; color: {config.palette['TEXT_DEFAULT']};">{sub_value}</span>
                </div>
            </div>
        """

    st.markdown(f"""
        <div data-testid="stMetric" style="border: 1px solid {config.palette['GREY_LIGHT']};">
            <div data-testid="stMetricLabel" style="display: flex; align-items: center; justify-content: center; gap: 8px;">
                <span style="font-size: 1.2em; color: {config.palette['PRIMARY']};">{icon}</span>
                <div style="font-size: 1.1em; font-weight: 500; color: {config.palette['GREY_DARK']}; text-align: center;">{label}</div>
            </div>
            <div data-testid="stMetricValue" style="font-size: 2.2em; font-weight: 800; color: {config.palette['PRIMARY']};">
                {formatar_inteiro(value)}
            </div>
            {sub_detail_html}
        </div>
    """, unsafe_allow_html=True)

def display_regional_goals(kpi_regional):
    """Exibe o KPI de meta regional com cores condicionais."""
    col_regional, col_meta, col_executado, col_restante = st.columns(4)
    
    with col_regional:
        st.markdown(f"""
            <div style="padding: 15px; border: 1px solid {config.palette['GREY_LIGHT']}; border-radius: 8px; text-align: center; height: 100%; display: flex; flex-direction: column; justify-content: center;">
                <p style="margin: 0; font-size: 1.1em; font-weight: 700; color: {config.palette['SECONDARY_ACCENT']};">🎯 Regional {kpi_regional['regional']}</p>
                <p style="margin: 5px 0 0 0; font-size: 1.8em; font-weight: 800; color: {config.palette['PRIMARY']};">{kpi_regional['percentual']:.1f}%</p>
            </div>
        """, unsafe_allow_html=True)
    
    with col_meta:
        display_kpi_metric("Meta", kpi_regional['meta'], "📈")
    
    with col_executado:
        display_kpi_metric("Executados (OK)", kpi_regional['executados'], "✅")
    
    with col_restante:
        restante_valor = kpi_regional['restante']
        
        if restante_valor <= 0:
            cor_kpi = config.palette["SUCCESS"]
            texto_kpi = "Meta Atingida!"
            label_kpi = "Status"
        else:
            cor_kpi = config.palette["DANGER"]
            texto_kpi = formatar_inteiro(restante_valor)
            label_kpi = "Restante"
        
        st.markdown(f"""
            <div data-testid="stMetric" style="border: 2px solid {cor_kpi}; background-color: {config.palette['WHITE']};">
                <div data-testid="stMetricLabel" style="color: {cor_kpi};">
                    <div style="font-weight: 700; font-size: 1.1em; color: {cor_kpi};">{label_kpi}</div>
                </div>
                <div data-testid="stMetricValue" style="color: {cor_kpi}; font-size: 2.2em; font-weight: bold;">
                    {texto_kpi}
                </div>
            </div>
        """, unsafe_allow_html=True)

# --- 4. Camada de Geração de Relatórios (PDF) ---
class RelatorioVisualPDF:
    def __init__(self, logo_path: Path, palette: dict, output_buffer: BytesIO):
        self.logo_path = logo_path
        self.palette = palette
        self.buffer = output_buffer
        self.doc = SimpleDocTemplate(self.buffer, pagesize=A4,
                                     leftMargin=0.75*inch, rightMargin=0.75*inch,
                                     topMargin=1.0*inch, bottomMargin=0.75*inch)
        self.story = []
        self._register_fonts()
        self._define_styles()

    def _register_fonts(self):
        try:
            pdfmetrics.registerFont(TTFont('Arial', str(config.font_path)))
            pdfmetrics.registerFont(TTFont('Arial-Bold', str(config.font_bold_path)))
            pdfmetrics.registerFontFamily('Arial', normal='Arial', bold='Arial-Bold')
        except Exception:
            pdfmetrics.registerFontFamily('Helvetica', normal='Helvetica', bold='Helvetica-Bold')

    def _define_styles(self):
        styles = getSampleStyleSheet()
        self.styles = {
            'h1': ParagraphStyle('h1', parent=styles['h1'], fontName='Arial-Bold', fontSize=18, leading=22, alignment=TA_LEFT, textColor=HexColor(self.palette["PRIMARY"]), spaceAfter=18),
            'h2': ParagraphStyle('h2', parent=styles['h2'], fontName='Arial-Bold', fontSize=14, leading=16, textColor=HexColor(self.palette["SECONDARY_ACCENT"]), spaceAfter=8),
            'body': ParagraphStyle('body', parent=styles['Normal'], fontName='Arial', fontSize=11, leading=14, textColor=HexColor(self.palette["TEXT_DEFAULT"]), spaceAfter=8),
            'kpi_card_label': ParagraphStyle('kpi_card_label', parent=styles['Normal'], fontName='Arial-Bold', fontSize=12, leading=14, alignment=TA_CENTER, textColor=HexColor(self.palette["GREY_DARK"])),
            'kpi_card_value': ParagraphStyle('kpi_card_value', parent=styles['Normal'], fontName='Arial-Bold', fontSize=18, leading=20, alignment=TA_CENTER, textColor=HexColor(self.palette["PRIMARY"])),
            'table_header': ParagraphStyle('table_header', parent=styles['Normal'], fontName='Arial-Bold', fontSize=9, leading=11, alignment=TA_CENTER, textColor=HexColor(self.palette["WHITE"])),
            'table_body': ParagraphStyle('table_body', parent=styles['Normal'], fontName='Arial', fontSize=9, leading=11, alignment=TA_LEFT, textColor=HexColor(self.palette["TEXT_DEFAULT"])),
            'table_body_center': ParagraphStyle('table_body_center', parent=styles['Normal'], fontName='Arial', fontSize=9, leading=11, alignment=TA_CENTER, textColor=HexColor(self.palette["TEXT_DEFAULT"])),
            'footer': ParagraphStyle('footer', parent=styles['Normal'], fontName='Arial-Italic', fontSize=8, leading=10, alignment=TA_CENTER, textColor=HexColor(self.palette["GREY_DARK"])), 
        }
    
    def _header_page(self, canvas, doc):
        canvas.saveState()
        # Título do Relatório (Centralizado)
        canvas.setFont('Arial-Bold', 16)
        canvas.setFillColor(HexColor(self.palette["PRIMARY"]))
        canvas.drawCentredString(A4[0]/2.0, A4[1] - 0.5*inch, "Relatório de Performance Busca Ativa")
        
        # Linha Divisória
        canvas.setStrokeColor(HexColor(self.palette["ACCENT"]))
        canvas.setLineWidth(2)
        canvas.line(0.75*inch, A4[1] - 0.75*inch, A4[0] - 0.75*inch, A4[1] - 0.75*inch)
        
        # Logo (Se existir)
        if self.logo_path.exists():
            logo_width = 0.5 * inch
            logo_height = 0.5 * inch
            try:
                canvas.drawImage(
                    str(self.logo_path),
                    0.75*inch,
                    A4[1] - 0.7*inch,
                    width=logo_width,
                    height=logo_height,
                    mask='auto'
                )
            except Exception:
                pass
        
        canvas.restoreState()

    def _footer_page(self, canvas, doc):
        canvas.saveState()
        canvas.setFont('Arial-Bold', 8)
        canvas.setFillColor(HexColor(self.palette["GREY_DARK"])) 
        canvas.drawCentredString(A4[0]/2.0, 0.5*inch, f"Página {doc.page} de {doc.pages}")
        canvas.restoreState()

    def add_kpi_summary(self, kpis: dict):
        self.story.append(Paragraph("Resumo da Base de Dados", self.styles['h2']))
        
        kpi_data = [
            [
                Paragraph("Total de Alvos:", self.styles['kpi_card_label']),
                Paragraph("Serviços Executados:", self.styles['kpi_card_label']),
                Paragraph("Serviços Produtivos:", self.styles['kpi_card_label']),
                Paragraph("Serviços Improdutivos:", self.styles['kpi_card_label'])
            ],
            [
                Paragraph(formatar_inteiro(kpis['total']), self.styles['kpi_card_value']),
                Paragraph(formatar_inteiro(kpis['executados_totais']), self.styles['kpi_card_value']),
                Paragraph(formatar_inteiro(kpis['executados_produtivos']), self.styles['kpi_card_value']),
                Paragraph(formatar_inteiro(kpis['executados_improdutivos']), self.styles['kpi_card_value'])
            ]
        ]
        
        kpi_data_row2 = [
            [
                Paragraph("Serviços em Campo:", self.styles['kpi_card_label']),
                Paragraph("Serviços a Atribuir:", self.styles['kpi_card_label']),
                Paragraph("Serviços Pendentes:", self.styles['kpi_card_label']),
                Paragraph("Produtivos (MF):", self.styles['kpi_card_label'])
            ],
            [
                Paragraph(formatar_inteiro(kpis['em_campo']), self.styles['kpi_card_value']),
                Paragraph(formatar_inteiro(kpis['a_atribuir']), self.styles['kpi_card_value']),
                Paragraph(formatar_inteiro(kpis['pendentes']), self.styles['kpi_card_value']),
                Paragraph(formatar_inteiro(kpis['executados_mf_produtivos']), self.styles['kpi_card_value'])
            ]
        ]
        
        table_style = TableStyle([
            ('GRID', (0,0), (-1,-1), 1, HexColor(self.palette["GREY_LIGHT"])),
            ('BACKGROUND', (0,0), (-1,0), HexColor(self.palette["BACKGROUND_LIGHT"])),
            ('BACKGROUND', (0,1), (-1,1), HexColor(self.palette["WHITE"])),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (-1,0), 10),
            ('TOPPADDING', (0,1), (-1,1), 10),
            ('BOTTOMPADDING', (0,1), (-1,1), 10),
        ])
        
        kpi_table = Table(kpi_data, colWidths=[self.doc.width / 4.0] * 4)
        kpi_table.setStyle(table_style)
        self.story.append(kpi_table)
        self.story.append(Spacer(1, 0.2 * inch))
        
        kpi_table_2 = Table(kpi_data_row2, colWidths=[self.doc.width / 4.0] * 4)
        kpi_table_2.setStyle(table_style)
        self.story.append(kpi_table_2)
        
        self.story.append(Spacer(1, 0.1 * inch))

    def add_dataframe_to_pdf(self, title: str, df: pd.DataFrame):
        self.story.append(Paragraph(title, self.styles['h2']))
        
        if df.empty:
            self.story.append(Paragraph("Nenhum dado disponível para esta tabela.", self.styles['body']))
            self.story.append(Spacer(1, 0.1 * inch))
            return

        # Adicionar formatação de porcentagem se a coluna existir
        df_display = df.copy()
        if 'Taxa_Produtividade' in df_display.columns:
            df_display['Taxa_Produtividade'] = df_display['Taxa_Produtividade'].apply(lambda x: f"{x:.1%}")
            
        headers = [Paragraph(col.replace('_', ' ').replace('Qtd', 'Qtd.'), self.styles['table_header']) for col in df_display.columns]
        data = [headers]
        for _, row in df_display.iterrows():
            row_data = []
            for col, item in row.items():
                style = self.styles['table_body_center']
                if col == config.coluna_colaborador:
                    style = self.styles['table_body']
                elif 'Qtd_' in col:
                    item = formatar_inteiro(item)
                
                row_data.append(Paragraph(str(item), style))
            data.append(row_data)

        num_cols = len(df_display.columns)
        available_width = self.doc.width
        col_widths = [available_width / num_cols] * num_cols
        
        table = Table(data, colWidths=col_widths)
        table_style_list = [
            ('BACKGROUND', (0,0), (-1,0), HexColor(self.palette["PRIMARY"])),
            ('TEXTCOLOR', (0,0), (-1,0), HexColor(self.palette["WHITE"])),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, HexColor(self.palette["GREY_LIGHT"])),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
            ('TOPPADDING', (0,0), (-1,-1), 6),
        ]
        for i in range(1, len(data)):
            bg_color = HexColor(self.palette["BACKGROUND_LIGHT"]) if i % 2 == 0 else HexColor(self.palette["WHITE"])
            table_style_list.append(('BACKGROUND', (0,i), (-1,i), bg_color))
        
        table.setStyle(TableStyle(table_style_list))
        self.story.append(table)
        self.story.append(Spacer(1, 0.15 * inch))

    def generate_report(self, df_base_total: pd.DataFrame, df_para_analise: pd.DataFrame, df_colab_performance: pd.DataFrame, colaboradores_nao_encontrados: list, selecao_regional: list):
        kpis_gerais = calcular_indicadores_totais(df_base_total, df_para_analise, config.colaboradores_list)
        metas_kpis = calcular_metas_por_regional(df_principal, config.metas_regionais, selecao_regional)
        
        self.story.append(Paragraph("Resumo de Performance Geral", self.styles['h1']))
        self.add_kpi_summary(kpis_gerais)

        if metas_kpis:
            self.story.append(Spacer(1, 0.25 * inch))
            self.story.append(Paragraph("Acompanhamento de Metas por Regional", self.styles['h2']))
            meta_data = [['Regional', 'Meta', 'Executados Produtivos', 'Restante', 'Percentual']]
            for item in metas_kpis:
                status_cor = HexColor(self.palette["SUCCESS"]) if item['restante'] <= 0 else HexColor(self.palette["DANGER"])
                meta_data.append([
                    Paragraph(item['regional'], self.styles['body']),
                    Paragraph(formatar_inteiro(item['meta']), self.styles['body']),
                    Paragraph(formatar_inteiro(item['executados']), self.styles['body']),
                    Paragraph(formatar_inteiro(item['restante']), ParagraphStyle('restante', parent=self.styles['body'], textColor=status_cor)),
                    Paragraph(f"{item['percentual']:.1f}%", self.styles['body'])
                ])
            
            meta_table = Table(meta_data, colWidths=[self.doc.width / 5.0] * 5)
            meta_table.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), HexColor(self.palette["PRIMARY"])),
                ('TEXTCOLOR', (0,0), (-1,0), HexColor(self.palette["WHITE"])),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('GRID', (0,0), (-1,-1), 0.5, HexColor(self.palette["GREY_LIGHT"])),
                ('BOTTOMPADDING', (0,0), (-1,0), 10),
                ('TOPPADDING', (0,1), (-1,-1), 10),
                ('BOTTOMPADDING', (0,1), (-1,-1), 10),
            ]))
            self.story.append(meta_table)
            self.story.append(Spacer(1, 0.15 * inch))
        
        self.story.append(PageBreak())
        
        self.story.append(Paragraph("Análise de Serviços por Região", self.styles['h1']))
        
        df_prod_reg = agregar_por_dimensao(df_base_total, 'REGIONAL', 'produtivos').rename(columns={'Métrica': 'Produtivos'})
        df_improd_reg = agregar_por_dimensao(df_base_total, 'REGIONAL', 'improdutivos').rename(columns={'Métrica': 'Improdutivos'})
        df_total_reg = agregar_por_dimensao(df_base_total, 'REGIONAL', 'executados').rename(columns={'Métrica': 'Total'})

        df_analise_regional = df_prod_reg.merge(df_improd_reg, on='Dimensão', how='outer').merge(df_total_reg, on='Dimensão', how='outer').fillna(0)
        df_analise_regional = df_analise_regional.sort_values(by='Total', ascending=False).rename(columns={'Dimensão': 'Regional'})
        self.add_dataframe_to_pdf("Serviços Executados por Regional", df_analise_regional)
        
        self.story.append(Paragraph("Desempenho dos seus Colaboradores", self.styles['h1']))
        
        if colaboradores_nao_encontrados:
            self.story.append(Paragraph(f"<b>Atenção:</b> Os seguintes colaboradores não foram encontrados na base de dados: {', '.join(colaboradores_nao_encontrados)}", self.styles['body']))
            self.story.append(Spacer(1, 0.1 * inch))
        
        df_colab_performance_sorted = df_colab_performance.sort_values(by='Qtd_Produtivos', ascending=False)
        self.add_dataframe_to_pdf("Tabela de Desempenho Individual Completa", df_colab_performance_sorted.rename(columns={config.coluna_colaborador: 'Colaborador'}))

        try:
            self.doc.build(self.story, onFirstPage=self._header_page, onLaterPages=self._header_page, canvasmaker=lambda *args, **kwargs: SimpleDocTemplate.Canvas(*args, **kwargs, pagesize=A4))
            return self.buffer.getvalue()
        except Exception as e:
            st.error(f"Erro ao construir o PDF: {e}")
            return None


# --- 5. Lógica de UI (Camada de Apresentação Principal) ---
st.markdown(f"""
    <style>
    /* Estilos Globais Aprimorados */
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700;900&display=swap');
    :root {{
        --primary-color: {config.palette["PRIMARY"]};
        --accent-color: {config.palette["ACCENT"]};
        --secondary-accent-color: {config.palette["SECONDARY_ACCENT"]};
        --bg-light-color: {config.palette["BACKGROUND_LIGHT"]};
        --text-default-color: {config.palette["TEXT_DEFAULT"]};
        --grey-light-color: {config.palette["GREY_LIGHT"]};
        --grey-dark-color: {config.palette["GREY_DARK"]}; 
        --white-color: {config.palette["WHITE"]};
        --shadow-light-color: {config.palette["SHADOW_LIGHT"]};
        --success-color: {config.palette["SUCCESS"]};
        --warning-color: {config.palette["WARNING"]};
        --danger-color: {config.palette["DANGER"]};
    }}
    html, body, .stApp {{ 
        background-color: var(--bg-light-color); 
        color: var(--text-default-color); 
        font-family: 'Roboto', sans-serif;
    }}
    /* Cabeçalho Profissional */
    .main-title-container {{ 
        display: flex; 
        align-items: center; 
        gap: 15px; 
        margin-bottom: 30px; 
        padding: 20px 30px;
        background-color: var(--white-color);
        border-radius: 15px;
        box-shadow: 0 5px 20px var(--shadow-light-color);
    }}
    .main-title-container h1 {{ 
        margin: 0; 
        line-height: 1.2; 
        font-size: 2.5em; 
        font-weight: 900; 
        color: var(--secondary-accent-color);
    }}
    /* Contêiner de Métricas (KPIs) */
    [data-testid="stMetric"] {{ 
        background-color: var(--white-color); 
        border-radius: 12px; 
        padding: 20px 25px; 
        box-shadow: 0 4px 10px var(--shadow-light-color); 
        text-align: center; 
        border: 1px solid var(--grey-light-color); 
        transition: transform 0.2s ease-in-out, box-shadow 0.2s ease-in-out; 
        margin-bottom: 15px; 
        height: 100%; /* Garantir que colunas com sub-métricas fiquem alinhadas */
    }}
    [data-testid="stMetric"]:hover {{ 
        transform: translateY(-3px); 
        box-shadow: 0 6px 15px var(--shadow-light-color); 
    }}
    [data-testid="stMetricValue"] {{ 
        font-size: 2.2em; 
        font-weight: 800; /* Mais negrito */
        color: var(--primary-color);
        word-wrap: break-word; 
        overflow-wrap: break-word; 
        white-space: normal; 
        margin-top: 8px; 
    }}
    [data-testid="stMetricLabel"] > div {{ 
        color: var(--grey-dark-color); 
        font-size: 1.1em; 
        font-weight: 500; 
    }}
    h3 {{ 
        color: var(--secondary-accent-color); 
        font-size: 1.8em; 
        font-weight: 700; 
        margin-top: 35px; 
        margin-bottom: 20px; 
        padding-bottom: 5px; 
        border-bottom: 3px solid var(--accent-color); 
    }}
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {{ 
        gap: 20px; 
        justify-content: center; 
        margin-bottom: 30px; 
        margin-top: 20px; 
    }}
    .stTabs [data-baseweb="tab"] {{ 
        height: 50px; 
        padding: 0 30px; 
        background-color: var(--white-color); 
        border-radius: 12px 12px 0 0; 
        border: 1px solid var(--grey-light-color); 
        font-weight: 700; 
        color: var(--text-default-color); 
        transition: all 0.2s ease-in-out; 
        font-size: 1.1em; 
    }}
    .stTabs [data-baseweb="tab"]:hover {{ 
        background-color: var(--accent-color); 
        color: var(--white-color); 
        border-color: var(--accent-color); 
    }}
    .stTabs [data-baseweb="tab"][aria-selected="true"] {{ 
        background-color: var(--primary-color); 
        color: var(--white-color); 
        border-top: 5px solid var(--accent-color); 
        border-color: var(--primary-color); 
        transform: translateY(-5px); 
        box-shadow: 0 6px 15px rgba(0,0,0,0.2); 
    }}
    /* Botão de Download */
    .stDownloadButton > button {{
        background-color: var(--success-color);
        color: var(--white-color);
        border: none;
        padding: 12px 25px;
        border-radius: 10px;
        font-weight: bold;
        transition: background-color 0.3s ease;
        box-shadow: 0 2px 5px rgba(0,0,0,0.2);
    }}
    .stDownloadButton > button:hover {{
        background-color: var(--secondary-accent-color);
    }}
    /* Expansor de Filtro */
    [data-testid="stExpander"] {{
        border-radius: 12px;
        overflow: hidden;
        border: 1px solid var(--grey-light-color);
        box-shadow: 0 2px 8px var(--shadow-light-color);
    }}
    [data-testid="stExpander"] [data-baseweb="toggle"] {{
        background-color: var(--white-color);
        padding: 15px;
    }}
    </style>
""", unsafe_allow_html=True)


# --- UI Principal ---
st.set_page_config(page_title="Painel de Performance Busca Ativa", layout="wide", initial_sidebar_state="expanded")

# Título Principal com Estilo
st.markdown('<div class="main-title-container">', unsafe_allow_html=True)
if config.logo_path.exists():
    try: 
        st.image(str(config.logo_path), width=80) # Reduzindo o tamanho da logo para o cabeçalho
    except: 
        st.warning("Não foi possível carregar a logo.")
st.markdown('<h1>Busca Ativa: Painel de Performance</h1>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# -------------------------------------------------------------
# | NOVO BLOCO DE DIAGNÓSTICO: O QUE O CÓDIGO ESTÁ VENDO?     |
# -------------------------------------------------------------
if df_principal.empty:
    st.error("ERRO CRÍTICO: O DataFrame principal está VAZIO (0 linhas).")
    st.warning(f"Verifique se o arquivo **'{config.excel_file}'** existe no mesmo diretório e se a planilha **'{config.excel_sheet}'** está correta.")
    st.info(f"O Streamlit não pode carregar o painel se a base de dados principal estiver vazia. Corrija o arquivo e reinicie o app.")
    st.stop()

with st.expander("🔬 VERIFICAÇÃO DE DADOS (APENAS DIAGNÓSTICO)", expanded=False):
    st.markdown("#### DataFrame Original (Primeiras 5 Linhas)")
    st.dataframe(df_principal.head(5))
    st.info(f"O DataFrame tem **{len(df_principal)}** linhas e as colunas detectadas são: **{list(df_principal.columns)}**")
    st.markdown("---")
# -------------------------------------------------------------


# --- Controles de Filtro (Revisão Final da Lógica) ---
with st.expander("🛠️ Configurações de Filtro", expanded=True):
    col_regional, col_municipio, col_data_devolucao = st.columns(3)

    opcoes_regional = sorted(df_principal['REGIONAL'].unique()) if not df_principal.empty else []
    selecao_regional = col_regional.multiselect("Selecione Regional:", options=opcoes_regional, default=opcoes_regional, key="ms_regional")
    
    # Filtra municípios com base na seleção regional
    df_municipios_base = df_principal[df_principal['REGIONAL'].isin(selecao_regional)] if selecao_regional else df_principal
    opcoes_municipio = sorted(df_municipios_base['MUNICIPIO'].unique()) if not df_municipios_base.empty else []
    
    # Se a seleção de municípios estiver vazia, usa todas as opções disponíveis (default behavior)
    selecao_municipio_raw = col_municipio.multiselect("Selecione Município:", options=opcoes_municipio, default=opcoes_municipio, key="ms_municipio")
    
    # Tratamento de Município (usa todas as opções se a seleção estiver vazia)
    selecao_municipio = selecao_municipio_raw if selecao_municipio_raw else opcoes_municipio

    # --- Filtro de Data ---
    # Usando DATA_FILTRO_AUX (tipo date) que é mais limpa e segura para o date_input
    opcoes_data_raw = sorted(df_principal['DATA_FILTRO_AUX'].dropna().unique()) if 'DATA_FILTRO_AUX' in df_principal.columns else []
    
    data_start = None
    data_end = None
    df_base_total_temp = pd.DataFrame()

    if opcoes_data_raw and len(opcoes_data_raw) > 0:
        data_min = opcoes_data_raw[0]
        data_max = opcoes_data_raw[-1]
        
        selecao_data_range = col_data_devolucao.date_input(
            "Selecione o Período de Devolução:",
            value=[data_min, data_max] if data_min and data_max else None,
            min_value=data_min,
            max_value=data_max,
            key="di_data_devolucao"
        )
        
        if isinstance(selecao_data_range, list) and len(selecao_data_range) == 2:
            data_start, data_end = selecao_data_range[0], selecao_data_range[1]
        elif isinstance(selecao_data_range, dt.date):
            data_start, data_end = selecao_data_range, selecao_data_range
        
        
    # --- PRÉ-FILTRAGEM DE ACORDO COM AS SELEÇÕES ---
    if not df_principal.empty and selecao_regional and selecao_municipio and data_start and data_end:
        
        # 1. Filtro de Região e Município
        df_base_total_temp = df_principal[
            (df_principal['REGIONAL'].isin(selecao_regional)) &
            (df_principal['MUNICIPIO'].isin(selecao_municipio))
        ].copy()
        
        # 2. Filtro de Data (Aplicado apenas se o pré-filtro não estiver vazio)
        if not df_base_total_temp.empty and data_start <= data_end:
            df_base_total_temp = df_base_total_temp[
                (df_base_total_temp['DATA_FILTRO_AUX'] >= data_start) &
                (df_base_total_temp['DATA_FILTRO_AUX'] <= data_end)
            ].copy()
        
        # 3. Aviso Visual se o DataFrame Filtrado estiver vazio
        if df_base_total_temp.empty:
            st.warning("A combinação de filtros selecionada (Região + Município + Data) não retornou nenhum dado. Tente expandir o intervalo de datas ou o filtro de Município.")
    
# --- Aplicação dos Filtros (Início da Lógica Principal) ---
if not df_principal.empty and not df_base_total_temp.empty:
    
    df_base_total = df_base_total_temp.copy()
    
    # Dataframe de análise para os colaboradores do gestor (MF)
    colaboradores_upper = [c.upper().strip() for c in config.colaboradores_list]
    df_para_analise = df_base_total[df_base_total[config.coluna_colaborador].isin(colaboradores_upper)].copy()

    kpis = calcular_indicadores_totais(df_base_total, df_para_analise, config.colaboradores_list)
    df_colab_performance = agregar_desempenho_colaborador(df_para_analise, config.colaboradores_list)
    
    # Metas são calculadas com base na BASE TOTAL para ser o progresso geral
    metas_kpis = calcular_metas_por_regional(df_principal, config.metas_regionais, config.metas_regionais.keys()) 
    metas_kpis_filtradas = [k for k in metas_kpis if k['regional'] in selecao_regional]


    # --- TABS DE VISUALIZAÇÃO ---
    tab_base, tab_colaboradores = st.tabs(["📊 ANÁLISE GERAL DA BASE", "👥 DESEMPENHO POR COLABORADOR MF"])

    with tab_base:
        st.markdown("### 🎯 Acompanhamento de Metas (Base Total)")
        
        # Exibe os KPIs de Meta com a nova função
        st.markdown(f"<div style='border: 1px solid {config.palette['GREY_LIGHT']}; border-radius: 12px; padding: 15px; background-color: {config.palette['WHITE']}; box-shadow: 0 2px 5px {config.palette['SHADOW_LIGHT']}'>", unsafe_allow_html=True)
        for kpi_regional in metas_kpis_filtradas:
            display_regional_goals(kpi_regional)
            st.markdown("---") # Separador visual entre as metas
        st.markdown("</div>", unsafe_allow_html=True)
        
        st.markdown("### 📈 Indicadores Chave de Performance (KPIs)")
        
        col1_base, col2_base, col3_base, col4_base, col5_base = st.columns(5)
        
        with col1_base: 
            display_kpi_metric("Total de Alvos (Cond. Obj. SIM)", kpis['total'], "📋")
        
        # Corrigido: Valor formatado para evitar injeção de HTML
        sub_value_exec = f"{formatar_inteiro(kpis['executados_produtivos'])} Prod. / {formatar_inteiro(kpis['executados_improdutivos'])} Improd."

        with col2_base: 
            display_kpi_metric(
                "Executados", 
                kpis['executados_totais'], 
                "✅",
                sub_label="Detalhe", 
                sub_value=sub_value_exec
            )
        with col3_base: 
            display_kpi_metric("Em Campo", kpis['em_campo'], "🛠️")
        with col4_base: 
            display_kpi_metric("A Atribuir", kpis['a_atribuir'], "🆕")
        with col5_base: 
            display_kpi_metric("Pendentes", kpis['pendentes'], "📤")
            if kpis['pendentes'] > 0:
                st.warning(f"⚠️ **Atenção!** Existem {formatar_inteiro(kpis['pendentes'])} serviços pendentes.")

        st.markdown("### 📊 Análise Detalhada da Base")
        
        # --- Gráfico de Executados (Produtivos e Improdutivos) ---
        st.markdown("#### Serviços Executados (Produtivos vs Improdutivos)")
        
        col_radio_base, col_vazio = st.columns([1, 3])
        visao_dimensao_executados_base = col_radio_base.radio(
            "Agrupar Executados por:", 
            ["Regional", "Município"], 
            key="radio_dimensao_executados_base", 
            horizontal=True
        )
        
        coluna_agregacao_base = 'REGIONAL' if visao_dimensao_executados_base == "Regional" else 'MUNICIPIO'
        
        df_produtivos = agregar_por_dimensao(df_base_total, coluna_agregacao_base, 'produtivos')
        df_improdutivos = agregar_por_dimensao(df_base_total, coluna_agregacao_base, 'improdutivos')
        
        if not df_produtivos.empty or not df_improdutivos.empty:
            df_plot_prod_improd = pd.concat([
                df_produtivos.assign(Tipo='Produtivo'), 
                df_improdutivos.assign(Tipo='Improdutivo')
            ]).sort_values(by='Métrica', ascending=False)
            
            color_map = {'Produtivo': config.palette['CHART_BAR_1'], 'Improdutivo': config.palette['ACCENT']} 
            
            fig = plot_bar_chart_v2(
                df_plot_prod_improd, 
                x_col='Dimensão', 
                y_col='Métrica', 
                title=f"Distribuição Produtivo/Improdutivo por {visao_dimensao_executados_base}",
                x_label=visao_dimensao_executados_base,
                y_label='Quantidade',
                color_col='Tipo',
                color_discrete_map=color_map,
                orientation='h'
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Nenhum dado de Produtivos/Improdutivos disponível para a seleção.")
            
        st.markdown("---") 

        # --- Gráfico de A Atribuir ---
        st.markdown("#### Serviços a Atribuir (Condição Objetiva SIM)")
        visao_atribuir_base = st.radio("Agrupar A Atribuir por:", ["Regional", "Município"], key="radio_atribuir_base", horizontal=True)
        coluna_atribuir_base = 'REGIONAL' if visao_atribuir_base == "Regional" else 'MUNICIPIO'
        df_agregado_atribuir = agregar_por_dimensao(df_base_total, coluna_atribuir_base, 'a_atribuir')
        
        if not df_agregado_atribuir.empty:
            st.plotly_chart(plot_bar_chart_v2(
                df_agregado_atribuir, 
                'Dimensão', 
                'Métrica', 
                'Serviços a Atribuir por ' + visao_atribuir_base, 
                visao_atribuir_base, 
                'Quantidade', 
                color_discrete_sequence=[config.palette['WARNING']] 
            ), use_container_width=True)
        else:
            st.info("Nenhum dado de 'Serviços a Atribuir' disponível para a seleção.")
        
        st.markdown("---") 

        # --- Gráfico de Pendentes ---
        st.markdown("#### Serviços Pendentes")
        visao_pendentes_base = st.radio("Agrupar Pendentes por:", ["Regional", "Município"], key="radio_pendentes_base", horizontal=True)
        coluna_pendentes_base = 'REGIONAL' if visao_pendentes_base == "Regional" else 'MUNICIPIO'
        df_agregado_pendentes = agregar_por_dimensao(df_base_total, coluna_pendentes_base, 'pendentes')
        
        if not df_agregado_pendentes.empty:
            st.plotly_chart(plot_bar_chart_v2(
                df_agregado_pendentes, 
                'Dimensão', 
                'Métrica', 
                'Serviços Pendentes por ' + visao_pendentes_base, 
                visao_pendentes_base, 
                'Quantidade', 
                color_discrete_sequence=[config.palette['DANGER']] 
            ), use_container_width=True)
        else:
            st.info("Nenhum dado de 'Serviços Pendentes' disponível para a seleção.")

    with tab_colaboradores:
        st.markdown("### 🧑‍💻 Desempenho Individual da Equipe")
        
        colaboradores_nao_encontrados = kpis['colaboradores_nao_encontrados']
        
        if colaboradores_nao_encontrados:
            st.warning(f"⚠️ **Atenção:** Os seguintes colaboradores da sua lista não foram encontrados na base de dados: {', '.join(colaboradores_nao_encontrados)}")
            st.markdown("---")
            
        if not df_colab_performance.empty:
            
            st.markdown("#### Top 10 Colaboradores por Produtivos")
            
            fig_top10 = px.bar(
                df_colab_performance.head(10).sort_values(by='Qtd_Produtivos', ascending=True),
                y=config.coluna_colaborador, 
                x='Qtd_Produtivos',
                text='Qtd_Produtivos',
                orientation='h',
                title="<b>Top 10 - Quantidade de Serviços Produtivos</b>",
                labels={config.coluna_colaborador: "Colaborador", "Qtd_Produtivos": "Serviços Produtivos"},
                color_discrete_sequence=[config.palette['PRIMARY']],
                template='plotly_white'
            )
            fig_top10.update_layout(
                yaxis={'categoryorder': 'total ascending'},
                title_font_color=config.palette['SECONDARY_ACCENT'],
                margin=dict(l=20, r=60, t=50, b=20),
                title_x=0.05,
            )
            fig_top10.update_traces(
                texttemplate='%{x}',
                textposition='outside',
                marker_line_width=1,
                marker_line_color=config.palette['WHITE']
            )
            st.plotly_chart(fig_top10, use_container_width=True)

            st.markdown("---")
            st.markdown("#### Tabela Completa de Desempenho")
            
            col_search, col_sort = st.columns([2, 2])
            with col_search:
                search_term = st.text_input("🔍 Pesquisar por nome:", "").upper()
            with col_sort:
                sort_col = st.selectbox("Ordenar por:", ['Qtd_Produtivos', 'Taxa_Produtividade', 'Qtd_Executados', 'Qtd_Alocados'], index=0)
            
            df_filtrado_colab = df_colab_performance
            if search_term:
                df_filtrado_colab = df_filtrado_colab[df_filtrado_colab[config.coluna_colaborador].str.contains(search_term, na=False)]

            df_filtrado_colab = df_filtrado_colab.sort_values(by=sort_col, ascending=False)
            
            # Formatando a tabela para exibição no Streamlit
            df_display_colab = df_filtrado_colab.assign(**{
                'Qtd_Executados': df_filtrado_colab['Qtd_Executados'].apply(formatar_inteiro),
                'Qtd_Produtivos': df_filtrado_colab['Qtd_Produtivos'].apply(formatar_inteiro),
                'Qtd_Improdutivos': df_filtrado_colab['Qtd_Improdutivos'].apply(formatar_inteiro),
                'Qtd_Em_Campo': df_filtrado_colab['Qtd_Em_Campo'].apply(formatar_inteiro),
                'Qtd_Alocados': df_filtrado_colab['Qtd_Alocados'].apply(formatar_inteiro),
                'Taxa_Produtividade': df_filtrado_colab['Taxa_Produtividade'].apply(lambda x: f"{x:.1%}") # Novo KPI de produtividade
            }).rename(columns={'Taxa_Produtividade': 'Taxa de Produtividade', config.coluna_colaborador: 'Colaborador'})
            
            st.dataframe(df_display_colab, use_container_width=True, hide_index=True)

            st.markdown("---")
            st.markdown("### 📄 Download do Relatório Executivo")
            
            buffer_pdf = BytesIO()
            # Passando df_base_total (filtrada) para o PDF para gerar as tabelas de resumo corretamente
            pdf_data = RelatorioVisualPDF(config.logo_path, config.palette, buffer_pdf).generate_report(df_base_total, df_para_analise, df_colab_performance, colaboradores_nao_encontrados, selecao_regional)
            
            if pdf_data:
                st.download_button(
                    label="📥 Gerar e Baixar Relatório em PDF",
                    data=pdf_data,
                    file_name=f"Relatorio_Busca_Ativa_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                    mime="application/pdf"
                )
        else:
            st.info("Nenhum dado de desempenho disponível para os colaboradores selecionados na base filtrada.")

else:
    st.info("Por favor, selecione as opções nos filtros para exibir os dados do painel. Verifique se o arquivo de dados está presente e os filtros estão preenchidos.")

st.markdown("---")
st.markdown(f"<p style='text-align:center; font-size:14px; color:{config.palette['GREY_DARK']};'>Criado por PLINIO M. RODRIGUES. &copy; {datetime.now().year}</p>", unsafe_allow_html=True)