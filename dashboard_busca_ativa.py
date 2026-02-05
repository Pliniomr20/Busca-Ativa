import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from pathlib import Path
from io import BytesIO

# --- IMPORTS PARA PDF ---
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak 
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor, white, lightgrey
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# --- 1. Camada de Configuração e Constantes ---
class Config:
    def __init__(self):
        self.base_dir = Path(__file__).resolve().parent
        self.logo_path = self.base_dir / "LOGO_M_F_1_-removebg-preview.png"
        self.font_path = self.base_dir / 'ARIALN.TTF'
        self.font_bold_path = self.base_dir / 'ARIALNB.TTF'
        self.excel_file = "BUSCA ATIVA.xlsx"
        self.excel_sheet = "Sheet1"
        self.coluna_colaborador = 'NOME_AGENTE'
        self.colaboradora_destaque = 'INGRITH LORENA PEREIRA DE OLIVEIRA' 

        # --- PALETA DE CORES INSTITUCIONAL ---
        self.palette = {
            "PRIMARY": "#1565C0",           # Azul Institucional (Forte e Sóbrio)
            "ACCENT": "#42A5F5",            # Azul Acento (Detalhes)
            "BG_PAGE": "#F4F6F9",           # Cinza-Azulado (Fundo Dashboard Corporativo)
            "BG_CARD": "#FFFFFF",           # Branco (Fundo dos Blocos)
            "TEXT_MAIN": "#2C3E50",         # Cinza Escuro (Leitura)
            "TEXT_LIGHT": "#7F8C8D",        # Cinza Médio (Labels)
            "BORDER": "#E0E0E0",            # Bordas Sutis
            "SUCCESS": "#27AE60",
            "WARNING": "#F39C12",
            "DANGER": "#C0392B",
            "WHITE": "#FFFFFF"
        }
        
        self.metas_regionais = {'NORTE': 4764, 'NORDESTE': 2418, 'SUL': 4547}
        self.servicos = {
            'executados': ['CONCLUIDO OK', 'DESCARREGADO COM IMPEDIMENTO', 'DESCARREGADO SEM IMPEDIMENTO', 'IMPROCEDENTE'],
            'produtivos': ['CONCLUIDO OK'],
        }

        self.colaboradores_list = [
            'ADNEY HENRIQUE NOGUEIRA LOPES', 'ADRIANO RIBEIRO SANTOS', 'ALAN ALVES AURELIANO', 'ALEX SILVA OLIVEIRA',
            'ALEXANDRE GURGEL DO AMARAL', 'ALONSO GONZAGA DA SILVA', 'ANTONIO SALIM GARCIA', 'BRENNER OLIVEIRA DE MELO',
            'BRENNO PEREIRA CAMPOS DE OLIVEIRA', 'BRUNO ALVES FERREIRA', 'BRUNO HENRIQUE DE MARINS CABRAL',
            'BRUNO HENRIQUE GOMES DE BRITO FREITAS', 'CAIO GUSTAVO DANTAS SILVA', 'CARLOS DANIEL CUSTODIO DA SILVA',
            'CARLOS HENRIQUE GONCALVES MELO', 'CLEBER PEREIRA CARDOSO', 'CLEITON ARAUJO DE OLIVEIRA',
            'CLEMILSON RODRIGUES DA TRINDADE', 'CRISTIANO DE JESUS MONTEIRO', 'DAMIAO PEREIRA DE MENESES',
            'DANIEL LUIZ CORREIA PANTA', 'DANILO MIGUEL DE OLIVEIRA', 'DAYANNE GABRIELLE DIAS', 'DEIVID SOUZA SILVA',
            'DHYOGO VIEIRA DE MOURA', 'DIEGO FONSECA DOS SANTOS', 'DJALMA MACIEL MARTINS', 'DORISMAR DUARTE SANTOS',
            'DOUGLAS ALVES DA COSTA', 'DOUGLAS KAIQUE DOS SANTOS REIS', 'EDIVALDO MOURA DE OLIVEIRA',
            'ELVIS DO NASCIMENTO RIBEIRO', 'FABIO WILLIAN OLIVEIRA DE MIRANDA', 'FERNANDO FERREIRA DE LIMA',
            'FILLIPE RODRIGUES DE SOUZA', 'FLAVIO DOURADO DE SOUZA', 'FLAVIO FERREIRA BORGES',
            'FRANCISCO DAS CHAGAS DE SOUSA SANTOS', 'GUILHERME SCOTT BASILIO ONOFRE',
            'HELLEN CRISTINA VALADARES FERREIRA', 'HENRIQUE BARBOSA NUNES', 'HIGOR VINICIUS DE CASTRO',
            'HYGOR DOS SANTOS SOUSA', 'IDAMAR VIEIRA DE OLIVEIRA FILHO', 'IGOR SILVA SANTOS',
            'IRISLAN SANTINNI TORRES DE SOUSA', 'IURY MIKAEL DE OLIVEIRA RICARDO', 'JEAN VITOR SOUZA MENDES',
            'JEFFERSON DOUGLAS DE SOUSA MAIA', 'JOAO NETO ROCHA DA SILVA', 'JOAO PAULO DOS SANTOS EUROPEU',
            'JOAO VITOR VIEIRA DOS SANTOS', 'JOENDERSON DE JESUS AVELINO', 'JONATAN RODRIGO BATISTA FELIX',
            'JONATHAN FARIA OLIVEIRA', 'JONATHAN LIMA DA ROCHA MACHADO', 'JOSE DOURADO DE OLIVEIRA FILHO',
            'JOSE WILLAME DA SILVA MOTA', 'JOVEHYRIS DE OLIVEIRA FRANCA', 'JULIANO DE ALENCAR RODRIGUES',
            'KEVERSON ANTONIO DE SOUZA SIQUEIRA', 'KILDERY VALVERDE DOS SANTOS', 'KLEBER FERNANDES DE AZEVEDO',
            'KLEVER PEREIRA DOS SANTOS', 'LAISIO DA SILVA ALEXANDRINO DE JESUS', 'LARISSON PEREIRA DIAS',
            'LAZARO BRAZ DE SOUSA', 'LUAN GABRIEL SANTANA SANTIAGO', 'LUCAS COSTA LINO', 'LUCIMAR DE MENDONCA',
            'LUIZ DA SILVA SANTOS', 'MAIK DA CONCEICAO SILVA', 'MARCELO ANTONIO MARTINS FILHO',
            'MARCELO MENDES RAMOS', 'MARCIO PAULO SILVA', 'MARCIO WAGNER JOSE LOPES SANCHES',
            'MARCO AURELIO DA SILVA LIMA', 'MARCOS ANTONIO RODRIGUES DA SILVA', 'MARK ETIENNE RODRIGUES DA COSTA',
            'MARLLON BRUNNO ALEM ALVES', 'MATEUS DIAS DOS SANTOS', 'MATEUS FERREIRA SOUZA',
            'MATEUS LIMA MENDONCA', 'MATHEUS DE JESUS SILVA', 'MAURICIO DIAS DA SILVA', 'MAYCON EDUARDO FIGUEREDO',
            'MICHAEL CAMPOS MORAIS', 'MICHAEL DOUGLAS DOTI PEREIRA', 'MURILLO GABRIEL DA SILVA LOBO',
            'MURILO MATHEUS BORGES RODRIGUES', 'NALLISSON THIAGO NASCIMENTO SILVA', 'NELSON NERES SOARES',
            'ODEILDO DA COSTA SANTANA', 'OTAVIO RODRIGUES OLIMPIO', 'PAULO VINICIOS HABERMANN DA ROCHA PINTO',
            'PEDRO HENRIQUE CIRINO DE MELO', 'PEDRO HENRIQUE DA CRUZ', 'PEDRO VICTOR NASCIMENTO DA SILVA',
            'RAFAEL DUARTE MARQUES', 'RAFAEL PEREIRA DE OLIVEIRA', 'RAPHAEL SILVA DE SOUZA',
            'RICARDO DA SILVA PEREIRA', 'RICARDO DE AMORIM CARNEIRO', 'RIVALDO JOSE DA SILVA',
            'RODRIGGO WAGNER CAMPOS DA SILVA', 'RONAN DA PENHA DE MORAIS', 'RONILSON DAS CHAGAS OLIVEIRA',
            'RONIS MARCIO CANDIDO FERREIRA', 'ROSIMAR PEREIRA LEITE', 'SAMUEL ALVES DIAS',
            'SANDOVAL JUNIOR NASCIMENTO DAS CHAGAS', 'SANDRO SANTOS ARAUJO', 'TIAGO DA SILVA RAMOS',
            'TIAGO LUCIO FERNANDES SOUSA', 'VALDEMAR DE ALMEIDA FILHO', 'VITOR DE SOUZA FERNANDES SANTOS',
            'WANDERSON MENDES DE MOURA', 'WELBESON RODRIGUES DA COSTA', 'WENDER DE CASTRO VIEIRA',
            'WENDER SOARES DA SILVA', 'WESLEY DE SOUSA PEREIRA', 'WEVERSON DA SILVA',
            'WEVERTON CARLAZAN DE ARAUJO', 'KLEVER PEREIRA DOS SANTOS', 'INGRITH LORENA PEREIRA DE OLIVEIRA', 'ANA PAULA FERREIRA DOS SANTOS',
            'DAVI BATISTA RIBEIRO', 'GLEICYANE VIEIRA DA SILVA', 'MOISES FERREIRA MAIA PEDROSA', 'PLINIO RODRIGUES', 'GABRIELA APARECIDA SAMPAIO LOIOLA',
            'WESLEY MARTINS DE CASTRO', 'BERENICE RIBEIRO DE SOUZA'
        ]

config = Config()

# --- Funções de Formatação ---
def formatar_inteiro(valor: float | int) -> str:
    if pd.isna(valor) or valor is None: return "0"
    try: valor = int(valor)
    except (ValueError, TypeError): return "Inválido"
    return f"{valor:,}".replace(",", ".")

# --- 2. Camada de Acesso e Processamento ---
@st.cache_data(ttl=3600, show_spinner=False)
def carregar_e_processar_dados(caminho_arquivo: Path) -> pd.DataFrame:
    if not caminho_arquivo.exists(): st.stop() # Falha silenciosa ou controlada
    
    try:
        df = pd.read_excel(caminho_arquivo, sheet_name=config.excel_sheet)
        df.columns = df.columns.str.strip().str.upper().str.replace(' ', '_').str.replace('[^A-Z0-9_]', '', regex=False)
        
        required_cols = ['REGIONAL', 'MUNICIPIO', 'NOME_FASE', 'ALVO_CONDICAO_OBJETIVA', config.coluna_colaborador, 'DATA_DEVOLUCAO']
        if not all(col in df.columns for col in required_cols): st.stop()
        if df.empty: st.stop()
            
        df['NOME_FASE'] = df['NOME_FASE'].str.upper().str.strip()
        df['REGIONAL'] = df['REGIONAL'].str.upper().str.strip()
        df['MUNICIPIO'] = df['MUNICIPIO'].str.upper().str.strip()
        df[config.coluna_colaborador] = df[config.coluna_colaborador].str.upper().str.strip()
        
        df['DATA_DEVOLUCAO'] = pd.to_datetime(df['DATA_DEVOLUCAO'], errors='coerce', dayfirst=True)
        df.dropna(subset=['DATA_DEVOLUCAO'], inplace=True)
        
        df['Ano'] = df['DATA_DEVOLUCAO'].dt.year
        df['Mes'] = df['DATA_DEVOLUCAO'].dt.month
        df['Dia'] = df['DATA_DEVOLUCAO'].dt.day
        
        df = df[df['REGIONAL'].isin(['NORTE', 'NORDESTE', 'SUL'])].copy()
        return df
    
    except Exception:
        st.stop()

df_principal = carregar_e_processar_dados(config.base_dir / config.excel_file)

# --- 3. Lógica de Negócio ---
def calcular_indicadores_totais(df_base_total: pd.DataFrame, df_para_analise: pd.DataFrame, colaboradores_list: list) -> dict:
    if df_base_total.empty: return {"colaboradores_nao_encontrados": []}
    colaboradores_na_base = set(df_base_total[config.coluna_colaborador].unique())
    return {"colaboradores_nao_encontrados": [c for c in colaboradores_list if c.upper().strip() not in colaboradores_na_base]}

def get_desempenho_visitas_por_regional(df: pd.DataFrame, colaboradores_list: list) -> pd.DataFrame:
    colaboradores_upper = [c.upper().strip() for c in colaboradores_list]
    df_filtrado = df[df[config.coluna_colaborador].isin(colaboradores_upper)].copy()
    if df_filtrado.empty: return pd.DataFrame()
        
    df_agregado = df_filtrado.groupby([config.coluna_colaborador, 'REGIONAL']).agg(
        Qtd_Visitas_Total=('NOME_FASE', lambda x: x.isin(config.servicos['executados']).sum()),      
        Qtd_Produtivos=('NOME_FASE', lambda x: x.isin(config.servicos['produtivos']).sum())         
    ).reset_index()
    
    df_agregado.columns = ['Colaborador', 'Regional', 'Qtd_Visitas_Total', 'Qtd_Produtivos']
    df_agregado['% Produtividade'] = (df_agregado['Qtd_Produtivos'] / df_agregado['Qtd_Visitas_Total']) * 100
    df_agregado['% Produtividade'] = df_agregado['% Produtividade'].fillna(0).round(2) 
    return df_agregado.sort_values(by=['Regional', 'Qtd_Produtivos'], ascending=[True, False]).reset_index(drop=True)

def get_performance_individual(df: pd.DataFrame, nome_colaborador: str) -> pd.DataFrame:
    df_individual = df[df['Colaborador'] == nome_colaborador.upper().strip()].copy()
    if df_individual.empty: return pd.DataFrame()
    df_individual['Total Visitas'] = df_individual['Qtd_Visitas_Total'].apply(formatar_inteiro)
    df_individual['Concluído OK'] = df_individual['Qtd_Produtivos'].apply(formatar_inteiro)
    df_individual['Produtividade (%)'] = df_individual['% Produtividade'].apply(lambda x: f"{x:.2f}%")
    return df_individual

def plot_bar_chart_produtividade_regional(df_data):
    df_prod_resumo = df_data.groupby('Regional')['Qtd_Produtivos'].sum().reset_index(name='Total_Produtivo')
    if df_prod_resumo.empty: return px.bar(title="Sem dados")
    
    # Visual extremamente limpo e corporativo
    fig = px.bar(
        df_prod_resumo, 
        x='Regional', y='Total_Produtivo', text='Total_Produtivo', 
        color='Regional',
        color_discrete_sequence=[config.palette['PRIMARY'], config.palette['ACCENT'], "#90CAF9"],
        height=320 # Altura controlada
    )
    
    fig.update_traces(
        texttemplate='%{y}', textposition='outside',
        textfont=dict(size=13, color=config.palette["TEXT_MAIN"], family="Arial"),
        marker_line_width=0
    )
    
    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        showlegend=False, 
        margin=dict(t=20, b=30, l=10, r=10),
        yaxis=dict(showgrid=False, showticklabels=False, title=None, visible=False),
        xaxis=dict(showgrid=False, showline=True, linecolor=config.palette["BORDER"], tickfont=dict(size=12, color=config.palette["TEXT_MAIN"]))
    )
    return fig

# --- 4. Relatórios PDF (Inalterado, apenas ajustes de import) ---
class RelatorioVisualPDF:
    def __init__(self, logo_path: Path, palette: dict, output_buffer: BytesIO):
        self.logo_path = logo_path
        self.palette = palette
        self.buffer = output_buffer
        self.doc = SimpleDocTemplate(self.buffer, pagesize=A4, leftMargin=0.75*inch, rightMargin=0.75*inch, topMargin=1.0*inch, bottomMargin=0.75*inch)
        self.story = []
        self._register_fonts()
        self._define_styles()

    def _define_styles(self):
        styles = getSampleStyleSheet()
        self.styles = {
            'h1': ParagraphStyle('h1', parent=styles['h1'], fontName='Arial-Bold', fontSize=18, leading=22, alignment=TA_CENTER, textColor=HexColor(self.palette["PRIMARY"]), spaceAfter=12),
            'h2': ParagraphStyle('h2', parent=styles['h2'], fontName='Arial-Bold', fontSize=14, leading=16, textColor=HexColor(self.palette["PRIMARY"]), spaceAfter=8),
            'body': ParagraphStyle('body', parent=styles['Normal'], fontName='Arial', fontSize=11, leading=14, textColor=HexColor(self.palette["TEXT_MAIN"]), spaceAfter=8),
            'table_header': ParagraphStyle('table_header', parent=styles['Normal'], fontName='Arial-Bold', fontSize=9, leading=11, alignment=TA_CENTER, textColor=HexColor(self.palette["WHITE"])),
            'table_body': ParagraphStyle('table_body', parent=styles['Normal'], fontName='Arial', fontSize=9, leading=11, alignment=TA_LEFT, textColor=HexColor(self.palette["TEXT_MAIN"])),
            'table_body_center': ParagraphStyle('table_body_center', parent=styles['Normal'], fontName='Arial', fontSize=9, leading=11, alignment=TA_CENTER, textColor=HexColor(self.palette["TEXT_MAIN"])),
            'table_body_prod': ParagraphStyle('table_body_prod', parent=styles['Normal'], fontName='Arial-Bold', fontSize=9, leading=11, alignment=TA_CENTER, textColor=HexColor(self.palette["SUCCESS"])),
        }
    
    def _register_fonts(self):
        try:
            pdfmetrics.registerFont(TTFont('Arial', str(config.font_path)))
            pdfmetrics.registerFont(TTFont('Arial-Bold', str(config.font_bold_path)))
            pdfmetrics.registerFontFamily('Arial', normal='Arial', bold='Arial-Bold')
        except Exception:
            pdfmetrics.registerFontFamily('Helvetica', normal='Helvetica', bold='Helvetica-Bold')

    def _header_page(self, canvas, doc):
        canvas.saveState()
        canvas.setFont('Arial-Bold', 14)
        canvas.setFillColor(HexColor(self.palette["PRIMARY"]))
        canvas.drawCentredString(A4[0]/2.0, A4[1] - 0.4*inch, "Busca Ativa M&F - Desempenho de Visitas Produtivas")
        canvas.setStrokeColor(HexColor(self.palette["ACCENT"]))
        canvas.setLineWidth(1.5)
        canvas.line(0.75*inch, A4[1] - 0.85*inch, A4[0] - 0.75*inch, A4[1] - 0.85*inch)
        canvas.restoreState()

    def add_regional_performance_table(self, regional: str, df: pd.DataFrame):
        self.story.append(Paragraph(f"Desempenho: Regional {regional}", self.styles['h2']))
        if df.empty:
            self.story.append(Paragraph(f"Nenhum dado encontrado para a Regional {regional}.", self.styles['body']))
            self.story.append(Spacer(1, 0.1 * inch))
            return

        df_pdf = df.rename(columns={'Colaborador': 'Nome Agente', 'Qtd_Visitas_Total': 'Total Visitas', 'Qtd_Produtivos': 'Concluído OK', '% Produtividade': 'Produtividade (%)'}).sort_values(by='Concluído OK', ascending=False)
        df_pdf['Produtividade (%)'] = df_pdf['Produtividade (%)'].apply(lambda x: f"{x:,.2f}%".replace(",", "X").replace(".", ",").replace("X", "."))
        
        headers = [Paragraph(col, self.styles['table_header']) for col in ['Nome Agente', 'Total Visitas', 'Concluído OK', 'Produtividade (%)']]
        data = [headers]
        for _, row in df_pdf.iterrows():
            data.append([
                Paragraph(str(row['Nome Agente']), self.styles['table_body']),
                Paragraph(str(formatar_inteiro(row['Total Visitas'])), self.styles['table_body_center']),
                Paragraph(str(formatar_inteiro(row['Concluído OK'])), self.styles['table_body_center']),
                Paragraph(str(row['Produtividade (%)']), self.styles['table_body_prod'] if float(str(row['Produtividade (%)']).replace('%','').replace(',','.')) >= 50 else self.styles['table_body_center'])
            ])

        col_widths = [self.doc.width * 0.45, self.doc.width * 0.18, self.doc.width * 0.18, self.doc.width * 0.19]
        table = Table(data, colWidths=col_widths)
        table_style_list = [
            ('BACKGROUND', (0,0), (-1,0), HexColor(self.palette["PRIMARY"])),
            ('TEXTCOLOR', (0,0), (-1,0), HexColor(self.palette["WHITE"])),
            ('ALIGN', (0,0), (-1,0), 'CENTER'), ('ALIGN', (1,1), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('GRID', (0,0), (-1,-1), 0.5, HexColor(self.palette["BORDER"])),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8), ('TOPPADDING', (0,0), (-1,-1), 8),
        ]
        for i in range(1, len(data)):
            table_style_list.append(('BACKGROUND', (0,i), (-1,i), HexColor(self.palette["BG_PAGE"]) if i % 2 == 0 else HexColor(self.palette["WHITE"])))
        
        table.setStyle(TableStyle(table_style_list))
        self.story.append(table)
        self.story.append(Spacer(1, 0.2 * inch))
        self.story.append(PageBreak())

    def generate_report(self, df_colab_desempenho: pd.DataFrame, colaboradores_nao_encontrados: list):
        self.story.append(Paragraph("Busca Ativa M&F - Relatório de Desempenho", self.styles['h1']))
        self.story.append(Paragraph(f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", self.styles['body']))
        self.story.append(Spacer(1, 0.2 * inch))
        if colaboradores_nao_encontrados:
            self.story.append(Paragraph(f"<b>ATENÇÃO:</b> {len(colaboradores_nao_encontrados)} colaboradores não encontrados.", self.styles['body']))
            self.story.append(Spacer(1, 0.1 * inch))

        for regional in ['SUL', 'NORTE', 'NORDESTE']:
            self.add_regional_performance_table(regional, df_colab_desempenho[df_colab_desempenho['Regional'] == regional].copy())
            
        try:
            self.doc.build(self.story, onFirstPage=self._header_page, onLaterPages=self._header_page)
            return self.buffer.getvalue()
        except Exception: return None

# --- 5. UI/UX Principal (CSS RÍGIDO E ESTRUTURAL) ---
st.set_page_config(page_title="Busca Ativa M&F", layout="wide", initial_sidebar_state="collapsed")

st.markdown(f"""
    <style>
    /* --- FONTE GLOBAL & BACKGROUND --- */
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap');
    
    html, body, .stApp, .stAppViewContainer {{
        background-color: {config.palette["BG_PAGE"]} !important;
        font-family: 'Roboto', sans-serif !important;
        color: {config.palette["TEXT_MAIN"]};
    }}
    
    /* REMOVER PADDING PADRÃO DO STREAMLIT QUE ESTRAGA O LAYOUT */
    .block-container {{
        padding-top: 1rem !important;
        padding-bottom: 5rem !important;
        max-width: 1300px !important;
    }}
    header {{ visibility: hidden; }} /* Esconde a barra colorida padrão */
    
    /* --- HEADER PERSONALIZADO --- */
    .custom-header {{
        background-color: white;
        padding: 1rem 1.5rem;
        border-radius: 8px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        display: flex;
        align-items: center;
        margin-bottom: 20px;
        border-bottom: 3px solid {config.palette["PRIMARY"]};
    }}
    .header-logo {{ font-size: 24px; margin-right: 15px; font-weight: bold; color: {config.palette["PRIMARY"]}; }}
    .header-title {{ font-size: 20px; font-weight: 700; color: {config.palette["TEXT_MAIN"]}; margin: 0; }}
    .header-subtitle {{ font-size: 14px; font-weight: 400; color: {config.palette["TEXT_LIGHT"]}; margin-top: 2px; }}

    /* --- CARDS (CONTAINER BRANCO) --- */
    .section-card {{
        background-color: white;
        padding: 20px;
        border-radius: 8px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        margin-bottom: 20px;
        border: 1px solid {config.palette["BORDER"]};
    }}
    
    /* --- EXPANDER (FILTROS) --- */
    .streamlit-expanderHeader {{
        background-color: white !important;
        border-radius: 6px !important;
        font-weight: 500 !important;
        color: {config.palette["PRIMARY"]} !important;
        border: 1px solid {config.palette["BORDER"]} !important;
        font-size: 15px !important;
    }}
    [data-testid="stExpander"] {{
        background-color: white;
        border-radius: 6px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        border: none;
    }}

    /* --- KPIS (MÉTRICAS) PADRONIZADAS --- */
    [data-testid="stMetric"] {{
        background-color: white;
        padding: 15px;
        border-radius: 8px;
        border: 1px solid {config.palette["BORDER"]};
        box-shadow: 0 1px 2px rgba(0,0,0,0.03);
        text-align: center;
        height: 100%; /* Força altura igual */
    }}
    /* Label do KPI */
    [data-testid="stMetricLabel"] {{
        font-size: 13px !important;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        color: {config.palette["TEXT_LIGHT"]} !important;
        margin-bottom: 5px !important;
    }}
    /* Valor do KPI (Tamanho controlado) */
    [data-testid="stMetricValue"] {{
        font-size: 28px !important; /* Tamanho fixo, nem gigante nem pequeno */
        font-weight: 700 !important;
        color: {config.palette["PRIMARY"]} !important;
    }}
    /* Delta (Pequeno) */
    [data-testid="stMetricDelta"] {{
        font-size: 12px !important;
        margin-top: 5px;
    }}

    /* --- TÍTULOS DE SEÇÃO (HIERARQUIA) --- */
    h2 {{
        font-size: 18px !important;
        font-weight: 600 !important;
        color: {config.palette["TEXT_MAIN"]} !important;
        margin-top: 0 !important;
        margin-bottom: 15px !important;
        padding-bottom: 8px;
        border-bottom: 1px solid {config.palette["BORDER"]};
    }}
    h3 {{ font-size: 16px !important; font-weight: 600 !important; color: {config.palette["PRIMARY"]}; }}

    /* --- TABELAS --- */
    .stDataEditor {{ border: 1px solid {config.palette["BORDER"]}; border-radius: 6px; }}
    
    /* Correção de espaçamento de colunas */
    div[data-testid="column"] {{ padding: 0 5px; }}

    /* --- DESTAQUE HERO (INGRITH) --- */
    .hero-container {{
        background: linear-gradient(180deg, #FFFFFF 0%, #FAFAFA 100%);
        border-left: 5px solid #FFC107; /* Amarelo Ouro Sutil */
    }}

    </style>
""", unsafe_allow_html=True)

# --- HEADER ESTRUTURAL ---
st.markdown(f"""
    <div class="custom-header">
        <div class="header-logo">M&F</div>
        <div>
            <div class="header-title">Busca Ativa - Dashboard Executivo</div>
            <div class="header-subtitle">Performance e Monitoramento de Visitas</div>
        </div>
    </div>
""", unsafe_allow_html=True)

# --- ÁREA DE FILTROS (PAINEL DE CONTROLE) ---
with st.expander("⚙️  Filtros Globais", expanded=True):
    c1, c2, c3, c4, c5 = st.columns([2, 3, 1, 1, 1])
    
    with c1:
        opcoes_regional_orig = sorted(df_principal['REGIONAL'].unique()) if not df_principal.empty else []
        selecao_regional = st.multiselect("Regional", options=opcoes_regional_orig, default=opcoes_regional_orig)
    
    df_municipios = df_principal[df_principal['REGIONAL'].isin(selecao_regional)] if selecao_regional else pd.DataFrame()
    opcoes_mun = ["TODOS"] + sorted(df_municipios['MUNICIPIO'].unique()) if not df_municipios.empty else []
    
    with c2:
        sel_mun_raw = st.multiselect("Município", options=opcoes_mun, default=["TODOS"])
        selecao_municipio = sorted(df_municipios['MUNICIPIO'].unique()) if "TODOS" in sel_mun_raw else sel_mun_raw

    # Lógica de Data simplificada e robusta
    opcoes_ano = ["TODOS"] + sorted(df_principal['Ano'].unique()) if 'Ano' in df_principal.columns else []
    with c3:
        sel_ano_raw = st.multiselect("Ano", options=opcoes_ano, default=["TODOS"])
        selecao_ano = sorted(df_principal['Ano'].unique()) if "TODOS" in sel_ano_raw else [int(x) for x in sel_ano_raw]
    
    df_mes = df_principal[df_principal['Ano'].isin(selecao_ano)] if selecao_ano else pd.DataFrame()
    opcoes_mes = ["TODOS"] + sorted(df_mes['Mes'].unique()) if not df_mes.empty else []
    with c4:
        sel_mes_raw = st.multiselect("Mês", options=opcoes_mes, default=["TODOS"])
        selecao_mes = sorted(df_mes['Mes'].unique()) if "TODOS" in sel_mes_raw else [int(x) for x in sel_mes_raw]

    df_dia = df_mes[df_mes['Mes'].isin(selecao_mes)] if selecao_mes else pd.DataFrame()
    opcoes_dia = ["TODOS"] + sorted(df_dia['Dia'].unique()) if not df_dia.empty else []
    with c5:
        sel_dia_raw = st.multiselect("Dia", options=opcoes_dia, default=["TODOS"])
        selecao_dia = sorted(df_dia['Dia'].unique()) if "TODOS" in sel_dia_raw else [int(x) for x in sel_dia_raw]

# --- PROCESSAMENTO ---
if not df_principal.empty and selecao_regional and selecao_municipio:
    df_base = df_principal[
        (df_principal['REGIONAL'].isin(selecao_regional)) &
        (df_principal['MUNICIPIO'].isin(selecao_municipio)) &
        (df_principal['Ano'].isin(selecao_ano)) &
        (df_principal['Mes'].isin(selecao_mes)) &
        (df_principal['Dia'].isin(selecao_dia))
    ].copy()
    
    if df_base.empty:
        st.warning("Nenhum dado para os filtros selecionados.")
        st.stop()
        
    df_analise = df_base[df_base[config.coluna_colaborador].isin([c.upper().strip() for c in config.colaboradores_list])].copy()
    df_desempenho = get_desempenho_visitas_por_regional(df_analise, config.colaboradores_list)
    kpis = calcular_indicadores_totais(df_base, df_analise, config.colaboradores_list)
    
    # --- UI BLOCO 1: HERO (DESTAQUE) ---
    df_ingrith = get_performance_individual(df_desempenho, config.colaboradora_destaque)
    if not df_ingrith.empty:
        st.markdown('<div class="section-card hero-container">', unsafe_allow_html=True)
        col_hero_title, col_hero_main, col_hero_details = st.columns([1.5, 1, 2.5])
        
        with col_hero_title:
            st.markdown(f"### 👑 Destaque Profissional")
            st.markdown(f"**{config.colaboradora_destaque.title()}**")
            st.caption("Monitoramento de performance individual em destaque.")
        
        with col_hero_main:
            total_prod_ingrith = df_ingrith['Qtd_Produtivos'].sum()
            total_vis_ingrith = df_ingrith['Qtd_Visitas_Total'].sum()
            media_prod = (total_prod_ingrith/total_vis_ingrith)*100 if total_vis_ingrith > 0 else 0
            st.metric("Total Produtivos (Geral)", formatar_inteiro(total_prod_ingrith), f"{media_prod:.1f}% Eficiência")
            
        with col_hero_details:
             cols_reg = st.columns(3)
             for i, reg in enumerate(['SUL', 'NORTE', 'NORDESTE']):
                 dfr = df_ingrith[df_ingrith['Regional'] == reg]
                 with cols_reg[i]:
                     if not dfr.empty:
                         st.metric(reg, dfr['Concluído OK'].iloc[0], f"Visitas: {dfr['Total Visitas'].iloc[0]}")
                     else:
                         st.metric(reg, "-", "Sem dados")
        st.markdown('</div>', unsafe_allow_html=True)

    # --- UI BLOCO 2: VISÃO GERAL (GRÁFICO + KPIS) ---
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown("<h2>📊 Visão Geral da Operação</h2>", unsafe_allow_html=True)
    
    c_chart, c_kpis = st.columns([2, 1])
    
    with c_chart:
        fig = plot_bar_chart_produtividade_regional(df_desempenho)
        st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
    
    with c_kpis:
        st.markdown("### Resumo por Regional")
        df_resumo = df_desempenho.groupby('Regional')['Qtd_Produtivos'].sum().reset_index()
        for _, row in df_resumo.iterrows():
            reg = row['Regional']
            total = row['Qtd_Produtivos']
            # Calcula total de visitas para a regional para o delta
            visitas_reg = df_desempenho[df_desempenho['Regional'] == reg]['Qtd_Visitas_Total'].sum()
            delta_val = (total / visitas_reg * 100) if visitas_reg > 0 else 0
            
            st.metric(f"Produtivos {reg}", formatar_inteiro(total), f"{delta_val:.1f}% Eficiência")
            st.markdown("<div style='margin-bottom: 10px'></div>", unsafe_allow_html=True) # Espaço extra
            
    st.markdown('</div>', unsafe_allow_html=True)

    # --- UI BLOCO 3: TABELAS DETALHADAS ---
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown("<h2>📋 Detalhamento por Colaborador</h2>", unsafe_allow_html=True)
    
    tabs = st.tabs(["SUL", "NORTE", "NORDESTE"])
    
    for aba, reg in zip(tabs, ['SUL', 'NORTE', 'NORDESTE']):
        with aba:
            dfr = df_desempenho[df_desempenho['Regional'] == reg].copy()
            if dfr.empty:
                st.info(f"Sem dados registrados para a regional {reg}.")
            else:
                dfr = dfr.sort_values(by='Qtd_Produtivos', ascending=False).reset_index(drop=True)
                dfr['Posição'] = dfr.index + 1
                dfr['Colaborador'] = dfr['Colaborador'].apply(lambda x: f"👑 {x}" if x == config.colaboradora_destaque.upper() else x)
                
                df_show = dfr[['Posição', 'Colaborador', 'Qtd_Visitas_Total', 'Qtd_Produtivos', '% Produtividade']].rename(columns={
                    'Qtd_Visitas_Total': 'Total Visitas', 'Qtd_Produtivos': 'Produtivos', '% Produtividade': 'Eficiência'
                })
                
                st.dataframe(
                    df_show,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Posição": st.column_config.NumberColumn("#", width="small"),
                        "Colaborador": st.column_config.TextColumn("Colaborador", width="large"),
                        "Total Visitas": st.column_config.NumberColumn(format="%d"),
                        "Produtivos": st.column_config.NumberColumn(format="%d"),
                        "Eficiência": st.column_config.ProgressColumn(format="%.1f%%", min_value=0, max_value=100)
                    }
                )
    st.markdown('</div>', unsafe_allow_html=True)

    # --- FOOTER / DOWNLOAD ---
    col_f1, col_f2 = st.columns([4, 1])
    with col_f1:
        if kpis['colaboradores_nao_encontrados']:
            with st.expander(f"⚠️ {len(kpis['colaboradores_nao_encontrados'])} Colaboradores não encontrados na base"):
                st.write(", ".join(kpis['colaboradores_nao_encontrados']))
    
    with col_f2:
        buffer = BytesIO()
        pdf_bytes = RelatorioVisualPDF(config.logo_path, config.palette, buffer).generate_report(df_desempenho, kpis['colaboradores_nao_encontrados'])
        if pdf_bytes:
            st.download_button("📥 Baixar PDF Oficial", data=pdf_bytes, file_name=f"Relatorio_MF_{datetime.now().strftime('%Y%m%d')}.pdf", mime="application/pdf", use_container_width=True, type="primary")