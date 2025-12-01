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
        self.colaboradora_destaque = 'INGRITH LORENA PEREIRA DE OLIVEIRA' # Nome para destaque

        # --- PALETA AZUL ---
        self.palette = {
            "PRIMARY": "#1976D2",           # Azul Escuro Principal
            "ACCENT": "#42A5F5",            # Azul Claro (Destaque)
            "SECONDARY_ACCENT": "#90CAF9",
            "BACKGROUND_LIGHT": "#E3F2FD",  # Azul Muito Claro (Fundo)
            "TEXT_DEFAULT": "#212529",
            "GREY_LIGHT": "#e0e0e0",
            "GREY_DARK": "#757575",
            "WHITE": "#FFFFFF",
            "SHADOW_LIGHT": "rgba(0,0,0,0.08)",
            "SUCCESS": "#388E3C",           # Verde Escuro (Para Produtivo)
            "WARNING": "#FFB300",           # Laranja (Atenção)
            "DANGER": "#D32F2F"             # Vermelho
        }
        
        self.metas_regionais = {
            'NORTE': 4764,
            'NORDESTE': 2418,
            'SUL': 4547
        }
        
        self.servicos = {
            'executados': ['CONCLUIDO OK', 'DESCARREGADO COM IMPEDIMENTO', 'DESCARREGADO SEM IMPEDIMENTO', 'IMPROCEDENTE'],
            'produtivos': ['CONCLUIDO OK'],
        }

        # Lista de Colaboradores (Mantida inalterada)
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
            'WEVERTON CARLAZAN DE ARAUJO', 'KLEVER PEREIRA DOS SANTOS', 'INGRITH LORENA PEREIRA DE OLIVEIRA'
        ]

config = Config()

# --- Funções de Formatação e Utilitários ---
def formatar_inteiro(valor: float | int) -> str:
    if pd.isna(valor) or valor is None: return "0"
    try: valor = int(valor)
    except (ValueError, TypeError): return "Inválido"
    return f"{valor:,}".replace(",", ".")

# --- 2. Camada de Acesso e Processamento de Dados ---
@st.cache_data(ttl=3600, show_spinner="Carregando e processando dados de Busca Ativa...")
def carregar_e_processar_dados(caminho_arquivo: Path) -> pd.DataFrame:
    if not caminho_arquivo.exists():
        st.error(f"Erro Crítico: Arquivo de dados não encontrado em '{caminho_arquivo}'.")
        st.stop()
    
    try:
        df = pd.read_excel(caminho_arquivo, sheet_name=config.excel_sheet)
        
        df.columns = df.columns.str.strip().str.upper().str.replace(' ', '_').str.replace('[^A-Z0-9_]', '', regex=False)
        
        required_cols = ['REGIONAL', 'MUNICIPIO', 'NOME_FASE', 'ALVO_CONDICAO_OBJETIVA', config.coluna_colaborador, 'DATA_DEVOLUCAO']
        if not all(col in df.columns for col in required_cols):
            missing_cols = [col for col in required_cols if col not in df.columns]
            st.error(f"Colunas essenciais faltando na planilha: {missing_cols}")
            st.stop()
            
        if df.empty:
            st.warning("Nenhum dado foi encontrado após a aplicação dos filtros iniciais. Verifique a planilha.")
            st.stop()
            
        df['NOME_FASE'] = df['NOME_FASE'].str.upper().str.strip()
        df['REGIONAL'] = df['REGIONAL'].str.upper().str.strip()
        df['MUNICIPIO'] = df['MUNICIPIO'].str.upper().str.strip()
        df[config.coluna_colaborador] = df[config.coluna_colaborador].str.upper().str.strip()
        
        df['DATA_DEVOLUCAO'] = pd.to_datetime(df['DATA_DEVOLUCAO'], errors='coerce', dayfirst=True)
        df.dropna(subset=['DATA_DEVOLUCAO'], inplace=True)
        
        df['Ano'] = df['DATA_DEVOLUCAO'].dt.year
        df['Mes'] = df['DATA_DEVOLUCAO'].dt.month
        df['Dia'] = df['DATA_DEVOLUCAO'].dt.day

        regionais_validas = ['NORTE', 'NORDESTE', 'SUL']
        df = df[df['REGIONAL'].isin(regionais_validas)].copy()

        return df
    
    except Exception as e:
        st.error(f"Erro fatal ao carregar o arquivo Excel: {e}")
        st.exception(e)
        st.stop()

df_principal = carregar_e_processar_dados(config.base_dir / config.excel_file)


# --- 3. Camada de Lógica de Negócio e Agregação ---
def calcular_indicadores_totais(df_base_total: pd.DataFrame, df_para_analise: pd.DataFrame, colaboradores_list: list) -> dict:
    if df_base_total.empty:
        return {"colaboradores_nao_encontrados": []}
        
    colaboradores_na_base = set(df_base_total[config.coluna_colaborador].unique())
    colaboradores_nao_encontrados = [c for c in colaboradores_list if c.upper().strip() not in colaboradores_na_base]
    
    return {
        "colaboradores_nao_encontrados": colaboradores_nao_encontrados,
    }

def get_desempenho_visitas_por_regional(df: pd.DataFrame, colaboradores_list: list) -> pd.DataFrame:
    """Calcula o total de visitas realizadas, as produtivas e o % Produtividade por colaborador e regional."""
    
    colaboradores_upper = [c.upper().strip() for c in colaboradores_list]
    df_filtrado = df[df[config.coluna_colaborador].isin(colaboradores_upper)].copy()
    
    if df_filtrado.empty:
        return pd.DataFrame()
        
    df_agregado = df_filtrado.groupby([config.coluna_colaborador, 'REGIONAL']).agg(
        Qtd_Visitas_Total=('NOME_FASE', lambda x: x.isin(config.servicos['executados']).sum()),      
        Qtd_Produtivos=('NOME_FASE', lambda x: x.isin(config.servicos['produtivos']).sum())         
    ).reset_index()
    
    df_agregado.columns = ['Colaborador', 'Regional', 'Qtd_Visitas_Total', 'Qtd_Produtivos']

    df_agregado['% Produtividade'] = (df_agregado['Qtd_Produtivos'] / df_agregado['Qtd_Visitas_Total']) * 100
    df_agregado['% Produtividade'] = df_agregado['% Produtividade'].fillna(0).round(2) 

    return df_agregado.sort_values(by=['Regional', 'Qtd_Produtivos'], ascending=[True, False]).reset_index(drop=True)

def get_performance_individual(df: pd.DataFrame, nome_colaborador: str) -> pd.DataFrame:
    """Extrai os dados de desempenho (Qtd_Produtivos) de um colaborador específico por regional."""
    
    df_individual = df[df['Colaborador'] == nome_colaborador.upper().strip()].copy()
    
    if df_individual.empty:
        return pd.DataFrame()
    
    df_individual['Total Visitas'] = df_individual['Qtd_Visitas_Total'].apply(formatar_inteiro)
    df_individual['Concluído OK'] = df_individual['Qtd_Produtivos'].apply(formatar_inteiro)
    df_individual['Produtividade (%)'] = df_individual['% Produtividade'].apply(lambda x: f"{x:.2f}%")
    
    return df_individual

def plot_bar_chart_produtividade_regional(df_data):
    """Cria um gráfico de barras simples de produtivos por Regional, sem título no gráfico e sem título no eixo Y."""
    df_prod_resumo = df_data.groupby('Regional')['Qtd_Produtivos'].sum().reset_index(name='Total_Produtivo')
    
    if df_prod_resumo.empty:
        return px.bar(title="<b>Nenhum dado Produtivo (CONCLUÍDO OK) disponível</b>")
        
    fig = px.bar(
        df_prod_resumo, 
        x='Regional', 
        y='Total_Produtivo', 
        text='Total_Produtivo', 
        title=None, # Título do gráfico removido
        labels={'Regional': 'Regional', 'Total_Produtivo': 'Quantidade de Serviços Produtivos'},
        color='Regional',
        color_discrete_sequence=[config.palette['PRIMARY'], config.palette['ACCENT'], config.palette['SECONDARY_ACCENT']],
        template='plotly_white',
        height=400
    )
    fig.update_traces(
        texttemplate='%{y}',
        textposition='outside',
        hovertemplate="<b>%{x}</b><br>Produtivos: %{y}<extra></extra>",
        marker_line_width=1,
        marker_line_color='white'
    )
    fig.update_layout(
        title_x=0.5, 
        showlegend=False, 
        font_color=config.palette["TEXT_DEFAULT"],
        margin=dict(t=50, b=50, l=50, r=50), 
        yaxis=dict(
            automargin=True,
            title_standoff=30,
            title=None # Título do eixo Y removido
        ),
        xaxis=dict(
            title=None
        )
    )
    return fig

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

    def _define_styles(self):
        styles = getSampleStyleSheet()
        self.styles = {
            'h1': ParagraphStyle('h1', parent=styles['h1'], fontName='Arial-Bold', fontSize=18, leading=22, alignment=TA_CENTER, textColor=HexColor(self.palette["PRIMARY"]), spaceAfter=12),
            'h2': ParagraphStyle('h2', parent=styles['h2'], fontName='Arial-Bold', fontSize=14, leading=16, textColor=HexColor(self.palette["PRIMARY"]), spaceAfter=8),
            'body': ParagraphStyle('body', parent=styles['Normal'], fontName='Arial', fontSize=11, leading=14, textColor=HexColor(self.palette["TEXT_DEFAULT"]), spaceAfter=8),
            'table_header': ParagraphStyle('table_header', parent=styles['Normal'], fontName='Arial-Bold', fontSize=9, leading=11, alignment=TA_CENTER, textColor=HexColor(self.palette["WHITE"])),
            'table_body': ParagraphStyle('table_body', parent=styles['Normal'], fontName='Arial', fontSize=9, leading=11, alignment=TA_LEFT, textColor=HexColor(self.palette["TEXT_DEFAULT"])),
            'table_body_center': ParagraphStyle('table_body_center', parent=styles['Normal'], fontName='Arial', fontSize=9, leading=11, alignment=TA_CENTER, textColor=HexColor(self.palette["TEXT_DEFAULT"])),
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

    def _footer_page(self, canvas, doc):
        canvas.saveState()
        canvas.setFont('Arial-Bold', 8)
        canvas.setFillColor(HexColor(self.palette["GREY_DARK"]))
        canvas.drawCentredString(A4[0]/2.0, 0.5*inch, f"Página {doc.page} de {doc.pages}")
        canvas.restoreState()

    def add_regional_performance_table(self, regional: str, df: pd.DataFrame):
        self.story.append(Paragraph(f"Desempenho: Regional {regional}", self.styles['h2']))
        
        if df.empty:
            self.story.append(Paragraph(f"Nenhum dado de Visitas Produtivas encontrado para a Regional {regional}.", self.styles['body']))
            self.story.append(Spacer(1, 0.1 * inch))
            return

        df_pdf = df.rename(columns={
            'Colaborador': 'Nome Agente', 
            'Qtd_Visitas_Total': 'Total Visitas',
            'Qtd_Produtivos': 'Concluído OK',
            '% Produtividade': 'Produtividade (%)' 
        }).sort_values(by='Concluído OK', ascending=False)
        
        df_pdf['Produtividade (%)'] = df_pdf['Produtividade (%)'].apply(lambda x: f"{x:,.2f}%".replace(",", "X").replace(".", ",").replace("X", "."))
        
        headers = [Paragraph(col, self.styles['table_header']) for col in ['Nome Agente', 'Total Visitas', 'Concluído OK', 'Produtividade (%)']]
        data = [headers]
        
        for _, row in df_pdf.iterrows():
            
            row_data = [
                Paragraph(str(row['Nome Agente']), self.styles['table_body']),
                Paragraph(str(formatar_inteiro(row['Total Visitas'])), self.styles['table_body_center']),
                Paragraph(str(formatar_inteiro(row['Concluído OK'])), self.styles['table_body_center']),
                Paragraph(str(row['Produtividade (%)']), self.styles['table_body_prod'] if float(str(row['Produtividade (%)']).replace('%', '').replace(',', '.')) >= 50 else self.styles['table_body_center'])
            ]
            data.append(row_data)

        num_cols = len(headers)
        available_width = self.doc.width
        col_widths = [available_width * 0.45, available_width * 0.18, available_width * 0.18, available_width * 0.19]
        
        table = Table(data, colWidths=col_widths)
        table_style_list = [
            ('BACKGROUND', (0,0), (-1,0), HexColor(self.palette["PRIMARY"])),
            ('TEXTCOLOR', (0,0), (-1,0), HexColor(self.palette["WHITE"])),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('ALIGN', (1,1), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, HexColor(self.palette["GREY_LIGHT"])),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ('TOPPADDING', (0,0), (-1,-1), 8),
        ]
        for i in range(1, len(data)):
            bg_color = HexColor(self.palette["BACKGROUND_LIGHT"]) if i % 2 == 0 else HexColor(self.palette["WHITE"])
            table_style_list.append(('BACKGROUND', (0,i), (-1,i), bg_color))
        
        table.setStyle(TableStyle(table_style_list))
        self.story.append(table)
        self.story.append(Spacer(1, 0.2 * inch))
        self.story.append(PageBreak())

    def generate_report(self, df_colab_desempenho: pd.DataFrame, colaboradores_nao_encontrados: list):
        self.story.append(Paragraph("Busca Ativa M&F - Relatório de Desempenho de Visitas Produtivas", self.styles['h1']))
        self.story.append(Paragraph(f"Período: {datetime.now().strftime('%d/%m/%Y %H:%M')}", self.styles['body']))
        self.story.append(Spacer(1, 0.2 * inch))
        
        if colaboradores_nao_encontrados:
            qtd_nao_encontrados = len(colaboradores_nao_encontrados)
            nomes_formatados = ', '.join(colaboradores_nao_encontrados)
            self.story.append(Paragraph(f"<b>ATENÇÃO:</b> {qtd_nao_encontrados} colaboradores não foram encontrados na base. Lista: {nomes_formatados}", self.styles['body']))
            self.story.append(Spacer(1, 0.1 * inch))

        regionais = ['SUL', 'NORTE', 'NORDESTE']
        
        for regional in regionais:
            df_regional = df_colab_desempenho[df_colab_desempenho['Regional'] == regional].copy()
            self.add_regional_performance_table(regional, df_regional)

        if self.story and isinstance(self.story[-1], PageBreak):
            self.story.pop()
            
        try:
            self.doc.build(self.story, onFirstPage=self._header_page, onLaterPages=self._header_page)
            return self.buffer.getvalue()
        except Exception:
            return None


# --- 5. Lógica de UI (Camada de Apresentação Principal) ---

# --- CSS CUSTOMIZADO COM PALETA AZUL ---
st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap');
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
    /* --- CORREÇÃO: BORDAS TABELAS E ALINHAMENTO GERAL --- */
    .stDataEditor [data-testid="stDataContainer"] {{
        border: none !important;
    }}
    .stDataEditor [role="row"] {{
        border-bottom: none !important; 
    }}
    .stDataEditor [role="columnheader"] {{
        border-bottom: none !important; 
    }}
    .block-container {{
        padding-top: 2rem; 
        padding-bottom: 2rem; 
        max-width: 1200px; 
        margin: 0 auto; 
    }}
    .main-title-container {{ 
        display: flex; 
        align-items: center; 
        justify-content: center; 
        gap: 15px; 
        margin-bottom: 30px; 
        padding: 20px;
        background-color: var(--white-color);
        border-radius: 15px;
        box-shadow: 0 5px 15px var(--shadow_light-color);
    }}
    .main-title-container h1 {{ 
        margin: 0; 
        line-height: 1.8; /* Aumenta o espaçamento vertical */
        font-size: 2.8em; 
        font-weight: 700; 
        color: var(--primary-color);
    }}
    /* --- TÍTULO DA SEÇÃO PRINCIPAL (CORRIGIDO) --- */
    h2 {{
        color: var(--primary-color);
        font-size: 2.2em; 
        font-weight: 700;
        text-align: center; /* Centralizado */
        margin-top: 30px;
        margin-bottom: 30px;
        width: 100%;
    }}
    h3 {{ 
        color: var(--primary-color); 
        font-size: 1.7em; 
        font-weight: 600; 
        margin-top: 25px; 
        margin-bottom: 15px; 
        padding-bottom: 5px; 
        border-bottom: 2px solid var(--accent-color); 
    }}
    /* --- Estilos para o KPI Lateral (Sumário Produtivo Geral) --- */
    [data-testid="stMetric"] {{
        background-color: var(--white-color); 
        border-radius: 10px; 
        padding: 15px;
        box-shadow: 0 2px 5px var(--shadow-light-color); 
        border-left: 5px solid var(--primary-color); 
        margin-bottom: 15px; 
    }}
    [data-testid="stMetricLabel"] > div {{
        font-weight: 600;
        color: var(--grey-dark-color);
        font-size: 1em;
    }}
    [data-testid="stMetricValue"] {{
        font-size: 2.2em;
        font-weight: 700;
        color: var(--primary-color);
    }}
    .ingrith-box {{
        background-color: var(--white-color);
        border-radius: 10px;
        padding: 20px 10px; 
        box-shadow: 0 2px 8px var(--shadow-light-color);
    }}
    /* --- CENTRALIZAÇÃO DO KPI DE DESTAQUE --- */
    /* Garante que o KPI de destaque (Geral Ingrith) fique centralizado */
    .kpi-center-wrapper {{
        display: flex;
        justify-content: center;
        width: 100%;
    }}

    </style>
""", unsafe_allow_html=True)


# --- UI Principal ---
st.set_page_config(page_title="Busca Ativa M&F", layout="wide", initial_sidebar_state="collapsed")
st.markdown('<div class="main-title-container">', unsafe_allow_html=True)
if config.logo_path.exists():
    try: st.image(str(config.logo_path), width=150)
    except: st.warning("Não foi possível carregar a logo.")
else: st.warning(f"Logo não encontrada em: {config.logo_path}.")
st.markdown('<h1>Busca Ativa M&F</h1>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# --- FILTROS APRIMORADOS (Ano, Mês, Dia) ---
with st.expander("Configurações de Filtro de Dados", expanded=True):
    col_regional, col_municipio = st.columns([2, 3]) 
    col_ano, col_mes, col_dia = st.columns([1, 1, 1])

    with col_regional:
        opcoes_regional_orig = sorted(df_principal['REGIONAL'].unique()) if not df_principal.empty else []
        selecao_regional = st.multiselect("Regional:", options=opcoes_regional_orig, default=opcoes_regional_orig, key="ms_regional")
    
    df_municipios_filtrados = df_principal[df_principal['REGIONAL'].isin(selecao_regional)] if selecao_regional else pd.DataFrame()
    opcoes_municipio_orig = sorted(df_municipios_filtrados['MUNICIPIO'].unique()) if not df_municipios_filtrados.empty else []
    opcoes_municipio_full = ["TODOS"] + opcoes_municipio_orig

    with col_municipio:
        selecao_municipio_raw = st.multiselect("Município:", options=opcoes_municipio_full, default=["TODOS"], key="ms_municipio")
        
        if "TODOS" in selecao_municipio_raw:
            selecao_municipio = opcoes_municipio_orig
        else:
            selecao_municipio = selecao_municipio_raw

    opcoes_ano_orig = sorted(df_principal['Ano'].unique()) if 'Ano' in df_principal.columns else []
    opcoes_ano_full = ["TODOS"] + opcoes_ano_orig
    with col_ano:
        selecao_ano_raw = st.multiselect("Ano:", options=opcoes_ano_full, default=["TODOS"], key="ms_ano")
        if "TODOS" in selecao_ano_raw:
            selecao_ano = opcoes_ano_orig
        else:
            selecao_ano = [int(a) for a in selecao_ano_raw]

    df_meses_filtrados = df_principal[df_principal['Ano'].isin(selecao_ano)] if selecao_ano else pd.DataFrame()
    opcoes_mes_orig = sorted(df_meses_filtrados['Mes'].unique()) if 'Mes' in df_meses_filtrados.columns else []
    opcoes_mes_full = ["TODOS"] + opcoes_mes_orig

    with col_mes:
        selecao_mes_raw = st.multiselect("Mês:", options=opcoes_mes_full, default=["TODOS"], key="ms_mes")
        if "TODOS" in selecao_mes_raw:
            selecao_mes = opcoes_mes_orig
        else:
            selecao_mes = [int(m) for m in selecao_mes_raw]
        
    df_dias_filtrados = df_meses_filtrados[df_meses_filtrados['Mes'].isin(selecao_mes)] if selecao_mes else pd.DataFrame()
    opcoes_dia_orig = sorted(df_dias_filtrados['Dia'].unique()) if 'Dia' in df_dias_filtrados.columns else []
    opcoes_dia_full = ["TODOS"] + opcoes_dia_orig
    
    with col_dia:
        selecao_dia_raw = st.multiselect("Dia:", options=opcoes_dia_full, default=["TODOS"], key="ms_dia")
        if "TODOS" in selecao_dia_raw:
            selecao_dia = opcoes_dia_orig
        else:
            selecao_dia = [int(d) for d in selecao_dia_raw]


if not df_principal.empty and selecao_regional and selecao_municipio and selecao_ano and selecao_mes and selecao_dia:
    df_base_total = df_principal[
        (df_principal['REGIONAL'].isin(selecao_regional)) &
        (df_principal['MUNICIPIO'].isin(selecao_municipio)) &
        (df_principal['Ano'].isin(selecao_ano)) &
        (df_principal['Mes'].isin(selecao_mes)) &
        (df_principal['Dia'].isin(selecao_dia))
    ].copy()
    
    if df_base_total.empty:
        st.info("Nenhum dado encontrado para a combinação de filtros selecionada.")
        st.stop()
        
    df_para_analise = df_base_total[df_base_total[config.coluna_colaborador].isin(
        [c.upper().strip() for c in config.colaboradores_list]
    )].copy()

    # --- LÓGICA PRINCIPAL ---
    
    df_desempenho = get_desempenho_visitas_por_regional(df_para_analise, config.colaboradores_list)
    kpis = calcular_indicadores_totais(df_base_total, df_para_analise, config.colaboradores_list)
    colaboradores_nao_encontrados = kpis['colaboradores_nao_encontrados']
    
    # Aviso Otimizado
    if colaboradores_nao_encontrados:
        qtd_nao_encontrados = len(colaboradores_nao_encontrados)
        nomes_formatados = ', '.join(colaboradores_nao_encontrados)
        
        mensagem_principal = f"⚠️ **ATENÇÃO:** **{qtd_nao_encontrados}** colaboradores da sua lista não foram encontrados na base de dados."
        
        with st.expander(mensagem_principal, expanded=False):
            st.markdown("---")
            st.markdown(f"**Lista de Nomes Não Encontrados:**")
            st.markdown(f"<p style='font-size: 14px; color: {config.palette['GREY_DARK']};'>{nomes_formatados}</p>", unsafe_allow_html=True)
    
    st.markdown("---") 

    # -------------------------------------------------------------------
    # --- DESTAQUE INGRID LORENA (CORREÇÃO DE ALINHAMENTO) ---
    # -------------------------------------------------------------------
    df_ingrith = get_performance_individual(df_desempenho, config.colaboradora_destaque)

    if not df_ingrith.empty:
        total_produtivo_ingrith = df_ingrith['Qtd_Produtivos'].sum()
        total_visitas_ingrith = df_ingrith['Qtd_Visitas_Total'].sum()
        media_produtividade = (total_produtivo_ingrith / total_visitas_ingrith) * 100 if total_visitas_ingrith > 0 else 0
        
        # Nome Centralizado
        st.markdown(f"<div style='text-align: center;'><h2>👑 Destaque: {config.colaboradora_destaque.title()}</h2></div>", unsafe_allow_html=True)
        
        # KPI GERAL (Topo) - Centralizado
        st.markdown('<div class="kpi-center-wrapper">', unsafe_allow_html=True)
        
        # Usando colunas vazias para centralizar o KPI que tem largura fixa
        col_left_spacer, col_kpi_ingrith, col_right_spacer = st.columns([1.5, 1, 1.5]) 
        
        with col_kpi_ingrith:
            # Replicando a estrutura do KPI lateral para o destaque
            st.markdown(f"""
                <div data-testid="stMetric">
                    <div data-testid="stMetricLabel" style="font-size: 1em;">Total de Produtivos (Todas Regionais)</div>
                    <div data-testid="stMetricValue">{formatar_inteiro(total_produtivo_ingrith)}</div>
                    <div style="font-size: 0.9em; font-weight: 500; color: {config.palette['SUCCESS']}; margin-top: 5px;">
                        <span style="font-size: 1.1em; margin-right: 5px;">↑</span> Prod. Média Geral: {media_produtividade:.2f}%
                    </div>
                </div>
            """, unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("### Resultados por Regional")

        # Resultados Regionais da Ingrith
        with st.container():
            st.markdown('<div class="ingrith-box">', unsafe_allow_html=True)
            col_reg_sul, col_reg_norte, col_reg_nordeste = st.columns(3)
            regionais = ['SUL', 'NORTE', 'NORDESTE']
            cols = [col_reg_sul, col_reg_norte, col_reg_nordeste]

            for regional, col in zip(regionais, cols):
                df_reg = df_ingrith[df_ingrith['Regional'] == regional]
                
                with col:
                    if not df_reg.empty:
                        concluido = df_reg['Concluído OK'].iloc[0]
                        total_vis = df_reg['Total Visitas'].iloc[0]
                        prod_perc = df_reg['Produtividade (%)'].iloc[0]
                        
                        st.markdown(f"**{regional}**", unsafe_allow_html=True)
                        st.metric(
                            f"Concluído OK ({regional})",
                            concluido,
                            delta=f"Total Visitas: {total_vis} | Prod. Regional: {prod_perc}"
                        )
                    else:
                        st.markdown(f"**{regional}**", unsafe_allow_html=True)
                        st.metric(f"Concluído OK ({regional})", "0", delta="Sem registros.", delta_color="off")

            st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown("---")
    # -------------------------------------------------------------------

    # 3. Gráfico de Barras de Resumo Regional (Produtivos)
    # CORREÇÃO 4: Título da seção principal
    st.markdown("## RESUMO DA PRODUTIVIDADE")
    
    with st.container(): 
        # Vamos usar colunas para alinhar o Sumário ao centro da tela
        col_left_spacer, col_kpi_content, col_right_spacer = st.columns([1, 2, 1])
        
        with col_kpi_content:
            # O Sumário Produtivo Geral agora ocupa a largura da coluna central.
            st.markdown("") 
            total_produtivo_geral = df_desempenho['Qtd_Produtivos'].sum()
            df_prod_resumo = df_desempenho.groupby('Regional')['Qtd_Produtivos'].sum().reset_index(name='Total_Produtivo')
            df_prod_resumo['Regional'] = pd.Categorical(df_prod_resumo['Regional'], categories=['SUL', 'NORTE', 'NORDESTE'], ordered=True)
            df_prod_resumo.sort_values('Regional', inplace=True)
            
            # --- RESTAURAÇÃO DO VISUAL ELEGANTE KPI LATERAL ---
            for index, row in df_prod_resumo.iterrows():
                total_visitas_regional = df_desempenho[df_desempenho['Regional'] == row['Regional']]['Qtd_Visitas_Total'].sum()
                prod_media = (row['Total_Produtivo'] / total_visitas_regional) * 100 if total_visitas_regional > 0 else 0
                
                st.metric(
                    f"**{row['Regional']}**",
                    f"{formatar_inteiro(row['Total_Produtivo'])}",
                    delta=f"Prod. Média: {prod_media:.2f}%",
                    delta_color="off" 
                )
            # --- FIM RESTAURAÇÃO KPI LATERAL ---

    st.markdown("---")

    # 4. Tabelas separadas por regional (AGORA EMPILHADAS)
    st.markdown("## 🎯 Desempenho Individual por Colaborador")
    
    regionais_selecionadas = sorted([r for r in ['SUL', 'NORTE', 'NORDESTE'] if r in selecao_regional])
    
    with st.container():
        primary_color = config.palette["PRIMARY"]
        
        for regional in regionais_selecionadas:
            df_regional = df_desempenho[df_desempenho['Regional'] == regional].copy()
            
            if df_regional.empty:
                st.markdown(f"📍{regional}")
                st.info(f"Nenhum colaborador com dados de visitas na Regional {regional}.")
                continue

            df_regional = df_regional.sort_values(by='Qtd_Produtivos', ascending=False).reset_index(drop=True)
            df_regional['Posição'] = df_regional.index + 1
            
            df_regional['Colaborador'] = df_regional['Colaborador'].apply(
                lambda x: "👑 " + x if x == config.colaboradora_destaque.upper() else x
            )
            
            df_final = df_regional[['Posição', 'Colaborador', 'Qtd_Visitas_Total', 'Qtd_Produtivos', '% Produtividade']].rename(columns={
                'Qtd_Visitas_Total': 'Total Visitas',
                'Qtd_Produtivos': 'Concluído OK',
                '% Produtividade': 'Produtividade (%)'
            })
            
            col_config = {
                "Posição": st.column_config.NumberColumn("Pos.", width="small"),
                "Colaborador": st.column_config.TextColumn("Colaborador", width="large"),
                "Total Visitas": st.column_config.NumberColumn("Total Visitas", format="%d", help="Total de serviços executados em campo"),
                "Concluído OK": st.column_config.NumberColumn("Concluído OK", format="%d", help="Total de visitas produtivas"),
                "Produtividade (%)": st.column_config.ProgressColumn(
                    "Produtividade (%)",
                    format="%.2f %%",
                    min_value=0,
                    max_value=100,
                    color=primary_color 
                ),
            }

            st.markdown(f"📍{regional}")
            
            st.data_editor(
                df_final,
                column_config=col_config,
                use_container_width=True,
                hide_index=True,
                disabled=df_final.columns 
            )
            st.markdown("---") 

    # 5. Lógica de Download do PDF 
    st.markdown("### 📥 Baixar Relatório")
    
    buffer_pdf = BytesIO()
    pdf_data = RelatorioVisualPDF(config.logo_path, config.palette, buffer_pdf).generate_report(
        df_desempenho, colaboradores_nao_encontrados
    )
    
    if pdf_data:
        st.download_button(
            label="✅ Gerar e Baixar Relatório em PDF (Desempenho Profissional)",
            data=pdf_data,
            file_name=f"Relatorio_Produtivo_MF_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
            mime="application/pdf"
        )
    else:
        st.error("Não foi possível gerar o relatório PDF.")

st.markdown("---")
st.markdown(f"<p style='text-align:center; font-size:14px; color:{config.palette['GREY_DARK']};'>Criado por PLINIO M. RODRIGUES. &copy; {datetime.now().year}</p>", unsafe_allow_html=True)