import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from pathlib import Path
from io import BytesIO

# Imports para geração de PDF (ReportLab)
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor, black, white, lightgrey, grey
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT

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

        # Paleta de cores padronizada em verde
        self.palette = {
            "PRIMARY": "#2e7d32",
            "ACCENT": "#4caf50",
            "SECONDARY_ACCENT": "#8bc34a",
            "BACKGROUND_LIGHT": "#f0f8f4",
            "TEXT_DEFAULT": "#212529",
            "GREY_LIGHT": "#e0e0e0",
            "GREY_DARK": "#757575",
            "WHITE": "#FFFFFF",
            "SHADOW_LIGHT": "rgba(0,0,0,0.08)",
            "SUCCESS": "#2e7d32",
            "WARNING": "#fdd835",
            "DANGER": "#d32f2f"
        }
        
        self.servicos = {
            'executados': ['CONCLUIDO OK', 'DESCARREGADO COM IMPEDIMENTO', 'DESCARREGADO SEM IMPEDIMENTO', 'IMPROCEDENTE'],
            'em_campo': ['ALVO EM CAMPO'],
            'a_atribuir': ['ALVO NAO ATRIBUIDO'],
            'pendentes': ['ALVO ENVIADO - NAO RECEBIDO'],
            'produtivos': ['CONCLUIDO OK', 'DESCARREGADO SEM IMPEDIMENTO'],
            'improdutivos': ['DESCARREGADO COM IMPEDIMENTO', 'IMPROCEDENTE']
        }

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
            'CLAUDINEIA MIRANDA SOUZA', 'LHORRAN FHILLYPHE TAVARES NOGUEIRA'
        ]

config = Config()

# --- Funções de Formatação e Utilitários ---
def formatar_inteiro(valor: float | int) -> str:
    if pd.isna(valor) or valor is None: return "0"
    try: valor = int(valor)
    except (ValueError, TypeError): return "Inválido"
    return f"{valor:,}".replace(",", ".")

def get_status_kpi_color(kpi_value, threshold, inverse=False):
    if inverse:
        return config.palette["DANGER"] if kpi_value > threshold else config.palette["SUCCESS"]
    else:
        return config.palette["SUCCESS"] if kpi_value > threshold else config.palette["DANGER"]

# --- 2. Camada de Acesso e Processamento de Dados (Data Access Layer) ---
@st.cache_data(ttl=3600, show_spinner="Carregando e processando dados de Busca Ativa...")
def carregar_e_processar_dados(caminho_arquivo: Path) -> pd.DataFrame:
    if not caminho_arquivo.exists():
        st.error(f"Erro Crítico: Arquivo de dados não encontrado em '{caminho_arquivo}'.")
        st.stop()
    
    try:
        df = pd.read_excel(caminho_arquivo, sheet_name=config.excel_sheet)
        
        df.columns = df.columns.str.strip().str.upper().str.replace(' ', '_').str.replace('[^A-Z0-9_]', '', regex=True)
        
        required_cols = ['REGIONAL', 'MUNICIPIO', 'NOME_FASE', 'ALVO_CONDICAO_OBJETIVA', config.coluna_colaborador]
        if not all(col in df.columns for col in required_cols):
            missing_cols = [col for col in required_cols if col not in df.columns]
            st.error(f"Colunas essenciais faltando na planilha: {missing_cols}")
            st.stop()
            
        df = df[df['ALVO_CONDICAO_OBJETIVA'].str.upper().str.strip() == 'SIM'].copy()
        
        if df.empty:
            st.warning("Nenhum dado foi encontrado após a aplicação dos filtros iniciais. Verifique a planilha.")
            st.stop()
            
        df['NOME_FASE'] = df['NOME_FASE'].str.upper().str.strip()
        df['REGIONAL'] = df['REGIONAL'].str.upper().str.strip()
        df['MUNICIPIO'] = df['MUNICIPIO'].str.upper().str.strip()
        df[config.coluna_colaborador] = df[config.coluna_colaborador].str.upper().str.strip()

        regionais_validas = ['NORTE', 'NORDESTE', 'SUL']
        df = df[df['REGIONAL'].isin(regionais_validas)].copy()

        return df
    
    except Exception as e:
        st.error(f"Erro fatal ao carregar o arquivo Excel: {e}")
        st.exception(e)
        st.stop()

df_principal = carregar_e_processar_dados(config.base_dir / config.excel_file)


# --- 3. Camada de Lógica de Negócio e Agregação (Service Layer) ---
def calcular_indicadores_totais(df_base_total: pd.DataFrame, df_para_analise: pd.DataFrame, colaboradores_list: list) -> dict:
    if df_base_total.empty:
        return {
            "total": 0, "executados_totais": 0, "executados_produtivos": 0, "executados_improdutivos": 0,
            "em_campo": 0, "a_atribuir": 0, "pendentes": 0, "colaboradores_nao_encontrados": []
        }
        
    qtd_total_servicos_base = len(df_base_total)
    qtd_a_atribuir_base = df_base_total['NOME_FASE'].isin(config.servicos['a_atribuir']).sum()

    if df_para_analise.empty:
        qtd_executados_totais = 0
        qtd_executados_produtivos = 0
        qtd_executados_improdutivos = 0
        qtd_em_campo_colab = 0
        qtd_pendentes_colab = 0
    else:
        qtd_executados_totais = df_para_analise['NOME_FASE'].isin(config.servicos['executados']).sum()
        qtd_executados_produtivos = df_para_analise['NOME_FASE'].isin(config.servicos['produtivos']).sum()
        qtd_executados_improdutivos = df_para_analise['NOME_FASE'].isin(config.servicos['improdutivos']).sum()
        qtd_em_campo_colab = df_para_analise['NOME_FASE'].isin(config.servicos['em_campo']).sum()
        qtd_pendentes_colab = df_para_analise['NOME_FASE'].isin(config.servicos['pendentes']).sum()

    colaboradores_na_base = set(df_base_total[config.coluna_colaborador].unique())
    colaboradores_nao_encontrados = [c for c in colaboradores_list if c.upper().strip() not in colaboradores_na_base]
    
    return {
        "total": qtd_total_servicos_base,
        "executados_totais": qtd_executados_totais,
        "executados_produtivos": qtd_executados_produtivos,
        "executados_improdutivos": qtd_executados_improdutivos,
        "em_campo": qtd_em_campo_colab,
        "a_atribuir": qtd_a_atribuir_base,
        "pendentes": qtd_pendentes_colab,
        "colaboradores_nao_encontrados": colaboradores_nao_encontrados
    }

def agregar_por_dimensao(df: pd.DataFrame, coluna_agregacao: str, servico_type: str) -> pd.DataFrame:
    if df.empty or coluna_agregacao not in df.columns:
        return pd.DataFrame(columns=['Dimensão', 'Métrica'])
    
    if servico_type in config.servicos:
        df_agregado = df.groupby(coluna_agregacao)['NOME_FASE'].apply(
            lambda x: x.isin(config.servicos[servico_type]).sum()
        ).reset_index()
    else:
        df_agregado = df.groupby(coluna_agregacao)['NOME_FASE'].count().reset_index()
        
    df_agregado.columns = ['Dimensão', 'Métrica']
    return df_agregado.sort_values(by='Métrica', ascending=False)

def agregar_desempenho_colaborador(df: pd.DataFrame, colaboradores_list: list) -> pd.DataFrame:
    df_filtrado = df[df[config.coluna_colaborador].isin([c.upper().strip() for c in colaboradores_list])].copy()
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

    return df_agregado.sort_values(by=config.coluna_colaborador)

def plot_bar_chart(df_data, x_col, y_col, title, x_label, y_label, color_discrete_sequence=None):
    fig = px.bar(
        df_data,
        x=x_col,
        y=y_col,
        title=title,
        labels={x_col: x_label, y_col: y_col},
        color_discrete_sequence=color_discrete_sequence,
        template='simple_white',
        height=500
    )
    fig.update_traces(
        texttemplate='%{y}',
        textposition='outside'
    )
    fig.update_layout(
        xaxis_tickangle=-45,
        title_font_color=config.palette["PRIMARY"],
        font_color=config.palette["TEXT_DEFAULT"],
        margin=dict(l=50, r=20, t=80, b=20),
        title_x=0.5,
        title_y=0.9
    )
    return fig


# --- 4. Camada de Geração de Relatórios (Reporting Layer - PDF) ---
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
            'h1': ParagraphStyle('h1', parent=styles['h1'], fontName='Arial-Bold', fontSize=18, leading=22, alignment=TA_CENTER, textColor=HexColor(self.palette["PRIMARY"]), spaceAfter=12),
            'h2': ParagraphStyle('h2', parent=styles['h2'], fontName='Arial-Bold', fontSize=14, leading=16, textColor=HexColor(self.palette["PRIMARY"]), spaceAfter=8),
            'body': ParagraphStyle('body', parent=styles['Normal'], fontName='Arial', fontSize=10, leading=12, textColor=HexColor(self.palette["TEXT_DEFAULT"]), spaceAfter=6),
            'kpi_label': ParagraphStyle('kpi_label', parent=styles['Normal'], fontName='Arial-Bold', fontSize=10, leading=12, textColor=HexColor(self.palette["GREY_DARK"])),
            'kpi_value': ParagraphStyle('kpi_value', parent=styles['Normal'], fontName='Arial-Bold', fontSize=14, leading=16, textColor=HexColor(self.palette["PRIMARY"])),
            'table_header': ParagraphStyle('table_header', parent=styles['Normal'], fontName='Arial-Bold', fontSize=8, leading=10, alignment=TA_CENTER, textColor=HexColor(self.palette["WHITE"])),
            'table_body': ParagraphStyle('table_body', parent=styles['Normal'], fontName='Arial', fontSize=7, leading=9, alignment=TA_LEFT, textColor=HexColor(self.palette["TEXT_DEFAULT"])),
            'table_body_center': ParagraphStyle('table_body_center', parent=styles['Normal'], fontName='Arial', fontSize=7, leading=9, alignment=TA_CENTER, textColor=HexColor(self.palette["TEXT_DEFAULT"])),
            'footer': ParagraphStyle('footer', parent=styles['Normal'], fontName='Arial-Italic', fontSize=8, leading=10, alignment=TA_CENTER, textColor=HexColor(self.palette["GREY_DARK"])),
        }

    def _header_page(self, canvas, doc):
        canvas.saveState()
        if self.logo_path.exists():
            logo_width = 0.5 * inch
            logo_height = 0.5 * inch
            try:
                canvas.drawImage(
                    str(self.logo_path),
                    A4[0] - 0.75*inch - logo_width,
                    A4[1] - 0.7*inch,
                    width=logo_width,
                    height=logo_height,
                    mask='auto'
                )
            except Exception:
                pass
        canvas.setFont('Arial-Bold', 14)
        canvas.setFillColor(HexColor(self.palette["PRIMARY"]))
        canvas.drawCentredString(A4[0]/2.0, A4[1] - 0.4*inch, "Relatório de Performance Busca Ativa")
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

    def add_kpi_summary(self, kpis: dict):
        self.story.append(Paragraph("Resumo da Base de Dados", self.styles['h2']))
        
        kpi_data = [
            [Paragraph(f"<b>Total de Serviços:</b>", self.styles['kpi_label']), Paragraph(formatar_inteiro(kpis['total']), self.styles['kpi_value'])],
            [Paragraph(f"<b>Serviços Executados:</b>", self.styles['kpi_label']), Paragraph(formatar_inteiro(kpis['executados_totais']), self.styles['kpi_value'])],
            [Paragraph(f"<b>Serviços Produtivos:</b>", self.styles['kpi_label']), Paragraph(formatar_inteiro(kpis['executados_produtivos']), self.styles['kpi_value'])],
            [Paragraph(f"<b>Serviços Improdutivos:</b>", self.styles['kpi_label']), Paragraph(formatar_inteiro(kpis['executados_improdutivos']), self.styles['kpi_value'])],
            [Paragraph(f"<b>Serviços em Campo:</b>", self.styles['kpi_label']), Paragraph(formatar_inteiro(kpis['em_campo']), self.styles['kpi_value'])],
            [Paragraph(f"<b>Serviços a Atribuir:</b>", self.styles['kpi_label']), Paragraph(formatar_inteiro(kpis['a_atribuir']), self.styles['kpi_value'])],
            [Paragraph(f"<b>Serviços Pendentes:</b>", self.styles['kpi_label']), Paragraph(formatar_inteiro(kpis['pendentes']), self.styles['kpi_value'])],
        ]
        kpi_table = Table(kpi_data, colWidths=[self.doc.width / 2.0, self.doc.width / 2.0])
        kpi_table.setStyle(TableStyle([
            ('ALIGN', (0,0), (0,-1), 'LEFT'),
            ('ALIGN', (1,0), (1,-1), 'LEFT'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        self.story.append(kpi_table)
        self.story.append(Spacer(1, 0.1 * inch))

    def add_dataframe_to_pdf(self, title: str, df: pd.DataFrame):
        self.story.append(Paragraph(title, self.styles['h2']))
        
        if df.empty:
            self.story.append(Paragraph("Nenhum dado disponível para esta tabela.", self.styles['body']))
            self.story.append(Spacer(1, 0.1 * inch))
            return

        headers = [Paragraph(col.replace('_', ' '), self.styles['table_header']) for col in df.columns]
        data = [headers]
        for _, row in df.iterrows():
            row_data = [
                Paragraph(str(formatar_inteiro(item)), self.styles['table_body_center']) if isinstance(item, (int, float)) else Paragraph(str(item), self.styles['table_body'])
                for item in row
            ]
            data.append(row_data)

        num_cols = len(df.columns)
        available_width = self.doc.width
        col_widths = [available_width / num_cols] * num_cols
        
        table = Table(data, colWidths=col_widths)
        table_style_list = [
            ('BACKGROUND', (0,0), (-1,0), HexColor(self.palette["PRIMARY"])),
            ('TEXTCOLOR', (0,0), (-1,0), HexColor(self.palette["WHITE"])),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, HexColor(self.palette["GREY_LIGHT"])),
        ]
        for i in range(1, len(data)):
            bg_color = HexColor(self.palette["BACKGROUND_LIGHT"]) if i % 2 == 0 else HexColor(self.palette["WHITE"])
            table_style_list.append(('BACKGROUND', (0,i), (-1,i), bg_color))
        
        table.setStyle(TableStyle(table_style_list))
        self.story.append(table)
        self.story.append(Spacer(1, 0.1 * inch))

    def generate_report(self, df_base_total: pd.DataFrame, df_para_analise: pd.DataFrame, df_colab_performance: pd.DataFrame, colaboradores_nao_encontrados: list):
        kpis = calcular_indicadores_totais(df_base_total, df_para_analise, config.colaboradores_list)
        
        self.story.append(Paragraph("Resumo de Performance", self.styles['h1']))
        self.add_kpi_summary(kpis)
        
        self.story.append(Paragraph("Análise por Regional e Município", self.styles['h1']))
        df_regional_analise = agregar_por_dimensao(df_para_analise, 'REGIONAL', 'produtivos')
        self.add_dataframe_to_pdf("Serviços Produtivos por Regional", df_regional_analise)

        df_municipio_analise = agregar_por_dimensao(df_para_analise, 'MUNICIPIO', 'improdutivos')
        self.add_dataframe_to_pdf("Serviços Improdutivos por Município", df_municipio_analise)

        if self.story and not df_colab_performance.empty:
            self.story.append(PageBreak())
        
        self.story.append(Paragraph("Desempenho dos seus Colaboradores", self.styles['h1']))
        
        if colaboradores_nao_encontrados:
            self.story.append(Paragraph(f"<b>Atenção:</b> Os seguintes colaboradores não foram encontrados na base de dados: {', '.join(colaboradores_nao_encontrados)}", self.styles['body']))
            self.story.append(Spacer(1, 0.1 * inch))
        
        self.add_dataframe_to_pdf("Tabela de Desempenho Individual Completa", df_colab_performance)

        try:
            self.doc.build(self.story, onFirstPage=self._header_page, onLaterPages=self._header_page)
            return self.buffer.getvalue()
        except Exception as e:
            st.error(f"Erro ao construir o PDF: {e}")
            return None


# --- 5. Lógica de UI (Camada de Apresentação Principal) ---
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
    .main-title-container {{ 
        display: flex; 
        align-items: center; 
        justify-content: center; 
        gap: 15px; 
        margin-bottom: 30px; 
        padding: 20px;
        background-color: var(--white-color);
        border-radius: 15px;
        box-shadow: 0 5px 15px var(--shadow-light-color);
    }}
    .main-title-container h1 {{ 
        margin: 0; 
        line-height: 1.2; 
        font-size: 2.8em; 
        font-weight: 700; 
        color: var(--primary-color);
    }}
    [data-testid="stMetric"] {{ 
        background-color: var(--white-color); 
        border-radius: 12px; 
        padding: 20px 25px; 
        box-shadow: 0 4px 10px var(--shadow-light-color); 
        text-align: center; 
        border: 1px solid var(--grey-light-color); 
        transition: transform 0.2s ease-in-out, box-shadow 0.2s ease-in-out; 
        margin-bottom: 15px; 
    }}
    [data-testid="stMetric"]:hover {{ 
        transform: translateY(-3px); 
        box-shadow: 0 6px 15px var(--shadow_light-color); 
    }}
    [data-testid="stMetricValue"] {{ 
        font-size: 2.2em; 
        font-weight: bold; 
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
        color: var(--primary-color); 
        font-size: 1.7em; 
        font-weight: 600; 
        margin-top: 25px; 
        margin-bottom: 15px; 
        padding-bottom: 5px; 
        border-bottom: 2px solid var(--accent-color); 
    }}
    .stDataFrame {{ 
        border-radius: 10px; 
        overflow: hidden; 
        box-shadow: 0 2px 8px var(--shadow_light-color); 
        border: 1px solid var(--grey-light-color); 
    }}
    .stTabs [data-baseweb="tab-list"] {{ 
        gap: 12px; 
        justify-content: center; 
        margin-bottom: 25px; 
        margin-top: 20px; 
    }}
    .stTabs [data-baseweb="tab"] {{ 
        height: 45px; 
        padding: 0 25px; 
        background-color: var(--bg-light-color); 
        border-radius: 10px 10px 0 0; 
        border: 1px solid var(--grey-light-color); 
        font-weight: 600; 
        color: var(--text-default-color); 
        transition: all 0.2s ease-in-out; 
        font-size: 1.05em; 
    }}
    .stTabs [data-baseweb="tab"]:hover {{ 
        background-color: var(--accent-color); 
        color: var(--white-color); 
        border-color: var(--accent-color); 
    }}
    .stTabs [data-baseweb="tab"][aria-selected="true"] {{ 
        background-color: var(--primary-color); 
        color: var(--white-color); 
        border-top: 4px solid var(--secondary-accent-color); 
        border-color: var(--primary-color); 
        transform: translateY(-3px); 
        box-shadow: 0 4px 8px rgba(0,0,0,0.1); 
    }}
    .stDownloadButton > button {{
        background-color: var(--success-color);
        color: var(--white-color);
        border: none;
        padding: 10px 20px;
        border-radius: 8px;
        font-weight: bold;
        transition: background-color 0.3s ease;
    }}
    .stDownloadButton > button:hover {{
        background-color: #218838;
    }}
    </style>
""", unsafe_allow_html=True)


# --- UI Principal ---
st.set_page_config(page_title="Painel de Performance Busca Ativa", layout="wide", initial_sidebar_state="collapsed")
st.markdown('<div class="main-title-container">', unsafe_allow_html=True)
if config.logo_path.exists():
    try: st.image(str(config.logo_path), width=150)
    except: st.warning("Não foi possível carregar a logo.")
else: st.warning(f"Logo não encontrada em: {config.logo_path}.")
st.markdown('<h1>Painel de Performance Busca Ativa</h1>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

with st.expander("Configurações de Filtro", expanded=True):
    col_regional, col_municipio = st.columns(2)
    with col_regional:
        opcoes_regional = sorted(df_principal['REGIONAL'].unique()) if not df_principal.empty else []
        selecao_regional = st.multiselect("Selecione Regional:", options=opcoes_regional, default=opcoes_regional, key="ms_regional")
    
    with col_municipio:
        df_municipios_filtrados = df_principal[df_principal['REGIONAL'].isin(selecao_regional)] if selecao_regional else pd.DataFrame()
        opcoes_municipio = sorted(df_municipios_filtrados['MUNICIPIO'].unique()) if not df_municipios_filtrados.empty else []
        selecao_municipio = st.multiselect("Selecione Município:", options=opcoes_municipio, default=opcoes_municipio, key="ms_municipio")

if not df_principal.empty and selecao_regional and selecao_municipio:
    df_base_total = df_principal[
        (df_principal['REGIONAL'].isin(selecao_regional)) &
        (df_principal['MUNICIPIO'].isin(selecao_municipio))
    ].copy()
    
    if df_base_total.empty:
        st.info("Nenhum dado encontrado para a combinação de filtros selecionada.")
        st.stop()
        
    df_para_analise = df_base_total[df_base_total[config.coluna_colaborador].isin(
        [c.upper().strip() for c in config.colaboradores_list]
    )].copy()

    kpis = calcular_indicadores_totais(df_base_total, df_para_analise, config.colaboradores_list)
    df_colab_performance = agregar_desempenho_colaborador(df_base_total, config.colaboradores_list)

    tab_base, tab_colaboradores = st.tabs(["📊 Análise da Base", "👥 Desempenho por Colaborador"])

    with tab_base:
        st.markdown("### KPIs da Base de Dados")
        col1_base, col2_base_container, col3_base, col4_base, col5_base = st.columns([1, 1, 1, 1, 1])
        
        with col1_base: st.metric("📋 Total de Serviços", formatar_inteiro(kpis['total']))
        
        with col2_base_container:
            st.markdown(f"""
                <div data-testid="stMetric">
                    <div data-testid="stMetricLabel" style="display: flex; align-items: center; justify-content: flex-start; gap: 5px;">
                        <span style="font-size: 1.1em;">✅</span>
                        <div style="font-size: 1.1em; font-weight: 500; color: var(--grey-dark-color);">Executados</div>
                    </div>
                    <div data-testid="stMetricValue" style="font-size: 2.2em; font-weight: bold; color: var(--primary-color);">
                        {formatar_inteiro(kpis['executados_totais'])}
                    </div>
                    <div style="margin-top: 15px; border-top: 1px solid var(--grey-light-color); padding-top: 10px;">
                        <div style="display: flex; justify-content: space-between; font-size: 14px; color: #666; margin-bottom: 5px;">
                            <span>Produtivos</span>
                            <span style="font-weight: bold; color: var(--text-default-color);">{formatar_inteiro(kpis['executados_produtivos'])}</span>
                        </div>
                        <div style="display: flex; justify-content: space-between; font-size: 14px; color: #666;">
                            <span>Improdutivos</span>
                            <span style="font-weight: bold; color: var(--text-default-color);">{formatar_inteiro(kpis['executados_improdutivos'])}</span>
                        </div>
                    </div>
                </div>
            """, unsafe_allow_html=True)

        with col3_base: st.metric("🛠️ Em Campo", formatar_inteiro(kpis['em_campo']))
        with col4_base: st.metric("🆕 A Atribuir", formatar_inteiro(kpis['a_atribuir']))
        with col5_base: st.metric("📤 Pendentes", formatar_inteiro(kpis['pendentes']))
        
        if kpis['pendentes'] > 0:
            st.warning(f"⚠️ **Atenção!** Existem **{formatar_inteiro(kpis['pendentes'])}** serviços pendentes na base de dados para seus colaboradores.")

        st.markdown("---")
        st.markdown("### Análise de Serviços")
        
        col_grafico_executados, col_grafico_atribuir = st.columns(2)
        with col_grafico_executados:
            st.markdown("#### Serviços Executados")
            
            selecao_visualizacao_executados = st.radio(
                "Visualização do Gráfico:",
                ["Produtivos e Improdutivos", "Total de Executados"],
                key="radio_visao_executados",
                horizontal=True
            )

            visao_dimensao_executados = st.radio(
                "Agrupar por:", 
                ["Regional", "Município"], 
                key="radio_dimensao_executados", 
                horizontal=True
            )
            
            coluna_agregacao = 'REGIONAL' if visao_dimensao_executados == "Regional" else 'MUNICIPIO'

            if selecao_visualizacao_executados == "Total de Executados":
                df_agregado = agregar_por_dimensao(df_para_analise, coluna_agregacao, 'executados')
                if not df_agregado.empty:
                    fig = plot_bar_chart(
                        df_agregado, 
                        'Dimensão', 
                        'Métrica', 
                        f"Total de Serviços Executados por {visao_dimensao_executados}", 
                        visao_dimensao_executados, 
                        'Quantidade', 
                        color_discrete_sequence=[config.palette['SUCCESS']]
                    )
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("Nenhum dado de 'Total de Executados' disponível para a seleção.")
            else: # "Produtivos e Improdutivos"
                df_produtivos = agregar_por_dimensao(df_para_analise, coluna_agregacao, 'produtivos')
                df_improdutivos = agregar_por_dimensao(df_para_analise, coluna_agregacao, 'improdutivos')
                
                if not df_produtivos.empty or not df_improdutivos.empty:
                    df_plot_prod_improd = pd.concat([df_produtivos.assign(Tipo='Produtivo'), df_improdutivos.assign(Tipo='Improdutivo')])
                    fig = px.bar(
                        df_plot_prod_improd, 
                        x='Dimensão', 
                        y='Métrica', 
                        color='Tipo',
                        title=f"Serviços Produtivos e Improdutivos por {visao_dimensao_executados}",
                        labels={'Dimensão': visao_dimensao_executados, 'Métrica': 'Quantidade'},
                        color_discrete_map={'Produtivo': config.palette['SUCCESS'], 'Improdutivo': config.palette['DANGER']},
                        height=500
                    )
                    fig.update_layout(xaxis_tickangle=-45)
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("Nenhum dado de 'Produtivos e Improdutivos' disponível para a seleção.")
        
        with col_grafico_atribuir:
            st.markdown("#### Serviços a Atribuir")
            visao_atribuir = st.radio("Filtrar a Atribuir por:", ["Regional", "Município"], key="radio_atribuir", horizontal=True)
            coluna_atribuir = 'REGIONAL' if visao_atribuir == "Regional" else 'MUNICIPIO'
            
            df_agregado_atribuir = agregar_por_dimensao(df_base_total, coluna_atribuir, 'a_atribuir')
            if not df_agregado_atribuir.empty:
                st.plotly_chart(plot_bar_chart(df_agregado_atribuir, 'Dimensão', 'Métrica', 'Serviços a Atribuir por ' + visao_atribuir, visao_atribuir, 'Quantidade', color_discrete_sequence=[config.palette['ACCENT']]), use_container_width=True)
        
        st.markdown("---")
        
        st.markdown("#### Serviços Pendentes")
        visao_pendentes = st.radio("Filtrar Pendentes por:", ["Regional", "Município"], key="radio_pendentes", horizontal=True)
        coluna_pendentes = 'REGIONAL' if visao_pendentes == "Regional" else 'MUNICIPIO'
        
        df_agregado_pendentes = agregar_por_dimensao(df_para_analise, coluna_pendentes, 'pendentes')
        if not df_agregado_pendentes.empty:
            st.plotly_chart(plot_bar_chart(df_agregado_pendentes, 'Dimensão', 'Métrica', 'Serviços Pendentes por ' + visao_pendentes, visao_pendentes, 'Quantidade', color_discrete_sequence=[config.palette['ACCENT']]), use_container_width=True)

    with tab_colaboradores:
        st.markdown("### Desempenho dos seus Colaboradores")
        
        if kpis['colaboradores_nao_encontrados']:
            st.warning(f"⚠️ **Atenção:** Os seguintes colaboradores da sua lista não foram encontrados na base de dados: {', '.join(kpis['colaboradores_nao_encontrados'])}")
            st.markdown("---")
            
        if not df_colab_performance.empty:
            # Adiciona o filtro de pesquisa
            col_search, _ = st.columns([2, 8])
            with col_search:
                search_term = st.text_input("Pesquisar por nome:", "").upper()
            
            df_filtrado_colab = df_colab_performance
            if search_term:
                df_filtrado_colab = df_filtrado_colab[df_filtrado_colab[config.coluna_colaborador].str.contains(search_term, na=False)]

            # Ordena a tabela pela quantidade de serviços executados
            df_filtrado_colab = df_filtrado_colab.sort_values(by='Qtd_Executados', ascending=False)
            
            st.markdown("#### Tabela de Desempenho Individual Completa")
            st.dataframe(df_filtrado_colab.assign(**{
                'Qtd_Executados': df_filtrado_colab['Qtd_Executados'].apply(formatar_inteiro),
                'Qtd_Produtivos': df_filtrado_colab['Qtd_Produtivos'].apply(formatar_inteiro),
                'Qtd_Improdutivos': df_filtrado_colab['Qtd_Improdutivos'].apply(formatar_inteiro),
                'Qtd_Em_Campo': df_filtrado_colab['Qtd_Em_Campo'].apply(formatar_inteiro),
                'Qtd_Alocados': df_filtrado_colab['Qtd_Alocados'].apply(formatar_inteiro),
            }), use_container_width=True, hide_index=True)

            st.markdown("---")
            st.markdown("#### Gráficos de Desempenho por Colaborador")
            
            metricas_opcoes = {
                'Qtd_Alocados': {'label': 'Quantidade de Serviços Alocados', 'color': config.palette['SECONDARY_ACCENT']},
                'Qtd_Executados': {'label': 'Quantidade de Serviços Executados', 'color': config.palette['PRIMARY']},
                'Qtd_Produtivos': {'label': 'Quantidade de Serviços Produtivos', 'color': config.palette['SUCCESS']},
                'Qtd_Improdutivos': {'label': 'Quantidade de Serviços Improdutivos', 'color': config.palette['DANGER']},
                'Qtd_Em_Campo': {'label': 'Quantidade de Serviços em Campo', 'color': config.palette['PRIMARY']},
                'Qtd_Pendentes': {'label': 'Quantidade de Serviços Pendentes', 'color': config.palette['ACCENT']},
            }
            
            selecao_metrica = st.selectbox(
                "Selecione a métrica para o gráfico:",
                options=list(metricas_opcoes.keys()),
                format_func=lambda x: metricas_opcoes[x]['label'],
                key='metrica_colaborador_chart'
            )
            
            metrica_selecionada = metricas_opcoes[selecao_metrica]
            
            if not df_filtrado_colab.empty:
                total_metrica = df_filtrado_colab[selecao_metrica].sum()
                total_formatado = formatar_inteiro(total_metrica)
                titulo_grafico = f"{metrica_selecionada['label']} por Colaborador (Total: {total_formatado})"
            else:
                titulo_grafico = f"{metrica_selecionada['label']} por Colaborador"
            
            fig_colab = plot_bar_chart(
                df_data=df_filtrado_colab, 
                x_col=config.coluna_colaborador, 
                y_col=selecao_metrica,
                title=titulo_grafico,
                x_label="Colaborador",
                y_label="Quantidade de Serviços",
                color_discrete_sequence=[metrica_selecionada['color']]
            )
            fig_colab.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig_colab, use_container_width=True)
            
            st.markdown("---")
            st.markdown("### Baixar Relatório")
            
            buffer_pdf = BytesIO()
            report_generator = RelatorioVisualPDF(config.logo_path, config.palette, buffer_pdf)
            pdf_data = report_generator.generate_report(df_base_total, df_para_analise, df_filtrado_colab, kpis['colaboradores_nao_encontrados'])
            
            if pdf_data:
                st.download_button(
                    label="📥 Gerar e Baixar Relatório em PDF",
                    data=pdf_data,
                    file_name=f"Relatorio_Busca_Ativa_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                    mime="application/pdf"
                )
        else:
            st.info("Nenhum dado de desempenho disponível para os colaboradores selecionados.")

else:
    st.info("Por favor, selecione as opções nos filtros no topo da página para exibir os dados do painel.")

st.markdown("---")
st.markdown(f"<p style='text-align:center; font-size:14px; color:{config.palette['GREY_DARK']};'>Criado por PLINIO M. RODRIGUES. &copy; {datetime.now().year}</p>", unsafe_allow_html=True)