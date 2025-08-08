import pandas as pd
from pathlib import Path

# --- Configurações Gerais ---
pd.set_option('display.max_rows', None)

# Caminho completo do arquivo Excel.
# ATENÇÃO: Substitua 'BUSCA ATIVA.xlsx' pelo nome e caminho exato do seu arquivo.
arquivo_excel = Path(r"C:\Users\USUARIO\Desktop\BUSCA ATIVA\BUSCA ATIVA.xlsx")

if not arquivo_excel.exists():
    print(f"Erro: Arquivo '{arquivo_excel}' não encontrado.")
    exit()

# --- Leitura e Pré-processamento da Planilha ---
try:
    df = pd.read_excel(arquivo_excel, sheet_name="Sheet1")
    
    # Normaliza nomes de colunas
    df.columns = df.columns.str.strip().str.upper().str.replace(' ', '_').str.replace('[^A-Z0-9_]', '', regex=True)
    
    # Valida a existência das colunas essenciais
    required_cols = ['DATA_ATRIBUICAO', 'DATA_CARREGOU', 'DATA_DEVOLUCAO', 'REGIONAL', 'MUNICIPIO', 'NOME_FASE', 'ALVO_CONDICAO_OBJETIVA']
    if not all(col in df.columns for col in required_cols):
        missing_cols = [col for col in required_cols if col not in df.columns]
        print(f"Erro: Colunas essenciais faltando na planilha: {missing_cols}")
        exit()

    # Filtro obrigatório: apenas dados com ALVO_CONDICAO_OBJETIVA = 'SIM'
    df = df[df['ALVO_CONDICAO_OBJETIVA'].str.upper() == 'SIM'].copy()
    
    # Filtro de Regionais: apenas NORTE, NORDESTE, SUL
    regionais_validas = ['NORTE', 'NORDESTE', 'SUL']
    df = df[df['REGIONAL'].str.upper().isin(regionais_validas)].copy()

    if df.empty:
        print("Atenção: Nenhum dado foi encontrado após a aplicação dos filtros iniciais. Verifique a planilha.")
        exit()

    # Converte colunas de data para o formato correto
    for col in ['DATA_ATRIBUICAO', 'DATA_CARREGOU', 'DATA_DEVOLUCAO']:
        df[col] = pd.to_datetime(df[col], errors='coerce')
        # Cria a coluna de mês para agrupamento no formato 'YYYY-MM'
        df[f'MES_{col.split("_")[1]}'] = df[col].dt.to_period('M').astype(str)

    # A coluna INSTALACAO não será mais usada para soma, apenas para contagem (se necessário).
    # Aqui, vamos focar na contagem total de serviços (linhas)
    
except Exception as e:
    print(f"Erro ao processar a planilha: {e}")
    exit()

# ============================
# TOTAIS GERAIS
# ============================

# Métrica principal alterada para contagem de serviços
qtd_total_servicos = len(df)
print("\nQUANTIDADE TOTAL DE SERVIÇOS: {:,.0f}".format(qtd_total_servicos))

# Mapeamento dos status de serviço
SERVICOS_EXECUTADOS = ['CONCLUIDO OK', 'DESCARREGADO COM IMPEDIMENTO', 'DESCARREGADO SEM IMPEDIMENTO', 'IMPROCEDENTE']
SERVICOS_EM_CAMPO = ['ALVO EM CAMPO']
SERVICOS_NAO_ATRIBUIDOS = ['ALVO NAO ATRIBUIDO']
SERVICOS_NAO_RECEBIDOS = ['ALVO ENVIADO - NAO RECEBIDO']

qtd_servicos_executados = df[df['NOME_FASE'].str.upper().isin(SERVICOS_EXECUTADOS)]['NOME_FASE'].count()
qtd_servicos_em_campo = df[df['NOME_FASE'].str.upper().isin(SERVICOS_EM_CAMPO)]['NOME_FASE'].count()
qtd_servicos_nao_atribuidos = df[df['NOME_FASE'].str.upper().isin(SERVICOS_NAO_ATRIBUIDOS)]['NOME_FASE'].count()
qtd_servicos_nao_recebidos = df[df['NOME_FASE'].str.upper().isin(SERVICOS_NAO_RECEBIDOS)]['NOME_FASE'].count()

print("QUANTIDADE DE SERVIÇOS EXECUTADOS:", qtd_servicos_executados)
print("QUANTIDADE DE SERVIÇOS EM CAMPO:", qtd_servicos_em_campo)
print("QUANTIDADE DE SERVIÇOS A ATRIBUIR:", qtd_servicos_nao_atribuidos)
print("QUANTIDADE DE SERVIÇOS NAO RECEBIDOS:", qtd_servicos_nao_recebidos)

# ============================
# RESUMO POR REGIONAL
# ============================

resumo_regional = df.groupby('REGIONAL').agg(
    QTD_TOTAL_SERVICOS=('REGIONAL', 'count'),  # Contagem por regional
    QTD_EXECUTADOS=('NOME_FASE', lambda x: x.str.upper().isin(SERVICOS_EXECUTADOS).sum()),
    QTD_EM_CAMPO=('NOME_FASE', lambda x: x.str.upper().isin(SERVICOS_EM_CAMPO).sum()),
    QTD_A_ATRIBUIR=('NOME_FASE', lambda x: x.str.upper().isin(SERVICOS_NAO_ATRIBUIDOS).sum()),
    QTD_NAO_RECEBIDOS=('NOME_FASE', lambda x: x.str.upper().isin(SERVICOS_NAO_RECEBIDOS).sum())
).reset_index().sort_values(by='QTD_TOTAL_SERVICOS', ascending=False)

print("\n--- RESUMO POR REGIONAL ---")
print(resumo_regional)

# ============================
# RESUMO POR MUNICIPIO
# ============================

resumo_municipio = df.groupby('MUNICIPIO').agg(
    QTD_TOTAL_SERVICOS=('MUNICIPIO', 'count'), # Contagem por município
    QTD_EXECUTADOS=('NOME_FASE', lambda x: x.str.upper().isin(SERVICOS_EXECUTADOS).sum()),
    QTD_EM_CAMPO=('NOME_FASE', lambda x: x.str.upper().isin(SERVICOS_EM_CAMPO).sum()),
    QTD_A_ATRIBUIR=('NOME_FASE', lambda x: x.str.upper().isin(SERVICOS_NAO_ATRIBUIDOS).sum()),
    QTD_NAO_RECEBIDOS=('NOME_FASE', lambda x: x.str.upper().isin(SERVICOS_NAO_RECEBIDOS).sum())
).reset_index().sort_values(by='QTD_TOTAL_SERVICOS', ascending=False)

print("\n--- RESUMO POR MUNICIPIO ---")
print(resumo_municipio)

# ============================
# SALVAR EM EXCEL
# ============================

salvar = True  # Altere para False se não quiser salvar
if salvar:
    try:
        with pd.ExcelWriter("resumo_busca_ativa_consolidado.xlsx") as writer:
            resumo_regional.to_excel(writer, sheet_name="Resumo_Por_Regional", index=False)
            resumo_municipio.to_excel(writer, sheet_name="Resumo_Por_Municipio", index=False)
        print("\n✅ Arquivo Excel salvo como 'resumo_busca_ativa_consolidado.xlsx'")
    except Exception as e:
        print(f"❌ Erro ao salvar o arquivo Excel: {e}")