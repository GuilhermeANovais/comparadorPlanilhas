import pandas as pd
import numpy as np

# --- CONFIGURAÇÃO ---
ARQUIVO_A = 'POSTO.xlsx'
ARQUIVO_B = 'PMM.xlsx'

COLUNA_DATA = 'Data'
COLUNA_NOTA = 'Numero_Nota'
COLUNA_TIPO = 'Tipo_Combustivel'
COLUNA_VALOR = 'Litragem'
COLUNA_SETOR = 'Setor'


# --------------------

def formatar_data(series):
    datas_convertidas = pd.to_datetime(series, errors='coerce')
    return datas_convertidas.dt.strftime('%d/%m/%Y').fillna('Data Inválida/Nula')


# Substitua sua função 'comparar_planilhas' por esta versão corrigida

def comparar_planilhas():
    """
    Função principal que carrega, compara as planilhas e exibe os resultados.
    """
    try:
        df_a = pd.read_excel(ARQUIVO_A)
        df_b = pd.read_excel(ARQUIVO_B)

        print(f"✔ Planilha '{ARQUIVO_A}' carregada com {len(df_a)} registros.")
        print(f"✔ Planilha '{ARQUIVO_B}' carregada com {len(df_b)} registros.\n")

    except FileNotFoundError as e:
        print(f"❌ ERRO: Arquivo não encontrado! Verifique se o nome '{e.filename}' está correto.")
        return
    except Exception as e:
        print(f"❌ ERRO inesperado ao ler as planilhas: {e}")
        return

    df_a[COLUNA_NOTA] = df_a[COLUNA_NOTA].astype(str).str.zfill(4)
    df_b[COLUNA_NOTA] = df_b[COLUNA_NOTA].astype(str).str.zfill(4)

    df_merged = pd.merge(df_a, df_b, on=COLUNA_NOTA, suffixes=('_A', '_B'), how='outer')

    print("--- INICIANDO ANÁLISE DE DIVERGÊNCIAS ---")

    # 1. NOTAS FALTANTES
    notas_apenas_em_a = df_merged[df_merged[f'{COLUNA_VALOR}_B'].isnull()]
    notas_apenas_em_b = df_merged[df_merged[f'{COLUNA_VALOR}_A'].isnull()]

    # BLOCO CORRIGIDO ABAIXO
    if not notas_apenas_em_a.empty:
        print(f"\n[AVISO] Notas encontradas APENAS na Planilha {ARQUIVO_A}:")
        for index, row in notas_apenas_em_a.iterrows():
            data_formatada = formatar_data(row[[f'{COLUNA_DATA}_A']]).iloc[0]
            # Corrigido para usar a coluna com sufixo _A e remover Setor
            print(f"  - Data: {data_formatada} | Nota: {row[COLUNA_NOTA]} | Litragem: {row[f'{COLUNA_VALOR}_A']}")

    if not notas_apenas_em_b.empty:
        print(f"\n[AVISO] Notas encontradas APENAS na Planilha {ARQUIVO_B}:")
        for index, row in notas_apenas_em_b.iterrows():
            data_formatada = formatar_data(row[[f'{COLUNA_DATA}_B']]).iloc[0]
            # Corrigido para usar a coluna com sufixo _B
            print(f"  - Data: {data_formatada} | Nota: {row[COLUNA_NOTA]} | Litragem: {row[f'{COLUNA_VALOR}_B']} (Setor: {row[COLUNA_SETOR]})")

    df_comum = df_merged.dropna(subset=[f'{COLUNA_VALOR}_A', f'{COLUNA_VALOR}_B']).copy()

    # 2. DATAS DIVERGENTES
    df_comum[f'{COLUNA_DATA}_A'] = pd.to_datetime(df_comum[f'{COLUNA_DATA}_A'], errors='coerce')
    df_comum[f'{COLUNA_DATA}_B'] = pd.to_datetime(df_comum[f'{COLUNA_DATA}_B'], errors='coerce')

    datas_erradas = df_comum[df_comum[f'{COLUNA_DATA}_A'] != df_comum[f'{COLUNA_DATA}_B']]
    if not datas_erradas.empty:
        print("\n[ERRO] Notas com DATAS divergentes:")
        for index, row in datas_erradas.iterrows():
            data_a = formatar_data(row[[f'{COLUNA_DATA}_A']]).iloc[0]
            data_b = formatar_data(row[[f'{COLUNA_DATA}_B']]).iloc[0]
            print(
                f"  - Nota: {row[COLUNA_NOTA]} (Setor: {row[COLUNA_SETOR]}) | Planilha {ARQUIVO_A}: {data_a} vs Planilha {ARQUIVO_B}: {data_b}")

    # 3. TIPO DE COMBUSTÍVEL DIVERGENTE
    tipos_trocados = df_comum[df_comum[f'{COLUNA_TIPO}_A'] != df_comum[f'{COLUNA_TIPO}_B']]
    if not tipos_trocados.empty:
        print("\n[ERRO] Notas com TIPO DE COMBUSTÍVEL divergente:")
        for index, row in tipos_trocados.iterrows():
            data_a = formatar_data(row[[f'{COLUNA_DATA}_A']]).iloc[0]
            print(
                f"  - Data: {data_a} | Nota: {row[COLUNA_NOTA]} (Setor: {row[COLUNA_SETOR]}) | Planilha {ARQUIVO_A}: {row[f'{COLUNA_TIPO}_A']} vs Planilha {ARQUIVO_B}: {row[f'{COLUNA_TIPO}_B']}")

    # 4. LITRAGEM DIVERGENTE
    valores_errados = df_comum[~np.isclose(df_comum[f'{COLUNA_VALOR}_A'], df_comum[f'{COLUNA_VALOR}_B'])]
    if not valores_errados.empty:
        print("\n[ERRO] Notas com LITRAGEM divergente:")
        for index, row in valores_errados.iterrows():
            data_a = formatar_data(row[[f'{COLUNA_DATA}_A']]).iloc[0]
            print(
                f"  - Data: {data_a} | Nota: {row[COLUNA_NOTA]} (Setor: {row[COLUNA_SETOR]}) | Planilha {ARQUIVO_A}: {row[f'{COLUNA_VALOR}_A']}L vs Planilha {ARQUIVO_B}: {row[f'{COLUNA_VALOR}_B']}L")

    print("\n--- ANÁLISE CONCLUÍDA ---")

if __name__ == "__main__":
    comparar_planilhas()