import pandas as pd
import numpy as np
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4


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

    # =====================================================================
    # 1. VALORES NEGATIVOS E LITRAGENS ANORMALMENTE ALTAS (sem limites fixos)
    # =====================================================================
    print("\n--- VERIFICANDO VALORES INVÁLIDOS DE LITRAGEM ---")

    negativos_a = df_a[df_a[COLUNA_VALOR] < 0]
    negativos_b = df_b[df_b[COLUNA_VALOR] < 0]

    if not negativos_a.empty:
        print(f"\n[ERRO] Litragens negativas na planilha {ARQUIVO_A}:")
        for _, row in negativos_a.iterrows():
            print(f"  - Nota: {row[COLUNA_NOTA]} | Litragem: {row[COLUNA_VALOR]}")

    if not negativos_b.empty:
        print(f"\n[ERRO] Litragens negativas na planilha {ARQUIVO_B}:")
        for _, row in negativos_b.iterrows():
            print(f"  - Nota: {row[COLUNA_NOTA]} | Litragem: {row[COLUNA_VALOR]}")

    # Detectar valores incomuns estatisticamente (acima do percentil 99)
    limite_alto_a = df_a[COLUNA_VALOR].quantile(0.99)
    limite_alto_b = df_b[COLUNA_VALOR].quantile(0.99)

    altos_a = df_a[df_a[COLUNA_VALOR] > limite_alto_a]
    altos_b = df_b[df_b[COLUNA_VALOR] > limite_alto_b]

    if not altos_a.empty:
        print(f"\n[ALERTA] Litragens incomuns (muito acima do normal) na planilha {ARQUIVO_A}:")
        for _, row in altos_a.iterrows():
            print(f"  - Nota: {row[COLUNA_NOTA]} | Litragem: {row[COLUNA_VALOR]}")

    if not altos_b.empty:
        print(f"\n[ALERTA] Litragens incomuns (muito acima do normal) na planilha {ARQUIVO_B}:")
        for _, row in altos_b.iterrows():
            print(f"  - Nota: {row[COLUNA_NOTA]} | Litragem: {row[COLUNA_VALOR]}")

    # =====================================================================
    # 2. LITROS INCOMPATÍVEIS COM O TIPO (estatístico, sem limites fixos)
    # =====================================================================
    print("\n--- ANALISANDO CONSISTÊNCIA ENTRE TIPO DE COMBUSTÍVEL E LITRAGEM ---")

    def analisar_tipo(df, nome_arquivo):
        for tipo in df[COLUNA_TIPO].dropna().unique():
            tipo_filtro = str(tipo).upper()
            subset = df[df[COLUNA_TIPO].str.upper() == tipo_filtro]

            if subset.empty:
                continue

            media = subset[COLUNA_VALOR].mean()

            # Procurar registros muito fora da média
            suspeitos = subset[subset[COLUNA_VALOR] > media * 2]

            for _, row in suspeitos.iterrows():
                print(f"[ALERTA] Litragem fora do padrão do tipo na planilha {nome_arquivo}:")
                print(f"  - Nota {row[COLUNA_NOTA]} | Tipo: {tipo_filtro} | Litragem: {row[COLUNA_VALOR]} (média: {media:.1f})")

    analisar_tipo(df_a, ARQUIVO_A)
    analisar_tipo(df_b, ARQUIVO_B)

    # =====================================================================
    # 3. NOTAS REPETIDAS NO MESMO DIA COM MESMA LITRAGEM
    # =====================================================================
    print("\n--- CHECANDO NOTAS DUPLICADAS NO MESMO DIA E MESMA LITRAGEM ---")

    def repetidas_mesmo_dia(df, nome_arquivo):
        df_aux = df.copy()
        df_aux['DataFmt'] = pd.to_datetime(df_aux[COLUNA_DATA], errors='coerce').dt.strftime("%d/%m/%Y")

        duplicadas = df_aux[df_aux.duplicated(subset=['DataFmt', COLUNA_NOTA, COLUNA_VALOR], keep=False)]

        if not duplicadas.empty:
            print(f"\n[ALERTA] Notas repetidas no mesmo dia e com a mesma litragem em {nome_arquivo}:")
            for _, row in duplicadas.iterrows():
                print(
                    f"  - Data: {row['DataFmt']} | Nota: {row[COLUNA_NOTA]} | Litragem: {row[COLUNA_VALOR]}"
                )

    repetidas_mesmo_dia(df_a, ARQUIVO_A)
    repetidas_mesmo_dia(df_b, ARQUIVO_B)

    # =====================================================================
    # 4. MESMA NOTA EM DIAS DIFERENTES
    # =====================================================================
    print("\n--- VERIFICANDO MESMA NOTA EM DIAS DIFERENTES ---")

    def nota_em_dias_diferentes(df, nome_arquivo):
        df_aux = df.copy()
        df_aux['DataFmt'] = pd.to_datetime(df_aux[COLUNA_DATA], errors='coerce').dt.strftime("%d/%m/%Y")

        conflitos = df_aux.groupby(COLUNA_NOTA)['DataFmt'].nunique()
        conflitos = conflitos[conflitos > 1]

        if not conflitos.empty:
            print(f"\n[ERRO] Mesma nota utilizada em dias diferentes na planilha {nome_arquivo}:")
            for nota in conflitos.index:
                datas = df_aux[df_aux[COLUNA_NOTA] == nota]['DataFmt'].unique()
                print(f"  - Nota {nota} usada nas datas: {', '.join(datas)}")

    nota_em_dias_diferentes(df_a, ARQUIVO_A)
    nota_em_dias_diferentes(df_b, ARQUIVO_B)

    # =====================================================================
    # 5. MERGE E VERIFICAÇÕES ORIGINAIS
    # =====================================================================
    df_merged = pd.merge(df_a, df_b, on=COLUNA_NOTA, suffixes=('_A', '_B'), how='outer')

    print("\n--- INICIANDO ANÁLISE DE DIVERGÊNCIAS ---")

    notas_apenas_em_a = df_merged[df_merged[f'{COLUNA_VALOR}_B'].isnull()]
    notas_apenas_em_b = df_merged[df_merged[f'{COLUNA_VALOR}_A'].isnull()]

    if not notas_apenas_em_a.empty:
        print(f"\n[AVISO] Notas APENAS na planilha {ARQUIVO_A}:")
        for _, row in notas_apenas_em_a.iterrows():
            data_formatada = formatar_data(pd.Series([row[f'{COLUNA_DATA}_A']])).iloc[0]
            print(f"  - Data: {data_formatada} | Nota: {row[COLUNA_NOTA]} | Litragem: {row[f'{COLUNA_VALOR}_A']}")

    if not notas_apenas_em_b.empty:
        print(f"\n[AVISO] Notas APENAS na planilha {ARQUIVO_B}:")
        for _, row in notas_apenas_em_b.iterrows():
            data_formatada = formatar_data(pd.Series([row[f'{COLUNA_DATA}_B']])).iloc[0]
            print(
                f"  - Data: {data_formatada} | Nota: {row[COLUNA_NOTA]} | Litragem: {row[f'{COLUNA_VALOR}_B']} (Setor: {row[COLUNA_SETOR]})"
            )

    df_comum = df_merged.dropna(subset=[f'{COLUNA_VALOR}_A', f'{COLUNA_VALOR}_B']).copy()

    # DATAS DIVERGENTES
    df_comum[f'{COLUNA_DATA}_A'] = pd.to_datetime(df_comum[f'{COLUNA_DATA}_A'], errors='coerce')
    df_comum[f'{COLUNA_DATA}_B'] = pd.to_datetime(df_comum[f'{COLUNA_DATA}_B'], errors='coerce')

    datas_erradas = df_comum[df_comum[f'{COLUNA_DATA}_A'] != df_comum[f'{COLUNA_DATA}_B']]
    if not datas_erradas.empty:
        print("\n[ERRO] Datas divergentes:")
        for _, row in datas_erradas.iterrows():
            print(
                f"  - Nota {row[COLUNA_NOTA]} | "
                f"{formatar_data(pd.Series([row[f'{COLUNA_DATA}_A']])).iloc[0]} vs "
                f"{formatar_data(pd.Series([row[f'{COLUNA_DATA}_B']])).iloc[0]}"
            )

    # TIPOS DIVERGENTES
    tipos_trocados = df_comum[df_comum[f'{COLUNA_TIPO}_A'] != df_comum[f'{COLUNA_TIPO}_B']]
    if not tipos_trocados.empty:
        print("\n[ERRO] Tipos divergentes:")
        for _, row in tipos_trocados.iterrows():
            print(
                f"  - Nota {row[COLUNA_NOTA]} | {row[f'{COLUNA_TIPO}_A']} vs {row[f'{COLUNA_TIPO}_B']}"
            )

    # LITRAGEM DIVERGENTE
    valores_errados = df_comum[~np.isclose(df_comum[f'{COLUNA_VALOR}_A'], df_comum[f'{COLUNA_VALOR}_B'])]
    if not valores_errados.empty:
        print("\n[ERRO] Litragem divergente:")
        for _, row in valores_errados.iterrows():
            print(
                f"  - Nota {row[COLUNA_NOTA]} | {row[f'{COLUNA_VALOR}_A']}L vs {row[f'{COLUNA_VALOR}_B']}L"
            )

    print("\n--- ANÁLISE CONCLUÍDA ---")

    # =====================================================================
    # 7. GERAR RELATÓRIO EM PDF
    # =====================================================================
    print("Gerando relatório em PDF...")

    styles = getSampleStyleSheet()
    pdf = SimpleDocTemplate("relatorio_abastecimentos.pdf", pagesize=A4)

    story = []

    def add(titulo, conteudo):
        story.append(Paragraph(f"<b>{titulo}</b>", styles["Heading3"]))
        story.append(Paragraph(conteudo, styles["BodyText"]))
        story.append(Spacer(1, 12))

    add("Resumo Executivo",
        f"""
        Total Registros A: {len(df_a)}<br/>
        Total Registros B: {len(df_b)}<br/>
        Notas apenas em A: {len(notas_apenas_em_a)}<br/>
        Notas apenas em B: {len(notas_apenas_em_b)}<br/>
        Datas Divergentes: {len(datas_erradas)}<br/>
        Tipos Divergentes: {len(tipos_trocados)}<br/>
        Litragens Divergentes: {len(valores_errados)}<br/>
        """)

    add("Observações Importantes",
        "Este relatório foi gerado automaticamente pelo sistema de auditoria de abastecimentos.")

    pdf.build(story)

    print("📑 PDF 'relatorio_abastecimentos.pdf' gerado com sucesso!")


if __name__ == "__main__":
    comparar_planilhas()