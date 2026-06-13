import pandas as pd
import os
import sys
import re


def formatar_mo(valor):
    """
    Padroniza MO / Marca Ótica para 9 dígitos.
    Exemplo:
    123456 vira 000123456
    123456.0 vira 000123456
    """
    if pd.isna(valor):
        return ""

    valor = str(valor).strip()

    if valor.lower() in ["nan", "nat", ""]:
        return ""

    valor = valor.replace(".0", "")
    valor = valor.replace(" ", "")

    return valor.zfill(9)


def normalizar_texto(valor):
    """
    Padroniza textos para comparação.
    Remove espaços extras e coloca tudo em maiúsculo.
    """
    if pd.isna(valor):
        return ""

    valor = str(valor).strip().upper()
    valor = re.sub(r"\s+", " ", valor)

    return valor


def normalizar_mes(valor):
    """
    Padroniza a coluna Mês para comparação.
    Exemplo:
    Janeiro/2026
    janeiro / 2026
    JANEIRO/2026
    Janiero/2026
    todos viram:
    JANEIRO/2026
    """
    if pd.isna(valor):
        return ""

    valor = str(valor).strip().upper()

    # Remove espaços extras
    valor = re.sub(r"\s+", " ", valor)

    # Remove espaços antes/depois da barra
    valor = valor.replace(" /", "/").replace("/ ", "/")

    # Correções comuns de digitação
    valor = valor.replace("JANIERO", "JANEIRO")
    valor = valor.replace("JANIER0", "JANEIRO")
    valor = valor.replace("JANEIRO ", "JANEIRO")

    return valor


def criar_comparacao_total(base_com_mo, relatorio_com_mo):
    """
    Compara sem considerar mês.

    Regra:
    - Base usa MO_PADRONIZADA
    - Relatório usa MO_PADRONIZADA

    Resultado:
    - Na base: mostra se aquela MO está no relatório em qualquer mês.
    - No relatório: mostra se aquela MO está na base em qualquer mês.
    """

    print("\nCriando comparação TOTAL, sem considerar mês...")

    base_total = base_com_mo.copy()
    relatorio_total = relatorio_com_mo.copy()

    mos_base_set = set(base_total["MO_PADRONIZADA"])
    mos_relatorio_set = set(relatorio_total["MO_PADRONIZADA"])

    # Base: está ou não está no relatório, independente do mês
    base_total["STATUS_MO_NO_RELATORIO_TOTAL"] = base_total["MO_PADRONIZADA"].apply(
        lambda x: "ESTÁ" if x in mos_relatorio_set else "NÃO ESTÁ"
    )

    # Relatório: está ou não está na base, independente do mês
    relatorio_total["STATUS_MO_NA_BASE_TOTAL"] = relatorio_total["MO_PADRONIZADA"].apply(
        lambda x: "ESTÁ" if x in mos_base_set else "NÃO ESTÁ"
    )

    # Resumo total da base
    resumo_base_total = (
        base_total
        .groupby("STATUS_MO_NO_RELATORIO_TOTAL")["MO_PADRONIZADA"]
        .nunique()
        .reset_index(name="Qtd MOs da base")
        .rename(columns={"STATUS_MO_NO_RELATORIO_TOTAL": "STATUS"})
    )

    # Resumo total do relatório
    resumo_relatorio_total = (
        relatorio_total
        .groupby("STATUS_MO_NA_BASE_TOTAL")["MO_PADRONIZADA"]
        .nunique()
        .reset_index(name="Qtd MOs do relatório")
        .rename(columns={"STATUS_MO_NA_BASE_TOTAL": "STATUS"})
    )

    # Junta os resumos
    resumo_total = resumo_base_total.merge(
        resumo_relatorio_total,
        how="outer",
        on="STATUS"
    ).fillna(0)

    resumo_total = resumo_total[
        [
            "STATUS",
            "Qtd MOs da base",
            "Qtd MOs do relatório"
        ]
    ]

    return resumo_total, base_total, relatorio_total


def criar_comparacao_por_mes(base_com_mo, relatorio_com_mo):
    """
    Compara usando a coluna 'Mês' nas duas planilhas.

    Regra:
    - Base usa MO + Mês
    - Relatório usa Marca Ótica + Mês

    Resultado:
    - Na base: mostra se aquela MO está no relatório no mesmo mês.
    - No relatório: mostra se aquela MO está na base no mesmo mês.
    """

    coluna_mes = "Mês"

    if coluna_mes not in base_com_mo.columns:
        print(f"ATENÇÃO: A coluna '{coluna_mes}' não foi encontrada na base.")
        print("A comparação por mês será ignorada.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    if coluna_mes not in relatorio_com_mo.columns:
        print(f"ATENÇÃO: A coluna '{coluna_mes}' não foi encontrada no relatório.")
        print("A comparação por mês será ignorada.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    print("\nCriando comparação usando a coluna Mês...")

    base_mes = base_com_mo.copy()
    relatorio_mes = relatorio_com_mo.copy()

    # Padroniza a coluna Mês
    base_mes["MES_PADRONIZADO"] = base_mes[coluna_mes].apply(normalizar_mes)
    relatorio_mes["MES_PADRONIZADO"] = relatorio_mes[coluna_mes].apply(normalizar_mes)

    # Remove linhas sem mês válido
    base_mes = base_mes[base_mes["MES_PADRONIZADO"] != ""].copy()
    relatorio_mes = relatorio_mes[relatorio_mes["MES_PADRONIZADO"] != ""].copy()

    # Chave de comparação = MO + Mês
    base_mes["CHAVE_MO_MES"] = base_mes["MO_PADRONIZADA"] + "_" + base_mes["MES_PADRONIZADO"]
    relatorio_mes["CHAVE_MO_MES"] = relatorio_mes["MO_PADRONIZADA"] + "_" + relatorio_mes["MES_PADRONIZADO"]

    chaves_base_mes = set(base_mes["CHAVE_MO_MES"])
    chaves_relatorio_mes = set(relatorio_mes["CHAVE_MO_MES"])

    # Base: está ou não está no relatório no mesmo mês
    comparacao_base_mes = base_mes.copy()
    comparacao_base_mes["STATUS_MO_NO_RELATORIO_MES"] = comparacao_base_mes["CHAVE_MO_MES"].apply(
        lambda x: "ESTÁ" if x in chaves_relatorio_mes else "NÃO ESTÁ"
    )

    # Relatório: está ou não está na base no mesmo mês
    comparacao_relatorio_mes = relatorio_mes.copy()
    comparacao_relatorio_mes["STATUS_MO_NA_BASE_MES"] = comparacao_relatorio_mes["CHAVE_MO_MES"].apply(
        lambda x: "ESTÁ" if x in chaves_base_mes else "NÃO ESTÁ"
    )

    # Resumo da base por mês
    resumo_base_mes = (
        comparacao_base_mes
        .groupby(["MES_PADRONIZADO", "STATUS_MO_NO_RELATORIO_MES"])["MO_PADRONIZADA"]
        .nunique()
        .reset_index(name="Qtd MOs da base")
    )

    # Resumo do relatório por mês
    resumo_relatorio_mes = (
        comparacao_relatorio_mes
        .groupby(["MES_PADRONIZADO", "STATUS_MO_NA_BASE_MES"])["MO_PADRONIZADA"]
        .nunique()
        .reset_index(name="Qtd MOs do relatório")
    )

    # Junta os resumos
    resumo_mensal = resumo_base_mes.merge(
        resumo_relatorio_mes,
        how="outer",
        left_on=["MES_PADRONIZADO", "STATUS_MO_NO_RELATORIO_MES"],
        right_on=["MES_PADRONIZADO", "STATUS_MO_NA_BASE_MES"]
    )

    resumo_mensal["STATUS"] = resumo_mensal["STATUS_MO_NO_RELATORIO_MES"].combine_first(
        resumo_mensal["STATUS_MO_NA_BASE_MES"]
    )

    resumo_mensal = resumo_mensal[
        [
            "MES_PADRONIZADO",
            "STATUS",
            "Qtd MOs da base",
            "Qtd MOs do relatório"
        ]
    ].fillna(0)

    resumo_mensal = resumo_mensal.sort_values(
        by=["MES_PADRONIZADO", "STATUS"]
    )

    return resumo_mensal, comparacao_base_mes, comparacao_relatorio_mes


def diagnosticar_base_relatorio(caminho_base, caminho_relatorio, caminho_saida):
    print("Lendo base...")
    base = pd.read_excel(caminho_base)

    print("Lendo relatório...")
    relatorio = pd.read_excel(caminho_relatorio)

    # Limpa nomes das colunas
    base.columns = base.columns.str.strip()
    relatorio.columns = relatorio.columns.str.strip()

    coluna_mo_base = "MO"
    coluna_marca_relatorio = "Marca Ótica"
    coluna_programa = "O paciente entrou em contato para mais informações/agendamento de qual Programa?"

    # Valida coluna MO da base
    if coluna_mo_base not in base.columns:
        print(f"ERRO: A coluna '{coluna_mo_base}' não existe na base.")
        print("Colunas encontradas na base:")
        print(list(base.columns))
        return

    # Valida coluna Marca Ótica do relatório
    if coluna_marca_relatorio not in relatorio.columns:
        print(f"ERRO: A coluna '{coluna_marca_relatorio}' não existe no relatório.")
        print("Colunas encontradas no relatório:")
        print(list(relatorio.columns))
        return

    # Cria colunas padronizadas
    base["MO_PADRONIZADA"] = base[coluna_mo_base].apply(formatar_mo)
    relatorio["MO_PADRONIZADA"] = relatorio[coluna_marca_relatorio].apply(formatar_mo)

    # Remove MOs vazias apenas para as análises principais
    base_com_mo = base[base["MO_PADRONIZADA"] != ""].copy()
    relatorio_com_mo = relatorio[relatorio["MO_PADRONIZADA"] != ""].copy()

    # Quantidades principais
    total_linhas_base = len(base)
    total_linhas_relatorio = len(relatorio)

    total_mo_unicas_base = base_com_mo["MO_PADRONIZADA"].nunique()
    total_mo_unicas_relatorio = relatorio_com_mo["MO_PADRONIZADA"].nunique()

    # MOs repetidas na base
    contagem_base = (
        base_com_mo
        .groupby("MO_PADRONIZADA")
        .size()
        .reset_index(name="Qtd vezes na base")
        .sort_values(by="Qtd vezes na base", ascending=False)
    )

    mos_repetidas_base = contagem_base[
        contagem_base["Qtd vezes na base"] > 1
    ].copy()

    # MOs repetidas no relatório
    contagem_relatorio = (
        relatorio_com_mo
        .groupby("MO_PADRONIZADA")
        .size()
        .reset_index(name="Qtd vezes no relatório")
        .sort_values(by="Qtd vezes no relatório", ascending=False)
    )

    mos_repetidas_relatorio = contagem_relatorio[
        contagem_relatorio["Qtd vezes no relatório"] > 1
    ].copy()

    # Conjuntos de MOs, independente do mês
    mos_relatorio_set = set(relatorio_com_mo["MO_PADRONIZADA"])
    mos_base_set = set(base_com_mo["MO_PADRONIZADA"])

    # MOs da base que não existem no relatório, independente do mês
    base_sem_relatorio = base_com_mo[
        ~base_com_mo["MO_PADRONIZADA"].isin(mos_relatorio_set)
    ].copy()

    # MOs do relatório que não existem na base, independente do mês
    relatorio_sem_base = relatorio_com_mo[
        ~relatorio_com_mo["MO_PADRONIZADA"].isin(mos_base_set)
    ].copy()

    # Lista única de MOs que estão no relatório, mas não estão na base
    mos_relatorio_nao_estao_base = (
        relatorio_sem_base[["MO_PADRONIZADA"]]
        .drop_duplicates()
        .sort_values(by="MO_PADRONIZADA")
        .copy()
    )

    # Lista única de MOs que estão na base, mas não estão no relatório
    mos_base_nao_estao_relatorio = (
        base_sem_relatorio[["MO_PADRONIZADA"]]
        .drop_duplicates()
        .sort_values(by="MO_PADRONIZADA")
        .copy()
    )

    # MOs que existem nas duas planilhas
    mos_em_comum = sorted(list(mos_base_set.intersection(mos_relatorio_set)))

    df_mos_em_comum = pd.DataFrame({
        "MO_PADRONIZADA": mos_em_comum
    })

    # Analisa programas diferentes por MO no relatório
    if coluna_programa in relatorio_com_mo.columns:
        programas_por_mo = (
            relatorio_com_mo
            .dropna(subset=[coluna_programa])
            .assign(PROGRAMA_LIMPO=lambda df: df[coluna_programa].astype(str).str.strip())
            .groupby("MO_PADRONIZADA")["PROGRAMA_LIMPO"]
            .nunique()
            .reset_index(name="Qtd programas diferentes")
            .sort_values(by="Qtd programas diferentes", ascending=False)
        )

        mos_com_mais_de_um_programa = programas_por_mo[
            programas_por_mo["Qtd programas diferentes"] > 1
        ].copy()

        detalhes_programas = (
            relatorio_com_mo
            .sort_values(by="MO_PADRONIZADA")
            [["MO_PADRONIZADA", coluna_programa]]
            .drop_duplicates()
            .copy()
        )
    else:
        mos_com_mais_de_um_programa = pd.DataFrame()
        detalhes_programas = pd.DataFrame()

    # ============================================================
    # COMPARAÇÃO TOTAL, SEM CONSIDERAR MÊS
    # ============================================================
    resumo_total, comparacao_base_total, comparacao_relatorio_total = criar_comparacao_total(
        base_com_mo,
        relatorio_com_mo
    )

    # ============================================================
    # COMPARAÇÃO POR MÊS USANDO A COLUNA "Mês"
    # ============================================================
    resumo_mensal, comparacao_base_mes, comparacao_relatorio_mes = criar_comparacao_por_mes(
        base_com_mo,
        relatorio_com_mo
    )

    # Resumo geral sem MOs vazias e sem duplicatas exatas
    resumo = pd.DataFrame({
        "Métrica": [
            "Total de linhas na base",
            "Total de linhas no relatório",
            "Total de MOs únicas na base",
            "Total de MOs únicas no relatório",
            "MOs repetidas na base",
            "MOs repetidas no relatório",
            "MOs da base que não existem no relatório",
            "MOs do relatório que não existem na base",
            "MOs que existem nas duas planilhas"
        ],
        "Valor": [
            total_linhas_base,
            total_linhas_relatorio,
            total_mo_unicas_base,
            total_mo_unicas_relatorio,
            len(mos_repetidas_base),
            len(mos_repetidas_relatorio),
            base_sem_relatorio["MO_PADRONIZADA"].nunique(),
            relatorio_sem_base["MO_PADRONIZADA"].nunique(),
            len(mos_em_comum)
        ]
    })

    # Salva o diagnóstico
    with pd.ExcelWriter(caminho_saida, engine="openpyxl") as writer:
        resumo.to_excel(writer, sheet_name="RESUMO", index=False)

        # Comparação total
        resumo_total.to_excel(writer, sheet_name="Resumo total", index=False)
        comparacao_base_total.to_excel(writer, sheet_name="Base total", index=False)
        comparacao_relatorio_total.to_excel(writer, sheet_name="Relatorio total", index=False)

        # Comparação por mês
        resumo_mensal.to_excel(writer, sheet_name="Resumo mensal", index=False)
        comparacao_base_mes.to_excel(writer, sheet_name="Base mensal", index=False)
        comparacao_relatorio_mes.to_excel(writer, sheet_name="Relatorio mensal", index=False)

        # Abas de repetição
        mos_repetidas_base.to_excel(writer, sheet_name="MO repetida base", index=False)
        mos_repetidas_relatorio.to_excel(writer, sheet_name="MO repetida relatorio", index=False)

        # Abas completas, independente do mês
        base_sem_relatorio.to_excel(writer, sheet_name="Base sem relatorio", index=False)
        relatorio_sem_base.to_excel(writer, sheet_name="Relatorio sem base", index=False)

        # Abas só com MOs únicas, independente do mês
        mos_base_nao_estao_relatorio.to_excel(writer, sheet_name="MOs base nao relat", index=False)
        mos_relatorio_nao_estao_base.to_excel(writer, sheet_name="MOs relat nao base", index=False)

        # Aba de MOs em comum, independente do mês
        df_mos_em_comum.to_excel(writer, sheet_name="MOs em comum", index=False)

        # Abas de programa
        mos_com_mais_de_um_programa.to_excel(writer, sheet_name="MO varios programas", index=False)
        detalhes_programas.to_excel(writer, sheet_name="Detalhe programas", index=False)

    print("-" * 60)
    print("DIAGNÓSTICO GERADO COM SUCESSO!")
    print(caminho_saida)
    print("-" * 60)

    print("\nResumo:")
    print(resumo)

    print("\nMOs do relatório que não estão na base, independente do mês:")
    print(len(mos_relatorio_nao_estao_base))

    print("\nMOs da base que não estão no relatório, independente do mês:")
    print(len(mos_base_nao_estao_relatorio))

    print("\nComparação TOTAL criada nas abas:")
    print("- Resumo total")
    print("- Base total")
    print("- Relatorio total")

    print("\nComparação por MÊS criada nas abas:")
    print("- Resumo mensal")
    print("- Base mensal")
    print("- Relatorio mensal")


if __name__ == "__main__":
    if getattr(sys, "frozen", False):
        diretorio_atual = os.path.dirname(sys.executable)
    else:
        diretorio_atual = os.path.dirname(os.path.abspath(__file__))

    caminho_base = os.path.join(diretorio_atual, "base.xlsx")
    caminho_relatorio = os.path.join(diretorio_atual, "relatorio.xlsx")
    caminho_saida = os.path.join(diretorio_atual, "diagnostico_base_relatorio.xlsx")

    diagnosticar_base_relatorio(
        caminho_base,
        caminho_relatorio,
        caminho_saida
    )

    input("\nPressione Enter para fechar...")