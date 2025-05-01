import pandas as pd

def comparar_planilhas(df1, df2, coluna1_combo, coluna2_combo, checkboxes):
    if df1 is None or df2 is None:
        raise ValueError("Carregue as duas planilhas antes de comparar.")

    coluna1 = f"{coluna1_combo}_1"
    coluna2 = f"{coluna2_combo}_2"

    if coluna1_combo == "Escolha a coluna da primeira planilha" or coluna2_combo == "Escolha a coluna da segunda planilha":
        raise ValueError("Por favor, selecione as duas colunas para comparar.")

    comparado = pd.merge(
        df1,
        df2,
        left_on=df1.columns[0],
        right_on=df2.columns[0],
        how="outer",
        suffixes=("_1", "_2"),
    )

    if checkboxes.get("soma"):
        comparado["Soma"] = comparado[coluna1] + comparado[coluna2]
    if checkboxes.get("diferenca"):
        comparado["Diferença"] = comparado[coluna1] - comparado[coluna2]
    if checkboxes.get("multiplicacao"):
        comparado["Multiplicação"] = comparado[coluna1] * comparado[coluna2]
    if checkboxes.get("divisao"):
        comparado["Divisão"] = comparado[coluna1] / comparado[coluna2]
    if checkboxes.get("media"):
        comparado["Média"] = (comparado[coluna1] + comparado[coluna2]) / 2
    if checkboxes.get("diferenca_percentual"):
        comparado["Diferença percentual"] = (
            (comparado[coluna1] - comparado[coluna2]) / comparado[coluna2]
        ) * 100

    return comparado