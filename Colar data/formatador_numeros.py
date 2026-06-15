import os
import sys
import pandas as pd

# --- CONFIGURAÇÕES ---
NOME_PLANILHA = "numeros.xlsx"
NOME_COLUNA = "mo"

def obter_diretorio_script():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def processar_planilha(caminho_arquivo, coluna, caminho_saida):
    try:
        df = pd.read_excel(caminho_arquivo, engine="openpyxl")

        if coluna not in df.columns:
            raise ValueError(f"A coluna '{coluna}' não foi encontrada.")

        serie = df[coluna].dropna()

        total = len(serie)

        with open(caminho_saida, 'w', encoding='utf-8') as f:
            for i, valor in enumerate(serie):
                numero = str(valor).strip()

                if not numero:
                    continue

                if i < total - 1:
                    f.write(f'"{numero}",\n')
                else:
                    f.write(f'"{numero}"')

        print(f"\n✅ Processados {total} números com sucesso!")

    except Exception as e:
        print(f"Erro: {e}")
        sys.exit(1)

if __name__ == "__main__":
    pasta = obter_diretorio_script()

    caminho_planilha = os.path.join(pasta, NOME_PLANILHA)
    caminho_saida = os.path.join(pasta, "numeros_formatados.txt")

    processar_planilha(caminho_planilha, NOME_COLUNA, caminho_saida)

    print(f"\nArquivo gerado em:\n{caminho_saida}")
    input("\nPressione Enter para sair...")