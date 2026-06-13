import pandas as pd
import os

# Pega a pasta onde o script está salvo
pasta_script = os.path.dirname(os.path.abspath(__file__))

# Caminho da planilha na mesma pasta do script
arquivo_entrada = os.path.join(pasta_script, "planilha.xlsx")
arquivo_saida = os.path.join(pasta_script, "planilha_sem_repetidos.xlsx")

# Lê a planilha
df = pd.read_excel(arquivo_entrada, engine="openpyxl")

# Remove espaços extras dos nomes das colunas
df.columns = df.columns.str.strip()

print("Colunas encontradas na planilha:")
for coluna in df.columns:
    print(f"- {coluna}")

# Tenta encontrar a coluna Nome Completo ignorando maiúsculas/minúsculas e espaços extras
coluna_nome = None

for coluna in df.columns:
    coluna_tratada = coluna.lower().strip()

    if coluna_tratada in ["nome completo", "nome colpeto"]:
        coluna_nome = coluna
        break

if coluna_nome is None:
    print("\nERRO: Não encontrei a coluna de nome completo.")
    print("Veja acima o nome exato das colunas e ajuste no código.")
else:
    # Remove espaços extras dos nomes dentro da coluna
    df[coluna_nome] = df[coluna_nome].astype(str).str.strip()

    # Remove nomes repetidos mantendo a primeira aparição
    df_sem_repetidos = df.drop_duplicates(subset=[coluna_nome], keep="first")

    # Salva a nova planilha
    df_sem_repetidos.to_excel(arquivo_saida, index=False, engine="openpyxl")

    print(f"\nColuna usada: {coluna_nome}")
    print(f"Linhas antes: {len(df)}")
    print(f"Linhas depois: {len(df_sem_repetidos)}")
    print(f"Duplicados removidos: {len(df) - len(df_sem_repetidos)}")
    print(f"Planilha criada com sucesso: {arquivo_saida}")