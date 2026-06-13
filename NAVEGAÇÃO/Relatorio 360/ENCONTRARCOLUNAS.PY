import pandas as pd

caminho = r"C:\Users\le37118890\OneDrive - Grupo Amil\Codigos_leleo\Relatorio 360\BASESSONA.xlsx"
df = pd.read_excel(caminho)

print("\n=== COLUNAS ENCONTRADAS NA BASE HISTÓRICA ===")
for col in df.columns:
    print(f"-> '{col}'")
print("=============================================\n")
