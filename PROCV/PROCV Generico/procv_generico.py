import pandas as pd
import unicodedata
from pathlib import Path


# =========================
# CONFIGURAÇÕES
# =========================

pasta_script = Path(__file__).parent

entrada = pasta_script / "entrada.xlsx"
base = pasta_script / "basee.xlsx"
saida = pasta_script / "resultado.xlsx"

# Escolha o tipo da busca:
# "nome" para buscar por nome
# "mo" para buscar por MO de 9 dígitos
tipo_busca = "mo"

# Coluna usada para procurar nas duas bases
coluna_busca = "nome"

# Coluna da base que você quer trazer
coluna_valor = "status"


# =========================
# FUNÇÕES
# =========================

def limpar_texto(texto):
    """
    Limpa texto/nome:
    - remove espaços no começo e no fim
    - remove espaços duplicados
    - transforma em minúsculo
    - remove acentos
    """
    if pd.isna(texto):
        return ""

    texto = str(texto).strip().lower()
    texto = " ".join(texto.split())

    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))

    return texto


def limpar_mo(valor):
    """
    Limpa MO:
    - transforma em texto
    - remove .0 quando vem do Excel como número
    - mantém somente números
    - completa com zero à esquerda até 9 dígitos
    """
    if pd.isna(valor):
        return ""

    valor = str(valor).strip()

    if valor.endswith(".0"):
        valor = valor[:-2]

    valor = "".join(c for c in valor if c.isdigit())

    if valor:
        valor = valor.zfill(9)

    return valor


def limpar_colunas(df):
    """
    Limpa os nomes das colunas.
    Exemplo:
    ' Nome ' vira 'nome'
    'STATUS' vira 'status'
    """
    df.columns = [limpar_texto(col) for col in df.columns]
    return df


def validar_coluna(df, nome_arquivo, coluna):
    """
    Verifica se a coluna existe.
    Se não existir, mostra as colunas disponíveis.
    """
    if coluna not in df.columns:
        print(f"\nERRO: A coluna '{coluna}' não foi encontrada em {nome_arquivo}.")
        print(f"Colunas encontradas em {nome_arquivo}:")
        for col in df.columns:
            print(f"- {col}")
        raise KeyError(f"Coluna não encontrada: {coluna}")


def limpar_chave(valor):
    """
    Decide automaticamente qual limpeza usar conforme o tipo_busca.
    """
    if tipo_busca == "nome":
        return limpar_texto(valor)

    elif tipo_busca == "mo":
        return limpar_mo(valor)

    else:
        raise ValueError("tipo_busca inválido. Use 'nome' ou 'mo'.")


# =========================
# LEITURA DOS ARQUIVOS
# =========================

df_entrada = pd.read_excel(entrada)
df_base = pd.read_excel(base)


# =========================
# LIMPAR NOMES DAS COLUNAS
# =========================

df_entrada = limpar_colunas(df_entrada)
df_base = limpar_colunas(df_base)


# =========================
# VALIDAR COLUNAS
# =========================

validar_coluna(df_entrada, "entrada.xlsx", coluna_busca)
validar_coluna(df_base, "basee.xlsx", coluna_busca)
validar_coluna(df_base, "basee.xlsx", coluna_valor)


# =========================
# LIMPEZA DA CHAVE DE BUSCA
# =========================

df_entrada["_chave_limpa"] = df_entrada[coluna_busca].apply(limpar_chave)
df_base["_chave_limpa"] = df_base[coluna_busca].apply(limpar_chave)


# =========================
# AGRUPAR TODAS AS APARIÇÕES DA BASE
# =========================

# Cria um dicionário assim:
# chave -> lista de status encontrados
# Exemplo:
# "joao silva" -> ["Ativo", "Pendente", "Cancelado"]

mapa_busca = (
    df_base
    .groupby("_chave_limpa")[coluna_valor]
    .apply(list)
    .to_dict()
)


# =========================
# DESCOBRIR MAIOR QUANTIDADE DE APARIÇÕES
# =========================

maior_qtd_aparicoes = 0

for chave in df_entrada["_chave_limpa"]:
    valores = mapa_busca.get(chave, [])
    if len(valores) > maior_qtd_aparicoes:
        maior_qtd_aparicoes = len(valores)


# =========================
# CRIAR COLUNAS status_1, status_2, status_3...
# =========================

for i in range(maior_qtd_aparicoes):
    nome_coluna_nova = f"{coluna_valor}_{i + 1}"

    df_entrada[nome_coluna_nova] = df_entrada["_chave_limpa"].apply(
        lambda chave: mapa_busca.get(chave, [])[i]
        if i < len(mapa_busca.get(chave, []))
        else ""
    )


# =========================
# REMOVER COLUNA AUXILIAR
# =========================

df_entrada = df_entrada.drop(columns=["_chave_limpa"])


# =========================
# SALVAR RESULTADO
# =========================

df_entrada.to_excel(saida, index=False)

print(f"Arquivo gerado com sucesso: {saida}")
print(f"Maior quantidade de aparições encontradas: {maior_qtd_aparicoes}")
