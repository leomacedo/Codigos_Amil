import pandas as pd
import os
import sys

def normalizar_baseleo():
    # 1. Define o diretório atual e procura pelos arquivos baseleo
    if getattr(sys, 'frozen', False):
        diretorio_script = os.path.dirname(sys.executable)
    else:
        diretorio_script = os.path.dirname(os.path.abspath(__file__))
        
    caminho_xlsx = os.path.join(diretorio_script, "baseleo.xlsx")
    caminho_csv = os.path.join(diretorio_script, "baseleo.csv")
    
    if os.path.exists(caminho_xlsx):
        caminho_arquivo = caminho_xlsx
        df = pd.read_excel(caminho_arquivo)
    elif os.path.exists(caminho_csv):
        caminho_arquivo = caminho_csv
        df = pd.read_csv(caminho_arquivo, sep=';', encoding='utf-8-sig') # Tenta utf-8, fallback pode ser latin1
    else:
        print(f"ERRO: Arquivo 'baseleo.xlsx' ou 'baseleo.csv' não encontrado na pasta:\n{diretorio_script}")
        return

    print(f"Processando arquivo: {caminho_arquivo}...")
    
    # Guarda os nomes originais para tentar restaurar no final (opcional)
    colunas_originais = df.columns.tolist()
    
    # Padroniza temporariamente as colunas para minúsculo para evitar erros de digitação (espaços)
    df.columns = df.columns.str.strip().str.lower()

    # --- TRATAMENTOS ---

    # 1. Marca ótica (Deixar com 9 dígitos, preenchendo com zero à esquerda)
    if 'marca_otica' in df.columns:
        df['marca_otica'] = df['marca_otica'].astype(str).str.replace(r'\.0$', '', regex=True).str.replace('nan', '')
        df['marca_otica'] = df['marca_otica'].apply(lambda x: x.zfill(9) if x.strip() != '' else '')

    # 2. Data de Nascimento (Converter de AAAAMMDD para DD/MM/AAAA)
    if 'data_nascimento' in df.columns:
        df['data_nascimento'] = df['data_nascimento'].astype(str).str.replace(r'\.0$', '', regex=True).str.replace('nan', '')
        # Transforma o formato 20090519 para datatime, e depois extrai a string em DD/MM/AAAA
        # fillna mantém o dado original se a conversão falhar (ex: data que já estava certa)
        df['data_nascimento'] = pd.to_datetime(df['data_nascimento'], format='%Y%m%d', errors='coerce').dt.strftime('%d/%m/%Y').fillna(df['data_nascimento'])

    # 3. Nível Rede (Extrair apenas os números, ex: NÍVEL 60 -> 60)
    if 'nivel_rede' in df.columns:
        df['nivel_rede'] = df['nivel_rede'].astype(str).str.extract(r'(\d+)')[0].fillna('')

    # 4. UF Município (Deixar apenas o UF os 2 primeiros dígitos)
    if 'uf_municipio' in df.columns:
        df['uf_municipio'] = df['uf_municipio'].astype(str).str.replace('nan', '').apply(lambda x: x[:2] if x else '')

    # 5. Juntar DDD e Celular
    if 'ddd_tel_celular' in df.columns and 'tel_celular' in df.columns:
        # Remove casas decimais caso os números tenham vindo como "11.0" em vez de "11"
        ddd = df['ddd_tel_celular'].astype(str).str.replace(r'\.0$', '', regex=True).str.replace('nan', '')
        celular = df['tel_celular'].astype(str).str.replace(r'\.0$', '', regex=True).str.replace('nan', '')
        
        # Junta as duas colunas na própria coluna "tel_celular"
        df['tel_celular'] = ddd + celular
        
        # (Opcional) Remove a coluna original de DDD já que ela foi aglutinada
        df = df.drop(columns=['ddd_tel_celular'])

    # Volta as colunas para o formato que você via ou deixa assim mesmo (opcional)
    # Para mapear exatamente: o pandas fará um salvamento limpo.
    
    # --- EXPORTAÇÃO ---
    caminho_saida = os.path.join(diretorio_script, "baseleo_normalizada.xlsx")
    df.to_excel(caminho_saida, index=False)
    
    print(f"SUCESSO! O arquivo normalizado foi salvo em:\n{caminho_saida}")

if __name__ == "__main__":
    normalizar_baseleo()
    input("\nPressione Enter para fechar a tela...")
