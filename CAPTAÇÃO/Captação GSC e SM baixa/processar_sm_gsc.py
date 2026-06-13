import pandas as pd
import os
import sys
from datetime import datetime
import traceback

def processar_arquivos(diretorio):
    lista_dfs = []
    
    for arquivo in os.listdir(diretorio):
        if arquivo.startswith('~$') or not (arquivo.endswith('.xlsx') or arquivo.endswith('.csv')):
            continue
            
        if arquivo == "base_captacao_gerada.xlsx":
            continue
            
        nome_arquivo_lower = arquivo.lower()
        if 'sm' in nome_arquivo_lower:
            tipo_planilha = 'sm'
        elif 'gsc' in nome_arquivo_lower:
            tipo_planilha = 'gsc'
        else:
            continue # Pula os arquivos que não têm "sm" ou "gsc" no nome
            
        caminho_arquivo = os.path.join(diretorio, arquivo)
        print(f"Processando: {arquivo} (Tipo identificado: {tipo_planilha.upper()})")
        
        try:
            # Lê o arquivo
            if arquivo.endswith('.csv'):
                df = pd.read_csv(caminho_arquivo, sep=';', encoding='utf-8-sig')
            else:
                df = pd.read_excel(caminho_arquivo)
                
            # 1. Usar a coluna exata 'status' (ou 'STATUS')
            col_status = 'status' if 'status' in df.columns else 'STATUS'
            if col_status not in df.columns:
                df[col_status] = '' # Cria a coluna se não existir
                
            # 2. Filtrar onde o status é vazio
            df[col_status] = df[col_status].fillna('').astype(str)
            mask_vazios = df[col_status].str.strip() == ''
            
            if not mask_vazios.any():
                print(f"  -> Nenhum registro vazio encontrado. Pulando...")
                continue
                
            df_vazios = df[mask_vazios].copy()
            
            # 3. Atualizar o status original e salvar o arquivo de volta
            df.loc[mask_vazios, col_status] = 'enviado captação'
            
            if arquivo.endswith('.csv'):
                df.to_csv(caminho_arquivo, sep=';', index=False, encoding='utf-8-sig')
            else:
                df.to_excel(caminho_arquivo, index=False)
            print(f"  -> Status preenchido com 'enviado captação' e arquivo original salvo!")
            
            # 4. Criar o df_captacao puxando os nomes exatos sem procurar
            df_captacao = pd.DataFrame(index=df_vazios.index)
            
            # Regras de Negócio
            fonte_caminho = "Planilha Total Care" if tipo_planilha == "sm" else "GSC"
            linha_cuidado = "Saúde Mental" if tipo_planilha == "sm" else df_vazios.get('linha', df_vazios.get('LINHA DE CUIDADO', ""))
            
            df_captacao['FONTE'] = fonte_caminho
            df_captacao['CAMINHO'] = fonte_caminho
            df_captacao['DATA INSERÇÃO'] = datetime.now().strftime('%d/%m/%Y')
            df_captacao['LINHA DE CUIDADO'] = linha_cuidado
            df_captacao['Nome'] = df_vazios.get('nome', df_vazios.get('Nome', ""))
            df_captacao['MO'] = df_vazios.get('mo', df_vazios.get('MO', ""))
            df_captacao['CPF'] = df_vazios.get('cpf', df_vazios.get('CPF', ""))
            df_captacao['DATA DE NASCIMENTO'] = df_vazios.get('nascimento', df_vazios.get('DATA DE NASCIMENTO', ""))
            df_captacao['NÍVEL DE REDE'] = ""
            df_captacao['UF'] = df_vazios.get('uf', df_vazios.get('UF', ""))
            df_captacao['MUNICÍPIO'] = df_vazios.get('municipio', df_vazios.get('MUNICÍPIO', ""))
            df_captacao['TELEFONE'] = df_vazios.get('telefone', df_vazios.get('TELEFONE', ""))
            df_captacao['STATUS'] = ""
            
            lista_dfs.append(df_captacao)
            
        except Exception as e:
            print(f"ERRO ao processar o arquivo {arquivo}: {str(e)}")
            traceback.print_exc()
            
    if lista_dfs:
        return pd.concat(lista_dfs, ignore_index=True)
    return pd.DataFrame()

if __name__ == "__main__":
    diretorio_script = os.path.dirname(os.path.abspath(__file__)) if not getattr(sys, 'frozen', False) else os.path.dirname(sys.executable)
    
    print("="*50)
    print("PROCESSANDO PLANILHAS SM E GSC NA PASTA ATUAL")
    df_consolidado = processar_arquivos(diretorio_script)
    
    if not df_consolidado.empty:
        caminho_saida = os.path.join(diretorio_script, "base_captacao_gerada.xlsx")
        df_consolidado.to_excel(caminho_saida, index=False)
        print(f"\nSUCESSO! {len(df_consolidado)} registros filtrados e salvos em:\n{caminho_saida}")
        
    input("\nProcesso finalizado. Pressione Enter para sair...")