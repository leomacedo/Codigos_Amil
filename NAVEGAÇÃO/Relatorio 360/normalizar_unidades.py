import pandas as pd
import unicodedata
import re
import os

def normalizar_unidade(nome_bruto, log_nao_mapeadas=None):
    if pd.isna(nome_bruto): return ""
    
    # Remove acentos e joga para maiúsculo
    val = unicodedata.normalize('NFD', str(nome_bruto))
    val = "".join([c for c in val if not unicodedata.combining(c)]).upper().strip()
    
    # Lixo ou preenchimento inválido
    if re.search(r'^X+$', val) or 'MARIANA PENNA' in val: return 'Não Preenchido/Inválido'
    
    # Digital / Telemedicina
    if re.search(r'TELEMEDICINA|VIRTUAL|FUMO ZERO', val): return 'Telemedicina'
    if 'DOM PEDRO' in val: return 'Centro Médico Dom Pedro'
    if 'CONEXA' in val: return 'Conexa'
    if 'ONLINE' in val: return 'Online'
    
    # Amil Espaço Saúde (AES)
    if 'CAMPO GRANDE'in val: return 'Amil Espaço Saúde - Campo Grande'
    if 'NOVA IGUACU' in val: return 'Amil Espaço Saúde - Nova Iguaçu'
    if 'TIJUCA' in val: return 'Amil Espaço Saúde - Tijuca'
    if 'BOTAFOGO' in val: return 'Amil Espaço Saúde - Botafogo'
    if 'CAXIAS' in val: return 'Amil Espaço Saúde - Caxias'
    if 'GUARULHOS' in val: return 'Amil Espaço Saúde - Guarulhos'
    if 'NITEROI' in val: return 'Amil Espaço Saúde - Niterói'
    if 'OSASCO' in val: return 'Amil Espaço Saúde - Osasco'
    if 'SANTANA' in val: return 'Amil Espaço Saúde - Santana'
    if 'TATUAPE' in val: return 'Amil Espaço Saúde - Tatuapé'
    if 'ANA ROSA' in val: return 'Amil Espaço Saúde - Ana Rosa'
    if 'SANTO AMARO' in val: return 'Amil Espaço Saúde - Santo Amaro'
    
    # Unidades Avançadas
    if 'BUTANTA' in val: return 'Unidade Avançada Luz Butantã'
    if 'JOAO DIAS' in val or 'LUZ JOAO' in val: return 'Unidade Avançada Luz João Dias'
    if 'CARLOS CHAGAS' in val: return 'Unidade Avançada Carlos Chagas'
    if 'VITORIA' in val: return 'Unidade Avançada Vitória'
    if 'AVANCADA LUZ' in val: return 'Unidade Avançada Luz'
    
    # Hospitais e Centros Médicos
    if 'IPIRANGA MOGI' in val: return 'Hospital Ipiranga Mogi'
    if 'PAULISTANO' in val: return 'Hospital Paulistano'
    if 'CUBATAO' in val: return 'Hospital da Luz - Cubatão'
    if 'DA LUZ' in val: return 'Hospital da Luz'
    if 'MEDICO JARDIM' in val: return 'AP Centro Médico Jardim'
    if 'JOAO AZEVEDO' in val: return 'AP Centro Médico João Azevedo (SBC)'
    if 'LUIZ FERREIRA' in val: return 'AP Centro Médico Luiz Ferreira (SBC)'
    if 'SAO LUCAS' in val: return 'Hospital São Lucas'
    if 'SAMARITANO HIGIENOPOLIS' in val: return 'Hospital Samaritano Higienópolis'
    if 'SAMARITANO PAULISTA' in val: return 'Hospital Samaritano Paulista'
    if 'ALVORADA' in val: return 'Hospital Alvorada'
    if 'AMERICAS' in val: return 'Hospital Américas'
    if 'SANTA JOANA' in val: return 'Hospital Santa Joana (Recife)'
    
    # Fallback de segurança (Limpa lixo com a palavra UNIDADE caso apareça unidade nova não mapeada)
    fallback = re.sub(r'^(DADOS UNIDADES )?UNIDADE:\s*', '', val)
    resultado_final = str(nome_bruto).strip() if fallback == val else fallback.title()
    
    # Se chegou até aqui e o texto não está vazio, é porque a unidade não estava no dicionário acima
    if log_nao_mapeadas is not None and resultado_final:
        log_nao_mapeadas.add(str(nome_bruto).strip())
        
    return resultado_final

def processar_arquivo(caminho_arquivo, nome_coluna):
    print(f"Lendo arquivo: {caminho_arquivo}")
    
    try:
        if caminho_arquivo.endswith('.csv'):
            try:
                df = pd.read_csv(caminho_arquivo, sep=';', encoding='utf-8-sig')
            except UnicodeDecodeError:
                df = pd.read_csv(caminho_arquivo, sep=';', encoding='latin1')
        else:
            df = pd.read_excel(caminho_arquivo)
            
        df.columns = df.columns.str.strip()
        
        if nome_coluna not in df.columns:
            print(f"ERRO: A coluna '{nome_coluna}' não foi encontrada na planilha.")
            print(f"Colunas disponíveis: {df.columns.tolist()}")
            return
            
        print(f"Formatando a coluna '{nome_coluna}'...")
        unidades_desconhecidas = set()
        df[nome_coluna] = df[nome_coluna].apply(normalizar_unidade, log_nao_mapeadas=unidades_desconhecidas)
        
        diretorio = os.path.dirname(caminho_arquivo)
        nome_base, ext = os.path.splitext(os.path.basename(caminho_arquivo))
        caminho_saida = os.path.join(diretorio, f"{nome_base}_formatado{ext}")
        
        if ext.lower() == '.csv':
            df.to_csv(caminho_saida, sep=';', index=False, encoding='utf-8-sig')
        else:
            df.to_excel(caminho_saida, index=False)
            
        print("-" * 50)
        print(f"SUCESSO! Arquivo salvo em:\n{caminho_saida}")
        
        if unidades_desconhecidas:
            print("\nATENÇÃO: Foram encontradas unidades que não estão no dicionário de regras:")
            for unid in sorted(unidades_desconhecidas):
                print(f" - {unid}")
            print(f"Total de unidades não mapeadas: {len(unidades_desconhecidas)}")
            
        print("-" * 50)
        
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

if __name__ == "__main__":
    # === CONFIGURAÇÕES ===
    # Mude o caminho abaixo para a planilha que você deseja limpar
    CAMINHO_PLANILHA = r"C:\Users\le37118890\OneDrive - Grupo Amil\Codigos_leleo\NAVEGAÇÃO\Relatorio 360\leo.xlsx"
    
    # Coloque aqui o nome exato da coluna onde ficam as unidades
    NOME_DA_COLUNA = "Porta de Entrada" 
    # =====================
    
    if os.path.exists(CAMINHO_PLANILHA):
        processar_arquivo(CAMINHO_PLANILHA, NOME_DA_COLUNA)
    else:
        print(f"Arquivo não encontrado: {CAMINHO_PLANILHA}")