import pandas as pd
import os
import sys

def analisar_dados_jornadas(caminho_entrada, caminho_saida):
    print(f"Lendo a planilha: {caminho_entrada}")
    try:
        # Lê a planilha
        df = pd.read_excel(caminho_entrada)
        
        # Verifica se as colunas necessárias existem
        colunas_necessarias = ['Jornada', 'Estado', 'Mês', 'Porta de Entrada']
        for col in colunas_necessarias:
            if col not in df.columns:
                print(f"ERRO: A coluna '{col}' não foi encontrada na planilha.")
                return
                
        # Trata os valores vazios para que eles também entrem na contagem conforme você pediu
        df['Porta de Entrada'] = df['Porta de Entrada'].fillna('(Vazio)')
        df['Estado'] = df['Estado'].fillna('(Vazio)')
        df['Jornada'] = df['Jornada'].fillna('Sem Jornada')
        df['Mês'] = df['Mês'].fillna('(Vazio)')
        
        jornadas = df['Jornada'].unique()
        resumo_geral = []
        resumo_por_mes = []
        
        # Prepara o resumo detalhado por mês
        for mes in df['Mês'].unique():
            df_mes = df[df['Mês'] == mes]
            for jornada in df_mes['Jornada'].unique():
                df_mes_jorn = df_mes[df_mes['Jornada'] == jornada]
                resumo_por_mes.append({
                    'Mês': mes,
                    'Jornada': jornada,
                    'Geral': len(df_mes_jorn),
                    'SP': len(df_mes_jorn[df_mes_jorn['Estado'].astype(str).str.upper().str.strip() == 'SP']),
                    'RJ': len(df_mes_jorn[df_mes_jorn['Estado'].astype(str).str.upper().str.strip() == 'RJ'])
                })
        
        with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
            for jornada in jornadas:
                # Filtra apenas os dados da jornada atual
                df_jornada = df[df['Jornada'] == jornada]
                
                # Separa SP e RJ (e previne erros de digitação como espaços extras ou maiúsculas/minúsculas)
                df_sp = df_jornada[df_jornada['Estado'].astype(str).str.upper().str.strip() == 'SP']
                df_rj = df_jornada[df_jornada['Estado'].astype(str).str.upper().str.strip() == 'RJ']
                
                total_geral = len(df_jornada)
                total_sp = len(df_sp)
                total_rj = len(df_rj)
                
                # 1. Criação do quadro superior: Visão Geral / SP / RJ
                df_resumo_aba = pd.DataFrame({
                    'Métrica': ['Geral', 'SP', 'RJ'],
                    'Quantidade': [total_geral, total_sp, total_rj]
                })
                
                # Função auxiliar para não repetir código na tabela das Portas de Entrada
                def gerar_tabela_portas(df_estado, nome_estado):
                    if df_estado.empty:
                        return pd.DataFrame({f'Porta de Entrada - {nome_estado}': ['Sem dados'], 'Quantidade': [0]})
                        
                    # Conta as portas de entrada e reseta o índice para virar coluna
                    tabela = df_estado['Porta de Entrada'].value_counts().reset_index()
                    tabela.columns = [f'Porta de Entrada - {nome_estado}', 'Quantidade']
                    
                    # Adiciona a linha de "Total Geral"
                    total = pd.DataFrame([{f'Porta de Entrada - {nome_estado}': 'Total Geral', 'Quantidade': tabela['Quantidade'].sum()}])
                    tabela = pd.concat([tabela, total], ignore_index=True)
                    return tabela
                
                # 2. Criação das tabelas de Portas de Entrada para SP e RJ
                tabela_sp = gerar_tabela_portas(df_sp, 'SP')
                tabela_rj = gerar_tabela_portas(df_rj, 'RJ')
                
                # Limpa o nome da jornada para não quebrar a aba do Excel (que não aceita caracteres especiais)
                nome_aba = str(jornada)
                for char in ['/', '\\', '?', '*', ':', '[', ']']:
                    nome_aba = nome_aba.replace(char, '')
                nome_aba = nome_aba[:31].strip() # O Excel tem limite de 31 caracteres no nome da aba
                
                # Exporta as tabelas em posições diferentes da planilha
                df_resumo_aba.to_excel(writer, sheet_name=nome_aba, startrow=0, startcol=0, index=False)
                tabela_sp.to_excel(writer, sheet_name=nome_aba, startrow=5, startcol=0, index=False) # Começa na Coluna A
                tabela_rj.to_excel(writer, sheet_name=nome_aba, startrow=5, startcol=3, index=False) # Começa na Coluna D
                
                # Ajusta as larguras das colunas para os nomes não ficarem apertados
                worksheet = writer.sheets[nome_aba]
                worksheet.column_dimensions['A'].width = 40
                worksheet.column_dimensions['B'].width = 15
                worksheet.column_dimensions['D'].width = 40
                worksheet.column_dimensions['E'].width = 15

                # Salva os totais para o "Resumão" final
                resumo_geral.append({
                    'Jornada': jornada,
                    'Geral': total_geral,
                    'SP': total_sp,
                    'RJ': total_rj
                })
                
            # Cria a aba final com o resumo de todas as jornadas juntas
            df_resumo_total = pd.DataFrame(resumo_geral)
            df_resumo_total.to_excel(writer, sheet_name='Resumo Geral', index=False)
            ws_resumo = writer.sheets['Resumo Geral']
            ws_resumo.column_dimensions['A'].width = 35
            ws_resumo.column_dimensions['B'].width = 15
            ws_resumo.column_dimensions['C'].width = 15
            ws_resumo.column_dimensions['D'].width = 15
            
            # Cria a aba final com o resumo separado por mês
            df_resumo_mes = pd.DataFrame(resumo_por_mes)
            df_resumo_mes.to_excel(writer, sheet_name='Resumo por Mês', index=False)
            ws_mes = writer.sheets['Resumo por Mês']
            ws_mes.column_dimensions['A'].width = 20
            ws_mes.column_dimensions['B'].width = 35
            ws_mes.column_dimensions['C'].width = 15
            ws_mes.column_dimensions['D'].width = 15
            ws_mes.column_dimensions['E'].width = 15
            
        print("-" * 50)
        print(f"SUCESSO! Relatório detalhado gerado em:\n{caminho_saida}")
        print("-" * 50)
        
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

if __name__ == "__main__":
    # Mantendo o seu padrão: localiza o diretório automático
    if getattr(sys, 'frozen', False):
        diretorio_atual = os.path.dirname(sys.executable)
    else:
        diretorio_atual = os.path.dirname(os.path.abspath(__file__))
        
    # Salve sua planilha como "dados_jornadas.xlsx" na mesma pasta
    caminho_in = os.path.join(diretorio_atual, "dados_jornadas.xlsx")
    caminho_out = os.path.join(diretorio_atual, "relatorio_jornadas_formatado.xlsx")
    
    if os.path.exists(caminho_in):
        analisar_dados_jornadas(caminho_in, caminho_out)
    else:
        print(f"ATENÇÃO: Por favor, salve a sua planilha original com o nome 'dados_jornadas.xlsx' na pasta:\n{diretorio_atual}")
        print("E rode o script novamente.")
        
    input("\nPressione Enter para sair...")
