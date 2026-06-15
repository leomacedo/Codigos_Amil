import re
import os
import sys

# --- COLE SUA COLUNA DE TELEFONES AQUI ---
TELEFONES_CRUS = """
21976535327
21973179293
21987618579
21968899982
21998657128
21-970559020
21981160720
21-970559020
21999078188
21965122393
21970460243
21992849987
21-996560478
21-994253609
21-980910360
11997064309
21976516379
11983342999
21989177455
11940188063
21996042832
21-994160133
(19) 998002501
(11)  994283335
(19)  993042935  
(11)  962221220
(21) 971836719
(11)  962221220
(11)  962856104
(11) 994478218
(21) 973254246;(21)  996687968
(11)  930061767
11984268949
11980812325
11912260375
11959427354
11967904790
"""

def limpar_telefones(texto_cru):
    # Quebra o texto em linhas
    linhas = texto_cru.strip().split('\n')
    
    resultado = []
    for linha in linhas:
        # Remove tudo que não for dígito (0-9), ponto-e-vírgula (;) ou barra (/)
        linha_limpa = re.sub(r'[^0-9;/]', '', linha)
        
        # Unifica os separadores (transforma / em ;) e separa os números
        telefones = linha_limpa.replace('/', ';').split(';')
        
        telefones_validos = []
        for tel in telefones:
            if not tel:
                continue
                
            # Se o número tem 10 dígitos, adiciona o '9' logo após o DDD (os 2 primeiros dígitos)
            if len(tel) == 10:
                tel = tel[:2] + '9' + tel[2:]

            # VERIFICAÇÃO: tem 11 dígitos E o 3º dígito (índice 2) é um '9'?
            if len(tel) == 11 and tel[2] == '9':
                telefones_validos.append(tel)
        
        # Junta os telefones válidos e adiciona na lista de resultados
        resultado.append('; '.join(telefones_validos))
            
    return '\n'.join(resultado)

if __name__ == "__main__":
    diretorio_script = os.path.dirname(os.path.abspath(__file__)) if not getattr(sys, 'frozen', False) else os.path.dirname(sys.executable)
        
    caminho_saida = os.path.join(diretorio_script, "telefones_limpos.txt")
    texto_formatado = limpar_telefones(TELEFONES_CRUS)
    
    with open(caminho_saida, 'w', encoding='utf-8') as f:
        f.write(texto_formatado)
        
    print(f"SUCESSO! O arquivo foi gerado em:\n{caminho_saida}\n")
    print("--- CONTEÚDO GERADO ---\n" + texto_formatado + "\n-----------------------")
    input("\nPressione Enter para sair...")
