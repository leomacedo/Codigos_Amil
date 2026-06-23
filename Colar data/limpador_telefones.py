import re
import os
import sys

# --- COLE SUA COLUNA DE TELEFONES AQUI ---
TELEFONES_CRUS = """
21976374908
21984304223
21969659697
11969282188
11969282188
11940179720
11970264484
11999999999
21998517390
11973681363
11995926483
11946728982
11918357007
11989775184
11949646503
11987346138
21986129699
11992334376
21979447630
21980105075
2199148252
95879928
21972355650
11994718055
95609610
21972279953
21990030621
21965021427
21989546094
11982729154
11975006310
93232575
11948411296
11951748184
21979863099
21999921227
21968358152
11916108634
21980963584
21967514328
2174800511
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
