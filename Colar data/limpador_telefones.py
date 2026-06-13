import re
import os
import sys

# --- COLE SUA COLUNA DE TELEFONES AQUI ---
TELEFONES_CRUS = """
11988582455
21-970559020
21-970559020
11 944801469
11946549993
11911235940
11989986471
11981005518
11932171157
11960250682
21964283388
11992061897
11966932277
(21) 9933-55456
(11) 9390-11969
(11) 9864-10853
(21) 9792-83954
(21) 9765-80789
(11) 9513-17787
(11) 9807-37842
(21) 9972-45635
(21) 9680-0620
11997097866
(11) 9102-02472
(11) 9670-18913
(11) 9429-37623
(11) 9623-89807
(21) 9758-0370
(11) 9884-44323
(11) 9477-49578
(11) 9658-44046
(11) 9735-11652
(21) 9683-65005
(11) 9762-14207
(11) 9474-97375
(21) 9959-58673
(21) 9812-01024
(11) 9415-57493
(21) 9809-02961
(11) 9848-42007
(21) 9804-04660
(11) 9895-29960
(11) 9776-79251
(21) 9943-01852
(11) 9838-62025
(21) 9841-08981
(11) 9946-48040
(11) 9937-02728
(21) 9608-94206
(21) 9817-18605
(11) 9719-41012
(11) 9549-26109
(21) 9854-20840
(11) 9723-35651
(11) 9527-13866
(11) 9637-56394
(11) 9692-42620
(21) 9824-65347
(11) 9852-99022
(21) 9840-38282
(21) 9825-16355
(21) 9960-23903
(21) 9944-55343
(11) 9977-21874
(11) 9824-80074
(11) 9444-03430
(21) 9722-06768
(21) 97189-2905
(21) 9997-00838
(21) 9818-67806
(21) 9943-06780
(21) 9646-46330
(21) 9807-56869
(21) 9826-5617
(21) 9953-93309
(21) 9835-79662
(21) 9810-03530
(21) 9830-9843
(11) 9897-80466
(11) 9863-69674
(85) 9881-36528
(85) 9884-22168
(85) 9892-56188
(85) 9893-71283
(85) 9881-25998
(85) 9979-84450
(85) 9994-74476
(85) 9986-09034
(11) 9720-00275
(21) 9934-44602
(11) 9447-24805
(11) 9693-84120
(11) 9472-00364
(11) 9964-63483
(11) 9706-74149
(11) 9152-00847
(11) 9805-13346
(11) 9654-62185
(11) 9728-83688
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
