import pandas as pd
import os

# --- CONFIGURAÇÃO ---
# Coloque aqui o nome exato do seu ficheiro CSV
NOME_DO_ARQUIVO_CSV = 'COMPONENTES NO PADRAO.csv'

# --- SCRIPT DE VERIFICAÇÃO ---

print(f"A analisar o ficheiro: {NOME_DO_ARQUIVO_CSV}...")

# Verifica se o ficheiro existe na pasta
if not os.path.exists(NOME_DO_ARQUIVO_CSV):
    print(f"\nERRO: O ficheiro '{NOME_DO_ARQUIVO_CSV}' não foi encontrado.")
    print("Por favor, certifique-se de que o script está na mesma pasta que a sua planilha.")
else:
    try:
        # Lê a planilha da mesma forma que o robô faz
        df = pd.read_csv(NOME_DO_ARQUIVO_CSV, encoding='latin-1', delimiter=';')

        # Remove espaços dos nomes das colunas para evitar erros
        df.columns = df.columns.str.strip()

        # 1. Conta o número total de linhas (parcelas)
        total_de_linhas = len(df)

        # 2. Conta o número de valores únicos na coluna 'DOCUMENTO' (incluindo vazios, como o Excel)
        total_de_unicos_bruto = df['DOCUMENTO'].nunique(dropna=False)

        # 3. Conta o número de valores únicos que NÃO SÃO VAZIOS (como o robô faz)
        total_de_notas_validas = df['DOCUMENTO'].nunique(dropna=True)

        # 4. Conta quantas linhas têm valores vazios na coluna DOCUMENTO
        linhas_vazias = df['DOCUMENTO'].isna().sum()

        print("\n--- RESULTADO DA ANÁLISE DETALHADA ---")
        print(f"Número total de linhas na planilha: {total_de_linhas}")
        print(f"Número de itens únicos (incluindo vazios) - Contagem do Excel: {total_de_unicos_bruto}")
        print(f"Número de notas fiscais VÁLIDAS e únicas - Contagem do Robô: {total_de_notas_validas}")
        print(f"Número de linhas com a coluna 'DOCUMENTO' vazia: {linhas_vazias}")
        print("-----------------------------------------")

        if total_de_unicos_bruto > total_de_notas_validas:
            print("\nConclusão: A diferença ocorre porque a planilha contém linhas vazias na coluna 'DOCUMENTO'.")
            print("O robô ignora corretamente estas linhas e processa apenas as 816 notas válidas.")
        else:
            print("\nConclusão: O robô está a contar corretamente as notas fiscais únicas.")


    except KeyError:
        print("\nERRO: Não consegui encontrar a coluna 'DOCUMENTO' na sua planilha.")
        print("Por favor, verifique se o nome da coluna está correto.")
    except Exception as e:
        print(f"\nOcorreu um erro inesperado ao ler o ficheiro: {e}")

