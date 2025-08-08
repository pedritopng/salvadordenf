import pyautogui
import time
import pandas as pd

# --- CONFIGURAÇÃO E FUNÇÕES AUXILIARES ---
# Medida de segurança: Leve o mouse para o canto superior esquerdo para parar o script.
pyautogui.FAILSAFE = True

# --- CARREGANDO OS DADOS DO ARQUIVO CSV ---
# O nome do arquivo deve ser exatamente igual ao que está na pasta.
nome_do_arquivo_csv = 'TESTE_LISTA_NF.csv'

try:
    # --- CORREÇÃO APLICADA ---
    # Adicionado delimiter=';' pois arquivos CSV no Brasil geralmente usam ponto e vírgula como separador.
    df = pd.read_csv(nome_do_arquivo_csv, encoding='latin-1', delimiter=';')

    # Remove espaços em branco do início e do fim dos nomes das colunas.
    df.columns = df.columns.str.strip()

    # Pega apenas os valores ÚNICOS da coluna "NOTA FISCAL" para evitar processar a mesma nota várias vezes.
    notas_unicas = df['NOTA FISCAL'].unique()

    # Converte a lista de notas para uma lista de strings para digitação.
    lista_de_notas = [str(nota) for nota in notas_unicas]

    print(f"Arquivo CSV '{nome_do_arquivo_csv}' carregado com sucesso.")
    print(f"{len(lista_de_notas)} notas fiscais ÚNICAS foram encontradas para automação.")

except FileNotFoundError:
    print(
        f"ERRO CRÍTICO: Planilha '{nome_do_arquivo_csv}' não encontrada. Verifique se o nome está correto e se o arquivo está na mesma pasta do script.")
    exit()  # Encerra o script se não encontrar a planilha
except KeyError:
    print(
        f"ERRO CRÍTICO: A coluna 'NOTA FISCAL' não foi encontrada na planilha. Verifique se o nome da coluna está correto no arquivo CSV.")
    exit()  # Encerra o script se a coluna não existir

# --- INÍCIO DA AUTOMAÇÃO ---
print("\nAutomação iniciando em 5 segundos...")
print("Por favor, deixe a janela do sistema NEO em primeiro plano e não mexa no mouse ou teclado.")
time.sleep(5)

# Loop principal que passa por cada nota da sua lista de notas únicas
for numero_da_nota_atual in lista_de_notas:
    print(f"\n==================================================")
    print(f"--- Processando a Nota Fiscal: {numero_da_nota_atual} ---")
    print(f"==================================================")

    try:
        pyautogui.doubleClick(131, 67, duration=0.25)
        time.sleep(0.1)
        pyautogui.write(numero_da_nota_atual, interval=0.05)
        print(f"OK - Número da nota '{numero_da_nota_atual}' digitado.")
        time.sleep(0.25)
        pyautogui.press('enter', presses=1, interval=0.5)
        print("OK - ENTER para buscar.")
        time.sleep(0.5)

        pyautogui.press('enter', presses=3, interval=0.5)
        print("OK - ENTER 3x para fechar avisos.")
        time.sleep(0.25)

        pyautogui.click(508, 1015, duration=0.25)
        print("OK - Clique em imprimir.")
        time.sleep(0.25)

#VERIFICAR SE APARECE MAIS DE UMA JANELA

        pyautogui.press('enter', presses=1, interval=0.5)
        print("OK - ENTER 1x para fechar aviso de certificado.")
        time.sleep(0.5)  # Tempo crucial para o sistema buscar a nota

        pyautogui.click(330, 1009, duration=0.25)
        print("OK - Clique em imprimir. ")
        time.sleep(1.5)

        pyautogui.click(83, 38, duration=0.25)
        time.sleep(1)
        pyautogui.press('enter')
        print("OK - Comando para gerar PDF enviado.")
        time.sleep(1)  # Tempo MUITO IMPORTANTE para a janela "Salvar como" do Windows abrir

        nome_arquivo = f"DANFE {numero_da_nota_atual}"
        print(f"OK - Salvando o arquivo como: {nome_arquivo}")
        pyautogui.write(nome_arquivo, interval=0.05)
        time.sleep(0.25)
        pyautogui.press('enter')
        time.sleep(1)  # Tempo para o arquivo ser salvo fisicamente no disco. Aumente se os arquivos forem grandes.

        print("OK - Voltando para a tela inicial...")
        pyautogui.press('esc', presses=2, interval=0.5)
        time.sleep(1)


    except (
            Exception) as e:
        print(f"!!!!!!!!!! ERRO INESPERADO !!!!!!!!!!")
        print(f"Ocorreu um erro grave ao processar a nota {numero_da_nota_atual}: {e}")
        print("Tentando se recuperar e ir para a próxima nota.")
        pyautogui.press('esc', presses=4, interval=0.5)  # Tenta fechar tudo
        continue

print("\n\n--- FIM DA AUTOMAÇÃO ---")
print("Todos os itens da lista foram processados.")

