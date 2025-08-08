import pyautogui
import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

# --- CONFIGURAÇÕES E PARÂMETROS ---

# Medida de segurança: Leve o mouse para o canto superior esquerdo para parar o script.
pyautogui.FAILSAFE = True

# Coordenadas dos cliques (ajuste aqui se a janela ou os botões mudarem de lugar)
CAMPO_NOTA_FISCAL = (131, 67)
BOTAO_IMPRIMIR_TELA_PRINCIPAL = (508, 1015)
BOTAO_IMPRIMIR_TELA_SECUNDARIA = (330, 1009)
ICONE_GERAR_PDF = (83, 38)


# --- FUNÇÕES DE AUTOMAÇÃO ---

def selecionar_arquivo(titulo):
    """Abre uma janela para o usuário selecionar um arquivo."""
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal do tkinter

    caminho_arquivo = filedialog.askopenfilename(
        title=titulo,
        filetypes=(("Arquivos de Planilha", "*.csv *.xlsx"), ("Todos os arquivos", "*.*"))
    )
    return caminho_arquivo


def selecionar_pasta_destino():
    """Abre uma janela para o usuário selecionar a pasta onde os PDFs serão salvos."""
    print("Por favor, selecione a pasta onde os PDFs serão salvos.")
    root = tk.Tk()
    root.withdraw()
    caminho_pasta = filedialog.askdirectory(
        title="Selecione a pasta de destino para os PDFs"
    )
    return caminho_pasta


def carregar_dados_notas(caminho_arquivo):
    """Lê o arquivo CSV ou Excel, trata os dados e retorna uma lista de notas únicas."""
    if not caminho_arquivo:
        print("Nenhum arquivo de planilha selecionado.")
        return None

    try:
        # Verifica a extensão do arquivo para usar a função de leitura correta
        if caminho_arquivo.endswith('.csv'):
            df = pd.read_csv(caminho_arquivo, encoding='latin-1', delimiter=';')
        elif caminho_arquivo.endswith('.xlsx'):
            df = pd.read_excel(caminho_arquivo)
        else:
            print(f"ERRO: Formato de arquivo não suportado: {caminho_arquivo}")
            return None

        df.columns = df.columns.str.strip()
        notas_unicas = df['DOCUMENTO'].unique()
        lista_de_notas = [str(int(nota)) for nota in notas_unicas if pd.notna(nota)]

        print(f"Arquivo '{os.path.basename(caminho_arquivo)}' carregado com sucesso.")
        print(f"{len(lista_de_notas)} notas fiscais ÚNICAS foram encontradas para automação.")
        return lista_de_notas

    except FileNotFoundError:
        print(f"ERRO CRÍTICO: Planilha '{caminho_arquivo}' não encontrada.")
        return None
    except KeyError:
        print(f"ERRO CRÍTICO: A coluna 'DOCUMENTO' não foi encontrada na planilha.")
        return None
    except Exception as e:
        print(f"Ocorreu um erro inesperado ao ler a planilha: {e}")
        return None


def processar_uma_nota(numero_nota, pasta_destino):
    """Executa a sequência completa de automação para um único número de nota fiscal."""

    # 1. Digita o número da nota e busca
    pyautogui.doubleClick(CAMPO_NOTA_FISCAL, duration=0.25)
    time.sleep(0.1)
    pyautogui.write(numero_nota, interval=0.05)
    print(f"OK - Número da nota '{numero_nota}' digitado.")
    time.sleep(0.25)
    pyautogui.press('enter')
    print("OK - ENTER para buscar.")
    time.sleep(0.5)

    # 2. Fecha possíveis avisos do sistema
    pyautogui.press('enter', presses=3, interval=0.5)
    print("OK - ENTER 3x para fechar avisos.")
    time.sleep(0.25)

    # 3. Navega para a tela de impressão
    pyautogui.click(BOTAO_IMPRIMIR_TELA_PRINCIPAL, duration=0.25)
    print("OK - Clique para tela de impressão.")
    time.sleep(0.25)

    # 4. Fecha aviso de certificado e clica em imprimir novamente
    pyautogui.press('enter')
    print("OK - ENTER para fechar aviso de certificado.")
    time.sleep(0.5)

    pyautogui.click(BOTAO_IMPRIMIR_TELA_SECUNDARIA, duration=0.25)
    print("OK - Clique final em imprimir.")
    time.sleep(1.5)

    # 5. Gera o PDF e salva o arquivo
    pyautogui.click(ICONE_GERAR_PDF, duration=0.25)
    time.sleep(1)
    pyautogui.press('enter')
    print("OK - Comando para gerar PDF enviado.")
    time.sleep(1)

    nome_arquivo = f"DANFE {numero_nota}"
    print(f"OK - Salvando o arquivo como: {nome_arquivo}")
    pyautogui.write(nome_arquivo, interval=0.05)
    time.sleep(0.25)
    pyautogui.press('enter')
    time.sleep(1)

    # 6. VERIFICAÇÃO DE SALVAMENTO
    nome_arquivo_pdf = nome_arquivo + ".pdf"
    caminho_completo_arquivo = os.path.join(pasta_destino, nome_arquivo_pdf)

    print(f"Verificando se o arquivo '{nome_arquivo_pdf}' foi salvo...")
    verificado = False
    for _ in range(5):  # Tenta verificar por 5 segundos
        if os.path.exists(caminho_completo_arquivo):
            print(f"SUCESSO: Arquivo confirmado na pasta de destino.")
            verificado = True
            break
        time.sleep(1)

    if not verificado:
        print(f"FALHA: Não foi possível confirmar o salvamento do arquivo '{nome_arquivo_pdf}'.")

    # 7. Volta para a tela inicial
    print("OK - Voltando para a tela inicial...")
    pyautogui.press('esc', presses=2, interval=0.5)
    time.sleep(1)


# --- SCRIPT PRINCIPAL ---

def main():
    """Função principal que orquestra a automação."""

    caminho_planilha = selecionar_arquivo("Selecione a PLANILHA de notas fiscais")
    if not caminho_planilha:
        print("Nenhuma planilha foi selecionada. Encerrando o programa.")
        return

    pasta_destino_pdf = selecionar_pasta_destino()
    if not pasta_destino_pdf:
        print("Nenhuma pasta de destino foi selecionada. Encerrando o programa.")
        return

    lista_de_notas = carregar_dados_notas(caminho_planilha)
    if lista_de_notas is None:
        print("Automação não pode continuar devido a erro no carregamento dos dados.")
        return

    print("\nAutomação iniciando em 5 segundos...")
    print("Por favor, deixe a janela do sistema NEO em primeiro plano e não mexa no mouse ou teclado.")
    time.sleep(5)

    for numero_da_nota_atual in lista_de_notas:
        print(f"\n==================================================")
        print(f"--- Processando a Nota Fiscal: {numero_da_nota_atual} ---")
        print(f"==================================================")

        try:
            processar_uma_nota(numero_da_nota_atual, pasta_destino_pdf)

        except Exception as e:
            print(f"!!!!!!!!!! ERRO INESPERADO !!!!!!!!!!")
            print(f"Ocorreu um erro grave ao processar a nota {numero_da_nota_atual}: {e}")
            print("Tentando se recuperar e ir para a próxima nota.")
            pyautogui.press('esc', presses=4, interval=0.5)
            continue

    print("\n\n--- FIM DA AUTOMAÇÃO ---")
    print("Todos os itens da lista foram processados.")


if __name__ == "__main__":
    main()
