import pyautogui
import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import sys
import threading
import keyboard
import json
import re

# --- CONFIGURAÇÕES E PARÂMETROS ---

# Flag global para controlar o estado da automação
automacao_ativa = True

# Medida de segurança: Leve o mouse para o canto superior esquerdo para parar o script.
pyautogui.FAILSAFE = True

# Coordenadas dos cliques (ajustadas por você)
CAMPO_NOTA_FISCAL = (260, 105)
BOTAO_MOSTRAR = (90, 1010)
BOTAO_IMPRIMIR = (325, 1010)
BOTAO_GERAR_PDF = (85, 40)
BOTAO_OK_PDF = (315, 760)
BOTAO_SALVAR_PDF = (1610, 810)
BOTAO_FECHAR_NF = (615, 40)
BOTAO_OK_AVISO = (960, 550)
BOTAO_GERAR_XML = (575, 1010)

# Arquivo para guardar as últimas seleções de pasta
NOME_ARQUIVO_CONFIG = 'config_automacao.json'


# --- FUNÇÕES DE AUTOMAÇÃO E KILLSWITCH ---

def monitorar_killswitch():
    """Função que roda em segundo plano para ouvir a tecla 'espaço'."""
    global automacao_ativa
    keyboard.wait('space')
    print("\n!!!!!!!!!! KILLSWITCH ATIVADO !!!!!!!!!!")
    print("A automação será interrompida de forma segura...")
    automacao_ativa = False


def executar_acao(acao, *args, **kwargs):
    """
    Executa uma ação do pyautogui ou time.sleep, mas antes verifica o killswitch.
    Lança uma exceção se o killswitch for ativado.
    """
    if not automacao_ativa:
        raise KeyboardInterrupt("Killswitch ativado pelo usuário.")

    # Executa a ação passada como argumento
    acao(*args, **kwargs)


def carregar_config():
    """Carrega as últimas pastas salvas do arquivo de configuração."""
    if os.path.exists(NOME_ARQUIVO_CONFIG):
        with open(NOME_ARQUIVO_CONFIG, 'r') as f:
            try:
                return json.load(f)
            except json.JSONDecodeError:
                return {}  # Retorna um dicionário vazio se o arquivo estiver corrompido
    return {}


def salvar_config(config):
    """Salva as configurações (caminhos de pasta) em um arquivo JSON."""
    with open(NOME_ARQUIVO_CONFIG, 'w') as f:
        json.dump(config, f, indent=4)


def selecionar_arquivo(titulo, ultimo_caminho):
    """Abre uma janela para o usuário selecionar um arquivo, começando no último caminho usado."""
    root = tk.Tk()
    root.withdraw()
    caminho_arquivo = filedialog.askopenfilename(
        title=titulo,
        initialdir=ultimo_caminho,
        filetypes=(("Arquivos de Planilha", "*.csv *.xlsx"), ("Todos os arquivos", "*.*"))
    )
    return caminho_arquivo


def selecionar_pasta_destino(ultimo_caminho):
    """Abre uma janela para o usuário selecionar a pasta de destino, começando no último caminho usado."""
    print("Por favor, selecione a pasta onde os PDFs serão salvos.")
    root = tk.Tk()
    root.withdraw()
    caminho_pasta = filedialog.askdirectory(
        title="Selecione a pasta de destino para os PDFs",
        initialdir=ultimo_caminho
    )
    return caminho_pasta


def carregar_dados_notas(caminho_arquivo):
    """Lê o arquivo CSV ou Excel, trata os dados e retorna uma lista de notas únicas."""
    if not caminho_arquivo:
        print("Nenhum arquivo de planilha selecionado.")
        return None
    try:
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
    except Exception as e:
        print(f"ERRO ao carregar a planilha: {e}")
        return None


def limpar_pasta_destino(pasta_destino):
    """Apaga todos os arquivos .pdf da pasta de destino especificada."""
    print(f"Iniciando limpeza da pasta: {pasta_destino}")
    arquivos_apagados = 0
    try:
        for nome_arquivo in os.listdir(pasta_destino):
            if nome_arquivo.lower().endswith('.pdf'):
                caminho_completo = os.path.join(pasta_destino, nome_arquivo)
                try:
                    os.remove(caminho_completo)
                    print(f"Arquivo apagado: {nome_arquivo}")
                    arquivos_apagados += 1
                except Exception as e:
                    print(f"Não foi possível apagar o arquivo {nome_arquivo}. Erro: {e}")
    except FileNotFoundError:
        print(f"A pasta de destino '{pasta_destino}' não foi encontrada para limpeza.")
        return
    print(f"Limpeza concluída. {arquivos_apagados} arquivos .pdf foram apagados.")


def GERAR_NOVO_XML():
    """
    Executa a sequência de ações para corrigir o erro de XML.
    """
    print("ERRO DE XML DETECTADO! Iniciando tentativa de resolução automática...")

    executar_acao(pyautogui.click, BOTAO_OK_AVISO, duration=0.1)
    print("OK - Clique botão ok aviso.")
    executar_acao(time.sleep, 1)

    executar_acao(pyautogui.click, BOTAO_MOSTRAR, duration=0.1)
    print("OK - Clique botão mostrar.")
    executar_acao(time.sleep, 1)

    executar_acao(pyautogui.click, BOTAO_GERAR_XML, duration=0.1)
    print("OK - Clique botão gerar xml.")
    executar_acao(time.sleep, 3)

    executar_acao(pyautogui.click, BOTAO_OK_AVISO, duration=0.1)
    print("OK - Clique botão ok gerado com sucesso.")
    executar_acao(time.sleep, 1)

    executar_acao(pyautogui.click, BOTAO_MOSTRAR, duration=0.1)
    print("OK - Clique botão mostrar.")
    executar_acao(time.sleep, 1)

    print("Tentativa de geração de novo XML finalizada. Tentando salvar o PDF novamente...")


def processar_uma_nota(numero_nota):
    """
    Executa a sequência de cliques para gerar e salvar o PDF.
    Esta função representa uma única tentativa de processamento.
    """
    executar_acao(pyautogui.doubleClick, CAMPO_NOTA_FISCAL, duration=0.1)
    print("OK - Campo Nota Fiscal selecionado.")
    executar_acao(time.sleep, 0.1)

    executar_acao(pyautogui.write, numero_nota, interval=0.01)
    print(f"OK - Número da nota '{numero_nota}' digitado.")
    executar_acao(time.sleep, 0.1)

    executar_acao(pyautogui.click, BOTAO_MOSTRAR, duration=0.1)
    print("OK - Clique botão mostrar.")
    executar_acao(time.sleep, 0.25)

    executar_acao(pyautogui.click, BOTAO_IMPRIMIR, duration=0.1)
    print("OK - Clique botão imprimir.")
    executar_acao(time.sleep, 0.5)

    executar_acao(pyautogui.click, BOTAO_GERAR_PDF, duration=0.1)
    print("OK - Clique no botão PDF.")
    executar_acao(time.sleep, 0.5)

    executar_acao(pyautogui.click, BOTAO_OK_PDF, duration=0.1)
    print("OK - Clique no botão OK.")
    executar_acao(time.sleep, 0.5)

    nome_arquivo = f"DANFE_{numero_nota}"
    executar_acao(pyautogui.write, nome_arquivo, interval=0.01)
    print(f"OK - Tentando salvar arquivo como: {nome_arquivo}")
    executar_acao(time.sleep, 0.2)

    executar_acao(pyautogui.click, BOTAO_SALVAR_PDF, duration=0.1)
    print("OK - Clique salvar pdf.")
    executar_acao(time.sleep, 1)


# --- SCRIPT PRINCIPAL ---

def main():
    """Função principal que orquestra a automação."""

    killswitch_thread = threading.Thread(target=monitorar_killswitch, daemon=True)
    killswitch_thread.start()

    config = carregar_config()
    ultimo_caminho_planilha = config.get("ultimo_caminho_planilha")
    ultimo_caminho_destino = config.get("ultimo_caminho_destino")

    caminho_planilha = selecionar_arquivo("Selecione a PLANILHA de notas fiscais", ultimo_caminho_planilha)
    if not caminho_planilha:
        print("Nenhuma planilha foi selecionada. Encerrando o programa.")
        return
    config["ultimo_caminho_planilha"] = os.path.dirname(caminho_planilha)
    salvar_config(config)

    pasta_destino_pdf = selecionar_pasta_destino(ultimo_caminho_destino)
    if not pasta_destino_pdf:
        print("Nenhuma pasta de destino foi selecionada. Encerrando o programa.")
        return
    config["ultimo_caminho_destino"] = pasta_destino_pdf
    salvar_config(config)

    lista_completa_notas = carregar_dados_notas(caminho_planilha)
    if lista_completa_notas is None:
        print("Automação não pode continuar devido a erro no carregamento dos dados.")
        return

    # --- LÓGICA DE RETOMADA ---
    notas_a_processar = lista_completa_notas
    try:
        arquivos_na_pasta = os.listdir(pasta_destino_pdf)
        pdfs_existentes = [f for f in arquivos_na_pasta if f.lower().endswith('.pdf')]

        if pdfs_existentes:
            print(f"\nAVISO: A pasta de destino já contém {len(pdfs_existentes)} arquivos PDF.")
            resposta = pyautogui.confirm(
                text='A pasta de destino não está vazia. Como deseja proceder?',
                title='Retomar Automação?',
                buttons=['Retomar trabalho (pular notas já salvas)', 'Começar do zero (apagar PDFs existentes)']
            )
            if resposta == 'Começar do zero (apagar PDFs existentes)':
                confirmacao = pyautogui.confirm(
                    text=f'TEM CERTEZA?\nIsso irá apagar TODOS os {len(pdfs_existentes)} arquivos .pdf da pasta:\n\n{pasta_destino_pdf}',
                    title='CONFIRMAÇÃO FINAL',
                    buttons=['Sim, apagar e começar do zero', 'Cancelar']
                )
                if confirmacao == 'Sim, apagar e começar do zero':
                    limpar_pasta_destino(pasta_destino_pdf)
                else:
                    print("Operação cancelada pelo usuário.")
                    return
            else:  # Retomar trabalho
                print("Modo de retomada ativado. Verificando notas já salvas...")
                notas_ja_salvas = set()
                for pdf in pdfs_existentes:
                    match = re.search(r'\d+', pdf)
                    if match:
                        notas_ja_salvas.add(match.group(0))

                notas_a_processar = [nota for nota in lista_completa_notas if nota not in notas_ja_salvas]
                print(
                    f"{len(notas_ja_salvas)} notas já salvas foram encontradas. {len(notas_a_processar)} notas restantes para processar.")

    except Exception as e:
        print(f"Ocorreu um erro ao verificar a pasta de destino: {e}")
        return

    if not notas_a_processar:
        print("\nNão há novas notas para processar. Todas as notas da planilha já foram salvas.")
        return

    print("\nAutomação iniciando em 3 segundos...")
    print(">>> Pressione a tecla 'ESPAÇO' a qualquer momento para parar. <<<")
    time.sleep(3)

    notas_com_falha = []
    interrompido_pelo_usuario = False
    erro_fatal = False

    for numero_da_nota_atual in notas_a_processar:
        if not automacao_ativa:
            interrompido_pelo_usuario = True
            break

        print(f"\n==================================================")
        print(f"--- Processando a Nota Fiscal: {numero_da_nota_atual} ---")
        print(f"==================================================")

        try:
            # --- NOVA LÓGICA DE TENTATIVA E RETENTATIVA ---

            # 1. PRIMEIRA TENTATIVA
            print("--- Iniciando primeira tentativa de salvamento...")
            processar_uma_nota(numero_da_nota_atual)

            # 2. PRIMEIRA VERIFICAÇÃO
            nome_arquivo_pdf = f"DANFE_{numero_da_nota_atual}.pdf"
            caminho_completo_arquivo = os.path.join(pasta_destino_pdf, nome_arquivo_pdf)
            verificado = os.path.exists(caminho_completo_arquivo)

            # 3. CORREÇÃO E RETENTATIVA (SE NECESSÁRIO)
            if not verificado:
                print("Arquivo não encontrado na primeira tentativa. Executando rotina de correção...")
                GERAR_NOVO_XML()

                print("--- Iniciando segunda tentativa de salvamento...")
                processar_uma_nota(numero_da_nota_atual)

                # VERIFICAÇÃO FINAL
                print("Verificando novamente se o arquivo foi salvo...")
                for _ in range(5):
                    if os.path.exists(caminho_completo_arquivo):
                        verificado = True
                        break
                    time.sleep(0.5)

            # 4. PEDIR AJUDA MANUAL (SE TUDO FALHOU)
            if not verificado:
                mensagem_erro = (
                    f"FALHA CRÍTICA NO SALVAMENTO!\n\n"
                    f"O arquivo para a nota fiscal '{numero_da_nota_atual}' não foi encontrado, mesmo após tentativa automática de correção.\n\n"
                    f"AÇÃO MANUAL NECESSÁRIA:\n"
                    f"1. Salve o PDF manualmente.\n"
                    f"2. Feche as janelas extras do NEO.\n"
                    f"3. Deixe o programa na tela principal de busca.\n\n"
                    f"Clique em 'Retomar' para continuar com a PRÓXIMA nota."
                )
                resposta = pyautogui.confirm(text=mensagem_erro, title="Erro na Automação - Ação Necessária",
                                             buttons=['Retomar com a próxima nota', 'Encerrar Robô'])
                if resposta == 'Encerrar Robô':
                    raise SystemExit(f"Automação encerrada pelo usuário após falha na nota {numero_da_nota_atual}.")
                else:
                    print(f"AVISO: A nota {numero_da_nota_atual} foi pulada após intervenção manual.")
                    notas_com_falha.append(numero_da_nota_atual)
            else:
                print(f"SUCESSO: Arquivo confirmado na pasta de destino.")
                executar_acao(pyautogui.click, BOTAO_FECHAR_NF, duration=0.1)
                print(f"OK - Clique botão fechar NF.")
                executar_acao(time.sleep, 0.5)

        except KeyboardInterrupt:
            interrompido_pelo_usuario = True
            break
        except SystemExit as e:
            print(f"\n--- ERRO FATAL ---")
            print(e)
            notas_com_falha.append(numero_da_nota_atual)
            erro_fatal = True
            break
        except Exception as e:
            print(f"!!!!!!!!!! ERRO INESPERADO !!!!!!!!!!")
            print(f"Ocorreu um erro grave ao processar a nota {numero_da_nota_atual}: {e}")
            notas_com_falha.append(numero_da_nota_atual)
            pyautogui.press('esc', presses=3, interval=0.5)
            continue

    print("\n\n--- FIM DA AUTOMAÇÃO ---")
    if interrompido_pelo_usuario:
        print("A automação foi interrompida pelo usuário via Killswitch (ESPAÇO).")
    elif erro_fatal:
        print("A automação foi interrompida devido a uma falha crítica ou por escolha do usuário.")

    if notas_com_falha:
        print("As seguintes notas falharam ou foram puladas durante a execução:")
        for nota in notas_com_falha:
            print(f"- {nota}")
    elif not interrompido_pelo_usuario and not erro_fatal:
        print("Todos os itens da lista foram processados!")


if __name__ == "__main__":
    main()
