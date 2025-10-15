import pyautogui
import keyboard
import time
import csv
from datetime import datetime
import sys
import os


# flag para parar a macro
parar_execucao = False

# Garante que a pasta "LOGs" existe
os.makedirs("LOGs", exist_ok=True)

# Nome do arquivo de log dinâmico dentro da pasta LOGs
nome_arquivo_log = os.path.join("LOGs", f"log_{datetime.now().strftime('%Y.%m.%d_%H.%M')}.txt")

# Inicializa o arquivo de log dinâmico para registrar todas as ações e mensagens durante a execução da macro.
def registrar_log(msg):
    # Registra mensagem no console e salva imediatamente no log
    horario_atual = datetime.now().strftime('%H:%M:%S')
    msg = f"[{horario_atual}] {msg}"
    print(msg)
    with open(nome_arquivo_log, "a", encoding="utf-8") as f:
        f.write(msg + "\n")

def escutador_tecla():
    # Escutador para interromper macro ao pressionar "F9"
    global parar_execucao
    if keyboard.is_pressed("F9"):
        os.system('cls' if os.name == 'nt' else 'clear')
        registrar_log("")
        registrar_log("[§§] Tecla 'F9' pressionada → Interrompendo execução...")
        parar_execucao = True

def main():
    registrar_log("Pressione F12 para iniciar a macro. Pressione F9 para parar.")
    keyboard.wait("F12")

    while True:  # loop para escolher/recarregar CSV
        caminho_csv = input("Digite o nome do arquivo CSV (sem a extensão .csv): ").strip() + ".csv"
        registrar_log(f"Arquivo CSV definido como: {caminho_csv}")
        registrar_log("")

        ArqCSV = carregarCsv(caminho_csv)
        if ArqCSV is None:
            registrar_log("Falha ao carregar o CSV. Tente novamente.")
            continue  # volta para pedir o CSV novamente

        # Menu de módulos — loop até o usuário pedir para voltar ao menu de CSV
        while True:
            registrar_log("Selecione o módulo:")
            registrar_log("[f] = Preencher faixas horárias")
            registrar_log("[p] = Gerar documentos (PDF) de Quadro e Faixa horária")
            registrar_log("[*] = Voltar para seleção de arquivo")
            registrar_log("")

            seletor = input("Módulo: ").strip().lower()

            if seletor == "f":
                pParametros(ArqCSV)
            elif seletor == "p":
                imprimir_FQH(ArqCSV)
            elif seletor == "*":
                registrar_log("Voltando para seleção de arquivo...")
                break  # sai do menu e retorna ao loop externo (escolher CSV)
            else:
                print("Opção inválida")

def carregarCsv(caminho_csv):
    try:
        with open(caminho_csv, newline="", encoding="utf-8") as csvfile:
            leitor = csv.reader(csvfile, delimiter=";")

            # Lê o cabeçalho (primeira linha)
            cabecalho = next(leitor)
            registrar_log(f"Cabeçalho: {cabecalho}")

            # Converte o restante das linhas em lista
            linhas = list(leitor)

            registrar_log(f"Arquivo '{caminho_csv}' carregado com sucesso. Total de registros: {len(linhas)}")
            registrar_log("")
            return linhas

    except FileNotFoundError:
        registrar_log(f"Arquivo '{caminho_csv}' não encontrado. Verifique o nome e tente novamente.")
        registrar_log("")
        return None

    except Exception as e:
        registrar_log(f"Erro ao ler o arquivo CSV: {e}")
        registrar_log("")
        return None

def imprimir_FQH(ArqCSV):
    global parar_execucao

    registrar_log("Módulo de [Impressão de Faixas e Quadros] iniciado")

    # Tempo para o usuário se preparar
    for _ in range(5):
        if parar_execucao: break
        registrar_log(f"Iniciando em {5 - _} segundos...")
        time.sleep(1)

    # Loop principal
    for i, (tipoOso, numOso, nomePDF) in enumerate(ArqCSV):
        escutador_tecla()
        if parar_execucao: break

        registrar_log("")
        registrar_log("--------------------------------------------------")
        registrar_log(f"[§§] Processando OSO: {numOso}") 

        # Digita OSO
        registrar_log(f"[§§] Digitando OSO [{numOso}]")
        pyautogui.write(numOso, interval=0.1)
        time.sleep(0.5)

        # Vai para Qualidade gráfica
        registrar_log(f"[{numOso}] TAB           | Navega até Qualidade gráfica")
        pyautogui.press("tab")
        time.sleep(0.3)

        # Verifica derivadas
        if tipoOso == "B" and i + 1 < len(ArqCSV):
            _, proximoNumOso, _ = ArqCSV[i + 1]
            if numOso[:4] == proximoNumOso[:4]:
                registrar_log(f"[§§] Oso [{numOso}] possui DERIVADAS")
                registrar_log(f"[§§] Fechando pop-up")
                pyautogui.press("tab")
                registrar_log(f"[{numOso}] TAB           | Navega até 'não' em ver osos filhas")
                pyautogui.press("enter")
                registrar_log(f"[{numOso}] ENTER         | Confirma seleção")
                time.sleep(0.3)

        # Configurar -> Imprimir
        for _ in range(2):
            escutador_tecla()
            if parar_execucao: break
            pyautogui.press("tab")
            registrar_log(f"[{numOso}] TAB X {_+1}       | Navega até Imprimir")
        if parar_execucao: break
        time.sleep(0.3)

        pyautogui.press("enter")
        registrar_log(f"[{numOso}] ENTER         | Confirma imprimir")
        time.sleep(4)

        # Digitar nome do PDF
        pyautogui.write(nomePDF, interval=0.1)
        registrar_log(f"[§§] Digitando nome do arquivo: {nomePDF}")
        time.sleep(0.5)

        pyautogui.press("enter")
        registrar_log(f"[{numOso}] ENTER         | Confirma salvar")
        time.sleep(4)

        # Voltar até campo OSO
        for _ in range(3):
            escutador_tecla()
            if parar_execucao: break
            pyautogui.hotkey("shift", "tab")
            registrar_log(f"[{numOso}] SHIFT+TAB X {_+1}  | Volta até 'Informe o Nº da OSO'")
            time.sleep(0.3)
        if parar_execucao: break

        # Apagar 6 dígitos
        for _ in range(6):
            escutador_tecla()
            if parar_execucao: break
            pyautogui.press("backspace")
            registrar_log(f"[{numOso}] BACKSPACE X {_+1} | Apaga 6 dígitos")
            time.sleep(0.3)
        if parar_execucao: break

        registrar_log(f"[§§] Arquivo '{nomePDF}' gerado com sucesso.")
        registrar_log("")

def pParametros(arquivoCSV):

    registrar_log("Módulo de [Preenchimento de parâmetros] iniciado")

    # Tempo para o usuário se preparar
    for _ in range(5):
        if parar_execucao: break
        registrar_log(f"Iniciando em {5 - _} segundos...")
        time.sleep(1)
    
    # Loop principal
    for i, (faixaIni, faixaFim, interv, tPerc, tTerm, frota, linha, dia) in enumerate(arquivoCSV):
        escutador_tecla()
        if parar_execucao: break

        registrar_log("")
        registrar_log("--------------------------------------------------")
        registrar_log(f"[§§] Processando a linha: {linha} - Dia: {dia} - Linha {i + 1} de {len(arquivoCSV)}")
        registrar_log("")
        time.sleep(1)

        # Formatação dos horários
        def formatar_hora(Modo, h):
            h = h.strip()
            if Modo == "log":
                if len(h) == 4:
                    return f"{h[:2]}:{h[2:]}"
                elif len(h) == 3:
                    return f"0{h[:1]}:{h[1:]}"
                return h
            elif Modo == "dig":
                if len(h) == 3:
                    return f"0{h}"
                return h

        faixaIni_ = formatar_hora("log", faixaIni)
        faixaFim_ = formatar_hora("log", faixaFim)
        faixaFim = formatar_hora("dig", faixaFim)
        faixaIni = formatar_hora("dig", faixaIni)

        # Preenchimento dos campos
        #Preenche o campo Faixa Inicial
        pyautogui.write(faixaIni)
        registrar_log(f"[{linha}] Horário inicial {faixaIni_} preenchido.")
    
        # Move para o campo horário final
        pyautogui.press("tab")
        registrar_log(f"[{linha}] Movendo para o campo horário final.")
        time.sleep(0.1)

        #Preenche o campo Faixa horária final
        pyautogui.write(faixaFim)
        registrar_log(f"[{linha}] Horário inicial {faixaFim_} preenchido.")

        # Move para o campo intervalo
        pyautogui.press("tab")
        time.sleep(0.1)
        registrar_log(f"[{linha}] Movendo para o campo intervalo.")
        time.sleep(0.1)

        # Preenche o campo intervalo
        pyautogui.write(interv)
        registrar_log(f"[{linha}] Intervalo preenchido.")
        time.sleep(0.1)

        # Move para o campo tempo de percurso
        pyautogui.press("tab")
        time.sleep(0.1)
        registrar_log(f"[{linha}] Movendo para o campo tempo de percurso.")

        # Preenche o campo tempo de percurso
        pyautogui.write(tPerc)
        registrar_log(f"[{linha}] Tempo de percurso preenchido.")
        time.sleep(0.1)

        # Move para o campo Terminal
        pyautogui.press("tab")
        registrar_log(f"[{linha}] Movendo para o campo Terminal")
        time.sleep(0.1)

        # Preenche o campo Terminal
        pyautogui.write(tTerm)
        registrar_log(f"[{linha}] Terminal preenchido.")
        time.sleep(0.1)

        # Move para o campo Frota
        pyautogui.press("tab")
        registrar_log(f"[{linha}] Movendo para o campo Frota")
        time.sleep(0.1)

        # Preenche o campo Frota
        pyautogui.write(frota)
        registrar_log(f"[{linha}] Frota preenchida.")
        time.sleep(0.1)

        # Move para o campo Linha
        pyautogui.press("tab")
        registrar_log(f"[{linha}] Movendo para o campo Linha")
        time.sleep(0.1)

        # Aguarda um momento para o usuário verificar os dados antes de salvar
        registrar_log(f"[{linha}] Aguardando verificação de erro...")
        
        #Verifica se o próximo dia ou linha são diferentes do atual
        if i + 1 < len(arquivoCSV):
            _, _, _, _, _, _, proximaLinha, proximoDia = arquivoCSV[i + 1]
            if dia != proximaLinha:
                registrar_log("")
                registrar_log("=======================================================")
                registrar_log(f"             Fim da faixa horária de {linha}")
                registrar_log("=======================================================")
                registrar_log("")
                registrar_log(f"[{linha}] Aguardando a mudança de dia no SIGET...")
                registrar_log(f"[{linha}] Pressione ENTER para o preenchimento da faixa de [{proximaLinha}] [{proximoDia}]")
                keyboard.wait("f10")
            elif linha != proximoDia:
                registrar_log("")
                registrar_log("=======================================================")
                registrar_log(f"             Fim da faixa horária de {dia}")
                registrar_log("=======================================================")
                registrar_log("")
                registrar_log(f"[{linha}] Aguardando a mudança de linha no SIGET...")
                registrar_log(f"[{linha}] Pressione ENTER para o preenchimento da linha [{proximaLinha}] [{proximoDia}]")
                keyboard.wait("f10")
            else:
                registrar_log(f"[{linha}] Pressione F10 para a próxima faixa")
                keyboard.wait("f10")
        time.sleep(0.5)
        escutador_tecla()
        if parar_execucao: break

instrucoes = """
[ Instruções de Uso da Macro SIGET ]

| Macro para automatizar impressão de
| Ordens de Serviço (OSO) no SIGET.
| Utiliza pyautogui para automação de
| interface e keyboard para controle
| de teclado.

Requisitos:
| Definir a impressora padrão como
| "Microsoft Print to PDF".
| Fazer login no SIGET.
| Deixa no SIGET o campo "Informe o Nº
| da OSO" selecionado.
| Definir o Excel com a aba da planilha
| correta aberta ("FH" ou "QH").

Instruções:
| Pressione F12 para iniciar.
| Pressione P a qualquer momento para
| interromper e salvar o log.
"""

print(instrucoes)
main()