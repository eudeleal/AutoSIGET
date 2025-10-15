import pyautogui
import keyboard
import time
import csv
from datetime import datetime
import os

# flag para parar a macro
parar_execucao = False

# Garante que a pasta "LOGs" existe
os.makedirs("LOGs", exist_ok=True)

# Nome do arquivo de log dinâmico dentro da pasta LOGs
nome_arquivo_log = os.path.join("LOGs", f"log_{datetime.now().strftime('%Y.%m.%d_%H.%M')}.txt")

# Inicializa o arquivo de log dinâmico para registrar todas as ações e mensagens durante a execução da macro.
def registrarLog(msg):
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
        print("")
        print("[§§] Tecla 'F9' pressionada → Interrompendo execução...")
        parar_execucao = True

def main():
    print("\033[32mPressione F12 para iniciar a macro. Pressione F9 para parar.\033[m")
    keyboard.wait("F12")

    while True:  # loop para escolher/recarregar CSV
        caminho_csv = input("Digite o nome do arquivo CSV (sem a extensão .csv): ").strip() + ".csv"
        print(f"Arquivo CSV definido como: {caminho_csv}")
        print("")

        ArqCSV, cabecalho = carregarCsv(caminho_csv)

        if ArqCSV is None or cabecalho is None:
            print("Falha ao carregar o CSV. Tente novamente.")
            continue  # volta para pedir o CSV novamente

        # Menu de módulos — loop até o usuário pedir para voltar ao menu de CSV
        while True:
            
            print("Selecione o módulo:")
            print("[f] = Preencher faixas horárias")
            print("[p] = Gerar documentos (PDF) de Quadro e Faixa horária")
            print("[*] = Voltar para seleção de arquivo")
            print("")

            seletor = input("Módulo: ").strip().lower()

            if seletor == "f":
                if len(cabecalho) == 8:
                    pParametros(ArqCSV)
                else:
                    print("Colunas necessárias = 8")
                    print(f"CSV selecionado possui: {len(cabecalho)} colunas.")
                    print("")
            elif seletor == "p":
                if len(cabecalho) == 3:
                    imprimir_FQH(ArqCSV)
                else:
                    print("Colunas necessárias = 3")
                    print(f"CSV selecionado possui: {len(cabecalho)} colunas.")
                    print("")
            elif seletor == "*":
                print("Voltando para seleção de arquivo...")
                break  # sai do menu e retorna ao loop externo (escolher CSV)
            else:
                print("Opção inválida")
                print("")

def carregarCsv(caminho_csv):
    try:
        with open(caminho_csv, newline="", encoding="utf-8") as csvfile:
            leitor = csv.reader(csvfile, delimiter=";")

            # Lê o cabeçalho (primeira linha)
            cabecalho = next(leitor)
            print(f"Numero de colunas: {len(cabecalho)}")
            print(f"Cabeçalho: {cabecalho}")

            # Converte o restante das linhas em lista
            linhas = list(leitor)

            print(f"Arquivo '{caminho_csv}' carregado com sucesso. Total de registros: {len(linhas)}")
            print("")
            return linhas, cabecalho

    except FileNotFoundError:
        print(f"Arquivo '{caminho_csv}' não encontrado. Verifique o nome e tente novamente.")
        print("")
        return None, None

    except Exception as e:
        print(f"Erro ao ler o arquivo CSV: {e}")
        print("")
        return None, None

def imprimir_FQH(ArqCSV):
    global parar_execucao

    print("Módulo de [Impressão de Faixas e Quadros] iniciado")
    ososAtivas = "_"
    
    while ososAtivas == "_":
        ososAtivas = input("OSOs ativas? S/N:").upper().strip()
        if ososAtivas == "S":
            ososAtivas = True
            print("\033[32mOsos configuradas como ATIVAS\033[32m")
        elif ososAtivas == "N":
            ososAtivas = False
            print("\033[32mOsos configuradas como DESATIVADAS\033[32m")
        else:
            ososAtivas = "_"
            print("Reposta inválida")
            continue

    # Tempo para o usuário se preparar
        for _ in range(5):
            if parar_execucao: break
            print(f"Iniciando em {5 - _} segundos...")
            time.sleep(1)

    # Loop principal
    for i, (tipoOso, numOso, nomePDF) in enumerate(ArqCSV):
        escutador_tecla()
        if parar_execucao: break

        print("")
        print("-" * 50)
        print(f"[§§] Processando OSO: {numOso}") 

        # Digita OSO
        print(f"[§§] Digitando OSO [{numOso}]")
        pyautogui.write(numOso, interval=0.1)
        time.sleep(0.5)

        # Vai para Qualidade gráfica
        print(f"[{numOso}] TAB           | Navega até Qualidade gráfica")
        pyautogui.press("tab")
        time.sleep(0.3)

        # Verifica derivadas
        if ososAtivas == True and tipoOso == "B" and i + 1 < len(ArqCSV):
            _, proximoNumOso, _ = ArqCSV[i + 1]
            if numOso[:4] == proximoNumOso[:4]:
                print(f"[§§] Oso [{numOso}] possui DERIVADAS")
                print(f"[§§] Fechando pop-up")
                pyautogui.press("tab")
                print(f"[{numOso}] TAB           | Navega até 'não' em ver osos filhas")
                pyautogui.press("enter")
                print(f"[{numOso}] ENTER         | Confirma seleção")
                time.sleep(0.3)

        # Configurar -> Imprimir
        for _ in range(2):
            escutador_tecla()
            if parar_execucao: break
            pyautogui.press("tab")
            print(f"[{numOso}] TAB X {_+1}       | Navega até Imprimir")
        if parar_execucao: break
        time.sleep(0.3)

        pyautogui.press("enter")
        print(f"[{numOso}] ENTER         | Confirma imprimir")
        time.sleep(4)

        # Digitar nome do PDF
        pyautogui.write(nomePDF, interval=0.1)
        print(f"[§§] Digitando nome do arquivo: {nomePDF}")
        time.sleep(0.5)

        pyautogui.press("enter")
        print(f"[{numOso}] ENTER         | Confirma salvar")
        time.sleep(4)

        # Voltar até campo OSO
        for _ in range(3):
            escutador_tecla()
            if parar_execucao: break
            pyautogui.hotkey("shift", "tab")
            print(f"[{numOso}] SHIFT+TAB X {_+1}  | Volta até 'Informe o Nº da OSO'")
            time.sleep(0.3)
        if parar_execucao: break

        # Apagar 6 dígitos
        for _ in range(6):
            escutador_tecla()
            if parar_execucao: break
            pyautogui.press("backspace")
            print(f"[{numOso}] BACKSPACE X {_+1} | Apaga 6 dígitos")
            time.sleep(0.3)
        if parar_execucao: break

        print(f"[§§] Arquivo '{nomePDF}' gerado com sucesso.")
        print("")

def pParametros(arquivoCSV):
    global parar_execucao

    print("Módulo de [Preenchimento de parâmetros] iniciado")

    # Tempo para o usuário se preparar
    for _ in range(5):
        if parar_execucao: return
        print(f"Iniciando em {5 - _} segundos...")
        time.sleep(1)

    # ==========================================================
    # 🔹 Identificar blocos únicos (Linha + Dia)
    blocos = []
    linha_anterior = None
    dia_anterior = None
    inicio_bloco = 0

    for i, (faixaIni, faixaFim, interv, tPerc, tTerm, frota, linha, dia) in enumerate(arquivoCSV):
        linha = linha.strip()
        dia = dia.strip()

        if linha_anterior is None:
            linha_anterior = linha
            dia_anterior = dia
            inicio_bloco = i
            continue

        if linha != linha_anterior or dia != dia_anterior:
            blocos.append({
                "nome": f"{linha_anterior} {dia_anterior}",
                "inicio": inicio_bloco,
                "fim": i - 1
            })
            inicio_bloco = i
            linha_anterior = linha
            dia_anterior = dia

    # adiciona último bloco (se existir)
    if linha_anterior is not None:
        blocos.append({
            "nome": f"{linha_anterior} {dia_anterior}",
            "inicio": inicio_bloco,
            "fim": len(arquivoCSV) - 1
        })

    print("")
    print("📋 Blocos identificados:")
    for idx, b in enumerate(blocos, start=1):
        print(f"[{idx}] {b['nome']}: linhas {b['inicio'] + 1} a {b['fim'] + 1}")
    print("=" * 50)
    print("")

    # ==========================================================
    # 🔹 Pergunta ao usuário se deseja começar de um bloco específico
    while True:
        try:
            escolha = input(f"Digite o número do bloco para começar (1–{len(blocos)}) ou pressione ENTER para iniciar do primeiro: ").strip()
            if escolha == "":
                bloco_inicial = 1
                break
            bloco_inicial = int(escolha)
            if 1 <= bloco_inicial <= len(blocos):
                break
            else:
                print(f"Valor inválido. Escolha entre 1 e {len(blocos)}.")
        except ValueError:
            print("Entrada inválida. Digite apenas números.")

    print(f"Iniciando a partir do bloco {bloco_inicial}: {blocos[bloco_inicial-1]['nome']}")
    print("=" * 50)

    # ==========================================================
    # 🔹 Processa bloco por bloco (a partir da escolha)
    for bloco in blocos[bloco_inicial-1:]:
        print("")
        print("=" * 50)
        print(f"Iniciando preenchimento do bloco: {bloco['nome']}")
        print("=" * 50)

        # trecho contém linhas do índice inicio..fim (inclusive)
        trecho = arquivoCSV[bloco["inicio"] : bloco["fim"] + 1]

        # itera sobre o trecho; i_global é o índice real no arquivoCSV (0-based)
        for offset, (faixaIni, faixaFim, interv, tPerc, tTerm, frota, linha, dia) in enumerate(trecho):
            i_global = bloco["inicio"] + offset

            escutador_tecla()
            if parar_execucao:
                print("🟥 Execução interrompida pelo escutador.")
                return

            print("")
            print("-" * 50)
            print(f"[§§] Processando a linha: {linha.strip()} - Dia: {dia.strip()} - Linha {i_global + 1} de {len(arquivoCSV)}")
            print("")

            # --------------------------
            # Função auxiliar para formatar horários
            def formatar_hora(Modo, h):
                h = (h or "").strip()
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

            faixaIni_orig = faixaIni or ""
            faixaFim_orig = faixaFim or ""

            faixaIni_ = formatar_hora("log", faixaIni_orig)
            faixaFim_ = formatar_hora("log", faixaFim_orig)
            faixaFim = formatar_hora("dig", faixaFim_orig)
            faixaIni = formatar_hora("dig", faixaIni_orig)

            # --------------------------
            # Preenchimento dos campos
            pyautogui.write(faixaIni)
            print(f"[{linha.strip()}] Horário inicial {faixaIni_} preenchido.")

            pyautogui.press("tab")
            print(f"[{linha.strip()}] Movendo para o campo horário final.")

            # 🔹 CORREÇÃO: Detecta virada de dia comparando apenas dentro do mesmo bloco
            try:
                if offset > 0:  # Usa offset (posição dentro do bloco) em vez de i_global
                    faixaIniAnterior = (trecho[offset - 1][0] or "").strip()
                    if faixaIniAnterior.isdigit() and faixaIni.isdigit():
                        if int(faixaIni) < int(faixaIniAnterior):
                            print(f"[{linha.strip()}] Faixa inicial indica virada de dia → Confirmando continuação no dia seguinte.")
                            time.sleep(0.3)
                            pyautogui.press("enter")
                            time.sleep(0.5)
            except Exception as e:
                print(f"[{linha.strip()}] (Aviso) Falha ao verificar virada de dia em faixaIni: {e}")

            pyautogui.write(faixaFim)
            print(f"[{linha.strip()}] Horário final {faixaFim_} preenchido.")

            pyautogui.press("tab")
            print(f"[{linha.strip()}] Movendo para o campo intervalo.")

            # 🔹 CORREÇÃO: Detecta virada de dia no horário final apenas dentro do mesmo bloco
            try:
                if offset > 0:  # Usa offset (posição dentro do bloco) em vez de i_global
                    faixaFimAnterior = (trecho[offset - 1][1] or "").strip()
                    if faixaFimAnterior.isdigit() and faixaFim.isdigit():
                        if int(faixaFim) < int(faixaFimAnterior):
                            print(f"[{linha.strip()}] Faixa final indica virada de dia → Confirmando continuação no dia seguinte.")
                            time.sleep(0.3)
                            pyautogui.press("enter")
                            time.sleep(0.5)
            except Exception as e:
                print(f"[{linha.strip()}] (Aviso) Falha ao verificar virada de dia em faixaFim: {e}")

            pyautogui.write(interv)
            print(f"[{linha.strip()}] Intervalo preenchido.")

            pyautogui.press("tab")
            print(f"[{linha.strip()}] Movendo para o campo tempo de percurso.")

            pyautogui.write(tPerc)
            print(f"[{linha.strip()}] Tempo de percurso preenchido.")

            pyautogui.press("tab")
            print(f"[{linha.strip()}] Movendo para o campo Terminal")

            pyautogui.write(tTerm)
            print(f"[{linha.strip()}] Terminal preenchido.")

            pyautogui.press("tab")
            print(f"[{linha.strip()}] Movendo para o campo Frota")

            pyautogui.write(frota)
            print(f"[{linha.strip()}] Frota preenchida.")

            pyautogui.press("tab")
            print(f"[{linha.strip()}] Movendo para o campo Linha")

            print(f"[{linha.strip()}] Aguardando verificação de erro...")

            # --------------------------
            # Controle de avanço dentro do bloco (espera por F10 para avançar faixa)
            is_last_in_block = (offset == len(trecho) - 1)

            if not is_last_in_block:
                print(f"[{linha.strip()}] Pressione F10 para a próxima faixa (ou F9 para encerrar)")
                # espera F10 ou F9
                while True:
                    if keyboard.is_pressed("F9"):
                        print("🟥 Execução encerrada pelo usuário.")
                        return
                    if keyboard.is_pressed("F10"):
                        # debouce
                        time.sleep(0.2)
                        break
                    time.sleep(0.05)
            else:
                # último da lista do bloco
                print("")
                print(f"🟩 Fim do bloco {bloco['nome']}")
                print("Pressione F10 para iniciar o próximo bloco (ou F9 para encerrar)")
                while True:
                    if keyboard.is_pressed("F9"):
                        print("🟥 Execução encerrada pelo usuário.")
                        return
                    if keyboard.is_pressed("F10"):
                        time.sleep(0.2)
                        break
                    time.sleep(0.05)

            # pequeno atraso antes de começar próxima faixa (reduz erros de leitura de tecla)
            time.sleep(0.1)

            escutador_tecla()
            if parar_execucao:
                print("🟥 Execução interrompida pelo escutador.")
                return

        # fim do bloco -- ao sair do for interno, o próximo bloco no for externo é processado

    print("")
    print("=" * 50)
    print(" Todos os blocos foram processados (ou execução parou) ")
    print("=" * 50)

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