
#Bibliotecas nativas
import csv
import os
import time
from datetime import datetime

# Bibliotecas
# pip install pyautogui, keyboard, tabulate
import pyautogui
import keyboard
from tabulate import tabulate

# -------------------------
# Configurações / Globals
# -------------------------
parar_execucao = False
LOG_DIR = "AutoSiget3000/LOGs"
os.makedirs(LOG_DIR, exist_ok=True)
nome_arquivo_log = os.path.join(LOG_DIR, f"log_{datetime.now().strftime('%Y%m%d_%H.%M')}.txt")

# -------------------------
# Escutador (marca parar_execucao)
# -------------------------
def interruptorGlobal():
    """Checa F12 (encerra). Deve ser chamado periodicamente dentro dos loops."""
    global parar_execucao
    if keyboard.is_pressed("F12"):
        parar_execucao = True
        log.user("Tecla F12 detectada — encerrando execução.")
        log.user("Retornando ao menu.")
        print("")
        print("")
        print("")
        time.sleep(5)
        parar_execucao = False
        main()

# -------------------------
# Utilitários de Log
# -------------------------
class log:

    def save(mensagem: str):
        """Registra eventos detalhados no arquivo de log (arquivo)."""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        linha = f"[{timestamp}] {mensagem}"
        # grava imediatamente
        with open(nome_arquivo_log, "a", encoding="utf-8") as f:
            f.write(linha + "\n")

    def user(msg: str):
        """Prints sucinto para o usuário (console). Também registra um resumo no log."""
        print(msg)
        # Log mais sucinto: registra que foi mostrado algo ao usuário
        log.save(f"[UI] {msg}")

# -------------------------
# Funções utilitárias de horário
# -------------------------
class formatHora:
    
    def log(h: str):
        """Formato para logs/prints: 'HHMM' -> 'HH:MM' """
        h = (h or "").strip()
        if len(h) == 4 and h.isdigit():
            return f"{h[:2]}:{h[2:]}"
        if len(h) == 3 and h.isdigit():
            return f"0{h[:1]}:{h[1:]}"
        return h

    def dig(h: str):
        """Formato para digitação (mantém '000' -> '000', '044' -> '0044')."""
        h = (h or "").strip()
        if len(h) == 3 and h.isdigit():
            return f"0{h}"
        return h

class csv:
    # -------------------------
    # Leitura do CSV (DictReader) + normalização das chaves
    # -------------------------
    def pedir():
        '''Função auxiliar para solicitar o nome do arquivo .csv ao usuário'''
        
        print("")
        print("Insira o nome do arquivo para iniciar, ENTER para confirmar o nome.")
        
        # Escolha do CSV
        caminho_csv = input("Nome do CSV (sem .csv): ").strip() + ".csv"
        log.user(f"Arquivo selecionado: {caminho_csv}")
        
        linhas, fieldnames = csv.load(caminho_csv)
        return linhas, fieldnames

    def load(caminho_csv: str):
        """
        Lê CSV com header; normaliza nomes das chaves (strip) e valores (strip).
        Retorna (lista_de_dicionarios, lista_de_cabecalhos_normalizados) ou (None, None) se arquivo não encontrado.
        """
        
        if not os.path.exists(caminho_csv):
            log.user(f"Arquivo '{caminho_csv}' não encontrado.")
            main()

        with open(caminho_csv, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f, delimiter=";")
            # Normaliza fieldnames (strip spaces)
            fieldnames = [fn.strip() for fn in (reader.fieldnames or [])]
            linhas = []
            for row in reader:
                # cria nova dict com chaves normalizadas e valores limpos (strip)
                nova = {}
                for k, v in row.items():
                    key = (k or "").strip()
                    # Retira possíveis espaços
                    nova[key] = (v or "").strip()
                linhas.append(nova)

        log.save(f"CSV '{caminho_csv}' carregado. Registros: {len(linhas)}. Campos: {fieldnames}")
        return linhas, fieldnames

class PrgmLinha:
    # -------------------------
    # Geração dos blocos (id_bloco = Linha_Dia_Sentido)
    # -------------------------
    def genBlocos(linhas):
        """
        Recebe lista de dicts (cada linha do CSV com campos normalizados).
        Retorna lista de blocos com estrutura:
        { 'id_bloco': '1001_Sab_0', 'inicio': i0, 'fim': i1, 'count': n }
        Mantém ordem original do CSV.
        """

        blocos = []
        chave_anterior = None
        inicio = 0

        for i, row in enumerate(linhas):
            # Acessa campos com nomes esperados (compatível com seu CSV unificado)
            linha = row.get("Linha", row.get("Linha ", "")).strip()
            dia = row.get("Dia", "").strip()
            sentido = row.get("Sentido", "").strip()

            id_atual = f"{linha}_{dia}_{sentido}"

            if chave_anterior is None:
                chave_anterior = id_atual
                inicio = i
                continue

            if id_atual != chave_anterior:
                blocos.append({
                    "id_bloco": chave_anterior,
                    "inicio": inicio,
                    "fim": i - 1,
                    "count": i - inicio
                })
                chave_anterior = id_atual
                inicio = i

        # último bloco
        if chave_anterior is not None:
            blocos.append({
                "id_bloco": chave_anterior,
                "inicio": inicio,
                "fim": len(linhas) - 1,
                "count": len(linhas) - inicio
            })

        log.save(f"{len(blocos)} blocos gerados.")
        return blocos

    # -------------------------
    # Seleção dos blocos
    # -------------------------
    def selec_Blocos(linhas, blocos):
        '''Escolha do bloco_inicial_index pelo usuário, definido o bloco de partida para execução dos preenchimentos'''
        # exibe blocos (resumido)
        print("")
        print("  Blocos detectados:")
        headers = ["N:", "ID Bloco", "Qtd. Faixas Horárias"]
        tabela = [
            [i + 1, b["id_bloco"], b["count"]]
            for i, b in enumerate(blocos)
        ]
        print("  " + tabulate(tabela, headers=headers, tablefmt="rounded_grid"))
        print("")
        time.sleep(1)

        # escolhe bloco inicial
        bloco_inicial = 0
        deNovo = True
        while deNovo:
            escolha = input(f"Digite o número do bloco para começar (1-{len(blocos)}) | [0] Para voltar ao menu inicial: ").strip()
            time.sleep(0.5)
            if escolha and escolha.isdigit():
                n = int(escolha)
                if 1 <= n <= len(blocos):
                    bloco_inicial = n
                    deNovo = False
                elif n == 0:
                    main()
                else:
                    log.user("Entrada inválida")
                    print("")
                    continue
            else:
                log.user("Entrada inválida")
                print("")
                continue
        
        # Corrige inicio do bloco inicial ignorando o cabeçalho
        bloco_inicio_index=bloco_inicial-1

        # confirma e inicia processamento
        log.user(f"Iniciando a partir do bloco {bloco_inicial}: {blocos[bloco_inicial-1]['id_bloco']}")
        log.save(f"USUARIO iniciou processamento a partir do bloco {bloco_inicial}")

        return (linhas, blocos, bloco_inicio_index)

    # -------------------------
    # Preenchimento de uma faixa
    # -------------------------
    def preenFaixa(row, linha_label, faixaFimAnterior=None):
        """
        Executa os comandos de pyautogui para preencher uma faixa com base em 'row' (dict).
        Recebe faixaFimAnterior (strings (FaixaInicio) para detectar virada de dia.
        Observação: não faz waits por F10 aqui — apenas executa o preenchimento da faixa.
        """

        # preserva strings originais
        faixaIni_raw = row.get("FaixaInicio", row.get("FaixaInicio ", "")).strip()
        faixaFim_raw = row.get("FaixaFinal", row.get("FaixaFinal ", "")).strip()
        interv = row.get("Intervalo", row.get("Interv", "")).strip()
        tPerc = row.get("Percurso", row.get("Perc.", row.get("Perc", ""))).strip()
        tTerm = row.get("TempTerm", row.get("T. Term", row.get("T. Term ", ""))).strip()
        frota = row.get("Frota", "").strip()

        # formata para print/log
        faixaIni_log = formatHora.log(faixaIni_raw)
        faixaFim_log = formatHora.log(faixaFim_raw)

        # prepara valores para digitar (mantendo a lógica existente)
        faixaIni_dig = formatHora.dig(faixaIni_raw)
        faixaFim_dig = formatHora.dig(faixaFim_raw)

        # --------------------------
        # Início do preenchimento (mesma sequência de ações)
        pyautogui.write(faixaIni_dig)
        log.save(f"[{linha_label}] escreveu FaixaInicio '{faixaIni_log}'")
        
        # console
            #log.user(f"{linha_label} • Preenchendo faixa de {faixaIni_log} à {faixaFim_log}")
        print(tabulate([[faixaIni_log, faixaFim_log, interv, tPerc, tTerm, frota]],
                headers=["F. Inicio", "F. Final", "Interv.", "T. Perc.", "T. Term", "Frota"],
                tablefmt="rounded_grid"))

        pyautogui.press("tab")
        log.save(f"[{linha_label}] tab → campo FaixaFinal")

        # variável para evitar dupla verificação de faixa para o dia seguinte
        faixaIniVerif = False

        # Detecta virada de dia: se faixaFimAnterior existe e hora atual < anterior (0000 case)
        if faixaFimAnterior:
            prev_ini = (faixaFimAnterior[0] or "").strip()
            if prev_ini.isdigit() and faixaIni_raw.isdigit():
                if int(faixaIni_raw) < int(prev_ini):
                    log.save(f"[{linha_label}] virada detectada em Faixa Inicio ({faixaIni_raw} < {prev_ini}) — confirmando Enter")
                    time.sleep(0.25)
                    pyautogui.press("tab")
                    pyautogui.press("enter")
                    time.sleep(0.4)
                    faixaIniVerif = True
                    log.save(f"[{linha_label}] faixaIniVerif definida como VERDADEIRA")

        pyautogui.write(faixaFim_dig)
        log.save(f"[{linha_label}] escreveu FaixaFinal '{faixaFim_log}'")

        pyautogui.press("tab")
        log.save(f"[{linha_label}] tab → campo Intervalo")

        # Detecta virada de dia no FaixaFinal em relação ao anterior
        if faixaFimAnterior and not faixaIniVerif:
            prev_fim = faixaIni_dig
            if prev_fim.isdigit() and faixaFim_dig.isdigit():
                if int(faixaFim_dig) < int(prev_fim):
                    log.save(f"[{linha_label}] virada detectada em FaixaFinal ({faixaFim_dig} < {prev_fim}) — confirmando Enter")
                    time.sleep(0.25)
                    pyautogui.press("tab")
                    pyautogui.press("enter")
                    time.sleep(0.4)

        pyautogui.write(interv)
        log.save(f"[{linha_label}] escreveu Intervalo '{interv}'")
        time.sleep(0.1)

        pyautogui.press("tab")
        log.save(f"[{linha_label}] tab → campo Tempo de Percurso")
        time.sleep(0.1)

        pyautogui.write(tPerc)
        log.save(f"[{linha_label}] escreveu Percurso '{tPerc}'")
        time.sleep(0.1)

        pyautogui.press("tab")
        log.save(f"[{linha_label}] tab → campo Terminal")
        time.sleep(0.1)

        pyautogui.write(tTerm)
        log.save(f"[{linha_label}] escreveu TempTerm '{tTerm}'")
        time.sleep(0.1)

        pyautogui.press("tab")
        log.save(f"[{linha_label}] tab → campo Frota")
        time.sleep(0.1)

        pyautogui.write(frota)
        log.save(f"[{linha_label}] escreveu Frota '{frota}'")
        time.sleep(0.1)

        pyautogui.press("tab")
        log.save(f"[{linha_label}] tab → campo Linha (pronto para verificação)")
        time.sleep(0.1)

        # final do preenchimento da faixa
        log.save(f"[{linha_label}] faixa preenchida ({faixaIni_log} → {faixaFim_log})")
        log.user(f"{linha_label} • Faixa preenchida | F10 para continuar | F9 para repetir | F12 para parar ")

    # -------------------------
    # Preenchimento de blocos (controle de interação F10/F12)
    # -------------------------
    def preenBlocos(linhas, blocos, bloco_inicio_index):
        """
        Percorre blocos a partir de bloco_inicio_index (0-based index na lista 'blocos').
        Para cada faixa dentro do bloco chama PrgmLinha_preenFaixa().
        Aguarda F10 entre faixas; F12 encerra tudo.
        """

        global parar_execucao

        total_blocos = len(blocos)
        # itera blocos a partir da escolha do usuário
        for idx_bloco, bloco in enumerate(blocos[bloco_inicio_index:], start=bloco_inicio_index): # Mantive 'start' ajustado
            id_bloco = bloco["id_bloco"]
            inicio = bloco["inicio"]
            fim = bloco["fim"]
            count = bloco["count"]

            log.save(f"INICIO_BLOCO {id_bloco} linhas {inicio+1}-{fim+1}")

            print(tabulate([[id_bloco]],
                        headers=["Bloco"],
                        tablefmt="rounded_grid"))

            # itera dentro do bloco: trecho de linhas
            trecho = linhas[inicio:fim+1]
            faixaFimAnterior = None  # para detectar virada comparando com a faixa anterior dentro do bloco

            offset = 0 # Inicializa o índice relativo dentro do 'trecho'
            while offset < count: # count é o len(trecho)

                row = trecho[offset]
                
                interruptorGlobal()
                if parar_execucao:
                    log.save("Parada solicitada — encerrando preenchimento.")
                    log.user("Execução encerrada pelo usuário.")
                    return

                i_global = inicio + offset
                linha_label = bloco["id_bloco"]  # label amigável para logs e prints

                # timer antes de iniciar na primeira vez
                if offset == 0:
                    log.user("")
                    log.user("  [+] Tempo de 3s para ajustar o cusor no SIGET [+] ")
                    log.user("")
                    time.sleep(3)

                # imprime resumo ao usuário: ação principal
                log.user(f">>> Processando faixa {offset+1}/{count}")

                # chama a rotina que faz os pyautogui.write / press etc.
                PrgmLinha.preenFaixa(row, linha_label, faixaFimAnterior=faixaFimAnterior)

                # atualiza faixaFimAnterior para a próxima iteração (usa raw values)
                faixaFimAnterior = (row.get("FaixaFim", ""))

                # controle de avanço: se não é a última faixa do bloco, espera F10;
                # se for a última, espera F10 para ir ao próximo bloco (ou F12)
                is_last = (offset == count - 1)
                
                log.save(f"AGUARDANDO_F10_F9_F12 faixa {offset+1} do bloco {id_bloco}")
                
                # loop leve aguardando tecla
                while True:
                    interruptorGlobal()
                    if parar_execucao:
                        log.save("Parada solicitada durante espera de F10.")
                        log.user(" >>> Execução encerrada pelo usuário.")
                        return
                    if keyboard.is_pressed("F10"):
                        time.sleep(0.18)
                        log.save(f" >>> F10 pressionado — avançando para faixa {offset+2} do bloco {id_bloco}")
                        offset += 1 
                        break
                    if keyboard.is_pressed("F9"):
                        time.sleep(0.18)
                        log.save(f" >>> F9 pressionado — repetindo faixa {offset+1} do bloco {id_bloco}")
                        is_last = False
                        break
                    time.sleep(0.06)

                if is_last:
                    # último da lista do bloco
                    offset += 1

                    print("")
                    log.user(f" [+] Bloco {id_bloco} finalizado ({count} faixas horárias).  [+] ")
                    log.save(f"FIM_BLOCO {id_bloco}")
                    print("")
                    
                    # Verifica se há um próximo bloco
                    prox_idx = idx_bloco - bloco_inicio_index  # posição relativa no enumerate
                    if prox_idx != 0:
                        log.save(f">>> Ultimo bloco")
                    elif prox_idx < total_blocos - 1:
                        prox_id_bloco = blocos[bloco_inicio_index + prox_idx + 1]["id_bloco"]
                        log.save(f">>> Próximo bloco: {prox_id_bloco}")
                        print(tabulate([[prox_id_bloco]],
                            headers=["Prox. Bloco"],
                            tablefmt="rounded_grid"))
                    else:
                        prox_id_bloco = None
                        log.user(">>> Este foi o último bloco.")
                    log.user(">>> Pressione F10 para iniciar o próximo bloco | F9 para selecionar bloco | F12 para encerrar")
                    # Espera F10, F9 ou F12
                    while True:
                        interruptorGlobal()
                        if parar_execucao:
                            log.save("Parada solicitada na finalização de bloco.")
                            log.user("Execução encerrada pelo usuário.")
                            return
                        if keyboard.is_pressed("F10"):
                            time.sleep(0.18)
                            if prox_id_bloco:
                                log.save(f" >>> F10 pressionado — iniciando próximo bloco ({prox_id_bloco})")
                            else:
                                log.save(" >>> F10 pressionado — não há próximo bloco (fim).")
                            break
                        if keyboard.is_pressed("F9"):
                            time.sleep(0.18)
                            log.save(f" >>> F9 pressionado iniciando seletor de blocos")
                            PrgmLinha.selec_Blocos(linhas, blocos)
                            break
                        time.sleep(0.06)
            
            # fim do bloco — segue para próximo bloco no for
            # pequeno delay entre blocos
            time.sleep(0.15)
            
        # todos blocos processados
        print("")
        log.user("Todos os blocos processados.")
        print("")
        print("")
        print("Retornando ao MENU...")
        log.save("PROCESSAMENTO_COMPLETO")
        main()

# -------------------------
# Normalização de OSOs
# -------------------------
class printPDF:
    
    def tratarOSOs(osoBruta):
        """
        Verifica a formatação das OSOs [123456], 
        identifica se é Base ou Derivada e 
        relaciona com a respectiva Linha 
        """
        
        QuantidadeDeOsos = 0
        for i , row in enumerate(osoBruta):
            oso = row.get("Oso", row.get("Oso ", "")).strip()
            if oso != "":
                QuantidadeDeOsos = QuantidadeDeOsos + 1
        
        osos = []

        for i, row in enumerate(osoBruta[:QuantidadeDeOsos]):

            linha = row.get("LinhaOso", row.get("LinhaOso ", "")).strip()
            oso = row.get("Oso", row.get("Oso ", "")).strip()
            
            linha_dig = row.get("LinhaOso", row.get("LinhaOso ", "")).strip()
            oso_dig = row.get("Oso", row.get("Oso ", "")).strip()

            if len(oso) == 6 and oso.isdigit():
                fimOso = int(oso[4:])

                if fimOso == 00:
                    oso_dig = oso[:4]
                    linha_dig = linha[:4]
                    osos.append({
                        "n_oso": i + 1,
                        "linha": linha_dig,
                        "oso_dig": oso_dig,
                        "oso": oso,
                        "tipo": "BASE"
                    })
                else:
                    oso_dig = oso[:4] + "-" + oso[4:]
                    linha_dig = linha[:4] + "-" + linha[4:]
                    osos.append({
                        "n_oso": i + 1,
                        "linha": linha_dig,
                        "oso_dig": oso_dig,
                        "oso": oso,
                        "tipo": "DERIVADA"
                    })
            else:
                log.user(f"OSO [{oso}] | formato inválido seguir o formato: [123456]")
                return
                
        filtoB = lambda x: x["tipo"] == "BASE"
        filtoD = lambda x: x["tipo"] == "DERIVADA"
        ososBase = list(filter(filtoB, osos))
        ososDerivada = list(filter(filtoD, osos))

        print("")
        log.user(f"{len(osos)} Osos idenficadas | B:{len(ososBase)} D:{len(ososDerivada)} ")
        print("")
        time.sleep(1)  # import time
        return osos

    # -------------------------
    # Impressão de OSOs (controle de interação F10/F12)
    # -------------------------
    def preencher_PDFs(osos, oso_inicial_index):
        """
        Percorre osos a partir de oso_inicial_index (0-based index na lista 'osos').
        Para cada faixa dentro do OSO chama PrgmLinha_preenFaixa().
        Aguarda F10 entre faixas; F12 encerra tudo.
        """

        global parar_execucao
        fQH = ""

        print('''
            Selecione o tipo de impressão:
                [1] - FH - Faixa Horária     [OSOS ATIVAS]
                [2] - QH - Quadro Horário    [OSOS ATIVAS]
                ===========================================
                [3] - FH - Faixa Horária  [OSOS DESATIVAS]
                [4] - QH - Quadro Horário [OSOS DESATIVAS]
            ''')

        # escolhe FH ou QH
        deNovo = True
        while deNovo:
            escolha = input("Digite o número do tipo de impressão: ").strip()
            time.sleep(0.5)
            if escolha and escolha.isdigit():
                n = int(escolha)
                if n == 1:
                    fQH = "FH"
                    OsoAtiva = True
                    break
                elif n == 2:
                    fQH = "QH"
                    OsoAtiva = True
                    break
                elif n == 3:
                    fQH = "FH"
                    OsoAtiva = False
                    break
                elif n == 4:
                    fQH = "QH"
                    OsoAtiva = False
                    break
                else:
                    log.user("Entrada inválida")
                    print("")
                    continue
            else:
                log.user("Entrada inválida")
                print("")
                continue

        # itera osos a partir da escolha do usuário
        for idx_oso, OSO in enumerate(osos[oso_inicial_index:], start=oso_inicial_index):

            oso = OSO["oso"]
            oso_label = OSO["oso_dig"]  # label amigável para logs e prints
            linha = OSO["linha"]

            log.save(f"INICIO OSO: {oso} | Linha: {linha}")
            
            tabela = [[oso_label, linha]]
            print(tabulate(tabela, headers=["OSO", "LINHA"],tablefmt="rounded_grid"))
            print("")

            RepetirOso = True
            while RepetirOso:
                
                interruptorGlobal()
                if parar_execucao:
                    log.save("Parada solicitada — encerrando preenchimento.")
                    log.user("Execução encerrada pelo usuário.")
                    return

                # imprime resumo ao usuário: ação principal
                log.user(f">>> Processando OSO {oso_label}")

                row = osos[idx_oso]

                # Verifica se tem OSO devivada após a o OSO base atual 
                OsoDerivada = False
                prox_idx = idx_oso + 1
                if prox_idx < len(osos):
                    prox_oso = osos[prox_idx]
                    if prox_oso["tipo"] == "DERIVADA":
                        OsoDerivada = True

                # chama a rotina que faz os pyautogui.write / press etc.
                printPDF.imprimirPDF(row, osos, OsoAtiva, OsoDerivada, fQH)
                
                log.save(f"AGUARDANDO_F10_F9_F12 OSO {oso_label}")
                
                # loop leve aguardando tecla
                while True:
                    interruptorGlobal()
                    if parar_execucao:
                        log.save("Parada solicitada durante espera de F10.")
                        log.user(" >>> Execução encerrada pelo usuário.")
                        RepetirOso = False
                        return
                    if keyboard.is_pressed("F10"):
                        time.sleep(0.18)
                        log.save(f" >>> F10 pressionado — preenchendo oso do OSO {oso}")
                        RepetirOso = False
                        break
                    if keyboard.is_pressed("F9"):
                        time.sleep(0.18)
                        log.save(f" >>> F9 pressionado — repetindo faixa do OSO {oso}")
                        RepetirOso = True
                        break
                    time.sleep(0.06)
            # fim da OSO — segue para próxima OSO no for
            # pequeno delay entre osos
            time.sleep(0.15)
            
        # todos osos processados
        print("")
        log.user("Todos os osos processados.")
        print("")
        print("")
        print("Retornando ao MENU...")
        time.sleep(1)
        log.save("PROCESSAMENTO_COMPLETO")

    # -------------------------
    # Selecionar OSO a ser preenchida
    # -------------------------
    def seletorOsos(OSOs):
        
        headers = ["Nº", "Linha", "OSO", "Tipo"]
        tabela = [
            [b["n_oso"],b["linha"], b["oso_dig"], b["tipo"]]
            for i, b in enumerate(OSOs)
        ]
        print("")
        print(tabulate(tabela, headers=headers, tablefmt="rounded_grid"))
        print("")
        time.sleep(1)
        
        # escolhe OSO inicial
        oso_inicial = 1
        deNovo = True
        while deNovo:
            escolha = input(f"Digite o número da OSO para começar (1-{len(OSOs)}) | [0] Para voltar ao menu inicial: ").strip()
            time.sleep(0.5)
            if escolha and escolha.isdigit():
                n = int(escolha)
                if 1 <= n <= len(OSOs):
                    oso_inicial = n
                    deNovo = False
                elif n == 0:
                    main()
                else:
                    log.user("Entrada inválida")
                    print("")
                    continue
            else:
                log.user("Entrada inválida")
                print("")
                continue
        print("")
        printPDF.preencher_PDFs(OSOs, oso_inicial_index=oso_inicial-1)

    # -------------------------
    # Impressão de um PDF
    # -------------------------
    def imprimirPDF(row, oso, OsoAtiva, OsoDerivada, fQH):
        '''
        Impressão de um PDF.
        '''

        # preserva strings originais
        linha = row.get("linha", row.get("linha ", "")).strip()
        oso = row.get("oso", row.get("oso ", "")).strip()
        oso_dig = row.get("oso_dig", row.get("oso_dig ", "")).strip()
        NomePDF = f"{fQH} {linha} {oso_dig}"


        # --------------------------
        # Início do preenchimento
        # --------------------------

        log.user("  [+] Tempo de 3s para ajustar o cusor no SIGET [+] ")
        time.sleep(3)

        # Preenche o campo Insira o Nº da OSO
        pyautogui.write(oso)
        log.save(f"[{oso_dig}] escreveu OSO '{oso}'")
        log.user(f"Preenchendo OSO:     {oso_dig}")
        time.sleep(0.3)

        # Navega de N° da OSO → Imprimir em
        pyautogui.press("tab")
        log.save(f"[{oso_dig}] tab → Imprimir em")

        # Confirma se é OSO derivada (filha)
        if OsoAtiva and OsoDerivada:
            aviso = f"Aviso: OSO filha detectada ({oso_dig})"
            log.user(aviso)
            log.save(f"[{oso_dig}] {aviso}")

            # Navega até o campo “SIM / NÃO” e escolhe “NÃO”
            pyautogui.press("right")  # muda SIM → NÃO
            log.save(f"[{oso_dig}] alterou opção para NÃO")
            
            pyautogui.press("enter")
            log.save(f"[{oso_dig}] confirmou NÃO")
            time.sleep(0.4)

        # Navega de Imprimir em → Configurar → Imprimir
        pyautogui.press("tab", presses=2, interval=0.2)
        log.save(f"[{oso_dig}] tab → botão Imprimir")
        
        pyautogui.press("enter")
        log.save(f"[{oso_dig}] confirmou impressão")
        time.sleep(2)

        # Digita o nome do PDF e confirma salvar
        pyautogui.write(NomePDF)
        log.save(f"[{oso_dig}] escreveu nome do PDF '{NomePDF}'")
        log.user(f"Preenchendo NomePDF: {NomePDF}")
        time.sleep(0.3)
        
        pyautogui.press("enter")
        log.save(f"[{oso_dig}] confirmou salvar PDF")
        time.sleep(3.5)

        # Verifica se o PDF foi salvo
        caminho_pdf = os.path.join("AutoSiget3000", "FQHs", f"{NomePDF}.pdf")

        '''
        # Cria um PDF fake (arquivo vazio ou com texto de teste)
        with open(caminho_pdf, "w") as f:
            f.write(f"Arquivo fake para debug: {NomePDF}.pdf\n")
        '''

        if os.path.exists(caminho_pdf):
            log.save(f"[{oso_dig}] PDF salvo com sucesso: {NomePDF}.pdf")
            log.user(f"   {NomePDF}.pdf | Salvo com sucesso")
            log.user(f"[{oso_dig}] • F10 para continuar | F9 para repetir | F12 para parar ")
        else:
            log.save(f"[{oso_dig}] ERRO: PDF não encontrado ({NomePDF}.pdf)")
            log.user("===============================================")
            log.user(f"   ERRO: {NomePDF}.pdf | PDF não encontrado   ")
            log.user("===============================================")
            log.user(f"[{oso_dig}] • F10 para continuar | F9 para repetir | F12 para parar ")

        # final do preenchimento da faixa

def IniciarModulo_printPDF(dados):
    
    LOG_DIR = "AutoSiget3000/FQHs"
    os.makedirs(LOG_DIR, exist_ok=True)

    # Tratar osos e adicionar dados calculados
    OSOs = printPDF.tratarOSOs(dados)

    # Instruções antes de iniciar o programa
    print("REQUISITOS E INSTRUÇÕES:")
    print('''
    Antes de iniciar preencha os seguintes requitos:
        [$] Prepare o CSV conforme instruções na documentação
            e deixe na mesma pasta deste programa.
        [$] Defina a impressora padrão como Microsoft Print to PDF.
            *Nescessário cada vez que abir o SIGET.
        [$] NÃO exclua, mova ou renomei a pasta FQHs criada na
            mesma pasta desse programa durante a execução.
            Após encerrado a pasta está livre para
            modificações ou exclusão.
        [$] Abra o SIGET.
        [$] Abra a aba "Documentos OSO".
        [$] Imprima uma faixa ou quadro para configurar o SIGET.
        [$] Salve na pasta "FQHs" criada na mesma pasta desde programa.
        [$] Posicione o cursor no inicio de "Informe o N° da OSO".

    ''')

    print("Pressione F10 para continuar...")
    keyboard.wait("F10")

    print('''
    INSTRUÇÕES E AVISOS:
        [#] Este prgrama apenas automatiza a impressão das
        faixas e quadros horários (programação da faixa)
        [#] Quaisquer erros e avisos que apareça continua
            as soluções que já são utiluzadas normalmente
        [#] SEMPRE deve se colocar o cursor no inico da
            "Informe o Nº da OSO" de cada OSO que será
            impressa
    ''')

    print("Pressione F10 para continuar...")
    keyboard.wait("F10")
    print("")

    # Iniciar o preenchimento de PDFs
    printPDF.printPDF.seletorOsos(OSOs) 

def IniciarModulo_PgrmLinha(dados):
    '''Inicia o processamento do Módulo de programação de linha'''
    # gera Blocos
    blocos = PrgmLinha.genBlocos(dados)

    if not blocos:
        log.user("Nenhum OSO identificado no CSV.")
        return
    
    # Instruções antes de iniciar o programa
    print("REQUISITOS E INSTRUÇÕES:")
    print('''
    Antes de iniciar preencha os seguintes requitos:
        [$] Prepare o CSV conforme instruções na documentação
        [$] Abra o SIGET
        [$] Abra a aba "Parâmetros de linha"
        [$] Insira a OSO da linha a ser preenchida
            se atentando ao SENTIDO
        [$] Limpe TODAS as faixas horárias do DIA e SENTIDO
            que será preenchio
        [$] Posicione o cursor na primeira CÉLULA da 1ª linha de preenchimento

    ''')

    print("Pressione F10 para continuar...")
    keyboard.wait("F10")

    print('''
    INSTRUÇÕES E AVISOS:
        [#] Este prgrama apenas automatiza a digitação das
        faixas horárias (programação da faixa)
        [#] Quaisquer erros e avisos que apareça continua
            as soluções que já são utiluzadas normalmente
        [#] SEMPRE deve se colocar o cursor no inico da
            FAIXA INICIAL de cada faixa horária que será
            preenchida
    ''')

    print("Pressione F10 para continuar...")
    keyboard.wait("F10")

    #Seleciona os Blocos
    PrgmLinha.selec_Blocos(dados, blocos)

# -------------------------
# Interface principal
# -------------------------
def main():
    '''Módulo central do programa, menu inicial'''
    dados, cabecalho = csv.pedir()
    if dados is None:
        return

    # valida campos mínimos (esperados)
    chavesObrigatorias = {"FaixaInicio", "FaixaFinal", "Intervalo", "Percurso", "TempTerm", "Frota", "Linha", "Dia", "Sentido", "Oso", "LinhaOso"}
    chavesCabecalho = set(k for k in (cabecalho or []))
    
    # Normaliza chavesCabecalho
    chavesCabecalho = set(h.strip() for h in (cabecalho or []))
    chaveFaltando = chavesObrigatorias - chavesCabecalho
    if chaveFaltando:
        log.user(f"Erro: o CSV está faltando colunas obrigatórias: {', '.join(chaveFaltando)}")
        main()
    
    while True:
        print('''
        Selecione o módulo:
            [1] - Digitar programação de linha
            [2] - Imprimir PDF

            [0] - Voltar ao MENU
        ''')

        # escolhe módulo
        deNovo = True
        while deNovo:
            escolha = input("Digite o número do módulo: ").strip()
            time.sleep(0.5)
            if escolha and escolha.isdigit():
                n = int(escolha)
                if n == 1:
                    IniciarModulo_PgrmLinha(dados)
                elif n == 2:
                    IniciarModulo_printPDF.ImprimirPDFs(dados)
                elif n == 0:
                    main()
                else:
                    log.user("Entrada inválida")
                    print("")
                    continue
            else:
                log.user("Entrada inválida")
                print("")
                continue

def intro():
    '''Mensagem de introdução ao sistema'''
    global parar_execucao

    print("=" * 60)
    print("                    AUTOMA SIGET3000")
    print("=" * 60)
    print("Bem Vindo(a)")
    
    main()

# -------------------------
# Roda a introdução
# -------------------------
if __name__ == "__main__":
    intro()




# ===================
#  Autor: Eude Leal
#  Github: eudeleal
# ===================

# Brasil
# Salvador, BA - 20255