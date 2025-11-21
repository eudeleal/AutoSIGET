
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
        Log.user("Tecla F12 detectada — encerrando execução.")
        Log.user("Retornando ao menu.")
        print("")
        print("")
        print("")
        time.sleep(5)
        parar_execucao = False
        main()

# -------------------------
# Utilitários de Log
# -------------------------
class Log:

    @staticmethod
    def save(mensagem: str):
        """Registra eventos detalhados no arquivo de log (arquivo)."""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        linha = f"[{timestamp}] {mensagem}"
        # grava imediatamente
        with open(nome_arquivo_log, "a", encoding="utf-8") as f:
            f.write(linha + "\n")

    @staticmethod
    def user(msg: str):
        """Reporte sucinto para o usuário (console). Também registra um resumo no log."""
        print(msg)
        # Log mais sucinto: registra que foi mostrado algo ao usuário
        Log.save(f"[UI] {msg}")

# -------------------------
# Funções utilitárias de horário
# -------------------------
class Util:

    @staticmethod
    def formatHora_log(hora: str):
        # Formato para logs/prints: 'HHMM' ≥ 'HH:MM'
        h = (hora or "").strip()
        if len(h) == 4 and h.isdigit():
            return f"{h[:2]}:{h[2:]}"
        if len(h) == 3 and h.isdigit():
            return f"0{h[:1]}:{h[1:]}"
        return h

    @staticmethod
    def formatHora_dig(hora: str):
        # Formato para digitação (mantém '000' ≥ '000', '044' ≥ '0044').
        h = (hora or "").strip()
        if len(h) == 3 and h.isdigit():
            return f"0{h}"
        return h

# -------------------------
# Funções utilitárias gerais
# -------------------------
class Menu:
    @staticmethod
    def escolha(texto1, texto2):
        resposta = str
        DeNovo = bool(True)

        while DeNovo:
            print("")
            print(f"{texto1} [1] | {texto2} [2]")
            resposta = input("(Enter para confirmar) | Resposta: ").strip()

            if resposta.isdigit():
                if resposta == '1':
                    resposta = int(resposta)
                    break
                elif resposta == '2':
                    resposta = int(resposta)
                    break
                else:
                    print("Resposta inválida")
            else:
                print("Resposta inválida")

        return resposta

    @staticmethod
    def blocoInicial(blocos):
        bloco_inicial = int
        deNovo = (bool(True))

        while deNovo:
            escolha = input(
                f"Digite o número para começar (1-{len(blocos)}) | [0] Para voltar ao menu inicial: ").strip()
            time.sleep(0.5)
            if escolha and escolha.isdigit():
                n = int(escolha)
                if 1 <= n <= len(blocos):
                    bloco_inicial = n
                    deNovo = False
                elif n == 0:
                    bloco_inicial = -16
                else:
                    Log.user("Entrada inválida")
                    print("")
                    continue
            else:
                Log.user("Entrada inválida")
                print("")
                continue

        return bloco_inicial

class LCsv:
    # -------------------------
    # Leitura do CSV (DictReader) + normalização das chaves
    # -------------------------

    @staticmethod
    def pedir():
        # Função auxiliar para solicitar o nome do arquivo .csv ao usuário

        print(" |> Selecione o arquivo csv |")
        print(" |>    F10 para continuar   | ")
        print("")

        keyboard.wait('F10')

        while True:
            caminho_csv = input("Digite o nome do arquivo CSV (sem '.csv'): ").strip() + ".csv"
    
            Log.user(f" |> Arquivo selecionado: {caminho_csv}")
    
            time.sleep(0.5)
    
            if not os.path.exists(caminho_csv):
                Log.user(f" |> Arquivo '{caminho_csv}' não encontrado.")
                print("")
                continue
            break

        linhas, fieldnames = LCsv.load(caminho_csv)
        return linhas, fieldnames

    @staticmethod
    def load(caminho_csv: str):
        """
        Lê CSV com header; normaliza nomes das chaves (strip) e valores (strip).
        Retorna (lista de dicionários, lista de cabeçalhos normalizados) ou (None, None) se arquivo não encontrado.
        """

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

        Log.save(f"CSV '{caminho_csv}' carregado. Registros: {len(linhas)}. Campos: {fieldnames}")
        return linhas, fieldnames

class ProgramaLinha:

    @staticmethod
    def instrucao():

        # Limpa o console
        os.system('cls' if os.name == 'nt' else 'clear')

        # Instruções antes de iniciar o programa
        print("REQUISITOS E INSTRUÇÕES:")
        print('''
        Antes de iniciar preencha os seguintes requisitos:
            [$] Prepare o CSV conforme instruções na documentação
            [$] Abra o SIGET
            [$] Abra a aba "Parâmetros de linha"
            [$] Insira a OSO da linha a ser preenchida
                se atentando ao SENTIDO
            [$] Limpe TODAS as faixas horárias do DIA e SENTIDO
                que será preenchido
            [$] Posicione o cursor na primeira CÉLULA da 1ª linha de preenchimento

        ''')

        print("Pressione F10 para continuar...")
        keyboard.wait("F10")

        print('''
        INSTRUÇÕES E AVISOS:
            [#] Este programa apenas automatiza a digitação das
            faixas horárias (programação da faixa)
            [#] Quaisquer erros e avisos que apareça continua
                as soluções que já são utilizadas normalmente
            [#] SEMPRE deve se colocar o cursor no inicio da
                FAIXA INICIAL de cada faixa horária que será
                preenchida
        ''')

        print("Pressione F10 para continuar...")
        keyboard.wait("F10")

        return None
    
    # -------------------------
    # Geração dos blocos (id_bloco = Linha_Dia_Sentido)
    # -------------------------
    @staticmethod
    def genBlocos(linhas):
        """
        Recebe lista de dicts (cada linha do CSV com campos normalizados).
        Retorna lista de blocos com estrutura:
        {'id_bloco': '1001_Sab_0', 'inicio': i0, 'fim': i1, 'count': n}
        Mantém ordem original do csv.
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

        Log.save(f"{len(blocos)} blocos gerados.")
        return blocos

    # -------------------------
    # Seleção dos blocos
    # -------------------------
    @staticmethod
    def select_Blocos(linhas, blocos):
        
        # Limpa o console
        os.system('cls' if os.name == 'nt' else 'clear')
        
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
        bloco_inicial = Menu.blocoInicial(blocos)

        if bloco_inicial == -16:
            return linhas, blocos, bloco_inicial

        # Corrige inicio do bloco inicial ignorando o cabeçalho
        bloco_inicio_index=bloco_inicial-1

        # confirma e inicia processamento
        Log.user(f"Iniciando a partir do bloco {bloco_inicial}: {blocos[bloco_inicial - 1]['id_bloco']}")
        Log.save(f"USUARIO iniciou processamento a partir do bloco {bloco_inicial}")

        return linhas, blocos, bloco_inicio_index

    # -------------------------
    # Preenchimento de uma faixa
    # -------------------------
    @staticmethod
    def preenFaixa(row, linha_label, faixaFimAnterior=None):
        """
        Executa os comandos de pyautogui para preencher uma faixa com base em 'row' (dict).
        Recebe faixaFimAnterior strings (FaixaInicio) para detectar virada de dia.
        Observação: não faz waits por F10 aqui — apenas executa o preenchimento da faixa.
        """

        # preserva strings originais
        faixaIni_raw = row.get("FaixaInicio", row.get("FaixaInicio ", "")).strip()
        faixaFim_raw = row.get("FaixaFinal", row.get("FaixaFinal ", "")).strip()
        Interv = row.get("Interv", row.get("Interv", "")).strip()
        tPerc = row.get("Percurso", row.get("Perc.", row.get("Perc", ""))).strip()
        tTerm = row.get("TempTerm", row.get("T. Term", row.get("T. Term ", ""))).strip()
        frota = row.get("Frota", "").strip()

        # formata para print/log
        faixaIni_log = Util.formatHora_log(faixaIni_raw)
        faixaFim_log = Util.formatHora_log(faixaFim_raw)

        # prepara valores para digitar (mantendo a lógica existente)
        faixaIni_dig = Util.formatHora_dig(faixaIni_raw)
        faixaFim_dig = Util.formatHora_dig(faixaFim_raw)

        # --------------------------
        # Início do preenchimento (mesma sequência de ações)
        pyautogui.write(faixaIni_dig)
        Log.save(f"[{linha_label}] escreveu FaixaInicio '{faixaIni_log}'")
        
        # console
            #log.user(f"{linha_label} • Preenchendo faixa de {faixaIni_log} à {faixaFim_log}")
        print(tabulate([[faixaIni_log, faixaFim_log, Interv, tPerc, tTerm, frota]],
                       headers=["F. Inicio", "F. Final", "Interv.", "T. Perc.", "T. Term", "Frota"],
                       tablefmt="rounded_grid"))

        pyautogui.press("tab")
        Log.save(f"[{linha_label}] tab → campo FaixaFinal")

        # variável para evitar dupla verificação de faixa para o dia seguinte
        faixaIniVerif = False

        # Detecta virada de dia: se faixaFimAnterior existe e hora atual < anterior (0000 case)
        if faixaFimAnterior:
            prev_ini = (faixaFimAnterior[0] or "").strip()
            if prev_ini.isdigit() and faixaIni_raw.isdigit():
                if int(faixaIni_raw) < int(prev_ini):
                    Log.save(f"[{linha_label}] virada detectada em Faixa Inicio ({faixaIni_raw} < {prev_ini}) — confirmando Enter")
                    time.sleep(0.25)
                    pyautogui.press("tab")
                    pyautogui.press("enter")
                    time.sleep(0.4)
                    faixaIniVerif = True
                    Log.save(f"[{linha_label}] faixaIniVerif definida como VERDADEIRA")

        pyautogui.write(faixaFim_dig)
        Log.save(f"[{linha_label}] escreveu FaixaFinal '{faixaFim_log}'")

        pyautogui.press("tab")
        Log.save(f"[{linha_label}] tab → campo Intervalo")

        # Detecta virada de dia no FaixaFinal em relação ao anterior
        if faixaFimAnterior and not faixaIniVerif:
            prev_fim = faixaIni_dig
            if prev_fim.isdigit() and faixaFim_dig.isdigit():
                if int(faixaFim_dig) < int(prev_fim):
                    Log.save(f"[{linha_label}] virada detectada em FaixaFinal ({faixaFim_dig} < {prev_fim}) — confirmando Enter")
                    time.sleep(0.25)
                    pyautogui.press("tab")
                    pyautogui.press("enter")
                    time.sleep(0.4)

        pyautogui.write(Interv)
        Log.save(f"[{linha_label}] escreveu Intervalo '{Interv}'")
        time.sleep(0.1)

        pyautogui.press("tab")
        Log.save(f"[{linha_label}] tab → campo Tempo de Percurso")
        time.sleep(0.1)

        pyautogui.write(tPerc)
        Log.save(f"[{linha_label}] escreveu Percurso '{tPerc}'")
        time.sleep(0.1)

        pyautogui.press("tab")
        Log.save(f"[{linha_label}] tab → campo Terminal")
        time.sleep(0.1)

        pyautogui.write(tTerm)
        Log.save(f"[{linha_label}] escreveu TempTerm '{tTerm}'")
        time.sleep(0.1)

        pyautogui.press("tab")
        Log.save(f"[{linha_label}] tab → campo Frota")
        time.sleep(0.1)

        pyautogui.write(frota)
        Log.save(f"[{linha_label}] escreveu Frota '{frota}'")
        time.sleep(0.1)

        pyautogui.press("tab")
        Log.save(f"[{linha_label}] tab → campo Linha (pronto para verificação)")
        time.sleep(0.1)

        # final do preenchimento da faixa
        Log.save(f"[{linha_label}] faixa preenchida ({faixaIni_log} → {faixaFim_log})")
        Log.user(f"{linha_label} • Faixa preenchida | F10 para continuar | F9 para repetir | F12 para parar ")

    # -------------------------
    # Preenchimento de blocos (controle de interação F10/F12)
    # -------------------------
    @staticmethod
    def preenBlocos(linhas, blocos, bloco_inicio_index):
        """
        Percorre blocos a partir de bloco_inicio_index (0-based index na lista 'blocos').
        Para cada faixa dentro do bloco chama PragmaLinha_preenFaixa().
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

            Log.save(f"INICIO_BLOCO {id_bloco} linhas {inicio + 1}-{fim + 1}")

            # Limpa o console
            os.system('cls' if os.name == 'nt' else 'clear')

            print(tabulate([[id_bloco]],
                        headers=["Bloco"],
                        tablefmt="rounded_grid"))

            # itera dentro do bloco: trecho de linhas
            trecho = linhas[inicio:fim+1]
            faixaFimAnterior = None  # para detectar virada comparando com a faixa anterior dentro do bloco
            prox_id_bloco = None # Zera o próximo bloco[ID]

            offset = 0 # Inicializa o índice relativo dentro do 'trecho'
            while offset < count: # count é o len(trecho)

                row = trecho[offset]
                
                interruptorGlobal()
                if parar_execucao:
                    Log.save("Parada solicitada — encerrando preenchimento.")
                    Log.user("Execução encerrada pelo usuário.")
                    return None

                linha_label = bloco["id_bloco"]  # label amigável para logs e prints

                # timer antes de iniciar na primeira vez
                if offset == 0:
                    Log.user("")
                    Log.user("  [+] Tempo de 3s para ajustar o cursor no SIGET [+] ")
                    Log.user("")
                    time.sleep(3)

                # imprime resumo ao usuário: ação principal
                Log.user(f">>> Processando faixa {offset + 1}/{count}")

                # chama a rotina que faz os pyautogui.write / press, etc.
                ProgramaLinha.preenFaixa(row, linha_label, faixaFimAnterior=faixaFimAnterior)

                # atualiza faixaFimAnterior para a próxima iteração (usa raw values)
                faixaFimAnterior = (row.get("FaixaFim", ""))

                # controle de avanço: se não é a última faixa do bloco, espera F10;
                # se for a última, espera F10 para ir ao próximo bloco (ou F12)
                is_last = (offset == count - 1)
                
                Log.save(f"AGUARDANDO_F10_F9_F12 faixa {offset + 1} do bloco {id_bloco}")
                
                # loop leve aguardando tecla
                while True:
                    if keyboard.is_pressed("F12"):
                        Log.save("Parada solicitada durante espera de F10.")
                        Log.user(" >>> Execução encerrada pelo usuário.")
                        return None
                    if keyboard.is_pressed("F10"):
                        time.sleep(0.18)
                        Log.save(f" >>> F10 pressionado — avançando para faixa {offset + 2} do bloco {id_bloco}")
                        offset += 1 
                        break
                    if keyboard.is_pressed("F9"):
                        time.sleep(0.18)
                        Log.save(f" >>> F9 pressionado — repetindo faixa {offset + 1} do bloco {id_bloco}")
                        is_last = False
                        break
                    time.sleep(0.06)

                if is_last:
                    # último da lista do bloco
                    offset += 1

                    print("")
                    Log.user(f" [+] Bloco {id_bloco} finalizado ({count} faixas horárias).  [+] ")
                    Log.save(f"FIM_BLOCO {id_bloco}")
                    print("")
                    
                    # Verifica se há um próximo bloco
                    prox_idx = idx_bloco - bloco_inicio_index  # posição relativa no enumerate
                    if prox_idx != 0:
                        Log.save(f">>> Ultimo bloco")
                    elif prox_idx < total_blocos - 1:
                        prox_id_bloco = blocos[bloco_inicio_index + prox_idx + 1]["id_bloco"]
                        Log.save(f">>> Próximo bloco: {prox_id_bloco}")
                        print(tabulate([[prox_id_bloco]],
                            headers=["Prox. Bloco"],
                            tablefmt="rounded_grid"))
                    else:
                        prox_id_bloco = None
                        Log.user(">>> Este foi o último bloco.")
                    Log.user(">>> Pressione F10 para iniciar o próximo bloco | F9 para selecionar bloco | F12 para encerrar")
                    # Espera F10, F9 ou F12
                    while True:
                        if keyboard.is_pressed("F12"):
                            Log.save("Parada solicitada na finalização de bloco.")
                            Log.user("Execução encerrada pelo usuário.")
                            return None
                        if keyboard.is_pressed("F10"):
                            time.sleep(0.18)
                            if prox_id_bloco:
                                Log.save(f" >>> F10 pressionado — iniciando próximo bloco ({prox_id_bloco})")
                            else:
                                Log.save(" >>> F10 pressionado — não há próximo bloco (fim).")
                            break
                        if keyboard.is_pressed("F9"):
                            time.sleep(0.18)
                            Log.save(f" >>> F9 pressionado iniciando seletor de blocos")
                            gerarBlocos = True
                            return gerarBlocos
                        time.sleep(0.06)
            
            # fim do bloco — segue para próximo bloco no for
            # pequeno delay entre blocos
            time.sleep(0.15)
            
        # todos blocos processados
        print("")
        Log.user("Todos os blocos processados.")
        print("")

        gerarBlocos = False
        return gerarBlocos

class PrintPDF:

    @staticmethod
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
                Log.user(f"OSO [{oso}] | formato inválido seguir o formato: [123456]")
                return None
                
        filtroB = lambda x: x["tipo"] == "BASE"
        filtroD = lambda x: x["tipo"] == "DERIVADA"
        ososBase = list(filter(filtroB, osos))
        ososDerivada = list(filter(filtroD, osos))

        print("")
        Log.user(f"{len(osos)} Osos identificadas | B:{len(ososBase)} D:{len(ososDerivada)} ")
        print("")
        time.sleep(1)  # import time
        return osos

    @staticmethod
    def SelecionarFQH():
        fQH = ""
        OsoAtiva = False

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
                    Log.user("Entrada inválida")
                    print("")
                    continue
            else:
                Log.user("Entrada inválida")
                print("")
                continue

        return fQH, OsoAtiva

    # -------------------------
    # Impressão de OSOs (controle de interação F10/F12)
    # -------------------------
    @staticmethod
    def preencher_PDFs(osos, oso_inicial_index, fQH, OsoAtiva):
        """
        Percorre osos a partir de oso_inicial_index (0-based index na lista 'osos').
        Para cada faixa dentro do OSO chama PragmaLinha_preenFaixa().
        Aguarda F10 entre faixas; F12 encerra tudo.
        """

        global parar_execucao
        
        # itera osos a partir da escolha do usuário
        for idx_oso, OSO in enumerate(osos[oso_inicial_index:], start=oso_inicial_index):
            
            n_oso = OSO["n_oso"]
            oso = OSO["oso"]
            oso_label = OSO["oso_dig"]  # label amigável para logs e prints
            linha = OSO["linha"]

            Log.save(f"INICIO OSO: {oso} | Linha: {linha}")

            tabela = [[n_oso ,oso_label, linha]]
            print(tabulate(tabela, headers=["n.º", "OSO", "LINHA"],tablefmt="rounded_grid"))
            print("")

            RepetirOso = True
            while RepetirOso:
                
                interruptorGlobal()
                if parar_execucao:
                    Log.save("Parada solicitada — encerrando preenchimento.")
                    Log.user("Execução encerrada pelo usuário.")
                    return

                # imprime resumo ao usuário: ação principal
                Log.user(f">>> Processando OSO {oso_label}")

                row = osos[idx_oso]

                ## Verifica se tem OSO derivada após a OSO base atual
                #OsoDerivada = False
                #prox_oso = None
                #if row["tipo"] == "BASE":
                #    prox_idx = idx_oso + 1
                #    if prox_idx < len(osos):
                #        prox_oso = osos[prox_idx]
                #        if prox_oso["tipo"] == "DERIVADA":
                #            OsoDerivada = True

                prox_oso = osos[idx_oso + 1]

                # chama a rotina que faz os pyautogui.write / press, etc.
                PrintPDF.imprimirPDF(row, fQH, OsoAtiva)
                
                print("Próxima OSO:")
                n_oso = prox_oso["n_oso"]
                oso = prox_oso["oso"]
                oso_label = prox_oso["oso_dig"]  # label amigável para logs e prints
                linha = prox_oso["linha"]

                tabela = [[n_oso ,oso_label, linha]]
                print(tabulate(tabela, headers=["n.º", "OSO", "LINHA"],tablefmt="rounded_grid"))
                print("")

                Log.save(f"AGUARDANDO_F10_F9_F12 OSO {oso_label}")
                
                # loop leve aguardando tecla
                while True:
                    if keyboard.is_pressed("F12"):
                        Log.save("Parada solicitada durante espera de F10.")
                        Log.user(" >>> Execução encerrada pelo usuário.")
                        return
                    if keyboard.is_pressed("F10"):
                        time.sleep(0.18)
                        Log.save(f" >>> F10 pressionado — preenchendo oso do OSO {oso}")
                        RepetirOso = False
                        break
                    if keyboard.is_pressed("F9"):
                        time.sleep(0.18)
                        Log.save(f" >>> F9 pressionado — repetindo faixa do OSO {oso}")
                        RepetirOso = True
                        break
                    time.sleep(0.06)
            # fim da OSO — segue para próxima OSO no for
            # pequeno delay entre osos
            time.sleep(0.15)
            
        # todos osos processados
        print("")
        Log.user("Todas as OSOs processados.")
        print("")
        time.sleep(1)
        Log.save("PROCESSAMENTO_COMPLETO")

    # -------------------------
    # Selecionar OSO a ser preenchida
    # -------------------------
    @staticmethod
    def seletorOsos(OSOs):
        
        headers = ["n.º", "Linha", "OSO", "Tipo"]
        tabela = [
            [b["n_oso"],b["linha"], b["oso_dig"], b["tipo"]]
            for i, b in enumerate(OSOs)
        ]
        print("")
        print(tabulate(tabela, headers=headers, tablefmt="rounded_grid"))
        print("")
        time.sleep(1)
        
        # escolhe OSO inicial
        oso_inicial = Menu.blocoInicial(OSOs)

        oso_inicial_index=oso_inicial-1
        return OSOs, oso_inicial_index

    @staticmethod
    def imprimirPDF(row, fQH, OsoAtiva):
        # Impressão de um PDF.

        # preserva strings originais
        linha = row.get("linha", row.get("linha ", "")).strip()
        oso = row.get("oso", row.get("oso ", "")).strip()
        oso_dig = row.get("oso_dig", row.get("oso_dig ", "")).strip()
        NomePDF = f"{fQH} {linha} {oso_dig}"
        caminho_pdf = os.path.join("AutoSiget3000", "FQHs", f"{NomePDF}.pdf")

        # --------------------------
        # Início do preenchimento
        # --------------------------

        # Preenche o campo Insira o n.º da OSO (Dígito por dígito, para evitar erros)
        for ch in oso:
            pyautogui.write(ch)
            time.sleep(0.1)
        Log.save(f"[{oso_dig}] escreveu OSO '{oso}'")
        Log.user(f"Preenchendo OSO:     {oso_dig}")
        time.sleep(0.2)

        pyautogui.press('left', presses=6, interval=0.1)
        Log.save(f"[{oso_dig}] 6x seta para esquerda")
        time.sleep(0.1)

        pyautogui.press('tab')
        Log.save(f"[{oso_dig}] Insira o n.º da OSO ≥ Impressão Gráfica")
        time.sleep(0.2)

        time.sleep(0.5)
        pyautogui.press("enter", presses=3, interval=0.3)
        Log.save(f"[{oso_dig}] fechou o POP-UP e botou para imprimir")
        time.sleep(3)

        # Digita o nome do PDF e confirma salvar
        pyautogui.write(NomePDF)
        Log.save(f"[{oso_dig}] escreveu nome do PDF '{NomePDF}'")
        Log.user(f"Preenchendo NomePDF: {NomePDF}")
        time.sleep(0.3)
        
        pyautogui.press("enter")
        Log.save(f"[{oso_dig}] confirmou salvar PDF")
        time.sleep(3.5)

        pyautogui.hotkey('shift', 'tab')
        Log.save(f"[{oso_dig}] botão Imprimir ≥ Informe o N° da oso")
        time.sleep(0.2)

        ''''
        # Cria um PDF fake (arquivo vazio ou com texto de teste)
        with open(caminho_pdf, "w") as f:
            f.write(f"Arquivo fake para debug: {NomePDF}.pdf\n")
        '''
        
        if os.path.exists(caminho_pdf):
            Log.save(f"[{oso_dig}] PDF salvo com sucesso: {NomePDF}.pdf")
            Log.user(f"   {NomePDF}.pdf | Salvo com sucesso")
            Log.user(f"[{oso_dig}] • F10 para continuar | F9 para repetir | F12 para parar ")
        else:
            Log.save(f"[{oso_dig}] ERRO: PDF não encontrado ({NomePDF}.pdf)")
            Log.user("===============================================")
            Log.user(f"   ERRO: {NomePDF}.pdf | PDF não encontrado   ")
            Log.user("===============================================")
            Log.user(f"[{oso_dig}] • F10 para continuar | F9 para repetir | F12 para parar ")

        # final do preenchimento da faixa

    @staticmethod
    def Instrucao():
        # Instruções antes de iniciar o programa
        print("REQUISITOS E INSTRUÇÕES:")
        print('''
        Antes de iniciar preencha os seguintes requisitos:
            [$] Prepare o CSV conforme instruções na documentação
                e deixe na mesma pasta deste programa.
            [$] Defina a impressora padrão como Microsoft Print to PDF.
                *Necessário cada vez que abir o SIGET.
            [$] NÃO exclua, mova ou renomeei a pasta FQHs criada na
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
            [#] Este programa apenas automatiza a impressão das
            faixas e quadros horários (programação da faixa)
            [#] Quaisquer erros e avisos que apareça continua
                as soluções que já são utilizados normalmente
            [#] SEMPRE deve se colocar o cursor no inicio da
                "Informe o n.º da OSO" de cada OSO que será
                impressa
        ''')

        print("Pressione F10 para continuar...")
        keyboard.wait("F10")
        print("")
        return None

def IniciarModulo_printPDF(dados):
    
    log_dir = "AutoSiget3000/FQHs"
    os.makedirs(log_dir, exist_ok=True)

    # Tratar osos e adicionar dados calculados
    OSOs = PrintPDF.tratarOSOs(dados)

    #Instruções para o usuário
    PrintPDF.Instrucao()

    #Seleciona o Tipo de preenchimento
    fQH, OsoAtiva = PrintPDF.SelecionarFQH()

    while True:
        # Iniciar o preenchimento de PDFs
        OSOs, oso_inicial_index = PrintPDF.seletorOsos(OSOs)

        #Tempo para o usuário trocar de tela Programa ≥ SIGET
        print("Aguardando 5s para posicionar o CURSOR no SIGET...", end=",\n")
        time.sleep(5)

        # Preenchimento das OSOs
        PrintPDF.preencher_PDFs(OSOs, oso_inicial_index, fQH, OsoAtiva)

        resposta = Menu.escolha("Encerrar programa", "Iniciar outra OSO")

        if resposta == '1':
            break
        else:
            continue
    
    main()

def IniciarModulo_ProgramaLinha(dados):
    #Inicia o processamento do Módulo de programação de linha
    # gera Blocos
    blocos = ProgramaLinha.genBlocos(dados)

    if not blocos:
        Log.user("Nenhum OSO identificado no CSV.")
        return
    
    # Instruções para o uso
    ProgramaLinha.instrucao()
    
    while True:
        #Seleciona os Blocos
        linhas, blocos, bloco_inicio_index = ProgramaLinha.select_Blocos(dados, blocos)

        #Preenche a partir do bloco selecionado
        gerarBlocos = ProgramaLinha.preenBlocos(linhas, blocos, bloco_inicio_index)

        if gerarBlocos:
            continue
        else:
            resposta = Menu.escolha("Encerrar programa", "Iniciar outra OSO")
            if resposta == "1":
                break
            else:
                continue
    
    main()

# -------------------------
# Interface principal
# -------------------------
def main():

    intro()

    '''Módulo central do programa, menu inicial'''
    dados, cabecalho = LCsv.pedir()
    if dados is None:
        return

    # valida campos mínimos (esperados)
    chavesObrigatorias = {"FaixaInicio", "FaixaFinal", "Intervalo", "Percurso", "TempTerm", "Frota", "Linha", "Dia", "Sentido", "Oso", "LinhaOso"}
    
    # Normaliza chavesCabecalho
    chavesCabecalho = set(h.strip() for h in (cabecalho or []))
    chaveFaltando = chavesObrigatorias - chavesCabecalho
    if chaveFaltando:
        Log.user(f"Erro: o CSV está faltando colunas obrigatórias: {', '.join(chaveFaltando)}")
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
            escolha = input(">>> Digite o número do módulo: ").strip()
            time.sleep(0.5)
            if escolha and escolha.isdigit():
                n = int(escolha)
                if n == 1:
                    IniciarModulo_ProgramaLinha(dados)
                elif n == 2:
                    IniciarModulo_printPDF(dados)
                elif n == 0:
                    main()
                else:
                    Log.user("Entrada inválida")
                    print("")
                    continue
            else:
                Log.user("Entrada inválida")
                print("")
                continue

def intro():
    # Mensagem de introdução ao sistema
    global parar_execucao

    print("")
    print("")
    print("")
    print("")
    print("  |============================================================|")
    print("  |                      AUTOMA SIGET3000                      |")
    print("  |============================================================|")
    print("")
    print("                           Bem Vindo(a)")
    print("                     |=====================|")
    print("")
    print("")
    return None

# -------------------------
# Roda a introdução
# -------------------------
if __name__ == "__main__":
    main()




# ===================
#  Autor: Eude Leal
#  Github: eudeleal
# ===================

# Brasil
# Salvador, BA - 20255