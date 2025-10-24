
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
nome_arquivo_log = os.path.join(LOG_DIR, f"log_{datetime.now().strftime('%Y.%m.%d_%H.%M.%S')}.txt")

# -------------------------
# Utilitários de Log
# -------------------------
def registrar_log(mensagem: str):
    """Registra eventos detalhados no arquivo de log (arquivo)."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    linha = f"[{timestamp}] {mensagem}"
    # grava imediatamente
    with open(nome_arquivo_log, "a", encoding="utf-8") as f:
        f.write(linha + "\n")

def print_user(msg: str):
    """Prints sucinto para o usuário (console). Também registra um resumo no log."""
    print(msg)
    # Log mais sucinto: registra que foi mostrado algo ao usuário
    registrar_log(f"[UI] {msg}")

# -------------------------
# Escutador (marca parar_execucao)
# -------------------------
def escutador_tecla():
    """Checa F12 (encerra). Deve ser chamado periodicamente dentro dos loops."""
    global parar_execucao
    if keyboard.is_pressed("F12"):
        parar_execucao = True
        print_user("Tecla F12 detectada — encerrando execução.")
        print_user("Retornando ao menu.")
        print("")
        print("")
        print("")
        time.sleep(5)
        parar_execucao = False
        main()

# -------------------------
# Leitura do CSV (DictReader) + normalização das chaves
# -------------------------
def pedirCSV():
    '''Função auxiliar para solicitar o nome do arquivo .csv ao usuário'''
    
    print("")
    print("Insira o nome do arquivo para iniciar, ENTER para confirmar o nome.")
    
    # Escolha do CSV
    caminho_csv = input("Nome do CSV (sem .csv): ").strip() + ".csv"
    print_user(f"Arquivo selecionado: {caminho_csv}")
    
    linhas, fieldnames = carregar_csv(caminho_csv)
    return linhas, fieldnames

def carregar_csv(caminho_csv: str):
    """
    Lê CSV com header; normaliza nomes das chaves (strip) e valores (strip).
    Retorna (lista_de_dicionarios, lista_de_cabecalhos_normalizados) ou (None, None) se arquivo não encontrado.
    """
    
    if not os.path.exists(caminho_csv):
        print_user(f"Arquivo '{caminho_csv}' não encontrado.")
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

    registrar_log(f"CSV '{caminho_csv}' carregado. Registros: {len(linhas)}. Campos: {fieldnames}")
    return linhas, fieldnames

# -------------------------
# Geração dos blocos (id_bloco = Linha_Dia_Sentido)
# -------------------------
def gerar_blocos(linhas):
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

    registrar_log(f"{len(blocos)} blocos gerados.")
    return blocos

# -------------------------
# Seleção dos blocos
# -------------------------
def seletorBlocos(linhas, blocos):
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
    #for idx, b in enumerate(blocos, start=1):
    #    print(f"[{idx}] {b['id_bloco']} → linhas ({b['count']} faixas horárias)")
    print("")
    time.sleep(1)

    # escolhe bloco inicial
    bloco_inicial = 1
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
                print_user("Entrada inválida")
                print("")
                continue
        else:
            print_user("Entrada inválida")
            print("")
            continue

    # confirma e inicia processamento
    print_user(f"Iniciando a partir do bloco {bloco_inicial}: {blocos[bloco_inicial-1]['id_bloco']}")
    registrar_log(f"USUARIO iniciou processamento a partir do bloco {bloco_inicial}")

    # chama rotina de preenchimento
    preencher_blocos(linhas, blocos, oso_inicial_index=bloco_inicial-1)

# -------------------------
# Funções utilitárias de horário
# -------------------------
def formatar_hora_log(h: str):
    """Formato para logs/prints: 'HHMM' -> 'HH:MM' """
    h = (h or "").strip()
    if len(h) == 4 and h.isdigit():
        return f"{h[:2]}:{h[2:]}"
    if len(h) == 3 and h.isdigit():
        return f"0{h[:1]}:{h[1:]}"
    return h

def formatar_hora_dig(h: str):
    """Formato para digitação (mantém '000' -> '000', '044' -> '0044')."""
    h = (h or "").strip()
    if len(h) == 3 and h.isdigit():
        return f"0{h}"
    return h

# -------------------------
# Preenchimento de uma faixa
# -------------------------
def preencher_faixa(row, linha_label, faixaFimAnterior=None):
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
    faixaIni_log = formatar_hora_log(faixaIni_raw)
    faixaFim_log = formatar_hora_log(faixaFim_raw)

    # prepara valores para digitar (mantendo a lógica existente)
    faixaIni_dig = formatar_hora_dig(faixaIni_raw)
    faixaFim_dig = formatar_hora_dig(faixaFim_raw)

    # --------------------------
    # Início do preenchimento (mesma sequência de ações)
    pyautogui.write(faixaIni_dig)
    registrar_log(f"[{linha_label}] escreveu FaixaInicio '{faixaIni_log}'")
    
    # console
        #print_user(f"{linha_label} • Preenchendo faixa de {faixaIni_log} à {faixaFim_log}")
    print(tabulate([[faixaIni_log, faixaFim_log, interv, tPerc, tTerm, frota]],
               headers=["F. Inicio", "F. Final", "Interv.", "T. Perc.", "T. Term", "Frota"],
               tablefmt="rounded_grid"))

    pyautogui.press("tab")
    registrar_log(f"[{linha_label}] tab → campo FaixaFinal")

    # variável para evitar dupla verificação de faixa para o dia seguinte
    faixaIniVerif = False

    # Detecta virada de dia: se faixaFimAnterior existe e hora atual < anterior (0000 case)
    if faixaFimAnterior:
        prev_ini = (faixaFimAnterior[0] or "").strip()
        if prev_ini.isdigit() and faixaIni_raw.isdigit():
            if int(faixaIni_raw) < int(prev_ini):
                registrar_log(f"[{linha_label}] virada detectada em Faixa Inicio ({faixaIni_raw} < {prev_ini}) — confirmando Enter")
                time.sleep(0.25)
                pyautogui.press("tab")
                pyautogui.press("enter")
                time.sleep(0.4)
                faixaIniVerif = True
                registrar_log(f"[{linha_label}] faixaIniVerif definida como VERDADEIRA")

    pyautogui.write(faixaFim_dig)
    registrar_log(f"[{linha_label}] escreveu FaixaFinal '{faixaFim_log}'")

    pyautogui.press("tab")
    registrar_log(f"[{linha_label}] tab → campo Intervalo")

    # Detecta virada de dia no FaixaFinal em relação ao anterior
    if faixaFimAnterior and not faixaIniVerif:
        prev_fim = faixaIni_dig
        if prev_fim.isdigit() and faixaFim_dig.isdigit():
            if int(faixaFim_dig) < int(prev_fim):
                registrar_log(f"[{linha_label}] virada detectada em FaixaFinal ({faixaFim_dig} < {prev_fim}) — confirmando Enter")
                time.sleep(0.25)
                pyautogui.press("tab")
                pyautogui.press("enter")
                time.sleep(0.4)

    pyautogui.write(interv)
    registrar_log(f"[{linha_label}] escreveu Intervalo '{interv}'")
    time.sleep(0.1)

    pyautogui.press("tab")
    registrar_log(f"[{linha_label}] tab → campo Tempo de Percurso")
    time.sleep(0.1)

    pyautogui.write(tPerc)
    registrar_log(f"[{linha_label}] escreveu Percurso '{tPerc}'")
    time.sleep(0.1)

    pyautogui.press("tab")
    registrar_log(f"[{linha_label}] tab → campo Terminal")
    time.sleep(0.1)

    pyautogui.write(tTerm)
    registrar_log(f"[{linha_label}] escreveu TempTerm '{tTerm}'")
    time.sleep(0.1)

    pyautogui.press("tab")
    registrar_log(f"[{linha_label}] tab → campo Frota")
    time.sleep(0.1)

    pyautogui.write(frota)
    registrar_log(f"[{linha_label}] escreveu Frota '{frota}'")
    time.sleep(0.1)

    pyautogui.press("tab")
    registrar_log(f"[{linha_label}] tab → campo Linha (pronto para verificação)")
    time.sleep(0.1)

    # final do preenchimento da faixa
    registrar_log(f"[{linha_label}] faixa preenchida ({faixaIni_log} → {faixaFim_log})")
    print_user(f"{linha_label} • Faixa preenchida | F10 para continuar | F9 para repetir | F12 para parar ")

# -------------------------
# Preenchimento de blocos (controle de interação F10/F12)
# -------------------------
def preencher_blocos(linhas, blocos, bloco_inicio_index):
    """
    Percorre blocos a partir de bloco_inicio_index (0-based index na lista 'blocos').
    Para cada faixa dentro do bloco chama preencher_faixa().
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

        registrar_log(f"INICIO_BLOCO {id_bloco} linhas {inicio+1}-{fim+1}")

        print(tabulate([[id_bloco]],
                       headers=["Bloco"],
                       tablefmt="rounded_grid"))

        # itera dentro do bloco: trecho de linhas
        trecho = linhas[inicio:fim+1]
        faixaFimAnterior = None  # para detectar virada comparando com a faixa anterior dentro do bloco

        offset = 0 # Inicializa o índice relativo dentro do 'trecho'
        while offset < count: # count é o len(trecho)

            row = trecho[offset]
            
            escutador_tecla()
            if parar_execucao:
                registrar_log("Parada solicitada — encerrando preenchimento.")
                print_user("Execução encerrada pelo usuário.")
                return

            i_global = inicio + offset
            linha_label = bloco["id_bloco"]  # label amigável para logs e prints

            # timer antes de iniciar na primeira vez
            if offset == 0:
                print_user("")
                print_user("  [+] Tempo de 5s para ajustar o cusor no SIGET [+] ")
                print_user("")
                time.sleep(5)

            # imprime resumo ao usuário: ação principal
            print_user(f">>> Processando faixa {offset+1}/{count}")

            # chama a rotina que faz os pyautogui.write / press etc.
            preencher_faixa(row, linha_label, faixaFimAnterior=faixaFimAnterior)

            # atualiza faixaFimAnterior para a próxima iteração (usa raw values)
            faixaFimAnterior = (row.get("FaixaFim", ""))

            # controle de avanço: se não é a última faixa do bloco, espera F10;
            # se for a última, espera F10 para ir ao próximo bloco (ou F12)
            is_last = (offset == count - 1)
            
            registrar_log(f"AGUARDANDO_F10_F9_F12 faixa {offset+1} do bloco {id_bloco}")
            
            # loop leve aguardando tecla
            while True:
                escutador_tecla()
                if parar_execucao:
                    registrar_log("Parada solicitada durante espera de F10.")
                    print_user(" >>> Execução encerrada pelo usuário.")
                    return
                if keyboard.is_pressed("F10"):
                    time.sleep(0.18)
                    registrar_log(f" >>> F10 pressionado — avançando para faixa {offset+2} do bloco {id_bloco}")
                    offset += 1 
                    break
                if keyboard.is_pressed("F9"):
                    time.sleep(0.18)
                    registrar_log(f" >>> F9 pressionado — repetindo faixa {offset+1} do bloco {id_bloco}")
                    is_last = False
                    break
                time.sleep(0.06)

            if is_last:
                # último da lista do bloco
                offset += 1

                print("")
                print_user(f" [+] Bloco {id_bloco} finalizado ({count} faixas horárias).  [+] ")
                registrar_log(f"FIM_BLOCO {id_bloco}")
                print("")
                
                # Verifica se há um próximo bloco
                prox_idx = idx_bloco - bloco_inicio_index  # posição relativa no enumerate
                if prox_idx < total_blocos - 1:
                    prox_id_bloco = blocos[bloco_inicio_index + prox_idx + 1]["id_bloco"]
                    registrar_log(f">>> Próximo bloco: {prox_id_bloco}")
                    print(tabulate([[prox_id_bloco, len(prox_id_bloco)]],
               headers=["Prox. Bloco", "Qtd Faixas"],
               tablefmt="rounded_grid"))
                else:
                    prox_id_bloco = None
                    print_user(">>> Este foi o último bloco.")
                print_user(">>> Pressione F10 para iniciar o próximo bloco | F9 para selecionar bloco | F12 para encerrar")
                # Espera F10, F9 ou F12
                while True:
                    escutador_tecla()
                    if parar_execucao:
                        registrar_log("Parada solicitada na finalização de bloco.")
                        print_user("Execução encerrada pelo usuário.")
                        return
                    if keyboard.is_pressed("F10"):
                        time.sleep(0.18)
                        if prox_id_bloco:
                            registrar_log(f" >>> F10 pressionado — iniciando próximo bloco ({prox_id_bloco})")
                        else:
                            registrar_log(" >>> F10 pressionado — não há próximo bloco (fim).")
                        break
                    if keyboard.is_pressed("F9"):
                        time.sleep(0.18)
                        registrar_log(f" >>> F9 pressionado iniciando seletor de blocos")
                        seletorBlocos(linhas, blocos)
                        break
                    time.sleep(0.06)
        
        # fim do bloco — segue para próximo bloco no for
        # pequeno delay entre blocos
        time.sleep(0.15)
        
    # todos blocos processados
    print_user("Todos os blocos processados.")
    registrar_log("PROCESSAMENTO_COMPLETO")
    main()

# -------------------------
# Interface principal
# -------------------------
def main():
    '''Módulo central do programa, menu inicial'''
    dados, cabecalho = pedirCSV()
    if dados is None:
        return

    # valida campos mínimos (esperados)
    chavesObrigatorias = {"FaixaInicio", "FaixaFinal", "Intervalo", "Percurso", "TempTerm", "Frota", "Linha", "Dia", "Sentido", "Oso", "LinhaOso"}
    chavesCabecalho = set(k for k in (cabecalho or []))
    
    # Normaliza chavesCabecalho
    chavesCabecalho = set(h.strip() for h in (cabecalho or []))
    chaveFaltando = chavesObrigatorias - chavesCabecalho
    if chaveFaltando:
        print_user(f"Erro: o CSV está faltando colunas obrigatórias: {', '.join(chaveFaltando)}")
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
                    IniciarModulo_ProgramacaoLinha(dados)
                elif n == 2:
                    IniciarModulo_ImprimirPDFs(dados)
                elif n == 0:
                    main()
                else:
                    print_user("Entrada inválida")
                    print("")
                    continue
            else:
                print_user("Entrada inválida")
                print("")
                continue

def IniciarModulo_ImprimirPDFs(dados):
    
    LOG_DIR = "AutoSiget3000/FQHs"
    os.makedirs(LOG_DIR, exist_ok=True)

    # Tratar osos e adicionar dados calculados
    OSOs = tratarOSOs(dados)

    print(OSOs)



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

    # Iniciar o preenchimento de PDFs
    preencher_PDFs(OSOs) 

def IniciarModulo_ProgramacaoLinha(dados):
    '''Inicia o processamento do Módulo de programação de linha'''
    # gera blocos
    blocos = gerar_blocos(dados)
    if not blocos:
        print_user("Nenhum bloco identificado no CSV.")
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

    #Seleciona os blocos
    seletorBlocos(dados, blocos)

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

#=======================================#

# -------------------------
# Normalização de OSOs
# -------------------------
def tratarOSOs(osoBruta):
    """
    Verifica a formatação das OSOs [123456], 
    identifica se é Base ou Derivada e 
    relaciona com a respectiva Linha 
    """

    osos = []

    for row in enumerate(osoBruta):
        
        linha = row.get("Linha", row.get("Linha ", "")).strip()
        oso = row.get("Oso", row.get("Oso ", "")).strip()

        if len(oso) == 6 and oso.isdigit():
            if oso[2:] == "00":
                osos.append({
                    "linha": linha,
                    "osos": oso,
                    "tipo": "B"
                })
            else:
                osos.append({
                    "linha": linha,
                    "osos": oso,
                    "tipo": "D"
                })
        else:
            print_user(f"OSO [{osoBruta}] | formato inválido seguir o formato: [123456]")
            return
            
    filtoB = lambda x: x["tipo"] == "B"
    filtoD = lambda x: x["tipo"] == "D"
    ososBase = list(filter(filtoB, osos))
    ososDerivada = list(filter(filtoD, osos))

    print_user(f"{len(osos)} osos idenficadas. | B:{len(ososBase)} D:{len(ososDerivada)} ")
    return osos

# -------------------------
# Impressão de um PDF
# -------------------------
def imprimirPDF(row, oso_label):
    """
    EXPLICAÇÃO AQUI

    """

    # preserva strings originais
    linha = row.get("Linha", row.get("Linha ", "")).strip()
    oso = row.get("Oso", row.get("Oso ", "")).strip()
    tipo = row.get("Tipo", row.get("Tipo ", "")).strip()
    fQH = row.get("FQH", row.get("FQH ", "")).strip()
    NomePDF = f"{fQH} {linha} {oso}"

    # --------------------------
    # Início do preenchimento
    
    # console
    print(tabulate([[fQH, linha]],
               headers=["FQH", "LINHA"],
               tablefmt="rounded_grid"))

    '''
    
    SEQUENCIA DE:

    pyautogui.press("tab")
    registrar_log(f"[{oso_label}] tab → campo OSO")

    pyautogui.write()
    registrar_log(f"[{oso_label}] escreveu OSO '{oso}'")
    time.sleep(0.1)

    LÓGICA:

    Tab     | Informe o N° da OSO -> Imprimir em

    if OsoAtiva:
        # Aviso de OSO filha detectada
        Tab     | SIM -> NÃO
        Enter   | Confirma NÃO

    Tab     | Imprimir em -> Configurar
    Tab     | Configurar -> Imprimir
    Enter   | Confirma Imprimir

    DIGITA NomePdf

    Enter   | Confirma SALVAR

    if PdfExiste:
        # PDF salvo com sucesso
    else:
        # PDF não encontrado, verifique a pasta SALVAR

        # Escolha repetir ou continuar
    
    # Escolha repetir ou continuar

    '''

    # final do preenchimento da faixa
    registrar_log(f"[{oso_label}] PDF {NomePDF} lançado")
    print_user(f"{oso_label} • OSO lançada | F10 para continuar | F9 para repetir | F12 para parar ")

# -------------------------
# Selecionar OSO a ser preenchida
# -------------------------
def seletorOsos(linhaOso, osos):
    
    # Escolha  a OSO
    
    oso_inicial_index = 0
    preencher_PDFs(linhaOso, osos, oso_inicial_index)

# -------------------------
# Impressão de OSOs (controle de interação F10/F12)
# -------------------------
def preencher_PDFs(linhaOso, osos, oso_inicial_index):
    """
    Percorre osos a partir de oso_inicial_index (0-based index na lista 'osos').
    Para cada faixa dentro do bloco chama preencher_faixa().
    Aguarda F10 entre faixas; F12 encerra tudo.
    """

    global parar_execucao

    total_osos = len(osos)
    # itera osos a partir da escolha do usuário
    for idx_bloco, bloco in enumerate(osos[oso_inicial_index:], start=oso_inicial_index): # Mantive 'start' ajustado
        id_bloco = bloco["id_bloco"]
        inicio = bloco["inicio"]
        fim = bloco["fim"]
        count = bloco["count"]

        registrar_log(f"INICIO_BLOCO {id_bloco} linhaOso {inicio+1}-{fim+1}")

        print(tabulate([[id_bloco]],
                       headers=["Bloco"],
                       tablefmt="rounded_grid"))

        # itera dentro do bloco: trecho de linhas
        trecho = linhaOso[inicio:fim+1]
        faixaFimAnterior = None  # para detectar virada comparando com a faixa anterior dentro do bloco

        offset = 0 # Inicializa o índice relativo dentro do 'trecho'
        while offset < count: # count é o len(trecho)

            row = trecho[offset]
            
            escutador_tecla()
            if parar_execucao:
                registrar_log("Parada solicitada — encerrando preenchimento.")
                print_user("Execução encerrada pelo usuário.")
                return

            i_global = inicio + offset
            oso_label = bloco["id_bloco"]  # label amigável para logs e prints

            # timer antes de iniciar na primeira vez
            if offset == 0:
                print_user("")
                print_user("  [+] Tempo de 5s para ajustar o cusor no SIGET [+] ")
                print_user("")
                time.sleep(5)

            # imprime resumo ao usuário: ação principal
            print_user(f">>> Processando faixa {offset+1}/{count}")

            # chama a rotina que faz os pyautogui.write / press etc.
            preencher_faixa(row, oso_label, faixaFimAnterior=faixaFimAnterior)

            # atualiza faixaFimAnterior para a próxima iteração (usa raw values)
            faixaFimAnterior = (row.get("FaixaFim", ""))

            # controle de avanço: se não é a última faixa do bloco, espera F10;
            # se for a última, espera F10 para ir ao próximo bloco (ou F12)
            is_last = (offset == count - 1)
            
            registrar_log(f"AGUARDANDO_F10_F9_F12 faixa {offset+1} do bloco {id_bloco}")
            
            # loop leve aguardando tecla
            while True:
                escutador_tecla()
                if parar_execucao:
                    registrar_log("Parada solicitada durante espera de F10.")
                    print_user(" >>> Execução encerrada pelo usuário.")
                    return
                if keyboard.is_pressed("F10"):
                    time.sleep(0.18)
                    registrar_log(f" >>> F10 pressionado — avançando para faixa {offset+2} do bloco {id_bloco}")
                    offset += 1 
                    break
                if keyboard.is_pressed("F9"):
                    time.sleep(0.18)
                    registrar_log(f" >>> F9 pressionado — repetindo faixa {offset+1} do bloco {id_bloco}")
                    is_last = False
                    break
                time.sleep(0.06)

            if is_last:
                # último da lista do bloco
                offset += 1

                print("")
                print_user(f" [+] Bloco {id_bloco} finalizado ({count} faixas horárias).  [+] ")
                registrar_log(f"FIM_BLOCO {id_bloco}")
                print("")
                
                # Verifica se há um próximo bloco
                prox_idx = idx_bloco - oso_inicial_index  # posição relativa no enumerate
                if prox_idx < total_osos - 1:
                    prox_id_bloco = osos[oso_inicial_index + prox_idx + 1]["id_bloco"]
                    registrar_log(f">>> Próximo bloco: {prox_id_bloco}")
                    print(tabulate([[prox_id_bloco, len(prox_id_bloco)]],
               headers=["Prox. Bloco", "Qtd Faixas"],
               tablefmt="rounded_grid"))
                else:
                    prox_id_bloco = None
                    print_user(">>> Este foi o último bloco.")
                print_user(">>> Pressione F10 para iniciar o próximo bloco | F9 para selecionar bloco | F12 para encerrar")
                # Espera F10, F9 ou F12
                while True:
                    escutador_tecla()
                    if parar_execucao:
                        registrar_log("Parada solicitada na finalização de bloco.")
                        print_user("Execução encerrada pelo usuário.")
                        return
                    if keyboard.is_pressed("F10"):
                        time.sleep(0.18)
                        if prox_id_bloco:
                            registrar_log(f" >>> F10 pressionado — iniciando próximo bloco ({prox_id_bloco})")
                        else:
                            registrar_log(" >>> F10 pressionado — não há próximo bloco (fim).")
                        break
                    if keyboard.is_pressed("F9"):
                        time.sleep(0.18)
                        registrar_log(f" >>> F9 pressionado iniciando seletor de osos")
                        seletorOsos(linhaOso, osos)
                        break
                    time.sleep(0.06)
        
        # fim do bloco — segue para próximo bloco no for
        # pequeno delay entre osos
        time.sleep(0.15)
        
    # todos osos processados
    print_user("Todos os osos processados.")
    registrar_log("PROCESSAMENTO_COMPLETO")
    main()








# ===================
#  Autor: Eude Leal
#  Github: eudeleal
# ===================

# Brasil
# Salvador, BA - 20255