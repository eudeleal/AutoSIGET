#!/usr/bin/env python3
# app_3.py
# Versão refatorada por Eude — unificação CSV e controle por blocos
# Mantém toda a lógica de preenchimento com pyautogui igual ao app_2.py

import csv
import os
import time
import pyautogui
import keyboard
from datetime import datetime
from tabulate import tabulate

# -------------------------
# Configurações / Globals
# -------------------------
parar_execucao = False
LOG_DIR = "LOGs"
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
    """Checa F9 (encerra). Deve ser chamado periodicamente dentro dos loops."""
    global parar_execucao
    if keyboard.is_pressed("F9"):
        parar_execucao = True
        print_user("Tecla F9 detectada — encerrando execução...")

# -------------------------
# Leitura do CSV (DictReader) + normalização das chaves
# -------------------------
def carregar_csv(caminho_csv: str):
    """
    Lê CSV com header; normaliza nomes das chaves (strip) e valores (strip).
    Retorna (lista_de_dicionarios, lista_de_cabecalhos_normalizados) ou (None, None) se arquivo não encontrado.
    """
    if not os.path.exists(caminho_csv):
        print_user(f"Arquivo '{caminho_csv}' não encontrado.")
        return None, None

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
                # Some CSVs have extra spaces; ensure consistent keys
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
# Funções utilitárias de horário
# -------------------------
def formatar_hora_log(h: str):
    """Formato amigável para logs/prints: 'HHMM' -> 'HH:MM' """
    h = (h or "").strip()
    if len(h) == 4 and h.isdigit():
        return f"{h[:2]}:{h[2:]}"
    if len(h) == 3 and h.isdigit():
        return f"0{h[:1]}:{h[1:]}"
    return h

def formatar_hora_dig(h: str):
    """Formato para digitação (mantém '000' -> '000', '044' -> '0044' logic handled elsewhere)."""
    h = (h or "").strip()
    if len(h) == 3 and h.isdigit():
        return f"0{h}"
    return h

# -------------------------
# Preenchimento de uma faixa
# -------------------------
def preencher_faixa(row, linha_label, faixaAnterior=None):
    """
    Executa os comandos de pyautogui para preencher uma faixa com base em 'row' (dict).
    Recebe faixaAnterior (tupla de strings (FaixaInicio, FaixaFinal)) para detectar virada de dia.
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
    print_user(f"{linha_label} • Preenchendo faixa de {faixaIni_log} à {faixaFim_log}")

    pyautogui.press("tab")
    registrar_log(f"[{linha_label}] tab → campo FaixaFinal")

    # Detecta virada de dia: se faixaAnterior existe e hora atual < anterior (0000 case)
    if faixaAnterior:
        prev_ini = (faixaAnterior[0] or "").strip()
        if prev_ini.isdigit() and faixaIni_raw.isdigit():
            if int(faixaIni_raw) < int(prev_ini):
                registrar_log(f"[{linha_label}] virada detectada em FaixaInicio ({faixaIni_raw} < {prev_ini}) — confirmando Enter")
                time.sleep(0.25)
                pyautogui.press("enter")
                time.sleep(0.4)

    pyautogui.write(faixaFim_dig)
    registrar_log(f"[{linha_label}] escreveu FaixaFinal '{faixaFim_log}'")

    pyautogui.press("tab")
    registrar_log(f"[{linha_label}] tab → campo Intervalo")

    # Detecta virada de dia no FaixaFinal em relação ao anterior
    if faixaAnterior:
        prev_fim = (faixaAnterior[1] or "").strip()
        if prev_fim.isdigit() and faixaFim_raw.isdigit():
            if int(faixaFim_raw) < int(prev_fim):
                registrar_log(f"[{linha_label}] virada detectada em FaixaFinal ({faixaFim_raw} < {prev_fim}) — confirmando Enter")
                time.sleep(0.25)
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
    print_user(f"{linha_label} • Faixa preenchida | F10 para continuar (F9 para parar)")

# -------------------------
# Preenchimento de blocos (controle de interação F10/F9)
# -------------------------
def preencher_blocos(linhas, blocos, bloco_inicio_index):
    """
    Percorre blocos a partir de bloco_inicio_index (0-based index na lista 'blocos').
    Para cada faixa dentro do bloco chama preencher_faixa().
    Aguarda F10 entre faixas; F9 encerra tudo.
    """
    global parar_execucao

    total_blocos = len(blocos)
    # itera blocos a partir da escolha do usuário
    for idx_bloco, bloco in enumerate(blocos[bloco_inicio_index:], start=bloco_inicio_index + 1):
        id_bloco = bloco["id_bloco"]
        inicio = bloco["inicio"]
        fim = bloco["fim"]
        count = bloco["count"]

        print_user(f" >>> Iniciando bloco {id_bloco} ({count} faixas horárias)")
        registrar_log(f"INICIO_BLOCO {id_bloco} linhas {inicio+1}-{fim+1}")

        # itera dentro do bloco: trecho de linhas
        trecho = linhas[inicio:fim+1]
        faixaAnterior = None  # para detectar virada comparando com a faixa anterior dentro do bloco

        for offset, row in enumerate(trecho):
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
            preencher_faixa(row, linha_label, faixaAnterior=faixaAnterior)

            # atualiza faixaAnterior para a próxima iteração (usa raw values)
            faixaAnterior = (row.get("FaixaInicio", ""), row.get("FaixaFinal", ""))

            # controle de avanço: se não é a última faixa do bloco, espera F10;
            # se for a última, espera F10 para ir ao próximo bloco (ou F9)
            is_last = (offset == len(trecho) - 1)
            if not is_last:
                #print_user(" >>> Aguardando F10 para próxima faixa (F9 para sair)")
                registrar_log(f"AGUARDANDO_F10 faixa {offset+1} do bloco {id_bloco}")
                # loop leve aguardando tecla
                while True:
                    escutador_tecla()
                    if parar_execucao:
                        registrar_log("Parada solicitada durante espera de F10.")
                        print_user(" >>> Execução encerrada pelo usuário.")
                        return
                    if keyboard.is_pressed("F10"):
                        # debouncing
                        time.sleep(0.18)
                        registrar_log(f" >>> F10 pressionado — avançando para faixa {offset+2} do bloco {id_bloco}")
                        break
                    time.sleep(0.06)
            else:
                # último da lista do bloco
                print("")
                print_user(f" [+] Bloco {id_bloco} finalizado ({count} faixas horárias).  [+] ")
                registrar_log(f"FIM_BLOCO {id_bloco}")
                print("")

                # Verifica se há um próximo bloco
                prox_idx = idx_bloco - bloco_inicio_index  # posição relativa no enumerate
                if prox_idx < total_blocos - 1:
                    prox_id_bloco = blocos[bloco_inicio_index + prox_idx]["id_bloco"]
                    print_user(f">>> Próximo bloco: {prox_id_bloco}")
                else:
                    prox_id_bloco = None
                    print_user(">>> Este foi o último bloco.")
                print_user(">>> Pressione F10 para iniciar o próximo bloco ou F9 para encerrar.")
                # Espera F10 ou F9
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
                    time.sleep(0.06)


        # fim do bloco — segue para próximo bloco no for
        # pequeno delay entre blocos
        time.sleep(0.15)

    # todos blocos processados
    print_user("🟢 Todos os blocos processados.")
    registrar_log("PROCESSAMENTO_COMPLETO")

# -------------------------
# Interface principal (CLI)
# -------------------------
def main():
    global parar_execucao

    print("=" * 60)
    print("                    AUTOMA SIGET3000")
    print("=" * 60)
    print("Bem Vindo(a)")
    print("Insira o nome do arquivo para iniciar")
    
    # escolha do CSV
    caminho_csv = input("Nome do CSV (sem .csv): ").strip() + ".csv"
    print_user(f"Arquivo selecionado: {caminho_csv}")

    linhas, cabecalho = carregar_csv(caminho_csv)
    if linhas is None:
        return

    # valida campos mínimos (esperados)
    chavesObrigatorias = {"FaixaInicio", "FaixaFinal", "Intervalo", "Percurso", "TempTerm", "Frota", "Linha", "Dia", "Sentido"} # "Tipo", "NovaOso"
    chavesCabecalho = set(k for k in (cabecalho or []))
    # make chavesCabecalho normalized (strip)
    chavesCabecalho = set(h.strip() for h in (cabecalho or []))
    chaveFaltando = chavesObrigatorias - chavesCabecalho
    if chaveFaltando:
        print_user(f"Erro: o CSV está faltando colunas obrigatórias: {', '.join(chaveFaltando)}")
        return

    # gera blocos
    blocos = gerar_blocos(linhas)
    if not blocos:
        print_user("Nenhum bloco identificado no CSV.")
        return

    # exibe blocos (resumido)
    print("")
    print("Blocos detectados:")
    for idx, b in enumerate(blocos, start=1):
        print(f"[{idx}] {b['id_bloco']} → linhas ({b['count']} faixas horárias)")
    print("")

    # escolhe bloco inicial
    bloco_inicial = 1
    escolha = input(f"Digite o número do bloco para começar (1–{len(blocos)}) ou ENTER para o 1º: ").strip()
    if escolha:
        try:
            n = int(escolha)
            if 1 <= n <= len(blocos):
                bloco_inicial = n
            else:
                print_user("Entrada inválida — iniciando do primeiro bloco.")
        except:
            print_user("Entrada inválida — iniciando do primeiro bloco.")

    # confirma e inicia processamento
    print_user(f"Iniciando a partir do bloco {bloco_inicial}: {blocos[bloco_inicial-1]['id_bloco']}")
    registrar_log(f"USUARIO iniciou processamento a partir do bloco {bloco_inicial}")

    # chama rotina de preenchimento
    preencher_blocos(linhas, blocos, bloco_inicio_index=bloco_inicial-1)

# -------------------------
# Roda o main
# -------------------------
if __name__ == "__main__":
    main()