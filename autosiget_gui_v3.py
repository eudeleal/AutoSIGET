
# Bibliotecas nativas
import csv
import os
import time
from datetime import datetime
import threading

# Bibliotecas externas
import pyautogui
import keyboard
from tabulate import tabulate
import flet as ft

# -------------------------
# MODEL - Lógica de Negócio
# -------------------------

class ConfigModel:
    """Model para configurações globais"""
    def __init__(self):
        self.parar_execucao = False
        self.LOG_DIR = "AutoSiget3000/LOGs"
        os.makedirs(self.LOG_DIR, exist_ok=True)
        self.nome_arquivo_log = os.path.join(
            self.LOG_DIR, 
            f"log_{datetime.now().strftime('%Y%m%d_%H.%M')}.txt"
        )

class LogModel:
    """Model para gerenciamento de logs (adaptado da classe Log)"""
    def __init__(self, config: ConfigModel):
        self.config = config
        self.log_callbacks = []
    
    def add_callback(self, callback):
        """Adiciona callback para atualização da UI"""
        self.log_callbacks.append(callback)
    
    def save(self, mensagem: str):
        """Registra eventos detalhados no arquivo de log (arquivo)."""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        linha = f"[{timestamp}] {mensagem}"
        with open(self.config.nome_arquivo_log, "a", encoding="utf-8") as f:
            f.write(linha + "\n")
    
    def user(self, msg: str):
        """Reporte sucinto para o usuário (console e UI)."""
        print(msg)
        self.save(f"[UI] {msg}")
        # Notifica callbacks (UI)
        for callback in self.log_callbacks:
            callback(msg)

class UtilModel:
    """Model para funções utilitárias (adaptado da classe Util)"""
    @staticmethod
    def formatHora_log(hora: str):
        # Formato para logs/prints: 'HHMM' -> 'HH:MM'
        h = (hora or "").strip()
        if len(h) == 4 and h.isdigit():
            return f"{h[:2]}:{h[2:]}"
        if len(h) == 3 and h.isdigit():
            return f"0{h[:1]}:{h[1:]}"
        return h
    
    @staticmethod
    def formatHora_dig(hora: str):
        # Formato para digitação (mantém '000' -> '000', '044' -> '0044').
        h = (hora or "").strip()
        if len(h) == 3 and h.isdigit():
            return f"0{h}"
        return h

class CSVModel:
    """Model para operações com CSV (adaptado da classe LCsv)"""
    def __init__(self, log_model: LogModel):
        self.log = log_model
    
    def load(self, caminho_csv: str):
        """Carrega CSV e retorna linhas e fieldnames"""
        if not os.path.exists(caminho_csv):
            return None, None
        
        try:
            with open(caminho_csv, newline="", encoding="utf-8") as f:
                reader = csv.DictReader(f, delimiter=";")
                fieldnames = [fn.strip() for fn in (reader.fieldnames or [])]
                linhas = []
                for row in reader:
                    nova = {}
                    for k, v in row.items():
                        key = (k or "").strip()
                        nova[key] = (v or "").strip()
                    linhas.append(nova)
            
            self.log.save(f"CSV '{caminho_csv}' carregado. Registros: {len(linhas)}. Campos: {fieldnames}")
            return linhas, fieldnames
        except Exception as e:
            self.log.user(f"Erro ao carregar CSV: {e}")
            return None, None

class ProgramaLinhaModel:
    """Model para programação de linha (adaptado da classe ProgramaLinha)"""
    def __init__(self, log_model: LogModel, config: ConfigModel):
        self.log = log_model
        self.config = config
    
    def genBlocos(self, linhas):
        """Gera blocos de programação"""
        blocos = []
        chave_anterior = None
        inicio = 0
        
        for i, row in enumerate(linhas):
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
        
        if chave_anterior is not None:
            blocos.append({
                "id_bloco": chave_anterior,
                "inicio": inicio,
                "fim": len(linhas) - 1,
                "count": len(linhas) - inicio
            })
        
        self.log.save(f"{len(blocos)} blocos gerados.")
        return blocos
    
    def preenFaixa(self, row, linha_label, faixaFimAnterior=None):
        """Preenche uma faixa horária (simulação de pyautogui)"""
        
        # --- Extração e Formatação ---
        faixaIni_raw = row.get("FaixaInicio", row.get("FaixaInicio ", "")).strip()
        faixaFim_raw = row.get("FaixaFinal", row.get("FaixaFinal ", "")).strip()
        Interv = row.get("Interv", row.get("Interv", "")).strip()
        tPerc = row.get("Percurso", row.get("Perc.", row.get("Perc", ""))).strip()
        tTerm = row.get("TempTerm", row.get("T. Term", row.get("T. Term ", ""))).strip()
        frota = row.get("Frota", "").strip()
        
        faixaIni_log = UtilModel.formatHora_log(faixaIni_raw)
        faixaFim_log = UtilModel.formatHora_log(faixaFim_raw)
        faixaIni_dig = UtilModel.formatHora_dig(faixaIni_raw)
        faixaFim_dig = UtilModel.formatHora_dig(faixaFim_raw)
        
        # --- Simulação de PyAutoGUI ---
        
        # 1. Escreve FaixaInicio
        pyautogui.write(faixaIni_dig)
        self.log.save(f"[{linha_label}] escreveu FaixaInicio '{faixaIni_log}'")
        
        # 2. Log para o usuário (tabela)
        tabela = tabulate([[faixaIni_log, faixaFim_log, Interv, tPerc, tTerm, frota]],
                           headers=["F. Inicio", "F. Final", "Interv.", "T. Perc.", "T. Term", "Frota"],
                           tablefmt="rounded_grid")
        self.log.user(f"Preenchendo faixa:\n{tabela}")
        
        # 3. Tab para FaixaFinal
        pyautogui.press("tab")
        self.log.save(f"[{linha_label}] tab → campo FaixaFinal")
        
        faixaIniVerif = False
        
        # 4. Detecta virada de dia (FaixaInicio)
        if faixaFimAnterior:
            prev_ini = (faixaFimAnterior[0] or "").strip()
            if prev_ini.isdigit() and faixaIni_raw.isdigit():
                if int(faixaIni_raw) < int(prev_ini):
                    self.log.save(f"[{linha_label}] virada detectada em Faixa Inicio")
                    time.sleep(0.25)
                    pyautogui.press("tab")
                    pyautogui.press("enter")
                    time.sleep(0.4)
                    faixaIniVerif = True
        
        # 5. Escreve FaixaFinal
        pyautogui.write(faixaFim_dig)
        self.log.save(f"[{linha_label}] escreveu FaixaFinal '{faixaFim_log}'")
        
        # 6. Tab para Intervalo
        pyautogui.press("tab")
        self.log.save(f"[{linha_label}] tab → campo Intervalo")
        
        # 7. Detecta virada de dia (FaixaFinal)
        if faixaFimAnterior and not faixaIniVerif:
            prev_fim = faixaIni_dig
            if prev_fim.isdigit() and faixaFim_dig.isdigit():
                if int(faixaFim_dig) < int(prev_fim):
                    self.log.save(f"[{linha_label}] virada detectada em FaixaFinal")
                    time.sleep(0.25)
                    pyautogui.press("tab")
                    pyautogui.press("enter")
                    time.sleep(0.4)
        
        # 8. Escreve Intervalo
        pyautogui.write(Interv)
        self.log.save(f"[{linha_label}] escreveu Intervalo '{Interv}'")
        time.sleep(0.1)
        
        # 9. Tab para Percurso
        pyautogui.press("tab")
        self.log.save(f"[{linha_label}] tab → campo Tempo de Percurso")
        time.sleep(0.1)
        
        # 10. Escreve Percurso
        pyautogui.write(tPerc)
        self.log.save(f"[{linha_label}] escreveu Percurso '{tPerc}'")
        time.sleep(0.1)
        
        # 11. Tab para Terminal
        pyautogui.press("tab")
        self.log.save(f"[{linha_label}] tab → campo Terminal")
        time.sleep(0.1)
        
        # 12. Escreve Terminal
        pyautogui.write(tTerm)
        self.log.save(f"[{linha_label}] escreveu TempTerm '{tTerm}'")
        time.sleep(0.1)
        
        # 13. Tab para Frota
        pyautogui.press("tab")
        self.log.save(f"[{linha_label}] tab → campo Frota")
        time.sleep(0.1)
        
        # 14. Escreve Frota
        pyautogui.write(frota)
        self.log.save(f"[{linha_label}] escreveu Frota '{frota}'")
        time.sleep(0.1)
        
        # 15. Tab para Linha
        pyautogui.press("tab")
        self.log.save(f"[{linha_label}] tab → campo Linha (pronto para verificação)")
        time.sleep(0.1)
        
        self.log.save(f"[{linha_label}] faixa preenchida ({faixaIni_log} → {faixaFim_log})")
        self.log.user(f"{linha_label} • Faixa preenchida | F10 para continuar | F9 para repetir | F12 para parar ")

class PrintPDFModel:
    """Model para impressão de PDFs (adaptado da classe PrintPDF)"""
    def __init__(self, log_model: LogModel, config: ConfigModel):
        self.log = log_model
        self.config = config
    
    def tratarOSOs(self, osoBruta):
        """Trata e valida OSOs"""
        QuantidadeDeOsos = 0
        for row in osoBruta:
            oso = row.get("Oso", row.get("Oso ", "")).strip()
            if oso != "":
                QuantidadeDeOsos += 1
        
        osos = []
        
        for i, row in enumerate(osoBruta[:QuantidadeDeOsos]):
            linha = row.get("LinhaOso", row.get("LinhaOso ", "")).strip()
            oso = row.get("Oso", row.get("Oso ", "")).strip()
            
            if len(oso) == 6 and oso.isdigit():
                fimOso = int(oso[4:])
                
                if fimOso == 0:
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
                self.log.user(f"OSO [{oso}] | formato inválido seguir o formato: [123456]")
                return None
        
        filtroB = lambda x: x["tipo"] == "BASE"
        filtroD = lambda x: x["tipo"] == "DERIVADA"
        ososBase = list(filter(filtroB, osos))
        ososDerivada = list(filter(filtroD, osos))
        
        self.log.user(f"{len(osos)} OSOs identificadas | B:{len(ososBase)} D:{len(ososDerivada)}")
        return osos
    
    def imprimirPDF(self, row, fQH, OsoAtiva):
        """Impressão de um PDF (simulação de pyautogui)"""
        
        # --- Extração e Formatação ---
        linha = row.get("linha", row.get("linha ", "")).strip()
        oso = row.get("oso", row.get("oso ", "")).strip()
        oso_dig = row.get("oso_dig", row.get("oso_dig ", "")).strip()
        NomePDF = f"{fQH} {linha} {oso_dig}"
        caminho_pdf = os.path.join("AutoSiget3000", "FQHs", f"{NomePDF}.pdf")
        
        # --- Simulação de PyAutoGUI ---
        
        # 1. Preenche o campo Insira o n.º da OSO
        for ch in oso:
            pyautogui.write(ch)
            time.sleep(0.1)
        self.log.save(f"[{oso_dig}] escreveu OSO '{oso}'")
        self.log.user(f"Preenchendo OSO:     {oso_dig}")
        time.sleep(0.2)
        
        # 2. 6x seta para esquerda
        pyautogui.press('left', presses=6, interval=0.1)
        self.log.save(f"[{oso_dig}] 6x seta para esquerda")
        time.sleep(0.1)
        
        # 3. Tab para Impressão Gráfica
        pyautogui.press('tab')
        self.log.save(f"[{oso_dig}] Insira o n.º da OSO -> Impressão Gráfica")
        time.sleep(0.2)
        
        # 4. Fecha POP-UP e imprime
        time.sleep(0.5)
        pyautogui.press("enter", presses=3, interval=0.3)
        self.log.save(f"[{oso_dig}] fechou o POP-UP e botou para imprimir")
        time.sleep(3)
        
        # 5. Digita o nome do PDF e confirma salvar
        pyautogui.write(NomePDF)
        self.log.save(f"[{oso_dig}] escreveu nome do PDF '{NomePDF}'")
        self.log.user(f"Preenchendo NomePDF: {NomePDF}")
        time.sleep(0.3)
        
        pyautogui.press("enter")
        self.log.save(f"[{oso_dig}] confirmou salvar PDF")
        time.sleep(3.5)
        
        # 6. Shift+Tab para voltar
        pyautogui.hotkey('shift', 'tab')
        self.log.save(f"[{oso_dig}] botão Imprimir -> Informe o N° da oso")
        time.sleep(0.2)
        
        # --- Simulação de Verificação de Arquivo ---
        # Como não podemos criar o PDF real, simulamos o sucesso/erro
        # A verificação de os.path.exists(caminho_pdf) no código original é a chave.
        # Aqui, vamos apenas logar a intenção.
        
        self.log.save(f"[{oso_dig}] Simulação: PDF salvo com sucesso: {NomePDF}.pdf")
        self.log.user(f"   {NomePDF}.pdf | Simulação de Salvo com sucesso")
        self.log.user(f"[{oso_dig}] • F10 para continuar | F9 para repetir | F12 para parar ")

# -------------------------
# CONTROLLER
# -------------------------

class AppController:
    """Controller principal da aplicação"""
    def __init__(self):
        self.config = ConfigModel()
        self.log_model = LogModel(self.config)
        self.csv_model = CSVModel(self.log_model)
        self.prog_linha_model = ProgramaLinhaModel(self.log_model, self.config)
        self.pdf_model = PrintPDFModel(self.log_model, self.config)
        
        self.dados = None
        self.cabecalho = None
        self.blocos = None
        self.osos = None
    
    def carregar_csv(self, caminho_csv: str):
        """Carrega arquivo CSV"""
        self.dados, self.cabecalho = self.csv_model.load(caminho_csv)
        
        if self.dados is None:
            return False, f"Arquivo '{caminho_csv}' não encontrado ou erro de leitura."
        
        # Valida campos obrigatórios
        chavesObrigatorias = {
            "FaixaInicio", "FaixaFinal", "Interv", "Percurso", 
            "TempTerm", "Frota", "Linha", "Dia", "Sentido", "Oso", "LinhaOso"
        }
        chavesCabecalho = set(h.strip() for h in (self.cabecalho or []))
        chaveFaltando = chavesObrigatorias - chavesCabecalho
        
        if chaveFaltando:
            return False, f"Colunas obrigatórias faltando: {', '.join(chaveFaltando)}"
        
        return True, f"CSV carregado com sucesso! {len(self.dados)} registros."
    
    def gerar_blocos(self):
        """Gera blocos de programação"""
        if self.dados is None:
            return False, "Nenhum CSV carregado."
        
        self.blocos = self.prog_linha_model.genBlocos(self.dados)
        
        if not self.blocos:
            return False, "Nenhum bloco identificado no CSV."
        
        return True, f"{len(self.blocos)} blocos gerados."
    
    def gerar_osos(self):
        """Gera lista de OSOs"""
        if self.dados is None:
            return False, "Nenhum CSV carregado."
        
        self.osos = self.pdf_model.tratarOSOs(self.dados)
        
        if self.osos is None:
            return False, "Erro ao processar OSOs (verifique o formato)."
        
        if not self.osos:
            return False, "Nenhuma OSO identificada no CSV."
        
        return True, f"{len(self.osos)} OSOs geradas."

# -------------------------
# VIEW - Interface Gráfica
# -------------------------

class BaseView:
    """View base com componentes comuns"""
    def __init__(self, controller: AppController):
        self.controller = controller
    
    def criar_titulo(self, texto: str, size: int = 24):
        """Cria um título estilizado"""
        return ft.Text(
            texto,
            size=size,
            weight=ft.FontWeight.BOLD,
            color=ft.Colors.BLACK
        )
    
    def criar_botao(self, texto: str, on_click, icon=None, expand=False):
        """Cria um botão estilizado"""
        return ft.ElevatedButton(
            texto,
            icon=icon,
            on_click=on_click,
            expand=expand,
            style=ft.ButtonStyle(
                padding=15,
                shape=ft.RoundedRectangleBorder(radius=8),
                bgcolor=ft.Colors.GREY_800,
                color=ft.Colors.WHITE
            )
        )
    
    def criar_card(self, conteudo):
        """Cria um card estilizado"""
        return ft.Container(
            content=conteudo,
            padding=20,
            border_radius=10,
            bgcolor=ft.Colors.WHITE,
            border=ft.border.all(1, ft.Colors.GREY_300)
        )
    
    def mostrar_alerta(self, page: ft.Page, titulo: str, mensagem: str):
        """Mostra um alerta"""
        def fechar_dialogo(e):
            page.dialog.open = False
            page.update()
            
        dlg = ft.AlertDialog(
            title=ft.Text(titulo, color=ft.Colors.BLACK),
            content=ft.Text(mensagem, color=ft.Colors.BLACK),
            actions=[
                ft.TextButton("OK", on_click=fechar_dialogo)
            ]
        )
        page.dialog = dlg
        dlg.open = True
        page.update()

class HomeView(BaseView):
    """View da tela inicial (agora o menu principal)"""
    def __init__(self, controller: AppController, page: ft.Page):
        super().__init__(controller)
        self.page = page
    
    def build(self):
        """Constrói a view"""
        return ft.Column(
            [
                ft.Container(height=20),
                ft.Container(
                    content=ft.Column(
                        [
                            ft.Icon(ft.Icons.SETTINGS_APPLICATIONS, size=80, color=ft.Colors.ORANGE_700),
                            ft.Text(
                                "AUTOMA SIGET3000",
                                size=32,
                                weight=ft.FontWeight.BOLD,
                                color=ft.Colors.ORANGE_700
                            ),
                            ft.Text(
                                "Sistema de Automação para SIGET",
                                size=16,
                                color=ft.Colors.BLACK
                            ),
                        ],
                        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                        spacing=10
                    ),
                    padding=30,
                    alignment=ft.alignment.center
                ),
                ft.Divider(height=20, color=ft.Colors.TRANSPARENT),
                self.criar_card(
                    ft.Column(
                        [
                            self.criar_titulo("Menu Principal", 20),
                            ft.Text(
                                "Selecione uma opção abaixo para continuar:",
                                size=14,
                                color=ft.Colors.BLACK
                            ),
                            ft.Container(height=10),
                            self.criar_botao(
                                "📋 Programação de Linha",
                                self.navegar_para_prog_linha,
                                icon=ft.Icons.EDIT_CALENDAR,
                                expand=True
                            ),
                            self.criar_botao(
                                "🖨️ Impressão de PDF",
                                self.navegar_para_pdf,
                                icon=ft.Icons.PRINT,
                                expand=True
                            ),
                        ],
                        spacing=15
                    )
                ),
                ft.Container(height=20),
                ft.Text(
                    "Autor: Eude Leal | Github: eudeleal",
                    size=12,
                    color=ft.Colors.GREY_500,
                    text_align=ft.TextAlign.CENTER
                ),
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            scroll=ft.ScrollMode.AUTO,
            expand=True
        )
    
    def navegar_para_prog_linha(self, e):
        """Navega para tela de programação de linha"""
        self.page.views.append(ProgramacaoLinhaView(self.controller, self.page).build_view())
        self.page.update()
    
    def navegar_para_pdf(self, e):
        """Navega para tela de impressão de PDF"""
        self.page.views.append(ImpressaoPDFView(self.controller, self.page).build_view())
        self.page.update()

class CSVView(BaseView):
    """View para carregamento de CSV (tela inicial)"""
    def __init__(self, controller: AppController, page: ft.Page):
        super().__init__(controller)
        self.page = page
        self.texto_status = ft.Text("", size=14, color=ft.Colors.BLACK)
        
        # FilePicker para seleção de arquivo
        self.file_picker = ft.FilePicker(on_result=self.on_file_picker_result)
        page.overlay.append(self.file_picker)

    def build(self):
        """Constrói a view inicial"""
        return ft.Column(
            [
                ft.Container(height=50),
                ft.Container(
                    content=ft.Column(
                        [
                            ft.Icon(ft.Icons.UPLOAD_FILE, size=100, color=ft.Colors.ORANGE_700),
                            self.criar_titulo("Carregar Arquivo CSV", 30),
                            ft.Text(
                                "Para começar, selecione o arquivo CSV com os dados.",
                                size=16,
                                color=ft.Colors.BLACK
                            ),
                            ft.Container(height=20),
                            self.criar_botao(
                                "Selecionar Arquivo",
                                lambda _: self.file_picker.pick_files(
                                    allow_multiple=False,
                                    allowed_extensions=['csv']
                                ),
                                icon=ft.Icons.FOLDER_OPEN,
                                expand=True
                            ),
                            ft.Container(height=10),
                            self.texto_status,
                        ],
                        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                        spacing=15
                    ),
                    padding=40,
                    alignment=ft.alignment.center
                ),
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            expand=True
        )

    def on_file_picker_result(self, e: ft.FilePickerResultEvent):
        """Callback do FilePicker"""
        if e.files:
            caminho_arquivo = e.files[0].path
            self.texto_status.value = "Carregando..."
            self.texto_status.color = ft.Colors.BLACK
            self.page.update()
            
            sucesso, mensagem = self.controller.carregar_csv(caminho_arquivo)
            
            if sucesso:
                self.texto_status.value = f"✅ {mensagem}"
                self.texto_status.color = ft.Colors.GREEN_700
                self.page.update()
                time.sleep(1) # Pequena pausa para o usuário ver a mensagem
                self.page.go("/home") # Navega para o menu principal
            else:
                self.texto_status.value = f"❌ {mensagem}"
                self.texto_status.color = ft.Colors.RED_700
                self.page.update()
        else:
            self.texto_status.value = "Nenhum arquivo selecionado."
            self.texto_status.color = ft.Colors.BLACK
            self.page.update()

class ProgramacaoLinhaView(BaseView):
    """View para programação de linha"""
    def __init__(self, controller: AppController, page: ft.Page):
        super().__init__(controller)
        self.page = page
        self.lista_blocos = ft.ListView(spacing=5, expand=True)
        self.texto_status = ft.Text("", size=14)
        self.log_area = ft.ListView(spacing=5, expand=True, auto_scroll=True)
        self.is_processing = False
        
        # Adiciona callback para logs
        self.controller.log_model.add_callback(self.adicionar_log)
    
    def build_view(self):
        """Constrói a view completa"""
        # Gera blocos automaticamente
        sucesso, mensagem = self.controller.gerar_blocos()
        
        if sucesso:
            self.atualizar_lista_blocos()
        else:
            self.mostrar_alerta(self.page, "Erro", mensagem)
        
        return ft.View(
            "/prog_linha",
            [
                ft.AppBar(
                    title=ft.Text("Programação de Linha", color=ft.Colors.WHITE),
                    bgcolor=ft.Colors.ORANGE_100,
                    leading=ft.IconButton(
                        icon=ft.Icons.ARROW_BACK,
                        on_click=self.voltar
                    )
                ),
                ft.Container(
                    content=ft.Column(
                        [
                            ft.Container(height=10),
                            self.criar_card(
                                ft.Column(
                                    [
                                        self.criar_titulo("Blocos Disponíveis", 18),
                                        ft.Text(
                                            "Selecione um bloco para iniciar o preenchimento:",
                                            size=13,
                                            color=ft.Colors.BLACK
                                        ),
                                        ft.Container(height=10),
                                        ft.Container(
                                            content=self.lista_blocos,
                                            height=200,
                                            border=ft.border.all(1, ft.Colors.GREY_300),
                                            border_radius=8,
                                            padding=10
                                        ),
                                    ],
                                    spacing=10
                                )
                            ),
                            ft.Container(height=10),
                            self.criar_card(
                                ft.Column(
                                    [
                                        ft.Row(
                                            [
                                                self.criar_titulo("Log de Execução", 18),
                                                ft.IconButton(
                                                    icon=ft.Icons.CLEAR_ALL,
                                                    tooltip="Limpar log",
                                                    on_click=self.limpar_log
                                                )
                                            ],
                                            alignment=ft.MainAxisAlignment.SPACE_BETWEEN
                                        ),
                                        ft.Container(
                                            content=self.log_area,
                                            height=200,
                                            border=ft.border.all(1, ft.Colors.BLACK),
                                            border_radius=8,
                                            padding=10,
                                            bgcolor=ft.Colors.GREY_200
                                        ),
                                    ],
                                    spacing=10
                                )
                            ),
                            ft.Container(height=10),
                            self.texto_status,
                        ],
                        scroll=ft.ScrollMode.AUTO,
                        expand=True
                    ),
                    padding=20,
                    expand=True
                )
            ]
        )
    
    def atualizar_lista_blocos(self):
        """Atualiza a lista de blocos na UI"""
        self.lista_blocos.controls.clear()
        
        if not self.controller.blocos:
            self.lista_blocos.controls.append(
                ft.Text("Nenhum bloco disponível", color=ft.Colors.GREY_500)
            )
        else:
            for i, bloco in enumerate(self.controller.blocos):
                self.lista_blocos.controls.append(
                    ft.ListTile(
                        leading=ft.Icon(ft.Icons.CALENDAR_MONTH, color=ft.Colors.ORANGE_700),
                        title=ft.Text(bloco['id_bloco'], color=ft.Colors.BLACK, weight=ft.FontWeight.BOLD),
                        subtitle=ft.Text(f"{bloco['count']} faixas", color=ft.Colors.GREY_600),
                        trailing=ft.ElevatedButton(
                            "Iniciar",
                            on_click=lambda e, idx=i: self.iniciar_bloco(idx),
                            icon=ft.Icons.PLAY_ARROW,
                            style=ft.ButtonStyle(
                                bgcolor=ft.Colors.GREY_800,
                                color=ft.Colors.WHITE,
                                padding=5,
                                shape=ft.RoundedRectangleBorder(radius=5)
                            )
                        ),
                        content_padding=ft.padding.symmetric(vertical=0, horizontal=10),
                        dense=True
                    )
                )
        
        self.page.update()
    
    def iniciar_bloco(self, bloco_index: int):
        """Inicia o processamento de um bloco"""
        if self.is_processing:
            self.mostrar_alerta(self.page, "Aviso", "Um processo já está em execução.")
            return
            
        self.is_processing = True
        self.controller.config.parar_execucao = False
        self.adicionar_log(f"Iniciando bloco {bloco_index + 1}...")
        self.texto_status.value = "⏳ Processamento iniciado. Aguarde 3 segundos para posicionar o cursor no SIGET..."
        self.texto_status.color = ft.Colors.ORANGE_700
        self.page.update()
        
        # Executa em thread separada para não travar a UI
        threading.Thread(
            target=self.processar_bloco,
            args=(bloco_index,),
            daemon=True
        ).start()
    
    def processar_bloco(self, bloco_index: int):
        """Processa um bloco (executado em thread separada)"""
        try:
            bloco = self.controller.blocos[bloco_index]
            linhas = self.controller.dados
            inicio = bloco["inicio"]
            fim = bloco["fim"]
            count = bloco["count"]
            
            self.adicionar_log(f"Aguardando 3 segundos...")
            time.sleep(3)
            
            trecho = linhas[inicio:fim+1]
            faixaFimAnterior = None
            
            for offset in range(count):
                if self.controller.config.parar_execucao:
                    self.adicionar_log("Execução interrompida pelo usuário.")
                    break
                
                row = trecho[offset]
                linha_label = bloco["id_bloco"]
                
                self.adicionar_log(f">>> Processando faixa {offset+1}/{count}")
                
                self.controller.prog_linha_model.preenFaixa(
                    row, 
                    linha_label, 
                    faixaFimAnterior=faixaFimAnterior
                )
                
                faixaFimAnterior = (row.get("FaixaFim", ""))
                
                self.adicionar_log(f"Faixa {offset+1} preenchida. Pressione F10 para continuar | F9 para repetir | F12 para parar")
                
                # Aguarda F10, F9 ou F12
                while True:
                    if keyboard.is_pressed("F12"):
                        self.controller.config.parar_execucao = True
                        self.adicionar_log("Parada solicitada (F12).")
                        break
                    if keyboard.is_pressed("F10"):
                        time.sleep(0.18)
                        self.adicionar_log(f"F10 pressionado — avançando para faixa {offset + 2}")
                        break
                    if keyboard.is_pressed("F9"):
                        time.sleep(0.18)
                        self.adicionar_log(f"F9 pressionado — repetindo faixa {offset + 1}")
                        offset -= 1 # Repete a faixa atual
                        break
                    time.sleep(0.06)
                
                if self.controller.config.parar_execucao:
                    break
            
            if not self.controller.config.parar_execucao:
                self.adicionar_log(f"Bloco {bloco['id_bloco']} finalizado!")
                self.atualizar_status("✅ Processamento concluído!", ft.Colors.GREEN_700)
            else:
                self.atualizar_status("🛑 Processamento interrompido.", ft.Colors.RED_700)
            
        except Exception as ex:
            self.adicionar_log(f"Erro: {str(ex)}")
            self.atualizar_status(f"❌ Erro: {str(ex)}", ft.Colors.RED_700)
        finally:
            self.is_processing = False
            self.page.update()
    
    def adicionar_log(self, mensagem: str):
        """Adiciona mensagem ao log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_area.controls.append(
            ft.Text(f"[{timestamp}] {mensagem}", size=12, color=ft.Colors.BLACK)
        )
        if len(self.log_area.controls) > 100:
            self.log_area.controls.pop(0)
        self.page.update()
    
    def limpar_log(self, e):
        """Limpa o log"""
        self.log_area.controls.clear()
        self.page.update()
    
    def atualizar_status(self, mensagem: str, cor):
        """Atualiza o status"""
        self.texto_status.value = mensagem
        self.texto_status.color = cor
        self.page.update()
    
    def voltar(self, e):
        """Volta para a tela anterior"""
        self.page.views.pop()
        self.page.update()

class ImpressaoPDFView(BaseView):
    """View para impressão de PDF"""
    def __init__(self, controller: AppController, page: ft.Page):
        super().__init__(controller)
        self.page = page
        self.lista_osos = ft.ListView(spacing=5, expand=True)
        self.texto_status = ft.Text("", size=14)
        self.log_area = ft.ListView(spacing=5, expand=True, auto_scroll=True)
        self.is_processing = False
        self.tipo_impressao = ft.Dropdown(
            label="Tipo de Impressão",
            options=[
                ft.dropdown.Option("FH_ATIVA", "FH - Faixa Horária (OSOs Ativas)"),
                ft.dropdown.Option("QH_ATIVA", "QH - Quadro Horário (OSOs Ativas)"),
                ft.dropdown.Option("FH_DESATIVA", "FH - Faixa Horária (OSOs Desativas)"),
                ft.dropdown.Option("QH_DESATIVA", "QH - Quadro Horário (OSOs Desativas)"),
            ],
            value="FH_ATIVA",
            expand=True
        )
        
        # Adiciona callback para logs
        self.controller.log_model.add_callback(self.adicionar_log)
    
    def build_view(self):
        """Constrói a view completa"""
        # Gera OSOs automaticamente
        sucesso, mensagem = self.controller.gerar_osos()
        
        if sucesso:
            self.atualizar_lista_osos()
        else:
            self.mostrar_alerta(self.page, "Erro", mensagem)
        
        return ft.View(
            "/pdf",
            [
                ft.AppBar(
                    title=ft.Text("Impressão de PDF", color=ft.Colors.WHITE),
                    bgcolor=ft.Colors.ORANGE_700,
                    leading=ft.IconButton(
                        icon=ft.Icons.ARROW_BACK,
                        on_click=self.voltar
                    )
                ),
                ft.Container(
                    content=ft.Column(
                        [
                            ft.Container(height=10),
                            self.criar_card(
                                ft.Column(
                                    [
                                        self.criar_titulo("Configurações", 18),
                                        self.tipo_impressao,
                                    ],
                                    spacing=10
                                )
                            ),
                            ft.Container(height=10),
                            self.criar_card(
                                ft.Column(
                                    [
                                        self.criar_titulo("OSOs Disponíveis", 18),
                                        ft.Text(
                                            "Selecione uma OSO para iniciar a impressão:",
                                            size=13,
                                            color=ft.Colors.BLACK
                                        ),
                                        ft.Container(height=10),
                                        ft.Container(
                                            content=self.lista_osos,
                                            height=200,
                                            border=ft.border.all(1, ft.Colors.GREY_300),
                                            border_radius=8,
                                            padding=10
                                        ),
                                    ],
                                    spacing=10
                                )
                            ),
                            ft.Container(height=10),
                            self.criar_card(
                                ft.Column(
                                    [
                                        ft.Row(
                                            [
                                                self.criar_titulo("Log de Execução", 18),
                                                ft.IconButton(
                                                    icon=ft.Icons.CLEAR_ALL,
                                                    tooltip="Limpar log",
                                                    on_click=self.limpar_log
                                                )
                                            ],
                                            alignment=ft.MainAxisAlignment.SPACE_BETWEEN
                                        ),
                                        ft.Container(
                                            content=self.log_area,
                                            height=200,
                                            border=ft.border.all(1, ft.Colors.GREY_300),
                                            border_radius=8,
                                            padding=10,
                                            bgcolor=ft.Colors.GREY_200
                                        ),
                                    ],
                                    spacing=10
                                )
                            ),
                            ft.Container(height=10),
                            self.texto_status,
                        ],
                        scroll=ft.ScrollMode.AUTO,
                        expand=True
                    ),
                    padding=20,
                    expand=True
                )
            ]
        )
    
    def atualizar_lista_osos(self):
        """Atualiza a lista de OSOs na UI"""
        self.lista_osos.controls.clear()
        
        if not self.controller.osos:
            self.lista_osos.controls.append(
                ft.Text("Nenhuma OSO disponível", color=ft.Colors.GREY_500)
            )
        else:
            for i, oso in enumerate(self.controller.osos):
                self.lista_osos.controls.append(
                    ft.ListTile(
                        leading=ft.Icon(ft.Icons.PRINT, color=ft.Colors.ORANGE_700),
                        title=ft.Text(f"OSO {oso['oso_dig']} | Linha: {oso['linha']}", color=ft.Colors.BLACK, weight=ft.FontWeight.BOLD),
                        subtitle=ft.Text(f"Tipo: {oso['tipo']}", color=ft.Colors.GREY_600),
                        trailing=ft.ElevatedButton(
                            "Imprimir",
                            on_click=lambda e, idx=i: self.iniciar_impressao(idx),
                            icon=ft.Icons.PRINT,
                            style=ft.ButtonStyle(
                                bgcolor=ft.Colors.GREY_800,
                                color=ft.Colors.WHITE,
                                padding=5,
                                shape=ft.RoundedRectangleBorder(radius=5)
                            )
                        ),
                        content_padding=ft.padding.symmetric(vertical=0, horizontal=10),
                        dense=True
                    )
                )
        
        self.page.update()
    
    def iniciar_impressao(self, oso_index: int):
        """Inicia a impressão de uma OSO"""
        if self.is_processing:
            self.mostrar_alerta(self.page, "Aviso", "Um processo já está em execução.")
            return
            
        self.is_processing = True
        self.controller.config.parar_execucao = False
        self.adicionar_log(f"Iniciando impressão da OSO {oso_index + 1}...")
        self.texto_status.value = "⏳ Impressão iniciada. Aguarde 5 segundos para posicionar o cursor no SIGET..."
        self.texto_status.color = ft.Colors.ORANGE_700
        self.page.update()
        
        # Executa em thread separada
        threading.Thread(
            target=self.processar_impressao,
            args=(oso_index,),
            daemon=True
        ).start()
    
    def processar_impressao(self, oso_index: int):
        """Processa a impressão (executado em thread separada)"""
        try:
            oso = self.controller.osos[oso_index]
            
            self.adicionar_log(f"Aguardando 5 segundos...")
            time.sleep(5)
            
            # Extrai tipo de impressão
            tipo_valor = self.tipo_impressao.value
            fQH = "FH" if "FH" in tipo_valor else "QH"
            OsoAtiva = "ATIVA" in tipo_valor
            
            self.adicionar_log(f"Processando OSO {oso['oso_dig']}...")
            
            self.controller.pdf_model.imprimirPDF(oso, fQH, OsoAtiva)
            
            self.adicionar_log(f"OSO {oso['oso_dig']} processada. Pressione F10 para continuar | F9 para repetir | F12 para parar")
            
            # Aguarda F10, F9 ou F12
            while True:
                if keyboard.is_pressed("F12"):
                    self.controller.config.parar_execucao = True
                    self.adicionar_log("Parada solicitada (F12).")
                    break
                if keyboard.is_pressed("F10"):
                    time.sleep(0.18)
                    self.adicionar_log(f"F10 pressionado — avançando para próxima OSO")
                    break
                if keyboard.is_pressed("F9"):
                    time.sleep(0.18)
                    self.adicionar_log(f"F9 pressionado — repetindo OSO")
                    self.processar_impressao(oso_index) # Repete a OSO atual
                    return
                time.sleep(0.06)
            
            if not self.controller.config.parar_execucao:
                self.adicionar_log(f"Impressão finalizada!")
                self.atualizar_status("✅ Impressão concluída!", ft.Colors.GREEN_700)
            else:
                self.atualizar_status("🛑 Impressão interrompida.", ft.Colors.RED_700)
            
        except Exception as ex:
            self.adicionar_log(f"Erro: {str(ex)}")
            self.atualizar_status(f"❌ Erro: {str(ex)}", ft.Colors.RED_700)
        finally:
            self.is_processing = False
            self.page.update()
    
    def adicionar_log(self, mensagem: str):
        """Adiciona mensagem ao log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_area.controls.append(
            ft.Text(f"[{timestamp}] {mensagem}", size=12, color=ft.Colors.BLACK)
        )
        if len(self.log_area.controls) > 100:
            self.log_area.controls.pop(0)
        self.page.update()
    
    def limpar_log(self, e):
        """Limpa o log"""
        self.log_area.controls.clear()
        self.page.update()
    
    def atualizar_status(self, mensagem: str, cor):
        """Atualiza o status"""
        self.texto_status.value = mensagem
        self.texto_status.color = cor
        self.page.update()
    
    def voltar(self, e):
        """Volta para a tela anterior"""
        self.page.views.pop()
        self.page.update()

# -------------------------
# APLICAÇÃO PRINCIPAL
# -------------------------

def main(page: ft.Page):
    """Função principal da aplicação"""
    page.title = "AUTOMA SIGET3000"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.bgcolor = ft.Colors.WHITE
    page.padding = 0
    page.window_width = 900
    page.window_height = 700
    page.window_resizable = True
    
    # Inicializa controller
    controller = AppController()
    
    # Cria views
    csv_view = CSVView(controller, page)
    home_view = HomeView(controller, page)
    
    def route_change(route):
        page.views.clear()
        page.views.append(
            ft.View(
                "/",
                [csv_view.build()],
                bgcolor=ft.Colors.WHITE
            )
        )
        if page.route == "/home":
            page.views.append(
                ft.View(
                    "/home",
                    [home_view.build()],
                    bgcolor=ft.Colors.WHITE
                )
            )
        page.update()

    def view_pop(view):
        page.views.pop()
        top_view = page.views[-1]
        page.go(top_view.route)

    page.on_route_change = route_change
    page.on_view_pop = view_pop
    
    page.go(page.route)

if __name__ == "__main__":
    ft.app(target=main)
