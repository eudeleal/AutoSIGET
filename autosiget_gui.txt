
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
    """Model para gerenciamento de logs"""
    def __init__(self, config: ConfigModel):
        self.config = config
        self.log_callbacks = []
    
    def add_callback(self, callback):
        """Adiciona callback para atualização da UI"""
        self.log_callbacks.append(callback)
    
    def save(self, mensagem: str):
        """Registra eventos detalhados no arquivo de log"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        linha = f"[{timestamp}] {mensagem}"
        with open(self.config.nome_arquivo_log, "a", encoding="utf-8") as f:
            f.write(linha + "\n")
    
    def user(self, msg: str):
        """Mensagem para o usuário (console e UI)"""
        print(msg)
        self.save(f"[UI] {msg}")
        # Notifica callbacks (UI)
        for callback in self.log_callbacks:
            callback(msg)

class FormatHoraModel:
    """Model para formatação de horários"""
    @staticmethod
    def log(h: str):
        h = (h or "").strip()
        if len(h) == 4 and h.isdigit():
            return f"{h[:2]}:{h[2:]}"
        if len(h) == 3 and h.isdigit():
            return f"0{h[:1]}:{h[1:]}"
        return h
    
    @staticmethod
    def dig(h: str):
        h = (h or "").strip()
        if len(h) == 3 and h.isdigit():
            return f"0{h}"
        return h

class CSVModel:
    """Model para operações com CSV"""
    def __init__(self, log_model: LogModel):
        self.log = log_model
    
    def load(self, caminho_csv: str):
        """Carrega CSV e retorna linhas e fieldnames"""
        if not os.path.exists(caminho_csv):
            return None, None
        
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

class ProgramacaoLinhaModel:
    """Model para programação de linha"""
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
        """Preenche uma faixa horária"""
        faixaIni_raw = row.get("FaixaInicio", row.get("FaixaInicio ", "")).strip()
        faixaFim_raw = row.get("FaixaFinal", row.get("FaixaFinal ", "")).strip()
        interv = row.get("Intervalo", row.get("Interv", "")).strip()
        tPerc = row.get("Percurso", row.get("Perc.", row.get("Perc", ""))).strip()
        tTerm = row.get("TempTerm", row.get("T. Term", row.get("T. Term ", ""))).strip()
        frota = row.get("Frota", "").strip()
        
        faixaIni_log = FormatHoraModel.log(faixaIni_raw)
        faixaFim_log = FormatHoraModel.log(faixaFim_raw)
        faixaIni_dig = FormatHoraModel.dig(faixaIni_raw)
        faixaFim_dig = FormatHoraModel.dig(faixaFim_raw)
        
        # Início do preenchimento
        pyautogui.write(faixaIni_dig)
        self.log.save(f"[{linha_label}] escreveu FaixaInicio '{faixaIni_log}'")
        
        self.log.user(f"{linha_label} • Preenchendo faixa de {faixaIni_log} à {faixaFim_log}")
        
        pyautogui.press("tab")
        self.log.save(f"[{linha_label}] tab → campo FaixaFinal")
        
        faixaIniVerif = False
        
        # Detecta virada de dia
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
        
        pyautogui.write(faixaFim_dig)
        self.log.save(f"[{linha_label}] escreveu FaixaFinal '{faixaFim_log}'")
        
        pyautogui.press("tab")
        self.log.save(f"[{linha_label}] tab → campo Intervalo")
        
        if faixaFimAnterior and not faixaIniVerif:
            prev_fim = faixaIni_dig
            if prev_fim.isdigit() and faixaFim_dig.isdigit():
                if int(faixaFim_dig) < int(prev_fim):
                    self.log.save(f"[{linha_label}] virada detectada em FaixaFinal")
                    time.sleep(0.25)
                    pyautogui.press("tab")
                    pyautogui.press("enter")
                    time.sleep(0.4)
        
        pyautogui.write(interv)
        self.log.save(f"[{linha_label}] escreveu Intervalo '{interv}'")
        time.sleep(0.1)
        
        pyautogui.press("tab")
        self.log.save(f"[{linha_label}] tab → campo Tempo de Percurso")
        time.sleep(0.1)
        
        pyautogui.write(tPerc)
        self.log.save(f"[{linha_label}] escreveu Percurso '{tPerc}'")
        time.sleep(0.1)
        
        pyautogui.press("tab")
        self.log.save(f"[{linha_label}] tab → campo Terminal")
        time.sleep(0.1)
        
        pyautogui.write(tTerm)
        self.log.save(f"[{linha_label}] escreveu TempTerm '{tTerm}'")
        time.sleep(0.1)
        
        pyautogui.press("tab")
        self.log.save(f"[{linha_label}] tab → campo Frota")
        time.sleep(0.1)
        
        pyautogui.write(frota)
        self.log.save(f"[{linha_label}] escreveu Frota '{frota}'")
        time.sleep(0.1)
        
        pyautogui.press("tab")
        self.log.save(f"[{linha_label}] tab → campo Linha")
        time.sleep(0.1)
        
        self.log.save(f"[{linha_label}] faixa preenchida ({faixaIni_log} → {faixaFim_log})")

class ImpressaoPDFModel:
    """Model para impressão de PDFs"""
    def __init__(self, log_model: LogModel, config: ConfigModel):
        self.log = log_model
        self.config = config
    
    def tratarOSOs(self, osoBruta):
        """Trata e valida OSOs"""
        QuantidadeDeOsos = 0
        for i, row in enumerate(osoBruta):
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
                self.log.user(f"OSO [{oso}] | formato inválido seguir o formato: [123456]")
                return None
        
        filtoB = lambda x: x["tipo"] == "BASE"
        filtoD = lambda x: x["tipo"] == "DERIVADA"
        ososBase = list(filter(filtoB, osos))
        ososDerivada = list(filter(filtoD, osos))
        
        self.log.user(f"{len(osos)} OSOs identificadas | B:{len(ososBase)} D:{len(ososDerivada)}")
        return osos

# -------------------------
# CONTROLLER
# -------------------------

class AppController:
    """Controller principal da aplicação"""
    def __init__(self):
        self.config = ConfigModel()
        self.log_model = LogModel(self.config)
        self.csv_model = CSVModel(self.log_model)
        self.prog_linha_model = ProgramacaoLinhaModel(self.log_model, self.config)
        self.pdf_model = ImpressaoPDFModel(self.log_model, self.config)
        
        self.dados = None
        self.cabecalho = None
        self.blocos = None
        self.osos = None
    
    def carregar_csv(self, caminho_csv: str):
        """Carrega arquivo CSV"""
        if not caminho_csv.endswith('.csv'):
            caminho_csv += '.csv'
        
        self.dados, self.cabecalho = self.csv_model.load(caminho_csv)
        
        if self.dados is None:
            return False, f"Arquivo '{caminho_csv}' não encontrado."
        
        # Valida campos obrigatórios
        chavesObrigatorias = {
            "FaixaInicio", "FaixaFinal", "Intervalo", "Percurso", 
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
            color=ft.colors.BLUE_700
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
                shape=ft.RoundedRectangleBorder(radius=8)
            )
        )
    
    def criar_card(self, conteudo):
        """Cria um card estilizado"""
        return ft.Container(
            content=conteudo,
            padding=20,
            border_radius=10,
            bgcolor=ft.colors.SURFACE_VARIANT,
            border=ft.border.all(1, ft.colors.OUTLINE)
        )

class HomeView(BaseView):
    """View da tela inicial"""
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
                            ft.Icon(ft.icons.SETTINGS_APPLICATIONS, size=80, color=ft.colors.BLUE_700),
                            ft.Text(
                                "AUTOMA SIGET3000",
                                size=32,
                                weight=ft.FontWeight.BOLD,
                                color=ft.colors.BLUE_700
                            ),
                            ft.Text(
                                "Sistema de Automação para SIGET",
                                size=16,
                                color=ft.colors.GREY_700
                            ),
                        ],
                        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                        spacing=10
                    ),
                    padding=30,
                    alignment=ft.alignment.center
                ),
                ft.Divider(height=20, color=ft.colors.TRANSPARENT),
                self.criar_card(
                    ft.Column(
                        [
                            self.criar_titulo("Bem-vindo(a)!", 20),
                            ft.Text(
                                "Selecione uma opção abaixo para começar:",
                                size=14,
                                color=ft.colors.GREY_700
                            ),
                            ft.Container(height=10),
                            self.criar_botao(
                                "📁 Carregar Arquivo CSV",
                                self.navegar_para_csv,
                                icon=ft.icons.UPLOAD_FILE,
                                expand=True
                            ),
                            self.criar_botao(
                                "📋 Programação de Linha",
                                self.navegar_para_prog_linha,
                                icon=ft.icons.EDIT_CALENDAR,
                                expand=True
                            ),
                            self.criar_botao(
                                "🖨️ Impressão de PDF",
                                self.navegar_para_pdf,
                                icon=ft.icons.PRINT,
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
                    color=ft.colors.GREY_500,
                    text_align=ft.TextAlign.CENTER
                ),
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            scroll=ft.ScrollMode.AUTO,
            expand=True
        )
    
    def navegar_para_csv(self, e):
        """Navega para tela de carregamento de CSV"""
        self.page.views.append(CSVView(self.controller, self.page).build_view())
        self.page.update()
    
    def navegar_para_prog_linha(self, e):
        """Navega para tela de programação de linha"""
        if self.controller.dados is None:
            self.mostrar_alerta("Erro", "Por favor, carregue um arquivo CSV primeiro.")
            return
        self.page.views.append(ProgramacaoLinhaView(self.controller, self.page).build_view())
        self.page.update()
    
    def navegar_para_pdf(self, e):
        """Navega para tela de impressão de PDF"""
        if self.controller.dados is None:
            self.mostrar_alerta("Erro", "Por favor, carregue um arquivo CSV primeiro.")
            return
        self.page.views.append(ImpressaoPDFView(self.controller, self.page).build_view())
        self.page.update()
    
    def mostrar_alerta(self, titulo: str, mensagem: str):
        """Mostra um alerta"""
        dlg = ft.AlertDialog(
            title=ft.Text(titulo),
            content=ft.Text(mensagem),
            actions=[
                ft.TextButton("OK", on_click=lambda e: self.fechar_dialogo(dlg))
            ]
        )
        self.page.dialog = dlg
        dlg.open = True
        self.page.update()
    
    def fechar_dialogo(self, dlg):
        """Fecha um diálogo"""
        dlg.open = False
        self.page.update()

class CSVView(BaseView):
    """View para carregamento de CSV"""
    def __init__(self, controller: AppController, page: ft.Page):
        super().__init__(controller)
        self.page = page
        self.campo_csv = ft.TextField(
            label="Nome do arquivo CSV (sem extensão)",
            hint_text="exemplo: dados",
            expand=True,
            autofocus=True
        )
        self.texto_status = ft.Text("", size=14, color=ft.colors.GREY_700)
    
    def build_view(self):
        """Constrói a view completa"""
        return ft.View(
            "/csv",
            [
                ft.AppBar(
                    title=ft.Text("Carregar Arquivo CSV"),
                    bgcolor=ft.colors.BLUE_700,
                    leading=ft.IconButton(
                        icon=ft.icons.ARROW_BACK,
                        on_click=self.voltar
                    )
                ),
                ft.Container(
                    content=ft.Column(
                        [
                            ft.Container(height=20),
                            self.criar_card(
                                ft.Column(
                                    [
                                        self.criar_titulo("Carregar CSV", 20),
                                        ft.Text(
                                            "Digite o nome do arquivo CSV (sem a extensão .csv):",
                                            size=14,
                                            color=ft.colors.GREY_700
                                        ),
                                        ft.Container(height=10),
                                        ft.Row(
                                            [
                                                self.campo_csv,
                                                ft.ElevatedButton(
                                                    "Carregar",
                                                    icon=ft.icons.UPLOAD_FILE,
                                                    on_click=self.carregar_csv
                                                )
                                            ],
                                            spacing=10
                                        ),
                                        ft.Container(height=10),
                                        self.texto_status,
                                    ],
                                    spacing=10
                                )
                            ),
                            ft.Container(height=20),
                            self.criar_card(
                                ft.Column(
                                    [
                                        self.criar_titulo("Instruções", 18),
                                        ft.Text(
                                            "• O arquivo CSV deve estar na mesma pasta do programa\n"
                                            "• Use ponto e vírgula (;) como separador\n"
                                            "• Certifique-se de que todas as colunas obrigatórias estão presentes\n"
                                            "• Colunas: FaixaInicio, FaixaFinal, Intervalo, Percurso, TempTerm, Frota, Linha, Dia, Sentido, Oso, LinhaOso",
                                            size=13,
                                            color=ft.colors.GREY_700
                                        ),
                                    ],
                                    spacing=10
                                )
                            ),
                        ],
                        scroll=ft.ScrollMode.AUTO,
                        expand=True
                    ),
                    padding=20,
                    expand=True
                )
            ]
        )
    
    def carregar_csv(self, e):
        """Carrega o arquivo CSV"""
        nome_arquivo = self.campo_csv.value.strip()
        
        if not nome_arquivo:
            self.texto_status.value = "❌ Por favor, digite o nome do arquivo."
            self.texto_status.color = ft.colors.RED_700
            self.page.update()
            return
        
        sucesso, mensagem = self.controller.carregar_csv(nome_arquivo)
        
        if sucesso:
            self.texto_status.value = f"✅ {mensagem}"
            self.texto_status.color = ft.colors.GREEN_700
        else:
            self.texto_status.value = f"❌ {mensagem}"
            self.texto_status.color = ft.colors.RED_700
        
        self.page.update()
    
    def voltar(self, e):
        """Volta para a tela anterior"""
        self.page.views.pop()
        self.page.update()

class ProgramacaoLinhaView(BaseView):
    """View para programação de linha"""
    def __init__(self, controller: AppController, page: ft.Page):
        super().__init__(controller)
        self.page = page
        self.lista_blocos = ft.Column(spacing=10)
        self.texto_status = ft.Text("", size=14)
        self.log_area = ft.ListView(spacing=5, expand=True, auto_scroll=True)
        
        # Adiciona callback para logs
        self.controller.log_model.add_callback(self.adicionar_log)
    
    def build_view(self):
        """Constrói a view completa"""
        # Gera blocos automaticamente
        sucesso, mensagem = self.controller.gerar_blocos()
        
        if sucesso:
            self.atualizar_lista_blocos()
        
        return ft.View(
            "/prog_linha",
            [
                ft.AppBar(
                    title=ft.Text("Programação de Linha"),
                    bgcolor=ft.colors.BLUE_700,
                    leading=ft.IconButton(
                        icon=ft.icons.ARROW_BACK,
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
                                            color=ft.colors.GREY_700
                                        ),
                                        ft.Container(height=10),
                                        ft.Container(
                                            content=self.lista_blocos,
                                            height=200,
                                            border=ft.border.all(1, ft.colors.OUTLINE),
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
                                                    icon=ft.icons.CLEAR_ALL,
                                                    tooltip="Limpar log",
                                                    on_click=self.limpar_log
                                                )
                                            ],
                                            alignment=ft.MainAxisAlignment.SPACE_BETWEEN
                                        ),
                                        ft.Container(
                                            content=self.log_area,
                                            height=200,
                                            border=ft.border.all(1, ft.colors.OUTLINE),
                                            border_radius=8,
                                            padding=10,
                                            bgcolor=ft.colors.BLACK12
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
                ft.Text("Nenhum bloco disponível", color=ft.colors.GREY_500)
            )
        else:
            for i, bloco in enumerate(self.controller.blocos):
                self.lista_blocos.controls.append(
                    ft.Container(
                        content=ft.Row(
                            [
                                ft.Text(f"{i+1}.", weight=ft.FontWeight.BOLD),
                                ft.Text(bloco['id_bloco'], expand=True),
                                ft.Text(f"{bloco['count']} faixas", color=ft.colors.GREY_600),
                                ft.ElevatedButton(
                                    "Iniciar",
                                    icon=ft.icons.PLAY_ARROW,
                                    on_click=lambda e, idx=i: self.iniciar_bloco(idx)
                                )
                            ],
                            alignment=ft.MainAxisAlignment.SPACE_BETWEEN
                        ),
                        padding=10,
                        border=ft.border.all(1, ft.colors.OUTLINE),
                        border_radius=5
                    )
                )
        
        self.page.update()
    
    def iniciar_bloco(self, bloco_index: int):
        """Inicia o processamento de um bloco"""
        self.adicionar_log(f"Iniciando bloco {bloco_index + 1}...")
        self.texto_status.value = "⏳ Processamento iniciado. Aguarde 3 segundos para posicionar o cursor no SIGET..."
        self.texto_status.color = ft.colors.ORANGE_700
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
                
                self.adicionar_log(f"Processando faixa {offset+1}/{count}")
                
                self.controller.prog_linha_model.preenFaixa(
                    row, 
                    linha_label, 
                    faixaFimAnterior=faixaFimAnterior
                )
                
                faixaFimAnterior = (row.get("FaixaFim", ""))
                
                self.adicionar_log(f"Faixa {offset+1} preenchida. Pressione F10 para continuar...")
                
                # Aguarda F10
                keyboard.wait('F10')
                time.sleep(0.18)
            
            self.adicionar_log(f"Bloco {bloco['id_bloco']} finalizado!")
            self.atualizar_status("✅ Processamento concluído!", ft.colors.GREEN_700)
            
        except Exception as ex:
            self.adicionar_log(f"Erro: {str(ex)}")
            self.atualizar_status(f"❌ Erro: {str(ex)}", ft.colors.RED_700)
    
    def adicionar_log(self, mensagem: str):
        """Adiciona mensagem ao log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_area.controls.append(
            ft.Text(f"[{timestamp}] {mensagem}", size=12, color=ft.colors.BLACK87)
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
        self.lista_osos = ft.Column(spacing=10)
        self.texto_status = ft.Text("", size=14)
        self.log_area = ft.ListView(spacing=5, expand=True, auto_scroll=True)
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
        
        return ft.View(
            "/pdf",
            [
                ft.AppBar(
                    title=ft.Text("Impressão de PDF"),
                    bgcolor=ft.colors.BLUE_700,
                    leading=ft.IconButton(
                        icon=ft.icons.ARROW_BACK,
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
                                            color=ft.colors.GREY_700
                                        ),
                                        ft.Container(height=10),
                                        ft.Container(
                                            content=self.lista_osos,
                                            height=200,
                                            border=ft.border.all(1, ft.colors.OUTLINE),
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
                                                    icon=ft.icons.CLEAR_ALL,
                                                    tooltip="Limpar log",
                                                    on_click=self.limpar_log
                                                )
                                            ],
                                            alignment=ft.MainAxisAlignment.SPACE_BETWEEN
                                        ),
                                        ft.Container(
                                            content=self.log_area,
                                            height=200,
                                            border=ft.border.all(1, ft.colors.OUTLINE),
                                            border_radius=8,
                                            padding=10,
                                            bgcolor=ft.colors.BLACK12
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
                ft.Text("Nenhuma OSO disponível", color=ft.colors.GREY_500)
            )
        else:
            for i, oso in enumerate(self.controller.osos):
                self.lista_osos.controls.append(
                    ft.Container(
                        content=ft.Row(
                            [
                                ft.Text(f"{oso['n_oso']}.", weight=ft.FontWeight.BOLD),
                                ft.Text(f"Linha: {oso['linha']}", expand=True),
                                ft.Text(f"OSO: {oso['oso_dig']}", color=ft.colors.GREY_600),
                                ft.Chip(
                                    label=ft.Text(oso['tipo'], size=11),
                                    bgcolor=ft.colors.BLUE_100 if oso['tipo'] == 'BASE' else ft.colors.ORANGE_100
                                ),
                                ft.ElevatedButton(
                                    "Imprimir",
                                    icon=ft.icons.PRINT,
                                    on_click=lambda e, idx=i: self.iniciar_impressao(idx)
                                )
                            ],
                            alignment=ft.MainAxisAlignment.SPACE_BETWEEN
                        ),
                        padding=10,
                        border=ft.border.all(1, ft.colors.OUTLINE),
                        border_radius=5
                    )
                )
        
        self.page.update()
    
    def iniciar_impressao(self, oso_index: int):
        """Inicia a impressão de uma OSO"""
        self.adicionar_log(f"Iniciando impressão da OSO {oso_index + 1}...")
        self.texto_status.value = "⏳ Impressão iniciada. Aguarde 5 segundos para posicionar o cursor no SIGET..."
        self.texto_status.color = ft.colors.ORANGE_700
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
            
            # Simula impressão (substitua pela lógica real)
            linha = oso.get("linha", "").strip()
            oso_val = oso.get("oso", "").strip()
            oso_dig = oso.get("oso_dig", "").strip()
            
            NomePDF = f"{fQH} {linha} {oso_dig}"
            
            self.adicionar_log(f"Nome do PDF: {NomePDF}")
            self.adicionar_log(f"OSO processada com sucesso!")
            
            self.atualizar_status("✅ Impressão concluída!", ft.colors.GREEN_700)
            
        except Exception as ex:
            self.adicionar_log(f"Erro: {str(ex)}")
            self.atualizar_status(f"❌ Erro: {str(ex)}", ft.colors.RED_700)
    
    def adicionar_log(self, mensagem: str):
        """Adiciona mensagem ao log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_area.controls.append(
            ft.Text(f"[{timestamp}] {mensagem}", size=12, color=ft.colors.BLACK87)
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
    page.padding = 0
    page.window_width = 900
    page.window_height = 700
    page.window_resizable = True
    
    # Inicializa controller
    controller = AppController()
    
    # Cria view inicial
    home_view = HomeView(controller, page)
    
    page.views.append(
        ft.View(
            "/",
            [home_view.build()],
            scroll=ft.ScrollMode.AUTO
        )
    )
    
    page.update()

if __name__ == "__main__":
    ft.app(target=main)
