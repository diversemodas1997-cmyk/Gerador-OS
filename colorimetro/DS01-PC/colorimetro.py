"""
Colorímetro DS-01 no PC — mede, compara padrão x amostra e exporta para Excel.

Uso:
    python colorimetro.py            (Bluetooth real)
    python colorimetro.py --simular  (aparelho simulado, para testar sem o DS-01)
"""
import asyncio
import csv
import datetime as dt
from decimal import Decimal, ROUND_HALF_UP
import json
import math
import os
import queue
import sys
import threading
import traceback
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from protocolo import (DS01, TransporteBLE, TransporteSimulado, TransporteCelular, Colorimetria, MODOS,
                       ErroAparelho, lch, delta_e_ab, delta_e_2000, delta_e_94, delta_e_cmc,
                       lab_para_hex)

AQUI = os.path.dirname(os.path.abspath(__file__))
ARQ_CONFIG = os.path.join(AQUI, "config.json")
PASTA_MEDICOES = os.path.join(AQUI, "medicoes")
SIMULAR = "--simular" in sys.argv

# Valores iguais às configurações do app no celular (capturas de 27/08/2026)
CONFIG_PADRAO = {
    "ultimo_aparelho": None,
    "iluminante1": "D65/10",
    "iluminante2": "A/2",
    "modo": 0,                                   # SCI
    "formula": "dE*ab",
    "media": 1,
    "cmc_l": 2.0, "cmc_c": 1.0,
    "tol": {"dL": 0.4, "da": 0.1, "db": 0.5, "dE*ab": 0.6, "dC": 0.6, "dH": 0.6,
            "dL'": 0.6, "dC'": 0.6, "dH'": 0.6, "dE00": 0.6, "dEcmc": 0.6, "dE94": 0.6},
    "nome_padrao": "Padrão", "nome_amostra": "Amostra",
    "k94": [1.0, 1.0, 1.0], "k00": [1.0, 1.0, 1.0],     # KL, KC, KH
}
FORMULAS = ["dE*ab", "dE00", "dEcmc", "dE94"]

# Cores e fontes do app do celular (tema escuro)
FUNDO = "#0b0f19"
CARTAO = "#212736"
CARTAO2 = "#2b3245"
CAIXA = "#353b54"                            # caixas da tela Configurações
TEXTO = "#ffffff"
FRACO = "#9aa3b5"
CIANO = "#5ccfe6"
CIANO_A, CIANO_B = "#5bbfea", "#6ee6df"      # degradê dos botões "Calibração"
VERDE = "#3ddc84"
VERMELHO = "#ff6b6b"
LARANJA = "#ffa94d"
CINZA_VAZIO = "#8c8c8c"
FONTE = "Segoe UI"


def _rgb(h):
    return tuple(int(h[i:i + 2], 16) for i in (1, 3, 5))


def misturar(c1, c2, t):
    """Mistura duas cores hex: t=0 -> c1, t=1 -> c2."""
    a, b = _rgb(c1), _rgb(c2)
    return "#%02x%02x%02x" % tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))


class Cartao(tk.Canvas):
    """Cartão de cantos arredondados; o conteúdo vai em .corpo."""

    def __init__(self, pai, pad=16, raio=12, ajustar_altura=False):
        super().__init__(pai, bg=FUNDO, highlightthickness=0, bd=0)
        self.pad, self.raio, self.ajustar = pad, raio, ajustar_altura
        self.corpo = tk.Frame(self, bg=CARTAO)
        self._janela = self.create_window(pad, pad, window=self.corpo, anchor="nw")
        self.bind("<Configure>", self._redesenhar)
        if ajustar_altura:
            self.corpo.bind("<Configure>", lambda e: self.configure(height=self.corpo.winfo_reqheight() + 2 * pad))

    def _redesenhar(self, e):
        w, h, r = e.width, e.height, self.raio
        self.delete("fundo")
        p = [r, 0, r, 0, w - r, 0, w - r, 0, w, 0, w, r, w, r, w, h - r, w, h - r, w, h,
             w - r, h, w - r, h, r, h, r, h, 0, h, 0, h - r, 0, h - r, 0, r, 0, r, 0, 0]
        self.create_polygon(p, smooth=True, fill=CARTAO, outline="", tags="fundo")
        self.tag_lower("fundo")
        self.itemconfigure(self._janela, width=max(1, w - 2 * self.pad))
        if not self.ajustar:
            self.itemconfigure(self._janela, height=max(1, h - 2 * self.pad))


class Pilula(tk.Canvas):
    """Botão arredondado do app: 'ciano' (degradê), 'branco' (texto ciano) ou 'contorno'."""

    def __init__(self, pai, texto, comando, estilo="ciano", altura=44, fonte=12, largura=None, fundo=CARTAO):
        super().__init__(pai, height=altura, bg=fundo, highlightthickness=0, bd=0, cursor="hand2")
        if largura:
            self.configure(width=largura)
        self.texto, self.comando, self.estilo, self.tam, self.fundo = texto, comando, estilo, fonte, fundo
        self.ativo = True
        self.bind("<Configure>", lambda e: self._desenhar())
        self.bind("<Button-1>", lambda e: self.ativo and self.comando and self.comando())

    def __setitem__(self, chave, valor):
        if chave == "state":
            ativo = valor != "disabled"
            if ativo != self.ativo:
                self.ativo = ativo
                self.configure(cursor="hand2" if ativo else "arrow")
                self._desenhar()
        elif chave == "text":
            if valor != self.texto:
                self.texto = valor
                self._desenhar()
        else:
            super().__setitem__(chave, valor)

    def _desenhar(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 10:
            return
        r = h / 2
        apagado = lambda c: c if self.ativo else misturar(c, self.fundo, 0.55)  # noqa: E731

        def meia(x):
            d = min(x, w - x)
            return r if d >= r else math.sqrt(max(0.0, r * r - (r - d) ** 2))

        if self.estilo == "contorno":
            c = apagado(CIANO)
            self.create_arc(1, 1, h - 1, h - 1, start=90, extent=180, style="arc", outline=c, width=1.5)
            self.create_arc(w - h + 1, 1, w - 1, h - 1, start=270, extent=180, style="arc", outline=c, width=1.5)
            self.create_line(r, 1, w - r, 1, fill=c, width=1.5)
            self.create_line(r, h - 1, w - r, h - 1, fill=c, width=1.5)
            cor_texto = c
        else:
            for x in range(w):
                if self.estilo == "branco":
                    cor = apagado("#ffffff")
                else:
                    cor = apagado(misturar(CIANO_A, CIANO_B, x / max(1, w - 1)))
                m = meia(x + 0.5)
                self.create_line(x, r - m, x, r + m + 1, fill=cor)
            cor_texto = apagado(CIANO) if self.estilo == "branco" else apagado("#ffffff")
        peso = "normal" if self.estilo == "contorno" else "bold"
        self.create_text(w / 2, h / 2, text=self.texto, fill=cor_texto, font=(FONTE, self.tam, peso))



def carregar_config():
    cfg = json.loads(json.dumps(CONFIG_PADRAO))
    try:
        with open(ARQ_CONFIG, encoding="utf-8") as f:
            salvo = json.load(f)
        tol = {**cfg["tol"], **salvo.get("tol", {})}
        cfg.update(salvo)
        cfg["tol"] = tol
    except (OSError, ValueError):
        pass
    return cfg


CASAS_TELA = 1      # o app mostra 1 casa decimal para o DS-01 (precision = 1)
VERDE_APP, VERMELHO_APP = "#40b12b", "#b6101f"   # barra Pass/Fail do app


def arred(v, casas=CASAS_TELA):
    """Arredonda como o app (toFixed do JavaScript: metade se afasta do zero)."""
    return float(Decimal(v).quantize(Decimal(1).scaleb(-casas), rounding=ROUND_HALF_UP))


def num(v, casas=CASAS_TELA):
    return f"{arred(v, casas):.{casas}f}".replace(".", ",")


def num_tol(v):
    return f"{v:g}".replace(".", ",")


def dentro(v, limite):
    """|v| <= limite, sem o erro de 0,4000000001 passar do limite de 0,4."""
    return abs(v) <= limite + 1e-9


class Leitura:
    """Uma medição já processada (média aplicada)."""

    def __init__(self, nome, tipo, espectro, wave_start, lab_aparelho, cor, cfg):
        self.nome, self.tipo = nome, tipo
        self.espectro, self.wave_start = espectro, wave_start
        self.horario = dt.datetime.now()
        self.lab_aparelho = lab_aparelho
        self.recalcular(cor, cfg)

    def recalcular(self, cor, cfg):
        if self.espectro:
            self.lab1 = cor.espectro_para_lab(self.espectro, self.wave_start, cfg["iluminante1"])
            self.lab2 = cor.espectro_para_lab(self.espectro, self.wave_start, cfg["iluminante2"])
        else:
            self.lab1, self.lab2 = self.lab_aparelho, None


class App:
    def __init__(self, root):
        self.root = root
        self.cfg = carregar_config()
        self.cor = Colorimetria()
        self.fila_ui = queue.Queue()
        self.loop = asyncio.new_event_loop()
        threading.Thread(target=self.loop.run_forever, daemon=True).start()

        self.transporte = None
        self.aparelho = None
        self.encontrados = []
        self.padrao = None
        self.leituras = []
        self.acumulando = []           # medições parciais para a média
        self.cont = {"padrao": 0, "amostra": 0}
        self.ocupado = False
        self.atual = None

        root.title("Colorímetro DS-01" + ("  [SIMULAÇÃO]" if SIMULAR else ""))
        try:
            root.iconbitmap(os.path.join(AQUI, "colorimetro.ico"))
        except tk.TclError:
            pass
        root.geometry("1280x840")
        root.minsize(1120, 760)
        self._estilos()
        self._montar()
        self._atualizar_botoes()
        root.after(50, self._processar_fila)
        root.report_callback_exception = lambda *exc: self.log("erro na tela:\n" +"".join(traceback.format_exception(*exc)))
        if not SIMULAR:
            # A página da ponte fica disponível assim que o programa abre
            self.em_segundo_plano(TransporteCelular.iniciar_servidor(), None,
                                  lambda e: self.log(f"servidor da ponte não iniciou: {e}"))
        root.protocol("WM_DELETE_WINDOW", self._fechar)
        root.bind("<F5>", lambda e: self.medir("amostra"))
        root.bind("<F6>", lambda e: self.medir("padrao"))
        if self.cfg["ultimo_aparelho"]:
            a = self.cfg["ultimo_aparelho"]
            self.encontrados = [a]
            self._atualizar_botoes()

    # ------------------------------------------------------------------ infraestrutura
    def _estilos(self):
        r = self.root
        r.configure(bg=FUNDO)
        s = ttk.Style()
        s.theme_use("clam")
        s.configure(".", background=FUNDO, foreground=TEXTO, fieldbackground=CARTAO2,
                    bordercolor=CARTAO2, lightcolor=CARTAO2, darkcolor=CARTAO2,
                    troughcolor=CARTAO, font=(FONTE, 10))
        s.configure("TFrame", background=FUNDO)
        s.configure("TLabel", background=FUNDO, foreground=TEXTO)
        s.configure("Titulo.TLabel", foreground=CIANO, font=(FONTE, 11))
        s.configure("TButton", background=CARTAO2, foreground=TEXTO, padding=(12, 6), borderwidth=0)
        s.map("TButton", background=[("active", misturar(CARTAO2, CIANO, 0.25))])
        for w in ("TCombobox", "TEntry", "TSpinbox"):
            s.configure(w, fieldbackground=CARTAO2, foreground=TEXTO, background=CARTAO2,
                        arrowcolor=CIANO, insertcolor=TEXTO, selectbackground=CARTAO2)
            s.map(w, fieldbackground=[("readonly", CARTAO2)], foreground=[("readonly", TEXTO)],
                  selectbackground=[("readonly", CARTAO2)], selectforeground=[("readonly", TEXTO)])
        r.option_add("*TCombobox*Listbox.background", CARTAO2)
        r.option_add("*TCombobox*Listbox.foreground", TEXTO)
        r.option_add("*TCombobox*Listbox.selectBackground", misturar(CARTAO2, CIANO, 0.4))
        s.configure("Treeview", background=CARTAO, fieldbackground=CARTAO, foreground=TEXTO,
                    rowheight=28, borderwidth=0, font=(FONTE, 10))
        s.map("Treeview", background=[("selected", "#2f4a63")], foreground=[("selected", TEXTO)])
        s.configure("Treeview.Heading", background=CARTAO, foreground=FRACO, borderwidth=0,
                    relief="flat", font=(FONTE, 10))
        s.map("Treeview.Heading", background=[("active", CARTAO2)])
        s.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        for sb in ("TScrollbar", "Vertical.TScrollbar"):
            s.configure(sb, background=CARTAO2, troughcolor=CARTAO, arrowcolor=FRACO, gripcount=0,
                        bordercolor=CARTAO, lightcolor=CARTAO2, darkcolor=CARTAO2)
            s.map(sb, background=[("active", misturar(CARTAO2, CIANO, 0.3)), ("disabled", CARTAO)],
                  arrowcolor=[("disabled", misturar(FRACO, CARTAO, 0.6))])

    def em_segundo_plano(self, coro, ok=None, erro=None):
        fut = asyncio.run_coroutine_threadsafe(coro, self.loop)

        def terminou(f):
            try:
                r = f.result()
                self.fila_ui.put((ok, r) if ok else None)
            except Exception as e:  # noqa: BLE001
                self.fila_ui.put(((erro or self._mostrar_erro), e))
        fut.add_done_callback(terminou)

    def na_interface(self, fn, *args):
        self.fila_ui.put((lambda _r: fn(*args), None))

    def _processar_fila(self):
        try:
            while True:
                item = self.fila_ui.get_nowait()
                if item:
                    fn, arg = item
                    fn(arg)
        except queue.Empty:
            pass
        self.root.after(50, self._processar_fila)

    def _mostrar_erro(self, e):
        self.ocupado = False
        self._atualizar_botoes()
        self.status(f"Erro: {e}")
        messagebox.showerror("Colorímetro", str(e) or e.__class__.__name__)

    def status(self, texto):
        self.lbl_status["text"] = texto

    def log(self, texto):
        print(texto)
        try:
            with open(os.path.join(AQUI, "registro.log"), "a", encoding="utf-8") as f:
                f.write(f"{dt.datetime.now():%d/%m/%Y %H:%M:%S}  {texto}\n")
        except OSError:
            pass

    # ------------------------------------------------------------------ montagem da tela
    def _montar(self):
        r = self.root

        # --- cabeçalho, como o do app: "Medir Cor"  •  Medir  •  •••
        topo = tk.Frame(r, bg=FUNDO, height=52)
        topo.pack(fill="x", padx=20, pady=(6, 0))
        topo.pack_propagate(False)
        tk.Label(topo, text="Medir Cor", fg=CIANO, bg=FUNDO, font=(FONTE, 16)).pack(side="left")
        tk.Label(topo, text="Medir", fg=TEXTO, bg=FUNDO, font=(FONTE, 17, "bold")).place(relx=0.5, rely=0.5, anchor="center")
        mais = tk.Label(topo, text="•••", fg=TEXTO, bg=FUNDO, font=(FONTE, 16, "bold"), cursor="hand2")
        mais.pack(side="right")
        self.menu_mais = tk.Menu(r, tearoff=0, bg=CARTAO2, fg=TEXTO, activebackground=misturar(CARTAO2, CIANO, 0.4),
                                 activeforeground=TEXTO, bd=0, font=(FONTE, 10))
        self.menu_mais.add_command(label="Calibrar branco", command=lambda: self._pelo_menu(self.calibrar, "branco"))
        self.menu_mais.add_command(label="Calibrar preto", command=lambda: self._pelo_menu(self.calibrar, "preto"))
        self.menu_mais.add_separator()
        self.menu_mais.add_command(label="Configurações", command=self.abrir_config)
        self.menu_mais.add_command(label="Exportar para Excel (CSV)…", command=self.exportar)
        self.menu_mais.add_command(label="Histórico (todos os dias)…", command=self.abrir_historico)
        self.menu_mais.add_command(label="Limpar lista", command=self.limpar)
        mais.bind("<Button-1>", lambda e: self.menu_mais.tk_popup(e.x_root - 180, e.y_root + 12))

        rodape = tk.Frame(r, bg=FUNDO)
        rodape.pack(side="bottom", fill="x", padx=22, pady=(0, 8))
        self.lbl_status = tk.Label(rodape, text="Desconectado", fg=FRACO, bg=FUNDO, font=(FONTE, 9), anchor="w")
        self.lbl_status.pack(side="left", fill="x", expand=True)
        tk.Label(rodape, text=f"Salvo automaticamente em: {PASTA_MEDICOES}", fg=misturar(FRACO, FUNDO, 0.35),
                 bg=FUNDO, font=(FONTE, 9)).pack(side="right")

        corpo = tk.Frame(r, bg=FUNDO)
        corpo.pack(fill="both", expand=True, padx=14, pady=(4, 6))

        # ================= coluna esquerda: igual à tela "Medir" do celular
        esq = self.col_esq = tk.Frame(corpo, bg=FUNDO, width=480)
        esq.pack(side="left", fill="y")
        esq.pack_propagate(False)

        # cartão 1: tabela com páginas (bolinhas)
        c1 = Cartao(esq, pad=18)
        c1.configure(height=292)
        c1.pack(fill="x", padx=6, pady=6)
        t = c1.corpo
        self.lbl_pagina = tk.Label(t, text="", fg=FRACO, bg=CARTAO, font=(FONTE, 9))
        self.lbl_pagina.grid(row=0, column=0, columnspan=4, sticky="w")
        for j in range(4):
            t.grid_columnconfigure(j, weight=1, uniform="tab")
        self.tab_cab = []
        for j in range(1, 4):
            l = tk.Label(t, text="", fg=TEXTO, bg=CARTAO, font=(FONTE, 13))
            l.grid(row=1, column=j, pady=(6, 10))
            self.tab_cab.append(l)
        self.tab_linhas = []
        for i in range(4):
            nome = tk.Label(t, text="", fg=TEXTO, bg=CARTAO, font=(FONTE, 13))
            nome.grid(row=2 + i, column=0, pady=6)
            vals = []
            for j in range(1, 4):
                l = tk.Label(t, text="", fg=TEXTO, bg=CARTAO, font=(FONTE, 13))
                l.grid(row=2 + i, column=j, pady=6)
                vals.append(l)
            self.tab_linhas.append((nome, vals))
        t.grid_rowconfigure(6, weight=1)
        self.cv_pontos = tk.Canvas(t, height=20, bg=CARTAO, highlightthickness=0, cursor="hand2")
        self.cv_pontos.grid(row=7, column=0, columnspan=4, sticky="ew")
        self.cv_pontos.bind("<Configure>", lambda e: self._desenhar_pontos())
        self.cv_pontos.bind("<Button-1>", self._clicou_pontos)
        for w in [t, self.lbl_pagina, *self.tab_cab] + [x for n, v in self.tab_linhas for x in (n, *v)]:
            w.bind("<MouseWheel>", lambda e: self._mudar_pagina(-1 if e.delta > 0 else 1))
        self.pagina = 0

        # cartão 2: modo/iluminante, dE grande e os dois blocos de cor
        c2 = Cartao(esq, pad=18)
        c2.pack(fill="both", expand=True, padx=6, pady=6)
        k = c2.corpo
        self.lbl_modo = tk.Label(k, text="SCI/D65/10°", fg=TEXTO, bg=CARTAO, font=(FONTE, 14))
        self.lbl_modo.pack()
        # barra verde/vermelha "dE*ab=0,3 Aprovado", como no app
        self.lbl_de_grande = tk.Label(k, text="dE*ab=**", fg=TEXTO, bg=CARTAO, font=(FONTE, 18, "bold"))
        self.lbl_de_grande.pack(fill="x", pady=(2, 6))
        self.lbl_descricao = tk.Label(k, text="", fg=FRACO, bg=CARTAO, font=(FONTE, 10), wraplength=400)
        self.lbl_descricao.pack(pady=(0, 10))
        blocos = tk.Frame(k, bg=CARTAO)
        blocos.pack(fill="both", expand=True, padx=14)
        self.cv_blocos = {}
        for tipo in ("padrao", "amostra"):
            cv = tk.Canvas(blocos, bg=CARTAO, highlightthickness=0, cursor="hand2", width=1, height=1)
            cv.pack(side="left", fill="both", expand=True, padx=(0, 7) if tipo == "padrao" else (7, 0))
            cv.bind("<Configure>", lambda e: self.desenhar_cores())
            cv.bind("<Button-1>", lambda e, tp=tipo: self._clicou_bloco(tp))
            self.cv_blocos[tipo] = cv

        # botão "Conectar" branco, como no app
        self.bt_conexao = Pilula(esq, "⇄  Conectar", self.abrir_conexao, estilo="branco", altura=60, fonte=15, fundo=FUNDO)
        self.bt_conexao.pack(fill="x", padx=6, pady=(10, 4))

        # ================= coluna direita: aparelho, espectro e lista
        dir_ = self.col_dir = tk.Frame(corpo, bg=FUNDO)
        dir_.pack(side="left", fill="both", expand=True)

        ca = Cartao(dir_, pad=16, ajustar_altura=True)
        ca.pack(fill="x", padx=6, pady=6)
        a = ca.corpo
        a.grid_columnconfigure(0, weight=1)
        a.grid_columnconfigure(1, weight=1)
        tk.Label(a, text="Aparelho", fg=CIANO, bg=CARTAO, font=(FONTE, 12)).grid(row=0, column=0, sticky="w")
        tk.Label(a, text="Calibração", fg=CIANO, bg=CARTAO, font=(FONTE, 12)).grid(row=0, column=1, sticky="w")
        self.lbl_info = tk.Label(a, text="—", fg=TEXTO, bg=CARTAO, font=(FONTE, 10), justify="left", anchor="w")
        self.lbl_info.grid(row=1, column=0, sticky="nw", pady=(4, 0))
        self.lbl_cal = tk.Label(a, text="—", fg=TEXTO, bg=CARTAO, font=(FONTE, 10), justify="left", anchor="w")
        self.lbl_cal.grid(row=1, column=1, sticky="nw", pady=(4, 0))
        botoes = tk.Frame(a, bg=CARTAO)
        botoes.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        self.bt_cal_branco = Pilula(botoes, "Calibrar branco", lambda: self.calibrar("branco"), altura=42)
        self.bt_cal_branco.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.bt_cal_preto = Pilula(botoes, "Calibrar preto", lambda: self.calibrar("preto"), altura=42)
        self.bt_cal_preto.pack(side="left", fill="x", expand=True, padx=(8, 0))
        lb = tk.Frame(a, bg=CARTAO)
        lb.grid(row=3, column=0, columnspan=2, sticky="w", pady=(10, 0))
        tk.Label(lb, text="Botão do aparelho mede como:", fg=FRACO, bg=CARTAO, font=(FONTE, 10)).pack(side="left")
        self.var_botao = tk.StringVar(value="amostra")
        for texto, valor in (("Amostra", "amostra"), ("Padrão", "padrao")):
            tk.Radiobutton(lb, text=texto, value=valor, variable=self.var_botao, fg=TEXTO, bg=CARTAO,
                           selectcolor=CARTAO2, activebackground=CARTAO, activeforeground=TEXTO,
                           font=(FONTE, 10), bd=0, highlightthickness=0).pack(side="left", padx=(10, 0))
        self.lbl_media = tk.Label(a, text="", fg=LARANJA, bg=CARTAO, font=(FONTE, 10))
        self.lbl_media.grid(row=4, column=0, columnspan=2, sticky="w")
        if SIMULAR:
            Pilula(a, "Simular botão do aparelho", lambda: self.transporte and self.transporte.apertar_botao(0),
                   estilo="contorno", altura=34, fonte=10).grid(row=5, column=0, columnspan=2, sticky="ew", pady=(6, 0))

        ce = Cartao(dir_, pad=12)
        ce.configure(height=250)
        ce.pack(fill="x", padx=6, pady=6)
        self.cv_espectro = tk.Canvas(ce.corpo, bg=CARTAO, highlightthickness=0)
        self.cv_espectro.pack(fill="both", expand=True)
        self.cv_espectro.bind("<Configure>", lambda e: self.desenhar_espectro())

        cl = Cartao(dir_, pad=12)
        cl.pack(fill="both", expand=True, padx=6, pady=6)
        topo_lista = tk.Frame(cl.corpo, bg=CARTAO)
        topo_lista.pack(fill="x", pady=(0, 6))
        tk.Label(topo_lista, text="Medições desta sessão", fg=CIANO, bg=CARTAO, font=(FONTE, 11)).pack(side="left")
        ver = tk.Label(topo_lista, text="Histórico de todos os dias  ›", fg=CIANO, bg=CARTAO,
                       font=(FONTE, 11, "underline"), cursor="hand2")
        ver.pack(side="right")
        ver.bind("<Button-1>", lambda e: self.abrir_historico())
        fh = tk.Frame(cl.corpo, bg=CARTAO)
        fh.pack(fill="both", expand=True)
        cols = ("nome", "hora", "L", "a", "b", "C", "h", "dE", "dE00", "res")
        self.tree = ttk.Treeview(fh, columns=cols, show="headings", height=6)
        larg = {"nome": 110, "hora": 74, "res": 96}
        tit = {"nome": "Nome", "hora": "Hora", "L": "L*", "a": "a*", "b": "b*", "C": "C*", "h": "h°",
               "dE": "dE*ab", "dE00": "dE00", "res": "Resultado"}
        for c in cols:
            self.tree.heading(c, text=tit[c])
            self.tree.column(c, width=larg.get(c, 62), anchor="w" if c in ("nome", "res") else "e")
        self.tree.tag_configure("ok", foreground=VERDE)
        self.tree.tag_configure("falha", foreground=VERMELHO)
        self.tree.tag_configure("padrao", background="#28344a")
        sb = ttk.Scrollbar(fh, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._selecionou)
        self.menu = tk.Menu(self.root, tearoff=0, bg=CARTAO2, fg=TEXTO, activebackground=misturar(CARTAO2, CIANO, 0.4),
                            activeforeground=TEXTO, bd=0, font=(FONTE, 10))
        self.menu.add_command(label="Usar como padrão", command=self._usar_como_padrao)
        self.menu.add_command(label="Renomear…", command=self._renomear)
        self.menu.add_command(label="Excluir", command=self._excluir)
        self.tree.bind("<Button-3>", self._menu_contexto)

        r.bind("<Left>", lambda e: self._mudar_pagina(-1))
        r.bind("<Right>", lambda e: self._mudar_pagina(1))
        self.modo_celular = False
        r.bind("<Configure>", self._ao_redimensionar, add="+")
        self.mostrar(None)

    def _atualizar_botoes(self):
        con = self.aparelho is not None
        livre = not self.ocupado
        est = lambda cond: "normal" if cond else "disabled"  # noqa: E731
        if con:
            nome = (self.aparelho.info.serie if self.aparelho.info else "") or "aparelho"
            self.bt_conexao["text"] = f"⇄  Desconectar  ·  {nome}"
        else:
            self.bt_conexao["text"] = "Conectando…" if self.ocupado else "⇄  Conectar"
        self.bt_conexao["state"] = est(livre)
        for b in (self.bt_cal_branco, self.bt_cal_preto):
            b["state"] = est(con and livre)
        for cv in self.cv_blocos.values():
            cv.configure(cursor="hand2" if con and livre else "arrow")

    def _ao_redimensionar(self, e):
        if e.widget is not self.root:
            return
        # Janela encaixada lado a lado (estreita em relação à altura): layout do celular, uma coluna
        celular = e.width < 1.35 * e.height or e.width < 900
        if celular == self.modo_celular:
            return
        self.modo_celular = celular
        if celular:
            self.col_dir.pack_forget()
            self.col_esq.pack_configure(fill="both", expand=True)
        else:
            self.col_esq.pack_configure(fill="y", expand=False)
            self.col_dir.pack(side="left", fill="both", expand=True)

    def _pelo_menu(self, fn, *args):
        if not self.aparelho or self.ocupado:
            self.status("Conecte o aparelho primeiro" if not self.aparelho else "Aguarde a operação em andamento")
            return
        fn(*args)

    def _clicou_bloco(self, tipo):
        # No app, tocar no alvo do bloco faz a medição daquele tipo
        if not self.aparelho:
            self.status("Conecte o aparelho para medir")
            return
        self.medir(tipo)

    # ---- páginas da tabela (bolinhas)
    PAGINAS = 4

    def _mudar_pagina(self, passo):
        self.pagina = (self.pagina + passo) % self.PAGINAS
        self.mostrar(self.atual)

    def _clicou_pontos(self, e):
        w = self.cv_pontos.winfo_width()
        x0 = w / 2 - (self.PAGINAS - 1) * 14
        i = round((e.x - x0) / 28)
        if 0 <= i < self.PAGINAS:
            self.pagina = i
            self.mostrar(self.atual)

    def _desenhar_pontos(self):
        cv = self.cv_pontos
        cv.delete("all")
        w = cv.winfo_width()
        x0 = w / 2 - (self.PAGINAS - 1) * 14
        for i in range(self.PAGINAS):
            x = x0 + 28 * i
            cv.create_oval(x - 7, 3, x + 7, 17, fill=CIANO if i == self.pagina else "#4a5163", outline="")

    # ---- janela "Conectando" (como a do app)
    def abrir_conexao(self):
        if self.aparelho:
            self.desconectar()
            return
        if getattr(self, "jan_conexao", None) and self.jan_conexao.winfo_exists():
            self.jan_conexao.lift()
            return
        j = tk.Toplevel(self.root, bg=FUNDO)
        j.title("Conectar")
        j.geometry(f"520x560+{self.root.winfo_rootx() + 60}+{self.root.winfo_rooty() + 60}")
        j.transient(self.root)
        try:
            j.iconbitmap(os.path.join(AQUI, "colorimetro.ico"))
        except tk.TclError:
            pass
        self.jan_conexao = j
        cab = tk.Frame(j, bg=FUNDO, height=54)
        cab.pack(fill="x", padx=18)
        cab.pack_propagate(False)
        tk.Label(cab, text="Conectando", fg=TEXTO, bg=FUNDO, font=(FONTE, 15, "bold")).place(relx=0.5, rely=0.5, anchor="center")
        de_novo = tk.Label(cab, text="Procurar de novo", fg=CIANO, bg=FUNDO, font=(FONTE, 11), cursor="hand2")
        de_novo.pack(side="right")
        de_novo.bind("<Button-1>", lambda e: self.ocupado or self.procurar())
        tk.Label(j, text="Aparelhos:", fg=TEXTO, bg=FUNDO, font=(FONTE, 13), anchor="w").pack(fill="x", padx=20, pady=(4, 6))
        self.lista_conexao = tk.Frame(j, bg=FUNDO)
        self.lista_conexao.pack(fill="both", expand=True, padx=16)
        dica = ("Pelo celular: no Chrome do celular abra " + " ou ".join(TransporteCelular.enderecos())
                + " e toque em \"Conectar ao colorímetro\".") if not SIMULAR else ""
        tk.Label(j, text=dica, fg=FRACO, bg=FUNDO, font=(FONTE, 9), wraplength=470, justify="left").pack(
            fill="x", padx=20, pady=(0, 14))
        self._preencher_conexao()
        if not self.ocupado:
            self.procurar()

    def _preencher_conexao(self, procurando=False):
        j = getattr(self, "jan_conexao", None)
        if not j or not j.winfo_exists():
            return
        for w in self.lista_conexao.winfo_children():
            w.destroy()
        for alvo in self.encontrados:
            c = Cartao(self.lista_conexao, pad=12, ajustar_altura=True)
            c.pack(fill="x", pady=5)
            tk.Label(c.corpo, text=("★ " if alvo.get("provavel") else "") + alvo["nome"], fg=TEXTO, bg=CARTAO,
                     font=(FONTE, 12), anchor="w").grid(row=0, column=0, sticky="w")
            sub = alvo["endereco"] + (f"   {alvo['rssi']} dBm" if alvo.get("rssi") else "")
            tk.Label(c.corpo, text=sub, fg=FRACO, bg=CARTAO, font=(FONTE, 9), anchor="w").grid(row=1, column=0, sticky="w")
            c.corpo.grid_columnconfigure(0, weight=1)
            Pilula(c.corpo, "Conectar", lambda a=alvo: self.conectar(a), estilo="contorno", altura=40,
                   fonte=12, largura=130).grid(row=0, column=1, rowspan=2, padx=(8, 0))
        if procurando:
            tk.Label(self.lista_conexao, text="Procurando aparelhos…", fg=FRACO, bg=FUNDO,
                     font=(FONTE, 11)).pack(pady=14)
        elif not self.encontrados:
            tk.Label(self.lista_conexao, text="Nenhum aparelho encontrado", fg=FRACO, bg=FUNDO,
                     font=(FONTE, 11)).pack(pady=14)

    # ------------------------------------------------------------------ conexão
    def novo_transporte(self, alvo):
        if SIMULAR:
            return TransporteSimulado()
        if alvo["endereco"] == TransporteCelular.ENDERECO:
            return TransporteCelular()
        return TransporteBLE()

    def procurar(self):
        self.ocupado = True
        self._atualizar_botoes()
        self._preencher_conexao(procurando=True)
        self.status("Procurando aparelhos Bluetooth (8 s)… deixe o DS-01 ligado e perto do PC")
        if SIMULAR:
            busca = TransporteSimulado.procurar(8)
        else:
            async def busca_real():
                lista = []
                try:
                    lista = await TransporteBLE.procurar(8)
                except Exception as e:  # noqa: BLE001  sem adaptador no PC: ainda dá pelo celular
                    self.log(f"busca Bluetooth no PC falhou: {e}")
                return await TransporteCelular.procurar() + lista
            busca = busca_real()

        def ok(lista):
            self.ocupado = False
            self.encontrados = lista
            self._preencher_conexao()
            if lista:
                self.status(f"{len(lista)} aparelho(s) encontrado(s). ★ = provável colorímetro")
            else:
                self.status("Nenhum aparelho encontrado")
            self._atualizar_botoes()

        def erro(e):
            msg = str(e)
            if "adapter" in msg.lower() or "radio" in msg.lower() or "bluetooth" in msg.lower():
                msg = ("Não encontrei um adaptador Bluetooth ligado neste PC.\n\n"
                       "Conecte o adaptador USB e confira se o Bluetooth está ativado em "
                       "Configurações > Dispositivos.\n\nDetalhe: " + str(e))
            self._preencher_conexao()
            self._mostrar_erro(Exception(msg))
        self.em_segundo_plano(busca, ok, erro)

    def conectar(self, alvo):
        if self.ocupado:
            return
        if getattr(self, "jan_conexao", None) and self.jan_conexao.winfo_exists():
            self.jan_conexao.destroy()
        self.ocupado = True
        self._atualizar_botoes()
        self.status(f"Conectando a {alvo['nome']}…")
        t = self.novo_transporte(alvo)
        if isinstance(t, TransporteCelular):
            self.status("Conectando pelo celular… abra no Chrome do celular: "
                        + " ou ".join(TransporteCelular.enderecos()) + " e toque em Conectar")
        t.ao_desconectar = lambda: self.na_interface(self._caiu)
        dev = DS01(t, ao_medir_no_botao=self._botao_aparelho, log=self.log)

        async def fluxo():
            await t.conectar(alvo["endereco"])
            self.log(f"conectado a {alvo['endereco']}")
            await asyncio.sleep(0.5)
            for tentativa in range(1, 4):
                try:
                    info = await dev.ler_info()
                    break
                except ErroAparelho as e:
                    # Às vezes a assinatura de notificações não "pega": assinar de novo
                    self.log(f"leitura de informações falhou (tentativa {tentativa}): {e}")
                    if tentativa == 3:
                        raise
                    await t.reassinar()
                    await asyncio.sleep(0.5)
            self.log(f"aparelho: {info}")
            cal = await dev.ler_calibracao()
            return info, cal

        def ok(res):
            info, cal = res
            self.transporte, self.aparelho = t, dev
            self.ocupado = False
            self.cfg["ultimo_aparelho"] = {k: alvo[k] for k in ("endereco", "nome", "provavel")}
            self.salvar_config()
            modo = self._modo()
            self.lbl_info["text"] = (f"Modelo: {info.modelo or '?'}\nSérie: {info.serie}\n"
                                     f"Firmware: {info.firmware}\nCódigo: {info.codigo}   Modo: {MODOS.get(modo, modo)}")
            self._mostrar_cal(cal)
            self.status(f"Conectado a {alvo['nome']}")
            self._atualizar_botoes()
            self.mostrar(self.atual)
            if not (cal.branco_ok and cal.preto_ok):
                messagebox.showinfo("Calibração", "O aparelho informa que precisa de calibração.\n"
                                    "Use os botões 'Calibrar branco' e 'Calibrar preto'.")

        def erro(e):
            self.log(f"falha ao conectar: {type(e).__name__}: {e}")
            self.em_segundo_plano(t.desconectar())
            self._mostrar_erro(Exception(f"Não foi possível conectar: {e}"))
        self.em_segundo_plano(fluxo(), ok, erro)

    def _mostrar_cal(self, cal):
        def fmt(ok_, ts):
            quando = dt.datetime.fromtimestamp(ts).strftime("%d/%m/%Y %H:%M") if ts else "nunca"
            return ("✔ " if ok_ else "✘ ") + quando
        self.lbl_cal["text"] = f"Branco: {fmt(cal.branco_ok, cal.branco_data)}\nPreto:  {fmt(cal.preto_ok, cal.preto_data)}"

    def desconectar(self):
        t = self.transporte
        self.transporte, self.aparelho = None, None
        if t:
            t.ao_desconectar = None
            self.em_segundo_plano(t.desconectar())
        self.status("Desconectado")
        self.lbl_info["text"] = "—"
        self.lbl_cal["text"] = "—"
        self._atualizar_botoes()

    def _caiu(self):
        if self.aparelho:
            self.transporte, self.aparelho = None, None
            self.ocupado = False
            self.status("⚠ Conexão perdida — clique em Conectar novamente")
            self._atualizar_botoes()

    def _modo(self):
        if self.cfg["modo"] != "auto":
            return int(self.cfg["modo"])
        return self.aparelho.info.modo_padrao() if self.aparelho and self.aparelho.info else 0

    # ------------------------------------------------------------------ calibração
    def calibrar(self, tipo):
        if tipo == "branco":
            msg = "Coloque o aparelho sobre a placa de calibração BRANCA e clique OK."
        else:
            msg = "Coloque o aparelho sobre a cavidade/placa de calibração PRETA (ou aponte para o vazio) e clique OK."
        if not messagebox.askokcancel("Calibração", msg):
            return
        self.ocupado = True
        self._atualizar_botoes()
        self.status(f"Calibrando {tipo}…")

        async def fluxo():
            ok_ = await self.aparelho.calibrar(tipo)
            return ok_, await self.aparelho.ler_calibracao()

        def ok(res):
            sucesso, cal = res
            self.ocupado = False
            self._mostrar_cal(cal)
            self._atualizar_botoes()
            self.status(f"Calibração {tipo}: {'concluída' if sucesso else 'FALHOU'}")
            if not sucesso:
                messagebox.showwarning("Calibração", f"O aparelho recusou a calibração do {tipo}. Tente novamente.")
        self.em_segundo_plano(fluxo(), ok)

    # ------------------------------------------------------------------ medição
    def medir(self, tipo):
        if not self.aparelho or self.ocupado:
            return
        self.ocupado = True
        self._atualizar_botoes()
        self.status("Medindo…")
        self.em_segundo_plano(self.aparelho.medir(self._modo()), lambda m: self._chegou(m, tipo))

    def _botao_aparelho(self, modo, _erro):
        """Chamado (na thread do Bluetooth) quando o botão do DS-01 é apertado."""
        def iniciar():
            if self.ocupado or not self.aparelho:
                return
            tipo = self.var_botao.get()
            if tipo == "amostra" and self.padrao is None and not self.acumulando:
                tipo = "padrao"
            self.ocupado = True
            self._atualizar_botoes()
            self.status("Medição feita no aparelho — baixando…")
            self.em_segundo_plano(self.aparelho.baixar_medicao(modo), lambda m: self._chegou(m, tipo))
        self.na_interface(iniciar)

    def _chegou(self, m, tipo):
        self.ocupado = False
        self._atualizar_botoes()
        n = max(1, int(self.cfg.get("media", 1)))
        if self.acumulando and self.acumulando[0][1] != tipo:
            self.acumulando = []
        self.acumulando.append((m, tipo))
        if len(self.acumulando) < n:
            self.lbl_media["text"] = f"Média: {len(self.acumulando)}/{n} — reposicione e meça de novo"
            self.status(f"Leitura {len(self.acumulando)} de {n} registrada")
            return
        medidas = [x[0] for x in self.acumulando]
        self.acumulando = []
        self.lbl_media["text"] = ""
        if medidas[0].espectro:
            esp = [sum(v) / len(v) for v in zip(*[x.espectro for x in medidas])]
            lab_ap = None
        else:
            esp = []
            lab_ap = tuple(sum(v) / len(v) for v in zip(*[x.lab_aparelho for x in medidas]))
        self.cont[tipo] += 1
        prefixo = self.cfg["nome_padrao"] if tipo == "padrao" else self.cfg["nome_amostra"]
        lei = Leitura(f"{prefixo} {self.cont[tipo]}", tipo, esp, medidas[0].wave_start, lab_ap, self.cor, self.cfg)
        self.leituras.append(lei)
        self._inserir_linha(lei)
        if tipo == "padrao":
            self.padrao = lei
            self._redesenhar_lista()
        self._salvar_automatico(lei)
        self.mostrar(lei)
        self.status(f"{lei.nome} medido às {lei.horario:%H:%M:%S}")

    # ------------------------------------------------------------------ cálculo e exibição
    def comparar(self, ref, amo):
        tol = self.cfg["tol"]
        l1, l2 = ref.lab1, amo.lab1
        dL, da, db = (l2[i] - l1[i] for i in range(3))
        C1, _ = lch(*l1)
        C2, _ = lch(*l2)
        dC = C2 - C1
        dEab = delta_e_ab(l1, l2)
        dH = math.copysign(math.sqrt(max(0.0, dEab ** 2 - dL ** 2 - dC ** 2)),
                           l1[1] * l2[2] - l2[1] * l1[2])
        des = {"dE*ab": dEab, "dE00": delta_e_2000(l1, l2, *self.cfg["k00"]),
               "dEcmc": delta_e_cmc(l1, l2, self.cfg["cmc_l"], self.cfg["cmc_c"]), "dE94": delta_e_94(l1, l2, *self.cfg["k94"])}
        f = self.cfg["formula"]
        # Como o app: aprova/reprova só pelo dE escolhido, com o valor exato (getColorDiffTips)
        aprovado = des[f] <= tol[f]
        # Indicações por coluna (Pass/White/Green...): o app usa a diferença dos valores já
        # arredondados na tela (getDiffValueTips). Não entram na aprovação.
        d1 = tuple(arred(l2[i]) - arred(l1[i]) for i in range(3))
        dicas, desc = [], []
        for v, chave, (mais, menos), (txt_mais, txt_menos) in zip(
                d1, ("dL", "da", "db"), (("Claro", "Escuro"), ("Vermelho", "Verde"), ("Amarelo", "Azul")),
                (("mais clara", "mais escura"), ("mais vermelha", "mais verde"), ("mais amarela", "mais azul"))):
            if dentro(v, tol[chave]):
                dicas.append(("Aprovado", True))
            else:
                dicas.append((mais if v > 0 else menos, False))
                desc.append(txt_mais if v > 0 else txt_menos)
        return {"d": (dL, da, db, dC, dH), "d1": d1, "dicas": dicas, "des": des, "aprovado": aprovado, "desc": desc}

    def mostrar(self, lei):
        self.atual = lei
        ref = self.padrao
        amo = lei if lei is not ref else None
        r = self.comparar(ref, amo) if ref and amo else None
        tol = self.cfg["tol"]
        il1, il2 = self.cfg["iluminante1"], self.cfg["iluminante2"]

        def f(v, sinal=False):
            return "-" if v is None else (("+" if sinal and v > 0 else "") + num(v))

        def cor_tol(v, chave):
            return TEXTO if v is None else (VERDE if dentro(v, tol[chave]) else VERMELHO)

        p = self.pagina
        if p == 0:
            titulo, cab = f"L*a*b*  ·  {il1}°", ["L*", "a*", "b*"]
            d = r["d1"] if r else (None,) * 3
            linhas = [("Padrão", [f(v) for v in (ref.lab1 if ref else (None,) * 3)], None),
                      ("Amostra", [f(v) for v in (amo.lab1 if amo else (None,) * 3)], None),
                      ("Diferença", [f(d[i], True) for i in range(3)], None)]
            if r:   # linha Pass/Green/... embaixo da diferença, como no app
                linhas.append(("", [t for t, _ in r["dicas"]], [VERDE if ok_ else VERMELHO for _, ok_ in r["dicas"]]))
        elif p == 1:
            titulo, cab = f"L*C*h°  ·  {il1}°", ["L*", "C*", "h°"]
            lch_ = lambda x: (x.lab1[0], *lch(*x.lab1)) if x else (None,) * 3  # noqa: E731
            d = r["d"] if r else (None,) * 5
            linhas = [("Padrão", [f(v) for v in lch_(ref)], None),
                      ("Amostra", [f(v) for v in lch_(amo)], None),
                      ("Diferença", [f(d[0], True), f(d[3], True), "ΔH " + f(d[4], True) if r else "-"],
                       [cor_tol(d[0], "dL"), cor_tol(d[3], "dC"), cor_tol(d[4], "dH")])]
        elif p == 2:
            titulo, cab = f"Diferença de cor  ·  {il1}°", ["ΔE", "Tolerância", "Resultado"]
            linhas = []
            for fo in FORMULAS:
                v = r["des"][fo] if r else None
                ok_ = None if v is None else v <= tol[fo]
                linhas.append((fo, [f(v), num_tol(tol[fo]), "-" if ok_ is None else ("OK" if ok_ else "Fora")],
                               [cor_tol(v, fo), FRACO, TEXTO if ok_ is None else (VERDE if ok_ else VERMELHO)]))
        else:
            titulo, cab = f"Segundo iluminante  ·  {il2}°  (metamerismo)", ["L*", "a*", "b*"]
            l2r = ref.lab2 if ref and ref.lab2 else None
            l2a = amo.lab2 if amo and amo.lab2 else None
            dif = [l2a[i] - l2r[i] for i in range(3)] if l2r and l2a else [None] * 3
            linhas = [("Padrão", [f(v) for v in (l2r or (None,) * 3)], None),
                      ("Amostra", [f(v) for v in (l2a or (None,) * 3)], None),
                      ("Diferença", [f(v, True) for v in dif], [cor_tol(dif[i], k) for i, k in enumerate(("dL", "da", "db"))])]
            if l2r and l2a:
                linhas.append(("ΔE*ab", [num(delta_e_ab(l2r, l2a)), "", ""], None))

        self.lbl_pagina["text"] = titulo
        for j, l in enumerate(self.tab_cab):
            l["text"] = cab[j] if j < len(cab) else ""
        for i, (nome, vals) in enumerate(self.tab_linhas):
            if i < len(linhas):
                rot, vs, cores = linhas[i]
                nome["text"] = rot
                for j, l in enumerate(vals):
                    l["text"] = vs[j] if j < len(vs) else ""
                    l["fg"] = cores[j] if cores and j < len(cores) else TEXTO
            else:
                nome["text"] = ""
                for l in vals:
                    l["text"] = ""
        self._desenhar_pontos()

        # cartão do meio: SCI/D65/10°  e  dE*ab=…
        modo = MODOS.get(self._modo(), "SCI") if hasattr(self, "_modo") else "SCI"
        self.lbl_modo["text"] = f"{modo}/{il1}°"
        fo = self.cfg["formula"]
        if r:
            self.lbl_de_grande.config(text=f"{fo}={num(r['des'][fo])} {'Aprovado' if r['aprovado'] else 'Reprovado'}",
                                      fg="#ffffff", bg=VERDE_APP if r["aprovado"] else VERMELHO_APP)
            self.lbl_descricao["text"] = ("Amostra " + ", ".join(r["desc"])) if r["desc"] else "Dentro da tolerância em L*, a* e b*"
        else:
            self.lbl_de_grande.config(text=f"{fo}=**", fg=TEXTO, bg=CARTAO)
            self.lbl_descricao["text"] = "Clique em Padrão ou Amostra para medir" if self.aparelho else "Conecte o aparelho para medir"
        self.desenhar_cores()
        self.desenhar_espectro()

    def desenhar_cores(self):
        ref = self.padrao
        amo = self.atual if self.atual is not ref else None
        for tipo, lei, rotulo in (("padrao", ref, "Padrão"), ("amostra", amo, "Amostra")):
            cv = self.cv_blocos[tipo]
            cv.delete("all")
            w, h = cv.winfo_width(), cv.winfo_height()
            if w < 20:
                continue
            cor = lab_para_hex(*lei.lab1) if lei and lei.lab1 else CINZA_VAZIO
            faixa = 58
            cv.create_rectangle(0, 0, w, h - faixa, fill=cor, outline="")
            cv.create_rectangle(0, h - faixa, w, h, fill=misturar(cor, "#000000", 0.12), outline="")
            claro = lei and lei.lab1 and lei.lab1[0] > 72
            tc = "#1d2230" if claro else "#ffffff"
            cy = h - faixa / 2
            texto = rotulo if not lei else lei.nome
            tw = len(texto) * 9
            cx = w / 2 - tw / 2 - 16
            cv.create_oval(cx - 12, cy - 12, cx + 12, cy + 12, outline=tc, width=1.6)
            cv.create_oval(cx - 3, cy - 3, cx + 3, cy + 3, outline=tc, width=1.6)
            for dx, dy in ((0, -1), (0, 1), (-1, 0), (1, 0)):
                cv.create_line(cx + dx * 9, cy + dy * 9, cx + dx * 15, cy + dy * 15, fill=tc, width=1.6)
            cv.create_text(cx + 24, cy, text=texto, anchor="w", fill=tc, font=(FONTE, 14))

    def desenhar_espectro(self):
        cv = self.cv_espectro
        cv.delete("all")
        w, h = cv.winfo_width(), cv.winfo_height()
        if w < 50:
            return
        my, x0, y1 = 26, 44, h - 24
        x1 = w - 26
        cv.create_text(x0 - 36, 8, text="Refletância espectral", anchor="w", fill=CIANO, font=(FONTE, 11))
        maxv = 100
        series = [(s, c) for s, c in ((self.padrao, CIANO),
                  (self.atual if self.atual is not self.padrao else None, LARANJA)) if s and s.espectro]
        for s, _ in series:
            maxv = max(maxv, max(s.espectro) * 1.05)
        X = lambda wl: x0 + (wl - 400) / 300 * (x1 - x0)  # noqa: E731
        Y = lambda v: y1 - v / maxv * (y1 - my)          # noqa: E731
        for v in range(0, int(maxv) + 1, 20):
            cv.create_line(x0, Y(v), x1, Y(v), fill="#2f3647")
            cv.create_text(x0 - 6, Y(v), text=f"{v}%", anchor="e", fill=FRACO, font=(FONTE, 8))
        for wl in range(400, 701, 50):
            cv.create_text(X(wl), y1 + 12, text=f"{wl} nm", fill=FRACO, font=(FONTE, 8))
        for s, cor in series:
            pts = []
            for i, v in enumerate(s.espectro):
                wl = s.wave_start + 10 * i
                if 400 <= wl <= 700:
                    pts += [X(wl), Y(v)]
            if len(pts) >= 4:
                cv.create_line(*pts, fill=cor, width=2.5, smooth=True)
        legenda = [(s.nome, c) for s, c in series]
        for k, (t, c) in enumerate(legenda):
            y = 8 + 0 * k
            xl = x1 - 150 + 80 * k
            cv.create_line(xl, y, xl + 18, y, fill=c, width=3)
            cv.create_text(xl + 24, y, text=t, anchor="w", fill=TEXTO, font=(FONTE, 9))
        if not series:
            cv.create_text((x0 + x1) / 2, (my + y1) / 2, text="A curva aparece após a primeira medição",
                           fill=FRACO, font=(FONTE, 10))

    # ------------------------------------------------------------------ histórico
    def _valores_linha(self, lei):
        C, hh = lch(*lei.lab1)
        dE = dE00 = res = ""
        tags = ["padrao"] if lei.tipo == "padrao" else []
        if lei.tipo == "amostra" and self.padrao:
            r = self.comparar(self.padrao, lei)
            dE, dE00 = num(r["des"]["dE*ab"]), num(r["des"]["dE00"])
            res = "Aprovado" if r["aprovado"] else "Reprovado"
            tags.append("ok" if r["aprovado"] else "falha")
        elif lei.tipo == "padrao":
            res = "(padrão)" if lei is self.padrao else "padrão antigo"
        return (lei.nome, f"{lei.horario:%H:%M:%S}", *(num(v) for v in lei.lab1), num(C), num(hh), dE, dE00, res), tags

    def _inserir_linha(self, lei):
        vals, tags = self._valores_linha(lei)
        iid = self.tree.insert("", 0, values=vals, tags=tags)
        lei.iid = iid
        self.tree.selection_set(iid)

    def _redesenhar_lista(self):
        for lei in self.leituras:
            vals, tags = self._valores_linha(lei)
            self.tree.item(lei.iid, values=vals, tags=tags)

    def _leitura_do_iid(self, iid):
        return next((l for l in self.leituras if getattr(l, "iid", None) == iid), None)

    def _selecionou(self, _e=None):
        sel = self.tree.selection()
        if sel:
            lei = self._leitura_do_iid(sel[0])
            if lei:
                self.mostrar(lei)

    def _menu_contexto(self, e):
        iid = self.tree.identify_row(e.y)
        if iid:
            self.tree.selection_set(iid)
            self.menu.tk_popup(e.x_root, e.y_root)

    def _usar_como_padrao(self):
        sel = self.tree.selection()
        lei = sel and self._leitura_do_iid(sel[0])
        if lei:
            lei.tipo = "padrao"
            self.padrao = lei
            self._redesenhar_lista()
            self.mostrar(lei)

    def _renomear(self):
        from tkinter import simpledialog
        sel = self.tree.selection()
        lei = sel and self._leitura_do_iid(sel[0])
        if lei:
            novo = simpledialog.askstring("Renomear", "Novo nome:", initialvalue=lei.nome, parent=self.root)
            if novo:
                lei.nome = novo
                self._redesenhar_lista()

    def _excluir(self):
        sel = self.tree.selection()
        lei = sel and self._leitura_do_iid(sel[0])
        if lei:
            self.leituras.remove(lei)
            self.tree.delete(lei.iid)
            if lei is self.padrao:
                self.padrao = None
            self.mostrar(None)

    def limpar(self):
        if self.leituras and messagebox.askyesno("Limpar", "Apagar a lista da tela?\n"
                                                 "(o arquivo salvo automaticamente é mantido)"):
            self.leituras.clear()
            self.tree.delete(*self.tree.get_children())
            self.padrao = None
            self.cont = {"padrao": 0, "amostra": 0}
            self.mostrar(None)

    # ------------------------------------------------------------------ histórico salvo (todos os dias)
    def _ler_dia(self, arq):
        """Lê um CSV de medicoes/ e devolve as leituras, já comparadas com as regras atuais."""
        def fl(s):
            return float(s.replace(",", ".")) if s else None
        lidas, padroes = [], {}
        with open(arq, encoding="utf-8-sig", newline="") as f:
            linhas = list(csv.reader(f, delimiter=";"))
        cab = linhas[0] if linhas else []
        col = {nome: i for i, nome in enumerate(cab)}
        ondas = [(i, int(n.split()[0])) for i, n in enumerate(cab) if n.endswith(" nm")]
        for ln in linhas[1:]:
            if len(ln) < len(cab):
                ln = ln + [""] * (len(cab) - len(ln))
            try:
                lab = tuple(fl(ln[col[k]]) for k in ("L*", "a*", "b*"))
                hora = dt.datetime.strptime(f"{ln[col['Data']]} {ln[col['Hora']]}", "%d/%m/%Y %H:%M:%S")
            except (KeyError, ValueError, TypeError):
                continue
            esp = [fl(ln[i]) for i, _ in ondas]
            espectro = esp if esp and all(v is not None for v in esp) else []
            tipo = "padrao" if ln[col["Tipo"]] == "Padrão" else "amostra"
            lei = Leitura(ln[col["Nome"]], tipo, espectro, ondas[0][1] if ondas else 400, lab, self.cor, self.cfg)
            lei.horario = hora
            # o contador reinicia a cada vez que o programa abre: vale o último padrão com esse nome
            lei.ref = padroes.get(ln[col["Padrão de referência"]]) if tipo == "amostra" else None
            if tipo == "padrao":
                padroes[lei.nome] = lei
            lidas.append(lei)
        return lidas

    def abrir_historico(self):
        j = getattr(self, "jan_hist", None)
        if j and j.winfo_exists():
            j.lift()
            return
        dias = sorted((f for f in os.listdir(PASTA_MEDICOES) if f.startswith("medicoes_") and f.endswith(".csv")),
                      reverse=True) if os.path.isdir(PASTA_MEDICOES) else []
        rotulo = {}
        for f in dias:
            try:
                rotulo[dt.datetime.strptime(f[9:19], "%Y-%m-%d").strftime("%d/%m/%Y")] = f
            except ValueError:
                pass
        if not rotulo:
            messagebox.showinfo("Histórico", "Ainda não há medições salvas.")
            return
        j = self.jan_hist = tk.Toplevel(self.root, bg=FUNDO)
        j.title("Histórico")
        j.geometry(f"1000x640+{self.root.winfo_rootx() + 60}+{self.root.winfo_rooty() + 40}")
        j.transient(self.root)
        try:
            j.iconbitmap(os.path.join(AQUI, "colorimetro.ico"))
        except tk.TclError:
            pass

        cab = tk.Frame(j, bg=FUNDO, height=54)
        cab.pack(fill="x", padx=18)
        cab.pack_propagate(False)
        tk.Label(cab, text="Histórico", fg=TEXTO, bg=FUNDO, font=(FONTE, 15, "bold")).place(relx=0.5, rely=0.5, anchor="center")
        tk.Label(cab, text="Dia:", fg=FRACO, bg=FUNDO, font=(FONTE, 11)).pack(side="left")
        var_dia = tk.StringVar(value=next(iter(rotulo)))
        cb = ttk.Combobox(cab, textvariable=var_dia, values=list(rotulo), state="readonly", width=12)
        cb.pack(side="left", padx=8)
        resumo = tk.Label(j, text="", fg=FRACO, bg=FUNDO, font=(FONTE, 10), anchor="w")
        resumo.pack(fill="x", padx=20)

        rod = tk.Frame(j, bg=FUNDO)
        rod.pack(side="bottom", fill="x", padx=18, pady=(6, 14))

        c = Cartao(j, pad=12)
        c.pack(fill="both", expand=True, padx=14, pady=4)
        cols = ("hora", "nome", "tipo", "L", "a", "b", "ref", "dE", "dE00", "res")
        tit = {"hora": "Hora", "nome": "Nome", "tipo": "Tipo", "L": "L*", "a": "a*", "b": "b*",
               "ref": "Comparado com", "dE": "dE*ab", "dE00": "dE00", "res": "Resultado"}
        larg = {"hora": 80, "nome": 120, "tipo": 80, "ref": 130, "res": 100}
        tree = ttk.Treeview(c.corpo, columns=cols, show="headings")
        for k in cols:
            tree.heading(k, text=tit[k])
            tree.column(k, width=larg.get(k, 64), anchor="w" if k in ("nome", "tipo", "ref", "res") else "e")
        tree.tag_configure("ok", foreground=VERDE)
        tree.tag_configure("falha", foreground=VERMELHO)
        tree.tag_configure("padrao", background="#28344a")
        sb = ttk.Scrollbar(c.corpo, command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        por_iid = {}

        def carregar(*_):
            tree.delete(*tree.get_children())
            por_iid.clear()
            try:
                lidas = self._ler_dia(os.path.join(PASTA_MEDICOES, rotulo[var_dia.get()]))
            except OSError as e:
                resumo["text"] = f"Não consegui ler o arquivo: {e}"
                return
            aprov = reprov = 0
            for lei in reversed(lidas):       # mais recente em cima
                dE = dE00 = res = ""
                tags = ["padrao"] if lei.tipo == "padrao" else []
                if lei.ref:
                    r = self.comparar(lei.ref, lei)
                    dE, dE00 = num(r["des"]["dE*ab"]), num(r["des"]["dE00"])
                    res = "Aprovado" if r["aprovado"] else "Reprovado"
                    tags.append("ok" if r["aprovado"] else "falha")
                    aprov += r["aprovado"]
                    reprov += not r["aprovado"]
                iid = tree.insert("", "end", tags=tags, values=(
                    f"{lei.horario:%H:%M:%S}", lei.nome, "Padrão" if lei.tipo == "padrao" else "Amostra",
                    *(num(v) for v in lei.lab1), lei.ref.nome if lei.ref else "", dE, dE00, res))
                por_iid[iid] = lei
            fo = self.cfg["formula"]
            resumo["text"] = (f"{len(lidas)} medições  ·  {aprov} aprovadas  ·  {reprov} reprovadas   "
                              f"(resultado pelas regras atuais: {fo} ≤ {num_tol(self.cfg['tol'][fo])})")
        cb.bind("<<ComboboxSelected>>", carregar)

        def usar_como_padrao():
            sel = tree.selection()
            lei = por_iid.get(sel[0]) if sel else None
            if not lei:
                messagebox.showinfo("Histórico", "Selecione uma medição na lista.", parent=j)
                return
            novo = Leitura(lei.nome, "padrao", lei.espectro, lei.wave_start, lei.lab_aparelho, self.cor, self.cfg)
            novo.horario = lei.horario
            self.leituras.append(novo)
            self._inserir_linha(novo)
            self.padrao = novo
            self._redesenhar_lista()
            self.mostrar(novo)
            self.status(f"{lei.nome} de {lei.horario:%d/%m %H:%M} virou o padrão. Agora meça a amostra")
            j.destroy()

        menu = tk.Menu(j, tearoff=0, bg=CARTAO2, fg=TEXTO, activebackground=misturar(CARTAO2, CIANO, 0.4),
                       activeforeground=TEXTO, bd=0, font=(FONTE, 10))
        menu.add_command(label="Usar como padrão", command=usar_como_padrao)

        def contexto(e):
            iid = tree.identify_row(e.y)
            if iid:
                tree.selection_set(iid)
                menu.tk_popup(e.x_root, e.y_root)
        tree.bind("<Button-3>", contexto)
        tree.bind("<Double-1>", lambda e: usar_como_padrao())

        Pilula(rod, "Usar como padrão", usar_como_padrao, altura=44, fonte=11, largura=1, fundo=FUNDO).pack(
            side="left", fill="x", expand=True, padx=(0, 8))
        Pilula(rod, "Abrir no Excel", lambda: os.startfile(os.path.join(PASTA_MEDICOES, rotulo[var_dia.get()])),
               estilo="contorno", altura=44, fonte=11, largura=1, fundo=FUNDO).pack(side="left", fill="x", expand=True, padx=8)
        Pilula(rod, "Abrir pasta", lambda: os.startfile(PASTA_MEDICOES), estilo="contorno", altura=44, fonte=11,
               largura=1, fundo=FUNDO).pack(side="left", fill="x", expand=True, padx=(8, 0))
        j.bind("<Escape>", lambda e: j.destroy())
        carregar()

    # ------------------------------------------------------------------ arquivos
    def _linha_csv(self, lei):
        C, hh = lch(*lei.lab1)
        linha = [lei.horario.strftime("%d/%m/%Y"), lei.horario.strftime("%H:%M:%S"), lei.nome,
                 "Padrão" if lei.tipo == "padrao" else "Amostra", self.cfg["iluminante1"],
                 *(num(v, 3) for v in (*lei.lab1, C, hh))]
        if lei.tipo == "amostra" and self.padrao and lei is not self.padrao:
            r = self.comparar(self.padrao, lei)
            linha += [self.padrao.nome, *(num(v, 3) for v in r["d"][:3]),
                      *(num(r["des"][f], 3) for f in FORMULAS), "Aprovado" if r["aprovado"] else "Reprovado"]
        else:
            linha += [""] * 9
        esp = {lei.wave_start + 10 * i: v for i, v in enumerate(lei.espectro)}
        linha += [num(esp[w], 3) if w in esp else "" for w in range(400, 701, 10)]
        return linha

    def _cabecalho_csv(self):
        return (["Data", "Hora", "Nome", "Tipo", "Iluminante", "L*", "a*", "b*", "C*", "h°",
                 "Padrão de referência", "dL*", "da*", "db*", *FORMULAS, "Resultado"]
                + [f"{w} nm" for w in range(400, 701, 10)])

    def _salvar_automatico(self, lei):
        try:
            os.makedirs(PASTA_MEDICOES, exist_ok=True)
            arq = os.path.join(PASTA_MEDICOES, f"medicoes_{lei.horario:%Y-%m-%d}.csv")
            novo = not os.path.exists(arq)
            with open(arq, "a", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f, delimiter=";")
                if novo:
                    w.writerow(self._cabecalho_csv())
                w.writerow(self._linha_csv(lei))
        except OSError as e:
            self.status(f"Aviso: não consegui salvar automaticamente ({e})")

    def exportar(self):
        if not self.leituras:
            messagebox.showinfo("Exportar", "Nenhuma medição na lista.")
            return
        arq = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV (Excel)", "*.csv")],
                                           initialfile=f"medicoes_{dt.datetime.now():%Y-%m-%d_%H%M}.csv")
        if not arq:
            return
        with open(arq, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(self._cabecalho_csv())
            for lei in self.leituras:
                w.writerow(self._linha_csv(lei))
        self.status(f"Exportado: {arq}")

    def salvar_config(self):
        try:
            with open(ARQ_CONFIG, "w", encoding="utf-8") as f:
                json.dump(self.cfg, f, ensure_ascii=False, indent=2)
        except OSError:
            pass

    # ------------------------------------------------------------------ configurações (como a tela do app)
    def abrir_config(self):
        j = getattr(self, "jan_config", None)
        if j and j.winfo_exists():
            j.lift()
            return
        j = self.jan_config = tk.Toplevel(self.root, bg=FUNDO)
        j.title("Configurações")
        alt = min(860, j.winfo_screenheight() - 130)   # cabe na tela com a barra de tarefas
        j.geometry(f"520x{alt}+{self.root.winfo_rootx() + 80}+{max(0, (j.winfo_screenheight() - alt - 90) // 2)}")
        j.minsize(440, 520)
        j.transient(self.root)
        try:
            j.iconbitmap(os.path.join(AQUI, "colorimetro.ico"))
        except tk.TclError:
            pass
        rasc = json.loads(json.dumps(self.cfg))     # rascunho: só vale ao Voltar/Aplicar
        invalidos = set()
        estado = {"aba": "Parâmetro"}
        fmt = lambda v: f"{v:g}".replace(".", ",")  # noqa: E731
        menu_kw = dict(tearoff=0, bg=CARTAO2, fg=TEXTO, activebackground=misturar(CARTAO2, CIANO, 0.4),
                       activeforeground=TEXTO, bd=0, font=(FONTE, 10))

        # --- cabeçalho:  ‹ Voltar   Configurações   Resetar
        cab = tk.Frame(j, bg=FUNDO, height=56)
        cab.pack(fill="x", padx=14)
        cab.pack_propagate(False)
        voltar = tk.Label(cab, text="‹ Voltar", fg=CIANO, bg=FUNDO, font=(FONTE, 14), cursor="hand2")
        voltar.pack(side="left")
        tk.Label(cab, text="Configurações", fg=TEXTO, bg=FUNDO, font=(FONTE, 15, "bold")).place(relx=0.5, rely=0.5, anchor="center")
        resetar = tk.Label(cab, text="Resetar", fg=TEXTO, bg=FUNDO, font=(FONTE, 13), cursor="hand2")
        resetar.pack(side="right")

        # --- abas com sublinhado ciano
        barra = tk.Frame(j, bg=FUNDO)
        barra.pack(fill="x", padx=16)
        abas = {}
        for i, nome in enumerate(("Parâmetro", "Tolerância", "Outros")):
            barra.grid_columnconfigure(i, weight=1, uniform="aba")
            c = tk.Frame(barra, bg=FUNDO)
            c.grid(row=0, column=i, pady=(2, 8))
            l = tk.Label(c, text=nome, fg=TEXTO, bg=FUNDO, font=(FONTE, 13), cursor="hand2")
            l.pack()
            s = tk.Frame(c, height=3, bg=FUNDO)
            s.pack(fill="x", pady=(4, 0))
            l.bind("<Button-1>", lambda e, n=nome: trocar(n))
            abas[nome] = (l, s)

        # --- rodapé: Calibração  |  Aplicar
        rod = tk.Frame(j, bg=FUNDO)
        rod.pack(side="bottom", fill="x", padx=14, pady=(8, 14))
        bt_cal = Pilula(rod, "Calibração", lambda: calibracao(), altura=48, fonte=11, largura=1, fundo=FUNDO)
        bt_cal.pack(side="left", fill="x", expand=True, padx=(0, 8))
        bt_cal["state"] = "normal" if self.aparelho else "disabled"
        Pilula(rod, "Aplicar", lambda: aplicar(False), altura=48, fonte=11, largura=1, fundo=FUNDO).pack(
            side="left", fill="x", expand=True, padx=(8, 0))

        # --- área rolável (roda do mouse), como a lista do celular
        cv = tk.Canvas(j, bg=FUNDO, highlightthickness=0, yscrollincrement=20)
        cv.pack(fill="both", expand=True)
        corpo = tk.Frame(cv, bg=FUNDO)
        janela = cv.create_window(0, 0, window=corpo, anchor="nw")
        cv.bind("<Configure>", lambda e: cv.itemconfigure(janela, width=e.width))
        corpo.bind("<Configure>", lambda e: cv.configure(scrollregion=(0, 0, e.width, e.height)))
        j.bind("<MouseWheel>", lambda e: cv.yview_scroll(-3 if e.delta > 0 else 3, "units"))

        # ---------- peças da tela
        def secao(p, texto):
            tk.Label(p, text=texto, fg=CIANO, bg=FUNDO, font=(FONTE, 14), anchor="w").pack(fill="x", pady=(8, 4))
            tk.Frame(p, height=1, bg="#d6dae2").pack(fill="x", pady=(0, 12))

        def campo(pai, obter, guardar, bg, **kw):
            """Entrada numérica; valor inválido fica vermelho e não é guardado."""
            var = tk.StringVar(value=fmt(obter()))
            e = tk.Entry(pai, textvariable=var, bg=bg, fg=TEXTO, insertbackground=TEXTO, relief="flat", bd=0,
                         highlightthickness=0, font=(FONTE, 12), **kw)

            def mudou(*_):
                try:
                    x = float(var.get().replace(",", "."))
                    if x < 0 or math.isnan(x):
                        raise ValueError
                except ValueError:
                    invalidos.add(e)
                    e["fg"] = VERMELHO
                    return
                invalidos.discard(e)
                e["fg"] = TEXTO
                guardar(x)
            var.trace_add("write", mudou)
            return e

        def item(p, rotulo, valor="", opcoes=None, escolher=None, numero=None):
            """Linha em cartão: rótulo à esquerda, valor e › à direita (como no app)."""
            f = tk.Frame(p, bg=CARTAO, height=54)
            f.pack(fill="x", pady=(0, 10))
            f.pack_propagate(False)
            tk.Label(f, text=rotulo, fg=TEXTO, bg=CARTAO, font=(FONTE, 12)).pack(side="left", padx=16)
            if not (opcoes or numero):
                tk.Label(f, text=valor, fg=TEXTO, bg=CARTAO, font=(FONTE, 12)).pack(side="right", padx=16)
                return
            seta = tk.Label(f, text="›", fg=TEXTO, bg=CARTAO, font=(FONTE, 18))
            seta.pack(side="right", padx=(6, 16))
            if numero:
                e = campo(f, *numero, CARTAO, width=8, justify="right")
                e.pack(side="right")

                def clicou(_ev):
                    e.focus_set()
                    e.select_range(0, "end")
                    e.icursor("end")
            else:
                lv = tk.Label(f, text=valor, fg=TEXTO, bg=CARTAO, font=(FONTE, 12))
                lv.pack(side="right")

                def clicou(_ev):
                    m = tk.Menu(j, **menu_kw)
                    for texto, val in opcoes:
                        m.add_command(label=texto, command=lambda t=texto, v=val: (escolher(v), lv.configure(text=t)))
                    m.tk_popup(f.winfo_rootx() + f.winfo_width() - 160, f.winfo_rooty() + f.winfo_height())
            for w in (f, *f.winfo_children()):
                if not isinstance(w, tk.Entry):
                    w.configure(cursor="hand2")
                    w.bind("<Button-1>", clicou)

        def caixa(pai, texto, cor=TEXTO, borda=None, **kw):
            return tk.Label(pai, text=texto, fg=cor, bg=CAIXA, font=(FONTE, 12), highlightthickness=1 if borda else 0,
                            highlightbackground=borda or CAIXA, highlightcolor=borda or CAIXA, **kw)

        # ---------- aba Parâmetro
        def aba_parametro(p):
            modelo = (self.aparelho.info.modelo if self.aparelho and self.aparelho.info else "") or "DS-01"
            item(p, "Séries", modelo)
            item(p, "Refletância/Transmitância", "Refletância")
            modos = [("Automático", "auto")] + [(n, k) for k, n in MODOS.items()]
            item(p, "Modo teste", "Automático" if rasc["modo"] == "auto" else MODOS.get(int(rasc["modo"]), "?"),
                 modos, lambda v: rasc.__setitem__("modo", v))
            secao(p, "Configurações do display")
            item(p, "Diferença de cor", rasc["formula"], [(x, x) for x in FORMULAS],
                 lambda v: rasc.__setitem__("formula", v))
            secao(p, "Iluminante/Observador")
            ilums = list(dict.fromkeys(n.split("/")[0] for n in self.cor.nomes()))
            for chave, ordem in (("iluminante1", "Primeiro"), ("iluminante2", "Segundo")):
                il, ang = rasc[chave].split("/")
                item(p, f"{ordem} Iluminante", il, [(x, x) for x in ilums],
                     lambda v, c=chave: rasc.__setitem__(c, f"{v}/{rasc[c].split('/')[1]}"))
                item(p, f"{ordem} Ângulo", ang, [("2", "2"), ("10", "10")],
                     lambda v, c=chave: rasc.__setitem__(c, f"{rasc[c].split('/')[0]}/{v}"))
            secao(p, "CMC(l:c)")
            for chave, rot in (("cmc_l", "l"), ("cmc_c", "c")):
                item(p, rot, numero=(lambda c=chave: rasc[c], lambda x, c=chave: rasc.__setitem__(c, x)))
            for chave, titulo in (("k94", "CIE94"), ("k00", "CIE00")):
                secao(p, titulo)
                for i, rot in enumerate(("KL", "KC", "KH")):
                    item(p, rot, numero=(lambda c=chave, i=i: rasc[c][i],
                                         lambda x, c=chave, i=i: rasc[c].__setitem__(i, x)))

        # ---------- aba Tolerância
        def aba_tolerancia(p):
            secao(p, "CIE LAB")
            g = tk.Frame(p, bg=FUNDO)
            g.pack(fill="x", pady=(0, 6))
            for c, peso in enumerate((3, 2, 3, 1, 3)):
                g.grid_columnconfigure(c, weight=peso, uniform="lab")
            linha = 0

            def fileira(nome, op, valor_w, rotulo):
                nonlocal linha
                tk.Label(g, text=nome, fg=TEXTO, bg=FUNDO, font=(FONTE, 12)).grid(row=linha, column=0, sticky="w")
                tk.Label(g, text=op, fg=TEXTO, bg=FUNDO, font=(FONTE, 11)).grid(row=linha, column=1)
                valor_w.grid(row=linha, column=2, sticky="ew", ipady=7, pady=4)
                caixa(g, rotulo).grid(row=linha, column=4, sticky="ew", ipady=6, pady=4)
                linha += 1

            grupos = (("dL", "dL*", ("Claro", "Escuro")), ("da", "da*", ("Vermelho", "Verde")),
                      ("db", "db*", ("Amarelo", "Azul")), ("dE*ab", "dE*ab", ("Reprovado", None)))
            for chave, nome, (mais, menos) in grupos:
                v = rasc["tol"][chave]
                l_mais = tk.Label(g, text=fmt(v), fg=TEXTO, bg=FUNDO, font=(FONTE, 12))
                l_menos = tk.Label(g, text=fmt(-v), fg=TEXTO, bg=FUNDO, font=(FONTE, 12)) if menos else None

                def guardar(x, ch=chave, a=l_mais, b=l_menos):
                    rasc["tol"][ch] = x
                    a["text"] = fmt(x)
                    if b:
                        b["text"] = fmt(-x)
                e = campo(g, lambda ch=chave: rasc["tol"][ch], guardar, CAIXA, width=7, justify="center")
                fileira(nome, "≤±" if menos else "≤", e, "Aprovado")
                fileira(nome, ">", l_mais, mais)
                if menos:
                    fileira(nome, "<", l_menos, menos)
            for titulo, chaves in (("CIE LCH", (("dC", "dC*"), ("dH", "dH*"))),
                                   ("CIEDE2000", (("dL'", "dL'"), ("dC'", "dC'"), ("dH'", "dH'"), ("dE00", "dE00"))),
                                   ("CMC(l:c)&CIE94", (("dEcmc", "dEcmc"), ("dE94", "dE94")))):
                secao(p, titulo)
                for chave, rot in chaves:
                    item(p, rot, numero=(lambda c=chave: rasc["tol"][c],
                                         lambda x, c=chave: rasc["tol"].__setitem__(c, x)))

        # ---------- aba Outros
        def aba_outros(p):
            secao(p, "Média")
            f = tk.Frame(p, bg=FUNDO)
            f.pack(fill="x", padx=20, pady=(0, 14))
            media = int(rasc["media"]) > 1
            for c, (texto, ativo, n) in enumerate((("Simples", not media, 1), ("Média", media, 3))):
                f.grid_columnconfigure(c, weight=1, uniform="media")
                b = caixa(f, texto, CIANO if ativo else TEXTO, CIANO if ativo else None, cursor="hand2")
                b.configure(font=(FONTE, 13))
                b.grid(row=0, column=c, sticky="ew", ipady=12, padx=(0, 10) if c == 0 else (10, 0))
                if not ativo:
                    b.bind("<Button-1>", lambda e, n=n: (rasc.__setitem__("media", n), construir()))
            if media:
                item(p, "Leituras por medição", numero=(lambda: rasc["media"],
                                                        lambda x: rasc.__setitem__("media", max(2, int(round(x))))))
            secao(p, "Regras de nomenclatura")
            g = tk.Frame(p, bg=FUNDO)
            g.pack(fill="x")
            g.grid_columnconfigure(1, weight=3)
            g.grid_columnconfigure(3, weight=2)
            for r, (rot, chave) in enumerate((("Padrão", "nome_padrao"), ("Amostra", "nome_amostra"))):
                tk.Label(g, text=rot, fg=TEXTO, bg=FUNDO, font=(FONTE, 12), width=8, anchor="w").grid(row=r, column=0, sticky="w")
                var = tk.StringVar(value=rasc[chave])
                var.trace_add("write", lambda *_, c=chave, v=var: rasc.__setitem__(c, v.get()))
                tk.Entry(g, textvariable=var, bg=CAIXA, fg=TEXTO, insertbackground=TEXTO, relief="flat", bd=0,
                         highlightthickness=0, font=(FONTE, 12), justify="center").grid(
                    row=r, column=1, sticky="ew", ipady=12, pady=6)
                tk.Label(g, text="+", fg=TEXTO, bg=FUNDO, font=(FONTE, 13)).grid(row=r, column=2, padx=18)
                caixa(g, "Nº", CIANO, CIANO).grid(row=r, column=3, sticky="ew", ipady=12, pady=6)

        def construir():
            for w in corpo.winfo_children():
                w.destroy()
            invalidos.clear()
            for n, (l, s) in abas.items():
                l["fg"] = CIANO if n == estado["aba"] else TEXTO
                s["bg"] = CIANO if n == estado["aba"] else FUNDO
            p = tk.Frame(corpo, bg=FUNDO)
            p.pack(fill="x", padx=16, pady=(6, 10))
            {"Parâmetro": aba_parametro, "Tolerância": aba_tolerancia, "Outros": aba_outros}[estado["aba"]](p)
            cv.yview_moveto(0)

        def trocar(nome):
            estado["aba"] = nome
            construir()

        def aplicar(fechar):
            if invalidos:
                messagebox.showerror("Configurações", "Algum valor numérico está inválido (em vermelho).", parent=j)
                return False
            novo = json.loads(json.dumps(rasc))
            for k in ("nome_padrao", "nome_amostra"):
                novo[k] = novo[k].strip() or CONFIG_PADRAO[k]
            novo["ultimo_aparelho"] = self.cfg.get("ultimo_aparelho")
            self.cfg = novo
            self.salvar_config()
            for lei in self.leituras:
                lei.recalcular(self.cor, self.cfg)
            self._redesenhar_lista()
            self.mostrar(self.atual)
            if fechar:
                j.destroy()
            else:
                self.status("Configurações aplicadas")
            return True

        def resetar_tudo():
            ultimo = self.cfg.get("ultimo_aparelho")
            rasc.clear()
            rasc.update(json.loads(json.dumps(CONFIG_PADRAO)))
            rasc["ultimo_aparelho"] = ultimo
            construir()

        def calibracao():
            m = tk.Menu(j, **menu_kw)
            for tipo in ("branco", "preto"):
                m.add_command(label=f"Calibrar {tipo}",
                              command=lambda t=tipo: aplicar(True) and self._pelo_menu(self.calibrar, t))
            m.tk_popup(bt_cal.winfo_rootx() + 10, bt_cal.winfo_rooty() - 60)

        voltar.bind("<Button-1>", lambda e: aplicar(True))
        resetar.bind("<Button-1>", lambda e: resetar_tudo())
        j.protocol("WM_DELETE_WINDOW", lambda: aplicar(True))
        j.bind("<Escape>", lambda e: aplicar(True))
        construir()

    # ------------------------------------------------------------------ saída
    def _fechar(self):
        if self.transporte:
            t = self.transporte
            t.ao_desconectar = None
            try:
                asyncio.run_coroutine_threadsafe(t.desconectar(), self.loop).result(timeout=5)
            except Exception:  # noqa: BLE001
                pass
        self.root.destroy()


def main():
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:  # noqa: BLE001
        pass
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
