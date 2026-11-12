import csv
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import FormulaRule
from datetime import datetime


"""
Lançador de Validades.

Interface gráfica (tkinter) que lê os arquivos .txt exportados pelo aplicativo
de inventário (formato SKU|CÓDIGO DE BARRAS|NOME|VALIDADE|QUANTIDADE) e
alimenta a planilha de controle de validades, respeitando as regras de
lote (produto novo, mês livre, mesmo mês já preenchido) e aplicando a
formatação condicional de vencimento.
"""

# =========================================================
#  PASTAS PADRÃO
# =========================================================

# Pasta onde ficam os arquivos .txt do coletor. Pode ser definida pela
# variável de ambiente VALIDADE_PASTA_TXT (ex.: um compartilhamento de rede).
PASTA_TXT  = os.environ.get("VALIDADE_PASTA_TXT", os.path.dirname(os.path.abspath(__file__)))

# Pasta onde está a planilha controle_validade.xlsx (mesma pasta do script)
PASTA_XLSX = os.path.dirname(os.path.abspath(__file__))


# =========================================================
#  CORES
# =========================================================
BG      = "#1E2230"
BG2     = "#252A3A"
BG3     = "#2E3447"
ACCENT  = "#4A9EFF"
ACCENT2 = "#2D7DD2"
SUCCESS = "#3DDC84"
ERROR   = "#FF5C5C"
TEXT    = "#E8EAF0"
TEXT2   = "#9DA3B4"
BORDER  = "#3A4058"


# =========================================================
#  LOGICA
# =========================================================

def ler_txt(caminho_txt):
    """Le arquivo no formato SKU|CODIGO DE BARRAS|NOME|VALIDADE|QUANTIDADE."""
    produtos = []
    encodings = ["utf-8", "latin-1", "cp1252"]
    for enc in encodings:
        try:
            with open(caminho_txt, newline="", encoding=enc) as f:
                reader = csv.reader(f, delimiter="|")
                for linha in reader:
                    if len(linha) < 5:
                        continue
                    sku      = linha[0].strip()
                    barcode  = linha[1].strip()
                    nome     = linha[2].strip()
                    validade = linha[3].strip()
                    qtd      = linha[4].strip()
                    if not barcode:
                        continue
                    produtos.append({
                        "sku":        sku,
                        "barcode":    barcode,
                        "nome":       nome,
                        "validade":   validade,
                        "quantidade": qtd,
                    })
            return produtos
        except UnicodeDecodeError:
            continue
    raise ValueError("Nao foi possivel decodificar o arquivo TXT.")


def parse_validade(texto):
    for fmt in ("%d-%m-%y", "%d-%m-%Y", "%d/%m/%y", "%d/%m/%Y"):
        try:
            return datetime.strptime(texto.strip(), fmt).date()
        except ValueError:
            continue
    return texto


DAT_FONT  = Font(name="Calibri", size=10)
ALN_CTR   = Alignment(horizontal="center", vertical="center")
ALN_LEFT  = Alignment(horizontal="left",   vertical="center")
NOME_ABA  = "Validades"


def _celula_real(ws, row, col):
    """Retorna a celula principal da mesclagem, ou a propria celula se nao estiver mesclada.
    Evita o erro 'MergedCell object attribute value is read-only'."""
    from openpyxl.utils import get_column_letter
    coord = f"{get_column_letter(col)}{row}"
    for merged in ws.merged_cells.ranges:
        if coord in merged:
            return ws.cell(row=merged.min_row, column=merged.min_col)
    return ws.cell(row=row, column=col)


def _mapa_colunas_meses(ws):
    """Le a linha 1 e monta dicionario {(ano, mes): coluna_data}."""
    mapa = {}
    for cell in ws[1]:
        if cell.value and hasattr(cell.value, 'year'):
            mapa[(cell.value.year, cell.value.month)] = cell.column
    return mapa


def _ultima_linha_com_dado(ws):
    """Retorna a ultima linha que tem valor em col B, ignorando linhas vazias."""
    ultima = 1
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
        if row[0].value is not None:
            ultima = row[0].row
    return ultima


def _ultima_linha_do_produto(ws, linha_inicio):
    """Retorna a ultima linha do bloco do produto baseado na mesclagem de col A."""
    coord_a = f"A{linha_inicio}"
    for merged in ws.merged_cells.ranges:
        if coord_a in merged:
            return merged.max_row
    return linha_inicio


def _expandir_mesclagem_abc(ws, linha_inicio, nova_ultima):
    """Remove mesclagem atual de A/B/C e recria abrangendo ate nova_ultima."""
    for col_letra in ("A", "B", "C"):
        for merged in list(ws.merged_cells.ranges):
            if f"{col_letra}{linha_inicio}" in merged:
                ws.unmerge_cells(str(merged))
                break
        if nova_ultima > linha_inicio:
            ws.merge_cells(f"{col_letra}{linha_inicio}:{col_letra}{nova_ultima}")


def atualizar_planilha(produtos, caminho_xlsx):
    if not os.path.exists(caminho_xlsx):
        print(f"[ERRO] Planilha nao encontrada: {caminho_xlsx}")
        return

    wb = load_workbook(caminho_xlsx)

    if NOME_ABA not in wb.sheetnames:
        print(f"[ERRO] Aba '{NOME_ABA}' nao encontrada na planilha.")
        return

    ws = wb[NOME_ABA]

    # Mapa (ano, mes) -> coluna da DATA (quantidade vai na coluna seguinte)
    mapa_meses = _mapa_colunas_meses(ws)

    # Monta indice: barcode (col B) -> primeira linha do bloco
    # Dados a partir da linha 2 (linha 1 agora e o cabecalho)
    indice = {}
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
        cb = row[0]
        if cb.value:
            linha  = cb.row
            bc_val = str(cb.value).strip()
            indice[bc_val] = linha
            # Garante formula SKU em col A se estiver vazia
            ca = _celula_real(ws, linha, 1)
            if ca.value is None:
                ca.value     = f'=IFERROR(VLOOKUP(B{linha},Produtos!A:B,2,0),"")'
                ca.font      = DAT_FONT
                ca.alignment = ALN_CTR

    novos = atualizados = sem_coluna = 0
    proxima_linha_nova = _ultima_linha_com_dado(ws) + 1

    for p in produtos:
        bc       = p["barcode"]
        validade = parse_validade(p["validade"])
        qtd      = p["quantidade"]
        sku      = p.get("sku", "")

        # Se o arquivo trouxer SKU e o produto ja existir na planilha, atualiza col A
        if sku and bc in indice:
            ca = _celula_real(ws, indice[bc], 1)
            if ca.value != sku:
                ca.value     = sku
                ca.font      = DAT_FONT
                ca.alignment = ALN_CTR

        # Coluna da data do mes e coluna da quantidade (sempre ao lado)
        if hasattr(validade, 'year'):
            chave    = (validade.year, validade.month)
            col_data = mapa_meses.get(chave)
        else:
            col_data = None

        if col_data is None:
            print(f"[AVISO] Coluna para validade '{p['validade']}' nao encontrada. Produto ignorado: {bc}")
            sem_coluna += 1
            continue

        col_qtd = col_data + 1  # Qtd sempre na coluna seguinte a data

        if bc in indice:
            linha_inicio = indice[bc]
            ultima_linha = _ultima_linha_do_produto(ws, linha_inicio)

            # Procura se ja existe data e quantidade neste mes no bloco
            linha_mes  = None
            for l in range(linha_inicio, ultima_linha + 1):
                if ws.cell(row=l, column=col_qtd).value is not None:
                    linha_mes = l
                    break

            if linha_mes is not None:
                # Mesmo mes ja preenchido: manter data mais antiga, somar qtd, comentario
                dc_existente = ws.cell(row=linha_mes, column=col_data)
                qc_existente = ws.cell(row=linha_mes, column=col_qtd)

                data_existente = dc_existente.value
                qtd_existente  = qc_existente.value

                # Converte quantidades para int para somar
                try:
                    qtd_soma = int(str(qtd_existente).split('.')[0]) + int(str(qtd).split('.')[0])
                except (ValueError, TypeError):
                    qtd_soma = qtd

                # Determina qual data e mais antiga
                nova_data = validade
                if hasattr(data_existente, 'date'):
                    data_existente = data_existente.date()

                if hasattr(data_existente, 'year') and hasattr(nova_data, 'year'):
                    data_mais_antiga = min(data_existente, nova_data)
                else:
                    data_mais_antiga = data_existente

                # Atualiza celula com data mais antiga e quantidade somada
                dc_existente.value         = data_mais_antiga
                dc_existente.number_format = "dd/mm/yyyy"
                qc_existente.value         = qtd_soma

                # Comentario: "Recebido X unidades dia Y" (acumula se ja houver)
                from openpyxl.comments import Comment
                hoje_str        = datetime.now().strftime("%d/%m/%Y")
                nova_linha_comt = f"Recebido {qtd} unidades dia {hoje_str}"
                comt_existente  = dc_existente.comment
                if comt_existente:
                    texto_comentario = comt_existente.text + "\n" + nova_linha_comt
                else:
                    texto_comentario = nova_linha_comt
                dc_existente.comment = Comment(texto_comentario, "Lançador de Validades")

            else:
                # Coluna livre no bloco: insere normalmente e ja adiciona comentario
                from openpyxl.comments import Comment
                dc = ws.cell(row=ultima_linha, column=col_data, value=validade)
                dc.font, dc.alignment = DAT_FONT, ALN_CTR
                dc.number_format = "dd/mm/yyyy"

                qc = ws.cell(row=ultima_linha, column=col_qtd, value=qtd)
                qc.font, qc.alignment = DAT_FONT, ALN_CTR

                hoje_str     = datetime.now().strftime("%d/%m/%Y")
                texto_comt   = f"Recebido {qtd} unidades dia {hoje_str}"
                dc.comment   = Comment(texto_comt, "Lançador de Validades")

            atualizados += 1

        else:
            # Novo produto: insere na proxima linha disponivel (rastreada internamente)
            nl = proxima_linha_nova
            ws.row_dimensions[nl].height = 19.5

            # Desfaz qualquer mesclagem existente na linha destino
            for merged in list(ws.merged_cells.ranges):
                if merged.min_row <= nl <= merged.max_row:
                    ws.unmerge_cells(str(merged))

            # Grava o barcode primeiro em col B para confirmar o numero real da linha
            bc_cell = ws.cell(row=nl, column=2, value=bc)
            bc_cell.font, bc_cell.alignment = DAT_FONT, ALN_CTR
            linha_real = bc_cell.row

            # Col A: SKU do arquivo se disponivel, caso contrario formula VLOOKUP
            sc = ws.cell(row=linha_real, column=1)
            if p.get("sku"):
                sc.value = p["sku"]
            else:
                sc.value = f'=IFERROR(VLOOKUP(B{linha_real},Produtos!A:B,2,0),"")'
            sc.font      = DAT_FONT
            sc.alignment = ALN_CTR

            nm_cell = ws.cell(row=linha_real, column=3, value=p["nome"])
            nm_cell.font, nm_cell.alignment = DAT_FONT, ALN_LEFT

            dc = ws.cell(row=linha_real, column=col_data, value=validade)
            dc.font, dc.alignment = DAT_FONT, ALN_CTR
            dc.number_format = "dd/mm/yyyy"

            qc = ws.cell(row=linha_real, column=col_qtd, value=qtd)
            qc.font, qc.alignment = DAT_FONT, ALN_CTR

            from openpyxl.comments import Comment
            hoje_str   = datetime.now().strftime("%d/%m/%Y")
            dc.comment = Comment(f"Recebido {qtd} unidades dia {hoje_str}", "Lançador de Validades")

            indice[bc] = linha_real
            proxima_linha_nova += 1
            novos += 1

    _aplicar_formatacao_condicional(ws, mapa_meses)
    wb.save(caminho_xlsx)
    print(f"[OK] Planilha atualizada: {caminho_xlsx}")
    print(f"     {novos} produto(s) novo(s) inserido(s)")
    print(f"     {atualizados} produto(s) atualizado(s)")
    if sem_coluna:
        print(f"     {sem_coluna} produto(s) ignorado(s) - mes/ano nao encontrado na planilha")


def _aplicar_formatacao_condicional(ws, mapa_meses):
    """Aplica formatacao condicional nas colunas de data:
    - Vermelho: vence em ate 30 dias
    - Laranja:  vence em ate 60 dias
    - Amarelo:  vence em ate 90 dias
    """
    from openpyxl.utils import get_column_letter

    fill_vermelho = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    fill_laranja  = PatternFill(start_color="FF8000", end_color="FF8000", fill_type="solid")
    fill_amarelo  = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")

    font_branca = Font(color="FFFFFF", bold=True, name="Calibri", size=10)
    font_preta  = Font(color="000000", bold=True, name="Calibri", size=10)

    # Remove regras anteriores para nao acumular a cada execucao
    ws.conditional_formatting = ws.conditional_formatting.__class__()

    for (ano, mes), col in mapa_meses.items():
        col_letra = get_column_letter(col)
        intervalo = f"{col_letra}3:{col_letra}1048576"

        # Ordem importa: vermelho primeiro (mais urgente), stopIfTrue garante prioridade
        ws.conditional_formatting.add(intervalo, FormulaRule(
            formula=[f'AND({col_letra}3<>"",{col_letra}3<=TODAY()+30)'],
            stopIfTrue=True, font=font_branca, fill=fill_vermelho
        ))
        ws.conditional_formatting.add(intervalo, FormulaRule(
            formula=[f'AND({col_letra}3<>"",{col_letra}3<=TODAY()+60)'],
            stopIfTrue=True, font=font_preta, fill=fill_laranja
        ))
        ws.conditional_formatting.add(intervalo, FormulaRule(
            formula=[f'AND({col_letra}3<>"",{col_letra}3<=TODAY()+90)'],
            stopIfTrue=True, font=font_preta, fill=fill_amarelo
        ))


def executar_conversao(csv_paths, xlsx_path, log_fn, done_fn):
    try:
        if isinstance(csv_paths, str):
            csv_paths = [csv_paths]

        log_fn(f"[..] {len(csv_paths)} arquivo(s) selecionado(s)")
        for p in csv_paths:
            log_fn(f"[..] - {Path(p).name}")

        if xlsx_path:
            log_fn(f"[..] XLSX: {xlsx_path}")

        # Le todos os produtos de todos os arquivos
        todos_produtos = []
        for cp in csv_paths:
            prods = ler_txt(cp)
            log_fn(f"[OK] {len(prods)} produto(s) lido(s) de {Path(cp).name}")
            todos_produtos.extend(prods)

        log_fn(f"[OK] Total: {len(todos_produtos)} produto(s) lido(s)")

        if not xlsx_path:
            log_fn(f"[..] Nenhuma planilha selecionada - nada a fazer")
            log_fn(f"[OK] Concluido!")
            done_fn(True, str(Path(csv_paths[0]).parent))
            return

        # Filtra apenas produtos com data de validade
        produtos_com_validade = [p for p in todos_produtos if p["validade"].strip()]
        sem_validade = len(todos_produtos) - len(produtos_com_validade)
        if sem_validade:
            log_fn(f"[..] {sem_validade} produto(s) sem data de validade - ignorados na planilha")

        atualizar_planilha(produtos_com_validade, xlsx_path)
        log_fn(f"[OK] Planilha de validades atualizada")
        log_fn(f"[OK] Concluido com sucesso!")
        done_fn(True, str(Path(csv_paths[0]).parent))

    except Exception as e:
        import traceback
        log_fn(f"[ERRO] {e}")
        for line in traceback.format_exc().splitlines():
            log_fn(f"[ERRO] {line}")
        done_fn(False, None)


# =========================================================
#  INTERFACE GRAFICA
# =========================================================

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Lançador de Validades")
        self.root.geometry("620x580")
        self.root.minsize(580, 520)
        self.root.configure(bg=BG)
        self.root.resizable(True, True)
        self._output_folder = None

        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth()  - 620) // 2
        y = (self.root.winfo_screenheight() - 580) // 2
        self.root.geometry(f"620x580+{x}+{y}")

        self._style()
        self._build()

    def _style(self):
        s = ttk.Style()
        s.theme_use("clam")
        s.configure("TFrame",      background=BG)
        s.configure("TLabel",      background=BG,  foreground=TEXT,  font=("Segoe UI", 10))
        s.configure("Ghost.TButton",
                    background=BG3, foreground=TEXT2,
                    font=("Segoe UI", 9), borderwidth=0, relief="flat", padding=(10, 5))
        s.map("Ghost.TButton",
              background=[("active", BORDER)],
              foreground=[("active", TEXT)])
        s.configure("green.Horizontal.TProgressbar",
                    troughcolor=BG3, background=ACCENT,
                    borderwidth=0, thickness=6)

    def _build(self):
        # -- Header
        hdr = tk.Frame(self.root, bg=BG)
        hdr.pack(fill="x", padx=24, pady=(20, 0))

        tk.Label(hdr, text="\U0001f4e6", font=("Segoe UI", 22, "bold"),
                 bg=BG, fg=ACCENT).pack(side="left", padx=(0, 10))

        info = tk.Frame(hdr, bg=BG)
        info.pack(side="left")
        tk.Label(info, text="Lançador de Validades",
                 font=("Segoe UI", 16, "bold"), bg=BG, fg=TEXT).pack(anchor="w")
        tk.Label(info, text="Atualiza a planilha de controle de validades",
                 font=("Segoe UI", 9), bg=BG, fg=TEXT2).pack(anchor="w")

        tk.Frame(self.root, bg=BORDER, height=1).pack(fill="x", padx=24, pady=(14, 0))

        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, padx=24, pady=12)

        self._card_csv(body)
        self._card_xlsx(body)
        self._buttons(body)
        self._log_area(body)

    def _card_csv(self, parent):
        card = tk.Frame(parent, bg=BG2, highlightthickness=1,
                        highlightbackground=BORDER)
        card.pack(fill="x", pady=(0, 10))
        inner = tk.Frame(card, bg=BG2, padx=14, pady=12)
        inner.pack(fill="x")

        top = tk.Frame(inner, bg=BG2)
        top.pack(fill="x", pady=(0, 6))
        tk.Label(top, text="Arquivo(s) TXT  ", font=("Segoe UI", 10, "bold"),
                 bg=BG2, fg=TEXT).pack(side="left")
        tk.Label(top, text="obrigatorio", font=("Segoe UI", 8),
                 bg=BG2, fg=ACCENT).pack(side="left")

        row = tk.Frame(inner, bg=BG2)
        row.pack(fill="x")

        # Listbox para mostrar os arquivos selecionados
        lista_frame = tk.Frame(row, bg=BG3, highlightthickness=1,
                               highlightbackground=BORDER)
        lista_frame.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.csv_listbox = tk.Listbox(lista_frame, bg=BG3, fg=TEXT,
                                      font=("Segoe UI", 9), relief="flat",
                                      selectbackground=ACCENT2, height=3,
                                      activestyle="none", borderwidth=0)
        self.csv_listbox.pack(fill="x", padx=4, pady=4)

        btn_col = tk.Frame(row, bg=BG2)
        btn_col.pack(side="left")
        ttk.Button(btn_col, text="Selecionar", style="Ghost.TButton",
                   command=self._sel_csv).pack(pady=(0, 4))
        ttk.Button(btn_col, text="Limpar", style="Ghost.TButton",
                   command=self._limpar_csv).pack()

        tk.Label(inner, text="Selecione um ou mais arquivos .txt do coletor",
                 font=("Segoe UI", 8), bg=BG2, fg=TEXT2).pack(anchor="w", pady=(4, 0))

    def _card_xlsx(self, parent):
        card = tk.Frame(parent, bg=BG2, highlightthickness=1,
                        highlightbackground=BORDER)
        card.pack(fill="x", pady=(0, 14))
        inner = tk.Frame(card, bg=BG2, padx=14, pady=12)
        inner.pack(fill="x")

        top = tk.Frame(inner, bg=BG2)
        top.pack(fill="x", pady=(0, 6))
        tk.Label(top, text="Planilha de Validades  ", font=("Segoe UI", 10, "bold"),
                 bg=BG2, fg=TEXT).pack(side="left")
        tk.Label(top, text="opcional", font=("Segoe UI", 8),
                 bg=BG2, fg=TEXT2).pack(side="left")

        row = tk.Frame(inner, bg=BG2)
        row.pack(fill="x")
        self.xlsx_var = tk.StringVar()
        xlsx_entry = tk.Entry(row, textvariable=self.xlsx_var, bg=BG3, fg=TEXT,
                 insertbackground=TEXT, relief="flat", font=("Segoe UI", 9),
                 highlightthickness=1, highlightbackground=BORDER,
                 highlightcolor=ACCENT)
        xlsx_entry.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 8))
        xlsx_btn = ttk.Button(row, text="Selecionar", style="Ghost.TButton",
                   command=self._sel_xlsx)
        xlsx_btn.pack(side="left")

        tk.Label(inner,
                 text="Selecione a planilha apenas se quiser atualizar as validades",
                 font=("Segoe UI", 8), bg=BG2, fg=TEXT2).pack(anchor="w", pady=(4, 0))

    def _buttons(self, parent):
        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", pady=(0, 10))

        self.run_btn = tk.Button(row, text=">  Atualizar Planilha",
                                 font=("Segoe UI", 10, "bold"),
                                 bg=ACCENT, fg="white",
                                 activebackground=ACCENT2, activeforeground="white",
                                 relief="flat", cursor="hand2",
                                 bd=0, padx=20, pady=8,
                                 command=self._run)
        self.run_btn.pack(side="left")

        self.open_btn = tk.Button(row, text="[DIR]  Abrir Pasta",
                                  font=("Segoe UI", 9),
                                  bg=BG3, fg=TEXT2,
                                  activebackground=BORDER, activeforeground=TEXT,
                                  relief="flat", cursor="hand2",
                                  bd=0, padx=14, pady=8,
                                  state="disabled",
                                  command=self._open_folder)
        self.open_btn.pack(side="left", padx=(10, 0))

        self.status_lbl = tk.Label(row, text="", font=("Segoe UI", 9),
                                   bg=BG, fg=TEXT2)
        self.status_lbl.pack(side="right")

    def _log_area(self, parent):
        tk.Label(parent, text="Log de execucao", font=("Segoe UI", 9, "bold"),
                 bg=BG, fg=TEXT2).pack(anchor="w", pady=(0, 4))

        self.progress = ttk.Progressbar(parent,
                                        style="green.Horizontal.TProgressbar",
                                        mode="indeterminate")
        self.progress.pack(fill="x", pady=(0, 6))

        box = tk.Frame(parent, bg="#0D1117", highlightthickness=1,
                       highlightbackground=BORDER)
        box.pack(fill="both", expand=True)

        self.log = tk.Text(box, bg="#0D1117", fg="#C9D1D9",
                           font=("Consolas", 9), relief="flat",
                           state="disabled", wrap="word",
                           insertbackground=TEXT, padx=10, pady=8,
                           selectbackground=ACCENT2)
        self.log.pack(side="left", fill="both", expand=True)

        sb = ttk.Scrollbar(box, command=self.log.yview)
        sb.pack(side="right", fill="y")
        self.log.configure(yscrollcommand=sb.set)

        self.log.tag_config("ok",   foreground=SUCCESS)
        self.log.tag_config("erro", foreground=ERROR)
        self.log.tag_config("info", foreground=ACCENT)

    # -- acoes
    def _sel_csv(self):
        arquivos = filedialog.askopenfilenames(
            title="Selecionar arquivo(s) TXT",
            initialdir=PASTA_TXT,
            filetypes=[("TXT", "*.txt"), ("Todos", "*.*")])
        if arquivos:
            self.csv_listbox.delete(0, "end")
            for a in arquivos:
                self.csv_listbox.insert("end", a)

    def _limpar_csv(self):
        self.csv_listbox.delete(0, "end")

    def _sel_xlsx(self):
        p = filedialog.askopenfilename(
            title="Selecionar planilha de validades",
            initialdir=PASTA_XLSX,
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")])
        if p:
            self.xlsx_var.set(p)

    def _open_folder(self):
        arquivos = list(self.csv_listbox.get(0, "end"))
        pasta = str(Path(arquivos[0]).parent) if arquivos else self._output_folder
        if pasta and os.path.exists(pasta):
            os.startfile(pasta)

    def _append_log(self, msg):
        self.log.configure(state="normal")
        tag = "ok" if "[OK]" in msg else "erro" if "[ERRO]" in msg else "info"
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _set_running(self, state):
        if state:
            self.run_btn.configure(state="disabled", text="  Processando...",
                                   bg=BG3, fg=TEXT2)
            self.open_btn.configure(state="disabled")
            self.status_lbl.configure(text="Aguarde...", fg=TEXT2)
            self.progress.start(12)
        else:
            self.run_btn.configure(state="normal", text=">  Atualizar Planilha",
                                   bg=ACCENT, fg="white")
            self.progress.stop()
            self.progress["value"] = 0

    def _on_done(self, ok, folder):
        self._set_running(False)
        if ok:
            self._output_folder = folder
            self.open_btn.configure(state="normal")
            self.status_lbl.configure(text="Concluido!", fg=SUCCESS)
        else:
            self.status_lbl.configure(text="Erro!", fg=ERROR)

    def _run(self):
        csv_paths = list(self.csv_listbox.get(0, "end"))
        xlsx_path = self.xlsx_var.get().strip() or None

        if not csv_paths:
            messagebox.showwarning("Atencao", "Selecione ao menos um arquivo TXT antes de continuar.")
            return
        for cp in csv_paths:
            if not os.path.exists(cp):
                messagebox.showerror("Erro", f"Arquivo nao encontrado:\n{cp}")
                return
        if xlsx_path and not os.path.exists(xlsx_path):
            messagebox.showerror("Erro", f"Planilha nao encontrada:\n{xlsx_path}")
            return

        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self._set_running(True)

        threading.Thread(
            target=executar_conversao,
            args=(
                csv_paths,
                xlsx_path,
                lambda m: self.root.after(0, self._append_log, m),
                lambda ok, f: self.root.after(0, self._on_done, ok, f),
            ),
            daemon=True,
        ).start()


if __name__ == "__main__":
    root = tk.Tk()
    # Define o icone da janela (mesmo diretorio do executavel/script)
    try:
        ico_path = Path(os.path.abspath(__file__)).parent / "lancador_validades.ico"
        if ico_path.exists():
            root.iconbitmap(str(ico_path))
    except Exception:
        pass
    App(root)
    root.mainloop()
