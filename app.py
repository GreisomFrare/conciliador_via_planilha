import calendar
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import date, datetime

from config_manager import load_config, save_config
from database import DatabaseManager
from excel_importer import import_excel


def _fmt(value):
    if value is None:
        return ''
    if isinstance(value, (date, datetime)):
        return value.strftime('%d/%m/%Y')
    return str(value)


def _fmt_num(value):
    if value is None or value == '':
        return ''
    try:
        return f'{float(value):,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    except (ValueError, TypeError):
        return str(value)


def _br_to_float(s) -> float:
    """Parse número formatado em BR (1.234,56) de volta para float."""
    try:
        return float(str(s).replace('.', '').replace(',', '.'))
    except (ValueError, TypeError):
        return 0.0


def _sort_key(val: str):
    s = str(val).strip()
    try:
        return (0, float(s.replace('.', '').replace(',', '.')))
    except ValueError:
        pass
    try:
        return (0, datetime.strptime(s, '%d/%m/%Y').timestamp())
    except ValueError:
        pass
    return (1, s.lower())


# ---------------------------------------------------------------------------
# Calendar popup
# ---------------------------------------------------------------------------

class _CalendarPopup(tk.Toplevel):
    _MONTHS = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
               'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
    _DAYS_HDR = ('D', 'S', 'T', 'Q', 'Q', 'S', 'S')

    def __init__(self, anchor_widget, current_date: date, callback):
        super().__init__(anchor_widget)
        self.overrideredirect(True)
        self.resizable(False, False)
        self.grab_set()
        self.focus_set()
        self._callback = callback
        self._year = current_date.year
        self._month = current_date.month
        self._selected = current_date
        x = anchor_widget.winfo_rootx()
        y = anchor_widget.winfo_rooty() + anchor_widget.winfo_height() + 2
        self.geometry(f'+{x}+{y}')
        self._build()
        self.bind('<Escape>', lambda _: self.destroy())

    def _build(self):
        outer = tk.Frame(self, bg='#aaaaaa', bd=1)
        outer.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        nav = tk.Frame(outer, bg='#f0f0f0')
        nav.pack(fill=tk.X)
        tk.Button(nav, text='◀', relief='flat', bg='#f0f0f0',
                  command=self._prev).pack(side=tk.LEFT, padx=2, pady=2)
        self._title_var = tk.StringVar()
        tk.Label(nav, textvariable=self._title_var, bg='#f0f0f0',
                 font=('Arial', 9, 'bold'), width=18, anchor='center').pack(
            side=tk.LEFT, expand=True)
        tk.Button(nav, text='▶', relief='flat', bg='#f0f0f0',
                  command=self._next).pack(side=tk.RIGHT, padx=2, pady=2)

        cal_frm = tk.Frame(outer, bg='white')
        cal_frm.pack(padx=4, pady=2)
        for col, h in enumerate(self._DAYS_HDR):
            tk.Label(cal_frm, text=h, width=3, anchor='center', bg='white',
                     font=('Arial', 8, 'bold'), fg='#555').grid(row=0, column=col, padx=1)

        self._btns: list[list[tk.Button]] = []
        for r in range(6):
            row_btns = []
            for c in range(7):
                b = tk.Button(cal_frm, text='', width=3, height=1,
                              relief='flat', font=('Arial', 8), bd=0,
                              bg='white', activebackground='#93c5fd')
                b.grid(row=r + 1, column=c, padx=1, pady=1)
                row_btns.append(b)
            self._btns.append(row_btns)

        tk.Button(outer, text='Hoje', relief='flat', bg='#e8e8e8',
                  font=('Arial', 8), command=self._pick_today).pack(pady=3)

        self._render()

    def _render(self):
        self._title_var.set(f'{self._MONTHS[self._month - 1]} {self._year}')
        cal = calendar.Calendar(firstweekday=6)
        weeks = cal.monthdayscalendar(self._year, self._month)
        all_days: list[int] = []
        for week in weeks:
            all_days.extend(week)
        while len(all_days) < 42:
            all_days.append(0)
        today = date.today()
        for r in range(6):
            for c in range(7):
                day = all_days[r * 7 + c]
                btn = self._btns[r][c]
                if day == 0:
                    btn.configure(text='', state='disabled', bg='#fafafa',
                                  fg='#ccc', cursor='', command=lambda: None)
                else:
                    d = date(self._year, self._month, day)
                    is_sel = (d == self._selected)
                    is_today = (d == today)
                    if is_sel:
                        bg, fg = '#2563eb', 'white'
                    elif is_today:
                        bg, fg = '#dbeafe', '#1d4ed8'
                    else:
                        bg, fg = 'white', '#111'
                    btn.configure(text=str(day), state='normal',
                                  bg=bg, fg=fg, cursor='hand2',
                                  command=lambda d=d: self._pick(d))

    def _prev(self):
        if self._month == 1:
            self._month, self._year = 12, self._year - 1
        else:
            self._month -= 1
        self._render()

    def _next(self):
        if self._month == 12:
            self._month, self._year = 1, self._year + 1
        else:
            self._month += 1
        self._render()

    def _pick_today(self):
        self._pick(date.today())

    def _pick(self, d: date):
        self._callback(d)
        self.destroy()


# ---------------------------------------------------------------------------
# Date entry with calendar picker
# ---------------------------------------------------------------------------

class _DateEntry(ttk.Frame):
    def __init__(self, parent, **kw):
        super().__init__(parent)
        self._var = tk.StringVar()
        self._entry = ttk.Entry(self, textvariable=self._var, width=10, **kw)
        self._entry.pack(side=tk.LEFT)
        tk.Button(self, text='📅', relief='flat', bd=0,
                  font=('Arial', 9), cursor='hand2', padx=1,
                  command=self._open_calendar).pack(side=tk.LEFT, padx=1)
        self._entry.bind('<FocusOut>', self._on_focus_out)

    def _open_calendar(self):
        _CalendarPopup(self, self.get_date() or date.today(), self._on_pick)

    def _on_pick(self, d: date):
        self._var.set(d.strftime('%d/%m/%Y'))
        self._entry.configure(style='TEntry')

    def _on_focus_out(self, _=None):
        val = self._var.get().strip()
        if val:
            try:
                datetime.strptime(val, '%d/%m/%Y')
                self._entry.configure(style='TEntry')
            except ValueError:
                self._entry.configure(style='Error.TEntry')

    def get_date(self):
        try:
            return datetime.strptime(self._var.get().strip(), '%d/%m/%Y').date()
        except ValueError:
            return None

    def get_str(self) -> str:
        return self._var.get().strip()

    def set(self, val):
        self._var.set(val)

    def get(self):
        return self._var.get()


# ---------------------------------------------------------------------------
# Main application
# ---------------------------------------------------------------------------

class ConciliacaoCartaoApp:

    # DB row order from get_records_by_arquivo:
    # [0]CONCILIADO [1]DATA_VENDA [2]ESTABELECIMENTO [3]NSU_DOC [4]AUTORIZACAO
    # [5]VALOR_BRUTO [6]TAXA_TARIFA [7]VALOR_LIQUIDO [8]BANDEIRA
    # [9]NUMERO_PARCELA [10]TOTAL_PARCELAS [11]FORMA_PAGAMENTO
    # [12]STATUS_PAGAMENTO [13]DATA_PAGAMENTO

    _COLS_ESQ = ('conciliado', 'data_venda', 'estabelecimento', 'estab_erp',
                 'nsu_doc', 'autorizacao', 'valor_bruto', 'valor_liquido',
                 'bandeira', 'parcela', 'total_parc', 'forma_pag', 'status_pag')
    _HDRS_ESQ = ('OK?', 'Data Venda', 'Estabelecimento', 'Estab ERP',
                 'NSU/DOC', 'Autorização', 'Valor Bruto', 'Valor Líquido',
                 'Bandeira', 'Parc.', 'Total Parc.', 'Forma Pgto', 'Status')
    _WIDS_ESQ = (40, 85, 105, 68, 72, 80, 85, 85, 65, 42, 62, 110, 110)

    _COLS_DIR = ('existenoarq', 'estab', 'nsudoc', 'autoriza',
                 'idcartao', 'parcela', 'parcelas',
                 'vlrparc', 'vlrtaxa', 'vlrl', 'taxa',
                 'dtvencto', 'datacredito')
    _HDRS_DIR = ('OK?', 'Estab.', 'NSU/DOC', 'Autorização',
                 'Cartão', 'Parc.', 'Total Parc.',
                 'Vlr Parcela', 'Taxa', 'Vlr Líquido', '% Taxa',
                 'Dt. Vencto', 'Dt. Crédito')
    _WIDS_DIR = (38, 50, 72, 80, 120, 42, 62, 85, 72, 85, 52, 82, 82)

    _SHOW_DIR = [0, 0, 3, 4, 5, 8, 6, 10, 11, 12, 13, 14, 16]
    _NUM_DIR  = {10, 11, 12, 13}

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title('Conciliação de Cartão')
        self.root.geometry('1500x820')
        self.root.minsize(1200, 640)

        self.db = DatabaseManager()
        self.current_file: str | None = None
        self.current_id_arquivo: str | None = None

        self._all_items_esq: list = []
        self._all_items_dir: list = []
        self._sort_esq: dict = {}
        self._sort_dir: dict = {}

        self._setup_style()
        self._build_ui()

    def _setup_style(self):
        s = ttk.Style()
        s.theme_use('clam')
        s.configure('TNotebook.Tab', padding=[12, 5], font=('Arial', 9))
        s.configure('TLabelframe.Label', font=('Arial', 9, 'bold'))
        s.configure('Error.TEntry', fieldbackground='#FFD0D0')
        s.configure('Treeview', rowheight=22, font=('Arial', 8))
        s.configure('Treeview.Heading', font=('Arial', 8, 'bold'))

    def _build_ui(self):
        nb = ttk.Notebook(self.root)
        nb.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        self.tab_principal = ttk.Frame(nb)
        nb.add(self.tab_principal, text='1 - Principal')
        self._build_principal(self.tab_principal)

        self.tab_depara = ttk.Frame(nb)
        nb.add(self.tab_depara, text='2 - De-Para Estabelecimento')
        self._build_de_para(self.tab_depara)

        self.tab_config = ttk.Frame(nb)
        nb.add(self.tab_config, text='3 - Configuração Oracle')
        self._build_config(self.tab_config)

    # ---- Principal tab -----------------------------------------------

    def _build_principal(self, parent):
        frm_arq = ttk.LabelFrame(parent, text='Arquivo')
        frm_arq.pack(fill=tk.X, padx=5, pady=(4, 2))

        row1 = ttk.Frame(frm_arq)
        row1.pack(fill=tk.X, padx=5, pady=3)
        ttk.Label(row1, text='Arquivo:').pack(side=tk.LEFT)
        self._file_var = tk.StringVar()
        ttk.Entry(row1, textvariable=self._file_var, width=85,
                  state='readonly').pack(side=tk.LEFT, padx=4)
        ttk.Button(row1, text='Selecionar...', command=self._cmd_selecionar).pack(side=tk.LEFT, padx=2)
        ttk.Button(row1, text='Importar Planilha', command=self._cmd_importar).pack(side=tk.LEFT, padx=2)
        ttk.Button(row1, text='Carregar Grid', command=self._cmd_carregar_grid).pack(side=tk.LEFT, padx=2)

        row2 = ttk.Frame(frm_arq)
        row2.pack(fill=tk.X, padx=5, pady=3)
        ttk.Label(row2, text='Data Inicial:').pack(side=tk.LEFT)
        self._dt_ini = _DateEntry(row2)
        self._dt_ini.pack(side=tk.LEFT, padx=(2, 10))
        ttk.Label(row2, text='Data Final:').pack(side=tk.LEFT)
        self._dt_fim = _DateEntry(row2)
        self._dt_fim.pack(side=tk.LEFT, padx=(2, 10))

        row3 = ttk.Frame(frm_arq)
        row3.pack(fill=tk.X, padx=5, pady=3)
        ttk.Button(row3, text='Buscar Lançamentos', command=self._cmd_buscar).pack(side=tk.LEFT, padx=2)
        ttk.Button(row3, text='Conciliar', command=self._cmd_conciliar).pack(side=tk.LEFT, padx=2)
        ttk.Button(row3, text='Desfazer Conciliação', command=self._cmd_desfazer).pack(side=tk.LEFT, padx=2)
        ttk.Button(row3, text='🔍 Debug SQL', command=self._cmd_debug_sql).pack(side=tk.LEFT, padx=8)

        today = date.today().strftime('%d/%m/%Y')
        self._dt_ini.set(today)
        self._dt_fim.set(today)

        frm_crit = ttk.LabelFrame(parent, text='Critérios de Conciliação')
        frm_crit.pack(fill=tk.X, padx=5, pady=2)
        crit_row = ttk.Frame(frm_crit)
        crit_row.pack(fill=tk.X, padx=8, pady=4)
        self._crit_nsu        = tk.BooleanVar(value=True)
        self._crit_aut        = tk.BooleanVar(value=False)
        self._crit_data_venda = tk.BooleanVar(value=True)
        self._crit_valor      = tk.BooleanVar(value=False)
        self._crit_estab      = tk.BooleanVar(value=False)
        self._crit_parcela    = tk.BooleanVar(value=False)
        ttk.Checkbutton(crit_row, text='NSU/Docto',
                        variable=self._crit_nsu).pack(side=tk.LEFT, padx=10)
        ttk.Checkbutton(crit_row, text='Autorização',
                        variable=self._crit_aut).pack(side=tk.LEFT, padx=10)
        ttk.Checkbutton(crit_row, text='Dt. Pagamento = Dt. Crédito',
                        variable=self._crit_data_venda).pack(side=tk.LEFT, padx=10)
        ttk.Checkbutton(crit_row, text='Valor Bruto = Vlr Parcela',
                        variable=self._crit_valor).pack(side=tk.LEFT, padx=10)
        ttk.Checkbutton(crit_row, text='Estab ERP = Estab.',
                        variable=self._crit_estab).pack(side=tk.LEFT, padx=10)
        ttk.Checkbutton(crit_row, text='Parcela',
                        variable=self._crit_parcela).pack(side=tk.LEFT, padx=10)
        ttk.Separator(crit_row, orient='vertical').pack(side=tk.LEFT, fill=tk.Y, padx=10)
        ttk.Label(crit_row, text='Variação (centavos):').pack(side=tk.LEFT)
        self._crit_variacao = tk.StringVar(value='0')
        ttk.Spinbox(crit_row, textvariable=self._crit_variacao,
                    from_=0, to=9999, increment=1, width=6,
                    justify='center').pack(side=tk.LEFT, padx=4)

        frm_grids = ttk.Frame(parent)
        frm_grids.pack(fill=tk.BOTH, expand=True, padx=5, pady=2)
        frm_grids.columnconfigure(0, weight=50)
        frm_grids.columnconfigure(1, weight=50)
        frm_grids.rowconfigure(0, weight=1)

        lf_esq = ttk.LabelFrame(frm_grids, text='Dados do Arquivo (Cielo)')
        lf_esq.grid(row=0, column=0, sticky='nsew', padx=(0, 2))
        self._tree_esq = self._build_grid(
            lf_esq,
            self._COLS_ESQ, self._HDRS_ESQ, self._WIDS_ESQ,
            [('conciliado', '#90EE90'), ('pendente', '#FFFACD')],
            '_esq',
        )

        lf_dir = ttk.LabelFrame(frm_grids, text='Lançamentos ERP (ACERCAR)')
        lf_dir.grid(row=0, column=1, sticky='nsew', padx=(2, 0))
        self._tree_dir = self._build_grid(
            lf_dir,
            self._COLS_DIR, self._HDRS_DIR, self._WIDS_DIR,
            [('conciliado', '#90EE90'),
             ('pendente',   '#ADD8E6'),
             ('nomatch',    '#FFFACD')],
            '_dir',
        )

        frm_tot = ttk.Frame(parent)
        frm_tot.pack(fill=tk.X, padx=5, pady=2)

        def _tlbl(p, text, var, bg='#e8e8e8', width=9):
            ttk.Label(p, text=text).pack(side=tk.LEFT)
            ttk.Label(p, textvariable=var, width=width, anchor='e',
                      background=bg, relief=tk.SUNKEN, padding=(2, 0)).pack(
                side=tk.LEFT, padx=(2, 8))

        self._tot_total          = tk.StringVar(value='0')
        self._tot_conc           = tk.StringVar(value='0')
        self._tot_vlr_conc_esq   = tk.StringVar(value='0,00')
        self._tot_pend           = tk.StringVar(value='0')
        self._tot_erp            = tk.StringVar(value='0')
        self._tot_erp_conc       = tk.StringVar(value='0')
        self._tot_vlr_conc_dir   = tk.StringVar(value='0,00')
        self._tot_erp_match      = tk.StringVar(value='0')
        self._tot_erp_nomatch    = tk.StringVar(value='0')

        _tlbl(frm_tot, 'Arquivo — Total:', self._tot_total)
        _tlbl(frm_tot, 'Conciliados:', self._tot_conc, '#90EE90')
        _tlbl(frm_tot, 'Vlr. Bruto Conc.:', self._tot_vlr_conc_esq, '#90EE90', width=13)
        _tlbl(frm_tot, 'Pendentes:', self._tot_pend, '#FFFACD')
        ttk.Separator(frm_tot, orient='vertical').pack(side=tk.LEFT, fill=tk.Y, padx=6)
        _tlbl(frm_tot, 'ERP — Total:', self._tot_erp)
        _tlbl(frm_tot, 'Conciliados:', self._tot_erp_conc, '#90EE90')
        _tlbl(frm_tot, 'Vlr. Parc. Conc.:', self._tot_vlr_conc_dir, '#90EE90', width=13)
        _tlbl(frm_tot, 'Match pend.:', self._tot_erp_match, '#ADD8E6')
        _tlbl(frm_tot, 'Sem match:', self._tot_erp_nomatch, '#FFFACD')

        self._status_var = tk.StringVar(value='Pronto')
        ttk.Label(parent, textvariable=self._status_var,
                  relief=tk.SUNKEN, anchor='w').pack(
            fill=tk.X, side=tk.BOTTOM, padx=5, pady=(0, 2))

    # ------------------------------------------------------------------
    # Generic grid builder
    # ------------------------------------------------------------------

    def _build_grid(self, parent, cols, hdrs, wids, tags, side: str) -> ttk.Treeview:
        hdr = ttk.Frame(parent)
        hdr.pack(fill=tk.X, padx=4, pady=(3, 0))
        ttk.Label(hdr, text='Filtros por coluna:', font=('Arial', 8)).pack(side=tk.LEFT)
        ttk.Button(hdr, text='✕ Limpar', width=8,
                   command=lambda s=side: self._clear_filters(s)).pack(side=tk.RIGHT, padx=2)

        frm = ttk.Frame(parent)
        frm.pack(fill=tk.BOTH, expand=True)
        frm.rowconfigure(1, weight=1)
        frm.columnconfigure(0, weight=1)

        # Per-column filter strip synchronized with horizontal scroll
        bg = ttk.Style().lookup('TFrame', 'background') or '#f0f0f0'
        f_canvas = tk.Canvas(frm, height=24, highlightthickness=0, bg=bg, bd=0)
        f_canvas.grid(row=0, column=0, sticky='ew')
        f_frame = tk.Frame(f_canvas, bg=bg)
        f_canvas.create_window((0, 0), window=f_frame, anchor='nw')

        filter_vars: dict = {}
        setattr(self, f'_filter_vars{side}', filter_vars)

        for col, wid in zip(cols, wids):
            fvar = tk.StringVar()
            filter_vars[col] = fvar
            cell = tk.Frame(f_frame, width=wid, height=22, bg=bg)
            cell.pack_propagate(False)
            cell.pack(side=tk.LEFT)
            ttk.Entry(cell, textvariable=fvar, font=('Arial', 8)).pack(fill=tk.BOTH, expand=True)
            fvar.trace_add('write', lambda *_, s=side: self._refresh(s))

        f_canvas.configure(scrollregion=(0, 0, sum(wids), 24))
        f_frame.bind('<Configure>',
                     lambda e, c=f_canvas: c.configure(scrollregion=c.bbox('all')))

        tv = ttk.Treeview(frm, columns=cols, show='headings', selectmode='browse')
        for c, h, w in zip(cols, hdrs, wids):
            tv.heading(c, text=h,
                       command=lambda c=c, h=h: self._sort_click(side, c, h))
            tv.column(c, width=w, minwidth=35,
                      anchor='center' if w <= 62 else 'w')

        for tag, color in tags:
            tv.tag_configure(tag, background=color)

        vsb = ttk.Scrollbar(frm, orient='vertical', command=tv.yview)
        hsb = ttk.Scrollbar(frm, orient='horizontal',
                            command=lambda *a, _tv=tv, _fc=f_canvas: (
                                _tv.xview(*a), _fc.xview(*a)))

        def _xscrollset(first, last, _hsb=hsb, _fc=f_canvas):
            _fc.xview_moveto(first)
            _hsb.set(first, last)

        tv.configure(yscrollcommand=vsb.set, xscrollcommand=_xscrollset)
        tv.grid(row=1, column=0, sticky='nsew')
        vsb.grid(row=1, column=1, sticky='ns')
        hsb.grid(row=2, column=0, sticky='ew')

        setattr(self, f'_vsb{side}', vsb)
        setattr(self, f'_hsb{side}', hsb)
        setattr(self, f'_xscrollset{side}', _xscrollset)
        return tv

    # ------------------------------------------------------------------
    # Sort / filter
    # ------------------------------------------------------------------

    def _sort_click(self, side: str, col: str, _original_header: str):
        sort = getattr(self, f'_sort{side}')
        if sort.get('col') == col:
            sort['asc'] = not sort.get('asc', True)
        else:
            sort['col'] = col
            sort['asc'] = True
        self._refresh(side)

    def _refresh(self, side: str):
        tree   = self._tree_esq  if side == '_esq' else self._tree_dir
        vsb    = self._vsb_esq   if side == '_esq' else self._vsb_dir
        xss    = getattr(self, f'_xscrollset{side}')
        hdrs   = self._HDRS_ESQ  if side == '_esq' else self._HDRS_DIR
        cols   = self._COLS_ESQ  if side == '_esq' else self._COLS_DIR
        master = self._all_items_esq if side == '_esq' else self._all_items_dir
        fvars  = getattr(self, f'_filter_vars{side}')
        sort   = getattr(self, f'_sort{side}')

        col_list = list(cols)
        active = [(col_list.index(c), ft)
                  for c, fv in fvars.items()
                  if (ft := fv.get().strip().lower())]

        if active:
            items = [(v, t) for v, t in master
                     if all(ft in str(v[i]).lower() for i, ft in active)]
        else:
            items = list(master)

        sort_col = sort.get('col')
        if sort_col and sort_col in cols:
            col_idx = list(cols).index(sort_col)
            items.sort(key=lambda it: _sort_key(it[0][col_idx]),
                       reverse=not sort.get('asc', True))

        for c, h in zip(cols, hdrs):
            arrow = (' ▲' if sort.get('asc', True) else ' ▼') if c == sort_col else ''
            tree.heading(c, text=h + arrow)

        self._bulk_fill(tree, vsb, xss, items)

    def _bulk_fill(self, tree, vsb, xscrollset, items: list):
        tree.configure(yscrollcommand='', xscrollcommand='')
        tree.delete(*tree.get_children())
        try:
            for vals, tag in items:
                tree.insert('', 'end', values=vals, tags=(tag,))
        finally:
            tree.configure(yscrollcommand=vsb.set, xscrollcommand=xscrollset)

    def _clear_filters(self, side: str):
        for fvar in getattr(self, f'_filter_vars{side}').values():
            fvar.set('')

    def _load_esq(self, items: list):
        self._all_items_esq = items
        self._sort_esq = {}
        self._refresh('_esq')

    def _load_dir(self, items: list):
        self._all_items_dir = items
        self._sort_dir = {}
        self._refresh('_dir')

    # ---- De-Para tab -------------------------------------------------

    def _build_de_para(self, parent):
        outer = ttk.Frame(parent)
        outer.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        ttk.Label(
            outer,
            text='Mapeamento entre código de Estabelecimento da Cielo e código do ERP (ACERCAR.ESTAB)',
            font=('Arial', 9),
        ).pack(anchor='w', pady=(0, 6))

        tv_frm = ttk.Frame(outer)
        tv_frm.pack(fill=tk.BOTH, expand=True)
        tv_frm.rowconfigure(0, weight=1)
        tv_frm.columnconfigure(0, weight=1)

        self._tv_depara = ttk.Treeview(tv_frm, columns=('cielo', 'erp'),
                                        show='headings', height=12)
        self._tv_depara.heading('cielo', text='Estab. Cielo')
        self._tv_depara.heading('erp',   text='Estab. ERP')
        self._tv_depara.column('cielo', width=200, anchor='center')
        self._tv_depara.column('erp',   width=120, anchor='center')

        vsb = ttk.Scrollbar(tv_frm, orient='vertical', command=self._tv_depara.yview)
        self._tv_depara.configure(yscrollcommand=vsb.set)
        self._tv_depara.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')

        add_frm = ttk.LabelFrame(outer, text='Adicionar / Editar')
        add_frm.pack(fill=tk.X, pady=(8, 4))
        af_row = ttk.Frame(add_frm)
        af_row.pack(padx=8, pady=6)

        ttk.Label(af_row, text='Estab. Cielo:').pack(side=tk.LEFT)
        self._dp_cielo_var = tk.StringVar()
        ttk.Entry(af_row, textvariable=self._dp_cielo_var, width=18).pack(side=tk.LEFT, padx=(4, 16))
        ttk.Label(af_row, text='Estab. ERP:').pack(side=tk.LEFT)
        self._dp_erp_var = tk.StringVar()
        ttk.Entry(af_row, textvariable=self._dp_erp_var, width=10).pack(side=tk.LEFT, padx=4)
        ttk.Button(af_row, text='Adicionar', command=self._dp_add).pack(side=tk.LEFT, padx=(16, 4))
        ttk.Button(af_row, text='Remover Selecionado', command=self._dp_remove).pack(side=tk.LEFT, padx=4)

        btn_frm = ttk.Frame(outer)
        btn_frm.pack(pady=6, anchor='w')
        ttk.Button(btn_frm, text='Salvar De-Para', command=self._dp_save).pack(side=tk.LEFT, padx=8)
        self._dp_status = ttk.Label(btn_frm, text='', foreground='green')
        self._dp_status.pack(side=tk.LEFT, padx=8)

        self._tv_depara.bind('<<TreeviewSelect>>', self._dp_on_select)
        self._dp_load()

    def _dp_load(self):
        cfg = load_config()
        self._tv_depara.delete(*self._tv_depara.get_children())
        for cielo, erp in cfg.get('de_para', []):
            self._tv_depara.insert('', 'end', values=(cielo, erp))

    def _dp_on_select(self, _=None):
        sel = self._tv_depara.selection()
        if sel:
            vals = self._tv_depara.item(sel[0], 'values')
            self._dp_cielo_var.set(vals[0])
            self._dp_erp_var.set(vals[1])

    def _dp_add(self):
        cielo = self._dp_cielo_var.get().strip()
        erp   = self._dp_erp_var.get().strip()
        if not cielo or not erp:
            messagebox.showwarning('Aviso', 'Preencha Estab. Cielo e Estab. ERP.', parent=self.root)
            return
        for item in self._tv_depara.get_children():
            if self._tv_depara.item(item, 'values')[0] == cielo:
                self._tv_depara.delete(item)
                break
        self._tv_depara.insert('', 'end', values=(cielo, erp))
        self._dp_cielo_var.set('')
        self._dp_erp_var.set('')

    def _dp_remove(self):
        sel = self._tv_depara.selection()
        if sel:
            self._tv_depara.delete(sel[0])
            self._dp_cielo_var.set('')
            self._dp_erp_var.set('')

    def _dp_save(self):
        pairs = []
        for item in self._tv_depara.get_children():
            vals = self._tv_depara.item(item, 'values')
            pairs.append([vals[0], vals[1]])
        cfg = load_config()
        cfg['de_para'] = pairs
        save_config(cfg)
        self._dp_status.configure(text='✔ De-Para salvo!', foreground='green')
        self.root.after(3000, lambda: self._dp_status.configure(text=''))

    def _get_de_para(self) -> dict:
        cfg = load_config()
        return {str(c): str(e) for c, e in cfg.get('de_para', []) if c and e}

    # ---- Config tab --------------------------------------------------

    def _build_config(self, parent):
        outer = ttk.Frame(parent)
        outer.pack(padx=30, pady=20, fill=tk.X)

        frm = ttk.LabelFrame(outer, text='Conexão Oracle — TNS (Thick Mode)')
        frm.pack(fill=tk.X)
        frm.columnconfigure(1, weight=1)

        self._cfg_vars: dict[str, tk.StringVar] = {}

        def _row(i, label, key, secret=False, width=38, tip=''):
            ttk.Label(frm, text=label, anchor='e', width=30).grid(
                row=i, column=0, padx=8, pady=7, sticky='e')
            var = tk.StringVar()
            self._cfg_vars[key] = var
            ttk.Entry(frm, textvariable=var, width=width,
                      show='*' if secret else '').grid(
                row=i, column=1, padx=8, pady=7, sticky='w')
            if tip:
                ttk.Label(frm, text=tip, foreground='gray').grid(
                    row=i, column=2, padx=6, sticky='w')

        def _row_browse(i, label, key, title, filetypes=None):
            ttk.Label(frm, text=label, anchor='e', width=30).grid(
                row=i, column=0, padx=8, pady=7, sticky='e')
            var = tk.StringVar()
            self._cfg_vars[key] = var
            ttk.Entry(frm, textvariable=var, width=52).grid(
                row=i, column=1, padx=8, pady=7, sticky='w')

            def browse():
                path = (filedialog.askopenfilename(title=title, filetypes=filetypes)
                        if filetypes else filedialog.askdirectory(title=title))
                if path:
                    var.set(path)

            ttk.Button(frm, text='...', width=3, command=browse).grid(
                row=i, column=2, padx=2, sticky='w')

        _row(0, 'TNS Alias (Network Alias):', 'tns_alias', width=22,
             tip='Ex: VIASUPER  (conforme tnsnames.ora)')
        _row(1, 'Usuário:', 'user', width=22)
        _row(2, 'Senha:', 'password', secret=True, width=22)
        _row(3, 'Estab. ERP (ACERCAR.ESTAB):', 'estab', width=12,
             tip='Código do estabelecimento no ERP (ex: 1003)')
        _row_browse(4, 'Pasta / arquivo tnsnames.ora:', 'tns_dir',
                    title='Selecionar tnsnames.ora ou pasta',
                    filetypes=[('tnsnames.ora', 'tnsnames.ora'), ('Todos', '*.*')])
        _row_browse(5, 'Oracle Client (pasta bin):', 'oracle_client_dir',
                    title='Selecionar pasta bin do Oracle Client')

        ttk.Label(
            frm,
            text='Dica: pasta bin ex: C:\\oracle\\product\\19.0.0\\client_1\\bin\n'
                 '      tnsnames.ora ex: C:\\oracle\\product\\19.0.0\\client_1\\network\\admin',
            foreground='#777', justify='left',
        ).grid(row=6, column=0, columnspan=3, padx=8, pady=(2, 6), sticky='w')

        cfg = load_config()
        for k, v in self._cfg_vars.items():
            v.set(cfg.get(k, ''))

        btn_frm = ttk.Frame(outer)
        btn_frm.pack(pady=12)
        ttk.Button(btn_frm, text='Salvar', command=self._cmd_salvar_config).pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_frm, text='Testar Conexão', command=self._cmd_testar_conexao).pack(side=tk.LEFT, padx=8)

        self._cfg_status = ttk.Label(outer, text='', foreground='green')
        self._cfg_status.pack()

    # ------------------------------------------------------------------
    # Commands
    # ------------------------------------------------------------------

    def _criteria(self) -> dict:
        try:
            variacao = max(0, int(self._crit_variacao.get() or '0')) / 100.0
        except (ValueError, TypeError):
            variacao = 0.0
        return {
            'nsu':         self._crit_nsu.get(),
            'autorizacao': self._crit_aut.get(),
            'data_venda':  self._crit_data_venda.get(),
            'valor_bruto': self._crit_valor.get(),
            'estab':       self._crit_estab.get(),
            'parcela':     self._crit_parcela.get(),
            'variacao':    variacao,
        }

    def _set_status(self, msg: str):
        self._status_var.set(msg)
        self.root.update_idletasks()

    def _cmd_selecionar(self):
        path = filedialog.askopenfilename(
            title='Selecionar Planilha Cielo',
            filetypes=[('Excel', '*.xlsx *.xls'), ('Todos', '*.*')]
        )
        if path:
            self.current_file = path
            self._file_var.set(path)
            self.current_id_arquivo = os.path.splitext(os.path.basename(path))[0]

    def _cmd_importar(self):
        if not self.current_file:
            messagebox.showwarning('Aviso', 'Selecione um arquivo primeiro.')
            return

        def run():
            try:
                self._set_status('Lendo planilha…')
                id_arq, records = import_excel(self.current_file)
                self.current_id_arquivo = id_arq
                self._set_status('Conectando ao banco…')
                self.db.connect()
                self.db.ensure_table_exists()
                if self.db.arquivo_ja_importado(id_arq):
                    resp = messagebox.askyesno(
                        'Arquivo já importado',
                        f'O arquivo "{id_arq}" já existe na base.\n'
                        'Deseja reimportar (apaga os dados anteriores)?'
                    )
                    if not resp:
                        self._set_status('Importação cancelada.')
                        return
                self._set_status(f'Gravando {len(records)} registros…')
                self.db.insert_conciliacao_records(id_arq, records)
                self._set_status(f'{len(records)} registros importados com sucesso.')
                messagebox.showinfo('Sucesso', f'{len(records)} registros gravados.')
                self.root.after(0, self._cmd_carregar_grid)
            except Exception as exc:
                messagebox.showerror('Erro na importação', str(exc))
                self._set_status('Erro na importação.')
            finally:
                self.db.disconnect()

        threading.Thread(target=run, daemon=True).start()

    def _cmd_carregar_grid(self):
        if not self.current_id_arquivo:
            messagebox.showwarning('Aviso', 'Selecione e importe um arquivo primeiro.')
            return

        de_para = self._get_de_para()

        def run():
            try:
                self._set_status('Carregando dados do arquivo…')
                self.db.connect()
                _, rows = self.db.get_records_by_arquivo(self.current_id_arquivo)
            except Exception as exc:
                messagebox.showerror('Erro', str(exc))
                self._set_status('Erro ao carregar grid.')
                return
            finally:
                self.db.disconnect()

            # DB indices: [0]CONCILIADO [1]DATA_VENDA [2]ESTABELECIMENTO [3]NSU_DOC
            # [4]AUTORIZACAO [5]VALOR_BRUTO [6]TAXA_TARIFA [7]VALOR_LIQUIDO
            # [8]BANDEIRA [9]NUMERO_PARCELA [10]TOTAL_PARCELAS
            # [11]FORMA_PAGAMENTO [12]STATUS_PAGAMENTO [13]DATA_PAGAMENTO
            items = []
            conciliados = 0
            for row in rows:
                conc  = row[0]
                estab = _fmt(row[2])
                if conc == 'S':
                    conciliados += 1
                vals = [
                    _fmt(row[0]),            # conciliado
                    _fmt(row[1]),            # data_venda
                    estab,                   # estabelecimento
                    de_para.get(estab, ''),  # estab_erp (de-para lookup)
                    _fmt(row[3]),            # nsu_doc
                    _fmt(row[4]),            # autorizacao
                    _fmt_num(row[5]),        # valor_bruto
                    _fmt_num(row[7]),        # valor_liquido (skip taxa_tarifa[6])
                    _fmt(row[8]),            # bandeira
                    _fmt(row[9]) or '1',     # parcela
                    _fmt(row[10]) or '1',    # total_parc
                    _fmt(row[11]),           # forma_pag
                    _fmt(row[12]),           # status_pag
                ]
                items.append((vals, 'conciliado' if conc == 'S' else 'pendente'))

            total = len(items)
            vlr_conc_esq = sum(_br_to_float(v[6]) for v, t in items if t == 'conciliado')

            def update_ui():
                self._load_esq(items)
                self._tot_total.set(str(total))
                self._tot_conc.set(str(conciliados))
                self._tot_vlr_conc_esq.set(_fmt_num(vlr_conc_esq))
                self._tot_pend.set(str(total - conciliados))
                self._set_status(
                    f'{total} registros carregados — '
                    f'{conciliados} conciliados, {total - conciliados} pendentes.'
                )

            self.root.after(0, update_ui)

        threading.Thread(target=run, daemon=True).start()

    def _cmd_buscar(self, _silent=False):
        crit = self._criteria()
        if not any(crit.values()):
            if not _silent:
                messagebox.showwarning('Aviso', 'Selecione pelo menos um critério de conciliação.')
            return
        dt_ini = self._dt_ini.get_str()
        dt_fim = self._dt_fim.get_str()
        if not self._dt_ini.get_date() or not self._dt_fim.get_date():
            if not _silent:
                messagebox.showwarning('Aviso', 'Informe as datas no formato DD/MM/AAAA.')
            return
        if self._dt_ini.get_date() > self._dt_fim.get_date():
            if not _silent:
                messagebox.showwarning('Aviso', 'Data inicial deve ser ≤ data final.')
            return
        cfg     = load_config()
        estab   = cfg.get('estab', '1003')
        de_para = self._get_de_para()

        SHOW     = self._SHOW_DIR
        NUM_COLS = self._NUM_DIR

        def run():
            try:
                self._set_status('Buscando lançamentos no ERP…')
                self.db.connect()
                _, rows = self.db.get_acercar_records(
                    dt_ini, dt_fim, estab, crit, de_para)
            except Exception as exc:
                messagebox.showerror('Erro', str(exc))
                self._set_status('Erro ao buscar lançamentos.')
                return
            finally:
                self.db.disconnect()

            items  = []
            n_conc = n_match = n_no = 0
            for row in rows:
                n_no += 1
                vals = []
                for idx in SHOW:
                    v = row[idx]
                    if idx in NUM_COLS:
                        vals.append(_fmt_num(v))
                    elif idx in {6, 8}:  # PARCELAS / PARCELA — null = 1
                        vals.append(_fmt(v) or '1')
                    else:
                        vals.append(_fmt(v))
                vals[0] = ''
                items.append((vals, 'nomatch'))

            total = len(items)
            vlr_conc_dir = 0.0

            def update_ui():
                self._load_dir(items)
                self._tot_erp.set(str(total))
                self._tot_erp_conc.set(str(n_conc))
                self._tot_vlr_conc_dir.set(_fmt_num(vlr_conc_dir))
                self._tot_erp_match.set(str(n_match))
                self._tot_erp_nomatch.set(str(n_no))
                self._set_status(
                    f'{total} lançamentos ERP — {n_conc} conciliados, '
                    f'{n_match} com match, {n_no} sem match.'
                )

            self.root.after(0, update_ui)

        threading.Thread(target=run, daemon=True).start()

    def _cmd_conciliar(self):
        if not self.current_id_arquivo:
            messagebox.showwarning('Aviso', 'Nenhum arquivo carregado.')
            return
        if not self._all_items_dir:
            messagebox.showwarning('Aviso', 'Execute "Buscar Lançamentos" antes de conciliar.')
            return
        crit = self._criteria()
        dt_ini = self._dt_ini.get_str()
        dt_fim = self._dt_fim.get_str()
        if not self._dt_ini.get_date() or not self._dt_fim.get_date():
            messagebox.showwarning('Aviso', 'Informe as datas.')
            return
        cfg     = load_config()
        estab   = cfg.get('estab', '1003')
        de_para = self._get_de_para()
        if not messagebox.askyesno('Confirmar', 'Executar a conciliação dos registros encontrados?'):
            return

        def run():
            try:
                self._set_status('Conciliando…')
                self.db.connect()
                n = self.db.conciliar(
                    self.current_id_arquivo, dt_ini, dt_fim, estab, crit, de_para)
                self._set_status(f'Conciliação concluída — {n} registro(s) atualizado(s).')
                messagebox.showinfo('Conciliação', f'{n} registro(s) marcado(s) como conciliado(s).')
                self.root.after(0, self._reload_both_grids)
            except Exception as exc:
                messagebox.showerror('Erro', str(exc))
                self._set_status('Erro na conciliação.')
            finally:
                self.db.disconnect()

        threading.Thread(target=run, daemon=True).start()

    def _cmd_desfazer(self):
        if not self.current_id_arquivo:
            messagebox.showwarning('Aviso', 'Nenhum arquivo carregado.')
            return
        if not messagebox.askyesno(
            'Confirmar',
            f'Desfazer a conciliação de todos os registros de\n"{self.current_id_arquivo}"?'
        ):
            return

        def run():
            try:
                self.db.connect()
                n = self.db.desfazer_conciliacao(self.current_id_arquivo)
                self._set_status(f'Conciliação desfeita — {n} registro(s) revertido(s).')
                self.root.after(0, self._reload_both_grids)
            except Exception as exc:
                messagebox.showerror('Erro', str(exc))
            finally:
                self.db.disconnect()

        threading.Thread(target=run, daemon=True).start()

    def _reload_both_grids(self):
        """Recarrega grid Cielo e grid ERP em sequência num único thread (sem conflito de conexão)."""
        if not self.current_id_arquivo:
            return
        crit   = self._criteria()
        dt_ini = self._dt_ini.get_str()
        dt_fim = self._dt_fim.get_str()
        d_ini  = self._dt_ini.get_date()
        d_fim  = self._dt_fim.get_date()
        can_erp = (any(crit.values()) and d_ini and d_fim and d_ini <= d_fim)
        cfg     = load_config()
        estab   = cfg.get('estab', '1003')
        de_para = self._get_de_para()
        SHOW     = self._SHOW_DIR
        NUM_COLS = self._NUM_DIR

        def run():
            # ---- Cielo grid ----
            try:
                self._set_status('Recarregando grids…')
                self.db.connect()
                _, rows_esq = self.db.get_records_by_arquivo(self.current_id_arquivo)
            except Exception as exc:
                messagebox.showerror('Erro', str(exc))
                return
            finally:
                self.db.disconnect()

            items_esq = []
            conciliados = 0
            for row in rows_esq:
                conc  = row[0]
                estab_c = _fmt(row[2])
                if conc == 'S':
                    conciliados += 1
                vals = [
                    _fmt(row[0]), _fmt(row[1]),
                    estab_c, de_para.get(estab_c, ''),
                    _fmt(row[3]), _fmt(row[4]),
                    _fmt_num(row[5]), _fmt_num(row[7]),
                    _fmt(row[8]), _fmt(row[9]) or '1', _fmt(row[10]) or '1',
                    _fmt(row[11]), _fmt(row[12]),
                ]
                items_esq.append((vals, 'conciliado' if conc == 'S' else 'pendente'))
            total_esq    = len(items_esq)
            vlr_conc_esq = sum(_br_to_float(v[6]) for v, t in items_esq if t == 'conciliado')

            # ---- ERP grid (only if params are valid) ----
            items_dir = None
            n_conc = n_match = n_no = 0
            vlr_conc_dir = 0.0
            if can_erp:
                try:
                    self.db.connect()
                    _, rows_dir = self.db.get_acercar_records(dt_ini, dt_fim, estab, crit, de_para)
                except Exception as exc:
                    messagebox.showerror('Erro ao recarregar ERP', str(exc))
                    rows_dir = []
                finally:
                    self.db.disconnect()

                items_dir = []
                for row in rows_dir:
                    n_no += 1
                    vals = []
                    for idx in SHOW:
                        v = row[idx]
                        if idx in NUM_COLS:
                            vals.append(_fmt_num(v))
                        elif idx in {6, 8}:  # PARCELAS / PARCELA — null = 1
                            vals.append(_fmt(v) or '1')
                        else:
                            vals.append(_fmt(v))
                    vals[0] = ''
                    items_dir.append((vals, 'nomatch'))
                vlr_conc_dir = 0.0

            def update_ui():
                self._load_esq(items_esq)
                self._tot_total.set(str(total_esq))
                self._tot_conc.set(str(conciliados))
                self._tot_vlr_conc_esq.set(_fmt_num(vlr_conc_esq))
                self._tot_pend.set(str(total_esq - conciliados))
                if items_dir is not None:
                    total_dir = len(items_dir)
                    self._load_dir(items_dir)
                    self._tot_erp.set(str(total_dir))
                    self._tot_erp_conc.set(str(n_conc))
                    self._tot_vlr_conc_dir.set(_fmt_num(vlr_conc_dir))
                    self._tot_erp_match.set(str(n_match))
                    self._tot_erp_nomatch.set(str(n_no))
                    self._set_status(
                        f'Cielo: {total_esq} reg ({conciliados} conciliados) | '
                        f'ERP: {total_dir} lanç ({n_conc} conciliados, {n_match} match, {n_no} sem match).'
                    )
                else:
                    self._set_status(
                        f'{total_esq} registros — {conciliados} conciliados, '
                        f'{total_esq - conciliados} pendentes.'
                    )

            self.root.after(0, update_ui)

        threading.Thread(target=run, daemon=True).start()

    def _cmd_debug_sql(self):
        if not self.current_id_arquivo:
            messagebox.showwarning('Aviso', 'Nenhum arquivo carregado.')
            return
        crit   = self._criteria()
        dt_ini = self._dt_ini.get_str()
        dt_fim = self._dt_fim.get_str()
        cfg    = load_config()
        estab  = cfg.get('estab', '1003')
        de_para = self._get_de_para()

        sql_text = self.db.debug_sql_conciliar(
            self.current_id_arquivo, dt_ini, dt_fim, estab, crit, de_para)

        win = tk.Toplevel(self.root)
        win.title('Debug SQL — Conciliação')
        win.geometry('900x520')
        win.grab_set()

        frm = ttk.Frame(win)
        frm.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        txt = tk.Text(frm, font=('Courier New', 9), wrap='none')
        vsb = ttk.Scrollbar(frm, orient='vertical', command=txt.yview)
        hsb = ttk.Scrollbar(frm, orient='horizontal', command=txt.xview)
        txt.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        txt.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        frm.rowconfigure(0, weight=1)
        frm.columnconfigure(0, weight=1)

        txt.insert('1.0', sql_text)
        txt.configure(state='disabled')

        btn_frm = ttk.Frame(win)
        btn_frm.pack(pady=4)

        def _copy():
            win.clipboard_clear()
            win.clipboard_append(sql_text)

        ttk.Button(btn_frm, text='Copiar SQL', command=_copy).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_frm, text='Fechar', command=win.destroy).pack(side=tk.LEFT, padx=6)

    def _cmd_salvar_config(self):
        cfg = load_config()
        cfg.update({k: v.get().strip() for k, v in self._cfg_vars.items()})
        save_config(cfg)
        self._cfg_status.configure(text='✔ Configuração salva!', foreground='green')

    def _cmd_testar_conexao(self):
        self._cmd_salvar_config()
        self._cfg_status.configure(text='Testando conexão…', foreground='blue')
        self.root.update_idletasks()

        def run():
            try:
                self.db.connect()
                cur = self.db.connection.cursor()
                cur.execute('SELECT 1 FROM DUAL')
                cur.close()
                self.root.after(0, lambda: self._cfg_status.configure(
                    text='✔ Conexão bem-sucedida!', foreground='green'))
            except Exception as exc:
                self.root.after(0, lambda e=exc: self._cfg_status.configure(
                    text=f'✘ Erro: {e}', foreground='red'))
            finally:
                self.db.disconnect()

        threading.Thread(target=run, daemon=True).start()
