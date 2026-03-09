# 株価モニター
# 株価の変動を監視し、一定以上の変動があったときに通知するアプリ

import threading
import time
import json
import os
import csv
import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import tkinter.font as tkfont
import warnings
import winsound

import yfinance as yf

try:
    import pandas as pd
except Exception:
    pd = None

warnings.filterwarnings("ignore", message="pkg_resources is deprecated as an API")

# ----------------- 設定 -----------------
COMPANIES_CSV = "companies.csv"   # 形式: code,JapaneseName
CONFIG_FILE = "kabu_config.json"

WINDOW_WIDTH = 1400
WINDOW_HEIGHT = 650
DEFAULT_FONT_SIZE = 9

TABLE_COUNT = 4
ROWS_PER_TABLE = 25
UPDATE_INTERVAL = 60
DEFAULT_THRESHOLD = 0.1

# 列幅の比率: code:name:price:change = 1:2:1:1
COLUMN_RATIO = [1, 2, 1, 1]
MIN_COLUMN_WIDTH = 40
# ---------------------------------------

# ----------------- 通知機能 -----------------
notification_backend = None
try:
    from plyer import notification as plyer_notification
    notification_backend = "plyer"
except Exception:
    try:
        from win10toast import ToastNotifier
        toaster = ToastNotifier()
        notification_backend = "win10toast"
    except Exception:
        notification_backend = None

# ----------------- 音声再生機能 -----------------
pygame_available = False
try:
    import pygame
    pygame_available = True
except Exception:
    pygame_available = False


def send_notification(title, message):
    """OS通知を送る。利用できない場合は標準出力に出す。"""
    try:
        if notification_backend == "plyer":
            plyer_notification.notify(title=title, message=message, timeout=5)
        elif notification_backend == "win10toast":
            toaster.show_toast(title, message, duration=5, threaded=True)
        else:
            print("[notify suppressed]", title, message)
    except Exception as e:
        print("通知例外:", e)


# ---------- CSV 読み書き ----------
def load_companies(path: str) -> dict:
    """companies.csv を読み込み、{銘柄コード: 社名} の辞書を返す。"""
    companies = {}

    if not os.path.exists(path):
        return companies

    try:
        with open(path, "r", encoding="cp932", errors="replace") as f:
            reader = csv.reader(f)
            for row in reader:
                if not row:
                    continue

                code = str(row[0]).strip().upper().replace(".T", "")
                if not code:
                    continue

                name = row[1].strip() if len(row) > 1 else ""
                companies[code] = name
    except Exception as e:
        print("companies.csv 読み込み失敗:", e)

    return companies


def save_companies(path: str, companies: dict):
    """companies.csv を銘柄コード順で保存する。"""
    def sort_key(code):
        try:
            return int(re.sub(r"\D", "", code) or 0)
        except Exception:
            return 10**9

    try:
        items = sorted(companies.items(), key=lambda item: sort_key(item[0]))
        with open(path, "w", encoding="cp932", errors="replace", newline="") as f:
            writer = csv.writer(f)
            for code, name in items:
                writer.writerow([code, name])
    except Exception as e:
        print("companies.csv 保存失敗:", e)


# ---------- 補助関数 ----------
def normalize_ticker(text: str) -> str:
    """入力された銘柄コードを yfinance 用の形式にそろえる。"""
    text = text.strip().upper()
    if not text:
        return ""
    return text if "." in text else text + ".T"


def ticker_to_code(ticker: str) -> str:
    """7203.T -> 7203 のように .T を除いたコードを返す。"""
    return ticker.replace(".T", "").upper()


def code_sort_key(code: str) -> int:
    """銘柄コード中の数字部分をソート用キーとして返す。"""
    try:
        return int(re.sub(r"\D", "", code) or 0)
    except Exception:
        return 10**9


def ticker_sort_key(ticker: str) -> int:
    """ティッカー文字列を銘柄コード順で並べるためのキー。"""
    return code_sort_key(ticker_to_code(ticker))


def fetch_name_from_yfinance(ticker: str) -> str:
    """yfinance から会社名を取得する。"""
    try:
        stock = yf.Ticker(ticker)
        info = stock.info or {}
        for key in ("shortName", "longName"):
            value = info.get(key)
            if isinstance(value, str) and value.strip():
                return value.strip()
    except Exception as e:
        print("yfinance name 取得例外:", e)
    return ""


# ---------- 価格取得 ----------
def get_latest_prices(tickers):
    """
    複数銘柄の最新価格を取得して
    {ticker: price or None} の辞書を返す。
    """
    prices = {ticker: None for ticker in tickers}

    if not tickers:
        return prices

    try:
        df = yf.download(
            tickers,
            period="1d",
            interval="1m",
            group_by="ticker",
            threads=False,
            progress=False
        )

        is_multi_index = (
            pd is not None
            and hasattr(df, "columns")
            and isinstance(df.columns, pd.MultiIndex)
        )

        if is_multi_index:
            for ticker in tickers:
                try:
                    prices[ticker] = float(df["Close"][ticker].iloc[-1])
                except Exception:
                    prices[ticker] = None
        else:
            try:
                if "Close" in df.columns and not df["Close"].empty:
                    prices[tickers[0]] = float(df["Close"].iloc[-1])
                else:
                    prices[tickers[0]] = float(df.iloc[-1][-1])
            except Exception:
                prices[tickers[0]] = None

    except Exception as e:
        print("一括取得例外:", e)

    # 一括取得できなかった銘柄は個別取得で補う
    for ticker, price in list(prices.items()):
        if price is not None:
            continue

        try:
            stock = yf.Ticker(ticker)
            info = stock.info or {}

            latest = info.get("regularMarketPrice") or info.get("previousClose")
            if latest is not None:
                prices[ticker] = float(latest)
                continue

            history = stock.history(period="1d", interval="1m")
            if not history.empty:
                prices[ticker] = float(history["Close"].iloc[-1])

        except Exception as e:
            print(f"[{ticker}] 個別取得失敗:", e)
            prices[ticker] = None

    return prices


# ---------- アプリ本体 ----------
class StockMonitorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("株価モニター")

        # 設定読み込み
        config = self.load_config()

        self.font_size = int(config.get("font_size", DEFAULT_FONT_SIZE))
        self.table_total_width = config.get("table_total_width", None)
        self.threshold = float(config.get("threshold", DEFAULT_THRESHOLD))
        self.sound_file = config.get("sound_file", "")

        try:
            self.root.state("zoomed")
        except Exception:
            self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")

        # 社名データ読み込み
        self.companies = load_companies(COMPANIES_CSV)

        # 社名キャッシュ
        self.name_cache = {}
        for code, name in self.companies.items():
            if name:
                self.name_cache[code + ".T"] = name

        # 監視対象銘柄リスト
        self.all_tickers = self.load_tickers_from_config(config)

        # 表示用データ
        self.tables = [[] for _ in range(TABLE_COUNT)]

        # 価格関連データ
        self.current_prices = {}
        self.previous_prices = {}
        self.significant_change_notified = {}

        # スレッド制御
        self.lock = threading.Lock()
        self.running = False

        # 右クリック対象の保持
        self.context_ticker = None

        # フォント設定
        self.base_font = tkfont.Font(
            family="TkDefaultFont",
            size=self.font_size,
            weight="bold"
        )
        self.apply_treeview_style()

        # UI構築
        self.create_menu()
        self.create_top_controls()
        self.create_tables()

        # 初期表示
        self.rebuild_tables()
        self.redraw_tables()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------- 初期化関連 ----------
    def load_config(self):
        """設定ファイルを読み込む。存在しない場合は空辞書。"""
        if not os.path.exists(CONFIG_FILE):
            return {}

        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    def load_tickers_from_config(self, config):
        """
        設定から監視対象銘柄を読み込む。
        新形式 all_tickers を優先し、旧形式 tables にも対応する。
        """
        tickers = []

        if isinstance(config.get("all_tickers"), list):
            for ticker in config["all_tickers"]:
                if isinstance(ticker, str) and ticker.strip():
                    tickers.append(normalize_ticker(ticker))
        else:
            old_tables = config.get("tables", [[] for _ in range(TABLE_COUNT)])
            while len(old_tables) < TABLE_COUNT:
                old_tables.append([])

            for table in old_tables[:TABLE_COUNT]:
                for ticker in table[:ROWS_PER_TABLE]:
                    if isinstance(ticker, str) and ticker.strip():
                        tickers.append(normalize_ticker(ticker))

        # 重複削除してコード順に並べる
        unique_tickers = list(dict.fromkeys(tickers))
        return sorted(unique_tickers, key=ticker_sort_key)

    def apply_treeview_style(self):
        """Treeview と見出しのフォントを設定する。"""
        style = ttk.Style()
        style.configure("Treeview", font=self.base_font)
        style.configure("Treeview.Heading", font=self.base_font)

    # ---------- UI構築 ----------
    def create_menu(self):
        menubar = tk.Menu(self.root)

        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="閾値設定（%）", command=self.set_threshold)
        settings_menu.add_command(label="通知音設定", command=self.set_sound)
        settings_menu.add_command(label="表示設定...", command=self.set_display_settings)
        settings_menu.add_separator()
        settings_menu.add_command(label="表を全消去...", command=self.clear_table)
        menubar.add_cascade(label="設定", menu=settings_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="使い方", command=self.show_help)
        menubar.add_cascade(label="ヘルプ", menu=help_menu)

        self.root.config(menu=menubar)

    def create_top_controls(self):
        top_frame = ttk.Frame(self.root)
        top_frame.pack(side="top", fill="x", padx=8, pady=6)

        ttk.Label(top_frame, text="銘柄コード:").pack(side="left")

        self.ticker_entry = ttk.Entry(top_frame, width=14)
        self.ticker_entry.pack(side="left", padx=(4, 4))

        ttk.Button(top_frame, text="追加", command=self.add_ticker_from_entry).pack(side="left", padx=(0, 8))
        ttk.Button(top_frame, text="開始", command=self.start_monitoring).pack(side="right", padx=(4, 0))
        ttk.Button(top_frame, text="停止", command=self.stop_monitoring).pack(side="right", padx=(4, 0))

        self.status_label = ttk.Label(top_frame, text="最終更新: N/A", width=30, anchor="e")
        self.status_label.pack(side="right", padx=(0, 8))

    def create_tables(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=8, pady=8)

        self.root.update_idletasks()
        total_width = self.root.winfo_width() or WINDOW_WIDTH

        if self.table_total_width and isinstance(self.table_total_width, int):
            usable_width = max(600, min(self.table_total_width, total_width))
        else:
            usable_width = max(600, total_width - 40)

        table_width = int(usable_width / TABLE_COUNT)

        self.table_widgets = []
        self.trees = []

        for index in range(TABLE_COUNT):
            frame = ttk.Frame(main_frame, borderwidth=1, relief="solid", width=table_width)
            frame.grid(row=0, column=index, padx=4, sticky="n")
            main_frame.columnconfigure(index, weight=1)

            label = ttk.Label(
                frame,
                text=f"表 {index + 1} (0/{ROWS_PER_TABLE})",
                font=self.base_font
            )
            label.pack(side="top", pady=(6, 4))

            columns = ("code", "name", "price", "change")
            tree = ttk.Treeview(
                frame,
                columns=columns,
                show="headings",
                height=ROWS_PER_TABLE,
                selectmode="none"
            )

            tree.heading("code", text="銘柄コード")
            tree.heading("name", text="社名")
            tree.heading("price", text="金額 (円)")
            tree.heading("change", text="前回比 (%)")

            column_widths = self.calculate_column_widths(table_width)
            tree.column("code", width=column_widths[0], anchor="center")
            tree.column("name", width=column_widths[1], anchor="w")
            tree.column("price", width=column_widths[2], anchor="e")
            tree.column("change", width=column_widths[3], anchor="e")

            tree.pack(fill="both", expand=True)

            tree.tag_configure("up", background="#ff6666", foreground="#800000")
            tree.tag_configure("down", background="#66ff66", foreground="#006400")
            tree.tag_configure("neutral", background="#ffffff", foreground="#000000")

            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="削除", command=self.delete_selected_ticker)
            menu.add_separator()
            menu.add_command(label="社名変更...", command=self.rename_selected_company)

            tree.bind("<Button-3>", lambda event, i=index: self.show_context_menu(event, i))

            self.table_widgets.append({
                "frame": frame,
                "label": label,
                "menu": menu,
            })
            self.trees.append(tree)

    def calculate_column_widths(self, table_width):
        """列幅比率から各列の幅を計算する。"""
        ratio_sum = sum(COLUMN_RATIO)
        return [
            max(MIN_COLUMN_WIDTH, int(table_width * (ratio / ratio_sum)))
            for ratio in COLUMN_RATIO
        ]

    # ---------- 設定メニュー ----------
    def show_help(self):
        text = (
            "使い方:\n"
            "- 開始ボタンでモニターを開始します。\n"
            "- 銘柄追加: 上部にコード（例 7203）を入力して追加できます。\n"
            "- 表は全体でコード順に整列し、左詰めで各表に割り振られます。\n"
            "- 右クリックで削除／社名変更が可能です。\n"
            "- 閾値/通知音/フォントサイズ/表幅はメニュー[設定]から変更できます。\n"
        )
        messagebox.showinfo("ヘルプ", text, parent=self.root)

    def set_threshold(self):
        value = simpledialog.askfloat(
            "閾値設定",
            "閾値（%）を入力してください（例 0.1）:",
            initialvalue=self.threshold,
            minvalue=0.0,
            parent=self.root
        )
        if value is None:
            return

        self.threshold = float(value)
        self.save_config()
        messagebox.showinfo("設定保存", f"閾値を {self.threshold}% に設定しました。", parent=self.root)

    def set_sound(self):
        path = filedialog.askopenfilename(
            title="通知音を選択（MP3/WAV）",
            filetypes=[("Audio", "*.mp3 *.wav"), ("MP3", "*.mp3"), ("WAV", "*.wav")],
            parent=self.root
        )
        if not path:
            return

        self.sound_file = path

        if pygame_available:
            try:
                pygame.mixer.init()
            except Exception:
                pass

        self.save_config()
        messagebox.showinfo("設定保存", f"通知音を設定しました:\n{path}", parent=self.root)

    def set_display_settings(self):
        font_size = simpledialog.askinteger(
            "フォントサイズ",
            "フォントサイズを入力してください（例 9）:",
            initialvalue=self.font_size,
            minvalue=6,
            maxvalue=24,
            parent=self.root
        )
        if font_size is None:
            return

        table_width = simpledialog.askinteger(
            "テーブル横幅（ウィンドウ幅の使用）",
            "テーブルを並べるときに使用するウィンドウ幅(px)を入力してください。\n0を入れるとウィンドウ幅に合わせます。",
            initialvalue=self.table_total_width or 0,
            minvalue=0,
            parent=self.root
        )
        if table_width is None:
            return

        self.font_size = int(font_size)
        self.table_total_width = int(table_width) if table_width > 0 else None

        self.base_font = tkfont.Font(
            family="TkDefaultFont",
            size=self.font_size,
            weight="bold"
        )
        self.apply_treeview_style()
        self.update_column_widths()
        self.save_config()

        messagebox.showinfo("設定保存", "表示設定を保存しました。", parent=self.root)

    def update_column_widths(self):
        """現在の設定に合わせて各表の列幅を再計算する。"""
        self.root.update_idletasks()
        total_width = self.root.winfo_width() or WINDOW_WIDTH

        if self.table_total_width and isinstance(self.table_total_width, int):
            usable_width = max(600, min(self.table_total_width, total_width))
        else:
            usable_width = max(600, total_width - 40)

        table_width = int(usable_width / TABLE_COUNT)
        column_widths = self.calculate_column_widths(table_width)

        for tree in self.trees:
            tree.column("code", width=column_widths[0])
            tree.column("name", width=column_widths[1])
            tree.column("price", width=column_widths[2])
            tree.column("change", width=column_widths[3])
            tree.configure(style="Treeview")

    def clear_table(self):
        table_number = simpledialog.askinteger(
            "表を全消去",
            f"消去する表番号を入力してください (1-{TABLE_COUNT}):",
            minvalue=1,
            maxvalue=TABLE_COUNT,
            parent=self.root
        )
        if not table_number:
            return

        table_index = table_number - 1

        if not messagebox.askyesno("確認", f"表 {table_number} を全て消しますか？", parent=self.root):
            return

        with self.lock:
            for ticker in list(self.tables[table_index]):
                if ticker in self.all_tickers:
                    self.all_tickers.remove(ticker)

                self.current_prices.pop(ticker, None)
                self.name_cache.pop(ticker, None)
                self.significant_change_notified.pop(ticker, None)

            self.rebuild_tables()
            self.redraw_tables()
            self.save_config()

        messagebox.showinfo("完了", f"表 {table_number} を消去しました。", parent=self.root)

    # ---------- 右クリック操作 ----------
    def show_context_menu(self, event, table_index):
        tree = self.trees[table_index]
        ticker = tree.identify_row(event.y)
        self.context_ticker = ticker if ticker else None

        try:
            self.table_widgets[table_index]["menu"].tk_popup(event.x_root, event.y_root)
        finally:
            self.table_widgets[table_index]["menu"].grab_release()

    def delete_selected_ticker(self):
        ticker = self.context_ticker
        if not ticker:
            messagebox.showinfo("情報", "削除対象がありません。行で右クリックしてください。", parent=self.root)
            return

        if not messagebox.askyesno("削除確認", f"{self.make_display_text(ticker)} を削除しますか？", parent=self.root):
            return

        with self.lock:
            if ticker in self.all_tickers:
                self.all_tickers.remove(ticker)

            self.current_prices.pop(ticker, None)
            self.name_cache.pop(ticker, None)
            self.significant_change_notified.pop(ticker, None)

            self.rebuild_tables()
            self.redraw_tables()
            self.save_config()

    def rename_selected_company(self):
        ticker = self.context_ticker
        if not ticker:
            messagebox.showinfo("情報", "編集したい行で右クリックしてください。", parent=self.root)
            return

        current_name = self.get_display_name(ticker)
        new_name = simpledialog.askstring(
            "社名変更",
            f"{self.make_display_text(ticker)} の新しい社名を入力してください:",
            initialvalue=current_name,
            parent=self.root
        )
        if new_name is None:
            return

        new_name = new_name.strip()
        if not new_name:
            messagebox.showwarning("入力エラー", "社名が空です。", parent=self.root)
            return

        code = ticker_to_code(ticker)

        with self.lock:
            self.companies[code] = new_name
            self.name_cache[ticker] = new_name

            for tree in self.trees:
                if tree.exists(ticker):
                    tree.set(ticker, "name", new_name)

            save_companies(COMPANIES_CSV, self.companies)

        messagebox.showinfo("完了", f"{self.make_display_text(ticker)} の社名を変更しました。", parent=self.root)

    # ---------- 銘柄追加 ----------
    def add_ticker_from_entry(self):
        raw = self.ticker_entry.get().strip()
        if not raw:
            return

        ticker = normalize_ticker(raw)

        if ticker in self.all_tickers:
            messagebox.showinfo("追加済み", f"{self.make_display_text(ticker)} は既に登録されています。", parent=self.root)
            return

        try:
            stock = yf.Ticker(ticker)
            info = stock.info or {}
            price = info.get("regularMarketPrice") or info.get("previousClose")

            if price is None:
                history = stock.history(period="1d", interval="1m")
                if history.empty:
                    ok = messagebox.askyesno(
                        "確認",
                        f"{self.make_display_text(ticker)} のデータが取得できません。追加しますか？",
                        parent=self.root
                    )
                    if not ok:
                        return

        except Exception as e:
            print("追加チェック例外:", e)
            ok = messagebox.askyesno(
                "確認",
                f"{self.make_display_text(ticker)} の確認中にエラーが発生しました。追加しますか？",
                parent=self.root
            )
            if not ok:
                return

        with self.lock:
            if not self.add_ticker(ticker):
                messagebox.showwarning("追加失敗", "追加に失敗しました（容量上限）。", parent=self.root)
                return

        self.ticker_entry.delete(0, tk.END)

    def add_ticker(self, ticker: str) -> bool:
        """銘柄を追加し、全体を再ソートして表を更新する。"""
        ticker = normalize_ticker(ticker)

        if ticker in self.all_tickers:
            return True

        if len(self.all_tickers) >= TABLE_COUNT * ROWS_PER_TABLE:
            return False

        self.all_tickers.append(ticker)
        self.all_tickers = sorted(list(dict.fromkeys(self.all_tickers)), key=ticker_sort_key)

        self.rebuild_tables()
        self.redraw_tables()
        self.save_config()
        return True

    # ---------- 表示更新 ----------
    def rebuild_tables(self):
        """all_tickers から表ごとの表示リストを作り直す。"""
        self.tables = [[] for _ in range(TABLE_COUNT)]

        for index, ticker in enumerate(self.all_tickers[:TABLE_COUNT * ROWS_PER_TABLE]):
            table_index = index // ROWS_PER_TABLE
            self.tables[table_index].append(ticker)

    def redraw_tables(self):
        """Treeview を現在の tables に合わせて再描画する。"""
        for tree in self.trees:
            for item_id in tree.get_children():
                tree.delete(item_id)

        for table_index in range(TABLE_COUNT):
            for ticker in self.tables[table_index]:
                code = ticker_to_code(ticker)
                name = self.get_display_name(ticker)
                price = self.current_prices.get(ticker)
                price_text = f"{price:,.2f}" if price is not None else "N/A"
                change_text = self.format_change(ticker)

                self.trees[table_index].insert(
                    "",
                    "end",
                    iid=ticker,
                    values=(code, name, price_text, change_text),
                    tags=("neutral",)
                )

            self.table_widgets[table_index]["label"].config(
                text=f"表 {table_index + 1} ({len(self.tables[table_index])}/{ROWS_PER_TABLE})"
            )

    def get_display_name(self, ticker: str) -> str:
        """
        表示用社名を返す。
        優先順位は companies.csv -> キャッシュ -> yfinance -> 銘柄コード。
        """
        code = ticker_to_code(ticker)

        if code in self.companies and self.companies[code]:
            self.name_cache[ticker] = self.companies[code]
            return self.companies[code]

        if ticker in self.name_cache and self.name_cache[ticker]:
            return self.name_cache[ticker]

        name = fetch_name_from_yfinance(ticker)
        if name:
            self.name_cache[ticker] = name
            return name

        self.name_cache[ticker] = code
        return code

    # ---------- 監視 ----------
    def start_monitoring(self):
        if self.running:
            return

        self.running = True
        thread = threading.Thread(target=self.monitor_loop, daemon=True)
        thread.start()

    def stop_monitoring(self):
        self.running = False

    def monitor_loop(self):
        while self.running:
            try:
                with self.lock:
                    tickers = list(self.all_tickers)

                if not tickers:
                    time.sleep(UPDATE_INTERVAL)
                    continue

                prices = get_latest_prices(tickers)
                self.root.after(0, lambda result=prices: self.update_prices(result))

            except Exception as e:
                print("監視ループ例外:", e)

            time.sleep(UPDATE_INTERVAL)

    def update_prices(self, prices):
        """取得した価格を画面に反映し、必要なら通知を出す。"""
        now = time.strftime("%Y-%m-%d %H:%M:%S")
        self.status_label.config(text=f"最終更新: {now}")

        for table_index in range(TABLE_COUNT):
            for ticker in self.tables[table_index]:
                latest_price = prices.get(ticker)
                if latest_price is not None:
                    self.current_prices[ticker] = latest_price

                previous_price = self.previous_prices.get(ticker)
                change_text = "N/A"
                tag = "neutral"

                if (
                    previous_price is not None
                    and self.current_prices.get(ticker) is not None
                    and previous_price > 0
                ):
                    change = (self.current_prices[ticker] - previous_price) / previous_price * 100.0
                    change_text = f"{change:+.2f} %"

                    if abs(change) >= self.threshold:
                        tag = "up" if change > 0 else "down"

                        if not self.significant_change_notified.get(ticker, False):
                            self.alert_price_change(ticker, change)
                            self.significant_change_notified[ticker] = True
                    else:
                        self.significant_change_notified[ticker] = False

                price_text = (
                    f"{self.current_prices[ticker]:,.2f}"
                    if self.current_prices.get(ticker) is not None
                    else "N/A"
                )

                if self.trees[table_index].exists(ticker):
                    self.trees[table_index].set(ticker, "price", price_text)
                    self.trees[table_index].set(ticker, "change", change_text)
                    self.trees[table_index].item(ticker, tags=(tag,))
                else:
                    code = ticker_to_code(ticker)
                    name = self.get_display_name(ticker)
                    self.trees[table_index].insert(
                        "",
                        "end",
                        iid=ticker,
                        values=(code, name, price_text, change_text),
                        tags=(tag,)
                    )

                if self.current_prices.get(ticker) is not None:
                    self.previous_prices[ticker] = self.current_prices[ticker]

    def format_change(self, ticker):
        """前回取得価格との差分を文字列で返す。"""
        previous_price = self.previous_prices.get(ticker)
        current_price = self.current_prices.get(ticker)

        if previous_price is None or current_price is None:
            return "N/A"

        try:
            change = (current_price - previous_price) / previous_price * 100.0
            return f"{change:+.2f} %"
        except Exception:
            return "N/A"

    def alert_price_change(self, ticker, change):
        """通知音とOS通知を出す。"""
        try:
            if pygame_available and self.sound_file and os.path.exists(self.sound_file):
                try:
                    if not pygame.mixer.get_init():
                        pygame.mixer.init()
                    pygame.mixer.music.load(self.sound_file)
                    pygame.mixer.music.play()
                except Exception as e:
                    print("pygame再生失敗:", e)
                    if self.sound_file.lower().endswith(".wav"):
                        winsound.PlaySound(self.sound_file, winsound.SND_FILENAME | winsound.SND_ASYNC)
                    else:
                        winsound.Beep(1000, 300)
            else:
                if self.sound_file and os.path.exists(self.sound_file) and self.sound_file.lower().endswith(".wav"):
                    winsound.PlaySound(self.sound_file, winsound.SND_FILENAME | winsound.SND_ASYNC)
                else:
                    winsound.Beep(1000, 300)
        except Exception as e:
            print("サウンド再生例外:", e)

        send_notification("株価変動アラート", f"{self.make_display_text(ticker)} {change:+.2f}%")

    # ---------- 保存・表示 ----------
    def make_display_text(self, ticker):
        name = self.get_display_name(ticker)
        code = ticker_to_code(ticker)
        return f"{name} ({code})"

    def save_config(self):
        config = {
            "all_tickers": self.all_tickers,
            "threshold": self.threshold,
            "sound_file": self.sound_file,
            "font_size": self.font_size,
            "table_total_width": self.table_total_width,
        }

        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print("config 保存失敗:", e)

    def on_close(self):
        self.running = False
        save_companies(COMPANIES_CSV, self.companies)
        self.save_config()
        self.root.destroy()


# ---------- 実行 ----------
if __name__ == "__main__":
    root = tk.Tk()
    app = StockMonitorApp(root)
    root.mainloop()