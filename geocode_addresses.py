# 利用 Google Geocoding API 將地址資料批量轉換為經緯度（GUI 版本）
import os
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd
import requests

API_URL = "https://maps.googleapis.com/maps/api/geocode/json"
MAX_RETRIES = 3
BASE_DELAY = 1.0  # 指數退避的基礎延遲秒數

# ── 顏色主題 ──────────────────────────────────────────────
BG_PRIMARY = "#1e1e2e"       # 深色背景
BG_SECONDARY = "#2a2a3c"     # 卡片背景
BG_INPUT = "#363649"         # 輸入框背景
FG_PRIMARY = "#cdd6f4"       # 主要文字
FG_SECONDARY = "#a6adc8"     # 次要文字
FG_ACCENT = "#89b4fa"        # 強調色（藍）
FG_SUCCESS = "#a6e3a1"       # 成功（綠）
FG_ERROR = "#f38ba8"         # 錯誤（紅）
FG_WARNING = "#fab387"       # 警告（橘）
BORDER_COLOR = "#45475a"     # 邊框
BUTTON_BG = "#89b4fa"        # 按鈕背景
BUTTON_FG = "#1e1e2e"        # 按鈕文字
BUTTON_HOVER = "#b4d0fb"     # 按鈕懸停
PROGRESS_BG = "#45475a"      # 進度條背景
PROGRESS_FG = "#89b4fa"      # 進度條填充


def geocode_address(
    address: str, api_key: str
) -> tuple[float | None, float | None, str | None]:
    """將單一地址轉換為經緯度，含重試機制。"""
    for attempt in range(MAX_RETRIES):
        try:
            response = requests.get(
                API_URL, params={"address": address, "key": api_key}, timeout=10
            )
        except requests.RequestException as exc:
            return None, None, f"連線錯誤: {exc}"

        if response.status_code != 200:
            return None, None, f"HTTP Error: {response.status_code}"

        data = response.json()
        status = data.get("status")

        if status == "OK" and data["results"]:
            location = data["results"][0]["geometry"]["location"]
            return location["lat"], location["lng"], None

        if status == "OVER_QUERY_LIMIT":
            delay = BASE_DELAY * (2**attempt)
            time.sleep(delay)
            continue

        # ZERO_RESULTS, REQUEST_DENIED, INVALID_REQUEST 等
        return None, None, f"{status}: {data.get('error_message', '無結果')}"

    return None, None, "OVER_QUERY_LIMIT: 已達最大重試次數"


# ── GUI 應用程式 ──────────────────────────────────────────
class GeocoderApp:
    """主視窗應用程式。"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("📍 地址批量轉經緯度工具")
        self.root.geometry("780x680")
        self.root.minsize(680, 600)
        self.root.configure(bg=BG_PRIMARY)

        self.df: pd.DataFrame | None = None
        self.input_path: str | None = None
        self.is_running = False
        self.cancel_flag = False

        self._configure_styles()
        self._build_ui()

    # ── ttk 樣式 ─────────────────────────────────────────
    def _configure_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        # 全域
        style.configure(".", background=BG_PRIMARY, foreground=FG_PRIMARY,
                         fieldbackground=BG_INPUT, borderwidth=0)

        # Frame
        style.configure("Card.TFrame", background=BG_SECONDARY, relief="flat")

        # Label
        style.configure("TLabel", background=BG_SECONDARY, foreground=FG_PRIMARY,
                         font=("Helvetica Neue", 12))
        style.configure("Title.TLabel", background=BG_PRIMARY, foreground=FG_PRIMARY,
                         font=("Helvetica Neue", 20, "bold"))
        style.configure("Subtitle.TLabel", background=BG_PRIMARY, foreground=FG_SECONDARY,
                         font=("Helvetica Neue", 11))
        style.configure("Section.TLabel", background=BG_SECONDARY, foreground=FG_ACCENT,
                         font=("Helvetica Neue", 13, "bold"))
        style.configure("Status.TLabel", background=BG_PRIMARY, foreground=FG_SECONDARY,
                         font=("Helvetica Neue", 11))

        # Entry
        style.configure("TEntry", fieldbackground=BG_INPUT, foreground=FG_PRIMARY,
                         insertcolor=FG_PRIMARY, padding=8,
                         font=("Menlo", 11))

        # Combobox
        style.configure("TCombobox", fieldbackground=BG_INPUT, foreground=FG_PRIMARY,
                         padding=8, font=("Helvetica Neue", 11))
        style.map("TCombobox",
                  fieldbackground=[("readonly", BG_INPUT)],
                  foreground=[("readonly", FG_PRIMARY)])

        # Button
        style.configure("Accent.TButton", background=BUTTON_BG, foreground=BUTTON_FG,
                         font=("Helvetica Neue", 12, "bold"), padding=(20, 10))
        style.map("Accent.TButton",
                  background=[("active", BUTTON_HOVER), ("disabled", BORDER_COLOR)])

        style.configure("Cancel.TButton", background=FG_ERROR, foreground=BUTTON_FG,
                         font=("Helvetica Neue", 12, "bold"), padding=(20, 10))
        style.map("Cancel.TButton",
                  background=[("active", "#e06c8a"), ("disabled", BORDER_COLOR)])

        style.configure("Browse.TButton", background=FG_ACCENT, foreground=BUTTON_FG,
                         font=("Helvetica Neue", 11), padding=(12, 6))
        style.map("Browse.TButton",
                  background=[("active", BUTTON_HOVER)])

        # Progressbar
        style.configure("Custom.Horizontal.TProgressbar",
                         troughcolor=PROGRESS_BG, background=PROGRESS_FG,
                         thickness=18, borderwidth=0)

    # ── UI 建構 ──────────────────────────────────────────
    def _build_ui(self):
        # 外層容器
        container = tk.Frame(self.root, bg=BG_PRIMARY, padx=24, pady=16)
        container.pack(fill="both", expand=True)

        # ── 標題 ──
        ttk.Label(container, text="📍 地址批量轉經緯度工具",
                  style="Title.TLabel").pack(anchor="w")
        ttk.Label(container, text="選取 Excel / CSV 檔案，自動將地址欄位轉換為經緯度座標",
                  style="Subtitle.TLabel").pack(anchor="w", pady=(2, 16))

        # ── 檔案選取卡片 ──
        file_card = ttk.Frame(container, style="Card.TFrame", padding=16)
        file_card.pack(fill="x", pady=(0, 10))

        ttk.Label(file_card, text="① 選擇檔案", style="Section.TLabel").pack(anchor="w")

        file_row = ttk.Frame(file_card, style="Card.TFrame")
        file_row.pack(fill="x", pady=(8, 0))

        self.file_var = tk.StringVar(value="尚未選取檔案")
        file_entry = ttk.Entry(file_row, textvariable=self.file_var,
                               state="readonly", font=("Menlo", 11))
        file_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        ttk.Button(file_row, text="瀏覽…", style="Browse.TButton",
                   command=self.browse_file).pack(side="right")

        # ── 設定卡片 ──
        settings_card = ttk.Frame(container, style="Card.TFrame", padding=16)
        settings_card.pack(fill="x", pady=(0, 10))

        ttk.Label(settings_card, text="② 設定", style="Section.TLabel").pack(anchor="w")

        settings_grid = ttk.Frame(settings_card, style="Card.TFrame")
        settings_grid.pack(fill="x", pady=(8, 0))
        settings_grid.columnconfigure(1, weight=1)

        # 地址欄位選擇
        ttk.Label(settings_grid, text="地址欄位：").grid(
            row=0, column=0, sticky="w", padx=(0, 12), pady=4)
        self.column_var = tk.StringVar()
        self.column_combo = ttk.Combobox(
            settings_grid, textvariable=self.column_var,
            state="disabled", values=[]
        )
        self.column_combo.grid(row=0, column=1, sticky="ew", pady=4)

        # API 金鑰
        ttk.Label(settings_grid, text="API 金鑰：").grid(
            row=1, column=0, sticky="w", padx=(0, 12), pady=4)
        self.key_var = tk.StringVar()
        self.key_entry = ttk.Entry(settings_grid, textvariable=self.key_var,
                                   show="•", font=("Menlo", 11))
        self.key_entry.grid(row=1, column=1, sticky="ew", pady=4)

        # 顯示/隱藏金鑰
        self.show_key = tk.BooleanVar(value=False)
        self.toggle_btn = tk.Checkbutton(
            settings_grid, text="顯示", variable=self.show_key,
            command=self._toggle_key_visibility,
            bg=BG_SECONDARY, fg=FG_SECONDARY, selectcolor=BG_INPUT,
            activebackground=BG_SECONDARY, activeforeground=FG_PRIMARY,
            font=("Helvetica Neue", 10)
        )
        self.toggle_btn.grid(row=1, column=2, padx=(8, 0))

        # ── 按鈕列 ──
        btn_row = tk.Frame(container, bg=BG_PRIMARY)
        btn_row.pack(fill="x", pady=(4, 10))

        self.start_btn = ttk.Button(
            btn_row, text="▶  開始轉換", style="Accent.TButton",
            command=self.start_geocoding
        )
        self.start_btn.pack(side="left")

        self.cancel_btn = ttk.Button(
            btn_row, text="■  取消", style="Cancel.TButton",
            command=self.cancel_geocoding, state="disabled"
        )
        self.cancel_btn.pack(side="left", padx=(10, 0))

        # ── 進度區 ──
        progress_frame = tk.Frame(container, bg=BG_PRIMARY)
        progress_frame.pack(fill="x", pady=(0, 6))

        self.progress_var = tk.DoubleVar(value=0)
        self.progressbar = ttk.Progressbar(
            progress_frame, variable=self.progress_var,
            maximum=100, style="Custom.Horizontal.TProgressbar"
        )
        self.progressbar.pack(fill="x")

        self.status_var = tk.StringVar(value="就緒")
        ttk.Label(progress_frame, textvariable=self.status_var,
                  style="Status.TLabel").pack(anchor="w", pady=(4, 0))

        # ── 日誌 ──
        log_label = ttk.Label(container, text="③ 執行記錄",
                              style="Section.TLabel",
                              background=BG_PRIMARY)
        log_label.pack(anchor="w", pady=(4, 4))

        log_frame = tk.Frame(container, bg=BORDER_COLOR, bd=1, relief="solid")
        log_frame.pack(fill="both", expand=True)

        self.log_text = tk.Text(
            log_frame, bg=BG_SECONDARY, fg=FG_PRIMARY,
            font=("Menlo", 10), wrap="word",
            insertbackground=FG_PRIMARY, selectbackground=FG_ACCENT,
            selectforeground=BUTTON_FG, relief="flat",
            padx=12, pady=10, state="disabled",
            highlightthickness=0
        )
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)

        # 日誌文字標籤（顏色）
        self.log_text.tag_configure("success", foreground=FG_SUCCESS)
        self.log_text.tag_configure("error", foreground=FG_ERROR)
        self.log_text.tag_configure("warning", foreground=FG_WARNING)
        self.log_text.tag_configure("info", foreground=FG_ACCENT)

    # ── 輔助方法 ─────────────────────────────────────────
    def _toggle_key_visibility(self):
        self.key_entry.configure(show="" if self.show_key.get() else "•")

    def log(self, message: str, tag: str = ""):
        """在日誌區新增一行文字。"""
        self.log_text.configure(state="normal")
        if tag:
            self.log_text.insert("end", message + "\n", tag)
        else:
            self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    # ── 檔案瀏覽 ─────────────────────────────────────────
    def browse_file(self):
        path = filedialog.askopenfilename(
            title="選取地址資料檔案",
            filetypes=[
                ("Excel 檔案", "*.xlsx *.xls"),
                ("CSV 檔案", "*.csv"),
                ("所有檔案", "*.*"),
            ],
        )
        if not path:
            return

        self.input_path = path
        self.file_var.set(os.path.basename(path))
        self.log(f"📂 已選取檔案: {path}", "info")
        self._load_columns(path)

    def _load_columns(self, path: str):
        """讀取檔案並填入欄位下拉選單。"""
        try:
            if path.lower().endswith(".csv"):
                self.df = pd.read_csv(path)
            else:
                self.df = pd.read_excel(path)
        except Exception as exc:
            messagebox.showerror("讀取失敗", f"無法讀取檔案:\n{exc}")
            self.df = None
            return

        columns = list(self.df.columns)
        self.column_combo["values"] = columns
        self.column_combo["state"] = "readonly"

        # 嘗試自動選中常見的地址欄位名稱
        auto_select = None
        for name in ("ADDR", "addr", "地址", "Address", "address"):
            if name in columns:
                auto_select = name
                break
        if auto_select:
            self.column_var.set(auto_select)
        elif columns:
            self.column_combo.current(0)

        self.log(
            f"✅ 成功讀取 {len(self.df)} 筆資料，共 {len(columns)} 個欄位: "
            f"{', '.join(columns)}",
            "success",
        )

    # ── 開始 / 取消 ──────────────────────────────────────
    def start_geocoding(self):
        # 驗證
        if self.df is None:
            messagebox.showwarning("提示", "請先選擇要轉換的檔案！")
            return
        if not self.column_var.get():
            messagebox.showwarning("提示", "請選擇地址欄位！")
            return
        api_key = self.key_var.get().strip()
        if not api_key:
            messagebox.showwarning("提示", "請輸入 Google Maps API 金鑰！")
            return

        # 詢問輸出位置
        default_name = os.path.splitext(os.path.basename(self.input_path))[0] + "_結果.xlsx"
        output_path = filedialog.asksaveasfilename(
            title="儲存結果檔案",
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel 檔案", "*.xlsx"), ("CSV 檔案", "*.csv")],
        )
        if not output_path:
            return

        # 鎖定 UI
        self.is_running = True
        self.cancel_flag = False
        self.start_btn["state"] = "disabled"
        self.cancel_btn["state"] = "normal"
        self.progress_var.set(0)

        thread = threading.Thread(
            target=self._geocode_thread,
            args=(api_key, self.column_var.get(), output_path),
            daemon=True,
        )
        thread.start()

    def cancel_geocoding(self):
        if self.is_running:
            self.cancel_flag = True
            self.log("⏹ 使用者要求取消，等待目前查詢完成…", "warning")

    def _geocode_thread(self, api_key: str, column: str, output_path: str):
        """在背景執行緒中執行地理編碼。"""
        df = self.df.copy()
        total = len(df)

        df["經度"] = None
        df["緯度"] = None
        df["錯誤原因"] = None

        success = 0
        fail = 0

        self.root.after(0, self.log,
                        f"\n{'─' * 50}", "info")
        self.root.after(0, self.log,
                        f"🚀 開始轉換，共 {total} 筆地址…\n", "info")

        for i, (index, row) in enumerate(df.iterrows()):
            if self.cancel_flag:
                self.root.after(0, self.log, f"⛔ 已取消（完成 {i}/{total}）", "warning")
                break

            address = str(row[column]).strip()
            if not address or address.lower() == "nan":
                df.at[index, "錯誤原因"] = "空白地址"
                fail += 1
                self.root.after(0, self.log,
                                f"  [{i + 1}/{total}] ⚠ （空白地址）", "warning")
            else:
                lat, lng, error = geocode_address(address, api_key)
                df.at[index, "經度"] = lng
                df.at[index, "緯度"] = lat
                df.at[index, "錯誤原因"] = error

                if error:
                    fail += 1
                    self.root.after(0, self.log,
                                    f"  [{i + 1}/{total}] ✗ {address} → {error}", "error")
                else:
                    success += 1
                    self.root.after(0, self.log,
                                    f"  [{i + 1}/{total}] ✓ {address} → ({lat}, {lng})",
                                    "success")

                time.sleep(0.05)

            # 更新進度
            pct = (i + 1) / total * 100
            self.root.after(0, self._update_progress, pct, i + 1, total, success, fail)

        # 儲存結果
        try:
            if output_path.lower().endswith(".csv"):
                df.to_csv(output_path, index=False, encoding="utf-8-sig")
            else:
                df.to_excel(output_path, index=False)

            self.root.after(0, self.log,
                            f"\n{'─' * 50}", "info")
            self.root.after(0, self.log,
                            f"✅ 完成！成功 {success}/{total} 筆，失敗 {fail} 筆",
                            "success")
            self.root.after(0, self.log,
                            f"📄 結果已儲存至: {output_path}", "info")
        except Exception as exc:
            self.root.after(0, self.log,
                            f"❌ 儲存檔案時發生錯誤: {exc}", "error")

        # 解鎖 UI
        self.root.after(0, self._finish)

    def _update_progress(self, pct: float, current: int, total: int,
                         success: int, fail: int):
        self.progress_var.set(pct)
        self.status_var.set(
            f"進度 {current}/{total}（{pct:.0f}%）　✓ 成功 {success}　✗ 失敗 {fail}"
        )

    def _finish(self):
        self.is_running = False
        self.cancel_flag = False
        self.start_btn["state"] = "normal"
        self.cancel_btn["state"] = "disabled"
        if self.progress_var.get() >= 100:
            self.status_var.set("✅ 轉換完成")


# ── 啟動 ──────────────────────────────────────────────────
def main():
    root = tk.Tk()
    GeocoderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
