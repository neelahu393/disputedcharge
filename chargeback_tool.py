"""
扣回/回補調帳上傳檔產生器 v6
比對邏輯：來源商店訂單編號 ↔ 查詢檔商店訂單編號
          有效條件：交易狀態回應訊息 = 交易成功（直接取交易成功那筆，不因同時存在其他狀態而排除）
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import sys


def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return relative_path


def process_files(source_path, query_path1, query_path2, output_path, log_callback, done_callback):
    try:
        log_callback("📦 載入必要套件...")
        import pandas as pd
        import pyxlsb
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils.datetime import from_excel

        # ── 讀取來源 xlsb ──────────────────────────────────────────
        log_callback("📂 讀取來源檔案（xlsb）...")
        with pyxlsb.open_workbook(source_path) as wb:
            with wb.get_sheet(wb.sheets[0]) as sheet:
                all_rows = list(sheet.rows())

        header = [c.v for c in all_rows[0]]
        log_callback(f"  來源欄位：{header}")

        col_store    = header.index("商店代號")
        col_order    = header.index("商店訂單編號")
        col_issuer   = header.index("發卡機構")
        col_type     = header.index("請退款")
        col_amount   = header.index("請退款金額")
        col_procdate = header.index("藍新處理日期")
        col_paydate  = header.index("實際撥款日期")

        def fmt_datetime(v):
            if v is None: return None
            if isinstance(v, str): return v
            if isinstance(v, (int, float)):
                try: return from_excel(v).strftime("%Y-%m-%d %H:%M:%S")
                except: return str(v)
            try: return v.strftime("%Y-%m-%d %H:%M:%S")
            except: return str(v)

        def fmt_date(v):
            if v is None: return None
            if isinstance(v, str): return v.split(" ")[0] if " " in v else v
            if isinstance(v, (int, float)):
                try: return from_excel(v).strftime("%Y-%m-%d")
                except: return str(v)
            try: return v.strftime("%Y-%m-%d")
            except: return str(v)

        rows_data = []
        count_chargeback = 0
        count_supplement = 0

        for r in all_rows[1:]:
            cells = [c.v for c in r]
            if not any(cells): continue
            issuer    = cells[col_issuer]   if col_issuer   < len(cells) else None
            proc_type = cells[col_type]     if col_type     < len(cells) else None
            if proc_type == "銀行扣回": count_chargeback += 1
            elif proc_type == "回補":   count_supplement += 1
            rows_data.append({
                "store":      cells[col_store]  if col_store  < len(cells) else None,
                "shop_order": str(cells[col_order]) if (col_order < len(cells) and cells[col_order] is not None) else None,
                "eztxn":      None,
                "payment":    1 if issuer else None,
                "proc_date":  fmt_datetime(cells[col_procdate] if col_procdate < len(cells) else None),
                "handle":     proc_type,
                "amount":     cells[col_amount] if col_amount < len(cells) else None,
                "pay_date":   fmt_date(cells[col_paydate] if col_paydate < len(cells) else None),
            })

        log_callback(f"✅ 讀取完成：共 {len(rows_data)} 筆")
        log_callback(f"  扣回：{count_chargeback} 筆　回補：{count_supplement} 筆")

        # ── 建立比對 mapping：只取交易成功那筆 ────────────────────
        dfs = []
        for idx, qpath in enumerate([query_path1, query_path2], 1):
            if qpath and os.path.exists(qpath):
                log_callback(f"🔍 讀取交易查詢檔 {idx}：{os.path.basename(qpath)}")
                df_q = pd.read_excel(qpath, dtype=str)
                log_callback(f"  共 {len(df_q)} 筆")
                dfs.append(df_q)

        if dfs:
            df_all = pd.concat(dfs, ignore_index=True)
            # 只保留交易成功，重複訂單號取第一筆
            df_valid = (df_all[df_all["交易狀態回應訊息"] == "交易成功"]
                        .drop_duplicates(subset="商店訂單編號", keep="first"))
            log_callback(f"  交易成功筆數：{len(df_valid)}")
            mapping = dict(zip(df_valid["商店訂單編號"],
                               zip(df_valid["ezAIO交易序號"], df_valid["門市代號"])))
        else:
            mapping = {}
            log_callback("⚠️  未提供查詢檔，略過門市代號比對")

        # ── 套用 mapping ───────────────────────────────────────────
        replaced = 0
        for row in rows_data:
            k = row["shop_order"]
            if k and k in mapping:
                row["eztxn"] = mapping[k][0]
                row["store"] = mapping[k][1]
                replaced += 1

        unmatched = sum(1 for r in rows_data if not r["eztxn"])
        log_callback(f"🔄 比對結果：{replaced} 筆成功　{unmatched} 筆未比對到")
        if unmatched > 0:
            log_callback("  ⚠️  未比對到的列，ezAIO欄位保留來源原值")

        # ── 輸出 xlsx ──────────────────────────────────────────────
        log_callback("📝 建立輸出 Excel 檔案...")
        col_handle_name = f"處理方式(扣回{count_chargeback}：/回補{count_supplement}：)"
        headers = [
            "ezAIO門市代號",
            "ezAIO交易序號",
            "支付方式(信用卡：1/電子錢包：2)",
            "藍新金流執行日期(YYYY-MM-DD HH:MM:SS)",
            col_handle_name,
            "訂單金額",
            "實際撥款日期(YYYY-MM-DD)",
        ]

        wb_out = Workbook()
        ws = wb_out.active
        ws.title = "調帳上傳"

        hf     = PatternFill("solid", fgColor="4472C4")
        hfont  = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        halign = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin   = Side(style="thin", color="AAAAAA")
        bdr    = Border(left=thin, right=thin, top=thin, bottom=thin)
        alt    = PatternFill("solid", fgColor="EEF2FF")
        dfont  = Font(name="Arial", size=10)
        dc     = Alignment(horizontal="center", vertical="center")
        dl     = Alignment(horizontal="left",   vertical="center")

        for ci, h in enumerate(headers, 1):
            cell = ws.cell(1, ci, h)
            cell.font = hfont; cell.fill = hf; cell.alignment = halign; cell.border = bdr
        ws.row_dimensions[1].height = 36

        for ri, row in enumerate(rows_data, 2):
            fill = alt if ri % 2 == 0 else None
            vals = [
                row["store"],
                row["eztxn"],
                row["payment"],
                row["proc_date"],
                row["handle"],
                row["amount"],
                row["pay_date"],
            ]
            for ci, val in enumerate(vals, 1):
                cell = ws.cell(ri, ci, val)
                cell.font = dfont; cell.border = bdr
                cell.alignment = dc if ci in (3, 5, 6, 7) else dl
                if fill: cell.fill = fill

        for i, w in enumerate([18, 22, 12, 26, 28, 12, 20], 1):
            ws.column_dimensions[ws.cell(1, i).column_letter].width = w
        ws.freeze_panes = "A2"

        wb_out.save(output_path)
        log_callback(f"💾 已儲存：{output_path}")
        done_callback(True, output_path, count_chargeback, count_supplement)

    except Exception as e:
        import traceback
        log_callback(f"❌ 發生錯誤：{e}")
        log_callback(traceback.format_exc())
        done_callback(False, str(e), 0, 0)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("扣回/回補調帳上傳檔產生器")
        self.geometry("680x580")
        self.resizable(False, False)
        self.configure(bg="#F0F4FF")
        self._build_ui()

    def _build_ui(self):
        tk.Label(self, text="扣回 / 回補 調帳上傳檔產生器",
                 font=("Microsoft JhengHei", 15, "bold"),
                 bg="#3B5BDB", fg="white", pady=12).pack(fill="x")

        frame = tk.Frame(self, bg="#F0F4FF", padx=20, pady=10)
        frame.pack(fill="x")

        def make_row(row_idx, label, var_attr, pick_cmd, optional=False):
            lbl = label + ("（選填）" if optional else "")
            tk.Label(frame, text=lbl,
                     font=("Microsoft JhengHei", 10), bg="#F0F4FF").grid(
                     row=row_idx, column=0, sticky="w", pady=5)
            var = tk.StringVar()
            setattr(self, var_attr, var)
            tk.Entry(frame, textvariable=var, width=44,
                     font=("Consolas", 9)).grid(row=row_idx, column=1, padx=6)
            tk.Button(frame, text="瀏覽", command=pick_cmd,
                      bg="#3B5BDB", fg="white",
                      font=("Microsoft JhengHei", 9),
                      relief="flat", padx=10).grid(row=row_idx, column=2)

        make_row(0, "來源檔（xlsb）：",       "src_var", self._pick_source)
        make_row(1, "交易查詢檔 1（xlsx）：",  "q1_var",  self._pick_q1)
        make_row(2, "交易查詢檔 2（xlsx）：",  "q2_var",  self._pick_q2, optional=True)
        make_row(3, "輸出路徑（xlsx）：",      "dst_var", self._pick_dest)

        self.run_btn = tk.Button(self, text="▶  開始產生",
                                 font=("Microsoft JhengHei", 11, "bold"),
                                 bg="#2F9E44", fg="white",
                                 relief="flat", padx=20, pady=8,
                                 command=self._run)
        self.run_btn.pack(pady=(8, 0))

        self.progress = ttk.Progressbar(self, mode="indeterminate", length=640)
        self.progress.pack(pady=(8, 0))

        tk.Label(self, text="執行記錄：",
                 font=("Microsoft JhengHei", 9), bg="#F0F4FF",
                 anchor="w").pack(fill="x", padx=20)
        log_frame = tk.Frame(self, bg="#F0F4FF", padx=20, pady=4)
        log_frame.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_frame, height=10, font=("Consolas", 9),
                                state="disabled", bg="#1E1E2E", fg="#CDD6F4",
                                insertbackground="white", relief="flat")
        scroll = tk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

    def _pick_source(self):
        path = filedialog.askopenfilename(
            title="選擇來源 xlsb 檔案",
            filetypes=[("Excel Binary", "*.xlsb"), ("所有檔案", "*.*")])
        if path:
            self.src_var.set(path)
            base = os.path.splitext(path)[0]
            self.dst_var.set(base + "_調帳上傳.xlsx")

    def _pick_q1(self):
        path = filedialog.askopenfilename(
            title="選擇交易查詢檔 1（xlsx）",
            filetypes=[("Excel 檔案", "*.xlsx"), ("所有檔案", "*.*")])
        if path: self.q1_var.set(path)

    def _pick_q2(self):
        path = filedialog.askopenfilename(
            title="選擇交易查詢檔 2（xlsx，選填）",
            filetypes=[("Excel 檔案", "*.xlsx"), ("所有檔案", "*.*")])
        if path: self.q2_var.set(path)

    def _pick_dest(self):
        path = filedialog.asksaveasfilename(
            title="選擇輸出路徑",
            defaultextension=".xlsx",
            filetypes=[("Excel 檔案", "*.xlsx"), ("所有檔案", "*.*")])
        if path: self.dst_var.set(path)

    def _log(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _run(self):
        src = self.src_var.get().strip()
        q1  = self.q1_var.get().strip()
        q2  = self.q2_var.get().strip()
        dst = self.dst_var.get().strip()

        if not src:
            messagebox.showwarning("提示", "請先選擇來源 xlsb 檔案。"); return
        if not dst:
            messagebox.showwarning("提示", "請先指定輸出路徑。"); return
        if not os.path.exists(src):
            messagebox.showerror("錯誤", f"找不到來源檔案：\n{src}"); return

        self.run_btn.configure(state="disabled")
        self.progress.start(12)
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        self._log("🚀 開始處理...")

        threading.Thread(
            target=process_files,
            args=(src, q1 or None, q2 or None, dst, self._safe_log, self._done),
            daemon=True
        ).start()

    def _safe_log(self, msg):
        self.after(0, self._log, msg)

    def _done(self, success, path_or_err, cb, cs):
        self.after(0, self._on_done, success, path_or_err, cb, cs)

    def _on_done(self, success, path_or_err, count_cb, count_cs):
        self.progress.stop()
        self.run_btn.configure(state="normal")
        if success:
            messagebox.showinfo(
                "完成",
                f"✅ 調帳上傳檔已產生！\n\n"
                f"扣回：{count_cb} 筆\n"
                f"回補：{count_cs} 筆\n\n"
                f"儲存位置：\n{path_or_err}")
            folder = os.path.dirname(path_or_err)
            if folder and os.path.exists(folder) and sys.platform == "win32":
                os.startfile(folder)
        else:
            messagebox.showerror("錯誤", f"處理失敗：\n{path_or_err}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
