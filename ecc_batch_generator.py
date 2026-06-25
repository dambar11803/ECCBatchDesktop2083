"""
ECC Batch Generator Dashboard
NCHL-ECC System — Tkinter GUI (900×620)
- White dashboard, black text only, buttons keep colour
- Display section: 5 stat rows (Total, Accepted, Insufficient, Express, Unaccepted)
- Commission batch generation: Regular cheques > threshold
- XLS output: plain, no fills, no colours at all
- BRANCHCODE = first 3 digits of MAINCODE
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import os
import threading
from datetime import datetime

try:
    import pandas as pd
    import xlwt
    import xlrd
    DEPS_OK = True
except ImportError:
    DEPS_OK = False

WIN_W, WIN_H    = 900, 620
OUTPUT_ACCEPTED = f"accepted_{datetime.now().strftime('%Y%m%d')}.xls"
OUTPUT_EXPRESS  = f"express_{datetime.now().strftime('%Y%m%d')}.xls"

# Column index map from master report (0-based)
COL_MAP = {
    "S.N.": 1, "Cheque Sequence": 7, "Cheque Number": 8, "Cheque Date": 14,
    "Cheque Amount": 15, "BFD Bank": 22, "BFD Branch": 24, "BFD Account": 25,
    "PAY Bank Code": 27, "PAY Branch Code": 29, "PAY Account": 31,
    "Status": 34, "Session Date": 36, "Reason": 38, "Urgency": 40,
    "Instrument": 41, "Currency": 43, "BFD Customer Name": 44,
    "Pay Customer Name": 48, "Exp. Date": 54, "Cheque Reference": 55,
    "Posting Date": 58, "Scan Date": 59, "Presentment Date": 61
}

ACC_COLS  = ["BRANCHCODE", "MAINCODE", "TRANCODE", "AMOUNT",
             "LCYAMOUNT",  "DESC1",    "DESC2",    "DESC3"]

# ── colour palette ─────────────────────────────────────────────────────────
W      = "#FFFFFF"
K      = "#000000"
LG     = "#D9D9D9"
HDR_BG = "#1A237E"

BTN_BROWSE  = "#1565C0"
BTN_LOAD    = "#2E7D32"
BTN_CLEAR   = "#C62828"
BTN_ACCEPT  = "#1565C0"
BTN_EXPRESS = "#6A1B9A"
BTN_COMM    = "#00695C"

ACCENT = {
    "total":      "#1565C0",
    "accepted":   "#2E7D32",
    "insuf":      "#E65100",
    "express":    "#6A1B9A",
    "unaccepted": "#C62828",
}


class ECCDashboard(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("NCHL-ECC Batch Generator  |  Rastriya Banijya Bank Ltd.")
        self.geometry(f"{WIN_W}x{WIN_H}")
        self.resizable(False, False)
        self.configure(bg=W)

        self.master_path = tk.StringVar(value="")
        self.master_data = None
        self.status_msg  = tk.StringVar(
            value="Ready — browse and load the Master Report to begin.")

        # Cheque summary stats
        self.v_total_cnt    = tk.StringVar(value="—")
        self.v_total_amt    = tk.StringVar(value="—")
        self.v_accepted_cnt = tk.StringVar(value="—")
        self.v_accepted_amt = tk.StringVar(value="—")
        self.v_insuf_cnt    = tk.StringVar(value="—")
        self.v_insuf_amt    = tk.StringVar(value="—")
        self.v_express_cnt  = tk.StringVar(value="—")
        self.v_express_amt  = tk.StringVar(value="—")
        self.v_unaccept_cnt = tk.StringVar(value="—")
        self.v_unaccept_amt = tk.StringVar(value="—")

        # Commission settings
        self.v_branch_code    = tk.StringVar(value="255")
        self.v_parking_acc    = tk.StringVar(value="9313102000")
        self.v_regular_comm   = tk.StringVar(value="15")
        self.v_comm_threshold = tk.StringVar(value="200001")
        self.v_comm_acc       = tk.StringVar(value="9505062601")

        self._build_ui()
        if not DEPS_OK:
            self._set_status("⚠  pip install pandas xlwt xlrd")

    # ═══════════════════════════════════════════════════════════════════════
    #  UI BUILD
    # ═══════════════════════════════════════════════════════════════════════
    def _build_ui(self):

        # ── HEADER BAND ───────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=HDR_BG, height=48)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr,
                 text="NCHL-ECC  |  ECC Batch Generator — Cheque Exchange Dashboard",
                 font=("Segoe UI", 11, "bold"), bg=HDR_BG, fg=W).place(x=14, y=8)
        tk.Label(hdr, text="Rastriya Banijya Bank Ltd.",
                 font=("Segoe UI", 8), bg=HDR_BG, fg="#BBDEFB").place(x=14, y=30)
        self.clock_lbl = tk.Label(hdr, text="",
                 font=("Segoe UI", 10, "bold"), bg=HDR_BG, fg=W)
        self.clock_lbl.place(x=730, y=6)
        tk.Label(hdr, text=datetime.now().strftime("%d %b %Y"),
                 font=("Segoe UI", 8), bg=HDR_BG, fg="#BBDEFB").place(x=748, y=28)
        self._tick()

        # ── UPLOAD STRIP ──────────────────────────────────────────────────
        up = tk.Frame(self, bg=W, height=40)
        up.pack(fill="x", padx=14, pady=(6, 0))
        up.pack_propagate(False)
        tk.Label(up, text="Master Report (.xls):",
                 font=("Segoe UI", 9, "bold"), bg=W, fg=K).place(x=0, y=10)
        eb = tk.Frame(up, bg=LG)
        eb.place(x=158, y=8, width=420, height=24)
        self.path_entry = tk.Entry(
            eb, textvariable=self.master_path,
            font=("Consolas", 8), bg=W, fg=K,
            insertbackground=K, relief="flat", bd=0)
        self.path_entry.place(x=1, y=1, width=418, height=22)
        self._btn(up, "Browse", BTN_BROWSE, self._browse,  x=588, y=6, w=80, h=28)
        self._btn(up, "Load",   BTN_LOAD,   self._process, x=676, y=6, w=70, h=28)
        self._btn(up, "Clear",  BTN_CLEAR,  self._clear,   x=754, y=6, w=70, h=28)

        tk.Frame(self, bg=LG, height=1).pack(fill="x", padx=14, pady=(5, 0))

        # ── MAIN BODY: summary (left) + commission settings (right) ───────
        body = tk.Frame(self, bg=W)
        body.pack(fill="x", padx=14, pady=(8, 0))

        # ── LEFT: CHEQUE SUMMARY TABLE ────────────────────────────────────
        left = tk.Frame(body, bg=W)
        left.pack(side="left", anchor="n")

        tk.Label(left, text="Cheque Summary",
                 font=("Segoe UI", 8, "bold"), bg=W, fg="#1A237E").pack(anchor="w")

        tbl = tk.Frame(left, bg=LG)
        tbl.pack(anchor="w", pady=(3, 0))

        TABLE_W = 486
        COL_CAT = 200
        COL_CNT = 110
        COL_AMT = 176

        # column header row
        ch = tk.Frame(tbl, bg="#DDEEFF", height=22, width=TABLE_W)
        ch.pack(fill="x")
        ch.pack_propagate(False)
        tk.Label(ch, text="  Category",
                 font=("Segoe UI", 8, "bold"), bg="#DDEEFF", fg="#0D47A1",
                 anchor="w").place(x=4,   y=2, width=COL_CAT-4, height=18)
        tk.Label(ch, text="Cheques",
                 font=("Segoe UI", 8, "bold"), bg="#DDEEFF", fg="#0D47A1",
                 anchor="center").place(x=COL_CAT, y=2, width=COL_CNT, height=18)
        tk.Label(ch, text="Amount (NPR)",
                 font=("Segoe UI", 8, "bold"), bg="#DDEEFF", fg="#0D47A1",
                 anchor="e").place(x=COL_CAT+COL_CNT, y=2, width=COL_AMT-6, height=18)

        rows_cfg = [
            ("total",      "Total Cheques",      self.v_total_cnt,    self.v_total_amt),
            ("accepted",   "Accepted Cheques",   self.v_accepted_cnt, self.v_accepted_amt),
            ("insuf",      "Insufficient Fund",  self.v_insuf_cnt,    self.v_insuf_amt),
            ("express",    "Express Cheques",    self.v_express_cnt,  self.v_express_amt),
            ("unaccepted", "Unaccepted Cheques", self.v_unaccept_cnt, self.v_unaccept_amt),
        ]
        for key, label, cnt_v, amt_v in rows_cfg:
            rf = tk.Frame(tbl, bg=W, height=30, width=TABLE_W)
            rf.pack(fill="x")
            rf.pack_propagate(False)
            tk.Frame(rf, bg=ACCENT[key], width=4).place(
                x=0, y=0, width=4, height=30)
            tk.Label(rf, text=f"   {label}",
                     font=("Segoe UI", 9), bg=W, fg=K, anchor="w").place(
                         x=4, y=0, width=COL_CAT-4, height=30)
            tk.Label(rf, textvariable=cnt_v,
                     font=("Segoe UI", 9, "bold"), bg=W,
                     fg=ACCENT[key], anchor="center").place(
                         x=COL_CAT, y=0, width=COL_CNT, height=30)
            tk.Label(rf, textvariable=amt_v,
                     font=("Segoe UI", 9, "bold"), bg=W, fg=K, anchor="e").place(
                         x=COL_CAT+COL_CNT, y=0, width=COL_AMT-8, height=30)
            tk.Frame(tbl, bg=LG, height=1, width=TABLE_W).pack(fill="x")

        # ── RIGHT: COMMISSION SETTINGS ────────────────────────────────────
        right = tk.Frame(body, bg=W)
        right.pack(side="left", anchor="n", padx=(16, 0))

        tk.Label(right, text="Commission Settings",
                 font=("Segoe UI", 8, "bold"), bg=W, fg="#1A237E").pack(anchor="w")
        tk.Frame(right, bg=LG, height=1).pack(fill="x", pady=(3, 5))

        def _field(parent, label, var, tip=None, fw=220):
            row = tk.Frame(parent, bg=W)
            row.pack(anchor="w", pady=(0, 5))
            tk.Label(row, text=label,
                     font=("Segoe UI", 8, "bold"), bg=W, fg=K,
                     anchor="w").pack(anchor="w")
            eb = tk.Frame(row, bg=LG)
            eb.pack(anchor="w", pady=(2, 0))
            tk.Entry(eb, textvariable=var,
                     font=("Consolas", 9), bg=W, fg=K,
                     insertbackground=K, relief="flat", bd=0,
                     width=26).pack(padx=1, pady=1)
            if tip:
                tk.Label(row, text=tip,
                         font=("Segoe UI", 7), bg=W, fg="#777777",
                         anchor="w", wraplength=fw).pack(anchor="w")

        _field(right, "Branch Code",
               self.v_branch_code)
        _field(right, "Parking Account",
               self.v_parking_acc)
        _field(right, "Commission Account",
               self.v_comm_acc,
               tip="Account that receives the +15 commission credit")
        _field(right, "Regular Commission (Rs.)",
               self.v_regular_comm,
               tip="Flat Rs.15 charged per qualifying cheque")
        _field(right, "Commission Threshold (Rs.)",
               self.v_comm_threshold,
               tip="Cheques with amount ≥ this value are charged")

        tk.Frame(self, bg=LG, height=1).pack(fill="x", padx=14, pady=(6, 0))

        # ── ACTION BUTTONS ────────────────────────────────────────────────
        act = tk.Frame(self, bg=W, height=44)
        act.pack(fill="x", padx=14, pady=(6, 0))
        act.pack_propagate(False)
        tk.Label(act, text="Generate:",
                 font=("Segoe UI", 9, "bold"), bg=W, fg=K).place(x=0, y=10)
        self._btn(act, "Accepted Report (.xls)",  BTN_ACCEPT,
                  self._generate_accepted,  x=80,  y=5, w=210, h=34)
        self._btn(act, "Express Report (.xls)",   BTN_EXPRESS,
                  self._generate_express,   x=300, y=5, w=200, h=34)
        self._btn(act, "Calculate Commission",    BTN_COMM,
                  self._generate_commission, x=510, y=5, w=200, h=34)

        tk.Frame(self, bg=LG, height=1).pack(fill="x", padx=14, pady=(6, 0))

        # ── STATUS BAR ────────────────────────────────────────────────────
        sb = tk.Frame(self, bg=W, height=24)
        sb.pack(fill="x", side="bottom")
        sb.pack_propagate(False)
        tk.Frame(sb, bg=LG, height=1).pack(fill="x")
        self.status_lbl = tk.Label(
            sb, textvariable=self.status_msg,
            font=("Segoe UI", 8), bg=W, fg=K, anchor="w")
        self.status_lbl.place(x=10, y=4, width=570)
        tk.Label(sb,
                 text="All Rights Reserved@Dambar Rai(11803)",
                 font=("Segoe UI", 7), bg=W, fg="#888888",
                 anchor="e").place(x=586, y=5, width=300)

    # ── widget helper ─────────────────────────────────────────────────────
    def _btn(self, parent, text, color, cmd, x, y, w=120, h=30):
        tk.Button(
            parent, text=text, font=("Segoe UI", 9, "bold"),
            bg=color, fg=W, relief="flat", cursor="hand2",
            activebackground=color, bd=0,
            command=cmd).place(x=x, y=y, width=w, height=h)

    # ═══════════════════════════════════════════════════════════════════════
    #  DATA LOGIC
    # ═══════════════════════════════════════════════════════════════════════
    def _browse(self):
        p = filedialog.askopenfilename(
            title="Select Master Report (.xls)",
            filetypes=[("Excel 97-2003", "*.xls"), ("All files", "*.*")])
        if p:
            self.master_path.set(p)
            self._set_status(f"Selected: {os.path.basename(p)}")

    def _process(self):
        if not DEPS_OK:
            messagebox.showerror("Missing packages",
                                 "pip install pandas xlwt xlrd"); return
        p = self.master_path.get().strip()
        if not p or not os.path.exists(p):
            messagebox.showwarning("No file",
                "Please browse and select the Master Report first."); return
        threading.Thread(target=self._do_load, daemon=True).start()

    def _do_load(self):
        try:
            self._set_status("Loading…")
            df = pd.read_excel(self.master_path.get().strip(),
                               engine="xlrd", header=None)
            self.master_data = df
            self._compute_stats(df)
            self._set_status(
                f"Loaded — "
                f"Total: {self.v_total_cnt.get()}  |  "
                f"Accepted: {self.v_accepted_cnt.get()}  |  "
                f"Insuf: {self.v_insuf_cnt.get()}  |  "
                f"Express: {self.v_express_cnt.get()}  |  "
                f"Unaccepted: {self.v_unaccept_cnt.get()}")
        except Exception as e:
            self._set_status(f"Error: {e}")
            messagebox.showerror("Load error", str(e))

    def _get_data_rows(self, df):
        rows = []
        for _, row in df.iloc[15:].dropna(subset=[1]).iterrows():
            try:
                int(float(row[1]))
                rows.append(row)
            except Exception:
                pass
        return rows

    def _parse_amt(self, raw):
        try:
            return float(str(raw).replace(",", "").strip())
        except Exception:
            return 0.0

    def _gv(self, row, ci):
        """Safe cell getter — returns empty string for NaN."""
        v = row[ci] if ci < len(row) else ""
        try:
            if pd.isna(v):
                return ""
        except Exception:
            pass
        return str(v).strip()

    def _compute_stats(self, df):
        data = self._get_data_rows(df)
        total_cnt  = len(data)
        acc_cnt = insuf_cnt = exp_cnt = unacc_cnt = 0
        total_amt = acc_amt = insuf_amt = exp_amt = unacc_amt = 0.0

        for row in data:
            reason  = str(row[38]).strip()
            urgency = str(row[40]).strip()
            amt     = self._parse_amt(row[15])
            total_amt += amt

            if urgency == "Express":
                exp_cnt += 1
                exp_amt += amt

            if reason == "0000-ACCEPTED":
                acc_cnt += 1;   acc_amt   += amt
            elif "Insufficient Fund" in reason:
                insuf_cnt += 1; insuf_amt += amt
            else:
                unacc_cnt += 1; unacc_amt += amt

        def _u():
            self.v_total_cnt.set(str(total_cnt))
            self.v_total_amt.set(f"{total_amt:,.2f}")
            self.v_accepted_cnt.set(str(acc_cnt))
            self.v_accepted_amt.set(f"{acc_amt:,.2f}")
            self.v_insuf_cnt.set(str(insuf_cnt))
            self.v_insuf_amt.set(f"{insuf_amt:,.2f}")
            self.v_express_cnt.set(str(exp_cnt))
            self.v_express_amt.set(f"{exp_amt:,.2f}")
            self.v_unaccept_cnt.set(str(unacc_cnt))
            self.v_unaccept_amt.set(f"{unacc_amt:,.2f}")
        self.after(0, _u)

    # ═══════════════════════════════════════════════════════════════════════
    #  ACCEPTED REPORT
    # ═══════════════════════════════════════════════════════════════════════
    def _generate_accepted(self):
        if not DEPS_OK:
            messagebox.showerror("Missing packages",
                                 "pip install pandas xlwt xlrd"); return
        if self.master_data is None:
            messagebox.showwarning("No data",
                                   "Load the Master Report first."); return
        p = filedialog.asksaveasfilename(
            title="Save Accepted Report", initialfile=OUTPUT_ACCEPTED,
            defaultextension=".xls",
            filetypes=[("Excel 97-2003", "*.xls")])
        if p:
            threading.Thread(target=self._do_accepted,
                             args=(p,), daemon=True).start()

    def _do_accepted(self, path):
        try:
            self._set_status("Building Accepted Report…")
            filtered = []
            for row in self._get_data_rows(self.master_data):
                if (str(row[40]).strip() == "Regular" and
                        str(row[38]).strip() == "0000-ACCEPTED"):
                    maincode = self._gv(row, 25)
                    amt_raw  = self._gv(row, 15)
                    filtered.append({
                        "BRANCHCODE": maincode[:3],
                        "MAINCODE":   maincode,
                        "TRANCODE":   "555",
                        "AMOUNT":     amt_raw,
                        "LCYAMOUNT":  amt_raw,
                        "DESC1":      self._gv(row, 48),
                        "DESC2":      f"{self._gv(row,27)} {self._gv(row,8)}",
                        "DESC3":      self._gv(row, 31),
                        "_num":       self._parse_amt(amt_raw),
                    })
            filtered.sort(key=lambda x: x["_num"])
            wb = xlwt.Workbook(encoding="utf-8")
            ws = wb.add_sheet("Accepted")
            st = self._xls_styles()
            col_widths = [4000, 7000, 3500, 5000, 5000, 10000, 6000, 8000]
            for ci, col in enumerate(ACC_COLS):
                ws.write(0, ci, col, st["hdr"])
                ws.col(ci).width = col_widths[ci]
            for ri, rec in enumerate(filtered, start=1):
                for ci, key in enumerate(ACC_COLS):
                    ws.write(ri, ci, rec[key], st["dat"])
            wb.save(path)
            kb = os.path.getsize(path) // 1024
            self._set_status(
                f"Accepted Report saved — {len(filtered)} records  |  "
                f"{os.path.basename(path)}  ({kb} KB)")
            messagebox.showinfo("Done",
                f"Accepted Report saved!\n\n{path}\n\n"
                f"Records : {len(filtered)}  (Regular + 0000-ACCEPTED)\n"
                f"File size : {kb} KB")
        except Exception as e:
            self._set_status(f"Export error: {e}")
            messagebox.showerror("Export error", str(e))

    # ═══════════════════════════════════════════════════════════════════════
    #  EXPRESS REPORT
    # ═══════════════════════════════════════════════════════════════════════
    def _generate_express(self):
        if not DEPS_OK:
            messagebox.showerror("Missing packages",
                                 "pip install pandas xlwt xlrd"); return
        if self.master_data is None:
            messagebox.showwarning("No data",
                                   "Load the Master Report first."); return
        p = filedialog.asksaveasfilename(
            title="Save Express Report", initialfile=OUTPUT_EXPRESS,
            defaultextension=".xls",
            filetypes=[("Excel 97-2003", "*.xls")])
        if p:
            threading.Thread(target=self._do_express,
                             args=(p,), daemon=True).start()

    def _do_express(self, path):
        try:
            self._set_status("Building Express Report…")
            filtered = []
            for row in self._get_data_rows(self.master_data):
                if str(row[40]).strip() == "Express":
                    maincode = self._gv(row, 25)
                    amt_raw  = self._gv(row, 15)
                    filtered.append({
                        "BRANCHCODE": maincode[:3],
                        "MAINCODE":   maincode,
                        "TRANCODE":   "555",
                        "AMOUNT":     amt_raw,
                        "LCYAMOUNT":  amt_raw,
                        "DESC1":      self._gv(row, 48),
                        "DESC2":      f"{self._gv(row,27)} {self._gv(row,8)}",
                        "DESC3":      self._gv(row, 31),
                        "_num":       self._parse_amt(amt_raw),
                    })
            filtered.sort(key=lambda x: x["_num"])
            wb = xlwt.Workbook(encoding="utf-8")
            ws = wb.add_sheet("Express")
            st = self._xls_styles()
            col_widths = [4000, 7000, 3500, 5000, 5000, 10000, 6000, 8000]
            for ci, col in enumerate(ACC_COLS):
                ws.write(0, ci, col, st["hdr"])
                ws.col(ci).width = col_widths[ci]
            for ri, rec in enumerate(filtered, start=1):
                for ci, key in enumerate(ACC_COLS):
                    ws.write(ri, ci, rec[key], st["dat"])
            wb.save(path)
            kb = os.path.getsize(path) // 1024
            self._set_status(
                f"Express Report saved — {len(filtered)} records  |  "
                f"{os.path.basename(path)}  ({kb} KB)")
            messagebox.showinfo("Done",
                f"Express Report saved!\n\n{path}\n\n"
                f"Records : {len(filtered)}  (Urgency = Express)\n"
                f"File size : {kb} KB")
        except Exception as e:
            self._set_status(f"Export error: {e}")
            messagebox.showerror("Export error", str(e))

    # ═══════════════════════════════════════════════════════════════════════
    #  CALCULATE COMMISSION
    # ═══════════════════════════════════════════════════════════════════════
    def _generate_commission(self):
        if not DEPS_OK:
            messagebox.showerror("Missing packages",
                                 "pip install pandas xlwt xlrd"); return
        if self.master_data is None:
            messagebox.showwarning("No data",
                                   "Load the Master Report first."); return

        # Validate settings fields
        try:
            threshold = float(self.v_comm_threshold.get().strip())
            comm_amt  = float(self.v_regular_comm.get().strip())
        except ValueError:
            messagebox.showerror("Invalid input",
                "Commission Threshold and Regular Commission must be numbers.")
            return

        branch_code = self.v_branch_code.get().strip()
        comm_acc    = self.v_comm_acc.get().strip()

        if not branch_code or not comm_acc:
            messagebox.showerror("Invalid input",
                "Branch Code and Commission Account must not be empty.")
            return

        fname = f"commission_{datetime.now().strftime('%Y%m%d')}.xls"
        p = filedialog.asksaveasfilename(
            title="Save Commission Batch", initialfile=fname,
            defaultextension=".xls",
            filetypes=[("Excel 97-2003", "*.xls")])
        if p:
            threading.Thread(
                target=self._do_commission,
                args=(p, threshold, comm_amt, branch_code, comm_acc),
                daemon=True).start()

    def _do_commission(self, path, threshold, comm_amt, branch_code, comm_acc):
        """
        Generate commission batch.

        For each Regular cheque with amount >= threshold (any reason):

          Row A  — debit customer (negative):
            BRANCHCODE  = first 3 digits of BFD Account (col 25)
            MAINCODE    = BFD Account (col 25)
            TRANCODE    = 055
            AMOUNT      = -comm_amt  (e.g. -15)
            LCYAMOUNT   = -comm_amt
            DESC1       = "ECC Charge for Rs.{cheque_amount}"
            DESC2       = ""
            DESC3       = ""

          Row B  — credit commission account (positive):
            BRANCHCODE  = branch_code field value (default 255)
            MAINCODE    = comm_acc field value     (default 9505062601)
            TRANCODE    = 555
            AMOUNT      = +comm_amt  (e.g. 15)
            LCYAMOUNT   = +comm_amt
            DESC1       = BFD Customer Name (col 44)
            DESC2       = "ECC Charge for Rs.{cheque_amount}"
            DESC3       = BFD Account (col 25)
        """
        try:
            self._set_status("Building Commission Batch…")

            neg_rows = []   # debit rows  (all together first)
            pos_rows = []   # credit rows (all together after)

            comm_str = str(int(comm_amt)) if comm_amt == int(comm_amt) \
                       else str(comm_amt)

            for row in self._get_data_rows(self.master_data):
                urgency = str(row[40]).strip()

                # Only Regular cheques (any reason)
                if urgency != "Regular":
                    continue

                raw_amt = self._parse_amt(row[15])
                if raw_amt < threshold:
                    continue

                bfd_acc  = self._gv(row, 25)   # BFD Account (col 25)
                bfd_cust = self._gv(row, 44)   # BFD Customer Name (col 44)

                # Format cheque amount — drop .0 if whole number
                disp_amt = str(int(raw_amt)) if raw_amt == int(raw_amt) \
                           else str(raw_amt)
                desc_charge = f"ECC Charge for Rs.{disp_amt}"

                # Row A: debit (charge the BFD account holder)
                neg_rows.append({
                    "BRANCHCODE": bfd_acc[:3],
                    "MAINCODE":   bfd_acc,
                    "TRANCODE":   "055",
                    "AMOUNT":     -comm_amt,
                    "LCYAMOUNT":  -comm_amt,
                    "DESC1":      desc_charge,
                    "DESC2":      "",
                    "DESC3":      "",
                })

                # Row B: credit commission account
                pos_rows.append({
                    "BRANCHCODE": branch_code,
                    "MAINCODE":   comm_acc,
                    "TRANCODE":   "555",
                    "AMOUNT":     comm_amt,
                    "LCYAMOUNT":  comm_amt,
                    "DESC1":      bfd_cust,
                    "DESC2":      desc_charge,
                    "DESC3":      bfd_acc,
                })

            total_records = len(neg_rows) + len(pos_rows)

            if total_records == 0:
                self._set_status(
                    "Commission Batch: no qualifying cheques found "
                    f"(Regular + Accepted + amount ≥ {threshold:,.0f})")
                messagebox.showinfo("No records",
                    f"No Regular Accepted cheques found with amount ≥ "
                    f"Rs.{threshold:,.0f}.\n\nNo file was saved.")
                return

            wb = xlwt.Workbook(encoding="utf-8")
            ws = wb.add_sheet("Commission")
            st = self._xls_styles()

            col_widths = [4000, 7000, 3500, 5000, 5000, 12000, 9000, 9000]
            for ci, col in enumerate(ACC_COLS):
                ws.write(0, ci, col, st["hdr"])
                ws.col(ci).width = col_widths[ci]

            ri = 1
            # Write all debit rows first, then all credit rows
            for rec in neg_rows:
                for ci, key in enumerate(ACC_COLS):
                    val = rec[key]
                    # write numeric amounts as numbers, rest as strings
                    if key in ("AMOUNT", "LCYAMOUNT"):
                        ws.write(ri, ci, val, st["dat"])
                    else:
                        ws.write(ri, ci, str(val) if val != "" else "", st["dat"])
                ri += 1

            for rec in pos_rows:
                for ci, key in enumerate(ACC_COLS):
                    val = rec[key]
                    if key in ("AMOUNT", "LCYAMOUNT"):
                        ws.write(ri, ci, val, st["dat"])
                    else:
                        ws.write(ri, ci, str(val) if val != "" else "", st["dat"])
                ri += 1

            wb.save(path)
            kb = os.path.getsize(path) // 1024
            cheque_count = len(neg_rows)
            self._set_status(
                f"Commission Batch saved — {cheque_count} cheques  |  "
                f"{total_records} rows  |  "
                f"{os.path.basename(path)}  ({kb} KB)")
            messagebox.showinfo("Done",
                f"Commission Batch saved!\n\n{path}\n\n"
                f"Qualifying cheques : {cheque_count}\n"
                f"Total rows written : {total_records} "
                f"({cheque_count} debit + {cheque_count} credit)\n"
                f"Threshold          : Rs.{threshold:,.0f}\n"
                f"Commission/cheque  : Rs.{comm_amt:,.0f}\n"
                f"File size          : {kb} KB")
        except Exception as e:
            self._set_status(f"Commission export error: {e}")
            messagebox.showerror("Export error", str(e))

    # ═══════════════════════════════════════════════════════════════════════
    #  XLS STYLES
    # ═══════════════════════════════════════════════════════════════════════
    def _xls_styles(self):
        def _s(bold=False, height=180, border=True):
            s = xlwt.XFStyle()
            f = xlwt.Font()
            f.bold = bold; f.height = height; f.colour_index = 0x08
            s.font = f
            a = xlwt.Alignment()
            a.horz = xlwt.Alignment.HORZ_LEFT
            s.alignment = a
            if border:
                b = xlwt.Borders()
                b.left = b.right = b.top = b.bottom = xlwt.Borders.THIN
                s.borders = b
            return s
        return {
            "hdr": _s(bold=True, height=180, border=True),
            "dat": _s(height=180,            border=True),
        }

    # ═══════════════════════════════════════════════════════════════════════
    #  UTILITIES
    # ═══════════════════════════════════════════════════════════════════════
    def _set_status(self, msg):
        self.after(0, lambda: (
            self.status_msg.set(msg),
            self.status_lbl.configure(fg=K)
        ))

    def _tick(self):
        self.clock_lbl.configure(text=datetime.now().strftime("%H:%M:%S"))
        self.after(1000, self._tick)

    def _clear(self):
        self.master_path.set("")
        self.master_data = None
        for v in (self.v_total_cnt,    self.v_total_amt,
                  self.v_accepted_cnt, self.v_accepted_amt,
                  self.v_insuf_cnt,    self.v_insuf_amt,
                  self.v_express_cnt,  self.v_express_amt,
                  self.v_unaccept_cnt, self.v_unaccept_amt):
            v.set("—")
        self._set_status("Cleared — ready for a new file.")


if __name__ == "__main__":
    app = ECCDashboard()
    app.mainloop()