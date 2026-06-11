# -*- coding: utf-8 -*-
"""Lab WhatsApp Sender - sends personalized lab reports to dialysis patients via WhatsApp Web."""
import os, sys, json, csv, time, random, threading, queue, re, urllib.parse, datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import pandas as pd

APP_DIR = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
SETTINGS_FILE = os.path.join(APP_DIR, "settings.json")
LOG_FILE = os.path.join(APP_DIR, "sent_log.csv")
WA_PROFILE = os.path.join(APP_DIR, "wa_profile")

DEFAULT_PARAMS = [
    {"name": "Calcium", "ar": "الكالسيوم", "type": "range", "low": "", "high": "", "unit": "mg/dl"},
    {"name": "Corrected Calcium", "ar": "الكالسيوم المصحح", "type": "range", "low": "", "high": "", "unit": "mg/dl"},
    {"name": "Albumin", "ar": "الألبومين", "type": "range", "low": "", "high": "", "unit": "g/dl"},
    {"name": "Phosphorus", "ar": "الفوسفور", "type": "range", "low": "", "high": "", "unit": "mg/dl"},
    {"name": "Hemoglobin", "ar": "الهيموجلوبين", "type": "range", "low": "", "high": "", "unit": "g/dl"},
    {"name": "Ferritin", "ar": "الفيريتين", "type": "range", "low": "", "high": "", "unit": "ng/ml"},
    {"name": "Transferrin Saturation", "ar": "تشبع الترانسفرين", "type": "range", "low": "", "high": "", "unit": "%"},
    {"name": "Potassium", "ar": "البوتاسيوم", "type": "range", "low": "", "high": "", "unit": "mmol/l"},
    {"name": "Kt/V", "ar": "كفاءة الغسيل Kt/V", "type": "min_only", "low": "", "high": "", "unit": ""},
    {"name": "Blood Urea Nitrogen", "ar": "يوريا الدم", "type": "range", "low": "", "high": "", "unit": "mg/dl"},
    {"name": "BUN Post Dialysis", "ar": "يوريا ما بعد الغسيل", "type": "range", "low": "", "high": "", "unit": "mg/dl"},
]

DEFAULT_TEMPLATE = (
    "السلام عليكم ورحمة الله وبركاته\n"
    "المريض: {Name}\n"
    "نتائج التحاليل بتاريخ {Date}:\n\n"
    "{LABS}\n\n"
    "{RECS}"
    "{NOTE}"
    "\nنتمنى لكم دوام الصحة والعافية"
)

FLAG_HIGH = "مرتفع ↑"
FLAG_LOW = "منخفض ↓"
FLAG_NORMAL = "طبيعي"


def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                s = json.load(f)
            for k, v in default_settings().items():
                s.setdefault(k, v)
            return s
        except Exception:
            pass
    return default_settings()


def default_settings():
    return {
        "country_code": "966",
        "gap_min": 120,
        "gap_max": 360,
        "max_per_session": 90,
        "parameters": [dict(p) for p in DEFAULT_PARAMS],
        "recommendations": {},
        "template": DEFAULT_TEMPLATE,
    }


def save_settings(s):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(s, f, ensure_ascii=False, indent=2)


def fmt_num(v):
    try:
        f = float(v)
        if f == int(f):
            return str(int(f))
        return str(round(f, 2))
    except Exception:
        return str(v)


def normalize_mobile(raw, cc):
    digits = re.sub(r"\D", "", str(raw))
    if not digits:
        return ""
    if digits.startswith("00"):
        digits = digits[2:]
    if digits.startswith(cc):
        return digits
    if digits.startswith("0"):
        return cc + digits[1:]
    if len(digits) == 9 and digits.startswith("5") and cc == "966":
        return cc + digits
    return cc + digits if len(digits) <= 9 else digits


def classify(value, param):
    try:
        v = float(value)
    except Exception:
        return None
    t = param.get("type", "range")
    low = param.get("low", "")
    high = param.get("high", "")
    try:
        low = float(low) if str(low).strip() != "" else None
    except Exception:
        low = None
    try:
        high = float(high) if str(high).strip() != "" else None
    except Exception:
        high = None
    if t == "min_only":
        if low is not None and v < low:
            return "low"
        return "normal"
    if t == "max_only":
        if high is not None and v > high:
            return "high"
        return "normal"
    if low is not None and v < low:
        return "low"
    if high is not None and v > high:
        return "high"
    if low is None and high is None:
        return None
    return "normal"


FLAG_TEXT = {"high": FLAG_HIGH, "low": FLAG_LOW, "normal": FLAG_NORMAL}


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Lab WhatsApp Sender")
        self.geometry("1050x720")
        ctk.set_appearance_mode("light")
        self.settings = load_settings()
        self.patients = []          # list of dicts: id, name, mobile, note, labs{param:value}, date
        self.lab_path = tk.StringVar()
        self.mob_path = tk.StringVar()
        self.sending = False
        self.stop_flag = False
        self.msg_queue = queue.Queue()

        self.tabs = ctk.CTkTabview(self)
        self.tabs.pack(fill="both", expand=True, padx=8, pady=8)
        for t in ["1. Files", "2. Ranges", "3. Recommendations", "4. Template", "5. Send"]:
            self.tabs.add(t)
        self.build_files_tab()
        self.build_ranges_tab()
        self.build_recs_tab()
        self.build_template_tab()
        self.build_send_tab()
        self.after(200, self.poll_queue)

    # ---------------- Tab 1: Files ----------------
    def build_files_tab(self):
        f = self.tabs.tab("1. Files")
        row1 = ctk.CTkFrame(f); row1.pack(fill="x", pady=6, padx=6)
        ctk.CTkLabel(row1, text="Lab results file (Excel):", width=200, anchor="w").pack(side="left", padx=4)
        ctk.CTkEntry(row1, textvariable=self.lab_path, width=520).pack(side="left", padx=4)
        ctk.CTkButton(row1, text="Browse", width=90, command=lambda: self.pick(self.lab_path)).pack(side="left")

        row2 = ctk.CTkFrame(f); row2.pack(fill="x", pady=6, padx=6)
        ctk.CTkLabel(row2, text="Mobile numbers file (Excel):", width=200, anchor="w").pack(side="left", padx=4)
        ctk.CTkEntry(row2, textvariable=self.mob_path, width=520).pack(side="left", padx=4)
        ctk.CTkButton(row2, text="Browse", width=90, command=lambda: self.pick(self.mob_path)).pack(side="left")

        row3 = ctk.CTkFrame(f); row3.pack(fill="x", pady=6, padx=6)
        ctk.CTkLabel(row3, text="Country code:", width=200, anchor="w").pack(side="left", padx=4)
        self.cc_entry = ctk.CTkEntry(row3, width=80)
        self.cc_entry.insert(0, self.settings.get("country_code", "966"))
        self.cc_entry.pack(side="left", padx=4)
        ctk.CTkButton(row3, text="Load && Match", width=140, command=self.load_data).pack(side="left", padx=16)

        self.files_info = ctk.CTkTextbox(f, height=420, font=("Segoe UI", 13))
        self.files_info.pack(fill="both", expand=True, padx=6, pady=6)
        self.files_info.insert("end", "Steps:\n1) Choose the lab results Excel file (the system export, one lab per row).\n"
            "2) Choose the mobile numbers Excel file. Required columns: ID , mobile . Optional column: note (physician note per patient).\n"
            "3) Press Load & Match.\n")

    def pick(self, var):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if p:
            var.set(p)

    def load_data(self):
        self.settings["country_code"] = self.cc_entry.get().strip() or "966"
        save_settings(self.settings)
        cc = self.settings["country_code"]
        out = self.files_info
        out.delete("1.0", "end")
        try:
            lab = pd.read_excel(self.lab_path.get())
        except Exception as e:
            messagebox.showerror("Error", f"Cannot read lab file:\n{e}")
            return
        need = {"ID", "Name", "Lab name", "Value"}
        missing = need - set(c.strip() for c in lab.columns.astype(str))
        lab.columns = [str(c).strip() for c in lab.columns]
        if missing:
            messagebox.showerror("Error", f"Lab file is missing columns: {', '.join(missing)}")
            return
        try:
            mob = pd.read_excel(self.mob_path.get())
        except Exception as e:
            messagebox.showerror("Error", f"Cannot read mobile file:\n{e}")
            return
        mob.columns = [str(c).strip().lower() for c in mob.columns]
        if "id" not in mob.columns or "mobile" not in mob.columns:
            messagebox.showerror("Missing column",
                "The mobile file must contain columns named exactly:\n\nID\nmobile\n\nPlease add them and load again.")
            return
        has_note = "note" in mob.columns
        if not has_note:
            ok = messagebox.askyesno("Missing column: note",
                "The mobile file does not contain a column named:  note\n\n"
                "This column is used for the physician note added at the end of each message.\n"
                "If you want notes, add a column named  note  to the file (leave it empty for patients with no note), then load again.\n\n"
                "Continue WITHOUT notes?")
            if not ok:
                return

        def cid(x):
            s = str(x).strip()
            if s.endswith(".0"):
                s = s[:-2]
            return re.sub(r"\D", "", s)

        mob_map, note_map = {}, {}
        for _, r in mob.iterrows():
            pid = cid(r["id"])
            nums = [normalize_mobile(x, cc) for x in re.split(r"[,/;\n]+", str(r["mobile"])) if re.sub(r"\D", "", str(x))]
            nums = [n for n in nums if n]
            mob_map.setdefault(pid, [])
            for n in nums:
                if n not in mob_map[pid]:
                    mob_map[pid].append(n)
            if has_note and pd.notna(r.get("note")) and str(r.get("note")).strip():
                note_map[pid] = str(r["note"]).strip()

        lab["ID"] = lab["ID"].map(cid)
        lab["Lab name"] = lab["Lab name"].astype(str).str.strip()
        patients = {}
        for _, r in lab.iterrows():
            pid = r["ID"]
            p = patients.setdefault(pid, {"id": pid, "name": "", "labs": {}, "date": "", "mobile": "", "note": ""})
            nm = str(r["Name"])
            p["name"] = re.sub(r"\s*\(\d+\)\s*$", "", nm).strip()
            p["labs"][r["Lab name"]] = r["Value"]
            d = r.get("Lab date")
            if pd.notna(d):
                if isinstance(d, (int, float)):
                    d = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(d))
                try:
                    p["date"] = pd.Timestamp(d).strftime("%Y-%m-%d")
                except Exception:
                    p["date"] = str(d)

        no_mobile = []
        expanded = []
        for pid, p in patients.items():
            p["note"] = note_map.get(pid, "")
            nums = mob_map.get(pid, [])
            if not nums:
                no_mobile.append(p["name"] or pid)
                continue
            for j, n in enumerate(nums, 1):
                e = dict(p)
                e["mobile"] = n
                e["slot"] = f"{j}/{len(nums)}" if len(nums) > 1 else ""
                expanded.append(e)
        self.patients = expanded
        self.patients.sort(key=lambda p: (p["name"], p["mobile"]))
        multi = sum(1 for p in self.patients if p["slot"])

        known = {pp["name"] for pp in self.settings["parameters"]}
        new_labs = sorted({l for p in patients.values() for l in p["labs"]} - known)
        for nl in new_labs:
            self.settings["parameters"].append({"name": nl, "ar": "", "type": "range", "low": "", "high": "", "unit": ""})
        if new_labs:
            save_settings(self.settings)
            self.refresh_ranges()
            self.refresh_chips()
            self.refresh_rec_params()

        out.insert("end", f"Patients in lab file: {len(patients)}\n")
        out.insert("end", f"Message slots (one per number): {len(self.patients)}\n")
        out.insert("end", f"Numbers belonging to multi-number patients: {multi}\n")
        out.insert("end", f"Patients with notes: {len({p['id'] for p in self.patients if p['note']})}\n")
        if new_labs:
            out.insert("end", f"New lab parameters detected and added to Ranges: {', '.join(new_labs)}\n")
        if no_mobile:
            out.insert("end", f"\nNO MOBILE FOUND ({len(no_mobile)}) - these will be SKIPPED:\n")
            for n in no_mobile:
                out.insert("end", f"   - {n}\n")
        out.insert("end", "\nPreview (name | normalized number | labs | note?):\n")
        for p in self.patients[:400]:
            sl = f" [{p['slot']}]" if p.get("slot") else ""
            out.insert("end", f"   {p['name']}{sl} | {p['mobile']} | {len(p['labs'])} labs | {'note' if p['note'] else '-'}\n")
        self.refresh_send_list()

    # ---------------- Tab 2: Ranges ----------------
    def build_ranges_tab(self):
        f = self.tabs.tab("2. Ranges")
        top = ctk.CTkFrame(f); top.pack(fill="x", padx=6, pady=4)
        ctk.CTkButton(top, text="+ Add parameter", command=self.add_param).pack(side="left", padx=4)
        ctk.CTkButton(top, text="Save ranges", command=self.save_ranges).pack(side="left", padx=4)
        hdr = ctk.CTkFrame(f); hdr.pack(fill="x", padx=6)
        for txt, w in [("Parameter", 190), ("Arabic name (optional)", 190), ("Type", 150), ("Low", 80), ("High", 80), ("Unit", 90), ("", 60)]:
            ctk.CTkLabel(hdr, text=txt, width=w, anchor="w").pack(side="left", padx=3)
        self.ranges_frame = ctk.CTkScrollableFrame(f, height=520)
        self.ranges_frame.pack(fill="both", expand=True, padx=6, pady=4)
        self.range_rows = []
        self.refresh_ranges()

    def refresh_ranges(self):
        for w in self.ranges_frame.winfo_children():
            w.destroy()
        self.range_rows = []
        for i, p in enumerate(self.settings["parameters"]):
            row = ctk.CTkFrame(self.ranges_frame)
            row.pack(fill="x", pady=2)
            e_name = ctk.CTkEntry(row, width=190); e_name.insert(0, p["name"]); e_name.pack(side="left", padx=3)
            e_ar = ctk.CTkEntry(row, width=190, justify="right"); e_ar.insert(0, p.get("ar", "")); e_ar.pack(side="left", padx=3)
            type_map = {"range": "Range (low-high)", "min_only": "Keep ABOVE low", "max_only": "Keep BELOW high"}
            cb = ctk.CTkComboBox(row, width=150, values=list(type_map.values()))
            cb.set(type_map.get(p.get("type", "range")))
            cb.pack(side="left", padx=3)
            e_low = ctk.CTkEntry(row, width=80); e_low.insert(0, str(p.get("low", ""))); e_low.pack(side="left", padx=3)
            e_high = ctk.CTkEntry(row, width=80); e_high.insert(0, str(p.get("high", ""))); e_high.pack(side="left", padx=3)
            e_unit = ctk.CTkEntry(row, width=90); e_unit.insert(0, p.get("unit", "")); e_unit.pack(side="left", padx=3)
            ctk.CTkButton(row, text="X", width=40, fg_color="#b33", command=lambda idx=i: self.del_param(idx)).pack(side="left", padx=3)
            self.range_rows.append((e_name, e_ar, cb, e_low, e_high, e_unit))

    def add_param(self):
        self.save_ranges(silent=True)
        self.settings["parameters"].append({"name": "New Parameter", "ar": "", "type": "range", "low": "", "high": "", "unit": ""})
        self.refresh_ranges()

    def del_param(self, idx):
        if messagebox.askyesno("Delete", f"Delete parameter '{self.settings['parameters'][idx]['name']}'?"):
            self.settings["parameters"].pop(idx)
            save_settings(self.settings)
            self.refresh_ranges(); self.refresh_chips(); self.refresh_rec_params()

    def save_ranges(self, silent=False):
        inv_map = {"Range (low-high)": "range", "Keep ABOVE low": "min_only", "Keep BELOW high": "max_only"}
        params = []
        for e_name, e_ar, cb, e_low, e_high, e_unit in self.range_rows:
            params.append({"name": e_name.get().strip(), "ar": e_ar.get().strip(),
                           "type": inv_map.get(cb.get(), "range"),
                           "low": e_low.get().strip(), "high": e_high.get().strip(),
                           "unit": e_unit.get().strip()})
        self.settings["parameters"] = [p for p in params if p["name"]]
        save_settings(self.settings)
        self.refresh_chips(); self.refresh_rec_params()
        if not silent:
            messagebox.showinfo("Saved", "Ranges saved.")

    # ---------------- Tab 3: Recommendations ----------------
    def build_recs_tab(self):
        f = self.tabs.tab("3. Recommendations")
        top = ctk.CTkFrame(f); top.pack(fill="x", padx=6, pady=6)
        ctk.CTkLabel(top, text="Parameter:").pack(side="left", padx=4)
        self.rec_param = ctk.CTkComboBox(top, width=240, values=[p["name"] for p in self.settings["parameters"]], command=lambda _: self.load_rec())
        self.rec_param.pack(side="left", padx=4)
        ctk.CTkLabel(top, text="When value is:").pack(side="left", padx=4)
        self.rec_dir = ctk.CTkComboBox(top, width=140, values=["high (مرتفع)", "low (منخفض)"], command=lambda _: self.load_rec())
        self.rec_dir.pack(side="left", padx=4)
        ctk.CTkButton(top, text="Save recommendation", command=self.save_rec).pack(side="left", padx=12)
        ctk.CTkLabel(f, text="Recommendation text (Arabic supported) - sent only when this parameter is abnormal in this direction:", anchor="w").pack(fill="x", padx=8)
        self.rec_text = ctk.CTkTextbox(f, height=180, font=("Segoe UI", 14))
        self.rec_text.pack(fill="x", padx=8, pady=4)
        ctk.CTkLabel(f, text="Saved recommendations:", anchor="w").pack(fill="x", padx=8)
        self.rec_list = ctk.CTkTextbox(f, height=260, font=("Segoe UI", 13))
        self.rec_list.pack(fill="both", expand=True, padx=8, pady=4)
        self.refresh_rec_list()
        self.load_rec()

    def rec_key(self):
        d = "high" if self.rec_dir.get().startswith("high") else "low"
        return f"{self.rec_param.get()}|{d}"

    def load_rec(self):
        self.rec_text.delete("1.0", "end")
        self.rec_text.insert("end", self.settings["recommendations"].get(self.rec_key(), ""))

    def save_rec(self):
        txt = self.rec_text.get("1.0", "end").strip()
        if txt:
            self.settings["recommendations"][self.rec_key()] = txt
        else:
            self.settings["recommendations"].pop(self.rec_key(), None)
        save_settings(self.settings)
        self.refresh_rec_list()
        messagebox.showinfo("Saved", "Recommendation saved.")

    def refresh_rec_list(self):
        self.rec_list.delete("1.0", "end")
        for k, v in self.settings["recommendations"].items():
            name, d = k.split("|")
            self.rec_list.insert("end", f"• {name} - {FLAG_TEXT[d]}:\n{v}\n\n")

    def refresh_rec_params(self):
        try:
            self.rec_param.configure(values=[p["name"] for p in self.settings["parameters"]])
        except Exception:
            pass

    # ---------------- Tab 4: Template ----------------
    def build_template_tab(self):
        f = self.tabs.tab("4. Template")
        left = ctk.CTkFrame(f); left.pack(side="left", fill="both", expand=True, padx=6, pady=6)
        right = ctk.CTkFrame(f, width=260); right.pack(side="right", fill="y", padx=6, pady=6)
        ctk.CTkLabel(left, text="Message template (Arabic supported). Click a chip on the right to insert a placeholder at the cursor:", anchor="w", wraplength=600).pack(fill="x")
        self.tpl_text = ctk.CTkTextbox(left, font=("Segoe UI", 15))
        self.tpl_text.pack(fill="both", expand=True, pady=4)
        self.tpl_text.insert("end", self.settings.get("template", DEFAULT_TEMPLATE))
        btns = ctk.CTkFrame(left); btns.pack(fill="x")
        ctk.CTkButton(btns, text="Save template", command=self.save_template).pack(side="left", padx=4, pady=4)
        ctk.CTkButton(btns, text="Preview first patient", command=self.preview_msg).pack(side="left", padx=4, pady=4)
        ctk.CTkLabel(right, text="Insert parameter:").pack(pady=4)
        self.chips_frame = ctk.CTkScrollableFrame(right, width=240, height=520)
        self.chips_frame.pack(fill="both", expand=True)
        self.refresh_chips()

    def refresh_chips(self):
        if not hasattr(self, "chips_frame"):
            return
        for w in self.chips_frame.winfo_children():
            w.destroy()
        base = [("{Name}", "Patient name"), ("{ID}", "Patient ID"), ("{Date}", "Lab date"),
                ("{LABS}", "All labs list"), ("{RECS}", "Recommendations"), ("{NOTE}", "Physician note")]
        for ph, desc in base:
            ctk.CTkButton(self.chips_frame, text=f"{ph}  ({desc})", anchor="w",
                          command=lambda p=ph: self.tpl_text.insert("insert", p)).pack(fill="x", pady=2)
        ctk.CTkLabel(self.chips_frame, text="— single lab values —").pack(pady=4)
        for p in self.settings["parameters"]:
            ph = "{" + p["name"] + "}"
            ctk.CTkButton(self.chips_frame, text=ph, anchor="w",
                          command=lambda x=ph: self.tpl_text.insert("insert", x)).pack(fill="x", pady=2)

    def save_template(self):
        self.settings["template"] = self.tpl_text.get("1.0", "end").rstrip("\n")
        save_settings(self.settings)
        messagebox.showinfo("Saved", "Template saved.")

    def preview_msg(self):
        if not self.patients:
            messagebox.showwarning("No data", "Load the files first (tab 1).")
            return
        self.save_template()
        msg = self.build_message(self.patients[0])
        win = ctk.CTkToplevel(self); win.title("Preview"); win.geometry("520x600")
        tb = ctk.CTkTextbox(win, font=("Segoe UI", 15)); tb.pack(fill="both", expand=True, padx=8, pady=8)
        tb.insert("end", msg)

    # ---------------- Message building ----------------
    def build_message(self, p):
        params = self.settings["parameters"]
        lab_lines, recs = [], []
        for prm in params:
            if prm["name"] not in p["labs"]:
                continue
            val = p["labs"][prm["name"]]
            flag = classify(val, prm)
            disp = prm.get("ar") or prm["name"]
            unit = (" " + prm["unit"]) if prm.get("unit") else ""
            line = f"{disp}: {fmt_num(val)}{unit}"
            if flag:
                line += f"  ({FLAG_TEXT[flag]})"
            lab_lines.append(line)
            if flag in ("high", "low"):
                r = self.settings["recommendations"].get(f"{prm['name']}|{flag}")
                if r:
                    recs.append(r)
        msg = self.settings.get("template", DEFAULT_TEMPLATE)
        msg = msg.replace("{Name}", p["name"]).replace("{ID}", p["id"]).replace("{Date}", p.get("date", ""))
        msg = msg.replace("{LABS}", "\n".join(lab_lines))
        recs_block = ("التوصيات:\n" + "\n".join("- " + r for r in recs) + "\n") if recs else ""
        msg = msg.replace("{RECS}", recs_block)
        note_block = ("\nملاحظة الطبيب: " + p["note"] + "\n") if p.get("note") else ""
        msg = msg.replace("{NOTE}", note_block)
        for prm in params:
            ph = "{" + prm["name"] + "}"
            if ph in msg:
                if prm["name"] in p["labs"]:
                    v = p["labs"][prm["name"]]
                    fl = classify(v, prm)
                    rep = fmt_num(v) + ((" " + prm["unit"]) if prm.get("unit") else "")
                    if fl:
                        rep += f" ({FLAG_TEXT[fl]})"
                    msg = msg.replace(ph, rep)
                else:
                    msg = msg.replace(ph, "")
        return re.sub(r"\n{3,}", "\n\n", msg).strip()

    # ---------------- Tab 5: Send ----------------
    def build_send_tab(self):
        f = self.tabs.tab("5. Send")
        top = ctk.CTkFrame(f); top.pack(fill="x", padx=6, pady=4)
        ctk.CTkLabel(top, text="Gap between messages (seconds)  min:").pack(side="left", padx=2)
        self.gap_min = ctk.CTkEntry(top, width=70); self.gap_min.insert(0, str(self.settings["gap_min"])); self.gap_min.pack(side="left")
        ctk.CTkLabel(top, text="max:").pack(side="left", padx=2)
        self.gap_max = ctk.CTkEntry(top, width=70); self.gap_max.insert(0, str(self.settings["gap_max"])); self.gap_max.pack(side="left")
        ctk.CTkLabel(top, text="   Max messages this session:").pack(side="left", padx=2)
        self.max_sess = ctk.CTkEntry(top, width=70); self.max_sess.insert(0, str(self.settings["max_per_session"])); self.max_sess.pack(side="left")

        top2 = ctk.CTkFrame(f); top2.pack(fill="x", padx=6, pady=4)
        self.btn_start = ctk.CTkButton(top2, text="START sending", fg_color="#2a7", command=self.start_sending)
        self.btn_start.pack(side="left", padx=4)
        self.btn_stop = ctk.CTkButton(top2, text="STOP (after current)", fg_color="#b33", command=self.stop_sending, state="disabled")
        self.btn_stop.pack(side="left", padx=4)
        ctk.CTkButton(top2, text="Refresh list", command=self.refresh_send_list).pack(side="left", padx=4)
        ctk.CTkButton(top2, text="Reset sent log", command=self.reset_log).pack(side="left", padx=4)
        ctk.CTkLabel(top2, text="Test number:").pack(side="left", padx=8)
        self.test_num = ctk.CTkEntry(top2, width=140); self.test_num.pack(side="left")
        ctk.CTkButton(top2, text="Send test (first patient msg)", command=self.send_test).pack(side="left", padx=4)

        self.send_log = ctk.CTkTextbox(f, font=("Consolas", 12))
        self.send_log.pack(fill="both", expand=True, padx=6, pady=6)
        self.refresh_send_list()

    def already_sent_ids(self):
        keys = set()
        if os.path.exists(LOG_FILE):
            with open(LOG_FILE, "r", encoding="utf-8") as fh:
                for row in csv.reader(fh):
                    if len(row) >= 5 and row[4] == "sent":
                        keys.add(row[0] + "|" + row[2])
        return keys

    def log_send(self, p, status):
        new = not os.path.exists(LOG_FILE)
        with open(LOG_FILE, "a", encoding="utf-8", newline="") as fh:
            w = csv.writer(fh)
            if new:
                w.writerow(["ID", "Name", "Mobile", "Time", "Status"])
            w.writerow([p["id"], p["name"], p["mobile"], datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), status])

    def reset_log(self):
        if messagebox.askyesno("Reset", "Delete the sent log? All patients will be considered NOT sent."):
            if os.path.exists(LOG_FILE):
                os.remove(LOG_FILE)
            self.refresh_send_list()

    def refresh_send_list(self):
        if not hasattr(self, "send_log"):
            return
        sent = self.already_sent_ids()
        self.send_log.delete("1.0", "end")
        pend = [p for p in self.patients if p["id"] + "|" + p["mobile"] not in sent]
        self.send_log.insert("end", f"Loaded patients with mobile: {len(self.patients)} | Already sent: {len(self.patients)-len(pend)} | Pending: {len(pend)}\n")
        self.send_log.insert("end", "-" * 80 + "\n")
        for p in self.patients:
            st = "SENT" if p["id"] + "|" + p["mobile"] in sent else "pending"
            self.send_log.insert("end", f"[{st}] {p['name']} - {p['mobile']}\n")

    def log_line(self, txt):
        self.msg_queue.put(txt)

    def poll_queue(self):
        try:
            while True:
                txt = self.msg_queue.get_nowait()
                self.send_log.insert("end", txt + "\n")
                self.send_log.see("end")
        except queue.Empty:
            pass
        self.after(300, self.poll_queue)

    def save_send_settings(self):
        try:
            self.settings["gap_min"] = int(self.gap_min.get())
            self.settings["gap_max"] = int(self.gap_max.get())
            self.settings["max_per_session"] = int(self.max_sess.get())
        except ValueError:
            messagebox.showerror("Error", "Gaps and max must be numbers.")
            return False
        save_settings(self.settings)
        return True

    def start_sending(self):
        if self.sending:
            return
        if not self.patients:
            messagebox.showwarning("No data", "Load files first (tab 1).")
            return
        if not self.save_send_settings():
            return
        sent = self.already_sent_ids()
        todo = [p for p in self.patients if p["id"] + "|" + p["mobile"] not in sent][: self.settings["max_per_session"]]
        if not todo:
            messagebox.showinfo("Done", "No pending patients.")
            return
        if not messagebox.askyesno("Confirm", f"Send to {len(todo)} patients now?\nGap: {self.settings['gap_min']}-{self.settings['gap_max']} sec."):
            return
        self.sending = True; self.stop_flag = False
        self.btn_start.configure(state="disabled"); self.btn_stop.configure(state="normal")
        threading.Thread(target=self.send_worker, args=(todo,), daemon=True).start()

    def stop_sending(self):
        self.stop_flag = True
        self.log_line(">>> STOP requested - will stop after current message.")

    def send_test(self):
        if not self.patients:
            messagebox.showwarning("No data", "Load files first.")
            return
        num = normalize_mobile(self.test_num.get(), self.settings["country_code"])
        if not num:
            messagebox.showwarning("No number", "Enter a test number.")
            return
        fake = dict(self.patients[0]); fake["mobile"] = num
        self.sending = True
        self.btn_start.configure(state="disabled")
        threading.Thread(target=self.send_worker, args=([fake],), daemon=True).start()

    # ---------------- WhatsApp ----------------
    def get_driver(self):
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options as ChromeOptions
        from selenium.webdriver.edge.options import Options as EdgeOptions
        try:
            opt = ChromeOptions()
            opt.add_argument(f"--user-data-dir={WA_PROFILE}")
            opt.add_argument("--no-first-run")
            opt.add_experimental_option("excludeSwitches", ["enable-logging"])
            return webdriver.Chrome(options=opt)
        except Exception as e1:
            self.log_line(f"Chrome failed ({e1.__class__.__name__}), trying Edge...")
            opt = EdgeOptions()
            opt.add_argument(f"--user-data-dir={WA_PROFILE}_edge")
            opt.add_argument("--no-first-run")
            return webdriver.Edge(options=opt)

    def wait_logged_in(self, driver, timeout=180):
        from selenium.webdriver.common.by import By
        self.log_line("Opening WhatsApp Web... if QR code appears, scan it with your phone.")
        driver.get("https://web.whatsapp.com")
        t0 = time.time()
        while time.time() - t0 < timeout:
            if self.stop_flag:
                return False
            try:
                if driver.find_elements(By.CSS_SELECTOR, 'div[contenteditable="true"]') or \
                   driver.find_elements(By.CSS_SELECTOR, 'div[aria-label]')[0:0] or \
                   driver.find_elements(By.CSS_SELECTOR, '#side'):
                    if driver.find_elements(By.CSS_SELECTOR, '#side'):
                        self.log_line("WhatsApp Web is ready.")
                        return True
            except Exception:
                pass
            time.sleep(2)
        self.log_line("Timed out waiting for WhatsApp Web login.")
        return False

    def send_one(self, driver, p, msg):
        from selenium.webdriver.common.by import By
        from selenium.webdriver.common.keys import Keys
        url = f"https://web.whatsapp.com/send?phone={p['mobile']}&text={urllib.parse.quote(msg)}"
        driver.get(url)
        t0 = time.time()
        box = None
        while time.time() - t0 < 60:
            if self.stop_flag and box is None:
                pass
            try:
                boxes = driver.find_elements(By.CSS_SELECTOR, 'footer div[contenteditable="true"]')
                if boxes:
                    box = boxes[-1]
                    break
            except Exception:
                pass
            # invalid number popup
            try:
                for el in driver.find_elements(By.CSS_SELECTOR, 'div[data-animate-modal-popup="true"]'):
                    if "invalid" in el.text.lower() or "غير صحيح" in el.text:
                        return "invalid_number"
            except Exception:
                pass
            time.sleep(1.5)
        if box is None:
            return "failed_open_chat"
        time.sleep(random.uniform(2, 4))
        try:
            btns = driver.find_elements(By.CSS_SELECTOR, 'button[aria-label="Send"], span[data-icon="send"], span[data-icon="wds-ic-send-filled"]')
            if btns:
                btns[-1].click()
            else:
                box.send_keys(Keys.ENTER)
        except Exception:
            try:
                box.send_keys(Keys.ENTER)
            except Exception:
                return "failed_send"
        time.sleep(random.uniform(3, 5))
        return "sent"

    def send_worker(self, todo):
        try:
            driver = self.get_driver()
        except Exception as e:
            self.log_line(f"ERROR: could not start browser: {e}")
            self.sending_done()
            return
        try:
            if not self.wait_logged_in(driver):
                self.sending_done(); driver.quit(); return
            total = len(todo)
            for i, p in enumerate(todo, 1):
                if self.stop_flag:
                    self.log_line(">>> Stopped by user.")
                    break
                msg = self.build_message(p)
                self.log_line(f"[{i}/{total}] Sending to {p['name']} ({p['mobile']}) ...")
                try:
                    status = self.send_one(driver, p, msg)
                except Exception as e:
                    status = f"error:{e.__class__.__name__}"
                self.log_send(p, status)
                self.log_line(f"      -> {status}")
                if i < total and not self.stop_flag:
                    gap = random.randint(self.settings["gap_min"], max(self.settings["gap_min"], self.settings["gap_max"]))
                    self.log_line(f"      waiting {gap} sec ...")
                    t0 = time.time()
                    while time.time() - t0 < gap:
                        if self.stop_flag:
                            break
                        time.sleep(1)
            self.log_line(">>> Session finished. Browser stays open; you may close it.")
        finally:
            self.sending_done()

    def sending_done(self):
        self.sending = False
        try:
            self.btn_start.configure(state="normal")
            self.btn_stop.configure(state="disabled")
        except Exception:
            pass


if __name__ == "__main__":
    app = App()
    app.mainloop()
