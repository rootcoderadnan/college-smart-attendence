# hrpg_attendance_full.py
"""
H. R. P. G. College - Smart Attendance System (Tkinter)
- QR generator (qrcode)
- Camera QR scanner (OpenCV)
- Daily & Monthly Excel reports (pandas + openpyxl)
- CSV storage (data/)
- Responsive layout, scaled logo, single system titlebar
"""

import os
import csv
import threading
import traceback
from datetime import datetime
from glob import glob
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog

# PIL is required for logo display and scaling
try:
    from PIL import Image, ImageTk
except Exception:
    Image = None
    ImageTk = None

# optional features (scanner/reports)
try:
    import cv2
except Exception:
    cv2 = None
try:
    import pandas as pd
except Exception:
    pd = None
try:
    import qrcode
except Exception:
    qrcode = None

# ---------------- Config ----------------
COLLEGE_NAME = "H. R. P. G. College, Khalilabad"
LOGO_FILENAMES = ["college_logo.png", "college_logo.jpg", "logo.png", "/mnt/data/college_logo.png.png"]
DATA_DIR = "data"
REPORTS_DIR = "reports"
QRCODES_DIR = "qrcodes"
STUDENTS_FILE = "students.csv"
COOLDOWN_SECONDS = 6
DEFAULT_CAMERA = "0"
BASE_WIDTH = 900
BASE_HEIGHT = 620
# ----------------------------------------

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(REPORTS_DIR, exist_ok=True)
os.makedirs(QRCODES_DIR, exist_ok=True)

# ---------------- Helpers ----------------
def find_logo_path():
    for name in LOGO_FILENAMES:
        if os.path.exists(name):
            return name
    return None

def load_logo_scaled(path, max_side):
    """Return (ImageTk.PhotoImage, (w,h)) or (None,(0,0))"""
    if Image is None or ImageTk is None:
        return None, (0, 0)
    try:
        img = Image.open(path)
        img.thumbnail((max_side, max_side), Image.LANCZOS)
        return ImageTk.PhotoImage(img), img.size
    except Exception:
        return None, (0, 0)

def ensure_demo_students():
    if not os.path.exists(STUDENTS_FILE):
        demo = [
            ("101", "Aman Kumar"),
            ("102", "Neha Sharma"),
            ("103", "Ravi Singh"),
            ("104", "Sonal Verma"),
            ("105", "Vivek Yadav"),
        ]
        with open(STUDENTS_FILE, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Roll", "Name"])
            for r, n in demo:
                writer.writerow([r, n])

def load_students_dict():
    if not os.path.exists(STUDENTS_FILE):
        raise FileNotFoundError(f"{STUDENTS_FILE} not found. Create it with header: Roll,Name")
    students = {}
    with open(STUDENTS_FILE, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames:
            raise RuntimeError("students.csv has no header")
        # tolerant header lookup
        cols = reader.fieldnames
        roll_col = None
        name_col = None
        for c in cols:
            s = c.strip().lower()
            if s in ("roll", "id", "student id", "student_id"):
                roll_col = c
            if s in ("name", "student", "student_name", "fullname"):
                name_col = c
        if roll_col is None or name_col is None:
            # fallback to first two columns
            if len(cols) >= 2:
                roll_col, name_col = cols[0], cols[1]
            else:
                raise RuntimeError("students.csv must contain 'Roll' and 'Name' headers")
        f.seek(0)
        reader = csv.DictReader(f)
        for r in reader:
            roll = str(r.get(roll_col, "")).strip()
            name = str(r.get(name_col, "")).strip()
            if roll:
                students[roll] = name if name else "Unknown"
    return students

def parse_payload(payload: str):
    s = str(payload).strip()
    if not s:
        return None
    if "," in s:
        return s.split(",", 1)[0].strip()
    if ":" in s:
        parts = s.split(":", 1)
        right = parts[1].strip()
        if right:
            return right
    return s

def attendance_filepath(subject):
    safe = "".join(ch if (ch.isalnum() or ch in (" ", "-", "_")) else "_" for ch in subject).strip().replace(" ", "_")
    return os.path.join(DATA_DIR, f"attendance_{safe}.csv")

def ensure_attendance_file(path):
    if not os.path.exists(path):
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Roll", "Name", "Day", "Date", "Time", "Subject"])

def append_attendance_row(path, roll, name, day, date_s, time_s, subject):
    with open(path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([roll, name, day, date_s, time_s, subject])

def read_attendance_rows(path):
    rows = []
    if not os.path.exists(path):
        return rows
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for r in reader:
            rows.append({k: (v if v is not None else "") for k, v in r.items()})
    return rows

def generate_qr_codes(payload_as="roll,name"):
    if qrcode is None:
        return False, "qrcode library not installed. Run: pip install qrcode"
    students = load_students_dict()
    for roll, name in students.items():
        payload = f"{roll},{name}" if payload_as == "roll,name" else roll
        img = qrcode.make(payload)
        out_path = os.path.join(QRCODES_DIR, f"{roll}.png")
        img.save(out_path)
    return True, f"QR codes saved to {QRCODES_DIR}/"

# ---------------- Scanner Worker ----------------
class ScannerWorker(threading.Thread):
    def __init__(self, subject, cam_arg, on_finish, update_callback=None):
        super().__init__(daemon=True)
        self.subject = subject
        self.cam_arg = cam_arg
        self.on_finish = on_finish
        self.update_callback = update_callback
        self._stop = threading.Event()

    def stop(self):
        self._stop.set()

    def stopped(self):
        return self._stop.is_set()

    def run(self):
        if cv2 is None:
            self.on_finish(False, "OpenCV is not installed (pip install opencv-python)")
            return

        try:
            students = load_students_dict()
        except Exception as e:
            self.on_finish(False, f"Failed to load students.csv: {e}")
            return

        file_path = attendance_filepath(self.subject)
        ensure_attendance_file(file_path)

        # open camera
        def open_capture(cam_arg):
            if not cam_arg:
                cam_arg = DEFAULT_CAMERA
            cam_arg = str(cam_arg).strip()
            try:
                idx = int(cam_arg)
                cap = cv2.VideoCapture(idx)
                if cap.isOpened():
                    return cap
            except Exception:
                pass
            cap = cv2.VideoCapture(cam_arg)
            if cap.isOpened():
                return cap
            return None

        cap = open_capture(self.cam_arg)
        if cap is None:
            self.on_finish(False, f"Cannot open camera: {self.cam_arg}")
            return

        detector = cv2.QRCodeDetector()
        last_seen = {}

        try:
            while not self.stopped():
                ret, frame = cap.read()
                if not ret or frame is None:
                    cv2.waitKey(50)
                    continue
                data, bbox, _ = detector.detectAndDecode(frame)
                roll_payload = ""
                if data:
                    roll_payload = parse_payload(data)
                    if roll_payload:
                        now = datetime.now()
                        last = last_seen.get(roll_payload)
                        if last and (now - last).total_seconds() < COOLDOWN_SECONDS:
                            pass
                        else:
                            last_seen[roll_payload] = now
                            name = students.get(roll_payload, "")
                            if not name:
                                if self.update_callback:
                                    self.update_callback(f"Unknown QR: {roll_payload}")
                            else:
                                day_name = now.strftime("%A")
                                if day_name == "Sunday":
                                    if self.update_callback:
                                        self.update_callback("Sunday - attendance disabled.")
                                else:
                                    date_s = now.strftime("%Y-%m-%d")
                                    time_s = now.strftime("%H:%M:%S")
                                    rows = read_attendance_rows(file_path)
                                    dup = any((r.get("Roll", "").strip() == roll_payload and r.get("Date", "").strip() == date_s and r.get("Subject", "").strip() == self.subject) for r in rows)
                                    if dup:
                                        if self.update_callback:
                                            self.update_callback(f"Already marked: {roll_payload}")
                                    else:
                                        append_attendance_row(file_path, roll_payload, name, day_name, date_s, time_s, self.subject)
                                        if self.update_callback:
                                            self.update_callback(f"Marked: {roll_payload} - {name}")

                # draw bbox if exists
                if bbox is not None and getattr(bbox, "shape", None):
                    try:
                        pts = None
                        if len(bbox.shape) == 3:
                            pts = bbox[0].astype(int)
                        else:
                            pts = bbox.astype(int)
                        for i in range(len(pts)):
                            pt1 = tuple(pts[i]); pt2 = tuple(pts[(i+1) % len(pts)])
                            cv2.line(frame, pt1, pt2, (10,160,10), 3)
                        if roll_payload:
                            x, y = int(pts[0][0]), int(pts[0][1])
                            cv2.putText(frame, roll_payload, (x, max(y-10,10)), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (10,160,10), 2)
                    except Exception:
                        pass

                cv2.imshow("Scanner - Press q to stop", frame)
                if cv2.waitKey(1) & 0xFF == ord("q"):
                    break

            try:
                cap.release()
            except Exception:
                pass
            try:
                cv2.destroyAllWindows()
            except Exception:
                pass
            self.on_finish(True, f"Scanner stopped. Attendance saved to: {file_path}")
        except Exception as e:
            try:
                cap.release()
                cv2.destroyAllWindows()
            except Exception:
                pass
            tb = traceback.format_exc()
            self.on_finish(False, f"Scanner error: {e}\n{tb}")

# ---------------- Reporting ----------------
def generate_daily_report(subject):
    if pd is None:
        return False, "pandas not installed (pip install pandas openpyxl)"
    safe = "".join(ch if (ch.isalnum() or ch in (" ", "-", "_")) else "_" for ch in subject).strip().replace(" ", "_")
    primary = os.path.join(DATA_DIR, f"attendance_{safe}.csv")
    files = [primary] if os.path.exists(primary) else sorted(glob(os.path.join(DATA_DIR, f"attendance_{safe}_*.csv")))
    if not files:
        return False, f"No attendance files found for subject {subject}."
    dfs = []
    for f in files:
        try:
            d = pd.read_csv(f, dtype=str)
            dfs.append(d)
        except Exception:
            pass
    if not dfs:
        return False, "No readable CSV files."
    big = pd.concat(dfs, ignore_index=True)
    os.makedirs(REPORTS_DIR, exist_ok=True)
    out = os.path.join(REPORTS_DIR, f"{safe}_daily_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    try:
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            big.to_excel(writer, sheet_name="Raw_Attendance", index=False)
        return True, out
    except Exception as e:
        return False, f"Failed to write report: {e}"

def generate_monthly_report(subject, month=None, year=None):
    if pd is None:
        return False, "pandas not installed (pip install pandas openpyxl)"
    now = datetime.now()
    if month is None:
        month = now.month
    if year is None:
        year = now.year
    safe = "".join(ch if (ch.isalnum() or ch in (" ", "-", "_")) else "_" for ch in subject).strip().replace(" ", "_")
    files = sorted(glob(os.path.join(DATA_DIR, f"attendance_{safe}*.csv")))
    if not files:
        return False, f"No attendance files found for subject {subject}."
    dfs = []
    for f in files:
        try:
            d = pd.read_csv(f, dtype=str)
            d.columns = [c.strip() for c in d.columns]
            if "Date" not in d.columns:
                for c in d.columns:
                    sample = d[c].dropna().astype(str)
                    if not sample.empty and any(ch.isdigit() for ch in sample.iloc[0]):
                        d.rename(columns={c: "Date"}, inplace=True)
                        break
            if "Date" not in d.columns:
                continue
            d["Date_norm"] = pd.to_datetime(d["Date"], errors="coerce")
            d = d.dropna(subset=["Date_norm"])
            d["Month"] = d["Date_norm"].dt.month
            d["Year"] = d["Date_norm"].dt.year
            dfs.append(d)
        except Exception:
            pass
    if not dfs:
        return False, "No valid attendance rows found."
    big = pd.concat(dfs, ignore_index=True)
    big_month = big[(big["Month"] == int(month)) & (big["Year"] == int(year))].copy()
    if big_month.empty:
        return False, f"No records for {month}/{year}."
    for c in ("Roll", "Name"):
        if c not in big_month.columns:
            return False, f"Attendance data must contain '{c}' column."
    big_month["Present"] = 1
    matrix = big_month.pivot_table(index=["Roll", "Name"], columns=big_month["Date_norm"].dt.strftime("%Y-%m-%d"), values="Present", aggfunc="max", fill_value=0)
    summary = matrix.sum(axis=1).reset_index(name="Present_Days")
    total_days = len(matrix.columns)
    summary["Total_Days"] = total_days
    summary["Attendance_%"] = (summary["Present_Days"] / total_days * 100).round(2)
    earliest_pivot = None
    if "Time" in big_month.columns:
        big_month["Time_norm"] = pd.to_datetime(big_month["Time"], errors="coerce").dt.strftime("%H:%M:%S")
        earliest = big_month.dropna(subset=["Time_norm"]).groupby(["Roll", "Name", big_month["Date_norm"].dt.strftime("%Y-%m-%d")], as_index=False)["Time_norm"].min()
        earliest_pivot = earliest.pivot_table(index=["Roll", "Name"], columns=big_month["Date_norm"].dt.strftime("%Y-%m-%d"), values="Time_norm", aggfunc="first")
        earliest_pivot = earliest_pivot.reindex(sorted(earliest_pivot.columns), axis=1)
    os.makedirs(REPORTS_DIR, exist_ok=True)
    out = os.path.join(REPORTS_DIR, f"{safe}_Monthly_{year}_{month}.xlsx")
    try:
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            summary.to_excel(writer, sheet_name="Summary", index=False)
            matrix.reset_index().to_excel(writer, sheet_name="Matrix", index=False)
            if earliest_pivot is not None:
                earliest_pivot.reset_index().to_excel(writer, sheet_name="DayTime", index=False)
            big_month.to_excel(writer, sheet_name="Raw", index=False)
        return True, out
    except Exception as e:
        return False, f"Failed to write monthly report: {e}"

# ---------------- GUI ----------------
class HRPGApp:
    def __init__(self, root):
        self.root = root
        root.title(f"{COLLEGE_NAME} - Smart Attendance")
        root.geometry(f"{BASE_WIDTH}x{BASE_HEIGHT}")
        root.minsize(780, 520)

        # main container
        container = tk.Frame(root, bg="#ffffff")
        container.pack(fill="both", expand=True)

        # header frame (logo + title)
        header = tk.Frame(container, bg="#ffffff")
        header.pack(fill="x", pady=(10, 6), padx=12)

        # logo
        self.logo_path = find_logo_path()
        self.logo_img = None
        self.logo_label = tk.Label(header, bg="#ffffff")
        self.logo_label.pack(side="left", padx=(6,12))
        if self.logo_path:
            imgtk, (w_img, h_img) = load_logo_scaled(self.logo_path, max_side=140)
            if imgtk:
                self.logo_img = imgtk
                self.logo_label.config(image=self.logo_img)
            else:
                self.logo_label.config(text="[Logo]", fg="red")

        # title and subtitle
        title_box = tk.Frame(header, bg="#ffffff")
        title_box.pack(side="left", anchor="w")
        self.title_lbl = tk.Label(title_box, text=COLLEGE_NAME, font=("Segoe UI", 20, "bold"), bg="#ffffff", fg="#0a3d62")
        self.title_lbl.pack(anchor="w")
        self.subtitle = tk.Label(title_box, text="Smart Attendance System", font=("Segoe UI", 12), bg="#ffffff", fg="#e67e22")
        self.subtitle.pack(anchor="w")

        # controls frame (subject, camera)
        ctrl = tk.Frame(container, bg="#f4f6f9")
        ctrl.pack(fill="x", padx=12, pady=(6,10))

        tk.Label(ctrl, text="Select Subject:", font=("Segoe UI", 10), bg="#f4f6f9").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.subject_var = tk.StringVar()
        self.subject_combo = ttk.Combobox(ctrl, textvariable=self.subject_var, values=["C Language", "Math", "Business Accounting"], width=30)
        self.subject_combo.grid(row=0, column=1, sticky="w", padx=6, pady=6)
        self.subject_combo.set("C Language")

        tk.Label(ctrl, text="Or enter subject:", font=("Segoe UI", 10), bg="#f4f6f9").grid(row=0, column=2, sticky="w", padx=6, pady=6)
        self.subject_custom = tk.Entry(ctrl, width=28)
        self.subject_custom.grid(row=0, column=3, sticky="w", padx=6, pady=6)

        tk.Label(ctrl, text="Camera (index or URL):", font=("Segoe UI", 10), bg="#f4f6f9").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        self.cam_var = tk.StringVar(value=DEFAULT_CAMERA)
        tk.Entry(ctrl, textvariable=self.cam_var, width=30).grid(row=1, column=1, sticky="w", padx=6, pady=6)

        # make grid responsive
        ctrl.grid_columnconfigure(1, weight=1)
        ctrl.grid_columnconfigure(3, weight=1)

        # buttons area - using a center frame that expands, buttons will stretch horizontally
        btn_area = tk.Frame(container, bg="#f4f6f9")
        btn_area.pack(fill="both", expand=True, padx=12, pady=6)

        # single-column grid for buttons
        btn_area.grid_columnconfigure(0, weight=1)

        # big colorful buttons
        btn_font = ("Helvetica", 12, "bold")
        pady = 10

        self.btn_scan = tk.Button(btn_area, text="Start Scanner", font=btn_font, bg="#2980b9", fg="white", height=2, command=self.start_scanner)
        self.btn_scan.grid(row=0, column=0, sticky="ew", padx=(120,120), pady=(pady,pady))

        self.btn_stop = tk.Button(btn_area, text="Stop Scanner", font=btn_font, bg="#c0392b", fg="white", height=2, command=self.stop_scanner, state="disabled")
        self.btn_stop.grid(row=1, column=0, sticky="ew", padx=(120,120), pady=(0,pady))

        self.btn_generate_qr = tk.Button(btn_area, text="Generate QR Codes", font=btn_font, bg="#27ae60", fg="white", height=2, command=self.generate_qr)
        self.btn_generate_qr.grid(row=2, column=0, sticky="ew", padx=(120,120), pady=(0,pady))

        self.btn_daily = tk.Button(btn_area, text="Generate Daily Report", font=btn_font, bg="#f39c12", fg="white", height=2, command=self.generate_daily_report)
        self.btn_daily.grid(row=3, column=0, sticky="ew", padx=(120,120), pady=(0,pady))

        self.btn_monthly = tk.Button(btn_area, text="Generate Monthly Report", font=btn_font, bg="#8e44ad", fg="white", height=2, command=self.generate_monthly_report)
        self.btn_monthly.grid(row=4, column=0, sticky="ew", padx=(120,120), pady=(0,pady))

        self.btn_open_data = tk.Button(btn_area, text="Open Data Folder", font=btn_font, bg="#16a085", fg="white", height=2, command=self.open_data_folder)
        self.btn_open_data.grid(row=5, column=0, sticky="ew", padx=(120,120), pady=(0,pady))

        self.btn_admin = tk.Button(btn_area, text="Open Admin - Today's Attendance", font=btn_font, bg="#34495e", fg="white", height=2, command=self.open_admin)
        self.btn_admin.grid(row=6, column=0, sticky="ew", padx=(120,120), pady=(0,pady))

        # status bar
        self.status_var = tk.StringVar(value="Ready")
        status = tk.Label(container, textvariable=self.status_var, anchor="w", bg="#f4f6f9", font=("Segoe UI", 10))
        status.pack(fill="x", padx=12, pady=(6,12))

        # hover effects
        for b in [self.btn_scan, self.btn_stop, self.btn_generate_qr, self.btn_daily, self.btn_monthly, self.btn_open_data, self.btn_admin]:
            b.bind("<Enter>", lambda e, btn=b: btn.config(relief="raised"))
            b.bind("<Leave>", lambda e, btn=b: btn.config(relief="flat"))

        # small debounce for resize so we don't do heavy work many times and cause lag
        self._resize_after_id = None
        root.bind("<Configure>", self._on_root_configure)

        self.scanner_worker = None
        ensure_demo_students()

        # initial scale pass so logo fits
        self._apply_scaling()

    # ---------- UI scaling (debounced) ----------
    def _on_root_configure(self, event):
        if self._resize_after_id:
            self.root.after_cancel(self._resize_after_id)
        # schedule apply scaling after 120ms
        self._resize_after_id = self.root.after(120, self._apply_scaling)

    def _apply_scaling(self):
        # adjust padding (horizontal) to keep buttons centered without pixel placing
        try:
            w = max(self.root.winfo_width(), BASE_WIDTH)
            # compute side padding as percent of width (so on big screens buttons stay centered)
            side_pad = int(w * 0.12)  # 12% left/right
            # apply padx for each button (we used same padx when creating)
            widgets = [self.btn_scan, self.btn_stop, self.btn_generate_qr, self.btn_daily, self.btn_monthly, self.btn_open_data, self.btn_admin]
            for wbtn in widgets:
                # grid_info exists; update padding
                info = wbtn.grid_info()
                if info:
                    wbtn.grid_configure(padx=(side_pad, side_pad))

            # scale fonts modestly based on width
            scale = max(0.8, min(1.6, w / BASE_WIDTH))
            base_font_size = int(12 * scale)
            heading_size = int(18 * scale)
            subtitle_size = int(11 * scale)
            btn_font = ("Helvetica", base_font_size, "bold")
            self.title_lbl.config(font=("Segoe UI", max(14, heading_size), "bold"))
            self.subtitle.config(font=("Segoe UI", max(10, subtitle_size)))
            for wbtn in widgets:
                wbtn.config(font=btn_font)

            # scale logo max side
            if self.logo_path:
                max_side = max(60, int(140 * scale))
                imgtk, (wi, hi) = load_logo_scaled(self.logo_path, max_side=max_side)
                if imgtk:
                    self.logo_img = imgtk
                    self.logo_label.config(image=self.logo_img, text="")
                else:
                    self.logo_label.config(text="[Logo]", image="")
            else:
                self.logo_label.config(text="[Logo Missing]", fg="red")
        except Exception:
            pass

    # ---------- helper UI methods ----------
    def chosen_subject(self):
        custom = self.subject_custom.get().strip()
        return custom if custom else self.subject_var.get().strip()

    def set_status(self, text):
        self.status_var.set(text)
        self.root.update_idletasks()

    # ---------- actions ----------
    def generate_qr(self):
        def worker():
            self.set_status("Generating QR codes...")
            try:
                ok, msg = generate_qr_codes(payload_as="roll,name")
                if ok:
                    messagebox.showinfo("QR Generator", msg)
                else:
                    messagebox.showerror("QR Generator", msg)
            except Exception as e:
                messagebox.showerror("QR Generator", f"Error: {e}")
            self.set_status("Ready")
        threading.Thread(target=worker, daemon=True).start()

    def start_scanner(self):
        subj = self.chosen_subject()
        cam = self.cam_var.get().strip()
        if not subj:
            messagebox.showwarning("Subject required", "Please enter a subject name.")
            return
        if self.scanner_worker and self.scanner_worker.is_alive():
            messagebox.showinfo("Scanner", "Scanner already running.")
            return

        if cv2 is None:
            messagebox.showerror("Scanner", "OpenCV (cv2) not installed. Install with: pip install opencv-python")
            return

        def on_finish(success, message):
            def ui_done():
                self.btn_scan.config(state="normal")
                self.btn_stop.config(state="disabled")
                self.set_status("Ready")
                if success:
                    messagebox.showinfo("Scanner", message)
                else:
                    messagebox.showerror("Scanner Error", message)
            self.root.after(0, ui_done)

        def update_cb(msg):
            self.set_status(msg)

        self.btn_scan.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.set_status("Starting scanner...")
        self.scanner_worker = ScannerWorker(subj, cam, on_finish, update_callback=update_cb)
        self.scanner_worker.start()
        self.set_status("Scanner running (press Stop Scanner or close scanner window)")

    def stop_scanner(self):
        if self.scanner_worker and self.scanner_worker.is_alive():
            self.set_status("Stopping scanner...")
            self.scanner_worker.stop()
        else:
            messagebox.showinfo("Scanner", "Scanner not running.")

    def generate_daily_report(self):
        subj = self.chosen_subject()
        if not subj:
            messagebox.showwarning("Subject required", "Please enter subject name.")
            return
        def worker():
            self.set_status("Generating daily report...")
            ok, info = generate_daily_report(subj)
            if ok:
                messagebox.showinfo("Daily Report", f"Saved: {info}")
            else:
                messagebox.showerror("Daily Report Error", info)
            self.set_status("Ready")
        threading.Thread(target=worker, daemon=True).start()

    def generate_monthly_report(self):
        subj = self.chosen_subject()
        if not subj:
            messagebox.showwarning("Subject required", "Please enter subject name.")
            return
        month = simpledialog.askinteger("Month", "Enter month number (1-12):", parent=self.root, minvalue=1, maxvalue=12)
        if month is None:
            return
        year = simpledialog.askinteger("Year", "Enter year (e.g. 2025):", parent=self.root, minvalue=2000, maxvalue=2100)
        if year is None:
            return
        def worker():
            self.set_status("Generating monthly report...")
            ok, info = generate_monthly_report(subj, month=month, year=year)
            if ok:
                messagebox.showinfo("Monthly Report", f"Saved: {info}")
            else:
                messagebox.showerror("Monthly Report Error", info)
            self.set_status("Ready")
        threading.Thread(target=worker, daemon=True).start()

    def open_data_folder(self):
        folder = os.path.join(os.getcwd(), DATA_DIR)
        os.makedirs(folder, exist_ok=True)
        try:
            if os.name == "nt":
                os.startfile(folder)
            else:
                import subprocess
                subprocess.run(["xdg-open", folder])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open data folder: {e}")

    def open_admin(self):
        subj = self.chosen_subject()
        if not subj:
            messagebox.showwarning("Subject required", "Please enter subject name.")
            return
        file_path = attendance_filepath(subj)
        rows = read_attendance_rows(file_path)
        today = datetime.now().strftime("%Y-%m-%d")
        todays = [r for r in rows if r.get("Date","").strip() == today]
        top = tk.Toplevel(self.root)
        top.title(f"Today's Attendance - {subj}")
        top.geometry("700x420")
        lb = tk.Listbox(top, font=("Segoe UI", 11))
        lb.pack(fill="both", expand=True, padx=12, pady=12)
        if not todays:
            lb.insert("end", "No attendance marked today.")
        else:
            lb.insert("end", f"Attendance for {today} - Subject: {subj}")
            lb.insert("end", "-"*80)
            for r in todays:
                txt = f"{r.get('Roll','')}  |  {r.get('Name','')}  |  {r.get('Time','')}"
                lb.insert("end", txt)

# ----------------- run -----------------
if __name__ == "__main__":
    root = tk.Tk()
    app = HRPGApp(root)
    root.mainloop()
