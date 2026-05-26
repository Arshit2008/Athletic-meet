import os
import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import openpyxl

# =========================================================================
# 1. EXCEL FILE PIPELINE DATA INTERFACE LAYER
# =========================================================================
class AthleticExcelPipeline:
    def __init__(self):
        self.conn = sqlite3.connect(":memory:")
        self.cursor = self.conn.cursor()
        self._create_database_schema()
        
    def _create_database_schema(self):
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS school_data (
                stu_id TEXT PRIMARY KEY, name TEXT, class TEXT, house TEXT,
                dob TEXT, height REAL, weight REAL, run_100 INTEGER, run_200 INTEGER
            )
        ''')
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS trials_active (
                stu_id TEXT, event TEXT, heat INTEGER, lane INTEGER, name TEXT, 
                house TEXT, class TEXT, division TEXT, track_time REAL DEFAULT 0.00
            )
        ''')
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS semifinals_output (
                event TEXT, heat INTEGER, lane INTEGER, name TEXT, house TEXT, seed_time REAL
            )
        ''')
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS house_points (
                house TEXT, event TEXT, points INTEGER,
                PRIMARY KEY (house, event)
            )
        ''')
        for h in ["Red", "Blue", "Green", "Yellow"]:
            for ev in ["100m", "200m"]:
                self.cursor.execute("INSERT OR IGNORE INTO house_points (house, event, points) VALUES (?, ?, 0)", (h, ev))
        self.conn.commit()

    def load_excel_workbook(self, filename="School_Reg.xlsx"):
        if not os.path.exists(filename):
            messagebox.showerror("File Error", f"Could not find '{filename}'!\nPlease check directory structures.")
            return False
            
        try:
            wb = openpyxl.load_workbook(filename, data_only=True)
            sheet = wb.active
            data_tuples = []
            
            for row in sheet.iter_rows(min_row=2, max_row=2000, values_only=True):
                if not row or row is None or str(row).strip() == "":
                    break 
                
                stu_id, name, class_and_section, house, dob_val, height, weight, r100, r200 = row
                
                if isinstance(dob_val, datetime):
                    dob_string = dob_val.strftime("%d-%m-%Y")
                else:
                    dob_string = str(dob_val).strip()
                
                data_tuples.append((
                    str(stu_id).strip(), str(name).strip(), str(class_and_section).strip(), str(house).strip(),
                    dob_string, float(height or 0.0), float(weight or 0.0),
                    1 if str(r100).strip().upper() == 'TRUE' else 0,
                    1 if str(r200).strip().upper() == 'TRUE' else 0
                ))
            
            self.cursor.execute("DELETE FROM school_data")
            self.cursor.executemany("INSERT INTO school_data VALUES (?,?,?,?,?,?,?,?,?)", data_tuples)
            self.conn.commit()
            return True
        except Exception as e:
            messagebox.showerror("Excel Parse Error", f"Failed reading from XLSX file:\n{str(e)}")
            return False

# =========================================================================
# 2. SEEDING LOGIC ENGINE
# =========================================================================
class MeetLogicProcessor:
    def __init__(self, pipeline):
        self.pipeline = pipeline

    def calculate_legacy_score(self, dob_str, height_cm, weight_kg):
        clean_dob = str(dob_str).strip()
        date_formats = ["%d-%m-%Y", "%Y-%m-%d"]
        dob = None
        for fmt in date_formats:
            try:
                dob = datetime.strptime(clean_dob, fmt)
                break
            except ValueError:
                continue
        if dob is None:
            raise ValueError(f"Date '{dob_str}' format invalid.")
            
        meet_date = datetime(2026, 5, 25)
        age_months = (meet_date - dob).days / 30.4
        score = (age_months / 9.0) + (height_cm / 7.62) + (weight_kg * 0.73) - 3.50
        return round(score, 2)

    def assign_division(self, score):
        if score >= 100: return "O"
        elif score >= 88: return "A"
        elif score >= 76: return "B"
        else: return "C"

    def execute_initial_seeding(self, target_event):
        flag_col = "run_100" if target_event == "100m" else "run_200"
        self.pipeline.cursor.execute(f"SELECT stu_id, name, class, house, dob, height, weight FROM school_data WHERE {flag_col} = 1")
        athletes = self.pipeline.cursor.fetchall()

        pools = {"O": [], "A": [], "B": [], "C": []}
        for stu_id, name, class_name, house, dob_string, height, weight in athletes:
            score = self.calculate_legacy_score(dob_string, height, weight)
            div = self.assign_division(score)
            pools[div].append({"id": stu_id, "name": name, "class": class_name, "house": house, "score": score})

        self.pipeline.cursor.execute("DELETE FROM trials_active WHERE event=?", (target_event,))
        heat_idx = 1
        rotation_state = 0

        for div, pool in pools.items():
            pool.sort(key=lambda x: x["score"], reverse=True)
            while len(pool) > 0:
                lanes = [None] * 6
                houses_in_heat = set()
                for i in range(4):
                    for runner in pool:
                        if runner["house"] not in houses_in_heat:
                            lanes[i] = runner
                            houses_in_heat.add(runner["house"])
                            pool.remove(runner)
                            break
                extras = ["Red", "Blue"] if rotation_state % 2 == 0 else ["Green", "Yellow"]
                for i, target_house in enumerate(extras, start=4):
                    found_wildcard = False
                    for runner in pool:
                        if runner["house"] == target_house:
                            lanes[i] = runner
                            pool.remove(runner)
                            found_wildcard = True
                            break
                    if not found_wildcard and len(pool) > 0:
                        lanes[i] = pool.pop(0)

                for idx, runner in enumerate(lanes, start=1):
                    if runner:
                        self.pipeline.cursor.execute('''
                            INSERT INTO trials_active (stu_id, event, heat, lane, name, house, class, division)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (runner["id"], target_event, heat_idx, idx, runner["name"], runner["house"], runner["class"], div))
                heat_idx += 1
                rotation_state += 1
        self.pipeline.conn.commit()

# =========================================================================
# 3. ADVANCED HIERARCHICAL TREE ENGINE INTERFACE
# =========================================================================
class ModernMeetTerminal:
    def __init__(self, window, pipeline, engine, selected_event):
        self.window = window
        self.pipeline = pipeline
        self.engine = engine
        self.selected_event = selected_event
        
        self.window.title(f"Athletic Tournament Dashboard - [{self.selected_event.upper()}]")
        self.window.geometry("850x600")
        self._assemble_components()
        
    def _assemble_components(self):
        banner = tk.Label(self.window, text=f"TOURNAMENT MASTER ENGINE: {self.selected_event.upper()}", bg="#1f497d", fg="white", font=("Arial", 13, "bold"), pady=6)
        banner.pack(fill=tk.X)
        
        btn_frame = tk.Frame(self.window, pady=8)
        btn_frame.pack(fill=tk.X, padx=15)
        
        tk.Button(btn_frame, text="1. Run Seeding", command=self.trigger_seeding, bg="#2e7d32", fg="white", font=("Arial", 9, "bold"), width=18).grid(row=0, column=0, padx=4)
        tk.Button(btn_frame, text="2. Timing Desk", command=self.display_input_desk, bg="#ef6c00", fg="white", font=("Arial", 9, "bold"), width=18).grid(row=0, column=1, padx=4)
        tk.Button(btn_frame, text="3. Semifinals", command=self.trigger_semifinals, bg="#c62828", fg="white", font=("Arial", 9, "bold"), width=18).grid(row=0, column=2, padx=4)
        tk.Button(btn_frame, text="4. House Points", command=self.trigger_house_points, bg="#fbc02d", fg="black", font=("Arial", 9, "bold"), width=18).grid(row=0, column=3, padx=4)
        
        self.tree_frame = tk.Frame(self.window)
        self.tree_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        self.tree = ttk.Treeview(self.tree_frame, columns=("Lane", "Name", "House", "Class", "Time"), show="tree headings")
        self.tree.heading("#0", text="Bracket Structure / Division / Heat", anchor="w")
        self.tree.heading("Lane", text="Lane")
        self.tree.heading("Name", text="Competitor Name")
        self.tree.heading("House", text="House")
        self.tree.heading("Class", text="Class Section")
        self.tree.heading("Time", text="Stopwatch Time")
        
        self.tree.column("#0", width=250, anchor="w")
        self.tree.column("Lane", width=50, anchor="center")
        self.tree.column("Name", width=180, anchor="w")
        self.tree.column("House", width=80, anchor="center")
        self.tree.column("Class", width=100, anchor="center")
        self.tree.column("Time", width=100, anchor="center")
        
        scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self._clear_and_log_msg("System Initialized. Ready to parse school registrations.")

    def _clear_and_log_msg(self, msg):
        self.tree.delete(*self.tree.get_children())
        self.tree.insert("", "end", text=msg, values=("", "", "", "", ""))

    def trigger_seeding(self):
        if self.pipeline.load_excel_workbook():
            try:
                self.engine.execute_initial_seeding(self.selected_event)
                self.tree.delete(*self.tree.get_children())
                
                self.pipeline.cursor.execute("SELECT DISTINCT division FROM trials_active WHERE event=? ORDER BY division", (self.selected_event,))
                divisions = self.pipeline.cursor.fetchall()
                
                for (div,) in divisions:
                    div_node = self.tree.insert("", "end", text=f"🏆 DIVISION {div}", open=True)
                    self.pipeline.cursor.execute("SELECT DISTINCT heat FROM trials_active WHERE event=? AND division=? ORDER BY heat", (self.selected_event, div))
                    heats = self.pipeline.cursor.fetchall()
                    
                    for (ht,) in heats:
                        ht_node = self.tree.insert(div_node, "end", text=f"🏃 Heat {ht}", open=False)
                        self.pipeline.cursor.execute("SELECT lane, name, house, class FROM trials_active WHERE event=? AND division=? AND heat=? ORDER BY lane", (self.selected_event, div, ht))
                        runners = self.pipeline.cursor.fetchall()
                        for ln, nm, hs, cl in runners:
                            self.tree.insert(ht_node, "end", text="", values=(ln, nm, hs, cl, "0.00s"))
            except Exception as e:
                messagebox.showerror("Error", f"Seeding Error: {str(e)}")

    def display_input_desk(self):
        desk = tk.Toplevel(self.window)
        desk.title("Timing Collection Desk")
        desk.geometry("480x450")
        desk.configure(bg="#f4f4f4")
        
        tk.Label(desk, text="Type Tracked Lane Times Below:", font=("Arial", 11, "bold"), bg="#f4f4f4", pady=10).pack()
        self.pipeline.cursor.execute("SELECT stu_id, heat, lane, name, house FROM trials_active WHERE event=?", (self.selected_event,))
        rows = self.pipeline.cursor.fetchall()
        box_map = {}
        
        if not rows:
            tk.Label(desk, text="⚠️ Warning: No athletes seeded yet!", fg="red", bg="#f4f4f4").pack(pady=30)
        else:
            canvas = tk.Canvas(desk, bg="#f4f4f4", borderwidth=0, highlightthickness=0)
            scrollbar = tk.Scrollbar(desk, orient="vertical", command=canvas.yview)
            scroll_frame = tk.Frame(canvas, bg="#f4f4f4")
            scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
            canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            canvas.pack(side="left", fill="both", expand=True, padx=10)
            scrollbar.pack(side="right", fill="y")

            for uid, heat, lane, name, house in rows:
                r_ui = tk.Frame(scroll_frame, bg="white", bd=1, relief="groove", pady=5)
                r_ui.pack(fill=tk.X, padx=10, pady=2)
                tk.Label(r_ui, text=f"Heat {heat} | Lane {lane} - {name} ({house}):", anchor="w", width=34, bg="white").pack(side=tk.LEFT, padx=5)
                box = tk.Entry(r_ui, width=8, font=("Arial", 10, "bold"), justify="center")
                box.pack(side=tk.RIGHT, padx=5)
                box_map[uid] = box

        def commit_desk():
            if not box_map:
                desk.destroy()
                return
            for uid, box in box_map.items():
                val = box.get()
                if val:
                    self.pipeline.cursor.execute("UPDATE trials_active SET track_time=? WHERE stu_id=? AND event=?", (float(val), uid, self.selected_event))
            self.pipeline.conn.commit()
            messagebox.showinfo("Saved", "Times recorded successfully!")
            desk.destroy()
            
        tk.Button(desk, text="💾 SAVE TIMINGS", command=commit_desk, bg="#1f497d", fg="white", font=("Arial", 11, "bold"), pady=6).pack(side=tk.BOTTOM, fill=tk.X, padx=15, pady=15)

    def trigger_semifinals(self):
        self.pipeline.cursor.execute("DELETE FROM semifinals_output WHERE event=?", (self.selected_event,))
        self.pipeline.cursor.execute("SELECT name, house, class, track_time FROM trials_active WHERE event=? AND track_time > 0.0 ORDER BY track_time ASC", (self.selected_event,))
        ranked_runners = self.pipeline.cursor.fetchall()
        
        if not ranked_runners:
            messagebox.showwarning("No Data", "Enter timing parameters first!")
            return
            
        self.tree.delete(*self.tree.get_children())
        root_node = self.tree.insert("", "end", text=f"🏁 OFFICIAL {self.selected_event.upper()} SEMIFINALS STRUCTURE", open=True)
        
        semi_heat = 1
        semi_lane = 1
        ht_node = self.tree.insert(root_node, "end", text=f"Semifinal Heat {semi_heat}", open=True)
        
        for name, house, cls, track_time in ranked_runners:
            self.pipeline.cursor.execute("INSERT INTO semifinals_output VALUES (?, ?, ?, ?, ?, ?)", (self.selected_event, semi_heat, semi_lane, name, house, track_time))
            self.tree.insert(ht_node, "end", text="", values=(semi_lane, name, house, cls, f"{track_time}s"))
            
            semi_lane += 1
            if semi_lane > 6:
                semi_lane = 1
                semi_heat += 1
                if ranked_runners.index((name, house, cls, track_time)) < len(ranked_runners) - 1:
                    ht_node = self.tree.insert(root_node, "end", text=f"Semifinal Heat {semi_heat}", open=True)
        self.pipeline.conn.commit()

    def trigger_house_points(self):
        self.pipeline.cursor.execute("SELECT name, house, class, track_time FROM trials_active WHERE event=? AND track_time > 0.0 ORDER BY track_time ASC LIMIT 3", (self.selected_event,))
        winners = self.pipeline.cursor.fetchall()
        
        if len(winners) < 3:
            messagebox.showwarning("Incomplete", "Not enough timed entries.")
            return
            
        self.pipeline.cursor.execute("UPDATE house_points SET points = 0 WHERE event=?", (self.selected_event,))
        self.tree.delete(*self.tree.get_children())
        
        podium_node = self.tree.insert("", "end", text="🏅 MEDAL PODIUM STANDINGS", open=True)
        points_scale =[6,5,4]
        medals = ["🥇 1st (Gold)", "🥈 2nd (Silver)", "🥉 3rd (Bronze)"]
        
        for idx, (name, house, cls, time) in enumerate(winners):
            pts = points_scale[idx]
            self.pipeline.cursor.execute("UPDATE house_points SET points = points + ? WHERE house = ? AND event = ?", (pts, house, self.selected_event))
            self.tree.insert(podium_node, "end", text=medals[idx], values=("", name, house, cls, f"{time}s (+{pts} Pts)"))
            
        self.pipeline.conn.commit()
        
        leaderboard_node = self.tree.insert("", "end", text="📊 FINAL HOUSE STANDINGS LEADERBOARD", open=True)
        self.pipeline.cursor.execute("SELECT house, points FROM house_points WHERE event=? ORDER BY points DESC", (self.selected_event,))
        standings = self.pipeline.cursor.fetchall()
        for h, pts in standings:
            self.tree.insert(leaderboard_node, "end", text=f"  House {h}", values=("", "", "", "", f"{pts} Total Points"))

# =========================================================================
# 4. PRIMARY EVENT SELECTION GATEWAY LAUNCHER SCREEN
# =========================================================================
class EventLauncherGateway:
    def __init__(self, root_window):
        self.root = root_window
        self.root.title("Athletic Launcher Gateway")
        self.root.geometry("400x250")
        self.root.configure(bg="#f4f4f4")
        
        self.file_pipeline = AthleticExcelPipeline()
        self.logic_engine = MeetLogicProcessor(self.file_pipeline)
        self._build_launcher_ui()
        
    def _build_launcher_ui(self):
        title_lbl = tk.Label(self.root, text="MEET TOURNAMENT GATEWAY", bg="#1f497d", fg="white", font=("Arial", 12, "bold"), pady=10)
        title_lbl.pack(fill=tk.X)
        
        msg_lbl = tk.Label(self.root, text="Select the race event you wish to manage:", bg="#f4f4f4", font=("Arial", 10), pady=15)
        msg_lbl.pack()
        
        self.event_combobox = ttk.Combobox(self.root, values=["100m", "200m"], state="readonly", font=("Arial", 11), justify="center")
        self.event_combobox.set("100m")
        self.event_combobox.pack(pady=5)
        
        launch_btn = tk.Button(self.root, text="🚀 LAUNCH ENGINE TERMINAL", command=self.boot_main_dashboard, bg="#2e7d32", fg="white", font=("Arial", 10, "bold"), pady=8, padx=15)
        launch_btn.pack(pady=20)
        
    def boot_main_dashboard(self):
        choice = self.event_combobox.get()
        self.root.withdraw()
        dashboard_window = tk.Toplevel()
        dashboard_window.protocol("WM_DELETE_WINDOW", lambda: self.root.destroy())
        ModernMeetTerminal(dashboard_window, self.file_pipeline, self.logic_engine, choice)

if __name__ == "__main__":
    launcher_root = tk.Tk()
    app_gateway = EventLauncherGateway(launcher_root)
    launcher_root.mainloop()
