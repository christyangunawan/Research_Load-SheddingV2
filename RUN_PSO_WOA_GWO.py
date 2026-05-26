import sys
import os
import time
import random
import numpy as np
import pandas as pd
import re

# --- KONFIGURASI SISTEM ---
TARGET_GEN_TRIP = "DG_3"
EXCEL_OUTPUT = "DG 3_RUN 6.xlsx"
TOTAL_RUNS = 1  # Jumlah eksekusi algoritma berulang (otomatis bertambah)

# Parameter Optimasi
SEARCH_AGENTS = 25
MAX_ITER = 50
DIMENSION = 33

USE_HARDCODED_TARGET = True
TARGET_DEFICIT_MW = 1.200

# Parameter Algoritma
W_MAX = 0.9
W_MIN = 0.4
C1_PSO = 2.0
C2_PSO = 2.0
VOLT_MIN = 0.95
VOLT_MAX = 1.05
TRANSFER_FUNCTION = "v_shape"  # Pilihan: "sigmoid" atau "v_shape"

# Parameter Simulasi RMS
RMS_TSTOP = 100.0        # Total waktu simulasi (detik)
RMS_TSAMPLE = 30.0       # Waktu pengambilan data tegangan (detik)
TIME_GEN_TRIP = 1.0      # Waktu generator trip di simulasi RMS (detik)
TIME_LOAD_SHED = 1.46    # Waktu pelepasan beban di simulasi RMS (detik)

load_shed_events = []    # List untuk menyimpan event load shedding

# Parameter Frekuensi Steady-State
FREQ_TOLERANCE = 0.05    # Toleransi steady-state (±Hz dari nilai akhir)
FREQ_SETTLE_WINDOW = 2.0 # Frekuensi harus bertahan dalam band selama (detik)
FREQ_AVG_WINDOW = 5.0    # Rata-rata frekuensi N detik terakhir sebagai nilai final
FREQ_REF_BUS_IDX = 0     # Indeks bus referensi untuk frekuensi (bus pertama)
VOLT_MONITOR_BUS_IDX = 17 # Indeks bus untuk monitoring tegangan vs waktu di export

# Path DIgSILENT PowerFactory
VERSION_PYTHON = "3.12"
PATH_APP = r"C:\Program Files\DIgSILENT\PowerFactory 2024"
PATH_API = fr"{PATH_APP}\Python\{VERSION_PYTHON}"
PROJECT_NAME = "Import(8)"

if PATH_API not in sys.path: sys.path.append(PATH_API)
os.environ['PATH'] = PATH_APP + ";" + os.environ['PATH']

try:
    import powerfactory
except ImportError:
    print("Error: Module powerfactory tidak ditemukan. Cek PATH API Anda.")
    exit()

app = powerfactory.GetApplication()
app.ActivateProject(PROJECT_NAME)
try:
    app.EchoOff() # Mematikan update teks/konsol di PowerFactory (sangat mempercepat simulasi)
except:
    pass

# Deklarasi variabel global yang akan diisi oleh fungsi initialize_network_objects()
inc, sim, elmres = None, None, None
all_loads, all_buses, all_gens = [], [], []
COST_MAP, ZONE_LABELS = [], []


# --- FUNGSI PENDUKUNG ---
def natural_keys(text):
    return [int(c) if c.isdigit() else c for c in re.split(r'(\d+)', text)]


def sigmoid(x):
    return 1.0 / (1.0 + np.exp(-x))


def v_shape(x):
    return np.abs(np.tanh(x))


def setup_rms_simulation_and_events():
    """Menyiapkan result variables dan dynamic events (trip) untuk RMS."""
    global load_shed_events
    # Link result object ke ComInc agar data terekam saat simulasi
    inc.p_resvar = elmres

    elmres.Clear()
    for bus in all_buses:
        elmres.AddVars(bus, "m:u")

    # Tambahkan frekuensi pada bus referensi
    ref_bus = all_buses[FREQ_REF_BUS_IDX]
    elmres.AddVars(ref_bus, "m:fehz")

    # --- Setup Dynamic Events ---
    sc = app.GetActiveStudyCase()
    evt_folder = app.GetFromStudyCase("IntEvt")
    if not evt_folder:
        evt_folder = sc.CreateObject("IntEvt", "Events")

    # Bersihkan event lama jika ada
    for evt in evt_folder.GetContents():
        evt.Delete()

    # Event Generator Trip
    target_gen = next((g for g in all_gens if g.GetAttribute("loc_name") in [TARGET_GEN_TRIP, TARGET_GEN_TRIP.replace("_", " ")]), None)
    if target_gen:
        gen_evt = evt_folder.CreateObject("EvtSwitch", f"Trip Gen {TARGET_GEN_TRIP}")
        gen_evt.p_target = target_gen
        gen_evt.time = TIME_GEN_TRIP
        try:
            gen_evt.i_all = 0
            gen_evt.i_switch = 0
        except: pass

    # Event Load Shedding
    load_shed_events = []
    for load in all_loads:
        load_evt = evt_folder.CreateObject("EvtSwitch", f"Shed {load.GetAttribute('loc_name')}")
        load_evt.p_target = load
        load_evt.time = TIME_LOAD_SHED
        try:
            load_evt.i_all = 0
            load_evt.i_switch = 0
        except: pass
        load_evt.outserv = 1  # Disable secara default
        load_shed_events.append(load_evt)

    print(f"RMS Configured: Voltage on {len(all_buses)} buses, Freq on '{ref_bus.GetAttribute('loc_name')}'.")
    print(f"RMS Events Configured: Gen Trip at {TIME_GEN_TRIP}s, Load Shed at {TIME_LOAD_SHED}s.")


def extract_rms_results():
    
    elmres.Load()
    n_rows = elmres.GetNumberOfRows()

    if n_rows == 0:
        elmres.Release()
        return [0.0], RMS_TSTOP

    ref_bus = all_buses[FREQ_REF_BUS_IDX]
    freq_col = elmres.FindColumn(ref_bus, "m:fehz")

    volt_cols = []
    for bus in all_buses:
        volt_cols.append(elmres.FindColumn(bus, "m:u"))

    # 1. Binary Search untuk mendapatkan baris terdekat dengan RMS_TSAMPLE (30s)
    low = 0
    high = n_rows - 1
    best_row = 0
    min_diff = float('inf')

    while low <= high:
        mid = (low + high) // 2
        t_val = elmres.GetValue(mid, -1)
        t = t_val[1] if isinstance(t_val, (list, tuple)) else t_val
        
        diff = t - RMS_TSAMPLE
        abs_diff = abs(diff)
        
        if abs_diff < min_diff:
            min_diff = abs_diff
            best_row = mid
            
        if diff == 0:
            break
        elif diff < 0:
            low = mid + 1
        else:
            high = mid - 1

    # Ambil tegangan HANYA pada best_row
    voltages = []
    for col_idx in volt_cols:
        if col_idx >= 0:
            val = elmres.GetValue(best_row, col_idx)
            voltages.append(val[1] if isinstance(val, (list, tuple)) else val)
    if not voltages:
        voltages = [0.0]

    # 2. Full-read frekuensi untuk settle time (presisi penuh, tidak ada data yang terlewat)
    settle_time = RMS_TSTOP
    if freq_col >= 0:
        times = []
        freqs = []

        for row in range(n_rows):
            t_val = elmres.GetValue(row, -1)
            t = t_val[1] if isinstance(t_val, (list, tuple)) else t_val
            times.append(t)

            f_val = elmres.GetValue(row, freq_col)
            freqs.append(f_val[1] if isinstance(f_val, (list, tuple)) else f_val)

        t_end = times[-1]
        final_freqs = [freqs[i] for i in range(len(times)) if times[i] >= (t_end - FREQ_AVG_WINDOW)]
        f_final = sum(final_freqs) / len(final_freqs) if final_freqs else freqs[-1]

        f_min = f_final - FREQ_TOLERANCE
        f_max = f_final + FREQ_TOLERANCE

        for i in range(len(times)):
            if f_min <= freqs[i] <= f_max:
                settle_ok = True
                window_covered = False
                for j in range(i, len(times)):
                    if times[j] - times[i] >= FREQ_SETTLE_WINDOW:
                        window_covered = True
                        break
                    if not (f_min <= freqs[j] <= f_max):
                        settle_ok = False
                        break
                if settle_ok and window_covered:
                    settle_time = times[i]
                    break

    elmres.Release()
    return voltages, settle_time


def simulate_and_extract_timeseries(pattern):
    """Re-simulate pola terbaik dan ekstrak time series frekuensi & tegangan."""
    # Terapkan pola load shedding ke event (bukan ke beban langsung)
    for i in range(len(all_loads)):
        load_shed_events[i].outserv = 0 if pattern[i] == 1 else 1

    # Jalankan simulasi RMS
    err_inc = inc.Execute()
    if err_inc != 0:
        return None, None
    err_sim = sim.Execute()
    if err_sim != 0:
        return None, None

    elmres.Load()
    n_rows = elmres.GetNumberOfRows()
    if n_rows == 0:
        elmres.Release()
        return None, None

    ref_bus = all_buses[FREQ_REF_BUS_IDX]
    freq_col = elmres.FindColumn(ref_bus, "m:fehz")
    mon_bus = all_buses[VOLT_MONITOR_BUS_IDX]
    volt_col = elmres.FindColumn(mon_bus, "m:u")

    times, freqs, volts = [], [], []
    for row in range(n_rows):
        t_val = elmres.GetValue(row, -1)
        t = t_val[1] if isinstance(t_val, (list, tuple)) else t_val
        times.append(t)

        if freq_col >= 0:
            f_val = elmres.GetValue(row, freq_col)
            freqs.append(f_val[1] if isinstance(f_val, (list, tuple)) else f_val)
        if volt_col >= 0:
            v_val = elmres.GetValue(row, volt_col)
            volts.append(v_val[1] if isinstance(v_val, (list, tuple)) else v_val)

    freq_df = pd.DataFrame({"Time_s": times, "Frequency_Hz": freqs}) if freqs else None
    volt_df = pd.DataFrame({"Time_s": times, f"Voltage_{mon_bus.GetAttribute('loc_name')}_pu": volts}) if volts else None

    elmres.Release()
    return freq_df, volt_df

def dominates(obj1, obj2):
    """Mengembalikan True jika obj1 mendominasi obj2 (Multi-Objective)."""
    return all(obj1[k] <= obj2[k] for k in range(len(obj1))) and any(obj1[k] < obj2[k] for k in range(len(obj1)))


# --- METRIK MOO ---
def normalize_pareto(pareto):
    objs = np.array([s["objectives"] for s in pareto])
    min_vals = objs.min(axis=0)
    max_vals = objs.max(axis=0)
    return (objs - min_vals) / (max_vals - min_vals + 1e-9)


def compute_hypervolume(pareto):
    if not pareto: return 0
    norm = normalize_pareto(pareto)
    n_obj = norm.shape[1]
    ref = np.full(n_obj, 1.1)
    return sum(np.prod(ref - p) for p in norm)


def compute_spacing(pareto):
    if len(pareto) < 2: return 0
    objs = np.array([s["objectives"] for s in pareto])
    distances = []
    for i in range(len(objs)):
        d = min(np.sum(np.abs(objs[i] - objs[j])) for j in range(len(objs)) if i != j)
        distances.append(d)
    return np.std(distances)


def find_knee(pareto):
    if not pareto: return None
    norm = normalize_pareto(pareto)
    ideal = np.zeros(norm.shape[1])
    dist = [np.linalg.norm(p - ideal) for p in norm]
    return pareto[int(np.argmin(dist))]


def initialize_network_objects():
    global all_loads, all_buses, all_gens, COST_MAP, ZONE_LABELS, inc, sim, elmres
    
    # Inisialisasi perintah simulasi RMS
    inc = app.GetFromStudyCase("ComInc")
    sim = app.GetFromStudyCase("ComSim")
    elmres = app.GetFromStudyCase("All calculations.ElmRes")
    sim.tstop = RMS_TSTOP
    
    raw_loads = app.GetCalcRelevantObjects("*.ElmLod")
    all_loads = [obj for obj in raw_loads if obj.GetClassName() == "ElmLod"]
    all_loads.sort(key=lambda x: natural_keys(x.GetAttribute("loc_name")))

    all_buses = app.GetCalcRelevantObjects("*.ElmTerm")
    all_buses.sort(key=lambda x: natural_keys(x.GetAttribute("loc_name")))

    all_gens = app.GetCalcRelevantObjects("*.ElmSym") + app.GetCalcRelevantObjects("*.ElmGenstat")

    # Mapping Prioritas Zona
    COST_MAP = []
    ZONE_LABELS = []
    for i in range(len(all_loads)):
        if 0 <= i <= 4:
            cost, zone = 1000.0, "Zone 1"
        elif 5 <= i <= 16:
            cost, zone = 50.0, "Zone 2"
        elif 17 <= i <= 23:
            cost, zone = 200.0, "Zone 3"
        else:
            cost, zone = 500.0, "Zone 4"
        COST_MAP.append(cost)
        ZONE_LABELS.append(zone)


# --- CORE LOGIC ---
def apply_contingency_and_reset():
    print("Resetting System (Ensuring all elements are in-service)...")
    for l in all_loads: l.SetAttribute("outserv", 0)
    for g in all_gens: g.SetAttribute("outserv", 0)
    return TARGET_DEFICIT_MW


def calculate_fitness(position_continuous, min_shed_required):
    if TRANSFER_FUNCTION == "v_shape":
        probs = v_shape(position_continuous)
    else:
        probs = sigmoid(position_continuous)
        
    pattern = (np.random.rand(len(position_continuous)) < probs).astype(int)

    # Terapkan pola load shedding ke dynamic events (Aktifkan event jika pattern = 1)
    for i in range(len(all_loads)):
        load_shed_events[i].outserv = 0 if pattern[i] == 1 else 1

    # Jalankan simulasi RMS (Initial Conditions + Simulation)
    err_inc = inc.Execute()
    if err_inc == 0:
        err_sim = sim.Execute()
    else:
        err_sim = 1
    err = 0 if (err_inc == 0 and err_sim == 0) else 1

    # Ambil tegangan dan settling time dalam satu pass baca
    if err == 0:
        voltages, settle_time = extract_rms_results()
    else:
        voltages, settle_time = [0.0], RMS_TSTOP

    total_mw = sum(all_loads[i].GetAttribute("plini") for i, val in enumerate(pattern) if val == 1)
    cost = sum(all_loads[i].GetAttribute("plini") * COST_MAP[i] for i, val in enumerate(pattern) if val == 1)
    vdi = sum((1.0 - v) ** 2 for v in voltages) if err == 0 else 100.0
    over = max(0.0, total_mw - min_shed_required)

    feasible = (err == 0) and (total_mw >= (min_shed_required - 0.001)) and (min(voltages) >= VOLT_MIN)

    return {
        "objectives": [cost, vdi, over, settle_time],
        "MW": total_mw, "Cost": cost, "VDI": vdi, "Over": over, "Settle": settle_time,
        "feasible": feasible, "voltages": voltages, "pattern": pattern, "min_v": min(voltages) if voltages else 0,
        "continuous": position_continuous.copy()
    }


def update_archive(archive, candidate):
    if not candidate["feasible"]: return archive
    n_obj = len(candidate["objectives"])
    # Dominance Check
    new_archive = [sol for sol in archive if not (
                all(candidate["objectives"][k] <= sol["objectives"][k] for k in range(n_obj)) and any(
            candidate["objectives"][k] < sol["objectives"][k] for k in range(n_obj)))]
    if not any(all(sol["objectives"][k] <= candidate["objectives"][k] for k in range(n_obj)) and any(
            sol["objectives"][k] < candidate["objectives"][k] for k in range(n_obj)) for sol in archive):
        new_archive.append(candidate)
    return new_archive


class Optimizer:
    def __init__(self, algo_name, min_shed_mw):
        self.algo_name, self.min_shed_mw = algo_name, min_shed_mw
        self.dim = len(all_loads)
        self.archive, self.detailed_log, self.convergence_curve = [], [], []

        self.X = np.random.uniform(-5, 5, (SEARCH_AGENTS, self.dim))
        self.V = np.zeros((SEARCH_AGENTS, self.dim))
        self.P_best_pos = self.X.copy()
        self.P_best_score = np.full(SEARCH_AGENTS, float('inf'))
        self.P_best_objs = np.full((SEARCH_AGENTS, 4), float('inf'))
        self.G_best_pos = np.zeros(self.dim)
        self.G_best_metrics = {'Cost': 0, 'MW': 0, 'VDI': 0, 'Over': 0, 'Vmin': 0, 'Settle': 0}

        # GWO: 3 leaders (alpha, beta, delta)
        self.alpha_pos = np.zeros(self.dim)
        self.beta_pos = np.zeros(self.dim)
        self.delta_pos = np.zeros(self.dim)

    def run(self):
        start_time = time.time()
        for t in range(MAX_ITER):
            a = 2.0 - t * (2.0 / MAX_ITER)
            w = W_MAX - t * ((W_MAX - W_MIN) / MAX_ITER)

            for i in range(SEARCH_AGENTS):
                sol = calculate_fitness(self.X[i], self.min_shed_mw)
                self.archive = update_archive(self.archive, sol)

                # Update P_best (Menggunakan Pareto Dominance untuk MOPSO)
                if sol["feasible"]:
                    if self.P_best_score[i] == float('inf'): # Belum ada P_best
                        self.P_best_score[i] = sol["Cost"]
                        self.P_best_objs[i] = sol["objectives"]
                        self.P_best_pos[i] = self.X[i].copy()
                    elif dominates(sol["objectives"], self.P_best_objs[i]): # Dominasi P_best lama
                        self.P_best_score[i] = sol["Cost"]
                        self.P_best_objs[i] = sol["objectives"]
                        self.P_best_pos[i] = self.X[i].copy()
                    elif not dominates(self.P_best_objs[i], sol["objectives"]): # Non-dominated
                        if random.random() < 0.5: # 50% peluang menggantikan
                            self.P_best_score[i] = sol["Cost"]
                            self.P_best_objs[i] = sol["objectives"]
                            self.P_best_pos[i] = self.X[i].copy()

                self.detailed_log.append({
                    "Iteration": t + 1, "Agent_ID": i + 1,
                    "Cost": sol["Cost"], "VDI": sol["VDI"], "Overshed": sol["Over"],
                    "Settle": sol["Settle"],
                    "Feasible": sol["feasible"], "Pattern": "".join(map(str, sol["pattern"]))
                })


            if self.archive:
                leader = find_knee(self.archive)
                self.G_best_pos = leader["continuous"].copy()
                self.G_best_metrics = {
                    'MW': leader["MW"], 'Cost': leader["Cost"],
                    'VDI': leader["VDI"], 'Over': leader["Over"], 'Vmin': leader["min_v"],
                    'Settle': leader["Settle"]
                }
                
                # Update Leader GWO (Alpha, Beta, Delta) dari Pareto Archive
                if self.algo_name == "GWO":
                    if len(self.archive) >= 3:
                        arch_copy = [s for s in self.archive if s is not leader]
                        random.shuffle(arch_copy)
                        self.alpha_pos = leader["continuous"].copy()
                        self.beta_pos = arch_copy[0]["continuous"].copy() if arch_copy else self.alpha_pos.copy()
                        self.delta_pos = arch_copy[1]["continuous"].copy() if len(arch_copy) > 1 else self.beta_pos.copy()
                    elif len(self.archive) == 2:
                        self.alpha_pos = self.archive[0]["continuous"].copy()
                        self.beta_pos = self.archive[1]["continuous"].copy()
                        self.delta_pos = self.beta_pos.copy()
                    else:
                        self.alpha_pos = self.archive[0]["continuous"].copy()
                        self.beta_pos = self.alpha_pos.copy()
                        self.delta_pos = self.alpha_pos.copy()

            self.convergence_curve.append(self.G_best_metrics['Cost'])
            print(
                f"[{self.algo_name}] Iter {t + 1} | Pareto Size: {len(self.archive)} | Knee Cost: {self.G_best_metrics['Cost']:.2f}")

            # Update Posisi
            for i in range(SEARCH_AGENTS):
                r1, r2 = np.random.random(), np.random.random()
                if self.algo_name == "PSO":
                    self.V[i] = (w * self.V[i]) + (C1_PSO * r1 * (self.P_best_pos[i] - self.X[i])) + (
                                C2_PSO * r2 * (self.G_best_pos - self.X[i]))
                    self.X[i] += np.clip(self.V[i], -5, 5)
                elif self.algo_name == "GWO":
                    # Alpha guidance
                    A1 = 2.0 * a * np.random.random(self.dim) - a
                    C1 = 2.0 * np.random.random(self.dim)
                    D_alpha = np.abs(C1 * self.alpha_pos - self.X[i])
                    X1 = self.alpha_pos - A1 * D_alpha

                    # Beta guidance
                    A2 = 2.0 * a * np.random.random(self.dim) - a
                    C2 = 2.0 * np.random.random(self.dim)
                    D_beta = np.abs(C2 * self.beta_pos - self.X[i])
                    X2 = self.beta_pos - A2 * D_beta

                    # Delta guidance
                    A3 = 2.0 * a * np.random.random(self.dim) - a
                    C3 = 2.0 * np.random.random(self.dim)
                    D_delta = np.abs(C3 * self.delta_pos - self.X[i])
                    X3 = self.delta_pos - A3 * D_delta

                    # Position update: average of 3 leaders
                    self.X[i] = (X1 + X2 + X3) / 3.0
                elif self.algo_name == "WOA":
                    p, l = np.random.random(), np.random.uniform(-1, 1)
                    if p < 0.5:
                        A = 2.0 * a * np.random.random() - a
                        D = np.abs(2.0 * np.random.random() * self.G_best_pos - self.X[i])
                        self.X[i] = self.G_best_pos - A * D
                    else:
                        D_p = np.abs(self.G_best_pos - self.X[i])
                        self.X[i] = D_p * np.exp(1 * l) * np.cos(2 * np.pi * l) + self.G_best_pos
            
            # --- Clear Cache PowerFactory ---
            # Hal ini dilakukan untuk mencegah penggunaan memori yang terus membengkak (berat) 
            # akibat data simulasi dan objek yang menumpuk.
            app.ResetCalculation()
            app.ClearRecycleBin()

        return self.G_best_metrics, self.convergence_curve, time.time() - start_time, self.detailed_log, self.G_best_pos, self.archive

def export_to_excel(results_dict, output_file):
    """Mengekspor semua hasil yang ada di memori ke file Excel. Aman dipanggil berkali-kali."""
    if not results_dict:
        print("Tidak ada data untuk di-export.")
        return
        
    print(f"Exporting (Auto-Save) to {output_file}...")
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 1. Summary Sheet
            summary_rows = []
            for a, data in results_dict.items():
                if "Metrics" in data:
                    summary_rows.append({
                        "Algorithm": a, 
                        "Knee_Cost": data["Metrics"].get("Cost", 0),
                        "Knee_Vmin": data["Metrics"].get("Vmin", 0), 
                        "Time_Sec": data.get("Time", 0)
                    })
            if summary_rows:
                pd.DataFrame(summary_rows).to_excel(writer, sheet_name='Summary', index=False)

            # 2. Pareto, MOO Metrics & Logs
            for algo, data in results_dict.items():
                pareto = data.get("Pareto", [])
                if pareto:
                    pd.DataFrame([{
                        "Cost": s["objectives"][0], "VDI": s["objectives"][1],
                        "Overshed": s["objectives"][2], "Settle_Time": s["objectives"][3], "MW": s["MW"]
                    } for s in pareto]).to_excel(writer, sheet_name=f"Pareto_{algo}", index=False)

                    hv, sp, knee = compute_hypervolume(pareto), compute_spacing(pareto), find_knee(pareto)
                    if knee:
                        pd.DataFrame([{
                            "Hypervolume": hv, "Spacing": sp,
                            "Knee_Cost": knee["objectives"][0], "Knee_VDI": knee["objectives"][1],
                            "Knee_Settle": knee["objectives"][3]
                        }]).to_excel(writer, sheet_name=f"MOO_{algo}", index=False)

                if "Log" in data and data["Log"]:
                    pd.DataFrame(data["Log"]).to_excel(writer, sheet_name=f"Log_{algo}", index=False)

            # 3. GBest Detail Sheets
            for algo, data in results_dict.items():
                pareto = data.get("Pareto", [])
                if not pareto: continue
                
                knee = find_knee(pareto)
                if not knee: continue

                best_pattern = knee["pattern"]
                shed_rows = []
                for i, load in enumerate(all_loads):
                    shed_rows.append({
                        "Load_Name": load.GetAttribute("loc_name"),
                        "Zone": ZONE_LABELS[i],
                        "Cost_Weight": COST_MAP[i],
                        "MW": load.GetAttribute("plini"),
                        "Shed": "YES" if best_pattern[i] == 1 else "NO"
                    })

                shed_df = pd.DataFrame(shed_rows)
                summary_row = pd.DataFrame([{
                    "Load_Name": "--- TOTAL ---", "Zone": "", "Cost_Weight": "",
                    "MW": knee["MW"],
                    "Shed": f"Cost={knee['Cost']:.2f}, VDI={knee['VDI']:.6f}, Over={knee['Over']:.3f}, Settle={knee['Settle']:.2f}s"
                }])
                shed_df = pd.concat([shed_df, summary_row], ignore_index=True)
                shed_df.to_excel(writer, sheet_name=f"GBest_{algo}", index=False)

                # Re-simulate untuk time series
                print(f"Re-simulating G_best for {algo} time-series export...")
                freq_df, volt_df = simulate_and_extract_timeseries(best_pattern)
                if freq_df is not None:
                    freq_df.to_excel(writer, sheet_name=f"Freq_{algo}", index=False)
                if volt_df is not None:
                    volt_df.to_excel(writer, sheet_name=f"Volt_{algo}", index=False)
                    
        print(f"Export berhasil. Data tersimpan di {output_file}")
    except Exception as e:
        print(f"Gagal melakukan export ke Excel: {e}")


if __name__ == "__main__":
    # Parse nama file dasar dan nomor RUN awal
    match = re.search(r'(.*?_RUN\s*)(\d+)(.*)', EXCEL_OUTPUT, re.IGNORECASE)
    if match:
        base_prefix = match.group(1)
        start_run = int(match.group(2))
        extension = match.group(3)
    else:
        base_prefix = EXCEL_OUTPUT.replace(".xlsx", "_RUN_")
        start_run = 1
        extension = ".xlsx"

    # Penentuan base nama project (misal: "Import(8)" -> base="Import(", idx=8, suffix=")")
    match_proj = re.search(r'(.*?\(?)(\d+)(\)?)', PROJECT_NAME)
    if match_proj:
        proj_base = match_proj.group(1)
        start_proj_idx = int(match_proj.group(2))
        proj_suffix = match_proj.group(3)
    else:
        proj_base = PROJECT_NAME
        start_proj_idx = 1
        proj_suffix = ""

    last_active_project = PROJECT_NAME

    for current_run in range(start_run, start_run + TOTAL_RUNS):
        current_excel_output = f"{base_prefix}{current_run}{extension}"
        
        if match_proj:
            current_proj_idx = start_proj_idx + (current_run - start_run)
            current_proj_name = f"{proj_base}{current_proj_idx}{proj_suffix}"
        else:
            current_proj_name = PROJECT_NAME

        print(f"\n{'='*50}\nMEMULAI SIMULASI {current_excel_output.upper()} ({current_run - start_run + 1}/{TOTAL_RUNS})\n{'='*50}")
        print(f"Mencoba mengaktifkan project: {current_proj_name}")
        
        err = app.ActivateProject(current_proj_name)
        if err != 0:
            print(f"[WARNING] Project '{current_proj_name}' gagal diaktifkan / tidak ditemukan.")
            print(f"-> Fallback: Tetap menggunakan project '{last_active_project}'")
            app.ActivateProject(last_active_project)
            current_proj_name = last_active_project
        else:
            print(f"-> Berhasil mengaktifkan '{current_proj_name}'")
            last_active_project = current_proj_name

        # Re-inisialisasi objek network untuk project yang aktif saat ini
        initialize_network_objects()
        
        calc_target = apply_contingency_and_reset()
        setup_rms_simulation_and_events()

        results_store = {}
        algos = ["PSO", "WOA", "GWO"]

        try:
            for algo in algos:
                opt = Optimizer(algo, calc_target)
                metrics, curve, duration, log, best_pat, pareto = opt.run()
                results_store[algo] = {
                    "Metrics": metrics, "Curve": curve, "Time": duration,
                    "Log": log, "Best_Pattern": best_pat, "Pareto": pareto
                }
                export_to_excel(results_store, current_excel_output)
                
        except KeyboardInterrupt:
            print(f"\n[!] Eksekusi {current_excel_output} dihentikan paksa (Ctrl+C).")
            export_to_excel(results_store, current_excel_output)
            break
            
        except Exception as e:
            print(f"\n[!] Error saat menjalankan {current_excel_output}: {e}")
            import traceback
            traceback.print_exc()
            export_to_excel(results_store, current_excel_output)
            continue

    print("\nSeluruh Proses Berurutan Selesai Berhasil.")