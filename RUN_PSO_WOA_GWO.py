import sys
import os
import time
import random
import numpy as np
import pandas as pd
import re

# --- KONFIGURASI SISTEM ---
TARGET_GEN_TRIP = "DG_3"
EXCEL_OUTPUT = "PARETO_Comparison_Result_Breakdown_MOO.xlsx"

# Parameter Optimasi
SEARCH_AGENTS = 30
MAX_ITER = 50
TARGET_DEFICIT_MW = 1.200

# Parameter Algoritma
W_MAX, W_MIN = 0.9, 0.4
C1_PSO, C2_PSO = 2.0, 2.0
VOLT_MIN, VOLT_MAX = 0.95, 1.05

# Path DIgSILENT PowerFactory
VERSION_PYTHON = "3.12"
PATH_APP = r"C:\Program Files\DIgSILENT\PowerFactory 2024"
PATH_API = fr"{PATH_APP}\Python\{VERSION_PYTHON}"
PROJECT_NAME = "Import(1)"

if PATH_API not in sys.path: sys.path.append(PATH_API)
os.environ['PATH'] = PATH_APP + ";" + os.environ['PATH']

try:
    import powerfactory
except ImportError:
    print("Error: Module powerfactory tidak ditemukan.")
    exit()

app = powerfactory.GetApplication()
app.ActivateProject(PROJECT_NAME)
ldf = app.GetFromStudyCase("ComLdf")

# --- INISIALISASI OBJEK ---
raw_loads = app.GetCalcRelevantObjects("*.ElmLod")
all_loads = [obj for obj in raw_loads if obj.GetClassName() == "ElmLod"]
# Mengurutkan beban secara alami (Bus 1, Bus 2, dst)
all_loads.sort(key=lambda x: [int(c) if c.isdigit() else c for c in re.split(r'(\d+)', x.GetAttribute("loc_name"))])

# FIX: Menentukan DIMENSION secara dinamis berdasarkan beban yang terdeteksi
DIMENSION = len(all_loads)
print(f"Sistem Mendeteksi {DIMENSION} Beban Aktif.")

all_buses = app.GetCalcRelevantObjects("*.ElmTerm")
all_buses.sort(key=lambda x: [int(c) if c.isdigit() else c for c in re.split(r'(\d+)', x.GetAttribute("loc_name"))])
all_gens = app.GetCalcRelevantObjects("*.ElmSym") + app.GetCalcRelevantObjects("*.ElmGenstat")

# Mapping Biaya per Zona
COST_MAP = []
for i in range(DIMENSION):
    if 0 <= i <= 4:
        COST_MAP.append(1000.0)
    elif 5 <= i <= 16:
        COST_MAP.append(50.0)
    elif 17 <= i <= 23:
        COST_MAP.append(200.0)
    else:
        COST_MAP.append(500.0)


# --- FUNGSI PENDUKUNG ---
def sigmoid(x): return 1.0 / (1.0 + np.exp(-x))


def calculate_fitness(pattern_binary, min_shed_required):
    pattern = pattern_binary.astype(int)
    for i, load in enumerate(all_loads):
        load.SetAttribute("outserv", int(pattern[i]))

    err = ldf.Execute()
    voltages = [b.GetAttribute("m:u") for b in all_buses]
    total_mw = sum(all_loads[i].GetAttribute("plini") for i, val in enumerate(pattern) if val == 1)
    cost = sum(all_loads[i].GetAttribute("plini") * COST_MAP[i] for i, val in enumerate(pattern) if val == 1)
    vdi = sum((1.0 - v) ** 2 for v in voltages) if err == 0 else 100.0
    over = max(0.0, total_mw - min_shed_required)

    const_shed = (total_mw >= (min_shed_required - 0.001))
    const_volt = (min(voltages) >= VOLT_MIN and max(voltages) <= VOLT_MAX) if voltages else False

    return {
        "objectives": [cost, vdi, over], "MW": total_mw, "Cost": cost, "VDI": vdi, "Over": over,
        "feasible": (const_shed and const_volt), "const_shed": const_shed, "const_volt": const_volt,
        "min_v": min(voltages) if voltages else 0, "max_v": max(voltages) if voltages else 0,
        "pattern": pattern
    }


def update_archive(archive, candidate):
    if candidate is None or not candidate["feasible"]: return archive
    new_archive = []
    dominated = False
    for sol in archive:
        if all(sol["objectives"][k] <= candidate["objectives"][k] for k in range(3)) and \
                any(sol["objectives"][k] < candidate["objectives"][k] for k in range(3)):
            dominated = True;
            break
        elif all(candidate["objectives"][k] <= sol["objectives"][k] for k in range(3)) and \
                any(candidate["objectives"][k] < sol["objectives"][k] for k in range(3)):
            continue
        else:
            new_archive.append(sol)
    if not dominated: new_archive.append(candidate)
    return new_archive


def find_knee(pareto):
    if not pareto: return None
    objs = np.array([s["objectives"] for s in pareto])
    # Normalisasi untuk perhitungan jarak Euclidean yang adil
    norm = (objs - objs.min(axis=0)) / (objs.max(axis=0) - objs.min(axis=0) + 1e-9)
    dist = [np.linalg.norm(p) for p in norm]
    return pareto[np.argmin(dist)]


# --- OPTIMIZER CLASS ---
class Optimizer:
    def __init__(self, algo_name, min_shed_mw):
        self.algo_name, self.min_shed_mw = algo_name, min_shed_mw
        self.dim = DIMENSION
        self.archive, self.archive_history, self.detailed_log = [], [], []
        # Inisialisasi posisi agen
        if algo_name == "NSGA-II":
            self.X = np.random.randint(0, 2, (SEARCH_AGENTS, self.dim))
        else:
            self.X = np.random.uniform(-5, 5, (SEARCH_AGENTS, self.dim))

        self.V = np.zeros((SEARCH_AGENTS, self.dim))
        self.P_best_pos = self.X.copy()
        self.P_best_score = np.full(SEARCH_AGENTS, float('inf'))
        self.G_best_pos = np.zeros(self.dim)
        self.G_best_metrics = {'Cost': 0, 'Vmin': 0}

    def run(self):
        for t in range(MAX_ITER):
            a = 2.0 - t * (2.0 / MAX_ITER)
            w = W_MAX - t * ((W_MAX - W_MIN) / MAX_ITER)

            current_pop_data = []
            for i in range(SEARCH_AGENTS):
                if self.algo_name != "NSGA-II":
                    binary_pattern = (np.random.rand(self.dim) < sigmoid(self.X[i])).astype(int)
                else:
                    binary_pattern = self.X[i].astype(int)

                sol = calculate_fitness(binary_pattern, self.min_shed_mw)
                current_pop_data.append(sol)

                # Maintenance Archive & History
                old_len = len(self.archive)
                self.archive = update_archive(self.archive, sol)
                if len(self.archive) > old_len or (sol in self.archive):
                    sol_entry = sol.copy();
                    sol_entry["found_at_iter"] = t + 1
                    self.archive_history.append(sol_entry)

                if sol["feasible"] and sol["Cost"] < self.P_best_score[i]:
                    self.P_best_score[i] = sol["Cost"];
                    self.P_best_pos[i] = self.X[i].copy()

                self.detailed_log.append({
                    "Iteration": t + 1, "Agent_ID": i + 1, "Cost": sol["Cost"], "VDI": sol["VDI"],
                    "Overshed": sol["Over"], "MW_Shed": sol["MW"], "Min_Voltage": sol["min_v"],
                    "Max_Voltage": sol["max_v"], "Feasible_Total": sol["feasible"],
                    "Pattern": "".join(map(str, sol["pattern"]))
                })

            if self.archive:
                knee = find_knee(self.archive)
                self.G_best_pos = knee["pattern"]
                self.G_best_metrics = {'MW': knee["MW"], 'Cost': knee["Cost"], 'VDI': knee["VDI"],
                                       'Vmin': knee["min_v"]}

            # --- UPDATE MECHANISM ---
            if self.algo_name == "NSGA-II":
                new_X = []
                for _ in range(SEARCH_AGENTS // 2):
                    p1, p2 = random.sample(current_pop_data, 2)
                    cp = random.randint(1, self.dim - 1)
                    c1 = np.concatenate([p1["pattern"][:cp], p2["pattern"][cp:]])
                    c2 = np.concatenate([p2["pattern"][:cp], p1["pattern"][cp:]])
                    # Mutasi
                    for child in [c1, c2]:
                        if random.random() < 0.1:
                            idx = random.randint(0, self.dim - 1)
                            child[idx] = 1 - child[idx]
                        new_X.append(child)
                self.X = np.array(new_X)
            else:
                for i in range(SEARCH_AGENTS):
                    r1, r2 = np.random.random(), np.random.random()
                    if self.algo_name == "PSO":
                        self.V[i] = (w * self.V[i]) + (C1_PSO * r1 * (self.P_best_pos[i] - self.X[i])) + (
                                    C2_PSO * r2 * (self.G_best_pos - self.X[i]))
                        self.X[i] += np.clip(self.V[i], -5, 5)
                    elif self.algo_name == "GWO":
                        A, C = 2.0 * a * np.random.random(self.dim) - a, 2.0 * np.random.random(self.dim)
                        D = np.abs(C * self.G_best_pos - self.X[i]);
                        self.X[i] = self.G_best_pos - A * D
                    elif self.algo_name == "WOA":
                        p, l = np.random.random(), np.random.uniform(-1, 1)
                        if p < 0.5:
                            A = 2.0 * a * np.random.random() - a
                            D = np.abs(2.0 * np.random.random() * self.G_best_pos - self.X[i]);
                            self.X[i] = self.G_best_pos - A * D
                        else:
                            D_p = np.abs(self.G_best_pos - self.X[i]);
                            self.X[i] = D_p * np.exp(1 * l) * np.cos(2 * np.pi * l) + self.G_best_pos

            print(
                f"[{self.algo_name}] Iter {t + 1} | Archive: {len(self.archive)} | Knee Cost: {self.G_best_metrics['Cost']:.2f}")
        return self.G_best_metrics, self.detailed_log, self.archive, self.archive_history


# --- MAIN EXECUTION ---
if __name__ == "__main__":
    # Contingency Trip Generator
    for l in all_loads: l.SetAttribute("outserv", 0)
    target = next((g for g in all_gens if g.GetAttribute("loc_name") == TARGET_GEN_TRIP), None)
    if target: target.SetAttribute("outserv", 1)

    algos = ["PSO", "WOA", "GWO", "NSGA-II"]
    results_store = {}

    for algo in algos:
        opt = Optimizer(algo, TARGET_DEFICIT_MW)
        metrics, log, arch, history = opt.run()
        results_store[algo] = {"Metrics": metrics, "Log": log, "Final_Pareto": arch, "History": history}

    print(f"Exporting data ke {EXCEL_OUTPUT}...")
    with pd.ExcelWriter(EXCEL_OUTPUT, engine='openpyxl') as writer:
        # Summary Sheet
        summary = [{"Algorithm": a, "Knee_Cost": results_store[a]["Metrics"]["Cost"],
                    "Knee_Vmin": results_store[a]["Metrics"]["Vmin"]} for a in algos]
        pd.DataFrame(summary).to_excel(writer, sheet_name='Summary', index=False)

        for algo in algos:
            # Log Sheet
            pd.DataFrame(results_store[algo]["Log"]).to_excel(writer, sheet_name=f"Log_{algo}", index=False)
            # Pareto Sheet (Riwayat optimalitas)
            hist_data = []
            for sol in results_store[algo]["History"]:
                hist_data.append({
                    "Iter": sol["found_at_iter"], "Cost": sol["Cost"], "VDI": sol["VDI"],
                    "Over": sol["Over"], "Min_V": sol["min_v"], "Max_V": sol["max_v"],
                    "Pattern": "".join(map(str, sol["pattern"]))
                })
            pd.DataFrame(hist_data).to_excel(writer, sheet_name=f"Pareto_{algo}", index=False)

    print("Proses Selesai. Seluruh data 4 algoritma berhasil disimpan.")