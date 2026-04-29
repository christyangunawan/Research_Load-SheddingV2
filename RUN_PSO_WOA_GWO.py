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
    print("Error: Module powerfactory tidak ditemukan. Cek PATH API Anda.")
    exit()

app = powerfactory.GetApplication()
app.ActivateProject(PROJECT_NAME)
ldf = app.GetFromStudyCase("ComLdf")


# --- FUNGSI PENDUKUNG ---
def natural_keys(text):
    return [int(c) if c.isdigit() else c for c in re.split(r'(\d+)', text)]


def sigmoid(x):
    return 1.0 / (1.0 + np.exp(-x))


def get_voltage_array():
    return [b.GetAttribute("m:u") for b in all_buses]


# --- METRIK MOO ---
def normalize_pareto(pareto):
    objs = np.array([s["objectives"] for s in pareto])
    min_vals = objs.min(axis=0)
    max_vals = objs.max(axis=0)
    return (objs - min_vals) / (max_vals - min_vals + 1e-9)


def compute_hypervolume(pareto):
    if not pareto: return 0
    norm = normalize_pareto(pareto)
    ref = np.array([1.1, 1.1, 1.1])
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


# --- INISIALISASI OBJEK POWERFACTORY ---
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
    print("Resetting System & Applying Contingency...")
    for l in all_loads: l.SetAttribute("outserv", 0)
    for g in all_gens: g.SetAttribute("outserv", 0)

    target = next(
        (g for g in all_gens if g.GetAttribute("loc_name") in [TARGET_GEN_TRIP, TARGET_GEN_TRIP.replace("_", " ")]),
        None)
    if target:
        target.SetAttribute("outserv", 1)
        print(f"Generator {target.GetAttribute('loc_name')} TRIPPED.")
    return TARGET_DEFICIT_MW


def calculate_fitness(position_continuous, min_shed_required):
    probs = sigmoid(position_continuous)
    pattern = (np.random.rand(len(position_continuous)) < probs).astype(int)

    for i, load in enumerate(all_loads):
        load.SetAttribute("outserv", int(pattern[i]))

    err = ldf.Execute()
    voltages = get_voltage_array()

    total_mw = sum(all_loads[i].GetAttribute("plini") for i, val in enumerate(pattern) if val == 1)
    cost = sum(all_loads[i].GetAttribute("plini") * COST_MAP[i] for i, val in enumerate(pattern) if val == 1)
    vdi = sum((1.0 - v) ** 2 for v in voltages) if err == 0 else 100.0
    over = max(0.0, total_mw - min_shed_required)

    feasible = (err == 0) and (total_mw >= (min_shed_required - 0.001)) and (min(voltages) >= VOLT_MIN)

    return {
        "objectives": [cost, vdi, over],
        "MW": total_mw, "Cost": cost, "VDI": vdi, "Over": over,
        "feasible": feasible, "voltages": voltages, "pattern": pattern, "min_v": min(voltages) if voltages else 0
    }


def update_archive(archive, candidate):
    if not candidate["feasible"]: return archive
    # Dominance Check
    new_archive = [sol for sol in archive if not (
                all(candidate["objectives"][k] <= sol["objectives"][k] for k in range(3)) and any(
            candidate["objectives"][k] < sol["objectives"][k] for k in range(3)))]
    if not any(all(sol["objectives"][k] <= candidate["objectives"][k] for k in range(3)) and any(
            sol["objectives"][k] < candidate["objectives"][k] for k in range(3)) for sol in archive):
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
        self.G_best_pos = np.zeros(self.dim)
        self.G_best_metrics = {'Cost': 0}

    def run(self):
        start_time = time.time()
        for t in range(MAX_ITER):
            a = 2.0 - t * (2.0 / MAX_ITER)
            w = W_MAX - t * ((W_MAX - W_MIN) / MAX_ITER)

            for i in range(SEARCH_AGENTS):
                sol = calculate_fitness(self.X[i], self.min_shed_mw)
                self.archive = update_archive(self.archive, sol)

                # Update P_best
                if sol["Cost"] < self.P_best_score[i] and sol["feasible"]:
                    self.P_best_score[i] = sol["Cost"]
                    self.P_best_pos[i] = self.X[i].copy()

                self.detailed_log.append({
                    "Iteration": t + 1, "Agent_ID": i + 1,
                    "Cost": sol["Cost"], "VDI": sol["VDI"], "Overshed": sol["Over"],
                    "Feasible": sol["feasible"], "Pattern": "".join(map(str, sol["pattern"]))
                })


            if self.archive:
                leader = random.choice(self.archive)
                self.G_best_pos = leader["pattern"]
                self.G_best_metrics = {
                    'MW': leader["MW"], 'Cost': leader["Cost"],
                    'VDI': leader["VDI"], 'Over': leader["Over"], 'Vmin': leader["min_v"]
                }

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
                    A, C = 2.0 * a * np.random.random(self.dim) - a, 2.0 * np.random.random(self.dim)
                    D = np.abs(C * self.G_best_pos - self.X[i])
                    self.X[i] = self.G_best_pos - A * D
                elif self.algo_name == "WOA":
                    p, l = np.random.random(), np.random.uniform(-1, 1)
                    if p < 0.5:
                        A = 2.0 * a * np.random.random() - a
                        D = np.abs(2.0 * np.random.random() * self.G_best_pos - self.X[i])
                        self.X[i] = self.G_best_pos - A * D
                    else:
                        D_p = np.abs(self.G_best_pos - self.X[i])
                        self.X[i] = D_p * np.exp(1 * l) * np.cos(2 * np.pi * l) + self.G_best_pos

        return self.G_best_metrics, self.convergence_curve, time.time() - start_time, self.detailed_log, self.G_best_pos, self.archive


if __name__ == "__main__":
    calc_target = apply_contingency_and_reset()
    results_store = {}
    algos = ["PSO", "WOA", "GWO"]

    for algo in algos:
        opt = Optimizer(algo, calc_target)
        metrics, curve, duration, log, best_pat, pareto = opt.run()
        results_store[algo] = {
            "Metrics": metrics, "Curve": curve, "Time": duration,
            "Log": log, "Best_Pattern": best_pat, "Pareto": pareto
        }

    print(f"Exporting to {EXCEL_OUTPUT}...")
    with pd.ExcelWriter(EXCEL_OUTPUT, engine='openpyxl') as writer:
        # 1. Summary Sheet
        summary_df = pd.DataFrame([{
            "Algorithm": a, "Knee_Cost": results_store[a]["Metrics"]["Cost"],
            "Knee_Vmin": results_store[a]["Metrics"]["Vmin"], "Time_Sec": results_store[a]["Time"]
        } for a in algos])
        summary_df.to_excel(writer, sheet_name='Summary', index=False)

        # 2. Pareto & MOO Metrics Loop
        for algo in algos:
            pareto = results_store[algo]["Pareto"]
            if not pareto: continue

            # Pareto List
            pd.DataFrame([{
                "Cost": s["objectives"][0], "VDI": s["objectives"][1],
                "Overshed": s["objectives"][2], "MW": s["MW"]
            } for s in pareto]).to_excel(writer, sheet_name=f"Pareto_{algo}", index=False)

            # MOO Quality Metrics
            hv, sp, knee = compute_hypervolume(pareto), compute_spacing(pareto), find_knee(pareto)
            pd.DataFrame([{
                "Hypervolume": hv, "Spacing": sp,
                "Knee_Cost": knee["objectives"][0], "Knee_VDI": knee["objectives"][1]
            }]).to_excel(writer, sheet_name=f"MOO_{algo}", index=False)

            # Logs
            pd.DataFrame(results_store[algo]["Log"]).to_excel(writer, sheet_name=f"Log_{algo}", index=False)

    print("Proses Selesai Berhasil.")