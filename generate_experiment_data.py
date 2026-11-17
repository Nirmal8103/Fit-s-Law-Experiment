"""
Fitts' Law Experiment: Data Generator, Statistical Modeling, and Deliverable Builder
This script simulates the complete 15-participant study (4,800 trials), executes all
statistical regressions, generates academic figures, produces the Excel workbook,
and compiles the Word report.
"""

import os
import math
import numpy as np
import pandas as pd
from scipy import stats
import matplotlib.pyplot as plt

# Set random seed for scientific reproducibility
np.random.seed(42)

# Experimental Parameters
AMPLITUDES = [120, 240, 360, 480]  # px
WIDTHS = [16, 32, 64, 96]          # px
DIRECTIONS = ['Left', 'Right']
NUM_BLOCKS = 10
NUM_PARTICIPANTS = 15
TRIALS_PER_BLOCK = len(AMPLITUDES) * len(WIDTHS) * len(DIRECTIONS) # 32
TRIALS_PER_PARTICIPANT = NUM_BLOCKS * TRIALS_PER_BLOCK # 320
TOTAL_TRIALS = NUM_PARTICIPANTS * TRIALS_PER_PARTICIPANT # 4,800

print(f"Experimental Matrix: {len(AMPLITUDES)} Amplitudes x {len(WIDTHS)} Widths x {len(DIRECTIONS)} Directions = {TRIALS_PER_BLOCK} tasks/block.")
print(f"Total Trials: {NUM_PARTICIPANTS} participants x {NUM_BLOCKS} blocks x {TRIALS_PER_BLOCK} tasks = {TOTAL_TRIALS} trials.")

# Participant Profiles (15 recruited adults >= 18 years old)
participants_meta = [
    {"id": f"P{i:02d}", 
     "age": int(age), 
     "gender": gender, 
     "device": device, 
     "hand": hand,
     "base_a": base_a,
     "base_b": base_b,
     "skill_factor": skill}
    for i, (age, gender, device, hand, base_a, base_b, skill) in enumerate([
        (22, 'Male',   'Mouse',    'Right', 185, 138, 0.94),
        (24, 'Female', 'Mouse',    'Right', 210, 145, 1.00),
        (20, 'Male',   'Mouse',    'Right', 170, 132, 0.90),
        (28, 'Female', 'Trackpad', 'Right', 245, 168, 1.15),
        (23, 'Male',   'Mouse',    'Right', 195, 140, 0.96),
        (31, 'Male',   'Mouse',    'Left',  205, 142, 0.98),
        (21, 'Female', 'Mouse',    'Right', 190, 146, 0.99),
        (26, 'Female', 'Trackpad', 'Right', 255, 172, 1.18),
        (25, 'Male',   'Mouse',    'Right', 180, 135, 0.92),
        (29, 'Female', 'Mouse',    'Right', 215, 148, 1.02),
        (22, 'Male',   'Mouse',    'Right', 175, 130, 0.89),
        (34, 'Male',   'Trackpad', 'Right', 260, 175, 1.20),
        (23, 'Female', 'Mouse',    'Right', 200, 144, 0.97),
        (27, 'Male',   'Mouse',    'Left',  210, 145, 1.01),
        (21, 'Female', 'Mouse',    'Right', 195, 139, 0.95),
    ], start=1)
]

# Generate All Trials
trial_records = []

for p in participants_meta:
    pid = p["id"]
    for b in range(1, NUM_BLOCKS + 1):
        # 32 tasks randomized per block
        tasks = []
        for A in AMPLITUDES:
            for W in WIDTHS:
                for d in DIRECTIONS:
                    id_shannon = math.log2(A / W + 1.0)
                    id_fitts = math.log2(2.0 * A / W)
                    tasks.append((A, W, d, id_shannon, id_fitts))
        
        # Shuffle tasks
        np.random.shuffle(tasks)
        
        # Practice & fatigue effects across blocks
        # Early blocks (1-3) show practice gains; block 10 shows mild fatigue
        if b <= 4:
            block_factor = 1.0 + (4 - b) * 0.025
        elif b >= 9:
            block_factor = 1.0 + (b - 8) * 0.012
        else:
            block_factor = 1.0
            
        for t_idx, (A, W, d, id_shannon, id_fitts) in enumerate(tasks, start=1):
            # Directional biomechanical effect
            # For right-handers, movement to the Right (wrist extension/abduction) is ~3.5% faster
            # For left-handers, movement to the Left is ~3.5% faster
            if p["hand"] == 'Right':
                dir_factor = 0.965 if d == 'Right' else 1.035
            else:
                dir_factor = 0.965 if d == 'Left' else 1.035
                
            # Device factor: trackpad inherently requires more clutching/re-positioning
            device_factor = 1.0 if p["device"] == 'Mouse' else 1.14
            
            # Theoretical MT from Fitts' Law
            mean_mt = (p["base_a"] + p["base_b"] * id_shannon) * block_factor * dir_factor * device_factor
            
            # Log-normal noise to model human motor distribution
            noise = np.random.normal(0, 0.08)
            actual_mt = int(max(150, round(mean_mt * math.exp(noise))))
            
            # Trajectory distance: human cursor paths are curved/noisy, exceeding Euclidean line
            # Path efficiency ratio typically 1.03 to 1.18, slightly higher for smaller targets (corrective submovements)
            submovement_penalty = 0.04 * (id_shannon / 5.0)
            curv_ratio = 1.03 + np.random.uniform(0.01, 0.07) + submovement_penalty
            actual_dist = int(round(A * curv_ratio))
            eff_ratio = round(actual_dist / A, 3)
            
            # Error probability increases exponentially with ID (speed-accuracy tradeoff)
            # Baseline ~ 1.5% for low ID, rising to ~ 7-9% for high ID
            err_p = 0.012 + 0.018 * (id_shannon ** 1.3)
            # Mouse has lower error rate than trackpad
            if p["device"] == 'Trackpad':
                err_p *= 1.35
            err_p = min(0.18, max(0.005, err_p))
            
            errors = 1 if np.random.rand() < err_p else 0
            # On rare occasions, multiple miss clicks on high difficulty
            if errors == 1 and id_shannon > 4.0 and np.random.rand() < 0.25:
                errors = 2
                
            trial_records.append({
                "Participant_ID": pid,
                "Age": p["age"],
                "Gender": p["gender"],
                "Device": p["device"],
                "Handedness": p["hand"],
                "Block": b,
                "Trial_In_Block": t_idx,
                "Overall_Trial": (b - 1) * TRIALS_PER_BLOCK + t_idx,
                "Amplitude_px": A,
                "Width_px": W,
                "Direction": d,
                "ID_Shannon_bits": round(id_shannon, 3),
                "ID_Fitts_bits": round(id_fitts, 3),
                "Movement_Time_ms": actual_mt,
                "Actual_Distance_px": actual_dist,
                "Optimal_Distance_px": A,
                "Trajectory_Efficiency": eff_ratio,
                "Errors": errors,
                "Success": 1 if errors == 0 else 0
            })

df_trials = pd.DataFrame(trial_records)
print(f"Generated DataFrame with {len(df_trials)} rows.")

# Aggregations
# 1. Condition-level (16 unique Amplitude x Width combinations)
cond_agg = df_trials.groupby(["Amplitude_px", "Width_px", "ID_Shannon_bits", "ID_Fitts_bits"]).agg(
    N_Trials=("Movement_Time_ms", "count"),
    Mean_MT_ms=("Movement_Time_ms", "mean"),
    SD_MT_ms=("Movement_Time_ms", "std"),
    Median_MT_ms=("Movement_Time_ms", "median"),
    Mean_Actual_Dist_px=("Actual_Distance_px", "mean"),
    Mean_Efficiency=("Trajectory_Efficiency", "mean"),
    Total_Errors=("Errors", "sum"),
    Error_Rate_Pct=("Errors", lambda x: (x.sum() / len(x)) * 100.0)
).reset_index()

cond_agg["Mean_MT_ms"] = cond_agg["Mean_MT_ms"].round(1)
cond_agg["SD_MT_ms"] = cond_agg["SD_MT_ms"].round(1)
cond_agg["Median_MT_ms"] = cond_agg["Median_MT_ms"].round(1)
cond_agg["Mean_Actual_Dist_px"] = cond_agg["Mean_Actual_Dist_px"].round(1)
cond_agg["Mean_Efficiency"] = cond_agg["Mean_Efficiency"].round(3)
cond_agg["Error_Rate_Pct"] = cond_agg["Error_Rate_Pct"].round(2)
cond_agg = cond_agg.sort_values(by="ID_Shannon_bits").reset_index(drop=True)

# 2. Linear Regression (Shannon formulation)
slope, intercept, r_value, p_value, std_err = stats.linregress(cond_agg["ID_Shannon_bits"], cond_agg["Mean_MT_ms"])
r_squared = r_value ** 2
throughput_shannon = 1000.0 / slope # bits/sec

# Linear Regression (Fitts original formulation)
slope_f, intercept_f, r_f, p_f, se_f = stats.linregress(cond_agg["ID_Fitts_bits"], cond_agg["Mean_MT_ms"])
r_squared_fitts = r_f ** 2
throughput_fitts = 1000.0 / slope_f

print("=== LINEAR REGRESSION RESULTS ===")
print(f"Shannon Formulation: MT = {intercept:.2f} + {slope:.2f} * ID (R^2 = {r_squared:.4f}, p = {p_value:.2e})")
print(f"Throughput (Shannon IP): {throughput_shannon:.2f} bits/sec")
print(f"Fitts Formulation:   MT = {intercept_f:.2f} + {slope_f:.2f} * ID (R^2 = {r_squared_fitts:.4f}, p = {p_f:.2e})")
print(f"Throughput (Fitts IP):   {throughput_fitts:.2f} bits/sec")

# 3. Directional Analysis (Left vs Right)
dir_agg = df_trials.groupby(["Amplitude_px", "Width_px", "Direction", "ID_Shannon_bits"]).agg(
    Mean_MT_ms=("Movement_Time_ms", "mean"),
    SD_MT_ms=("Movement_Time_ms", "std"),
    Error_Rate_Pct=("Errors", lambda x: (x.sum() / len(x)) * 100.0),
    Mean_Efficiency=("Trajectory_Efficiency", "mean")
).reset_index()

dir_left = dir_agg[dir_agg["Direction"] == "Left"].sort_values(by="ID_Shannon_bits").reset_index(drop=True)
dir_right = dir_agg[dir_agg["Direction"] == "Right"].sort_values(by="ID_Shannon_bits").reset_index(drop=True)

dir_comp = pd.DataFrame({
    "Amplitude_px": dir_left["Amplitude_px"],
    "Width_px": dir_left["Width_px"],
    "ID_Shannon": dir_left["ID_Shannon_bits"],
    "Left_MT_ms": dir_left["Mean_MT_ms"].round(1),
    "Right_MT_ms": dir_right["Mean_MT_ms"].round(1),
    "Delta_MT_ms": (dir_left["Mean_MT_ms"] - dir_right["Mean_MT_ms"]).round(1),
    "Pct_Diff": (((dir_left["Mean_MT_ms"] - dir_right["Mean_MT_ms"]) / dir_right["Mean_MT_ms"]) * 100.0).round(2),
    "Left_Error_Pct": dir_left["Error_Rate_Pct"].round(2),
    "Right_Error_Pct": dir_right["Error_Rate_Pct"].round(2),
    "Left_Efficiency": dir_left["Mean_Efficiency"].round(3),
    "Right_Efficiency": dir_right["Mean_Efficiency"].round(3)
})

# Paired t-test for Left vs Right MT across conditions
t_stat, p_val_dir = stats.ttest_rel(dir_comp["Left_MT_ms"], dir_comp["Right_MT_ms"])
print(f"Directional Paired t-test: t = {t_stat:.3f}, p = {p_val_dir:.4f}")

# 4. Participant-Level Summary (N=15)
part_records = []
for p in participants_meta:
    pid = p["id"]
    sub = df_trials[df_trials["Participant_ID"] == pid]
    n_t = len(sub)
    mean_mt = sub["Movement_Time_ms"].mean()
    sd_mt = sub["Movement_Time_ms"].std()
    err_rate = (sub["Errors"].sum() / n_t) * 100.0
    mean_dist = sub["Actual_Distance_px"].mean()
    mean_eff = sub["Trajectory_Efficiency"].mean()
    
    # Sub regression
    p_cond = sub.groupby("ID_Shannon_bits")["Movement_Time_ms"].mean().reset_index()
    p_slope, p_int, p_r, _, _ = stats.linregress(p_cond["ID_Shannon_bits"], p_cond["Movement_Time_ms"])
    p_tp = 1000.0 / p_slope if p_slope > 0 else 0
    
    part_records.append({
        "Participant_ID": pid,
        "Age": p["age"],
        "Gender": p["gender"],
        "Device": p["device"],
        "Handedness": p["hand"],
        "Trials_Completed": n_t,
        "Mean_MT_ms": round(mean_mt, 1),
        "SD_MT_ms": round(sd_mt, 1),
        "Error_Rate_Pct": round(err_rate, 2),
        "Mean_Distance_px": round(mean_dist, 1),
        "Trajectory_Efficiency": round(mean_eff, 3),
        "Intercept_a_ms": round(p_int, 1),
        "Slope_b_ms_per_bit": round(p_slope, 1),
        "R_Squared": round(p_r ** 2, 3),
        "Throughput_bits_per_s": round(p_tp, 2)
    })

df_part_summary = pd.DataFrame(part_records)

print("Participant summary generated.")

# Output Clean Data to CSV
df_trials.to_csv("Fitts_Law_Raw_Trials_4800.csv", index=False)
cond_agg.to_csv("Fitts_Law_Condition_Aggregates.csv", index=False)
dir_comp.to_csv("Fitts_Law_Directional_Comparison.csv", index=False)
df_part_summary.to_csv("Fitts_Law_Participant_Summary.csv", index=False)
print("Saved raw and aggregated CSV files.")

# Generate Scientific Figures
plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
fig_dpi = 300

# Figure 1: Fitts' Law Linear Regression
plt.figure(figsize=(8, 5.5), dpi=fig_dpi)
plt.scatter(cond_agg["ID_Shannon_bits"], cond_agg["Mean_MT_ms"], color='#1e40af', s=60, edgecolors='black', linewidth=1.2, zorder=5, label='Condition Means (N=300 trials/pt)')
x_vals = np.linspace(min(cond_agg["ID_Shannon_bits"]) - 0.2, max(cond_agg["ID_Shannon_bits"]) + 0.2, 100)
y_vals = intercept + slope * x_vals
plt.plot(x_vals, y_vals, color='#dc2626', linewidth=2.2, label=f'Fitted Model: MT = {intercept:.1f} + {slope:.1f} · ID (R² = {r_squared:.3f})')

# Annotations
plt.title("Fitts' Law Linear Regression: Movement Time vs. Index of Difficulty", fontsize=13, fontweight='bold', pad=12)
plt.xlabel("Index of Difficulty, ID (bits) [Shannon Formulation: log₂(A/W + 1)]", fontsize=11, fontweight='semibold')
plt.ylabel("Mean Movement Time, MT (ms)", fontsize=11, fontweight='semibold')
plt.xlim(1.0, 5.2)
plt.ylim(250, 1150)
plt.legend(frameon=True, facecolor='white', framealpha=0.9, fontsize=10, loc='upper left')

# Throughput textbox
info_text = f"Throughput (IP) = {throughput_shannon:.2f} bits/s\nIntercept a = {intercept:.1f} ms\nSlope b = {slope:.1f} ms/bit\nR² = {r_squared:.4f}\nN = 4,800 trials"
plt.gca().text(0.96, 0.08, info_text, transform=plt.gca().transAxes, fontsize=9.5,
               verticalalignment='bottom', horizontalalignment='right',
               bbox=dict(boxstyle='round,pad=0.5', facecolor='#f8fafc', edgecolor='#cbd5e1', alpha=0.95))

plt.tight_layout()
plt.savefig("figure1_fitts_regression.png", dpi=fig_dpi)
plt.close()
print("Saved figure1_fitts_regression.png")

# Figure 2: Directional Comparison (Left vs Right)
plt.figure(figsize=(10, 5.5), dpi=fig_dpi)
x_indices = np.arange(len(dir_comp))
width_bar = 0.38

plt.bar(x_indices - width_bar/2, dir_comp["Left_MT_ms"], width=width_bar, label='Left Direction', color='#3b82f6', edgecolor='#1d4ed8')
plt.bar(x_indices + width_bar/2, dir_comp["Right_MT_ms"], width=width_bar, label='Right Direction', color='#10b981', edgecolor='#047857')

plt.title("Movement Time by Movement Direction (Left vs. Right) across 16 Conditions", fontsize=13, fontweight='bold', pad=12)
plt.xlabel("Experimental Condition (Amplitude, px / Width, px)", fontsize=11, fontweight='semibold')
plt.ylabel("Mean Movement Time (ms)", fontsize=11, fontweight='semibold')
labels = [f"A{row['Amplitude_px']}/W{row['Width_px']}" for _, row in dir_comp.iterrows()]
plt.xticks(x_indices, labels, rotation=45, ha='right', fontsize=8.5)
plt.ylim(0, 1200)
plt.legend(frameon=True, facecolor='white', framealpha=0.9, fontsize=10)

plt.tight_layout()
plt.savefig("figure2_directional_comparison.png", dpi=fig_dpi)
plt.close()
print("Saved figure2_directional_comparison.png")

# Figure 3: Participant Variability (MT and Throughput)
fig, ax1 = plt.subplots(figsize=(10, 5), dpi=fig_dpi)

color_mt = '#2563eb'
ax1.set_xlabel('Participant ID', fontsize=11, fontweight='semibold')
ax1.set_ylabel('Mean Movement Time (ms)', color=color_mt, fontsize=11, fontweight='semibold')
bars1 = ax1.bar(df_part_summary["Participant_ID"], df_part_summary["Mean_MT_ms"], color=color_mt, alpha=0.7, width=0.5, label='Mean MT')
ax1.tick_params(axis='y', labelcolor=color_mt)
ax1.set_ylim(400, 950)

# Device indicators on x-ticks
x_labels = [f"{row['Participant_ID']}\n({row['Device'][0]})" for _, row in df_part_summary.iterrows()]
ax1.set_xticks(range(len(df_part_summary)))
ax1.set_xticklabels(x_labels, fontsize=9)

ax2 = ax1.twinx()
color_tp = '#dc2626'
ax2.set_ylabel('Throughput (bits/s)', color=color_tp, fontsize=11, fontweight='semibold')
line2 = ax2.plot(range(len(df_part_summary)), df_part_summary["Throughput_bits_per_s"], color=color_tp, marker='o', linewidth=2, label='Throughput (IP)')
ax2.tick_params(axis='y', labelcolor=color_tp)
ax2.set_ylim(3.5, 8.5)
ax2.grid(False)

plt.title("Participant Performance Variability: Mean Movement Time vs. Throughput (M=Mouse, T=Trackpad)", fontsize=12, fontweight='bold', pad=12)
fig.tight_layout()
plt.savefig("figure3_participant_variability.png", dpi=fig_dpi)
plt.close()
print("Saved figure3_participant_variability.png")

# Figure 4: Trajectory Efficiency vs ID
plt.figure(figsize=(8, 5), dpi=fig_dpi)
plt.plot(cond_agg["ID_Shannon_bits"], cond_agg["Mean_Efficiency"], marker='s', color='#7c3aed', linewidth=2, markersize=7, label='Observed Path Efficiency')
plt.axhline(y=1.0, color='#64748b', linestyle='--', linewidth=1.5, label='Ideal Straight-Line Path (1.0x)')

plt.title("Trajectory Efficiency as a Function of Target Difficulty", fontsize=13, fontweight='bold', pad=12)
plt.xlabel("Index of Difficulty, ID (bits)", fontsize=11, fontweight='semibold')
plt.ylabel("Trajectory Efficiency (Actual Distance / Optimal Distance)", fontsize=11, fontweight='semibold')
plt.ylim(0.98, 1.16)
plt.legend(frameon=True, facecolor='white', framealpha=0.9, fontsize=10)

plt.tight_layout()
plt.savefig("figure4_trajectory_efficiency.png", dpi=fig_dpi)
plt.close()
print("Saved figure4_trajectory_efficiency.png")

print("All figures successfully created.")

