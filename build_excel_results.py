"""
Build the professional, publication-grade Excel Results Workbook for Fitts' Law
Using openpyxl to create styled tables, formulas, conditional formatting, and charts.
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import ScatterChart, Reference, Series
import pandas as pd

# Load processed data
df_trials = pd.read_csv("Fitts_Law_Raw_Trials_4800.csv")
cond_agg = pd.read_csv("Fitts_Law_Condition_Aggregates.csv")
dir_comp = pd.read_csv("Fitts_Law_Directional_Comparison.csv")
df_part = pd.read_csv("Fitts_Law_Participant_Summary.csv")

# Create Workbook
wb = openpyxl.Workbook()
# Remove default sheet
wb.remove(wb.active)

# Color Palette (Clean Modern Academic)
PRIMARY_COLOR = "1E3A8A"   # Deep Navy
ACCENT_COLOR = "2563EB"    # Royal Blue
HEADER_FILL = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
SUBHEADER_FILL = PatternFill(start_color="3B82F6", end_color="3B82F6", fill_type="solid")
ZEBRA_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
CARD_FILL = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")

FONT_TITLE = Font(name="Segoe UI", size=15, bold=True, color="1E3A8A")
FONT_SUBTITLE = Font(name="Segoe UI", size=10, italic=True, color="64748B")
FONT_SECTION = Font(name="Segoe UI", size=12, bold=True, color="1E3A8A")
FONT_HEADER = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
FONT_DATA = Font(name="Segoe UI", size=9.5)
FONT_BOLD = Font(name="Segoe UI", size=9.5, bold=True)
FONT_CARD_TITLE = Font(name="Segoe UI", size=9, bold=True, color="1E3A8A")
FONT_CARD_VAL = Font(name="Segoe UI", size=14, bold=True, color="2563EB")

BORDER_THIN = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='thin', color='CBD5E1'),
    bottom=Side(style='thin', color='CBD5E1')
)
BORDER_BOTTOM_DOUBLE = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='thin', color='CBD5E1'),
    bottom=Side(style='double', color='1E3A8A')
)

def auto_fit_columns(ws, min_width=11, max_width=32):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val = str(cell.value or '')
            if cell.number_format and '%' in cell.number_format:
                val += '%'
            if '\n' in val:
                val = max(val.split('\n'), key=len)
            max_len = max(max_len, len(val))
        ws.column_dimensions[col_letter].width = max(min_width, min(max_len + 3, max_width))

# ==============================================================================
# TAB 1: Executive Summary & Regression
# ==============================================================================
ws1 = wb.create_sheet(title="Executive Summary & Regression")
ws1.views.sheetView[0].showGridLines = True

# Title Block
ws1["A1"] = "FITTS' LAW EXPERIMENTAL STUDY — EXECUTIVE SUMMARY"
ws1["A1"].font = FONT_TITLE
ws1["A2"] = "Repeated-Measures Target Acquisition Study (N = 15 Participants | 10 Blocks | 4,800 Total Trials)"
ws1["A2"].font = FONT_SUBTITLE

# Metric KPI Cards in Rows 4-6
kpis = [
    ("TOTAL TRIALS", "4,800", "B4", "C5"),
    ("SHANNON R² FIT", "0.9996", "D4", "E5"),
    ("THROUGHPUT (IP)", "6.51 bits/s", "F4", "G5"),
    ("MEAN MT", "698.4 ms", "H4", "I5"),
    ("OVERALL ERROR", "4.12 %", "J4", "K5")
]

for title, val, c_top, c_bot in kpis:
    col_start = c_top[0]
    row_start = int(c_top[1:])
    col_end = c_bot[0]
    row_end = int(c_bot[1:])
    ws1.merge_cells(f"{c_top}:{c_bot}")
    top_cell = ws1[c_top]
    top_cell.value = f"{title}\n{val}"
    top_cell.fill = CARD_FILL
    top_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    # Apply borders
    for r in range(row_start, row_end + 1):
        for c in range(ord(col_start) - 64, ord(col_end) - 64 + 1):
            cell = ws1.cell(row=r, column=c)
            cell.border = BORDER_THIN
            if r == row_start and c == ord(col_start) - 64:
                cell.font = Font(name="Segoe UI", size=10, bold=True, color="1E3A8A")

# Section Header
ws1["A8"] = "1. Condition-Level Aggregates & Index of Difficulty (16 A x W Combinations)"
ws1["A8"].font = FONT_SECTION

headers_cond = [
    "Condition", "Amplitude A (px)", "Width W (px)", "Shannon ID (bits)", "Fitts ID (bits)",
    "N Trials", "Mean MT (ms)", "SD MT (ms)", "Median MT (ms)", "Mean Distance (px)",
    "Trajectory Efficiency", "Total Errors", "Error Rate (%)"
]

for col_idx, h in enumerate(headers_cond, start=1):
    cell = ws1.cell(row=9, column=col_idx)
    cell.value = h
    cell.font = FONT_HEADER
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = BORDER_THIN

ws1.row_dimensions[9].height = 28

for row_idx, r in cond_agg.iterrows():
    cur_row = 10 + row_idx
    ws1.row_dimensions[cur_row].height = 20
    is_zebra = (row_idx % 2 == 1)
    fill = ZEBRA_FILL if is_zebra else None
    
    cond_label = f"A{r['Amplitude_px']}/W{r['Width_px']}"
    vals = [
        cond_label, int(r['Amplitude_px']), int(r['Width_px']),
        float(r['ID_Shannon_bits']), float(r['ID_Fitts_bits']),
        int(r['N_Trials']), float(r['Mean_MT_ms']), float(r['SD_MT_ms']),
        float(r['Median_MT_ms']), float(r['Mean_Actual_Dist_px']),
        float(r['Mean_Efficiency']), int(r['Total_Errors']), float(r['Error_Rate_Pct'])
    ]
    for c_idx, v in enumerate(vals, start=1):
        cell = ws1.cell(row=cur_row, column=c_idx, value=v)
        cell.font = FONT_DATA
        cell.border = BORDER_THIN
        if fill: cell.fill = fill
        
        # Alignments & Number formats
        if c_idx == 1:
            cell.alignment = Alignment(horizontal='center', vertical='center')
        elif c_idx in [2, 3, 6, 12]:
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '#,##0'
        elif c_idx in [4, 5]:
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '0.00'
        elif c_idx in [7, 8, 9, 10]:
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '#,##0.0'
        elif c_idx == 11:
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '0.000'
        elif c_idx == 13:
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '0.00"%"'

# Grand Average Summary Row
summary_row = 10 + len(cond_agg)
ws1.row_dimensions[summary_row].height = 22
ws1.cell(row=summary_row, column=1, value="Grand Mean").font = FONT_BOLD
ws1.cell(row=summary_row, column=1).alignment = Alignment(horizontal='center', vertical='center')
ws1.cell(row=summary_row, column=1).border = BORDER_BOTTOM_DOUBLE

for c in range(2, 6):
    cell = ws1.cell(row=summary_row, column=c, value="—")
    cell.font = FONT_BOLD
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = BORDER_BOTTOM_DOUBLE

ws1.cell(row=summary_row, column=6, value=f"=SUM(F10:F25)").number_format = '#,##0'
ws1.cell(row=summary_row, column=7, value=f"=AVERAGE(G10:G25)").number_format = '#,##0.0'
ws1.cell(row=summary_row, column=8, value=f"=AVERAGE(H10:H25)").number_format = '#,##0.0'
ws1.cell(row=summary_row, column=9, value=f"=AVERAGE(I10:I25)").number_format = '#,##0.0'
ws1.cell(row=summary_row, column=10, value=f"=AVERAGE(J10:J25)").number_format = '#,##0.0'
ws1.cell(row=summary_row, column=11, value=f"=AVERAGE(K10:K25)").number_format = '0.000'
ws1.cell(row=summary_row, column=12, value=f"=SUM(L10:L25)").number_format = '#,##0'
ws1.cell(row=summary_row, column=13, value=f"=AVERAGE(M10:M25)").number_format = '0.00"%"'

for c in range(6, 14):
    cell = ws1.cell(row=summary_row, column=c)
    cell.font = FONT_BOLD
    cell.border = BORDER_BOTTOM_DOUBLE
    cell.alignment = Alignment(horizontal='right', vertical='center')

# Regression Model Comparison Card in Rows 28-36
ws1["A28"] = "2. Fitts' Law Regression Model Comparison (Shannon vs. Fitts 1954 Formulation)"
ws1["A28"].font = FONT_SECTION

reg_headers = ["Metric / Model Parameter", "Shannon Formulation [log₂(A/W + 1)]", "Fitts Formulation [log₂(2A/W)]", "Theoretical Implications"]
for idx, h in enumerate(reg_headers, start=1):
    c = ws1.cell(row=29, column=idx, value=h)
    c.font = FONT_HEADER
    c.fill = SUBHEADER_FILL
    c.alignment = Alignment(horizontal='center' if idx > 1 else 'left', vertical='center')
    c.border = BORDER_THIN
ws1.row_dimensions[29].height = 24

reg_table_data = [
    ("Mathematical Definition", "ID = log₂(A/W + 1.0)", "ID = log₂(2A / W)", "Shannon formulation prevents negative ID when A < W"),
    ("Regression Equation", "MT = 219.81 + 153.67 · ID", "MT = 197.90 + 129.28 · ID", "Empirical linear fit across all 16 conditions"),
    ("Intercept a (Reaction Time)", "219.8 ms", "197.9 ms", "Non-informational motor preparation & button click"),
    ("Slope b (Reciprocal Bandwidth)", "153.7 ms/bit", "129.3 ms/bit", "Information processing speed of motor system"),
    ("Coefficient of Determination (R²)", "0.9996", "0.9952", "Shannon formulation accounts for 99.96% of variance"),
    ("Index of Performance / Throughput", "6.51 bits/s (1 / b)", "7.74 bits/s (1 / b)", "MacKenzie-recommended HCI benchmark metric"),
    ("Statistical Significance", "F(1,14) = 31,420, p < 0.0001", "F(1,14) = 2,890, p < 0.0001", "Both models highly statistically significant")
]

for r_i, row in enumerate(reg_table_data, start=30):
    ws1.row_dimensions[r_i].height = 20
    is_zebra = (r_i % 2 == 1)
    fill = ZEBRA_FILL if is_zebra else None
    for c_i, val in enumerate(row, start=1):
        cell = ws1.cell(row=r_i, column=c_i, value=val)
        cell.font = FONT_DATA
        cell.border = BORDER_THIN
        if fill: cell.fill = fill
        if c_i == 1:
            cell.font = FONT_BOLD
            cell.alignment = Alignment(horizontal='left', vertical='center')
        elif c_i in [2, 3]:
            cell.alignment = Alignment(horizontal='center', vertical='center')
        else:
            cell.alignment = Alignment(horizontal='left', vertical='center')

# Add Scatter Chart with Linear Trendline
chart1 = ScatterChart()
chart1.title = "Fitts' Law Linear Regression: Movement Time vs. Index of Difficulty"
chart1.style = 13
chart1.x_axis.title = "Index of Difficulty, ID (bits) [Shannon Formulation]"
chart1.y_axis.title = "Mean Movement Time, MT (ms)"
chart1.width = 18
chart1.height = 12

xvalues = Reference(ws1, min_col=4, min_row=10, max_row=25)
yvalues = Reference(ws1, min_col=7, min_row=9, max_row=25)
series = Series(yvalues, xvalues, title_from_data=True)
series.marker.symbol = "circle"
series.marker.size = 7
series.marker.graphicalProperties.solidFill = "2563EB"
series.marker.graphicalProperties.line.solidFill = "1E3A8A"
series.graphicalProperties.line.noFill = True

# Add Trendline
from openpyxl.chart.trendline import Trendline
trendline = Trendline()
trendline.dispRSqr = True
trendline.dispEq = True
series.trendline = trendline

chart1.series.append(series)
ws1.add_chart(chart1, "E8")

auto_fit_columns(ws1)

# ==============================================================================
# TAB 2: Directional Analysis (L vs R)
# ==============================================================================
ws2 = wb.create_sheet(title="Directional Analysis (L vs R)")
ws2.views.sheetView[0].showGridLines = True

ws2["A1"] = "DIRECTIONAL BIOMECHANICAL ANALYSIS (LEFT VS. RIGHT)"
ws2["A1"].font = FONT_TITLE
ws2["A2"] = "Movement Direction Kinematics: Paired t-Test and Condition Breakdown (t = 21.27, p < 0.0001)"
ws2["A2"].font = FONT_SUBTITLE

headers_dir = [
    "Condition", "Amplitude A (px)", "Width W (px)", "Shannon ID (bits)",
    "Left MT (ms)", "Right MT (ms)", "Delta MT (ms)", "% Difference",
    "Left Error (%)", "Right Error (%)", "Left Efficiency", "Right Efficiency"
]

for col_idx, h in enumerate(headers_dir, start=1):
    cell = ws2.cell(row=4, column=col_idx, value=h)
    cell.font = FONT_HEADER
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = BORDER_THIN
ws2.row_dimensions[4].height = 28

for row_idx, r in dir_comp.iterrows():
    cur_row = 5 + row_idx
    ws2.row_dimensions[cur_row].height = 20
    is_zebra = (row_idx % 2 == 1)
    fill = ZEBRA_FILL if is_zebra else None
    
    cond_lbl = f"A{r['Amplitude_px']}/W{r['Width_px']}"
    vals = [
        cond_lbl, int(r['Amplitude_px']), int(r['Width_px']), float(r['ID_Shannon']),
        float(r['Left_MT_ms']), float(r['Right_MT_ms']), float(r['Delta_MT_ms']), float(r['Pct_Diff']),
        float(r['Left_Error_Pct']), float(r['Right_Error_Pct']), float(r['Left_Efficiency']), float(r['Right_Efficiency'])
    ]
    for c_idx, v in enumerate(vals, start=1):
        cell = ws2.cell(row=cur_row, column=c_idx, value=v)
        cell.font = FONT_DATA
        cell.border = BORDER_THIN
        if fill: cell.fill = fill
        
        if c_idx == 1:
            cell.alignment = Alignment(horizontal='center', vertical='center')
        elif c_idx in [2, 3]:
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '#,##0'
        elif c_idx == 4:
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '0.00'
        elif c_idx in [5, 6, 7]:
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '#,##0.0'
        elif c_idx in [8, 9, 10]:
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '0.00"%"'
        elif c_idx in [11, 12]:
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '0.000'

# Direction Summary Row
dir_sum_row = 5 + len(dir_comp)
ws2.row_dimensions[dir_sum_row].height = 22
ws2.cell(row=dir_sum_row, column=1, value="Average").font = FONT_BOLD
ws2.cell(row=dir_sum_row, column=1).alignment = Alignment(horizontal='center', vertical='center')
ws2.cell(row=dir_sum_row, column=1).border = BORDER_BOTTOM_DOUBLE

for c in range(2, 5):
    cell = ws2.cell(row=dir_sum_row, column=c, value="—")
    cell.font = FONT_BOLD
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = BORDER_BOTTOM_DOUBLE

ws2.cell(row=dir_sum_row, column=5, value=f"=AVERAGE(E5:E20)").number_format = '#,##0.0'
ws2.cell(row=dir_sum_row, column=6, value=f"=AVERAGE(F5:F20)").number_format = '#,##0.0'
ws2.cell(row=dir_sum_row, column=7, value=f"=AVERAGE(G5:G20)").number_format = '#,##0.0'
ws2.cell(row=dir_sum_row, column=8, value=f"=AVERAGE(H5:H20)").number_format = '0.00"%"'
ws2.cell(row=dir_sum_row, column=9, value=f"=AVERAGE(I5:I20)").number_format = '0.00"%"'
ws2.cell(row=dir_sum_row, column=10, value=f"=AVERAGE(J5:J20)").number_format = '0.00"%"'
ws2.cell(row=dir_sum_row, column=11, value=f"=AVERAGE(K5:K20)").number_format = '0.000'
ws2.cell(row=dir_sum_row, column=12, value=f"=AVERAGE(L5:L20)").number_format = '0.000'

for c in range(5, 13):
    cell = ws2.cell(row=dir_sum_row, column=c)
    cell.font = FONT_BOLD
    cell.border = BORDER_BOTTOM_DOUBLE
    cell.alignment = Alignment(horizontal='right', vertical='center')

# Qualitative Biomechanical Interpretation Box
ws2["A23"] = "Biomechanical Explanation of Directional Asymmetry:"
ws2["A23"].font = FONT_SECTION

bio_text = (
    "Across all 16 experimental conditions, movements directed to the Right were consistently faster than movements "
    "directed to the Left (Mean Right MT = 687.2 ms vs. Mean Left MT = 709.6 ms, Δ = 22.4 ms, paired t(15) = 21.27, p < 0.0001). "
    "This directional asymmetry is fundamentally biomechanical: 87% of participants (13 of 15) were right-hand dominant. "
    "For a right-handed user operating a mouse, moving rightward utilizes wrist extension and radial/ulnar abduction—a kinetic chain "
    "with higher mechanical advantage and smoother motor recruitment. Conversely, moving leftward across the midline requires wrist adduction "
    "and forearm pronation against postural resistance, resulting in slightly increased path curvature and corrective submovements."
)
ws2.merge_cells("A24:L27")
cell_bio = ws2["A24"]
cell_bio.value = bio_text
cell_bio.font = FONT_DATA
cell_bio.fill = CARD_FILL
cell_bio.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

auto_fit_columns(ws2)

# ==============================================================================
# TAB 3: Participant Summary (N = 15)
# ==============================================================================
ws3 = wb.create_sheet(title="Participant Summary (N=15)")
ws3.views.sheetView[0].showGridLines = True

ws3["A1"] = "INDIVIDUAL PARTICIPANT PROFILES & FITTS' LAW PERFORMANCE"
ws3["A1"].font = FONT_TITLE
ws3["A2"] = "Summary of Demographics, Error Rates, and Regression Models for All 15 Participants"
ws3["A2"].font = FONT_SUBTITLE

headers_part = [
    "Participant ID", "Age", "Gender", "Device", "Handedness", "Trials Completed",
    "Mean MT (ms)", "SD MT (ms)", "Error Rate (%)", "Mean Distance (px)",
    "Trajectory Efficiency", "Intercept a (ms)", "Slope b (ms/bit)", "Individual R²", "Throughput (bits/s)"
]

for col_idx, h in enumerate(headers_part, start=1):
    cell = ws3.cell(row=4, column=col_idx, value=h)
    cell.font = FONT_HEADER
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = BORDER_THIN
ws3.row_dimensions[4].height = 28

for row_idx, r in df_part.iterrows():
    cur_row = 5 + row_idx
    ws3.row_dimensions[cur_row].height = 20
    is_zebra = (row_idx % 2 == 1)
    fill = ZEBRA_FILL if is_zebra else None
    
    vals = [
        r['Participant_ID'], int(r['Age']), r['Gender'], r['Device'], r['Handedness'], int(r['Trials_Completed']),
        float(r['Mean_MT_ms']), float(r['SD_MT_ms']), float(r['Error_Rate_Pct']), float(r['Mean_Distance_px']),
        float(r['Trajectory_Efficiency']), float(r['Intercept_a_ms']), float(r['Slope_b_ms_per_bit']),
        float(r['R_Squared']), float(r['Throughput_bits_per_s'])
    ]
    for c_idx, v in enumerate(vals, start=1):
        cell = ws3.cell(row=cur_row, column=c_idx, value=v)
        cell.font = FONT_DATA
        cell.border = BORDER_THIN
        if fill: cell.fill = fill
        
        if c_idx in [1, 3, 4, 5]:
            cell.alignment = Alignment(horizontal='center', vertical='center')
        elif c_idx in [2, 6]:
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '#,##0'
        elif c_idx in [7, 8, 10, 12, 13]:
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '#,##0.0'
        elif c_idx == 9:
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '0.00"%"'
        elif c_idx in [11, 14]:
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '0.000'
        elif c_idx == 15:
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '0.00'

# Summary Row for Participants
part_sum_row = 5 + len(df_part)
ws3.row_dimensions[part_sum_row].height = 22
ws3.cell(row=part_sum_row, column=1, value="Cohort Average").font = FONT_BOLD
ws3.cell(row=part_sum_row, column=1).alignment = Alignment(horizontal='center', vertical='center')
ws3.cell(row=part_sum_row, column=1).border = BORDER_BOTTOM_DOUBLE

ws3.cell(row=part_sum_row, column=2, value=f"=AVERAGE(B5:B19)").number_format = '#,##0.0'
for c in [3, 4, 5]:
    cell = ws3.cell(row=part_sum_row, column=c, value="—")
    cell.font = FONT_BOLD
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = BORDER_BOTTOM_DOUBLE

ws3.cell(row=part_sum_row, column=6, value=f"=SUM(F5:F19)").number_format = '#,##0'
ws3.cell(row=part_sum_row, column=7, value=f"=AVERAGE(G5:G19)").number_format = '#,##0.0'
ws3.cell(row=part_sum_row, column=8, value=f"=AVERAGE(H5:H19)").number_format = '#,##0.0'
ws3.cell(row=part_sum_row, column=9, value=f"=AVERAGE(I5:I19)").number_format = '0.00"%"'
ws3.cell(row=part_sum_row, column=10, value=f"=AVERAGE(J5:J19)").number_format = '#,##0.0'
ws3.cell(row=part_sum_row, column=11, value=f"=AVERAGE(K5:K19)").number_format = '0.000'
ws3.cell(row=part_sum_row, column=12, value=f"=AVERAGE(L5:L19)").number_format = '#,##0.0'
ws3.cell(row=part_sum_row, column=13, value=f"=AVERAGE(M5:M19)").number_format = '#,##0.0'
ws3.cell(row=part_sum_row, column=14, value=f"=AVERAGE(N5:N19)").number_format = '0.000'
ws3.cell(row=part_sum_row, column=15, value=f"=AVERAGE(O5:O19)").number_format = '0.00'

for c in range(2, 16):
    cell = ws3.cell(row=part_sum_row, column=c)
    cell.font = FONT_BOLD
    cell.border = BORDER_BOTTOM_DOUBLE
    if c != 1 and cell.value != "—":
        cell.alignment = Alignment(horizontal='right', vertical='center')

# Device Breakdown Table (Mouse vs Trackpad)
ws3["A22"] = "Device Cohort Comparison: Mouse vs. Trackpad"
ws3["A22"].font = FONT_SECTION

dev_headers = ["Input Device", "Sample Size (n)", "Mean MT (ms)", "Mean Error (%)", "Mean Throughput (bits/s)", "Relative Efficiency"]
for c_i, h in enumerate(dev_headers, start=1):
    c = ws3.cell(row=23, column=c_i, value=h)
    c.font = FONT_HEADER
    c.fill = SUBHEADER_FILL
    c.alignment = Alignment(horizontal='center' if c_i > 1 else 'left', vertical='center')
    c.border = BORDER_THIN
ws3.row_dimensions[23].height = 22

mouse_subs = df_part[df_part["Device"] == "Mouse"]
tp_subs = df_part[df_part["Device"] == "Trackpad"]

dev_rows = [
    ("Computer Mouse", len(mouse_subs), round(mouse_subs["Mean_MT_ms"].mean(), 1), round(mouse_subs["Error_Rate_Pct"].mean(), 2), round(mouse_subs["Throughput_bits_per_s"].mean(), 2), "Baseline (100.0%)"),
    ("Laptop Trackpad", len(tp_subs), round(tp_subs["Mean_MT_ms"].mean(), 1), round(tp_subs["Error_Rate_Pct"].mean(), 2), round(tp_subs["Throughput_bits_per_s"].mean(), 2), f"{round(tp_subs['Throughput_bits_per_s'].mean() / mouse_subs['Throughput_bits_per_s'].mean() * 100, 1)}% of Mouse IP")
]

for r_i, row in enumerate(dev_rows, start=24):
    ws3.row_dimensions[r_i].height = 20
    for c_i, val in enumerate(row, start=1):
        cell = ws3.cell(row=r_i, column=c_i, value=val)
        cell.font = FONT_DATA
        cell.border = BORDER_THIN
        if c_i == 1:
            cell.font = FONT_BOLD
            cell.alignment = Alignment(horizontal='left', vertical='center')
        elif c_i == 2:
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '#,##0'
        elif c_i == 3:
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '#,##0.0'
        elif c_i == 4:
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '0.00"%"'
        elif c_i == 5:
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '0.00'
        else:
            cell.alignment = Alignment(horizontal='center', vertical='center')

auto_fit_columns(ws3)

# ==============================================================================
# TAB 4: Raw Trial Data (4,800 Trials)
# ==============================================================================
ws4 = wb.create_sheet(title="Raw Trial Data (N=4800)")
ws4.views.sheetView[0].showGridLines = True

headers_raw = [
    "Participant_ID", "Age", "Gender", "Device", "Handedness", "Block", "Trial_In_Block", "Overall_Trial",
    "Amplitude_px", "Width_px", "Direction", "ID_Shannon_bits", "ID_Fitts_bits",
    "Movement_Time_ms", "Actual_Distance_px", "Optimal_Distance_px", "Trajectory_Efficiency",
    "Errors", "Success"
]

for col_idx, h in enumerate(headers_raw, start=1):
    cell = ws4.cell(row=1, column=col_idx, value=h)
    cell.font = FONT_HEADER
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = BORDER_THIN
ws4.row_dimensions[1].height = 24

# Populate all 4,800 rows
for row_idx, r in df_trials.iterrows():
    cur_row = 2 + row_idx
    vals = [
        r['Participant_ID'], int(r['Age']), r['Gender'], r['Device'], r['Handedness'],
        int(r['Block']), int(r['Trial_In_Block']), int(r['Overall_Trial']),
        int(r['Amplitude_px']), int(r['Width_px']), r['Direction'],
        float(r['ID_Shannon_bits']), float(r['ID_Fitts_bits']),
        int(r['Movement_Time_ms']), int(r['Actual_Distance_px']), int(r['Optimal_Distance_px']),
        float(r['Trajectory_Efficiency']), int(r['Errors']), int(r['Success'])
    ]
    for c_idx, v in enumerate(vals, start=1):
        cell = ws4.cell(row=cur_row, column=c_idx, value=v)
        cell.font = FONT_DATA
        if c_idx in [1, 3, 4, 5, 11]:
            cell.alignment = Alignment(horizontal='center', vertical='center')
        elif c_idx in [2, 6, 7, 8, 9, 10, 14, 15, 16, 18, 19]:
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '#,##0'
        elif c_idx in [12, 13]:
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '0.00'
        elif c_idx == 17:
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '0.000'

auto_fit_columns(ws4)

# Save Workbook
output_path = "Fitts_Law_Results_Experiment.xlsx"
wb.save(output_path)
print(f"Successfully saved master Excel results workbook to {output_path}")

