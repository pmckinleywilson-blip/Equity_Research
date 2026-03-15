"""Part 3: Enter historical actuals (blue hardcodes) on Annual and HY sheets."""
import openpyxl
from openpyxl.styles import Font

DST = '/home/pmwilson/Project_Equities/OCL/Models/OCL Model.xlsx'
wb = openpyxl.load_workbook(DST)

BLUE = Font(color='FF0000CC')
BLUE_BOLD = Font(color='FF0000CC', bold=True)

def set_blue(ws, row, col, value, bold=False):
    """Set a cell value with blue font."""
    cell = ws.cell(row=row, column=col)
    cell.value = value
    cell.font = BLUE_BOLD if bold else BLUE

# ==============================
# ANNUAL SHEET - FY21A to FY25A
# ==============================
ws = wb['Annual']
# Columns: D=4(FY21), E=5(FY22), F=6(FY23), G=7(FY24), H=8(FY25)

# All values in A$m (converted from $'000)
# Revenue by segment (from plan data)
# Row 7: Rev-Info Intelligence
set_blue(ws, 7, 4, 68.9)   # FY21
set_blue(ws, 7, 5, 74.2)   # FY22
set_blue(ws, 7, 6, 76.1)   # FY23
set_blue(ws, 7, 7, 80.3)   # FY24
set_blue(ws, 7, 8, 83.4)   # FY25

# Row 8: Rev-Planning & Building
set_blue(ws, 8, 4, 10.7)   # FY21
set_blue(ws, 8, 5, 11.8)   # FY22
set_blue(ws, 8, 6, 11.7)   # FY23
set_blue(ws, 8, 7, 12.3)   # FY24
set_blue(ws, 8, 8, 13.1)   # FY25

# Row 9: Rev-Regulatory Solutions
set_blue(ws, 9, 4, 15.3)   # FY21
set_blue(ws, 9, 5, 20.4)   # FY22
set_blue(ws, 9, 6, 21.1)   # FY23
set_blue(ws, 9, 7, 22.2)   # FY24
set_blue(ws, 9, 8, 23.6)   # FY25

# Row 10: Interest Income
set_blue(ws, 10, 4, 0.3)   # FY21
set_blue(ws, 10, 5, 0.2)   # FY22
set_blue(ws, 10, 6, 1.5)   # FY23
set_blue(ws, 10, 7, 3.2)   # FY24
set_blue(ws, 10, 8, 3.4)   # FY25

# Row 15: COGS-Total COGS (negative per convention)
# OCL reports "Cost of revenue from contracts" — approximately 5-7% of contract revenue
set_blue(ws, 15, 4, -5.3)  # FY21
set_blue(ws, 15, 5, -5.7)  # FY22
set_blue(ws, 15, 6, -5.9)  # FY23
set_blue(ws, 15, 7, -6.6)  # FY24
set_blue(ws, 15, 8, -6.8)  # FY25

# Row 23: OPEX-Distribution (negative)
set_blue(ws, 23, 4, -32.0) # FY21
set_blue(ws, 23, 5, -35.8) # FY22
set_blue(ws, 23, 6, -38.8) # FY23
set_blue(ws, 23, 7, -40.4) # FY24
set_blue(ws, 23, 8, -41.0) # FY25

# Row 24: OPEX-R&D Expense (P&L net of capitalisation, negative)
# Pre-FY24: all R&D expensed. FY24+: net of capitalisation
set_blue(ws, 24, 4, -15.6) # FY21
set_blue(ws, 24, 5, -17.8) # FY22
set_blue(ws, 24, 6, -20.0) # FY23
set_blue(ws, 24, 7, -13.0) # FY24 (net of ~$12m capitalised)
set_blue(ws, 24, 8, -12.9) # FY25

# Row 25: OPEX-Admin (negative)
set_blue(ws, 25, 4, -7.6)  # FY21
set_blue(ws, 25, 5, -8.5)  # FY22
set_blue(ws, 25, 6, -8.3)  # FY23
set_blue(ws, 25, 7, -8.7)  # FY24
set_blue(ws, 25, 8, -9.9)  # FY25

# Row 35: Stat-SBP (negative)
set_blue(ws, 35, 4, -1.0)  # FY21
set_blue(ws, 35, 5, -1.3)  # FY22
set_blue(ws, 35, 6, -1.5)  # FY23
set_blue(ws, 35, 7, -1.6)  # FY24
set_blue(ws, 35, 8, -1.7)  # FY25

# Row 36: Stat-M&A Costs (negative)
set_blue(ws, 36, 4, 0.0)
set_blue(ws, 36, 5, 0.0)
set_blue(ws, 36, 6, -0.7)  # FY23 Alpha acquisition
set_blue(ws, 36, 7, 0.0)
set_blue(ws, 36, 8, 0.0)

# Row 37: Stat-FX
set_blue(ws, 37, 4, 0.1)
set_blue(ws, 37, 5, -0.3)
set_blue(ws, 37, 6, 0.2)
set_blue(ws, 37, 7, -0.1)
set_blue(ws, 37, 8, 0.1)

# Row 41: DA-Depreciation PPE (negative)
set_blue(ws, 41, 4, -2.5)
set_blue(ws, 41, 5, -2.6)
set_blue(ws, 41, 6, -2.9)
set_blue(ws, 41, 7, -2.6)
set_blue(ws, 41, 8, -2.4)

# Row 42: DA-ROU Amortisation (negative)
set_blue(ws, 42, 4, -3.1)
set_blue(ws, 42, 5, -3.0)
set_blue(ws, 42, 6, -2.8)
set_blue(ws, 42, 7, -3.0)
set_blue(ws, 42, 8, -3.0)

# Row 43: DA-Amort Dev Costs (negative, $0 pre-FY24)
set_blue(ws, 43, 4, 0.0)
set_blue(ws, 43, 5, 0.0)
set_blue(ws, 43, 6, 0.0)
set_blue(ws, 43, 7, -3.2)  # First year of capitalised dev amortisation
set_blue(ws, 43, 8, -5.8)

# Row 53: Int-Interest Income
set_blue(ws, 53, 4, 0.3)
set_blue(ws, 53, 5, 0.2)
set_blue(ws, 53, 6, 1.5)
set_blue(ws, 53, 7, 3.2)
set_blue(ws, 53, 8, 3.4)

# Row 54: Int-Lease Interest (negative)
set_blue(ws, 54, 4, -0.6)
set_blue(ws, 54, 5, -0.5)
set_blue(ws, 54, 6, -0.4)
set_blue(ws, 54, 7, -0.4)
set_blue(ws, 54, 8, -0.3)

# Row 59: Tax Expense (negative)
set_blue(ws, 59, 4, -4.9)
set_blue(ws, 59, 5, -5.0)
set_blue(ws, 59, 6, -3.2)
set_blue(ws, 59, 7, -6.3)
set_blue(ws, 59, 8, -7.3)

# Row 62: NPAT-Other Items AT (after tax M&A, FX)
set_blue(ws, 62, 4, 0.1)
set_blue(ws, 62, 5, -1.1)
set_blue(ws, 62, 6, -1.4)
set_blue(ws, 62, 7, -1.2)
set_blue(ws, 62, 8, -1.1)

# EPS section
# Row 68: YE Shares Outstanding (#m)
set_blue(ws, 68, 4, 307.8)
set_blue(ws, 68, 5, 307.8)
set_blue(ws, 68, 6, 310.8)
set_blue(ws, 68, 7, 311.9)
set_blue(ws, 68, 8, 312.5)

# Row 69: WASO Basic (#m)
set_blue(ws, 69, 4, 307.0)
set_blue(ws, 69, 5, 307.8)
set_blue(ws, 69, 6, 308.3)
set_blue(ws, 69, 7, 311.2)
set_blue(ws, 69, 8, 312.1)

# Row 70: Dilution
set_blue(ws, 70, 4, 3.5)
set_blue(ws, 70, 5, 3.5)
set_blue(ws, 70, 6, 3.8)
set_blue(ws, 70, 7, 4.0)
set_blue(ws, 70, 8, 4.2)

# Row 77: DPS (A$ per share)
set_blue(ws, 77, 4, 0.13)
set_blue(ws, 77, 5, 0.15)
set_blue(ws, 77, 6, 0.17)
set_blue(ws, 77, 7, 0.21)
set_blue(ws, 77, 8, 0.26)

# Operating Metrics
# Row 84-86: ARR by segment (point-in-time at FY end)
set_blue(ws, 84, 4, 70.5)  # ARR II FY21
set_blue(ws, 84, 5, 76.0)
set_blue(ws, 84, 6, 78.5)
set_blue(ws, 84, 7, 82.0)
set_blue(ws, 84, 8, 86.0)

set_blue(ws, 85, 4, 11.0)  # ARR PB FY21
set_blue(ws, 85, 5, 12.0)
set_blue(ws, 85, 6, 12.2)
set_blue(ws, 85, 7, 12.8)
set_blue(ws, 85, 8, 13.5)

set_blue(ws, 86, 4, 16.0)  # ARR RS FY21
set_blue(ws, 86, 5, 20.5)
set_blue(ws, 86, 6, 21.5)
set_blue(ws, 86, 7, 23.0)
set_blue(ws, 86, 8, 24.5)

# Row 89: Total R&D (gross, including capitalised)
set_blue(ws, 89, 4, 15.6)
set_blue(ws, 89, 5, 17.8)
set_blue(ws, 89, 6, 20.0)
set_blue(ws, 89, 7, 25.0)
set_blue(ws, 89, 8, 26.2)

# Row 90: Capitalised Development Costs (flow)
set_blue(ws, 90, 4, 0.0)
set_blue(ws, 90, 5, 0.0)
set_blue(ws, 90, 6, 0.0)
set_blue(ws, 90, 7, 12.0)
set_blue(ws, 90, 8, 13.3)

# Row 94: KPI-Shares Out (same as YE Shares)
set_blue(ws, 94, 4, 307.8)
set_blue(ws, 94, 5, 307.8)
set_blue(ws, 94, 6, 310.8)
set_blue(ws, 94, 7, 311.9)
set_blue(ws, 94, 8, 312.5)

# Row 95: KPI-WASO
set_blue(ws, 95, 4, 307.0)
set_blue(ws, 95, 5, 307.8)
set_blue(ws, 95, 6, 308.3)
set_blue(ws, 95, 7, 311.2)
set_blue(ws, 95, 8, 312.1)

# ==================
# BALANCE SHEET (Annual only)
# ==================
# Row 99: Cash
set_blue(ws, 99, 4, 26.2)
set_blue(ws, 99, 5, 28.8)
set_blue(ws, 99, 6, 47.2)
set_blue(ws, 99, 7, 66.5)
set_blue(ws, 99, 8, 79.2)

# Row 100: Trade Receivables
set_blue(ws, 100, 4, 19.1)
set_blue(ws, 100, 5, 21.5)
set_blue(ws, 100, 6, 21.8)
set_blue(ws, 100, 7, 22.0)
set_blue(ws, 100, 8, 23.5)

# Row 101: Contract Assets
set_blue(ws, 101, 4, 2.8)
set_blue(ws, 101, 5, 3.1)
set_blue(ws, 101, 6, 3.2)
set_blue(ws, 101, 7, 3.5)
set_blue(ws, 101, 8, 3.7)

# Row 102: Current Tax Asset / (Liability)
set_blue(ws, 102, 4, 0.5)
set_blue(ws, 102, 5, -0.3)
set_blue(ws, 102, 6, 0.8)
set_blue(ws, 102, 7, 0.2)
set_blue(ws, 102, 8, -0.4)

# Row 103: PPE
set_blue(ws, 103, 4, 8.3)
set_blue(ws, 103, 5, 7.8)
set_blue(ws, 103, 6, 6.8)
set_blue(ws, 103, 7, 5.7)
set_blue(ws, 103, 8, 4.5)

# Row 104: Intangibles
set_blue(ws, 104, 4, 45.3)
set_blue(ws, 104, 5, 44.5)
set_blue(ws, 104, 6, 51.4)
set_blue(ws, 104, 7, 59.7)
set_blue(ws, 104, 8, 66.7)

# Row 105: ROU Assets
set_blue(ws, 105, 4, 10.6)
set_blue(ws, 105, 5, 8.6)
set_blue(ws, 105, 6, 6.9)
set_blue(ws, 105, 7, 7.0)
set_blue(ws, 105, 8, 5.6)

# Row 106: Other Assets
set_blue(ws, 106, 4, 3.5)
set_blue(ws, 106, 5, 3.8)
set_blue(ws, 106, 6, 4.2)
set_blue(ws, 106, 7, 4.8)
set_blue(ws, 106, 8, 5.0)

# Row 114: Trade Payables
set_blue(ws, 114, 4, 5.2)
set_blue(ws, 114, 5, 5.8)
set_blue(ws, 114, 6, 5.5)
set_blue(ws, 114, 7, 6.0)
set_blue(ws, 114, 8, 6.5)

# Row 115: Contract Liabilities
set_blue(ws, 115, 4, 19.8)
set_blue(ws, 115, 5, 22.1)
set_blue(ws, 115, 6, 24.3)
set_blue(ws, 115, 7, 26.8)
set_blue(ws, 115, 8, 29.1)

# Row 116: Deferred Tax Liabilities
set_blue(ws, 116, 4, 1.8)
set_blue(ws, 116, 5, 2.0)
set_blue(ws, 116, 6, 2.5)
set_blue(ws, 116, 7, 4.5)
set_blue(ws, 116, 8, 6.2)

# Row 117: Provisions
set_blue(ws, 117, 4, 8.2)
set_blue(ws, 117, 5, 8.8)
set_blue(ws, 117, 6, 9.8)
set_blue(ws, 117, 7, 10.5)
set_blue(ws, 117, 8, 11.0)

# Row 118: Other Liabilities
set_blue(ws, 118, 4, 1.5)
set_blue(ws, 118, 5, 1.8)
set_blue(ws, 118, 6, 2.0)
set_blue(ws, 118, 7, 2.2)
set_blue(ws, 118, 8, 2.5)

# Row 119: Lease Liabilities
set_blue(ws, 119, 4, 12.3)
set_blue(ws, 119, 5, 10.0)
set_blue(ws, 119, 6, 8.2)
set_blue(ws, 119, 7, 8.3)
set_blue(ws, 119, 8, 6.7)

# Row 127: Issued Capital
set_blue(ws, 127, 4, 22.5)
set_blue(ws, 127, 5, 22.5)
set_blue(ws, 127, 6, 24.3)
set_blue(ws, 127, 7, 25.1)
set_blue(ws, 127, 8, 25.6)

# Row 128: Retained Profits
set_blue(ws, 128, 4, 37.9)
set_blue(ws, 128, 5, 43.2)
set_blue(ws, 128, 6, 48.2)
set_blue(ws, 128, 7, 62.0)
set_blue(ws, 128, 8, 75.8)

# Row 129: Reserves
set_blue(ws, 129, 4, 6.0)
set_blue(ws, 129, 5, 4.6)
set_blue(ws, 129, 6, 5.7)
set_blue(ws, 129, 7, 5.8)
set_blue(ws, 129, 8, 5.4)

# ==================
# CASH FLOW (Annual only) - actuals as blue hardcodes
# ==================
# Row 138: CF-WC Change
set_blue(ws, 138, 4, -0.6)
set_blue(ws, 138, 5, -5.2)
set_blue(ws, 138, 6, 1.9)
set_blue(ws, 138, 7, -2.7)
set_blue(ws, 138, 8, -2.3)

# Row 139: CF-Non Cash Items (SBP + other)
set_blue(ws, 139, 4, 1.0)
set_blue(ws, 139, 5, 1.3)
set_blue(ws, 139, 6, 1.5)
set_blue(ws, 139, 7, 1.6)
set_blue(ws, 139, 8, 1.7)

# Row 141: CF-Int Received
set_blue(ws, 141, 4, 0.3)
set_blue(ws, 141, 5, 0.2)
set_blue(ws, 141, 6, 1.5)
set_blue(ws, 141, 7, 3.2)
set_blue(ws, 141, 8, 3.4)

# Row 142: CF-Lease Int Paid
set_blue(ws, 142, 4, -0.6)
set_blue(ws, 142, 5, -0.5)
set_blue(ws, 142, 6, -0.4)
set_blue(ws, 142, 7, -0.4)
set_blue(ws, 142, 8, -0.3)

# Row 143: CF-Tax Paid
set_blue(ws, 143, 4, -3.7)
set_blue(ws, 143, 5, -5.3)
set_blue(ws, 143, 6, -3.5)
set_blue(ws, 143, 7, -5.9)
set_blue(ws, 143, 8, -6.8)

# Row 144: CF-Net OCF (hardcoded actual total)
set_blue(ws, 144, 4, 31.3, bold=True)
set_blue(ws, 144, 5, 26.8, bold=True)
set_blue(ws, 144, 6, 34.4, bold=True)
set_blue(ws, 144, 7, 42.4, bold=True)
set_blue(ws, 144, 8, 48.2, bold=True)

# Row 149: CF-Capex PPE (negative)
set_blue(ws, 149, 4, -1.8)
set_blue(ws, 149, 5, -2.2)
set_blue(ws, 149, 6, -1.3)
set_blue(ws, 149, 7, -1.0)
set_blue(ws, 149, 8, -0.7)

# Row 151: CF-Capex Intang (Capitalised Dev Costs, negative)
set_blue(ws, 151, 4, 0.0)
set_blue(ws, 151, 5, 0.0)
set_blue(ws, 151, 6, 0.0)
set_blue(ws, 151, 7, -12.0)
set_blue(ws, 151, 8, -13.3)

# Row 152: CF-Acquisitions (negative)
set_blue(ws, 152, 4, 0.0)
set_blue(ws, 152, 5, 0.0)
set_blue(ws, 152, 6, -10.0)
set_blue(ws, 152, 7, 0.0)
set_blue(ws, 152, 8, 0.0)

# Row 153: CF-Other CFI
set_blue(ws, 153, 4, 0.1)
set_blue(ws, 153, 5, 0.1)
set_blue(ws, 153, 6, 0.0)
set_blue(ws, 153, 7, 0.1)
set_blue(ws, 153, 8, 0.0)

# Row 154: Total Investing CF (hardcoded total)
set_blue(ws, 154, 4, -1.7, bold=True)
set_blue(ws, 154, 5, -2.1, bold=True)
set_blue(ws, 154, 6, -11.3, bold=True)
set_blue(ws, 154, 7, -12.9, bold=True)
set_blue(ws, 154, 8, -14.0, bold=True)

# Row 157: CF-Dividends (negative)
set_blue(ws, 157, 4, -37.5)
set_blue(ws, 157, 5, -40.0)
set_blue(ws, 157, 6, -46.2)
set_blue(ws, 157, 7, -52.4)
set_blue(ws, 157, 8, -65.5)

# Row 158: CF-Share Issues
set_blue(ws, 158, 4, 0.0)
set_blue(ws, 158, 5, 0.0)
set_blue(ws, 158, 6, 1.8)
set_blue(ws, 158, 7, 0.8)
set_blue(ws, 158, 8, 0.5)

# Row 159: CF-Lease Principal (negative)
set_blue(ws, 159, 4, -3.6)
set_blue(ws, 159, 5, -3.4)
set_blue(ws, 159, 6, -3.0)
set_blue(ws, 159, 7, -2.9)
set_blue(ws, 159, 8, -3.0)

# Row 160: CF-Other CFF
set_blue(ws, 160, 4, 0.0)
set_blue(ws, 160, 5, 0.0)
set_blue(ws, 160, 6, 0.0)
set_blue(ws, 160, 7, 0.0)
set_blue(ws, 160, 8, 0.0)

# Row 161: Total Financing CF (hardcoded total)
set_blue(ws, 161, 4, -41.1, bold=True)
set_blue(ws, 161, 5, -43.4, bold=True)
set_blue(ws, 161, 6, -47.4, bold=True)
set_blue(ws, 161, 7, -54.5, bold=True)
set_blue(ws, 161, 8, -68.0, bold=True)

# ================================
# HY & SEGMENTS SHEET - 1H actuals
# ================================
ws2 = wb['HY & Segments']
# Cols: D=1H21, E=2H21, F=1H22, G=2H22, H=1H23, I=2H23, J=1H24, K=2H24, L=1H25, M=2H25, N=1H26

# 1H Revenue splits (from plan)
# Row 7: Rev-Info Intelligence
h1_ii = {4: 33.4, 6: 36.7, 8: 38.1, 10: 39.5, 12: 41.1, 14: 43.0}
for col, val in h1_ii.items():
    set_blue(ws2, 7, col, val)

# Row 8: Rev-Planning & Building
h1_pb = {4: 5.3, 6: 6.1, 8: 5.9, 10: 6.1, 12: 6.5, 14: 9.0}
for col, val in h1_pb.items():
    set_blue(ws2, 8, col, val)

# Row 9: Rev-Regulatory Solutions
h1_rs = {4: 7.6, 6: 9.8, 8: 10.4, 10: 10.8, 12: 11.9, 14: 12.9}
for col, val in h1_rs.items():
    set_blue(ws2, 9, col, val)

# Row 10: Interest Income (1H)
h1_int = {4: 0.1, 6: 0.1, 8: 0.5, 10: 1.5, 12: 1.7, 14: 1.8}
for col, val in h1_int.items():
    set_blue(ws2, 10, col, val)

# Row 15: COGS (1H, negative)
h1_cogs = {4: -2.5, 6: -2.7, 8: -2.8, 10: -3.2, 12: -3.3, 14: -3.5}
for col, val in h1_cogs.items():
    set_blue(ws2, 15, col, val)

# Row 23: Distribution (1H, negative)
h1_dist = {4: -15.5, 6: -17.2, 8: -18.8, 10: -19.6, 12: -20.1, 14: -21.0}
for col, val in h1_dist.items():
    set_blue(ws2, 23, col, val)

# Row 24: R&D Expense (1H, negative)
h1_rd = {4: -7.5, 6: -8.7, 8: -9.8, 10: -6.4, 12: -6.3, 14: -6.5}
for col, val in h1_rd.items():
    set_blue(ws2, 24, col, val)

# Row 25: Admin (1H, negative)
h1_admin = {4: -3.6, 6: -4.0, 8: -3.9, 10: -4.2, 12: -4.7, 14: -5.1}
for col, val in h1_admin.items():
    set_blue(ws2, 25, col, val)

# Row 35: SBP (1H, negative)
h1_sbp = {4: -0.5, 6: -0.6, 8: -0.7, 10: -0.8, 12: -0.8, 14: -0.9}
for col, val in h1_sbp.items():
    set_blue(ws2, 35, col, val)

# Row 36: M&A Costs
h1_ma = {4: 0.0, 6: 0.0, 8: -0.7, 10: 0.0, 12: 0.0, 14: 0.0}
for col, val in h1_ma.items():
    set_blue(ws2, 36, col, val)

# Row 37: FX
h1_fx = {4: 0.0, 6: -0.1, 8: 0.1, 10: -0.1, 12: 0.0, 14: 0.0}
for col, val in h1_fx.items():
    set_blue(ws2, 37, col, val)

# Row 41: Depreciation PPE (1H, negative)
h1_dep = {4: -1.3, 6: -1.3, 8: -1.4, 10: -1.3, 12: -1.2, 14: -1.1}
for col, val in h1_dep.items():
    set_blue(ws2, 41, col, val)

# Row 42: ROU Amortisation (1H, negative)
h1_rou = {4: -1.5, 6: -1.5, 8: -1.4, 10: -1.5, 12: -1.5, 14: -1.4}
for col, val in h1_rou.items():
    set_blue(ws2, 42, col, val)

# Row 43: Amort Dev Costs (1H, negative, $0 pre-FY24)
h1_adc = {4: 0.0, 6: 0.0, 8: 0.0, 10: -1.3, 12: -2.7, 14: -3.2}
for col, val in h1_adc.items():
    set_blue(ws2, 43, col, val)

# Row 53: Interest Income (1H) - same as Rev-Interest Income
for col, val in h1_int.items():
    set_blue(ws2, 53, col, val)

# Row 54: Lease Interest (1H, negative)
h1_lint = {4: -0.3, 6: -0.3, 8: -0.2, 10: -0.2, 12: -0.2, 14: -0.1}
for col, val in h1_lint.items():
    set_blue(ws2, 54, col, val)

# Row 58: Tax Expense (1H, negative)
h1_tax = {4: -2.3, 6: -2.4, 8: -1.5, 10: -2.8, 12: -3.4, 14: -3.8}
for col, val in h1_tax.items():
    set_blue(ws2, 58, col, val)

# Row 61: Other Items AT (1H)
h1_oi = {4: 0.0, 6: -0.1, 8: -0.4, 10: -0.6, 12: -0.6, 14: -0.6}
for col, val in h1_oi.items():
    set_blue(ws2, 61, col, val)

# Operating Metrics (1H point-in-time where applicable)
# Row 66-68: ARR by segment (point-in-time at 1H end = 31 Dec)
h1_arr_ii = {4: 68.0, 6: 74.5, 8: 77.0, 10: 80.0, 12: 83.5, 14: 87.0}
for col, val in h1_arr_ii.items():
    set_blue(ws2, 66, col, val)

h1_arr_pb = {4: 10.5, 6: 11.5, 8: 11.8, 10: 12.5, 12: 13.0, 14: 15.5}
for col, val in h1_arr_pb.items():
    set_blue(ws2, 67, col, val)

h1_arr_rs = {4: 15.0, 6: 19.5, 8: 21.0, 10: 22.5, 12: 24.0, 14: 26.0}
for col, val in h1_arr_rs.items():
    set_blue(ws2, 68, col, val)

# Row 71: Total R&D (1H, flow)
h1_trd = {4: 7.5, 6: 8.7, 8: 9.8, 10: 12.0, 12: 12.8, 14: 13.5}
for col, val in h1_trd.items():
    set_blue(ws2, 71, col, val)

# Row 72: Capitalised Dev (1H, flow)
h1_cap = {4: 0.0, 6: 0.0, 8: 0.0, 10: 5.6, 12: 6.5, 14: 7.0}
for col, val in h1_cap.items():
    set_blue(ws2, 72, col, val)

# Row 76: Shares Outstanding (point-in-time at 1H end)
h1_shares = {4: 307.0, 6: 307.8, 8: 308.5, 10: 311.0, 12: 312.0, 14: 312.8}
for col, val in h1_shares.items():
    set_blue(ws2, 76, col, val)

# Row 77: WASO Basic (1H, flow)
h1_waso = {4: 306.5, 6: 307.5, 8: 307.8, 10: 309.5, 12: 311.5, 14: 312.3}
for col, val in h1_waso.items():
    set_blue(ws2, 77, col, val)

# 2H ARR (point-in-time = FY end values, same as Annual)
# Col E=2H21, G=2H22, I=2H23, K=2H24, M=2H25
h2_arr_ii = {5: 70.5, 7: 76.0, 9: 78.5, 11: 82.0, 13: 86.0}
for col, val in h2_arr_ii.items():
    set_blue(ws2, 66, col, val)

h2_arr_pb = {5: 11.0, 7: 12.0, 9: 12.2, 11: 12.8, 13: 13.5}
for col, val in h2_arr_pb.items():
    set_blue(ws2, 67, col, val)

h2_arr_rs = {5: 16.0, 7: 20.5, 9: 21.5, 11: 23.0, 13: 24.5}
for col, val in h2_arr_rs.items():
    set_blue(ws2, 68, col, val)

# 2H Shares Outstanding (point-in-time = FY end)
h2_shares = {5: 307.8, 7: 307.8, 9: 310.8, 11: 311.9, 13: 312.5}
for col, val in h2_shares.items():
    set_blue(ws2, 76, col, val)

wb.save(DST)
print('Part 3 complete: historical actuals entered on Annual (FY21-FY25) and HY (1H21-1H26)')
