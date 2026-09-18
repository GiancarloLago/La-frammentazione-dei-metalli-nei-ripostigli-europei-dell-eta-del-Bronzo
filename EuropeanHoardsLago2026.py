"""
Figure_Monografia.py

Companion script for the figures published in:
"La frammentazione dei metalli nei ripostigli europei dell'eta del Bronzo"
(Lago, Scienze dell'Antichita monograph series).

This script reproduces, in sequence, all the data visualizations used
in the volume. Each section is clearly labelled with the corresponding
figure number(s) as printed in the book.

Requirements: pandas, numpy, matplotlib, openpyxl, python-docx (for the
one section that also exports a supporting Word table), tqdm (optional,
for a progress bar during a Monte Carlo simulation).

Usage: run this script from the same directory as DB.xlsx (the source
database, sheet "DB"). 
"""



# %%

# ============================================================
# COMMON IMPORTS
# Shared libraries used across the notebook's figure-generating cells.
# Run this cell first.
# ============================================================

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import re
from matplotlib.collections import LineCollection
import matplotlib.colors as mcolors
from matplotlib.patches import Patch
from matplotlib.lines import Line2D
from matplotlib.patches import Rectangle
import colorsys

# %%

# ============================================================
# FIGURE 4 — Comparison of Nordic and Central European chronologies
# Each horizontal segment represents the reconstructed chronological
# interval (years BC) of a single hoard, split into two panels:
# Nordic Chronology (NC) and Central European Chronology (SC).
# ============================================================

# ================== PARAMETERS ==================
filename = "DB.xlsx"
sheet = "DB"

# half page ~ 160 x 110 mm
FIGSIZE = (6.30, 4.33)   # inch
DPI = 300
LW = 1.4                 # line width

# Desired X axis
START_NC = 1700; RIGHT_NC = 800    # NC fixed 1700 -> 800
START_SC = 2150                    # SC starts at 2150 (ends at min(End) from data)

# Sub-period order (OR_1 oldest, at the top)
OR_ORDER = ["OR_1","OR_2","OR_3","OR_4","OR_5"]

# Vertical lines (now BLACK)
VLINES_NC = [1500, 1330, 1100, 950, 800]
VLINES_SC = [1550, 1330, 1080, 960]

# Custom ticks
TICKS_NC = [1700, 1500, 1300, 1100, 900]
TICKS_SC = [2150, 1850, 1550, 1250, 950]

# Base colors
BASE_NC = "#2e7d6e"   # green (Nordic)
BASE_SC = "#b24c4c"   # red (Central)

# ================== HELPER FUNCTIONS ==================
def col_like(df, name):
    low = {c.lower().strip(): c for c in df.columns}
    return low.get(name.lower().strip(), name)

def parse_bc_range(s):
    m = re.search(r"(\d+)\s*-\s*(\d+)\s*BC", str(s))
    if not m:
        return np.nan, np.nan
    a, b = int(m.group(1)), int(m.group(2))
    return a, b

def hex_to_hls(hexcolor):
    r, g, b = mcolors.to_rgb(hexcolor)
    return colorsys.rgb_to_hls(r, g, b)

def hls_to_hex(h, l, s):
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    return mcolors.to_hex((r, g, b))

def get_muted_dark_color(base_hex, sat_factor=0.55, lightness=0.40):
    h, l, s = hex_to_hls(base_hex)
    l = np.clip(lightness, 0, 1)
    s = np.clip(s * sat_factor, 0, 1)
    return hls_to_hex(h, l, s)

def prepare_panel(data, col_or):
    rows, y = [], 0.0
    GAP = 0.8
    for or_lab in OR_ORDER:
        block = data[data[col_or] == or_lab].copy()
        if block.empty:
            continue
        block = block.sort_values(by=["Start","Dur","End"],
                                  ascending=[False, True, False]).reset_index(drop=True)
        for i, r in block.iterrows():
            r["y"] = y + i
            rows.append(r)
        y += len(block) + GAP
    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=list(data.columns)+["y"])

def xlims(panel, left_override=None, right_override=None):
    if panel.empty:
        return (0, 0)
    left  = left_override  if left_override  is not None else int(panel["Start"].max())
    right = right_override if right_override is not None else int(panel["End"].min())
    if left <= right:
        left, right = right+1, right-1
    return (left, right)


def draw(ax, panel, title, family, xlim, vlines, ticks, col_or,
         add_left_edge=False, add_right_edge=False):
    if panel.empty:
        ax.set_visible(False); return

    # Single dark color for the whole panel
    base_hex = BASE_NC if family == "NC" else BASE_SC
    color = get_muted_dark_color(base_hex, sat_factor=0.55, lightness=0.40)

    segs = [[(r["End"], r["y"]), (r["Start"], r["y"])] for _, r in panel.iterrows()]
    lc = LineCollection(segs, colors=[color]*len(segs), linewidths=LW,
                        capstyle='butt', antialiased=False)
    ax.add_collection(lc)

    # Linee verticali NERE
    for x in vlines:
        ax.axvline(x, color="black", linewidth=0.8, zorder=0)
    if add_left_edge:
        ax.axvline(xlim[0], color="black", linewidth=0.8, zorder=0)   # bordo sinistro
    if add_right_edge:
        ax.axvline(xlim[1], color="black", linewidth=0.8, zorder=0)   # bordo destro

    # Assi
    ax.set_xlim(xlim[0], xlim[1])                        
    ax.set_ylim(panel["y"].min()-0.8, panel["y"].max()+0.8)
    ax.invert_yaxis()                                    

    ticks_in = [t for t in ticks if xlim[1] <= t <= xlim[0]]
    ax.set_xticks(ticks_in)
    ax.set_xticklabels([str(t) for t in ticks_in], fontsize=8)

    ax.set_title(title, fontsize=10, pad=6)
    ax.set_xlabel("Years BC", fontsize=8, labelpad=4)
    ax.set_yticks([])
    for s in ("left","top","right"): ax.spines[s].set_visible(False)
    ax.grid(False)

    ax.text(0.01, 0.02, f"n = {len(panel)}",
            transform=ax.transAxes, ha="left", va="bottom",
            fontsize=8, color="0.25")


# ================== READ & PREP DATA ==================
df = pd.read_excel(filename, sheet_name=sheet)

col_id   = col_like(df, "ID_Sito")
col_date = col_like(df, "Date_Chrono")
col_ns   = col_like(df, "N/S Chrono")
col_or   = col_like(df, "Or_fin")

df = df.drop_duplicates(subset=col_id).copy()
df[["Start","End"]] = df[col_date].apply(lambda x: pd.Series(parse_bc_range(x)))
df = df.dropna(subset=["Start","End", col_ns, col_or]).copy()
df["Start"] = df["Start"].astype(int)
df["End"]   = df["End"].astype(int)
df["Dur"]   = df["Start"] - df["End"]
df[col_or] = (df[col_or].astype(str).str.upper().str.replace(r"\s+","_", regex=True))
df[col_ns] = (df[col_ns].astype(str).str.upper().str.replace(r"\s+","", regex=True))

df_nc = df[df[col_ns] == "NC"].copy()
df_sc = df[df[col_ns] == "SC"].copy()

nc_plot = prepare_panel(df_nc, col_or)
sc_plot = prepare_panel(df_sc, col_or)

x_nc = xlims(nc_plot, left_override=START_NC, right_override=RIGHT_NC)
x_sc = xlims(sc_plot, left_override=START_SC, right_override=None)

# ================== PLOT AFFIANCATO ==================
fig, (ax_nc, ax_sc) = plt.subplots(1, 2, figsize=FIGSIZE, dpi=DPI)

draw(ax_nc, nc_plot, "Nordic Chronology", "NC", x_nc, VLINES_NC, TICKS_NC, col_or,
     add_left_edge=True, add_right_edge=False)

draw(ax_sc, sc_plot, "Central European Chronology", "SC", x_sc, VLINES_SC, TICKS_SC, col_or,
     add_left_edge=True, add_right_edge=True)

plt.tight_layout(w_pad=0.9)

# Save + Jupyter display
#fig.savefig("Fig04_Chronologies_side_by_side_HALF_dark.jpg", dpi=DPI, format="jpeg",
#            pil_kwargs={"quality":95, "subsampling":0})
plt.show()


# %%

# ============================================================
# FIGURE 5 — Quantity of fragments, damaged and intact objects
# from the Bronze Age in Central Europe, by chronological horizon.
# ============================================================

# ================== PARAMETERS FILE ==================
FILENAME = "DB.xlsx"
SHEET = "DB"

# ================== COLUMN NAMES (adjust if needed) ==================
COL_OR   = "Or_fin"
COL_COMP = "Complete"
COL_FRAG = "Fragmented"      # if the file uses "Fragment", it is corrected below
COL_MATCH= "Matching fr"
COL_BELT = "Belted"

# ================== FIGURE ==================
FIGSIZE = (6.30, 4.33)   # ~160 x 110 mm (half page)
DPI = 300

# ================== COLORS ==================
COL_COMPLETE_BASE = "#000000"   # black
COL_FRAGMENT_BASE = "#CFCFCF"   # light gray
COL_MATCHING      = "#2e7d6e"   # muted green (Matching)
COL_DAMAGED       = "#a85d4d"   # ruggine (Damaged)
COL_NS   = "N/S Chrono"   

# ================== READ DATA ==================
df = pd.read_excel(FILENAME, sheet_name=SHEET)

def col_like(df, name):
    low = {c.lower().strip(): c for c in df.columns}
    return low.get(name.lower().strip(), name)

# allinea eventuali varianti
COL_OR   = col_like(df, COL_OR)
COL_COMP = col_like(df, COL_COMP)
COL_FRAG = COL_FRAG if COL_FRAG in df.columns else col_like(df, "Fragment")
COL_MATCH= col_like(df, COL_MATCH)
COL_BELT = col_like(df, COL_BELT)
COL_NS   = col_like(df, COL_NS)

# **Exclude rows where N/S Chrono = IC (any case/spacing)**
df = df[~df[COL_NS].astype(str).str.upper().str.strip().eq("IC")].copy()

# helper: "true" if any numeric value > 0 (also works with numeric-like strings)
def as_bool_any_positive(s):
    s_num = pd.to_numeric(s, errors="coerce")
    if s_num.notna().any():
        return (s_num.fillna(0) > 0)
    return s.astype(str).str.strip().str.lower().isin(["1","true","yes","y","si","s"])

def as_int_safe(s):
    return pd.to_numeric(s, errors="coerce").fillna(0).astype(int)

or_fin      = df[COL_OR].astype(str).str.upper().str.replace(r"\s+","_", regex=True)
is_complete = as_bool_any_positive(df[COL_COMP])
is_fragment = as_bool_any_positive(df[COL_FRAG])
is_belted   = as_bool_any_positive(df[COL_BELT])
match_num   = as_int_safe(df[COL_MATCH])

df_std = pd.DataFrame({
    "Or_fin": or_fin,
    "Complete": is_complete,
    "Fragment": is_fragment,
    "Belted": is_belted,
    "Match": match_num
})

OR_ORDER = ["OR_1","OR_2","OR_3","OR_4","OR_5"]
df_std = df_std[df_std["Or_fin"].isin(OR_ORDER)]

# ================== AGGREGAZIONE ==================
rows = []
for or_lab in OR_ORDER:
    sub = df_std[df_std["Or_fin"] == or_lab]

    # --- COMPLETE (bar height = number of complete objects)
    C_total = sub["Complete"].sum()
    C_belt  = (sub["Complete"] & sub["Belted"]).sum()
    C_match = (sub["Complete"] & (sub["Match"] > 0)).sum()
    C_both  = (sub["Complete"] & sub["Belted"] & (sub["Match"] > 0)).sum()
    C_plain       = int(C_total - (C_belt + (C_match - C_both)))  # parte nera
    C_match_only  = int(C_match - C_both)                         # green overlay
    C_belt_any    = int(C_belt)                                   # overlay ruggine

    # --- FRAGMENTS
    # Base height = number of true fragments
    F_base = sub["Fragment"].sum()
    # Belted among fragments = only those within the base (subset, does not add height)
    F_belt = (sub["Fragment"] & sub["Belted"]).sum()
    # Matching that is **added on top**: all rows with Match>0 and **not Complete**
    F_match_sum = int(sub.loc[(~sub["Complete"]) & (sub["Match"] > 0), "Match"].sum())
    F_plain = int(F_base - F_belt)
    F_total_plus_match = int(F_base + F_match_sum)

    rows.append({
        "Or": or_lab,
        # completi
        "C_total": int(C_total),
        "C_plain": C_plain,
        "C_match_only": C_match_only,
        "C_belt": C_belt_any,
        # fragments
        "F_total_plus_match": F_total_plus_match,
        "F_plain": F_plain,
        "F_belt": int(F_belt),
        "F_match_sum": F_match_sum
    })

agg = pd.DataFrame(rows)

# ================== CLEAN Y AXIS + HEADROOM ==================
y_max = max(agg["C_total"].max(), agg["F_total_plus_match"].max())
def nice_step(v):
    m = 10 ** int(np.floor(np.log10(max(v,1))))
    for k in [1,2,5,10]:
        if v / (m*k) <= 8:
            return m*k
    return m*10
step = nice_step(y_max)
y_lim = int(np.ceil((y_max*1.25) / step) * step)  # +25% headroom above

# ================== PLOT ==================
fig, ax = plt.subplots(figsize=FIGSIZE, dpi=DPI)

x_idx = np.arange(len(OR_ORDER))
width = 0.32                      # slightly narrower bars
gap = 0.06                        # piccolo stacco tra le due barre
x_complete = x_idx - width/2 - gap/2
x_fragment = x_idx + width/2 + gap/2

# thin white edge (use edgecolor="none" if only the gap is preferred)
edge_kw = dict(edgecolor="white", linewidth=0.6)

# COMPLETI (somma = C_total)
ax.bar(x_complete, agg["C_plain"], width, color=COL_COMPLETE_BASE, **edge_kw, label="Complete")
ax.bar(x_complete, agg["C_match_only"], width, bottom=agg["C_plain"], color=COL_MATCHING, **edge_kw, label="Matching")
ax.bar(x_complete, agg["C_belt"], width, bottom=agg["C_plain"] + agg["C_match_only"], color=COL_DAMAGED, **edge_kw, label="Damaged (Comp/Fragm)")

# --- FRAGMENTS (sum = F_base + sum(Matching))
F_plain = agg["F_plain"].values           # = fragmented - fragmented_belted
F_belt  = agg["F_belt"].values            # rust-colored overlay (subset, does not add height)
F_match = agg["F_match_sum"].values       # si AGGIUNGE in cima

# 1) base grigia
ax.bar(x_fragment, F_plain, width,
       color=COL_FRAGMENT_BASE, **edge_kw,
       label="Fragmented", zorder=2)

# 2) overlay "Damaged" (ruggine) dentro la quota base
ax.bar(x_fragment, F_belt, width, bottom=F_plain,
       color=COL_DAMAGED, **edge_kw,
       zorder=3)

# 3) "Matching" overlay (green) ADDED on top
#    -> bottom = F_plain + F_belt  (so it does not cover the rust-colored portion)
ax.bar(x_fragment, F_match, width, bottom=F_plain + F_belt,
       color=COL_MATCHING, **edge_kw,
       label="Matching", zorder=4)



# X & separatori
ax.set_xticks(x_idx)
ax.set_xticklabels([f"Or. {i+1}" for i in range(len(OR_ORDER))])
for i in range(len(OR_ORDER)-1):
    ax.axvline(i + 0.5, color="0.85", lw=0.7, ls=(0,(4,3)), zorder=0)

# Y
ax.set_ylabel("Counts", fontsize=10)  # era 9
ax.tick_params(axis='both', labelsize=9)


# Compact legend (single "Damaged (Comp/Fragm)" and single "Matching")
legend_patches = [
    Patch(facecolor=COL_COMPLETE_BASE, edgecolor='none', label="Complete"),
    Patch(facecolor=COL_FRAGMENT_BASE, edgecolor='none', label="Fragmented"),
    Patch(facecolor=COL_DAMAGED, edgecolor='none', label="Damaged (Comp/Fragm)"),
    Patch(facecolor=COL_MATCHING, edgecolor='none', label="Matching"),
]

leg = ax.legend(
    handles=legend_patches,
    loc="upper center",
    bbox_to_anchor=(0.5, 1.10),   # nudge slightly higher
    ncol=2,
    fontsize=10,                  # ← increase here (10-11 recommended)
    frameon=False,
    handlelength=2.0,
    handletextpad=0.6,
    labelspacing=0.6,
    columnspacing=1.2,
    borderaxespad=0.2,
)


# Stile pulito
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
plt.tight_layout()

# Save + display
#plt.savefig("Fig05_Complete_Fragment_by_Or.jpg", dpi=300, format="jpeg")
plt.show()


# %%

# ============================================================
# FIGURE 6 — Share of fragmentation by period, compared between
# the Nordic Chronology (NC) and the Central European Chronology
# (SC) areas. Side-by-side panels with Complete / Fragmented /
# Matching fragment counts for each chronological horizon.
# ============================================================

# ================== FILE & SHEET ==================
FILENAME = "DB.xlsx"
SHEET = "DB"

# ================== DIMENSIONS & COLORS ==================
FIGSIZE = (6.30, 4.33)   # ~160 x 110 mm (half page)
DPI = 300
EDGE_KW = dict(edgecolor="white", linewidth=0.6)  # thin white edge

COL_COMPLETE  = "#000000"   # black
COL_FRAGMENT  = "#CFCFCF"   # light gray
COL_MATCHING  = "#2e7d6e"   # muted green

# ================== COLUMN NAMES (adjust if different) ==================
COL_OR   = "Or_fin"
COL_NS   = "N/S Chrono"     # "NC" / "SC"
COL_COMP = "Complete"
COL_FRAG = "Fragmented"     # if the file uses "Fragment", it is caught below
COL_MATCH= "Matching fr"

# ------------------ utilities ------------------
def col_like(df, name):
    low = {c.lower().strip(): c for c in df.columns}
    return low.get(name.lower().strip(), name)

def as_bool_any_positive(s):
    s_num = pd.to_numeric(s, errors="coerce")
    if s_num.notna().any():
        return (s_num.fillna(0) > 0)
    return s.astype(str).str.strip().str.lower().isin(["1","true","yes","y","si","s"])

def as_int_safe(s):
    return pd.to_numeric(s, errors="coerce").fillna(0).astype(int)

# ================== READ & NORMALIZE ==================
df = pd.read_excel(FILENAME, sheet_name=SHEET)

COL_OR   = col_like(df, COL_OR)
COL_NS   = col_like(df, COL_NS)
COL_COMP = col_like(df, COL_COMP)
COL_FRAG = COL_FRAG if COL_FRAG in df.columns else col_like(df, "Fragment")
COL_MATCH= col_like(df, COL_MATCH)

df_std = pd.DataFrame({
    "Or_fin": df[COL_OR].astype(str).str.upper().str.replace(r"\s+","_", regex=True),
    "NS":     df[COL_NS].astype(str).str.upper().str.strip(),  # "NC"/"SC"
    "Complete": as_bool_any_positive(df[COL_COMP]),
    "Fragment": as_bool_any_positive(df[COL_FRAG]),
    "Match":    as_int_safe(df[COL_MATCH]),
})

OR_ORDER = ["OR_1","OR_2","OR_3","OR_4","OR_5"]
df_std = df_std[df_std["Or_fin"].isin(OR_ORDER)]

# ================== AGGREGATION FUNCTION ==================
def aggregate_region(df_region):
    rows = []
    for or_lab in OR_ORDER:
        sub = df_region[df_region["Or_fin"] == or_lab]

        # Complete (plain count)
        n_complete = int(sub["Complete"].sum())

        # Fragmented (plain count)
        n_fragment = int(sub["Fragment"].sum())

        # Matching fragments:
        # - for Complete objects: count complete objects with Match>0 (1 each)
        match_from_complete = int((sub["Complete"] & (sub["Match"] > 0)).sum())
        # - for non-Complete objects: sum of Match (even if Fragment==0)
        match_from_non_complete = int(sub.loc[(~sub["Complete"]) & (sub["Match"] > 0), "Match"].sum())
        n_matching_total = match_from_complete + match_from_non_complete

        rows.append({
            "Or": or_lab,
            "Complete": n_complete,
            "Fragmented": n_fragment,
            "Matching": n_matching_total
        })
    out = pd.DataFrame(rows)
    # for convenience, also compute the panel's Y max
    out["Total_max_like"] = out[["Complete","Fragmented","Matching"]].sum(axis=1)
    return out

nc_agg = aggregate_region(df_std[df_std["NS"]=="NC"])
sc_agg = aggregate_region(df_std[df_std["NS"]=="SC"])

# ================== PLOT SIDE-BY-SIDE ==================
fig, (ax_nc, ax_sc) = plt.subplots(1, 2, figsize=FIGSIZE, dpi=DPI)

def nice_step(v):
    m = 10 ** int(np.floor(np.log10(max(v,1))))
    for k in [1,2,5,10]:
        if v / (m*k) <= 8:
            return m*k
    return m*10

def draw_panel(ax, agg, title):
    x = np.arange(len(agg))
    width = 0.22
    gap = 0.04
    # positions for 3 bars: Complete | Fragmented | Matching
    xC = x - width - gap
    xF = x
    xM = x + width + gap

    # bars
    ax.bar(xC, agg["Complete"],  width, color=COL_COMPLETE, **EDGE_KW, label="Complete")
    ax.bar(xF, agg["Fragmented"],width, color=COL_FRAGMENT, **EDGE_KW, label="Fragmented")
    ax.bar(xM, agg["Matching"],  width, color=COL_MATCHING, **EDGE_KW, label="Matching")

    # x axis
    ax.set_xticks(x)
    ax.set_xticklabels([f"Or. {i+1}" for i in range(len(agg))])

    # separators between groups
    for i in range(len(agg)-1):
        ax.axvline(i + 0.5, color="0.85", lw=0.7, ls=(0,(4,3)), zorder=0)

    # clean Y axis with headroom
    y_max = max(agg[["Complete","Fragmented","Matching"]].to_numpy().max(), 1)
    step = nice_step(y_max)
    y_lim = int(np.ceil((y_max*1.25) / step) * step)
    ax.set_ylim(0, y_lim)
    ax.set_yticks(np.arange(0, y_lim+1, step))
    ax.set_ylabel("Counts", fontsize=9)

    # style
    ax.set_title(title, fontsize=10, pad=6)
    for s in ("top","right"): ax.spines[s].set_visible(False)

draw_panel(ax_nc, nc_agg, "Nordic (NC)")
draw_panel(ax_sc, sc_agg, "Central European (SC)")

# --- LEGEND INSIDE THE CHART (replaces the previous fig.legend block) ---
handles = [
    plt.Rectangle((0,0),1,1,color=COL_COMPLETE),
    plt.Rectangle((0,0),1,1,color=COL_FRAGMENT),
    plt.Rectangle((0,0),1,1,color=COL_MATCHING),
]
labels = ["Complete", "Fragmented", "Matching fragments"]

ax_nc.legend(handles, labels,
             loc="upper right",      # inside the Nordic panel area
             fontsize=10,            # larger for print
             frameon=False,
             ncol=1,
             handlelength=1.6,
             handletextpad=0.5,
             labelspacing=0.4)

plt.tight_layout()
# (remove/ignore any fig.subplots_adjust(top=...) previously used for the external legend)
fig.subplots_adjust(top=0.85)  # room for the legend above

# Save + display
#fig.savefig("Fig06_NC_SC_bar_triple_counts.jpg", dpi=DPI, format="jpeg",
#            pil_kwargs={"quality":95, "subsampling":0})
plt.show()


# %%

# ============================================================
# FIGURE 16 — Summary of intact/fragmented objects by regions,
# states, and areas (Horizons 3-5). Grid of 12 panels (one per
# region/state), grouped by artefact category, with dashed
# rectangles marking functional families (Weapons, Tools,
# Jewels, Ingots).
# ============================================================

# ================== GRID 3x4 - REGIONS x CATEGORIES (Complete vs Fragmented) ==================
import re
import unicodedata
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib.patches import Rectangle

# ===== File / foglio =====
FILENAME = "DB.xlsx"
SHEET = "DB"

# ===== columns ("logical" names; verranno risolti in modo robusto) =====
COL_OR = "Or_fin"
COL_STATE = "State"
COL_REGION = "Region"
COL_TYPE = "Artefact"
COL_COMP = "Complete"
COL_FRAG = "Fragmented"   # fallback anche a "Fragment"
COL_MATCH = "Matching fr"

# ===== Period filter =====
OR_KEEP = ["OR_3", "OR_4", "OR_5"]

# ===== Final categories (x-axis order) =====
CATS = [
    "Spearheads", "Swords", "Daggers", "Knives",
    "Axes", "Sickles",
    "Bracelets", "Necklaces", "Other Jewels",
    "Ingots", "Bronze Sheets", "Golden obj.", "Other",
]

# ============== FIGURE (page approx. 160x235 mm) ==============
FIGSIZE = (6.30, 9.25)   # in pollici
DPI = 300
BAR_W = 0.72
EDGE_KW = dict(edgecolor="white", linewidth=0.6)

# --- bar colors ---
COL_COMPLETE = "#000000"  # black
COL_FRAGMIX = "#CFCFCF"   # light gray

# Weapons: 0-2; Tools: 4-5; Jewels: 6-8; Ingots: 9-9
grp_idx = {
    "Weapons": (0, 2),   # Spearheads-Swords-Daggers
    "Tools":   (4, 5),   # Axes-Sickles
    "Jewels":  (6, 8),   # Bracelets-Necklaces-Other Jewels
    "Ingots":  (9, 9),   # solo Ingots
}
DASH = (0, (6, 3))
grp_style = {
    "Weapons": dict(color="#1f77b4", ls=DASH, lw=0.9),
    "Tools":   dict(color="#ff7f0e", ls=DASH, lw=0.9),
    "Jewels":  dict(color="#2ca02c", ls=DASH, lw=0.9),
    "Ingots":  dict(color="#9467bd", ls=DASH, lw=0.9),
}

def col_like(df, name):
    low = {c.lower().strip(): c for c in df.columns}
    alts = {
        "artefact": ["artefact", "artifact", "type", "object", "category", "class", "item", "typology"],
        "fragmented": ["fragmented", "fragment"],
        "matching fr": ["matching fr", "matching_fr", "matchingfr", "matching"],
    }
    k = name.lower().strip()
    if k in alts:
        for cand in alts[k]:
            if cand in low:
                return low[cand]
    return low.get(k, name)

def to_str_u(s: pd.Series) -> pd.Series:
    return s.astype(str).str.upper().str.strip()

def as_int(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").fillna(0).astype(int)

def strip_accents(s: str) -> str:
    s = s.replace("ß", "ss")
    return "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))

def norm(s: str) -> str:
    s = strip_accents(str(s)).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

# ---- classificatore categorie ----
def categorize(artefact_raw: str) -> str:
    s = norm(artefact_raw)

    # Ingots (varie forme)
    if re.search(r"\b(ingot|ingots|bun[- ]?ingot|bun|barren|barrens)\b", s):
        return "Ingots"

    # Golden obj.: anything STARTING with "gold" (but not ingots)
    if s.startswith("gold"):
        return "Golden obj."

    # Mappature principali
    if "spearhead" in s: return "Spearheads"
    if "sword" in s:     return "Swords"
    if "dagger" in s:    return "Daggers"
    if "knife" in s or re.search(r"\bknive\b", s): return "Knives"
    if "axe" in s or re.search(r"\bax\b", s):      return "Axes"
    if "sickle" in s:    return "Sickles"

    if ("bronze" in s and "sheet" in s) or "sheet bronze" in s:
        return "Bronze Sheets"

    # Jewellery (bracelets, necklaces, others)
    if any(k in s for k in ["anklet", "armring", "armband", "fussberg", "fussring", "beinberg", "beinring", "bracelet"]):
        return "Bracelets"
    if any(k in s for k in ["necklace", "torque", "torc", "neckring"]):
        return "Necklaces"
    if any(re.search(p, s) for p in [r"\bring\b", r"\bfinger ring\b", r"\bhair ring\b", r"\bpendant\b", r"\bpin\b", r"ring finger ring"]):
        return "Other Jewels"

    # fallback
    return "Other"

df = pd.read_excel(FILENAME, sheet_name=SHEET)

COL_OR     = col_like(df, COL_OR)
COL_STATE  = col_like(df, COL_STATE)
COL_REGION = col_like(df, COL_REGION)
COL_TYPE   = col_like(df, COL_TYPE)
COL_COMP   = col_like(df, COL_COMP)
COL_FRAG   = col_like(df, COL_FRAG)
COL_MATCH  = col_like(df, COL_MATCH)

df_std = pd.DataFrame({
    "Or_fin":   to_str_u(df[COL_OR]).str.replace(r"\s+", "_", regex=True),
    "State":    to_str_u(df[COL_STATE]),
    "Region":   to_str_u(df[COL_REGION]),
    "Artefact": df[COL_TYPE].astype(str).str.strip(),
    "Complete": as_int(df[COL_COMP]),
    "Fragment": as_int(df[COL_FRAG]),
    "Match":    as_int(df[COL_MATCH]) if COL_MATCH in df.columns else 0,
})

df_std = df_std[df_std["Or_fin"].isin(OR_KEEP)].copy()

df_std["Cat"] = df_std["Artefact"].apply(categorize)

# ===== 12 areas (as shown in the figure) =====
def R(s): return s.upper().strip()

AREAS = [
    dict(title="MV/SH/NIE", mask=lambda d: (d["State"].eq(R("GERMANY")) &
                                            d["Region"].isin({R("MECKLENBURG-VORPOMMERN"), R("SCHLESWIG-HOLSTEIN"), R("NIEDERSACHSEN")}))),
    dict(title="Brandenburg",    mask=lambda d: d["Region"].eq(R("BRANDENBURG"))),
    dict(title="Poland",         mask=lambda d: d["State"].eq(R("POLAND"))),
    dict(title="Thüringen/Hessen/Rheinland-Pfalz",
         mask=lambda d: (d["State"].eq(R("GERMANY")) &
                         d["Region"].isin({R("THÜRINGEN"), R("HESSEN"), R("RHEINLAND-PFALZ")}))),
    dict(title="Sachsen-Anhalt", mask=lambda d: d["Region"].eq(R("SACHSEN-ANHALT"))),
    dict(title="Sachsen",        mask=lambda d: d["Region"].eq(R("SACHSEN"))),
    dict(title="Baden-Württemberg", mask=lambda d: d["Region"].eq(R("BADEN-WÜRTTEMBERG"))),
    dict(title="Bayern",         mask=lambda d: d["Region"].eq(R("BAYERN"))),
    dict(title="Austria",        mask=lambda d: d["State"].eq(R("AUSTRIA"))),
    dict(title="Switzerland",    mask=lambda d: d["State"].eq(R("SWITZERLAND"))),
    dict(title="Italy",          mask=lambda d: d["State"].isin({R("ITALY"), R("SAN MARINO")})),
    dict(title="Slovenia",       mask=lambda d: d["State"].eq(R("SLOVENIA"))),
]

# ===== Aggregation per area =====
def aggregate_area(df_area: pd.DataFrame):
    """Ritorna due array allineati a CATS:
       - baseC = somma Complete
       - topF  = Fragment + Matching (solo dove Complete < 1)
    """
    comp = []
    fragmix = []
    for cat in CATS:
        sub = df_area[df_area["Cat"] == cat]
        c = int(sub["Complete"].sum())
        f = int(sub["Fragment"].sum()) + int(sub.loc[sub["Complete"] < 1, "Match"].sum())
        comp.append(c)
        fragmix.append(f)
    return np.array(comp), np.array(fragmix)

# ===== Plot 4x3 =====
fig, axes = plt.subplots(4, 3, figsize=FIGSIZE, dpi=DPI)
axes = axes.ravel()

for ax, area in zip(axes, AREAS):
    df_area = df_std[area["mask"](df_std)].copy()
    baseC, topF = aggregate_area(df_area)

    x = np.arange(len(CATS))
    # barre stacked
    ax.bar(x, baseC, BAR_W, color=COL_COMPLETE, zorder=2, **EDGE_KW, label="Complete")
    ax.bar(x, topF,  BAR_W, bottom=baseC, color=COL_FRAGMIX, zorder=3, **EDGE_KW, label="Fragmented")

    # headroom
    y_max = max((baseC + topF).max(), 1)
    ax.set_ylim(0, np.ceil(y_max * 1.30))

    for g, (i0, i1) in grp_idx.items():
        x0, x1 = i0 - 0.5, i1 + 0.5
        rect = Rectangle((x0, 0), x1 - x0, ax.get_ylim()[1],
                         fill=False, clip_on=False, zorder=1, **grp_style[g])
        ax.add_patch(rect)

    # axes
    ax.set_title(area["title"], fontsize=10, pad=4)
    ax.set_xticks(x)
    ax.set_xticklabels(CATS, rotation=55, ha="right", rotation_mode="anchor", fontsize=6)
    ax.tick_params(axis="x", pad=5)
    ax.tick_params(axis="y", labelsize=8)
    ax.yaxis.grid(True, which="major", color="0.90", linestyle="-", linewidth=0.7)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

# hide any extra axes (if AREAS < 12)
for ax in axes[len(AREAS):]:
    ax.axis("off")

# spaces
fig.subplots_adjust(hspace=0.60, wspace=0.35)

# legends
bar_handles = [Line2D([0], [0], lw=8, color=COL_COMPLETE),
               Line2D([0], [0], lw=8, color=COL_FRAGMIX)]
fig.legend(bar_handles, ["Complete", "Fragmented"],
           loc="lower center", bbox_to_anchor=(0.5, 0.045), ncol=2,
           fontsize=10, frameon=False)

grp_handles = [Line2D([0], [0], **grp_style[g]) for g in ["Weapons", "Tools", "Jewels", "Ingots"]]
fig.legend(grp_handles, ["Weapons", "Tools", "Jewels", "Ingots"],
           loc="lower center", bbox_to_anchor=(0.5, 0.01), ncol=4,
           fontsize=9, frameon=False)

plt.tight_layout(rect=[0, 0.08, 1, 1])
#fig.savefig("Fig16_grid_3x4_regions_stacked_simplified.jpg",
#            dpi=DPI, format="jpeg", pil_kwargs={"quality": 95, "subsampling": 0})
plt.show()


# %%

# ============================================================
# FIGURE 17 — Quantification of object fragmentation/integrity
# in the Late Bronze Age across areas A, B, C (Horizons 3-5).
# Three stacked panels (Area C, B, A), full page, with dashed
# rectangles marking functional families (Weapons, Tools,
# Jewels, Ingots).
# ============================================================

import pandas as pd, numpy as np, re, unicodedata
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib.patches import Rectangle

# ---------- EDITORIAL PARAMETERS ----------
FIGSIZE = (6.30, 6.94)   # approx. (160 mm, 176.25 mm) in inches
DPI     = 300
COL_COMPLETE = "#000000"  # black
COL_FRAGMIX  = "#CFCFCF"  # light gray
OR_KEEP = {"OR_3","OR_4","OR_5"}

# ---------- UTILS ----------
def col_like(df, name):
    low = {c.lower().strip(): c for c in df.columns}
    alts = {"artefact":["artefact","artifact","type","object","category","class","item","typology"],
            "fragmented":["fragmented","fragment"],
            "matching fr":["matching fr","matching_fr","matchingfr","matching"]}
    k = name.lower().strip()
    if k in alts:
        for cand in alts[k]:
            if cand in low: return low[cand]
    return low.get(k, name)

def to_str_u(s):  # upper + strip
    return s.astype(str).str.upper().str.strip()

def as_int(s):
    return pd.to_numeric(s, errors="coerce").fillna(0).astype(int)

def strip_accents(s):
    s = s.replace("ß","ss")
    return "".join(ch for ch in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(ch))

def norm(s):
    s = strip_accents(str(s)).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

# Categorie (plurale)
CATS = [
    "Spearheads","Swords","Daggers","Knives",
    "Axes","Sickles",
    "Bracelets","Necklaces","Other Jewels",
    "Ingots","Bronze Sheets","Golden obj.","Other"
]

# === NEW === dashed boxes for functional families (indices on CATS)
# 0:Spearheads 1:Swords 2:Daggers 3:Knives 4:Axes 5:Sickles
# 6:Bracelets 7:Necklaces 8:Other Jewels 9:Ingots 10:Bronze Sheets 11:Golden obj. 12:Other
grp_idx = {
    "Weapons": (0, 2),   # Spearheads-Swords-Daggers (Knives excluded)
    "Tools":   (4, 5),   # Axes-Sickles
    "Jewels":  (6, 8),   # Bracelets-Necklaces-Other Jewels
    "Ingots":  (9, 9),   # solo Ingots
}
DASH = (0, (6, 3))  # uniform dash pattern for all
grp_style = {
    "Weapons": dict(color="#1f77b4", ls=DASH, lw=0.9),  # blu
    "Tools":   dict(color="#ff7f0e", ls=DASH, lw=0.9),  # orange
    "Jewels":  dict(color="#2ca02c", ls=DASH, lw=0.9),  # green
    "Ingots":  dict(color="#9467bd", ls=DASH, lw=0.9),  # purple
}
# === /NEW ===

def categorize(artefact_raw):
    s = norm(artefact_raw)
    # Ingots checked before Gold...
    if re.search(r"\b(ingot|ingots|bun[- ]?ingot|bun|barren|rohbarren|spitzbarren|plattbarren|scheibenbarren)\b", s):
        return "Ingots"
    if s.startswith("gold"):
        return "Golden obj."
    if "spearhead" in s:   return "Spearheads"
    if "sword" in s:       return "Swords"
    if "dagger" in s:      return "Daggers"
    if "knife" in s or "knive" in s:  return "Knives"
    if "axe" in s or re.search(r"\bax\b", s): return "Axes"
    if "sickle" in s:      return "Sickles"
    if ("bronze" in s and "sheet" in s) or "sheet bronze" in s: return "Bronze Sheets"
    if ("anklet" in s or "armring" in s or "armband" in s or
        "fussberg" in s or "fussring" in s or "beinberg" in s or "beinring" in s or
        "bracelet" in s or "bracelet anklet" in s):
        return "Bracelets"
    if ("necklace" in s or "torque" in s or "torc" in s or "neckring" in s):
        return "Necklaces"
    if ("ring" in s or "finger ring" in s or "hair ring" in s or
        "pendant" in s or re.search(r"\bpin\b", s) or "ring finger ring" in s):
        return "Other Jewels"
    return "Other"

def counts_by_category(df, cats_order=CATS):
    comp, fragmix = [], []
    for cat in cats_order:
        sub = df[df["Cat"] == cat]
        c = int(sub["Complete"].sum())
        f = int(sub["Fragment"].sum()) + int(sub.loc[sub["Complete"] < 1, "Match"].sum())
        comp.append(c); fragmix.append(f)
    return np.array(comp), np.array(fragmix)

# ---------- READ & FILTERS ----------
FILENAME = "DB.xlsx"; SHEET = "DB"
df_raw = pd.read_excel(FILENAME, sheet_name=SHEET)

c_or    = col_like(df_raw, "Or_fin")
c_state = col_like(df_raw, "State")
c_reg   = col_like(df_raw, "Region")
c_type  = col_like(df_raw, "Artefact")
c_comp  = col_like(df_raw, "Complete")
c_frag  = col_like(df_raw, "Fragmented")
c_match = col_like(df_raw, "Matching fr")
c_ns    = col_like(df_raw, "N/S Chrono") if "N/S Chrono" in df_raw.columns or "n/s chrono" in {x.lower() for x in df_raw.columns} else None

df = pd.DataFrame({
    "Or_fin":   to_str_u(df_raw[c_or]).str.replace(r"\s+","_", regex=True),
    "State":    to_str_u(df_raw[c_state]),
    "Region":   to_str_u(df_raw[c_reg]),
    "Artefact": df_raw[c_type].astype(str).str.strip(),
    "Complete": as_int(df_raw[c_comp]),
    "Fragment": as_int(df_raw[c_frag]),
    "Match":    as_int(df_raw[c_match]),
})
# ignore N/S Chrono = IC
if c_ns is not None:
    ns = to_str_u(df_raw[c_ns])
    df = df[~ns.eq("IC")].copy()

# period
df = df[df["Or_fin"].isin(OR_KEEP)].copy()

# categorize
df["Cat"] = df["Artefact"].apply(categorize)

# ---------- areas (A/B/C) ----------
U = lambda s: s.upper().strip()
GER = U("GERMANY")

mask_C = (
    (df["State"].eq(GER) & df["Region"].isin({U("MECKLENBURG-VORPOMMERN"), U("SCHLESWIG-HOLSTEIN"), U("NIEDERSACHSEN")})) |
    (df["State"].eq(GER) & df["Region"].isin({U("THÜRINGEN"), U("HESSEN"), U("RHEINLAND-PFALZ")})) |
    (df["Region"].eq(U("SACHSEN-ANHALT")))
)

mask_B = (
    df["Region"].eq(U("BRANDENBURG")) |
    df["State"].eq(U("POLAND")) |
    df["Region"].eq(U("SACHSEN"))
)

mask_A = (
    df["Region"].isin({U("BADEN-WÜRTTEMBERG"), U("BAYERN")}) |
    df["State"].isin({U("AUSTRIA"), U("ITALY"), U("SAN MARINO"), U("SLOVENIA")})
)
# NB: Switzerland excluded from A as requested

areas = [
    ("Area C", df[mask_C].copy()),
    ("Area B", df[mask_B].copy()),
    ("Area A", df[mask_A].copy()),
]

# ---------- PLOT (3 panels stacked) ----------
fig, axes = plt.subplots(
    3, 1, figsize=FIGSIZE, dpi=DPI, sharex=True
)
axes = np.atleast_1d(axes)

x = np.arange(len(CATS))
bar_w = 0.70
edge_kw = dict(edgecolor="white", linewidth=0.6)

for ax, (title, dfa) in zip(axes, areas):   # <-- 'areas' is the same A/B/C list created above
    C, F = counts_by_category(dfa, CATS)
    ax.bar(x, C, bar_w, color=COL_COMPLETE, **edge_kw, label="Complete", zorder=2)
    ax.bar(x, F, bar_w, bottom=C, color=COL_FRAGMIX, **edge_kw, label="Fragmented", zorder=3)

    y_max = max((C+F).max(), 1)
    ax.set_ylim(0, np.ceil(y_max*1.25))
    ax.set_ylabel("Counts", fontsize=9)
    ax.set_title(title, fontsize=11, pad=6)
    ax.yaxis.grid(True, color="0.9", lw=0.7)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)

    # === NEW === dashed boxes for functional families (per panel)
    for g, (i0, i1) in grp_idx.items():
        x0, x1 = i0 - 0.5, i1 + 0.5
        rect = Rectangle((x0, 0), x1 - x0, ax.get_ylim()[1],
                         fill=False, clip_on=False, zorder=1, **grp_style[g])
        ax.add_patch(rect)
    # === /NEW ===

axes[-1].set_xticks(x)
axes[-1].set_xticklabels(CATS, rotation=55, ha="right", fontsize=8)

# === NEW === two legends at the bottom (bars + groups)
# bar legend (black/gray)
bar_handles = [Line2D([0],[0], lw=8, color=COL_COMPLETE),
               Line2D([0],[0], lw=8, color=COL_FRAGMIX)]
fig.legend(bar_handles, ["Complete","Fragmented"],
           loc="lower center", bbox_to_anchor=(0.5, 0.04),
           ncol=2, fontsize=10, frameon=False)

# legend with dotted groups
grp_order = ["Weapons","Tools","Jewels","Ingots"]
grp_handles = [Line2D([0],[0], lw=2, ls=grp_style[g]["ls"], color=grp_style[g]["color"])
               for g in grp_order]
fig.legend(grp_handles, grp_order,
           loc="lower center", bbox_to_anchor=(0.5, 0.09),
           ncol=4, fontsize=9, frameon=False)

plt.tight_layout(rect=[0, 0.13, 1, 1])  # extra room for the two legends
# === /NEW ===

#fig.savefig("Fig17_Areas_ABC_fullpage.jpg", dpi=DPI, format="jpeg",
#            pil_kwargs={"quality":95, "subsampling":0})
plt.show()


# %%

# ============================================================
# FIGURE 20 — Quantification of complete, fragmented and 'damaged'
# objects in the examined areas and periods. 2x2 grid (Area A/B/C/CH),
# one panel per area, four chronological periods on the x axis.
# ============================================================

import pandas as pd, numpy as np, re, unicodedata
import matplotlib.pyplot as plt

# ============ EDITORIAL PARAMETERS ============
FIGSIZE = (6.30, 4.33)   # ~160 x 110 mm (half page)
DPI     = 300

# Palette (consistent with earlier charts)
COL_COMPLETE   = "#000000"  # black
COL_FRAGMENT   = "#CFCFCF"  # light gray
COL_MATCH_COMP = "#2e7d6e"  # muted green
COL_MATCH_FRAG = "#4c7da8"  # blu smorzato
COL_BENT_COMP  = "#a85d4d"  # marrone/ruggine
COL_BENT_FRAG  = "#d08c00"  # ocra

BAR_WIDTH   = 0.12          # single bar width
EDGE_KW     = dict(edgecolor="white", linewidth=0.6)

# ============ UTILS ============
def col_like(df, name):
    low = {c.lower().strip(): c for c in df.columns}
    alts = {"fragmented":["fragmented","fragment"],
            "matching fr":["matching fr","matching_fr","matchingfr","matching"]}
    k = name.lower().strip()
    if k in alts:
        for cand in alts[k]:
            if cand in low: return low[cand]
    return low.get(k, name)

def to_str_u(s):  # UPPER + strip
    return s.astype(str).str.upper().str.strip()

def as_int(s):
    return pd.to_numeric(s, errors="coerce").fillna(0).astype(int)

# ============ READING ============
FILENAME, SHEET = "DB.xlsx", "DB"
raw = pd.read_excel(FILENAME, sheet_name=SHEET)

c_or    = col_like(raw, "Or_fin")
c_state = col_like(raw, "State")
c_reg   = col_like(raw, "Region")
c_comp  = col_like(raw, "Complete")
c_frag  = col_like(raw, "Fragmented")
c_match = col_like(raw, "Matching fr")
c_belt  = col_like(raw, "Belted")
c_ns    = col_like(raw, "N/S Chrono") if any(x.lower()=="n/s chrono" for x in raw.columns.str.lower()) else None

df = pd.DataFrame({
    "Or_fin":   to_str_u(raw[c_or]).str.replace(r"\s+","_", regex=True),
    "State":    to_str_u(raw[c_state]),
    "Region":   to_str_u(raw[c_reg]),
    "Complete": as_int(raw[c_comp]),
    "Fragment": as_int(raw[c_frag]),
    "Match":    as_int(raw[c_match]),
    "Belted":   as_int(raw[c_belt]),
})
# exclude N/S Chrono = IC if present
if c_ns is not None:
    ns = to_str_u(raw[c_ns])
    df = df[~ns.eq("IC")].copy()

# ============ areas (A/B/C/CH) ============
U = lambda s: s.upper().strip()
GER = U("GERMANY")

mask_C = (
    (df["State"].eq(GER) & df["Region"].isin({U("MECKLENBURG-VORPOMMERN"), U("SCHLESWIG-HOLSTEIN"), U("NIEDERSACHSEN")})) |
    (df["State"].eq(GER) & df["Region"].isin({U("THÜRINGEN"), U("HESSEN"), U("RHEINLAND-PFALZ")})) |
    (df["Region"].eq(U("SACHSEN-ANHALT")))
)
mask_B = (
    df["Region"].eq(U("BRANDENBURG")) |
    df["State"].eq(U("POLAND")) |
    df["Region"].eq(U("SACHSEN"))
)
mask_A = (
    df["Region"].isin({U("BADEN-WÜRTTEMBERG"), U("BAYERN")}) |
    df["State"].isin({U("AUSTRIA"), U("SLOVENIA")})  # <-- Italy and San Marino excluded
)
mask_CH = df["State"].eq(U("SWITZERLAND"))

AREAS = [
    ("Area A", df[mask_A].copy()),
    ("Area B", df[mask_B].copy()),
    ("Area C", df[mask_C].copy()),
    ("Switzerland", df[mask_CH].copy()),
]

# ============ periods ============
PERIODS = [
    ("Or. 1",   {"OR_1"}),
    ("Or. 2",  {"OR_2"}),
    ("Or. 3", {"OR_3"}),
    ("Or. 4–5",{"OR_4","OR_5"}),
]

def agg_period(sub):
    """Returns a tuple (C, F, Mcomp, Mfrag, Cbent, Fbent) for a sub-dataframe already filtered by area+period."""
    C     = int(sub["Complete"].sum())
    F     = int(sub["Fragment"].sum())
    Mcomp = int(((sub["Complete"] > 0) & (sub["Match"] > 0)).sum())
    Mfrag = int(((sub["Complete"] < 1) & (sub["Match"] > 0)).sum())
    Cbent = int(sub.loc[sub["Belted"] > 0, "Complete"].sum())
    Fbent = int(sub.loc[sub["Belted"] > 0, "Fragment"].sum())
    return C, F, Mcomp, Mfrag, Cbent, Fbent

# ============ PLOT 2x2 ============
fig, axes = plt.subplots(2, 2, figsize=FIGSIZE, dpi=DPI, sharex=True)
axes = axes.ravel()

n_bars = 6
x = np.arange(len(PERIODS))
group_span = BAR_WIDTH * n_bars * 1.15
offsets = np.linspace(-group_span/2 + BAR_WIDTH/2, group_span/2 - BAR_WIDTH/2, n_bars)

for ax, (title, dfa) in zip(axes, AREAS):
    # arrays for the 6 series
    C_list, F_list, MC_list, MF_list, CB_list, FB_list = [], [], [], [], [], []
    for _, pset in PERIODS:
        sub = dfa[dfa["Or_fin"].isin(pset)]
        C, F, Mcomp, Mfrag, Cbent, Fbent = agg_period(sub)
        C_list.append(C); F_list.append(F)
        MC_list.append(Mcomp); MF_list.append(Mfrag)
        CB_list.append(Cbent); FB_list.append(Fbent)

    # draw bars (same order for all panels)
    ax.bar(x + offsets[0], C_list,  BAR_WIDTH, color=COL_COMPLETE,   **EDGE_KW, label="Complete",            zorder=3)
    ax.bar(x + offsets[1], F_list,  BAR_WIDTH, color=COL_FRAGMENT,   **EDGE_KW, label="Fragmented",          zorder=2)
    ax.bar(x + offsets[2], MC_list, BAR_WIDTH, color=COL_MATCH_COMP, **EDGE_KW, label="Matching fr. (comp)", zorder=3)
    ax.bar(x + offsets[3], MF_list, BAR_WIDTH, color=COL_MATCH_FRAG, **EDGE_KW, label="Matching fr. (fragm.)", zorder=3)
    ax.bar(x + offsets[4], CB_list, BAR_WIDTH, color=COL_BENT_COMP,  **EDGE_KW, label="Complete bent",       zorder=3)
    ax.bar(x + offsets[5], FB_list, BAR_WIDTH, color=COL_BENT_FRAG,  **EDGE_KW, label="Fragmented bent",     zorder=3)

    # style
    ymax = max([max(arr) if arr else 0 for arr in [C_list,F_list,MC_list,MF_list,CB_list,FB_list]] + [1])
    ax.set_ylim(0, np.ceil(ymax*1.25))
    ax.set_title(title, fontsize=11, pad=4)
    ax.set_ylabel("Counts", fontsize=9)
    ax.yaxis.grid(True, color="0.90", lw=0.7)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)

# period labels on the x axis (bottom panels)
axes[2].set_xticks(x); axes[3].set_xticks(x)
period_labels = [lab for lab, _ in PERIODS]
axes[2].set_xticklabels(period_labels, fontsize=9)
axes[3].set_xticklabels(period_labels, fontsize=9)

# --- more spacing between X labels and legend ---
for ax in axes:
    ax.tick_params(axis='x', pad=10)   # distance from labels to axis

handles = [
    plt.Line2D([0],[0], lw=8, color=COL_COMPLETE),
    plt.Line2D([0],[0], lw=8, color=COL_FRAGMENT),
    plt.Line2D([0],[0], lw=8, color=COL_MATCH_COMP),
    plt.Line2D([0],[0], lw=8, color=COL_MATCH_FRAG),
    plt.Line2D([0],[0], lw=8, color=COL_BENT_COMP),
    plt.Line2D([0],[0], lw=8, color=COL_BENT_FRAG),
]
labels  = ["Complete","Fragmented","Matching fr. (comp)",
           "Matching fr. (fragm.)","Complete bent","Fragmented bent"]

# Legend lower down, with ample margin above
fig.legend(handles, labels,
           loc="lower center", bbox_to_anchor=(0.5, 0.03),
           ncol=3, fontsize=9, frameon=False)

# Reserve more room at the bottom for the legend (hence more gap from the labels)
plt.tight_layout(rect=[0, 0.18, 1, 1])   # increase if you want even more room

#fig.savefig("Fig20_Area_ABC_CH_periods_2x2_halfpage.jpg", dpi=DPI, format="jpeg",
#            pil_kwargs={"quality":95, "subsampling":0})
plt.show()


# %%

# ============================================================
# FIGURE 21 — Deformation of complete or fragmented objects
# Bent ('Belted') complete vs fragmented counts, by artefact type.
# ============================================================

import pandas as pd, numpy as np, re, unicodedata
import matplotlib.pyplot as plt

# ------------------- EDITORIAL PARAMETERS -------------------
FIGSIZE = (6.30, 4.33)   # ~160 x 110 mm (half page)
DPI = 300
COL_COMPLETE = "#000000"   # black (Complete belted)
COL_FRAGMIX  = "#CFCFCF"   # gray (Fragmented belted)
BAR_W = 0.38
EDGE_KW = dict(edgecolor="white", linewidth=0.6)

# ------------------- Utils base -------------------
def col_like(df, name):
    low = {c.lower().strip(): c for c in df.columns}
    alts = {"artefact":["artefact","artifact","type","object","category","class","item","typology"],
            "fragmented":["fragmented","fragment"],
            "matching fr":["matching fr","matching_fr","matchingfr","matching"]}
    k = name.lower().strip()
    if k in alts:
        for cand in alts[k]:
            if cand in low: return low[cand]
    return low.get(k, name)

def to_str_u(s):  # normalize string to UPPER & strip
    return s.astype(str).str.upper().str.strip()

def as_int(s):
    return pd.to_numeric(s, errors="coerce").fillna(0).astype(int)

def strip_accents(s):
    s = s.replace("ß","ss")
    return "".join(ch for ch in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(ch))

def norm(s):
    s = strip_accents(str(s)).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

# ------------------- Categories and mapping -------------------
ORDER = [
    "Spearheads","Swords","Daggers","Greaves","Knives",
    "Axes","Sickles","Chisels",
    "Bracelets","Necklaces","Rings",
    "Pendants","Pins","Fibulae",
    "Ringbarren","Spangenbarren",   # <- new, after Fibulae
    "Bronze Sheets","Wires","Others"
]

def cat_belted(artefact_raw):
    s = norm(artefact_raw)

    # --- new categories 
    if "ringbarren" in s:
        return "Ringbarren"
    if "spangenbarren" in s:
        return "Spangenbarren"

    # Bronze sheets
    if ("bronze" in s and "sheet" in s) or "sheet bronze" in s:
        return "Bronze Sheets"

    # Jewellery (as agreed)
    if ("anklet" in s or "armring" in s or "armband" in s or
        "fussberg" in s or "fussring" in s or "beinberg" in s or "beinring" in s or
        "bracelet" in s or "bracelet anklet" in s):
        return "Bracelets"
    if ("necklace" in s or "torque" in s or "torc" in s or "neckring" in s):
        return "Necklaces"
    if ("ring" in s or "finger ring" in s or "hair ring" in s):
        return "Rings"

    # Pendants / Pins / Fibulae
    if "pendant" in s:          return "Pendants"
    if re.search(r"\bpin\b", s): return "Pins"
    if "fibula" in s:           return "Fibulae"

    # Weapons and tools
    if "spearhead" in s:        return "Spearheads"
    if "sword" in s:            return "Swords"
    if "dagger" in s:           return "Daggers"
    if "greave" in s:           return "Greaves"
    if "knife" in s or "knive" in s:  return "Knives"
    if "axe" in s or re.search(r"\bax\b", s): return "Axes"
    if "sickle" in s:           return "Sickles"

    # Chisels = chisel + awl
    if "chisel" in s or "awl" in s:   return "Chisels"

    # Wires
    if "wire" in s:             return "Wires"

    # All generic "ingot" entries now fall into Others
    if "ingot" in s or "barren" in s:
        return "Others"

    # Other
    return "Others"


# ------------------- READ DATA -------------------
FILENAME = "DB.xlsx"; SHEET = "DB"
df_raw = pd.read_excel(FILENAME, sheet_name=SHEET)

c_state = col_like(df_raw, "State")
c_type  = col_like(df_raw, "Artefact")
c_comp  = col_like(df_raw, "Complete")
c_frag  = col_like(df_raw, "Fragmented")
c_belt  = col_like(df_raw, "Belted")  # needed here

df = pd.DataFrame({
    "State":    to_str_u(df_raw[c_state]),
    "Artefact": df_raw[c_type].astype(str).str.strip(),
    "Complete": as_int(df_raw[c_comp]),
    "Fragment": as_int(df_raw[c_frag]),
    "Belted":   as_int(df_raw[c_belt]),
})

# Ignore Italy and San Marino
df = df[~df["State"].isin({"ITALY","SAN MARINO"})].copy()

# Category for this chart
df["Cat"] = df["Artefact"].apply(cat_belted)

# ------------------- Aggregation: only BELTED -------------------
rows = []
for cat in ORDER:
    sub = df[df["Cat"] == cat]
    # Complete belted: sum of "Complete" where Belted > 0
    c_belt = int(sub.loc[sub["Belted"] > 0, "Complete"].sum())
    # Fragmented belted: sum of "Fragment" where Belted > 0
    f_belt = int(sub.loc[sub["Belted"] > 0, "Fragment"].sum())
    rows.append((cat, c_belt, f_belt))

agg = pd.DataFrame(rows, columns=["Cat","CompleteBelted","FragmentedBelted"])

# ------------------- Plot -------------------
x = np.arange(len(ORDER))
fig, ax = plt.subplots(figsize=FIGSIZE, dpi=DPI)

ax.bar(x - BAR_W/2, agg["CompleteBelted"], BAR_W,
       color=COL_COMPLETE, **EDGE_KW, label="Complete (bent)", zorder=3)
ax.bar(x + BAR_W/2, agg["FragmentedBelted"], BAR_W,
       color=COL_FRAGMIX, **EDGE_KW, label="Fragmented (bent)", zorder=2)

# Headroom
ymax = max(agg["CompleteBelted"].max(), agg["FragmentedBelted"].max(), 1)
ax.set_ylim(0, np.ceil(ymax * 1.25))

# Axes & style
ax.set_xticks(x)
ax.set_xticklabels(ORDER, rotation=45, ha="right", fontsize=8)
ax.set_ylabel("Counts", fontsize=9)
ax.yaxis.grid(True, color="0.9", lw=0.7)
ax.set_axisbelow(True)
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)

# Legend
ax.legend(loc="upper left", fontsize=9, frameon=False, ncol=1)

plt.tight_layout()
#fig.savefig("Fig21_Belted_by_type_halfpage.jpg", dpi=DPI, format="jpeg",
#            pil_kwargs={"quality":95, "subsampling":0})
plt.show()


# %%

# ============================================================
# FIGURE 22 — Estimation of the preserved portion of the object
# compared to the intact object, and its quantification. 4x4 grid
# (areas x periods) showing the preserved-portion classes (QP column).
# ============================================================

import pandas as pd, numpy as np, re, unicodedata
import matplotlib.pyplot as plt
from matplotlib.patches import Patch


# =============== PARAMETERS ===============
FIGSIZE = (6.30, 6.94)   # ~160 x 176 mm (3/4 page)
DPI     = 300
BAR_W   = 0.65

COL_Q1 = "#000000"   # < 1/3
COL_Q2 = "#9e9e9e"   # 1/3-2/3
COL_Q3 = "#FFFFFF"   # 2/3-1
EDGE_KW_ALL = dict(edgecolor="#000000", linewidth=0.5)
LAB_Q1, LAB_Q2, LAB_Q3 = "<⅓", "⅓–⅔", "⅔–<1" 


# =============== READING ===============
FILENAME, SHEET = "DB.xlsx", "DB"
raw = pd.read_excel(FILENAME, sheet_name=SHEET)
c_or    = col_like(raw, "Or_fin")
c_state = col_like(raw, "State")
c_reg   = col_like(raw, "Region")
c_type  = col_like(raw, "Artefact")
c_qp    = col_like(raw, "Q P")

df = pd.DataFrame({
    "Or_fin":   to_str_u(raw[c_or]).str.replace(r"\s+","_", regex=True),
    "State":    to_str_u(raw[c_state]),
    "Region":   to_str_u(raw[c_reg]),
    "Artefact": raw[c_type].astype(str).str.strip(),
    "QP":       as_int(raw[c_qp]),
})

# =============== areas (rows) ===============
U = lambda s: s.upper().strip()
GER = U("GERMANY")
mask_C = (
    (df["State"].eq(GER) & df["Region"].isin({U("MECKLENBURG-VORPOMMERN"), U("SCHLESWIG-HOLSTEIN"), U("NIEDERSACHSEN")})) |
    (df["State"].eq(GER) & df["Region"].isin({U("THÜRINGEN"), U("HESSEN"), U("RHEINLAND-PFALZ")})) |
    (df["Region"].eq(U("SACHSEN-ANHALT")))
)
mask_B = (
    df["Region"].eq(U("BRANDENBURG")) |
    df["State"].eq(U("POLAND")) |
    df["Region"].eq(U("SACHSEN"))
)
mask_A = (
    df["Region"].isin({U("BADEN-WÜRTTEMBERG"), U("BAYERN")}) |
    df["State"].isin({U("AUSTRIA"), U("SLOVENIA")})
)
mask_I = df["State"].isin({U("ITALY"), U("SAN MARINO")})

ROW_AREAS = [
    ("Area A", df[mask_A].copy()),
    ("Area B", df[mask_B].copy()),
    ("Area C", df[mask_C].copy()),
    ("Italy",  df[mask_I].copy()),
]

# =============== columns (periods e categorie) ===============
COL_PERIODS = [
    ("Or. 1",  "OR_1",   ["Ringbarren","Spangenbarren","Axe","Dagger","Bracelet","Pin"]),
    ("Or. 2",  "OR_2",   ["Axe","Sickle","Bracelet","Pin","Ingot"]),
    ("Or. 3",  "OR_3",   ["Spearhead","Sword","Axe","Sickle","Bracelet","Ingot"]),
    ("Or. 4–5","OR_4_5", ["Spearhead","Sword","Axe","Sickle","Bracelet","Ingot"]),
]

PLURAL_LABEL = {
    "Ringbarren": "Ringbarren",    
    "Spangenbarren": "Spangenbarren",  
    "Bracelet": "Bracelets",        # already in the desired plural form
    "Axe": "Axes",
    "Dagger": "Daggers",
    "Spearhead": "Spearheads",
    "Sword": "Swords",
    "Sickle": "Sickles",
    "Pin": "Pins",
    "Ingot": "Ingots",
}
def label_plural(cat: str) -> str:
    return PLURAL_LABEL.get(cat, cat + "s")


# --- helper to normalize and classify by period ---
import re, unicodedata
import numpy as np

def _strip_accents(s):
    s = str(s).replace("ß", "ss")
    return "".join(ch for ch in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(ch))

def _norm(s):
    s = _strip_accents(s).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def _is_bracelet(s_norm):
    # Bracelets = Anklet/Armring/Armband/Beinberg/Fussring/Bracelet...
    return any(w in s_norm for w in [
        "bracelet", "anklet", "armring", "armband",
        "beinberg", "beinring", "fussberg", "fussring",
        "bracelet anklet", "anklet armring"
    ])

def map_cat_for_period(artefact_raw, period_key):
    s = _norm(artefact_raw)

    if period_key == "OR_1":
        if "ringbarren"     in s: return "Ringbarren"
        if "spangenbarren"  in s: return "Spangenbarren"
        if "dagger"         in s: return "Dagger"
        if "axe" in s or re.search(r"\bax\b", s): return "Axe"
        if _is_bracelet(s): return "Bracelet"
        if re.search(r"\bpin\b", s): return "Pin"
        return np.nan

    if period_key == "OR_2":
        if "sickle"         in s: return "Sickle"
        if "axe" in s or re.search(r"\bax\b", s): return "Axe"
        if _is_bracelet(s): return "Bracelet"
        if re.search(r"\bpin\b", s): return "Pin"
        if "ingot"          in s: return "Ingot"
        return np.nan

    if period_key == "OR_3":
        if "spearhead"      in s: return "Spearhead"
        if "sword"          in s: return "Sword"
        if "sickle"         in s: return "Sickle"
        if "axe" in s or re.search(r"\bax\b", s): return "Axe"
        if _is_bracelet(s): return "Bracelet"
        if "ingot"          in s: return "Ingot"
        return np.nan

    if period_key == "OR_4_5":
        if "spearhead"      in s: return "Spearhead"
        if "sword"          in s: return "Sword"
        if "sickle"         in s: return "Sickle"
        if "axe" in s or re.search(r"\bax\b", s): return "Axe"
        if _is_bracelet(s): return "Bracelet"
        if "ingot"          in s: return "Ingot"
        return np.nan

    return np.nan



def counts_area_period(area_df, period_key, cats):
    sub = area_df.copy()
    sub = sub[sub["Or_fin"].isin({"OR_4","OR_5"})] if period_key=="OR_4_5" else sub[sub["Or_fin"].eq(period_key)]
    sub["Cat"] = sub["Artefact"].apply(lambda x: map_cat_for_period(x, period_key))
    sub = sub.dropna(subset=["Cat"])
    out_q1, out_q2, out_q3 = [], [], []
    for cat in cats:
        q1 = int((sub["Cat"].eq(cat) & (sub["QP"] == 1)).sum())
        q2 = int((sub["Cat"].eq(cat) & (sub["QP"] == 2)).sum())
        q3 = int((sub["Cat"].eq(cat) & (sub["QP"] == 3)).sum())
        out_q1.append(q1); out_q2.append(q2); out_q3.append(q3)
    return np.array(out_q1), np.array(out_q2), np.array(out_q3)


# =============== PLOT 4x4 (single figure) ===============
fig, axes = plt.subplots(4, 4, figsize=FIGSIZE, dpi=DPI)
axes = np.array(axes)

# Draw bars
for r, (row_label, area_df) in enumerate(ROW_AREAS):
    for c, (per_label, per_key, cats) in enumerate(COL_PERIODS):
        ax = axes[r, c]
        q1, q2, q3 = counts_area_period(area_df, per_key, cats)
        x = np.arange(len(cats))

        ax.bar(x, q1, BAR_W, color=COL_Q1, **EDGE_KW_ALL, zorder=3, label=LAB_Q1)
        ax.bar(x, q2, BAR_W, bottom=q1,    color=COL_Q2, **EDGE_KW_ALL, zorder=3, label=LAB_Q2)
        ax.bar(x, q3, BAR_W, bottom=q1+q2, color=COL_Q3, **EDGE_KW_ALL, zorder=3, label=LAB_Q3)


        ymax = max((q1+q2+q3).max(), 1)
        ax.set_ylim(0, np.ceil(ymax*1.30))
        ax.yaxis.grid(True, color="0.88", lw=0.7)
        ax.set_axisbelow(True)
        ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)

        ax.set_xticks(x)
        if r == len(ROW_AREAS) - 1:
            disp = [label_plural(c) for c in cats]
            ax.set_xticklabels(disp, rotation=45, ha="right", fontsize=7)
            ax.tick_params(axis='x', pad=10)
        else:
            ax.set_xticklabels([])
        ax.tick_params(axis='y', labelsize=8)


# --- layout BEFORE positioning the text labels ---
plt.tight_layout(pad=1.3, w_pad=1.0, h_pad=1.2, rect=[0.08, 0.10, 0.98, 0.95])
fig.canvas.draw()

# Column titles (periods), above and centered
TITLE_Y_OFFSET = 0.038
for c, (per_label, _, _) in enumerate(COL_PERIODS):
    pos = axes[0, c].get_position()
    fig.text(pos.x0 + pos.width/2, pos.y1 + TITLE_Y_OFFSET,
             per_label, ha="center", va="bottom", fontsize=11)

# Row labels (areas), left-aligned and vertically centered
ROWLABEL_X_OFFSET = 0.045
for r, (row_label, _) in enumerate(ROW_AREAS):
    row_boxes = [axes[r, i].get_position() for i in range(4)]
    y0, y1 = min(bb.y0 for bb in row_boxes), max(bb.y1 for bb in row_boxes)
    x = row_boxes[0].x0 - ROWLABEL_X_OFFSET
    y = (y0 + y1) / 2.0
    fig.text(x, y, row_label, ha="right", va="center", rotation=90, fontsize=11,
         bbox=dict(boxstyle="round,pad=0.15", fc="white", ec="none", alpha=0.75))


# Global legend with colored boxes and black border
handles = [
    Patch(facecolor=COL_Q1, edgecolor="black", linewidth=0.5, label=LAB_Q1),
    Patch(facecolor=COL_Q2, edgecolor="black", linewidth=0.5, label=LAB_Q2),
    Patch(facecolor=COL_Q3, edgecolor="black", linewidth=0.5, label=LAB_Q3),
]
fig.legend(
    handles, [LAB_Q1, LAB_Q2, LAB_Q3],
    loc="lower center", bbox_to_anchor=(0.5, 0.03),
    ncol=3, frameon=False, fontsize=9,
    handlelength=3.0,     # ← wider box
    handleheight=1.2,     # ← taller box
    handletextpad=0.6,    # distance box↔testo
    columnspacing=1.0,    # spacing between legend columns
    borderpad=0.3         # internal legend padding
)

#plt.savefig("Fig22_QP_areas_periods_4x4_labels_compact.jpg",
#            dpi=DPI, format="jpeg", bbox_inches="tight",
#            pil_kwargs={"quality":95, "subsampling":0})
plt.show()


# %%

# ============================================================
# FIGURES 27, 30, 33, 36 — Boxplots of total hoard weight by area
# One figure per chronological horizon (Or. 1, 2, 3, 4-5), with one
# boxplot per area (A, B, C, Switzerland, Italy) and individual data
# points shown as jittered markers.
# ============================================================

import pandas as pd, numpy as np, re, unicodedata
import matplotlib.pyplot as plt

# ---------- PARAMETERS ----------
FILENAME = "DB.xlsx"; SHEET = "DB"
FIGSIZE = (6.30, 4.33)   # half page
DPI = 300
JITTER_SD = 0.06
rng = np.random.default_rng(42)

# Box colors (one per period: blue, green, orange, purple - consistent and well distinguished)
PERIOD_COLORS = {
    "OR_1":   "#1f77b4",  # blue
    "OR_2":   "#2ca02c",  # green
    "OR_3":   "#ff7f0e",  # orange
    "OR_4_5": "#9467bd",  # purple
}
BOX_EDGE = "black"
BOX_ALPHA = 0.55     # slightly transparent fill

# ---------- UTILS ----------
def col_like(df, name):
    low = {c.lower().strip(): c for c in df.columns}
    alts = {
        "id_sito": ["id_sito","id","site_id","id sito","idsite"],
        "or_fin":  ["or_fin","or","period","orizzonte"],
        "state":   ["state","stato","country"],
        "region":  ["region","regione"],
        "tot w":   ["tot w","tot_w","total w","total_w","total weight","peso totale","totweight","weight"]
    }
    k = name.lower().strip()
    if k in alts:
        for cand in alts[k]:
            if cand in low: return low[cand]
    return low.get(k, name)

def to_str_u(s):  # UPPER + strip
    return s.astype(str).str.upper().str.strip()

def as_float(s):
    return pd.to_numeric(s, errors="coerce").astype(float)

# ---------- READING & NORMALIZATION ----------
raw = pd.read_excel(FILENAME, sheet_name=SHEET)

c_id   = col_like(raw, "ID_Sito")
c_or   = col_like(raw, "Or_fin")
c_st   = col_like(raw, "State")
c_reg  = col_like(raw, "Region")
c_totw = col_like(raw, "Tot W")

df0 = pd.DataFrame({
    "ID":     to_str_u(raw[c_id]),
    "Or_fin": to_str_u(raw[c_or]).str.replace(r"\s+","_", regex=True),
    "State":  to_str_u(raw[c_st]),
    "Region": to_str_u(raw[c_reg]),
    "TotW":   as_float(raw[c_totw]),
})

# ---------- areas ----------
U = lambda s: s.upper().strip()
GER = U("GERMANY")

def masks(frame):
    mask_C = (
        (frame["State"].eq(GER) & frame["Region"].isin({U("MECKLENBURG-VORPOMMERN"), U("SCHLESWIG-HOLSTEIN"), U("NIEDERSACHSEN")})) |
        (frame["State"].eq(GER) & frame["Region"].isin({U("THÜRINGEN"), U("HESSEN"), U("RHEINLAND-PFALZ")})) |
        (frame["Region"].eq(U("SACHSEN-ANHALT")))
    )
    mask_B = (
        frame["Region"].eq(U("BRANDENBURG")) |
        frame["State"].eq(U("POLAND")) |
        frame["Region"].eq(U("SACHSEN"))
    )
    mask_A = (
        frame["Region"].isin({U("BADEN-WÜRTTEMBERG"), U("BAYERN")}) |
        frame["State"].isin({U("AUSTRIA"), U("SLOVENIA")})
    )
    mask_CH = frame["State"].eq(U("SWITZERLAND"))
    mask_IT = frame["State"].isin({U("ITALY"), U("SAN MARINO")})
    return {
        "Area A": mask_A,
        "Area B": mask_B,
        "Area C": mask_C,
        "Switzerland": mask_CH,
        "Italy": mask_IT,
    }

# ---------- DATA FOR A SINGLE period ----------

def data_for_period(df_all, period_key):
    if period_key == "OR_4_5":
        df = df_all[df_all["Or_fin"].isin({"OR_4","OR_5"})].copy()
        period_label = "Or. 4–5"
    else:
        df = df_all[df_all["Or_fin"] == period_key].copy()
        period_label = f"Or. {period_key.split('_')[1]}"
    # deduplicate by ID and filter TotW > 0
    df = df.sort_values("ID").drop_duplicates(subset="ID", keep="first")
    df = df[(df["TotW"] > 0) & np.isfinite(df["TotW"])]
    ms = masks(df)
    labels = list(ms.keys())
    data = [df[m]["TotW"].values for m in ms.values()]
    return period_label, labels, data

# ---------- DRAW A SINGLE FIGURE ----------

def plot_boxes(period_key):
    per_label, labels, data = data_for_period(df0, period_key)
    color = PERIOD_COLORS[period_key]

    fig, ax = plt.subplots(figsize=FIGSIZE, dpi=DPI)

    # boxplot (Matplotlib 3.9+: tick_labels)
    bp = ax.boxplot(
        data,
        tick_labels=labels,
        widths=0.55,
        patch_artist=True,
        showfliers=False
    )
    # box style
    for box in bp['boxes']:
        box.set(facecolor=color, alpha=BOX_ALPHA, edgecolor=BOX_EDGE, linewidth=1.0)
    for w in bp['whiskers'] + bp['caps']:
        w.set(color=BOX_EDGE, linewidth=1.0)
    for med in bp['medians']:
        med.set(color=BOX_EDGE, linewidth=1.4)

    # jitter
    for xi, vals in enumerate(data, start=1):
        if len(vals) == 0: 
            continue
        xj = xi + rng.normal(0, JITTER_SD, size=len(vals))
        ax.scatter(xj, vals, s=18, facecolors="white",
                   edgecolors="0.25", alpha=0.65, zorder=3)

    # log scale (safe: values > 0 already filtered)
    allpos = np.concatenate([v for v in data if len(v)>0])
    ax.set_yscale("log")
    ax.set_ylim(allpos.min()*0.8, allpos.max()*1.5)


    ax.yaxis.grid(True, which="major", color="0.90", lw=0.7)
    ax.set_ylabel("Weight (Hoards)", fontsize=9)
    ax.tick_params(axis='x', labelsize=9, pad=6)
    ax.tick_params(axis='y', labelsize=9)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    # period label in the top-right corner, slightly outside the panel
    ax.text(0.98, 1.02, per_label, transform=ax.transAxes,
            ha="right", va="bottom", fontsize=11)

    plt.tight_layout()
    out = f"Fig27_30_33_36_TotW_boxplot_{period_key}.jpg"
    #plt.savefig(out, dpi=DPI, format="jpeg",
    #            pil_kwargs={"quality":95, "subsampling":0})
    plt.show()
    print(f"Salvato: {out}")

# ---------- EXECUTION (4 separate figures) ----------
for pk in ["OR_1","OR_2","OR_3","OR_4_5"]:
    plot_boxes(pk)


# %%

# ============================================================
# FIGURES 25, 29, 32, 35 — Frequency distribution of total hoard weight
# One figure per chronological horizon, with 1x5 panels (one per
# area: A, B, C, Switzerland, Italy).
# ============================================================

import pandas as pd, numpy as np, re, unicodedata
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

# ===== EDITORIAL PARAMETERS =====
FILENAME, SHEET = "DB.xlsx", "DB"
FIGSIZE = (6.30, 2.90)   # ~160 x 74 mm, approx. 1/3 page (same width, reduced height)
DPI = 300
BINS = 12                # numero classi (lineari)
EDGE_COLOR = "black"
EDGE_LW = 0.6
ALPHA = 0.55
HEADROOM = 0.40          # extra headroom above the bars (+25%)

# Desired ticks (values in kg)
TICKS_BY_PERIOD_KG = {
    "OR_1":   [0, 50, 100, 150],
    "OR_2":   [0, 10, 20, 30],
    "OR_3":   [0, 60, 120, 180],
    "OR_4_5": [0, 30, 60, 90],
}

# Period palette (consistent with the earlier boxplots)
PERIOD_COLORS = {
    "OR_1":   "#1f77b4",  # blue
    "OR_2":   "#2ca02c",  # green
    "OR_3":   "#ff7f0e",  # orange
    "OR_4_5": "#9467bd",  # purple
}

# ===== Utils =====
def col_like(df, name):
    low = {c.lower().strip(): c for c in df.columns}
    alts = {
        "id_sito": ["id_sito","id","site_id","id sito","idsite"],
        "or_fin":  ["or_fin","or","period","orizzonte"],
        "state":   ["state","stato","country"],
        "region":  ["region","regione","land"],
        "tot w":   ["tot w","tot_w","total w","total_w","total weight","peso totale","totweight","weight"],
    }
    k = name.lower().strip()
    if k in alts:
        for cand in alts[k]:
            if cand in low: return low[cand]
    return low.get(k, name)

def to_str_u(s):  # UPPER + strip
    return s.astype(str).str.upper().str.strip()

def as_float(s):
    return pd.to_numeric(s, errors="coerce").astype(float)

# ===== Reading & normalization =====
raw = pd.read_excel(FILENAME, sheet_name=SHEET)
c_id   = col_like(raw, "ID_Sito")
c_or   = col_like(raw, "Or_fin")
c_st   = col_like(raw, "State")
c_reg  = col_like(raw, "Region")
c_totw = col_like(raw, "Tot W")

df0 = pd.DataFrame({
    "ID":     to_str_u(raw[c_id]),
    "Or_fin": to_str_u(raw[c_or]).str.replace(r"\s+","_", regex=True),
    "State":  to_str_u(raw[c_st]),
    "Region": to_str_u(raw[c_reg]),
    "TotW":   as_float(raw[c_totw]),
})

# Deduplicate by site and filter > 0
df0 = (df0.sort_values("ID")
          .drop_duplicates(subset="ID", keep="first"))
df0 = df0[(df0["TotW"] > 0) & np.isfinite(df0["TotW"])].copy()

# ===== areas =====
U = lambda s: s.upper().strip()
GER = U("GERMANY")

def masks(frame):
    mask_C = (
        (frame["State"].eq(GER) & frame["Region"].isin({U("MECKLENBURG-VORPOMMERN"),
                                                        U("SCHLESWIG-HOLSTEIN"),
                                                        U("NIEDERSACHSEN")})) |
        (frame["State"].eq(GER) & frame["Region"].isin({U("THÜRINGEN"),
                                                        U("HESSEN"),
                                                        U("RHEINLAND-PFALZ")})) |
        (frame["Region"].eq(U("SACHSEN-ANHALT")))
    )
    mask_B = (
        frame["Region"].eq(U("BRANDENBURG")) |
        frame["State"].eq(U("POLAND")) |
        frame["Region"].eq(U("SACHSEN"))
    )
    mask_A = (
        frame["Region"].isin({U("BADEN-WÜRTTEMBERG"), U("BAYERN")}) |
        frame["State"].isin({U("AUSTRIA"), U("SLOVENIA")})
    )
    mask_CH = frame["State"].eq(U("SWITZERLAND"))
    mask_IT = frame["State"].isin({U("ITALY"), U("SAN MARINO")})
    return [
        ("Area A", mask_A),
        ("Area B", mask_B),
        ("Area C", mask_C),
        ("Switzerland", mask_CH),
        ("Italy", mask_IT),
    ]

def period_subset(df, key):
    if key == "OR_4_5":
        return df[df["Or_fin"].isin({"OR_4","OR_5"})].copy(), "Or. 4–5"
    else:
        return df[df["Or_fin"].eq(key)].copy(), f"Or. {key.split('_')[1]}"

UNIT = "kg"   # "g" for grams (labels 100k), "kg" per kilos

def plot_hist_period(df_all, period_key):
    sub, per_label = period_subset(df_all, period_key)
    areas = masks(sub)

    # --- scale/labels ---
    if UNIT.lower() == "kg":
        scale = 1/1000.0
        x_label = "Weight kg (Hoards)"
        xfmt = FuncFormatter(lambda x, pos: f"{int(x)}" if x >= 1 else f"{x:g}")
    else:  # grams
        scale = 1.0
        x_label = "Weight gr (Hoards)"
        def _kfmt(x, pos):
            if x >= 1_000_000: return f"{x/1_000_000:.0f}M"
            if x >= 1_000:     return f"{x/1_000:.0f}k"
            return f"{int(x)}"
        xfmt = FuncFormatter(_kfmt)

    # --- fixed ticks per period (defined in kg) -> convert to the chosen UNIT ---
    base_ticks_kg = TICKS_BY_PERIOD_KG[period_key]
    tick_vals = np.array(base_ticks_kg, dtype=float) * (1.0 if UNIT.lower()=="kg" else 1000.0)
    vmin, vmax = float(tick_vals[0]), float(tick_vals[-1])

    # --- shared linear bins: from 0 to the last tick (in the same unit as the axis) ---
    bins = np.linspace(vmin, vmax, BINS + 1)

    # Pre-compute the maxima for uniform headroom across all panels
    counts_max = 0
    for _, m in areas:
        vals = sub[m]["TotW"].values
        vals = vals[np.isfinite(vals) & (vals > 0)]
        vals_sc = vals * scale
        if vals_sc.size:
            c, _ = np.histogram(vals_sc, bins=bins)
            counts_max = max(counts_max, int(c.max()))
    ylim_top = max(1, counts_max) * (1 + HEADROOM)

    # figure 1x5
    fig, axes = plt.subplots(1, 5, figsize=FIGSIZE, dpi=DPI, sharex=True, sharey=True)
    color = PERIOD_COLORS[period_key]

    for ax, (name, m) in zip(axes, areas):
        vals = sub[m]["TotW"].values
        vals = vals[np.isfinite(vals) & (vals > 0)]
        vals_sc = vals * scale

        if vals_sc.size:
            ax.hist(vals_sc, bins=bins, color=color, alpha=ALPHA,
                    edgecolor=EDGE_COLOR, linewidth=EDGE_LW)
            # n in the top-right corner
            ax.text(0.98, 0.98, f"n = {len(vals_sc)}",
                    transform=ax.transAxes, ha="right", va="top",
                    fontsize=8, color="0.25")

        # panel style
        ax.set_title(name, fontsize=10, pad=6)
        ax.set_xlim(vmin, vmax)
        ax.set_xticks(tick_vals)
        ax.set_ylim(0, ylim_top)                 # <-- headroom uniforme
        ax.xaxis.set_major_formatter(xfmt)
        ax.yaxis.grid(True, color="0.90", lw=0.7)
        ax.set_axisbelow(True)
        for sp in ("top","right"):
            ax.spines[sp].set_visible(False)

    # common labels
    axes[0].set_ylabel("Counts", fontsize=9)
    axes[len(axes)//2].set_xlabel(x_label, fontsize=9, labelpad=6)

    # period label in the top-right corner
    fig.text(0.985, 0.995, per_label, ha="right", va="top", fontsize=11)

    plt.tight_layout(w_pad=0.8)
    suffix = "KG" if UNIT.lower()=="kg" else "GR"
    out = f"Fig25_29_32_35_TotW_hist_{period_key}_LINEAR_{suffix}.jpg"
    #plt.savefig(out, dpi=DPI, format="jpeg", pil_kwargs={"quality":95, "subsampling":0})
    plt.show()
    print(f"Salvato: {out}")

# Run for the desired periods
for pk in ["OR_1","OR_2","OR_3","OR_4_5"]:
    plot_hist_period(df0, pk)


# %%

# ============================================================
# FIGURE 41 — Frequency distribution of the weight of Ösenhalsringe
# (Ringbarren) in the Early Bronze Age (Horizon 1).
# ============================================================

import pandas as pd, numpy as np, re, unicodedata
import matplotlib.pyplot as plt

# ===== PARAMETERS =====
FILENAME, SHEET = "DB.xlsx", "DB"
FIGSIZE = (6.30, 4.33)   # ~160 x 110 mm (half page)
DPI = 300
BIN_W = 2.85             # ampiezza bin in grammi
BAR_COLOR = "#000000"   # pure black
ALPHA = 1.0             # no transparency
EDGE_COLOR = "white"    # same as the earlier Complete bars
EDGE_LW = 0.6


# ===== Utils =====
def col_like(df, name):
    low = {c.lower().strip(): c for c in df.columns}
    alts = {
        "artefact":   ["artefact","artifact","type","object","category","class","item","typology"],
        "complete":   ["complete","is_complete"],
        "weight_obj": ["weight_obj","weight obj","obj weight","weight (obj)","peso oggetto","peso_obj","object weight"],
        "state":      ["state","stato","country"],
        "region":     ["region","regione","land"],
    }
    k = name.lower().strip()
    if k in alts:
        for cand in alts[k]:
            if cand in low: return low[cand]
    return low.get(k, name)

def to_str_u(s):  # upper + strip
    return s.astype(str).str.upper().str.strip()

def as_float(s):
    return pd.to_numeric(s, errors="coerce").astype(float)

def as_int(s):
    return pd.to_numeric(s, errors="coerce").fillna(0).astype(int)

def strip_accents(s):
    s = s.replace("ß","ss")
    return "".join(ch for ch in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(ch))

def norm(s):
    s = strip_accents(str(s)).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

# ===== Reading =====
raw = pd.read_excel(FILENAME, sheet_name=SHEET)

c_art  = col_like(raw, "Artefact")
c_comp = col_like(raw, "Complete")
c_w    = col_like(raw, "Weight_obj")
c_st   = col_like(raw, "State")
c_reg  = col_like(raw, "Region")

df = pd.DataFrame({
    "Artefact": raw[c_art].astype(str),
    "Complete": as_int(raw[c_comp]),
    "Wobj":     as_float(raw[c_w]),
    "State":    to_str_u(raw[c_st]),
    "Region":   to_str_u(raw[c_reg]),
})

# ===== Filters: Ringbarren + Complete>0 + Weight_obj>0 =====
is_ringbarren = df["Artefact"].apply(lambda x: "ringbarren" in norm(x))
df = df[is_ringbarren & (df["Complete"] > 0) & np.isfinite(df["Wobj"]) & (df["Wobj"] > 0)].copy()

# ===== Areas (everything except Italy/San Marino) -> keep ONLY area A, B, C, Switzerland =====
U = lambda s: s.upper().strip()
GER = U("GERMANY")

mask_C = (
    (df["State"].eq(GER) & df["Region"].isin({U("MECKLENBURG-VORPOMMERN"), U("SCHLESWIG-HOLSTEIN"), U("NIEDERSACHSEN")})) |
    (df["State"].eq(GER) & df["Region"].isin({U("THÜRINGEN"), U("HESSEN"), U("RHEINLAND-PFALZ")})) |
    (df["Region"].eq(U("SACHSEN-ANHALT")))
)
mask_B = (
    df["Region"].eq(U("BRANDENBURG")) |
    df["State"].eq(U("POLAND")) |
    df["Region"].eq(U("SACHSEN"))
)
mask_A = (
    df["Region"].isin({U("BADEN-WÜRTTEMBERG"), U("BAYERN")}) |
    df["State"].isin({U("AUSTRIA"), U("SLOVENIA")})
)
mask_CH = df["State"].eq(U("SWITZERLAND"))

df_combined = df[mask_A | mask_B | mask_C | mask_CH].copy()

# ===== Shared bins (0 -> max, step 2.85 g) =====
# ===== Shared bins (0 -> 340 g, step 2.85 g) =====
vals = df_combined["Wobj"].to_numpy()
if vals.size == 0:
    print("No complete Ringbarren with Weight_obj>0 in areas A/B/C/Switzerland.")
else:
    BIN_W  = 2.85
    MAX_X  = 340.0
    shown  = vals[np.isfinite(vals) & (vals > 0) & (vals <= MAX_X)]
    bins   = np.arange(0.0, MAX_X + BIN_W, BIN_W)

    fig, ax = plt.subplots(figsize=FIGSIZE, dpi=DPI)
    fig.suptitle("Ringbarren", fontsize=12, y=0.985)
    plt.tight_layout(rect=[0, 0, 1, 0.96])  # leave room for the title
    ax.hist(shown, bins=bins, color=BAR_COLOR, alpha=ALPHA,
            edgecolor=EDGE_COLOR, linewidth=EDGE_LW)

    # n (only what is visible, <= 340 g) in the top-right corner
    ax.text(0.98, 0.98, f"n = {len(shown)}",
            transform=ax.transAxes, ha="right", va="top",
            fontsize=9, color="0.25")

    # style
    ax.set_xlabel("Weight (gr)", fontsize=10)
    ax.set_ylabel("Counts", fontsize=10)
    ax.set_xlim(0, MAX_X)
    ax.yaxis.grid(True, color="0.90", lw=0.7)
    ax.set_axisbelow(True)
    for sp in ("top","right"): ax.spines[sp].set_visible(False)

    plt.tight_layout()
    out = "Fig41_Ringbarren_hist_ABC_CH_SINGLE_halfpage_0_340g.jpg"
    #plt.savefig(out, dpi=DPI, format="jpeg",
    #            pil_kwargs={"quality":95, "subsampling":0})
    plt.show()
    print(f"Salvato: {out}")


# %%

# ============================================================
# FIGURE 43 — Frequency distribution of the weight of complete Spangenbarren
# Panel A: whole sample; Panel B: 'small' complete Spangenbarren;
# Panel C: 'large' complete Spangenbarren.
# ============================================================

import pandas as pd, numpy as np, re, unicodedata
import matplotlib.pyplot as plt

# ========== PARAMETERS ==========
FILENAME, SHEET = "DB.xlsx", "DB"

# "half" (half page) o "third" (1/3 page)
PAGE = "half"
FIGSIZE = (6.30, 4.33) if PAGE=="half" else (6.30, 3.07)

DPI     = 300
BAR_COLOR = "#000000"   # pure black
ALPHA = 1.0             # no transparency
EDGE_COLOR = "white"    # same as the earlier Complete bars
EDGE_LW = 0.6

# bin width specific to each panel
BIN_ALL = 5.76
BIN_B   = 3.11
BIN_C   = 3.67

# ========== UTILS ==========
def col_like(df, name):
    low = {c.lower().strip(): c for c in df.columns}
    alts = {
        "artefact":   ["artefact","artifact","type","object","category","class","item","typology"],
        "complete":   ["complete","is_complete"],
        "weight_obj": ["weight_obj","weight obj","obj weight","weight (obj)","peso oggetto","peso_obj","object weight"],
    }
    k = name.lower().strip()
    if k in alts:
        for cand in alts[k]:
            if cand in low: return low[cand]
    return low.get(k, name)

def as_int(s):   return pd.to_numeric(s, errors="coerce").fillna(0).astype(int)
def as_float(s): return pd.to_numeric(s, errors="coerce").astype(float)

def strip_accents(s):
    s = s.replace("ß","ss")
    return "".join(ch for ch in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(ch))
def norm(s):
    s = strip_accents(str(s)).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def make_bins(lo, hi, step):
    if hi <= lo: hi = lo + step
    right = np.ceil((hi - lo) / step) * step + lo
    return np.arange(lo, right + step/2.0, step)

# ========== READ & FILTERS ==========
raw = pd.read_excel(FILENAME, sheet_name=SHEET)
c_art  = col_like(raw, "Artefact")
c_comp = col_like(raw, "Complete")
c_w    = col_like(raw, "Weight_obj")

df = pd.DataFrame({
    "Artefact": raw[c_art].astype(str),
    "Complete": as_int(raw[c_comp]),
    "Wobj":     as_float(raw[c_w]),
})

is_spangen = df["Artefact"].apply(lambda x: "spangenbarren" in norm(x))
df = df[is_spangen & (df["Complete"] > 0) & np.isfinite(df["Wobj"]) & (df["Wobj"] > 0)].copy()

# Series for the three panels (data)
vals_all      = df["Wobj"].to_numpy()
vals_20_125   = vals_all[(vals_all >=  20) & (vals_all <= 125)]
vals_125_250  = vals_all[(vals_all >= 125) & (vals_all <= 250)]

# Bins per panel
bins_all      = make_bins(0,    vals_all.max() if vals_all.size else 0, BIN_ALL)
bins_20_125   = make_bins(20,   125,   BIN_B)
bins_125_250  = make_bins(125,  250,   BIN_C)

# ========== PLOT 1x3 ==========
fig, axes = plt.subplots(1, 3, figsize=FIGSIZE, dpi=DPI, sharey=False)
fig.suptitle("Spangenbarren", fontsize=12, y=0.98)

panels = [
    ("A", vals_all,     bins_all,      (0,   bins_all[-1] if bins_all.size else 10)),
    ("B", vals_20_125,  bins_20_125,   (20,  140)),  # only visualization
    ("C", vals_125_250, bins_125_250,  (120, 260)),  # only visualization
]

for j, (ax, (letter, arr, bins, xlim)) in enumerate(zip(axes, panels)):
    counts = np.array([])
    if arr.size:
        counts, edges, patches = ax.hist(
            arr, bins=bins, color=BAR_COLOR, alpha=ALPHA,
            edgecolor=EDGE_COLOR, linewidth=EDGE_LW
        )
    # n in the top-right corner
    ax.text(0.98, 0.98, f"n = {len(arr)}",
            transform=ax.transAxes, ha="right", va="top",
            fontsize=9, color="0.25")
    # panel letter in the top-left corner
    ax.text(0.015, 0.985, letter, transform=ax.transAxes,
            ha="left", va="top", fontsize=11, fontweight="bold")

    # limits 
    ax.set_xlim(*xlim)

    # headroom above the bars (~30% over the max count)
    if counts.size:
        top = counts.max()
        ax.set_ylim(0, np.ceil(top * 1.30))
    ax.yaxis.grid(True, color="0.90", lw=0.7)
    ax.set_axisbelow(True)
    for sp in ("top","right"): ax.spines[sp].set_visible(False)

    # labels
    ax.set_xlabel("Weight (gr)", fontsize=9)
    if j == 0:
        ax.set_ylabel("Counts", fontsize=9)   # only on the first
    else:
        ax.set_ylabel("")

plt.tight_layout(w_pad=1.0, rect=[0, 0, 1, 0.955])
outfile = f"Fig43_Spangenbarren_hist_1x3_bins_custom_{PAGE}.jpg"
#plt.savefig(outfile, dpi=DPI, format="jpeg",
#            pil_kwargs={"quality":95, "subsampling":0})
plt.show()
print(f"Salvato: {outfile}")


# %%

# ============================================================
# FIGURE 46 — Frequency distribution of the weight of Early Bronze Age
# flanged axes (Horizon 1). 5 stacked panels, one per area.
# ============================================================

import pandas as pd, numpy as np, re, unicodedata
import matplotlib.pyplot as plt

# ========== PARAMETERS ==========
FILENAME, SHEET = "DB.xlsx", "DB"

FIGSIZE = (6.30, 9.25)   # full page recommended for 5 rows
DPI       = 300
BAR_COLOR = "#000000"   # pure black
ALPHA = 1.0             # no transparency
EDGE_COLOR = "white"    # same as the earlier Complete bars
EDGE_LW = 0.6

# Bin width specific to each area (grams)
BIN_W_MAP = {
    "Area A":      28.00,
    "Area B":      12.13,
    "Area C":      13.06,
    "Switzerland": 36.88,
    "Italy":       48.22,
}

# ========== UTILS ==========
def col_like(df, name):
    low = {c.lower().strip(): c for c in df.columns}
    alts = {
        "or_fin":   ["or_fin","or","period","orizzonte"],
        "state":    ["state","stato","country"],
        "region":   ["region","regione","land"],
        "artefact": ["artefact","artifact","type","object","category","class","item","typology"],
        "typology": ["typology","type detail","tipo","sottotipo","subtype","sub type"],
        "complete": ["complete","is_complete"],
        "weight_obj":["weight_obj","weight obj","obj weight","weight (obj)","peso oggetto","peso_obj","object weight"],
    }
    k = name.lower().strip()
    if k in alts:
        for cand in alts[k]:
            if cand in low: return low[cand]
    return low.get(k, name)

def as_int(s):   return pd.to_numeric(s, errors="coerce").fillna(0).astype(int)
def as_float(s): return pd.to_numeric(s, errors="coerce").astype(float)

def strip_accents(s):
    s = str(s).replace("ß","ss")
    return "".join(ch for ch in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(ch))
def norm(s):
    s = strip_accents(str(s)).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def make_bins(lo, hi, step):
    if not np.isfinite(lo): lo = 0.0
    if not np.isfinite(hi) or hi <= lo: hi = lo + step
    right = np.ceil((hi - lo) / step) * step + lo
    return np.arange(lo, right + step/2.0, step)

# ========== READING ==========
raw = pd.read_excel(FILENAME, sheet_name=SHEET)
c_or   = col_like(raw, "Or_fin")
c_st   = col_like(raw, "State")
c_reg  = col_like(raw, "Region")
c_art  = col_like(raw, "Artefact")
c_typo = col_like(raw, "Typology")
c_comp = col_like(raw, "Complete")
c_wobj = col_like(raw, "Weight_obj")

df = pd.DataFrame({
    "Or_fin":   raw[c_or].astype(str).str.upper().str.strip().str.replace(r"\s+","_", regex=True),
    "State":    raw[c_st].astype(str).str.upper().str.strip(),
    "Region":   raw[c_reg].astype(str).str.upper().str.strip(),
    "Artefact": raw[c_art].astype(str),
    "Typology": raw[c_typo].astype(str),
    "Complete": as_int(raw[c_comp]),
    "Wobj":     as_float(raw[c_wobj]),
})

# ========== COMMON FILTER ==========
is_or1   = df["Or_fin"].eq("OR_1")
is_comp  = df["Complete"] > 0
has_w    = np.isfinite(df["Wobj"]) & (df["Wobj"] > 0)
is_axe   = df["Artefact"].apply(lambda x: "axe" in norm(x))
base = df[is_or1 & is_comp & has_w & is_axe].copy()

# ========== area MASKS ==========
U = lambda s: s.upper().strip()
GER = U("GERMANY")
mask_C = (
    (base["State"].eq(GER) & base["Region"].isin({U("MECKLENBURG-VORPOMMERN"), U("SCHLESWIG-HOLSTEIN"), U("NIEDERSACHSEN")})) |
    (base["State"].eq(GER) & base["Region"].isin({U("THÜRINGEN"), U("HESSEN"), U("RHEINLAND-PFALZ")})) |
    (base["Region"].eq(U("SACHSEN-ANHALT")))
)
mask_B = (
    base["Region"].eq(U("BRANDENBURG")) |
    base["State"].eq(U("POLAND")) |
    base["Region"].eq(U("SACHSEN"))
)
mask_A = (
    base["Region"].isin({U("BADEN-WÜRTTEMBERG"), U("BAYERN")}) |
    base["State"].isin({U("AUSTRIA"), U("SLOVENIA")})
)
mask_CH    = base["State"].eq(U("SWITZERLAND"))
mask_IT_SM = base["State"].isin({U("ITALY"), U("SAN MARINO")})

# type for A/B/C/CH
is_flanged = base["Typology"].apply(lambda x: ("flanged" in norm(x)) and ("axe" in norm(x)))

panels = [
    ("Area A",      base[mask_A  & is_flanged].copy()),
    ("Area B",      base[mask_B  & is_flanged].copy()),
    ("Area C",      base[mask_C  & is_flanged].copy()),
    ("Switzerland", base[mask_CH & is_flanged].copy()),
    ("Italy",       base[mask_IT_SM].copy()),  # Italy (+SM), all Axes
]

# ========== Plot 5x1 ==========
fig, axes = plt.subplots(5, 1, figsize=FIGSIZE, dpi=DPI, sharey=True)
fig.suptitle("Flanged Axes – Or. 1", fontsize=12, y=0.955)

# first pass: shared ylim with headroom
tops, vals_cache, bins_cache = [], [], []
for name, dfa in panels:
    vals = dfa["Wobj"].to_numpy()
    vals = vals[np.isfinite(vals) & (vals > 0)]
    step = BIN_W_MAP.get(name, 50.0)
    bins = make_bins(0.0, vals.max() if vals.size else 0.0, step)
    if vals.size and bins.size > 1:
        counts, _ = np.histogram(vals, bins=bins)
        tops.append(counts.max() if counts.size else 0)
    else:
        tops.append(0)
    vals_cache.append(vals); bins_cache.append(bins)

ylim_top = np.ceil(max(tops) * 1.30) if any(tops) else 1

# second pass: drawing
for i, (ax, (name, _)) in enumerate(zip(axes, panels)):
    vals = vals_cache[i]; bins = bins_cache[i]
    if vals.size and bins.size > 1:
        ax.hist(vals, bins=bins, color=BAR_COLOR, alpha=ALPHA,
                edgecolor=EDGE_COLOR, linewidth=EDGE_LW)
    # n in the top-right corner
    ax.text(0.985, 0.97, f"n = {len(vals)}",
            transform=ax.transAxes, ha="right", va="top",
            fontsize=8.5, color="0.25")
    ax.set_title(name, fontsize=10, pad=4, loc="left")
    ax.set_ylim(0, ylim_top)
    ax.yaxis.grid(True, color="0.90", lw=0.7)
    ax.set_axisbelow(True)
    for sp in ("top","right"): ax.spines[sp].set_visible(False)
    if i == 0:
        ax.set_ylabel("Counts", fontsize=9)
    else:
        ax.set_ylabel("Counts")

# x label only on the bottom panel
axes[-1].set_xlabel("Weight (gr)", fontsize=10)

plt.tight_layout(h_pad=0.50, rect=[0.06, 0.03, 0.98, 0.996])
outfile = "Fig46_Hist_OR1_Axes_Flanged_5panels_vertical_bins_by_area.jpg"
#plt.savefig(outfile, dpi=DPI, format="jpeg",
#            pil_kwargs={"quality":95, "subsampling":0})
plt.show()
print(f"Salvato: {outfile}")


# %%

# ============================================================
# FIGURE 47 — Frequency distribution of the weight of complete objects
# from Area A in the Middle Bronze Age (Horizon 2). Two panels:
# 0-100 g and 100-700 g.
# ============================================================

import pandas as pd, numpy as np, re, unicodedata
import matplotlib.pyplot as plt

# ===== PARAMETERS =====
FILENAME, SHEET = "DB.xlsx", "DB"
FIGSIZE = (6.30, 4.33)   # ~160 x 110 mm (half page)
DPI = 300
BAR_COLOR = "#000000"   # pure black
ALPHA = 1.0             # no transparency
EDGE_COLOR = "white"    # same as the earlier Complete bars
EDGE_LW = 0.6

BIN_TOP    = 3.33    # 0-100 g
BIN_BOTTOM = 11.52   # 100-700 g

# ===== Utils =====
def col_like(df, name):
    low = {c.lower().strip(): c for c in df.columns}
    alts = {
        "or_fin":   ["or_fin","or","period","orizzonte"],
        "state":    ["state","stato","country"],
        "region":   ["region","regione","land"],
        "complete": ["complete","is_complete"],
        "weight_obj":["weight_obj","weight obj","obj weight","weight (obj)","peso oggetto","peso_obj","object weight"],
    }
    k = name.lower().strip()
    if k in alts:
        for cand in alts[k]:
            if cand in low: return low[cand]
    return low.get(k, name)

def to_str_u(s):  # UPPER + strip
    return s.astype(str).str.upper().str.strip()

def as_int(s):   return pd.to_numeric(s, errors="coerce").fillna(0).astype(int)
def as_float(s): return pd.to_numeric(s, errors="coerce").astype(float)

def make_bins(lo, hi, step):
    if not np.isfinite(lo): lo = 0.0
    if not np.isfinite(hi) or hi <= lo: hi = lo + step
    right = np.ceil((hi - lo) / step) * step + lo
    return np.arange(lo, right + step/2.0, step)

# ===== Reading =====
raw = pd.read_excel(FILENAME, sheet_name=SHEET)
c_or   = col_like(raw, "Or_fin")
c_st   = col_like(raw, "State")
c_reg  = col_like(raw, "Region")
c_comp = col_like(raw, "Complete")
c_wobj = col_like(raw, "Weight_obj")

df = pd.DataFrame({
    "Or_fin": to_str_u(raw[c_or]).str.replace(r"\s+","_", regex=True),
    "State":  to_str_u(raw[c_st]),
    "Region": to_str_u(raw[c_reg]),
    "Comp":   as_int(raw[c_comp]),
    "Wobj":   as_float(raw[c_wobj]),
})

# ===== Filters =====
is_or2  = df["Or_fin"].eq("OR_2")
is_comp = df["Comp"] > 0
has_w   = np.isfinite(df["Wobj"]) & (df["Wobj"] > 0)

U = lambda s: s.upper().strip()
mask_areaA = (
    df["Region"].isin({U("BADEN-WÜRTTEMBERG"), U("BAYERN")}) |
    df["State"].isin({U("AUSTRIA"), U("SLOVENIA")})
)

sub = df[is_or2 & is_comp & has_w & mask_areaA].copy()
w = sub["Wobj"].to_numpy()

# split
w_top    = w[(w > 0)    & (w < 100)]
w_bottom = w[(w >= 100) & (w <= 700)]

# ===== Histograms 2x1 =====
fig, axes = plt.subplots(2, 1, figsize=FIGSIZE, dpi=DPI, sharex=False)

# Single centered title
fig.suptitle("Complete obj. (Area A – Or. 2)", fontsize=12, y=0.885)

# --- top panel: 0-100 g (bin 3.33) ---
ax = axes[0]
bins_top = make_bins(0.0, 100.0, BIN_TOP)
counts_top = np.array([])
if w_top.size and bins_top.size > 1:
    counts_top, edges, patches = ax.hist(
        w_top, bins=bins_top, color=BAR_COLOR, alpha=ALPHA,
        edgecolor=EDGE_COLOR, linewidth=EDGE_LW
    )
ax.text(0.985, 0.97, f"n = {len(w_top)}", transform=ax.transAxes,
        ha="right", va="top", fontsize=9, color="0.25")
ax.set_ylabel("Counts", fontsize=10)
if counts_top.size: ax.set_ylim(0, np.ceil(counts_top.max() * 1.30))
ax.set_xlim(0, 100)
ax.yaxis.grid(True, color="0.90", lw=0.7); ax.set_axisbelow(True)
for sp in ("top","right"): ax.spines[sp].set_visible(False)

# --- bottom panel: 100-700 g (bin 11.52) ---
ax = axes[1]
bins_bottom = make_bins(100.0, 700.0, BIN_BOTTOM)
counts_bottom = np.array([])
if w_bottom.size and bins_bottom.size > 1:
    counts_bottom, edges, patches = ax.hist(
        w_bottom, bins=bins_bottom, color=BAR_COLOR, alpha=ALPHA,
        edgecolor=EDGE_COLOR, linewidth=EDGE_LW
    )
ax.text(0.985, 0.97, f"n = {len(w_bottom)}", transform=ax.transAxes,
        ha="right", va="top", fontsize=9, color="0.25")
ax.set_xlabel("Weight (gr)", fontsize=10)
ax.set_ylabel("Counts", fontsize=10)
if counts_bottom.size: ax.set_ylim(0, np.ceil(counts_bottom.max() * 1.30))
ax.set_xlim(100, 700)
ax.yaxis.grid(True, color="0.90", lw=0.7); ax.set_axisbelow(True)
for sp in ("top","right"): ax.spines[sp].set_visible(False)

plt.tight_layout(h_pad=0.5, rect=[0.02, 0.02, 0.98, 0.97])
#plt.savefig("Fig47_Hist_OR2_AreaA_0-100_and_100-700_header_center.jpg", dpi=DPI, format="jpeg",
#            pil_kwargs={"quality":95, "subsampling":0})
plt.show()


# %%

# ============================================================
# FIGURE 48 — Frequency distribution of the weight of fragments
# from Area A in the Middle Bronze Age (Horizon 2). Two panels.
# ============================================================

import pandas as pd, numpy as np, re, unicodedata
import matplotlib.pyplot as plt

# ===== PARAMETERS =====
FILENAME, SHEET = "DB.xlsx", "DB"
FIGSIZE = (6.30, 4.33)   # ~160 x 110 mm (half page)
DPI = 300
COL_FRAGMENTED = "#CFCFCF"   # light gray
EDGE_COLOR     = "#000000"   # bordo black
EDGE_LW        = 0.5         # thin line width
ALPHA          = 1.0         # solid, as in the other charts

BIN_TOP    = 2.77    # 0-100 g
BIN_BOTTOM = 11.11   # 100-700 g

# ===== Utils =====
def col_like(df, name):
    low = {c.lower().strip(): c for c in df.columns}
    alts = {
        "or_fin":     ["or_fin","or","period","orizzonte"],
        "state":      ["state","stato","country"],
        "region":     ["region","regione","land"],
        "complete":   ["complete","is_complete"],
        "fragmented": ["fragmented","fragment"],
        "matching fr":["matching fr","matching_fr","matchingfr","matching"],
        "weight_obj": ["weight_obj","weight obj","obj weight","weight (obj)","peso oggetto","peso_obj","object weight"],
    }
    k = name.lower().strip()
    if k in alts:
        for cand in alts[k]:
            if cand in low: return low[cand]
    return low.get(k, name)

def to_str_u(s):  # UPPER + strip
    return s.astype(str).str.upper().str.strip()

def as_int(s):   return pd.to_numeric(s, errors="coerce").fillna(0).astype(int)
def as_float(s): return pd.to_numeric(s, errors="coerce").astype(float)

def make_bins(lo, hi, step):
    if not np.isfinite(lo): lo = 0.0
    if not np.isfinite(hi) or hi <= lo: hi = lo + step
    right = np.ceil((hi - lo) / step) * step + lo
    return np.arange(lo, right + step/2.0, step)

# ===== Reading =====
raw = pd.read_excel(FILENAME, sheet_name=SHEET)
c_or   = col_like(raw, "Or_fin")
c_st   = col_like(raw, "State")
c_reg  = col_like(raw, "Region")
c_comp = col_like(raw, "Complete")
c_frag = col_like(raw, "Fragmented")
c_match= col_like(raw, "Matching fr")
c_wobj = col_like(raw, "Weight_obj")

df = pd.DataFrame({
    "Or_fin": to_str_u(raw[c_or]).str.replace(r"\s+","_", regex=True),
    "State":  to_str_u(raw[c_st]),
    "Region": to_str_u(raw[c_reg]),
    "Comp":   as_int(raw[c_comp]),
    "Frag":   as_int(raw[c_frag]),
    "Match":  as_int(raw[c_match]),
    "Wobj":   as_float(raw[c_wobj]),
})

# ===== Filters =====
is_or2   = df["Or_fin"].eq("OR_2")
has_w    = np.isfinite(df["Wobj"]) & (df["Wobj"] > 0)

U = lambda s: s.upper().strip()
mask_areaA = (
    df["Region"].isin({U("BADEN-WÜRTTEMBERG"), U("BAYERN")}) |
    df["State"].isin({U("AUSTRIA"), U("SLOVENIA")})
)

cond_frag                = df["Frag"] > 0
cond_match_non_complete  = (df["Comp"] < 1) & (df["Match"] > 0)

sub = df[is_or2 & has_w & mask_areaA & (cond_frag | cond_match_non_complete)].copy()
w = sub["Wobj"].to_numpy()

# split 
w_top    = w[(w > 0)    & (w < 100)]
w_bottom = w[(w >= 100) & (w <= 700)]

# ===== Histograms 2x1 =====
fig, axes = plt.subplots(2, 1, figsize=FIGSIZE, dpi=DPI, sharex=False)
fig.suptitle("Fragmented obj. (Area A – Or. 2)", fontsize=12, y=0.985)

# --- top panel: 0-100 g (bin 2.77) ---
ax = axes[0]
bins_top = make_bins(0.0, 100.0, BIN_TOP)
counts_top = np.array([])
if w_top.size and bins_top.size > 1:
    counts_top, edges, patches = ax.hist(
        w_top, bins=bins_top, color=COL_FRAGMENTED, alpha=ALPHA,
        edgecolor=EDGE_COLOR, linewidth=EDGE_LW
    )
ax.text(0.985, 0.97, f"n = {len(w_top)}", transform=ax.transAxes,
        ha="right", va="top", fontsize=9, color="0.25")
ax.set_ylabel("Counts", fontsize=10)
if counts_top.size: ax.set_ylim(0, np.ceil(counts_top.max() * 1.30))
ax.set_xlim(0, 100)
ax.yaxis.grid(True, color="0.90", lw=0.7); ax.set_axisbelow(True)
for sp in ("top","right"): ax.spines[sp].set_visible(False)

# --- bottom panel: 100-700 g (bin 11.11) ---
ax = axes[1]
bins_bottom = make_bins(100.0, 700.0, BIN_BOTTOM)
counts_bottom = np.array([])
if w_bottom.size and bins_bottom.size > 1:
    counts_bottom, edges, patches = ax.hist(
        w_bottom, bins=bins_bottom, color=COL_FRAGMENTED, alpha=ALPHA,
        edgecolor=EDGE_COLOR, linewidth=EDGE_LW
    )
ax.text(0.985, 0.97, f"n = {len(w_bottom)}", transform=ax.transAxes,
        ha="right", va="top", fontsize=9, color="0.25")
ax.set_xlabel("Weight (gr)", fontsize=10)
ax.set_ylabel("Counts", fontsize=10)
if counts_bottom.size: ax.set_ylim(0, np.ceil(counts_bottom.max() * 1.30))
ax.set_xlim(100, 700)
ax.yaxis.grid(True, color="0.90", lw=0.7); ax.set_axisbelow(True)
for sp in ("top","right"): ax.spines[sp].set_visible(False)

plt.tight_layout(h_pad=0.8, rect=[0.02, 0.02, 0.98, 0.97])
#plt.savefig("Fig48_Hist_OR2_AreaA_Fragmented_plusMatchingNonComplete_2panels.jpg",
#            dpi=DPI, format="jpeg", pil_kwargs={"quality":95, "subsampling":0})
plt.show()


# %%

# ============================================================
# FIGURES 50, 51, 52, 53, 54 — Frequency distribution of the weight
# of complete objects from the Late Bronze Age (Horizons 3-5), one
# figure per area: A, B, C, Switzerland, Italy.
# ============================================================

import pandas as pd, numpy as np, re, unicodedata
import matplotlib.pyplot as plt

# ===== PARAMETERS  =====
FILENAME, SHEET = "DB.xlsx", "DB"
FIGSIZE = (6.30, 4.33)   # ~160 x 110 mm (half page)
DPI = 300
BAR_COLOR = "#000000"   # black
ALPHA = 1            
EDGE_COLOR = "white"   
EDGE_LW = 0.4


BIN_TOP     = 3.33   # 0-300 g
BIN_BOTTOM  = 6.66   # 300-600 g
X_TOP       = (0.0, 300.0)
X_BOTTOM    = (300.0, 600.0)

# ===== Utils =====
def col_like(df, name):
    low = {c.lower().strip(): c for c in df.columns}
    alts = {
        "or_fin":     ["or_fin","or","period","orizzonte"],
        "state":      ["state","stato","country"],
        "region":     ["region","regione","land"],
        "complete":   ["complete","is_complete"],
        "weight_obj": ["weight_obj","weight obj","obj weight","weight (obj)","peso oggetto","peso_obj","object weight"],
    }
    k = name.lower().strip()
    if k in alts:
        for cand in alts[k]:
            if cand in low: return low[cand]
    return low.get(k, name)

def to_str_u(s):  # UPPER + strip
    return s.astype(str).str.upper().str.strip()

def as_int(s):   return pd.to_numeric(s, errors="coerce").fillna(0).astype(int)
def as_float(s): return pd.to_numeric(s, errors="coerce").astype(float)

def make_bins(lo, hi, step):
    if not np.isfinite(lo): lo = 0.0
    if not np.isfinite(hi) or hi <= lo: hi = lo + step
    right = np.ceil((hi - lo) / step) * step + lo
    return np.arange(lo, right + step/2.0, step)

# ===== Reading =====
raw = pd.read_excel(FILENAME, sheet_name=SHEET)
c_or   = col_like(raw, "Or_fin")
c_st   = col_like(raw, "State")
c_reg  = col_like(raw, "Region")
c_comp = col_like(raw, "Complete")
c_wobj = col_like(raw, "Weight_obj")

df = pd.DataFrame({
    "Or_fin": to_str_u(raw[c_or]).str.replace(r"\s+","_", regex=True),
    "State":  to_str_u(raw[c_st]),
    "Region": to_str_u(raw[c_reg]),
    "Comp":   as_int(raw[c_comp]),
    "Wobj":   as_float(raw[c_wobj]),
})

# ===== area MASKS =====
U = lambda s: s.upper().strip()
GER = U("GERMANY")

mask_A = (
    df["Region"].isin({U("BADEN-WÜRTTEMBERG"), U("BAYERN")}) |
    df["State"].isin({U("AUSTRIA"), U("SLOVENIA")})
)

mask_B = (
    df["Region"].eq(U("BRANDENBURG")) |
    df["State"].eq(U("POLAND")) |
    df["Region"].eq(U("SACHSEN"))
)

mask_C = (
    (df["State"].eq(GER) & df["Region"].isin({U("MECKLENBURG-VORPOMMERN"), U("SCHLESWIG-HOLSTEIN"), U("NIEDERSACHSEN")})) |
    (df["State"].eq(GER) & df["Region"].isin({U("THÜRINGEN"), U("HESSEN"), U("RHEINLAND-PFALZ")})) |
    (df["Region"].eq(U("SACHSEN-ANHALT")))
)

mask_CH = df["State"].eq(U("SWITZERLAND"))
mask_IT = df["State"].isin({U("ITALY"), U("SAN MARINO")})  # label: "Italy"

AREAS = [
    ("Area A", mask_A),
    ("Area B", mask_B),
    ("Area C", mask_C),
    ("Switzerland", mask_CH),
    ("Italy", mask_IT),
]

# ===== Plotting function for a single area =====
def plot_area_complete_3to5(area_label, area_mask):
    is_or_3to5 = df["Or_fin"].isin({"OR_3","OR_4","OR_5"})
    is_comp    = df["Comp"] > 0
    has_w      = np.isfinite(df["Wobj"]) & (df["Wobj"] > 0)

    sub = df[is_or_3to5 & is_comp & has_w & area_mask].copy()
    w = sub["Wobj"].to_numpy()

    # split 
    w_top    = w[(w > 0) & (w <  X_TOP[1])]
    w_bottom = w[(w >= X_BOTTOM[0]) & (w <= X_BOTTOM[1])]

    fig, axes = plt.subplots(2, 1, figsize=FIGSIZE, dpi=DPI, sharex=False)
    fig.suptitle(f"Complete obj. ({area_label} – Or. 3–5)", fontsize=12, y=0.985)

    # top panel
    ax = axes[0]
    bins_top = make_bins(*X_TOP, BIN_TOP)
    counts_top = np.array([])
    if w_top.size and bins_top.size > 1:
        counts_top, edges, patches = ax.hist(
            w_top, bins=bins_top, color=BAR_COLOR, alpha=ALPHA,
            edgecolor=EDGE_COLOR, linewidth=EDGE_LW
        )
    ax.text(0.985, 0.97, f"n = {len(w_top)}", transform=ax.transAxes,
            ha="right", va="top", fontsize=9, color="0.25")
    ax.set_ylabel("Counts", fontsize=10)
    if counts_top.size: ax.set_ylim(0, np.ceil(counts_top.max() * 1.30))
    ax.set_xlim(*X_TOP)
    ax.yaxis.grid(True, color="0.90", lw=0.7); ax.set_axisbelow(True)
    for sp in ("top","right"): ax.spines[sp].set_visible(False)

    # bottom panel
    ax = axes[1]
    bins_bottom = make_bins(*X_BOTTOM, BIN_BOTTOM)
    counts_bottom = np.array([])
    if w_bottom.size and bins_bottom.size > 1:
        counts_bottom, edges, patches = ax.hist(
            w_bottom, bins=bins_bottom, color=BAR_COLOR, alpha=ALPHA,
            edgecolor=EDGE_COLOR, linewidth=EDGE_LW
        )
    ax.text(0.985, 0.97, f"n = {len(w_bottom)}", transform=ax.transAxes,
            ha="right", va="top", fontsize=9, color="0.25")
    ax.set_xlabel("Weight (gr)", fontsize=10)
    ax.set_ylabel("Counts", fontsize=10)
    if counts_bottom.size: ax.set_ylim(0, np.ceil(counts_bottom.max() * 1.30))
    ax.set_xlim(*X_BOTTOM)
    ax.yaxis.grid(True, color="0.90", lw=0.7); ax.set_axisbelow(True)
    for sp in ("top","right"): ax.spines[sp].set_visible(False)

    plt.tight_layout(h_pad=0.8, rect=[0.02, 0.02, 0.98, 0.97])
    out = f"Fig50_51_52_53_54_Hist_{area_label.replace(' ','')}_OR3-5_CompleteObj_0-300_300-600.jpg"
    #plt.savefig(out, dpi=DPI, format="jpeg", pil_kwargs={"quality":95, "subsampling":0})
    plt.show()
    print(f"Salvato: {out}")

# ===== Run for A, B, C, Switzerland, Italy =====
for label, m in AREAS:
    plot_area_complete_3to5(label, m)


# %%

# ============================================================
# FIGURE 56 — Frequency distribution of the weight of complete ingots
# (casting cakes) from the Late Bronze Age (Horizons 3-5).
# 5x1 panels, all areas in a single figure.
# ============================================================

import pandas as pd, numpy as np, re, unicodedata
import matplotlib.pyplot as plt

# ===== EDITORIAL PARAMETERS =====
FILENAME, SHEET = "DB.xlsx", "DB"
FIGSIZE = (6.30, 6.94)   # ~160 x 176 mm (3/4 page)
DPI = 300
BAR_COLOR = "#000000"    # black (Complete)
ALPHA = 1.0              # solid
EDGE_COLOR = "white"     # thin white edge
EDGE_LW = 0.6

TITLE = "Complete Ingots (Or. 3–5)"
BIN_WIDTH = 150.0        # <<< bin 150 g

# ===== Utils =====
def col_like(df, name):
    low = {c.lower().strip(): c for c in df.columns}
    alts = {
        "or_fin":     ["or_fin","or","period","orizzonte"],
        "state":      ["state","stato","country"],
        "region":     ["region","regione","land"],
        "complete":   ["complete","is_complete"],
        "artefact":   ["artefact","artifact","type","object","category","class","item","typology"],
        "weight_obj": ["weight_obj","weight obj","obj weight","weight (obj)","peso oggetto","peso_obj","object weight"],
    }
    k = name.lower().strip()
    if k in alts:
        for cand in alts[k]:
            if cand in low: return low[cand]
    return low.get(k, name)

def to_str_u(s):  # UPPER + strip
    return s.astype(str).str.upper().str.strip()

def as_int(s):   return pd.to_numeric(s, errors="coerce").fillna(0).astype(int)
def as_float(s): return pd.to_numeric(s, errors="coerce").astype(float)

def strip_accents(s):
    s = s.replace("ß","ss")
    return "".join(ch for ch in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(ch))

def norm(s):
    s = strip_accents(str(s)).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

# ===== Reading =====
raw = pd.read_excel(FILENAME, sheet_name=SHEET)
c_or   = col_like(raw, "Or_fin")
c_st   = col_like(raw, "State")
c_reg  = col_like(raw, "Region")
c_comp = col_like(raw, "Complete")
c_art  = col_like(raw, "Artefact")
c_wobj = col_like(raw, "Weight_obj")

df = pd.DataFrame({
    "Or_fin":   to_str_u(raw[c_or]).str.replace(r"\s+","_", regex=True),
    "State":    to_str_u(raw[c_st]),
    "Region":   to_str_u(raw[c_reg]),
    "Complete": as_int(raw[c_comp]),
    "Artefact": raw[c_art].astype(str),
    "Wobj":     as_float(raw[c_wobj]),
})

# ===== Filters: OR_3-5, Complete>0, Weight_obj>0, Artefact=Ingot (variants included) =====
is_or_3to5 = df["Or_fin"].isin({"OR_3","OR_4","OR_5"})
is_comp    = df["Complete"] > 0
has_w      = np.isfinite(df["Wobj"]) & (df["Wobj"] > 0)

def is_ingot(s):
    t = norm(s)
    return bool(re.search(r"\b(ingot|ingots|bun\s*ingot|barren|rohbarren|spitzbarren|plattbarren|scheibenbarren)\b", t))

is_ing = df["Artefact"].apply(is_ingot)

df = df[is_or_3to5 & is_comp & has_w & is_ing].copy()

# ===== areas =====
U = lambda s: s.upper().strip()
GER = U("GERMANY")
mask_A  = (df["Region"].isin({U("BADEN-WÜRTTEMBERG"), U("BAYERN")}) | df["State"].isin({U("AUSTRIA"), U("SLOVENIA")}))
mask_B  = (df["Region"].eq(U("BRANDENBURG")) | df["State"].eq(U("POLAND")) | df["Region"].eq(U("SACHSEN")))
mask_C  = ((df["State"].eq(GER) & df["Region"].isin({U("MECKLENBURG-VORPOMMERN"), U("SCHLESWIG-HOLSTEIN"), U("NIEDERSACHSEN")})) |
           (df["State"].eq(GER) & df["Region"].isin({U("THÜRINGEN"), U("HESSEN"), U("RHEINLAND-PFALZ")})) |
           (df["Region"].eq(U("SACHSEN-ANHALT"))))
mask_CH = df["State"].eq(U("SWITZERLAND"))
mask_IT = df["State"].isin({U("ITALY"), U("SAN MARINO")})  # label: Italy

AREAS = [
    ("Area A", mask_A),
    ("Area B", mask_B),
    ("Area C", mask_C),
    ("Switzerland", mask_CH),
    ("Italy", mask_IT),
]

# ===== fixed bins: 0, 150, 300, ... =====
all_vals = df["Wobj"].values
all_vals = all_vals[np.isfinite(all_vals) & (all_vals > 0)]
if all_vals.size == 0:
    raise ValueError("No data available for Complete Ingots (Or. 3-5).")

vmin, vmax_raw = 0.0, float(all_vals.max())
vmax = np.ceil(vmax_raw / BIN_WIDTH) * BIN_WIDTH
bins = np.arange(vmin, vmax + BIN_WIDTH, BIN_WIDTH)

# ===== Plot 5x1 =====
fig, axes = plt.subplots(5, 1, figsize=FIGSIZE, dpi=DPI, sharex=True)
fig.suptitle(TITLE, fontsize=12, y=0.985)

for ax, (label, m) in zip(axes, AREAS):
    vals = df[m]["Wobj"].values
    vals = vals[np.isfinite(vals) & (vals > 0)]
    counts = np.array([])
    if vals.size and bins.size > 1:
        counts, _, _ = ax.hist(vals, bins=bins,
                               color=BAR_COLOR, alpha=ALPHA,
                               edgecolor=EDGE_COLOR, linewidth=EDGE_LW)
    # n in the top-right corner
    ax.text(0.985, 0.95, f"n = {len(vals)}",
            transform=ax.transAxes, ha="right", va="top",
            fontsize=8, color="0.25")
    # panel title
    ax.set_title(label, fontsize=10, pad=4, loc="left")
    # style
    ax.yaxis.grid(True, color="0.90", lw=0.7); ax.set_axisbelow(True)
    for sp in ("top","right"): ax.spines[sp].set_visible(False)
    if counts.size: ax.set_ylim(0, np.ceil(counts.max()*1.30))

axes[-1].set_xlabel("Weight (gr)", fontsize=10)
for ax in axes: ax.set_ylabel("Counts", fontsize=9)

plt.tight_layout(h_pad=0.8, rect=[0.04, 0.02, 0.98, 0.97])
out = "Fig56_Complete_Ingots_OR3-5_hist_5x1_bin150.jpg"
#plt.savefig(out, dpi=DPI, format="jpeg",
#            pil_kwargs={"quality":95, "subsampling":0})
plt.show()
print(f"Salvato: {out}")


# %%

# ============================================================
# FIGURES 57, 58, 59, 60, 61 — Frequency distribution of the weight
# of fragmented ingots (casting cakes) from the Late Bronze Age
# (Horizons 3-5), one figure per area: A, B, C, Switzerland, Italy.
# ============================================================

# ================== FRAGMENTED INGOTS - 1/3 PAGE, ALL AREAS (SEPARATE FIGURES) ==================
from matplotlib.ticker import MultipleLocator, FormatStrFormatter

# --- PARAMETERS layout ---
FIGSIZE_HALF = (6.30, 4.33)    # approx. 160 x 110 mm (half page)
DPI = 300
FRAG_GRAY = "#CFCFCF"
EDGE_KW = dict(edgecolor="#000000", linewidth=0.5)
HEADROOM = 0.25  # extra headroom above the bars (25%)

# --- required columns (reusing the col_like / to_str_u helpers) ---
c_or     = col_like(raw, "Or_fin")
c_state  = col_like(raw, "State")
c_reg    = col_like(raw, "Region")
c_type   = col_like(raw, "Artefact")
c_weight = col_like(raw, "Weight_obj")
c_frag   = col_like(raw, "Fragmented")
c_comp   = col_like(raw, "Complete")
c_match  = col_like(raw, "Matching fr")

# Working DataFrame
d = pd.DataFrame({
    "Or_fin":       to_str_u(raw[c_or]).str.replace(r"\s+","_", regex=True),
    "State":        to_str_u(raw[c_state]),
    "Region":       to_str_u(raw[c_reg]),
    "Artefact":     raw[c_type].astype(str).str.strip(),
    "Weight_obj":   pd.to_numeric(raw[c_weight], errors="coerce"),
    "Fragmented":   pd.to_numeric(raw[c_frag],   errors="coerce").fillna(0),
    "Complete":     pd.to_numeric(raw[c_comp],   errors="coerce").fillna(0),
    "Matching_fr":  pd.to_numeric(raw[c_match],  errors="coerce").fillna(0),
})

# util
U = lambda s: str(s).upper().strip()

def _strip_accents(s):
    s = str(s).replace("ß","ss")
    return "".join(ch for ch in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(ch))

def _norm(s):
    s = _strip_accents(s).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def is_ingot(s_raw: str) -> bool:
    s = _norm(s_raw)
    # exclude Ring/Spangenbarren should they ever appear
    if "ringbarren" in s or "spangenbarren" in s:
        return False
    return ("ingot" in s) or ("barren" in s)

# --- masks for Areas (as defined) ---
GER = U("GERMANY")
mask_A = (
    d["Region"].isin({U("BADEN-WÜRTTEMBERG"), U("BAYERN")}) |
    d["State"].isin({U("AUSTRIA"), U("SLOVENIA")})
)
mask_B = (
    d["Region"].eq(U("BRANDENBURG")) |
    d["State"].eq(U("POLAND")) |
    d["Region"].eq(U("SACHSEN"))
)
mask_C = (
    (d["State"].eq(GER) & d["Region"].isin({U("MECKLENBURG-VORPOMMERN"), U("SCHLESWIG-HOLSTEIN"), U("NIEDERSACHSEN")})) |
    (d["State"].eq(GER) & d["Region"].isin({U("THÜRINGEN"), U("HESSEN"), U("RHEINLAND-PFALZ")})) |
    (d["Region"].eq(U("SACHSEN-ANHALT")))
)
mask_CH = d["State"].eq(U("SWITZERLAND"))
mask_I  = d["State"].isin({U("ITALY"), U("SAN MARINO")})

AREAS = [
    ("Area A",     mask_A),
    ("Area B",     mask_B),
    ("Area C",     mask_C),
    ("Switzerland",mask_CH),
    ("Italy",      mask_I),
]

def make_fragmented_ingots_hist(area_label: str, area_mask):
    # required global filters
    mask_period = d["Or_fin"].isin({"OR_3","OR_4","OR_5"})
    mask_weight = d["Weight_obj"] > 0
    mask_ingot  = d["Artefact"].apply(is_ingot)
    # fragments: Fragmented > 0  or  (Matching_fr > 0 and Complete < 1)
    mask_frag   = (d["Fragmented"] > 0) | ((d["Matching_fr"] > 0) & (d["Complete"] < 1))

    sel = d[area_mask & mask_period & mask_weight & mask_ingot & mask_frag].copy()

    w = sel["Weight_obj"].dropna().values
    w_top    = w[(w > 0)   & (w <= 100)]
    w_bottom = w[(w > 100) & (w <= 1000)]

    # bins 
    bw_top = 1.9
    bw_bot = 8.3
    bins_top = np.arange(0,   100 + bw_top, bw_top)
    bins_bot = np.arange(100, 1000 + bw_bot, bw_bot)

    # 1/3-page style
    FIGSIZE_THIRD = (6.30, 3.08)  # approx. 160 x 78 mm
    DPI = 300
    FRAG_GRAY = "#CFCFCF"
    EDGE_KW = dict(edgecolor="#000000", linewidth=0.5)

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=FIGSIZE_THIRD, dpi=DPI, sharex=False, gridspec_kw={"hspace":0.35})

    # top panel: 0-100
    ax1.hist(w_top, bins=bins_top, color=FRAG_GRAY, **EDGE_KW)
    ax1.set_xlim(0, 100)
    ax1.xaxis.set_major_locator(MultipleLocator(10))
    ax1.xaxis.set_major_formatter(FormatStrFormatter('%d'))
    ax1.set_ylabel("Counts", fontsize=9)
    # ← extra headroom above the bars
    ytop = ax1.get_ylim()[1]
    ax1.set_ylim(0, max(1, ytop) * (1 + HEADROOM))
    
    ax1.yaxis.grid(True, color="0.90", lw=0.7)
    ax1.set_axisbelow(True)
    ax1.spines["top"].set_visible(False); ax1.spines["right"].set_visible(False)
    ax1.tick_params(axis="both", labelsize=8)
    ax1.text(0.98, 0.96, f"n = {len(w_top)}", transform=ax1.transAxes,
             ha="right", va="top", fontsize=8)


    # bottom panel: 100-1000
    ax2.hist(w_bottom, bins=bins_bot, color=FRAG_GRAY, **EDGE_KW)
    ax2.set_xlim(100, 1000)
    ax2.xaxis.set_major_locator(MultipleLocator(100))
    ax2.xaxis.set_major_formatter(FormatStrFormatter('%d'))
    ax2.set_ylabel("Counts", fontsize=9)
    # ← extra headroom above the bars
    ybot = ax2.get_ylim()[1]
    ax2.set_ylim(0, max(1, ybot) * (1 + HEADROOM))
    
    ax2.yaxis.grid(True, color="0.90", lw=0.7)
    ax2.set_axisbelow(True)
    ax2.spines["top"].set_visible(False); ax2.spines["right"].set_visible(False)
    ax2.tick_params(axis="both", labelsize=8)
    ax2.set_xlabel("Weight (gr)", fontsize=9)
    ax2.text(0.98, 0.96, f"n = {len(w_bottom)}", transform=ax2.transAxes,
             ha="right", va="top", fontsize=8)


    # single title
    fig.suptitle(f"Fragmented ingots ({area_label} – Or. 3–5)", fontsize=10.5)

    plt.tight_layout(pad=0.8, rect=[0.06, 0.06, 0.98, 0.93])
    out = f"Fig57_58_59_60_61_Fragmented_ingots_{area_label.replace(' ','')}_OR3-5_thirdpage.jpg"
    #plt.savefig(out, dpi=DPI, format="jpeg", bbox_inches="tight",
    #            pil_kwargs={"quality":95, "subsampling":0})
    plt.show()
    print(f"Salvato: {out}  |  n_top={len(w_top)}  n_bottom={len(w_bottom)}  n_tot={len(w)}")

# --- generate the 5 separate figures ---
for label, m in AREAS:
    make_fragmented_ingots_hist(label, m)


# %%

# ============================================================
# FIGURES 63, 64, 65, 66, 67 — Frequency distribution of the weight
# of fragmented objects from the Late Bronze Age (Horizons 3-5),
# excluding ingots and gold objects, one figure per area:
# A, B, C, Switzerland, Italy.
# ============================================================

# ================== FRAGMENTED OBJ. - NO INGOTS / NO GOLD - 1/2 PAGE (SEPARATE AREAS, HIGHLIGHT BIN) ==================
    
from matplotlib.ticker import MultipleLocator, FormatStrFormatter
import numpy as np
import re, unicodedata
import matplotlib.pyplot as plt

# --- PARAMETERS layout ---
FIGSIZE_HALF = (6.30, 4.33)    # approx. 160 x 110 mm (half page)
DPI = 300
FRAG_GRAY = "#CFCFCF"
EDGE_KW = dict(edgecolor="#000000", linewidth=0.5)
HEADROOM = 0.25  # extra headroom above the bars (25%)

# --- highlight theme (top panel only) ---
HIGHLIGHT_THEMES = {
    "blue": {"facecolor": "#DBE9FF", "alpha": 0.45, "edgecolor": "none", "zorder": 1},
    "warm": {"facecolor": "#F7F1D5", "alpha": 0.50, "edgecolor": "none", "zorder": 1},
    "gray": {"facecolor": "#EFEFEF", "alpha": 0.70, "edgecolor": "none", "zorder": 1},
}
HIGHLIGHT_STYLE = dict(facecolor="darkred", alpha=0.30, edgecolor="none", zorder=1)

# --- required columns (reusing the col_like / to_str_u helpers) ---
c_or     = col_like(raw, "Or_fin")
c_state  = col_like(raw, "State")
c_reg    = col_like(raw, "Region")
c_type   = col_like(raw, "Artefact")
c_weight = col_like(raw, "Weight_obj")
c_frag   = col_like(raw, "Fragmented")
c_comp   = col_like(raw, "Complete")
c_match  = col_like(raw, "Matching fr")

# Working DataFrame
d = pd.DataFrame({
    "Or_fin":       to_str_u(raw[c_or]).str.replace(r"\s+","_", regex=True),
    "State":        to_str_u(raw[c_state]),
    "Region":       to_str_u(raw[c_reg]),
    "Artefact":     raw[c_type].astype(str).str.strip(),
    "Weight_obj":   pd.to_numeric(raw[c_weight], errors="coerce"),
    "Fragmented":   pd.to_numeric(raw[c_frag],   errors="coerce").fillna(0),
    "Complete":     pd.to_numeric(raw[c_comp],   errors="coerce").fillna(0),
    "Matching_fr":  pd.to_numeric(raw[c_match],  errors="coerce").fillna(0),
})

# --- util ---
U = lambda s: str(s).upper().strip()

def _strip_accents(s):
    s = str(s).replace("ß","ss")
    return "".join(ch for ch in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(ch))

def _norm(s):
    s = _strip_accents(s).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def is_ingot(s_raw: str) -> bool:
    s = _norm(s_raw)
    # exclude Ring/Spangenbarren by ingots
    if "ringbarren" in s or "spangenbarren" in s:
        return False
    return ("ingot" in s) or ("barren" in s)

def starts_with_gold(s_raw: str) -> bool:
    # excludes anything STARTING with "gold " (gold sheet, gold pin, ...)
    s = _norm(s_raw)
    return s.startswith("gold ")

# --- masks for Areas (as defined) ---
    
GER = U("GERMANY")
mask_A = (
    d["Region"].isin({U("BADEN-WÜRTTEMBERG"), U("BAYERN")}) |
    d["State"].isin({U("AUSTRIA"), U("SLOVENIA")})
)
mask_B = (
    d["Region"].eq(U("BRANDENBURG")) |
    d["State"].eq(U("POLAND")) |
    d["Region"].eq(U("SACHSEN"))
)
mask_C = (
    (d["State"].eq(GER) & d["Region"].isin({U("MECKLENBURG-VORPOMMERN"), U("SCHLESWIG-HOLSTEIN"), U("NIEDERSACHSEN")})) |
    (d["State"].eq(GER) & d["Region"].isin({U("THÜRINGEN"), U("HESSEN"), U("RHEINLAND-PFALZ")})) |
    (d["Region"].eq(U("SACHSEN-ANHALT")))
)
mask_CH = d["State"].eq(U("SWITZERLAND"))
mask_I  = d["State"].isin({U("ITALY"), U("SAN MARINO")})

AREAS = [
    ("Area A",      mask_A),
    ("Area B",      mask_B),
    ("Area C",      mask_C),
    ("Switzerland", mask_CH),
    ("Italy",       mask_I),
]

# ---------- helper highlighting based on real BINS ----------
def _nearest_bin_index(bins, value):
    """Index of the bin whose center is nearest to 'value'."""
    centers = (bins[:-1] + bins[1:]) / 2.0
    return int(np.argmin(np.abs(centers - float(value))))

def _indices_around(bins, center, n_before, n_after, include_center=True):
    """Set of indices around the bin nearest to 'center' (clipped to the limits)."""
    i0 = _nearest_bin_index(bins, center)
    idxs = set()
    if include_center:
        idxs.add(i0)
    for k in range(1, n_before + 1):
        j = i0 - k
        if j >= 0:
            idxs.add(j)
    for k in range(1, n_after + 1):
        j = i0 + k
        if j <= len(bins) - 2:
            idxs.add(j)
    return idxs

def _merge_consecutive(idxs):
        """Converts indices [i1, i2, ...] into continuous blocks [(start, end), ...]."""
    if not idxs:
        return []
    idxs = sorted(idxs)
    blocks = []
    start = prev = idxs[0]
    for i in idxs[1:]:
        if i == prev + 1:
            prev = i
        else:
            blocks.append((start, prev))
            start = prev = i
    blocks.append((start, prev))
    return blocks

def compute_all_highlight_bins(bins_top):
    """
    Rules:
      10  → central bin only
      20  → 1 bin before, central bin, 1 bin after
      30  → 2 bins before, central bin, 1 bin after
      40, 50, 60 → 2 bins before, central bin, 2 bins after
      80  → 3 bins before, central bin, 3 bins after
    """
    idxs = set()
    idxs |= _indices_around(bins_top, 10, 0, 0, include_center=True)
    idxs |= _indices_around(bins_top, 20, 1, 1, include_center=True)
    idxs |= _indices_around(bins_top, 30, 2, 1, include_center=True)
    for c in (40, 50, 60):
        idxs |= _indices_around(bins_top, c, 2, 2, include_center=True)
    idxs |= _indices_around(bins_top, 80, 3, 3, include_center=True)
    return sorted(idxs)

# ---------- function producing the figure for each area ----------
    
def make_fragmented_noningot_nogold_hist(area_label: str, area_mask):
    # global filters: OR_3-5, weight>0, fragments, exclusions
    mask_period = d["Or_fin"].isin({"OR_3","OR_4","OR_5"})
    mask_weight = d["Weight_obj"] > 0
    mask_frag   = (d["Fragmented"] > 0) | ((d["Matching_fr"] > 0) & (d["Complete"] < 1))
    mask_excl   = (~d["Artefact"].apply(is_ingot)) & (~d["Artefact"].apply(starts_with_gold))

    sel = d[area_mask & mask_period & mask_weight & mask_frag & mask_excl].copy()

    w = sel["Weight_obj"].dropna().values
    w_top    = w[(w > 0)   & (w <= 100)]
    w_bottom = w[(w > 100) & (w <= 600)]

    # bins
    bw_top = 1.11
    bw_bot = 8.5
    bins_top = np.arange(0,   100 + bw_top, bw_top)
    bins_bot = np.arange(100, 600 + bw_bot, bw_bot)

    # figure
    fig, (ax1, ax2) = plt.subplots(
        2, 1, figsize=FIGSIZE_HALF, dpi=DPI, sharex=False,
        gridspec_kw={"hspace": 0.15}, constrained_layout=True
    )

    # ---------- top panel: 0-100 (with highlighted bands per bin) ----------
    ax1.set_xlim(0, 100)

    # highlighted bands behind the bars (merging adjacent bins)
    idxs = compute_all_highlight_bins(bins_top)
    for a, b in _merge_consecutive(idxs):
        x0, x1 = float(bins_top[a]), float(bins_top[b+1])
        ax1.axvspan(x0, x1, **HIGHLIGHT_STYLE)

    # histogram
    ax1.hist(w_top, bins=bins_top, color=FRAG_GRAY, **EDGE_KW, zorder=3)
    ax1.xaxis.set_major_locator(MultipleLocator(10))
    ax1.xaxis.set_major_formatter(FormatStrFormatter('%d'))
    ax1.set_ylabel("Counts", fontsize=9)
    # headroom
    ytop = ax1.get_ylim()[1]
    ax1.set_ylim(0, max(1, ytop) * (1 + HEADROOM))

    ax1.yaxis.grid(True, color="0.90", lw=0.7)
    ax1.set_axisbelow(True)
    ax1.spines["top"].set_visible(False); ax1.spines["right"].set_visible(False)
    ax1.tick_params(axis="both", labelsize=8)
    ax1.text(0.98, 0.96, f"n = {len(w_top)}", transform=ax1.transAxes,
             ha="right", va="top", fontsize=8)

    # ---------- bottom panel: 100-600 ----------
    ax2.hist(w_bottom, bins=bins_bot, color=FRAG_GRAY, **EDGE_KW)
    ax2.set_xlim(100, 600)
    ax2.xaxis.set_major_locator(MultipleLocator(100))       # tick every 100
    ax2.xaxis.set_major_formatter(FormatStrFormatter('%d'))
    ax2.set_ylabel("Counts", fontsize=9)
    # headroom
    ybot = ax2.get_ylim()[1]
    ax2.set_ylim(0, max(1, ybot) * (1 + HEADROOM))

    ax2.yaxis.grid(True, color="0.90", lw=0.7)
    ax2.set_axisbelow(True)
    ax2.spines["top"].set_visible(False); ax2.spines["right"].set_visible(False)
    ax2.tick_params(axis="both", labelsize=8)
    ax2.set_xlabel("Weight (gr)", fontsize=9)
    ax2.text(0.98, 0.96, f"n = {len(w_bottom)}", transform=ax2.transAxes,
             ha="right", va="top", fontsize=8)

    # single title
    fig.suptitle(f"Fragmented obj. ({area_label} – Or. 3–5)", fontsize=10.5)

    out = f"Fig63_64_65_66_67_Fragmented_obj_{area_label.replace(' ','')}_OR3-5_halfpage.jpg"
    #plt.savefig(out, dpi=DPI, format="jpeg", bbox_inches="tight",
    #            pil_kwargs={"quality":95, "subsampling":0})
    plt.show()
    print(f"Salvato: {out}  |  n_top={len(w_top)}  n_bottom={len(w_bottom)}  n_tot={len(w)}")

# --- generate the 5 separate figures ---
for label, m in AREAS:
    make_fragmented_noningot_nogold_hist(label, m)


# %%

# ============================================================
# FIGURE 68 — Comparison between the frequency distribution of the
# weight of Late Bronze Age fragmented sickles and the frequency
# distribution of fragmented objects (Horizons 3-5).
# ============================================================

# ================== FRAGMENTED - SICKLES vs OTHERS - 1/2 PAGE (ALL AREAS) ==================
    
from matplotlib.ticker import MultipleLocator, FormatStrFormatter
import numpy as np
import re, unicodedata
import matplotlib.pyplot as plt

# --- PARAMETERS layout ---
FIGSIZE_HALF = (6.30, 4.33)    # approx. 160 x 110 mm (half page)
DPI = 300
FRAG_GRAY = "#CFCFCF"
EDGE_KW = dict(edgecolor="#000000", linewidth=0.5)
HEADROOM = 0.25  # +25% headroom above the bars
ALPHA_BAR = 1

# --- highlighting (top panel only) ---
# Highlighting in the top panel: combined rules
HIGHLIGHT_RULES = {
    10: (0, 0),   
    20: (1, 1),   
    30: (2, 1),   
    40: (2, 2),   
    50: (2, 2),
    60: (2, 2),
    80: (3, 3),   
    100: (2, 2),  
    120: (2, 2),
}
HIGHLIGHT_STYLE = dict(facecolor="darkred", alpha=0.30, edgecolor="none", zorder=1)


# --- required columns (reusing the col_like / to_str_u helpers) ---
c_or     = col_like(raw, "Or_fin")
c_state  = col_like(raw, "State")
c_reg    = col_like(raw, "Region")
c_type   = col_like(raw, "Artefact")
c_weight = col_like(raw, "Weight_obj")
c_frag   = col_like(raw, "Fragmented")
c_comp   = col_like(raw, "Complete")
c_match  = col_like(raw, "Matching fr")

# Working DataFrame (no area filter)
d = pd.DataFrame({
    "Or_fin":       to_str_u(raw[c_or]).str.replace(r"\s+","_", regex=True),
    "State":        to_str_u(raw[c_state]),
    "Region":       to_str_u(raw[c_reg]),
    "Artefact":     raw[c_type].astype(str).str.strip(),
    "Weight_obj":   pd.to_numeric(raw[c_weight], errors="coerce"),
    "Fragmented":   pd.to_numeric(raw[c_frag],   errors="coerce").fillna(0),
    "Complete":     pd.to_numeric(raw[c_comp],   errors="coerce").fillna(0),
    "Matching_fr":  pd.to_numeric(raw[c_match],  errors="coerce").fillna(0),
})

# --- category utilities ---
def _strip_accents(s):
    s = str(s).replace("ß","ss")
    return "".join(ch for ch in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(ch))

def _norm(s):
    s = _strip_accents(s).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def is_ingot(s_raw: str) -> bool:
    s = _norm(s_raw)
    if "ringbarren" in s or "spangenbarren" in s:
        return False
    return ("ingot" in s) or ("barren" in s)

def starts_with_gold(s_raw: str) -> bool:
    s = _norm(s_raw)
    return s.startswith("gold ")

def is_sickle(s_raw: str) -> bool:
    s = _norm(s_raw)
    return "sickle" in s

# --- common filters ---
mask_period = d["Or_fin"].isin({"OR_3","OR_4","OR_5"})
mask_weight = d["Weight_obj"] > 0
mask_complete = d["Complete"] < 1
mask_fragcond = (d["Fragmented"] > 0) | (d["Matching_fr"] > 0)

# --- selections top/bottom ---
sel_top = d[mask_period & mask_weight & mask_complete & mask_fragcond & d["Artefact"].apply(is_sickle)].copy()

mask_excl_bottom = (~d["Artefact"].apply(is_sickle)) & (~d["Artefact"].apply(is_ingot)) & (~d["Artefact"].apply(starts_with_gold))
sel_bottom = d[mask_period & mask_weight & mask_complete & mask_fragcond & mask_excl_bottom].copy()

# --- weight vectors and bins ---
bw = 1.11
bins = np.arange(0, 200 + bw, bw)

w_top = sel_top["Weight_obj"].dropna().values
w_top = w_top[(w_top > 0) & (w_top <= 200)]

w_bottom = sel_bottom["Weight_obj"].dropna().values
w_bottom = w_bottom[(w_bottom > 0) & (w_bottom <= 200)]

# --- highlighting helper based on actual bins ---
def _nearest_bin_index(bins, value):
    centers = (bins[:-1] + bins[1:]) / 2.0
    return int(np.argmin(np.abs(centers - float(value))))

def _indices_around(bins, center, n_before, n_after, include_center=True):
    i0 = _nearest_bin_index(bins, center)
    idxs = set([i0] if include_center else [])
    for k in range(1, n_before + 1):
        j = i0 - k
        if j >= 0:
            idxs.add(j)
    for k in range(1, n_after + 1):
        j = i0 + k
        if j <= len(bins) - 2:
            idxs.add(j)
    return idxs

def _merge_consecutive(idxs):
    if not idxs:
        return []
    idxs = sorted(idxs)
    blocks, start, prev = [], idxs[0], idxs[0]
    for i in idxs[1:]:
        if i == prev + 1:
            prev = i
        else:
            blocks.append((start, prev))
            start = prev = i
    blocks.append((start, prev))
    return blocks

def compute_highlight_blocks(bins, rules_dict):
    idxs = set()
    for c, (nb, na) in rules_dict.items():
        idxs |= _indices_around(bins, c, nb, na, include_center=True)
    return _merge_consecutive(idxs)

# --- figure ---
fig, (ax1, ax2) = plt.subplots(
    2, 1, figsize=FIGSIZE_HALF, dpi=DPI, sharex=True,
    gridspec_kw={"hspace": 0.15}, constrained_layout=True
)

# ---------- top panel: Fragmented Sickles ----------
ax1.set_xlim(0, 200)
ax1.xaxis.set_major_locator(MultipleLocator(10))
ax1.xaxis.set_major_formatter(FormatStrFormatter('%d'))
ax1.tick_params(axis='x', which='both', labelbottom=True, bottom=True)  # etichette anche in alto
ax1.set_ylabel("Counts", fontsize=9)

# highlighting: all bins per HIGHLIGHT_RULES, in dark red
for a, b in compute_highlight_blocks(bins, HIGHLIGHT_RULES):
    x0, x1 = float(bins[a]), float(bins[b+1])
    ax1.axvspan(x0, x1, **HIGHLIGHT_STYLE)

ax1.hist(w_top, bins=bins, color=FRAG_GRAY, **EDGE_KW, alpha=ALPHA_BAR)

# headroom
ytop = ax1.get_ylim()[1]
ax1.set_ylim(0, max(1, ytop) * (1 + HEADROOM))

ax1.yaxis.grid(True, color="0.90", lw=0.7)
ax1.set_axisbelow(True)
ax1.spines["top"].set_visible(False); ax1.spines["right"].set_visible(False)
ax1.tick_params(axis="both", labelsize=8)
ax1.set_title("Fragmented Sickles (Or. 3–5)", fontsize=10.5)
ax1.text(0.98, 0.96, f"n = {len(w_top)}", transform=ax1.transAxes, ha="right", va="top", fontsize=8)

# ---------- bottom panel: Fragmented obj., excl. sickles ----------
ax2.set_xlim(0, 200)
ax2.xaxis.set_major_locator(MultipleLocator(10))
ax2.xaxis.set_major_formatter(FormatStrFormatter('%d'))
ax2.set_ylabel("Counts", fontsize=9)

# >>> bands as above, in dark red
for a, b in compute_highlight_blocks(bins, HIGHLIGHT_RULES):
    x0, x1 = float(bins[a]), float(bins[b+1])
    ax2.axvspan(x0, x1, **HIGHLIGHT_STYLE)

ax2.hist(w_bottom, bins=bins, color=FRAG_GRAY, **EDGE_KW, alpha=ALPHA_BAR)

# headroom
ybot = ax2.get_ylim()[1]
ax2.set_ylim(0, max(1, ybot) * (1 + HEADROOM))

ax2.yaxis.grid(True, color="0.90", lw=0.7)
ax2.set_axisbelow(True)
ax2.spines["top"].set_visible(False); ax2.spines["right"].set_visible(False)
ax2.tick_params(axis="both", labelsize=8)
ax2.set_xlabel("Weight (gr)", fontsize=9)
ax2.set_title("Fragmented obj., excl. sickles (Or. 3–5)", fontsize=10.5)
ax2.text(0.98, 0.96, f"n = {len(w_bottom)}", transform=ax2.transAxes, ha="right", va="top", fontsize=8)

# --- export (optional) ---
out = "Fig68_Fragmented_Sickles_vs_others_OR3-5_halfpage.jpg"
#plt.savefig(out, dpi=DPI, format="jpeg", bbox_inches="tight", pil_kwargs={"quality":95, "subsampling":0})
plt.show()
print(f"Top n={len(w_top)} | Bottom n={len(w_bottom)}  | Output: {out}")


# %%

# ============================================================
# FIGURE 69 — Frequency distribution of the weight of fragmented
# objects from the Late Bronze Age (Horizons 3-5), stacked by area
# (0-120 g range), half page with legend below.
# ============================================================

# ================== FRAGMENTED OBJ. - STACKED PER area (0-120 g) - HALF PAGE + LEGEND BELOW ==================
import numpy as np
import pandas as pd
import re, unicodedata
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator, FormatStrFormatter
from matplotlib.patches import Patch, Rectangle
from matplotlib.colors import to_rgba

# --- layout ---
FIGSIZE_HALF = (6.30, 4.33)   # approx. 160 x 110 mm (1/2 page)
DPI = 300
HEADROOM = 0.25               # +25% headroom above the bars
EDGE_KW = dict(edgecolor="#000000", linewidth=0.5)
ALPHA_STACK = 0.35            # fill transparency (edges remain opaque)

# --- palette greys (order A,B,C,CH,Italy) ---
COLS = ["#000000", "#5A5A5A", "#9E9E9E", "#CFCFCF", "#FFFFFF"]
color_list = [to_rgba(c, ALPHA_STACK) for c in COLS]

# --- highlighting (same rules as before) ---
HIGHLIGHT_RULES = {
    10: (0, 0),   20: (1, 1),  30: (2, 1),
    40: (2, 2),   50: (2, 2),  60: (2, 2),
    80: (3, 3),   100: (2, 2), 120: (2, 2),
}
HIGHLIGHT_STYLE = dict(facecolor="darkred", alpha=0.20, edgecolor="none")

# --- normalization / categories ---
def _strip_accents(s):
    s = str(s).replace("ß","ss")
    return "".join(ch for ch in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(ch))
def _norm(s):
    s = _strip_accents(str(s)).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()
def is_ingot(s_raw: str) -> bool:
    s = _norm(s_raw)
    if "ringbarren" in s or "spangenbarren" in s: return False
    return ("ingot" in s) or ("barren" in s)
def starts_with_gold(s_raw: str) -> bool:
    return _norm(s_raw).startswith("gold")

# --- columns from the 'raw' DataFrame (reusing the existing helpers) ---
c_or     = col_like(raw, "Or_fin")
c_state  = col_like(raw, "State")
c_reg    = col_like(raw, "Region")
c_type   = col_like(raw, "Artefact")
c_weight = col_like(raw, "Weight_obj")
c_frag   = col_like(raw, "Fragmented")
c_comp   = col_like(raw, "Complete")
c_match  = col_like(raw, "Matching fr")

d = pd.DataFrame({
    "Or_fin":       to_str_u(raw[c_or]).str.replace(r"\s+","_", regex=True),
    "State":        to_str_u(raw[c_state]),
    "Region":       to_str_u(raw[c_reg]),
    "Artefact":     raw[c_type].astype(str).str.strip(),
    "Weight_obj":   pd.to_numeric(raw[c_weight], errors="coerce"),
    "Fragmented":   pd.to_numeric(raw[c_frag],   errors="coerce").fillna(0),
    "Complete":     pd.to_numeric(raw[c_comp],   errors="coerce").fillna(0),
    "Matching_fr":  pd.to_numeric(raw[c_match],  errors="coerce").fillna(0),
})

# --- global filters ---
mask_period   = d["Or_fin"].isin({"OR_3","OR_4","OR_5"})
mask_weight   = d["Weight_obj"] > 0
mask_complete = d["Complete"] < 1
mask_fragcond = (d["Fragmented"] > 0) | (d["Matching_fr"] > 0)
mask_excl     = (~d["Artefact"].apply(is_ingot)) & (~d["Artefact"].apply(starts_with_gold))
df_filt = d[mask_period & mask_weight & mask_complete & mask_fragcond & mask_excl].copy()

# --- areas ---
U = lambda s: str(s).upper().strip()
GER = U("GERMANY")
mask_A  = (df_filt["Region"].isin({U("BADEN-WÜRTTEMBERG"), U("BAYERN")}) | df_filt["State"].isin({U("AUSTRIA"), U("SLOVENIA")}))
mask_B  = (df_filt["Region"].eq(U("BRANDENBURG")) | df_filt["State"].eq(U("POLAND")) | df_filt["Region"].eq(U("SACHSEN")))
mask_C  = ((df_filt["State"].eq(GER) & df_filt["Region"].isin({U("MECKLENBURG-VORPOMMERN"), U("SCHLESWIG-HOLSTEIN"), U("NIEDERSACHSEN")})) |
           (df_filt["State"].eq(GER) & df_filt["Region"].isin({U("THÜRINGEN"), U("HESSEN"), U("RHEINLAND-PFALZ")})) |
           (df_filt["Region"].eq(U("SACHSEN-ANHALT"))))
mask_CH = df_filt["State"].eq(U("SWITZERLAND"))
mask_I  = df_filt["State"].isin({U("ITALY"), U("SAN MARINO")})
AREAS = [("Area A",mask_A), ("Area B",mask_B), ("Area C",mask_C), ("Switzerland",mask_CH), ("Italy",mask_I)]

# --- weight series per area (A,B,C,CH,Italy) ---
series, tot_n = [], 0
for _, m in AREAS:
    w = df_filt.loc[m, "Weight_obj"].dropna().values
    w = w[(w > 0) & (w <= 120)]
    series.append(w); tot_n += len(w)

# --- bins & conteggi totali ---
bw = 1.11
bins = np.arange(0, 120 + bw, bw)
counts_list  = [np.histogram(s, bins=bins)[0] for s in series]
counts_total = np.sum(np.stack(counts_list), axis=0) if counts_list else np.zeros(len(bins)-1, dtype=int)

# --- helper highlight ---
def _nearest_bin_index(bins, value):
    centers = (bins[:-1] + bins[1:]) / 2.0
    return int(np.argmin(np.abs(centers - float(value))))
def _indices_around(bins, center, n_before, n_after, include_center=True):
    i0 = _nearest_bin_index(bins, center)
    idxs = set([i0] if include_center else [])
    for k in range(1, n_before+1):
        j = i0 - k
        if j >= 0: idxs.add(j)
    for k in range(1, n_after+1):
        j = i0 + k
        if j <= len(bins)-2: idxs.add(j)
    return idxs
def _merge_consecutive(idxs):
    if not idxs: return []
    idxs = sorted(idxs); blocks=[]; start=prev=idxs[0]
    for i in idxs[1:]:
        if i == prev+1: prev=i
        else: blocks.append((start,prev)); start=prev=i
    blocks.append((start,prev)); return blocks
def compute_highlight_blocks(bins, rules_dict):
    idxs=set()
    for c,(nb,na) in rules_dict.items():
        idxs |= _indices_around(bins, c, nb, na, include_center=True)
    return _merge_consecutive(idxs)

# --- figure ---
fig, ax = plt.subplots(1, 1, figsize=FIGSIZE_HALF, dpi=DPI, constrained_layout=True)

# histogram stacked )
ax.hist(series, bins=bins, stacked=True, color=color_list, **EDGE_KW)

# limit axis and headroom
ax.set_xlim(0, 120)
ymax = max(1, counts_total.max()) * (1 + HEADROOM)
ax.set_ylim(0, ymax)

# highlighting ABOVE the bars ONLY (rectangles from the top bin to ymax)
for a, b in compute_highlight_blocks(bins, HIGHLIGHT_RULES):
    for i in range(a, b+1):
        left, right = float(bins[i]), float(bins[i+1])
        base = float(counts_total[i])
        if base < ymax:
            ax.add_patch(Rectangle((left, base), right-left, ymax-base, **HIGHLIGHT_STYLE))

# Axes / style
ax.xaxis.set_major_locator(MultipleLocator(10))
ax.xaxis.set_major_formatter(FormatStrFormatter('%d'))
ax.set_ylabel("Counts", fontsize=9)
ax.set_xlabel("Weight (gr)", fontsize=9)
ax.yaxis.grid(True, color="0.90", lw=0.7)
ax.set_axisbelow(True)
ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
ax.tick_params(axis="both", labelsize=8)

# legend at the bottom (outside the plot)
# --- total n in the top-right corner (inside the axes)
ax.text(0.98, 0.96, f"n = {tot_n}", transform=ax.transAxes,
        ha="right", va="top", fontsize=8)

# --- legend at the bottom, area names ONLY (no title)
legend_handles = [
    Patch(facecolor=to_rgba(c, ALPHA_STACK), edgecolor="#000000",
          linewidth=0.5, label=lbl)
    for (lbl,_), c in zip(AREAS, COLS)
]
fig.legend(
    handles=legend_handles,
    loc="lower center", bbox_to_anchor=(0.5, -0.12),
    ncol=5, frameon=False, fontsize=8,
    handlelength=1.6, handletextpad=0.6, columnspacing=1.0
)

# title
ax.set_title("Fragmented obj. (Or. 3–5)", fontsize=10.5)

# margins to avoid overlapping the legend
ax.set_xlabel("Weight (gr)", fontsize=9, labelpad=6)
fig.set_constrained_layout_pads(w_pad=0.02, h_pad=0.02, hspace=0.02)

# EXPORT
#plt.savefig("Fig69_Fragmented_stacked_areas_OR3-5_halfpage.jpg",
#            dpi=DPI, format="jpeg", bbox_inches="tight",
#            pil_kwargs={"quality":95, "subsampling":0})
plt.show()


# %%

# ============================================================
# FIGURE 70 — Frequency distribution of the weight of Italian and
# Central European balance weights in the shekel range (0-120 g).
# ============================================================

# ================== BALANCE WEIGHTS - HISTOGRAM (0-120 g), NOT STACKED, BLACK BARS ==================
import numpy as np
import pandas as pd
import re, unicodedata
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator, FormatStrFormatter
from matplotlib.patches import Rectangle

# ---------- PARAMETERS ----------
FILE  = "Balance_Weights_2024.xlsx"
SHEET = 0  

FIGSIZE_HALF = (6.30, 4.33)   # approx. 160 x 110 mm (1/2 page)
DPI = 300
HEADROOM = 0.25               # +25% headroom above the bars
EDGE_KW = dict(edgecolor="#000000", linewidth=0.5)

# Highlighting
HIGHLIGHT_RULES = {
    10: (0, 0),  20: (1, 1),  30: (2, 1),
    40: (2, 2),  50: (2, 2),  60: (2, 2),
    80: (3, 3),  100: (2, 2), 120: (2, 2),
}
HIGHLIGHT_STYLE = dict(facecolor="darkred", alpha=0.30, edgecolor="none")

# ---------- HELPER ----------
def _strip_accents(s):
    s = str(s).replace("ß","ss")
    return "".join(ch for ch in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(ch))
def _norm(s):
    s = _strip_accents(str(s)).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def col_like(df, *needles):
    cols = list(df.columns)
    low  = [c.lower() for c in cols]
    for n in needles:
        n_low = str(n).lower()
        for i, c in enumerate(low):
            if n_low in c:
                return cols[i]
    raise KeyError(f"Column not found for: {needles}")

def _nearest_bin_index(bins, value):
    centers = (bins[:-1] + bins[1:]) / 2.0
    return int(np.argmin(np.abs(centers - float(value))))

def _indices_around(bins, center, n_before, n_after, include_center=True):
    i0 = _nearest_bin_index(bins, center)
    idxs = set([i0] if include_center else [])
    for k in range(1, n_before+1):
        j = i0 - k
        if j >= 0: idxs.add(j)
    for k in range(1, n_after+1):
        j = i0 + k
        if j <= len(bins)-2: idxs.add(j)
    return idxs

def _merge_consecutive(idxs):
    if not idxs: return []
    idxs = sorted(idxs); blocks=[]; start=prev=idxs[0]
    for i in idxs[1:]:
        if i == prev+1: prev=i
        else: blocks.append((start,prev)); start=prev=i
    blocks.append((start,prev)); return blocks

def compute_highlight_blocks(bins, rules_dict):
    idxs=set()
    for c,(nb,na) in rules_dict.items():
        idxs |= _indices_around(bins, c, nb, na, include_center=True)
    return _merge_consecutive(idxs)

# ---------- READING ----------
bw_raw = pd.read_excel(FILE, sheet_name=SHEET)

# columns 
c_mass    = col_like(bw_raw, "Mass in g (complete/after reconstruction)", "mass in g (complete)", "mass in g", "mass", "weight")
c_country = None
for cand in ("country", "nation", "state"):
    try:
        c_country = col_like(bw_raw, cand)
        break
    except Exception:
        pass
c_typo = None
for cand in ("typology", "type", "form", "category"):
    try:
        c_typo = col_like(bw_raw, cand)
        break
    except Exception:
        pass

# Working DataFrame
d = pd.DataFrame({"Mass_g": pd.to_numeric(bw_raw[c_mass], errors="coerce")})
if c_country is not None:
    d["Country"] = bw_raw[c_country].astype(str)
else:
    d["Country"] = ""
if c_typo is not None:
    d["Typology"] = bw_raw[c_typo].astype(str)
else:
    d["Typology"] = ""

# ---------- FILTERS ----------
EXCL_TYPO = {"kannelurenstein", "piriform", "other hanging"}
mask_typo_ok = ~d["Typology"].map(lambda s: _norm(s) in EXCL_TYPO) if c_typo is not None else True

# Ignore countries
EXCL_COUNTRIES = {"ENGLAND", "PORTUGAL", "SPAIN"}
mask_ctry_ok = ~d["Country"].map(lambda s: str(s).strip().upper() in EXCL_COUNTRIES) if c_country is not None else True

# Valid weights (0-120]
mask_weight = d["Mass_g"].notna() & (d["Mass_g"] > 0)

df = d[mask_typo_ok & mask_ctry_ok & mask_weight].copy()

# ---------- DATA FOR THE HISTOGRAM ----------
w = df["Mass_g"].dropna().values
w = w[(w > 0) & (w <= 120)]

bw   = 1.11
bins = np.arange(0, 120 + bw, bw)
counts, _ = np.histogram(w, bins=bins)

# ---------- figure ----------
fig, ax = plt.subplots(1, 1, figsize=FIGSIZE_HALF, dpi=DPI, constrained_layout=True)

ax.hist(w, bins=bins, color="#000000", edgecolor="#FFFFFF", linewidth=0.6)

# Limits & headroom
ax.set_xlim(0, 120)
ymax = max(1, counts.max()) * (1 + HEADROOM)
ax.set_ylim(0, ymax)

# Highlighting ABOVE the bars ONLY
for a, b in compute_highlight_blocks(bins, HIGHLIGHT_RULES):
    for i in range(a, b+1):
        left, right = float(bins[i]), float(bins[i+1])
        base = float(counts[i])
        if base < ymax:
            ax.add_patch(Rectangle((left, base), right-left, ymax-base, **HIGHLIGHT_STYLE))

# Axes / style
ax.xaxis.set_major_locator(MultipleLocator(10))
ax.xaxis.set_major_formatter(FormatStrFormatter('%d'))
ax.set_ylabel("Counts", fontsize=9)
ax.set_xlabel("Weight (g)", fontsize=9)
ax.yaxis.grid(True, color="0.90", lw=0.7)
ax.set_axisbelow(True)
ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
ax.tick_params(axis="both", labelsize=8)

# n in the top-right corner
ax.text(0.98, 0.96, f"n = {len(w)}", transform=ax.transAxes,
        ha="right", va="top", fontsize=8)

# title (adjust as preferred)
ax.set_title("Balance weights (0–120 g)", fontsize=10.5)

# Export (optional)
# plt.savefig("Fig70_Balance_weights_hist_black_halfpage.jpg", dpi=DPI, format="jpeg", bbox_inches="tight",
#             pil_kwargs={"quality":95, "subsampling":0})
plt.show()

# Riepilogo
print(f"n total: {len(w)}  |  excluded by country/type/weight limits: {len(d) - len(df)}")


# %%

# ============================================================
# SUPPORTING TABLE (not a figure) — Summary statistics for
# intact vs fragmented objects/ingots by area and period, with
# estimates for unweighed items, exported to Excel and Word.
# ============================================================

# load the main DB (the one with Or_fin / State / Region)
raw_db = pd.read_excel("DB.xlsx", sheet_name="DB")        # o "DB_SI.xlsx"

# load the balance weights (the one with Country / Typology)
bw = pd.read_excel("Balance_Weights_2024.xlsx")


# ================== OBJECTS/INGOTS - INTACT vs FRAGMENTED - ESTIMATES + EXPORT ==================
import pandas as pd, numpy as np, re, unicodedata
from datetime import datetime

# ---- fallback helpers if missing ----
def _ensure_helper_funcs():
    g = globals()
    if "col_like" not in g:
        def col_like(df, key):  # fallback minimale
            return key
        g["col_like"] = col_like
    if "to_str_u" not in g:
        def to_str_u(s):
            return s.astype(str).str.upper().str.strip()
        g["to_str_u"] = to_str_u
_ensure_helper_funcs()

# ---- columns from 'raw' ----
c_or     = col_like(raw, "Or_fin")
c_state  = col_like(raw, "State")
c_reg    = col_like(raw, "Region")
c_type   = col_like(raw, "Artefact")
c_weight = col_like(raw, "Weight_obj")
c_frag   = col_like(raw, "Fragmented")
c_comp   = col_like(raw, "Complete")
c_match  = col_like(raw, "Matching fr")

# ---- normalization / categories ----
def _strip_accents(s):
    s = str(s).replace("ß","ss")
    return "".join(ch for ch in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(ch))
def _norm(s):
    s = _strip_accents(str(s)).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def is_ingot(s_raw: str) -> bool:
    s = _norm(s_raw)
    # Ring-/Spangenbarren do NOT count as ingots
    if "ringbarren" in s or "spangenbarren" in s:
        return False
    return ("ingot" in s) or ("barren" in s)

def starts_with_gold(s_raw: str) -> bool:
    return _norm(s_raw).startswith("gold")

# ---- Working DataFrame ----
d = pd.DataFrame({
    "Or_fin":      to_str_u(raw[c_or]).str.replace(r"\s+","_", regex=True),
    "State":       to_str_u(raw[c_state]),
    "Region":      to_str_u(raw[c_reg]),
    "Artefact":    raw[c_type].astype(str).str.strip(),
    "Weight_obj":  pd.to_numeric(raw[c_weight], errors="coerce"),
    "Fragmented":  pd.to_numeric(raw[c_frag] , errors="coerce").fillna(0),
    "Complete":    pd.to_numeric(raw[c_comp] , errors="coerce").fillna(0),
    "Matching_fr": pd.to_numeric(raw[c_match], errors="coerce").fillna(0),
})

# ---- map periods (OR_4 and OR_5 together) ----
def map_period(or_fin):
    if or_fin in {"OR_1","OR_2","OR_3"}: return or_fin
    if or_fin in {"OR_4","OR_5"}:        return "OR_4_5"
    return np.nan

d["PER"] = d["Or_fin"].map(map_period)
d = d.dropna(subset=["PER"])

# ---- areas (as defined) ----
U = lambda s: str(s).upper().strip()
GER = U("GERMANY")
mask_A  = (d["Region"].isin({U("BADEN-WÜRTTEMBERG"), U("BAYERN")}) |
           d["State"].isin({U("AUSTRIA"), U("SLOVENIA")}))
mask_B  = (d["Region"].eq(U("BRANDENBURG")) |
           d["State"].eq(U("POLAND")) |
           d["Region"].eq(U("SACHSEN")))
mask_C  = ((d["State"].eq(GER) & d["Region"].isin({U("MECKLENBURG-VORPOMMERN"), U("SCHLESWIG-HOLSTEIN"), U("NIEDERSACHSEN")})) |
           (d["State"].eq(GER) & d["Region"].isin({U("THÜRINGEN"), U("HESSEN"), U("RHEINLAND-PFALZ")})) |
           (d["Region"].eq(U("SACHSEN-ANHALT"))))
mask_CH = d["State"].eq(U("SWITZERLAND"))
mask_I  = d["State"].isin({U("ITALY"), U("SAN MARINO")})

AREE = [
    ("Area A",      mask_A),
    ("Area B",      mask_B),
    ("Area C",      mask_C),
    ("Switzerland", mask_CH),
    ("Italy",       mask_I),
]

# ---- definitions  ----
# Fragmented: Fragmented > 0  or  (Matching_fr > 0 and Complete < 1)
is_fragmented = (d["Fragmented"] > 0) | ((d["Matching_fr"] > 0) & (d["Complete"] < 1))
# Intact: Complete >= 1 and not fragmented
is_intact     = (d["Complete"] >= 1) & (~is_fragmented)

# Objects = everything except ingot & Gold*
is_obj  = (~d["Artefact"].apply(is_ingot)) & (~d["Artefact"].apply(starts_with_gold))
is_ing  = ( d["Artefact"].apply(is_ingot))

# ---- labelled periods (Or. 1, 2, 3, 4-5) ----
PERIODI = [("Or. 1","OR_1"), ("Or. 2","OR_2"), ("Or. 3","OR_3"), ("Or. 4–5","OR_4_5")]
period_labels = [p for p,_ in PERIODI]

# ---- metrics function on a subset ----
def _metrics(w):
    # measured: >0 ; not measured: NaN o 0
    w = pd.to_numeric(w, errors="coerce")
    measured = w > 0
    missing  = w.isna() | (w == 0)
    s = float(w[measured].sum())
    m = float(w[measured].mean()) if measured.any() else np.nan
    c_meas = int(measured.sum())
    c_miss = int(missing.sum())
    return s, m, c_meas, c_miss

# ---- build table: Period x Group x Metric ----
groups = [
    ("Objects — Intact",      is_obj  & is_intact),
    ("Objects — Fragmented",  is_obj  & is_fragmented),
    ("Ingots — Intact",       is_ing  & is_intact),
    ("Ingots — Fragmented",   is_ing  & is_fragmented),
]
metrics_base = ["Sum","Mean","Count >0","Count 0/blank"]
metrics_all  = metrics_base + ["Stima non pesati", "Totale incl. stima"]

cols = pd.MultiIndex.from_product([period_labels, [g for g,_ in groups], metrics_all],
                                  names=["Period","Group","Metric"])
index_labels = [lbl for lbl,_ in AREE]
stats_frag_df = pd.DataFrame(index=index_labels, columns=cols, dtype="float")

# base metrics
for area_label, area_mask in AREE:
    df_area = d[area_mask].copy()
    for p_lab, p_key in PERIODI:
        df_p = df_area[df_area["PER"].eq(p_key)]
        for g_label, g_mask in groups:
            sel = df_p[g_mask.reindex(df_p.index, fill_value=False)]
            s, m, c_meas, c_miss = _metrics(sel["Weight_obj"])
            stats_frag_df.loc[area_label, (p_lab, g_label, "Sum")]           = s
            stats_frag_df.loc[area_label, (p_lab, g_label, "Mean")]          = m
            stats_frag_df.loc[area_label, (p_lab, g_label, "Count >0")]      = c_meas
            stats_frag_df.loc[area_label, (p_lab, g_label, "Count 0/blank")] = c_miss

# estimates (Mean x Count 0/blank) and Total incl. estimate
for p_lab in period_labels:
    for g_label, _ in groups:
        mean_vals = pd.to_numeric(stats_frag_df[(p_lab, g_label, "Mean")], errors="coerce").fillna(0.0)
        miss_vals = pd.to_numeric(stats_frag_df[(p_lab, g_label, "Count 0/blank")], errors="coerce").fillna(0.0)
        sum_vals  = pd.to_numeric(stats_frag_df[(p_lab, g_label, "Sum")], errors="coerce").fillna(0.0)

        est_vals = (mean_vals * miss_vals)
        stats_frag_df[(p_lab, g_label, "Stima non pesati")] = est_vals.round(0).astype("Int64")
        tot_vals = sum_vals + est_vals
        stats_frag_df[(p_lab, g_label, "Totale incl. stima")] = tot_vals.round(0).astype("Int64")

# ---- formattazione numerica (arrotonda / tipi) ----
metric_lv   = stats_frag_df.columns.get_level_values("Metric")
sum_like    = metric_lv.isin(["Sum","Stima non pesati","Totale incl. stima"])
mean_mask   = (metric_lv == "Mean")
counts_mask = metric_lv.isin(["Count >0", "Count 0/blank"])

stats_frag_df.loc[:, sum_like]    = stats_frag_df.loc[:, sum_like].round(0).astype("Int64")
stats_frag_df.loc[:, mean_mask]   = stats_frag_df.loc[:, mean_mask].round(2)
stats_frag_df.loc[:, counts_mask] = stats_frag_df.loc[:, counts_mask].astype("Int64")

# ================== EXPORT EXCEL (XlsxWriter) ==================
def export_frag_excel(df: pd.DataFrame, path: str = "stats_FRAG_intact_by_area_period.xlsx"):
    import xlsxwriter  
    flat = df.copy()
    flat.columns = [f"{p} | {g} | {m}" for (p,g,m) in flat.columns]

    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        # MultiIndex
        df.to_excel(writer, sheet_name="MultiIndex")
        ws1 = writer.sheets["MultiIndex"]
        ws1.freeze_panes(1, 1)
        ws1.set_column(0, 0, 18)
        ws1.set_column(1, 2000, 15)

        # Flat
        flat.to_excel(writer, sheet_name="Flat", index=True)
        ws2 = writer.sheets["Flat"]
        ws2.freeze_panes(1, 1)
        ws2.autofilter(0, 0, flat.shape[0], flat.shape[1])
        ws2.set_column(0, 0, 18)

        wb      = writer.book
        fmt_hdr = wb.add_format({"bold": True})
        fmt_int = wb.add_format({"num_format": "#,##0"})
        fmt_d2  = wb.add_format({"num_format": "#,##0.00"})
        ws2.set_row(0, None, fmt_hdr)

        for j, col in enumerate(flat.columns, start=1):
            metric = col.split(" | ")[-1]
            if metric in ("Sum","Count >0","Count 0/blank","Stima non pesati","Totale incl. stima"):
                ws2.set_column(j, j, 16, fmt_int)
            elif metric == "Mean":
                ws2.set_column(j, j, 14, fmt_d2)
            else:
                ws2.set_column(j, j, 14)

    print(f"Excel scritto: {path}")

try:
    export_frag_excel(stats_frag_df)
except PermissionError:
    alt = f"stats_FRAG_intact_by_area_period_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    export_frag_excel(stats_frag_df, path=alt)

# ================== EXPORT WORD (.docx) - Areas as columns ==================
try:
    from docx import Document
    from docx.shared import Pt
except ModuleNotFoundError:
    print("python-docx non installato: salta export Word. Installa con: pip install python-docx")
else:
    # reorganize: columns = Areas, rows = Period | Group | Metric (estimates included)
    area_order = ["Area A","Area B","Area C","Switzerland","Italy"]
    areas = [a for a in area_order if a in stats_frag_df.index.tolist()] or stats_frag_df.index.tolist()

    period_labels = list(dict.fromkeys(stats_frag_df.columns.get_level_values("Period")))
    groups_out    = list(dict.fromkeys(stats_frag_df.columns.get_level_values("Group")))
    metrics_out   = ["Sum","Mean","Count >0","Count 0/blank","Stima non pesati","Totale incl. stima"]

    rows = []
    for p in period_labels:
        for g in groups_out:
            for m in metrics_out:
                label = f"{p} | {g} | {m}"
                vals = [stats_frag_df.loc[a, (p, g, m)] if (p,g,m) in stats_frag_df.columns else np.nan
                        for a in areas]
                rows.append([label] + vals)

    def _fmt_cell(info, x):
        if pd.isna(x): return ""
        metric = info.split(" | ")[-1]
        if metric == "Mean":
            return f"{float(x):,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
        else:
            try:
                xi = int(pd.to_numeric(x))
                return f"{xi:,}".replace(",", ".")
            except Exception:
                return str(x)

    table_df = pd.DataFrame(rows, columns=["Info"] + areas)
    for col in areas:
        table_df[col] = [_fmt_cell(info, val) for info, val in zip(table_df["Info"], table_df[col])]

    doc = Document()
    doc.add_heading("Objects / Ingots - Intact vs Fragmented - with Estimates", level=1)
    rows_n, cols_n = table_df.shape
    table = doc.add_table(rows=rows_n+1, cols=cols_n)
    table.style = "Light List"

    # header
    hdr = table.rows[0].cells
    hdr[0].text = "Info"
    for j, col in enumerate(table_df.columns[1:], start=1):
        hdr[j].text = col
    for c in table.rows[0].cells:
        for p in c.paragraphs:
            for r in p.runs:
                r.font.bold = True
                r.font.size = Pt(10)

    # body
    for i in range(rows_n):
        cells = table.rows[i+1].cells
        cells[0].text = str(table_df.iloc[i, 0])
        for j in range(1, cols_n):
            cells[j].text = str(table_df.iloc[i, j])

    out_docx = "stats_FRAG_intact_by_area_period.docx"
    doc.save(out_docx)
    print(f"Documento Word creato: {out_docx}")


# %%

# ============================================================
# FIGURE 38 — Published and estimated quantity of metal deposited
# in Bronze Age hoards from Central Europe. Stacked grid of
# measured/estimated weights (kg), 3/4 page.
# ============================================================

# ================== GRID - WEIGHTS (kg) MEASURED & ESTIMATED (STACKED) - 3/4 PAGE ==================
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import Patch

# ---------- prerequisite ----------
if "stats_frag_df" not in globals():
    raise NameError("Missing 'stats_frag_df'. Run the cell that creates it first.")

# ---------- CHART PARAMETERS ----------
FIGSIZE = (6.30, 6.94)   # ~ 3/4 page (160 x 176 mm)
DPI = 300
HEADROOM = 0.25
BAR_W = 0.68

COL_COMPLETE = "#000000"   # black for complete
COL_FRAG     = "#CFCFCF"   # light gray for fragments
EDGE_NONE    = dict(edgecolor="none", linewidth=0.0)

area_order   = ["Area A", "Area B", "Area C", "Switzerland", "Italy"]
period_order = ["Or. 1", "Or. 2", "Or. 3", "Or. 4–5"]

# helper: extract a value (in kg) from stats_frag_df
def Vkg(area, per, group_label, metric):
    key = (per, group_label, metric)
    val = float(stats_frag_df.loc[area, key]) if key in stats_frag_df.columns else 0.0
    return val / 1000.0  # g -> kg

groups = {
    "obj_meas_int": ("Objects — Intact",     "Sum"),
    "obj_meas_frag":("Objects — Fragmented", "Sum"),
    "obj_est_int":  ("Objects — Intact",     "Stima non pesati"),
    "obj_est_frag": ("Objects — Fragmented", "Stima non pesati"),
    "ing_meas_int": ("Ingots — Intact",      "Sum"),
    "ing_meas_frag":("Ingots — Fragmented",  "Sum"),
    "ing_est_int":  ("Ingots — Intact",      "Stima non pesati"),
    "ing_est_frag": ("Ingots — Fragmented",  "Stima non pesati"),
}

# ---------- figure ----------
fig, axes = plt.subplots(len(area_order), len(period_order),
                         figsize=FIGSIZE, dpi=DPI, sharex=True, sharey=False)
axes = np.atleast_2d(axes)

xticks = np.arange(4)
xtick_labels = ["Objects' weight", "Est. obj. weight", "Ingots' weight", "Est. ing. weight"]

for r, area in enumerate(area_order):
    if area not in stats_frag_df.index:
        for c in range(len(period_order)):
            axes[r, c].axis("off")
        continue

    for c, per in enumerate(period_order):
        ax = axes[r, c]

        base = [
            Vkg(area, per, *groups["obj_meas_int"]),
            Vkg(area, per, *groups["obj_est_int"]),
            Vkg(area, per, *groups["ing_meas_int"]),
            Vkg(area, per, *groups["ing_est_int"]),
        ]
        top = [
            Vkg(area, per, *groups["obj_meas_frag"]),
            Vkg(area, per, *groups["obj_est_frag"]),
            Vkg(area, per, *groups["ing_meas_frag"]),
            Vkg(area, per, *groups["ing_est_frag"]),
        ]

        # measured bars (0,2)
        ax.bar(xticks[[0,2]], [base[0], base[2]], BAR_W, color=COL_COMPLETE, zorder=3, **EDGE_NONE)
        ax.bar(xticks[[0,2]], [top[0],  top[2]],  BAR_W, bottom=[base[0], base[2]],
               color=COL_FRAG, zorder=3, **EDGE_NONE)

        # estimated bars (1,3)
        ax.bar(xticks[[1,3]], [base[1], base[3]], BAR_W, color=COL_COMPLETE, zorder=3, **EDGE_NONE)
        ax.bar(xticks[[1,3]], [top[1],  top[3]],  BAR_W, bottom=[base[1], base[3]],
               color=COL_FRAG, zorder=3, **EDGE_NONE)

        # axes style
        ymax = max([b+t for b, t in zip(base, top)] + [1.0])
        ax.set_ylim(0, ymax * (1 + HEADROOM))
        ax.yaxis.grid(True, color="0.90", lw=0.7)
        ax.set_axisbelow(True)
        ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
        ax.tick_params(axis="y", labelsize=7)

        # x-ticks sempre; etichette solo all'ultima riga
        ax.set_xticks(xticks)
        if r == len(area_order) - 1:
            ax.set_xticklabels(xtick_labels, fontsize=7, rotation=45, ha="right")
        else:
            ax.set_xticklabels([])

# Y label only on the first column (all equal and aligned)
YLABEL_TEXT = "Weight (kg)"
YLABEL_FS   = 6
YLABEL_X    = -0.45
for r in range(len(area_order)):
    ax = axes[r, 0]
    ax.set_ylabel(YLABEL_TEXT, fontsize=YLABEL_FS)
    ax.yaxis.set_label_coords(YLABEL_X, 0.5)

# legend
handles = [
    Patch(facecolor=COL_COMPLETE, edgecolor="none", label="Complete"),
    Patch(facecolor=COL_FRAG,     edgecolor="none", label="Fragmented"),
]
fig.legend(handles=handles, loc="lower center", bbox_to_anchor=(0.5, 0.02),
           ncol=2, frameon=False, fontsize=9)

# layout with extra room at the top for the period titles
plt.tight_layout(pad=1.4, w_pad=1.2, h_pad=1.25, rect=[0.12, 0.06, 0.98, 0.90])
fig.canvas.draw()  # compute the final bounding boxes

# ---------- (A) area names: positioned higher ----------
AREA_LABEL_XOFFSET = 0.20   # larger = further left
AREA_LABEL_YSHIFT  = 0.01  # larger = further up

x_left = min(axes[r, 0].get_position().x0 for r in range(len(area_order))) - AREA_LABEL_XOFFSET
x_left = max(x_left, 0.01) 

for r, area in enumerate(area_order):
    y_top = max(axes[r, c].get_position().y1 for c in range(len(period_order)))
    fig.text(x_left, y_top + AREA_LABEL_YSHIFT, area,
             ha="left", va="bottom", fontsize=10)

# ---------- (B) Period titles above each column ----------
TITLE_Y_OFFSET = 0.02
for c, per in enumerate(period_order):
    pos = axes[0, c].get_position()
    fig.text(pos.x0 + pos.width/2, pos.y1 + TITLE_Y_OFFSET,
             per, ha="center", va="bottom", fontsize=12)

# EXPORT
out = "Fig38_Grid_weights_measured_estimated_by_area_period_three_quarters_kg_with_titles.jpg"
#plt.savefig(out, dpi=DPI, format="jpeg", bbox_inches="tight",
#            pil_kwargs={"quality":95, "subsampling":0})
plt.show()
print(f"Salvato: {out}")


# %%

# ============================================================
# FIGURE 39 — Average weight per year of metal hoards
# Duration of periods: Horizon 1 = 600 years (2150-1550 BCE);
# Horizon 2 = 220 years (1550-1330 BCE); Horizon 3 = 230 years
# (1330-1100 BCE); Horizon 4-5 = 300 years (1100-800 BCE).
# ============================================================

# ================== ROW (1/3 PAGE): WEIGHT PER YEAR (g/yr), MEASURED vs ESTIMATED - + ORANGE LINE ==================
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from glob import glob

# ---------- try to load stats_frag_df if it's not already in memory ----------
def _load_stats_frag_df():
    if "stats_frag_df" in globals():
        return globals()["stats_frag_df"]

    # preferred filename + any timestamped variants
    candidates = ["stats_FRAG_intact_by_area_period.xlsx"] + \
                 sorted(glob("stats_FRAG_intact_by_area_period*.xlsx"), reverse=True)
    for path in candidates:
        try:
            # try the MultiIndex sheet first (already hierarchical)
            df = pd.read_excel(path, sheet_name="MultiIndex", header=[0,1,2], index_col=0)
            print(f"[INFO] Caricato: {path} (sheet: MultiIndex)")
            return df
        except Exception:
            try:
                # fallback: rebuild MultiIndex from the "Flat" sheet
                flat = pd.read_excel(path, sheet_name="Flat", index_col=0)
                tuples = [tuple(map(str.strip, c.split("|"))) for c in flat.columns]
                df = flat.copy()
                df.columns = pd.MultiIndex.from_tuples(tuples, names=["Period","Group","Metric"])
                print(f"[INFO] Loaded: {path} (sheet: Flat -> MultiIndex rebuilt)")
                return df
            except Exception:
                continue
    raise FileNotFoundError(
        "Could not find 'stats_frag_df' in memory or a 'stats_FRAG_intact_by_area_period*.xlsx' file on disk."
    )

stats_frag_df = _load_stats_frag_df()

# ---------- setup ----------
AREAS   = ["Area A", "Area B", "Area C", "Switzerland", "Italy"]
PERIODS = ["Or. 1", "Or. 2", "Or. 3", "Or. 4–5"]

# Durations (years) used for the weight/year calculation
YEARS   = {"Or. 1": 600, "Or. 2": 220, "Or. 3": 230, "Or. 4–5": 350}

GROUPS   = ["Objects — Intact", "Objects — Fragmented", "Ingots — Intact", "Ingots — Fragmented"]
MET_MEAS = "Sum"
MET_EST  = "Stima non pesati"

# style
FIGSIZE   = (6.30, 3.08)   # ~ 1/3 page
DPI       = 300
BAR_W     = 0.60
HEADROOM  = 0.25           # a bit more room at the top
COL_MEAS  = "#000000"      # bottom = measured
COL_EST   = "#CFCFCF"      # top    = estimated
LINE_COL  = "orange"       # linea di collegamento
EDGE_NONE = dict(edgecolor="none", linewidth=0.0)

# ---------- helpers ----------
def _safe_float(x):
    try:
        return float(x)
    except Exception:
        return 0.0

def get_g(area, per, group, metric):
    key = (per, group, metric)
    try:
        return _safe_float(stats_frag_df.loc[area, key])
    except Exception:
        return 0.0

# ---------- data in g/year ----------
data_meas_py = {area: [] for area in AREAS}
data_est_py  = {area: [] for area in AREAS}

for area in AREAS:
    for per in PERIODS:
        g_meas = sum(get_g(area, per, g, MET_MEAS) for g in GROUPS)
        g_est  = sum(get_g(area, per, g, MET_EST)  for g in GROUPS)
        yrs = YEARS[per]
        data_meas_py[area].append(g_meas / yrs if yrs else 0.0)
        data_est_py[area].append(g_est  / yrs if yrs else 0.0)

# scala Y comune
global_ymax = 1.0
for area in AREAS:
    s = np.array(data_meas_py[area], dtype=float) + np.array(data_est_py[area], dtype=float)
    if s.size:
        global_ymax = max(global_ymax, float(np.nanmax(s)))

# ---------- plot ----------
fig, axes = plt.subplots(1, len(AREAS), figsize=FIGSIZE, dpi=DPI, sharey=True)
x = np.arange(len(PERIODS))

for i, area in enumerate(AREAS):
    ax = axes[i]
    meas = np.array(data_meas_py[area], dtype=float)  # g/yr
    est  = np.array(data_est_py[area],  dtype=float)  # g/yr
    tot  = meas + est

    # stacked bars
    ax.bar(x, meas, BAR_W, color=COL_MEAS, zorder=3, **EDGE_NONE)
    ax.bar(x, est,  BAR_W, bottom=meas, color=COL_EST, zorder=3, **EDGE_NONE)

    # orange line connecting the bar tops
    ax.plot(x, tot, color=LINE_COL, linewidth=0.9, zorder=4)

    ax.set_ylim(0, global_ymax * (1 + HEADROOM))
    ax.set_xticks(x)
    ax.set_xticklabels(PERIODS, rotation=30, ha="right", fontsize=7)
    ax.yaxis.grid(True, color="0.90", lw=0.7)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
    ax.set_title(area, fontsize=9, pad=4)

# y label 
axes[0].set_ylabel("Weight per year (g)", fontsize=8)

# legend
handles = [Patch(facecolor=COL_MEAS, label="Measured (g/yr)"),
           Patch(facecolor=COL_EST,  label="Estimated (g/yr)")]
fig.legend(handles=handles, loc="lower center", bbox_to_anchor=(0.5, 0.02),
           ncol=2, frameon=False, fontsize=8, handlelength=1.8, handletextpad=0.6)

plt.tight_layout(pad=0.8, w_pad=0.9, rect=[0.05, 0.10, 0.98, 0.95])
#plt.savefig("Fig39_Row_weight_per_year_meas_vs_est_with_line_thirdpage.jpg",
#             dpi=DPI, format="jpeg", bbox_inches="tight",
#             pil_kwargs={"quality":95, "subsampling":0})
plt.show()


# %%

# ============================================================
# CQA DATA PREPARATION (no direct figure)
# Helper functions and series preparation feeding the CQA (Cosine
# Quantogram Analysis) plots used in Figures 49, 55, 62, 71.
# ============================================================

# ============== PREPARE SERIES for CQA (areas as defined from State/Region columns) ==============
import pandas as pd
import numpy as np
import re, unicodedata
import math
try:
    from tqdm import tqdm
except ImportError:
    # silent fallback if tqdm is not available
    def tqdm(x, **k): 
        return x


# ---------- helper robusti ----------
def _strip_accents(s):
    s = str(s).replace("ß","ss")
    return "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))

def _norm(s):
    s = _strip_accents(str(s)).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def to_str_u(s):  # upper + strip
    return s.astype(str).str.upper().str.strip()

def col_like(df, *needles):
    """Find a column by name/partial match (case-insensitive)."""
    cols = list(df.columns)
    lows = [c.lower().strip() for c in cols]
    for n in needles:
        n_low = str(n).lower().strip()
        for i, c in enumerate(lows):
            if (n_low == c) or (n_low in c):
                return cols[i]
    raise KeyError(f"Column not found for: {needles}")

def is_ingot(s_raw: str) -> bool:
    s = _norm(s_raw)
    # Ring-/Spangenbarren EXCLUDED from ingots
    if "ringbarren" in s or "spangenbarren" in s:
        return False
    return ("ingot" in s) or ("barren" in s)

def starts_with_gold(s_raw: str) -> bool:
    return _norm(s_raw).startswith("gold")

U = lambda s: str(s).upper().strip()

# ---------- data source ----------
if "raw" in globals():
    df_src = raw.copy()
else:
    # fallback: load a DB with State/Region (NOT the balance weights)
    for fname in ("DB.xlsx", "DB_SI.xlsx"):
        try:
            df_src = pd.read_excel(fname, sheet_name="DB")
            print(f"[INFO] Caricato: {fname}")
            break
        except Exception:
            df_src = None
    if df_src is None:
        raise FileNotFoundError("Could not find 'raw' in memory or the DB.xlsx / DB_SI.xlsx files on disk.")

# ---------- main column mapping ----------
# Or_fin (accepts variants with/without underscores/spaces)
c_or     = col_like(df_src, "Or_fin", "or fin", "period", "or_fin")
c_state  = col_like(df_src, "State")
c_reg    = col_like(df_src, "Region")
c_type   = col_like(df_src, "Artefact", "artifact", "type")
c_weight = col_like(df_src, "Weight_obj", "weight obj", "weight_obj", "weight", "mass", "weight in g", "mass in g")
c_frag   = col_like(df_src, "Fragmented", "fragmented")
c_comp   = col_like(df_src, "Complete", "complete")
c_match  = col_like(df_src, "Matching fr", "matching_fr", "matching fr.", "matching")

# ---------- normalized dataframe for filtering ----------
d = pd.DataFrame({
    "Or_fin":      to_str_u(df_src[c_or]).str.replace(r"\s+", "_", regex=True),
    "State":       to_str_u(df_src[c_state]),
    "Region":      to_str_u(df_src[c_reg]),
    "Artefact":    df_src[c_type].astype(str).str.strip(),
    "Weight_obj":  pd.to_numeric(df_src[c_weight], errors="coerce"),
    "Fragmented":  pd.to_numeric(df_src[c_frag],  errors="coerce").fillna(0),
    "Complete":    pd.to_numeric(df_src[c_comp],  errors="coerce").fillna(0),
    "Matching_fr": pd.to_numeric(df_src[c_match], errors="coerce").fillna(0),
})

# ---------- FILTERS as specified ----------
# periods: OR_3, OR_4, OR_5
mask_period = d["Or_fin"].isin({"OR_3","OR_4","OR_5"})
# weight range 7-200 g (7 included, 200 excluded)
mask_range  = d["Weight_obj"].between(7, 200, inclusive="left")
# keep EVERYTHING except Gold* (ingots included)
mask_keep   = ~d["Artefact"].apply(starts_with_gold)
# fragments: Fragmented>0  or  (Matching_fr>0 and Complete<1)
mask_frag   = (d["Fragmented"] > 0) | ((d["Matching_fr"] > 0) & (d["Complete"] < 1))

filt_common = mask_period & mask_range & mask_keep & mask_frag

# ---------- areas (EXACTLY as in the original snippet) ----------
GER = U("GERMANY")
mask_A  = (d["Region"].isin({U("BADEN-WÜRTTEMBERG"), U("BAYERN")}) |
           d["State"].isin({U("AUSTRIA"), U("SLOVENIA")}))
mask_B  = (d["Region"].eq(U("BRANDENBURG")) |
           d["State"].eq(U("POLAND")) |
           d["Region"].eq(U("SACHSEN")))
mask_C  = ((d["State"].eq(GER) & d["Region"].isin({U("MECKLENBURG-VORPOMMERN"), U("SCHLESWIG-HOLSTEIN"), U("NIEDERSACHSEN")})) |
           (d["State"].eq(GER) & d["Region"].isin({U("THÜRINGEN"), U("HESSEN"), U("RHEINLAND-PFALZ")})) |
           (d["Region"].eq(U("SACHSEN-ANHALT"))))
mask_CH = d["State"].eq(U("SWITZERLAND"))
mask_I  = d["State"].isin({U("ITALY"), U("SAN MARINO")})
mask_union_areas = (mask_A | mask_B | mask_C | mask_CH | mask_I)

# ---------- series for CQA ----------
w_non_it = d.loc[filt_common & (mask_A | mask_B | mask_C | mask_CH), "Weight_obj"].dropna().astype(float).values
w_it     = d.loc[filt_common & mask_I,                                 "Weight_obj"].dropna().astype(float).values
w_all = d.loc[filt_common & mask_union_areas, "Weight_obj"].dropna().astype(float).values

print(f"[CHECK] n non-Italy (A+B+C+CH): {w_non_it.size} | n Italy: {w_it.size}")

# quick breakdown for a sanity check
for name, m in [("Area A", mask_A), ("Area B", mask_B), ("Area C", mask_C), ("Switzerland", mask_CH), ("Italy", mask_I)]:
    n_here = int((filt_common & m & d["Weight_obj"].notna()).sum())
    print(f"  - {name}: {n_here}")

# (optional) unique Or_fin values after filtering, to check OR_3/4/5
print("Unique Or_fin values after filtering:", sorted(d.loc[filt_common, "Or_fin"].dropna().unique().tolist()))

# ---------- auto-detect period column ----------
def guess_or_column(df):
    # 1) look for a column with values like OR_3 / OR 4 / OR.5
    patt = re.compile(r"^\s*OR[\s_\.]*([1-5])\s*$", re.IGNORECASE)
    candidates = []
    for c in df.columns:
        ser = df[c].astype(str)

# ---------- BALANCE WEIGHTS (consistent 7-200 g) ----------
BW_FILE  = "Balance_Weights_2024.xlsx"; BW_SHEET = 0
try:
    bw = pd.read_excel(BW_FILE, sheet_name=BW_SHEET)
    def col_like_loose(df, *needles):
        cols = list(df.columns); low = [c.lower() for c in cols]
        for n in needles:
            n_low = str(n).lower()
            for i,c in enumerate(low):
                if n_low in c:
                    return cols[i]
        return None
    c_mass    = col_like_loose(bw, "mass in g (complete/after reconstruction)", "mass in g", "mass", "weight")
    c_country = col_like_loose(bw, "country","nation","state")
    c_typo    = col_like_loose(bw, "typology","type","form","category")

    BW = pd.DataFrame({
        "Mass_g":  pd.to_numeric(bw[c_mass], errors="coerce") if c_mass else np.nan,
        "Country": bw[c_country].astype(str) if c_country else "",
        "Typology":bw[c_typo].astype(str)    if c_typo   else "",
    })
    EXCL_TYPO = {"kannelurenstein","piriform","other hanging"}
    EXCL_COUNTRIES = {"ENGLAND","PORTUGAL","SPAIN"}
    def _ok_typo(s):  return _norm(s) not in EXCL_TYPO
    def _ok_ctry(s):  return str(s).strip().upper() not in EXCL_COUNTRIES
    mask_bw_ok = BW["Mass_g"].notna() & (BW["Mass_g"] > 0) \
                 & BW["Typology"].map(_ok_typo) & BW["Country"].map(_ok_ctry)
    w_bal = BW.loc[mask_bw_ok & BW["Mass_g"].between(7,200, inclusive="left"), "Mass_g"].dropna().values
except Exception:
    w_bal = np.array([])

# ===== tqdm =====
import sys, os
os.environ["TQDM_NOTEBOOK"] = "0"   # non usare widget notebook
try:
    from tqdm import tqdm as _tqdm  # text-based version
    def tqdm_iter(it, **kw):
        kw0 = dict(leave=True, dynamic_ncols=True)
        kw0.update(kw)
        return _tqdm(it, file=sys.stdout, **kw0)
except Exception:
    def tqdm_iter(it, **kw): 
        return it

def run_cqa_mc_from_series(series,
                           min_data_sample=7, max_data_sample=200,
                           min_quantum=2, max_quantum=16, step=0.02,
                           mc_parameter=0.15, mc_iterations=1000):
    s = pd.to_numeric(pd.Series(series), errors="coerce").dropna()
    s = s[(s >= min_data_sample) & (s < max_data_sample)]
    n = int(s.shape[0])
    quanta_arr = np.arange(min_quantum, max_quantum + step, step)

    if n == 0:
        phi_q_df = pd.DataFrame({"Phi_q_values": np.nan, "quanta": quanta_arr})
        return {"phi_q_df": phi_q_df, "quanta_arr": quanta_arr,
                "alpha_1": np.nan, "alpha_5": np.nan,
                "quantum_max": np.nan, "phi_max": np.nan, "n": 0}

    # ===== CQA =====
    df = pd.DataFrame({'Filtered Values': s})
    sample_size = df.count()
    coeff = 2 / sample_size
    sample_coefficient = np.sqrt(coeff).iloc[0]

    cosine_results = pd.concat([(2 * math.pi * df.iloc[:, 0]) / q for q in quanta_arr], axis=1)
    cosine_results.columns = [f'Results {round(q, 2)}' for q in quanta_arr]
    cosine_results = np.cos(cosine_results)

    cosine_sum_series = cosine_results.sum(axis=0)
    cosine_sum = cosine_sum_series.to_frame()

    phi_q_df = cosine_sum.multiply(sample_coefficient)
    phi_q_df.columns = ['Phi_q_values']
    phi_q_df['quanta'] = quanta_arr

    if phi_q_df['Phi_q_values'].dropna().empty:
        quantum_max = np.nan
        phi_max = np.nan
    else:
        idx_best = phi_q_df['Phi_q_values'].idxmax()
        quantum_max = float(phi_q_df.loc[idx_best, 'quanta'])
        phi_max = float(phi_q_df.loc[idx_best, 'Phi_q_values'])

    # ===== Monte Carlo =====
    def apply_variation(df_copy, mc_parameter):
        percent_diff = df_copy * mc_parameter
        random_matrix = np.random.uniform(-1, 1, size=df_copy.shape)
        variations = random_matrix * percent_diff
        return df_copy + variations

    # generate the perturbed datasets
    intermediate_dfs = []
    for _ in tqdm_iter(range(mc_iterations), desc="Monte Carlo", unit="iter"):
        df_copy = df.copy()
        mc_df = apply_variation(df_copy, mc_parameter)
        intermediate_dfs.append(mc_df[df.columns[0]])

    final_mc_df = pd.concat(intermediate_dfs, axis=1)
    final_mc_df.columns = [f'MC_{i+1}' for i in range(mc_iterations)]

    def calculate_Phi_q_values_mc(df_mc, quanta_arr):
        Phi_q_values_mc_list = []
        for q in quanta_arr:
            cosine_result_mc = np.cos((2 * math.pi * df_mc.iloc[:, 0]) / q)
            Phi_q_values_mc = np.sum(cosine_result_mc) * sample_coefficient
            Phi_q_values_mc_list.append(Phi_q_values_mc)
        return Phi_q_values_mc_list

    Montecarlo_phi_q_values = []
    for col in final_mc_df.columns:
        df_mc = final_mc_df[[col]]
        Montecarlo_phi_q_values.append(calculate_Phi_q_values_mc(df_mc, quanta_arr))
    Montecarlo_phi_q = pd.DataFrame(Montecarlo_phi_q_values, columns=quanta_arr).T

    alpha_1_share = round(mc_iterations * 0.01)
    alpha_5_share = round(mc_iterations * 0.05)
    mc_max_column = Montecarlo_phi_q.max()

    if mc_max_column.dropna().empty:
        alpha_1 = np.nan
        alpha_5 = np.nan
    else:
        alpha_1 = float(mc_max_column.nlargest(alpha_1_share).iloc[-1])
        alpha_5 = float(mc_max_column.nlargest(alpha_5_share).iloc[-1])

    return {"phi_q_df": phi_q_df, "quanta_arr": quanta_arr,
            "alpha_1": alpha_1, "alpha_5": alpha_5,
            "quantum_max": quantum_max, "phi_max": phi_max, "n": n}

# ---- EXECUTION (reuses w_non_it, w_it, w_all, w_bal already prepared) ----
cqa_non_it = run_cqa_mc_from_series(w_non_it)
cqa_it     = run_cqa_mc_from_series(w_it)
cqa_all    = run_cqa_mc_from_series(w_all)
cqa_bal    = run_cqa_mc_from_series(w_bal)

def _fmt(x):
    try:
        if x is None or (isinstance(x, float) and (np.isnan(x) or np.isinf(x))):
            return "–"
        return f"{float(x):.2f}"
    except Exception:
        return str(x)

print(
    "n (non-Italy / Italy / ALL / BW):",
    cqa_non_it["n"], cqa_it["n"], cqa_all["n"], cqa_bal["n"]
)
print(
    "Best quantum (non-IT / IT / ALL / BW):",
    _fmt(cqa_non_it["quantum_max"]),
    _fmt(cqa_it["quantum_max"]),
    _fmt(cqa_all["quantum_max"]),
    _fmt(cqa_bal["quantum_max"])
)
print(
    "Max Φ(q) (non-IT / IT / ALL / BW):",
    _fmt(cqa_non_it["phi_max"]),
    _fmt(cqa_it["phi_max"]),
    _fmt(cqa_all["phi_max"]),
    _fmt(cqa_bal["phi_max"])
)

print(
    "alpha 5% (non-IT / IT / ALL / BW):",
    _fmt(cqa_non_it["alpha_5"]),
    _fmt(cqa_it["alpha_5"]),
    _fmt(cqa_all["alpha_5"]),
    _fmt(cqa_bal["alpha_5"])
)
print(
    "alpha 1% (non-IT / IT / ALL / BW):",
    _fmt(cqa_non_it["alpha_1"]),
    _fmt(cqa_it["alpha_1"]),
    _fmt(cqa_all["alpha_1"]),
    _fmt(cqa_bal["alpha_1"])
)

# %%

# ============================================================
# CQA DATA PREPARATION (no direct figure)
# Series preparation for the main DB and the PNAS balance-weight
# dataset (column W), feeding Figure 71.
# ============================================================

# ============== PREPARE SERIES for CQA (main DB) + PNAS (column W) ==============
import pandas as pd, numpy as np, math, re, unicodedata, sys, os

# ---------- tqdm ----------
os.environ["TQDM_NOTEBOOK"] = "0"
try:
    from tqdm import tqdm as _tqdm
    def tqdm_iter(it, **kw):
        kw0 = dict(leave=True, dynamic_ncols=True)
        kw0.update(kw)
        return _tqdm(it, file=sys.stdout, **kw0)
except Exception:
    def tqdm_iter(it, **kw): 
        return it

# ---------- helper ----------
def _strip_accents(s):
    s = str(s).replace("ß","ss")
    return "".join(ch for ch in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(ch))
def _norm(s):
    s = _strip_accents(str(s)).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def to_str_u(s):  # upper + strip
    return s.astype(str).str.upper().str.strip()

def col_like(df, *needles):
    """Find a column by name/partial match (case-insensitive)."""
    cols = list(df.columns)
    lows = [c.lower().strip() for c in cols]
    for n in needles:
        n_low = str(n).lower().strip()
        # match 
        for i, c in enumerate(lows):
            if n_low == c:
                return cols[i]
        # match 
        for i, c in enumerate(lows):
            if n_low in c:
                return cols[i]
    raise KeyError(f"Column not found for: {needles}")

def starts_with_gold(s_raw: str) -> bool:
    return _norm(s_raw).startswith("gold")

U = lambda s: str(s).upper().strip()

# ---------- main data source ----------
if "raw" in globals():
    df_src = raw.copy()
else:
    for fname in ("DB.xlsx", "DB_SI.xlsx"):
        try:
            df_src = pd.read_excel(fname, sheet_name="DB")
            print(f"[INFO] Caricato DB principale: {fname}")
            break
        except Exception:
            df_src = None
    if df_src is None:
        raise FileNotFoundError("Could not find 'raw' in memory or the DB.xlsx / DB_SI.xlsx files on disk.")

# ---------- column mapping, main DB ----------
c_or     = col_like(df_src, "Or_fin", "or fin", "period", "or_fin")
c_state  = col_like(df_src, "State")
c_reg    = col_like(df_src, "Region")
c_type   = col_like(df_src, "Artefact", "artifact", "type")
c_weight = col_like(df_src, "Weight_obj", "weight obj", "weight_obj", "weight", "mass", "weight in g", "mass in g")
c_frag   = col_like(df_src, "Fragmented", "fragmented")
c_comp   = col_like(df_src, "Complete", "complete")
c_match  = col_like(df_src, "Matching fr", "matching_fr", "matching fr.", "matching")

# ---------- dataframe normalized ----------
d = pd.DataFrame({
    "Or_fin":      to_str_u(df_src[c_or]).str.replace(r"\s+", "_", regex=True),
    "State":       to_str_u(df_src[c_state]),
    "Region":      to_str_u(df_src[c_reg]),
    "Artefact":    df_src[c_type].astype(str).str.strip(),
    "Weight_obj":  pd.to_numeric(df_src[c_weight], errors="coerce"),
    "Fragmented":  pd.to_numeric(df_src[c_frag],  errors="coerce").fillna(0),
    "Complete":    pd.to_numeric(df_src[c_comp],  errors="coerce").fillna(0),
    "Matching_fr": pd.to_numeric(df_src[c_match], errors="coerce").fillna(0),
})

# ---------- FILTERS (ingots INCLUDED, only Gold* excluded) ----------
mask_period = d["Or_fin"].isin({"OR_3","OR_4","OR_5"})
mask_range  = d["Weight_obj"].between(7, 200, inclusive="left")
mask_keep   = ~d["Artefact"].apply(starts_with_gold)
mask_frag   = (d["Fragmented"] > 0) | ((d["Matching_fr"] > 0) & (d["Complete"] < 1))
filt_common = mask_period & mask_range & mask_keep & mask_frag

# ---------- areas (as defined) ----------
GER = U("GERMANY")
mask_A  = (d["Region"].isin({U("BADEN-WÜRTTEMBERG"), U("BAYERN")}) |
           d["State"].isin({U("AUSTRIA"), U("SLOVENIA")}))
mask_B  = (d["Region"].eq(U("BRANDENBURG")) |
           d["State"].eq(U("POLAND")) |
           d["Region"].eq(U("SACHSEN")))
mask_C  = ((d["State"].eq(GER) & d["Region"].isin({U("MECKLENBURG-VORPOMMERN"),
                                                   U("SCHLESWIG-HOLSTEIN"),
                                                   U("NIEDERSACHSEN")})) |
           (d["State"].eq(GER) & d["Region"].isin({U("THÜRINGEN"),
                                                   U("HESSEN"),
                                                   U("RHEINLAND-PFALZ")})) |
           (d["Region"].eq(U("SACHSEN-ANHALT"))))
mask_CH = d["State"].eq(U("SWITZERLAND"))
mask_I  = d["State"].isin({U("ITALY"), U("SAN MARINO")})
mask_union_areas = (mask_A | mask_B | mask_C | mask_CH | mask_I)

# ---------- main series ----------
w_non_it = d.loc[filt_common & (mask_A | mask_B | mask_C | mask_CH), "Weight_obj"].dropna().astype(float).values
w_it     = d.loc[filt_common & mask_I,                                 "Weight_obj"].dropna().astype(float).values
w_all    = d.loc[filt_common & mask_union_areas,                        "Weight_obj"].dropna().astype(float).values

print(f"[CHECK] n Central Europe (A+B+C+CH): {w_non_it.size} | n Italy: {w_it.size}")
for name, m in [("Area A", mask_A), ("Area B", mask_B), ("Area C", mask_C),
                ("Switzerland", mask_CH), ("Italy", mask_I)]:
    n_here = int((filt_common & m & d["Weight_obj"].notna()).sum())
    print(f"  - {name}: {n_here}")
print("Unique Or_fin values after filtering:", sorted(d.loc[filt_common, "Or_fin"].dropna().unique().tolist()))

# ---------- PNAS: read & filter ----------
PNAS_FILE = "pnas.xlsx"; PNAS_SHEET = 0
EXCL_SYSTEM = {"indus valley", "mesopotamia", "aegean-anatolia"}  # case-insensitive

pnas_raw = pd.read_excel(PNAS_FILE, sheet_name=PNAS_SHEET)
c_mass   = col_like(pnas_raw, "W", "mass in g (complete/after reconstruction)", "mass in g", "mass", "weight")
try:
    c_system = col_like(pnas_raw, "system")
except KeyError:
    c_system = None

dp = pd.DataFrame({"Mass_g": pd.to_numeric(pnas_raw[c_mass], errors="coerce")})
if c_system:
    dp["System"] = pnas_raw[c_system].astype(str)

mask_mass_p = dp["Mass_g"].between(7, 200, inclusive="left")
mask_sys_p  = True if not c_system else ~dp["System"].str.strip().str.casefold().isin(EXCL_SYSTEM)
w_pnas      = dp.loc[mask_mass_p & mask_sys_p, "Mass_g"].dropna().astype(float).values
print(f"[PNAS] n after filters (7-200 g; systems excl.): {w_pnas.size}")

# ---------- CQA + Monte Carlo (original logic, with a text progress bar) ----------
def run_cqa_mc_from_series(series,
                           min_data_sample=7, max_data_sample=200,
                           min_quantum=2, max_quantum=16, step=0.02,
                           mc_parameter=0.15, mc_iterations=1000):
    s = pd.to_numeric(pd.Series(series), errors="coerce").dropna()
    s = s[(s >= min_data_sample) & (s < max_data_sample)]
    n = int(s.shape[0])
    quanta_arr = np.arange(min_quantum, max_quantum + step, step)

    if n == 0:
        phi_q_df = pd.DataFrame({"Phi_q_values": np.nan, "quanta": quanta_arr})
        return {"phi_q_df": phi_q_df, "quanta_arr": quanta_arr,
                "alpha_1": np.nan, "alpha_5": np.nan,
                "quantum_max": np.nan, "phi_max": np.nan, "n": 0}

    df = pd.DataFrame({'Filtered Values': s})
    coeff = 2 / df.count()
    sample_coefficient = float(np.sqrt(coeff).iloc[0])

    # quantogram
    cos = pd.concat([(2 * math.pi * df.iloc[:, 0]) / q for q in quanta_arr], axis=1)
    cos = np.cos(cos)
    phi_q_df = pd.DataFrame(cos.sum(axis=0), columns=['Phi_q_values']) * sample_coefficient
    phi_q_df['quanta'] = quanta_arr

    # best
    idx_best    = phi_q_df['Phi_q_values'].to_numpy().argmax()
    quantum_max = float(phi_q_df['quanta'].iloc[idx_best])
    phi_max     = float(phi_q_df['Phi_q_values'].iloc[idx_best])

    # Monte Carlo
    def apply_variation(df_copy, mc_parameter):
        percent = df_copy * mc_parameter
        return df_copy + np.random.uniform(-1, 1, size=df_copy.shape) * percent

    inter = []
    for _ in tqdm_iter(range(mc_iterations), desc="Monte Carlo", unit="iter"):
        inter.append(apply_variation(df, mc_parameter)[df.columns[0]])
    final_mc_df = pd.concat(inter, axis=1)
    final_mc_df.columns = [f"MC_{i+1}" for i in range(mc_iterations)]

    # Phi(q) for each Monte Carlo column
    def phi_for_df(df_mc, quanta_arr):
        vals = []
        for q in quanta_arr:
            vals.append(float(np.sum(np.cos((2 * math.pi * df_mc.iloc[:, 0]) / q)) * sample_coefficient))
        return vals

    mc_vals = []
    for col in final_mc_df.columns:
        mc_vals.append(phi_for_df(final_mc_df[[col]], quanta_arr))
    MC = pd.DataFrame(mc_vals, columns=quanta_arr).T

    # thresholds (1% e 5%)
    share1 = max(1, round(mc_iterations * 0.01))
    share5 = max(1, round(mc_iterations * 0.05))
    mc_max = MC.max()
    alpha_1 = float(mc_max.nlargest(share1).iloc[-1])
    alpha_5 = float(mc_max.nlargest(share5).iloc[-1])

    return {"phi_q_df": phi_q_df, "quanta_arr": quanta_arr,
            "alpha_1": alpha_1, "alpha_5": alpha_5,
            "quantum_max": quantum_max, "phi_max": phi_max, "n": n}

# ---------- MONTE CARLO PARAMETERS ----------
MC_ITERATIONS = 1000   # increase (e.g. 500-1000) for more stable estimates

# ---------- CQA EXECUTION ----------
cqa_non_it = run_cqa_mc_from_series(w_non_it, mc_iterations=MC_ITERATIONS)
cqa_it     = run_cqa_mc_from_series(w_it,     mc_iterations=MC_ITERATIONS)
cqa_all    = run_cqa_mc_from_series(w_all,    mc_iterations=MC_ITERATIONS)
cqa_pnas   = run_cqa_mc_from_series(w_pnas,   mc_iterations=MC_ITERATIONS)

# ---------- REPORT ----------
def _fmt(x):
    try:
        if x is None or (isinstance(x, float) and (np.isnan(x) or np.isinf(x))):
            return "–"
        return f"{float(x):.2f}"
    except Exception:
        return str(x)

print(
    "n (Central Europe / Italy / Overall / PNAS):",
    cqa_non_it["n"], cqa_it["n"], cqa_all["n"], cqa_pnas["n"]
)
print(
    "Best quantum (CE / IT / ALL / PNAS):",
    _fmt(cqa_non_it["quantum_max"]),
    _fmt(cqa_it["quantum_max"]),
    _fmt(cqa_all["quantum_max"]),
    _fmt(cqa_pnas["quantum_max"])
)
print(
    "Max Φ(q) (CE / IT / ALL / PNAS):",
    _fmt(cqa_non_it["phi_max"]),
    _fmt(cqa_it["phi_max"]),
    _fmt(cqa_all["phi_max"]),
    _fmt(cqa_pnas["phi_max"])
)
print(
    "alpha 5% (CE / IT / ALL / PNAS):",
    _fmt(cqa_non_it["alpha_5"]),
    _fmt(cqa_it["alpha_5"]),
    _fmt(cqa_all["alpha_5"]),
    _fmt(cqa_pnas["alpha_5"])
)
print(
    "alpha 1% (CE / IT / ALL / PNAS):",
    _fmt(cqa_non_it["alpha_1"]),
    _fmt(cqa_it["alpha_1"]),
    _fmt(cqa_all["alpha_1"]),
    _fmt(cqa_pnas["alpha_1"])
)


# %%

# ============================================================
# FIGURE 71 — CQA of fragments of objects from Italian and Central
# European hoards, and of balance weights, in the Late Bronze Age.
# 2x2 grid: Central European fragm. bronze, whole sample, Italian
# fragm. bronze, European balance weights.
# ============================================================

# ===== CQA - 2x2 half-page, unified style (black line, gray fill) =====
import numpy as np
import matplotlib.pyplot as plt

FIGSIZE_HALF = (6.30, 4.33)  # ~160 x 110 mm

COL_LINE  = "black"
COL_FILL  = "#D9D9D9"
COL_ALPHA = "#FF7F0E"  # orange for alpha_5
LS_DASH   = (0, (4, 3))

def _fill_between_zero(ax, x, y, face, alpha=0.35, z=1):
    ax.fill_between(x, 0, y, facecolor=face, alpha=alpha, zorder=z, linewidth=0)

def _style_axes(ax, title, y_lim=None):
    ax.set_xlim(4, 16)
    ax.set_xticks(np.arange(4, 17, 1))
    if y_lim is None:
        y_lim = (-6, 9)
    ax.set_ylim(*y_lim)
    ax.grid(axis="y", color="0.90")
    ax.set_ylabel("φ(q)")
    ax.set_title(title, loc="left", fontsize=9)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.tick_params(labelsize=7)

def plot_panel(ax, result, title, show_alpha=False, show_band_lines=False, y_lim=None):
    phi = result["phi_q_df"]
    x = phi["quanta"].to_numpy()
    y = phi["Phi_q_values"].to_numpy()

    # area under the curve
    _fill_between_zero(ax, x, y, face=COL_FILL, alpha=0.35, z=1)
    # line
    ax.plot(x, y, color=COL_LINE, lw=1.1, zorder=2)

    # ONLY dashed lines at 9 and 11 (no fill)
    if show_band_lines:
        ax.axvline(9,  color=COL_LINE, ls=LS_DASH, lw=0.8)
        ax.axvline(11, color=COL_LINE, ls=LS_DASH, lw=0.8)

    # alpha 5% (only where required)
    if show_alpha:
        a5 = result.get("alpha_1", np.nan)
        if np.isfinite(a5):
            ax.axhline(a5, color=COL_ALPHA, ls=LS_DASH, lw=0.7)

    _style_axes(ax, title, y_lim=y_lim)

# figure 2x2
fig, axes = plt.subplots(2, 2, figsize=FIGSIZE_HALF, dpi=300, sharex=True)
ax_tl, ax_tr, ax_bl, ax_br = axes[0,0], axes[0,1], axes[1,0], axes[1,1]

# draw
plot_panel(ax_tl, cqa_non_it, "Central European fragm. bronze",
           show_alpha=False, show_band_lines=True, y_lim=None)         # default [-6, 9]
plot_panel(ax_tr, cqa_all,    "Fragm. bronze (Tot sample)",
           show_alpha=True,  show_band_lines=True, y_lim=None)         # default [-6, 9]
plot_panel(ax_bl, cqa_it,     "Italian fragm. bronze",
           show_alpha=False, show_band_lines=True, y_lim=None)         # default [-6, 9]
plot_panel(ax_br, cqa_pnas,    "European Balance weights",
           show_alpha=True,  show_band_lines=True, y_lim=(-5, 6))      # SOLO qui [-5, 6]

# x labels only at the bottom
ax_bl.set_xlabel("quanta", fontsize=6)
ax_br.set_xlabel("quanta", fontsize=6)

plt.tight_layout()
plt.subplots_adjust(bottom=0.16, hspace=0.36, wspace=0.28)

#plt.savefig(
#    "Fig71_CQA_halfpage_2x2_unified.jpg",
#    dpi=300, bbox_inches="tight",
#    format="jpeg",
#    pil_kwargs={"quality":95, "subsampling":0}
#)
plt.show()


# %%

# ============================================================
# CQA DATA PREPARATION for Figure 55 (no direct figure)
# Series preparation: OR_3-5, Complete>0, Weight_obj, Areas
# A/B/C/Switzerland (ingots included, Gold* excluded).
# ============================================================

# ============== CQA - OR_3-5, Complete>0, Weight_obj - Areas A/B/C/Switzerland (ingots included, Gold* excluded) ==============
import pandas as pd, numpy as np, math, re, unicodedata, sys, os

# ---------- tqdm ----------
os.environ["TQDM_NOTEBOOK"] = "0"
try:
    from tqdm import tqdm as _tqdm
    def tqdm_iter(it, **kw):
        kw0 = dict(leave=True, dynamic_ncols=True)
        kw0.update(kw)
        return _tqdm(it, file=sys.stdout, **kw0)
except Exception:
    def tqdm_iter(it, **kw): 
        return it

# ---------- helper ----------
def _strip_accents(s):
    s = str(s).replace("ß","ss")
    return "".join(ch for ch in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(ch))
def _norm(s):
    s = _strip_accents(str(s)).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()
def to_str_u(s):  # upper + strip
    return s.astype(str).str.upper().str.strip()
def col_like(df, *needles):
    cols = list(df.columns)
    lows = [c.lower().strip() for c in cols]
    for n in needles:
        n_low = str(n).lower().strip()
        # match 
        for i, c in enumerate(lows):
            if n_low == c:
                return cols[i]
        # match 
        for i, c in enumerate(lows):
            if n_low in c:
                return cols[i]
    raise KeyError(f"Column not found for: {needles}")
def starts_with_gold(s_raw: str) -> bool:
    return _norm(s_raw).startswith("gold")
U = lambda s: str(s).upper().strip()

# ---------- Reading ----------
if "raw" in globals():
    df_src = raw.copy()
else:
    for fname in ("DB.xlsx", "DB_SI.xlsx"):
        try:
            df_src = pd.read_excel(fname, sheet_name="DB")
            print(f"[INFO] Caricato DB principale: {fname}")
            break
        except Exception:
            df_src = None
    if df_src is None:
        raise FileNotFoundError("Could not find 'raw' in memory or the DB.xlsx / DB_SI.xlsx files on disk.")

# ---------- column mapping ----------
c_or     = col_like(df_src, "Or_fin", "or fin", "period", "or_fin")
c_state  = col_like(df_src, "State")
c_reg    = col_like(df_src, "Region")
c_type   = col_like(df_src, "Artefact", "artifact", "type")
c_weight = col_like(df_src, "Weight_obj", "weight obj", "weight_obj", "weight", "mass", "weight in g", "mass in g")
c_comp   = col_like(df_src, "Complete", "complete")

# ---------- dataframe normalized ----------
d = pd.DataFrame({
    "Or_fin":      to_str_u(df_src[c_or]).str.replace(r"\s+", "_", regex=True),
    "State":       to_str_u(df_src[c_state]),
    "Region":      to_str_u(df_src[c_reg]),
    "Artefact":    df_src[c_type].astype(str).str.strip(),
    "Weight_obj":  pd.to_numeric(df_src[c_weight], errors="coerce"),
    "Complete":    pd.to_numeric(df_src[c_comp],  errors="coerce").fillna(0),
})

# ---------- filters for this analysis ----------
# periods: OR_3, OR_4, OR_5
mask_period = d["Or_fin"].isin({"OR_3","OR_4","OR_5"})
# Weight range: 7-200 g (>=7, <200)
mask_range  = d["Weight_obj"].between(7, 200, inclusive="left")
# Intact objects/ingots: Complete > 0
mask_intact = d["Complete"] > 0
# Ignore Gold*; ingots included
mask_keep   = ~d["Artefact"].apply(starts_with_gold)

filt_common = mask_period & mask_range & mask_intact & mask_keep

# ---------- areas ----------
GER = U("GERMANY")
mask_A  = (d["Region"].isin({U("BADEN-WÜRTTEMBERG"), U("BAYERN")}) |
           d["State"].isin({U("AUSTRIA"), U("SLOVENIA")}))
mask_B  = (d["Region"].eq(U("BRANDENBURG")) |
           d["State"].eq(U("POLAND")) |
           d["Region"].eq(U("SACHSEN")))
mask_C  = ((d["State"].eq(GER) & d["Region"].isin({U("MECKLENBURG-VORPOMMERN"),
                                                   U("SCHLESWIG-HOLSTEIN"),
                                                   U("NIEDERSACHSEN")})) |
           (d["State"].eq(GER) & d["Region"].isin({U("THÜRINGEN"),
                                                   U("HESSEN"),
                                                   U("RHEINLAND-PFALZ")})) |
           (d["Region"].eq(U("SACHSEN-ANHALT"))))
mask_IT = d["State"].eq(U("ITALY"))

AREAS = [
    ("Area A",      mask_A),
    ("Area B",      mask_B),
    ("Area C",      mask_C),
    ("Italy", mask_IT),
]

# ---------- series per area ----------
series_by_area = {}
for label, m in AREAS:
    w = d.loc[filt_common & m, "Weight_obj"].dropna().astype(float).values
    series_by_area[label] = w

print("[CHECK] n per area (OR_3-5, Complete>0, 7-200g, no Gold*)")
for label in series_by_area:
    print(f"  - {label}: {series_by_area[label].size}")

# ---------- CQA + Monte Carlo function ----------
def run_cqa_mc_from_series(series,
                           min_data_sample=7, max_data_sample=201,
                           min_quantum=2, max_quantum=16, step=0.02,
                           mc_parameter=0.15, mc_iterations=100):
    import numpy as np, pandas as pd, math

    # --- prepare filtered series ---
    s = pd.to_numeric(pd.Series(series), errors="coerce").dropna()
    s = s[(s >= min_data_sample) & (s < max_data_sample)]
    n = int(s.shape[0])
    quanta_arr = np.arange(min_quantum, max_quantum + step, step)

    if n == 0:
        phi_q_df = pd.DataFrame({"Phi_q_values": np.nan, "quanta": quanta_arr})
        return {"phi_q_df": phi_q_df, "quanta_arr": quanta_arr,
                "alpha_1": np.nan, "alpha_5": np.nan,
                "quantum_max": np.nan, "phi_max": np.nan, "n": 0}

    df = pd.DataFrame({'Filtered Values': s})

    coeff = 2 / df.count()
    sample_coefficient = float(np.sqrt(coeff).iloc[0])

    # --- robust QUANTOGRAM (keeps alignment with quanta_arr) ---
    cos = pd.concat([(2 * math.pi * df.iloc[:, 0]) / q for q in quanta_arr], axis=1)
    cos.columns = [round(q, 2) for q in quanta_arr]          # for readability only
    cos = np.cos(cos)                                        # resta un DataFrame
    phi_vals = cos.sum(axis=0).to_numpy() * sample_coefficient

    phi_q_df = pd.DataFrame({
        "Phi_q_values": phi_vals,
        "quanta": quanta_arr
    })

    # best (positional, avoids index ambiguity)
    i_best     = int(np.nanargmax(phi_q_df["Phi_q_values"].to_numpy()))
    quantum_max= float(phi_q_df["quanta"].to_numpy()[i_best])
    phi_max    = float(phi_q_df["Phi_q_values"].to_numpy()[i_best])

    # --- Monte Carlo ---
    def apply_variation(df_copy, mc_parameter):
        percent = df_copy * mc_parameter
        noise   = np.random.uniform(-1, 1, size=df_copy.shape)
        return df_copy + noise * percent

    inter = []
    for _ in tqdm_iter(range(mc_iterations), desc="Monte Carlo", unit="iter"):
        inter.append(apply_variation(df, mc_parameter)[df.columns[0]])
    final_mc_df = pd.concat(inter, axis=1)
    final_mc_df.columns = [f"MC_{i+1}" for i in range(mc_iterations)]

    # Phi(q) for each Monte Carlo column
    def phi_for_df(df_mc, quanta_arr):
        out = []
        col = df_mc.columns[0]
        for q in quanta_arr:
            val = float(np.sum(np.cos((2 * math.pi * df_mc[col]) / q)) * sample_coefficient)
            out.append(val)
        return out

    MC_vals = []
    for col in final_mc_df.columns:
        MC_vals.append(phi_for_df(final_mc_df[[col]], quanta_arr))
    MC = pd.DataFrame(MC_vals, columns=quanta_arr).T

    # threshold (1% e 5%)
    share1 = max(1, round(mc_iterations * 0.01))
    share5 = max(1, round(mc_iterations * 0.05))
    mc_max = MC.max()
    alpha_1 = float(mc_max.nlargest(share1).iloc[-1])
    alpha_5 = float(mc_max.nlargest(share5).iloc[-1])

    return {
        "phi_q_df": phi_q_df,
        "quanta_arr": quanta_arr,
        "alpha_1": alpha_1,
        "alpha_5": alpha_5,
        "quantum_max": quantum_max,
        "phi_max": phi_max,
        "n": n
    }


# ---------- Monte Carlo parameters (increase to 500-1000 if desired) ----------
MC_ITER = 10

# ---------- run for each area ----------
cqa_results = {}
for label, w in series_by_area.items():
    cqa_results[label] = run_cqa_mc_from_series(w, mc_iterations=MC_ITER)

# export as variables used by the plot
cqa_A  = cqa_results["Area A"]
cqa_B  = cqa_results["Area B"]
cqa_C  = cqa_results["Area C"]
cqa_IT_comp = cqa_results["Italy"]

# --------------------
def _fmt(x):
    try:
        if x is None or (isinstance(x, float) and (np.isnan(x) or np.isinf(x))):
            return "–"
        return f"{float(x):.2f}"
    except Exception:
        return str(x)

print("\n=== CQA SUMMARY (OR_3-5, Complete>0, 7-200g, no Gold*) ===")
for lbl in ["Area A","Area B","Area C","Italy"]:
    r = cqa_results[lbl]
    print(f"{lbl:12s}  n={r['n']:5d}  q*={_fmt(r['quantum_max'])}  Φmax={_fmt(r['phi_max'])}  α5%={_fmt(r['alpha_5'])}  α1%={_fmt(r['alpha_1'])}")


# %%

# ============================================================
# FIGURE 55 — CQA of the weight of complete objects from the Late
# Bronze Age (Horizons 3-5) in areas A, B, C and Italy.
# 2x2 grid, unified style, Y = [-6, 10].
# ============================================================

# ===== CQA - 2x2 half-page, unified style (black line, gray fill), Y = [-4, 6], no alpha/band =====
import numpy as np
import matplotlib.pyplot as plt

FIGSIZE_HALF = (6.30, 4.33)  # ~160 x 110 mm

COL_LINE  = "black"
COL_FILL  = "#D9D9D9"

def _fill_between_zero(ax, x, y, face, alpha=0.35, z=1):
    ax.fill_between(x, 0, y, facecolor=face, alpha=alpha, zorder=z, linewidth=0)

def _style_axes(ax, title, y_lim=(-4, 6)):
    ax.set_xlim(4, 16)
    ax.set_xticks(np.arange(4, 17, 1))
    ax.set_ylim(*y_lim)
    ax.grid(axis="y", color="0.90")
    ax.set_ylabel("φ(q)")
    ax.set_title(title, loc="left", fontsize=9)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.tick_params(labelsize=7)

def plot_panel(ax, result, title, y_lim=(-6, 10)):
    phi = result["phi_q_df"]
    x = phi["quanta"].to_numpy()
    y = phi["Phi_q_values"].to_numpy()

    # area under the curve + line
    _fill_between_zero(ax, x, y, face=COL_FILL, alpha=0.35, z=1)
    ax.plot(x, y, color=COL_LINE, lw=1.1, zorder=2)

    _style_axes(ax, title, y_lim=y_lim)

# figure 2x2
fig, axes = plt.subplots(2, 2, figsize=FIGSIZE_HALF, dpi=300, sharex=True, sharey=True)
ax_tl, ax_tr, ax_bl, ax_br = axes[0,0], axes[0,1], axes[1,0], axes[1,1]

# draw (no band/alpha; same Y scale for all)
plot_panel(ax_tl, cqa_A,        "Area A", y_lim=(-6, 10))
plot_panel(ax_tr, cqa_B,        "Area B", y_lim=(-6, 10))
plot_panel(ax_bl, cqa_C,        "Area C", y_lim=(-6, 10))
plot_panel(ax_br, cqa_IT_comp,  "Italy",  y_lim=(-6, 10))

# x labels only at the bottom
ax_bl.set_xlabel("quanta", fontsize=6)
ax_br.set_xlabel("quanta", fontsize=6)

plt.tight_layout()
plt.subplots_adjust(bottom=0.16, hspace=0.36, wspace=0.28)

#plt.savefig(
#    "Fig55_CQA_2x2_areas_complete.jpg",
#    dpi=300, bbox_inches="tight",
#    format="jpeg",
#    pil_kwargs={"quality":95, "subsampling":0}
#)
plt.show()


# %%

# ============================================================
# CQA DATA PREPARATION for Figure 49 (no direct figure)
# Series preparation: OR_2 fragmented objects, Area A vs whole
# sample (Weight_obj, Gold* excluded).
# ============================================================

# ============== CQA - OR_2, FRAGMENTED - area A vs Whole sample (Weight_obj, excludes Gold*) ==============
import pandas as pd, numpy as np, math, re, unicodedata, sys
# --- text-based tqdm (no widget) ---
try:
    from tqdm import tqdm as _tqdm
    def tqdm_iter(it, **kw):
        kw0 = dict(leave=True, dynamic_ncols=True)
        kw0.update(kw)
        return _tqdm(it, file=sys.stdout, **kw0)
except Exception:
    def tqdm_iter(it, **kw): return it

# ---------- helper ----------
def _strip_accents(s):
    s = str(s).replace("ß","ss")
    return "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
def _norm(s):
    s = _strip_accents(str(s)).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()
def to_str_u(s):  # upper + strip
    return s.astype(str).str.upper().str.strip()
def col_like(df, *needles):
    cols = list(df.columns); lows = [c.lower().strip() for c in cols]
    for n in needles:
        n_low = str(n).lower().strip()
        for i, c in enumerate(lows):
            if (n_low == c) or (n_low in c): return cols[i]
    raise KeyError(f"Column not found for: {needles}")
def starts_with_gold(s_raw: str) -> bool:
    return _norm(s_raw).startswith("gold")
U = lambda s: str(s).upper().strip()

# ---------- data source ----------
if "raw" in globals():
    df_src = raw.copy()
else:
    for fname in ("DB.xlsx", "DB_SI.xlsx"):
        try:
            df_src = pd.read_excel(fname, sheet_name="DB"); print(f"[INFO] Caricato: {fname}"); break
        except Exception: df_src = None
    if df_src is None:
        raise FileNotFoundError("Could not find 'raw' in memory or the DB.xlsx / DB_SI.xlsx files on disk.")

# ---------- mapping columns ----------
c_or     = col_like(df_src, "Or_fin", "or fin", "period", "or_fin")
c_state  = col_like(df_src, "State")
c_reg    = col_like(df_src, "Region")
c_type   = col_like(df_src, "Artefact", "artifact", "type")
c_weight = col_like(df_src, "Weight_obj", "weight obj", "weight_obj", "weight", "mass", "weight in g", "mass in g")
c_frag   = col_like(df_src, "Fragmented", "fragmented")
c_comp   = col_like(df_src, "Complete", "complete")
c_match  = col_like(df_src, "Matching fr", "matching_fr", "matching fr.", "matching")

# ---------- dataframe normalized ----------
d = pd.DataFrame({
    "Or_fin":      to_str_u(df_src[c_or]).str.replace(r"\s+", "_", regex=True),
    "State":       to_str_u(df_src[c_state]),
    "Region":      to_str_u(df_src[c_reg]),
    "Artefact":    df_src[c_type].astype(str).str.strip(),
    "Weight_obj":  pd.to_numeric(df_src[c_weight], errors="coerce"),
    "Fragmented":  pd.to_numeric(df_src[c_frag],  errors="coerce").fillna(0),
    "Complete":    pd.to_numeric(df_src[c_comp],  errors="coerce").fillna(0),
    "Matching_fr": pd.to_numeric(df_src[c_match], errors="coerce").fillna(0),
})

# ---------- required filters ----------
mask_or2   = d["Or_fin"].eq("OR_2")
mask_range = d["Weight_obj"].between(7, 200, inclusive="left")             # >=7 & <200
mask_keep = (
    ~d["Artefact"].apply(starts_with_gold)
) & (
    ~d["Artefact"].astype(str).str.lower()
       .str.replace(r"[^a-z0-9]+", " ", regex=True)   # normalize as in _norm
       .str.contains(r"\bingot(s)?\b", regex=True, na=False)
)                     # ⟵ ESCLUDE Gold*
mask_frag  = (d["Fragmented"] > 0) | ((d["Matching_fr"] > 0) & (d["Complete"] < 1))

# ---------- area A ----------
GER = U("GERMANY")
mask_A  = (d["Region"].isin({U("BADEN-WÜRTTEMBERG"), U("BAYERN")}) |
           d["State"].isin({U("AUSTRIA"), U("SLOVENIA")}))

# ---------- weight series (OR_2 fragmented, Gold* excluded) ----------
filt_or2_frag = mask_or2 & mask_range & mask_keep & mask_frag
w_or2_A_frag     = d.loc[filt_or2_frag & mask_A, "Weight_obj"].dropna().astype(float).values
w_or2_ALL_frag   = d.loc[filt_or2_frag,          "Weight_obj"].dropna().astype(float).values
print(f"[CHECK OR_2] n Area A (frag, no Gold*): {w_or2_A_frag.size}  |  n All (frag, no Gold*): {w_or2_ALL_frag.size}")

# ---------- CQA + Monte Carlo ----------
if "run_cqa_mc_from_series" not in globals():
    def run_cqa_mc_from_series(series,
                               min_data_sample=7, max_data_sample=200,
                               min_quantum=2, max_quantum=16, step=0.02,
                               mc_parameter=0.15, mc_iterations=1000):
        s = pd.to_numeric(pd.Series(series), errors="coerce").dropna()
        s = s[(s >= min_data_sample) & (s < max_data_sample)]
        n = int(s.shape[0]); quanta_arr = np.arange(min_quantum, max_quantum + step, step)
        if n == 0:
            phi_q_df = pd.DataFrame({"Phi_q_values": np.nan, "quanta": quanta_arr})
            return {"phi_q_df": phi_q_df, "quanta_arr": quanta_arr,
                    "alpha_1": np.nan, "alpha_5": np.nan,
                    "quantum_max": np.nan, "phi_max": np.nan, "n": 0}
        df = pd.DataFrame({'Filtered Values': s})
        sample_coefficient = float(np.sqrt(2 / df.count()).iloc[0])
        cos = pd.concat([(2 * math.pi * df.iloc[:,0]) / q for q in quanta_arr], axis=1)
        cos.columns = [f"r{round(q,2)}" for q in quanta_arr]
        phi_q_df = np.cos(cos).sum(axis=0).to_frame("Phi_q_values") * sample_coefficient
        phi_q_df["quanta"] = quanta_arr
        idx_best    = phi_q_df['Phi_q_values'].idxmax()
        quantum_max = float(phi_q_df.loc[idx_best, 'quanta'])
        phi_max     = float(phi_q_df.loc[idx_best, 'Phi_q_values'])

        def apply_variation(df_copy, mc_parameter):
            percent_diff = df_copy * mc_parameter
            return df_copy + (np.random.uniform(-1, 1, size=df_copy.shape) * percent_diff)

        cols_mc = []
        for _ in tqdm_iter(range(mc_iterations), desc="Monte Carlo", unit="iter"):
            cols_mc.append(apply_variation(df.copy(), mc_parameter)[df.columns[0]])
        final_mc_df = pd.concat(cols_mc, axis=1)

        def phi_for_df(df_mc, quanta_arr):
            vals = []
            for q in quanta_arr:
                vals.append(float(np.sum(np.cos((2 * math.pi * df_mc.iloc[:,0]) / q)) * sample_coefficient))
            return vals

        mc_vals = [phi_for_df(final_mc_df[[col]], quanta_arr) for col in final_mc_df.columns]
        Montecarlo_phi_q = pd.DataFrame(mc_vals, columns=quanta_arr).T
        a1_share = max(1, round(mc_iterations * 0.01))
        a5_share = max(1, round(mc_iterations * 0.05))
        mc_max = Montecarlo_phi_q.max()
        alpha_1 = float(mc_max.nlargest(a1_share).iloc[-1])
        alpha_5 = float(mc_max.nlargest(a5_share).iloc[-1])
        return {"phi_q_df": phi_q_df, "quanta_arr": quanta_arr,
                "alpha_1": alpha_1, "alpha_5": alpha_5,
                "quantum_max": quantum_max, "phi_max": phi_max, "n": n}

# ---------- execution (new variables) ----------
cqa_or2_A_frag     = run_cqa_mc_from_series(w_or2_A_frag,   mc_iterations=100)
cqa_or2_all_frag   = run_cqa_mc_from_series(w_or2_ALL_frag, mc_iterations=100)

# resume
def _fmt(x):
    try:
        if x is None or (isinstance(x, float) and (np.isnan(x) or np.isinf(x))): return "–"
        return f"{float(x):.2f}"
    except Exception: return str(x)

print("\n=== OR_2 - FRAGMENTED (Gold* excluded) ===")
print("n (Area A / All):", cqa_or2_A_frag["n"], cqa_or2_all_frag["n"])
print("Best q (Area A / All):", _fmt(cqa_or2_A_frag["quantum_max"]), _fmt(cqa_or2_all_frag["quantum_max"]))
print("Max Φ(q) (Area A / All):", _fmt(cqa_or2_A_frag["phi_max"]), _fmt(cqa_or2_all_frag["phi_max"]))
print("alpha 5% (Area A / All):", _fmt(cqa_or2_A_frag["alpha_5"]), _fmt(cqa_or2_all_frag["alpha_5"]))


# %%

# ============================================================
# FIGURE 49 — CQA of the weight of fragments (Middle Bronze Age,
# Horizon 2) from Area A, compared with the whole sample.
# ============================================================

# ===== CQA - full-width x quarter-height (1x2), dashed alpha line =====
import numpy as np
import matplotlib.pyplot as plt

FIGSIZE_FULL_W_QUARTER_H = (6.30, 2.16)  # ~full page width, 1/4 page height

COL_LINE  = "black"
COL_FILL  = "#D9D9D9"
COL_ALPHA = "#FF7F0E"  # orange for alpha_5
LS_DASH   = (0, (4, 3))  # dash pattern as in the other chart

def _fill_between_zero(ax, x, y, face, alpha=0.35, z=1):
    ax.fill_between(x, 0, y, facecolor=face, alpha=alpha, zorder=z, linewidth=0)

def _style_axes(ax, title, y_lim=(-6, 8)):
    ax.set_xlim(4, 16)
    ax.set_xticks(np.arange(4, 17, 1))
    ax.set_ylim(*y_lim)
    ax.grid(axis="y", color="0.90")
    ax.set_ylabel("φ(q)", fontsize=7)
    ax.set_title(title, loc="left", fontsize=8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.tick_params(labelsize=6)

def plot_cqa_row(ax, result, title, y_lim=(-6, 8), show_alpha=True):
    phi = result["phi_q_df"]
    x = phi["quanta"].to_numpy()
    y = phi["Phi_q_values"].to_numpy()

    _fill_between_zero(ax, x, y, face=COL_FILL, alpha=0.35, z=1)
    ax.plot(x, y, color=COL_LINE, lw=1.0, zorder=2)

    if show_alpha:
        a5 = result.get("alpha_1", np.nan)
        if np.isfinite(a5):
            ax.axhline(a5, color=COL_ALPHA, alpha=0.7, lw=0.9, ls=LS_DASH)

    _style_axes(ax, title, y_lim=y_lim)

# --- figure 1x2: area A + Total sample ---
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=FIGSIZE_FULL_W_QUARTER_H, dpi=300, sharex=True, sharey=True)

plot_cqa_row(ax1, cqa_or2_A_frag,   "Fragm. bronze (Area A)",     y_lim=(-6, 8), show_alpha=True)
plot_cqa_row(ax2, cqa_or2_all_frag, "Fragm. bronze (Tot sample)", y_lim=(-6, 8), show_alpha=True)

ax1.set_xlabel("quanta", fontsize=6)
ax2.set_xlabel("quanta", fontsize=6)

plt.tight_layout()
#plt.savefig("Fig49_CQA_OR2_fullwidth_quarterheight_alpha_dashed.jpg", dpi=300, bbox_inches="tight")
plt.show()


# %%

# ============================================================
# CQA DATA PREPARATION — exploratory OR_2 ingots series
# (no figure in the published volume)
# Series preparation: OR_2 / Artefact = ingot / fragmented condition,
# Area A vs whole sample.
# ============================================================

# ============== CQA - OR_2 / Artefact = ingot / Fragmented cond. - area A vs Whole sample ==============
import pandas as pd
import numpy as np
import math
import re
import unicodedata
import sys

# --- tqdm ---
try:
    from tqdm import tqdm as _tqdm
    def tqdm_iter(it, **kw):
        kw0 = dict(leave=True, dynamic_ncols=True)
        kw0.update(kw)
        return _tqdm(it, file=sys.stdout, **kw0)
except Exception:
    def tqdm_iter(it, **kw): return it

# ---------- helper ----------
def _strip_accents(s):
    s = str(s).replace("ß","ss")
    return "".join(ch for ch in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(ch))

def _norm(s):
    s = _strip_accents(str(s)).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def to_str_u(s):  # upper + strip
    return s.astype(str).str.upper().str.strip()

def col_like(df, *needles):
    cols = list(df.columns)
    lows = [c.lower().strip() for c in cols]
    for n in needles:
        n_low = str(n).lower().strip()
        # match
        for i, c in enumerate(lows):
            if n_low == c:
                return cols[i]
        # match
        for i, c in enumerate(lows):
            if n_low in c:
                return cols[i]
    raise KeyError(f"Column not found for: {needles}")

def is_ingot_only(s_raw: str) -> bool:
    """True solo se il testo contiene la parola 'ingot/ingots' (non esclude/ include 'barren', 'ringbarren', ecc.)."""
    s = _norm(s_raw)
    return re.search(r"\bingot(s)?\b", s) is not None

U = lambda s: str(s).upper().strip()

# ---------- data source ----------
if "raw" in globals():
    df_src = raw.copy()
else:
    for fname in ("DB.xlsx", "DB_SI.xlsx"):
        try:
            df_src = pd.read_excel(fname, sheet_name="DB")
            print(f"[INFO] Caricato: {fname}")
            break
        except Exception:
            df_src = None
    if df_src is None:
        raise FileNotFoundError("Could not find 'raw' in memory or the DB.xlsx / DB_SI.xlsx files on disk.")

# ---------- column mapping ----------
c_or     = col_like(df_src, "Or_fin", "or fin", "period", "or_fin")
c_state  = col_like(df_src, "State")
c_reg    = col_like(df_src, "Region")
c_type   = col_like(df_src, "Artefact", "artifact", "type")
c_weight = col_like(df_src, "Weight_obj", "weight obj", "weight_obj", "weight", "mass", "weight in g", "mass in g")
c_frag   = col_like(df_src, "Fragmented", "fragmented")
c_comp   = col_like(df_src, "Complete", "complete")
c_match  = col_like(df_src, "Matching fr", "matching_fr", "matching fr.", "matching", "matching dr", "matching_dr")

# ---------- DataFrame normalized ----------
d = pd.DataFrame({
    "Or_fin":      to_str_u(df_src[c_or]).str.replace(r"\s+", "_", regex=True),
    "State":       to_str_u(df_src[c_state]),
    "Region":      to_str_u(df_src[c_reg]),
    "Artefact":    df_src[c_type].astype(str).str.strip(),
    "Weight_obj":  pd.to_numeric(df_src[c_weight], errors="coerce"),
    "Fragmented":  pd.to_numeric(df_src[c_frag],  errors="coerce").fillna(0),
    "Complete":    pd.to_numeric(df_src[c_comp],  errors="coerce").fillna(0),
    "Matching_fr": pd.to_numeric(df_src[c_match], errors="coerce").fillna(0),
})

# ---------- required filters ----------
mask_or2   = d["Or_fin"].eq("OR_2")
mask_ingot = d["Artefact"].apply(is_ingot_only)
mask_frag  = (d["Fragmented"] > 0) | ((d["Matching_fr"] > 0) & (d["Complete"] < 1))
mask_range = d["Weight_obj"].between(7, 200, inclusive="left")  # remove if the full range is wanted

filt_base = mask_or2 & mask_ingot & mask_frag & mask_range

# ---------- areas ----------
GER = U("GERMANY")
mask_A = (
    d["Region"].isin({U("BADEN-WÜRTTEMBERG"), U("BAYERN")})
    | d["State"].isin({U("AUSTRIA"), U("SLOVENIA")})
)

# area A vs Whole sample
w_or2_ingot_areaA = (
    d.loc[filt_base & mask_A, "Weight_obj"].dropna().astype(float).values
)
w_or2_ingot_all = (
    d.loc[filt_base, "Weight_obj"].dropna().astype(float).values
)

print(f"[CHECK] n OR_2 ingot — Area A: {w_or2_ingot_areaA.size} | Whole sample: {w_or2_ingot_all.size}")

# ---------- CQA + Monte Carlo ----------
def run_cqa_mc_from_series(series,
                           min_quantum=2, max_quantum=16, step=0.02,
                           mc_parameter=0.15, mc_iterations=1000):
    s = pd.to_numeric(pd.Series(series), errors="coerce").dropna()
    n = int(s.shape[0])
    quanta_arr = np.arange(min_quantum, max_quantum + step, step)

    if n == 0:
        phi_q_df = pd.DataFrame({"Phi_q_values": np.nan, "quanta": quanta_arr})
        return {"phi_q_df": phi_q_df, "quanta_arr": quanta_arr,
                "alpha_1": np.nan, "alpha_5": np.nan,
                "quantum_max": np.nan, "phi_max": np.nan, "n": 0}

    df = pd.DataFrame({'Filtered Values': s})
    sample_size = df.count()
    coeff = 2 / sample_size
    sample_coefficient = float(np.sqrt(coeff).iloc[0])

    # quantogram
    cosine_results = pd.concat([(2 * math.pi * df.iloc[:, 0]) / q for q in quanta_arr], axis=1)
    cosine_results.columns = [f'Results {round(q, 2)}' for q in quanta_arr]
    cosine_results = np.cos(cosine_results)
    cosine_sum_series = cosine_results.sum(axis=0)
    phi_q_df = cosine_sum_series.to_frame(name="Phi_q_values") * sample_coefficient
    phi_q_df["quanta"] = quanta_arr

    # best
    idx_best    = phi_q_df['Phi_q_values'].idxmax()
    quantum_max = float(phi_q_df.loc[idx_best, 'quanta'])
    phi_max     = float(phi_q_df.loc[idx_best, 'Phi_q_values'])

    # Monte Carlo
    def apply_variation(df_copy, mc_parameter):
        percent_diff = df_copy * mc_parameter
        random_matrix = np.random.uniform(-1, 1, size=df_copy.shape)
        variations = random_matrix * percent_diff
        return df_copy + variations

    inter = []
    for _ in tqdm_iter(range(mc_iterations), desc="Monte Carlo", unit="iter"):
        df_copy = df.copy()
        mc_df = apply_variation(df_copy, mc_parameter)
        inter.append(mc_df[df.columns[0]])
    final_mc_df = pd.concat(inter, axis=1)
    final_mc_df.columns = [f"MC_{i+1}" for i in range(mc_iterations)]

    # Phi(q) for each Monte Carlo run
    def phi_for_df(df_mc, quanta_arr):
        vals = []
        for q in quanta_arr:
            cos_mc = np.cos((2 * math.pi * df_mc.iloc[:, 0]) / q)
            vals.append(float(np.sum(cos_mc) * sample_coefficient))
        return vals

    Montecarlo_phi_q_values = []
    for col in final_mc_df.columns:
        Montecarlo_phi_q_values.append(phi_for_df(final_mc_df[[col]], quanta_arr))
    Montecarlo_phi_q = pd.DataFrame(Montecarlo_phi_q_values, columns=quanta_arr).T

    # threshold 1% e 5%
    a1_share = max(1, round(mc_iterations * 0.01))
    a5_share = max(1, round(mc_iterations * 0.05))
    mc_max_column = Montecarlo_phi_q.max()
    alpha_1 = float(mc_max_column.nlargest(a1_share).iloc[-1])
    alpha_5 = float(mc_max_column.nlargest(a5_share).iloc[-1])

    return {
        "phi_q_df": phi_q_df,
        "quanta_arr": quanta_arr,
        "alpha_1": alpha_1,
        "alpha_5": alpha_5,
        "quantum_max": quantum_max,
        "phi_max": phi_max,
        "n": n
    }

# ---------- run and store in NEW variables ----------
cqa_or2_ingot_areaA = run_cqa_mc_from_series(w_or2_ingot_areaA, mc_iterations=100)
cqa_or2_ingot_all   = run_cqa_mc_from_series(w_or2_ingot_all,   mc_iterations=100)

# resume
def _fmt(x):
    try:
        if x is None or (isinstance(x, float) and (np.isnan(x) or np.isinf(x))):
            return "–"
        return f"{float(x):.2f}"
    except Exception:
        return str(x)

print("\n=== CQA OR_2 - Ingot ===")
print("n (Area A / Whole):", cqa_or2_ingot_areaA["n"], cqa_or2_ingot_all["n"])
print("Best q (Area A / Whole):", _fmt(cqa_or2_ingot_areaA["quantum_max"]), _fmt(cqa_or2_ingot_all["quantum_max"]))
print("Max Φ(q) (Area A / Whole):", _fmt(cqa_or2_ingot_areaA["phi_max"]), _fmt(cqa_or2_ingot_all["phi_max"]))
print("α5% (Area A / Whole):", _fmt(cqa_or2_ingot_areaA["alpha_5"]), _fmt(cqa_or2_ingot_all["alpha_5"]))


# %%

# ============================================================
# CQA DATA PREPARATION for Figure 62 (no direct figure)
# Series preparation: exact-weight fragmented ingots, Horizons 3-5,
# Area A vs whole sample, with a Monte Carlo robustness check.
# ============================================================

# === CQA - INGOTS (exact) FRAGMENTED - OR_3/4/5 - area A vs Total + Monte Carlo ===
import pandas as pd, numpy as np, math, re, unicodedata, sys, os
import matplotlib.pyplot as plt

# --- tqdm ---
os.environ["TQDM_NOTEBOOK"] = "0"
try:
    from tqdm import tqdm as _tqdm
    def tqdm_iter(it, **kw):
        kw0 = dict(leave=True, dynamic_ncols=True)
        kw0.update(kw)
        return _tqdm(it, file=sys.stdout, **kw0)
except Exception:
    def tqdm_iter(it, **kw): return it

# ---------- helper ----------
def _strip_accents(s):
    s = str(s).replace("ß","ss")
    return "".join(ch for ch in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(ch))
def _norm(s):
    s = _strip_accents(str(s)).lower()
    s = re.sub(r"[^a-z0-9]+"," ", s)
    return re.sub(r"\s+"," ", s).strip()
def to_str_u(s):  # UPPER + strip
    return s.astype(str).str.upper().str.strip()
def col_like(df, *needles):
    cols = list(df.columns); low = [c.lower().strip() for c in cols]
    for n in needles:
        nlow = str(n).lower().strip()
        # match 
        for i,c in enumerate(low):
            if nlow == c: return cols[i]
        # match 
        for i,c in enumerate(low):
            if nlow in c: return cols[i]
    raise KeyError(f"Column not found for: {needles}")

# ---------- reading ----------
if "raw" in globals():
    df_src = raw.copy()
else:
    for fname in ("DB.xlsx","DB_SI.xlsx"):
        try:
            df_src = pd.read_excel(fname, sheet_name="DB")
            print(f"[INFO] Caricato: {fname}")
            break
        except Exception:
            df_src = None
    if df_src is None:
        raise FileNotFoundError("Could not find 'raw' or the DB.xlsx / DB_SI.xlsx files.")

# ---------- column mapping ----------
c_or     = col_like(df_src, "Or_fin","or fin","period","or_fin")
c_state  = col_like(df_src, "State")
c_reg    = col_like(df_src, "Region")
c_type   = col_like(df_src, "Artefact","artifact","type")
c_weight = col_like(df_src, "Weight_obj","weight obj","weight_obj","weight","mass","weight in g","mass in g")
c_frag   = col_like(df_src, "Fragmented","fragmented")
c_comp   = col_like(df_src, "Complete","complete")
c_match  = col_like(df_src, "Matching fr","matching_fr","matching fr.","matching")

# ---------- dataframe normalized ----------
d = pd.DataFrame({
    "Or_fin":      to_str_u(df_src[c_or]).str.replace(r"\s+","_", regex=True),
    "State":       to_str_u(df_src[c_state]),
    "Region":      to_str_u(df_src[c_reg]),
    "Artefact":    df_src[c_type].astype(str).str.strip(),
    "Weight_obj":  pd.to_numeric(df_src[c_weight], errors="coerce"),
    "Fragmented":  pd.to_numeric(df_src[c_frag],  errors="coerce").fillna(0),
    "Complete":    pd.to_numeric(df_src[c_comp],  errors="coerce").fillna(0),
    "Matching_fr": pd.to_numeric(df_src[c_match], errors="coerce").fillna(0),
})

# ---------- ONLY Artefact = 'ingot' (match, case-insensitive) ----------
def is_ingot_exact(s_raw: str) -> bool:
    return _norm(s_raw) == "ingot"
mask_ingot = d["Artefact"].apply(is_ingot_exact)

# ---------- other filters ----------
mask_period = d["Or_fin"].isin({"OR_3","OR_4","OR_5"})
mask_range  = d["Weight_obj"].between(7, 200, inclusive="left")
mask_frag   = (d["Fragmented"] > 0) | ((d["Matching_fr"] > 0) & (d["Complete"] < 1))

# ---------- areas ----------
U = lambda s: str(s).upper().strip()
GER = U("GERMANY")
mask_A  = (d["Region"].isin({U("BADEN-WÜRTTEMBERG"), U("BAYERN")}) |
           d["State"].isin({U("AUSTRIA"), U("SLOVENIA")}))
mask_B  = (d["Region"].eq(U("BRANDENBURG")) |
           d["State"].eq(U("POLAND")) |
           d["Region"].eq(U("SACHSEN")))
mask_C  = ((d["State"].eq(GER) & d["Region"].isin({U("MECKLENBURG-VORPOMMERN"),
                                                   U("SCHLESWIG-HOLSTEIN"),
                                                   U("NIEDERSACHSEN")})) |
           (d["State"].eq(GER) & d["Region"].isin({U("THÜRINGEN"),
                                                   U("HESSEN"),
                                                   U("RHEINLAND-PFALZ")})) |
           (d["Region"].eq(U("SACHSEN-ANHALT"))))
mask_CH = d["State"].eq(U("SWITZERLAND"))
mask_I  = d["State"].isin({U("ITALY"), U("SAN MARINO")})
mask_union = (mask_A | mask_B | mask_C | mask_CH | mask_I)

# ---------- series: area A vs Total ----------
filt = mask_ingot & mask_period & mask_range & mask_frag
w_ing_AreaA = d.loc[filt & mask_A,      "Weight_obj"].dropna().astype(float).values
w_ing_All   = d.loc[filt & mask_union,  "Weight_obj"].dropna().astype(float).values

print(f"[CHECK] n (Area A ingots fragm.): {w_ing_AreaA.size} | n (Tot sample ingots fragm.): {w_ing_All.size}")

# ---------- CQA + Monte Carlo ----------
def run_cqa_mc_from_series(series,
                           min_quantum=2, max_quantum=16, step=0.02,
                           mc_parameter=0.15, mc_iterations=1000):
    s = pd.to_numeric(pd.Series(series), errors="coerce").dropna()
    if s.empty:
        q = np.arange(min_quantum, max_quantum + step, step)
        return {"phi_q_df": pd.DataFrame({"Phi_q_values": np.nan, "quanta": q}),
                "quanta_arr": q, "alpha_1": np.nan, "alpha_5": np.nan,
                "quantum_max": np.nan, "phi_max": np.nan, "n": 0}

    n = int(s.shape[0])
    q = np.arange(min_quantum, max_quantum + step, step)
    df = pd.DataFrame({"Filtered Values": s})
    coeff = 2 / df.count()
    sample_coeff = float(np.sqrt(coeff).iloc[0])

    cos = pd.concat([(2 * math.pi * df.iloc[:,0]) / qq for qq in q], axis=1)
    cos = np.cos(cos)
    phi_q_df = pd.DataFrame(cos.sum(axis=0), columns=["Phi_q_values"]) * sample_coeff
    phi_q_df["quanta"] = q

    # --- best
    y_vals = phi_q_df['Phi_q_values'].to_numpy()
    x_vals = phi_q_df['quanta'].to_numpy()
    if np.all(np.isnan(y_vals)):
        quantum_max = np.nan
        phi_max     = np.nan
    else:
        idx_best    = int(np.nanargmax(y_vals))
        quantum_max = float(x_vals[idx_best])
        phi_max     = float(y_vals[idx_best])


    # Monte Carlo
    # ---------- Monte Carlo  ----------
    rng = np.random.default_rng()  # optional: pass a seed for reproducibility
    # generate Monte Carlo columns (same logic as the jitter)
    inter = []
    for _ in tqdm_iter(range(mc_iterations), desc="Monte Carlo", unit="iter"):
        noise = rng.uniform(-1, 1, size=(df.shape[0], 1)) * (df * mc_parameter).to_numpy()
        inter.append((df + noise)[df.columns[0]])
    mc_df = pd.concat(inter, axis=1)          # shape: (n, mc_iterations)

    # quantogram for all q values and all iterations at once
    arr = mc_df.to_numpy()                    # (n, K)
    q_col = np.array(q, dtype=float)[:, None] # (Q,1)
    # cos((2*pi/q) * x) --> for broadcasting: (Q,1) over (1,n) via arr.T
    # sum over n (axis=1 after cos, since we are operating on arr.T)
    cos_sums = []
    for qq in q:
        cos_sums.append(np.cos((2 * np.pi / qq) * arr).sum(axis=0))
    cos_sums = np.vstack(cos_sums)            # (Q, K)
    MC = cos_sums * sample_coeff              # (Q, K) Monte Carlo phi(q) values

    # quantogram maxima per iteration (K,)
    mc_max = MC.max(axis=0)

    # thresholds (quantiles) - use 'method' for numpy >= 1.22, otherwise 'interpolation'
    alpha_1 = float(np.quantile(mc_max, 0.99, method="linear"))
    alpha_5 = float(np.quantile(mc_max, 0.95, method="linear"))

    # (for later inspection if needed)
    # return anche 'mc_max' dentro il dict


    return {"phi_q_df": phi_q_df, "quanta_arr": q,
            "alpha_1": alpha_1, "alpha_5": alpha_5,
            "quantum_max": quantum_max, "phi_max": phi_max, "n": n}

MC_ITERS = 1000  
cqa_ing_AreaA = run_cqa_mc_from_series(w_ing_AreaA, mc_iterations=MC_ITERS)
cqa_ing_All   = run_cqa_mc_from_series(w_ing_All,   mc_iterations=MC_ITERS)

print("Best q (AreaA / All):",
      f"{cqa_ing_AreaA['quantum_max']:.2f}", f"{cqa_ing_All['quantum_max']:.2f}")
print("alpha 5% (AreaA / All):",
      f"{cqa_ing_AreaA['alpha_5']:.2f}", f"{cqa_ing_All['alpha_5']:.2f}")

# ---------- PLOT 1x2 (black line + gray fill + dashed orange alpha_5) ----------
COL_LINE="black"; COL_FILL="#D9D9D9"; COL_ALPHA="#FF7F0E"; LS_DASH=(0,(4,3))
def _fill_between_zero(ax, x, y, face, alpha=0.35, z=1):
    ax.fill_between(x, 0, y, facecolor=face, alpha=alpha, zorder=z, linewidth=0)

def plot_panel(ax, result, title, y_lim=(-6,9)):
    phi = result["phi_q_df"]
    x = phi["quanta"].to_numpy(); y = phi["Phi_q_values"].to_numpy()
    _fill_between_zero(ax, x, y, face=COL_FILL, alpha=0.35, z=1)
    ax.plot(x, y, color=COL_LINE, lw=1.1, zorder=2)
    if np.isfinite(result.get("alpha_5", np.nan)):
        ax.axhline(result["alpha_5"], color=COL_ALPHA, ls=LS_DASH, lw=0.9)
    ax.set_xlim(4,16); ax.set_xticks(np.arange(4,17,1))
    ax.set_ylim(*y_lim); ax.grid(axis="y", color="0.9")
    ax.set_ylabel("φ(q)"); ax.set_xlabel("quanta")
    ax.set_title(title, loc="left", fontsize=12)
    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)

fig, (axL, axR) = plt.subplots(1,2, figsize=(11,3.2), dpi=300, sharey=True)
plot_panel(axL, cqa_ing_AreaA, "Fragm. ingots (Area A)")
plot_panel(axR, cqa_ing_All,   "Fragm. ingots (Tot sample)")
plt.tight_layout(w_pad=1.4)
plt.show()


# %%

# ============================================================
# FIGURE 62 — CQA of fragmented ingots (casting cakes) from the
# Late Bronze Age (Horizons 3-5), Area A vs the whole sample.
# ============================================================

# ===== CQA - Fragm. ingots (area A vs Total sample) - 1/4 page + JPEG export =====
import numpy as np
import matplotlib.pyplot as plt

# 1/4 page ~ 160 x 55 mm
FIGSIZE_QUARTER = (6.30, 2.16)   # inch
DPI = 300

COL_LINE  = "black"
COL_FILL  = "#D9D9D9"
COL_ALPHA = "#FF7F0E"
LS_DASH   = (0, (4, 3))

def _fill_between_zero(ax, x, y, face=COL_FILL, alpha=0.35, z=1):
    ax.fill_between(x, 0, y, facecolor=face, alpha=alpha, zorder=z, linewidth=0)

def _style_axes(ax, title):
    ax.set_xlim(4, 16)
    ax.set_xticks(np.arange(4, 17, 1))
    ax.set_ylim(-6, 8)
    ax.set_yticks(np.arange(-6, 9, 2)) 
    ax.grid(axis="y", color="0.90")
    ax.set_ylabel("φ(q)")
    ax.set_title(title, loc="left", fontsize=12)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

def plot_panel(ax, result, title, show_alpha=True):
    phi = result["phi_q_df"]
    x = phi["quanta"].to_numpy()
    y = phi["Phi_q_values"].to_numpy()

    _fill_between_zero(ax, x, y)
    ax.plot(x, y, color=COL_LINE, lw=1.1, zorder=2)

    if show_alpha and np.isfinite(result.get("alpha_1", np.nan)):
        ax.axhline(result["alpha_1"], color=COL_ALPHA, ls=LS_DASH, lw=0.9)

    _style_axes(ax, title)

fig, (ax_l, ax_r) = plt.subplots(1, 2, figsize=FIGSIZE_QUARTER, dpi=DPI, sharex=True, sharey=True)

plot_panel(ax_l, cqa_ing_AreaA, "Fragm. ingots (Area A)",       show_alpha=True)
plot_panel(ax_r, cqa_ing_All,   "Fragm. ingots (Tot sample)",   show_alpha=True)

# etichetta X
ax_l.set_xlabel("quanta", fontsize=9)
ax_r.set_xlabel("quanta", fontsize=9)

plt.tight_layout(pad=0.6, w_pad=0.9)

# JPEG export (high quality)
#plt.savefig(
#    "Fig62_CQA_ingots_OR345_MC_quarterpage.jpg",
#    dpi=DPI, bbox_inches="tight", format="jpeg",
#    pil_kwargs={"quality": 95, "subsampling": 0}
#)
plt.show()


# %%

# ============================================================
# SUPPORTING TABLE (not a figure) — Summary statistics for the
# CQA of fragmented ingots (Area A vs whole sample): n, best quantum,
# max phi(q), and alpha thresholds at 1% and 5%.
# ============================================================

import numpy as np
import pandas as pd

def _fmt(x):
    try:
        return "–" if x is None or (isinstance(x, float) and (np.isnan(x) or np.isinf(x))) else f"{float(x):.2f}"
    except Exception:
        return str(x)

summary = pd.DataFrame({
    "n":            [cqa_ing_AreaA["n"],         cqa_ing_All["n"]],
    "Best q":       [cqa_ing_AreaA["quantum_max"], cqa_ing_All["quantum_max"]],
    "Max φ(q)":     [cqa_ing_AreaA["phi_max"],   cqa_ing_All["phi_max"]],
    "α1%":          [cqa_ing_AreaA["alpha_1"],   cqa_ing_All["alpha_1"]],
    "α5%":          [cqa_ing_AreaA["alpha_5"],   cqa_ing_All["alpha_5"]],
}, index=["Fragm. ingots (Area A)","Fragm. ingots (Tot sample)"])

print(summary.applymap(_fmt).to_string())


# %%

