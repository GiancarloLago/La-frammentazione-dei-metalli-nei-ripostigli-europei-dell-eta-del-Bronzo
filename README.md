# La-frammentazione-dei-metalli-nei-ripostigli-europei-dell-et-del-Bronzo
Code and database for Lago (2026), La frammentazione dei metalli nei ripostigli europei dell'età del Bronzo (Scienze dell'Antichità). Reproduces all data visualisations in the volume. Python/Jupyter; requires pandas, numpy, matplotlib. Source database in Excel (DB.xlsx)

Code used to generate the data visualizations published in:
> Lago, G. *La frammentazione dei metalli nei ripostigli europei dell'età del Bronzo*
> (Scienze dell'Antichità monografie 8).
Contents
`EuropeanHoardsLago2026.ipynb` — Jupyter notebook, one cell per figure (or per
group of figures sharing the same code), in the order they appear in the book.
`EuropeanHoardsLago2026.py` — the same code as a single Python script, with
`# %%` cell markers (compatible with VS Code / Spyder cell execution).
`DB.xlsx` — source database (not included in this repository; see below).
Each cell/section is labelled with the number of the figure it reproduces,
e.g. `FIGURE 17`, or `FIGURES 50, 51, 52, 53, 54` when a single block of code
generates one figure per area. A few cells produce a supporting table rather
than a figure; these are labelled `SUPPORTING TABLE (not a figure)`. A few
data-preparation cells with no plotting output of their own (feeding a later
CQA figure) are labelled `CQA DATA PREPARATION`.
Requirements
```
pandas
numpy
matplotlib
openpyxl
python-docx      # only used by the supporting-table export section
tqdm             # optional; used for a progress bar during one Monte Carlo
                 # simulation, with a silent fallback if not installed
```
Install with:
```
pip install pandas numpy matplotlib openpyxl python-docx tqdm
```
Usage
Place `DB.xlsx` in the same directory as the script/notebook, then either:
open `Figure_Monografia_EN.ipynb` in Jupyter and run cells in order, or
run `Figure_Monografia_EN.py` directly (e.g. `python Figure_Monografia_EN.py`),
or open it in an editor that supports `# %%` cell execution.
All `savefig` calls are commented out by default (`#plt.savefig(...)`), so
running the script only displays the figures inline without writing files to
disk. To export a figure, uncomment the relevant `savefig` line; output
filenames are prefixed with the corresponding figure number, e.g.
`Fig17_Areas_ABC_fullpage.jpg`.
Figure index
Figure(s) in the book	Section title in the code
Fig. 4	Comparison of Nordic and Central European chronologies
Fig. 5	Quantity of fragments, damaged and intact objects
Fig. 6	Share of fragmentation by period, Nordic vs Central European chronology
Fig. 16	Summary of intact/fragmented objects by regions, states, and areas
Fig. 17	Quantification of fragmentation/integrity, areas A/B/C
Fig. 20	Quantification of complete/fragmented/damaged objects by area and period
Fig. 21	Deformation of complete or fragmented objects
Fig. 22	Estimated preserved portion of fragmented objects
Fig. 25, 29, 32, 35	Frequency distribution of total hoard weight, by area
Fig. 27, 30, 33, 36	Boxplots of total hoard weight, by area
Fig. 38	Published and estimated quantity of metal deposited
Fig. 39	Average weight per year of metal hoards
Fig. 41	Frequency distribution of Ösenhalsringe (Ringbarren) weight
Fig. 43	Frequency distribution of Spangenbarren weight
Fig. 46	Frequency distribution of flanged-axe weight (Horizon 1)
Fig. 47	Frequency distribution of complete objects, Area A (Horizon 2)
Fig. 48	Frequency distribution of fragments, Area A (Horizon 2)
Fig. 49	CQA of fragments, Area A vs whole sample (Horizon 2)
Fig. 50–54	Frequency distribution of complete objects (Horizons 3–5), by area
Fig. 55	CQA of complete objects, areas A/B/C/Italy (Horizons 3–5)
Fig. 56	Frequency distribution of complete ingots (Horizons 3–5)
Fig. 57–61	Frequency distribution of fragmented ingots, by area
Fig. 62	CQA of fragmented ingots, Area A vs whole sample
Fig. 63–67	Frequency distribution of fragmented objects, by area
Fig. 68	Fragmented sickles vs other fragmented objects
Fig. 69	Frequency distribution of fragmented objects, stacked by area
Fig. 70	Frequency distribution of balance-weight weight
Fig. 71	CQA of fragments (Italy, Central Europe, balance weights)
Notes
A handful of exploratory cells from the original working notebook
(alternative/discarded chart variants not used in the final volume) have
been removed from this version for clarity.
Column names from the source spreadsheet are resolved through small
"fuzzy" helper functions (`col_like`, alias lists, etc.) so the code keeps
working even if a column header changes slightly (capitalization, spacing,
British vs. American spelling).
