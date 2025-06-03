#pip install pandas matplotlib colour-science mpld3 openpyxl
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import colour
from colour.plotting import colour_style, plot_chromaticity_diagram_CIE1976UCS
import matplotlib.pyplot as plt
from matplotlib.patches import Ellipse
import mpld3
from mpld3 import plugins

# === Step 1: Open Excel manually ===
root = tk.Tk()
root.withdraw()
xlsx_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
if not xlsx_path:
    raise SystemExit("No Excel file selected.")

# === Step 2: Read the 'LED SPEC' sheet (no header, as it's custom layout) ===
df = pd.read_excel(xlsx_path, sheet_name="LED SPEC", header=None)

# === Step 3: Helper to extract ellipses from the calibration section ===
def extract_ellipses(df):
    ellipses = []
    for i in range(len(df)):
        cell = str(df.iat[i, 0])
        if "Target Coordinates" in cell:
            try:
                # Parse center u', v'
                center_str = str(df.iat[i + 1, 1])
                u, v = map(float, center_str.split(','))

                # Parse ellipse axes and rotation angle
                major = float(df.iat[i + 2, 1])
                minor = float(df.iat[i + 3, 1])
                angle = float(df.iat[i + 4, 1])
                
                ellipses.append((u, v, major, minor, angle))
            except Exception as e:
                print(f"⚠️ Skipping row {i}: {e}")
                continue
    return ellipses

ellipses = extract_ellipses(df)

# === Step 4: Plot base CIE1976 chart ===
colour_style()
plot_chromaticity_diagram_CIE1976UCS(standalone=False)
ax = plt.gca()

def format_uv_coord(x, y):
    return f"u'={x:.4f}, v'={y:.4f}"
ax.format_coord = format_uv_coord

# === Step 5: Plot ellipses ===
for u, v, major, minor, angle in ellipses:
    ellipse = Ellipse(xy=(u, v), width=major, height=minor, angle=angle,
                      edgecolor='black', facecolor='none', linewidth=1)
    ax.add_patch(ellipse)

plt.axis([-0.1, 0.7, -0.1, 0.7])
plt.xlabel("u'")
plt.ylabel("v'")
plt.title("CIE 1976 Chromaticity Diagram")
plt.subplots_adjust(left=0.1, right=0.95, top=0.93, bottom=0.13)
plt.tight_layout()

# === Step 6: Output HTML ===
html_path = xlsx_path.replace('.xlsx', '_cie1976.html')
with open(html_path, "w") as f:
    f.write(mpld3.fig_to_html(plt.gcf()))

print(f"✅ Chart saved to: {html_path}")
plt.show()
