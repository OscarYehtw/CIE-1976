import tkinter as tk
from tkinter import filedialog
import pandas as pd
import colour
from colour.plotting import colour_style
import matplotlib.pyplot as plt
from matplotlib.patches import Ellipse, PathPatch
import mpld3
from mpld3 import plugins
import numpy as np
from matplotlib.path import Path

class MousePositionPlugin(plugins.PluginBase):
    """Plugin to display u′ and v′ coordinates on hover (D3 v3 compatible)."""
    JAVASCRIPT = """
    mpld3.register_plugin("mouseposition", MousePositionPlugin);
    MousePositionPlugin.prototype = Object.create(mpld3.Plugin.prototype);
    MousePositionPlugin.prototype.constructor = MousePositionPlugin;

    function MousePositionPlugin(fig, props){
        mpld3.Plugin.call(this, fig, props);
        this.fig = fig;
    }

    MousePositionPlugin.prototype.draw = function(){
        var fig = this.fig;
        var ax = fig.axes[0];

        var coordsDiv = document.createElement("div");
        coordsDiv.style.position = "absolute";
        coordsDiv.style.top = "10px";
        coordsDiv.style.left = "10px";
        coordsDiv.style.padding = "5px 8px";
        coordsDiv.style.background = "rgba(255,255,255,0.9)";
        coordsDiv.style.border = "1px solid #ccc";
        coordsDiv.style.borderRadius = "5px";
        coordsDiv.style.font = "12px sans-serif";
        coordsDiv.style.zIndex = 1000;
        coordsDiv.innerHTML = "u′ = ---, v′ = ---";
        document.body.appendChild(coordsDiv);

        fig.canvas.on("mousemove", function(){
            var coords = d3.mouse(this);
            var x = ax.x.invert(coords[0]);
            var y = ax.y.invert(coords[1]);
            coordsDiv.innerHTML = "u′ = " + x.toFixed(4) + ", v′ = " + y.toFixed(4);
        });

        fig.canvas.on("mouseleave", function(){
            coordsDiv.innerHTML = "u′ = ---, v′ = ---";
        });
    };
    """
    def __init__(self):
        self.dict_ = {"type": "mouseposition"}

# === Step 1: Select Excel file ===
root = tk.Tk()
root.withdraw()
xlsx_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
if not xlsx_path:
    raise SystemExit("No Excel file selected.")

# === Step 2: Read Excel sheet without headers ===
df = pd.read_excel(xlsx_path, sheet_name="LED SPEC", header=None)

# === Step 3: Extract ellipse data ===
def extract_ellipses(df):
    ellipses = []
    for i in range(len(df)):
        if "Target Coordinates" in str(df.iat[i, 0]):
            try:
                u, v = map(float, str(df.iat[i + 1, 1]).split(','))
                major = float(df.iat[i + 2, 1])
                minor = float(df.iat[i + 3, 1])
                angle = float(df.iat[i + 4, 1])
                ellipses.append((u, v, major, minor, angle))
            except Exception as e:
                print(f"⚠️ Skipping row {i}: {e}")
    return ellipses

ellipses = extract_ellipses(df)

# === Step 4: Plot base CIE1976 diagram ===
colour_style()
fig, ax = plt.subplots(figsize=(8, 6))
colour.plotting.plot_chromaticity_diagram_CIE1976UCS(standalone=False, axes=ax)

# === Step 5: Draw ellipses ===
for u, v, major, minor, angle in ellipses:
    ellipse = Ellipse(xy=(u, v), width=major, height=minor, angle=angle,
                      edgecolor='black', facecolor='none', linewidth=1)
    ax.add_patch(ellipse)

plt.axis([0, 0.7, 0, 0.7])
plt.xlabel("u′")
plt.ylabel("v′")
plt.title("CIE 1976 Chromaticity Diagram")
plt.tight_layout()

# === Step 6: Add live cursor coordinates plugin ===

# === Step 7: Save HTML ===
fig = plt.gcf()
plugins.connect(fig, MousePositionPlugin())

html_path = xlsx_path.replace('.xlsx', '_cie1976.html')
with open(html_path, "w", encoding="utf-8") as f:
    f.write(mpld3.fig_to_html(fig))

print(f"✅ HTML with dynamic u′, v′ tracking saved to: {html_path}")
plt.show()
