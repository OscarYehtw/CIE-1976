import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import colour
from colour.plotting import colour_style
import matplotlib.pyplot as plt
from matplotlib.legend_handler import HandlerPatch
from matplotlib.patches import Ellipse, Patch, PathPatch
import mpld3
from mpld3 import plugins
import numpy as np
from matplotlib.path import Path
from matplotlib.lines import Line2D

class TickStylePlugin(plugins.PluginBase):
    JAVASCRIPT = """
    mpld3.register_plugin("tickstyle", TickStylePlugin);
    TickStylePlugin.prototype = Object.create(mpld3.Plugin.prototype);
    TickStylePlugin.prototype.constructor = TickStylePlugin;
    function TickStylePlugin(fig, props){
        mpld3.Plugin.call(this, fig, props);
        this.fig = fig;
    }
    TickStylePlugin.prototype.draw = function(){
        d3.selectAll("g.tick line").style("stroke-dasharray", "2,2")
                                   .style("stroke-width", "0.5")
                                   .style("stroke", "gray");
    };
    """
    def __init__(self):
        self.dict_ = {"type": "tickstyle"}

class CrosshairPlugin(plugins.PluginBase):
    JAVASCRIPT = """
    mpld3.register_plugin("crosshair", CrosshairPlugin);
    CrosshairPlugin.prototype = Object.create(mpld3.Plugin.prototype);
    CrosshairPlugin.prototype.constructor = CrosshairPlugin;

    function CrosshairPlugin(fig, props){
        mpld3.Plugin.call(this, fig, props);
        this.fig = fig;
        this.ax = fig.axes[0];
    }

    CrosshairPlugin.prototype.draw = function(){
        var fig = this.fig;
        var ax = this.ax;

        // var svg = d3.select(fig.canvas);
        var svg = d3.select("svg.mpld3-figure");

        var width = +svg.attr("width");
        var height = +svg.attr("height");

        this.vline = svg.append("line")
            .attr("stroke", "gray")
            .attr("stroke-width", 1)
            .attr("pointer-events", "none")
            .style("display", "none");

        this.hline = svg.append("line")
            .attr("stroke", "gray")
            .attr("stroke-width", 1)
            .attr("pointer-events", "none")
            .style("display", "none");

        var vline = this.vline;
        var hline = this.hline;

        svg.on("mousemove", function(event){
            var coords = d3.mouse(this);
            var x = ax.x.invert(coords[0]);
            var y = ax.y.invert(coords[1]);

            if(x < 0 || x > 0.7 || y < 0 || y > 0.7){
                vline.style("display", "none");
                hline.style("display", "none");
                return;
            }

            var px = ax.x(x);
            var py = ax.y(y);

            vline
                .attr("x1", px).attr("y1", 0)
                .attr("x2", px).attr("y2", height)
                .style("display", null);

            hline
                .attr("x1", 0).attr("y1", py)
                .attr("x2", width).attr("y2", py)
                .style("display", null);
        });

        svg.on("mouseleave", function(){
            vline.style("display", "none");
            hline.style("display", "none");
        });
    };
    """

    def __init__(self):
        self.dict_ = {"type": "crosshair"}

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
            var rect = this.getBoundingClientRect();
            var mouseX = d3.event.clientX - rect.left;
            var mouseY = d3.event.clientY - rect.top;

            var x = ax.x.invert(mouseX - ax.position[0]);
            var y = ax.y.invert(mouseY - ax.position[1]);

            if (x >= 0 && x <= 0.7 && y >= 0 && y <= 0.7) {
                coordsDiv.innerHTML = "u′ = " + x.toFixed(4) + ", v′ = " + y.toFixed(4);
                coordsDiv.style.display = "block";
            } else {
                coordsDiv.style.display = "none";
            }
        });

        fig.canvas.on("mouseleave", function(){
            coordsDiv.innerHTML = "u′ = ---, v′ = ---";
            coordsDiv.style.display = "block";
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

# === Reading CSV ===
#csv_path = "r-255_g-0_b-0_L-1to255period10.csv"
#csv_path = "r-0_g-255_b-0_L-1to255period10.csv"
csv_path = "r-0_g-0_b-255_L-1to255period10.csv"
df = pd.read_csv(csv_path)

# Specific columns for u′ and v′
u_values = df["u`"].values
v_values = df["v`"].values

# === Step 4: Plot base CIE1976 diagram ===
colour_style()
fig, ax = plt.subplots(figsize=(8, 6))
colour.plotting.plot_chromaticity_diagram_CIE1976UCS(standalone=False, axes=ax)

# === Step 5: Draw ellipses ===
for i, (u, v, major, minor, angle) in enumerate(ellipses):
    ellipse = Ellipse(xy=(u, v), width=major, height=minor, angle=angle,
                      edgecolor='black', facecolor='none', linewidth=0.5,
                      label=os.path.basename(xlsx_path) if i == 0 else None)
    ax.add_patch(ellipse)

# === Draw point ===
#scatter = ax.scatter(u_values, v_values, color='black', s=5, alpha=0.7, label=os.path.basename(csv_path))

# === Function to check whether a point lies inside an ellipse ===
def is_point_in_ellipse(x, y, cx, cy, major, minor, angle_deg):
    angle = np.radians(angle_deg)
    cos_a = np.cos(angle)
    sin_a = np.sin(angle)
    
    dx = x - cx
    dy = y - cy

    # Rotate the coordinates
    x_rot = dx * cos_a + dy * sin_a
    y_rot = -dx * sin_a + dy * cos_a

    # Use the standard ellipse equation
    return (x_rot / (major / 2))**2 + (y_rot / (minor / 2))**2 <= 1

# === Classify each point as PASS / FAIL ===
pass_x, pass_y, fail_x, fail_y = [], [], [], []

for x, y in zip(u_values, v_values):
    if any(is_point_in_ellipse(x, y, u, v, major, minor, angle) for (u, v, major, minor, angle) in ellipses):
        pass_x.append(x)
        pass_y.append(y)
    else:
        fail_x.append(x)
        fail_y.append(y)

# === Plot the points ===
scatter_pass = ax.scatter(pass_x, pass_y, color='green', s=20, alpha=0.7, label="PASS")
scatter_fail = ax.scatter(fail_x, fail_y, color='red', s=20, alpha=0.7, label="FAIL")

# === Tag each point with PASS / FAIL and add the result to the original DataFrame ===
results = []
for x, y in zip(u_values, v_values):
    if any(is_point_in_ellipse(x, y, u, v, major, minor, angle) for (u, v, major, minor, angle) in ellipses):
        results.append("PASS")
    else:
        results.append("FAIL")

df["Result"] = results

# === Output result CSV ===
result_csv_path = os.path.splitext(csv_path)[0] + "_result.csv"
df.to_csv(result_csv_path, index=False)

# === Define legend ===
def make_ellipse_legend(legend, orig_handle, xdescent, ydescent, width, height, fontsize):
    return Ellipse((width / 2, height / 2), width=10, height=5, edgecolor='black', facecolor='none', linewidth=1)

handles, labels = ax.get_legend_handles_labels()
by_label = dict(zip(labels, handles))
ax.legend(by_label.values(), by_label.keys(), loc='upper right', fontsize=9,
          handler_map={Ellipse: HandlerPatch(patch_func=make_ellipse_legend)})

plt.axis([0, 0.7, 0, 0.7])
plt.xlabel("u′")
plt.ylabel("v′")
plt.title("CIE 1976 Chromaticity Diagram")
plt.tight_layout()

# === Step 6: Add live cursor coordinates plugin ===
fig = plt.gcf()
plugins.connect(fig, MousePositionPlugin())
plugins.connect(fig, TickStylePlugin())
#plugins.connect(fig, CrosshairPlugin())

# === Step 7: Save HTML ===
#html_path = xlsx_path.replace('.xlsx', '_cie1976.html')
html_path = 'cie1976_out.html'
with open(html_path, "w", encoding="utf-8") as f:
    f.write(mpld3.fig_to_html(fig))

print(f"✅ HTML with dynamic u′, v′ tracking saved to: {html_path}")
plt.show()
