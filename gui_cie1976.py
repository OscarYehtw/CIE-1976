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

# === Step 4: Plot base CIE1976 diagram ===
colour_style()
fig, ax = plt.subplots(figsize=(8, 6))
colour.plotting.plot_chromaticity_diagram_CIE1976UCS(standalone=False, axes=ax)

# === Step 5: Draw ellipses ===
for u, v, major, minor, angle in ellipses:
    ellipse = Ellipse(xy=(u, v), width=major, height=minor, angle=angle,
                      edgecolor='black', facecolor='none', linewidth=0.5)
    ax.add_patch(ellipse)

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
