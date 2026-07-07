import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np
import sys

# -----------------------------
# DYNAMIC VALUE FROM EXCEL (0 - 100)
# -----------------------------
# Read the value passed from the main script, fallback to 100 if something goes wrong
value = int(float(sys.argv[1])) if len(sys.argv) > 1 else 100

fig, ax = plt.subplots(figsize=(10, 6))
ax.set_aspect('equal')
ax.axis('off')

# -----------------------------
# OUTER GREY ARC
# -----------------------------
outer = patches.Wedge(
    (0, 0),
    1.05,
    0,
    180,
    width=0.09,
    facecolor="#d8d8e5",
    edgecolor="none"
)
ax.add_patch(outer)

# -----------------------------
# RED ARC
# -----------------------------
red = patches.Wedge(
    (0, 0),
    0.93,
    108,
    180,
    width=0.10,
    facecolor="#b13a3a",
    edgecolor="none"
)
ax.add_patch(red)

# -----------------------------
# YELLOW ARC
# -----------------------------
yellow = patches.Wedge(
    (0, 0),
    0.93,
    72,
    108,
    width=0.10,
    facecolor="#f1b434",  
    edgecolor="none"
)
ax.add_patch(yellow)

# -----------------------------
# GREEN ARC
# -----------------------------
green = patches.Wedge(
    (0, 0),
    0.93,
    0,
    72,
    width=0.10,
    facecolor="#238b45",
    edgecolor="none"
)
ax.add_patch(green)

# -----------------------------
# TICK MARKS
# -----------------------------
for i in range(11):
    angle = np.deg2rad(180 - i * 18)

    x1 = 0.97 * np.cos(angle)
    y1 = 0.97 * np.sin(angle)

    x2 = 1.05 * np.cos(angle)
    y2 = 1.05 * np.sin(angle)

    ax.plot([x1, x2], [y1, y2],
            color="white",
            linewidth=3)

# -----------------------------
# NUMBERS
# -----------------------------
for i in range(11):
    angle = np.deg2rad(180 - i * 18)

    x = 1.18 * np.cos(angle)
    y = 1.18 * np.sin(angle)

    ax.text(
        x,
        y,
        str(i * 10),
        fontsize=18,
        ha="center",
        va="center",
        color="#222222"
    )

# -----------------------------
# NEEDLE
# -----------------------------
needle_angle = np.deg2rad(180 - value * 1.8)

x = 0.95 * np.cos(needle_angle)
y = 0.95 * np.sin(needle_angle)

ax.plot(
    [0, x],
    [0, y],
    color="#222222",
    linewidth=8,
    solid_capstyle="round"
)

ax.add_patch(
    plt.Circle(
        (0, 0),
        0.05,
        color="#222222"
    )
)

# -----------------------------
# CENTER VALUE
# -----------------------------
ax.text(
    0,
    0.33,
    f"{value}%",
    fontsize=46,
    fontweight="bold",
    ha="center",
    color="#222222"
)

# -----------------------------
# LABEL BOX
# -----------------------------
box = patches.FancyBboxPatch(
    (-0.55, -0.42),
    1.10,
    0.16,
    boxstyle="round,pad=0.02",
    facecolor="#e8e8ef",
    edgecolor="none"
)

ax.add_patch(box)

ax.text(
    0,
    -0.34,
    "% Cycle SLA MET",
    fontsize=20,
    fontweight="bold",
    ha="center",
    color="#222222"
)

# -----------------------------
# LIMITS
# -----------------------------
ax.set_xlim(-1.3, 1.3)
ax.set_ylim(-0.5, 1.25)

# -----------------------------
# SAVE IMAGE
# -----------------------------
plt.savefig("cycle_gauge.png", dpi=300, bbox_inches="tight")

print("Image saved successfully as cycle_gauge.png")

