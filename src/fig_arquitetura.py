"""Diagrama da arquitetura do SolarSync e da camada de ML deste trabalho."""
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

FIG = Path(__file__).resolve().parents[1] / "reports" / "figures"
AZUL, AQUA, EIXO, TXT = "#2a78d6", "#1baf7a", "#898781", "#0b0b0b"


def caixa(ax, x, y, w, h, texto, cor, fill):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.02,rounding_size=0.06",
                                fc=fill, ec=cor, lw=1.6))
    ax.text(x + w / 2, y + h / 2, texto, ha="center", va="center", fontsize=8, color=TXT)


def seta(ax, x0, y0, x1, y1, cor=EIXO, ls="-"):
    ax.add_patch(FancyArrowPatch((x0, y0), (x1, y1), arrowstyle="-|>", mutation_scale=11,
                                 color=cor, lw=1.4, linestyle=ls))


fig, ax = plt.subplots(figsize=(6.3, 3.3))
ax.set_xlim(0, 10); ax.set_ylim(0, 5.2); ax.axis("off")

# infraestrutura existente (linha superior)
nomes = ["Inversor híbrido\n(Anenji)", "MAX485 + ESP32\n(RS485 / Modbus)", "API Node.js", "MongoDB Atlas\n(telemetries)", "Dashboard\nReact + Vite"]
xs = [0.1, 2.1, 4.1, 6.1, 8.1]
for x, n in zip(xs, nomes):
    caixa(ax, x, 3.5, 1.7, 1.1, n, AZUL, "#eaf2fc")
for a, b in zip(xs[:-1], xs[1:]):
    seta(ax, a + 1.7, 4.05, b, 4.05)
ax.text(5.0, 4.95, "Infraestrutura existente do SolarSync", ha="center", fontsize=8, color="#52514e", style="italic")

# camada de ML (este trabalho)
caixa(ax, 3.9, 0.5, 3.4, 1.2, "Camada de aprendizagem de máquina\n(este trabalho):\nEDA · previsão · anomalias", AQUA, "#e6f6ef")
seta(ax, 6.9, 3.5, 6.2, 1.7, cor=AQUA)          # Mongo -> ML (somente leitura)
ax.text(6.95, 2.6, "consulta\n(somente leitura)", fontsize=7, color="#52514e", ha="left", va="center")
caixa(ax, 8.0, 0.5, 1.8, 1.2, "Resultados\n(V2: integração\nao dashboard)", AQUA, "#ffffff")
seta(ax, 7.3, 1.1, 8.0, 1.1, cor=AQUA, ls="--")
seta(ax, 8.9, 1.7, 8.9, 3.5, cor=AQUA, ls="--")

fig.tight_layout(); fig.savefig(FIG / "fig_arquitetura.png", dpi=200); plt.close(fig)
print("ok")
