"""Gráficos sobrios aptos para APA 7 (serif, sin cuadrícula vertical, paleta legible en grises)."""
from __future__ import annotations

import io
import os
import tempfile

def _mpl_dir() -> str:
    """Carpeta de caché de matplotlib, con reservas si el temporal del sistema no es usable."""
    home = os.path.expanduser("~")
    local = os.environ.get("LOCALAPPDATA") or os.path.join(home, "AppData", "Local")
    candidates = []
    try:
        candidates.append(os.path.join(tempfile.gettempdir(), "apa7_mpl"))
    except Exception:
        pass
    candidates += [
        os.path.join(local, "apa7-mcp", "mpl"),
        os.path.join(home, ".apa7_mpl"),
        os.path.join(os.getcwd(), ".apa7_mpl"),
    ]
    for path in candidates:
        try:
            os.makedirs(path, exist_ok=True)
            return path
        except Exception:
            continue
    return candidates[-1]


COLORS = ["#1f3a5f", "#8a8a8a", "#c9c9c9", "#3b7ea1", "#4d4d4d", "#a9c4d6"]


def render_chart(spec: dict) -> bytes:
    """spec: {type: bar|line|pie|scatter|hist, labels[], series[{name, values[]}], x_label?, y_label?}"""
    # matplotlib tarda ~1 s en importarse: solo se carga cuando realmente hay que dibujar un gráfico.
    os.environ.setdefault("MPLCONFIGDIR", _mpl_dir())
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    plt.rcParams["font.family"] = "serif"
    plt.rcParams["font.serif"] = ["Times New Roman", "Liberation Serif", "DejaVu Serif"]
    plt.rcParams["font.size"] = 12
    kind = spec.get("type", "bar")
    labels = spec.get("labels", [])
    series = spec.get("series", [])
    fig, ax = plt.subplots(figsize=(6.5, 4.0), dpi=200)

    if kind == "pie":
        vals = series[0]["values"]
        ax.pie(vals, labels=labels, colors=COLORS, autopct="%1.0f%%", startangle=90, textprops={"color": "black"},
               wedgeprops={"edgecolor": "white"})
        ax.axis("equal")
    elif kind == "hist":
        ax.hist(series[0]["values"], bins=spec.get("bins", 10), color=COLORS[0], edgecolor="white")
    elif kind == "scatter":
        for k, s in enumerate(series):
            ax.scatter(labels if labels else range(len(s["values"])), s["values"], color=COLORS[k % len(COLORS)], label=s["name"])
    else:
        n = max(len(series), 1)
        w = 0.8 / n
        for k, s in enumerate(series):
            xs = [i + (k - (n - 1) / 2) * w for i in range(len(labels))]
            if kind == "line":
                ax.plot(range(len(s["values"])), s["values"], marker="o", color=COLORS[k % len(COLORS)], label=s["name"], linewidth=2)
            else:
                ax.bar(xs, s["values"], width=w, color=COLORS[k % len(COLORS)], label=s["name"])
        ax.set_xticks(range(len(labels)))
        ax.set_xticklabels(labels)
        ax.yaxis.grid(True, color="#dddddd")
        ax.set_axisbelow(True)
        for side in ("top", "right"):
            ax.spines[side].set_visible(False)
    if kind != "pie":
        if spec.get("x_label"):
            ax.set_xlabel(spec["x_label"])
        if spec.get("y_label"):
            ax.set_ylabel(spec["y_label"])
        if len(series) > 1 and kind != "hist":
            ax.legend(frameon=False)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png")
    plt.close(fig)
    return buf.getvalue()
