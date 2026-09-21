"""Análise exploratória da telemetria SolarSync -> figuras + estatísticas (JSON)
e série reamostrada em 15 min para a modelagem.

Uso: python -m src.eda
"""
from __future__ import annotations

import json
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
FIG = ROOT / "reports" / "figures"
PROC = ROOT / "data" / "processed"

AZUL, LARANJA, AQUA = "#2a78d6", "#eb6834", "#1baf7a"
GRID, EIXO, TXT = "#e1e0d9", "#c3c2b7", "#0b0b0b"
plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 9, "axes.edgecolor": EIXO,
    "axes.labelcolor": TXT, "text.color": TXT, "xtick.color": "#52514e",
    "ytick.color": "#52514e", "axes.spines.top": False, "axes.spines.right": False,
})

NUM = ["pvVoltage", "pvCurrent", "pvPower", "acInputVoltage", "acInputFrequency",
       "acOutputVoltage", "acOutputFrequency", "batteryVoltage", "batteryCurrent",
       "batteryPower", "batterySOC", "loadWatts", "loadVA", "inverterLoadPercent"]


def carregar() -> pd.DataFrame:
    arq = sorted((ROOT / "data" / "raw").glob("telemetries_*.csv"))[-1]
    df = pd.read_csv(arq, parse_dates=["timestamp"])
    # MongoDB grava datas em UTC; converte para hora local (America/Sao_Paulo, UTC-3, sem horário de verão)
    df["timestamp"] = df["timestamp"] - pd.Timedelta(hours=3)
    return df.sort_values("timestamp").reset_index(drop=True), arq.name


def energia_kwh(df: pd.DataFrame, col: str) -> pd.Series:
    dt = df["timestamp"].diff().shift(-1).dt.total_seconds().clip(upper=60).fillna(0)
    e = df[col].clip(lower=0) * dt / 3600 / 1000
    return e.groupby(df["timestamp"].dt.date).sum()


def main() -> dict:
    FIG.mkdir(parents=True, exist_ok=True)
    PROC.mkdir(parents=True, exist_ok=True)
    df, nome = carregar()
    out: dict = {"arquivo": nome, "n_registros": int(len(df))}

    # ---- qualidade ----
    dt = df["timestamp"].diff().dt.total_seconds().dropna()
    out["qualidade"] = {
        "periodo": [str(df.timestamp.min()), str(df.timestamp.max())],
        "dias_calendario": int(df.timestamp.dt.normalize().nunique()),
        "timestamps_duplicados": int(df.timestamp.duplicated().sum()),
        "nulos_total": int(df[NUM].isna().sum().sum()),
        "intervalo_mediana_s": float(dt.median()),
        "intervalo_p95_s": float(dt.quantile(.95)),
        "intervalo_max_s": float(dt.max()),
        "lacunas_gt_60s": int((dt > 60).sum()),
        "lacunas_gt_600s": int((dt > 600).sum()),
        "leituras_por_dia": {str(k): int(v) for k, v in df.groupby(df.timestamp.dt.date).size().items()},
    }
    out["modo_inversor"] = {k: int(v) for k, v in df.inverterMode.value_counts().items()}
    out["status_inversor"] = {k: int(v) for k, v in df.inverterStatus.value_counts().items()}
    out["rede_ausente_acInputVoltage_lt_100"] = int((df.acInputVoltage < 100).sum())

    # ---- descritiva ----
    desc = df[NUM].describe(percentiles=[.5]).T[["mean", "std", "min", "50%", "max"]]
    out["descritiva"] = {k: {c: float(v) for c, v in r.items()} for k, r in desc.iterrows()}

    # ---- energia diária ----
    epv = energia_kwh(df, "pvPower")
    eld = energia_kwh(df, "loadWatts")
    out["energia_diaria_kwh"] = {str(d): {"pv": round(float(epv[d]), 2), "carga": round(float(eld[d]), 2)} for d in epv.index}

    # ---- correlações ----
    cor = df[["pvPower", "pvVoltage", "pvCurrent", "loadWatts", "batteryPower", "batterySOC"]].corr()
    out["correlacao"] = {a: {b: round(float(cor.loc[a, b]), 3) for b in cor.columns} for a in cor.index}

    # ---- reamostragem 15 min ----
    d = df.set_index("timestamp")
    r = d[NUM].resample("15min").mean()
    r["n"] = d["pvPower"].resample("15min").count()
    r.loc[r["n"] < 10, NUM] = np.nan
    r.to_csv(PROC / "serie_15min.csv")
    out["serie_15min"] = {"linhas": int(len(r)), "linhas_validas": int(r.pvPower.notna().sum())}

    # ---- FIGURA 1: pvPower ao longo do período (média 15 min) ----
    fig, ax = plt.subplots(figsize=(6.3, 2.8))
    ax.plot(r.index, r["pvPower"], color=AZUL, lw=0.8)
    ax.set_ylabel("Potência FV média (W)")
    ax.set_xlabel("Data")
    ax.grid(axis="y", color=GRID, lw=.6)
    fig.autofmt_xdate(rotation=0, ha="center")
    ax.xaxis.set_major_formatter(matplotlib.dates.DateFormatter("%d/%m"))
    fig.tight_layout(); fig.savefig(FIG / "fig_pv_periodo.png", dpi=200); plt.close(fig)

    # ---- FIGURA 2: perfil horário (mediana + intervalo interquartil) ----
    h = df.assign(h=df.timestamp.dt.hour).groupby("h")["pvPower"]
    q1, med, q3 = h.quantile(.25), h.median(), h.quantile(.75)
    fig, ax = plt.subplots(figsize=(6.3, 2.8))
    ax.fill_between(q1.index, q1, q3, color=AZUL, alpha=.2, lw=0, label="Intervalo interquartil")
    ax.plot(med.index, med, color=AZUL, lw=2, label="Mediana")
    ax.plot(h.mean().index, h.mean(), color=LARANJA, lw=1.6, ls="--", label="Média")
    ax.set_xlabel("Hora local (UTC−3)"); ax.set_ylabel("Potência FV (W)")
    ax.set_xticks(range(0, 24, 2)); ax.grid(axis="y", color=GRID, lw=.6)
    ax.legend(frameon=False, fontsize=8)
    fig.tight_layout(); fig.savefig(FIG / "fig_perfil_horario.png", dpi=200); plt.close(fig)

    # ---- FIGURA 3: energia diária FV (kWh) ----
    fig, ax = plt.subplots(figsize=(6.3, 2.8))
    dias = [pd.Timestamp(x).strftime("%d/%m") for x in epv.index]
    ax.bar(dias, epv.values, color=AZUL, width=.65)
    ax.set_ylabel("Energia FV estimada (kWh)"); ax.set_xlabel("Dia")
    ax.grid(axis="y", color=GRID, lw=.6); ax.set_axisbelow(True)
    plt.setp(ax.get_xticklabels(), fontsize=7)
    fig.tight_layout(); fig.savefig(FIG / "fig_energia_diaria.png", dpi=200); plt.close(fig)

    # ---- FIGURA 4: SOC da bateria ----
    fig, ax = plt.subplots(figsize=(6.3, 2.8))
    ax.plot(r.index, r["batterySOC"], color=AQUA, lw=.9)
    ax.set_ylabel("Estado de carga (%)"); ax.set_xlabel("Data"); ax.set_ylim(0, 105)
    ax.grid(axis="y", color=GRID, lw=.6)
    ax.xaxis.set_major_formatter(matplotlib.dates.DateFormatter("%d/%m"))
    fig.tight_layout(); fig.savefig(FIG / "fig_soc.png", dpi=200); plt.close(fig)

    # ---- limitação por bateria cheia (SOC >= 99 %) ----
    dia = r[r.pvPower > 50]
    cheia = dia.batterySOC >= 99
    out["limitacao_soc"] = {
        "linhas_diurnas": int(len(dia)), "linhas_soc_ge_99": int(cheia.sum()),
        "fracao": float(cheia.mean()),
        "por_dia": {str(k): round(float(v), 3) for k, v in cheia.groupby(dia.index.date).mean().items()},
        "razao_pv_carga_mediana": float((dia.pvPower[cheia] / dia.loadWatts[cheia]).median()),
    }
    # ---- FIGURA 5: exemplo 19/09 (potência FV, carga e SOC) ----
    ex = d.loc["2026-09-19 08:00":"2026-09-19 15:00"][["pvPower", "loadWatts", "batterySOC"]].resample("2min").mean()
    fig, (a1, a2) = plt.subplots(2, 1, figsize=(6.3, 4.0), sharex=True, gridspec_kw={"height_ratios": [2, 1]})
    a1.plot(ex.index, ex.pvPower, color=AZUL, lw=1.4, label="Potência FV")
    a1.plot(ex.index, ex.loadWatts, color=LARANJA, lw=1.2, label="Potência da carga")
    a1.set_ylabel("Potência (W)"); a1.grid(axis="y", color=GRID, lw=.6); a1.legend(frameon=False, fontsize=8)
    a2.plot(ex.index, ex.batterySOC, color=AQUA, lw=1.4)
    a2.set_ylabel("SOC (%)"); a2.set_xlabel("Hora local em 19/09/2026"); a2.set_ylim(float(ex.batterySOC.min()) - 3, 102)
    a2.grid(axis="y", color=GRID, lw=.6)
    a2.xaxis.set_major_formatter(matplotlib.dates.DateFormatter("%Hh"))
    fig.tight_layout(); fig.savefig(FIG / "fig_limitacao_soc.png", dpi=200); plt.close(fig)

    (ROOT / "reports" / "eda_resultados.json").write_text(json.dumps(out, indent=2, ensure_ascii=False), encoding="utf-8")
    return out


if __name__ == "__main__":
    o = main()
    print(json.dumps({k: o[k] for k in ("n_registros", "qualidade", "modo_inversor", "energia_diaria_kwh", "correlacao", "serie_15min")}, indent=1, ensure_ascii=False)[:4000])
