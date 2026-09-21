"""Baseline de detecção de anomalias (Isolation Forest) — V1 (relatório parcial).

Não há rótulos de falha (inverterStatus é sempre ONLINE), logo a abordagem é
não supervisionada. Treino cronológico: 06–15/09; escore em 16–20/09.
Compara com a regra atual do sistema (tensão AC de saída fora de 212–230 V).

Uso: python -m src.anomaly
"""
from __future__ import annotations

import json
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from sklearn.ensemble import IsolationForest
from sklearn.preprocessing import StandardScaler

SEED = 42
ROOT = Path(__file__).resolve().parents[1]
FIG = ROOT / "reports" / "figures"
CORTE = pd.Timestamp("2026-09-16")
CONTAMINACAO = 0.01
FEATS = ["acOutputVoltage", "acOutputFrequency", "batteryVoltage", "batteryPower",
         "loadWatts", "loadVA", "inverterLoadPercent", "pvPower", "pvVoltage", "batterySOC"]
V_MIN, V_MAX = 212.0, 230.0
AZUL, LARANJA, EIXO, GRID = "#2a78d6", "#eb6834", "#c3c2b7", "#e1e0d9"
plt.rcParams.update({"font.size": 9, "axes.edgecolor": EIXO, "axes.spines.top": False,
                     "axes.spines.right": False, "xtick.color": "#52514e", "ytick.color": "#52514e"})


def main() -> dict:
    arq = sorted((ROOT / "data" / "raw").glob("telemetries_*.csv"))[-1]
    raw = pd.read_csv(arq, parse_dates=["timestamp"])
    raw["timestamp"] = raw["timestamp"] - pd.Timedelta(hours=3)  # UTC -> hora local (UTC-3)
    raw = raw.set_index("timestamp").sort_index()
    m = raw[FEATS + ["acInputVoltage"]].resample("1min").mean().dropna()
    res: dict = {"n_minutos": int(len(m)), "contaminacao": CONTAMINACAO, "seed": SEED, "features": FEATS}

    # regra atual (PDF): tensão de saída AC fora de [212, 230] V
    regra = (m.acOutputVoltage < V_MIN) | (m.acOutputVoltage > V_MAX)
    res["regra_atual"] = {"minutos_alertados": int(regra.sum()), "fracao": float(regra.mean()),
                          "brutos_fora_faixa": int(((raw.acOutputVoltage < V_MIN) | (raw.acOutputVoltage > V_MAX)).sum())}

    tr, te = m[m.index < CORTE], m[m.index >= CORTE]
    sc = StandardScaler().fit(tr[FEATS])
    Xtr, Xte = sc.transform(tr[FEATS]), sc.transform(te[FEATS])
    iso = IsolationForest(n_estimators=300, contamination=CONTAMINACAO, random_state=SEED).fit(Xtr)
    lim = np.quantile(iso.score_samples(Xtr), CONTAMINACAO)  # limiar definido só no treino
    s_te = iso.score_samples(Xte)
    flag = pd.Series(s_te < lim, index=te.index)
    res["treino_dias"], res["teste_dias"] = "06–15/09", "16–20/09"
    res["n_treino"], res["n_teste"] = int(len(tr)), int(len(te))
    res["fracao_sinalizada_teste"] = float(flag.mean())
    res["minutos_sinalizados_teste"] = int(flag.sum())

    # sobreposição com a regra atual (teste)
    r_te = regra.loc[te.index]
    res["sobreposicao_teste"] = {
        "regra_e_modelo": int((flag & r_te).sum()), "so_regra": int((~flag & r_te).sum()),
        "so_modelo": int((flag & ~r_te).sum()), "regra_total": int(r_te.sum()),
    }

    # estabilidade entre sementes (Jaccard do conjunto sinalizado)
    conj = []
    for sd in range(5):
        i2 = IsolationForest(n_estimators=300, contamination=CONTAMINACAO, random_state=sd).fit(Xtr)
        l2 = np.quantile(i2.score_samples(Xtr), CONTAMINACAO)
        conj.append(set(te.index[i2.score_samples(Xte) < l2]))
    js = [len(a & b) / len(a | b) for i, a in enumerate(conj) for b in conj[i + 1:] if len(a | b)]
    res["jaccard_medio_entre_sementes"] = float(np.mean(js))

    # o que caracteriza os sinalizados: desvio (z) médio por variável vs. resto
    z = pd.DataFrame(Xte, index=te.index, columns=FEATS)
    res["perfil_sinalizados_z_medio"] = {f: float(z.loc[flag, f].mean()) for f in FEATS}
    # episódios (agrupa minutos consecutivos com folga de 5 min)
    ts = flag[flag].index
    ep = []
    if len(ts):
        ini = prev = ts[0]
        for t in ts[1:]:
            if (t - prev) > pd.Timedelta(minutes=5):
                ep.append((ini, prev)); ini = t
            prev = t
        ep.append((ini, prev))
    res["episodios_teste"] = len(ep)
    res["episodios_lista"] = [{"inicio": str(a), "fim": str(b), "minutos": int((b - a).total_seconds() / 60) + 1,
                               "loadWatts_max": float(te.loc[a:b, "loadWatts"].max()),
                               "batteryVoltage_min": float(te.loc[a:b, "batteryVoltage"].min()),
                               "acOutputVoltage_min": float(te.loc[a:b, "acOutputVoltage"].min()),
                               "acOutputVoltage_max": float(te.loc[a:b, "acOutputVoltage"].max())} for a, b in ep]

    # eventos de ausência de rede (não são falha, mas mudam o regime de operação)
    sem = (m.acInputVoltage < 100)
    grp = (sem != sem.shift()).cumsum()
    dur = sem.groupby(grp).agg(["first", "size"])
    quedas = dur[dur["first"]]["size"]
    res["ausencia_rede"] = {"episodios": int(len(quedas)), "minutos_total": int(quedas.sum()),
                            "duracao_mediana_min": float(quedas.median()) if len(quedas) else 0.0,
                            "duracao_max_min": int(quedas.max()) if len(quedas) else 0}

    # ---- figura ----
    fig, axs = plt.subplots(2, 1, figsize=(6.3, 4.4), sharex=True)
    a = axs[0]
    a.plot(te.index, te.acOutputVoltage, color=AZUL, lw=.5)
    a.axhline(V_MIN, color=LARANJA, ls="--", lw=1); a.axhline(V_MAX, color=LARANJA, ls="--", lw=1, label="Limites da regra atual (212 e 230 V)")
    a.scatter(flag[flag].index, te.loc[flag, "acOutputVoltage"], s=8, color="#0b0b0b", zorder=3, label="Sinalizado pelo modelo")
    a.set_ylabel("Tensão AC de saída (V)"); a.legend(frameon=False, fontsize=7, loc="lower left")
    a.grid(axis="y", color=GRID, lw=.6)
    b = axs[1]
    b.plot(te.index, te.loadWatts, color=AZUL, lw=.5)
    b.scatter(flag[flag].index, te.loc[flag, "loadWatts"], s=8, color="#0b0b0b", zorder=3)
    b.set_ylabel("Potência da carga (W)"); b.set_xlabel("Data")
    b.xaxis.set_major_formatter(matplotlib.dates.DateFormatter("%d/%m"))
    b.grid(axis="y", color=GRID, lw=.6)
    fig.tight_layout(); fig.savefig(FIG / "fig_anomalias.png", dpi=200); plt.close(fig)

    (ROOT / "reports" / "anomalia_resultados.json").write_text(json.dumps(res, indent=2, ensure_ascii=False), encoding="utf-8")
    return res


if __name__ == "__main__":
    r = main()
    print(json.dumps({k: v for k, v in r.items() if k != "episodios_lista"}, indent=1, ensure_ascii=False))
    for e in r["episodios_lista"]:
        print(e)
