"""Forecasting de curto prazo da potência FV (SolarSync) — V1 (relatório parcial).

Resolução 15 min; horizontes 15 min, 1 h e 3 h (modelos diretos). Validação
walk-forward diária (treina só com o passado, testa no dia seguinte).
Métricas apenas nos horários com luz solar (definidos a partir do TREINO).

Uso: python -m src.forecast
"""
from __future__ import annotations

import json
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from sklearn.ensemble import HistGradientBoostingRegressor
from sklearn.impute import SimpleImputer
from sklearn.linear_model import Ridge
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import StandardScaler

SEED = 42
ROOT = Path(__file__).resolve().parents[1]
FIG = ROOT / "reports" / "figures"
PASSO = 96  # passos de 15 min por dia
HORIZONTES = {"15 min": 1, "1 h": 4, "3 h": 12}
PRIMEIRO_DIA_TESTE = pd.Timestamp("2026-09-13")
LIMIAR_DIA_W = 10.0  # slot é "diurno" se média histórica de treino > 10 W

AZUL, LARANJA, AQUA, EIXO, GRID = "#2a78d6", "#eb6834", "#1baf7a", "#c3c2b7", "#e1e0d9"
plt.rcParams.update({"font.size": 9, "axes.edgecolor": EIXO, "axes.spines.top": False,
                     "axes.spines.right": False, "xtick.color": "#52514e", "ytick.color": "#52514e"})

MODELOS = ["Persistência", "Sazonal ingênuo (24 h)", "Climatologia", "Ridge", "Gradient Boosting"]


def carregar() -> pd.DataFrame:
    df = pd.read_csv(ROOT / "data" / "processed" / "serie_15min.csv", index_col=0, parse_dates=True)
    df = df.asfreq("15min")
    for c in ["pvPower", "pvVoltage", "loadWatts", "batterySOC"]:
        df[c] = df[c].interpolate(limit=2, limit_area="inside")
    return df


def montar(df: pd.DataFrame, h: int) -> pd.DataFrame:
    s = df["pvPower"]
    X = pd.DataFrame(index=df.index)
    X["pv0"] = s
    for k in (1, 2, 3, 4):
        X[f"pv_l{k}"] = s.shift(k)
    X["pv_ma4"] = s.rolling(4).mean()
    X["pv_ma12"] = s.rolling(12).mean()
    X["pv_max4"] = s.rolling(4).max()
    X["pvV0"] = df["pvVoltage"]
    X["carga0"] = df["loadWatts"]
    X["soc0"] = df["batterySOC"]
    X["pv_dia_ant"] = s.shift(PASSO - h)  # valor no mesmo horário do alvo, 24 h antes
    alvo_t = df.index + pd.Timedelta(minutes=15 * h)
    minutos = alvo_t.hour * 60 + alvo_t.minute
    X["tod_sin"] = np.sin(2 * np.pi * minutos / 1440)
    X["tod_cos"] = np.cos(2 * np.pi * minutos / 1440)
    X["y"] = s.shift(-h)
    X["alvo_t"] = alvo_t
    X["slot"] = (minutos // 15).astype(int)
    return X


def climatologia(df: pd.DataFrame, ate: pd.Timestamp) -> pd.Series:
    tr = df.loc[df.index < ate, "pvPower"].dropna()
    return tr.groupby((tr.index.hour * 60 + tr.index.minute) // 15).mean()


def metricas(y, p) -> dict:
    e = np.asarray(p) - np.asarray(y)
    return {"mae": float(np.mean(np.abs(e))), "rmse": float(np.sqrt(np.mean(e ** 2))), "n": int(len(e))}


def main() -> dict:
    FIG.mkdir(parents=True, exist_ok=True)
    df = carregar()
    dias_teste = [d for d in pd.date_range(PRIMEIRO_DIA_TESTE, df.index.max().normalize())]
    feats = None
    linhas = []
    imp_acum: dict[int, dict[str, list]] = {}

    for hn, h in HORIZONTES.items():
        X = montar(df, h)
        feats = [c for c in X.columns if c not in ("y", "alvo_t", "slot")]
        for D in dias_teste:
            D1 = D + pd.Timedelta(days=1)
            tr = X[(X["alvo_t"] < D) & X["y"].notna() & X["pv0"].notna()]
            te = X[(X["alvo_t"] >= D) & (X["alvo_t"] < D1) & X["y"].notna() & X["pv0"].notna()]
            if len(te) < 10:
                continue
            clim = climatologia(df, D)
            diurno = clim[clim > LIMIAR_DIA_W].index
            te = te[te["slot"].isin(diurno)]
            if te.empty:
                continue

            ridge = make_pipeline(SimpleImputer(strategy="median"), StandardScaler(), Ridge(alpha=1.0))
            ridge.fit(tr[feats], tr["y"])
            gb = HistGradientBoostingRegressor(max_iter=200, learning_rate=0.05, max_depth=3,
                                               min_samples_leaf=20, l2_regularization=1.0,
                                               random_state=SEED)
            gb.fit(tr[feats], tr["y"])

            preds = {
                "Persistência": te["pv0"].to_numpy(),
                "Sazonal ingênuo (24 h)": te["pv_dia_ant"].fillna(te["pv0"]).to_numpy(),
                "Climatologia": te["slot"].map(clim).to_numpy(),
                "Ridge": np.clip(ridge.predict(te[feats]), 0, None),
                "Gradient Boosting": np.clip(gb.predict(te[feats]), 0, None),
            }
            for m, p in preds.items():
                for t, yv, pv in zip(te["alvo_t"], te["y"], p):
                    linhas.append({"horizonte": hn, "h": h, "dia": D, "modelo": m, "alvo_t": t, "y": yv, "p": pv})

            if h == 4:  # importância por permutação (Gradient Boosting, horizonte 1 h)
                base = np.mean(np.abs(np.clip(gb.predict(te[feats]), 0, None) - te["y"]))
                rng = np.random.default_rng(SEED)
                for f in feats:
                    inc = []
                    for _ in range(5):
                        tp = te[feats].copy()
                        tp[f] = rng.permutation(tp[f].to_numpy())
                        inc.append(np.mean(np.abs(np.clip(gb.predict(tp), 0, None) - te["y"])) - base)
                    imp_acum.setdefault(h, {}).setdefault(f, []).append(float(np.mean(inc)))

    R = pd.DataFrame(linhas)
    R.to_csv(ROOT / "data" / "processed" / "previsoes_walkforward.csv", index=False)

    # ---------- métricas agregadas ----------
    res: dict = {"config": {"resolucao": "15 min", "dias_teste": [str(d.date()) for d in dias_teste],
                            "limiar_diurno_W": LIMIAR_DIA_W, "seed": SEED, "features": feats,
                            "n_dias_teste": len(dias_teste)}, "horizonte": {}}
    for hn in HORIZONTES:
        sub = R[R.horizonte == hn]
        res["horizonte"][hn] = {}
        pers = metricas(sub[sub.modelo == "Persistência"].y, sub[sub.modelo == "Persistência"].p)["mae"]
        for m in MODELOS:
            s = sub[sub.modelo == m]
            r = metricas(s.y, s.p)
            r["nmae"] = r["mae"] / float(s.y.mean())
            r["skill_vs_persistencia"] = 1 - r["mae"] / pers
            diario = s.groupby("dia").apply(lambda g: np.mean(np.abs(g.p - g.y)), include_groups=False)
            r["mae_diario_media"], r["mae_diario_dp"] = float(diario.mean()), float(diario.std(ddof=1))
            r["mae_por_dia"] = {str(k.date()): float(v) for k, v in diario.items()}
            res["horizonte"][hn][m] = r
        res["horizonte"][hn]["media_y_diurna_W"] = float(sub[sub.modelo == "Persistência"].y.mean())

    # vitórias diárias do GB contra persistência (1 h)
    d1 = res["horizonte"]["1 h"]
    a = d1["Gradient Boosting"]["mae_por_dia"]; b = d1["Persistência"]["mae_por_dia"]
    res["gb_vence_persistencia_1h_dias"] = int(sum(a[k] < b[k] for k in a)); res["n_dias_avaliados_1h"] = len(a)

    # importância
    imp = {f: float(np.mean(v)) for f, v in imp_acum.get(4, {}).items()}
    res["importancia_permutacao_1h"] = dict(sorted(imp.items(), key=lambda kv: -kv[1]))

    # ---------- day-ahead (sem clima): apenas baselines ----------
    S = df["pvPower"]
    da = []
    for D in dias_teste:
        D1 = D + pd.Timedelta(days=1)
        idx = S[(S.index >= D) & (S.index < D1)].index
        real = S.loc[idx]
        clim = climatologia(df, D)
        slot = (idx.hour * 60 + idx.minute) // 15
        p_clim = pd.Series(slot.map(clim).to_numpy(), index=idx)
        p_naive = S.shift(PASSO).loc[idx]
        prev3 = pd.concat([S.shift(PASSO * k) for k in (1, 2, 3)], axis=1).mean(axis=1).loc[idx]
        diurno = clim[clim > LIMIAR_DIA_W].index
        msk = pd.Series(slot.isin(diurno), index=idx) & real.notna() & p_naive.notna() & prev3.notna()
        da.append({"dia": str(D.date()),
                   "kwh_real": float(real[msk | real.notna()].sum() * .25 / 1000),
                   "kwh_naive": float(p_naive.fillna(0).sum() * .25 / 1000),
                   "kwh_clim": float(p_clim.sum() * .25 / 1000),
                   "kwh_ma3": float(prev3.fillna(0).sum() * .25 / 1000),
                   "mae_naive": float(np.mean(np.abs(p_naive[msk] - real[msk]))),
                   "mae_clim": float(np.mean(np.abs(p_clim[msk] - real[msk]))),
                   "mae_ma3": float(np.mean(np.abs(prev3[msk] - real[msk])))})
    DA = pd.DataFrame(da)
    res["day_ahead"] = {
        "tabela": da,
        "mae_medio": {k: float(DA[k].mean()) for k in ("mae_naive", "mae_clim", "mae_ma3")},
        "erro_energia_kwh_medio": {
            "naive": float((DA.kwh_naive - DA.kwh_real).abs().mean()),
            "clim": float((DA.kwh_clim - DA.kwh_real).abs().mean()),
            "ma3": float((DA.kwh_ma3 - DA.kwh_real).abs().mean()),
        },
        "energia_real_media_kwh": float(DA.kwh_real.mean()),
    }

    # ---------- figuras ----------
    # F1: MAE por modelo, 3 painéis (horizontes)
    fig, axs = plt.subplots(1, 3, figsize=(6.3, 2.7), sharex=False)
    for ax, hn in zip(axs, HORIZONTES):
        vals = [res["horizonte"][hn][m]["mae"] for m in MODELOS]
        cores = [AQUA if m == "Gradient Boosting" else AZUL for m in MODELOS]
        ax.barh(MODELOS, vals, color=cores, height=.62)
        ax.set_title(f"Horizonte {hn}", fontsize=9)
        ax.invert_yaxis(); ax.set_xlabel("MAE (W)")
        ax.grid(axis="x", color=GRID, lw=.6); ax.set_axisbelow(True)
        for i, v in enumerate(vals):
            ax.text(v, i, f" {v:.0f}", va="center", fontsize=7)
        if ax is not axs[0]:
            ax.set_yticklabels([])
        ax.set_xlim(0, max(vals) * 1.25)
    fig.tight_layout(); fig.savefig(FIG / "fig_mae_modelos.png", dpi=200); plt.close(fig)

    # F2: série real x previsões (1 h), últimos 3 dias de teste
    ult = dias_teste[-3:]
    sub = R[(R.horizonte == "1 h") & R.dia.isin(ult)]
    fig, axs = plt.subplots(1, 3, figsize=(6.3, 3.0), sharey=True)
    for ax, D in zip(axs, ult):
        s1 = sub[sub.dia == D]
        for m, cor, ls, lw, lab in (("Gradient Boosting", AZUL, "-", 1.6, "Observado"),):
            g = s1[s1.modelo == m].set_index("alvo_t")
            ax.plot(g.index.hour + g.index.minute / 60, g["y"], color=AZUL, lw=1.6, label="Observado")
        for m, cor, ls in (("Persistência", LARANJA, "--"), ("Gradient Boosting", AQUA, "-")):
            g = s1[s1.modelo == m].set_index("alvo_t")
            ax.plot(g.index.hour + g.index.minute / 60, g["p"], color=cor, lw=1.1, ls=ls, label=m)
        ax.set_title(pd.Timestamp(D).strftime("%d/%m/%Y"), fontsize=9)
        ax.set_xlabel("Hora local"); ax.grid(axis="y", color=GRID, lw=.6)
    axs[0].set_ylabel("Potência FV (W)")
    h_, l_ = axs[0].get_legend_handles_labels()
    fig.legend(h_, l_, loc="lower center", ncol=3, frameon=False, fontsize=8)
    fig.tight_layout(rect=(0, .08, 1, 1)); fig.savefig(FIG / "fig_previsao_1h.png", dpi=200); plt.close(fig)

    # F3: importância
    fig, ax = plt.subplots(figsize=(6.3, 3.0))
    items = list(res["importancia_permutacao_1h"].items())[:8][::-1]
    ax.barh([k for k, _ in items], [v for _, v in items], color=AZUL, height=.62)
    ax.set_xlabel("Aumento do MAE ao permutar a variável (W)")
    ax.grid(axis="x", color=GRID, lw=.6); ax.set_axisbelow(True)
    fig.tight_layout(); fig.savefig(FIG / "fig_importancia.png", dpi=200); plt.close(fig)

    # F4: energia diária real x baselines (day-ahead)
    fig, ax = plt.subplots(figsize=(6.3, 3.0))
    x = np.arange(len(DA)); w = .27
    ax.bar(x - w, DA.kwh_real, w, color=AZUL, label="Observado")
    ax.bar(x, DA.kwh_naive, w, color=LARANJA, label="Sazonal ingênuo (dia anterior)")
    ax.bar(x + w, DA.kwh_clim, w, color=AQUA, label="Climatologia")
    ax.set_xticks(x); ax.set_xticklabels([pd.Timestamp(d).strftime("%d/%m") for d in DA.dia])
    ax.set_ylabel("Energia FV diária (kWh)"); ax.set_xlabel("Dia previsto")
    ax.set_ylim(0, float(max(DA.kwh_real.max(), DA.kwh_naive.max(), DA.kwh_clim.max())) * 1.3)
    ax.grid(axis="y", color=GRID, lw=.6); ax.set_axisbelow(True); ax.legend(frameon=False, fontsize=8, ncol=3, loc="upper left")
    fig.tight_layout(); fig.savefig(FIG / "fig_day_ahead.png", dpi=200); plt.close(fig)

    (ROOT / "reports" / "forecast_resultados.json").write_text(json.dumps(res, indent=2, ensure_ascii=False), encoding="utf-8")
    return res


if __name__ == "__main__":
    r = main()
    for hn, d in r["horizonte"].items():
        print(f"\n== {hn} (média diurna {d['media_y_diurna_W']:.0f} W) ==")
        for m in MODELOS:
            x = d[m]
            print(f"  {m:26s} MAE {x['mae']:7.1f}  RMSE {x['rmse']:7.1f}  nMAE {x['nmae']:.2f}  skill {x['skill_vs_persistencia']:+.2f}  dp_dia {x['mae_diario_dp']:.0f}")
    print("\nGB vence persistência (1h):", r["gb_vence_persistencia_1h_dias"], "/", r["n_dias_avaliados_1h"])
    print("importância:", {k: round(v, 1) for k, v in r["importancia_permutacao_1h"].items()})
    print("day-ahead:", json.dumps({k: r["day_ahead"][k] for k in ("mae_medio", "erro_energia_kwh_medio", "energia_real_media_kwh")}, indent=1))
