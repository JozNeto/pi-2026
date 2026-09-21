"""Exporta a coleção `telemetries` do MongoDB Atlas para CSV local.

SOMENTE LEITURA: usa apenas `find` com projeção e `count_documents`. Nenhuma
operação de escrita é executada no banco. A connection string vem da variável
de ambiente MONGO_URI (nunca gravada em arquivo).

Uso:  MONGO_URI=... python -m src.export_telemetry
"""
from __future__ import annotations

import hashlib
import json
import os
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
from pymongo import MongoClient

RAW = Path(__file__).resolve().parents[1] / "data" / "raw"
CAMPOS = [
    "timestamp", "pvVoltage", "pvCurrent", "pvPower", "acInputVoltage",
    "acInputFrequency", "acOutputVoltage", "acOutputFrequency",
    "batteryVoltage", "batteryCurrent", "batteryPower", "batterySOC",
    "loadWatts", "loadVA", "loadCurrent", "inverterLoadPercent",
    "inverterMode", "inverterStatus",
]


def main() -> Path:
    col = MongoClient(os.environ["MONGO_URI"], serverSelectionTimeoutMS=20000) \
        .get_default_database()["telemetries"]
    proj = {c: 1 for c in CAMPOS} | {"_id": 0}
    df = pd.DataFrame(list(col.find({}, proj).sort("timestamp", 1)))
    df = df[CAMPOS]
    RAW.mkdir(parents=True, exist_ok=True)
    saida = RAW / f"telemetries_{datetime.now().strftime('%Y%m%d')}.csv"
    df.to_csv(saida, index=False)
    meta = {
        "colecao": "anenji_monitor.telemetries",
        "extraido_em_utc": datetime.now(timezone.utc).isoformat(),
        "n_registros": int(len(df)),
        "periodo": [str(df["timestamp"].min()), str(df["timestamp"].max())],
        "sha256": hashlib.sha256(saida.read_bytes()).hexdigest(),
        "operacoes": "somente leitura (find com projecao)",
    }
    saida.with_suffix(".meta.json").write_text(json.dumps(meta, indent=2), encoding="utf-8")
    return saida


if __name__ == "__main__":
    print("Gravado:", main())
