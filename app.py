from __future__ import annotations

import json
import math
from datetime import date, datetime
from io import BytesIO
from typing import Any

import numpy as np
import pandas as pd
from flask import Flask, jsonify, render_template, request, send_file

from analyzer import PREVIEW_COLUMNS, analyze_supply_data

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024


def _sanitize_for_json(obj: Any) -> Any:
    """Convert NaN/NaT/inf to None so the response is valid JSON."""
    if obj is None:
        return None
    if isinstance(obj, dict):
        return {k: _sanitize_for_json(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_sanitize_for_json(v) for v in obj]
    if isinstance(obj, (np.floating, float)):
        if math.isnan(obj) or math.isinf(obj):
            return None
        return float(obj)
    if isinstance(obj, (np.integer,)):
        return int(obj)
    if isinstance(obj, (np.bool_,)):
        return bool(obj)
    if isinstance(obj, (pd.Timestamp, datetime, date)):
        return obj.isoformat()
    try:
        if pd.isna(obj):
            return None
    except (TypeError, ValueError):
        pass
    return obj


def _df_to_records(df: pd.DataFrame, limit: int | None = None) -> list[dict]:
    if limit is not None:
        df = df.head(limit).copy()
    else:
        df = df.copy()

    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].map(
                lambda x: x.isoformat() if pd.notna(x) else None
            )

    records = df.replace({np.nan: None}).to_dict(orient="records")
    return _sanitize_for_json(records)


def _preview_dataframe(df: pd.DataFrame, limit: int = 200) -> list[dict]:
    cols = [c for c in PREVIEW_COLUMNS if c in df.columns]
    return _df_to_records(df[cols], limit=limit)


@app.route("/")
def index():
    return render_template("index.html")


@app.post("/api/analyze")
def analyze():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "Empty filename"}), 400

    if not file.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify({"error": "Please upload an Excel file (.xlsx)"}), 400

    try:
        result = analyze_supply_data(file.read())
    except Exception as exc:
        return jsonify({"error": f"Failed to analyze file: {exc}"}), 400

    payload = _sanitize_for_json(
        {
            "stats": result["stats"],
            "charts": result["charts"],
            "status_distribution": result["status_distribution"],
            "assumptions": result["assumptions"],
            "validation": result["validation"],
            "summary": _df_to_records(result["summary"]),
            "sourcer_summary": _df_to_records(result["sourcer_summary"]),
            "sourcer_breakdown": _df_to_records(result["sourcer_breakdown"], limit=300),
            "designation_breakdown": _df_to_records(
                result["designation_breakdown"], limit=300
            ),
            "processed_preview": _preview_dataframe(result["processed_df"]),
        }
    )
    return jsonify(payload)


@app.post("/api/download")
def download():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    try:
        result = analyze_supply_data(file.read())
    except Exception as exc:
        return jsonify({"error": f"Failed to analyze file: {exc}"}), 400

    return send_file(
        BytesIO(result["excel_bytes"]),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="Supply_Mapping_Analyzed.xlsx",
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)
