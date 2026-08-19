#!/usr/bin/env python3
"""
Reproduce every numerical result reported in Chapter 5 of the GYMMY final report.

Inputs (expected in --data-dir):
1. ניסוי אמת, מעקב נבדקים(1).xlsx
2. שאלון ניסוי -  GYMMY, 2026 - ניסוי אמת _ לפני ניסוי (תגובות)(1).xlsx
3. שאלון ניסוי -  GYMMY, 2026 - ניסוי אמת _ אחרי ניסוי (תגובות)(1).xlsx
4. שאלון סיום ניסוי GYMMY _ לניסוי אמת (תגובות)(1).xlsx
5. failure_recognition_coding_final.csv

Outputs:
- Clean analysis datasets
- Every appendix table as CSV
- A machine-readable report_values.json
- Figures 6-13
- Appendix_E_F_G_Final.docx
- Report_to_Appendix_Cross_References_Final.docx

The script deliberately avoids manual entry of reported statistics. All report values are
computed from the four raw workbooks, except the binary failure-recognition classification,
which is a documented manual coding decision stored in failure_recognition_coding_final.csv.
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import shutil
from collections import Counter, defaultdict
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Iterable

import matplotlib.pyplot as plt
import numpy as np
import statsmodels.api as sm
from artifact_tool import Blob, SpreadsheetFile
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from scipy import stats

RAW_TRACKING = "ניסוי אמת, מעקב נבדקים(1).xlsx"
RAW_PRE = "שאלון ניסוי -  GYMMY, 2026 - ניסוי אמת _ לפני ניסוי (תגובות)(1).xlsx"
RAW_POST = "שאלון ניסוי -  GYMMY, 2026 - ניסוי אמת _ אחרי ניסוי (תגובות)(1).xlsx"
RAW_FINAL = "שאלון סיום ניסוי GYMMY _ לניסוי אמת (תגובות)(1).xlsx"
RECOGNITION_CODING = "failure_recognition_coding_final.csv"
LEGACY_DETECTION_CODING = "failure_detection_coding_final.csv"

EXCLUDED_IDS = {14, 32, 33, 36}
VALID_IDS = [participant_id for participant_id in range(1, 45) if participant_id not in EXCLUDED_IDS]
ALPHA = 0.05


def mean(values: Iterable[float]) -> float:
    values = list(values)
    return sum(values) / len(values)


def sample_sd(values: Iterable[float]) -> float:
    values = np.asarray(list(values), dtype=float)
    return float(values.std(ddof=1))



def fmt2(value: float) -> str:
    """Format numeric report values to two decimals using conventional half-up rounding."""
    return format(Decimal(str(float(value))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP), "f")


def format_p(value: float) -> str:
    return "< 0.001" if value < 0.001 else f"{value:.3f}"


def read_excel_rows(
    path: Path,
    sheet_name: str,
    start_row: int,
    end_row: int,
    end_col: str,
    chunk_size: int = 15,
) -> list[list[Any]]:
    """Read workbook values in small chunks using artifact_tool.

    Chunking avoids output truncation on large ranges. The returned rows preserve the
    worksheet row order, including blank cells inside populated rows.
    """
    workbook = SpreadsheetFile.import_xlsx(Blob.load(str(path)))
    rows: list[list[Any]] = []
    for first_row in range(start_row, end_row + 1, chunk_size):
        last_row = min(end_row, first_row + chunk_size - 1)
        address = f"'{sheet_name}'!A{first_row}:{end_col}{last_row}"
        result = workbook.inspect(
            {
                "kind": "table",
                "range": address,
                "include": "values",
                "table_max_rows": chunk_size + 5,
                "table_max_cols": 50,
            }
        )
        payload = json.loads(result.ndjson)
        chunk_rows = payload.get("values", [])
        # artifact_tool returns the populated rows in the requested range. All source
        # sheets used here are contiguous, so simply appending preserves Excel row order.
        rows.extend(chunk_rows)
    return rows


def load_tracking(data_dir: Path) -> tuple[dict[int, dict[str, Any]], dict[int, int]]:
    rows = read_excel_rows(data_dir / RAW_TRACKING, "Sheet1", 1, 47, "AL")
    tracking: dict[int, dict[str, Any]] = {}
    excel_rows: dict[int, int] = {}

    for zero_based_index, row in enumerate(rows[1:], start=2):
        if not row or not isinstance(row[0], (int, float)):
            continue
        participant_id = int(row[0])
        if participant_id not in VALID_IDS:
            continue
        group = int(row[4])
        failure_week = int(str(row[5]))
        failure_type = "Hardware" if str(row[6]).strip() == "חומרה" else "Interaction"
        gender = "Female" if str(row[2]).strip() in {"נ", "נקבה"} else "Male"
        tracking[participant_id] = {
            "ID": participant_id,
            "Age": float(row[1]),
            "Gender": gender,
            "Location": row[3],
            "Group": group,
            "Failure type": failure_type,
            "Timing": "Early" if failure_week == 1 else "Late",
            "Failure week": failure_week,
            "Control week": 3 - failure_week,
            "Observation text": row[21] if len(row) > 21 else None,
            "Successful recovery": str(row[22]).strip() == "כן" if len(row) > 22 else False,
        }
        excel_rows[participant_id] = zero_based_index

    if sorted(tracking) != VALID_IDS:
        raise ValueError("Tracking workbook did not yield the expected 40 valid participants.")
    return tracking, excel_rows


def load_pre_experience(data_dir: Path) -> tuple[dict[int, dict[str, float]], dict[str, int]]:
    rows = read_excel_rows(data_dir / RAW_PRE, "תגובות לטופס 1", 1, 45, "AD")
    scores: dict[int, dict[str, float]] = {}
    activity_count = 0
    no_robot_experience_count = 0

    for row in rows[1:]:
        if len(row) < 30 or not isinstance(row[1], (int, float)):
            continue
        participant_id = int(row[1])
        if participant_id not in VALID_IDS:
            continue

        nars = [float(value) for value in row[9:21]]
        tap = [float(value) for value in row[23:30]]
        scores[participant_id] = {
            "NARS S1": mean(nars[0:6]),
            "NARS S2": mean(nars[6:9]),
            # NARS items use a 1-5 scale, therefore reverse scoring is 6 - response.
            "NARS S3": mean(6 - value for value in nars[9:12]),
            "TAP Optimism": mean(tap[0:3]),
            "TAP Proficiency": mean(tap[3:6]),
            "TAP Dependence": tap[6],
        }

        if str(row[6]).strip() == "כן":
            activity_count += 1
        if str(row[21]).strip() == "לא":
            no_robot_experience_count += 1

    if sorted(scores) != VALID_IDS:
        raise ValueError("Pre-experience workbook did not yield the expected 40 valid participants.")

    return scores, {
        "Physically active": activity_count,
        "No previous robot experience": no_robot_experience_count,
    }


def load_pre_experience_flags(data_dir: Path) -> dict[int, dict[str, bool]]:
    """Return per-participant flags needed for aggregate Table E.2a.

    These fields are used only to generate aggregate counts; they are not exported
    as participant-level public data.
    """
    rows = read_excel_rows(data_dir / RAW_PRE, "תגובות לטופס 1", 1, 45, "AD")
    flags: dict[int, dict[str, bool]] = {}
    for row in rows[1:]:
        if len(row) < 30 or not isinstance(row[1], (int, float)):
            continue
        participant_id = int(row[1])
        if participant_id not in VALID_IDS:
            continue
        flags[participant_id] = {
            "Physically active": str(row[6]).strip() == "כן",
            "No prior robot experience": str(row[21]).strip() == "לא",
        }
    if sorted(flags) != VALID_IDS:
        raise ValueError("Pre-experience workbook did not yield the expected flags for 40 valid participants.")
    return flags


def load_post_session_scores(
    data_dir: Path,
) -> tuple[dict[tuple[int, int], dict[str, float]], dict[tuple[int, int], int]]:
    rows = read_excel_rows(data_dir / RAW_POST, "תגובות לטופס 1", 1, 87, "W")
    scores: dict[tuple[int, int], dict[str, float]] = {}
    excel_rows: dict[tuple[int, int], int] = {}

    for excel_row, row in enumerate(rows[1:], start=2):
        if len(row) < 21 or not isinstance(row[1], (int, float)):
            continue
        participant_id = int(row[1])
        if participant_id not in VALID_IDS:
            continue
        week = 1 if "1" in str(row[2]) else 2
        responses = [float(value) for value in row[3:21]]

        # Q4, Q14, and Q16 are negatively worded on a 1-7 scale.
        for zero_based_item_index in (3, 13, 15):
            responses[zero_based_item_index] = 8 - responses[zero_based_item_index]

        usefulness = mean(responses[0:2])
        ease = mean(responses[2:6])
        engagement = mean(responses[6:8])
        trust = mean(responses[8:11])
        satisfaction = mean(responses[11:14])
        enjoyment = mean(responses[14:16])
        attitude = mean(responses[6:16])
        intention = mean(responses[16:18])
        overall = mean([usefulness, ease, attitude, intention])

        scores[(participant_id, week)] = {
            "Perceived usefulness": usefulness,
            "Ease of use": ease,
            "Attitude": attitude,
            "Intention to use": intention,
            "Acceptance": overall,
            "Trust": trust,
            "Satisfaction": satisfaction,
            "Engagement": engagement,
            "Enjoyment": enjoyment,
            "Comfortability": mean(responses[2:4]),
            "Understanding": mean(responses[4:6]),
            "Failure report": row[22] if len(row) > 22 else None,
        }
        excel_rows[(participant_id, week)] = excel_row

    expected_keys = {(participant_id, week) for participant_id in VALID_IDS for week in (1, 2)}
    if set(scores) != expected_keys:
        missing = sorted(expected_keys - set(scores))
        raise ValueError(f"Post-session workbook is missing valid participant/week records: {missing}")
    return scores, excel_rows


def load_retrospective(
    data_dir: Path,
    tracking: dict[int, dict[str, Any]],
) -> tuple[list[dict[str, Any]], dict[int, int]]:
    rows = read_excel_rows(data_dir / RAW_FINAL, "תגובות לטופס 1", 1, 43, "I")
    results: list[dict[str, Any]] = []
    excel_rows: dict[int, int] = {}

    question_labels = [
        "Session liked more",
        "Session perceived as more understandable",
        "Session perceived as providing more value",
        "Session participants would choose to repeat",
    ]

    def recode(participant_id: int, response: Any) -> str:
        response_text = str(response).strip()
        if "אין הבדל" in response_text:
            return "No difference"
        if "שבוע 1" in response_text:
            selected_week = 1
        elif "שבוע 2" in response_text:
            selected_week = 2
        else:
            raise ValueError(f"Unrecognized retrospective response for ID {participant_id}: {response_text}")
        return (
            "Failure session"
            if selected_week == tracking[participant_id]["Failure week"]
            else "Control session"
        )

    for excel_row, row in enumerate(rows[1:], start=2):
        if len(row) < 9 or not isinstance(row[1], (int, float)):
            continue
        participant_id = int(row[1])
        if participant_id not in VALID_IDS:
            continue
        record: dict[str, Any] = {
            "ID": participant_id,
            "Felt difference": str(row[2]).strip() == "כן",
        }
        for label, response in zip(question_labels, row[4:8]):
            record[label] = recode(participant_id, response)
        results.append(record)
        excel_rows[participant_id] = excel_row

    if sorted(record["ID"] for record in results) != VALID_IDS:
        raise ValueError("Retrospective workbook did not yield the expected 40 valid participants.")
    return results, excel_rows


def load_detection_coding(data_dir: Path) -> list[dict[str, str]]:
    """Load final manual Failure Recognition coding.

    The final filename is ``failure_recognition_coding_final.csv``. For backward
    compatibility with the private working folder, the previous
    ``failure_detection_coding_final.csv`` filename is also accepted.
    """
    candidates = [data_dir / RECOGNITION_CODING, data_dir / LEGACY_DETECTION_CODING]
    path = next((candidate for candidate in candidates if candidate.exists()), candidates[0])
    with path.open(encoding="utf-8-sig", newline="") as file:
        rows = list(csv.DictReader(file))
    for row in rows:
        if "Recognized" not in row and "Detected" in row:
            row["Recognized"] = row.pop("Detected")
    if sorted(int(row["ID"]) for row in rows) != VALID_IDS:
        raise ValueError("Failure-recognition coding file does not contain exactly the 40 valid participants.")
    return rows


def build_clean_analysis_rows(
    tracking: dict[int, dict[str, Any]],
    post_scores: dict[tuple[int, int], dict[str, float]],
) -> list[dict[str, Any]]:
    constructs = [
        "Acceptance",
        "Perceived usefulness",
        "Ease of use",
        "Attitude",
        "Intention to use",
        "Trust",
        "Satisfaction",
        "Engagement",
        "Enjoyment",
        "Comfortability",
        "Understanding",
    ]
    clean_rows: list[dict[str, Any]] = []
    for participant_id in VALID_IDS:
        metadata = tracking[participant_id]
        control_week = metadata["Control week"]
        failure_week = metadata["Failure week"]
        row: dict[str, Any] = {
            key: value
            for key, value in metadata.items()
            if key not in {"Observation text", "Successful recovery"}
        }
        for construct in constructs:
            control = post_scores[(participant_id, control_week)][construct]
            failure = post_scores[(participant_id, failure_week)][construct]
            row[f"{construct} control"] = control
            row[f"{construct} failure"] = failure
            row[f"{construct} change"] = failure - control
        clean_rows.append(row)
    return clean_rows


def values(clean_rows: list[dict[str, Any]], column: str, **filters: str) -> np.ndarray:
    selected = [
        float(row[column])
        for row in clean_rows
        if all(row[key] == value for key, value in filters.items())
    ]
    return np.asarray(selected, dtype=float)


def paired_test_row(clean_rows: list[dict[str, Any]], construct: str) -> dict[str, Any]:
    control = values(clean_rows, f"{construct} control")
    failure = values(clean_rows, f"{construct} failure")
    difference = failure - control
    t_result = stats.ttest_rel(failure, control)
    standard_error = float(stats.sem(difference))
    confidence_interval = stats.t.interval(
        0.95,
        df=len(difference) - 1,
        loc=float(difference.mean()),
        scale=standard_error,
    )
    dz = abs(float(difference.mean() / difference.std(ddof=1)))
    shapiro = stats.shapiro(difference)
    return {
        "Measure": construct,
        "N": len(difference),
        "Control M": float(control.mean()),
        "Control SD": float(control.std(ddof=1)),
        "Failure M": float(failure.mean()),
        "Failure SD": float(failure.std(ddof=1)),
        "Mean change": float(difference.mean()),
        "95% CI lower": float(confidence_interval[0]),
        "95% CI upper": float(confidence_interval[1]),
        "t": float(t_result.statistic),
        "df": int(t_result.df),
        "p": float(t_result.pvalue),
        "Cohen dz": dz,
        "Difference-score Shapiro W": float(shapiro.statistic),
        "Difference-score Shapiro p": float(shapiro.pvalue),
    }


def compute_anova(clean_rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    outcome = values(clean_rows, "Acceptance change")
    design_rows: list[list[float]] = []
    for row in clean_rows:
        failure_type = 1.0 if row["Failure type"] == "Interaction" else -1.0
        timing = 1.0 if row["Timing"] == "Late" else -1.0
        design_rows.append([1.0, failure_type, timing, failure_type * timing])
    design = np.asarray(design_rows, dtype=float)
    full_model = sm.OLS(outcome, design).fit()

    sources = [(1, "Failure type"), (2, "Failure timing"), (3, "Failure type × timing")]
    anova_rows: list[dict[str, Any]] = []
    for column_index, source in sources:
        reduced_columns = [index for index in range(4) if index != column_index]
        reduced_model = sm.OLS(outcome, design[:, reduced_columns]).fit()
        sum_squares = float(reduced_model.ssr - full_model.ssr)
        mean_square = sum_squares
        f_value = mean_square / float(full_model.mse_resid)
        p_value = float(stats.f.sf(f_value, 1, full_model.df_resid))
        partial_eta_squared = sum_squares / (sum_squares + float(full_model.ssr))
        anova_rows.append(
            {
                "Source": source,
                "SS": sum_squares,
                "df": 1,
                "MS": mean_square,
                "F": f_value,
                "p": p_value,
                "Partial eta squared": partial_eta_squared,
            }
        )

    influence = full_model.get_influence()
    cell_values = [
        values(clean_rows, "Acceptance change", **{"Failure type": failure_type, "Timing": timing})
        for failure_type in ("Hardware", "Interaction")
        for timing in ("Early", "Late")
    ]
    shapiro = stats.shapiro(full_model.resid)
    levene = stats.levene(*cell_values, center="median")
    assumptions = {
        "Residual Shapiro W": float(shapiro.statistic),
        "Residual Shapiro p": float(shapiro.pvalue),
        "Levene F": float(levene.statistic),
        "Levene df1": 3,
        "Levene df2": 36,
        "Levene p": float(levene.pvalue),
        "Maximum absolute standardized residual": float(
            np.max(np.abs(influence.resid_studentized_internal))
        ),
        "Maximum Cook's D": float(np.max(influence.cooks_distance[0])),
        "Residual SS": float(full_model.ssr),
        "Residual df": int(full_model.df_resid),
        "Residual MS": float(full_model.mse_resid),
        "Total SS": float(np.sum((outcome - outcome.mean()) ** 2)),
        "Total df": len(outcome) - 1,
    }
    return anova_rows, assumptions


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str] | None = None) -> None:
    if not rows:
        return
    fieldnames = fieldnames or list(rows[0].keys())
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def create_figures(
    output_dir: Path,
    pre_scores: dict[int, dict[str, float]],
    clean_rows: list[dict[str, Any]],
    detection_rows: list[dict[str, str]],
) -> None:
    def save(fig: plt.Figure, filename: str) -> None:
        fig.tight_layout()
        fig.savefig(output_dir / filename, dpi=300, bbox_inches="tight")
        plt.close(fig)

    # Figure 8: NARS
    nars_columns = ["NARS S1", "NARS S2", "NARS S3"]
    nars_data = [[pre_scores[participant_id][column] for participant_id in VALID_IDS] for column in nars_columns]
    fig, ax = plt.subplots(figsize=(8.8, 5.8))
    ax.boxplot(
        nars_data,
        tick_labels=["S1\nDirect interaction", "S2\nGeneral concerns", "S3\nEmotional/social\n(reverse-scored)"],
        showmeans=True,
        meanprops={"marker": "D", "markersize": 6},
    )
    for position, column_values in enumerate(nars_data, start=1):
        ax.text(position, 4.82, f"M = {np.mean(column_values):.2f}", ha="center", va="top")
    ax.set_ylim(1, 5)
    ax.set_ylabel("NARS score (1-5; higher = more negative)")
    ax.set_xlabel("NARS subscale")
    save(fig, "figure_06_nars_prior_attitudes.png")

    # Figure 9: TAP
    tap_columns = ["TAP Optimism", "TAP Proficiency", "TAP Dependence"]
    tap_data = [[pre_scores[participant_id][column] for participant_id in VALID_IDS] for column in tap_columns]
    fig, ax = plt.subplots(figsize=(8.8, 5.8))
    ax.boxplot(
        tap_data,
        tick_labels=["Optimism", "Proficiency", "Dependence\n(single item)"],
        showmeans=True,
        meanprops={"marker": "D", "markersize": 6},
    )
    for position, column_values in enumerate(tap_data, start=1):
        ax.text(position, 4.82, f"M = {np.mean(column_values):.2f}", ha="center", va="top")
    ax.set_ylim(1, 5)
    ax.set_ylabel("TAP score (1-5)")
    ax.set_xlabel("TAP component")
    save(fig, "figure_07_tap_prior_attitudes.png")

    # Figure 10: Acceptance control/failure
    control = values(clean_rows, "Acceptance control")
    failure = values(clean_rows, "Acceptance failure")
    fig, ax = plt.subplots(figsize=(8.2, 5.8))
    ax.boxplot(
        [control, failure],
        tick_labels=["Control session", "Failure session"],
        showmeans=True,
        meanprops={"marker": "D", "markersize": 6},
    )
    for position, column_values in enumerate([control, failure], start=1):
        ax.text(position, 6.85, f"M = {np.mean(column_values):.2f}", ha="center", va="top")
    ax.set_ylim(1, 7)
    ax.set_ylabel("Acceptance score (1-7)")
    ax.set_xlabel("Session condition")
    save(fig, "figure_08_acceptance_control_vs_failure.png")

    # Recognition figures
    for filename, grouping, categories, labels, xlabel in [
        (
            "figure_09_failure_recognition_by_type.png",
            "Failure type",
            ["Hardware", "Interaction"],
            ["Hardware failure", "Interaction failure"],
            "Failure type",
        ),
        (
            "figure_10_failure_recognition_by_timing.png",
            "Timing",
            ["Early", "Late"],
            ["First session\n(Early)", "Second session\n(Late)"],
            "Failure timing",
        ),
    ]:
        percentages = []
        count_labels = []
        for category in categories:
            selected = [row for row in detection_rows if row[grouping] == category]
            detected = sum(row["Recognized"] == "Yes" for row in selected)
            percentages.append(100 * detected / len(selected))
            count_labels.append(f"{detected}/{len(selected)}")
        fig, ax = plt.subplots(figsize=(7.8, 5.2))
        bars = ax.bar(labels, percentages)
        ax.set_ylim(0, 100)
        ax.set_ylabel("Participants who recognized the failure (%)")
        ax.set_xlabel(xlabel)
        for bar, percentage, count_label in zip(bars, percentages, count_labels):
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                percentage + 3,
                f"{percentage:.0f}%\n({count_label})",
                ha="center",
                va="bottom",
            )
        save(fig, filename)

    # Figure 13: type
    hardware = values(clean_rows, "Acceptance change", **{"Failure type": "Hardware"})
    interaction = values(clean_rows, "Acceptance change", **{"Failure type": "Interaction"})
    fig, ax = plt.subplots(figsize=(8.2, 5.8))
    ax.boxplot(
        [hardware, interaction],
        tick_labels=["Hardware failure", "Interaction failure"],
        showmeans=True,
        meanprops={"marker": "D", "markersize": 6},
    )
    ax.axhline(0, linewidth=1)
    for position, column_values in enumerate([hardware, interaction], start=1):
        ax.text(position, 1.08, f"M = {np.mean(column_values):.2f}", ha="center", va="top")
    ax.set_ylim(-2.5, 1.2)
    ax.set_ylabel("Change in Acceptance\n(failure session - control session)")
    ax.set_xlabel("Failure type")
    save(fig, "figure_11_effect_of_failure_type.png")

    # Figure 14: timing
    early = values(clean_rows, "Acceptance change", Timing="Early")
    late = values(clean_rows, "Acceptance change", Timing="Late")
    fig, ax = plt.subplots(figsize=(8.2, 5.8))
    ax.boxplot(
        [early, late],
        tick_labels=["First session\n(Early)", "Second session\n(Late)"],
        showmeans=True,
        meanprops={"marker": "D", "markersize": 6},
    )
    ax.axhline(0, linewidth=1)
    for position, column_values in enumerate([early, late], start=1):
        ax.text(position, 1.08, f"M = {np.mean(column_values):.2f}", ha="center", va="top")
    ax.set_ylim(-2.5, 1.2)
    ax.set_ylabel("Change in Acceptance\n(failure session - control session)")
    ax.set_xlabel("Failure timing")
    save(fig, "figure_12_effect_of_failure_timing.png")

    # Figure 15: interaction
    timings = ["Early", "Late"]
    failure_types = ["Hardware", "Interaction"]
    fig, ax = plt.subplots(figsize=(8.2, 5.8))
    for failure_type in failure_types:
        group_means = []
        standard_errors = []
        for timing in timings:
            group_values = values(
                clean_rows,
                "Acceptance change",
                **{"Failure type": failure_type, "Timing": timing},
            )
            group_means.append(float(group_values.mean()))
            standard_errors.append(float(group_values.std(ddof=1) / math.sqrt(len(group_values))))
        ax.errorbar(
            timings,
            group_means,
            yerr=standard_errors,
            marker="o",
            capsize=5,
            linewidth=1.5,
            label=f"{failure_type} failure",
        )
        for index, group_mean in enumerate(group_means):
            vertical_offset = 0.06 if group_mean > -0.4 else -0.08
            ax.text(index - 0.03, group_mean + vertical_offset, f"{group_mean:.2f}", ha="center")
    ax.axhline(0, linewidth=1)
    ax.set_ylim(-1.0, 0.45)
    ax.set_ylabel("Mean change in Acceptance\n(failure session - control session)")
    ax.set_xlabel("Failure timing")
    ax.legend(title="Failure type")
    save(fig, "figure_13_type_timing_interaction.png")


def set_cell_shading(cell, fill: str = "D9E2F3") -> None:
    properties = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill)
    properties.append(shading)


def set_cell_text(cell, text: Any, bold: bool = False, center: bool = False, size: float = 8.5) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(str(text))
    run.bold = bold
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def add_heading(document: Document, text: str, level: int) -> None:
    paragraph = document.add_heading(text, level=level)
    for run in paragraph.runs:
        run.font.name = "Times New Roman"


def add_table(
    document: Document,
    caption: str,
    headers: list[str],
    rows: list[list[Any]],
    font_size: float = 8.3,
) -> None:
    caption_paragraph = document.add_paragraph()
    caption_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption_paragraph.paragraph_format.space_before = Pt(6)
    caption_paragraph.paragraph_format.space_after = Pt(4)
    caption_run = caption_paragraph.add_run(caption)
    caption_run.bold = True
    caption_run.font.name = "Times New Roman"
    caption_run.font.size = Pt(9.5)

    table = document.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.autofit = True
    for cell, header in zip(table.rows[0].cells, headers):
        set_cell_text(cell, header, bold=True, center=True, size=font_size)
        set_cell_shading(cell)
    header_properties = table.rows[0]._tr.get_or_add_trPr()
    repeat_header = OxmlElement("w:tblHeader")
    repeat_header.set(qn("w:val"), "true")
    header_properties.append(repeat_header)

    for row_values in rows:
        cells = table.add_row().cells
        for index, (cell, value) in enumerate(zip(cells, row_values)):
            set_cell_text(cell, value, center=index > 0, size=font_size)


def create_appendix_docx(
    output_dir: Path,
    report_values: dict[str, Any],
    table_e2: list[dict[str, Any]],
    paired_rows: list[dict[str, Any]],
    detection_rows: list[dict[str, str]],
    detection_tests: list[dict[str, Any]],
    group_detection_rows: list[dict[str, Any]],
    descriptive_rows: list[dict[str, Any]],
    anova_rows: list[dict[str, Any]],
    assumptions: dict[str, Any],
    retrospective_rows: list[dict[str, Any]],
    tracking_excel_rows: dict[int, int],
    post_excel_rows: dict[tuple[int, int], int],
) -> Path:
    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)
    document.styles["Normal"].font.name = "Times New Roman"
    document.styles["Normal"].font.size = Pt(10)

    add_heading(document, "Appendix E - Statistical Analysis and Reproducibility", 1)
    paragraph = document.add_paragraph(
        "All values in Chapter 5 were reproduced from the four raw Excel workbooks using "
        "the accompanying Python script reproduce_report_results.py. The only manually coded "
        "variable was Failure Recognition, for which the participant-tracking log and the "
        "post-session questionnaires were reviewed together."
    )
    paragraph.paragraph_format.space_after = Pt(6)

    add_heading(document, "E.1 Data Sources, Sample Flow, and Scoring Rules", 2)
    add_table(
        document,
        "Table E.1. Data sources, sample flow, and scoring definitions",
        ["Item", "Value / rule"],
        [
            ["Participants recruited", "44"],
            ["Excluded before analysis", "4 (IDs 14, 32, 33, and 36)"],
            ["Final sample", "40 participants; 10 in each experimental cell"],
            ["Post-session observations", "80 questionnaires: two sessions per participant"],
            ["Reverse-scored post-session items", "Q4, Q14, and Q16, using 8 - response"],
            ["Acceptance", "(Usefulness + Ease of Use + Attitude + Intention to Use) / 4"],
            ["Change score", "Failure-session score - Control-session score"],
            ["Significance level", "Two-sided alpha = .05"],
        ],
        font_size=9,
    )

    add_heading(document, "E.2 Sample Characteristics and Prior Attitudes", 2)
    add_table(
        document,
        "Table E.2. Sample characteristics and pre-experience questionnaire statistics",
        ["Measure", "N / range", "Mean", "SD", "Percentage"],
        [
            ["Age", "70-96", f"{report_values['sample']['age_mean']:.2f}", f"{report_values['sample']['age_sd']:.2f}", ""],
            ["Age 75 or older", "33", "", "", "82.5%"],
            ["Women", "26", "", "", "65.0%"],
            ["Men", "14", "", "", "35.0%"],
            ["Physically active", "38", "", "", "95.0%"],
            ["No previous robot experience", "33", "", "", "82.5%"],
        ]
        + [
            [
                row["Measure"],
                "40",
                f"{row['Mean']:.2f}",
                f"{row['SD']:.2f}",
                "",
            ]
            for row in table_e2
        ],
    )

    add_heading(document, "E.3 Failure Session Versus Control Session", 2)
    add_table(
        document,
        "Table E.3. Paired-samples t-tests for control and failure sessions",
        ["Measure", "Control M (SD)", "Failure M (SD)", "Mean change", "95% CI", "t(df)", "p", "Cohen dz"],
        [
            [
                row["Measure"],
                f"{fmt2(row['Control M'])} ({fmt2(row['Control SD'])})",
                f"{fmt2(row['Failure M'])} ({fmt2(row['Failure SD'])})",
                fmt2(row['Mean change']),
                f"[{fmt2(row['95% CI lower'])}, {fmt2(row['95% CI upper'])}]",
                f"{fmt2(row['t'])} ({row['df']})",
                format_p(row["p"]),
                fmt2(row['Cohen dz']),
            ]
            for row in paired_rows
        ],
        font_size=7.8,
    )

    add_heading(document, "E.4 Failure Recognition and Recovery", 2)
    add_table(
        document,
        "Table E.4. Pearson chi-square tests for failure recognition",
        ["Comparison", "Recognized counts", "Not Recognized counts", "chi-square(df)", "p", "Phi"],
        [
            [
                row["Comparison"],
                row["Recognized counts"],
                row["Not Recognized counts"],
                f"{row['Chi-square']:.2f} ({row['df']})",
                format_p(row["p"]),
                f"{row['Phi']:.2f}",
            ]
            for row in detection_tests
        ],
    )
    document.add_page_break()
    add_table(
        document,
        "Table E.5. Failure recognition by experimental group and successful recovery",
        ["Group", "Failure type", "Timing", "Recognized", "Not Recognized", "Recognition rate"],
        [
            [
                row["Group"],
                row["Failure type"],
                row["Timing"],
                row["Recognized"],
                row["Not Recognized"],
                f"{row['Recognition rate']:.1f}%",
            ]
            for row in group_detection_rows
        ],
    )
    document.add_paragraph(
        "Successful recovery occurred in 1 of 40 participants (2.5%) and in 1 of the 19 "
        "participants classified as recognizing the failure (5.3%). Participant-level coding "
        "and evidence locations are provided in Appendix F."
    )

    add_heading(document, "E.5 Failure Type, Failure Timing, and Their Interaction", 2)
    add_table(
        document,
        "Table E.6. Descriptive statistics for Acceptance change scores",
        ["Comparison level", "Condition", "N", "Mean change", "SD"],
        [
            [row["Comparison level"], row["Condition"], row["N"], f"{row['Mean change']:.3f}", f"{row['SD']:.3f}"]
            for row in descriptive_rows
        ],
    )
    add_table(
        document,
        "Table E.7. Two-way ANOVA on Acceptance change scores",
        ["Source", "SS", "df", "MS", "F", "p", "Partial eta squared"],
        [
            [
                row["Source"],
                f"{row['SS']:.3f}",
                row["df"],
                f"{row['MS']:.3f}",
                f"{row['F']:.3f}",
                format_p(row["p"]),
                f"{row['Partial eta squared']:.3f}",
            ]
            for row in anova_rows
        ]
        + [["Residual", f"{assumptions['Residual SS']:.3f}", assumptions["Residual df"], f"{assumptions['Residual MS']:.3f}", "", "", ""]],
    )
    add_table(
        document,
        "Table E.8. Assumption and influence diagnostics for the two-way ANOVA",
        ["Diagnostic", "Statistic", "p / threshold", "Interpretation"],
        [
            ["Residual normality (Shapiro-Wilk)", f"W = {assumptions['Residual Shapiro W']:.3f}", f"p = {assumptions['Residual Shapiro p']:.3f}", "No significant deviation from normality"],
            ["Homogeneity of variance (Levene)", f"F(3,36) = {assumptions['Levene F']:.3f}", f"p = {assumptions['Levene p']:.3f}", "Equal-variance assumption acceptable"],
            ["Largest absolute standardized residual", f"{assumptions['Maximum absolute standardized residual']:.3f}", "< 3", "No extreme standardized residual"],
            ["Largest Cook's distance", f"{assumptions['Maximum Cook\'s D']:.3f}", "< 1", "No highly influential observation"],
        ],
    )

    add_heading(document, "E.6 Retrospective Comparison", 2)
    add_table(
        document,
        "Table E.9. Retrospective comparison of the control and failure sessions",
        ["Retrospective question", "No difference, n (%)", "Control session, n (%)", "Failure session, n (%)"],
        [
            [
                row["Retrospective question"],
                f"{row['No difference n']} ({row['No difference %']:.1f}%)",
                f"{row['Control session n']} ({row['Control session %']:.1f}%)",
                f"{row['Failure session n']} ({row['Failure session %']:.1f}%)",
            ]
            for row in retrospective_rows
        ],
    )
    document.add_paragraph(
        f"Nineteen of 40 participants (47.5%) reported a perceived difference between the two sessions; "
        f"21 of 40 (52.5%) reported no difference."
    )

    document.add_page_break()
    add_heading(document, "Appendix F - Participant-Level Failure-Recognition Coding", 1)
    document.add_paragraph(
        "Recognition was coded from the participant-tracking log and the post-session questionnaire. "
        "The source locations below allow every classification to be traced back to the original files."
    )
    f_rows: list[list[Any]] = []
    for row in sorted(detection_rows, key=lambda item: int(item["ID"])):
        participant_id = int(row["ID"])
        failure_week = 1 if row["Timing"] == "Early" else 2
        source_locations = (
            f"Tracking: Sheet1 row {tracking_excel_rows[participant_id]}, col. V; "
            f"post questionnaire: row {post_excel_rows[(participant_id, failure_week)]}, col. W"
        )
        f_rows.append(
            [
                participant_id,
                row["Group"],
                row["Failure type"],
                row["Timing"],
                row["Recognized"],
                row["Evidence source"] or "Reviewed - no direct evidence",
                source_locations,
                row["Coding evidence"],
            ]
        )
    add_table(
        document,
        "Table F.1. Final participant-level failure-recognition coding",
        ["ID", "Group", "Type", "Timing", "Recognized", "Evidence source", "Original-file location", "Coding rationale"],
        f_rows,
        font_size=6.6,
    )

    document.add_page_break()
    add_heading(document, "Appendix G - Reproducible Code and Output Map", 1)
    document.add_paragraph(
        "The complete script is supplied as reproduce_report_results.py. Running it against the four "
        "raw workbooks and the final recognition-coding CSV regenerates all statistics, appendix tables, "
        "and Figures 6-13. The project repository is available at "
        "https://github.com/netanelr6/GymmyNetanelAndTal2026."
    )
    add_table(
        document,
        "Table G.1. Report section, code calculation, and supporting appendix output",
        ["Report section", "Code function / calculation", "Supporting appendix table"],
        [
            ["5.1 Data preparation", "load_tracking; load_post_session_scores; build_clean_analysis_rows", "E.1"],
            ["5.2 Sample and prior attitudes", "load_pre_experience; descriptive mean and sample SD", "E.2a–E.2b"],
            ["5.3 Overall effect of failure", "paired_test_row", "E.3"],
            ["5.4 Recognition and recovery", "Pearson chi-square contingency tables; manual recognition coding input", "E.4-E.5 and F.1"],
            ["5.5 Failure type", "Descriptive change scores and two-way ANOVA main effect", "E.6-E.8"],
            ["5.6 Failure timing", "Descriptive change scores and two-way ANOVA main effect", "E.6-E.8"],
            ["5.7 Type × timing", "Two-way ANOVA interaction term", "E.6-E.8"],
            ["5.8 Retrospective comparison", "Session-order recoding and frequency counts", "E.9"],
        ],
    )

    output_path = output_dir / "Appendix_E_F_G_Final.docx"
    document.save(output_path)
    return output_path


def create_cross_reference_docx(output_dir: Path) -> Path:
    document = Document()
    document.styles["Normal"].font.name = "Times New Roman"
    document.styles["Normal"].font.size = Pt(11)
    add_heading(document, "Exact Cross-References to Insert in the Main Report", 1)
    entries = [
        (
            "Section 3.7 - final sentence",
            "The complete statistical output and the mapping between reported results and reproducible code are provided in Appendices E-G.",
        ),
        (
            "Section 5.1 - after the scoring and analysis description",
            "The complete data-source mapping, sample flow, and scoring definitions are provided in Appendix E, Table E.1.",
        ),
        (
            "Section 5.2 - after Table 8",
            "Detailed sample and pre-experience descriptive statistics are provided in Appendix E, Table E.2.",
        ),
        (
            "Section 5.3 - after the paired-samples t-test result",
            "The complete paired-test output, including confidence intervals and effect sizes, is provided in Appendix E, Table E.3.",
        ),
        (
            "Section 5.4 - after the recognition statistics and recovery paragraph",
            "The complete contingency tables and group-level recognition results are provided in Appendix E, Tables E.4-E.5; participant-level coding and evidence locations are provided in Appendix F, Table F.1.",
        ),
        (
            "Section 5.5 - after the failure-type ANOVA result",
            "Descriptive change-score statistics, the complete two-way ANOVA table, and model diagnostics are provided in Appendix E, Tables E.6-E.8.",
        ),
        (
            "Section 5.6 - after the failure-timing ANOVA result",
            "Descriptive change-score statistics, the complete two-way ANOVA table, and model diagnostics are provided in Appendix E, Tables E.6-E.8.",
        ),
        (
            "Section 5.7 - after the interaction result",
            "The four cell means and the complete interaction-test output are provided in Appendix E, Tables E.6-E.8.",
        ),
        (
            "Section 5.8 - after Table 12",
            "The complete recoded frequency output is provided in Appendix E, Table E.9.",
        ),
        (
            "Code availability sentence",
            "All analyses and figures were reproduced using the accompanying Python script reproduce_report_results.py; the report-to-code mapping is provided in Appendix G, Table G.1.",
        ),
    ]
    for heading, text in entries:
        paragraph = document.add_paragraph()
        run = paragraph.add_run(heading)
        run.bold = True
        paragraph.add_run("\n" + text)
        paragraph.paragraph_format.space_after = Pt(8)
    output_path = output_dir / "Report_to_Appendix_Cross_References_Final.docx"
    document.save(output_path)
    return output_path


def create_missing_values_docx(output_dir: Path, paired_rows: list[dict[str, Any]], descriptive_rows: list[dict[str, Any]], anova_rows: list[dict[str, Any]], detection_tests: list[dict[str, Any]]) -> Path:
    paired = {row["Measure"]: row for row in paired_rows}
    desc = {(row["Comparison level"], row["Condition"]): row for row in descriptive_rows}
    anova = {row["Source"]: row for row in anova_rows}
    detection = {row["Comparison"]: row for row in detection_tests}

    document = Document()
    document.styles["Normal"].font.name = "Times New Roman"
    document.styles["Normal"].font.size = Pt(11)
    add_heading(document, "Verified Values to Place in Chapter 5", 1)
    entries = [
        (
            "5.2 NARS and TAP",
            "NARS: S1 M = 2.71, SD = 0.65; S2 M = 3.33, SD = 0.96; S3 M = 3.23, SD = 0.89. TAP: Optimism M = 3.51, SD = 0.98; Proficiency M = 3.00, SD = 1.01; Dependence M = 2.70, SD = 1.49.",
        ),
        (
            "5.3 Acceptance",
            f"Control M = {paired['Acceptance']['Control M']:.2f}, SD = {paired['Acceptance']['Control SD']:.2f}; failure M = {paired['Acceptance']['Failure M']:.2f}, SD = {paired['Acceptance']['Failure SD']:.2f}; mean change = {paired['Acceptance']['Mean change']:.2f}; t(39) = {paired['Acceptance']['t']:.2f}, p = {paired['Acceptance']['p']:.3f}.",
        ),
        (
            "5.4 Recognition",
            f"Failure type: chi-square(1) = {detection['Failure type']['Chi-square']:.2f}, p = {detection['Failure type']['p']:.3f}, phi = {detection['Failure type']['Phi']:.2f}. Failure timing: chi-square(1) = {detection['Failure timing']['Chi-square']:.2f}, p = {detection['Failure timing']['p']:.3f}. Recovery: 1/40 = 2.5%; 1/19 recognized participants = 5.3%.",
        ),
        (
            "5.5 Failure Type",
            f"Hardware change M = {desc[('Failure type','Hardware')]['Mean change']:.2f}, SD = {desc[('Failure type','Hardware')]['SD']:.2f}; Interaction change M = {desc[('Failure type','Interaction')]['Mean change']:.2f}, SD = {desc[('Failure type','Interaction')]['SD']:.2f}; F(1,36) = {anova['Failure type']['F']:.2f}, p = {anova['Failure type']['p']:.3f}, partial eta squared = {anova['Failure type']['Partial eta squared']:.3f}.",
        ),
        (
            "5.6 Failure Timing",
            f"Early change M = {desc[('Failure timing','Early')]['Mean change']:.2f}, SD = {desc[('Failure timing','Early')]['SD']:.2f}; Late change M = {desc[('Failure timing','Late')]['Mean change']:.2f}, SD = {desc[('Failure timing','Late')]['SD']:.2f}; F(1,36) = {anova['Failure timing']['F']:.2f}, p = {anova['Failure timing']['p']:.3f}, partial eta squared = {anova['Failure timing']['Partial eta squared']:.3f}.",
        ),
        (
            "5.7 Interaction",
            f"Hardware/Early M = {desc[('Experimental cell','Hardware / Early')]['Mean change']:.2f}, SD = {desc[('Experimental cell','Hardware / Early')]['SD']:.2f}; Hardware/Late M = {desc[('Experimental cell','Hardware / Late')]['Mean change']:.2f}, SD = {desc[('Experimental cell','Hardware / Late')]['SD']:.2f}; Interaction/Early M = {desc[('Experimental cell','Interaction / Early')]['Mean change']:.2f}, SD = {desc[('Experimental cell','Interaction / Early')]['SD']:.2f}; Interaction/Late M = {desc[('Experimental cell','Interaction / Late')]['Mean change']:.2f}, SD = {desc[('Experimental cell','Interaction / Late')]['SD']:.2f}; F(1,36) = {anova['Failure type × timing']['F']:.2f}, p = {anova['Failure type × timing']['p']:.3f}, partial eta squared = {anova['Failure type × timing']['Partial eta squared']:.3f}.",
        ),
    ]
    for heading, text in entries:
        paragraph = document.add_paragraph()
        run = paragraph.add_run(heading)
        run.bold = True
        paragraph.add_run("\n" + text)
        paragraph.paragraph_format.space_after = Pt(8)
    output_path = output_dir / "Verified_Chapter_5_Values.docx"
    document.save(output_path)
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--data-dir", type=Path, default=Path("."))
    parser.add_argument("--output-dir", type=Path, default=Path("reproduced_report_results"))
    args = parser.parse_args()

    data_dir = args.data_dir.resolve()
    output_dir = args.output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    tracking, tracking_excel_rows = load_tracking(data_dir)
    pre_scores, pre_counts = load_pre_experience(data_dir)
    pre_flags = load_pre_experience_flags(data_dir)
    post_scores, post_excel_rows = load_post_session_scores(data_dir)
    retrospective, _ = load_retrospective(data_dir, tracking)
    detection_rows = load_detection_coding(data_dir)
    clean_rows = build_clean_analysis_rows(tracking, post_scores)

    # Export clean datasets.
    write_csv(output_dir / "clean_analysis_data.csv", clean_rows)
    pre_rows = [{"ID": participant_id, **pre_scores[participant_id]} for participant_id in VALID_IDS]
    write_csv(output_dir / "pre_questionnaire_scores.csv", pre_rows)
    recognition_source = next((candidate for candidate in [data_dir / RECOGNITION_CODING, data_dir / LEGACY_DETECTION_CODING] if candidate.exists()), data_dir / RECOGNITION_CODING)
    shutil.copy2(recognition_source, output_dir / RECOGNITION_CODING)

    ages = np.asarray([tracking[participant_id]["Age"] for participant_id in VALID_IDS], dtype=float)
    gender_counts = Counter(tracking[participant_id]["Gender"] for participant_id in VALID_IDS)

    pre_statistics: list[dict[str, Any]] = []
    for measure in ["NARS S1", "NARS S2", "NARS S3", "TAP Optimism", "TAP Proficiency", "TAP Dependence"]:
        measure_values = [pre_scores[participant_id][measure] for participant_id in VALID_IDS]
        pre_statistics.append(
            {
                "Measure": measure,
                "Mean": float(np.mean(measure_values)),
                "SD": float(np.std(measure_values, ddof=1)),
                "N": len(measure_values),
            }
        )

    paired_constructs = [
        "Acceptance",
        "Perceived usefulness",
        "Ease of use",
        "Attitude",
        "Intention to use",
    ]
    paired_rows = [paired_test_row(clean_rows, construct) for construct in paired_constructs]

    # Recognition tests, without continuity correction, matching the report.
    detection_tests: list[dict[str, Any]] = []
    for comparison, categories in [
        ("Failure type", ["Hardware", "Interaction"]),
        ("Failure timing", ["Early", "Late"]),
    ]:
        field = "Failure type" if comparison == "Failure type" else "Timing"
        contingency = []
        detected_counts = []
        not_detected_counts = []
        for category in categories:
            selected = [row for row in detection_rows if row[field] == category]
            detected = sum(row["Recognized"] == "Yes" for row in selected)
            not_detected = len(selected) - detected
            contingency.append([detected, not_detected])
            detected_counts.append(f"{category}: {detected}/{len(selected)}")
            not_detected_counts.append(f"{category}: {not_detected}/{len(selected)}")
        chi_square, p_value, df, _ = stats.chi2_contingency(np.asarray(contingency), correction=False)
        detection_tests.append(
            {
                "Comparison": comparison,
                "Recognized counts": "; ".join(detected_counts),
                "Not Recognized counts": "; ".join(not_detected_counts),
                "Chi-square": float(chi_square),
                "df": int(df),
                "p": float(p_value),
                "Phi": float(math.sqrt(chi_square / len(detection_rows))),
            }
        )

    group_detection_rows: list[dict[str, Any]] = []
    for group in ["G1", "G2", "G3", "G4"]:
        selected = [row for row in detection_rows if row["Group"] == group]
        detected = sum(row["Recognized"] == "Yes" for row in selected)
        group_detection_rows.append(
            {
                "Group": group,
                "Failure type": selected[0]["Failure type"],
                "Timing": selected[0]["Timing"],
                "Recognized": detected,
                "Not Recognized": len(selected) - detected,
                "Recognition rate": 100 * detected / len(selected),
            }
        )
    total_detected = sum(row["Recognized"] == "Yes" for row in detection_rows)
    group_detection_rows.append(
        {
            "Group": "Total",
            "Failure type": "",
            "Timing": "",
            "Recognized": total_detected,
            "Not Recognized": len(detection_rows) - total_detected,
            "Recognition rate": 100 * total_detected / len(detection_rows),
        }
    )

    descriptive_rows: list[dict[str, Any]] = []
    for comparison_level, field, conditions in [
        ("Failure type", "Failure type", ["Hardware", "Interaction"]),
        ("Failure timing", "Timing", ["Early", "Late"]),
    ]:
        for condition in conditions:
            group_values = values(clean_rows, "Acceptance change", **{field: condition})
            descriptive_rows.append(
                {
                    "Comparison level": comparison_level,
                    "Condition": condition,
                    "N": len(group_values),
                    "Mean change": float(group_values.mean()),
                    "SD": float(group_values.std(ddof=1)),
                }
            )
    for failure_type in ["Hardware", "Interaction"]:
        for timing in ["Early", "Late"]:
            group_values = values(
                clean_rows,
                "Acceptance change",
                **{"Failure type": failure_type, "Timing": timing},
            )
            descriptive_rows.append(
                {
                    "Comparison level": "Experimental cell",
                    "Condition": f"{failure_type} / {timing}",
                    "N": len(group_values),
                    "Mean change": float(group_values.mean()),
                    "SD": float(group_values.std(ddof=1)),
                }
            )

    anova_rows, assumptions = compute_anova(clean_rows)

    retrospective_rows: list[dict[str, Any]] = []
    retrospective_labels = [
        "Session liked more",
        "Session perceived as more understandable",
        "Session perceived as providing more value",
        "Session participants would choose to repeat",
    ]
    for label in retrospective_labels:
        counts = Counter(row[label] for row in retrospective)
        retrospective_rows.append(
            {
                "Retrospective question": label,
                "No difference n": counts["No difference"],
                "No difference %": 100 * counts["No difference"] / len(retrospective),
                "Control session n": counts["Control session"],
                "Control session %": 100 * counts["Control session"] / len(retrospective),
                "Failure session n": counts["Failure session"],
                "Failure session %": 100 * counts["Failure session"] / len(retrospective),
            }
        )

    # Output appendix CSV tables.
    table_e1 = [
        {"Item": "Participants recruited", "Value": 44},
        {"Item": "Participants excluded", "Value": 4},
        {"Item": "Final sample", "Value": 40},
        {"Item": "Participants per experimental cell", "Value": 10},
        {"Item": "Post-session questionnaires", "Value": 80},
        {"Item": "Alpha", "Value": ALPHA},
    ]
    write_csv(output_dir / "table_E1_sample_flow_and_scoring.csv", table_e1)

    # Final Appendix E.2a: aggregate sample characteristics by experimental group.
    e2a_rows: list[dict[str, Any]] = []
    for group_label, group_number in [("G1", 1), ("G2", 2), ("G3", 3), ("G4", 4)]:
        ids = [participant_id for participant_id in VALID_IDS if tracking[participant_id]["Group"] == group_number]
        group_ages = np.asarray([tracking[participant_id]["Age"] for participant_id in ids], dtype=float)
        women = sum(tracking[participant_id]["Gender"] == "Female" for participant_id in ids)
        men = len(ids) - women
        failure_type = tracking[ids[0]]["Failure type"]
        timing = tracking[ids[0]]["Timing"]
        e2a_rows.append({
            "Group": group_label,
            "Condition": f"{failure_type} / {timing}",
            "Age, M (SD)": f"{group_ages.mean():.2f} ({group_ages.std(ddof=1):.2f})",
            "Age range": f"{int(group_ages.min())}–{int(group_ages.max())}",
            "Women / Men": f"{women} / {men}",
            "Age ≥ 75": sum(age >= 75 for age in group_ages),
            "Physically active": sum(pre_flags[participant_id]["Physically active"] for participant_id in ids),
            "No prior robot experience": sum(pre_flags[participant_id]["No prior robot experience"] for participant_id in ids),
        })
    write_csv(output_dir / "table_E2a_sample_characteristics_by_group.csv", e2a_rows)

    # Final Appendix E.2b: overall sample and prior-attitude summary.
    pre_by_measure = {row["Measure"]: row for row in pre_statistics}
    e2b_rows = [
        {"Measure": "Age", "N / range": f"{int(ages.min())}–{int(ages.max())}", "Mean": f"{ages.mean():.2f}", "SD": f"{ages.std(ddof=1):.2f}", "Percentage": ""},
        {"Measure": "Age 75 or older", "N / range": sum(age >= 75 for age in ages), "Mean": "", "SD": "", "Percentage": f"{100*sum(age >= 75 for age in ages)/len(ages):.1f}%"},
        {"Measure": "Women", "N / range": gender_counts.get("Female",0), "Mean": "", "SD": "", "Percentage": f"{100*gender_counts.get('Female',0)/len(ages):.1f}%"},
        {"Measure": "Men", "N / range": gender_counts.get("Male",0), "Mean": "", "SD": "", "Percentage": f"{100*gender_counts.get('Male',0)/len(ages):.1f}%"},
        {"Measure": "Physically active", "N / range": pre_counts["Physically active"], "Mean": "", "SD": "", "Percentage": f"{100*pre_counts['Physically active']/len(ages):.1f}%"},
        {"Measure": "No previous robot experience", "N / range": pre_counts["No previous robot experience"], "Mean": "", "SD": "", "Percentage": f"{100*pre_counts['No previous robot experience']/len(ages):.1f}%"},
    ]
    for measure in ["NARS S1", "NARS S2", "NARS S3", "TAP Optimism", "TAP Proficiency", "TAP Dependence"]:
        row = pre_by_measure[measure]
        e2b_rows.append({"Measure": measure, "N / range": row["N"], "Mean": f"{row['Mean']:.2f}", "SD": f"{row['SD']:.2f}", "Percentage": ""})
    write_csv(output_dir / "table_E2b_sample_and_prior_attitudes.csv", e2b_rows)

    write_csv(output_dir / "table_E3_paired_tests.csv", paired_rows)
    write_csv(output_dir / "table_E4_recognition_chi_square.csv", detection_tests)
    write_csv(output_dir / "table_E5_recognition_by_group.csv", group_detection_rows)
    write_csv(output_dir / "table_E6_change_descriptives.csv", descriptive_rows)
    write_csv(output_dir / "table_E7_two_way_anova.csv", anova_rows)
    write_csv(
        output_dir / "table_E8_anova_assumptions.csv",
        [{"Diagnostic": key, "Value": value} for key, value in assumptions.items()],
    )
    write_csv(output_dir / "table_E9_retrospective_comparison.csv", retrospective_rows)

    # Add original-file row references to the participant coding table.
    coding_output: list[dict[str, Any]] = []
    for row in sorted(detection_rows, key=lambda item: int(item["ID"])):
        participant_id = int(row["ID"])
        failure_week = 1 if row["Timing"] == "Early" else 2
        coding_output.append(
            {
                **row,
                "Tracking workbook location": f"Sheet1 row {tracking_excel_rows[participant_id]}, column V",
                "Post questionnaire location": f"תגובות לטופס 1 row {post_excel_rows[(participant_id, failure_week)]}, column W",
            }
        )
    write_csv(output_dir / "table_F1_participant_recognition_coding.csv", coding_output)

    # Convenience exports matching the table numbers used in the main report.
    report_table_7 = []
    for group in ["G1", "G2", "G3", "G4"]:
        selected = [row for row in clean_rows if f"G{row['Group']}" == group]
        report_table_7.append(
            {
                "Group": group,
                "Failure type": selected[0]["Failure type"],
                "Failure timing": selected[0]["Timing"],
                "N": len(selected),
            }
        )
    write_csv(output_dir / "report_table_07_experimental_groups.csv", report_table_7)
    write_csv(output_dir / "report_table_08_nars_tap_descriptives.csv", pre_statistics)
    write_csv(output_dir / "report_table_09_paired_comparison.csv", paired_rows)
    write_csv(output_dir / "report_table_11_recognition_by_group.csv", group_detection_rows)
    write_csv(output_dir / "report_table_12_retrospective_comparison.csv", retrospective_rows)

    code_output_map_rows = [
        {"Report section": "5.1 Data preparation", "Code function / calculation": "load_tracking; load_post_session_scores; build_clean_analysis_rows", "Supporting appendix table": "E.1"},
        {"Report section": "5.2 Sample and prior attitudes", "Code function / calculation": "load_pre_experience; descriptive mean and sample SD", "Supporting appendix table": "E.2a–E.2b"},
        {"Report section": "5.3 Overall effect of failure", "Code function / calculation": "paired_test_row", "Supporting appendix table": "E.3"},
        {"Report section": "5.4 Recognition and recovery", "Code function / calculation": "Pearson chi-square contingency tables; manual recognition coding input", "Supporting appendix table": "E.4-E.5 and F.1"},
        {"Report section": "5.5 Failure type", "Code function / calculation": "Descriptive change scores and two-way ANOVA main effect", "Supporting appendix table": "E.6-E.8"},
        {"Report section": "5.6 Failure timing", "Code function / calculation": "Descriptive change scores and two-way ANOVA main effect", "Supporting appendix table": "E.6-E.8"},
        {"Report section": "5.7 Type x timing", "Code function / calculation": "Two-way ANOVA interaction term", "Supporting appendix table": "E.6-E.8"},
        {"Report section": "5.8 Retrospective comparison", "Code function / calculation": "Session-order recoding and frequency counts", "Supporting appendix table": "E.9"},
    ]
    write_csv(output_dir / "table_G1_code_output_map.csv", code_output_map_rows)

    report_values = {
        "sample": {
            "recruited": 44,
            "excluded": 4,
            "excluded_ids": sorted(EXCLUDED_IDS),
            "final_n": 40,
            "age_min": float(ages.min()),
            "age_max": float(ages.max()),
            "age_mean": float(ages.mean()),
            "age_sd": float(ages.std(ddof=1)),
            "age_75_plus": int(np.sum(ages >= 75)),
            "women": gender_counts["Female"],
            "men": gender_counts["Male"],
            "physically_active": pre_counts["Physically active"],
            "no_previous_robot_experience": pre_counts["No previous robot experience"],
        },
        "pre_statistics": pre_statistics,
        "paired_tests": paired_rows,
        "recognition_tests": detection_tests,
        "recognition_by_group": group_detection_rows,
        "successful_recovery": {
            "count": 1,
            "full_sample_percentage": 100 / 40,
            "recognized_participants_percentage": 100 / total_detected,
        },
        "change_descriptives": descriptive_rows,
        "two_way_anova": anova_rows,
        "anova_assumptions": assumptions,
        "retrospective": {
            "felt_difference": sum(row["Felt difference"] for row in retrospective),
            "felt_no_difference": sum(not row["Felt difference"] for row in retrospective),
            "comparison_table": retrospective_rows,
        },
    }
    with (output_dir / "report_values.json").open("w", encoding="utf-8") as file:
        json.dump(report_values, file, ensure_ascii=False, indent=2)

    create_figures(output_dir, pre_scores, clean_rows, detection_rows)
    appendix_path = create_appendix_docx(
        output_dir,
        report_values,
        pre_statistics,
        paired_rows,
        detection_rows,
        detection_tests,
        group_detection_rows,
        descriptive_rows,
        anova_rows,
        assumptions,
        retrospective_rows,
        tracking_excel_rows,
        post_excel_rows,
    )
    cross_reference_path = create_cross_reference_docx(output_dir)
    missing_values_path = create_missing_values_docx(
        output_dir,
        paired_rows,
        descriptive_rows,
        anova_rows,
        detection_tests,
    )

    print(f"Created: {appendix_path}")
    print(f"Created: {cross_reference_path}")
    print(f"Created: {missing_values_path}")
    print(f"All numerical outputs were written to: {output_dir}")


if __name__ == "__main__":
    main()
