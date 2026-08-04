#!/usr/bin/env python3
"""Generate Figures 6-13 from the anonymized processed GYMMY data.

This is the portable public script for the figures used in Chapter 5 of the report.
It does not require the raw participant workbooks.

Expected files in DATA_DIR:
- pre_questionnaire_scores.csv
- clean_analysis_data_public.csv
- failure_detection_coding_public.csv

Usage:
    python generate_all_report_figures.py
    python generate_all_report_figures.py --data-dir data --output-dir figures
"""

from __future__ import annotations

import argparse
import csv
import math
from pathlib import Path
from typing import Iterable

import matplotlib.pyplot as plt
import numpy as np


def read_rows(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8-sig", newline="") as file:
        return list(csv.DictReader(file))


def numeric(rows: Iterable[dict[str, str]], column: str) -> np.ndarray:
    return np.asarray([float(row[column]) for row in rows], dtype=float)


def save(fig: plt.Figure, output_dir: Path, filename: str) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    fig.tight_layout()
    fig.savefig(output_dir / filename, dpi=300, bbox_inches="tight")
    plt.close(fig)


def make_all_figures(data_dir: Path, output_dir: Path) -> None:
    pre_rows = read_rows(data_dir / "pre_questionnaire_scores.csv")
    clean_rows = read_rows(data_dir / "clean_analysis_data_public.csv")
    detection_rows = read_rows(data_dir / "failure_detection_coding_public.csv")

    # Figure 6: NARS prior attitudes.
    nars_columns = ["NARS S1", "NARS S2", "NARS S3"]
    nars_data = [numeric(pre_rows, column) for column in nars_columns]
    fig, ax = plt.subplots(figsize=(8.8, 5.8))
    ax.boxplot(
        nars_data,
        tick_labels=[
            "S1\nDirect interaction",
            "S2\nGeneral concerns",
            "S3\nEmotional/social\n(reverse-scored)",
        ],
        showmeans=True,
        meanprops={"marker": "D", "markersize": 6},
    )
    for position, values in enumerate(nars_data, start=1):
        ax.text(position, 4.82, f"M = {values.mean():.2f}", ha="center", va="top")
    ax.set_ylim(1, 5)
    ax.set_yticks([1, 2, 3, 4, 5])
    ax.set_ylabel("NARS score (1-5; higher = more negative)")
    ax.set_xlabel("NARS subscale")
    save(fig, output_dir, "figure_06_nars_prior_attitudes.png")

    # Figure 7: TAP prior attitudes.
    tap_columns = ["TAP Optimism", "TAP Proficiency", "TAP Dependence"]
    tap_data = [numeric(pre_rows, column) for column in tap_columns]
    fig, ax = plt.subplots(figsize=(8.8, 5.8))
    ax.boxplot(
        tap_data,
        tick_labels=["Optimism", "Proficiency", "Dependence\n(single item)"],
        showmeans=True,
        meanprops={"marker": "D", "markersize": 6},
    )
    for position, values in enumerate(tap_data, start=1):
        ax.text(position, 4.82, f"M = {values.mean():.2f}", ha="center", va="top")
    ax.set_ylim(1, 5)
    ax.set_yticks([1, 2, 3, 4, 5])
    ax.set_ylabel("TAP score (1-5)")
    ax.set_xlabel("TAP component")
    save(fig, output_dir, "figure_07_tap_prior_attitudes.png")

    # Figure 8: Overall Acceptance in control and failure sessions.
    control = numeric(clean_rows, "Overall acceptance control")
    failure = numeric(clean_rows, "Overall acceptance failure")
    fig, ax = plt.subplots(figsize=(8.2, 5.8))
    ax.boxplot(
        [control, failure],
        tick_labels=["Control session", "Failure session"],
        showmeans=True,
        meanprops={"marker": "D", "markersize": 6},
    )
    for position, values in enumerate([control, failure], start=1):
        ax.text(position, 6.85, f"M = {values.mean():.2f}", ha="center", va="top")
    ax.set_ylim(1, 7)
    ax.set_yticks([1, 2, 3, 4, 5, 6, 7])
    ax.set_ylabel("Overall Acceptance score (1-7)")
    ax.set_xlabel("Session condition")
    save(fig, output_dir, "figure_08_overall_acceptance_control_vs_failure.png")

    # Figures 9-10: detection by type and timing.
    detection_specs = [
        (
            "Failure type",
            ["Hardware", "Interaction"],
            ["Hardware failure", "Interaction failure"],
            "Failure type",
            "figure_09_failure_detection_by_type.png",
        ),
        (
            "Timing",
            ["Early", "Late"],
            ["First session\n(Early)", "Second session\n(Late)"],
            "Failure timing",
            "figure_10_failure_detection_by_timing.png",
        ),
    ]
    for field, categories, labels, xlabel, filename in detection_specs:
        percentages: list[float] = []
        counts: list[str] = []
        for category in categories:
            selected = [row for row in detection_rows if row[field] == category]
            detected = sum(row["Detected"] == "Yes" for row in selected)
            percentages.append(100.0 * detected / len(selected))
            counts.append(f"{detected}/{len(selected)}")
        fig, ax = plt.subplots(figsize=(7.8, 5.2))
        bars = ax.bar(labels, percentages)
        ax.set_ylim(0, 100)
        ax.set_yticks([0, 20, 40, 60, 80, 100])
        ax.set_ylabel("Participants who detected the failure (%)")
        ax.set_xlabel(xlabel)
        for bar, percentage, count in zip(bars, percentages, counts):
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                percentage + 3,
                f"{percentage:.0f}%\n({count})",
                ha="center",
                va="bottom",
            )
        save(fig, output_dir, filename)

    # Figure 11: change by failure type.
    hardware = numeric(
        [row for row in clean_rows if row["Failure type"] == "Hardware"],
        "Overall acceptance change",
    )
    interaction = numeric(
        [row for row in clean_rows if row["Failure type"] == "Interaction"],
        "Overall acceptance change",
    )
    fig, ax = plt.subplots(figsize=(8.2, 5.8))
    ax.boxplot(
        [hardware, interaction],
        tick_labels=["Hardware failure", "Interaction failure"],
        showmeans=True,
        meanprops={"marker": "D", "markersize": 6},
    )
    ax.axhline(0, linewidth=1)
    for position, values in enumerate([hardware, interaction], start=1):
        ax.text(position, 1.08, f"M = {values.mean():.2f}", ha="center", va="top")
    ax.set_ylim(-2.5, 1.2)
    ax.set_ylabel("Change in Overall Acceptance\n(failure session - control session)")
    ax.set_xlabel("Failure type")
    save(fig, output_dir, "figure_11_effect_of_failure_type.png")

    # Figure 12: change by failure timing.
    early = numeric(
        [row for row in clean_rows if row["Timing"] == "Early"],
        "Overall acceptance change",
    )
    late = numeric(
        [row for row in clean_rows if row["Timing"] == "Late"],
        "Overall acceptance change",
    )
    fig, ax = plt.subplots(figsize=(8.2, 5.8))
    ax.boxplot(
        [early, late],
        tick_labels=["First session\n(Early)", "Second session\n(Late)"],
        showmeans=True,
        meanprops={"marker": "D", "markersize": 6},
    )
    ax.axhline(0, linewidth=1)
    for position, values in enumerate([early, late], start=1):
        ax.text(position, 1.08, f"M = {values.mean():.2f}", ha="center", va="top")
    ax.set_ylim(-2.5, 1.2)
    ax.set_ylabel("Change in Overall Acceptance\n(failure session - control session)")
    ax.set_xlabel("Failure timing")
    save(fig, output_dir, "figure_12_effect_of_failure_timing.png")

    # Figure 13: failure type x timing interaction.
    timings = ["Early", "Late"]
    failure_types = ["Hardware", "Interaction"]
    fig, ax = plt.subplots(figsize=(8.2, 5.8))
    for failure_type in failure_types:
        group_means: list[float] = []
        standard_errors: list[float] = []
        for timing in timings:
            selected = [
                row
                for row in clean_rows
                if row["Failure type"] == failure_type and row["Timing"] == timing
            ]
            values = numeric(selected, "Overall acceptance change")
            group_means.append(float(values.mean()))
            standard_errors.append(float(values.std(ddof=1) / math.sqrt(len(values))))
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
    ax.set_ylabel("Mean change in Overall Acceptance\n(failure session - control session)")
    ax.set_xlabel("Failure timing")
    ax.legend(title="Failure type")
    save(fig, output_dir, "figure_13_type_timing_interaction.png")


def main() -> None:
    base = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--data-dir", type=Path, default=base / "data")
    parser.add_argument("--output-dir", type=Path, default=base / "figures")
    args = parser.parse_args()
    make_all_figures(args.data_dir.resolve(), args.output_dir.resolve())
    print(f"Generated Figures 6-13 in: {args.output_dir.resolve()}")


if __name__ == "__main__":
    main()
