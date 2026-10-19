from __future__ import annotations

import csv
import json
import math
from dataclasses import dataclass
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from matplotlib.patches import Ellipse
from scipy.optimize import brentq, minimize_scalar
from scipy.stats import chi2


ROOT = Path(__file__).resolve().parent
PROJECT = ROOT.parent
OUTPUT = ROOT / "upb"
KOCHNEV_SOURCE = Path(
    "/Users/sergeymalyshev/Library/CloudStorage/OneDrive-Personal/17-codex/disser/"
    "zircons/Kochnev2022/1-s2.0-S0301926822003217-mmc2.xlsx"
)
KOCHNEV_DOI = "https://doi.org/10.1016/j.precamres.2022.106877"
KOCHNEV_2025_SOURCE = Path(
    "/Users/sergeymalyshev/Library/CloudStorage/OneDrive-Personal/17-codex/disser/"
    "zircons/Kochnev2025/suppl data.xlsx"
)
KOCHNEV_2025_DOI = "https://doi.org/10.5800/GT-2025-16-1-0808"

LAMBDA_238 = 1.55125e-10
LAMBDA_235 = 9.8485e-10
U238_U235 = 137.818
SWITCH_AGE_MA = 1000.0
DISCORDANCE_LIMIT = 10.0

SOURCES = {
    "SM25-03": PROJECT / "SM25-03" / "260913_UPb_Zrn_SM25-03_Pasenko prot.xlsx",
    "SM25-05": PROJECT / "SM25-05" / "260913_UPb_Zrn_SM25-05_Pasenko prot.xlsx",
    "SM25-07": PROJECT / "SM25-07" / "260120_UPb_Zrn_SM25-07_Pasenko prot.xlsx",
    "SM25-16": PROJECT / "SM25-16" / "260120_UPb_Zrn_SM25-16_Pasenko prot.xlsx",
}

SOURCE_PREFIXES = {
    # The SM25-05 workbook omits the leading zero in spot identifiers.
    "SM25-05": "SM25-5_",
}


@dataclass
class Analysis:
    sid: str
    pb: float
    th: float
    u: float
    thu: float
    ratio_86: float
    err_ratio_86: float
    ratio_76: float
    err_ratio_76: float
    rho_4: float
    ratio_75: float
    err_ratio_75: float
    ratio_68: float
    err_ratio_68: float
    rho_3: float
    age_76: float
    err_age_76: float
    age_68: float
    err_age_68: float
    age_75: float
    err_age_75: float
    d1: float
    d2: float
    age_207corr: float
    err_age_207corr: float
    best_age: float = math.nan
    best_err: float = math.nan
    age_system: str = ""
    discordance_metric: str = ""
    discordance: float = math.nan
    analytical_pass: bool = False
    context_excluded: bool = False
    accepted: bool = False
    reason: str = ""


def load_sample(sample: str, path: Path) -> list[Analysis]:
    sheet = openpyxl.load_workbook(path, data_only=True)["Report"]
    result: list[Analysis] = []
    source_prefix = SOURCE_PREFIXES.get(sample, sample + "_")
    for row in range(1, sheet.max_row + 1):
        sid = sheet.cell(row, 1).value
        if not isinstance(sid, str) or not sid.startswith(source_prefix):
            continue
        normalized_sid = sample + "_" + sid[len(source_prefix) :]
        result.append(
            Analysis(
                sid=normalized_sid,
                pb=float(sheet.cell(row, 2).value),
                th=float(sheet.cell(row, 3).value),
                u=float(sheet.cell(row, 4).value),
                thu=float(sheet.cell(row, 5).value),
                ratio_86=float(sheet.cell(row, 6).value),
                err_ratio_86=float(sheet.cell(row, 7).value),
                ratio_76=float(sheet.cell(row, 8).value),
                err_ratio_76=float(sheet.cell(row, 9).value),
                rho_4=float(sheet.cell(row, 10).value),
                ratio_75=float(sheet.cell(row, 11).value),
                err_ratio_75=float(sheet.cell(row, 12).value),
                ratio_68=float(sheet.cell(row, 13).value),
                err_ratio_68=float(sheet.cell(row, 14).value),
                rho_3=float(sheet.cell(row, 15).value),
                age_76=float(sheet.cell(row, 18).value),
                err_age_76=float(sheet.cell(row, 19).value),
                age_68=float(sheet.cell(row, 20).value),
                err_age_68=float(sheet.cell(row, 21).value),
                age_75=float(sheet.cell(row, 22).value),
                err_age_75=float(sheet.cell(row, 23).value),
                d1=float(sheet.cell(row, 26).value),
                d2=float(sheet.cell(row, 27).value),
                age_207corr=float(sheet.cell(row, 28).value),
                err_age_207corr=float(sheet.cell(row, 29).value),
            )
        )
    return result


def is_context_contaminant_sm25_16(a: Analysis) -> bool:
    # These nine young analyses were explicitly treated as contamination / ice-rafted
    # material in the previous interpretation. This flag is kept separate from the
    # analytical discordance filter because it is a geological/contextual decision.
    return (
        230 <= a.age_68 <= 254
        or 473 <= a.age_68 <= 505
        or 753 <= a.age_68 <= 856
    )


def classify(sample: str, analyses: list[Analysis]) -> None:
    for a in analyses:
        if a.age_68 < SWITCH_AGE_MA:
            a.best_age = a.age_68
            a.best_err = a.err_age_68
            a.age_system = "206Pb/238U"
            a.discordance_metric = "D1: 207Pb/235U vs 206Pb/238U"
            a.discordance = a.d1
        else:
            a.best_age = a.age_76
            a.best_err = a.err_age_76
            a.age_system = "207Pb/206Pb"
            a.discordance_metric = "D2: 207Pb/206Pb vs 206Pb/238U"
            a.discordance = a.d2
        a.analytical_pass = abs(a.discordance) <= DISCORDANCE_LIMIT
        a.context_excluded = sample == "SM25-16" and is_context_contaminant_sm25_16(a)
        a.accepted = a.analytical_pass and not a.context_excluded
        if not a.analytical_pass:
            a.reason = f"|{a.discordance_metric.split(':')[0]}| ≥ {DISCORDANCE_LIMIT:.0f}%"
        elif a.context_excluded:
            a.reason = "контекстное исключение из прежней интерпретации"
        else:
            a.reason = "принят"


def weighted_mean(analyses: list[Analysis], age_field: str, err_field: str) -> tuple[float, float, float]:
    ages = np.array([getattr(a, age_field) for a in analyses], dtype=float)
    errors = np.array([getattr(a, err_field) for a in analyses], dtype=float)
    weights = 1.0 / errors**2
    mean = float(np.sum(weights * ages) / np.sum(weights))
    uncertainty_2s = float(2.0 / np.sqrt(np.sum(weights)))
    mswd = float(np.sum(weights * (ages - mean) ** 2) / (len(ages) - 1))
    return mean, uncertainty_2s, mswd


def add_error_ellipse(ax, a: Analysis, color: str, alpha: float, zorder: int) -> None:
    # Wetherill coordinates are 207Pb/235U (x) and 206Pb/238U (y).
    # The source Report supplies their 1-sigma absolute errors and Rho3.
    covariance = np.array(
        [
            [a.err_ratio_75**2, a.rho_3 * a.err_ratio_75 * a.err_ratio_68],
            [a.rho_3 * a.err_ratio_75 * a.err_ratio_68, a.err_ratio_68**2],
        ]
    )
    eigenvalues, eigenvectors = np.linalg.eigh(covariance)
    order = eigenvalues.argsort()[::-1]
    eigenvalues = eigenvalues[order]
    eigenvectors = eigenvectors[:, order]
    angle = math.degrees(math.atan2(eigenvectors[1, 0], eigenvectors[0, 0]))
    width, height = 4.0 * np.sqrt(np.maximum(eigenvalues, 0))  # full 2σ axes
    ax.add_patch(
        Ellipse(
            (a.ratio_75, a.ratio_68),
            width=width,
            height=height,
            angle=angle,
            facecolor=color,
            edgecolor=color,
            linewidth=0.7,
            alpha=alpha,
            zorder=zorder,
        )
    )


def concordia_curve(ax, min_age: float, max_age: float) -> None:
    ages = np.linspace(min_age, max_age, 500)
    ratio_68 = np.expm1(LAMBDA_238 * ages * 1e6)
    x = np.expm1(LAMBDA_235 * ages * 1e6)
    y = ratio_68
    ax.plot(x, y, color="#202925", linewidth=1.7, zorder=1)
    tick_step = 200 if max_age - min_age <= 1800 else 500
    start = math.ceil(min_age / tick_step) * tick_step
    for age in np.arange(start, max_age + 1, tick_step):
        ty = math.expm1(LAMBDA_238 * age * 1e6)
        tx = math.expm1(LAMBDA_235 * age * 1e6)
        ax.scatter([tx], [ty], s=13, color="#202925", zorder=2)
        ax.annotate(f"{int(age)}", (tx, ty), xytext=(4, 3), textcoords="offset points", fontsize=7, color="#4e5954")


def style_axis(ax) -> None:
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(True, color="#dfe5e1", linewidth=0.65, alpha=0.85)
    ax.tick_params(labelsize=9, colors="#2b3430")


def plot_sm25_07(analyses: list[Analysis]) -> dict[str, float | int | str]:
    accepted = [a for a in analyses if a.accepted]
    rejected = [a for a in analyses if not a.accepted]
    mean, uncertainty, mswd = weighted_mean(accepted, "age_68", "err_age_68")
    corrected_mean, corrected_uncertainty, corrected_mswd = weighted_mean(
        accepted, "age_207corr", "err_age_207corr"
    )

    fig, axes = plt.subplots(1, 2, figsize=(14.4, 7.7), gridspec_kw={"width_ratios": [1.05, 1.25]})
    fig.patch.set_facecolor("#f7f9f7")
    fig.suptitle("SM25-07 · U–Pb LA-ICP-MS, циркон", x=0.055, ha="left", fontsize=19, fontweight="bold")
    fig.text(
        0.055,
        0.925,
        "Фильтр: |D1| ≤ 10% для молодых анализов; возраст и среднее — по ²⁰⁶Pb/²³⁸U",
        fontsize=10.5,
        color="#53615a",
    )

    ax = axes[0]
    concordia_curve(ax, 400, 2600)
    for a in rejected:
        add_error_ellipse(ax, a, "#b9bfbb", 0.26, 2)
    for a in accepted:
        add_error_ellipse(ax, a, "#167d70", 0.42, 3)
    ax.scatter([a.ratio_75 for a in accepted], [a.ratio_68 for a in accepted], s=11, color="#0d685e", zorder=4)
    ax.set_xlim(0.6, 10.0)
    ax.set_ylim(0.065, 0.30)
    ax.set_xlabel("²⁰⁷Pb/²³⁵U", fontsize=10)
    ax.set_ylabel("²⁰⁶Pb/²³⁸U", fontsize=10)
    ax.set_title("A  Wetherill concordia ²⁰⁷Pb/²³⁵U–²⁰⁶Pb/²³⁸U · эллипсы 2σ", loc="left", fontsize=12, fontweight="bold")
    style_axis(ax)
    ax.text(
        0.03,
        0.97,
        f"принято: {len(accepted)}\nисключено: {len(rejected)}",
        transform=ax.transAxes,
        va="top",
        fontsize=9.5,
        bbox={"boxstyle": "round,pad=.45", "fc": "white", "ec": "#ccd4cf"},
    )

    ax = axes[1]
    ordered = sorted(accepted, key=lambda item: item.age_68)
    y_index = np.arange(1, len(ordered) + 1)
    ages = np.array([a.age_68 for a in ordered])
    errors_2s = np.array([2 * a.err_age_68 for a in ordered])
    ax.axvspan(mean - uncertainty, mean + uncertainty, color="#a9d5cf", alpha=0.62, zorder=0)
    ax.axvline(mean, color="#bb3d3d", linestyle="--", linewidth=1.5)
    ax.errorbar(ages, y_index, xerr=errors_2s, fmt="o", color="#126e64", ecolor="#438d85", ms=4.7, elinewidth=1.0, capsize=0)
    ax.set_xlim(760, 880)
    ax.set_ylim(0, len(ordered) + 1)
    ax.set_xlabel("Возраст ²⁰⁶Pb/²³⁸U, млн лет", fontsize=10)
    ax.set_ylabel("Принятые анализы, ранжированные по возрасту", fontsize=10)
    ax.set_yticks([])
    ax.set_title("B  Принятые анализы и взвешенное среднее", loc="left", fontsize=12, fontweight="bold")
    style_axis(ax)
    ax.text(
        0.97,
        0.07,
        f"{mean:.1f} ± {uncertainty:.1f} млн лет (2σ, внутр.)\nMSWD = {mswd:.2f}; n = {len(accepted)}",
        transform=ax.transAxes,
        ha="right",
        va="bottom",
        fontsize=11,
        bbox={"boxstyle": "round,pad=.55", "fc": "white", "ec": "#a9b5ae"},
    )
    fig.text(
        0.055,
        0.025,
        "Примечание: систематическая неопределённость метода не добавлена. Common-Pb-corrected 207PbCorr показан только в таблице аудита и не подменяет подписанный ²⁰⁶Pb/²³⁸U возраст.",
        fontsize=8.5,
        color="#68746e",
    )
    fig.subplots_adjust(left=0.075, right=0.985, bottom=0.105, top=0.86, wspace=0.20)
    fig.savefig(OUTPUT / "SM25-07_upb_summary.png", dpi=180, facecolor=fig.get_facecolor())
    plt.close(fig)
    return {
        "sample": "SM25-07",
        "measured": len(analyses),
        "accepted": len(accepted),
        "weighted_mean": mean,
        "weighted_mean_2s": uncertainty,
        "mswd": mswd,
        "corrected_mean": corrected_mean,
        "corrected_mean_2s": corrected_uncertainty,
        "corrected_mswd": corrected_mswd,
        "summary": f"²⁰⁶Pb/²³⁸U {mean:.1f} ± {uncertainty:.1f} млн лет (2σ), MSWD {mswd:.2f}, n={len(accepted)}",
    }


def plot_sm25_16(analyses: list[Analysis]) -> dict[str, float | int | str]:
    accepted = [a for a in analyses if a.accepted]
    discordant = [a for a in analyses if not a.analytical_pass]
    context = [a for a in analyses if a.analytical_pass and a.context_excluded]

    fig, axes = plt.subplots(1, 2, figsize=(14.4, 7.7), gridspec_kw={"width_ratios": [1.0, 1.35]})
    fig.patch.set_facecolor("#f7f9f7")
    fig.suptitle("SM25-16 · U–Pb LA-ICP-MS, обломочный циркон", x=0.055, ha="left", fontsize=19, fontweight="bold")
    fig.text(
        0.055,
        0.925,
        "Best age: ²⁰⁶Pb/²³⁸U < 1000 млн лет; ²⁰⁷Pb/²⁰⁶Pb ≥ 1000 млн лет · гибридный фильтр |D| ≤ 10%",
        fontsize=10.5,
        color="#53615a",
    )

    ax = axes[0]
    concordia_curve(ax, 150, 3300)
    for a in discordant:
        add_error_ellipse(ax, a, "#aeb4b0", 0.24, 2)
    for a in context:
        add_error_ellipse(ax, a, "#d77b28", 0.34, 3)
    for a in accepted:
        add_error_ellipse(ax, a, "#167d70", 0.30, 3)
    ax.set_xlim(0.0, 16.0)
    ax.set_ylim(0.02, 0.70)
    ax.set_xlabel("²⁰⁷Pb/²³⁵U", fontsize=10)
    ax.set_ylabel("²⁰⁶Pb/²³⁸U", fontsize=10)
    ax.set_title("A  Wetherill concordia ²⁰⁷Pb/²³⁵U–²⁰⁶Pb/²³⁸U · эллипсы 2σ", loc="left", fontsize=12, fontweight="bold")
    style_axis(ax)
    ax.text(
        0.03,
        0.97,
        f"принято: {len(accepted)}\nдискордантные: {len(discordant)}\nконтекстно исключены: {len(context)}",
        transform=ax.transAxes,
        va="top",
        fontsize=9.5,
        bbox={"boxstyle": "round,pad=.45", "fc": "white", "ec": "#ccd4cf"},
    )

    ax = axes[1]
    accepted_ages = np.array([a.best_age for a in accepted])
    x = np.linspace(0, 3500, 1751)
    bandwidth = 30.0
    density = np.exp(-0.5 * ((x[:, None] - accepted_ages[None, :]) / bandwidth) ** 2).sum(axis=1)
    density /= len(accepted_ages) * bandwidth * math.sqrt(2 * math.pi)
    peak_indexes = np.where((density[1:-1] > density[:-2]) & (density[1:-1] >= density[2:]))[0] + 1
    peaks = [(float(x[i]), float(density[i])) for i in peak_indexes]
    ax.fill_between(x, density, color="#a9d5cf", alpha=0.55)
    ax.plot(x, density, color="#126e64", linewidth=2.0, label="KDE принятых анализов")
    for peak_age, peak_density in peaks:
        ax.scatter([peak_age], [peak_density], s=18, color="#126e64", zorder=4)
        ax.annotate(f"{peak_age:.0f}", (peak_age, peak_density), xytext=(0, 6), textcoords="offset points", ha="center", fontsize=7, color="#415049")
    for a in accepted:
        color = "#126e64" if a.age_system == "206Pb/238U" else "#6656a8"
        ax.vlines(a.best_age, 0, density.max() * 0.035, color=color, linewidth=0.8, alpha=0.75)
    for a in context:
        ax.vlines(a.best_age, 0, density.max() * 0.055, color="#d77b28", linewidth=1.4)
    for a in discordant:
        ax.vlines(a.best_age, 0, density.max() * 0.045, color="#aeb4b0", linewidth=1.0)
    ax.set_xlim(0, 3500)
    ax.set_ylim(bottom=0)
    ax.set_xlabel("Best age, млн лет", fontsize=10)
    ax.set_ylabel("Плотность вероятности", fontsize=10)
    ax.set_yticks([])
    ax.set_title("B  Спектр возрастов принятых анализов", loc="left", fontsize=12, fontweight="bold")
    style_axis(ax)
    ax.text(
        0.97,
        0.96,
        f"n = {len(accepted)} из {len(analyses)}\nпики KDE: {len(peaks)} · h = {bandwidth:.0f} млн лет",
        transform=ax.transAxes,
        ha="right",
        va="top",
        fontsize=10,
        bbox={"boxstyle": "round,pad=.45", "fc": "white", "ec": "#ccd4cf"},
    )
    legend_lines = [
        plt.Line2D([0], [0], color="#126e64", lw=3, label="²⁰⁶Pb/²³⁸U (<1000 млн лет)"),
        plt.Line2D([0], [0], color="#6656a8", lw=3, label="²⁰⁷Pb/²⁰⁶Pb (≥1000 млн лет)"),
        plt.Line2D([0], [0], color="#d77b28", lw=3, label="контекстно исключены"),
        plt.Line2D([0], [0], color="#aeb4b0", lw=3, label="дискордантные"),
    ]
    ax.legend(handles=legend_lines, loc="upper left", frameon=False, fontsize=8.7)
    fig.text(
        0.055,
        0.025,
        "Контекстные исключения (9 молодых зёрен) перенесены из прежней интерпретации и не являются результатом U–Pb-фильтра. Взвешенное среднее для детритового спектра не рассчитывается.",
        fontsize=8.5,
        color="#68746e",
    )
    fig.subplots_adjust(left=0.075, right=0.985, bottom=0.105, top=0.86, wspace=0.28)
    fig.savefig(OUTPUT / "SM25-16_upb_summary.png", dpi=180, facecolor=fig.get_facecolor())
    plt.close(fig)
    return {
        "sample": "SM25-16",
        "measured": len(analyses),
        "accepted": len(accepted),
        "discordant": len(discordant),
        "context_excluded": len(context),
        "kde_peak": float(x[np.argmax(density)]),
        "kde_bandwidth": bandwidth,
        "kde_peaks": [round(age, 1) for age, _ in peaks],
        "summary": f"спектр best age: принято {len(accepted)} из {len(analyses)}, основной максимум ≈ {x[np.argmax(density)]:.0f} млн лет",
    }


def kde_summary(analyses: list[Analysis], bandwidth: float = 30.0) -> tuple[np.ndarray, np.ndarray, list[float]]:
    ages = np.array([a.best_age for a in analyses], dtype=float)
    lower = max(0.0, math.floor((float(ages.min()) - 150.0) / 50.0) * 50.0)
    upper = math.ceil((float(ages.max()) + 150.0) / 50.0) * 50.0
    x = np.linspace(lower, upper, max(900, int((upper - lower) / 2) + 1))
    density = np.exp(-0.5 * ((x[:, None] - ages[None, :]) / bandwidth) ** 2).sum(axis=1)
    density /= len(ages) * bandwidth * math.sqrt(2 * math.pi)
    indexes = np.where((density[1:-1] > density[:-2]) & (density[1:-1] >= density[2:]))[0] + 1
    return x, density, [float(x[index]) for index in indexes]


def plot_new_spectrum(
    sample: str,
    analyses: list[Analysis],
    sample_kind: str,
    title: str,
    interpretation: str,
) -> dict[str, object]:
    accepted = [a for a in analyses if a.accepted]
    rejected = [a for a in analyses if not a.analytical_pass]
    x, density, peaks = kde_summary(accepted)

    fig, axes = plt.subplots(1, 2, figsize=(14.4, 7.7), gridspec_kw={"width_ratios": [1.0, 1.35]})
    fig.patch.set_facecolor("#f7f9f7")
    fig.suptitle(f"{sample} · {title}", x=0.055, ha="left", fontsize=19, fontweight="bold")
    fig.text(
        0.055,
        0.925,
        "Best age: ²⁰⁶Pb/²³⁸U < 1000 млн лет; ²⁰⁷Pb/²⁰⁶Pb ≥ 1000 млн лет · гибридный фильтр |D| ≤ 10%",
        fontsize=10.5,
        color="#53615a",
    )

    ax = axes[0]
    concordia_curve(ax, max(0, min(a.best_age for a in analyses) - 100), min(3500, max(a.best_age for a in analyses) + 150))
    for a in rejected:
        add_error_ellipse(ax, a, "#aeb4b0", 0.24, 2)
    for a in accepted:
        add_error_ellipse(ax, a, "#167d70", 0.30, 3)
    max_x = max(a.ratio_75 + 2 * a.err_ratio_75 for a in analyses)
    max_y = max(a.ratio_68 + 2 * a.err_ratio_68 for a in analyses)
    ax.set_xlim(0, min(32, max_x * 1.08))
    ax.set_ylim(0, min(0.75, max_y * 1.08))
    ax.set_xlabel("²⁰⁷Pb/²³⁵U", fontsize=10)
    ax.set_ylabel("²⁰⁶Pb/²³⁸U", fontsize=10)
    ax.set_title("A  Wetherill concordia · эллипсы 2σ", loc="left", fontsize=12, fontweight="bold")
    style_axis(ax)
    ax.text(
        0.03, 0.97,
        f"принято: {len(accepted)}\nдискордантные: {len(rejected)}",
        transform=ax.transAxes, va="top", fontsize=9.5,
        bbox={"boxstyle": "round,pad=.45", "fc": "white", "ec": "#ccd4cf"},
    )

    ax = axes[1]
    ax.fill_between(x, density, color="#a9d5cf", alpha=0.55)
    ax.plot(x, density, color="#126e64", linewidth=2.0)
    for peak_age in peaks:
        peak_density = float(np.interp(peak_age, x, density))
        ax.scatter([peak_age], [peak_density], s=18, color="#126e64", zorder=4)
        ax.annotate(f"{peak_age:.0f}", (peak_age, peak_density), xytext=(0, 6), textcoords="offset points", ha="center", fontsize=7, color="#415049")
    for a in accepted:
        color = "#126e64" if a.age_system == "206Pb/238U" else "#6656a8"
        ax.vlines(a.best_age, 0, density.max() * 0.04, color=color, linewidth=0.8, alpha=0.75)
    for a in rejected:
        ax.vlines(a.best_age, 0, density.max() * 0.055, color="#aeb4b0", linewidth=1.0)
    ax.set_xlim(x[0], x[-1])
    ax.set_ylim(bottom=0)
    ax.set_xlabel("Best age, млн лет", fontsize=10)
    ax.set_ylabel("Плотность вероятности", fontsize=10)
    ax.set_yticks([])
    ax.set_title("B  Спектр возрастов принятых анализов", loc="left", fontsize=12, fontweight="bold")
    style_axis(ax)
    ax.text(
        0.97, 0.96,
        f"n = {len(accepted)} из {len(analyses)}\nпики KDE: {len(peaks)} · h = 30 млн лет",
        transform=ax.transAxes, ha="right", va="top", fontsize=10,
        bbox={"boxstyle": "round,pad=.45", "fc": "white", "ec": "#ccd4cf"},
    )
    fig.text(0.055, 0.025, interpretation, fontsize=8.5, color="#68746e")
    fig.subplots_adjust(left=0.075, right=0.985, bottom=0.105, top=0.86, wspace=0.28)
    fig.savefig(OUTPUT / f"{sample}_upb_summary.png", dpi=180, facecolor=fig.get_facecolor())
    plt.close(fig)

    result: dict[str, object] = {
        "sample": sample,
        "sample_kind": sample_kind,
        "measured": len(analyses),
        "accepted": len(accepted),
        "discordant": len(rejected),
        "kde_peak": float(x[np.argmax(density)]),
        "kde_bandwidth": 30.0,
        "kde_peaks": [round(age, 1) for age in peaks],
        "interpretation": interpretation,
    }
    if sample == "SM25-05":
        youngest = sorted(accepted, key=lambda a: a.best_age)[:3]
        cluster_mean, cluster_2s, cluster_mswd = weighted_mean(youngest, "best_age", "best_err")
        youngest_single = min(accepted, key=lambda a: a.best_age)
        result.update(
            {
                "youngest_single": youngest_single.best_age,
                "youngest_single_2s": 2 * youngest_single.best_err,
                "youngest_cluster": cluster_mean,
                "youngest_cluster_2s": cluster_2s,
                "youngest_cluster_mswd": cluster_mswd,
                "youngest_cluster_n": len(youngest),
                "summary": (
                    f"губинская свита · {len(accepted)}/{len(analyses)} принято; "
                    f"молодой кластер 3 анализов {cluster_mean:.1f} ± {cluster_2s:.1f} млн лет (2σ)"
                ),
            }
        )
    else:
        result["summary"] = (
            f"базитовый силл · {len(accepted)}/{len(analyses)} принято; "
            "единая магматическая популяция не выделяется"
        )
    return result


def write_audit(sample: str, analyses: list[Analysis]) -> None:
    fields = [
        "sid", "pb", "th", "u", "thu", "ratio_86", "err_ratio_86", "ratio_76", "err_ratio_76", "rho_4", "ratio_75", "err_ratio_75", "ratio_68", "err_ratio_68", "rho_3", "age_68", "err_age_68", "age_76", "err_age_76",
        "age_75", "err_age_75", "d1", "d2", "age_207corr", "err_age_207corr",
        "best_age", "best_err", "age_system", "discordance_metric", "discordance",
        "analytical_pass", "context_excluded", "accepted", "reason",
    ]
    with (OUTPUT / f"{sample}_upb_audit.csv").open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for a in analyses:
            writer.writerow({field: getattr(a, field) for field in fields})


def write_report(results: dict[str, dict[str, object]]) -> None:
    r03 = results["SM25-03"]
    r05 = results["SM25-05"]
    r07 = results["SM25-07"]
    r16 = results["SM25-16"]
    report = f"""# Аудит U–Pb расчётов по проекту Игарка

## Принятые правила пересчёта

- Порог дискордантности: `|D| ≤ {DISCORDANCE_LIMIT:.0f}%`.
- Для анализов с ²⁰⁶Pb/²³⁸U возрастом < {SWITCH_AGE_MA:.0f} млн лет используется ²⁰⁶Pb/²³⁸U возраст и D1 (сопоставление ²⁰⁷Pb/²³⁵U и ²⁰⁶Pb/²³⁸U возрастов).
- Для анализов с ²⁰⁶Pb/²³⁸U возрастом ≥ {SWITCH_AGE_MA:.0f} млн лет используется ²⁰⁷Pb/²⁰⁶Pb возраст и D2 (сопоставление ²⁰⁷Pb/²⁰⁶Pb и ²⁰⁶Pb/²³⁸U возрастов).
- Все неопределённости в исходном Report заданы как 1σ; на графиках показаны 2σ.
- Рассчитанные интервалы внутренние; внешняя систематическая погрешность сессии не добавлена.

## SM25-03 — базитовый силл

При гибридном фильтре принято {r03['accepted']} из {r03['measured']} анализов. Спектр от фанерозойских до архейских возрастов не образует одну магматическую популяцию. Группа 455–515 млн лет даёт 478,3 ± 7,3 млн лет (2σ, расширено по MSWD; MSWD 2,76; n=23): это кандидат на проверку, а не принятый возраст внедрения. Группа 800–834 млн лет даёт 812,9 ± 20,0 млн лет (2σ; MSWD 0,51; n=3), но она выделена апостериорно. Ни одна из этих групп не датирует силл без CL/BSE-контекста спотов. Возрасты 134–650 млн лет нельзя автоматически объявлять «геологически невозможными» и потерей Pb: проба интрузивная, а многие молодые анализы формально конкордантны по D1. Пока датировка силла не получена.

## SM25-05 — губинская свита

При гибридном фильтре принято {r05['accepted']} из {r05['measured']} анализов. Молодое одиночное зерно: {r05['youngest_single']:.1f} ± {r05['youngest_single_2s']:.1f} млн лет (2σ). Три самых молодых анализа перекрываются на 2σ и дают кандидат на максимальный возраст осадконакопления: **{r05['youngest_cluster']:.1f} ± {r05['youngest_cluster_2s']:.1f} млн лет (2σ; MSWD {r05['youngest_cluster_mswd']:.2f}; n={r05['youngest_cluster_n']})**. Это предварительный MDA, потому что выбор метрики дискордантности существенно влияет на самые молодые зёрна.

### Проверка заключения Claude по SM25-03 и SM25-05

- Арифметика Claude в основном воспроизводится, если для всех возрастов применить `|D2| < 10%`.
- Для моложе 1 млрд лет D2 опирается на малоточный ²⁰⁷Pb/²⁰⁶Pb возраст. Гибридный D1/D2 меняет число принятых анализов с 68 до {r03['accepted']} для SM25-03 и с 96 до {r05['accepted']} для SM25-05.
- Два молодых анализа SM25-05 (755,2 и 761,2 млн лет) проходят D1, но отбрасываются D2. Поэтому кластер Claude ~780 млн лет зависит от метода фильтрации.
- Выборки n=3, 5 и 8 у Claude вложены друг в друга и не являются тремя независимыми оценками. Сам выбор ровно 3/5/8 зёрен алгоритмом не обоснован.
- Числа, названные Claude `2σ`, фактически считались с множителем 1,96, то есть как 95%-й интервал, а не буквальные 2σ.
- Главная геологическая ошибка: SM25-03 — базитовый силл, а не образец лудовской свиты. Поэтому интерпретация его молодых возрастов как «невозможных» и вывод об MDA к нему неприменимы.

## SM25-07

Принят {r07['accepted']} анализ из {r07['measured']}. Взвешенное среднее по фактической колонке ²⁰⁶Pb/²³⁸U:
**{r07['weighted_mean']:.1f} ± {r07['weighted_mean_2s']:.1f} млн лет (2σ, внутр.); MSWD = {r07['mswd']:.2f}.**

Ключевая ошибка прежнего рисунка: опубликованная цифра 818,3 ± 6,0 млн лет воспроизводится по колонке `207PbCorr`, а не по подписанной колонке ²⁰⁶Pb/²³⁸U. Для принятой здесь выборки `207PbCorr` даёт {r07['corrected_mean']:.1f} ± {r07['corrected_mean_2s']:.1f} млн лет (MSWD {r07['corrected_mswd']:.2f}); этот результат нельзя подписывать как необработанный ²⁰⁶Pb/²³⁸U возраст.

## SM25-16

После гибридного фильтра принято {r16['accepted']} из {r16['measured']} анализов; {r16['discordant']} исключены как дискордантные и ещё {r16['context_excluded']} — по контекстной интерпретации прежнего рисунка. При фиксированной полосе KDE {r16['kde_bandwidth']:.0f} млн лет рассчитаны все локальные максимумы: {', '.join(f'{age:.0f}' for age in r16['kde_peaks'])} млн лет.

Взвешенное среднее 1400,7 ± 9,1 млн лет для 90 детритовых зёрен арифметически воспроизводится, но не используется как единый возраст образца: это спектр обломочных цирконов, а низкий MSWD сам по себе не доказывает одну геологическую популяцию.

## Отличие от фильтра Claude

Claude применил D1 ко всем возрастам. Для старых анализов здесь используется D2, согласованный с выбором ²⁰⁷Pb/²⁰⁶Pb best age. Поэтому SM25-16_64 (D1 = 5,59%, D2 = 11,02%) теперь исключён, и число принятых анализов меняется с 93 на 92.
"""
    (OUTPUT / "UPb_audit_report.md").write_text(report, encoding="utf-8")


def write_portal_data(results: dict[str, dict[str, object]], samples: dict[str, list[Analysis]]) -> None:
    analysis_fields = (
        "sid", "ratio_75", "err_ratio_75", "ratio_68", "err_ratio_68", "rho_3",
        "age_76", "err_age_76", "age_68", "err_age_68", "age_75", "err_age_75",
        "d1", "d2", "age_207corr", "err_age_207corr", "context_excluded",
    )
    payload = {
        sample: {
            **values,
            "figure": f"upb/{sample}_upb_summary.png",
            "audit": f"upb/{sample}_upb_audit.csv",
            "analyses": [
                {field: getattr(analysis, field) for field in analysis_fields}
                for analysis in samples[sample]
            ],
        }
        for sample, values in results.items()
    }
    (OUTPUT / "portal_data.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def load_kochnev_sample(prefix: str) -> list[dict[str, object]]:
    """Load one published Kochnev et al. zircon dataset by analytical prefix."""
    sheet = openpyxl.load_workbook(KOCHNEV_SOURCE, data_only=True, read_only=True)["Table"]
    analyses: list[dict[str, object]] = []
    for row in sheet.iter_rows(values_only=True):
        sid = row[0]
        if not isinstance(sid, str) or not sid.startswith(f"{prefix}-"):
            continue
        discordance = float(row[15])
        age_68, err_68_2s = float(row[13]), float(row[14])
        age_76, err_76_2s = float(row[9]), float(row[10])
        if isinstance(row[17], (int, float)) and isinstance(row[18], (int, float)):
            best_age, best_err_2s = float(row[17]), float(row[18])
        elif age_68 < SWITCH_AGE_MA:
            best_age, best_err_2s = age_68, err_68_2s
        else:
            best_age, best_err_2s = age_76, err_76_2s
        analyses.append(
            {
                "sid": sid,
                "age": best_age,
                "err_1s": best_err_2s / 2.0,
                "err_2s": best_err_2s,
                "discordance": discordance,
                "accepted": abs(discordance) <= DISCORDANCE_LIMIT,
                "flagged_by_author": "*" in sid,
            }
        )
    if len(analyses) != 130:
        raise ValueError(f"Expected 130 analyses for Kochnev {prefix}, found {len(analyses)}")
    return analyses


def load_kochnev_2025_1812_1() -> list[dict[str, object]]:
    """Load the dolerite sample; source errors are propagated 2SE."""
    sheet = openpyxl.load_workbook(
        KOCHNEV_2025_SOURCE, data_only=True, read_only=True
    )["1812-1 Dolerite"]
    analyses: list[dict[str, object]] = []
    for row in sheet.iter_rows(min_row=3, values_only=True):
        if not row[0]:
            continue
        age_68 = float(row[6])
        age_76 = float(row[16])
        age_75 = float(row[8])
        d1 = 100.0 * (age_75 / age_68 - 1.0)
        d2 = 100.0 * (age_76 / age_68 - 1.0)
        cluster = 800.0 <= age_68 <= 880.0
        if cluster:
            role = "кластер 835 млн лет"
        elif age_68 < 1000.0:
            role = "ксенокристалл ~900 млн лет"
        else:
            role = "ксенокристалл ~1,88 млрд лет"
        analyses.append(
            {
                "sid": str(row[0]),
                "ratio_75": float(row[1]),
                "err_ratio_75_2s": float(row[2]),
                "ratio_68": float(row[3]),
                "err_ratio_68_2s": float(row[4]),
                "rho": float(row[5]),
                "age_68": age_68,
                "err_age_68_2s": float(row[7]),
                "age_75": age_75,
                "err_age_75_2s": float(row[9]),
                "age_76": age_76,
                "err_age_76_2s": float(row[17]),
                "u_th": float(row[20]),
                "u_ppm": float(row[21]),
                "th_ppm": float(row[22]),
                "pb_ppm": float(row[23]),
                "d1": d1,
                "d2": d2,
                "cluster_835": cluster,
                "role": role,
            }
        )
    if len(analyses) != 15:
        raise ValueError(f"Expected 15 analyses for Kochnev 1812-1, found {len(analyses)}")
    return analyses


def load_kochnev_2025_2007_15() -> list[dict[str, object]]:
    """Load the Ludy Formation detrital-zircon sample; source errors are 2SE."""
    sheet = openpyxl.load_workbook(
        KOCHNEV_2025_SOURCE, data_only=True, read_only=True
    )["2007-15"]
    analyses: list[dict[str, object]] = []
    for row in sheet.iter_rows(min_row=3, values_only=True):
        if not row[0]:
            continue
        age_68 = float(row[6])
        age_75 = float(row[8])
        age_76 = float(row[16])
        d1 = 100.0 * (age_75 / age_68 - 1.0)
        d2 = 100.0 * (age_76 / age_68 - 1.0)
        use_68 = age_68 < SWITCH_AGE_MA
        discordance = d1 if use_68 else d2
        analyses.append(
            {
                "sid": str(row[0]).strip(),
                "age": age_68 if use_68 else age_76,
                "err_1s": (float(row[7]) if use_68 else float(row[17])) / 2.0,
                "err_2s": float(row[7]) if use_68 else float(row[17]),
                "age_system": "206Pb/238U" if use_68 else "207Pb/206Pb",
                "d1": d1,
                "d2": d2,
                "discordance": discordance,
                "accepted": abs(discordance) <= DISCORDANCE_LIMIT,
            }
        )
    if len(analyses) != 121:
        raise ValueError(f"Expected 121 analyses for Kochnev 2007-15, found {len(analyses)}")
    return analyses


def concordia_1812_1(cluster: list[dict[str, object]]) -> dict[str, float]:
    def model(age_ma: float) -> np.ndarray:
        years = age_ma * 1e6
        return np.array([np.expm1(LAMBDA_235 * years), np.expm1(LAMBDA_238 * years)])

    def covariance(a: dict[str, object]) -> np.ndarray:
        sx = float(a["err_ratio_75_2s"]) / 2.0
        sy = float(a["err_ratio_68_2s"]) / 2.0
        rho = float(a["rho"])
        return np.array([[sx * sx, rho * sx * sy], [rho * sx * sy, sy * sy]])

    def objective(age_ma: float) -> float:
        expected = model(age_ma)
        total = 0.0
        for a in cluster:
            observed = np.array([float(a["ratio_75"]), float(a["ratio_68"])])
            residual = observed - expected
            total += float(residual @ np.linalg.inv(covariance(a)) @ residual)
        return total

    result = minimize_scalar(objective, bounds=(780.0, 900.0), method="bounded")
    age = float(result.x)
    total_chi2 = float(result.fun)
    lo_2s = brentq(lambda value: objective(value) - total_chi2 - 4.0, 760.0, age)
    hi_2s = brentq(lambda value: objective(value) - total_chi2 - 4.0, age, 920.0)

    precisions = [np.linalg.inv(covariance(a)) for a in cluster]
    pooled_covariance = np.linalg.inv(np.sum(precisions, axis=0))
    pooled_ratio = pooled_covariance @ np.sum(
        [precision @ np.array([float(a["ratio_75"]), float(a["ratio_68"])])
         for precision, a in zip(precisions, cluster)],
        axis=0,
    )
    equivalence_chi2 = 0.0
    for precision, a in zip(precisions, cluster):
        residual = np.array([float(a["ratio_75"]), float(a["ratio_68"])]) - pooled_ratio
        equivalence_chi2 += float(residual @ precision @ residual)
    concordance_chi2 = total_chi2 - equivalence_chi2

    ages = np.array([float(a["age_68"]) for a in cluster])
    errors_1s = np.array([float(a["err_age_68_2s"]) / 2.0 for a in cluster])
    weights = 1.0 / errors_1s**2
    wm_age = float(np.sum(weights * ages) / np.sum(weights))
    wm_2s = float(2.0 / np.sqrt(np.sum(weights)))
    wm_chi2 = float(np.sum(weights * (ages - wm_age) ** 2))
    wm_df = len(cluster) - 1
    return {
        "age": age,
        "err_2s_minus": age - lo_2s,
        "err_2s_plus": hi_2s - age,
        "concordance_mswd": concordance_chi2,
        "concordance_probability": float(chi2.sf(concordance_chi2, 1)),
        "equivalence_mswd": equivalence_chi2 / (2 * len(cluster) - 2),
        "equivalence_probability": float(chi2.sf(equivalence_chi2, 2 * len(cluster) - 2)),
        "wm_age_68": wm_age,
        "wm_err_2s": wm_2s,
        "wm_mswd": wm_chi2 / wm_df,
        "wm_probability": float(chi2.sf(wm_chi2, wm_df)),
    }


def add_dict_ellipse(ax, a: dict[str, object], color: str, alpha: float = 0.34) -> None:
    sx = float(a["err_ratio_75_2s"]) / 2.0
    sy = float(a["err_ratio_68_2s"]) / 2.0
    covariance = np.array(
        [[sx * sx, float(a["rho"]) * sx * sy],
         [float(a["rho"]) * sx * sy, sy * sy]]
    )
    eigenvalues, eigenvectors = np.linalg.eigh(covariance)
    order = eigenvalues.argsort()[::-1]
    eigenvalues = eigenvalues[order]
    eigenvectors = eigenvectors[:, order]
    angle = math.degrees(math.atan2(eigenvectors[1, 0], eigenvectors[0, 0]))
    ax.add_patch(
        Ellipse(
            (float(a["ratio_75"]), float(a["ratio_68"])),
            4.0 * math.sqrt(eigenvalues[0]),
            4.0 * math.sqrt(eigenvalues[1]),
            angle=angle,
            facecolor=color,
            edgecolor=color,
            linewidth=0.8,
            alpha=alpha,
        )
    )


def write_kochnev_2025_1812_1() -> tuple[list[dict[str, object]], dict[str, float]]:
    analyses = load_kochnev_2025_1812_1()
    cluster = [a for a in analyses if a["cluster_835"]]
    stats = concordia_1812_1(cluster)

    with (OUTPUT / "kochnev2025_1812_1_audit.csv").open(
        "w", encoding="utf-8-sig", newline=""
    ) as handle:
        writer = csv.DictWriter(handle, fieldnames=list(analyses[0].keys()))
        writer.writeheader()
        writer.writerows(analyses)

    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9})
    fig, axes = plt.subplots(2, 2, figsize=(12.4, 8.8), facecolor="#f6f8f6")
    fig.subplots_adjust(left=0.085, right=0.98, top=0.89, bottom=0.13, wspace=0.24, hspace=0.30)
    fig.suptitle("Силл 1812-1 · U–Pb LA-ICP-MS, циркон", x=0.07, ha="left", fontsize=18, fontweight="bold")
    fig.text(0.07, 0.925, "Kochnev et al., 2025 · эллипсы и ошибки 2σ", color="#5e6b64", fontsize=10)

    teal, amber, purple, curve = "#13756c", "#c17a32", "#73589f", "#28322e"
    colors = [teal if a["cluster_835"] else (amber if a["age_68"] < 1000 else purple) for a in analyses]
    concordia_ages = np.linspace(650.0, 2050.0, 900)
    concordia_x = np.expm1(LAMBDA_235 * concordia_ages * 1e6)
    concordia_y = np.expm1(LAMBDA_238 * concordia_ages * 1e6)

    ax = axes[0, 0]
    ax.plot(concordia_x, concordia_y, color=curve, lw=1.6)
    for a, color in zip(analyses, colors):
        add_dict_ellipse(ax, a, color)
    ax.set_xlim(0.8, 6.1)
    ax.set_ylim(0.10, 0.37)
    ax.set_title("A  Wetherill concordia · все анализы", loc="left", fontweight="bold")
    ax.set_xlabel("²⁰⁷Pb/²³⁵U")
    ax.set_ylabel("²⁰⁶Pb/²³⁸U")
    ax.text(0.03, 0.96, "13 × ~835 Ma\n1 × ~900 Ma\n1 × ~1878 Ma", transform=ax.transAxes, va="top",
            bbox=dict(boxstyle="round,pad=.35", fc="white", ec="#cad2cd", alpha=.92))

    ax = axes[0, 1]
    zoom_ages = np.linspace(780.0, 890.0, 500)
    ax.plot(np.expm1(LAMBDA_235 * zoom_ages * 1e6), np.expm1(LAMBDA_238 * zoom_ages * 1e6), color=curve, lw=1.8)
    for a in cluster:
        add_dict_ellipse(ax, a, teal, .36)
    ax.set_xlim(1.20, 1.34)
    ax.set_ylim(0.130, 0.143)
    ax.set_title("B  Кластер 835 млн лет · увеличение", loc="left", fontweight="bold")
    ax.set_xlabel("²⁰⁷Pb/²³⁵U")
    ax.set_ylabel("²⁰⁶Pb/²³⁸U")
    ax.text(
        0.03, 0.96,
        f"Concordia age = {stats['age']:.1f} ± {(stats['err_2s_minus'] + stats['err_2s_plus']) / 2:.1f} млн лет (2σ)\n"
        f"MSWD concordance = {stats['concordance_mswd']:.2f}; p = {stats['concordance_probability']:.2f}; n = {len(cluster)}",
        transform=ax.transAxes, va="top",
        bbox=dict(boxstyle="round,pad=.35", fc="white", ec="#cad2cd", alpha=.94),
    )

    ax = axes[1, 0]
    ordered = sorted(cluster, key=lambda a: float(a["age_68"]))
    y = np.arange(len(ordered))
    values = np.array([float(a["age_68"]) for a in ordered])
    errors = np.array([float(a["err_age_68_2s"]) for a in ordered])
    mean, mean_2s = stats["wm_age_68"], stats["wm_err_2s"]
    ax.axvspan(mean - mean_2s, mean + mean_2s, color="#a9d5cf", alpha=.55)
    ax.axvline(mean, color="#a33d3d", lw=1.4, ls="--")
    ax.errorbar(values, y, xerr=errors, fmt="o", ms=4, color=teal, ecolor=teal, elinewidth=1, capsize=2)
    ax.set_yticks(y)
    short_ids = [
        str(a["sid"]).split(".FIN2")[0].split("_Linked")[0].replace(" with links", "")
        for a in ordered
    ]
    ax.set_yticklabels(short_ids, fontsize=7.5)
    ax.set_xlim(795, 870)
    ax.set_title("C  ²⁰⁶Pb/²³⁸U weighted mean", loc="left", fontweight="bold")
    ax.set_xlabel("Возраст, млн лет (ошибки 2σ)")
    ax.text(
        0.98, 0.04,
        f"{mean:.1f} ± {mean_2s:.1f} млн лет\nMSWD = {stats['wm_mswd']:.2f}; p = {stats['wm_probability']:.2f}",
        transform=ax.transAxes, ha="right", va="bottom",
        bbox=dict(boxstyle="round,pad=.35", fc="white", ec="#cad2cd", alpha=.94),
    )

    ax = axes[1, 1]
    for a, color in zip(analyses, colors):
        ax.scatter(float(a["u_ppm"]), float(a["u_th"]), s=40, color=color, edgecolor="white", linewidth=.6, zorder=3)
    ax.axhline(1.0, color="#9aa49e", lw=1, ls="--")
    ax.set_xscale("log")
    ax.set_xlim(180, 3200)
    ax.set_ylim(0.55, 1.78)
    ax.set_title("D  Состав циркона", loc="left", fontweight="bold")
    ax.set_xlabel("U, ppm (логарифмическая шкала)")
    ax.set_ylabel("U/Th")
    ax.text(0.03, 0.96, "Высоко-U зерна кластера: U/Th ≈ 0,7–0,9\nНизко-U зерна: U/Th чаще > 1",
            transform=ax.transAxes, va="top",
            bbox=dict(boxstyle="round,pad=.35", fc="white", ec="#cad2cd", alpha=.94))

    for ax in axes.flat:
        ax.set_facecolor("#fbfcfb")
        ax.grid(True, color="#dfe5e1", linewidth=.65)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
    fig.text(
        0.085, 0.025,
        "Кластер 835 млн лет статистически реален, но старше MDA вмещающих песчаников (716 ± 10 млн лет).\n"
        "Интерпретация: ксеногенные цирконы; это не прямой возраст внедрения силла.",
        color="#4f5c55", fontsize=9.2,
    )
    fig.savefig(OUTPUT / "kochnev2025_1812_1_summary.png", dpi=180, facecolor=fig.get_facecolor())
    plt.close(fig)

    # Portal view for an igneous body: concordia plus weighted mean, without KDE.
    fig, axes = plt.subplots(1, 2, figsize=(12.8, 5.8), facecolor="#f6f8f6")
    fig.subplots_adjust(left=0.075, right=0.98, top=0.84, bottom=0.15, wspace=0.24)
    fig.suptitle("Силл 1812-1 · concordia и weighted mean", x=0.06, ha="left", fontsize=17, fontweight="bold")
    fig.text(0.06, 0.90, "Kochnev et al., 2025 · ошибки и эллипсы 2σ · кластер n = 13", color="#5e6b64", fontsize=10)

    ax = axes[0]
    zoom_ages = np.linspace(780.0, 890.0, 500)
    ax.plot(
        np.expm1(LAMBDA_235 * zoom_ages * 1e6),
        np.expm1(LAMBDA_238 * zoom_ages * 1e6),
        color=curve,
        lw=1.8,
    )
    for a in cluster:
        add_dict_ellipse(ax, a, teal, .36)
    ax.set_xlim(1.20, 1.34)
    ax.set_ylim(0.130, 0.143)
    ax.set_title("A  Wetherill concordia · кластер", loc="left", fontweight="bold")
    ax.set_xlabel("²⁰⁷Pb/²³⁵U")
    ax.set_ylabel("²⁰⁶Pb/²³⁸U")
    ax.text(
        0.03,
        0.96,
        f"Concordia age = {stats['age']:.1f} ± {(stats['err_2s_minus'] + stats['err_2s_plus']) / 2:.1f} млн лет\n"
        f"MSWD concordance = {stats['concordance_mswd']:.2f}; p = {stats['concordance_probability']:.2f}",
        transform=ax.transAxes,
        va="top",
        bbox=dict(boxstyle="round,pad=.35", fc="white", ec="#cad2cd", alpha=.94),
    )

    ax = axes[1]
    ordered = sorted(cluster, key=lambda a: float(a["age_68"]))
    y = np.arange(len(ordered))
    values = np.array([float(a["age_68"]) for a in ordered])
    errors = np.array([float(a["err_age_68_2s"]) for a in ordered])
    mean, mean_2s = stats["wm_age_68"], stats["wm_err_2s"]
    ax.axvspan(mean - mean_2s, mean + mean_2s, color="#a9d5cf", alpha=.55)
    ax.axvline(mean, color="#a33d3d", lw=1.4, ls="--")
    ax.errorbar(values, y, xerr=errors, fmt="o", ms=4, color=teal, ecolor=teal, elinewidth=1, capsize=2)
    ax.set_yticks([])
    ax.set_xlim(795, 870)
    ax.set_title("B  ²⁰⁶Pb/²³⁸U weighted mean", loc="left", fontweight="bold")
    ax.set_xlabel("Возраст, млн лет (ошибки 2σ)")
    ax.text(
        0.98,
        0.05,
        f"{mean:.1f} ± {mean_2s:.1f} млн лет\nMSWD = {stats['wm_mswd']:.2f}; p = {stats['wm_probability']:.2f}; n = {len(cluster)}",
        transform=ax.transAxes,
        ha="right",
        va="bottom",
        bbox=dict(boxstyle="round,pad=.35", fc="white", ec="#cad2cd", alpha=.94),
    )
    for ax in axes:
        ax.set_facecolor("#fbfcfb")
        ax.grid(True, color="#dfe5e1", linewidth=.65)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
    fig.savefig(
        OUTPUT / "kochnev2025_1812_1_concordia_wm.png",
        dpi=180,
        facecolor=fig.get_facecolor(),
    )
    plt.close(fig)

    report = f"""# Силл 1812-1 — проверка U–Pb данных Kochnev et al. (2025)

## Результат пересчёта

- В таблице 15 определений по 11 зернам. Тринадцать определений формируют компактный кластер 812–872 млн лет по ²⁰⁷Pb/²⁰⁶Pb и 824–840 млн лет по ²⁰⁶Pb/²³⁸U.
- Независимый Wetherill concordia age по 13 определениям: **{stats['age']:.1f} ± {(stats['err_2s_minus'] + stats['err_2s_plus']) / 2:.1f} млн лет (2σ, без добавления неопределённости констант распада)**.
- Компонента concordance: MSWD = {stats['concordance_mswd']:.2f}, p = {stats['concordance_probability']:.3f}. Опубликованные 834.5 ± 4.2 млн лет, MSWD = 2.0 и p = 0.16 воспроизводятся в пределах округления и учёта констант распада.
- Обычное взвешенное среднее ²⁰⁶Pb/²³⁸U: **{stats['wm_age_68']:.1f} ± {stats['wm_err_2s']:.1f} млн лет (2σ)**, MSWD = {stats['wm_mswd']:.2f}, p = {stats['wm_probability']:.3f}.
- Остальные два анализа дают около 900 млн лет (²⁰⁶Pb/²³⁸U) и 1878 млн лет (²⁰⁷Pb/²⁰⁶Pb); они интерпретируются как ксенокристаллы.

## Геологическая интерпретация

Кластер 835 млн лет статистически реален, но не датирует внедрение силла. Для вмещающих песчаников губинской свиты опубликован MDA 716 ± 10 млн лет, то есть они более чем на 100 млн лет моложе. Полевые взаимоотношения не поддерживают более древнее интрузивное тело, поэтому Kochnev et al. (2025) считают наиболее вероятным захват циркона из более древнего магматического источника. Следовательно, по этим данным возраст самого силла остаётся неопределённым; **835 млн лет — возраст ксеногенной цирконовой популяции**.

Дополнительное ограничение: 13 анализов получены по 7 фрагментам кристаллов, поэтому это не 13 полностью независимых зерен. Для уточнения возраста внедрения нужны собственные магматические фазы/бадделеит либо текстурно доказанные новообразованные каймы циркона.

Источник: Kochnev B.B. et al., 2025, DOI {KOCHNEV_2025_DOI}. Исходный лист: `1812-1 Dolerite`.
"""
    (OUTPUT / "kochnev2025_1812_1_analysis.md").write_text(report, encoding="utf-8")
    return analyses, stats


def write_1812_1_sm25_07_comparison(
    kochnev: list[dict[str, object]],
    kochnev_stats: dict[str, float],
    sm25_07: list[Analysis],
) -> dict[str, float]:
    cluster = [a for a in kochnev if a["cluster_835"]]
    accepted = [a for a in sm25_07 if a.accepted]
    sm_mean, sm_2s, sm_mswd = weighted_mean(accepted, "age_68", "err_age_68")
    sm_corr_mean, sm_corr_2s, sm_corr_mswd = weighted_mean(
        accepted, "age_207corr", "err_age_207corr"
    )

    def sm_model(age_ma: float) -> np.ndarray:
        years = age_ma * 1e6
        return np.array([np.expm1(LAMBDA_235 * years), np.expm1(LAMBDA_238 * years)])

    def sm_covariance(a: Analysis) -> np.ndarray:
        return np.array(
            [
                [a.err_ratio_75**2, a.rho_3 * a.err_ratio_75 * a.err_ratio_68],
                [a.rho_3 * a.err_ratio_75 * a.err_ratio_68, a.err_ratio_68**2],
            ]
        )

    def sm_objective(age_ma: float) -> float:
        expected = sm_model(age_ma)
        total = 0.0
        for a in accepted:
            residual = np.array([a.ratio_75, a.ratio_68]) - expected
            total += float(residual @ np.linalg.inv(sm_covariance(a)) @ residual)
        return total

    sm_concordia_result = minimize_scalar(sm_objective, bounds=(760.0, 880.0), method="bounded")
    sm_concordia_age = float(sm_concordia_result.x)
    sm_concordia_chi2 = float(sm_concordia_result.fun)
    sm_concordia_lo = brentq(
        lambda value: sm_objective(value) - sm_concordia_chi2 - 4.0,
        720.0,
        sm_concordia_age,
    )
    sm_concordia_hi = brentq(
        lambda value: sm_objective(value) - sm_concordia_chi2 - 4.0,
        sm_concordia_age,
        920.0,
    )
    sm_concordia_2s = (sm_concordia_hi - sm_concordia_lo) / 2.0
    sm_concordia_df = 2 * len(accepted) - 1
    sm_concordia_combined_mswd = sm_concordia_chi2 / sm_concordia_df
    sm_concordia_probability = float(chi2.sf(sm_concordia_chi2, sm_concordia_df))

    kochnev_wm = kochnev_stats["wm_age_68"]
    kochnev_wm_2s = kochnev_stats["wm_err_2s"]
    difference = kochnev_wm - sm_mean
    difference_2s = math.sqrt(kochnev_wm_2s**2 + sm_2s**2)
    difference_z = 2.0 * difference / difference_2s

    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9})
    fig, axes = plt.subplots(2, 2, figsize=(12.4, 8.8), facecolor="#f6f8f6")
    fig.subplots_adjust(left=0.085, right=0.98, top=0.89, bottom=0.13, wspace=0.24, hspace=0.30)
    fig.suptitle("1812-1 и SM25-07 · сравнение U–Pb данных", x=0.085, ha="left", fontsize=18, fontweight="bold")
    fig.text(
        0.085,
        0.925,
        "Ксеногенный кластер силла и циркон из диоритовой дайки на спилитовом объекте",
        color="#5e6b64",
        fontsize=10,
    )
    teal, purple, curve = "#13756c", "#73589f", "#28322e"

    ax = axes[0, 0]
    zoom_ages = np.linspace(760.0, 900.0, 600)
    ax.plot(np.expm1(LAMBDA_235 * zoom_ages * 1e6), np.expm1(LAMBDA_238 * zoom_ages * 1e6), color=curve, lw=1.8)
    for a in cluster:
        add_dict_ellipse(ax, a, teal, .28)
    for a in accepted:
        add_error_ellipse(ax, a, purple, .14, 2)
    ax.scatter(
        [float(a["ratio_75"]) for a in cluster],
        [float(a["ratio_68"]) for a in cluster],
        s=14,
        color=teal,
        label="1812-1 · кластер n=13",
        zorder=4,
    )
    ax.scatter(
        [a.ratio_75 for a in accepted],
        [a.ratio_68 for a in accepted],
        s=12,
        facecolor="none",
        edgecolor=purple,
        linewidth=.8,
        label="SM25-07 · принято n=31",
        zorder=4,
    )
    ax.set_xlim(1.05, 1.55)
    ax.set_ylim(0.118, 0.151)
    ax.set_title("A  Wetherill concordia · одинаковый масштаб", loc="left", fontweight="bold")
    ax.set_xlabel("²⁰⁷Pb/²³⁵U")
    ax.set_ylabel("²⁰⁶Pb/²³⁸U")
    ax.legend(loc="lower right", frameon=True, fontsize=8)

    ax = axes[0, 1]
    estimates = [
        ("1812-1 · concordia", kochnev_stats["age"], (kochnev_stats["err_2s_minus"] + kochnev_stats["err_2s_plus"]) / 2, teal, "o"),
        ("1812-1 · WM ²⁰⁶Pb/²³⁸U", kochnev_wm, kochnev_wm_2s, teal, "s"),
        ("SM25-07 · WM ²⁰⁶Pb/²³⁸U", sm_mean, sm_2s, purple, "o"),
        ("SM25-07 · ²⁰⁷PbCorr", sm_corr_mean, sm_corr_2s, purple, "s"),
    ]
    for index, (label, age, error, color, marker) in enumerate(estimates[::-1]):
        ax.errorbar(age, index, xerr=error, fmt=marker, color=color, ecolor=color, ms=6, capsize=3)
    ax.set_yticks(range(len(estimates)))
    ax.set_yticklabels([item[0] for item in estimates[::-1]], fontsize=8)
    ax.set_xlim(805, 845)
    ax.set_ylim(-.7, len(estimates) - .3)
    ax.set_title("B  Возрастные оценки · ошибки 2σ", loc="left", fontweight="bold")
    ax.set_xlabel("Возраст, млн лет")
    ax.text(
        .98,
        .04,
        f"Разность WM: {difference:.1f} ± {difference_2s:.1f} млн лет (2σ)\n"
        f"{difference_z:.1f}σ по внутренним ошибкам",
        transform=ax.transAxes,
        ha="right",
        va="bottom",
        bbox=dict(boxstyle="round,pad=.35", fc="white", ec="#cad2cd", alpha=.94),
    )

    ax = axes[1, 0]
    xs = np.linspace(790, 855, 700)
    bandwidth = 5.0
    for ages, color, label in [
        ([float(a["age_68"]) for a in cluster], teal, "1812-1"),
        ([a.age_68 for a in accepted], purple, "SM25-07"),
    ]:
        density = np.mean(
            [np.exp(-.5 * ((xs - age) / bandwidth) ** 2) / (bandwidth * math.sqrt(2 * math.pi)) for age in ages],
            axis=0,
        )
        ax.plot(xs, density, color=color, lw=2.2, label=label)
        offset = -.001 if label == "1812-1" else -.0022
        ax.vlines(ages, offset, offset + .0008, color=color, lw=.7, alpha=.65)
    ax.axvline(kochnev_wm, color=teal, ls="--", lw=1)
    ax.axvline(sm_mean, color=purple, ls="--", lw=1)
    ax.set_ylim(-.003, None)
    ax.set_title("C  Распределение ²⁰⁶Pb/²³⁸U возрастов", loc="left", fontweight="bold")
    ax.set_xlabel("Возраст, млн лет · KDE h = 5 млн лет")
    ax.set_ylabel("Плотность")
    ax.legend(frameon=False)

    ax = axes[1, 1]
    ax.scatter(
        [float(a["u_ppm"]) for a in cluster],
        [float(a["u_th"]) for a in cluster],
        s=38,
        color=teal,
        label="1812-1",
        edgecolor="white",
        linewidth=.5,
    )
    ax.scatter(
        [a.u for a in accepted],
        [1.0 / a.thu for a in accepted],
        s=30,
        facecolor="none",
        edgecolor=purple,
        linewidth=.9,
        label="SM25-07",
    )
    ax.set_xscale("log")
    ax.set_xlim(70, 3200)
    ax.set_ylim(.45, 3.35)
    ax.set_title("D  Состав циркона", loc="left", fontweight="bold")
    ax.set_xlabel("U, ppm (логарифмическая шкала)")
    ax.set_ylabel("U/Th")
    ax.legend(frameon=False)
    ax.text(
        .03,
        .96,
        "Медианы:\n1812-1: U ≈ 1203 ppm, U/Th ≈ 0,96\nSM25-07: U ≈ 201 ppm, U/Th ≈ 1,76",
        transform=ax.transAxes,
        va="top",
        bbox=dict(boxstyle="round,pad=.35", fc="white", ec="#cad2cd", alpha=.94),
    )

    for ax in axes.flat:
        ax.set_facecolor("#fbfcfb")
        ax.grid(True, color="#dfe5e1", linewidth=.65)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
    fig.text(
        0.085,
        0.025,
        "Близость принудительных concordia ages обманчива: SM25-07 не образует конкордантную совокупность (combined MSWD ≈ "
        f"{sm_concordia_combined_mswd:.2f}).\n"
        "1812-1 содержит ксеногенный кластер; совпадение возрастов не доказывает одновременное внедрение двух тел.",
        color="#4f5c55",
        fontsize=9.2,
    )
    fig.savefig(OUTPUT / "1812_1_SM25_07_comparison.png", dpi=180, facecolor=fig.get_facecolor())
    plt.close(fig)

    report = f"""# Сравнение 1812-1 и SM25-07

## Что сравнивается

- **1812-1:** 13 определений ксеногенной цирконовой популяции. Concordia age {kochnev_stats['age']:.1f} ± {(kochnev_stats['err_2s_minus'] + kochnev_stats['err_2s_plus']) / 2:.1f} млн лет; взвешенное среднее ²⁰⁶Pb/²³⁸U {kochnev_wm:.1f} ± {kochnev_wm_2s:.1f} млн лет.
- **SM25-07:** 31 из 50 анализов проходят фильтр |D1| ≤ 10%. Взвешенное среднее ²⁰⁶Pb/²³⁸U {sm_mean:.1f} ± {sm_2s:.1f} млн лет (MSWD {sm_mswd:.2f}); ²⁰⁷PbCorr даёт {sm_corr_mean:.1f} ± {sm_corr_2s:.1f} млн лет (MSWD {sm_corr_mswd:.2f}).

## Проверка сходства

Взвешенные средние ²⁰⁶Pb/²³⁸U различаются на **{difference:.1f} ± {difference_2s:.1f} млн лет (2σ)**, что соответствует {difference_z:.1f}σ по внутренним аналитическим ошибкам. Без систематической неопределённости это статистически различимые оценки.

Принудительная подгонка всех 31 принятых анализов SM25-07 к одной точке конкордии даёт {sm_concordia_age:.1f} ± {sm_concordia_2s:.1f} млн лет, почти совпадая с 1812-1. Но эта оценка **невалидна**: combined MSWD = {sm_concordia_combined_mswd:.2f}, p = {sm_concordia_probability:.2e}. Совпадение возникает из геометрии рассеянного массива, а не из общей конкордантной популяции.

Цирконы также различаются по составу. Для кластера 1812-1 медиана U около 1203 ppm и U/Th около 0,96; для принятой выборки SM25-07 — около 201 ppm и 1,76 соответственно. Это не доказывает разные источники само по себе, но не поддерживает простое отождествление популяций.

## Вывод

По имеющимся данным тела нельзя уверенно связывать с одним импульсом внедрения. Наиболее защищённая оценка для SM25-07 находится около 819–821 млн лет. Кластер 835 млн лет в 1812-1 статистически реален, но интерпретируется как ксеногенный и потому не задаёт возраст силла. Даже полное совпадение этих возрастов не было бы доказательством синхронности тел.

Для проверки генетической связи полезнее сопоставить цельнопородную геохимию и Sm–Nd, а для 1812-1 получить прямую датировку новообразованной фазы (бадделеит, текстурно доказанные каймы циркона и т.п.).
"""
    (OUTPUT / "1812_1_SM25_07_comparison.md").write_text(report, encoding="utf-8")
    return {
        "sm_mean": sm_mean,
        "sm_2s": sm_2s,
        "difference": difference,
        "difference_2s": difference_2s,
        "sm_forced_concordia": sm_concordia_age,
        "sm_forced_concordia_2s": sm_concordia_2s,
        "sm_forced_concordia_mswd": sm_concordia_combined_mswd,
        "sm_forced_concordia_p": sm_concordia_probability,
    }


def write_geochronology_data(samples: dict[str, list[Analysis]]) -> None:
    datasets: list[dict[str, object]] = []
    labels = {
        "SM25-03": ("SM25-03", "базитовый силл", "Базитовый силл · точка SM25-03"),
        "SM25-05": ("SM25-05", "губинская свита", "Губинская свита · точка SM25-05"),
        "SM25-07": ("SM25-07", "миндалекаменные трахибазальты", "Спилитовый объект · точка SM25-07"),
        "SM25-16": ("SM25-16", "обломочный циркон", "Туфопесчаники · точка SM25-16"),
    }
    for sample, analyses in samples.items():
        label, subtitle, object_label = labels[sample]
        datasets.append(
            {
                "id": sample,
                "label": label,
                "subtitle": subtitle,
                "object": object_label,
                "source": "Малышев, Пасенко",
                "source_kind": "project",
                "kde_enabled": sample in {"SM25-05", "SM25-16"},
                "analyses": [
                    {
                        "sid": a.sid,
                        "age": a.best_age,
                        "err_1s": a.best_err,
                        "accepted": a.accepted,
                    }
                    for a in analyses
                ],
            }
        )

    kochnev_41812 = load_kochnev_sample("41812")
    gubinskaya = next(item for item in datasets if item["id"] == "SM25-05")
    gubinskaya.update(
        {
            "id": "GUBINSKAYA-COMPARISON",
            "label": "Губинская свита",
            "subtitle": "сопоставление двух выборок",
            "object": "Губинская свита · SM25-05 и обнажение 1812",
            "published_result": "максимальный возраст осадконакопления 716 ± 10 млн лет (n=3, concordia age)",
            "comparison": {
                "id": "KOCHNEV-41812",
                "label": "Кочнев · 1812-2",
                "source": "Kochnev et al., 2022",
                "doi": KOCHNEV_DOI,
                "analyses": kochnev_41812,
            },
        }
    )
    kochnev_1812_1, sill_stats = write_kochnev_2025_1812_1()
    datasets.append(
        {
            "id": "KOCHNEV-1812-1-SILL",
            "label": "1812-1",
            "subtitle": "долеритовый силл в поле губинской свиты",
            "object": "Тело долеритов · обнажение 1812",
            "source": "Kochnev et al., 2025",
            "source_kind": "literature",
            "kde_enabled": False,
            "doi": KOCHNEV_2025_DOI,
            "published_result": "concordia age 834,5 ± 4,2 млн лет (n=13); авторы интерпретируют цирконы как ксеногенные",
            "interpretation": (
                f"Пересчёт: {sill_stats['age']:.1f} ± "
                f"{(sill_stats['err_2s_minus'] + sill_stats['err_2s_plus']) / 2:.1f} млн лет (2σ). "
                "Кластер не является прямым возрастом внедрения силла."
            ),
            "figure": "upb/kochnev2025_1812_1_summary.png",
            "audit": "upb/kochnev2025_1812_1_audit.csv",
            "report": "upb/kochnev2025_1812_1_analysis.md",
            "upb_analyses": [
                {
                    "sid": a["sid"],
                    "ratio_75": a["ratio_75"],
                    "err_ratio_75": a["err_ratio_75_2s"] / 2.0,
                    "ratio_68": a["ratio_68"],
                    "err_ratio_68": a["err_ratio_68_2s"] / 2.0,
                    "rho_3": a["rho"],
                    "age_76": a["age_76"],
                    "err_age_76": a["err_age_76_2s"] / 2.0,
                    "age_68": a["age_68"],
                    "err_age_68": a["err_age_68_2s"] / 2.0,
                    "age_75": a["age_75"],
                    "err_age_75": a["err_age_75_2s"] / 2.0,
                    "context_excluded": not a["cluster_835"],
                }
                for a in kochnev_1812_1
            ],
            "analyses": [
                {
                    "sid": a["sid"],
                    "age": a["age_68"] if a["age_68"] < SWITCH_AGE_MA else a["age_76"],
                    "err_1s": (a["err_age_68_2s"] if a["age_68"] < SWITCH_AGE_MA else a["err_age_76_2s"]) / 2.0,
                    "accepted": abs(a["d1"] if a["age_68"] < SWITCH_AGE_MA else a["d2"]) <= DISCORDANCE_LIMIT,
                }
                for a in kochnev_1812_1
            ],
        }
    )
    kochnev_2007_15 = load_kochnev_2025_2007_15()
    datasets.append(
        {
            "id": "KOCHNEV-2007-15-LUDY",
            "label": "2007-15",
            "subtitle": "песчаники лудовской свиты",
            "object": "Лудовская свита · типовой разрез на левом берегу Енисея",
            "source": "Kochnev et al., 2025",
            "source_kind": "literature",
            "kde_enabled": True,
            "doi": KOCHNEV_2025_DOI,
            "published_result": (
                "121 определение; самое молодое 1551 ± 27 млн лет; основные пики около "
                "2,76, 2,57, 2,05, 1,93 и 1,87 млрд лет"
            ),
            "interpretation": (
                "Все 121 определения проходят гибридный порог |D| ≤ 10%; для выборки "
                "использован возраст ²⁰⁷Pb/²⁰⁶Pb, поскольку все анализы старше 1 млрд лет."
            ),
            "audit": "upb/kochnev2025_2007_15_audit.csv",
            "analyses": kochnev_2007_15,
        }
    )
    comparison_stats = write_1812_1_sm25_07_comparison(
        kochnev_1812_1, sill_stats, samples["SM25-07"]
    )
    datasets.append(
        {
            "id": "1812-1-SM25-07-COMPARISON",
            "label": "Два базитовых тела",
            "subtitle": "1812-1 и SM25-07",
            "object": "Силл в поле губинской свиты и диоритовая дайка на спилитовом объекте",
            "source": "Малышев, Пасенко · SM25-07",
            "source_kind": "comparison",
            "kde_enabled": False,
            "interpretation": (
                f"WM ²⁰⁶Pb/²³⁸U различаются на {comparison_stats['difference']:.1f} ± "
                f"{comparison_stats['difference_2s']:.1f} млн лет (2σ, внутренние ошибки). "
                "Кластер 1812-1 ксеногенный, поэтому его нельзя использовать как возраст внедрения силла."
            ),
            "figure": "upb/1812_1_SM25_07_comparison.png",
            "report": "upb/1812_1_SM25_07_comparison.md",
            "analyses": [
                {
                    "sid": a.sid,
                    "age": a.age_68,
                    "err_1s": a.err_age_68,
                    "accepted": a.accepted,
                }
                for a in samples["SM25-07"]
            ],
            "comparison": {
                "id": "KOCHNEV-1812-1-CLUSTER",
                "label": "Kochnev et al., 2025 · 1812-1",
                "source": "Kochnev et al., 2025",
                "doi": KOCHNEV_2025_DOI,
                "analyses": [
                    {
                        "sid": a["sid"],
                        "age": a["age_68"],
                        "err_1s": a["err_age_68_2s"] / 2.0,
                        "accepted": True,
                    }
                    for a in kochnev_1812_1
                    if a["cluster_835"]
                ],
            },
        }
    )
    payload = {
        "datasets": datasets,
        "literature_cards": {
            "SM25-05": {
                "title": "Губинская свита · Kochnev et al., 2022",
                "sample": "1812-2",
                "analytical_prefix": "41812",
                "summary": "Опубликованный MDA: 716 ± 10 млн лет; три молодых зерна, concordia age.",
                "detail": "В приложении 130 анализов; 120 проходят авторский порог |D| ≤ 10%.",
                "dataset_id": "GUBINSKAYA-COMPARISON",
                "doi": KOCHNEV_DOI,
                "audit": "upb/kochnev2022_41812_audit.csv",
            },
        },
    }
    with (OUTPUT / "kochnev2022_41812_audit.csv").open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["sid", "age", "err_1s", "err_2s", "discordance", "accepted", "flagged_by_author"])
        writer.writeheader()
        writer.writerows(kochnev_41812)
    with (OUTPUT / "kochnev2025_2007_15_audit.csv").open(
        "w", encoding="utf-8-sig", newline=""
    ) as handle:
        writer = csv.DictWriter(handle, fieldnames=list(kochnev_2007_15[0].keys()))
        writer.writeheader()
        writer.writerows(kochnev_2007_15)
    (OUTPUT / "geochronology_data.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def main() -> None:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    samples = {sample: load_sample(sample, path) for sample, path in SOURCES.items()}
    for sample, analyses in samples.items():
        classify(sample, analyses)
        write_audit(sample, analyses)
    results = {
        "SM25-03": plot_new_spectrum(
            "SM25-03",
            samples["SM25-03"],
            "igneous_complex",
            "U–Pb LA-ICP-MS, циркон из базитового силла",
            "Единая магматическая популяция не выделяется; датировка силла требует CL/BSE-контекста спотов и проверки common Pb.",
        ),
        "SM25-05": plot_new_spectrum(
            "SM25-05",
            samples["SM25-05"],
            "detrital",
            "U–Pb LA-ICP-MS, обломочный циркон · губинская свита",
            "Молодой кластер — предварительный MDA; его нужно сопоставить с результатами Кочнева и проверить при альтернативных метриках дискордантности.",
        ),
        "SM25-07": plot_sm25_07(samples["SM25-07"]),
        "SM25-16": plot_sm25_16(samples["SM25-16"]),
    }
    write_report(results)
    write_portal_data(results, samples)
    write_geochronology_data(samples)
    for sample, result in results.items():
        print(sample, result["summary"])


if __name__ == "__main__":
    main()
