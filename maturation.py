"""
maturation.py — WHEN the marks are set: the maturation curve of every mark whose extract dates it.

A mark (softcancel, dormant…) is not there from the start: a licence that falls due in twelve months may be
marked next week, or in the month it falls due. The extract keeps the LAST state of a mark, so the closed months
show every mark already set, and the future months only the part set so far. With the maturation columns
(<mark>_month_k: the units of the row marked k months before due, k = 0 … maturation_months, the last one "that
many months or more") the history says how a cohort fills its marks, and that is predicted like the rate and the
uplift: the third prediction of the forecast.

THE CURVE (per mark, learned from closed cohorts, the marks final)
  For every group and due month: U the units due, M the units marked at the end, V_h the units already marked h
  months before due (Σ of the columns k ≥ h). Then, summed over the cohorts:
      visible share   V_h / M                    how much of the final mark is already seen h months before
      p(h)            Σ (M − V_h) / Σ (U − V_h)  the probability that a unit NOT marked h months before due has
                                                 the mark when it falls due
  A unit not marked today, h months before its due month, ends marked with probability p(h). The expected final
  units of the mark are   V + (U − V) × p(h)   (V: marked today; the marks only grow).
  The groups are those of step 17 (the forecast series without its marks), its mandatory cell and the
  portfolio: the first one with at least maturation_min_units units not marked h months before.

WHERE IT IS USED
  step 17  the final share of a mark with maturation columns comes from the curve (today's share + the expected
           migration), not from the historical share of the calendar month; the split among the combinations
           that carry it stays the historical one
  step 19  the honest exam: the marks of every exam month as they were h months before (the columns k ≥ h),
           the curve learned with the cohorts known at the origin, the rates of the framework. The exam with the
           final marks stays as the CEILING (what predicting the marks perfectly would give)

h, THE MONTHS BEFORE DUE THAT ARE ALREADY KNOWN
  In the forecast: h = due month − last closed month (the current month: h = 1). In the exam: h = the horizon
  (a technique predicting T at horizon h sees the months up to T − h). A unit is visible iff it was marked at
  least h months before due: Σ of the columns k ≥ h.
"""

# ─── imports ─────────────────────────────────────────────────────────────────────
import numpy as np
import pandas as pd

from config import ACTIVE_FLAG_VALUES, Config, join_columns
from vocabulario import CALENDAR_ROLE_COLUMN, TRUTH_ROLES

# ─── named constants ─────────────────────────────────────────────────────────────
GROUP_LEVEL, CELL_LEVEL, PORTFOLIO_LEVEL = "_grupo", "_celda", "_cartera"
LEVELS = (GROUP_LEVEL, CELL_LEVEL, PORTFOLIO_LEVEL)
PORTFOLIO_KEY = "cartera"


def group_dims(configuration: Config) -> list:
    """The group of the curve: every dim of the forecast series but the timevarying ones (as step 17)."""
    return [column_name for column_name in configuration.rate_series_columns
            if column_name not in configuration.structural_timevarying_dims]


def add_level_keys(rows: pd.DataFrame, configuration: Config) -> pd.DataFrame:
    """The keys of the three levels of the curve on every row."""
    keyed = rows.copy()
    keyed[GROUP_LEVEL] = join_columns(keyed, group_dims(configuration))
    keyed[CELL_LEVEL] = join_columns(keyed, configuration.business_mandatory_dims)
    keyed[PORTFOLIO_LEVEL] = PORTFOLIO_KEY
    return keyed


def is_marked(rows: pd.DataFrame, mark: str) -> np.ndarray:
    return rows[mark].isin(ACTIVE_FLAG_VALUES).to_numpy() if mark in rows.columns else np.zeros(len(rows), dtype=bool)


def visible_units(rows: pd.DataFrame, mark: str, months_before, configuration: Config) -> np.ndarray:
    """The units of every row already marked `months_before` months before due (one value, or one per row):
    Σ of its columns k ≥ months_before; 0 on a row without the mark. Past the last column, only the last one
    (it means that many months or more)."""
    columns = configuration.maturation_columns_of(mark)
    # rows without the columns (the projected and simulated rows of the extended horizon) have nothing dated: 0
    timing = rows.reindex(columns=columns).fillna(0.0).to_numpy(dtype=float)
    seen_from_the_end = np.cumsum(timing[:, ::-1], axis=1)[:, ::-1]          # [:, k] = Σ of the columns ≥ k
    first = np.clip(np.broadcast_to(np.asarray(months_before, dtype=int), (len(rows),)), 0, len(columns) - 1)
    visible = seen_from_the_end[np.arange(len(rows)), first] if len(rows) else np.zeros(0)
    return np.where(is_marked(rows, mark), visible, 0.0)


class MaturationCurve:
    """The curve of one mark, learned from closed cohorts (their marks final): for every level key and h the
    sums Σ (M − V_h) and Σ (U − V_h); p(h) is their ratio at the first level with enough support."""

    def __init__(self, closed_rows: pd.DataFrame, mark: str, configuration: Config):
        self.mark = mark
        self.configuration = configuration
        self.months = int(configuration.maturation_months)
        self.minimum = configuration.maturation_min_units
        units_column = configuration.pipeline_units_col
        rows = add_level_keys(closed_rows, configuration)
        marked = is_marked(rows, mark)
        columns = configuration.maturation_columns_of(mark)
        timing = rows[columns].fillna(0.0).to_numpy(dtype=float) * marked[:, None]
        seen_from_the_end = np.cumsum(timing[:, ::-1], axis=1)[:, ::-1]        # V_h of every row, h = 0 … months
        units = rows[units_column].to_numpy(dtype=float)
        marked_units = units * marked
        # per row: still to be marked after h (M − V_h) and not marked h months before (U − V_h)
        to_come = marked_units[:, None] - seen_from_the_end
        not_marked_yet = units[:, None] - seen_from_the_end
        self.sums = {}
        for level in LEVELS:
            keys = rows[level].to_numpy()
            frame = pd.DataFrame(np.hstack([to_come, not_marked_yet, marked_units[:, None], seen_from_the_end]))
            frame["_key"] = keys
            grouped = frame.groupby("_key").sum()
            width = self.months + 1
            self.sums[level] = {key: (values[:width], values[width:2 * width], values[2 * width], values[2 * width + 1:])
                                for key, values in zip(grouped.index, grouped.to_numpy())}

    def probability(self, keys: dict, months_before: int) -> tuple:
        """(p, level): the probability that a unit not marked `months_before` months before due ends marked, at
        the first level with at least maturation_min_units units not marked then (the portfolio otherwise)."""
        h = int(np.clip(months_before, 0, self.months))
        for level in LEVELS:
            sums = self.sums[level].get(keys.get(level, PORTFOLIO_KEY))
            if sums is None:
                continue
            to_come, not_marked_yet = sums[0][h], sums[1][h]
            if not_marked_yet >= self.minimum or (level == PORTFOLIO_LEVEL and not_marked_yet > 0):
                return float(max(to_come, 0.0) / not_marked_yet), level
        return 0.0, PORTFOLIO_LEVEL

    def portfolio_table(self) -> pd.DataFrame:
        """The curve of the whole portfolio: h, the share of the final mark already seen, p(h)."""
        to_come, not_marked_yet, marked_units, seen = self.sums[PORTFOLIO_LEVEL][PORTFOLIO_KEY]
        return pd.DataFrame({"meses_antes": np.arange(self.months + 1),
                             "visible_del_final": seen / marked_units if marked_units > 0 else np.nan,
                             "p_marcarse": np.where(not_marked_yet > 0, to_come / np.maximum(not_marked_yet, 1e-9), np.nan)})


def closed_rows_of(fine_table: pd.DataFrame, configuration: Config, up_to=None) -> pd.DataFrame:
    """The closed rows (truth months) with something due, optionally only up to a month (inclusive)."""
    rows = fine_table[fine_table[CALENDAR_ROLE_COLUMN].isin(TRUTH_ROLES) & (fine_table[configuration.pipeline_units_col] > 0)]
    if up_to is not None:
        rows = rows[rows[configuration.period_col] <= up_to]
    return rows


def curve_history(fine_table: pd.DataFrame, configuration: Config, up_to) -> pd.DataFrame:
    """The cohorts a curve is learned from at an origin: the last maturation_history_months closed months up to it
    (inclusive), the same window the forecast learns its curve with (step 17), so the exam examines that curve."""
    rows = closed_rows_of(fine_table, configuration, up_to=up_to)
    months = sorted(rows[configuration.period_col].unique())[-configuration.maturation_history_months:]
    return rows[rows[configuration.period_col].isin(months)]


def examine_the_curve(fine_table: pd.DataFrame, mark: str, exam_months: list, horizons: list,
                      configuration: Config) -> pd.DataFrame:
    """The exam of the curve on the closed months: for every exam month T and horizon h, the units of the mark
    predicted with what was seen h months before (V_h + (U − V_h) × p(h), the curve learned with the cohorts up to
    T − h) against the units marked at the end. Also what was seen then (V_h, the marks of "today")."""
    period, units_column = configuration.period_col, configuration.pipeline_units_col
    closed = closed_rows_of(fine_table, configuration)
    rows = []
    for target_month in exam_months:
        month_rows = add_level_keys(closed[closed[period] == target_month], configuration)
        if month_rows.empty:
            continue
        for horizon in horizons:
            curve = MaturationCurve(curve_history(fine_table, configuration, target_month - horizon), mark, configuration)
            grouped = month_rows.assign(_marcadas=month_rows[units_column] * is_marked(month_rows, mark),
                                        _visibles=visible_units(month_rows, mark, horizon, configuration))
            grouped = grouped.groupby([GROUP_LEVEL, CELL_LEVEL])[[units_column, "_marcadas", "_visibles"]].sum().reset_index()
            probability = np.array([curve.probability({GROUP_LEVEL: group, CELL_LEVEL: cell}, horizon)[0]
                                    for group, cell in zip(grouped[GROUP_LEVEL], grouped[CELL_LEVEL])])
            predicted = grouped["_visibles"] + (grouped[units_column] - grouped["_visibles"]) * probability
            total = grouped[units_column].sum()
            rows.append({"mes": target_month, "h": horizon, "vencen": total,
                         "cuota_final": grouped["_marcadas"].sum() / total,
                         "cuota_vista_entonces": grouped["_visibles"].sum() / total,
                         "cuota_curva": predicted.sum() / total})
    exam = pd.DataFrame(rows)
    if len(exam):
        exam["error_curva_pp"] = 100 * (exam["cuota_curva"] - exam["cuota_final"])
        exam["error_vista_pp"] = 100 * (exam["cuota_vista_entonces"] - exam["cuota_final"])
    return exam


def honest_total_exam(framework_rows: pd.DataFrame, fine_table: pd.DataFrame, configuration: Config) -> pd.DataFrame:
    """The honest exam of the total: every exam month T and horizon h predicted with the marks as they were h
    months before, and their maturation from the curve known at the origin.

    INPUT:   the framework predictions of step 19 (one row per forecast series × exam month × h: fs_id, pred_rate,
             due_units, real_units), the fine table (the maturation columns, the dims of every series).
    OUTPUT:  one row per exam month × h: the real renewals, and the predicted ones three ways
               techo_marcas_finales  with the FINAL marks, as step 19 has always predicted: the ceiling, what
                                     predicting the marks perfectly would give
               marcas_hoy            with the marks seen at the origin and nothing more (no maturation)
               honesto               with the marks seen at the origin plus the migration the curve expects
    RULES:   for every mark with maturation columns, its units marked fewer than h months before due go back to
             the twin series without the mark (where they were at the origin); then a share p(h) of the units not
             marked moves to the marked twin. Every series renews at the rate the framework predicted for it; a
             twin with no prediction takes the rate of the other one times (or over) the ratio of the two rates in
             the portfolio history known at the origin. With several marks, their effects are added (each one
             measured against the ceiling).
    """
    period = configuration.period_col
    marks = [mark for mark in configuration.maturation_dims if mark in fine_table.columns]
    dims = [column for column in configuration.rate_series_columns if column in fine_table.columns]
    attributes = fine_table.drop_duplicates("fs_id").set_index("fs_id")[dims]
    timing_columns = [column for mark in marks for column in configuration.maturation_columns_of(mark)]
    timing = (fine_table.groupby(["fs_id", period])[timing_columns].sum(min_count=1).reset_index()
              if timing_columns else None)
    results = []
    for (target_month, horizon), rows in framework_rows.groupby([period, "h"]):
        origin = target_month - horizon
        known = curve_history(fine_table, configuration, origin)
        frame = rows[["fs_id", "pred_rate", "due_units", "real_units"]].join(attributes, on="fs_id")
        if timing is not None:
            frame = frame.merge(timing[timing[period] == target_month].drop(columns=[period]), on="fs_id", how="left")
        ceiling = float((frame["pred_rate"] * frame["due_units"]).sum())
        today, honest = ceiling, ceiling
        for mark in marks:
            curve = MaturationCurve(known, mark, configuration)
            ratio = marked_to_unmarked_ratio(known, mark, configuration)
            today += predicted_at_the_origin(frame, mark, horizon, curve, ratio, configuration, migrate=False) - ceiling
            honest += predicted_at_the_origin(frame, mark, horizon, curve, ratio, configuration, migrate=True) - ceiling
        results.append({"mes": target_month, "h": horizon, "renovadas_reales": float(frame["real_units"].sum()),
                        "techo_marcas_finales_pred": ceiling, "marcas_hoy_pred": today, "honesto_pred": honest})
    exam = pd.DataFrame(results)
    for method in ("techo_marcas_finales", "marcas_hoy", "honesto"):
        exam[f"{method}_error_total"] = exam[f"{method}_pred"] / exam["renovadas_reales"] - 1
    return exam


def marked_to_unmarked_ratio(known: pd.DataFrame, mark: str, configuration: Config) -> float:
    """The rate of the marked units over the rate of the unmarked ones, in the closed months known."""
    marked = is_marked(known, mark)
    due = known[configuration.pipeline_units_col].to_numpy(float)
    renewed = known[configuration.renewed_units_col].fillna(0).to_numpy(float)
    rate_marked = renewed[marked].sum() / max(due[marked].sum(), 1e-9)
    rate_unmarked = renewed[~marked].sum() / max(due[~marked].sum(), 1e-9)
    return float(rate_marked / rate_unmarked) if rate_unmarked > 0 and rate_marked > 0 else 1.0


def predicted_at_the_origin(frame: pd.DataFrame, mark: str, horizon: int, curve: MaturationCurve, ratio: float,
                            configuration: Config, migrate: bool) -> float:
    """The renewals predicted for one exam month with ONE mark as it was `horizon` months before due.

    Every twin pair (the same series with and without the mark): the marked side keeps only the units marked at
    least `horizon` months before (V); the rest (hidden) goes back to the side without the mark; with migrate, a
    share p(horizon) of the side without the mark moves to the marked one. Each side renews at its framework rate."""
    marked = is_marked(frame, mark)
    units = frame["due_units"].to_numpy(float)
    visible = np.minimum(visible_units(frame, mark, horizon, configuration), units)
    work = pd.DataFrame({"_gemela": join_columns(frame, [column for column in configuration.rate_series_columns
                                                         if column in frame.columns and column != mark]),
                         "_visible": np.where(marked, visible, 0.0),
                         "_sin": np.where(marked, units - visible, units),
                         "_marcada": marked, "_tasa": frame["pred_rate"].to_numpy(float)}, index=frame.index)
    pairs = work.groupby("_gemela")[["_visible", "_sin"]].sum()
    pairs["tasa_marcada"] = work[work["_marcada"]].groupby("_gemela")["_tasa"].mean().reindex(pairs.index)
    pairs["tasa_sin"] = work[~work["_marcada"]].groupby("_gemela")["_tasa"].mean().reindex(pairs.index)
    pairs["tasa_marcada"] = pairs["tasa_marcada"].fillna(pairs["tasa_sin"] * ratio)
    pairs["tasa_sin"] = pairs["tasa_sin"].fillna(pairs["tasa_marcada"] / ratio if ratio > 0 else pairs["tasa_marcada"])
    if migrate:
        keys = add_level_keys(frame.assign(_gemela=work["_gemela"]).drop_duplicates("_gemela"), configuration).set_index("_gemela")
        probability = np.array([curve.probability({GROUP_LEVEL: keys.at[twin, GROUP_LEVEL], CELL_LEVEL: keys.at[twin, CELL_LEVEL]},
                                                  horizon)[0] for twin in pairs.index])
        moved = pairs["_sin"].to_numpy() * probability
    else:
        moved = np.zeros(len(pairs))
    marked_units = pairs["_visible"].to_numpy() + moved
    unmarked_units = pairs["_sin"].to_numpy() - moved
    return float(np.nansum(marked_units * pairs["tasa_marcada"].to_numpy() + unmarked_units * pairs["tasa_sin"].to_numpy()))
