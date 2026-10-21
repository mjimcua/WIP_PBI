"""
step_03_fine_table.py — The fine table: every row of the raw, with the ids that name it.

The forecast of a future row is: what falls due × renewal rate × uplift. Each row is
therefore named by the two contexts it takes its numbers from:
  · RATE SERIES (fs_id): its mandatory, timevarying and extra_renovacion values. Every
    distinct combination is one series with its own renewal rate over time.
  · FORECAST UNIT (fu_id): its series in its month (fs_id | month). The rate is
    predicted per unit.
  · UPLIFT CELL (uplift_cell_id): its uplift mandatory dims, its extra_revalorizacion
    values and its discount bucket (region, product, 60-70 %...). The price at which it
    renews is estimated per cell, over every month of the cell.
The discount bucket is DERIVED here from the exact discount with the edges of the Config
(a null discount goes to "sin_dato"). A fine row is its unit, its extra_revalorizacion
values and its exact discount (the extract comes one row per exact discount): the grain is
checked on them. Nothing is aggregated here and no measure changes; step 04 adds the fine
rows of each unit. The ids are text: the core and the BI join on them (no hash keys).

Actions (logged as they are done):
  1. build fs_id from the rate series columns
  2. build fu_id: fs_id | month
  3. derive the discount bucket from the exact discount
  4. build uplift_cell_id from the uplift cell columns
  5. check the bucket, the grain and that nothing else changed            checks 1-3
  6. keep the fine table in memory (sff_nucleo writes every fine row)
  7. count the checks; stop if any failed
  8. show the rows per discount bucket and a few rows with their ids, as tables

Checks (logged as they are made, numbered, at the level of their status):
   1. every exact discount falls in one bucket (a null one in "sin_dato")
   2. one row per forecast unit, revaluation values and exact discount (the grain of the raw)
   3. same rows, same order, no column of step 02 changed

Output: the raw of step 02 with the discount bucket and three id columns (fs_id, fu_id, uplift_cell_id);
not written to SQL (sff_nucleo carries every fine row).
"""

# ─── imports ─────────────────────────────────────────────────────────────────────
import pandas as pd

from config import (COMBINED_ID_SEPARATOR, ID_FIELD_SEPARATOR, NO_REVALUATION_VALUES, UNKNOWN_DISCOUNT_BUCKET, Config,
                    discount_bucket_labels, join_columns)
from vocabulario import SERIES_ID_COLUMN, UNIT_ID_COLUMN, UPLIFT_CELL_ID_COLUMN


# ─── the step ────────────────────────────────────────────────────────────────────
STEP_LABEL = "03"
STEP_NAME = "FINE TABLE"
STEP_PURPOSE = ("name every row of the raw by the contexts it takes its numbers from: its rate series (fs_id), its "
                "forecast unit (series in its month, fu_id) and its uplift cell (price context, uplift_cell_id); "
                "nothing is aggregated")
STEP_ACTIONS = ["build fs_id from the rate series columns",
                "build fu_id: fs_id | month",
                "derive the discount bucket from the exact discount",
                "build uplift_cell_id from the uplift cell columns",
                "check the bucket, the grain and that nothing else changed (checks 1-3)",
                "keep the fine table in memory (sff_nucleo writes every fine row)",
                "count the checks; stop if any failed",
                "show the rows per discount bucket and a few rows with their ids, as tables"]
STEP_OUTPUT = ("the raw with the discount bucket and three id columns (one row per raw row; in memory, "
               "sff_nucleo writes it)")

# ─── named constants ─────────────────────────────────────────────────────────────
# The id columns this step adds, and the temporary name of a fine row inside the grain check.
ADDED_ID_COLUMNS = (SERIES_ID_COLUMN, UNIT_ID_COLUMN, UPLIFT_CELL_ID_COLUMN)
FINE_ROW_GROUP = "_fila"


def build_fine_table(calendared: pd.DataFrame, configuration: Config) -> pd.DataFrame:
    """Add the ids that name every row; check the grain."""
    configuration.log_step_start(STEP_LABEL, STEP_NAME, STEP_PURPOSE, STEP_ACTIONS, STEP_OUTPUT)
    check_log = []
    fine_table = calendared.copy()

    # [1] the rate series of every row
    fine_table[SERIES_ID_COLUMN] = join_columns(fine_table, configuration.rate_series_columns)
    configuration.log_action(STEP_LABEL, 1, f"fs_id from {len(configuration.rate_series_columns)} columns "
                                            f"({', '.join(configuration.rate_series_columns)}): "
                                            f"{fine_table[SERIES_ID_COLUMN].nunique():,} series")

    # [2] the forecast unit: the series in its month
    fine_table[UNIT_ID_COLUMN] = (fine_table[SERIES_ID_COLUMN] + ID_FIELD_SEPARATOR
                                  + fine_table[configuration.period_col].astype(str))
    configuration.log_action(STEP_LABEL, 2, f"fu_id = fs_id | month: {fine_table[UNIT_ID_COLUMN].nunique():,} forecast units")

    # [3] the discount bucket, from the exact discount (the Config's edges; null → sin_dato)
    if configuration.discount_value_column:
        fine_table[configuration.discount_bucket_column] = discount_bucket_labels(
            fine_table[configuration.discount_value_column], configuration.discount_bucket_edges)
        unknown_share = (fine_table[configuration.discount_bucket_column] == UNKNOWN_DISCOUNT_BUCKET).mean()
        configuration.log_action(STEP_LABEL, 3, f"{configuration.discount_bucket_column} from "
                                                f"{configuration.discount_value_column}: "
                                                f"{fine_table[configuration.discount_bucket_column].nunique()} buckets; "
                                                f"{unknown_share:.1%} of the rows '{UNKNOWN_DISCOUNT_BUCKET}'")
    else:
        configuration.log_action(STEP_LABEL, 3, "no exact discount declared: no bucket")

    # [4] the uplift cell: the price context of the row
    fine_table[UPLIFT_CELL_ID_COLUMN] = join_columns(fine_table, configuration.uplift_cell_columns)
    configuration.log_action(STEP_LABEL, 4, f"uplift_cell_id from {', '.join(configuration.uplift_cell_columns)}: "
                                            f"{fine_table[UPLIFT_CELL_ID_COLUMN].nunique():,} uplift cells")

    # [5] the checks
    configuration.log_action(STEP_LABEL, 5, "checking the bucket, the grain and that nothing else changed")
    check_discount_bucket(fine_table, configuration, check_log)
    check_grain(fine_table, configuration, check_log)
    new_columns = [column_name for column_name in fine_table.columns if column_name not in calendared.columns]
    expected_new_columns = len(ADDED_ID_COLUMNS) + (1 if configuration.discount_value_column else 0)
    configuration.log_check(STEP_LABEL, check_log, "same rows, same order, no column of step 02 changed",
                            fine_table[list(calendared.columns)].equals(calendared)
                            and len(new_columns) == expected_new_columns,
                            failure_detail="the fine table changed the rows or the columns it received",
                            context=f"{len(fine_table):,} rows · {len(new_columns)} columns added")

    # [6] the fine table is NOT written to SQL: sff_nucleo carries every fine row with its ids (step NU), and writing
    #     the 1M rows here cost 3.5 minutes of the real run for a table nobody read
    configuration.log_action(STEP_LABEL, 6, "the fine table stays in memory (every fine row is written in sff_nucleo, step NU)")

    # [7] the count of the checks; stop if anything failed
    configuration.log_action(STEP_LABEL, 7, "counting the checks")
    configuration.log_check_summary(STEP_LABEL, STEP_NAME, check_log)

    # [8] the rows per discount bucket, and a few rows with their ids
    if configuration.discount_value_column:
        configuration.log_action(STEP_LABEL, 8, "rows and money due per discount bucket:")
        bucket_column = configuration.discount_bucket_column
        configuration.show_table(fine_table.groupby(bucket_column, dropna=False)
                                 .agg(filas=(bucket_column, "size"),
                                      descuento_min=(configuration.discount_value_column, "min"),
                                      descuento_max=(configuration.discount_value_column, "max"),
                                      usd_vence=(configuration.pipeline_usd_col, "sum"))
                                 .reset_index())
    configuration.logger.doc(f"[{STEP_LABEL}] a few rows with their ids:")
    configuration.show_table(fine_table[[configuration.period_col, UNIT_ID_COLUMN, UPLIFT_CELL_ID_COLUMN]
                                        + configuration.fine_row_columns].head(configuration.example_rows_shown))
    return fine_table


def fine_row_values(fine_table: pd.DataFrame, configuration: Config) -> pd.Series:
    """The extra_revalorizacion values and the exact discount of every row, joined ("na"
    when there is none): with the unit, they are what makes a fine row unique."""
    if configuration.fine_row_columns:
        return join_columns(fine_table, configuration.fine_row_columns)
    return pd.Series(NO_REVALUATION_VALUES, index=fine_table.index)


def check_discount_bucket(fine_table: pd.DataFrame, configuration: Config, check_log: list) -> None:
    """Check 1: every exact discount has a bucket; only a null discount is 'sin_dato'."""
    if not configuration.discount_value_column:
        configuration.log_not_evaluated(STEP_LABEL, check_log, "every exact discount falls in one bucket",
                                        "no exact discount declared")
        return
    buckets = fine_table[configuration.discount_bucket_column]
    discount_values = fine_table[configuration.discount_value_column]
    without_bucket = fine_table[buckets.isna()]
    wrongly_unknown = int(((buckets == UNKNOWN_DISCOUNT_BUCKET) & discount_values.notna()).sum())
    configuration.log_check(STEP_LABEL, check_log, f"every exact discount falls in one bucket "
                                                   f"(a null one in '{UNKNOWN_DISCOUNT_BUCKET}')",
                            without_bucket.empty and wrongly_unknown == 0,
                            failure_detail=f"{len(without_bucket):,} rows without a bucket · {wrongly_unknown:,} known "
                                           f"discounts labelled '{UNKNOWN_DISCOUNT_BUCKET}'",
                            context=f"edges {configuration.discount_bucket_edges} %",
                            examples=without_bucket[[configuration.period_col, UNIT_ID_COLUMN,
                                                     configuration.discount_value_column]])


def check_grain(fine_table: pd.DataFrame, configuration: Config, check_log: list) -> None:
    """Check 2: no two raw rows are the same unit with the same revaluation values. If they
    were, the raw would not be at the grain the Config declares and every later sum would
    mix them. When it fails, the check says WHY: which columns of the raw take different
    values inside the repeated rows (a column missing from the taxonomy), or that only the
    money differs (the extract is below the declared grain and must be aggregated). The
    first repeated group is shown whole, with those columns."""
    grain_columns = [UNIT_ID_COLUMN] + configuration.fine_row_columns
    repeated_mask = fine_table.duplicated(grain_columns, keep=False)
    repeated_rows_count = int(fine_table.duplicated(grain_columns).sum())
    check_description = "one row per forecast unit, revaluation values and exact discount (the grain of the raw)"
    if repeated_rows_count == 0:
        configuration.log_check(STEP_LABEL, check_log, check_description, True,
                                context=f"{len(fine_table):,} rows = {len(fine_table):,} distinct fine rows")
        return

    # which columns separate the repeated rows, and how many repeated groups each one explains (a repeated group:
    # the rows of the same fine row, named by its unit, its revaluation values and its exact discount)
    repeated_rows = fine_table[repeated_mask].copy()
    repeated_rows[FINE_ROW_GROUP] = (repeated_rows[UNIT_ID_COLUMN] + COMBINED_ID_SEPARATOR
                                     + fine_row_values(repeated_rows, configuration))
    repeated_groups = int(repeated_rows[FINE_ROW_GROUP].nunique())
    splitting_columns = columns_that_split_the_grain(repeated_rows, grain_columns, configuration)
    groups_explained = groups_explained_by(repeated_rows, list(splitting_columns))
    if splitting_columns:
        cause_text = ("columns of the raw that differ inside the repeated rows (repeated groups where they differ): "
                      + ", ".join(f"{column_name} ({group_count:,})" for column_name, group_count in splitting_columns.items())
                      + f"; together they explain {groups_explained:,} of the {repeated_groups:,} repeated groups"
                      + (f", the other {repeated_groups - groups_explained:,} differ only in the money"
                         if groups_explained < repeated_groups else "")
                      + " → declare them as dimensions, or aggregate the extract without them")
    else:
        cause_text = (f"no column but the money differs inside the {repeated_groups:,} repeated groups: the extract is "
                      f"below the declared grain (e.g. one row per transaction) → aggregate it to the declared dimensions")

    first_row_group = repeated_rows.iloc[0][FINE_ROW_GROUP]
    first_group = repeated_rows[repeated_rows[FINE_ROW_GROUP] == first_row_group]
    shown_columns = ([configuration.period_col, UNIT_ID_COLUMN] + configuration.fine_row_columns
                     + [column_name for column_name in splitting_columns if column_name not in configuration.fine_row_columns]
                     + configuration.core_measures)
    configuration.log_check(STEP_LABEL, check_log, check_description, False,
                            failure_detail=f"{repeated_rows_count:,} rows repeat a fine row already seen "
                                           f"({repeated_mask.sum():,} rows in {repeated_groups:,} groups). {cause_text}",
                            examples=first_group[shown_columns])

    # the same problem, in SQL, to reproduce it on the extract
    configuration.logger.doc(f"[{STEP_LABEL}] to reproduce it in SQL (the declared grain, the groups with more than "
                             f"one row, and what differs inside them):")
    print(grain_reproduction_sql(configuration, list(splitting_columns)))


def groups_explained_by(repeated_rows: pd.DataFrame, splitting_columns: list) -> int:
    """How many repeated groups have at least one of the splitting columns taking more than one value."""
    if not splitting_columns:
        return 0
    varies_somewhere = (repeated_rows.groupby(FINE_ROW_GROUP)[splitting_columns].nunique(dropna=False) > 1).any(axis=1)
    return int(varies_somewhere.sum())


def grain_reproduction_sql(configuration: Config, splitting_columns: list) -> str:
    """A SQL query, built from the Config, that finds the groups of the declared grain with
    more than one row on the extract, and says how many of them each splitting column explains.
    The extract is taken from configuration.raw_extract_sql (a placeholder when not declared)."""
    grain_columns = ([configuration.period_col] + configuration.dimension_columns
                     + [column_name for column_name in configuration.fine_row_columns
                        if column_name not in configuration.dimension_columns])
    source = configuration.raw_extract_sql or "<la query o la tabla de tu extracto>"
    select_grain = ",\n           ".join(grain_columns)
    splitting_counts = "".join(f"\n           COUNT(DISTINCT {column_name}) AS distintos_{column_name},"
                               f"\n           MIN({column_name}) AS min_{column_name}, MAX({column_name}) AS max_{column_name},"
                               for column_name in splitting_columns)
    explained_sums = "".join(f",\n       SUM(CASE WHEN distintos_{column_name} > 1 THEN 1 ELSE 0 END) AS grupos_que_separa_{column_name}"
                             for column_name in splitting_columns)
    return (f"WITH raw AS (\n    {source}\n),\n"
            f"grupos AS (\n    SELECT {select_grain},\n           COUNT(*) AS filas,{splitting_counts}"
            f"\n           SUM({configuration.pipeline_units_col}) AS unidades_vencen"
            f"\n    FROM raw\n    GROUP BY {select_grain}\n    HAVING COUNT(*) > 1\n)\n"
            f"SELECT COUNT(*) AS grupos_repetidos,\n       SUM(filas) AS filas_implicadas{explained_sums}\nFROM grupos;\n"
            f"-- to see the groups: SELECT * FROM grupos ORDER BY filas DESC;\n"
            f"-- COUNT(DISTINCT) ignores NULL: a group with a value and a NULL is not counted as split")


def columns_that_split_the_grain(repeated_rows: pd.DataFrame, grain_columns: list, configuration: Config) -> dict:
    """The raw columns (not measures, not the ids added here) that take more than one value
    inside at least one repeated fine row, with how many groups: they are what makes the
    raw finer than the declared dimensions. Sorted by groups, most first."""
    measure_columns = (set(configuration.core_measures) | set(configuration.extra_measure_cols)
                       | set(configuration.closed_month_measure_cols) | set(configuration.maturation_columns))
    added_columns = set(ADDED_ID_COLUMNS) | {FINE_ROW_GROUP, configuration.discount_bucket_column}
    candidate_columns = [column_name for column_name in repeated_rows.columns
                         if column_name not in measure_columns and column_name not in added_columns
                         and column_name not in grain_columns and not column_name.startswith("s0_")]
    grouped = repeated_rows.groupby(FINE_ROW_GROUP, sort=False)
    splitting_columns = {}
    for column_name in candidate_columns:
        groups_where_it_varies = int((grouped[column_name].nunique(dropna=False) > 1).sum())
        if groups_where_it_varies:
            splitting_columns[column_name] = groups_where_it_varies
    return dict(sorted(splitting_columns.items(), key=lambda item: -item[1]))
