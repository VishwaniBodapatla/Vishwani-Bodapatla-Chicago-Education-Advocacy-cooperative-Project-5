import pandas as pd
import streamlit as st
import plotly.express as px
from difflib import get_close_matches
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
OUTPUT_DIR = PROJECT_DIR / "output"

# =========================
# CONFIG
# =========================
EXCEL_PATH_DEFAULT = str(OUTPUT_DIR / "jobright_jobs.xlsx")
SUMMARY_OUT_DEFAULT = str(OUTPUT_DIR / "branch_salary_dashboard_ready.xlsx")
JOBS_SHEET = "jobs"
SUMMARY_SHEET = "dashboard_summary"

SENIORITY_COL = "seniority"          # if you have a real seniority column
DEGREE_COL = "degree_level"


# =========================
# COMMON HELPERS
# =========================
def present(series: pd.Series) -> pd.Series:
    """True if cell has a real value (not blank, not NaN string)."""
    s = series.astype(str).fillna("").str.strip()
    sl = s.str.lower()
    return (s != "") & (sl != "nan") & (sl != "none")


def add_value_labels(fig, chart_type: str):
    """Show value labels on top of marks/bars where it makes sense."""
    if chart_type == "bar":
        fig.update_traces(texttemplate="%{y}", textposition="outside")
        fig.update_layout(uniformtext_minsize=8, uniformtext_mode="hide")
    elif chart_type == "line":
        fig.update_traces(mode="lines+markers+text", texttemplate="%{y}", textposition="top center")
    elif chart_type == "scatter":
        fig.update_traces(mode="markers+text", texttemplate="%{y}", textposition="top center")
    return fig


def add_bar_labels(fig):
    """Force numeric labels to show on top of bars (avoid overlap clipping)."""
    fig.update_traces(
        texttemplate="%{y:.3s}",
        textposition="outside",
        cliponaxis=False,
    )
    fig.update_layout(
        uniformtext_minsize=8,
        uniformtext_mode="show",
        margin=dict(t=80, b=80, l=60, r=40),
    )
    return fig


def safe_mean(s: pd.Series):
    s2 = pd.to_numeric(s, errors="coerce")
    return float(s2.mean()) if s2.notna().any() else None


def safe_median(s: pd.Series):
    s2 = pd.to_numeric(s, errors="coerce")
    return float(s2.median()) if s2.notna().any() else None


def safe_min(s: pd.Series):
    s2 = pd.to_numeric(s, errors="coerce")
    return float(s2.min()) if s2.notna().any() else None


def safe_max(s: pd.Series):
    s2 = pd.to_numeric(s, errors="coerce")
    return float(s2.max()) if s2.notna().any() else None


def topn_by(df: pd.DataFrame, value_col: str, n: int, ascending: bool = False) -> pd.DataFrame:
    t = df.dropna(subset=[value_col]).copy()
    return t.sort_values(value_col, ascending=ascending).head(n)


def pick_sort_col(chosen_sort: str, plotted_metrics: list[str], default_col: str = "count") -> str:
    """chosen_sort can be a real column name or '(use plotted metric)'."""
    if chosen_sort == "(use plotted metric)":
        return plotted_metrics[0] if plotted_metrics else default_col
    return chosen_sort


def enforce_category_order(fig, ordered_values: list[str]):
    """Force Plotly to render categories in the exact order we computed."""
    fig.update_xaxes(categoryorder="array", categoryarray=ordered_values)
    return fig


# =========================
# DATA
# =========================
def load_jobs(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=JOBS_SHEET)
    for c in ["salary_min", "salary_max"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


@st.cache_data(show_spinner=True)
def cached_jobs(path: str) -> pd.DataFrame:
    return load_jobs(path)


# =========================
# MAIN DASHBOARD HELPERS
# =========================
def guess_chart(df: pd.DataFrame):
    cols = df.columns.tolist()
    numeric = [c for c in cols if pd.api.types.is_numeric_dtype(df[c])]
    cat = [c for c in cols if not pd.api.types.is_numeric_dtype(df[c])]

    if "company_name" in cols and "salary_min" in numeric:
        return ("bar", ["company_name"], ["salary_min", "salary_max"] if "salary_max" in numeric else ["salary_min"])
    if "posted" in cols and numeric:
        return ("line", ["posted"], [numeric[0]])
    if cat and numeric:
        return ("bar", [cat[0]], [numeric[0]])
    if len(numeric) >= 2:
        return ("scatter", [numeric[0]], [numeric[1]])
    return ("bar", [cols[0]], cols[1:2] if len(cols) > 1 else [cols[0]])


def parse_ai_query(df: pd.DataFrame, q: str):
    """Lightweight “AI” parser: maps words to columns using close-match + finds chart keyword."""
    ql = (q or "").lower().strip()
    if not ql:
        return None

    chart = None
    for k in ["bar", "line", "scatter", "box", "histogram"]:
        if k in ql:
            chart = k
            break

    tokens = [t.strip(" ,;:.") for t in ql.split() if t.strip()]
    cols = df.columns.tolist()

    matched = []
    for t in tokens:
        m = get_close_matches(t, cols, n=1, cutoff=0.72)
        if m:
            matched.append(m[0])

    matched = list(dict.fromkeys(matched))
    if len(matched) >= 2:
        return (chart or "bar", [matched[0]], matched[1:])
    return None


def build_multi_x(df: pd.DataFrame, x_cols: list[str]) -> tuple[pd.DataFrame, str]:
    """If multiple X columns are selected, combine them into a single key for plotting."""
    if not x_cols:
        return df, ""
    if len(x_cols) == 1:
        return df, x_cols[0]
    tmp = df.copy()
    key = "__x_multi__"
    tmp[key] = tmp[x_cols].astype(str).agg(" | ".join, axis=1)
    return tmp, key


# =========================
# SENIORITY MARKER SUPPORT (your ask) ✅
# =========================
def detect_seniority_marker_cols(df: pd.DataFrame) -> list[str]:
    """
    Detect marker-like columns that look like seniority categories.
    Example marker columns: entry, lead, director (each cell contains text when present)
    """
    keywords = {
        "intern", "entry", "junior", "jr", "mid", "senior", "sr",
        "staff", "principal", "lead", "manager", "director", "vp", "vice", "cto", "ciso"
    }

    cols = []
    for c in df.columns:
        cl = str(c).strip().lower()

        looks_like_seniority_name = (cl in keywords) or any(k in cl for k in keywords)
        if not looks_like_seniority_name:
            continue

        # marker-like: small unique set (usually blank + few tokens)
        vals = df[c].dropna().astype(str).str.strip().str.lower().unique()
        if len(vals) <= 5:
            cols.append(c)

    return cols


def build_virtual_seniority_from_markers(
    df: pd.DataFrame,
    marker_cols: list[str],
    out_col: str = "__seniority_virtual__",
) -> tuple[pd.DataFrame, str]:
    """
    For each row, assign the FIRST marker column that is present.
    If multiple are present, first wins (based on marker_cols order).
    """
    if not marker_cols:
        return df, out_col

    tmp = df.copy()
    tmp[out_col] = ""

    for c in marker_cols:
        if c not in tmp.columns:
            continue
        m = present(tmp[c])
        tmp.loc[m & (tmp[out_col] == ""), out_col] = str(c)

    tmp[out_col] = tmp[out_col].replace("", pd.NA)
    return tmp, out_col


# =========================
# BRANCH SUMMARY GENERATION
# =========================
def generate_branch_salary_summary_excel_sliced(
    jobs_excel_path: str,
    out_excel_path: str = SUMMARY_OUT_DEFAULT,
    sheet_in: str = JOBS_SHEET,
    sheet_out: str = SUMMARY_SHEET,
    slice_col: str | None = None,
    slice_value: str | None = None,
) -> Path:
    """
    Generates summary Excel from jobright_jobs.xlsx.
    Uses marker-like columns (<=5 unique values) as branch/role/field columns.
    Optionally filters jobs where df[slice_col] == slice_value (case-insensitive exact match).
    """
    in_path = Path(jobs_excel_path)
    if not in_path.exists():
        raise FileNotFoundError(f"Input Excel not found: {jobs_excel_path}")

    df = pd.read_excel(in_path, sheet_name=sheet_in)

    if slice_col and slice_col in df.columns and slice_value and slice_value != "(All)":
        s = df[slice_col].astype(str).fillna("").str.strip().str.lower()
        target = str(slice_value).strip().lower()
        df = df[s == target].copy()

    for c in ["salary_min", "salary_max"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    exclude = {"salary_min", "salary_max"}
    candidate_cols = [c for c in df.columns if c not in exclude]

    field_cols = []
    for c in candidate_cols:
        vals = df[c].dropna().astype(str).str.strip().str.lower().unique()
        if len(vals) <= 5:
            field_cols.append(c)

    total_jobs = len(df)
    rows = []

    for col in field_cols:
        mask = present(df[col])
        cnt = int(mask.sum())
        if cnt == 0:
            continue

        sub = df.loc[mask].copy()
        avg_min = sub["salary_min"].mean() if "salary_min" in sub.columns else None
        avg_max = sub["salary_max"].mean() if "salary_max" in sub.columns else None
        med_min = sub["salary_min"].median() if "salary_min" in sub.columns else None
        med_max = sub["salary_max"].median() if "salary_max" in sub.columns else None

        salary_range_avg = (avg_max - avg_min) if pd.notna(avg_min) and pd.notna(avg_max) else None
        avg_avg = ((avg_min + avg_max) / 2) if pd.notna(avg_min) and pd.notna(avg_max) else None

        rows.append(
            {
                "branch": col,
                "count": cnt,
                "percent_of_jobs": (cnt / total_jobs) * 100.0 if total_jobs else None,
                "avg_salary_min": float(avg_min) if pd.notna(avg_min) else None,
                "avg_salary_max": float(avg_max) if pd.notna(avg_max) else None,
                "avg_salary_avg": float(avg_avg) if pd.notna(avg_avg) else None,
                "median_salary_min": float(med_min) if pd.notna(med_min) else None,
                "median_salary_max": float(med_max) if pd.notna(med_max) else None,
                "salary_range_avg": float(salary_range_avg) if pd.notna(salary_range_avg) else None,
                "std_salary_min": float(sub["salary_min"].std()) if "salary_min" in sub.columns else None,
                "std_salary_max": float(sub["salary_max"].std()) if "salary_max" in sub.columns else None,
            }
        )

    out = pd.DataFrame(rows)
    if out.empty:
        slice_desc = f"{slice_col}={slice_value}" if slice_col and slice_value else "no slice"
        raise RuntimeError(f"No marker-like columns produced results ({slice_desc}).")

    if "avg_salary_max" in out.columns:
        out["rank_by_avg_max"] = out["avg_salary_max"].rank(ascending=False, method="dense")
    if "avg_salary_min" in out.columns:
        out["rank_by_avg_min"] = out["avg_salary_min"].rank(ascending=False, method="dense")
    if "count" in out.columns:
        out["rank_by_count"] = out["count"].rank(ascending=False, method="dense")

    if "rank_by_avg_max" in out.columns:
        out = out.sort_values("rank_by_avg_max", ascending=True)

    out_round = out.copy()
    for c in out_round.columns:
        if c != "branch":
            out_round[c] = pd.to_numeric(out_round[c], errors="coerce").round(2)

    out_path = Path(out_excel_path)
    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        out_round.to_excel(w, sheet_name=sheet_out, index=False)
        out.to_excel(w, sheet_name="raw_values", index=False)

    return out_path


@st.cache_data(show_spinner=True)
def load_branch_summary(path: str, sheet: str = SUMMARY_SHEET) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet)

    if "avg_salary_avg" not in df.columns and "avg_salary_min" in df.columns and "avg_salary_max" in df.columns:
        df = df.copy()
        df["avg_salary_avg"] = (
            pd.to_numeric(df["avg_salary_min"], errors="coerce") + pd.to_numeric(df["avg_salary_max"], errors="coerce")
        ) / 2

    num_cols = [
        "count",
        "percent_of_jobs",
        "avg_salary_min",
        "avg_salary_max",
        "avg_salary_avg",
        "median_salary_min",
        "median_salary_max",
        "salary_range_avg",
        "std_salary_min",
        "std_salary_max",
        "rank_by_avg_max",
        "rank_by_avg_min",
        "rank_by_count",
    ]
    for c in num_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


# =========================
# SLICE SUMMARY (grouped bars by group_col values)
# =========================
def compute_slice_summary_from_df(df: pd.DataFrame, group_col: str, slice_values: tuple[str, ...]) -> pd.DataFrame:
    """
    Output long form:
      branch, group_value, count, avg_salary_min, avg_salary_max, avg_salary_avg, median_salary_min, median_salary_max
    - branch is marker-like column name (<=5 unique values)
    - group_value is each selected value in df[group_col]
    """
    if group_col not in df.columns:
        return pd.DataFrame()

    tmp = df.copy()
    tmp[group_col] = tmp[group_col].astype(str).fillna("").str.strip()

    selected = set([str(v).strip() for v in slice_values if str(v).strip()])
    if not selected:
        return pd.DataFrame()

    tmp = tmp[tmp[group_col].isin(selected)].copy()
    if tmp.empty:
        return pd.DataFrame()

    for c in ["salary_min", "salary_max"]:
        if c in tmp.columns:
            tmp[c] = pd.to_numeric(tmp[c], errors="coerce")

    exclude = {"salary_min", "salary_max"}
    candidate_cols = [c for c in tmp.columns if c not in exclude]

    field_cols = []
    for c in candidate_cols:
        vals = tmp[c].dropna().astype(str).str.strip().str.lower().unique()
        if len(vals) <= 5:
            field_cols.append(c)

    rows = []
    for branch_col in field_cols:
        for gv in selected:
            subg = tmp[tmp[group_col] == gv].copy()
            if subg.empty:
                continue

            mask = present(subg[branch_col])
            cnt = int(mask.sum())
            if cnt == 0:
                continue

            sub = subg.loc[mask].copy()
            avg_min = sub["salary_min"].mean() if "salary_min" in sub.columns else None
            avg_max = sub["salary_max"].mean() if "salary_max" in sub.columns else None
            med_min = sub["salary_min"].median() if "salary_min" in sub.columns else None
            med_max = sub["salary_max"].median() if "salary_max" in sub.columns else None
            avg_avg = ((avg_min + avg_max) / 2) if pd.notna(avg_min) and pd.notna(avg_max) else None

            rows.append(
                {
                    "branch": branch_col,
                    "group_value": gv,
                    "count": cnt,
                    "avg_salary_min": float(avg_min) if pd.notna(avg_min) else None,
                    "avg_salary_max": float(avg_max) if pd.notna(avg_max) else None,
                    "avg_salary_avg": float(avg_avg) if pd.notna(avg_avg) else None,
                    "median_salary_min": float(med_min) if pd.notna(med_min) else None,
                    "median_salary_max": float(med_max) if pd.notna(med_max) else None,
                }
            )

    out = pd.DataFrame(rows)
    if out.empty:
        return out

    for c in ["count", "avg_salary_min", "avg_salary_max", "avg_salary_avg", "median_salary_min", "median_salary_max"]:
        if c in out.columns:
            out[c] = pd.to_numeric(out[c], errors="coerce")
    return out


@st.cache_data(show_spinner=True)
def compute_slice_summary(
    jobs_excel_path: str,
    group_col: str,
    slice_values: tuple[str, ...],
    sheet_in: str = JOBS_SHEET,
) -> pd.DataFrame:
    in_path = Path(jobs_excel_path)
    if not in_path.exists():
        raise FileNotFoundError(f"Input Excel not found: {jobs_excel_path}")
    df = pd.read_excel(in_path, sheet_name=sheet_in)
    return compute_slice_summary_from_df(df, group_col=group_col, slice_values=slice_values)


# =========================
# SUMMARY DASHBOARD (Bar split supports: degree_level, seniority column, OR seniority marker fields) ✅
# =========================
def render_branch_summary_dashboard(
    sdf: pd.DataFrame,
    render_key: str = "base",
    jobs_excel_path: str | None = None,
):
    sdf = sdf[sdf["branch"].astype(str).str.strip().ne("")].copy()
    if sdf.empty:
        st.warning("No rows to display.")
        return

    metric_cols = [c for c in sdf.columns if c != "branch" and pd.api.types.is_numeric_dtype(sdf[c])]
    if not metric_cols:
        st.error("No numeric metric columns found in summary sheet.")
        return

    pretty_metric = {
        "count": "Count (demand)",
        "percent_of_jobs": "Percent of jobs",
        "avg_salary_min": "Avg Salary (Min)",
        "avg_salary_max": "Avg Salary (Max)",
        "avg_salary_avg": "Avg Salary (Average)",
        "median_salary_min": "Median Salary (Min)",
        "median_salary_max": "Median Salary (Max)",
        "salary_range_avg": "Avg Salary Range (Max-Min)",
        "std_salary_min": "Salary Spread Std (Min)",
        "std_salary_max": "Salary Spread Std (Max)",
        "rank_by_avg_max": "Rank by Avg Max",
        "rank_by_avg_min": "Rank by Avg Min",
        "rank_by_count": "Rank by Count",
    }

    st.divider()
    st.subheader("Field Selection")

    all_fields = sorted(sdf["branch"].astype(str).unique().tolist())
    default_exclude_candidates = {"role_type", "role_selected", "work_model", "platform", "degree_level", "seniority"}
    default_excluded = [b for b in all_fields if str(b).strip().lower() in default_exclude_candidates]

    include_mode = st.radio(
        "Selection mode",
        options=["Exclude list", "Include only list"],
        index=0,
        horizontal=True,
        key=f"include_mode_{render_key}",
    )

    if include_mode == "Exclude list":
        selected_excluded = st.multiselect(
            "Exclude fields (rows)",
            options=all_fields,
            default=default_excluded,
            help="Remove rows like role_type/work_model/platform so they don't dominate charts.",
            key=f"excluded_{render_key}",
        )
        f_sdf = sdf[~sdf["branch"].astype(str).isin(set(map(str, selected_excluded)))].copy()
    else:
        selected_included = st.multiselect(
            "Include ONLY these fields (rows)",
            options=all_fields,
            default=[b for b in all_fields if b not in default_excluded],
            key=f"included_{render_key}",
        )
        f_sdf = sdf[sdf["branch"].astype(str).isin(set(map(str, selected_included)))].copy()

    recalc_percent = st.checkbox(
        "Recalculate demand share (%) based on current selection",
        value=True,
        key=f"recalc_{render_key}",
    )
    if recalc_percent and "count" in f_sdf.columns:
        total = pd.to_numeric(f_sdf["count"], errors="coerce").fillna(0).sum()
        f_sdf = f_sdf.copy()
        f_sdf["percent_of_jobs"] = (
            (pd.to_numeric(f_sdf["count"], errors="coerce").fillna(0) / total) * 100.0 if total > 0 else 0.0
        )

    show_labels = st.checkbox(
        "Show value labels (bar charts)",
        value=True,
        key=f"labels_{render_key}",
    )

    st.divider()
    st.subheader("Global Ordering (applies to ALL charts)")

    global_sort_metric = st.selectbox(
        "Order by metric (applies to all charts)",
        options=list(dict.fromkeys(["(use plotted metric)"] + metric_cols)),
        index=0,
        key=f"global_sort_{render_key}",
    )

    global_order = st.radio(
        "Order direction",
        options=["Descending", "Ascending"],
        index=0,
        horizontal=True,
        key=f"global_dir_{render_key}",
    )
    global_ascending = (global_order == "Ascending")

    # -------------------------
    # BAR CHART (split supports seniority marker fields) ✅
    # -------------------------
    st.divider()
    st.subheader("Bar Chart (All metrics + Global ordering)")

    split_by = st.selectbox(
        "Group bars by",
        options=["(None)", DEGREE_COL, SENIORITY_COL, "seniority_fields"],  # ✅ NEW: seniority_fields
        index=0,
        key=f"bar_split_by_{render_key}",
    )

    split_values: list[str] = []
    marker_cols: list[str] = []
    use_marker_mode = False

    jdf = None
    if split_by != "(None)":
        if not jobs_excel_path:
            st.warning("Missing jobs_excel_path. Pass jobs_excel_path=excel_path when calling render_branch_summary_dashboard().")
        else:
            jdf = cached_jobs(jobs_excel_path)

            # 1) Normal column mode (degree_level or seniority column)
            if split_by in (DEGREE_COL, SENIORITY_COL) and split_by in jdf.columns:
                opts = (
                    jdf[split_by]
                    .dropna()
                    .astype(str)
                    .str.strip()
                    .replace("", pd.NA)
                    .dropna()
                    .unique()
                    .tolist()
                )
                opts = sorted(opts)
                split_values = st.multiselect(
                    f"Select {split_by} values",
                    options=opts,
                    default=opts[:4] if len(opts) >= 4 else opts,
                    key=f"bar_split_vals_{render_key}",
                )

            # 2) Marker field mode (entry/lead/director as separate columns)
            elif split_by == "seniority_fields":
                marker_cols = detect_seniority_marker_cols(jdf)
                if not marker_cols:
                    st.warning("No seniority marker fields found. (Expected columns like entry/lead/director...)")
                else:
                    use_marker_mode = True
                    split_values = st.multiselect(
                        "Select seniority fields (marker columns)",
                        options=marker_cols,
                        default=marker_cols[:6] if len(marker_cols) >= 6 else marker_cols,
                        key=f"bar_split_marker_vals_{render_key}",
                    )
            else:
                st.warning(f"'{split_by}' not found in jobs data.")

    bar_n = st.slider(
        "Top N (Bar)",
        min_value=3,
        max_value=100,
        value=12,
        step=1,
        key=f"bar_n_{render_key}",
    )

    default_bar_metrics = [c for c in ["avg_salary_min", "avg_salary_max"] if c in metric_cols]
    if not default_bar_metrics:
        default_bar_metrics = ["count"] if "count" in metric_cols else metric_cols[:1]

    bar_metrics = st.multiselect(
        "Metrics to plot",
        options=metric_cols,
        default=default_bar_metrics,
        key=f"bar_metrics_{render_key}",
    )

    if bar_metrics:
        sort_col = pick_sort_col(global_sort_metric, bar_metrics, default_col="count")

        # ✅ Split mode ON
        if split_by != "(None)" and jobs_excel_path and split_values:
            # Build long_df
            if use_marker_mode:
                # Build virtual group column from marker columns, then compute slice summary
                if jdf is None:
                    jdf = cached_jobs(jobs_excel_path)
                df2, gcol = build_virtual_seniority_from_markers(jdf, marker_cols=list(split_values), out_col="__seniority_virtual__")
                long_df = compute_slice_summary_from_df(df2, group_col=gcol, slice_values=tuple(split_values))
            else:
                long_df = compute_slice_summary(
                    jobs_excel_path,
                    group_col=split_by,
                    slice_values=tuple(split_values),
                    sheet_in=JOBS_SHEET,
                )

            # respect your field selection filters
            keep_branches = set(f_sdf["branch"].astype(str).tolist())
            long_df = long_df[long_df["branch"].astype(str).isin(keep_branches)].copy()

            if long_df.empty:
                st.info("No grouped data for those selections.")
            else:
                # ordering for branches
                if sort_col == "count":
                    ord_df = long_df.groupby("branch", as_index=False)["count"].sum()
                else:
                    ord_df = long_df.groupby("branch", as_index=False)[sort_col].mean(numeric_only=True)

                ord_df = ord_df.dropna(subset=[sort_col]).sort_values(sort_col, ascending=global_ascending).head(bar_n)
                ordered_x = ord_df["branch"].astype(str).tolist()

                plot_df = long_df[long_df["branch"].astype(str).isin(set(ordered_x))].copy()
                plot_df["branch"] = pd.Categorical(plot_df["branch"].astype(str), categories=ordered_x, ordered=True)
                plot_df = plot_df.sort_values("branch")

                title_prefix = "Bottom" if global_ascending else "Top"
                split_label = "seniority_fields" if use_marker_mode else split_by
                title = f"{title_prefix} {bar_n} by {pretty_metric.get(sort_col, sort_col)} (split by {split_label})"

                if len(bar_metrics) == 1:
                    m = bar_metrics[0]
                    fig_bar = px.bar(
                        plot_df,
                        x="branch",
                        y=m,
                        color="group_value",
                        barmode="group",
                        title=title,
                    )
                    fig_bar.update_layout(xaxis_tickangle=-30)
                    fig_bar.update_xaxes(title="Field / Role / Branch")
                    fig_bar.update_yaxes(title=pretty_metric.get(m, m))
                    if show_labels:
                        fig_bar = add_bar_labels(fig_bar)
                    st.plotly_chart(fig_bar, use_container_width=True)
                else:
                    melt = plot_df.melt(
                        id_vars=["branch", "group_value"],
                        value_vars=bar_metrics,
                        var_name="metric",
                        value_name="value",
                    )
                    melt["metric"] = melt["metric"].map(lambda x: pretty_metric.get(x, x))
                    fig_bar = px.bar(
                        melt,
                        x="branch",
                        y="value",
                        color="group_value",
                        barmode="group",
                        facet_col="metric",
                        title=title,
                    )
                    fig_bar.update_layout(xaxis_tickangle=-30)
                    fig_bar.update_xaxes(title="Field / Role / Branch")
                    fig_bar.update_yaxes(title="Value")
                    if show_labels:
                        fig_bar = add_bar_labels(fig_bar)
                    st.plotly_chart(fig_bar, use_container_width=True)

        # ✅ Split mode OFF (original)
        else:
            bar_df = topn_by(f_sdf, sort_col, bar_n, ascending=global_ascending).copy()
            bar_df = bar_df.sort_values(sort_col, ascending=global_ascending).copy()
            ordered_x = bar_df["branch"].astype(str).tolist()
            title_prefix = "Bottom" if global_ascending else "Top"
            title = f"{title_prefix} {bar_n} by {pretty_metric.get(sort_col, sort_col)}"

            if len(bar_metrics) == 1:
                m = bar_metrics[0]
                fig_bar = px.bar(bar_df, x="branch", y=m, title=title)
                fig_bar = enforce_category_order(fig_bar, ordered_x)
            else:
                melt = bar_df.melt(id_vars=["branch"], value_vars=bar_metrics, var_name="metric", value_name="value")
                melt["metric"] = melt["metric"].map(lambda x: pretty_metric.get(x, x))
                fig_bar = px.bar(melt, x="branch", y="value", color="metric", barmode="group", title=title)
                fig_bar = enforce_category_order(fig_bar, ordered_x)

            fig_bar.update_xaxes(title="Field / Role / Branch")
            fig_bar.update_yaxes(title="Value")
            fig_bar.update_layout(xaxis_tickangle=-30)
            if show_labels:
                fig_bar = add_bar_labels(fig_bar)
            st.plotly_chart(fig_bar, use_container_width=True)
    else:
        st.info("Pick at least one metric for the bar chart.")

    # -------------------------
    # PIE CHART
    # -------------------------
    st.divider()
    st.subheader("Pie Chart (All metrics + Global ordering)")

    pie_n = st.slider(
        "Top N (Pie)",
        min_value=3,
        max_value=50,
        value=12,
        step=1,
        key=f"pie_n_{render_key}",
    )
    pie_metric = st.selectbox(
        "Pie value metric",
        options=metric_cols,
        index=metric_cols.index("percent_of_jobs") if "percent_of_jobs" in metric_cols else 0,
        key=f"pie_metric_{render_key}",
    )

    pie_sort_col = pick_sort_col(global_sort_metric, [pie_metric], default_col="count")
    pie_df = topn_by(f_sdf, pie_sort_col, pie_n, ascending=global_ascending).copy()
    pie_df = pie_df.sort_values(pie_sort_col, ascending=global_ascending).copy()
    title_prefix = "Bottom" if global_ascending else "Top"

    fig_pie = px.pie(
        pie_df,
        names="branch",
        values=pie_metric,
        title=f"{title_prefix} {pie_n} by {pretty_metric.get(pie_sort_col, pie_sort_col)} (Pie uses {pretty_metric.get(pie_metric, pie_metric)})",
    )
    st.plotly_chart(fig_pie, use_container_width=True)

    # -------------------------
    # SCATTER CHART
    # -------------------------
    st.divider()
    st.subheader("Scatter (All metrics + Global ordering + Different colors per role/field)")

    scatter_n = st.slider(
        "Top N (Scatter points)",
        min_value=3,
        max_value=200,
        value=25,
        step=1,
        key=f"scatter_n_{render_key}",
    )
    x_metric = st.selectbox(
        "X-axis",
        options=metric_cols,
        index=metric_cols.index("count") if "count" in metric_cols else 0,
        key=f"scatter_x_{render_key}",
    )
    y_metric = st.selectbox(
        "Y-axis",
        options=metric_cols,
        index=metric_cols.index("avg_salary_max")
        if "avg_salary_max" in metric_cols
        else (1 if len(metric_cols) > 1 else 0),
        key=f"scatter_y_{render_key}",
    )
    size_metric = st.selectbox(
        "Point size (optional)",
        options=["(none)"] + metric_cols,
        index=0,
        key=f"scatter_size_{render_key}",
    )

    show_toolbar = st.toggle(
        "Show chart options toolbar",
        value=True,
        key=f"scatter_toolbar_{render_key}",
    )
    marker_size = st.slider(
        "Base point size",
        6,
        18,
        10,
        1,
        key=f"scatter_pt_{render_key}",
    )

    scatter_sort_col = pick_sort_col(global_sort_metric, [y_metric], default_col="count")
    scatter_df = topn_by(f_sdf, scatter_sort_col, scatter_n, ascending=global_ascending).copy()
    scatter_df = scatter_df.dropna(subset=[x_metric, y_metric]).copy()

    if scatter_df.empty:
        st.warning("Scatter chart has no data after filtering.")
    else:
        args = dict(
            data_frame=scatter_df,
            x=x_metric,
            y=y_metric,
            color="branch",
            hover_name="branch",
            hover_data={x_metric: True, y_metric: True, "count": True, "percent_of_jobs": True}
            if "count" in scatter_df.columns and "percent_of_jobs" in scatter_df.columns
            else None,
            title=f"{'Bottom' if global_ascending else 'Top'} {scatter_n} points by {pretty_metric.get(scatter_sort_col, scatter_sort_col)}",
        )
        if size_metric != "(none)":
            args["size"] = size_metric

        fig_scatter = px.scatter(**args)
        fig_scatter.update_traces(marker=dict(size=marker_size))
        fig_scatter.update_xaxes(title=pretty_metric.get(x_metric, x_metric))
        fig_scatter.update_yaxes(title=pretty_metric.get(y_metric, y_metric))

        st.plotly_chart(
            fig_scatter,
            use_container_width=True,
            config={
                "displayModeBar": show_toolbar,
                "displaylogo": False,
                "modeBarButtonsToRemove": [
                    "lasso2d",
                    "select2d",
                    "autoScale2d",
                    "toggleSpikelines",
                    "hoverCompareCartesian",
                    "hoverClosestCartesian",
                ],
            },
        )

    # -------------------------
    # COMPARISON CHART
    # -------------------------
    st.divider()
    st.subheader("Comparison (All metrics + Global ordering)")

    comp_n = st.slider(
        "Top N (Comparison)",
        min_value=3,
        max_value=100,
        value=12,
        step=1,
        key=f"comp_n_{render_key}",
    )
    default_comp_metrics = [c for c in ["avg_salary_min", "avg_salary_max", "avg_salary_avg"] if c in metric_cols]
    if not default_comp_metrics:
        default_comp_metrics = metric_cols[:1]

    comp_metrics = st.multiselect(
        "Metrics to compare (grouped bars)",
        options=metric_cols,
        default=default_comp_metrics,
        key=f"comp_metrics_{render_key}",
    )

    if comp_metrics:
        comp_sort_col = pick_sort_col(global_sort_metric, comp_metrics, default_col="count")
        comp_df = topn_by(f_sdf, comp_sort_col, comp_n, ascending=global_ascending).copy()
        comp_df = comp_df.sort_values(comp_sort_col, ascending=global_ascending).copy()
        ordered_x = comp_df["branch"].astype(str).tolist()
        title_prefix = "Bottom" if global_ascending else "Top"
        title = f"{title_prefix} {comp_n} by {pretty_metric.get(comp_sort_col, comp_sort_col)} (Comparison)"

        comp_long = comp_df.melt(id_vars=["branch"], value_vars=comp_metrics, var_name="metric", value_name="value")
        comp_long["metric"] = comp_long["metric"].map(lambda x: pretty_metric.get(x, x))

        fig_comp = px.bar(comp_long, x="branch", y="value", color="metric", barmode="group", title=title)
        fig_comp = enforce_category_order(fig_comp, ordered_x)
        fig_comp.update_xaxes(title="Field / Role / Branch")
        fig_comp.update_yaxes(title="Value")
        fig_comp.update_layout(xaxis_tickangle=-30)
        if show_labels:
            fig_comp = add_bar_labels(fig_comp)
        st.plotly_chart(fig_comp, use_container_width=True)
    else:
        st.info("Select at least one metric for comparison.")

    st.divider()
    st.subheader("Summary Table (sortable)")

    table_sort_col = pick_sort_col(
        global_sort_metric,
        ["count"] if "count" in metric_cols else metric_cols[:1],
        default_col=metric_cols[0],
    )
    cols_to_show = ["branch"] + metric_cols
    cols_to_show = [c for c in cols_to_show if c in f_sdf.columns]

    table_df = f_sdf[cols_to_show].copy()
    if table_sort_col in table_df.columns:
        table_df = table_df.sort_values(table_sort_col, ascending=global_ascending, na_position="last")

    st.dataframe(table_df, use_container_width=True, height=520)


# =========================
# APP
# =========================
st.set_page_config(page_title="Jobright Dashboard (Integrated)", layout="wide")
st.title("Jobright Dashboard (Integrated: Main + Branch Salary Summary)")

with st.sidebar:
    st.header("Data Source")
    excel_path = st.text_input("Excel path", value=EXCEL_PATH_DEFAULT)
    refresh = st.button("Reload Excel / Clear Cache")

if refresh:
    st.cache_data.clear()

tab_main, tab_summary = st.tabs(["Main Dashboard", "Branch Salary Summary"])

# =========================
# TAB 1: MAIN
# =========================
with tab_main:
    try:
        df = cached_jobs(excel_path)
    except Exception as e:
        st.error(f"Failed to load Excel: {e}")
        st.stop()

    st.write(f"Rows: **{len(df)}**  |  Columns: **{len(df.columns)}**")
    st.dataframe(df, use_container_width=True, height=260)

    st.divider()
    st.subheader("Filters")
    fcol1, fcol2, fcol3, fcol4 = st.columns(4)

    with fcol1:
        company = st.multiselect(
            "Company",
            sorted(df["company_name"].dropna().unique().tolist()) if "company_name" in df.columns else [],
        )
    with fcol2:
        work_model = st.multiselect(
            "Work Model",
            sorted(df["work_model"].dropna().unique().tolist()) if "work_model" in df.columns else [],
        )
    with fcol3:
        role_type = st.multiselect(
            "Role Type",
            sorted(df["role_type"].dropna().unique().tolist()) if "role_type" in df.columns else [],
        )
    with fcol4:
        seniority = st.multiselect(
            "Seniority",
            sorted(df[SENIORITY_COL].dropna().unique().tolist()) if SENIORITY_COL in df.columns else [],
        )

    fdf = df.copy()
    if company and "company_name" in fdf.columns:
        fdf = fdf[fdf["company_name"].isin(company)]
    if work_model and "work_model" in fdf.columns:
        fdf = fdf[fdf["work_model"].isin(work_model)]
    if role_type and "role_type" in fdf.columns:
        fdf = fdf[fdf["role_type"].isin(role_type)]
    if seniority and SENIORITY_COL in fdf.columns:
        fdf = fdf[fdf[SENIORITY_COL].isin(seniority)]

    st.write(f"Filtered Rows: **{len(fdf)}**")

    st.divider()
    st.subheader("Branch Metrics (Count + Salary Averages)")

    default_branches = [c for c in ["information systems", "computer science"] if c in fdf.columns]
    branch_cols = st.multiselect(
        "Select branch canonical columns",
        options=fdf.columns.tolist(),
        default=default_branches,
    )

    metric = st.selectbox(
        "Y-axis metric",
        options=[
            "Count",
            "Avg Salary (Min)",
            "Avg Salary (Max)",
            "Avg Salary (Average)",
            "Median Salary (Min)",
            "Median Salary (Max)",
            "Min Salary (Min)",
            "Max Salary (Max)",
        ],
        index=0,
    )

    show_values_branch = st.checkbox("Show values on chart (branch metrics)", value=True)

    rows = []
    if branch_cols:
        for b in branch_cols:
            if b not in fdf.columns:
                continue

            mask = present(fdf[b])
            sub = fdf.loc[mask].copy()

            if metric == "Count":
                val = int(mask.sum())
            elif metric == "Avg Salary (Min)":
                val = safe_mean(sub["salary_min"]) if "salary_min" in sub.columns else None
            elif metric == "Avg Salary (Max)":
                val = safe_mean(sub["salary_max"]) if "salary_max" in sub.columns else None
            elif metric == "Avg Salary (Average)":
                a = safe_mean(sub["salary_min"]) if "salary_min" in sub.columns else None
                b = safe_mean(sub["salary_max"]) if "salary_max" in sub.columns else None
                val = float((a + b) / 2) if (a is not None and b is not None) else None
            elif metric == "Median Salary (Min)":
                val = safe_median(sub["salary_min"]) if "salary_min" in sub.columns else None
            elif metric == "Median Salary (Max)":
                val = safe_median(sub["salary_max"]) if "salary_max" in sub.columns else None
            elif metric == "Min Salary (Min)":
                val = safe_min(sub["salary_min"]) if "salary_min" in sub.columns else None
            elif metric == "Max Salary (Max)":
                val = safe_max(sub["salary_max"]) if "salary_max" in sub.columns else None
            else:
                val = None

            rows.append({"branch": b, "value": val})

        mdf = pd.DataFrame(rows).dropna(subset=["value"])
        if not mdf.empty:
            mdf = mdf.sort_values("value", ascending=False)
            fig_m = px.bar(mdf, x="branch", y="value")
            fig_m.update_yaxes(title=metric)

            if show_values_branch:
                fig_m.update_traces(texttemplate="%{y}", textposition="outside")
                fig_m.update_layout(uniformtext_minsize=8, uniformtext_mode="hide")

            st.plotly_chart(fig_m, use_container_width=True)
        else:
            st.warning("No valid data for the selected metric (e.g., salaries missing).")
    else:
        st.info("Select branch columns (example: information systems, computer science).")

    st.divider()
    st.subheader("Optional: Multi-X Chart (combination view)")

    enable_multi_x_chart = st.checkbox("Enable Multi-X Chart", value=False)

    if enable_multi_x_chart:
        st.caption("This section is for combinations (e.g., IS + CS together). It removes 'nan | nan' rows.")

        ai_q = st.text_input(
            "AI assist (optional): e.g. 'bar company_name salary_min salary_max' or 'line posted salary_min'",
            value="",
            key="ai_multi_x",
        )
        suggest = st.button("Suggest Chart", key="suggest_multi_x")

        default_chart, default_xcols, default_ys = guess_chart(fdf)

        if suggest:
            parsed = parse_ai_query(fdf, ai_q)
            if parsed:
                st.session_state["chart_type"], st.session_state["x_cols"], st.session_state["y_cols"] = parsed
            else:
                st.session_state["chart_type"], st.session_state["x_cols"], st.session_state["y_cols"] = (
                    default_chart,
                    default_xcols,
                    default_ys,
                )

        cols = fdf.columns.tolist()

        chart_type = st.selectbox(
            "Chart type",
            ["bar", "line", "scatter", "box", "histogram"],
            index=["bar", "line", "scatter", "box", "histogram"].index(st.session_state.get("chart_type", default_chart)),
            key="chart_type_sel",
        )

        x_default = st.session_state.get("x_cols", default_xcols)
        x_cols = st.multiselect(
            "X column(s) (multi-select)",
            cols,
            default=[c for c in x_default if c in cols],
            key="x_cols_sel",
        )

        y_candidates = [c for c in cols if c not in (x_cols or [])]
        y_default = st.session_state.get("y_cols", default_ys)
        y_cols = st.multiselect(
            "Y column(s) (multi-valued)",
            y_candidates,
            default=[c for c in y_default if c in y_candidates],
            key="y_cols_sel",
        )

        base_df = fdf.copy()
        base_df, x_key = build_multi_x(base_df, x_cols)

        if x_cols:
            mask_any_present = None
            for c in x_cols:
                m = present(base_df[c])
                mask_any_present = m if mask_any_present is None else (mask_any_present | m)
            if mask_any_present is not None:
                base_df = base_df[mask_any_present]

        if chart_type in ("bar", "line", "box") and not x_key:
            st.warning("Pick at least 1 X column.")
            st.stop()

        agg = st.selectbox(
            "Aggregation (when X has duplicates)",
            ["none", "mean", "median", "sum", "min", "max", "count"],
            key="agg_sel",
        )

        plot_df = base_df.copy()

        if agg != "none" and chart_type in ("bar", "line") and x_key:
            group = plot_df.groupby(x_key, dropna=False)

            if agg in ("mean", "median", "sum", "min", "max"):
                if not y_cols:
                    st.warning("Pick at least 1 Y column for numeric aggregation.")
                    st.stop()

                tmp = plot_df.copy()
                for c in y_cols:
                    tmp[c] = pd.to_numeric(tmp[c], errors="coerce")
                group = tmp.groupby(x_key, dropna=False)

                if agg == "mean":
                    plot_df = group[y_cols].mean(numeric_only=True).reset_index()
                elif agg == "median":
                    plot_df = group[y_cols].median(numeric_only=True).reset_index()
                elif agg == "sum":
                    plot_df = group[y_cols].sum(numeric_only=True).reset_index()
                elif agg == "min":
                    plot_df = group[y_cols].min(numeric_only=True).reset_index()
                elif agg == "max":
                    plot_df = group[y_cols].max(numeric_only=True).reset_index()

            elif agg == "count":
                plot_df = group.size().reset_index(name="count")
                y_cols = ["count"]

        if chart_type not in ("histogram", "scatter") and not y_cols:
            st.warning("Pick at least 1 Y column.")
            st.stop()

        st.subheader("Chart")
        show_values_on_chart = st.checkbox("Show values on chart", value=True, key="show_vals_multi")

        if chart_type == "bar":
            fig = px.bar(plot_df, x=x_key, y=y_cols, barmode="group")
        elif chart_type == "line":
            fig = px.line(plot_df, x=x_key, y=y_cols)
        elif chart_type == "scatter":
            if not y_cols:
                nums = [c for c in cols if pd.api.types.is_numeric_dtype(base_df[c])]
                if len(nums) >= 2:
                    fig = px.scatter(plot_df, x=nums[0], y=nums[1])
                else:
                    st.warning("Scatter needs at least 1 Y column (or 2 numeric columns).")
                    st.stop()
            else:
                fig = px.scatter(plot_df, x=x_key, y=y_cols[0])
        elif chart_type == "box":
            fig = px.box(plot_df, x=x_key, y=y_cols[0])
        elif chart_type == "histogram":
            if not x_key:
                st.warning("Pick at least 1 X column for histogram.")
                st.stop()
            fig = px.histogram(plot_df, x=x_key)
        else:
            fig = px.bar(plot_df, x=x_key, y=y_cols, barmode="group")

        if show_values_on_chart:
            fig = add_value_labels(fig, chart_type)

        st.plotly_chart(fig, use_container_width=True)


# =========================
# TAB 2: BASE SUMMARY
# =========================
with tab_summary:
    st.subheader("Branch / Role Salary Summary Dashboard (Base)")

    out_excel = st.text_input("Summary output Excel", value=SUMMARY_OUT_DEFAULT, key="out_excel_base")

    c1, _ = st.columns([1, 2])
    with c1:
        if st.button("Generate / Refresh Summary Excel", key="gen_summary_base"):
            try:
                created = generate_branch_salary_summary_excel_sliced(excel_path, out_excel_path=out_excel)
                st.cache_data.clear()
                st.success(f"Created: {created}")
            except Exception as e:
                st.error(str(e))

    if not Path(out_excel).exists():
        st.info("Click **Generate / Refresh Summary Excel** to create the summary file, then charts will appear here.")
        st.stop()

    try:
        sdf = load_branch_summary(out_excel, SUMMARY_SHEET)
    except Exception as e:
        st.error(f"Failed to read summary Excel: {e}")
        st.stop()

    # ✅ IMPORTANT: pass jobs_excel_path so the grouped bars can read degree_level / seniority / seniority_fields
    render_branch_summary_dashboard(sdf, render_key="base", jobs_excel_path=excel_path)