import os
import calendar
from datetime import datetime, timedelta

import pandas as pd
from flask import Flask, jsonify, render_template, request

app = Flask(__name__)

EXCEL_PATH = os.path.join(os.path.dirname(__file__), "RELATORIO - FLUXO DE CAIXA.xlsx")
_df: pd.DataFrame | None = None


def load_data() -> pd.DataFrame:
    global _df
    df = pd.read_excel(EXCEL_PATH, sheet_name="Dados")
    df.columns = [
        "Empresa", "Origem", "DataEntrada", "Vencimento", "UltPagto",
        "CodFor", "FornecedorCliente", "TituloTrans", "Valor", "Sit",
        "ContaFinanceira", "CentroCusto", "CodPRJ", "Projeto", "Observacao",
    ]
    df["Vencimento"] = pd.to_datetime(df["Vencimento"], errors="coerce")
    df["DataEntrada"] = pd.to_datetime(df["DataEntrada"], errors="coerce")
    df["UltPagto"] = pd.to_datetime(df["UltPagto"], errors="coerce")
    df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)
    df["ContaFinanceira"] = df["ContaFinanceira"].fillna("Sem Categoria").str.strip()
    df["Projeto"] = df["Projeto"].fillna("PROJETO PADRÃO").str.strip()
    df["Sit"] = df["Sit"].fillna("").str.strip()
    df["Origem"] = df["Origem"].fillna("").str.strip()

    # Effective cash-flow date: UltPagto for settled items, Vencimento otherwise
    df["DataEfetiva"] = df["Vencimento"].copy()
    mask_lq = df["Sit"].isin(["LQ", "CR", "DB"]) & df["UltPagto"].notna()
    df.loc[mask_lq, "DataEfetiva"] = df.loc[mask_lq, "UltPagto"]

    _df = df
    return df


load_data()


def _filter_project(df: pd.DataFrame, project: str) -> pd.DataFrame:
    if project and project != "all":
        return df[df["Projeto"] == project]
    return df


# ──────────────────────────────────────────────────────────────────────────────
# Routes
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/reload", methods=["POST"])
def reload_data():
    load_data()
    return jsonify({"status": "ok", "rows": len(_df)})


@app.route("/api/projects")
def get_projects():
    projects = sorted(_df["Projeto"].dropna().unique().tolist())
    return jsonify(projects)


@app.route("/api/months")
def get_months():
    months = (
        _df["Vencimento"].dropna()
        .dt.to_period("M")
        .astype(str)
        .unique()
        .tolist()
    )
    return jsonify(sorted(months))


@app.route("/api/dashboard")
def get_dashboard():
    project = request.args.get("project", "all")
    today = datetime.now().date()

    df = _filter_project(_df.copy(), project)

    # Settled (realised) and open (forecast) subsets
    df_lq = df[df["Sit"].isin(["LQ", "CR", "DB"])].copy()
    df_ab = df[df["Sit"] == "AB"].copy()

    df_rec_lq = df_lq[df_lq["Origem"] == "Contas a Receber"].copy()
    df_pag_lq = df_lq[df_lq["Origem"] == "Contas a Pagar"].copy()

    # Monthly breakdown – last 12 months realised
    periods = pd.period_range(end=pd.Period(today, "M"), periods=12, freq="M")

    def _monthly(df_sub: pd.DataFrame) -> dict:
        df_sub = df_sub.copy()
        df_sub["Mes"] = df_sub["DataEfetiva"].dt.to_period("M")
        grouped = df_sub.groupby("Mes")["Valor"].sum()
        return {str(p): float(grouped.get(p, 0)) for p in periods}

    monthly_recv = _monthly(df_rec_lq)
    monthly_pay = _monthly(df_pag_lq)

    # Overdue (AB, vencimento < today)
    df_atrasado = df_ab[df_ab["Vencimento"].dt.date < today]
    atraso_receber = float(df_atrasado[df_atrasado["Origem"] == "Contas a Receber"]["Valor"].sum())
    atraso_pagar = float(df_atrasado[df_atrasado["Origem"] == "Contas a Pagar"]["Valor"].sum())

    # Future (AB, vencimento >= today)
    df_futuro = df_ab[df_ab["Vencimento"].dt.date >= today]
    futuro_receber = float(df_futuro[df_futuro["Origem"] == "Contas a Receber"]["Valor"].sum())
    futuro_pagar = float(df_futuro[df_futuro["Origem"] == "Contas a Pagar"]["Valor"].sum())

    # Current month realised
    cur_period = pd.Period(today, "M")
    df_this = df_lq[df_lq["DataEfetiva"].dt.to_period("M") == cur_period]
    recv_mes = float(df_this[df_this["Origem"] == "Contas a Receber"]["Valor"].sum())
    pay_mes = float(df_this[df_this["Origem"] == "Contas a Pagar"]["Valor"].sum())

    # Next 30 days (AB)
    next30 = today + timedelta(days=30)
    df_n30 = df_ab[
        (df_ab["Vencimento"].dt.date >= today) &
        (df_ab["Vencimento"].dt.date <= next30)
    ]
    next30_rec = float(df_n30[df_n30["Origem"] == "Contas a Receber"]["Valor"].sum())
    next30_pag = float(df_n30[df_n30["Origem"] == "Contas a Pagar"]["Valor"].sum())

    # Top accounts
    def _top(df_sub: pd.DataFrame, n: int = 10) -> list:
        return (
            df_sub.groupby("ContaFinanceira")["Valor"].sum()
            .nlargest(n)
            .reset_index()
            .rename(columns={"ContaFinanceira": "conta", "Valor": "valor"})
            .assign(valor=lambda d: d["valor"].round(2))
            .to_dict("records")
        )

    # Monthly trend (all periods available) for chart
    df_rec_lq["MesP"] = df_rec_lq["DataEfetiva"].dt.to_period("M")
    df_pag_lq["MesP"] = df_pag_lq["DataEfetiva"].dt.to_period("M")

    return jsonify({
        "kpis": {
            "recv_mes": recv_mes,
            "pay_mes": pay_mes,
            "resultado_mes": recv_mes - pay_mes,
            "atraso_receber": atraso_receber,
            "atraso_pagar": atraso_pagar,
            "futuro_receber": futuro_receber,
            "futuro_pagar": futuro_pagar,
            "next30_receber": next30_rec,
            "next30_pagar": next30_pag,
        },
        "monthly_recv": monthly_recv,
        "monthly_pay": monthly_pay,
        "top_despesas": _top(df_pag_lq),
        "top_receitas": _top(df_rec_lq),
    })


@app.route("/api/daily-flow")
def get_daily_flow():
    month_str = request.args.get("month", "")
    project = request.args.get("project", "all")
    today = datetime.now().date()

    if not month_str:
        month_str = today.strftime("%Y-%m")

    year, month = int(month_str[:4]), int(month_str[5:7])
    days_in_month = calendar.monthrange(year, month)[1]
    days = list(range(1, days_in_month + 1))

    df = _filter_project(_df.copy(), project)

    # Realised: effective date falls in this month
    df_real = df[
        df["Sit"].isin(["LQ", "CR", "DB"]) &
        (df["DataEfetiva"].dt.year == year) &
        (df["DataEfetiva"].dt.month == month)
    ].copy()
    df_real["Dia"] = df_real["DataEfetiva"].dt.day

    # Forecast: due date falls in this month
    df_prev = df[
        (df["Sit"] == "AB") &
        (df["Vencimento"].dt.year == year) &
        (df["Vencimento"].dt.month == month)
    ].copy()
    df_prev["Dia"] = df_prev["Vencimento"].dt.day

    def _pivot(df_sub: pd.DataFrame, day_col: str = "Dia") -> dict:
        """Returns {conta: {day: value}}"""
        if df_sub.empty:
            return {}
        result: dict = {}
        for _, row in df_sub.iterrows():
            conta = row["ContaFinanceira"]
            dia = int(row[day_col])
            val = float(row["Valor"])
            result.setdefault(conta, {})
            result[conta][dia] = result[conta].get(dia, 0.0) + val
        return result

    real_recv = _pivot(df_real[df_real["Origem"] == "Contas a Receber"])
    real_pag = _pivot(df_real[df_real["Origem"] == "Contas a Pagar"])
    real_tes_cr = _pivot(df_real[(df_real["Origem"] == "Tesouraria") & (df_real["Sit"] == "CR")])
    real_tes_db = _pivot(df_real[(df_real["Origem"] == "Tesouraria") & (df_real["Sit"] == "DB")])
    prev_recv = _pivot(df_prev[df_prev["Origem"] == "Contas a Receber"])
    prev_pag = _pivot(df_prev[df_prev["Origem"] == "Contas a Pagar"])

    # Merge Tesouraria into the respective directions
    for conta, day_vals in real_tes_cr.items():
        real_recv.setdefault(conta, {})
        for d, v in day_vals.items():
            real_recv[conta][d] = real_recv[conta].get(d, 0.0) + v

    for conta, day_vals in real_tes_db.items():
        real_pag.setdefault(conta, {})
        for d, v in day_vals.items():
            real_pag[conta][d] = real_pag[conta].get(d, 0.0) + v

    def _build_rows(real_m: dict, prev_m: dict) -> list:
        all_contas = sorted(set(real_m) | set(prev_m))
        rows = []
        for conta in all_contas:
            real_vals = [real_m.get(conta, {}).get(d, 0.0) for d in days]
            prev_vals = [prev_m.get(conta, {}).get(d, 0.0) for d in days]
            rows.append({
                "conta": conta,
                "real": real_vals,
                "prev": prev_vals,
                "total_real": round(sum(real_vals), 2),
                "total_prev": round(sum(prev_vals), 2),
            })
        return rows

    recv_rows = _build_rows(real_recv, prev_recv)
    pag_rows = _build_rows(real_pag, prev_pag)

    n = len(days)

    def _col_sum(rows, key):
        return [round(sum(r[key][i] for r in rows), 2) for i in range(n)]

    recv_real_tot = _col_sum(recv_rows, "real")
    recv_prev_tot = _col_sum(recv_rows, "prev")
    pag_real_tot = _col_sum(pag_rows, "real")
    pag_prev_tot = _col_sum(pag_rows, "prev")

    saldo_dia = [
        round(recv_real_tot[i] + recv_prev_tot[i] - pag_real_tot[i] - pag_prev_tot[i], 2)
        for i in range(n)
    ]

    saldo_acum = []
    acc = 0.0
    for v in saldo_dia:
        acc += v
        saldo_acum.append(round(acc, 2))

    today_day = today.day if today.year == year and today.month == month else None

    return jsonify({
        "month": month_str,
        "days": days,
        "today_day": today_day,
        "recv": recv_rows,
        "pag": pag_rows,
        "summary": {
            "recv_real": recv_real_tot,
            "recv_prev": recv_prev_tot,
            "pag_real": pag_real_tot,
            "pag_prev": pag_prev_tot,
            "saldo_dia": saldo_dia,
            "saldo_acum": saldo_acum,
        },
    })


@app.route("/api/years")
def get_years():
    """Return sorted list of years present in the data."""
    years = set()
    if _df is not None:
        years.update(_df["Vencimento"].dropna().dt.year.unique().tolist())
        years.update(_df["DataEfetiva"].dropna().dt.year.unique().tolist())
    return jsonify(sorted(int(y) for y in years))


@app.route("/api/monthly-flow")
def get_monthly_flow():
    project = request.args.get("project", "all")
    year_filter = request.args.get("year", "")

    df = _filter_project(_df.copy(), project)

    df_real = df[df["Sit"].isin(["LQ", "CR", "DB"])].copy()
    df_prev = df[df["Sit"] == "AB"].copy()

    df_real["MesP"] = df_real["DataEfetiva"].dt.to_period("M")
    df_prev["MesP"] = df_prev["Vencimento"].dt.to_period("M")

    # All available months (optionally filtered by year)
    all_periods = set()
    if not df_real.empty:
        all_periods.update(df_real["MesP"].dropna().unique())
    if not df_prev.empty:
        all_periods.update(df_prev["MesP"].dropna().unique())

    if year_filter:
        all_periods = {p for p in all_periods if str(p).startswith(year_filter)}

    months = sorted(p for p in all_periods if not pd.isna(p))
    months_str = [str(m) for m in months]

    def _pivot_month(df_sub: pd.DataFrame) -> dict:
        """Returns {conta: {month_str: value}}"""
        if df_sub.empty:
            return {}
        grouped = df_sub.groupby(["ContaFinanceira", "MesP"])["Valor"].sum()
        result: dict = {}
        for (conta, period), val in grouped.items():
            result.setdefault(conta, {})[str(period)] = float(val)
        return result

    real_recv = _pivot_month(df_real[df_real["Origem"] == "Contas a Receber"])
    real_pag = _pivot_month(df_real[df_real["Origem"] == "Contas a Pagar"])
    real_tes_cr = _pivot_month(df_real[(df_real["Origem"] == "Tesouraria") & (df_real["Sit"] == "CR")])
    real_tes_db = _pivot_month(df_real[(df_real["Origem"] == "Tesouraria") & (df_real["Sit"] == "DB")])
    prev_recv = _pivot_month(df_prev[df_prev["Origem"] == "Contas a Receber"])
    prev_pag = _pivot_month(df_prev[df_prev["Origem"] == "Contas a Pagar"])

    for conta, mv in real_tes_cr.items():
        real_recv.setdefault(conta, {})
        for m, v in mv.items():
            real_recv[conta][m] = real_recv[conta].get(m, 0.0) + v

    for conta, mv in real_tes_db.items():
        real_pag.setdefault(conta, {})
        for m, v in mv.items():
            real_pag[conta][m] = real_pag[conta].get(m, 0.0) + v

    def _build_rows(real_m: dict, prev_m: dict) -> list:
        all_contas = sorted(set(real_m) | set(prev_m))
        rows = []
        for conta in all_contas:
            real_vals = [round(real_m.get(conta, {}).get(ms, 0.0), 2) for ms in months_str]
            prev_vals = [round(prev_m.get(conta, {}).get(ms, 0.0), 2) for ms in months_str]
            rows.append({
                "conta": conta,
                "real": real_vals,
                "prev": prev_vals,
                "total_real": round(sum(real_vals), 2),
                "total_prev": round(sum(prev_vals), 2),
            })
        return rows

    recv_rows = _build_rows(real_recv, prev_recv)
    pag_rows = _build_rows(real_pag, prev_pag)

    n = len(months_str)

    def _col_sum(rows, key):
        return [round(sum(r[key][i] for r in rows), 2) for i in range(n)]

    return jsonify({
        "months": months_str,
        "recv": recv_rows,
        "pag": pag_rows,
        "summary": {
            "recv_real": _col_sum(recv_rows, "real"),
            "recv_prev": _col_sum(recv_rows, "prev"),
            "pag_real": _col_sum(pag_rows, "real"),
            "pag_prev": _col_sum(pag_rows, "prev"),
        },
    })


if __name__ == "__main__":
    app.run(debug=True, port=5000, use_reloader=False)
