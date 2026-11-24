"""
app.py — 电网导出数据整理

上传供电所/电网的 96点或24点累计读数导出表, 转成逐区间电量/功率。

界面分三块:
    侧边栏  上传文件 + 参数(表单, 改参数不触发重算)
    主区    结果分 4 个页签: 运行报告 / 曲线 / 每日汇总 / 区间明细
    底部    导出按钮

"""
from __future__ import annotations

import os
import tempfile

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from core import (build_intervals, build_labels, build_workbook, list_sheets,
                  mode_curve, probe, resolve_multipliers, select_rows,
                  summarize, to_float)

st.set_page_config(page_title="电网导出数据整理", layout="wide", page_icon="⚡")
st.title("⚡ 电网导出数据整理")
st.caption("累计读数 → 逐区间电量/功率。自动识别 15分钟/1小时 步长与表结构。")


@st.cache_data(show_spinner=False)
def cached_probe(path: str, sheet: str | None, _mtime: float):
    return probe(path, sheet)


@st.cache_data(show_spinner=False)
def cached_sheets(path: str, _mtime: float):
    return list_sheets(path)


# ================= 侧边栏: 文件 =================
with st.sidebar:
    st.header("📂 数据文件")
    uploaded = st.file_uploader(
        "上传导出表", type=["xls", "xlsx", "csv"],
        help="支持真 .xls (BIFF) / .xlsx / .csv。",
    )
    path, fname = None, None
    if uploaded is not None:
        path = os.path.join(tempfile.gettempdir(), "grid_clean_" + uploaded.name)
        with open(path, "wb") as f:
            f.write(uploaded.getvalue())
        fname = uploaded.name

    sheet_name = None
    if path:
        try:
            sheets = cached_sheets(path, os.path.getmtime(path))
            if len(sheets) > 1:
                sheet_name = st.selectbox("工作表", sheets)
        except Exception as e:
            st.error(f"读取失败: {e}")

    info = None
    if path:
        try:
            info = cached_probe(path, sheet_name, os.path.getmtime(path))
        except Exception as e:
            st.error(f"探测失败: {e}")

    # ================= 侧边栏: 参数 =================
    st.divider()
    st.header("⚙️ 处理参数")

    if info:
        st.caption(f"📐 {info['container']} | {info['nrows']} 行 × {info['ncols']} 列")

        with st.form("params", border=False):
            # 电表: 多表文件必须先选一个表再差分。
            meter_id = None
            meters = info["meters"]
            if len(meters) == 1:
                st.caption(f"🔌 只有一个电表: {meters[0]}")
                meter_id = meters[0]
            elif len(meters) > 1:
                st.caption(f"🔌 本文件含 {len(meters)} 个电表, 请选一个处理")
                meter_id = st.selectbox("选用电表", meters)
            dtype = st.selectbox("数据类型", ["全部"] + info["dtypes"])

            with st.expander("倍率与量纲", expanded=True):
                ratio_const = None
                if info["ratio_col"] is not None:
                    st.caption(f"表内『综合倍率』列: 第 {info['ratio_col'] + 1} 列")
                    if st.checkbox("改用固定倍率常数", value=False):
                        ratio_const = st.number_input("倍率常数", value=1.0)
                else:
                    st.caption("表内无『综合倍率』列, 需手填")
                    ratio_const = st.number_input("倍率常数", value=1.0)

                want_kw = st.radio(
                    "输出量纲", ["平均功率 kW", "区间电量 kWh"], horizontal=True,
                    help="kW: 再乘 60/step (15min→×4, 1h→×1), 数值即该区间平均功率。\n\n"
                         "kWh: 不再乘, 数值即该区间电量。",
                ) == "平均功率 kW"

            with st.expander("跨日差放哪一格", expanded=True):
                anchor = st.radio(
                    "跨日差放哪一格", ["末格 (left)", "首格 (right)"], horizontal=True,
                    help="末格: 每行最后一格放跨日差, 表首格是第一个减完的数。\n\n"
                         "首格: 每行第一格放跨日差, 整表首格 = 0。\n\n"
                         "混用错位。",
                )
                anchor = "left" if anchor.startswith("末格") else "right"

            submitted = st.form_submit_button("▶️ 开始处理", type="primary", width='stretch')
    else:
        st.caption("先上传一个导出表")
        submitted = False


# ================= 处理 =================
if submitted and path and info:
    t0, t1 = info["time_block"]
    if t0 is None:
        st.error("表头里没找到连续的时间列 (HH:MM), 无法处理")
    else:
        step = info["step"]
        n_pts = info["n_intervals"]
        # 自检: 一天的点数 x 步长 应该正好铺满 1440 分钟。
        # 对不上说明表头标签不全或步长识别有误 —— 这时标签和 60/step 都会算错，break。
        covered = n_pts * step
        if covered != 1440:
            st.error(
                f"⚠️ 步长自检不通过: {n_pts} 个时间点 × {step} 分钟 = {covered} 分钟, "
                f"应为 1440 分钟 (一天)。表头标签可能不全或步长识别有误, "
                f"继续处理会得到错误的区间标签和电量。请检查表头。"
            )
            st.stop()
        rows, stats = select_rows(
            info["body"], meter_col=info["meter_col"], meter_id=meter_id,
            dtype_col=info["dtype_col"], dtype=dtype,
        )
        if not rows:
            st.error("筛选后没有数据行")
        else:
            raw = [[to_float(r[c]) for c in range(t0, t1 + 1)] for r in rows]
            mults, mult_src = resolve_multipliers(rows, info["ratio_col"], ratio_const)

            pf = 60.0 / step if want_kw else 1.0
            out = build_intervals(raw, mults, anchor=anchor, power_factor=pf)
            labels = build_labels(len(raw[0]), step, anchor)
            daily, miss_info, negs = summarize(out, labels, step_min=step,
                                               values_are_power=want_kw)

            st.session_state["job"] = {
                "info": info, "rows": rows, "stats": stats, "raw": raw, "out": out,
                "labels": labels, "daily": daily, "miss": miss_info, "negs": negs,
                "step": step, "pf": pf, "anchor": anchor, "mult_src": mult_src,
                "ratio_const": ratio_const, "want_kw": want_kw, "fname": fname,
                "meter_id": meter_id, "dtype": dtype, "t0": t0, "t1": t1,
            }


# ================= 结果 =================
job = st.session_state.get("job")
if not job:
    st.info("👈 上传导出表, 在左侧调好参数后点 **开始处理**。")
    st.stop()

info = job["info"]
step = job["step"]
labels = job["labels"]
daily = job["daily"]
out = job["out"]
rows = job["rows"]
unit = "kW" if job["want_kw"] else "kWh"
day_names = [rows[i][info["date_col"]] if info["date_col"] is not None else i + 1
             for i in range(len(rows))]

tab_rep, tab_curve, tab_daily, tab_detail = st.tabs(
    ["📋 运行报告", "📈 曲线", "📊 每日汇总", "🔍 区间明细"])

# ---------- 运行报告 ----------
with tab_rep:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("原始行数", job["stats"]["total"])
    c2.metric("筛选后行数", job["stats"]["after_dtype"])
    c3.metric("步长", f"{step} 分钟")
    c4.metric("区间列数", len(labels))

    st.markdown(f"""
| 项 | 值 |
|---|---|
| 容器格式 | **{info['container']}** (按文件头判断) |
| 时间列块 | 第 {job['t0'] + 1} ~ {job['t1'] + 1} 列, 共 {len(labels)} 点 |
| 数据性质 | 累计读数 (已差分) |
| 倍率来源 | {job['mult_src']} |
| 计算式 | `Δ读数 × 综合倍率 × {job['pf']:g}` |
| 每格含义 | **{'平均功率 kW' if job['want_kw'] else '区间电量 kWh'}** (1格 = {step} 分钟) |
| 日电量 | 各格之和{' ÷ ' + format(job['pf'], 'g') if job['want_kw'] else ''} |
| 跨日差锚点 | **{job['anchor']}** ({'每行末格装跨日差' if job['anchor'] == 'left' else '每行首格装跨日差, 整表首格=0'}) |
| 首格 | **{out[0][0]}** ({'第一个减完的数' if job['anchor'] == 'left' else '0'}) |
| 电表筛选 | {job['meter_id'] or '不筛选'} |
| 数据类型 | {job['dtype']} |
| 日期跨度 | {info['date_first']} → {info['date_last']} |
""")

    miss_rows = [(i, day_names[i], m) for i, m in enumerate(job["miss"]) if m]
    with st.expander(f"⚠️ 缺失与异常 ({len(miss_rows)} 行有缺失, {len(job['negs'])} 个负值)",
                     expanded=bool(miss_rows)):
        if miss_rows:
            st.write("**缺失区间** (原表读数为 `--`, 未修改, 当日汇总留空不做部分和):")
            st.dataframe(pd.DataFrame(miss_rows, columns=["行号", "日期", "缺失格数"]),
                         hide_index=True, width='stretch')
        if job["negs"]:
            st.write("**负值区间** (照实保留):")
            st.dataframe(pd.DataFrame([(i + 1, j + 1, lb, v) for i, j, lb, v in job["negs"]],
                                      columns=["行号", "列号", "区间", "值"]),
                         hide_index=True, width='stretch')

# ---------- 曲线 ----------
with tab_curve:
    prof = []
    for j in range(len(labels)):
        vs = [out[i][j] for i in range(len(out)) if out[i][j] is not None]
        if vs:
            prof.append((labels[j], min(vs), max(vs)))
    df_prof = pd.DataFrame(prof, columns=["区间", "最小", "最大"])

    def base_layout(fig, h=340):
        fig.update_layout(
            height=h,
            hovermode="x unified",          # 同一时刻各条线一起显示
            dragmode="zoom",
            margin=dict(l=8, r=8, t=34, b=8),
            legend=dict(orientation="h", yanchor="bottom", y=1.02,
                        xanchor="left", x=0, bgcolor="rgba(0,0,0,0)"),
            xaxis=dict(showgrid=False, title=None),
            yaxis=dict(title=f"区间{unit}", gridcolor="rgba(128,128,128,0.15)"),
        )
        return fig

    # 1) 典型日曲线: 阴影带 = 每日最小~最大, 实线 = 众数
    modes, shares = mode_curve(out)
    avg_share = sum(s for s in shares if s) / max(1, len([s for s in shares if s]))

    fig_band = go.Figure()
    fig_band.add_trace(go.Scatter(
        x=df_prof["区间"], y=df_prof["最大"], mode="lines",
        line=dict(width=0), showlegend=False, hoverinfo="skip"))
    fig_band.add_trace(go.Scatter(
        x=df_prof["区间"], y=df_prof["最小"], mode="lines",
        line=dict(width=0), fill="tonexty", fillcolor="rgba(76,120,168,0.22)",
        name="每日最小~最大", hoverinfo="skip"))
    fig_band.add_trace(go.Scatter(
        x=df_prof["区间"], y=modes, mode="lines+markers",
        line=dict(color="#4C78A8", width=2), marker=dict(size=5),
        name="众数",
        hovertemplate="%{y:,.1f} " + unit + "<extra></extra>"))

    st.subheader("典型日曲线")
    st.caption("阴影带 = 每天的最小~最大范围, 实线 = 众数 ( "
               f"平均覆盖 {avg_share * 100:.0f}% 的天数)。鼠标悬停看数值, "
               "框选可放大, 双击还原。")
    st.plotly_chart(base_layout(fig_band), width='stretch',
                    config={"displaylogo": False, "scrollZoom": True,
                            "modeBarButtonsToRemove": ["lasso2d", "select2d"]})

    # 2) 单日曲线叠加对比
    pick = st.selectbox("选一天看原始曲线 (与众数曲线对比)", list(range(len(rows))),
                        format_func=lambda i: f"{day_names[i]}  (第 {i + 1} 行)")
    df_day = pd.DataFrame({"区间": labels, "值": out[pick]})
    df_day["值"] = pd.to_numeric(df_day["值"], errors="coerce")
    df_day = df_day.dropna(subset=["值"])

    fig_day = go.Figure()
    for tr in fig_band.data:
        fig_day.add_trace(tr)
    fig_day.add_trace(go.Scatter(
        x=df_day["区间"], y=df_day["值"], mode="lines+markers",
        line=dict(color="#F58518", width=1.8), marker=dict(size=4),
        name=f"{day_names[pick]} (选中日)",
        hovertemplate="%{y:,.1f} " + unit + "<extra></extra>"))

    st.subheader(f"叠加对比: {day_names[pick]}")
    st.plotly_chart(base_layout(fig_day, h=320), width='stretch',
                    config={"displaylogo": False, "scrollZoom": True,
                            "modeBarButtonsToRemove": ["lasso2d", "select2d"]})

    # 3) 日电量趋势 (柱状)
    dd_all = pd.DataFrame([{"日期": day_names[i], "日电量(kWh)": d[0]}
                           for i, d in enumerate(daily)])
    dd_ok = dd_all.dropna(subset=["日电量(kWh)"])
    if len(dd_ok):
        fig_bar = go.Figure()
        fig_bar.add_trace(go.Bar(
            x=dd_ok["日期"], y=dd_ok["日电量(kWh)"], name="日电量",
            marker=dict(color="#4C78A8"),
            hovertemplate="%{x}<br>%{y:,.1f} kWh<extra></extra>"))

        fig_bar.update_layout(
            height=300, barmode="overlay", hovermode="x unified",
            dragmode="zoom",
            margin=dict(l=8, r=8, t=34, b=8),
            legend=dict(orientation="h", yanchor="bottom", y=1.02,
                        xanchor="left", x=0, bgcolor="rgba(0,0,0,0)"),
            xaxis=dict(title=None, type="category"),
            yaxis=dict(title="日电量 kWh", gridcolor="rgba(128,128,128,0.15)"),
        )
        st.subheader("日电量趋势")
        st.caption("柱 = 每日电量。框选可放大某段时间, 滚轮缩放。")
        st.plotly_chart(fig_bar, width='stretch',
                        config={"displaylogo": False, "scrollZoom": True,
                                "modeBarButtonsToRemove": ["lasso2d", "select2d"]})

# ---------- 每日汇总 ----------
with tab_daily:
    dd = pd.DataFrame([
        {"日期": day_names[i], "日电量(kWh)": d[0],
         "日均功率(kW)": d[1], "缺失格数": d[2]}
        for i, d in enumerate(daily)
    ])
    ok = [d[0] for d in daily if d[0] is not None]
    if ok:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("日电量 均值", f"{sum(ok) / len(ok):,.0f} kWh")
        c2.metric("日电量 最小", f"{min(ok):,.0f} kWh")
        c3.metric("日电量 最大", f"{max(ok):,.0f} kWh")
        c4.metric("完整天数", f"{len(ok)}/{len(daily)}")
    st.dataframe(dd, hide_index=True, width='stretch')

# ---------- 区间明细 ----------
with tab_detail:
    prev = pd.DataFrame(out, columns=labels)
    prev.insert(0, "日期", day_names)
    st.dataframe(prev, hide_index=True, width='stretch')

# ================= 导出 =================
st.divider()
out_name = os.path.splitext(job["fname"] or "result")[0] + "_整理结果.xlsx"
c1, c2 = st.columns([1, 3])
with c1:
    st.download_button("💾 下载整理结果 xlsx", build_workbook(job),
                       file_name=out_name,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")
with c2:
    st.caption("输出 2 个 sheet: 区间值 (处理结果, 右侧带日汇总列) / 原始读数")
