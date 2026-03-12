import io
import zipfile

import numpy as np
import pandas as pd
import streamlit as st

# ─────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="模考成績分析系統",
    page_icon="📊",
    layout="centered",
)

st.title("📊 三次模考成績分析系統")
st.caption("上傳三次模考的 Excel 原始檔，系統將自動產生成績分析 CSV 與學生個人成績分析單。")

# ─────────────────────────────────────────────
# Sidebar — file uploaders
# ─────────────────────────────────────────────
with st.sidebar:
    st.header("📂 上傳檔案")
    file1 = st.file_uploader("第一次模考 Excel", type=["xlsx", "xls"], key="f1")
    file2 = st.file_uploader("第二次模考 Excel", type=["xlsx", "xls"], key="f2")
    file3 = st.file_uploader("第三次模考 Excel", type=["xlsx", "xls"], key="f3")
    class_name = st.text_input("班級名稱", value="922班", help="將顯示在報告標頭及檔名中")
    run_btn = st.button("🚀 執行分析", type="primary", use_container_width=True,
                        disabled=not (file1 and file2 and file3))

# ─────────────────────────────────────────────
# Parser functions (same logic as 模考檔案轉csv檔.py)
# ─────────────────────────────────────────────
def parse_exam1(file_obj, prefix):
    df = pd.read_excel(file_obj, header=None)
    data = df.iloc[9:].reset_index(drop=True)
    data = data[data.iloc[:, 2].notna() & (data.iloc[:, 2].astype(str).str.strip() != '')]
    r = pd.DataFrame()
    r['座號'] = data.iloc[:, 2].astype(str).str.strip()
    r['姓名'] = data.iloc[:, 3].astype(str).str.strip()
    r[f'{prefix}_總分'] = pd.to_numeric(data.iloc[:, 4], errors='coerce')
    r[f'{prefix}_國文'] = pd.to_numeric(data.iloc[:, 8], errors='coerce')
    r[f'{prefix}_數學'] = pd.to_numeric(data.iloc[:, 15], errors='coerce')
    r[f'{prefix}_英文'] = pd.to_numeric(data.iloc[:, 22], errors='coerce')
    r[f'{prefix}_社會'] = pd.to_numeric(data.iloc[:, 25], errors='coerce')
    r[f'{prefix}_自然'] = pd.to_numeric(data.iloc[:, 28], errors='coerce')
    r[f'{prefix}_寫作'] = pd.to_numeric(data.iloc[:, 30], errors='coerce')
    r[f'{prefix}_班排'] = pd.to_numeric(data.iloc[:, 31], errors='coerce')
    r[f'{prefix}_校排'] = pd.to_numeric(data.iloc[:, 32], errors='coerce')
    r[f'{prefix}_區排'] = pd.to_numeric(data.iloc[:, 33], errors='coerce')
    return r


def parse_exam2(file_obj, prefix):
    df = pd.read_excel(file_obj, header=None)
    data = df.iloc[8:].reset_index(drop=True)
    data = data[data.iloc[:, 0].notna() & (data.iloc[:, 0].astype(str).str.strip() != '')]
    r = pd.DataFrame()
    r['座號'] = pd.to_numeric(data.iloc[:, 0], errors='coerce').fillna(0).astype(int).astype(str).str.zfill(2)
    r['姓名'] = data.iloc[:, 1].astype(str).str.strip()
    r[f'{prefix}_總分'] = pd.to_numeric(data.iloc[:, 5], errors='coerce')
    r[f'{prefix}_國文'] = pd.to_numeric(data.iloc[:, 11], errors='coerce')
    r[f'{prefix}_數學'] = pd.to_numeric(data.iloc[:, 19], errors='coerce')
    r[f'{prefix}_英文'] = pd.to_numeric(data.iloc[:, 27], errors='coerce')
    r[f'{prefix}_社會'] = pd.to_numeric(data.iloc[:, 31], errors='coerce')
    r[f'{prefix}_自然'] = pd.to_numeric(data.iloc[:, 35], errors='coerce')
    total_with_writing = pd.to_numeric(data.iloc[:, 5], errors='coerce')
    five_subject = pd.to_numeric(data.iloc[:, 3], errors='coerce')
    r[f'{prefix}_寫作'] = (total_with_writing - five_subject).round(1)
    r[f'{prefix}_班排'] = pd.to_numeric(data.iloc[:, 6], errors='coerce')
    r[f'{prefix}_校排'] = pd.to_numeric(data.iloc[:, 7], errors='coerce')
    r[f'{prefix}_區排'] = pd.to_numeric(data.iloc[:, 38], errors='coerce')
    return r


def parse_exam3(file_obj, prefix):
    df = pd.read_excel(file_obj, header=None)
    data = df.iloc[7:].reset_index(drop=True)
    data = data[data.iloc[:, 1].notna() & (data.iloc[:, 1].astype(str).str.strip() != '')]
    r = pd.DataFrame()
    r['座號'] = data.iloc[:, 1].astype(str).str.strip()
    r['姓名'] = data.iloc[:, 2].astype(str).str.strip()
    r[f'{prefix}_總分'] = pd.to_numeric(data.iloc[:, 5], errors='coerce')
    r[f'{prefix}_國文'] = pd.to_numeric(data.iloc[:, 8], errors='coerce')
    r[f'{prefix}_數學'] = pd.to_numeric(data.iloc[:, 13], errors='coerce')
    r[f'{prefix}_英文'] = pd.to_numeric(data.iloc[:, 20], errors='coerce')
    r[f'{prefix}_社會'] = pd.to_numeric(data.iloc[:, 23], errors='coerce')
    r[f'{prefix}_自然'] = pd.to_numeric(data.iloc[:, 26], errors='coerce')
    r[f'{prefix}_寫作'] = pd.to_numeric(data.iloc[:, 28], errors='coerce')
    r[f'{prefix}_班排'] = pd.to_numeric(data.iloc[:, 29], errors='coerce')
    r[f'{prefix}_校排'] = pd.to_numeric(data.iloc[:, 30], errors='coerce')
    r[f'{prefix}_區排'] = pd.to_numeric(data.iloc[:, 31], errors='coerce')
    return r


def build_merged_df(f1, f2, f3):
    parsers = [(f1, '一模', parse_exam1), (f2, '二模', parse_exam2), (f3, '三模', parse_exam3)]
    merged = None
    for fobj, prefix, parser in parsers:
        df = parser(fobj, prefix)
        df['座號'] = pd.to_numeric(df['座號'], errors='coerce').fillna(0).astype(int).astype(str).str.zfill(2)
        df = df[df['座號'] != '00'].dropna(subset=['姓名'])
        df = df[df['姓名'].str.strip() != 'nan']
        merged = df if merged is None else pd.merge(merged, df, on=['座號', '姓名'], how='outer')

    merged = merged.sort_values('座號').reset_index(drop=True)

    subjects = ['總分', '國文', '英文', '數學', '社會', '自然', '寫作']
    for subj in subjects:
        if f'三模_{subj}' in merged.columns and f'二模_{subj}' in merged.columns:
            merged[f'近期{subj}變化'] = merged[f'三模_{subj}'] - merged[f'二模_{subj}']
        if f'三模_{subj}' in merged.columns and f'一模_{subj}' in merged.columns:
            merged[f'總體{subj}成長'] = merged[f'三模_{subj}'] - merged[f'一模_{subj}']
    for r in ['班排', '校排', '區排']:
        if f'三模_{r}' in merged.columns and f'二模_{r}' in merged.columns:
            merged[f'近期{r}進步'] = merged[f'二模_{r}'] - merged[f'三模_{r}']
        if f'三模_{r}' in merged.columns and f'一模_{r}' in merged.columns:
            merged[f'總體{r}進步'] = merged[f'一模_{r}'] - merged[f'三模_{r}']
    return merged


# ─────────────────────────────────────────────
# HTML report helpers (same logic as 產生學生成績分析單.py)
# ─────────────────────────────────────────────
EXAMS = ['一模', '二模', '三模']
SUBJECTS = ['國文', '數學', '英文', '社會', '自然']
RANKS = ['班排', '校排', '區排']
RANK_LABELS = {'班排': '班排名', '校排': '校排名', '區排': '區排名'}


def arrow(val):
    if pd.isna(val):
        return '<span style="color:#999">—</span>'
    v = float(val)
    if v > 0:
        return f'<span class="up">▲ {v:+.1f}</span>'
    elif v < 0:
        return f'<span class="down">▼ {v:.1f}</span>'
    return '<span class="flat">→ 0</span>'


def rank_arrow(val):
    if pd.isna(val):
        return '<span style="color:#999">—</span>'
    v = float(val)
    if v > 0:
        return f'<span class="up">▲ 進步 {v:.0f}</span>'
    elif v < 0:
        return f'<span class="down">▼ 退步 {abs(v):.0f}</span>'
    return '<span class="flat">→ 持平</span>'


def fmt(val, decimals=1):
    if pd.isna(val):
        return '—'
    return f'{float(val):.{decimals}f}'


def make_svg(values, labels, color, lower_is_better=False, width=260, height=100):
    pad_l, pad_r, pad_t, pad_b = 30, 30, 18, 18
    chart_w, chart_h = width - pad_l - pad_r, height - pad_t - pad_b
    valid = [(i, v) for i, v in enumerate(values) if not pd.isna(v)]
    if not valid:
        return f'<svg viewBox="0 0 {width} {height}" width="100%"><text x="50%" y="50%" dominant-baseline="middle" text-anchor="middle" fill="#aaa" font-size="11">無資料</text></svg>'
    vals = [v for _, v in valid]
    vmin, vmax = min(vals), max(vals)
    if vmin == vmax:
        vmin -= 1; vmax += 1
    n = len(labels)
    xs = [pad_l + i * chart_w / (n - 1) if n > 1 else pad_l + chart_w / 2 for i in range(n)]

    def to_y(v):
        return pad_t + (v - vmin) / (vmax - vmin) * chart_h if lower_is_better \
            else pad_t + (1 - (v - vmin) / (vmax - vmin)) * chart_h

    points = [(xs[i], to_y(v)) for i, v in valid]
    polyline = ' '.join(f'{x:.1f},{y:.1f}' for x, y in points)
    dots = ''
    for (idx, v), (x, y) in zip(valid, points):
        dots += f'<circle cx="{x:.1f}" cy="{y:.1f}" r="4" fill="{color}" stroke="white" stroke-width="1.5"/>'
        ly = y - 7 if y > pad_t + 12 else y + 16
        dots += f'<text x="{x:.1f}" y="{ly:.1f}" text-anchor="middle" font-size="10" fill="{color}" font-weight="600">{fmt(v)}</text>'
    xlbls = ''.join(f'<text x="{xs[i]:.1f}" y="{height-2}" text-anchor="middle" font-size="10" fill="#888">{lb}</text>'
                    for i, lb in enumerate(labels))
    return f'<svg viewBox="0 0 {width} {height}" width="100%" xmlns="http://www.w3.org/2000/svg"><polyline points="{polyline}" fill="none" stroke="{color}" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round"/>{dots}{xlbls}</svg>'


CSS = """
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;600;700;900&display=swap');
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Noto Sans TC', sans-serif; background: #f5f6fa; padding: 24px; color: #1a1a2e; font-size: 13px; }
  .card { background: white; border-radius: 12px; box-shadow: 0 2px 16px rgba(0,0,0,0.09); max-width: 820px; margin: 0 auto; padding: 28px 32px; }
  .header { display: flex; align-items: center; justify-content: space-between; border-bottom: 3px solid #4e6ef2; padding-bottom: 14px; margin-bottom: 22px; }
  .header-left h1 { font-size: 22px; font-weight: 900; color: #1a1a2e; }
  .header-left .subtitle { font-size: 12px; color: #888; margin-top: 3px; }
  .header-badge { background: linear-gradient(135deg, #4e6ef2, #7b5ea7); color: white; border-radius: 8px; padding: 8px 18px; text-align: center; }
  .header-badge .seat { font-size: 11px; opacity: 0.85; }
  .header-badge .name { font-size: 20px; font-weight: 900; letter-spacing: 2px; }
  .section-title { font-size: 14px; font-weight: 700; color: #4e6ef2; border-left: 4px solid #4e6ef2; padding-left: 10px; margin: 22px 0 12px; }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  th { background: #f0f3ff; color: #4e6ef2; font-weight: 700; padding: 8px 10px; text-align: center; border-bottom: 2px solid #c8d3f7; }
  td { padding: 7px 10px; text-align: center; border-bottom: 1px solid #eef0f8; }
  td.label { text-align: left; font-weight: 600; color: #333; background: #fcfcff; }
  tr.total-row td { font-weight: 700; background: #f0f3ff; font-size: 13.5px; }
  .up { color: #e05c5c; font-weight: 700; }
  .down { color: #27ae60; font-weight: 700; }
  .flat { color: #888; }
  .charts-grid { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 6px; }
  .chart-item { background: #f9faff; border: 1px solid #e8ecff; border-radius: 8px; padding: 8px 8px 4px; flex: 1 1 140px; min-width: 0; overflow: hidden; }
  .chart-item svg { display: block; width: 100%; }
  .chart-title { font-size: 12px; font-weight: 700; margin-bottom: 4px; text-align: center; }
  .note { font-size: 11px; color: #999; margin-top: 16px; text-align: right; }
  .note span { color: #e05c5c; font-weight: 600; }
  .note .green { color: #27ae60; }
  @media print {
    body { background: white; padding: 0; }
    .card { box-shadow: none; max-width: 100%; page-break-after: always; }
    .card:last-child { page-break-after: avoid; }
  }
"""


def generate_card(row, class_name="922班"):
    sid, name = row['座號'], row['姓名']
    today = pd.Timestamp.now().strftime("%Y年%m月%d日")

    score_rows = ''
    for subj in SUBJECTS:
        vals = [row.get(f'{e}_{subj}', np.nan) for e in EXAMS]
        score_rows += f'<tr><td class="label">{subj}</td><td>{fmt(vals[0])}</td><td>{fmt(vals[1])}</td><td>{fmt(vals[2])}</td><td>{arrow(row.get(f"近期{subj}變化", np.nan))}</td><td>{arrow(row.get(f"總體{subj}成長", np.nan))}</td></tr>'

    t_vals = [row.get(f'{e}_總分', np.nan) for e in EXAMS]
    total_row = f'<tr class="total-row"><td class="label">總分(含寫作)</td><td>{fmt(t_vals[0])}</td><td>{fmt(t_vals[1])}</td><td>{fmt(t_vals[2])}</td><td>{arrow(row.get("近期總分變化", np.nan))}</td><td>{arrow(row.get("總體總分成長", np.nan))}</td></tr>'

    w_vals = [row.get(f'{e}_寫作', np.nan) for e in EXAMS]
    writing_row = f'<tr><td class="label">寫作</td><td>{fmt(w_vals[0])}</td><td>{fmt(w_vals[1])}</td><td>{fmt(w_vals[2])}</td><td>{arrow(row.get("近期寫作變化", np.nan))}</td><td>{arrow(row.get("總體寫作成長", np.nan))}</td></tr>'

    colors = ['#4e6ef2', '#e05c5c', '#27ae60', '#e67e22', '#8e44ad']
    subj_charts = ''
    for subj, color in zip(SUBJECTS, colors):
        vals = [row.get(f'{e}_{subj}', np.nan) for e in EXAMS]
        subj_charts += f'<div class="chart-item"><div class="chart-title" style="color:{color}">{subj}</div>{make_svg(vals, ["一模","二模","三模"], color)}</div>'

    t_chart = make_svg(t_vals, ['一模', '二模', '三模'], '#1a1a2e', width=300, height=110)

    rcc = {'班排': '#c0392b', '校排': '#2980b9', '區排': '#16a085'}
    rank_rows = ''
    rank_charts = ''
    for r in RANKS:
        r_vals = [row.get(f'{e}_{r}', np.nan) for e in EXAMS]
        rank_rows += f'<tr><td class="label">{RANK_LABELS[r]}</td><td>{fmt(r_vals[0],0)}</td><td>{fmt(r_vals[1],0)}</td><td>{fmt(r_vals[2],0)}</td><td>{rank_arrow(row.get(f"近期{r}進步", np.nan))}</td><td>{rank_arrow(row.get(f"總體{r}進步", np.nan))}</td></tr>'
        rank_charts += f'<div class="chart-item"><div class="chart-title" style="color:{rcc[r]}">{RANK_LABELS[r]}</div>{make_svg(r_vals, ["一模","二模","三模"], rcc[r], lower_is_better=True)}</div>'

    return f"""<div class="card">
  <div class="header">
    <div class="header-left"><h1>三次模考個人成績分析單</h1><div class="subtitle">{class_name} ｜ 分析日期：{today}</div></div>
    <div class="header-badge"><div class="seat">座號 {sid}</div><div class="name">{name}</div></div>
  </div>
  <div class="section-title">📊 各科成績比較</div>
  <table><thead><tr><th style="text-align:left;width:110px">科目</th><th>第一次模考</th><th>第二次模考</th><th>第三次模考</th><th>近期變化<br><small style="font-weight:400;color:#888">(三模－二模)</small></th><th>整體成長<br><small style="font-weight:400;color:#888">(三模－一模)</small></th></tr></thead>
  <tbody>{score_rows}{writing_row}{total_row}</tbody></table>
  <div class="section-title">📈 各科積分趨勢</div>
  <div class="charts-grid">{subj_charts}</div>
  <div class="section-title">🏆 總積分趨勢</div>
  <div class="charts-grid"><div class="chart-item" style="flex:1"><div class="chart-title" style="color:#1a1a2e">含寫作總積分</div>{t_chart}</div></div>
  <div class="section-title">🎯 排名比較</div>
  <table><thead><tr><th style="text-align:left;width:110px">排名類型</th><th>第一次模考</th><th>第二次模考</th><th>第三次模考</th><th>近期進退步<br><small style="font-weight:400;color:#888">(二模排－三模排)</small></th><th>整體進退步<br><small style="font-weight:400;color:#888">(一模排－三模排)</small></th></tr></thead>
  <tbody>{rank_rows}</tbody></table>
  <div class="section-title">📉 排名趨勢（越低越好）</div>
  <div class="charts-grid">{rank_charts}</div>
  <div class="note">＊成績與排名：<span>▲ 紅色</span> 表示積分上升（名次退步），<span class="green">▼ 綠色</span> 表示積分下降（名次進步）；排名欄位正相反。</div>
</div>"""


def wrap_html(body, title):
    return f'<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="UTF-8"><title>{title}</title><style>{CSS}</style></head><body>{body}</body></html>'


# ─────────────────────────────────────────────
# Main logic — runs when button is pressed
# ─────────────────────────────────────────────
if run_btn:
    with st.spinner("🔄 正在讀取並分析資料…"):
        try:
            merged = build_merged_df(file1, file2, file3)
        except Exception as e:
            st.error(f"❌ 讀取失敗：{e}")
            st.stop()

    st.success(f"✅ 成功讀取 {len(merged)} 位學生資料")

    # ── Summary metrics ──
    col1, col2, col3 = st.columns(3)
    col1.metric("學生人數", len(merged))
    if '三模_總分' in merged.columns and '一模_總分' in merged.columns:
        avg_growth = (merged['三模_總分'] - merged['一模_總分']).mean()
        col2.metric("平均總分成長（三模－一模）", f"{avg_growth:+.2f}")
    if '三模_班排' in merged.columns and '一模_班排' in merged.columns:
        avg_rank = (merged['一模_班排'] - merged['三模_班排']).mean()
        col3.metric("平均班排進步（正數=進步）", f"{avg_rank:+.1f}")

    # ── Data preview ──
    with st.expander("📋 瀏覽合併資料表", expanded=False):
        st.dataframe(merged, use_container_width=True)

    # ── Generate CSV bytes ──
    csv_bytes = merged.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')

    # ── Generate HTML reports ──
    with st.spinner("📄 正在產生學生成績分析單…"):
        cards = []
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            for _, row in merged.iterrows():
                sid = str(row['座號']).strip().zfill(2)
                name = str(row['姓名']).strip()
                card = generate_card(row, class_name)
                cards.append(card)
                individual_html = wrap_html(card, f"{sid} {name} 三次模考成績分析單")
                zf.writestr(f"{sid}_{name}_成績分析單.html", individual_html.encode('utf-8'))
        zip_buf.seek(0)

        combined_body = '\n'.join(cards)
        combined_css = CSS + "\n  .card { margin-bottom: 40px; }"
        combined_html = f'<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="UTF-8"><title>{class_name} 全班成績分析單</title><style>{combined_css}</style></head><body>{combined_body}</body></html>'
        combined_bytes = combined_html.encode('utf-8')

    # ── Download buttons ──
    st.divider()
    st.subheader("⬇️ 下載結果")
    dl1, dl2, dl3 = st.columns(3)
    with dl1:
        st.download_button(
            label="📥 下載 CSV 成績表",
            data=csv_bytes,
            file_name=f"{class_name}_三次模考綜合分析表.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with dl2:
        st.download_button(
            label="📦 下載個人分析單 (ZIP)",
            data=zip_buf,
            file_name=f"{class_name}_個人成績分析單.zip",
            mime="application/zip",
            use_container_width=True,
        )
    with dl3:
        st.download_button(
            label="📄 下載全班合併分析單",
            data=combined_bytes,
            file_name=f"{class_name}_全班成績分析單.html",
            mime="text/html",
            use_container_width=True,
        )

elif not (file1 and file2 and file3):
    st.info("👈 請在左側上傳三次模考 Excel 檔案，再按「執行分析」。")
