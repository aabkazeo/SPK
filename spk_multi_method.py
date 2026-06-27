"""
SPK Multi-Method — UI modern, tab-based, AHP upper-triangular input
"""
import streamlit as st
import numpy as np
import pandas as pd

st.set_page_config(
    page_title="SPK Multi-Method",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Engines ───────────────────────────────────────────────────────────────────
from ahp_engine        import compute_full_ahp, generate_ahp_excel
from electre_engine    import run_electre,      generate_electre_excel
from saw_topsis_engine import run_saw, run_topsis, generate_saw_excel, generate_topsis_excel

# ── Design tokens ─────────────────────────────────────────────────────────────
ACCENT = {
    "ELECTRE": "#7c3aed",
    "AHP":     "#b45309",
    "SAW":     "#16a34a",
    "TOPSIS":  "#1d4ed8",
}
MEDAL  = ["🥇","🥈","🥉","4️⃣","5️⃣","6️⃣","7️⃣","8️⃣","9️⃣"]
RANK_BG = {1:"#fef9c3",2:"#f1f5f9",3:"#fff7ed"}

SAATY_OPTIONS = {
    "1  —  Sama penting":          1,
    "2":                           2,
    "3  —  Sedikit lebih penting": 3,
    "4":                           4,
    "5  —  Lebih penting":         5,
    "6":                           6,
    "7  —  Sangat lebih penting":  7,
    "8":                           8,
    "9  —  Mutlak lebih penting":  9,
    "1/2":                         0.5,
    "1/3":                         1/3,
    "1/4":                         0.25,
    "1/5":                         0.2,
    "1/6":                         1/6,
    "1/7":                         1/7,
    "1/8":                         0.125,
    "1/9":                         1/9,
}
SAATY_LABELS = list(SAATY_OPTIONS.keys())

# ── Global CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.main .block-container { padding: 1.5rem 2rem 4rem; max-width: 1200px; }

/* ── Sidebar ─────────────────────────────────────────────────── */
section[data-testid="stSidebar"]                { background: #0f172a; }
section[data-testid="stSidebar"] > div          { padding: 1.4rem 1.1rem; }
section[data-testid="stSidebar"] *              { color: #cbd5e1 !important; }
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3             { color: #f8fafc !important; }
section[data-testid="stSidebar"] .stSelectbox > div > div {
    background:#1e293b !important; border:1px solid #334155 !important;
    border-radius:8px !important; }

/* ── Page header ─────────────────────────────────────────────── */
.page-hero {
    padding: 20px 26px 18px;
    background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
    border-radius: 12px; margin-bottom: 22px; color: white;
}
.page-hero .eyebrow { font-size:11px; font-weight:700; letter-spacing:.12em;
    text-transform:uppercase; opacity:.6; margin-bottom:5px; }
.page-hero h1 { font-size:24px; font-weight:700; margin:0 0 4px; line-height:1.2; }
.page-hero p  { font-size:13px; opacity:.65; margin:0; }

/* ── Step label ──────────────────────────────────────────────── */
.step-lbl {
    display:inline-flex; align-items:center; gap:8px;
    font-size:11px; font-weight:700; letter-spacing:.08em;
    text-transform:uppercase; color:#64748b; margin:14px 0 8px;
}
.step-n {
    width:20px; height:20px; border-radius:50%;
    display:inline-flex; align-items:center; justify-content:center;
    font-size:11px; font-weight:800; color:white; flex-shrink:0;
}

/* ── Crit header row ─────────────────────────────────────────── */
.crit-head {
    display:grid; gap:8px; margin-bottom:4px;
    font-size:11px; font-weight:700; color:#94a3b8;
    text-transform:uppercase; letter-spacing:.07em; padding:0 2px;
}

/* ── AHP pair row ────────────────────────────────────────────── */
.pair-row {
    background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px;
    padding:10px 14px; margin-bottom:6px;
}
.pair-end { font-size:13px; font-weight:600; color:#334155; padding-top:6px; }

/* ── Result section ──────────────────────────────────────────── */
.res-title {
    font-size:11px; font-weight:700; letter-spacing:.1em;
    text-transform:uppercase; color:#94a3b8;
    display:flex; align-items:center; gap:8px; margin:26px 0 16px;
}
.res-title::after { content:''; flex:1; height:1px; background:#e2e8f0; }

/* ── Metric cards ────────────────────────────────────────────── */
.metric-grid { display:flex; gap:12px; flex-wrap:wrap; margin-bottom:20px; }
.metric-card {
    flex:1; min-width:160px; border-radius:12px; padding:18px 20px;
    border:2px solid transparent; text-align:center; position:relative;
}
.metric-card.m1 { background:#fefce8; border-color:#facc15; }
.metric-card.m2 { background:#f0fdf4; border-color:#86efac; }
.metric-card.m3 { background:#fff7ed; border-color:#fdba74; }
.metric-card.mn { background:#f8fafc; border-color:#e2e8f0; }
.metric-medal  { font-size:28px; }
.metric-name   { font-size:15px; font-weight:700; color:#0f172a; margin:6px 0 3px; }
.metric-score  { font-size:12px; color:#64748b; font-family:'JetBrains Mono',monospace; }
.metric-rank   { position:absolute; top:10px; right:12px;
    font-size:10px; font-weight:700; color:#64748b; }

/* ── Stat chips ──────────────────────────────────────────────── */
.chip-row { display:flex; gap:8px; flex-wrap:wrap; margin:10px 0 18px; }
.chip {
    background:white; border:1px solid #e2e8f0; border-radius:6px;
    padding:6px 12px; font-size:12px; color:#334155;
}
.chip b { font-family:'JetBrains Mono',monospace; color:#0f172a; }

/* ── Consistency banner ──────────────────────────────────────── */
.cons-ok  { background:#f0fdf4; border:1px solid #86efac; border-radius:8px;
    padding:10px 14px; font-size:13px; font-weight:600; color:#166534; margin:8px 0; }
.cons-bad { background:#fef2f2; border:1px solid #fca5a5; border-radius:8px;
    padding:10px 14px; font-size:13px; font-weight:600; color:#991b1b; margin:8px 0; }

/* ── Button ──────────────────────────────────────────────────── */
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg,#6366f1,#4f46e5) !important;
    color:white !important; border:none !important; border-radius:10px !important;
    padding:12px 40px !important; font-size:15px !important; font-weight:700 !important;
    box-shadow:0 4px 14px rgba(99,102,241,.35) !important; width:100%;
}
.stDownloadButton > button { border-radius:8px !important; font-weight:600 !important; }

/* ── Expander ────────────────────────────────────────────────── */
div[data-testid="stExpander"] {
    border:1px solid #e2e8f0 !important; border-radius:10px !important;
    overflow:hidden; margin-bottom:8px !important;
}
div[data-testid="stExpander"] summary {
    background:#f8fafc !important; padding:10px 16px !important;
    font-weight:600 !important; font-size:13px !important;
}

/* ── Tabs ────────────────────────────────────────────────────── */
button[data-baseweb="tab"] { font-size:13px !important; font-weight:600 !important; }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# Sidebar
# ═══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 🎯 SPK Multi-Method")
    st.markdown("---")
    method = st.selectbox(
        "Metode SPK",
        ["ELECTRE", "AHP", "SAW", "TOPSIS"],
        key="method_select",
    )
    st.markdown("---")
    info_map = {
        "ELECTRE": ("ELimination Et Choix Traduisant la REalité",
                    "Ranking lewat analisis keunggulan berpasangan (concordance & discordance)."),
        "AHP":     ("Analytical Hierarchy Process",
                    "Bobot kriteria dari perbandingan berpasangan Skala Saaty + uji konsistensi (CR)."),
        "SAW":     ("Simple Additive Weighting",
                    "Normalisasi nilai lalu jumlahkan berdasarkan bobot — metode paling sederhana."),
        "TOPSIS":  ("Technique for Order Preference by Similarity to Ideal Solution",
                    "Ranking berdasarkan jarak ke solusi ideal positif dan negatif."),
    }
    full, desc = info_map[method]
    ac = ACCENT[method]
    st.markdown(f"<p style='font-size:12px;font-weight:700;color:{ac};'>{full}</p>", unsafe_allow_html=True)
    st.markdown(f"<p style='font-size:12px;color:#94a3b8;line-height:1.6;'>{desc}</p>", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# Helpers
# ═══════════════════════════════════════════════════════════════════════════════

ICONS = {"ELECTRE":"⚡","AHP":"🔢","SAW":"➕","TOPSIS":"📐"}

def page_hero(method):
    full, desc = info_map[method]
    st.markdown(f"""
    <div class="page-hero">
      <div class="eyebrow">{ICONS[method]} {method}</div>
      <h1>{full}</h1>
      <p>Isi form di bawah secara berurutan, lalu tekan <b>Hitung</b> untuk melihat hasil &amp; download Excel.</p>
    </div>""", unsafe_allow_html=True)

def step_lbl(n, title, accent):
    st.markdown(f"""
    <div class="step-lbl">
      <span class="step-n" style="background:{accent};">{n}</span>{title}
    </div>""", unsafe_allow_html=True)

def metric_cards(alt_names, scores, rank_map, score_label="Skor"):
    MC = {1:"m1",2:"m2",3:"m3"}
    ranked = sorted(range(len(alt_names)), key=lambda i: rank_map[i])
    cards  = ""
    for pos, ai in enumerate(ranked):
        r  = rank_map[ai]
        mc = MC.get(r, "mn")
        md = MEDAL[r-1] if r-1 < len(MEDAL) else f"#{r}"
        cards += f"""
        <div class="metric-card {mc}">
          <div class="metric-rank">#{r}</div>
          <div class="metric-medal">{md}</div>
          <div class="metric-name">{alt_names[ai]}</div>
          <div class="metric-score">{score_label}: {scores[ai]:.5f}</div>
        </div>"""
    st.markdown(f'<div class="metric-grid">{cards}</div>', unsafe_allow_html=True)

def res_title(txt):
    st.markdown(f'<div class="res-title">{txt}</div>', unsafe_allow_html=True)

def chip_row(chips: dict):
    html = '<div class="chip-row">' + "".join(
        f'<div class="chip">{k} &nbsp;<b>{v}</b></div>' for k, v in chips.items()
    ) + '</div>'
    st.markdown(html, unsafe_allow_html=True)

def crit_input_section(prefix, n_crit):
    """Returns crit_names, weights, criteria_types"""
    st.markdown("""
    <div class="crit-head" style="grid-template-columns:2.2fr .8fr 1.1fr;">
      <span>Nama Kriteria</span><span style="text-align:center">Bobot</span>
      <span style="text-align:center">Jenis</span>
    </div>""", unsafe_allow_html=True)
    crit_names, weights, types = [], [], []
    for j in range(n_crit):
        c1, c2, c3 = st.columns([2.2, .8, 1.1])
        cn = c1.text_input("", value=f"K{j+1}", key=f"{prefix}_cn_{j}",
                            label_visibility="collapsed")
        w  = c2.number_input("", 1, 99, 1, key=f"{prefix}_w_{j}",
                              label_visibility="collapsed")
        t  = c3.selectbox("", ["benefit ↑","cost ↓"], key=f"{prefix}_t_{j}",
                           label_visibility="collapsed")
        crit_names.append(cn.strip() or f"K{j+1}")
        weights.append(float(w))
        types.append("benefit" if "benefit" in t else "cost")
    return crit_names, weights, types

def matrix_input(prefix, n_alt, n_crit, alt_names, crit_names):
    """Decision matrix — returns np.ndarray"""
    cols_h = st.columns([1.4] + [1]*n_crit)
    cols_h[0].markdown("<p style='font-size:11px;color:#94a3b8;font-weight:700;margin:0;text-align:right;padding-right:8px;'>↓ Alt \\ Krit →</p>", unsafe_allow_html=True)
    for j in range(n_crit):
        cols_h[j+1].markdown(f"<p style='font-size:11px;font-weight:700;color:#475569;text-align:center;background:#f1f5f9;border-radius:4px;padding:4px 0;margin:0;'>{crit_names[j]}</p>", unsafe_allow_html=True)
    X = np.zeros((n_alt, n_crit))
    for i in range(n_alt):
        rc = st.columns([1.4] + [1]*n_crit)
        rc[0].markdown(f"<p style='font-size:13px;font-weight:600;color:#334155;text-align:right;padding:6px 8px 0 0;margin:0;'>{alt_names[i]}</p>", unsafe_allow_html=True)
        for j in range(n_crit):
            X[i,j] = rc[j+1].number_input("", value=0.0, format="%.2f",
                          key=f"{prefix}_x_{i}_{j}", label_visibility="collapsed")
    return X

def alt_inputs(prefix, n_alt, pfx="A"):
    cols = st.columns(min(n_alt, 5))
    return [(cols[i%5].text_input(f"Alternatif {i+1}", value=f"{pfx}{i+1}",
             key=f"{prefix}_alt_{i}").strip() or f"{pfx}{i+1}")
            for i in range(n_alt)]

def hitung_btn(label):
    _, bc, _ = st.columns([1.5, 2, 1.5])
    return bc.button(f"✦ {label}", use_container_width=True, type="primary")

def download_row(buf, fname, label="⬇ Download Laporan Excel"):
    _, dc, _ = st.columns([1, 2, 1])
    dc.download_button(label, data=buf, file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════════
#  ELECTRE
# ═══════════════════════════════════════════════════════════════════════════════
if method == "ELECTRE":
    ac = ACCENT["ELECTRE"]
    page_hero("ELECTRE")

    tab_in, tab_res = st.tabs(["📋 Input Data", "📊 Hasil Perhitungan"])

    # ── INPUT TAB ─────────────────────────────────────────────────────────────
    with tab_in:
        c1, c2 = st.columns(2)
        n_alt  = int(c1.number_input("Jumlah Alternatif", 2, 15, 3, key="e_na"))
        n_crit = int(c2.number_input("Jumlah Kriteria",   2, 15, 5, key="e_nc"))

        step_lbl(1, "Nama Alternatif", ac)
        alt_names = alt_inputs("e", n_alt)

        step_lbl(2, "Kriteria, Bobot & Jenis", ac)
        st.caption("💡 **benefit ↑** = nilai besar lebih baik &nbsp;|&nbsp; **cost ↓** = nilai kecil lebih baik")
        crit_names, weights, criteria_types = crit_input_section("e", n_crit)

        step_lbl(3, "Matriks Nilai Keputusan", ac)
        st.caption("Isi nilai tiap alternatif terhadap tiap kriteria.")
        X_in = matrix_input("e", n_alt, n_crit, alt_names, crit_names)

        st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
        if hitung_btn("Hitung ELECTRE"):
            if np.all(X_in == 0):
                st.error("⚠️ Semua nilai 0 — isi matriks keputusan dulu.")
            else:
                with st.spinner("Menghitung…"):
                    res = run_electre(X_in, weights, criteria_types, alt_names, crit_names)
                st.session_state["e_res"] = res
                st.success("✅ Selesai! Lihat tab **Hasil Perhitungan**.")

    # ── HASIL TAB ─────────────────────────────────────────────────────────────
    with tab_res:
        if "e_res" not in st.session_state:
            st.info("Isi data di tab **Input Data** lalu tekan Hitung.")
        else:
            res = st.session_state["e_res"]

            res_title("Peringkat Akhir")
            scores_arr = np.array([r["E_score"] for r in sorted(res["ranking"], key=lambda x:x["alt_idx"])])
            metric_cards(res["alt_names"], scores_arr, res["rank_map"], "Skor E")

            chip_row({
                "Threshold C (c̄)":  f"{res['threshold_c']:.4f}",
                "Threshold D (d̄)":  f"{res['threshold_d']:.6f}",
                "Alternatif terbaik": res["alt_names"][min(res["rank_map"], key=res["rank_map"].get)],
            })

            t1,t2,t3,t4,t5 = st.tabs(["R (Normalisasi)","V (Terbobot)",
                                        "Concordance","Discordance","Dominan & E"])
            with t1:
                st.caption("r_ij = x_ij / √Σx_ij²")
                st.dataframe(pd.DataFrame(res["R"], index=res["alt_names"],
                             columns=res["crit_names"]).style.format("{:.6f}")
                             .background_gradient(cmap="Purples"), use_container_width=True)
            with t2:
                st.caption("v_ij = w_j × r_ij")
                st.dataframe(pd.DataFrame(res["V"], index=res["alt_names"],
                             columns=res["crit_names"]).style.format("{:.6f}")
                             .background_gradient(cmap="Greens"), use_container_width=True)
            with t3:
                pairs = [(k,l) for k in range(res["n_alt"])
                              for l in range(res["n_alt"]) if k!=l]
                rows  = [{"Pasangan": f"C{k+1}{l+1}",
                          "Himpunan C": "{"+",".join(str(j+1) for j in res["c_sets"][(k,l)])+"}",
                          "Himpunan D": "{"+",".join(str(j+1) for j in res["d_sets"][(k,l)])+"}"
                          } for k,l in pairs]
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
                st.markdown(f"**Matriks Concordance** · threshold c̄ = `{res['threshold_c']:.4f}`")
                C_df = pd.DataFrame(
                    [["-" if k==l else res["C"][k,l] for l in range(res["n_alt"])]
                     for k in range(res["n_alt"])],
                    index=res["alt_names"], columns=res["alt_names"])
                st.dataframe(C_df, use_container_width=True)
            with t4:
                st.markdown(f"**Matriks Discordance** · threshold d̄ = `{res['threshold_d']:.6f}`")
                D_df = pd.DataFrame(
                    [["-" if k==l else res["D"][k,l] for l in range(res["n_alt"])]
                     for k in range(res["n_alt"])],
                    index=res["alt_names"], columns=res["alt_names"])
                st.dataframe(D_df, use_container_width=True)
            with t5:
                g1,g2,g3 = st.columns(3)
                for col, (lbl, mat) in zip([g1,g2,g3],[
                    ("Dominan C (F)", res["F"]),
                    ("Dominan D (G)", res["G"]),
                    ("Agregat (E)",   res["E"])]):
                    df_m = pd.DataFrame(
                        [["-" if k==l else int(mat[k,l]) for l in range(res["n_alt"])]
                         for k in range(res["n_alt"])],
                        index=res["alt_names"], columns=res["alt_names"])
                    col.markdown(f"**{lbl}**")
                    col.dataframe(df_m, use_container_width=True)

                ranked = sorted(res["ranking"], key=lambda x: x["rank"])
                st.markdown("**Tabel Skor & Ranking**")
                rdf = pd.DataFrame([{"#": r["rank"],
                                      "Alternatif": res["alt_names"][r["alt_idx"]],
                                      "Skor E": round(r["E_score"],6)} for r in ranked])
                st.dataframe(rdf, use_container_width=True, hide_index=True)

            st.markdown("---")
            buf = generate_electre_excel(res)
            download_row(buf, "laporan_electre.xlsx",
                         "⬇ Download Laporan ELECTRE (Excel — dengan formula)")


# ═══════════════════════════════════════════════════════════════════════════════
#  AHP
# ═══════════════════════════════════════════════════════════════════════════════
elif method == "AHP":
    ac = ACCENT["AHP"]
    page_hero("AHP")

    tab_krit, tab_alt, tab_res = st.tabs([
        "🔢 Perbandingan Kriteria",
        "🔢 Perbandingan Alternatif",
        "📊 Hasil Perhitungan",
    ])

    # shared config — di luar tab agar persisten
    with st.sidebar:
        st.markdown("---")
        n_crit = int(st.number_input("Jumlah Kriteria",    2, 10, 4, key="a_nc"))
        n_alt  = int(st.number_input("Jumlah Alternatif", 2, 10, 3, key="a_na"))

    # ── Tab: Perbandingan Kriteria ──────────────────────────────────────────────
    with tab_krit:
        step_lbl(1, "Nama Kriteria", ac)
        cols_cn = st.columns(min(n_crit, 5))
        crit_names = [(cols_cn[j % 5].text_input(f"Kriteria {j+1}",
                       value=f"K{j+1}", key=f"a_cn_{j}").strip() or f"K{j+1}")
                      for j in range(n_crit)]

        step_lbl(2, "Matriks Perbandingan Berpasangan Kriteria", ac)
        st.caption(
            "Hanya isi segitiga atas. Diagonal otomatis = 1. "
            "Segitiga bawah otomatis = kebalikan (1/nilai)."
        )

        crit_matrix = np.ones((n_crit, n_crit))
        n_pairs = n_crit * (n_crit - 1) // 2

        if n_pairs == 0:
            st.info("Tambahkan lebih dari 1 kriteria.")
        else:
            use_two_col = n_pairs > 4
            left_pairs, right_pairs = [], []
            pair_list = [(i, j) for i in range(n_crit) for j in range(i+1, n_crit)]
            for idx, (i, j) in enumerate(pair_list):
                if use_two_col:
                    (left_pairs if idx % 2 == 0 else right_pairs).append((i, j))
                else:
                    left_pairs.append((i, j))

            if use_two_col:
                col_a, col_b = st.columns(2)
                for idx, (i, j) in enumerate(pair_list):
                    target = col_a if idx % 2 == 0 else col_b
                    with target:
                        lc, mc, rc = st.columns([1, 2.5, 1])
                        lc.markdown(f"<p class='pair-end' style='text-align:right'>{crit_names[i]}</p>",
                                    unsafe_allow_html=True)
                        sel = mc.selectbox("", SAATY_LABELS, index=0,
                                           key=f"a_cp_{i}_{j}",
                                           label_visibility="collapsed")
                        rc.markdown(f"<p class='pair-end'>{crit_names[j]}</p>",
                                    unsafe_allow_html=True)
                    v = SAATY_OPTIONS[sel]
                    crit_matrix[i, j] = v
                    crit_matrix[j, i] = 1.0 / v
            else:
                for i, j in pair_list:
                    lc, mc, rc = st.columns([1, 2.5, 1])
                    lc.markdown(f"<p class='pair-end' style='text-align:right'>{crit_names[i]}</p>",
                                unsafe_allow_html=True)
                    sel = mc.selectbox("", SAATY_LABELS, index=0,
                                       key=f"a_cp_{i}_{j}",
                                       label_visibility="collapsed")
                    rc.markdown(f"<p class='pair-end'>{crit_names[j]}</p>",
                                unsafe_allow_html=True)
                    v = SAATY_OPTIONS[sel]
                    crit_matrix[i, j] = v
                    crit_matrix[j, i] = 1.0 / v

        with st.expander("👁 Preview matriks kriteria lengkap"):
            st.dataframe(pd.DataFrame(crit_matrix, index=crit_names,
                         columns=crit_names).style.format("{:.4f}"),
                         use_container_width=True)

    # ── Tab: Perbandingan Alternatif ────────────────────────────────────────────
    with tab_alt:
        step_lbl(1, "Nama Alternatif", ac)
        cols_an = st.columns(min(n_alt, 5))
        alt_names = [(cols_an[i % 5].text_input(f"Alternatif {i+1}",
                      value=f"Alt{i+1}", key=f"a_an_{i}").strip() or f"Alt{i+1}")
                     for i in range(n_alt)]

        step_lbl(2, "Matriks Perbandingan Alternatif per Kriteria", ac)
        st.caption(
            "Untuk setiap kriteria, isi segitiga atas perbandingan antar alternatif. "
            "Diagonal = 1, segitiga bawah = otomatis 1/nilai."
        )

        alt_matrices = []
        n_alt_pairs  = n_alt * (n_alt - 1) // 2

        for j in range(n_crit):
            cn_label = crit_names[j] if j < len(crit_names) else f"K{j+1}"
            with st.expander(f"Kriteria: **{cn_label}** — {n_alt_pairs} pasang perbandingan",
                             expanded=(j == 0)):
                am = np.ones((n_alt, n_alt))
                alt_pair_list = [(a, b) for a in range(n_alt) for b in range(a+1, n_alt)]
                use_two = n_alt_pairs > 4

                if use_two:
                    col_a, col_b = st.columns(2)
                    for idx, (a, b) in enumerate(alt_pair_list):
                        target = col_a if idx % 2 == 0 else col_b
                        with target:
                            lc, mc, rc = st.columns([1, 2.5, 1])
                            lc.markdown(f"<p class='pair-end' style='text-align:right'>{alt_names[a]}</p>",
                                        unsafe_allow_html=True)
                            sel = mc.selectbox("", SAATY_LABELS, index=0,
                                               key=f"a_ap_{j}_{a}_{b}",
                                               label_visibility="collapsed")
                            rc.markdown(f"<p class='pair-end'>{alt_names[b]}</p>",
                                        unsafe_allow_html=True)
                        v = SAATY_OPTIONS[sel]
                        am[a, b] = v
                        am[b, a] = 1.0 / v
                else:
                    for a, b in alt_pair_list:
                        lc, mc, rc = st.columns([1, 2.5, 1])
                        lc.markdown(f"<p class='pair-end' style='text-align:right'>{alt_names[a]}</p>",
                                    unsafe_allow_html=True)
                        sel = mc.selectbox("", SAATY_LABELS, index=0,
                                           key=f"a_ap_{j}_{a}_{b}",
                                           label_visibility="collapsed")
                        rc.markdown(f"<p class='pair-end'>{alt_names[b]}</p>",
                                    unsafe_allow_html=True)
                        v = SAATY_OPTIONS[sel]
                        am[a, b] = v
                        am[b, a] = 1.0 / v

                st.dataframe(pd.DataFrame(am, index=alt_names,
                             columns=alt_names).style.format("{:.4f}"),
                             use_container_width=True)
                alt_matrices.append(am)

        # Pastikan jumlah alt_matrices = n_crit (isi sisa jika expander belum dibuka)
        while len(alt_matrices) < n_crit:
            alt_matrices.append(np.ones((n_alt, n_alt)))

        st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)
        if hitung_btn("Hitung AHP"):
            with st.spinner("Menghitung eigenvalue & konsistensi…"):
                full_res = compute_full_ahp(crit_matrix, alt_matrices,
                                            crit_names, alt_names)
            st.session_state["a_res"] = full_res
            st.success("✅ Selesai! Buka tab **Hasil Perhitungan**.")

    # ── Tab: Hasil ──────────────────────────────────────────────────────────────
    with tab_res:
        if "a_res" not in st.session_state:
            st.info("Isi perbandingan di tab sebelumnya lalu tekan Hitung.")
        else:
            fr  = st.session_state["a_res"]
            cr  = fr["crit_res"]
            alt_names_res  = fr["alt_names"]
            crit_names_res = fr["crit_names"]

            res_title("Peringkat Akhir")
            metric_cards(alt_names_res, fr["scores"], fr["rank_map"])

            # Konsistensi kriteria
            if cr["consistent"]:
                st.markdown(f'<div class="cons-ok">✅ Matriks kriteria KONSISTEN — CR = {cr["CR"]:.6f} (batas: 0.10)</div>',
                            unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="cons-bad">❌ Matriks kriteria TIDAK KONSISTEN — CR = {cr["CR"]:.6f} ≥ 0.10 · Perlu direvisi!</div>',
                            unsafe_allow_html=True)

            chip_row({"λ_max": f"{cr['lambda_max']:.6f}",
                      "CI": f"{cr['CI']:.6f}",
                      "RI": f"{cr['RI']:.2f}",
                      "CR": f"{cr['CR']:.6f}"})

            t1, t2, t3 = st.tabs(["Bobot Kriteria","Bobot Alternatif","Skor & Ranking"])

            with t1:
                bdf = pd.DataFrame({
                    "Kriteria": crit_names_res,
                    "Bobot (Prioritas)": cr["priority"],
                    "Eigen Value": cr["lambda_vals"],
                })
                st.dataframe(bdf.style.format({"Bobot (Prioritas)":"{:.8f}",
                                               "Eigen Value":"{:.8f}"})
                             .background_gradient(cmap="Oranges", subset=["Bobot (Prioritas)"]),
                             use_container_width=True, hide_index=True)

            with t2:
                all_ok = all(ar["consistent"] for ar in fr["alt_results"])
                if not all_ok:
                    st.warning("⚠️ Satu atau lebih matriks alternatif tidak konsisten (CR ≥ 0.10).")
                for j, (cn, ar) in enumerate(zip(crit_names_res, fr["alt_results"])):
                    icon = "✅" if ar["consistent"] else "❌"
                    with st.expander(f"{icon} {cn}  —  CR = {ar['CR']:.6f}"):
                        adf = pd.DataFrame({
                            "Alternatif": alt_names_res,
                            "Bobot": ar["priority"],
                            "Eigen Value": ar["lambda_vals"],
                        })
                        st.dataframe(adf.style.format({"Bobot":"{:.8f}","Eigen Value":"{:.8f}"}),
                                     use_container_width=True, hide_index=True)

            with t3:
                W_df = pd.DataFrame(fr["W"], index=alt_names_res, columns=crit_names_res)
                W_df["Skor Total"] = fr["scores"]
                st.dataframe(W_df.style.format("{:.6f}")
                             .background_gradient(cmap="YlOrRd", subset=["Skor Total"]),
                             use_container_width=True)
                ranked = sorted(range(len(alt_names_res)), key=lambda i: fr["rank_map"][i])
                rdf = pd.DataFrame([{"#": fr["rank_map"][i],
                                      "Alternatif": alt_names_res[i],
                                      "Skor Total": round(fr["scores"][i], 8)}
                                     for i in ranked])
                st.dataframe(rdf, use_container_width=True, hide_index=True)

            st.markdown("---")
            buf = generate_ahp_excel(fr)
            download_row(buf, "laporan_ahp.xlsx",
                         "⬇ Download Laporan AHP (Excel — dengan formula)")


# ═══════════════════════════════════════════════════════════════════════════════
#  SAW
# ═══════════════════════════════════════════════════════════════════════════════
elif method == "SAW":
    ac = ACCENT["SAW"]
    page_hero("SAW")

    tab_in, tab_res = st.tabs(["📋 Input Data","📊 Hasil Perhitungan"])

    with tab_in:
        c1, c2 = st.columns(2)
        n_alt  = int(c1.number_input("Jumlah Alternatif", 2, 15, 3, key="s_na"))
        n_crit = int(c2.number_input("Jumlah Kriteria",   2, 15, 4, key="s_nc"))

        step_lbl(1, "Nama Alternatif", ac)
        alt_names = alt_inputs("s", n_alt)

        step_lbl(2, "Kriteria, Bobot & Jenis", ac)
        st.caption("💡 **benefit ↑** = besar lebih baik &nbsp;|&nbsp; **cost ↓** = kecil lebih baik")
        crit_names, weights, criteria_types = crit_input_section("s", n_crit)

        step_lbl(3, "Matriks Nilai Keputusan", ac)
        X_in = matrix_input("s", n_alt, n_crit, alt_names, crit_names)

        st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
        if hitung_btn("Hitung SAW"):
            if np.all(X_in == 0):
                st.error("⚠️ Semua nilai 0 — isi matriks dulu.")
            else:
                with st.spinner("Menghitung…"):
                    res = run_saw(X_in, weights, criteria_types, alt_names, crit_names)
                st.session_state["s_res"] = res
                st.success("✅ Selesai! Buka tab Hasil Perhitungan.")

    with tab_res:
        if "s_res" not in st.session_state:
            st.info("Isi data di tab Input Data lalu tekan Hitung.")
        else:
            res = st.session_state["s_res"]
            alt_names  = res["alt_names"]
            crit_names = res["crit_names"]
            n_alt  = len(alt_names)
            n_crit = len(crit_names)

            res_title("Peringkat Akhir SAW")
            metric_cards(alt_names, res["scores"], res["rank_map"])

            ranked = sorted(range(n_alt), key=lambda i: res["rank_map"][i])
            chip_row({
                "Alternatif terbaik": alt_names[ranked[0]],
                "Skor tertinggi":     f"{res['scores'][ranked[0]]:.6f}",
            })

            t1, t2 = st.tabs(["Matriks Normalisasi (R)","Skor & Ranking"])
            with t1:
                st.caption("Benefit: r = x/MAX(x) &nbsp;|&nbsp; Cost: r = MIN(x)/x")
                st.dataframe(pd.DataFrame(res["R"], index=alt_names, columns=crit_names)
                             .style.format("{:.6f}").background_gradient(cmap="Greens"),
                             use_container_width=True)
                st.dataframe(pd.DataFrame([res["w_norm"]], columns=crit_names,
                             index=["Bobot Ternormalisasi"]).style.format("{:.4f}"),
                             use_container_width=True)
            with t2:
                rdf = pd.DataFrame([{"#": res["rank_map"][i], "Alternatif": alt_names[i],
                                      "Skor SAW": res["scores"][i]} for i in ranked])
                st.dataframe(rdf.style.format({"Skor SAW":"{:.6f}"})
                             .background_gradient(cmap="Greens", subset=["Skor SAW"]),
                             use_container_width=True, hide_index=True)

            st.markdown("---")
            buf = generate_saw_excel(res)
            download_row(buf, "hasil_saw.xlsx",
                         "⬇ Download Laporan SAW (Excel — dengan formula)")


# ═══════════════════════════════════════════════════════════════════════════════
#  TOPSIS
# ═══════════════════════════════════════════════════════════════════════════════
elif method == "TOPSIS":
    ac = ACCENT["TOPSIS"]
    page_hero("TOPSIS")

    tab_in, tab_res = st.tabs(["📋 Input Data","📊 Hasil Perhitungan"])

    with tab_in:
        c1, c2 = st.columns(2)
        n_alt  = int(c1.number_input("Jumlah Alternatif", 2, 15, 3, key="t_na"))
        n_crit = int(c2.number_input("Jumlah Kriteria",   2, 15, 4, key="t_nc"))

        step_lbl(1, "Nama Alternatif", ac)
        alt_names = alt_inputs("t", n_alt)

        step_lbl(2, "Kriteria, Bobot & Jenis", ac)
        st.caption("💡 **benefit ↑** = besar lebih baik &nbsp;|&nbsp; **cost ↓** = kecil lebih baik")
        crit_names, weights, criteria_types = crit_input_section("t", n_crit)

        step_lbl(3, "Matriks Nilai Keputusan", ac)
        X_in = matrix_input("t", n_alt, n_crit, alt_names, crit_names)

        st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
        if hitung_btn("Hitung TOPSIS"):
            if np.all(X_in == 0):
                st.error("⚠️ Semua nilai 0 — isi matriks dulu.")
            else:
                with st.spinner("Menghitung…"):
                    res = run_topsis(X_in, weights, criteria_types, alt_names, crit_names)
                st.session_state["to_res"] = res
                st.success("✅ Selesai! Buka tab Hasil Perhitungan.")

    with tab_res:
        if "to_res" not in st.session_state:
            st.info("Isi data di tab Input Data lalu tekan Hitung.")
        else:
            res = st.session_state["to_res"]
            alt_names  = res["alt_names"]
            crit_names = res["crit_names"]
            n_alt  = len(alt_names)
            n_crit = len(crit_names)

            ranked = sorted(range(n_alt), key=lambda i: res["rank_map"][i])

            res_title("Peringkat Akhir TOPSIS")
            metric_cards(alt_names, res["C"], res["rank_map"], "Nilai Ci")

            chip_row({
                "Alternatif terbaik": alt_names[ranked[0]],
                "Ci tertinggi":       f"{res['C'][ranked[0]]:.6f}",
                "D+ terkecil":        f"{res['D_pos'][ranked[0]]:.6f}",
            })

            t1,t2,t3,t4 = st.tabs(["R (Normalisasi)","V (Terbobot) & Ideal","Jarak D+/D–","Ranking"])

            with t1:
                st.caption("r_ij = x_ij / √Σx_ij²")
                st.dataframe(pd.DataFrame(res["R"], index=alt_names, columns=crit_names)
                             .style.format("{:.6f}").background_gradient(cmap="Blues"),
                             use_container_width=True)
            with t2:
                st.caption("v_ij = w_j × r_ij")
                st.dataframe(pd.DataFrame(res["V"], index=alt_names, columns=crit_names)
                             .style.format("{:.6f}").background_gradient(cmap="Blues"),
                             use_container_width=True)
                ca, cb = st.columns(2)
                ca.markdown("**A⁺ — Solusi Ideal Positif**")
                ca.dataframe(pd.DataFrame([res["A_pos"]], columns=crit_names,
                             index=["A+"]).style.format("{:.6f}"), use_container_width=True)
                cb.markdown("**A⁻ — Solusi Ideal Negatif**")
                cb.dataframe(pd.DataFrame([res["A_neg"]], columns=crit_names,
                             index=["A-"]).style.format("{:.6f}"), use_container_width=True)
            with t3:
                ddf = pd.DataFrame({
                    "Alternatif": alt_names,
                    "D⁺ (ke A+)": res["D_pos"],
                    "D⁻ (ke A-)": res["D_neg"],
                    "Ci (Preferensi)": res["C"],
                })
                st.dataframe(ddf.style.format({
                    "D⁺ (ke A+)":"{:.6f}","D⁻ (ke A-)":"{:.6f}","Ci (Preferensi)":"{:.6f}"
                }).background_gradient(cmap="Blues", subset=["Ci (Preferensi)"]),
                use_container_width=True, hide_index=True)
                st.caption("Ci = D⁻ / (D⁺ + D⁻) · semakin besar = semakin baik")
            with t4:
                rdf = pd.DataFrame([{"#": res["rank_map"][i], "Alternatif": alt_names[i],
                                      "Nilai Ci": res["C"][i]} for i in ranked])
                st.dataframe(rdf.style.format({"Nilai Ci":"{:.6f}"})
                             .background_gradient(cmap="Blues", subset=["Nilai Ci"]),
                             use_container_width=True, hide_index=True)

            st.markdown("---")
            buf = generate_topsis_excel(res)
            download_row(buf, "hasil_topsis.xlsx",
                         "⬇ Download Laporan TOPSIS (Excel — dengan formula)")

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("<div style='height:40px'></div>", unsafe_allow_html=True)
st.markdown("""
<div style='text-align:center;color:#94a3b8;font-size:11px;letter-spacing:.06em;'>
  SPK Multi-Method &nbsp;·&nbsp; ELECTRE · AHP · SAW · TOPSIS
</div>""", unsafe_allow_html=True)