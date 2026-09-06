from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# ============================================================
# KONFIGURASI HALAMAN
# ============================================================
st.set_page_config(
    page_title="Dashboard Peta Geotag Bangunan",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    /* Jarak bagian atas diperbesar agar judul tidak tertutup header Streamlit */
    .block-container {
        padding-top: 3.2rem !important;
        padding-bottom: 2.5rem;
        max-width: 1550px;
    }

    [data-testid="stSidebar"] {
        background: #f4f7fb;
    }

    [data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid #e4eaf2;
        padding: 14px 16px;
        border-radius: 14px;
        box-shadow: 0 2px 10px rgba(16, 24, 40, 0.05);
    }

    [data-testid="stMetricLabel"] {
        font-weight: 600;
    }

    .dashboard-title {
        font-size: clamp(1.8rem, 3vw, 2.35rem);
        line-height: 1.25;
        font-weight: 800;
        margin: 0 0 0.35rem 0;
        padding-top: 0.25rem;
        color: #152238;
    }

    .dashboard-subtitle {
        color: #667085;
        margin: 0 0 1.2rem 0;
        font-size: 0.98rem;
    }

    .selected-card {
        background: #ffffff;
        border: 1px solid #dce5ef;
        border-radius: 16px;
        padding: 18px 20px;
        box-shadow: 0 3px 14px rgba(16, 24, 40, 0.06);
        margin-top: 0.4rem;
        margin-bottom: 0.8rem;
    }

    .selected-name {
        font-size: 1.15rem;
        font-weight: 800;
        color: #152238;
        margin-bottom: 4px;
    }

    .selected-subtitle {
        color: #667085;
        font-size: 0.92rem;
        margin-bottom: 8px;
    }

    .detail-label {
        color: #667085;
        font-size: 0.78rem;
        text-transform: uppercase;
        letter-spacing: 0.03em;
        margin-bottom: 2px;
    }

    .detail-value {
        font-weight: 650;
        color: #1f2937;
        margin-bottom: 9px;
    }

    div[data-testid="stLinkButton"] a {
        border-radius: 10px;
        font-weight: 650;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# KONSTANTA
# ============================================================
DATA_FILE = Path(__file__).with_name("Map.xlsx")

REQUIRED_COLUMNS = [
    "rn",
    "assignment_id",
    "nama_assignment",
    "kecamatan",
    "desa_kelurahan",
    "sls",
    "kode_sub_sls",
    "jenis_prelist",
    "kode_bang_label",
    "no_bang",
    "geotag_latitude",
    "geotag_longitude",
    "assignment_status_alias",
    "sumber_data",
    "link_fasih",
]

STATUS_COLOR_MAP = {
    "SUBMITTED BY Pencacah": "#F59E0B",
    "APPROVED BY Pengawas": "#16A34A",
    "REJECTED BY Pengawas": "#DC2626",
    "DRAFT": "#2563EB",
    "OPEN": "#64748B",
}

BANGUNAN_COLOR_MAP = {
    "1. Bangunan Khusus Usaha": "#7C3AED",
    "2. Bangunan Campuran": "#0284C7",
    "3. Bangunan Tempat Tinggal": "#16A34A",
    "6. Bangunan Lainnya yang Tidak Tercakup (Tempat Judi, Tempat Layanan Kencan, Bangunan Kosong/ Rusak)": "#DC2626",
}


# ============================================================
# FUNGSI
# ============================================================
@st.cache_data(show_spinner=False)
def load_data(path_string: str) -> pd.DataFrame:
    df = pd.read_excel(path_string, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            "Kolom berikut tidak ditemukan pada Map.xlsx: " + ", ".join(missing)
        )

    df = df.copy()

    # ID internal untuk menghubungkan titik peta dengan baris data.
    df["_point_id"] = df.index.astype(str)

    # Koordinat
    df["geotag_latitude"] = pd.to_numeric(
        df["geotag_latitude"].astype(str).str.replace(",", ".", regex=False),
        errors="coerce",
    )
    df["geotag_longitude"] = pd.to_numeric(
        df["geotag_longitude"].astype(str).str.replace(",", ".", regex=False),
        errors="coerce",
    )

    # Bersihkan teks
    text_columns = [
        "rn",
        "assignment_id",
        "nama_assignment",
        "kecamatan",
        "desa_kelurahan",
        "sls",
        "kode_sub_sls",
        "jenis_prelist",
        "kode_bang_label",
        "no_bang",
        "assignment_status_alias",
        "sumber_data",
        "link_fasih",
    ]

    for col in text_columns:
        df[col] = df[col].fillna("").astype(str).str.strip()

    df["jenis_prelist_tampil"] = df["jenis_prelist"].replace("", "Tidak terisi")

    # Status ringkas
    status_upper = df["assignment_status_alias"].str.upper()
    df["status_ringkas"] = "Lainnya"
    df.loc[status_upper.str.contains("SUBMITTED", na=False), "status_ringkas"] = "Submitted"
    df.loc[status_upper.str.contains("APPROVED", na=False), "status_ringkas"] = "Approved"
    df.loc[status_upper.str.contains("REJECTED", na=False), "status_ringkas"] = "Rejected"

    # Validasi koordinat
    df["koordinat_valid"] = (
        df["geotag_latitude"].between(-90, 90)
        & df["geotag_longitude"].between(-180, 180)
    )

    return df


def multiselect_filter(data: pd.DataFrame, column: str, label: str) -> pd.DataFrame:
    options = sorted(
        x
        for x in data[column].dropna().astype(str).unique().tolist()
        if x.strip() != ""
    )
    selected = st.sidebar.multiselect(label, options)
    if selected:
        return data[data[column].isin(selected)]
    return data


def fmt_int(value) -> str:
    return f"{int(value):,}".replace(",", ".")


def make_map_style(style_name: str):
    """Basemap satelit/hybrid tanpa Google/Mapbox API key."""
    if style_name == "Satellite + Label":
        return {
            "version": 8,
            "sources": {
                "esri_imagery": {
                    "type": "raster",
                    "tiles": [
                        "https://server.arcgisonline.com/ArcGIS/rest/services/"
                        "World_Imagery/MapServer/tile/{z}/{y}/{x}"
                    ],
                    "tileSize": 256,
                    "attribution": "Tiles © Esri",
                },
                "esri_transportation": {
                    "type": "raster",
                    "tiles": [
                        "https://server.arcgisonline.com/ArcGIS/rest/services/"
                        "Reference/World_Transportation/MapServer/tile/{z}/{y}/{x}"
                    ],
                    "tileSize": 256,
                },
                "esri_places": {
                    "type": "raster",
                    "tiles": [
                        "https://server.arcgisonline.com/ArcGIS/rest/services/"
                        "Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}"
                    ],
                    "tileSize": 256,
                },
            },
            "layers": [
                {"id": "imagery", "type": "raster", "source": "esri_imagery"},
                {"id": "roads", "type": "raster", "source": "esri_transportation"},
                {"id": "places", "type": "raster", "source": "esri_places"},
            ],
        }

    if style_name == "Satellite":
        return {
            "version": 8,
            "sources": {
                "esri_imagery": {
                    "type": "raster",
                    "tiles": [
                        "https://server.arcgisonline.com/ArcGIS/rest/services/"
                        "World_Imagery/MapServer/tile/{z}/{y}/{x}"
                    ],
                    "tileSize": 256,
                    "attribution": "Tiles © Esri",
                }
            },
            "layers": [
                {"id": "imagery", "type": "raster", "source": "esri_imagery"}
            ],
        }

    return "open-street-map"


def get_selection_points(plot_event):
    """Kompatibel dengan bentuk return PlotlyState Streamlit."""
    if plot_event is None:
        return []

    try:
        return plot_event.selection.points
    except Exception:
        pass

    try:
        return plot_event["selection"]["points"]
    except Exception:
        return []


def selected_row_from_point(point, valid_data):
    customdata = point.get("customdata")
    if customdata is None:
        return None

    if isinstance(customdata, (list, tuple)):
        point_id = str(customdata[0])
    else:
        point_id = str(customdata)

    matched = valid_data[valid_data["_point_id"] == point_id]
    if matched.empty:
        return None

    return matched.iloc[0]


def google_directions_url(row) -> str:
    return (
        "https://www.google.com/maps/dir/?api=1"
        f"&destination={row['geotag_latitude']},{row['geotag_longitude']}"
        "&travelmode=driving"
    )


def google_point_url(row) -> str:
    return (
        "https://www.google.com/maps/search/?api=1"
        f"&query={row['geotag_latitude']},{row['geotag_longitude']}"
    )


# ============================================================
# DATA FIX
# ============================================================
if not DATA_FILE.exists():
    st.error(
        "File **Map.xlsx** tidak ditemukan. "
        "Pastikan Map.xlsx berada pada folder yang sama dengan file dashboard Python."
    )
    st.stop()

try:
    df = load_data(str(DATA_FILE))
except Exception as exc:
    st.error(f"Gagal membaca Map.xlsx: {exc}")
    st.stop()


# ============================================================
# HEADER
# ============================================================
st.markdown(
    '<div class="dashboard-title">Dashboard Peta Geotag Bangunan</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="dashboard-subtitle">'
    'Visualisasi lokasi assignment berdasarkan koordinat geotag — '
    'Kecamatan Sirimau, Kota Ambon.'
    '</div>',
    unsafe_allow_html=True,
)


# ============================================================
# SIDEBAR / FILTER
# ============================================================
st.sidebar.title("🗺️ Kontrol Dashboard")
st.sidebar.caption("Filter data dan pengaturan tampilan peta.")

st.sidebar.subheader("Filter Data")

filtered = df.copy()
filtered = multiselect_filter(filtered, "kecamatan", "Kecamatan")
filtered = multiselect_filter(filtered, "desa_kelurahan", "Desa/Kelurahan")
filtered = multiselect_filter(filtered, "sls", "SLS")
filtered = multiselect_filter(filtered, "jenis_prelist_tampil", "Jenis Prelist")
filtered = multiselect_filter(filtered, "kode_bang_label", "Jenis Bangunan")
filtered = multiselect_filter(filtered, "assignment_status_alias", "Status Assignment")

search_name = st.sidebar.text_input(
    "Cari nama assignment",
    placeholder="Contoh: RUMAH KOSONG",
).strip()

if search_name:
    filtered = filtered[
        filtered["nama_assignment"].str.contains(search_name, case=False, na=False)
    ]

st.sidebar.markdown("---")

color_by = st.sidebar.radio(
    "Warna titik berdasarkan",
    ["Status Assignment", "Jenis Bangunan", "Jenis Prelist"],
)

map_style_name = st.sidebar.selectbox(
    "Gaya peta",
    ["Satellite + Label", "Satellite", "OpenStreetMap"],
    index=0,
)

map_style = make_map_style(map_style_name)


# ============================================================
# KPI
# ============================================================
valid_map = filtered[filtered["koordinat_valid"]].copy()

submitted = (filtered["status_ringkas"] == "Submitted").sum()
approved = (filtered["status_ringkas"] == "Approved").sum()
rejected = (filtered["status_ringkas"] == "Rejected").sum()

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Total Data", fmt_int(len(filtered)))
k2.metric("Titik di Peta", fmt_int(len(valid_map)))
k3.metric("Submitted", fmt_int(submitted))
k4.metric("Approved", fmt_int(approved))
k5.metric("Rejected", fmt_int(rejected))

st.caption(
    f"Sumber data tetap: **Map.xlsx** · "
    f"Koordinat tidak valid/kosong: **{fmt_int(len(filtered) - len(valid_map))}**"
)


# ============================================================
# PETA INTERAKTIF
# ============================================================
st.subheader("📍 Peta Sebaran Geotag")
st.caption(
    "Klik salah satu titik pada peta untuk menampilkan detail lokasi, "
    "membuka titik di Google Maps, atau mendapatkan petunjuk arah."
)

if valid_map.empty:
    st.warning("Tidak ada titik dengan koordinat valid untuk filter yang dipilih.")
    st.stop()


if color_by == "Status Assignment":
    color_col = "assignment_status_alias"
    color_map = STATUS_COLOR_MAP
    legend_title = "Status Assignment"
elif color_by == "Jenis Bangunan":
    color_col = "kode_bang_label"
    color_map = BANGUNAN_COLOR_MAP
    legend_title = "Jenis Bangunan"
else:
    color_col = "jenis_prelist_tampil"
    color_map = None
    legend_title = "Jenis Prelist"

center = {
    "lat": float(valid_map["geotag_latitude"].median()),
    "lon": float(valid_map["geotag_longitude"].median()),
}

hover_data = {
    "geotag_latitude": ":.7f",
    "geotag_longitude": ":.7f",
    "kecamatan": True,
    "desa_kelurahan": True,
    "sls": True,
    "kode_sub_sls": True,
    "jenis_prelist_tampil": True,
    "kode_bang_label": True,
    "no_bang": True,
    "assignment_status_alias": True,
    "_point_id": False,
}

fig = px.scatter_mapbox(
    valid_map,
    lat="geotag_latitude",
    lon="geotag_longitude",
    color=color_col,
    color_discrete_map=color_map,
    hover_name="nama_assignment",
    hover_data=hover_data,
    custom_data=["_point_id"],
    center=center,
    zoom=17,
    height=690,
    opacity=0.92,
)

fig.update_traces(
    marker={"size": 13},
    selected={"marker": {"size": 19, "opacity": 1}},
    unselected={"marker": {"opacity": 0.72}},
)

# Highlight titik terpilih dibuat sebagai trace yang SELALU ada.
# Dengan begitu struktur figure tidak berubah setelah marker diklik,
# sehingga Plotly lebih mudah mempertahankan posisi center/zoom pengguna.
current_selected_id = st.session_state.get("selected_point_id")
highlight = valid_map[valid_map["_point_id"] == str(current_selected_id)]

if not highlight.empty:
    h = highlight.iloc[0]
    highlight_lat = [h["geotag_latitude"]]
    highlight_lon = [h["geotag_longitude"]]
else:
    highlight_lat = []
    highlight_lon = []

fig.add_trace(
    go.Scattermapbox(
        lat=highlight_lat,
        lon=highlight_lon,
        mode="markers",
        marker={"size": 23, "color": "white", "opacity": 0.85},
        hoverinfo="skip",
        showlegend=False,
    )
)
fig.add_trace(
    go.Scattermapbox(
        lat=highlight_lat,
        lon=highlight_lon,
        mode="markers",
        marker={"size": 14, "color": "#111827", "opacity": 1},
        hoverinfo="skip",
        showlegend=False,
    )
)

fig.update_layout(
    mapbox_style=map_style,
    clickmode="event+select",
    dragmode="pan",

    # Jangan kembalikan peta ke center/zoom awal ketika Streamlit rerun.
    # Nilai konstan membuat Plotly mempertahankan viewport dari browser.
    uirevision="keep-map-viewport-v1",
    selectionrevision="keep-map-selection-v1",

    margin=dict(l=0, r=0, t=0, b=0),
    legend=dict(
        title=legend_title,
        orientation="h",
        yanchor="bottom",
        y=0.01,
        xanchor="left",
        x=0.01,
        bgcolor="rgba(255,255,255,0.88)",
    ),
)

plot_event = st.plotly_chart(
    fig,
    use_container_width=True,
    key="geotag_map",
    on_select="rerun",
    selection_mode="points",
    config={
        "scrollZoom": True,
        "displaylogo": False,
    },
)

points = get_selection_points(plot_event)

if points:
    clicked_row = selected_row_from_point(points[-1], valid_map)
    if clicked_row is not None:
        new_id = str(clicked_row["_point_id"])
        if st.session_state.get("selected_point_id") != new_id:
            # Tidak ada st.rerun() kedua di sini.
            # on_select="rerun" sudah cukup dan viewport tetap dipertahankan.
            st.session_state["selected_point_id"] = new_id


# ============================================================
# DETAIL TITIK YANG DIKLIK
# ============================================================
selected_id = st.session_state.get("selected_point_id")
selected_match = valid_map[valid_map["_point_id"] == str(selected_id)] if selected_id is not None else pd.DataFrame()

if selected_match.empty:
    # Jika filter membuat titik lama hilang, bersihkan pilihan.
    if selected_id is not None:
        st.session_state.pop("selected_point_id", None)

    st.info("👆 Klik satu titik pada peta untuk melihat informasi lengkap lokasi.")
else:
    row = selected_match.iloc[0]

    st.subheader("📌 Informasi Titik Terpilih")

    st.markdown(
        f"""
        <div class="selected-card">
            <div class="selected-name">{row['nama_assignment'] or '-'}</div>
            <div class="selected-subtitle">
                {row['desa_kelurahan'] or '-'} · {row['sls'] or '-'}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    d1, d2, d3, d4 = st.columns(4)

    with d1:
        st.markdown('<div class="detail-label">Kecamatan</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="detail-value">{row["kecamatan"] or "-"}</div>', unsafe_allow_html=True)

        st.markdown('<div class="detail-label">Desa/Kelurahan</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="detail-value">{row["desa_kelurahan"] or "-"}</div>', unsafe_allow_html=True)

    with d2:
        st.markdown('<div class="detail-label">SLS</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="detail-value">{row["sls"] or "-"}</div>', unsafe_allow_html=True)

        st.markdown('<div class="detail-label">Kode Sub-SLS</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="detail-value">{row["kode_sub_sls"] or "-"}</div>', unsafe_allow_html=True)

    with d3:
        st.markdown('<div class="detail-label">Jenis Prelist</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="detail-value">{row["jenis_prelist"] or "-"}</div>', unsafe_allow_html=True)

        st.markdown('<div class="detail-label">Jenis Bangunan</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="detail-value">{row["kode_bang_label"] or "-"}</div>', unsafe_allow_html=True)

    with d4:
        st.markdown('<div class="detail-label">No. Bangunan</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="detail-value">{row["no_bang"] or "-"}</div>', unsafe_allow_html=True)

        st.markdown('<div class="detail-label">Status Assignment</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="detail-value">{row["assignment_status_alias"] or "-"}</div>', unsafe_allow_html=True)

    st.markdown("##### Koordinat")
    st.code(
        f"{row['geotag_latitude']:.7f}, {row['geotag_longitude']:.7f}",
        language=None,
    )

    direction_url = google_directions_url(row)
    point_url = google_point_url(row)

    b1, b2, b3 = st.columns(3)

    with b1:
        st.link_button(
            "🚗 Dapatkan Arah",
            direction_url,
            use_container_width=True,
        )

    with b2:
        st.link_button(
            "📍 Buka Titik di Google Maps",
            point_url,
            use_container_width=True,
        )

    with b3:
        if row["link_fasih"]:
            st.link_button(
                "🔗 Buka Assignment Fasih",
                row["link_fasih"],
                use_container_width=True,
            )
        else:
            st.button(
                "🔗 Link Fasih Tidak Tersedia",
                disabled=True,
                use_container_width=True,
            )

    with st.expander("Informasi tambahan"):
        st.write(f"**Assignment ID:** {row['assignment_id'] or '-'}")
        st.write(f"**Nomor urut data:** {row['rn'] or '-'}")
        st.write(f"**Sumber data:** {row['sumber_data'] or '-'}")


# ============================================================
# RINGKASAN
# ============================================================
st.markdown("---")
left, right = st.columns(2)

with left:
    st.subheader("Status Assignment")
    status_summary = (
        filtered["assignment_status_alias"]
        .replace("", "Tidak terisi")
        .value_counts()
        .rename_axis("Status")
        .reset_index(name="Jumlah")
    )

    fig_status = px.bar(
        status_summary,
        x="Jumlah",
        y="Status",
        orientation="h",
        text="Jumlah",
    )
    fig_status.update_layout(
        height=310,
        margin=dict(l=0, r=10, t=10, b=0),
        xaxis_title="Jumlah",
        yaxis_title=None,
        showlegend=False,
    )
    st.plotly_chart(fig_status, use_container_width=True)

with right:
    st.subheader("Jenis Bangunan")
    bang_summary = (
        filtered["kode_bang_label"]
        .replace("", "Tidak terisi")
        .value_counts()
        .rename_axis("Jenis Bangunan")
        .reset_index(name="Jumlah")
    )

    fig_bang = px.bar(
        bang_summary,
        x="Jumlah",
        y="Jenis Bangunan",
        orientation="h",
        text="Jumlah",
    )
    fig_bang.update_layout(
        height=310,
        margin=dict(l=0, r=10, t=10, b=0),
        xaxis_title="Jumlah",
        yaxis_title=None,
        showlegend=False,
    )
    st.plotly_chart(fig_bang, use_container_width=True)


# ============================================================
# TABEL DETAIL
# ============================================================
st.subheader("📋 Detail Data")

table = filtered[
    [
        "rn",
        "nama_assignment",
        "kecamatan",
        "desa_kelurahan",
        "sls",
        "kode_sub_sls",
        "jenis_prelist",
        "kode_bang_label",
        "no_bang",
        "geotag_latitude",
        "geotag_longitude",
        "assignment_status_alias",
        "link_fasih",
    ]
].copy()

st.dataframe(
    table,
    use_container_width=True,
    height=430,
    hide_index=True,
    column_config={
        "rn": "No.",
        "nama_assignment": "Nama Assignment",
        "kecamatan": "Kecamatan",
        "desa_kelurahan": "Desa/Kelurahan",
        "sls": "SLS",
        "kode_sub_sls": "Kode Sub-SLS",
        "jenis_prelist": "Jenis Prelist",
        "kode_bang_label": "Jenis Bangunan",
        "no_bang": "No. Bangunan",
        "geotag_latitude": st.column_config.NumberColumn("Latitude", format="%.7f"),
        "geotag_longitude": st.column_config.NumberColumn("Longitude", format="%.7f"),
        "assignment_status_alias": "Status",
        "link_fasih": st.column_config.LinkColumn(
            "Fasih",
            display_text="Buka Assignment",
        ),
    },
)
