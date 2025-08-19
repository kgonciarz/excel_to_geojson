import io
import json
import pandas as pd
import geopandas as gpd
from shapely import wkt
import streamlit as st

st.set_page_config(page_title="Excel → GeoJSON", page_icon="🗺️", layout="wide")
st.title("🗺️ Excel / CSV → GeoJSON Converter")

st.markdown(
    "Upload an **Excel (.xlsx/.xls)** or **CSV** file, choose how to form geometries "
    "(WKT or X/Y), set the input CRS, and download a **GeoJSON**."
)

# --- Helpers -----------------------------------------------------------------
COMMON_EPSG = {
    "WGS84 (EPSG:4326)": 4326,
    "Web Mercator (EPSG:3857)": 3857,
    "CH LV95 (EPSG:2056)": 2056,
    "ETRS89 / LAEA Europe (EPSG:3035)": 3035,
}

def detect_xy_columns(cols):
    # lightweight heuristics for common lon/lat column names
    lower = [c.lower() for c in cols]
    candidates = {
        "x": ["x", "lon", "long", "longitude", "easting"],
        "y": ["y", "lat", "latitude", "northing"],
    }
    x_guess = next((cols[lower.index(c)] for c in candidates["x"] if c in lower), None)
    y_guess = next((cols[lower.index(c)] for c in candidates["y"] if c in lower), None)
    return x_guess, y_guess

def make_gdf_from_xy(df, x_col, y_col, epsg_in):
    if x_col not in df.columns or y_col not in df.columns:
        raise ValueError("Selected X or Y column not found in the table.")
    # Coerce to numeric and drop bad rows
    df = df.copy()
    df[x_col] = pd.to_numeric(df[x_col], errors="coerce")
    df[y_col] = pd.to_numeric(df[y_col], errors="coerce")
    df = df.dropna(subset=[x_col, y_col])
    gdf = gpd.GeoDataFrame(
        df,
        geometry=gpd.points_from_xy(df[x_col], df[y_col]),
        crs=f"EPSG:{epsg_in}",
    )
    return gdf

def make_gdf_from_wkt(df, wkt_col, epsg_in):
    if wkt_col not in df.columns:
        raise ValueError(f"WKT column '{wkt_col}' not found.")
    df = df.copy()
    # Parse WKT; invalid rows become NaN then dropped
    df["__geom__"] = df[wkt_col].apply(lambda s: wkt.loads(s) if pd.notna(s) else None)
    df = df.dropna(subset=["__geom__"])
    gdf = gpd.GeoDataFrame(
        df.drop(columns=["__geom__"]),
        geometry=df["__geom__"],
        crs=f"EPSG:{epsg_in}",
    )
    return gdf

def to_geojson_bytes(gdf: gpd.GeoDataFrame) -> bytes:
    # Use GeoPandas' to_json to avoid temp files; ensure RFC 7946 (lon/lat) if CRS=4326
    text = gdf.to_json(drop_id=True)
    return text.encode("utf-8")

# --- Upload ------------------------------------------------------------------
uploaded = st.file_uploader(
    "Drag & drop Excel or CSV",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=False,
    help="Large files are fine; only the first sheet is used by default for Excel.",
)

if not uploaded:
    st.info("Upload a file to begin.")
    st.stop()

# --- Read file ---------------------------------------------------------------
filetype = uploaded.name.lower().rsplit(".", 1)[-1]

with st.spinner("Reading file..."):
    if filetype in {"xlsx", "xls"}:
        xls = pd.ExcelFile(uploaded)
        sheet = st.selectbox("Sheet", options=xls.sheet_names, index=0)
        df = xls.parse(sheet)
    else:  # CSV
        # Try to sniff delimiter; fallback to comma
        try:
            df = pd.read_csv(uploaded)
        except Exception:
            uploaded.seek(0)
            df = pd.read_csv(uploaded, sep=";")

if df.empty:
    st.error("No data found in the selected sheet/file.")
    st.stop()

st.success(f"Loaded {len(df):,} rows, {len(df.columns)} columns.")

# --- Geometry mode -----------------------------------------------------------
st.subheader("Geometry settings")

geom_mode = st.radio(
    "Choose geometry source",
    ["X/Y columns (points)", "WKT column (points/lines/polygons)"],
    horizontal=True,
)

epsg_label = st.selectbox(
    "Input CRS (EPSG)",
    list(COMMON_EPSG.keys()) + ["Other…"],
    index=0,
)
if epsg_label == "Other…":
    epsg_in = st.number_input("Enter EPSG code", value=4326, step=1, min_value=1000)
else:
    epsg_in = COMMON_EPSG[epsg_label]

epsg_out_label = st.selectbox(
    "Output CRS (EPSG)",
    list(COMMON_EPSG.keys()) + ["Other…"],
    index=0,
)
if epsg_out_label == "Other…":
    epsg_out = st.number_input("Enter output EPSG", value=4326, step=1, min_value=1000, key="out_epsg")
else:
    epsg_out = COMMON_EPSG[epsg_out_label]

gdf = None
error_box = st.empty()

try:
    if geom_mode.startswith("X/Y"):
        x_guess, y_guess = detect_xy_columns(df.columns)
        cols = list(df.columns)
        c1, c2 = st.columns(2)
        with c1:
            x_col = st.selectbox("X / Longitude / Easting", options=cols, index=cols.index(x_guess) if x_guess in cols else 0)
        with c2:
            y_col = st.selectbox("Y / Latitude / Northing", options=cols, index=cols.index(y_guess) if y_guess in cols else 1)
        if st.button("Build points"):
            gdf = make_gdf_from_xy(df, x_col, y_col, epsg_in)

    else:
        # WKT mode
        wkt_candidates = [c for c in df.columns if c.lower() in {"wkt", "geom", "geometry"}]
        wkt_col = st.selectbox(
            "WKT geometry column",
            options=list(df.columns),
            index=(df.columns.tolist().index(wkt_candidates[0]) if wkt_candidates else 0),
        )
        if st.button("Build geometries from WKT"):
            gdf = make_gdf_from_wkt(df, wkt_col, epsg_in)

except Exception as e:
    error_box.error(f"⚠️ {e}")

# --- When we have a GeoDataFrame --------------------------------------------
if isinstance(gdf, gpd.GeoDataFrame):
    # Reproject if requested
    if epsg_in != epsg_out:
        with st.spinner(f"Reprojecting to EPSG:{epsg_out}…"):
            gdf = gdf.to_crs(epsg=epsg_out)

    st.success(f"Geometry created. CRS: EPSG:{epsg_out}. Rows with valid geometry: {len(gdf):,}")

    # Preview
    with st.expander("Preview attribute table (first 200 rows)"):
        st.dataframe(gdf.drop(columns="geometry", errors="ignore").head(200), use_container_width=True)

    # Quick map (points will show best; polygons/lines also render)
    try:
        st.subheader("Quick map preview")
        st.map(gdf.to_crs(epsg=4326), use_container_width=True)
    except Exception:
        st.info("Map preview unavailable for this geometry; you can still download the GeoJSON.")

    # Download
    geojson_bytes = to_geojson_bytes(gdf.to_crs(epsg=4326))  # RFC 7946 convention (lon/lat)
    st.download_button(
        label="⬇️ Download GeoJSON (EPSG:4326)",
        data=geojson_bytes,
        file_name="data.geojson",
        mime="application/geo+json",
        help="GeoJSON is exported in EPSG:4326 (lon/lat) for web compatibility.",
    )

else:
    st.info("Configure geometry above, then click **Build**.")
