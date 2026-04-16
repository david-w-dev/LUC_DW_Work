# =========================
# CSV → 4326 + ENU (site-origin): EPSG:3857 input (abs or local) with orthometric Z
# - Handles inputs as ABS_3857 or LOCAL_3857 (adds SOURCE_ORIGIN first).
# - Uses your Unreal/Cesium site georeference as ENU origin.
# - Outputs:
#     x_4326 (lon°), y_4326 (lat°), z_4326 (ellipsoidal m, EGM2008),
#     x_m, y_m, z_m (ENU meters at site origin).
#
# Requirements:
#   - pygeodesy available to Blender's Python
#   - egm2008-1.pgm path set
# =========================

import csv, math, os, sys

#----------- Easy Access --------------------------------------------------------
# Carreg Wen - -3.49315 51.691848 55.3
# Clocaenog Dau - -3.473858 53.067066 55.3
# Glyn Cothi - -4.090242 52.001487 55.3
# Engage Demo - -3.379333859 51.83580198
# Afan_NeddTawe - -3.671571263 51.68789031 0.0

# ---------- CONFIG -------------------------------------------------------------
# Input CSV path (absolute); output will default to "<input>_converted.csv" if empty.
CSV_INPUT_PATH  = r"T:\134\13447_Trydan Gwyrdd Cymru Trydan 3D Interactive Model Support\VIS\QGIS\Afan_TurbinePoints_3857.csv"
CSV_OUTPUT_PATH = r""  # optional; leave blank to auto-name

# Column names for absolute EPSG:3857 meters and orthometric Z (meters)
IN_X_COL = "x"   # easting (meters, EPSG:3857)  -- case-insensitive
IN_Y_COL = "y"   # northing (meters, EPSG:3857) -- case-insensitive
IN_Z_COL = "z"   # orthometric height (meters)  -- case-insensitive

# Your input points are real georeferenced (absolute 3857) by default.
# Switch to "LOCAL_3857" if providing local offsets instead.
INPUT_COORD_MODE = "ABS_3857"  # "ABS_3857" or "LOCAL_3857"

# Use your Unreal/Cesium site georeference as the ENU origin
# Option 1: LLH (recommended) – provide lon/lat in degrees and explicit ellipsoidal height (meters)
ORIGIN_MODE   = "LLH"          # "LLH" or "EPSG3857"
ORIGIN_LON    = -3.671571263     # deg (site lon)
ORIGIN_LAT    = 51.68789031     # deg (site lat)
ORIGIN_H_ELLIP = 0.0          # ellipsoidal height (m). Set to your Cesium georeference height.

# If INPUT_COORD_MODE == "LOCAL_3857", provide the local origin in 3857 (meters)
SOURCE_ORIGIN_X_3857 = -386708.0
SOURCE_ORIGIN_Y_3857 = 6995413.0

# Option 2: EPSG:3857 – specify origin in 3857 and derive lon/lat; set explicit ellipsoidal height
ORIGIN_X_3857 = -386708.104    # meters (from your lon/lat)
ORIGIN_Y_3857 = 6995412.963    # meters (from your lon/lat)
ORIGIN_H_ELLIP_3857 = 0.0      # ellipsoidal height (m) when using EPSG3857 origin

# Geoid grid (EGM2008). Use the verified working path on your machine:
EGM2008_PGM_PATH = r"C:\Users\williams_d\geographiclib\geoids\egm2008-1\egm2008-1.pgm"

# If geoid is missing, optionally fall back to constant Z offset instead of erroring
FALLBACK_TO_CONSTANT_OFFSET = True
Z_OFFSET_METERS = 0.0
# ------------------------------------------------------------------------------

# Make Blender see pygeodesy
mods = os.path.expandvars(r"%APPDATA%\Blender Foundation\Blender\4.2\scripts\modules")
user_site = os.path.expandvars(r"%APPDATA%\Python\Python311\site-packages")
if os.path.isdir(mods) and mods not in sys.path: sys.path.append(mods)
if os.path.isdir(user_site) and user_site not in sys.path: sys.path.append(user_site)

try:
    from pygeodesy.geoids import GeoidKarney
except Exception as e:
    raise RuntimeError(
        "PyGeodesy not available. Install with:\n"
        r'"C:\Program Files\Blender Foundation\Blender 4.2\4.2\python\bin\python.exe" -m pip install --user pygeodesy'
    ) from e

# ---------- robust path resolver (keeps your U: path, auto-falls back to UNC) --
def _resolve_csv_path(path_in: str) -> str:
    p = os.path.expandvars(path_in)
    if os.path.exists(p):
        return p
    # Windows UNC fallback (map drive letter -> \\server\share)
    if os.name == 'nt' and len(p) >= 2 and p[1] == ':':
        try:
            import ctypes
            from ctypes import wintypes
            WNetGetConnectionW = ctypes.windll.mpr.WNetGetConnectionW
            WNetGetConnectionW.argtypes = [wintypes.LPCWSTR, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
            WNetGetConnectionW.restype  = wintypes.DWORD
            remote = ctypes.create_unicode_buffer(2048)
            size   = wintypes.DWORD(len(remote))
            rc = WNetGetConnectionW(p[:2], remote, ctypes.byref(size))  # 'U:' -> \\server\share
            if rc == 0:
                unc_root = remote.value
                rest = p[2:].lstrip("\\/")
                candidate = os.path.join(unc_root, rest)
                if os.path.exists(candidate):
                    print(f"[INFO] Using UNC fallback for {p[:2]} -> {unc_root}")
                    return candidate
        except Exception:
            pass
    # Forward-slash variant as a last try
    pfwd = p.replace("\\", "/")
    if os.path.exists(pfwd):
        return pfwd
    raise FileNotFoundError(f"CSV_INPUT_PATH not found: {path_in}")

# ---------- math / proj helpers -----------------------------------------------
R_MERC = 6378137.0  # Web Mercator sphere radius
A = 6378137.0
F = 1.0 / 298.257223563
E2 = F * (2.0 - F)

def mercator_to_lonlat_deg(x_m, y_m):
    lon_rad = x_m / R_MERC
    lat_rad = math.atan(math.sinh(max(min(y_m / R_MERC, 20.0), -20.0)))
    return lon_rad * 180.0 / math.pi, lat_rad * 180.0 / math.pi

def llh_to_ecef(lon_deg, lat_deg, h_m):
    lon = math.radians(lon_deg); lat = math.radians(lat_deg)
    sin_lat, cos_lat = math.sin(lat), math.cos(lat)
    sin_lon, cos_lon = math.sin(lon), math.cos(lon)
    N = A / math.sqrt(1.0 - E2 * sin_lat * sin_lat)
    X = (N + h_m) * cos_lat * cos_lon
    Y = (N + h_m) * cos_lat * sin_lon
    Z = (N * (1.0 - E2) + h_m) * sin_lat
    return X, Y, Z

def ecef_to_enu(x, y, z, lon0_deg, lat0_deg, h0_m):
    X0, Y0, Z0 = llh_to_ecef(lon0_deg, lat0_deg, h0_m)
    dx, dy, dz = x - X0, y - Y0, z - Z0
    lon0 = math.radians(lon0_deg); lat0 = math.radians(lat0_deg)
    sin_lat0, cos_lat0 = math.sin(lat0), math.cos(lat0)
    sin_lon0, cos_lon0 = math.sin(lon0), math.cos(lon0)
    e = -sin_lon0 * dx +  cos_lon0 * dy
    n = -sin_lat0 * cos_lon0 * dx - sin_lat0 * sin_lon0 * dy + cos_lat0 * dz
    u =  cos_lat0 * cos_lon0 * dx +  cos_lat0 * sin_lon0 * dy + sin_lat0 * dz
    return e, n, u

# ---------- geoid --------------------------------------------------------------
def _resolve_pgm():
    if EGM2008_PGM_PATH and os.path.isfile(EGM2008_PGM_PATH):
        return EGM2008_PGM_PATH
    for p in (
        r"C:\ProgramData\GeographicLib\geoids\egm2008-1.pgm",
        os.path.expanduser(r"~\geographiclib\geoids\egm2008-1.pgm"),
        r"C:\geographiclib\geoids\egm2008-1.pgm",
    ):
        if os.path.isfile(p): return p
    return None

_PGM = _resolve_pgm()
_GEOID = GeoidKarney(_PGM) if _PGM else None
if _GEOID is None and not FALLBACK_TO_CONSTANT_OFFSET:
    raise FileNotFoundError("EGM2008 PGM not found; set EGM2008_PGM_PATH or enable FALLBACK_TO_CONSTANT_OFFSET.")

def ortho_to_ellip(z_ortho_m, lon_deg, lat_deg):
    if _GEOID is not None:
        return z_ortho_m + _GEOID.height(lat_deg, lon_deg)  # (lat, lon)
    return z_ortho_m + Z_OFFSET_METERS

# ---------- utils --------------------------------------------------------------
def safe_float(v):
    try: return float(v)
    except Exception: return None

def _resolve_cols_case_insensitive(fields, want_names):
    """
    Map desired column names to actual CSV header names, case-insensitively.
    Example: 'x' -> 'X' if the file uses uppercase.
    """
    lower_map = {c.lower(): c for c in fields}
    resolved = {}
    for w in want_names:
        key = w.lower()
        if key in lower_map:
            resolved[w] = lower_map[key]
        else:
            raise KeyError(f"Missing input column '{w}' (case-insensitive) in headers: {fields}")
    return resolved

# ---------- main ---------------------------------------------------------------
def process_csv():
    CSV_IN = _resolve_csv_path(CSV_INPUT_PATH)

    out_path = CSV_OUTPUT_PATH or os.path.join(
        os.path.dirname(CSV_IN),
        os.path.splitext(os.path.basename(CSV_IN))[0] + "_converted.csv"
    )

    # Pass 1: compute lon/lat/h_ellip for each row
    rows = []
    with open(CSV_IN, "r", encoding="utf-8-sig", newline="") as f:
        rdr = csv.DictReader(f)
        fields = rdr.fieldnames or []

        # --- Case-insensitive header resolution for x/y/z
        colmap = _resolve_cols_case_insensitive(fields, (IN_X_COL, IN_Y_COL, IN_Z_COL))
        x_col = colmap[IN_X_COL]
        y_col = colmap[IN_Y_COL]
        z_col = colmap[IN_Z_COL]

        for rec in rdr:
            x = safe_float(rec.get(x_col))
            y = safe_float(rec.get(y_col))
            z = safe_float(rec.get(z_col))
            if x is None or y is None or z is None:
                rec["x_4326"]=rec["y_4326"]=rec["z_4326"]=""
                rows.append(rec)
                continue

            if INPUT_COORD_MODE.upper() == "LOCAL_3857":
                x_abs = SOURCE_ORIGIN_X_3857 + x
                y_abs = SOURCE_ORIGIN_Y_3857 + y
            else:
                x_abs, y_abs = x, y

            lon_deg, lat_deg = mercator_to_lonlat_deg(x_abs, y_abs)
            h_ellip = ortho_to_ellip(z, lon_deg, lat_deg)

            rec["x_4326"] = f"{lon_deg:.9f}"
            rec["y_4326"] = f"{lat_deg:.9f}"
            rec["z_4326"] = f"{h_ellip:.3f}"

            rows.append(rec)

    # Decide ENU origin lon, lat, h based on ORIGIN_MODE, using explicit height
    if ORIGIN_MODE.upper() == "LLH":
        lon0, lat0 = float(ORIGIN_LON), float(ORIGIN_LAT)
        h0 = float(ORIGIN_H_ELLIP)  # explicit
    elif ORIGIN_MODE.upper() == "EPSG3857":
        lon0, lat0 = mercator_to_lonlat_deg(float(ORIGIN_X_3857), float(ORIGIN_Y_3857))
        h0 = float(ORIGIN_H_ELLIP_3857)  # explicit
    else:
        raise ValueError("ORIGIN_MODE must be 'LLH' or 'EPSG3857'.")

    # Pass 2: compute ENU vs site origin
    for rec in rows:
        try:
            lon_deg = float(rec["x_4326"]); lat_deg = float(rec["y_4326"]); h_ellip = float(rec["z_4326"])
        except Exception:
            rec["x_m"]=rec["y_m"]=rec["z_m"]=""
            continue
        X, Y, Z = llh_to_ecef(lon_deg, lat_deg, h_ellip)
        e, n, u = ecef_to_enu(X, Y, Z, lon0, lat0, h0)
        rec["x_m"], rec["y_m"], rec["z_m"] = f"{e:.3f}", f"{n:.3f}", f"{u:.3f}"

    # Write output (preserve original cols; append new if needed)
    new_cols = ["x_4326","y_4326","z_4326","x_m","y_m","z_m"]
    with open(CSV_IN, "r", encoding="utf-8-sig", newline="") as f:
        base_fields = csv.DictReader(f).fieldnames or []
    out_fields = base_fields + [c for c in new_cols if c not in base_fields]

    with open(out_path, "w", encoding="utf-8", newline="") as fo:
        w = csv.DictWriter(fo, fieldnames=out_fields)
        w.writeheader()
        for rec in rows: w.writerow(rec)

    print(f"Done. Site origin (lon,lat,h) = ({lon0:.8f}, {lat0:.8f}, {h0:.3f} m). Wrote: {out_path}")

if __name__ == "__main__":
    process_csv()
