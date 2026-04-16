
# =========================
# LOD2 GLB: Convert EPSG:3857 -> EPSG:4326 per-vertex, with orthometric Z -> ellipsoidal Z
# and write the duplicate in practical Cartesian meters for Blender (ENU or ECEF-centered).
#
# Pipeline per vertex:
#   (x_local, y_local, z_ortho) + ORIGIN_3857  ->  (lon, lat) [4326]
#   z_ellip = z_ortho + N(lat,lon)  [EGM2008 geoid undulation]
#   (lon, lat, z_ellip) -> ENU or ECEF-centered meters (or lon/lat degrees if requested)
#
# Requirements:
#   - PyGeodesy installed where Blender can see it
#   - EGM2008 grid file (.pgm)
# =========================

import bpy, math, os, sys

# --- CONFIG -------------------------------------------------------------------
# World origin of the object in EPSG:3857 meters (easting, northing).
ORIGIN_X_3857 = -408717.4434   # <-- set this (meters)
ORIGIN_Y_3857 = 6743887.568   # <-- set this (meters)

# Where is the geoid grid file? (explicit path easiest)
EGM2008_PGM_PATH = r"D:\GeographicLib\egm2008-1\geoids\egm2008-1.pgm"

# Output coordinate system for the DUPLICATE:
#   "ENU_ORIGIN"        -> Local East/North/Up meters, tangent at origin lon/lat (recommended)
#   "ECEF_CENTERED"     -> Earth-Centered, Earth-Fixed XYZ meters, then translated so origin ~ (0,0,0)
#   "LONLAT_DEGREES"    -> Write lon/lat deg directly to X/Y (z in meters)  [not great for Blender]
#   "METERS_LOCAL_3857" -> Keep original local 3857 offsets for X/Y; only fix Z to ellipsoidal
OUTPUT_COORDS = "ENU_ORIGIN"

# Vertical options:
# If geoid missing, optionally fall back to a constant Z offset (meters) instead of erroring.
FALLBACK_TO_CONSTANT_OFFSET = True
Z_OFFSET_METERS = 0.0

# Apply object transforms before processing? (recommended True)
APPLY_OBJECT_TRANSFORMS = True

# ENU reference ellipsoidal height for the origin:
#   - Set a float (meters) to force a reference height,
#   - or None to auto-use the mean ellipsoidal height of all vertices.
ORIGIN_H_ELLIPSOIDAL = None
# ------------------------------------------------------------------------------

# --- Make sure Blender can import pygeodesy (path shim) -----------------------
mods = os.path.expandvars(r"%APPDATA%\Blender Foundation\Blender\4.2\scripts\modules")
if os.path.isdir(mods) and mods not in sys.path:
    sys.path.append(mods)
user_site = os.path.expandvars(r"%APPDATA%\Python\Python311\site-packages")
if os.path.isdir(user_site) and user_site not in sys.path:
    sys.path.append(user_site)

try:
    from pygeodesy.geoids import GeoidKarney
except Exception as e:
    raise RuntimeError(
        "PyGeodesy is not installed or not discoverable in Blender's Python.\n"
        "Install with:\n"
        r'"C:\Program Files\Blender Foundation\Blender 4.2\4.2\python\bin\python.exe" -m pip install --user pygeodesy'
    ) from e

# --- Helpers: projections, ellipsoid, frames ---------------------------------
R_MERC = 6378137.0  # EPSG:3857 sphere radius (meters)

def mercator_to_lonlat_deg(x_m: float, y_m: float):
    """Inverse Web Mercator (EPSG:3857) -> (lon_deg, lat_deg)."""
    lon_rad = x_m / R_MERC
    lat_rad = math.atan(math.sinh(max(min(y_m / R_MERC, 20.0), -20.0)))
    return (lon_rad * 180.0 / math.pi, lat_rad * 180.0 / math.pi)

# WGS84 ellipsoid
A = 6378137.0
F = 1.0 / 298.257223563
E2 = F * (2.0 - F)  # first eccentricity squared

def llh_to_ecef(lon_deg: float, lat_deg: float, h_m: float):
    """(lon, lat in degrees, h in meters) -> ECEF XYZ (meters)."""
    lon = math.radians(lon_deg)
    lat = math.radians(lat_deg)
    sin_lat = math.sin(lat)
    cos_lat = math.cos(lat)
    sin_lon = math.sin(lon)
    cos_lon = math.cos(lon)
    N = A / math.sqrt(1.0 - E2 * sin_lat * sin_lat)
    X = (N + h_m) * cos_lat * cos_lon
    Y = (N + h_m) * cos_lat * sin_lon
    Z = (N * (1.0 - E2) + h_m) * sin_lat
    return X, Y, Z

def ecef_to_enu(x: float, y: float, z: float, lon0_deg: float, lat0_deg: float, h0_m: float):
    """ECEF XYZ -> ENU at reference (lon0,lat0,h0). Returns (e,n,u) meters."""
    X0, Y0, Z0 = llh_to_ecef(lon0_deg, lat0_deg, h0_m)
    dx = x - X0
    dy = y - Y0
    dz = z - Z0
    lon0 = math.radians(lon0_deg)
    lat0 = math.radians(lat0_deg)
    sin_lat0 = math.sin(lat0)
    cos_lat0 = math.cos(lat0)
    sin_lon0 = math.sin(lon0)
    cos_lon0 = math.cos(lon0)
    # ECEF -> ENU rotation
    e = -sin_lon0 * dx +  cos_lon0 * dy
    n = -sin_lat0 * cos_lon0 * dx - sin_lat0 * sin_lon0 * dy + cos_lat0 * dz
    u =  cos_lat0 * cos_lon0 * dx +  cos_lat0 * sin_lon0 * dy + sin_lat0 * dz
    return e, n, u

# --- Geoid (EGM2008) ----------------------------------------------------------
def _find_egm2008_pgm():
    if EGM2008_PGM_PATH and os.path.isfile(EGM2008_PGM_PATH):
        return EGM2008_PGM_PATH
    for p in (
        r"C:\ProgramData\GeographicLib\geoids\egm2008-1.pgm",
        os.path.expanduser(r"~\geographiclib\geoids\egm2008-1.pgm"),
        r"C:\geographiclib\geoids\egm2008-1.pgm",
    ):
        if os.path.isfile(p):
            return p
    return None

EGM2008_PATH = _find_egm2008_pgm()
_GEOID = None

def get_geoid():
    global _GEOID
    if _GEOID is not None:
        return _GEOID
    if EGM2008_PATH and os.path.isfile(EGM2008_PATH):
        _GEOID = GeoidKarney(EGM2008_PATH)
        return _GEOID
    if FALLBACK_TO_CONSTANT_OFFSET:
        print("[WARN] Geoid PGM not found. Falling back to constant Z offset =", Z_OFFSET_METERS)
        return None
    raise FileNotFoundError(
        "EGM2008 .pgm not found.\n"
        "Set EGM2008_PGM_PATH to your file (e.g. r'C:\\Users\\you\\geographiclib\\geoids\\egm2008-1\\egm2008-1.pgm')\n"
        "or place it at C:\\ProgramData\\GeographicLib\\geoids\\egm2008-1.pgm"
    )

def orthometric_to_ellipsoidal(z_ortho_m: float, lon_deg: float, lat_deg: float):
    """h_ellip = H_orthometric + N(lat,lon)."""
    g = get_geoid()
    if g is None:
        return z_ortho_m + Z_OFFSET_METERS
    N = g.height(lat_deg, lon_deg)  # PyGeodesy expects (lat, lon)
    return z_ortho_m + N

# --- Core conversion -----------------------------------------------------------
def convert_active_object():
    obj = bpy.context.active_object
    if obj is None or obj.type != 'MESH':
        raise RuntimeError("Select a Mesh object first (active object must be a Mesh).")

    if APPLY_OBJECT_TRANSFORMS:
        bpy.ops.object.transform_apply(location=True, rotation=True, scale=True)

    # Duplicate object and data
    bpy.ops.object.duplicate()
    dup = bpy.context.active_object
    dup.name = f"{obj.name}_4326"
    dup.data = obj.data.copy()
    dup.data.name = f"{obj.data.name}_4326"

    me = dup.data
    was_edit = (dup.mode == 'EDIT')
    if was_edit:
        bpy.ops.object.mode_set(mode='OBJECT')

    # Origin lon/lat from 3857
    lon0_deg, lat0_deg = mercator_to_lonlat_deg(ORIGIN_X_3857, ORIGIN_Y_3857)

    mode = OUTPUT_COORDS.upper()
    if mode not in ("ENU_ORIGIN", "ECEF_CENTERED", "LONLAT_DEGREES", "METERS_LOCAL_3857"):
        raise RuntimeError("OUTPUT_COORDS must be one of: ENU_ORIGIN, ECEF_CENTERED, LONLAT_DEGREES, METERS_LOCAL_3857")

    verts = me.vertices
    total_verts = len(verts)
    print(f"Processing {total_verts} vertices...")

    # Process in chunks to avoid memory issues with large meshes
    # Use much smaller chunks for very large models
    if total_verts > 100000:
        chunk_size = 5000  # Very small chunks for large models
    elif total_verts > 50000:
        chunk_size = 10000  # Small chunks for medium models
    else:
        chunk_size = min(25000, max(1000, total_verts // 5))  # Adaptive for smaller models
    
    print(f"Using chunk size: {chunk_size} for {total_verts} vertices")
    
    # If we need an origin reference height and it's not provided, compute mean h_ellip first.
    ref_h = ORIGIN_H_ELLIPSOIDAL
    if mode in ("ENU_ORIGIN", "ECEF_CENTERED") and ref_h is None:
        print("Computing reference height from vertex sample...")
        total_h = 0.0
        count = 0
        # Sample every 100th vertex for reference height to save memory
        sample_step = max(1, total_verts // 1000)
        for i in range(0, total_verts, sample_step):
            v = verts[i]
            x_abs = ORIGIN_X_3857 + v.co.x
            y_abs = ORIGIN_Y_3857 + v.co.y
            lon_deg, lat_deg = mercator_to_lonlat_deg(x_abs, y_abs)
            h_ellip = orthometric_to_ellipsoidal(v.co.z, lon_deg, lat_deg)
            total_h += h_ellip
            count += 1
        ref_h = (total_h / max(count, 1))
        print(f"Reference height computed: {ref_h:.3f} m from {count} sample vertices")

    # Write converted coordinates in chunks with forced garbage collection
    print(f"Converting coordinates in chunks of {chunk_size}...")
    if mode == "LONLAT_DEGREES":
        import gc
        for chunk_start in range(0, total_verts, chunk_size):
            chunk_end = min(chunk_start + chunk_size, total_verts)
            progress = (chunk_end / total_verts) * 100
            print(f"Processing vertices {chunk_start}-{chunk_end} ({progress:.1f}%)...")
            
            for i in range(chunk_start, chunk_end):
                v = verts[i]
                x_abs = ORIGIN_X_3857 + v.co.x
                y_abs = ORIGIN_Y_3857 + v.co.y
                lon_deg, lat_deg = mercator_to_lonlat_deg(x_abs, y_abs)
                h_ellip = orthometric_to_ellipsoidal(v.co.z, lon_deg, lat_deg)
                v.co.x = lon_deg
                v.co.y = lat_deg
                v.co.z = h_ellip
            
            # Force garbage collection after each chunk
            if chunk_end % (chunk_size * 5) == 0:  # Every 5 chunks
                gc.collect()
                print("  └─ Memory cleanup...")
                
        dup.location = (0.0, 0.0, 0.0)

    elif mode == "METERS_LOCAL_3857":
        import gc
        for chunk_start in range(0, total_verts, chunk_size):
            chunk_end = min(chunk_start + chunk_size, total_verts)
            progress = (chunk_end / total_verts) * 100
            print(f"Processing vertices {chunk_start}-{chunk_end} ({progress:.1f}%)...")
            
            for i in range(chunk_start, chunk_end):
                v = verts[i]
                x_abs = ORIGIN_X_3857 + v.co.x
                y_abs = ORIGIN_Y_3857 + v.co.y
                lon_deg, lat_deg = mercator_to_lonlat_deg(x_abs, y_abs)
                h_ellip = orthometric_to_ellipsoidal(v.co.z, lon_deg, lat_deg)
                # keep local 3857 offsets for X/Y
                v.co.z = h_ellip
            
            # Force garbage collection after each chunk
            if chunk_end % (chunk_size * 5) == 0:  # Every 5 chunks
                gc.collect()
                print("  └─ Memory cleanup...")

    elif mode == "ECEF_CENTERED":
        # Compute origin ECEF at (lon0, lat0, ref_h) for centering
        X0, Y0, Z0 = llh_to_ecef(lon0_deg, lat0_deg, ref_h)
        import gc
        for chunk_start in range(0, total_verts, chunk_size):
            chunk_end = min(chunk_start + chunk_size, total_verts)
            progress = (chunk_end / total_verts) * 100
            print(f"Processing vertices {chunk_start}-{chunk_end} ({progress:.1f}%)...")
            
            for i in range(chunk_start, chunk_end):
                v = verts[i]
                x_abs = ORIGIN_X_3857 + v.co.x
                y_abs = ORIGIN_Y_3857 + v.co.y
                lon_deg, lat_deg = mercator_to_lonlat_deg(x_abs, y_abs)
                h_ellip = orthometric_to_ellipsoidal(v.co.z, lon_deg, lat_deg)
                X, Y, Z = llh_to_ecef(lon_deg, lat_deg, h_ellip)
                v.co.x = X - X0
                v.co.y = Y - Y0
                v.co.z = Z - Z0
            
            # Force garbage collection after each chunk
            if chunk_end % (chunk_size * 5) == 0:  # Every 5 chunks
                gc.collect()
                print("  └─ Memory cleanup...")

    elif mode == "ENU_ORIGIN":
        # Local tangent plane at origin (lon0, lat0, ref_h)
        import gc
        for chunk_start in range(0, total_verts, chunk_size):
            chunk_end = min(chunk_start + chunk_size, total_verts)
            progress = (chunk_end / total_verts) * 100
            print(f"Processing vertices {chunk_start}-{chunk_end} ({progress:.1f}%)...")
            
            for i in range(chunk_start, chunk_end):
                v = verts[i]
                x_abs = ORIGIN_X_3857 + v.co.x
                y_abs = ORIGIN_Y_3857 + v.co.y
                lon_deg, lat_deg = mercator_to_lonlat_deg(x_abs, y_abs)
                h_ellip = orthometric_to_ellipsoidal(v.co.z, lon_deg, lat_deg)
                X, Y, Z = llh_to_ecef(lon_deg, lat_deg, h_ellip)
                e, n, u = ecef_to_enu(X, Y, Z, lon0_deg, lat0_deg, ref_h)
                v.co.x, v.co.y, v.co.z = e, n, u
            
            # Force garbage collection after each chunk
            if chunk_end % (chunk_size * 5) == 0:  # Every 5 chunks
                gc.collect()
                print("  └─ Memory cleanup...")

    if was_edit:
        bpy.ops.object.mode_set(mode='EDIT')

    # In all Cartesian modes the object can stay at (0,0,0)
    dup.location = (0.0, 0.0, 0.0)
    print(f"Done: '{dup.name}' created. OUTPUT_COORDS={OUTPUT_COORDS}, Z=ellipsoidal. Ref h={ref_h:.3f} m")
    

# Run
if __name__ == "__main__":
    convert_active_object()
