"""
Project Nandi — Sprayer Coverage Simulator
===========================================
A single-file Streamlit tool for an autonomous weeding / spot-spraying robot.

It models:
  1. Static boom coverage (rasterized union / overlap / gap engine)
  2. Camera FOV & look-ahead geometry (near/far edge, time-to-boom)
  3. Performance / "ARCES" analysis (latency-bound Vmax, acre timing)

Stack: streamlit, numpy, pandas, plotly.graph_objects
Author: R&D Perception Engineering
"""

import math
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# --------------------------------------------------------------------------- #
# Page + lightweight theme tokens (kept consistent across all Plotly charts)
# --------------------------------------------------------------------------- #
st.set_page_config(
    page_title="Sprayer Coverage Simulator",
    page_icon="🌱",
    layout="wide",
)

C_ACTIVE = "rgba(34, 160, 80, 0.30)"      # green fill (active nozzle)
C_ACTIVE_L = "rgba(34, 160, 80, 0.95)"    # green line
C_OFF = "rgba(150, 150, 150, 0.18)"       # gray fill (inactive nozzle)
C_OFF_L = "rgba(120, 120, 120, 0.75)"     # gray line
C_GAP = "rgba(220, 40, 40, 0.28)"         # red fill (gap)
C_GAP_L = "rgba(200, 20, 20, 0.95)"
C_CAM = "rgba(30, 110, 200, 0.22)"        # blue fill (camera FOV)
C_CAM_L = "rgba(30, 110, 200, 0.95)"
C_GROUND = "rgba(90, 70, 50, 0.9)"
C_BOOM = "rgba(40, 40, 40, 0.95)"

PLOT_BG = "rgba(0,0,0,0)"
ACRE_M2 = 4046.86  # 1 international acre in square metres

ENGINE_GRID_RES = 3.0       # mm per raster cell (nominal)
ENGINE_MAX_CELLS = 1_800_000  # safety cap on raster size


# --------------------------------------------------------------------------- #
# CORE ENGINE — rasterized coverage of an arbitrary nozzle array
# --------------------------------------------------------------------------- #
def compute_coverage(x_centers_mm, active, nozzle_type, swath_w_mm,
                     fan_thickness_mm, grid_res=ENGINE_GRID_RES):
    """
    Rasterize the ground footprints of a nozzle array and measure coverage.

    Parameters
    ----------
    x_centers_mm : (N,) array  -> across-boom centre position of each nozzle (mm)
    active       : (N,) bool   -> nozzle ON/OFF state
    nozzle_type  : "Flat Fan" (rectangular footprint) | "Conical" (circular)
    swath_w_mm   : ground width of one nozzle footprint (mm) = 2*H*tan(angle/2)
    fan_thickness_mm : pass-direction depth of a flat-fan footprint (mm)
    grid_res     : raster resolution (mm/cell)

    Returns a dict with raw / union / overlap areas (m^2), per-nozzle table
    (footprint / unique / shared in m^2) and horizontal gap segments (mm).
    """
    x_centers_mm = np.asarray(x_centers_mm, dtype=float)
    active = np.asarray(active, dtype=bool)
    n = len(x_centers_mm)
    radius = swath_w_mm / 2.0

    out = {
        "n_total": n,
        "n_active": int(active.sum()),
        "raw_area_m2": 0.0,
        "union_area_m2": 0.0,
        "overlap_area_m2": 0.0,
        "overlap_pct": 0.0,
        "per_nozzle": [0.0] * n,        # footprint area (m^2) per nozzle index
        "unique": [0.0] * n,
        "shared": [0.0] * n,
        "gaps_mm": [],                  # list of (x_start, x_end, width)
        "covered_x_span_mm": 0.0,
        "grid_res": grid_res,
    }

    act_idx = np.where(active)[0]
    if len(act_idx) == 0 or radius <= 0:
        return out

    # ---- bounding box of the ACTIVE footprints ----
    xmin = float((x_centers_mm[act_idx] - radius).min())
    xmax = float((x_centers_mm[act_idx] + radius).max())
    if nozzle_type == "Conical":
        ymin, ymax = -radius, radius
    else:  # Flat Fan
        half_t = max(fan_thickness_mm, grid_res) / 2.0
        ymin, ymax = -half_t, half_t

    width_mm = xmax - xmin
    depth_mm = ymax - ymin

    # ---- adaptive resolution so the raster never explodes ----
    cells_est = (width_mm / grid_res) * (depth_mm / grid_res)
    if cells_est > ENGINE_MAX_CELLS:
        grid_res = math.sqrt(width_mm * depth_mm / ENGINE_MAX_CELLS)
    out["grid_res"] = grid_res

    nx = max(int(math.ceil(width_mm / grid_res)), 1)
    ny = max(int(math.ceil(depth_mm / grid_res)), 1)
    cell_area_m2 = (grid_res ** 2) / 1.0e6

    xs = xmin + (np.arange(nx) + 0.5) * grid_res      # cell-centre x (mm)
    ys = ymin + (np.arange(ny) + 0.5) * grid_res      # cell-centre y (mm)
    X, Y = np.meshgrid(xs, ys)                        # (ny, nx)

    # ---- pass 1: coverage count grid ----
    count = np.zeros((ny, nx), dtype=np.int16)
    masks = {}
    for i in act_idx:
        cx = x_centers_mm[i]
        if nozzle_type == "Conical":
            m = (X - cx) ** 2 + Y ** 2 <= radius ** 2
        else:
            m = np.abs(X - cx) <= radius           # full depth strip
        masks[i] = m
        count += m

    covered = count >= 1
    union_cells = int(covered.sum())
    union_area = union_cells * cell_area_m2

    # ---- pass 2: per-nozzle unique / shared ----
    raw_area = 0.0
    for i in act_idx:
        m = masks[i]
        fp_cells = int(m.sum())
        uniq_cells = int((m & (count == 1)).sum())
        fp = fp_cells * cell_area_m2
        uq = uniq_cells * cell_area_m2
        out["per_nozzle"][i] = fp
        out["unique"][i] = uq
        out["shared"][i] = fp - uq
        raw_area += fp

    overlap_area = max(raw_area - union_area, 0.0)
    out["raw_area_m2"] = raw_area
    out["union_area_m2"] = union_area
    out["overlap_area_m2"] = overlap_area
    out["overlap_pct"] = (overlap_area / raw_area * 100.0) if raw_area > 0 else 0.0
    out["covered_x_span_mm"] = width_mm

    # ---- horizontal gaps: x-columns NEVER covered (untreated crop strips) ----
    col_covered = covered.any(axis=0)        # (nx,) any y covered for this x
    interior = np.where(col_covered)[0]
    if len(interior) > 0:
        lo, hi = interior.min(), interior.max()
        gap_open = None
        for c in range(lo, hi + 1):
            if not col_covered[c] and gap_open is None:
                gap_open = c
            elif col_covered[c] and gap_open is not None:
                xs_g = xmin + gap_open * grid_res
                xe_g = xmin + c * grid_res
                out["gaps_mm"].append((xs_g, xe_g, xe_g - xs_g))
                gap_open = None
    return out


# --------------------------------------------------------------------------- #
# Geometry helpers
# --------------------------------------------------------------------------- #
def swath_from_geometry(height_mm, angle_deg):
    """Ground width of a single nozzle footprint from boom height + spray angle."""
    a = math.radians(min(max(angle_deg, 1.0), 178.0) / 2.0)
    return 2.0 * height_mm * math.tan(a)


def circle_polygon(cx, cy, r, npts=48):
    t = np.linspace(0, 2 * np.pi, npts)
    return cx + r * np.cos(t), cy + r * np.sin(t)


def fmt(v, unit="", dp=3):
    return f"{v:,.{dp}f}{unit}"


# --------------------------------------------------------------------------- #
# SIDEBAR — Section 1: controls
# --------------------------------------------------------------------------- #
st.sidebar.title("🌱 Project Nandi")
st.sidebar.caption("Sprayer Coverage Simulator")

mode = st.sidebar.selectbox("Simulation Mode", ["Static Boom Coverage"])

st.sidebar.header("Global Parameters")
nozzle_type = st.sidebar.radio("Nozzle Type", ["Flat Fan", "Conical"], horizontal=True)
boom_height = st.sidebar.slider("Boom Height (mm)", 100, 1500, 500, 10)
spray_angle = st.sidebar.slider("Spray Angle (°)", 10, 150, 80, 1)
fan_thickness = st.sidebar.slider(
    "Fan Pass Thickness (mm)", 10, 400, 80, 5,
    help="Pass-direction depth of a flat-fan footprint. Ignored for conical nozzles.",
)
n_nozzles = st.sidebar.slider("Number of Nozzles", 1, 50, 8, 1)

# Single-nozzle ground swath
swath_w = swath_from_geometry(boom_height, spray_angle)

st.sidebar.header("Spacing Configuration")
spacing_mode = st.sidebar.radio("Spacing", ["Uniform", "Custom"], horizontal=True)

if spacing_mode == "Uniform":
    default_sp = max(int(round(swath_w * 0.85)), 20)  # ~15% overlap default
    spacing = st.sidebar.slider("Nozzle Spacing (mm)", 20, 1200, min(default_sp, 1200), 5)
    centers = (np.arange(n_nozzles) - (n_nozzles - 1) / 2.0) * spacing

    cfg = pd.DataFrame({
        "Nozzle": [f"N{i + 1}" for i in range(n_nozzles)],
        "Active": [True] * n_nozzles,
    })
    st.sidebar.caption("Toggle individual nozzles ON / OFF:")
    edited = st.sidebar.data_editor(
        cfg, hide_index=True, use_container_width=True, key="cfg_uniform",
        column_config={
            "Nozzle": st.column_config.TextColumn(disabled=True, width="small"),
            "Active": st.column_config.CheckboxColumn(width="small"),
        },
    )
    active = edited["Active"].to_numpy(dtype=bool)

else:  # Custom
    base_sp = max(int(round(swath_w * 0.85)), 20)
    cfg = pd.DataFrame({
        "Nozzle": [f"N{i + 1}" for i in range(n_nozzles)],
        "Gap to prev (mm)": [0] + [base_sp] * (n_nozzles - 1),
        "Active": [True] * n_nozzles,
    })
    st.sidebar.caption("Edit per-nozzle spacing & ON/OFF state:")
    edited = st.sidebar.data_editor(
        cfg, hide_index=True, use_container_width=True, key="cfg_custom",
        column_config={
            "Nozzle": st.column_config.TextColumn(disabled=True, width="small"),
            "Gap to prev (mm)": st.column_config.NumberColumn(min_value=0, max_value=2000, step=5),
            "Active": st.column_config.CheckboxColumn(width="small"),
        },
    )
    gaps = edited["Gap to prev (mm)"].to_numpy(dtype=float)
    gaps[0] = 0.0
    positions = np.cumsum(gaps)
    centers = positions - positions.mean()  # centre the boom on 0
    active = edited["Active"].to_numpy(dtype=bool)

# --------------------------------------------------------------------------- #
# RUN THE ENGINE ONCE (results reused across all tabs)
# --------------------------------------------------------------------------- #
cov = compute_coverage(centers, active, nozzle_type, swath_w, fan_thickness)
radius = swath_w / 2.0


# =========================================================================== #
# MAIN CANVAS
# =========================================================================== #
st.title("Sprayer Coverage Simulator")
st.caption(
    f"Autonomous weeding & spot-spraying robot · **{nozzle_type}** nozzles · "
    f"boom @ {boom_height} mm · {spray_angle}° spray angle"
)

tab_boom, tab_cam, tab_perf = st.tabs(
    ["🟢 Static Boom Coverage", "📷 Camera FOV & Look-ahead", "🚜 Max Operating Speed"]
)

# --------------------------------------------------------------------------- #
# SECTION 2 — Static Boom Coverage
# --------------------------------------------------------------------------- #
with tab_boom:
    st.subheader("Section 2 · Static Boom Coverage")

    if cov["n_active"] == 0:
        st.warning("All nozzles are OFF — enable at least one nozzle in the sidebar.")
    else:
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Active Nozzles", f"{cov['n_active']} / {cov['n_total']}")
        m2.metric("Swath / Nozzle", f"{swath_w:,.0f} mm")
        m3.metric("Net Covered Area", f"{cov['union_area_m2']:.3f} m²")
        m4.metric("Total Overlap Area", f"{cov['overlap_area_m2']:.3f} m²")
        m5.metric("Overlap %", f"{cov['overlap_pct']:.1f} %")

        # Gap warning
        if cov["gaps_mm"]:
            total_gap = sum(g[2] for g in cov["gaps_mm"])
            st.error(
                f"⚠️ **Coverage gaps detected:** {len(cov['gaps_mm'])} untreated strip(s), "
                f"{total_gap:,.0f} mm total width. Crop in these lanes receives no spray."
            )
        else:
            st.success("✅ No horizontal coverage gaps across the active swath.")

        # ---- Chart 1: Top-down footprint ----
        st.markdown("**Top-Down · Ground Footprint**")
        fig_td = go.Figure()
        if nozzle_type == "Conical":
            ylo, yhi = -radius, radius
        else:
            ylo, yhi = -fan_thickness / 2.0, fan_thickness / 2.0

        for i, cx in enumerate(centers):
            on = bool(active[i])
            fill = C_ACTIVE if on else C_OFF
            line = C_ACTIVE_L if on else C_OFF_L
            if nozzle_type == "Conical":
                px, py = circle_polygon(cx, 0.0, radius)
            else:
                px = [cx - radius, cx + radius, cx + radius, cx - radius, cx - radius]
                py = [ylo, ylo, yhi, yhi, ylo]
            fig_td.add_trace(go.Scatter(
                x=px, y=py, fill="toself", mode="lines",
                line=dict(color=line, width=1.2), fillcolor=fill,
                name=f"N{i+1}", showlegend=False,
                hovertemplate=f"N{i+1} · {'ON' if on else 'OFF'}<extra></extra>",
            ))

        # gaps as red bands
        for (gs, ge, gw) in cov["gaps_mm"]:
            fig_td.add_trace(go.Scatter(
                x=[gs, ge, ge, gs, gs], y=[ylo, ylo, yhi, yhi, ylo],
                fill="toself", mode="lines", line=dict(color=C_GAP_L, width=1),
                fillcolor=C_GAP, name="gap", showlegend=False,
                hovertemplate=f"GAP {gw:,.0f} mm<extra></extra>",
            ))

        fig_td.update_layout(
            height=320, margin=dict(l=10, r=10, t=10, b=10),
            plot_bgcolor=PLOT_BG, paper_bgcolor=PLOT_BG,
            xaxis_title="Across-boom (mm)", yaxis_title="Pass direction (mm)",
        )
        fig_td.update_yaxes(scaleanchor="x", scaleratio=1, zeroline=False)
        fig_td.update_xaxes(zeroline=False)
        st.plotly_chart(fig_td, use_container_width=True)

        # ---- Chart 2: Front profile (spray cones/triangles) ----
        st.markdown("**Front-Profile · Spray Cones (boom → ground)**")
        fig_fp = go.Figure()
        for i, cx in enumerate(centers):
            on = bool(active[i])
            fill = C_ACTIVE if on else C_OFF
            line = C_ACTIVE_L if on else C_OFF_L
            fig_fp.add_trace(go.Scatter(
                x=[cx, cx - radius, cx + radius, cx],
                y=[boom_height, 0, 0, boom_height],
                fill="toself", mode="lines",
                line=dict(color=line, width=1.0), fillcolor=fill,
                showlegend=False,
                hovertemplate=f"N{i+1} · {'ON' if on else 'OFF'}<extra></extra>",
            ))
        # boom + ground lines
        xspan = [min(centers) - radius, max(centers) + radius]
        fig_fp.add_trace(go.Scatter(x=xspan, y=[boom_height, boom_height],
                                    mode="lines", line=dict(color=C_BOOM, width=4),
                                    name="Boom", showlegend=False))
        fig_fp.add_trace(go.Scatter(x=xspan, y=[0, 0], mode="lines",
                                    line=dict(color=C_GROUND, width=3),
                                    name="Ground", showlegend=False))
        fig_fp.update_layout(
            height=300, margin=dict(l=10, r=10, t=10, b=10),
            plot_bgcolor=PLOT_BG, paper_bgcolor=PLOT_BG,
            xaxis_title="Across-boom (mm)", yaxis_title="Height (mm)",
        )
        st.plotly_chart(fig_fp, use_container_width=True)

        # ---- Per-nozzle table with solenoid mapping ----
        st.markdown("**Per-Nozzle Coverage & Solenoid Channel Map**")
        rows = []
        for i, cx in enumerate(centers):
            ch = i + 1
            bank = i // 8 + 1
            ch_in_bank = i % 8 + 1
            rows.append({
                "Nozzle": f"N{i+1}",
                "Center X (mm)": round(cx, 1),
                "State": "ON" if active[i] else "OFF",
                "Footprint (m²)": round(cov["per_nozzle"][i], 4),
                "Unique (m²)": round(cov["unique"][i], 4),
                "Shared (m²)": round(cov["shared"][i], 4),
                "Solenoid Ch": f"CH{ch:02d}",
                "Driver Bank": f"B{bank}.{ch_in_bank}",
            })
        df = pd.DataFrame(rows)
        st.dataframe(df, hide_index=True, use_container_width=True)

        st.caption(
            f"Raster resolution: {cov['grid_res']:.2f} mm/cell · Raw (summed) "
            f"footprint area: {cov['raw_area_m2']:.3f} m² · "
            f"Covered swath span: {cov['covered_x_span_mm']/1000:.3f} m"
        )

# --------------------------------------------------------------------------- #
# SECTION 3 — Camera FOV & Look-ahead Geometry
# --------------------------------------------------------------------------- #
with tab_cam:
    st.subheader("Section 3 · Camera FOV & Look-ahead Geometry")

    c1, c2, c3 = st.columns(3)
    with c1:
        cam_h = st.slider("Camera Height (mm)", 100, 2000, 900, 10)
        cam_fwd = st.slider("Forward Offset (mm)", 0, 2000, 600, 10,
                            help="Camera mounted ahead of the boom by this much.")
    with c2:
        cam_lat = st.slider("Lateral Offset (mm)", -1000, 1000, 0, 10)
        cam_tilt = st.slider("Tilt / Depression (°)", 5, 85, 35, 1,
                            help="Optical-axis angle below horizontal.")
    with c3:
        hfov = st.slider("Horizontal FOV (°)", 20, 120, 70, 1)
        vfov = st.slider("Vertical FOV (°)", 20, 120, 50, 1)
    cam_speed = st.slider("Robot Speed (mm/s)", 50, 4000, 800, 10)

    # --- look-ahead trigonometry ---
    # depression angle of upper / lower FOV edges (below horizontal)
    a_near = cam_tilt + vfov / 2.0     # steepest ray -> nearest ground point
    a_far = cam_tilt - vfov / 2.0      # shallowest ray -> furthest ground point

    a_near = min(a_near, 89.5)
    far_above_horizon = a_far <= 0.5

    gd_near = cam_h / math.tan(math.radians(a_near))            # from camera ground pt
    gd_far = (cam_h / math.tan(math.radians(max(a_far, 0.5))))  # from camera ground pt

    # measured FROM the boom (boom at Y=0, camera at Y=cam_fwd)
    L_near = cam_fwd + gd_near
    L_far = cam_fwd + gd_far

    t_near = L_near / cam_speed if cam_speed > 0 else float("inf")  # s (mm/(mm/s))
    t_far = L_far / cam_speed if cam_speed > 0 else float("inf")

    g1, g2, g3, g4 = st.columns(4)
    g1.metric("Near-edge Look-ahead", f"{L_near:,.0f} mm")
    g2.metric("Far-edge Look-ahead", "∞ (above horizon)" if far_above_horizon else f"{L_far:,.0f} mm")
    g3.metric("Time-to-boom (near)", f"{t_near*1000:,.0f} ms")
    g4.metric("FOV Depth Window", "—" if far_above_horizon else f"{(L_far-L_near):,.0f} mm")

    if far_above_horizon:
        st.warning(
            "The far FOV edge points at/above the horizon (tilt too shallow for the "
            "vertical FOV). Far-edge look-ahead is unbounded — increase tilt to bound it."
        )

    # store for Section 5
    st.session_state["L_near_mm"] = L_near
    st.session_state["L_far_mm"] = L_far if not far_above_horizon else None

    # ---- Chart 1: Side view ----
    st.markdown("**Side-View · Camera FOV projecting forward onto the ground**")
    fig_side = go.Figure()
    x_max = (L_far if not far_above_horizon else L_near * 1.6) * 1.1
    # ground
    fig_side.add_trace(go.Scatter(x=[-200, x_max], y=[0, 0], mode="lines",
                                  line=dict(color=C_GROUND, width=3),
                                  name="Ground", showlegend=False))
    # boom mast at Y=0
    fig_side.add_trace(go.Scatter(x=[0, 0], y=[0, boom_height], mode="lines",
                                  line=dict(color=C_BOOM, width=5),
                                  name="Boom", showlegend=False))
    fig_side.add_trace(go.Scatter(x=[0], y=[boom_height], mode="markers+text",
                                  marker=dict(color=C_BOOM, size=10), text=["Boom"],
                                  textposition="top center", showlegend=False))
    # camera + FOV triangle
    cam_pt = (cam_fwd, cam_h)
    near_pt = (L_near, 0)
    far_pt = (L_far if not far_above_horizon else x_max, 0)
    fig_side.add_trace(go.Scatter(
        x=[cam_pt[0], near_pt[0], far_pt[0], cam_pt[0]],
        y=[cam_pt[1], 0, 0, cam_pt[1]],
        fill="toself", mode="lines", line=dict(color=C_CAM_L, width=1.5),
        fillcolor=C_CAM, name="FOV", showlegend=False,
    ))
    fig_side.add_trace(go.Scatter(x=[cam_fwd], y=[cam_h], mode="markers+text",
                                  marker=dict(color=C_CAM_L, size=12, symbol="square"),
                                  text=["Cam"], textposition="top center", showlegend=False))
    for px, label in [(L_near, "near"), (L_far if not far_above_horizon else None, "far")]:
        if px is not None:
            fig_side.add_trace(go.Scatter(x=[px], y=[0], mode="markers+text",
                                          marker=dict(color=C_CAM_L, size=8),
                                          text=[label], textposition="bottom center",
                                          showlegend=False))
    fig_side.update_layout(height=340, margin=dict(l=10, r=10, t=10, b=10),
                           plot_bgcolor=PLOT_BG, paper_bgcolor=PLOT_BG,
                           xaxis_title="Forward distance from boom (mm)",
                           yaxis_title="Height (mm)")
    st.plotly_chart(fig_side, use_container_width=True)

    # ---- Chart 2: Top-down FOV trapezoid vs boom width ----
    st.markdown("**Top-Down · Camera FOV trapezoid vs boom width**")
    hw_near = gd_near * math.tan(math.radians(hfov / 2.0))
    hw_far = gd_far * math.tan(math.radians(hfov / 2.0))
    far_y = L_far if not far_above_horizon else L_near + (L_near - cam_fwd)
    hw_far_draw = hw_far if not far_above_horizon else hw_near * 1.8

    fig_ftd = go.Figure()
    # FOV trapezoid
    fig_ftd.add_trace(go.Scatter(
        x=[cam_lat - hw_near, cam_lat + hw_near, cam_lat + hw_far_draw, cam_lat - hw_far_draw, cam_lat - hw_near],
        y=[L_near, L_near, far_y, far_y, L_near],
        fill="toself", mode="lines", line=dict(color=C_CAM_L, width=1.5),
        fillcolor=C_CAM, name="FOV footprint", showlegend=False,
    ))
    # boom coverage line at Y=0
    if cov["n_active"] > 0:
        bx0, bx1 = float(min(centers) - radius), float(max(centers) + radius)
    else:
        bx0, bx1 = -500, 500
    fig_ftd.add_trace(go.Scatter(x=[bx0, bx1], y=[0, 0], mode="lines",
                                 line=dict(color=C_BOOM, width=6),
                                 name="Boom width", showlegend=False))
    fig_ftd.add_trace(go.Scatter(x=[cam_lat], y=[cam_fwd], mode="markers+text",
                                 marker=dict(color=C_CAM_L, size=12, symbol="square"),
                                 text=["Cam"], textposition="top center", showlegend=False))
    fig_ftd.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10),
                          plot_bgcolor=PLOT_BG, paper_bgcolor=PLOT_BG,
                          xaxis_title="Lateral / across-boom (mm)",
                          yaxis_title="Forward distance from boom (mm)")
    fig_ftd.update_yaxes(scaleanchor="x", scaleratio=1)
    st.plotly_chart(fig_ftd, use_container_width=True)

    st.caption(
        f"Near ground reach (from camera): {gd_near:,.0f} mm · "
        f"FOV half-width near/far: {hw_near:,.0f}/{hw_far:,.0f} mm. "
        "Horizontal half-width uses small-angle ground projection (visual approximation)."
    )

# --------------------------------------------------------------------------- #
# SECTION 5 — Maximum Operating Speed
# --------------------------------------------------------------------------- #
with tab_perf:
    st.subheader("Section 5 · Maximum Operating Speed")
    st.caption(
        "Give the actuation latency and the look-ahead distance — the simulator "
        "returns the fastest the robot can drive while still spraying every "
        "detected weed before it passes under the boom."
    )

    # ---- Inputs ----
    with st.container(border=True):
        i1, i2 = st.columns(2)
        with i1:
            latency_ms = st.number_input(
                "Total actuation latency (ms)",
                min_value=1.0, max_value=2000.0, value=120.0, step=5.0,
                help="Full pipeline: capture → detect → decide → valve fully open.",
            )
        with i2:
            cam_la = st.session_state.get("L_near_mm", None)
            use_cam = st.toggle(
                "Use camera look-ahead (Section 3)",
                value=False, disabled=cam_la is None,
                help="Pull the near-edge look-ahead computed in the Camera tab.",
            )
            if use_cam and cam_la:
                look_mm = float(cam_la)
                st.number_input("Look-ahead distance (mm)",
                                value=round(look_mm, 1), disabled=True)
            else:
                look_mm = st.number_input(
                    "Look-ahead distance (mm)",
                    min_value=10.0, max_value=10000.0,
                    value=float(round(cam_la)) if cam_la else 600.0, step=10.0,
                    help="Distance from the boom to where a weed is first detected.",
                )

    # ---- Core result:  V = look-ahead / latency   (mm/ms == m/s) ----
    v_max_ms = look_mm / latency_ms if latency_ms > 0 else 0.0
    v_max_kmh = v_max_ms * 3.6

    with st.container(border=True):
        st.markdown("#### 🚜 Maximum operating speed")
        r1, r2 = st.columns(2)
        r1.metric("Vmax", f"{v_max_ms:.2f} m/s")
        r2.metric("Vmax", f"{v_max_kmh:.2f} km/h")
        st.caption(
            f"V = look-ahead ÷ latency = {look_mm:,.0f} mm ÷ {latency_ms:.0f} ms. "
            "At this speed a weed seen at the look-ahead line reaches the boom "
            "exactly as the nozzle fires — so run a little below it for margin."
        )

    # ---- Sensitivity: how Vmax moves with latency (fixed look-ahead) ----
    st.markdown("**Max speed vs. latency** (at the current look-ahead)")
    lat_axis = np.linspace(max(latency_ms * 0.3, 5.0), latency_ms * 2.0, 80)
    v_axis_kmh = (look_mm / lat_axis) * 3.6

    fig_v = go.Figure()
    fig_v.add_trace(go.Scatter(
        x=lat_axis, y=v_axis_kmh, mode="lines",
        line=dict(color=C_CAM_L, width=2.5),
        fill="tozeroy", fillcolor="rgba(30,110,200,0.10)", showlegend=False,
    ))
    fig_v.add_trace(go.Scatter(
        x=[latency_ms], y=[v_max_kmh], mode="markers+text",
        marker=dict(color=C_ACTIVE_L, size=13),
        text=[f"  {v_max_kmh:.1f} km/h"], textposition="middle right",
        showlegend=False,
    ))
    fig_v.update_layout(
        height=300, margin=dict(l=10, r=10, t=10, b=10),
        plot_bgcolor=PLOT_BG, paper_bgcolor=PLOT_BG,
        xaxis_title="Latency (ms)", yaxis_title="Max speed (km/h)",
    )
    fig_v.update_xaxes(zeroline=False)
    fig_v.update_yaxes(zeroline=False)
    st.plotly_chart(fig_v, use_container_width=True)


st.sidebar.divider()
st.sidebar.caption("Engine: rasterized union/overlap/gap · areas in m² · grid auto-scaled.")
