# suivi_app.py
# Dashboard de Suivi DGID — Reflex
# v2.7 — CORRECTIONS : last_recorded_week depuis DB + durée négative abs + graphique pleine largeur

import reflex as rx
from datetime import datetime, timedelta
import math
import asyncio
import duckdb
import os

# =============================================================================
# CONSTANTES DE STYLE
# =============================================================================
GOLD = "#D4AF37"
BROWN = "#4A3525"
CREAM = "#FAF7F2"
WHITE = "#FFFFFF"
TEXT_MAIN = "#1E1E1E"
TEXT_MUTED = "#8A8A8A"
BORDER_HDR = "#E8E8E8"
BORDER_CARD = "#F0F0F0"
ICON_BG = "#F5F0EB"
CREAM_CARD = "#FEF7E0"
BROWN_DARK = "#4A3525"
GREEN_LIGHT = "#E6F4EA"
GREEN_TEXT = "#1E8E3E"
RED_LIGHT = "#FCE8E6"
RED_TEXT = "#D93025"
CADASTRE_BG = "#F5E6C8"
CADASTRE_TEXT = "#8B6914"
CONSERVATION_BG = "#E8E8E8"
CONSERVATION_TEXT = "#666666"
DOMAINES_BG = "#E6D5C3"
DOMAINES_TEXT = "#6B4423"
SHADOW = "0 2px 12px rgba(0,0,0,0.04)"

DUCKDB_PATH = os.getenv("DUCKDB_PATH", "/app/duckdb_store/analytics.duckdb")


def _build_where(filters: dict) -> tuple[str, list]:
    clauses = []
    params = []
    if filters.get("start_date") and filters.get("end_date"):
        clauses.append("registration_date::DATE BETWEEN ?::DATE AND ?::DATE")
        params.extend([filters["start_date"], filters["end_date"]])
    if filters.get("service") and filters["service"] != "Tous":
        clauses.append("service_origine = ?")
        params.append(filters["service"])
    if filters.get("csf") and filters["csf"] != "Tous":
        clauses.append("csf_geographique = ?")
        params.append(filters["csf"])
    if filters.get("type") and filters["type"] != "Tous":
        clauses.append("motif_enregistrement = ?")
        params.append(filters["type"])
    where_sql = " AND ".join(clauses) if clauses else "1=1"
    return where_sql, params


def _safe_float(val, default=0.0):
    if val is None:
        return default
    try:
        return float(val)
    except (TypeError, ValueError):
        return default


def _safe_int(val, default=0):
    if val is None:
        return default
    try:
        return int(val)
    except (TypeError, ValueError):
        return default


class DashboardState(rx.State):
    start_date: str = ""
    end_date: str = ""
    selected_service: str = "Tous"
    selected_csf: str = "Tous"
    selected_type: str = "Tous"
    flux_period: str = "Mois"
    is_loading: bool = False
    objectif_attendu: int = 1000
    db_error: str = ""
    type_options: list[str] = ["Tous"]

    kpi_objectif: int = 1000
    kpi_taux_indexation: float = 0.0
    kpi_taux_num: float = 0.0
    kpi_total_dossiers: int = 0
    kpi_attente_num: int = 0
    kpi_pieces_indexees: int = 0
    kpi_taux_restitution: float = 0.0
    kpi_duree_moyenne: int = 0
    kpi_dossiers_jour: float = 0.0
    kpi_top_archiviste: str = "—"
    kpi_top_total: int = 0

    flux_data: list[dict] = []
    flux_svg_html: str = ""
    volumes_bureau: list[dict] = []
    donut_html: str = ""
    productivite: list[dict] = []
    productivite_count: int = 0
    derniers_enregistrements: list[dict] = []

    # =================================================================
    # CORRECTION : Graphique SVG pleine largeur
    # =================================================================
    def _build_flux_svg(self, data_list: list[dict]) -> str:
        if not data_list:
            return '<svg width="100%" height="100%" viewBox="0 0 700 280" xmlns="http://www.w3.org/2000/svg"><text x="50%" y="50%" text-anchor="middle" fill="#AAA" font-size="14" font-family="sans-serif">Aucune donnée</text></svg>'
        
        # Dimensions adaptatives au conteneur
        width, height, padding = 700, 280, 45
        chart_w = width - padding * 2
        chart_h = height - padding * 2
        max_val = max(d["value"] for d in data_list) * 1.2 or 100
        n = len(data_list)
        step_x = chart_w / max(n - 1, 1)
        points = []
        for i, d in enumerate(data_list):
            x = padding + i * step_x
            y = padding + chart_h - (d["value"] / max_val * chart_h)
            points.append((x, y))
        path_d = f"M {points[0][0]:.1f} {points[0][1]:.1f}"
        for x, y in points[1:]:
            path_d += f" L {x:.1f} {y:.1f}"
        fill_d = path_d + f" L {points[-1][0]:.1f} {padding + chart_h:.1f} L {points[0][0]:.1f} {padding + chart_h:.1f} Z"
        grid_lines = ""
        for i in range(5):
            y_line = padding + (chart_h / 4) * i
            grid_lines += f'<line x1="{padding}" y1="{y_line:.1f}" x2="{width - padding}" y2="{y_line:.1f}" stroke="#F0F0F0" stroke-dasharray="4,4" stroke-width="1"/>'
        x_labels = ""
        for i, d in enumerate(data_list):
            if i % 2 == 0 or n <= 8:
                x_pos = padding + i * step_x
                x_labels += f'<text x="{x_pos:.1f}" y="{height - 12}" text-anchor="middle" font-size="10" fill="#AAA" font-family="sans-serif">{d["date"]}</text>'
        circles = ""
        for x, y in points:
            circles += f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3.5" fill="{GOLD}" stroke="{WHITE}" stroke-width="2"/>'
        # viewBox ajouté pour scaling responsive
        return f"""<svg width="100%" height="100%" viewBox="0 0 {width} {height}" xmlns="http://www.w3.org/2000/svg" preserveAspectRatio="none">
            <defs><linearGradient id="areaGrad" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="{GOLD}" stop-opacity="0.25"/><stop offset="100%" stop-color="{GOLD}" stop-opacity="0.02"/></linearGradient></defs>
            {grid_lines}<path d="{fill_d}" fill="url(#areaGrad)" stroke="none"/><path d="{path_d}" fill="none" stroke="{GOLD}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>{circles}{x_labels}
        </svg>"""

    def _build_donut_html(self, vol_data: list[dict]) -> str:
        if not vol_data:
            return '<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#AAA;font-family:sans-serif;">Aucune donnée</div>'
        total = sum(d["volume"] for d in vol_data)
        cx, cy, r = 150, 150, 85
        circumference = 2 * math.pi * r
        circles_html = ""
        offset = 0
        for d in vol_data:
            dash = (d["volume"] / total) * circumference
            gap = circumference - dash
            circles_html += f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="none" stroke="{d["color"]}" stroke-width="32" stroke-dasharray="{dash:.2f} {gap:.2f}" stroke-dashoffset="{-offset:.2f}" transform="rotate(-90 {cx} {cy})" stroke-linecap="butt"/>'
            offset += dash
        legend_items = ""
        for d in vol_data:
            legend_items += f'<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;font-family:sans-serif;font-size:13px;"><div style="display:flex;align-items:center;gap:8px;"><div style="width:10px;height:10px;border-radius:2px;background:{d["color"]};"></div><span style="color:#333;">{d["bureau"]}</span></div><div style="display:flex;align-items:center;gap:12px;"><span style="font-weight:600;color:#1E1E1E;">{d["volume"]}</span><span style="color:#8A8A8A;font-size:12px;">{d["pct"]}%</span></div></div>'
        return f'<div style="display:flex;align-items:center;gap:20px;width:100%;height:100%;padding:20px;box-sizing:border-box;"><div style="position:relative;width:300px;height:300px;flex-shrink:0;"><svg width="300" height="300" viewBox="0 0 300 300">{circles_html}</svg><div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);text-align:center;"><div style="font-size:22px;font-weight:bold;color:#1E1E1E;font-family:sans-serif;">{total}</div><div style="font-size:11px;color:#8A8A8A;font-family:sans-serif;">Total</div></div></div><div style="flex:1;min-width:140px;">{legend_items}</div></div>'

    async def load_data(self):
        self.is_loading = True
        self.db_error = ""
        try:
            await asyncio.to_thread(self._fetch_all_data)
        except Exception as e:
            self.db_error = f"Erreur base de données : {str(e)}"
            print(self.db_error)
        finally:
            self.is_loading = False

    def _fetch_all_data(self):
        if not os.path.exists(DUCKDB_PATH):
            raise FileNotFoundError(f"Base DuckDB introuvable : {DUCKDB_PATH}")
        con = duckdb.connect(DUCKDB_PATH, read_only=True)
        try:
            filters = {"start_date": self.start_date, "end_date": self.end_date, "service": self.selected_service, "csf": self.selected_csf, "type": self.selected_type}
            where_sql, params = _build_where(filters)
            try:
                type_rows = con.execute("SELECT DISTINCT motif_enregistrement FROM fact_suivi_global WHERE motif_enregistrement IS NOT NULL ORDER BY motif_enregistrement").fetchall()
                self.type_options = ["Tous"] + [r[0] for r in type_rows if r[0]]
            except Exception:
                self.type_options = ["Tous", "Entrée", "Sortie"]
            row = con.execute(f"SELECT COUNT(DISTINCT numero_dossier) FROM fact_suivi_global WHERE {where_sql}", params).fetchone()
            self.kpi_total_dossiers = _safe_int(row[0])
            row = con.execute(f"SELECT COUNT(DISTINCT numero_dossier) FROM fact_suivi_global WHERE is_for_scan = TRUE AND date_indexation IS NULL AND {where_sql}", params).fetchone()
            self.kpi_attente_num = _safe_int(row[0])
            row = con.execute(f"SELECT COALESCE(SUM(count_item), 0) FROM fact_suivi_global WHERE date_indexation IS NOT NULL AND {where_sql}", params).fetchone()
            self.kpi_pieces_indexees = _safe_int(row[0])
            row = con.execute(f"SELECT COALESCE(COUNT(DISTINCT CASE WHEN date_retour IS NOT NULL THEN numero_dossier END) * 100.0 / NULLIF(COUNT(DISTINCT numero_dossier), 0), 0) FROM fact_suivi_global WHERE {where_sql}", params).fetchone()
            self.kpi_taux_restitution = round(_safe_float(row[0]), 1)
            
            # =================================================================
            # CORRECTION : ABS pour éviter les durées négatives
            # =================================================================
            row = con.execute(f"SELECT COALESCE(AVG(ABS(EXTRACT(EPOCH FROM (date_indexation - registration_date))) / 3600), 0) FROM fact_suivi_global WHERE date_indexation IS NOT NULL AND registration_date IS NOT NULL AND {where_sql}", params).fetchone()
            self.kpi_duree_moyenne = round(_safe_float(row[0]))
            
            row = con.execute(f"SELECT COALESCE(COUNT(DISTINCT numero_dossier) * 1.0 / NULLIF(COUNT(DISTINCT registration_date::DATE), 0), 0) FROM fact_suivi_global WHERE {where_sql}", params).fetchone()
            self.kpi_dossiers_jour = round(_safe_float(row[0]), 1)
            row = con.execute(f"SELECT COUNT(DISTINCT numero_dossier) FROM fact_suivi_global WHERE date_indexation IS NOT NULL AND {where_sql}", params).fetchone()
            indexes = _safe_int(row[0])
            self.kpi_taux_indexation = round((indexes / max(self.objectif_attendu, 1)) * 100, 1)
            row = con.execute(f"SELECT COUNT(DISTINCT numero_dossier) FROM fact_suivi_global WHERE is_for_scan = TRUE AND {where_sql}", params).fetchone()
            numerises = _safe_int(row[0])
            self.kpi_taux_num = round((numerises / max(self.objectif_attendu, 1)) * 100, 1)
            row = con.execute(f"SELECT operateur, COUNT(DISTINCT numero_dossier) as total FROM fact_suivi_global WHERE {where_sql} GROUP BY operateur ORDER BY total DESC LIMIT 1", params).fetchone()
            if row and row[0]:
                self.kpi_top_archiviste = str(row[0])
                self.kpi_top_total = _safe_int(row[1])
            else:
                self.kpi_top_archiviste = "—"
                self.kpi_top_total = 0

            # =================================================================
            # CORRECTION SQL : GROUP BY identique au SELECT + CAST pour TIMESTAMP_NS
            # =================================================================
            if self.flux_period == "Mois":
                period_expr = "STRFTIME('%Y-%m', registration_date::TIMESTAMP)"
            else:
                period_expr = "STRFTIME('%Y', registration_date::TIMESTAMP) || '-Q' || QUARTER(registration_date::TIMESTAMP)"

            flux_sql = (
                f"SELECT {period_expr} as periode, COUNT(DISTINCT numero_dossier) as value "
                f"FROM fact_suivi_global WHERE registration_date IS NOT NULL AND {where_sql} "
                f"GROUP BY {period_expr} ORDER BY {period_expr}"
            )
            rows = con.execute(flux_sql, params).fetchall()
            self.flux_data = [{"date": str(r[0]), "value": _safe_int(r[1])} for r in rows]
            self.flux_svg_html = self._build_flux_svg(self.flux_data)

            rows = con.execute(f"SELECT service_origine as bureau, COUNT(DISTINCT numero_dossier) as volume FROM fact_suivi_global WHERE service_origine IS NOT NULL AND {where_sql} GROUP BY service_origine ORDER BY volume DESC", params).fetchall()
            total_vol = sum(r[1] for r in rows) if rows else 1
            colors = {"Cadastre": GOLD, "Conservation": BROWN, "Domaines": "#A0826D"}
            self.volumes_bureau = [{"bureau": str(r[0]), "volume": _safe_int(r[1]), "pct": round(_safe_int(r[1]) / total_vol * 100, 1), "color": colors.get(str(r[0]), "#999")} for r in rows]
            self.donut_html = self._build_donut_html(self.volumes_bureau)
            rows = con.execute(f"SELECT operateur, service_origine as bureau, COUNT(DISTINCT numero_dossier) as dossiers, COALESCE(COUNT(DISTINCT numero_dossier) * 1.0 / NULLIF(COUNT(DISTINCT registration_date::DATE), 0), 0) as moyenne FROM fact_suivi_global WHERE operateur IS NOT NULL AND {where_sql} GROUP BY operateur, service_origine ORDER BY dossiers DESC LIMIT 11", params).fetchall()
            max_dossiers = max(_safe_int(r[2]) for r in rows) if rows else 1
            self.productivite = [{"rang": i + 1, "agent": str(r[0]), "csf": self.selected_csf if self.selected_csf != "Tous" else "Dakar Plateau", "bureau": str(r[1]), "dossiers": _safe_int(r[2]), "moyenne": round(_safe_float(r[3]), 1), "pct": round((_safe_int(r[2]) / max_dossiers) * 100, 1)} for i, r in enumerate(rows)]
            self.productivite_count = len(self.productivite)
            rows = con.execute(f"SELECT numero_dossier as code, COALESCE(NULLIF(folder_label, ''), motif_enregistrement, '—') as desc, COALESCE(motif_enregistrement, '—') as type, registration_date FROM fact_suivi_global WHERE {where_sql} ORDER BY registration_date DESC LIMIT 5", params).fetchall()
            self.derniers_enregistrements = [{"code": str(r[0]), "desc": str(r[1]), "type_badge": "Entrée" if "entrée" in str(r[2]).lower() or "réception" in str(r[2]).lower() else "Sortie"} for r in rows]
        finally:
            con.close()

    async def set_start_date(self, value: str):
        self.start_date = value
        await self.load_data()

    async def set_end_date(self, value: str):
        self.end_date = value
        await self.load_data()

    async def set_service(self, value: str):
        self.selected_service = value
        await self.load_data()

    async def set_csf(self, value: str):
        self.selected_csf = value
        await self.load_data()

    async def set_type(self, value: str):
        self.selected_type = value
        await self.load_data()

    async def reset_service(self):
        self.selected_service = "Tous"
        await self.load_data()

    async def reset_csf(self):
        self.selected_csf = "Tous"
        await self.load_data()

    async def reset_type(self):
        self.selected_type = "Tous"
        await self.load_data()

    async def select_current_week(self):
        today = datetime.today()
        start = today - timedelta(days=today.weekday())
        self.start_date = start.strftime("%Y-%m-%d")
        self.end_date = today.strftime("%Y-%m-%d")
        await self.load_data()

    # =================================================================
    # CORRECTION : Dernière semaine ENREGISTRÉE dans la base (pas calendaire)
    # =================================================================
    async def select_last_recorded_week(self):
        if not os.path.exists(DUCKDB_PATH):
            self.db_error = "Base de données inaccessible"
            return
        con = duckdb.connect(DUCKDB_PATH, read_only=True)
        try:
            # Trouver la date d'enregistrement la plus récente
            max_date_row = con.execute(
                "SELECT MAX(registration_date::DATE) FROM fact_suivi_global WHERE registration_date IS NOT NULL"
            ).fetchone()
            if max_date_row and max_date_row[0]:
                last_recorded = max_date_row[0]
                # Début de la semaine (lundi) de cette date
                start = last_recorded - timedelta(days=last_recorded.weekday())
                end = start + timedelta(days=6)
                self.start_date = start.strftime("%Y-%m-%d")
                self.end_date = end.strftime("%Y-%m-%d")
            else:
                # Fallback : semaine calendaire dernière
                today = datetime.today()
                last_week = today - timedelta(days=7)
                start = last_week - timedelta(days=last_week.weekday())
                end = start + timedelta(days=6)
                self.start_date = start.strftime("%Y-%m-%d")
                self.end_date = end.strftime("%Y-%m-%d")
        finally:
            con.close()
        await self.load_data()

    async def reset_date_filters(self):
        self.start_date = ""
        self.end_date = ""
        await self.load_data()

    async def set_flux_mois(self):
        self.flux_period = "Mois"
        await self.load_data()

    async def set_flux_trimestre(self):
        self.flux_period = "Trimestre"
        await self.load_data()

    async def set_objectif_attendu(self, value: str):
        try:
            self.objectif_attendu = int(value)
            self.kpi_objectif = int(value)
        except ValueError:
            pass
        await self.load_data()

    async def on_load(self):
        await self.select_current_week()

    def export_excel(self):
        print("Export Excel demandé")


def bureau_badge(bureau: str) -> rx.Component:
    bg = rx.cond(bureau == "Cadastre", CADASTRE_BG, rx.cond(bureau == "Conservation", CONSERVATION_BG, DOMAINES_BG))
    color = rx.cond(bureau == "Cadastre", CADASTRE_TEXT, rx.cond(bureau == "Conservation", CONSERVATION_TEXT, DOMAINES_TEXT))
    return rx.badge(bureau, bg=bg, color=color, border_radius="6px", padding_x="10px", padding_y="2px", font_size="11px", font_weight="500")


def type_badge(type_badge_str: str) -> rx.Component:
    bg = rx.cond(type_badge_str == "Entrée", GREEN_LIGHT, RED_LIGHT)
    color = rx.cond(type_badge_str == "Entrée", GREEN_TEXT, RED_TEXT)
    return rx.badge(type_badge_str, bg=bg, color=color, border_radius="20px", padding_x="12px", padding_y="3px", font_size="11px", font_weight="500")


def header() -> rx.Component:
    return rx.hstack(
        rx.hstack(
            rx.box(rx.icon("landmark", size=22, color=GOLD), bg=ICON_BG, padding="10px", border_radius="10px", display="flex", align_items="center", justify_content="center"),
            rx.vstack(rx.heading("Dashboard de Suivi", size="5", color=TEXT_MAIN, font_weight="bold", line_height="1.2"), rx.text("Numérisation & Indexation — CSF Dakar Plateau", color=TEXT_MUTED, font_size="12px", line_height="1.2"), spacing="1", align_items="start"),
            spacing="3", align_items="center",
        ),
        rx.spacer(),
        rx.button(rx.hstack(rx.icon("download", size=14, color=BROWN), rx.text("Exporter en Excel", font_size="13px", color=BROWN, font_weight="500"), spacing="2", align_items="center"), bg=WHITE, color=BROWN, border=f"1px solid {BROWN}", border_radius="8px", padding="8px 18px", height="40px", cursor="pointer", _hover={"bg": BROWN, "color": WHITE, "transition": "all 0.2s ease"}, on_click=DashboardState.export_excel),
        width="100%", padding="16px 32px", bg=WHITE, border_bottom=f"1px solid {BORDER_HDR}", align_items="center",
    )


# =============================================================================
# CORRECTION UI : filtres alignés à droite, espacement corrigé
# =============================================================================
def filter_bar() -> rx.Component:
    label_style = {"color": TEXT_MUTED, "font_size": "10px", "font_weight": "600", "letter_spacing": "0.5px", "text_transform": "uppercase"}
    input_style = {"border_radius": "8px", "border": f"1px solid {BORDER_CARD}", "padding": "8px 12px", "font_size": "13px", "color": TEXT_MAIN, "width": "140px", "height": "38px"}
    btn_ghost = {"variant": "ghost", "size": "2", "color": BROWN, "font_size": "12px", "border_radius": "20px", "border": f"1px solid {BORDER_CARD}", "padding": "8px 16px", "height": "38px", "cursor": "pointer", "_hover": {"bg": BROWN, "color": WHITE, "transition": "all 0.2s ease"}}
    btn_muted = {"variant": "ghost", "size": "2", "color": TEXT_MUTED, "font_size": "12px", "border_radius": "20px", "border": f"1px solid {BORDER_CARD}", "padding": "8px 16px", "height": "38px", "cursor": "pointer", "_hover": {"color": BROWN, "border_color": BROWN, "transition": "all 0.2s ease"}}
    select_style = {"border_radius": "8px", "border": f"1px solid {BORDER_CARD}", "font_size": "13px", "color": TEXT_MAIN, "width": "130px", "height": "38px", "bg": WHITE}

    return rx.box(
        rx.hstack(
            # --- Gauche : dates + boutons rapides ---
            rx.hstack(
                rx.vstack(rx.text("DU", **label_style), rx.input(type="date", value=DashboardState.start_date, on_change=DashboardState.set_start_date, **input_style), spacing="1", align_items="start"),
                rx.vstack(rx.text("AU", **label_style), rx.input(type="date", value=DashboardState.end_date, on_change=DashboardState.set_end_date, **input_style), spacing="1", align_items="start"),
                rx.button("Semaine actuelle", on_click=DashboardState.select_current_week, **btn_ghost),
                rx.button("Dernière semaine", on_click=DashboardState.select_last_recorded_week, **btn_ghost),
                rx.button("Réinitialiser", on_click=DashboardState.reset_date_filters, **btn_muted),
                spacing="4",
                align_items="end",
            ),
            rx.spacer(),
            # --- Droite : filtres BUREAU / CSF / TYPE ---
            rx.hstack(
                rx.vstack(rx.text("BUREAU", **label_style), rx.select(["Tous", "Cadastre", "Conservation", "Domaines"], value=DashboardState.selected_service, on_change=DashboardState.set_service, **select_style), spacing="1", align_items="start"),
                rx.vstack(rx.text("CSF", **label_style), rx.select(["Tous", "Dakar Plateau", "Ngor-Almadies", "Mbour"], value=DashboardState.selected_csf, on_change=DashboardState.set_csf, **select_style), spacing="1", align_items="start"),
                rx.vstack(rx.text("TYPE", **label_style), rx.select(DashboardState.type_options, value=DashboardState.selected_type, on_change=DashboardState.set_type, **select_style), spacing="1", align_items="start"),
                spacing="5",
                align_items="end",
            ),
            width="100%",
            align_items="end",
            max_width="1400px",
            margin="0 auto",
        ),
        padding="14px 32px",
        bg=WHITE,
        border_bottom=f"1px solid {BORDER_HDR}",
    )


def objectives_section() -> rx.Component:
    return rx.vstack(
        rx.hstack(
            rx.hstack(rx.icon("target", size=14, color=BROWN), rx.text("PILOTAGE SUPERVISEUR", color=BROWN, font_size="10px", font_weight="700", letter_spacing="1px"), spacing="2", align_items="center"),
            rx.spacer(),
            rx.vstack(rx.text("DOSSIERS ATTENDUS", color=TEXT_MUTED, font_size="10px", font_weight="600", letter_spacing="0.5px"), rx.input(value=DashboardState.objectif_attendu.to_string(), on_change=DashboardState.set_objectif_attendu, width="100px", height="32px", font_size="14px", font_weight="600", text_align="center", border_radius="8px", border=f"1px solid {BORDER_CARD}"), spacing="1", align_items="end"),
            width="100%", align_items="start", margin_bottom="8px",
        ),
        rx.heading("Objectifs de la période", size="4", color=TEXT_MAIN, font_weight="bold", margin_bottom="16px"),
        rx.grid(
            rx.box(rx.vstack(rx.text("Objectif fixé", color=TEXT_MUTED, font_size="12px", font_weight="500"), rx.hstack(rx.text(DashboardState.kpi_objectif.to_string(), font_size="32px", font_weight="bold", color=TEXT_MAIN, line_height="1"), rx.text("dossiers", font_size="14px", color=TEXT_MUTED, margin_top="8px"), spacing="2", align_items="end"), rx.text("CSF actuel : Tous", color=TEXT_MUTED, font_size="11px", margin_top="4px"), spacing="1", align_items="start"), padding="20px", bg=CREAM_CARD, border_radius="12px", border=f"1px solid #F5E6C8", width="100%", height="100%"),
            rx.box(rx.vstack(rx.hstack(rx.text("Taux d'indexation", color=TEXT_MUTED, font_size="12px", font_weight="500"), rx.spacer(), rx.box(rx.icon("layers", size=16, color=GOLD), bg=ICON_BG, padding="6px", border_radius="8px"), width="100%", align_items="center"), rx.hstack(rx.text(f"{DashboardState.kpi_taux_indexation}", font_size="28px", font_weight="bold", color=TEXT_MAIN, line_height="1"), rx.text("%", font_size="16px", color=TEXT_MUTED, margin_top="4px"), spacing="1", align_items="end"), rx.box(rx.box(width="100%", height="6px", bg=GOLD, border_radius="3px"), width="100%", bg="#F0E6D0", border_radius="3px", height="6px", margin_top="8px"), rx.hstack(rx.text("Indexés vs objectif de la période", color=TEXT_MUTED, font_size="10px"), rx.spacer(), rx.text("×7,8", color=GOLD, font_size="10px", font_weight="bold"), width="100%", margin_top="4px"), spacing="2", align_items="start", width="100%"), padding="20px", bg=WHITE, border_radius="12px", border=f"1px solid {BORDER_CARD}", box_shadow=SHADOW, width="100%", height="100%"),
            rx.box(rx.vstack(rx.hstack(rx.text("Taux de numérisation", color=TEXT_MUTED, font_size="12px", font_weight="500"), rx.spacer(), rx.box(rx.icon("scan", size=16, color=GOLD), bg=ICON_BG, padding="6px", border_radius="8px"), width="100%", align_items="center"), rx.hstack(rx.text(f"{DashboardState.kpi_taux_num}", font_size="28px", font_weight="bold", color=TEXT_MAIN, line_height="1"), rx.text("%", font_size="16px", color=TEXT_MUTED, margin_top="4px"), spacing="1", align_items="end"), rx.box(rx.box(width="60.9%", height="6px", bg=GOLD, border_radius="3px"), width="100%", bg="#F0E6D0", border_radius="3px", height="6px", margin_top="8px"), rx.text("Numérisés vs objectif de la période", color=TEXT_MUTED, font_size="10px", margin_top="4px"), spacing="2", align_items="start", width="100%"), padding="20px", bg=WHITE, border_radius="12px", border=f"1px solid {BORDER_CARD}", box_shadow=SHADOW, width="100%", height="100%"),
            columns="3", spacing="4", width="100%",
        ),
        width="100%", spacing="0",
    )


def metrics_section() -> rx.Component:
    def std_card(title, value, subtitle, icon_name):
        return rx.box(rx.vstack(rx.hstack(rx.text(title, color=TEXT_MUTED, font_size="12px", font_weight="500"), rx.spacer(), rx.box(rx.icon(icon_name, size=16, color=GOLD), bg=ICON_BG, padding="6px", border_radius="8px"), width="100%", align_items="center"), rx.text(value, font_size="24px", font_weight="bold", color=TEXT_MAIN, line_height="1", margin_top="8px"), rx.text(subtitle, color=TEXT_MUTED, font_size="11px", margin_top="4px"), spacing="1", align_items="start", width="100%"), padding="18px", bg=WHITE, border_radius="12px", border=f"1px solid {BORDER_CARD}", box_shadow=SHADOW, width="100%", height="100%")
    def dark_card(title, value, subtitle, icon_name):
        return rx.box(rx.vstack(rx.hstack(rx.text(title, color="rgba(255,255,255,0.7)", font_size="12px", font_weight="500"), rx.spacer(), rx.box(rx.icon(icon_name, size=16, color=GOLD), bg="rgba(255,255,255,0.1)", padding="6px", border_radius="8px"), width="100%", align_items="center"), rx.text(value, font_size="24px", font_weight="bold", color=WHITE, line_height="1", margin_top="8px"), rx.text(subtitle, color="rgba(255,255,255,0.7)", font_size="11px", margin_top="4px"), spacing="1", align_items="start", width="100%"), padding="18px", bg=BROWN_DARK, border_radius="12px", width="100%", height="100%")
    return rx.grid(
        std_card("Total de dossiers", DashboardState.kpi_total_dossiers.to_string(), "Dossiers uniques", "folder"),
        std_card("Attente numérisation", DashboardState.kpi_attente_num.to_string(), "Flux restant", "scan"),
        std_card("Pièces indexées", DashboardState.kpi_pieces_indexees.to_string(), "Feuillets indexés", "file-text"),
        dark_card("Taux de restitution", f"{DashboardState.kpi_taux_restitution} %", "Dossiers restitués", "refresh-cw"),
        std_card("Durée moyenne scan", f"{DashboardState.kpi_duree_moyenne} h", "≈ 21 j de traitement", "clock"),
        std_card("Dossiers par jour", f"{DashboardState.kpi_dossiers_jour} /j", "Cadence moyenne", "trending-up"),
        std_card("Top archiviste", DashboardState.kpi_top_archiviste, f"Total : {DashboardState.kpi_top_total} dossiers", "user"),
        columns="4", spacing="4", width="100%", margin_top="16px",
    )


# =============================================================================
# CORRECTION : Graphique SVG avec preserveAspectRatio="none" pour pleine largeur
# =============================================================================
def activity_chart() -> rx.Component:
    return rx.box(
        rx.vstack(
            rx.hstack(
                rx.hstack(rx.icon("activity", size=16, color=GOLD), rx.text("Flux d'activité", color=TEXT_MAIN, font_size="14px", font_weight="600"), spacing="2", align_items="center"),
                rx.spacer(),
                rx.hstack(
                    rx.button("Mois", on_click=DashboardState.set_flux_mois, bg=rx.cond(DashboardState.flux_period == "Mois", "#F0E6D0", WHITE), color=rx.cond(DashboardState.flux_period == "Mois", BROWN, TEXT_MUTED), border=f"1px solid {rx.cond(DashboardState.flux_period == 'Mois', GOLD, BORDER_CARD)}", border_radius="6px", padding="4px 12px", font_size="12px", height="32px", cursor="pointer"),
                    rx.button("Trimestre", on_click=DashboardState.set_flux_trimestre, bg=rx.cond(DashboardState.flux_period == "Trimestre", "#F0E6D0", WHITE), color=rx.cond(DashboardState.flux_period == "Trimestre", BROWN, TEXT_MUTED), border=f"1px solid {rx.cond(DashboardState.flux_period == 'Trimestre', GOLD, BORDER_CARD)}", border_radius="6px", padding="4px 12px", font_size="12px", height="32px", cursor="pointer"),
                    spacing="1",
                ),
                width="100%",
                align_items="center",
                padding="16px 20px 0 20px",
            ),
            # CORRECTION : width="100%" et height explicite pour le conteneur SVG
            rx.box(
                rx.html(DashboardState.flux_svg_html),
                width="100%",
                height="260px",
            ),
            width="100%",
            spacing="0",
        ),
        bg=WHITE,
        border_radius="12px",
        border=f"1px solid {BORDER_CARD}",
        box_shadow=SHADOW,
        width="100%",
        height="320px",
    )


def volume_donut() -> rx.Component:
    return rx.box(rx.vstack(
        rx.hstack(rx.icon("pie-chart", size=16, color=GOLD), rx.text("Volume par bureau", color=TEXT_MAIN, font_size="14px", font_weight="600"), spacing="2", align_items="center", padding="16px 20px 0 20px"),
        rx.html(DashboardState.donut_html), width="100%", spacing="0",
    ), bg=WHITE, border_radius="12px", border=f"1px solid {BORDER_CARD}", box_shadow=SHADOW, width="100%", height="320px")


def productivity_table() -> rx.Component:
    def agent_row(row: dict) -> rx.Component:
        return rx.hstack(
            rx.center(rx.text(row["rang"].to_string(), color=TEXT_MUTED, font_size="12px", font_weight="600"), width="32px", height="32px", border_radius="16px", bg=rx.cond(row["rang"] == 1, CREAM_CARD, "transparent")),
            rx.vstack(rx.text(row["agent"], color=TEXT_MAIN, font_size="13px", font_weight="600"), rx.text(row["csf"], color=TEXT_MUTED, font_size="11px"), spacing="0", align_items="start", width="160px"),
            rx.box(bureau_badge(row["bureau"]), width="110px"),
            rx.hstack(rx.box(rx.box(width=f"{row['pct']}%", height="6px", bg=GOLD, border_radius="3px"), width="120px", bg="#F0F0F0", border_radius="3px", height="6px"), rx.text(row["dossiers"].to_string(), color=TEXT_MAIN, font_size="13px", font_weight="600", width="50px", text_align="right"), spacing="2", align_items="center", flex="1"),
            rx.text(row["moyenne"].to_string(), color=TEXT_MAIN, font_size="13px", font_weight="500", width="60px", text_align="right"),
            width="100%", align_items="center", padding="10px 0", border_bottom=f"1px solid {BORDER_CARD}",
        )
    return rx.box(rx.vstack(
        rx.hstack(rx.hstack(rx.icon("users", size=16, color=GOLD), rx.text("Productivité équipe", color=TEXT_MAIN, font_size="14px", font_weight="600"), spacing="2", align_items="center"), rx.spacer(), rx.text(f"{DashboardState.productivite_count} agents", color=TEXT_MUTED, font_size="12px"), width="100%", align_items="center", padding="16px 20px"),
        rx.hstack(rx.text("AGENT", color=TEXT_MUTED, font_size="10px", font_weight="700", letter_spacing="0.5px", width="220px"), rx.text("BUREAU", color=TEXT_MUTED, font_size="10px", font_weight="700", letter_spacing="0.5px", width="110px"), rx.text("DOSSIERS TRAITÉS", color=TEXT_MUTED, font_size="10px", font_weight="700", letter_spacing="0.5px", flex="1"), rx.text("MOY./JOUR", color=TEXT_MUTED, font_size="10px", font_weight="700", letter_spacing="0.5px", width="70px", text_align="right"), width="100%", padding="0 20px 8px 20px", border_bottom=f"2px solid {BORDER_CARD}"),
        rx.vstack(rx.foreach(DashboardState.productivite, agent_row), width="100%", padding="0 20px", spacing="0"),
        width="100%", spacing="0",
    ), bg=WHITE, border_radius="12px", border=f"1px solid {BORDER_CARD}", box_shadow=SHADOW, width="100%")


def recent_records() -> rx.Component:
    def record_item(item: dict) -> rx.Component:
        return rx.hstack(
            rx.vstack(rx.text(item["code"], color=TEXT_MAIN, font_size="13px", font_weight="600"), rx.text(item["desc"], color=TEXT_MUTED, font_size="12px", max_width="280px"), spacing="1", align_items="start"),
            rx.spacer(),
            type_badge(item["type_badge"]),
            width="100%", align_items="center", padding="12px 0", border_bottom=f"1px solid {BORDER_CARD}",
        )
    return rx.box(rx.vstack(
        rx.hstack(rx.icon("file-text", size=16, color=GOLD), rx.text("Derniers enregistrements", color=TEXT_MAIN, font_size="14px", font_weight="600"), spacing="2", align_items="center", padding="16px 20px"),
        rx.vstack(rx.foreach(DashboardState.derniers_enregistrements, record_item), width="100%", padding="0 20px", spacing="0"),
        width="100%", spacing="0",
    ), bg=WHITE, border_radius="12px", border=f"1px solid {BORDER_CARD}", box_shadow=SHADOW, width="100%")


def loading_spinner() -> rx.Component:
    return rx.center(rx.vstack(rx.spinner(size="3", color=GOLD, thickness="4px"), rx.text("Chargement des données...", color=TEXT_MUTED, font_size="14px", margin_top="16px")), width="100%", height="100vh", bg=CREAM)


def error_banner() -> rx.Component:
    return rx.cond(DashboardState.db_error != "", rx.box(rx.hstack(rx.icon("triangle_alert", size=16, color=RED_TEXT), rx.text(DashboardState.db_error, color=RED_TEXT, font_size="13px"), spacing="2", align_items="center"), padding="12px 24px", bg=RED_LIGHT, border_radius="8px", margin="16px 32px", width="calc(100% - 64px)"), rx.box())


def index() -> rx.Component:
    return rx.box(
        rx.cond(DashboardState.is_loading & (DashboardState.kpi_total_dossiers == 0), loading_spinner(),
            rx.vstack(header(), filter_bar(), error_banner(), rx.box(rx.vstack(
                objectives_section(), metrics_section(),
                rx.hstack(rx.box(activity_chart(), width="65%"), rx.box(volume_donut(), width="35%"), spacing="4", width="100%", align_items="stretch"),
                rx.hstack(rx.box(productivity_table(), width="60%"), rx.box(recent_records(), width="40%"), spacing="4", width="100%", align_items="stretch"),
                spacing="6", width="100%", max_width="1400px", margin="0 auto", padding="24px 0",
            ), width="100%", bg=CREAM, min_height="100vh"), spacing="0", width="100%"),
        ),
        width="100%", min_height="100vh", bg=CREAM, on_mount=DashboardState.on_load,
    )


app = rx.App(theme=rx.theme(appearance="light", accent_color="orange", radius="medium"), stylesheets=[])
app.add_page(index, route="/", title="Dashboard de Suivi — DGID")