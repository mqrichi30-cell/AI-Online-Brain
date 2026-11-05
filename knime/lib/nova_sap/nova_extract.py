"""Extracción de la lista de órdenes (POs) de NOVA en Retail Link.

Retail Link usa SSO de Walmart + MFA, así que se usa un perfil de navegador
PERSISTENTE: una persona inicia sesión una vez (modo --setup, con ventana visible)
y las corridas automáticas reutilizan esa sesión hasta que caduque.

Dos modos (config["nova"]["mode"]):
  - "export": da clic en el botón de exportar de NOVA y lee el Excel/CSV descargado
              (lo más robusto si NOVA tiene exportación).
  - "table":  lee la tabla HTML de la pantalla, recorriendo la paginación.

IMPORTANTE: los selectores CSS dependen de cómo se ve NOVA en su cuenta. Vienen
como ejemplo en config.example.json y hay que ajustarlos con el inspector del
navegador (o `playwright codegen https://retaillink2.wal-mart.com/novarl/`).

Uso manual:
  python -m nova_sap.nova_extract --config config.json --setup   # login inicial
  python -m nova_sap.nova_extract --config config.json           # prueba de extracción
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import pandas as pd

LOGIN_HINTS = ("login", "signin", "sso", "auth", "pingone", "okta")


class NovaLoginRequired(RuntimeError):
    """La sesión caducó: hay que correr --setup e iniciar sesión manualmente."""


def _browser_opts(cfg: dict, launch: dict) -> None:
    """browser_channel: "msedge"/"chrome" (navegador instalado en Windows) o
    browser_executable: ruta a un ejecutable de Chromium."""
    if cfg.get("browser_executable"):
        launch["executable_path"] = cfg["browser_executable"]
    elif cfg.get("browser_channel"):
        launch["channel"] = cfg["browser_channel"]


def _is_login_page(page) -> bool:
    url = page.url.lower()
    return any(h in url for h in LOGIN_HINTS) and "novarl" not in url


def _wait_ready(page, cfg: dict) -> None:
    page.wait_for_load_state("networkidle", timeout=cfg.get("timeout_ms", 90_000))
    if _is_login_page(page):
        raise NovaLoginRequired(
            "NOVA redirigió al login. Ejecuta el modo --setup en la máquina del robot e inicia sesión.")
    ready = cfg["selectors"].get("ready")
    if ready:
        page.wait_for_selector(ready, timeout=cfg.get("timeout_ms", 90_000))


def _apply_filters(page, cfg: dict) -> None:
    """Pasos previos opcionales (elegir filtros, rango de fechas, botón buscar...).
    config: "pre_steps": [{"action": "click"|"fill"|"select"|"wait", "selector": ..., "value": ...}]"""
    for step in cfg.get("pre_steps", []):
        action, sel, val = step["action"], step.get("selector"), step.get("value")
        if action.startswith("_"):  # paso desactivado / ejemplo
            continue
        if action not in ("click", "fill", "select", "wait"):
            raise ValueError(f"Acción desconocida en nova.pre_steps: {action}")
        if action == "click":
            page.click(sel)
        elif action == "fill":
            page.fill(sel, str(val))
        elif action == "select":
            page.select_option(sel, str(val))
        elif action == "wait":
            if sel:
                page.wait_for_selector(sel)
            else:
                page.wait_for_timeout(int(val or 1000))
        page.wait_for_load_state("networkidle")


def _read_export(page, cfg: dict, download_dir: Path) -> pd.DataFrame:
    with page.expect_download(timeout=cfg.get("timeout_ms", 90_000)) as dl_info:
        page.click(cfg["selectors"]["export_button"])
    dl = dl_info.value
    target = download_dir / dl.suggested_filename
    dl.save_as(str(target))
    if target.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(target, dtype=str)
    return pd.read_csv(target, dtype=str, sep=None, engine="python")


def _read_table(page, cfg: dict) -> pd.DataFrame:
    sel = cfg["selectors"]
    rows: list[list[str]] = []
    headers: list[str] | None = None
    for _ in range(cfg.get("max_pages", 200)):
        table = page.locator(sel["table"]).first
        if headers is None:
            headers = [h.strip() for h in table.locator(sel.get("header_cells", "thead th")).all_inner_texts()]
        for tr in table.locator(sel.get("body_rows", "tbody tr")).all():
            rows.append([c.strip() for c in tr.locator("td").all_inner_texts()])
        nxt = sel.get("next_page")
        if not nxt:
            break
        btn = page.locator(nxt).first
        if btn.count() == 0 or not btn.is_enabled() or (btn.get_attribute("aria-disabled") == "true"):
            break
        btn.click()
        page.wait_for_load_state("networkidle")
    width = max((len(r) for r in rows), default=len(headers or []))
    headers = (headers or []) + [f"col_{i}" for i in range(len(headers or []), width)]
    return pd.DataFrame([r + [""] * (width - len(r)) for r in rows], columns=headers[:width] or None)


def extract_nova_orders(config: dict, headless: bool = True) -> pd.DataFrame:
    """Devuelve un DataFrame con todas las columnas de NOVA + columna normalizada 'po'."""
    from playwright.sync_api import sync_playwright

    cfg = config["nova"]
    profile = Path(cfg["browser_profile_dir"]).expanduser()
    profile.mkdir(parents=True, exist_ok=True)
    download_dir = Path(cfg.get("download_dir", profile / "downloads")).expanduser()
    download_dir.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        launch = dict(headless=headless, accept_downloads=True)
        _browser_opts(cfg, launch)
        ctx = p.chromium.launch_persistent_context(str(profile), **launch)
        try:
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            page.goto(cfg["url"], timeout=cfg.get("timeout_ms", 90_000))
            _wait_ready(page, cfg)
            _apply_filters(page, cfg)
            df = _read_export(page, cfg, download_dir) if cfg.get("mode", "export") == "export" else _read_table(page, cfg)
        finally:
            ctx.close()

    po_col = cfg["po_column"]
    if po_col not in df.columns:
        raise KeyError(f"No existe la columna '{po_col}' en NOVA. Columnas: {list(df.columns)}")
    df = df.rename(columns={po_col: "po"})
    return df


def interactive_setup(config: dict) -> None:
    """Abre el navegador visible con el perfil persistente para iniciar sesión (SSO+MFA)."""
    from playwright.sync_api import sync_playwright

    cfg = config["nova"]
    profile = Path(cfg["browser_profile_dir"]).expanduser()
    profile.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        launch = dict(headless=False)
        _browser_opts(cfg, launch)
        ctx = p.chromium.launch_persistent_context(str(profile), **launch)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto(cfg["url"])
        input("Inicia sesión en Retail Link / NOVA en la ventana y presiona ENTER aquí cuando veas la lista de órdenes...")
        ctx.close()
    print("Sesión guardada en", profile)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--config", required=True)
    ap.add_argument("--setup", action="store_true", help="Login manual para guardar la sesión")
    ap.add_argument("--headed", action="store_true", help="Ver el navegador durante la extracción")
    ap.add_argument("--out", help="Guardar resultado en CSV")
    args = ap.parse_args(argv)
    config = json.loads(Path(args.config).read_text(encoding="utf-8"))
    if args.setup:
        interactive_setup(config)
        return 0
    df = extract_nova_orders(config, headless=not args.headed)
    print(f"{len(df)} filas, {df['po'].nunique()} POs únicas")
    if args.out:
        df.to_csv(args.out, index=False)
    return 0


if __name__ == "__main__":
    sys.exit(main())
