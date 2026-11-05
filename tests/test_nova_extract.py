"""Prueba el extractor contra una página local que imita NOVA (tabla paginada + exportar).
Se salta si no hay Playwright/Chromium disponible."""
import os
from pathlib import Path

import pytest

pytest.importorskip("playwright")
CHROMIUM = os.environ.get("PW_CHROMIUM") or next(
    (str(p) for p in Path("/opt/pw-browsers").glob("chromium-*/chrome-linux*/chrome")), None)
if not CHROMIUM:
    pytest.skip("Chromium no disponible", allow_module_level=True)

from nova_sap.nova_extract import NovaLoginRequired, extract_nova_orders  # noqa: E402

PAGE = """<!doctype html><html><body>
<h1>Purchase Orders</h1>
<button id="export">Export</button>
<table id="t"><thead><tr><th>PO Number</th><th>DC</th></tr></thead><tbody></tbody></table>
<button id="next" aria-label="Next page">Next</button>
<script>
const pages = [[["1001","6006"],["1002","6006"]], [["1002","7038"],["1003","7038"]]];
let i = 0;
function render(){
  document.querySelector('#t tbody').innerHTML =
    pages[i].map(r => '<tr><td>'+r[0]+'</td><td>'+r[1]+'</td></tr>').join('');
  document.getElementById('next').disabled = i >= pages.length - 1;
}
document.getElementById('next').onclick = () => { i++; render(); };
document.getElementById('export').onclick = () => {
  const csv = 'PO Number,DC\\n' + pages.flat().map(r => r.join(',')).join('\\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csv], {type: 'text/csv'}));
  a.download = 'nova_orders.csv'; a.click();
};
render();
</script></body></html>"""


def cfg(tmp_path, mode, url=None):
    page = tmp_path / "nova.html"
    page.write_text(PAGE)
    return {"nova": {
        "url": url or page.as_uri(), "mode": mode, "browser_executable": CHROMIUM,
        "browser_profile_dir": str(tmp_path / "profile"), "download_dir": str(tmp_path / "dl"),
        "timeout_ms": 15000, "po_column": "PO Number",
        "pre_steps": [{"action": "wait", "selector": "text=Purchase Orders"}, {"action": "_ejemplo"}],
        "selectors": {"ready": "#t tbody tr", "export_button": "#export", "table": "#t",
                      "next_page": "#next"},
    }}


def test_table_mode_with_pagination(tmp_path):
    df = extract_nova_orders(cfg(tmp_path, "table"))
    assert df["po"].tolist() == ["1001", "1002", "1002", "1003"]
    assert list(df.columns) == ["po", "DC"]


def test_export_mode(tmp_path):
    df = extract_nova_orders(cfg(tmp_path, "export"))
    assert sorted(df["po"].unique()) == ["1001", "1002", "1003"] and len(df) == 4


def test_login_redirect_detected(tmp_path):
    login = tmp_path / "login.html"
    login.write_text("<html><body>Sign in</body></html>")
    with pytest.raises(NovaLoginRequired):
        extract_nova_orders(cfg(tmp_path, "table", url=login.as_uri()))
