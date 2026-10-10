"""Static integrity checks for the Excel add-in; no workbook or tenant access required."""
from pathlib import Path
import sys
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
MANIFEST = ROOT / "manifest.xml"
NS = {
    "a": "http://schemas.microsoft.com/office/appforoffice/1.1",
    "bt": "http://schemas.microsoft.com/office/officeappbasictypes/1.0",
    "ov": "http://schemas.microsoft.com/office/taskpaneappversionoverrides",
}

def fail(message: str) -> None:
    print(f"FAIL: {message}", file=sys.stderr)
    raise SystemExit(1)

if not MANIFEST.is_file():
    fail("manifest.xml is missing")

try:
    tree = ET.parse(MANIFEST)
except ET.ParseError as exc:
    fail(f"manifest.xml is not well-formed XML: {exc}")

root = tree.getroot()
if not root.tag.endswith("OfficeApp"):
    fail("manifest root is not OfficeApp")

manifest_text = MANIFEST.read_text(encoding="utf-8")
required_tokens = [
    "<Permissions>ReadWriteDocument</Permissions>",
    'Name="ExcelApi" MinVersion="1.4"',
    'Name="IdentityAPI" MinVersion="1.3"',
]
for token in required_tokens:
    if token not in manifest_text:
        fail(f"required Office capability is missing: {token}")

# All manifest URL paths that are part of this repository must map to checked-in files.
url_nodes = []
for element in root.iter():
    for attr in ("DefaultValue",):
        value = element.attrib.get(attr, "")
        if value.startswith("https://addin-registros-envios.vercel.app/"):
            url_nodes.append(value)

missing = []
for url in sorted(set(url_nodes)):
    path = url.split("addin-registros-envios.vercel.app/", 1)[1].split("?", 1)[0].split("#", 1)[0]
    if path and not (ROOT / path).is_file():
        missing.append((url, path))

if missing:
    for url, path in missing:
        print(f"Missing local asset for manifest URL: {url} -> {path}", file=sys.stderr)
    raise SystemExit(1)

# Critical navigation documents are expected in the current Office add-in.
for relative in ("taskpane.html", "commands.html", "index.html",
                 "email.html", "sms.html", "historico.html",
                 "resumo.html", "config.html"):
    if not (ROOT / relative).is_file():
        fail(f"expected page is missing: {relative}")

print("PASS: manifest XML is well formed.")
print("PASS: Office permission/capability declarations are present.")
print(f"PASS: {len(set(url_nodes))} unique manifest asset URLs resolve to repository files.")
print("PASS: core add-in pages are present.")
print("NOTE: these checks do not prove Entra consent, Office runtime behavior, workbook compatibility, Graph access, or successful writes.")
