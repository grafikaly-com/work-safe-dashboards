"""Extract the semantic-model schema of each .pbix into JSON.

Reads tables, columns, measures, relationships and Power Query (M) sources with
pbixray, plus the fields each report page uses. Row data is never read, and M
code is redacted before it is written (URL query strings, key-like values and
the embedded data of "Enter Data" tables).

Usage:
    python scripts/extract_pbix.py                    # all pbix/*.pbix
    python scripts/extract_pbix.py "pbix/Survey Dash.pbix"
"""

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path

from pbixray import PBIXRay

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "docs" / "models" / "_extracted"

# Power BI's hidden auto date/time tables; they add noise, not schema.
AUTO_TABLE = re.compile(r"^(LocalDateTable|DateTableTemplate)_")

# TOM enums (Microsoft.AnalysisServices.Tabular)
TOM_DATA_TYPES = {2: "String", 6: "Int64", 8: "Double", 9: "DateTime", 10: "Decimal",
                  11: "Boolean", 17: "Binary", 20: "Variant"}
TOM_COLUMN_TYPES = {1: "Data", 2: "Calculated", 3: "RowNumber", 4: "CalculatedTableColumn"}
PANDAS_DATA_TYPES = {"string": "String", "object": "String", "Int64": "Int64", "int64": "Int64",
                     "float64": "Double", "Float64": "Double", "datetime64[ns]": "DateTime",
                     "bool": "Boolean", "boolean": "Boolean"}

GUID = r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"


# --- redaction ---------------------------------------------------------------

def redact_m(expr):
    """Strip secrets and embedded data from M code before it leaves the model."""
    if not expr:
        return expr
    # "Enter Data" tables keep their rows as a compressed base64 blob.
    expr = re.sub(r'Binary\.FromText\("[^"]*"', 'Binary.FromText("<embedded data redacted>"', expr)
    expr = re.sub(r'"[A-Za-z0-9+/=]{80,}"', '"<blob redacted>"', expr)
    # URL query strings (API keys are usually passed here).
    expr = re.sub(r'(https?://[^"?\s]+)\?[^"\s]*', r"\1?<query redacted>", expr)
    # Key-like record fields: Headers=[Authorization="..."], ApiKey="...", #"x-api-key"="..."
    secret = r"(?:key|token|secret|password|pwd|authorization|bearer|sig)"
    expr = re.sub(rf'(?i)(#"[^"]*{secret}[^"]*"|\b[\w-]*{secret}[\w-]*)(\s*=\s*)"[^"]*"',
                  r'\1\2"<redacted>"', expr)
    return expr


# --- Power Query source parsing ----------------------------------------------

M_KEYWORDS = {"let", "in", "each", "if", "then", "else", "and", "or", "not", "null", "true", "false",
              "type", "meta", "as", "is", "try", "otherwise", "error", "section", "shared",
              "any", "anynonnull", "binary", "date", "datetime", "datetimezone", "duration", "function",
              "list", "logical", "none", "number", "record", "table", "text", "time", "nullable"}
M_TOKEN = re.compile(r'#"(?:[^"]|"")*"|"(?:[^"]|"")*"|//[^\n]*|/\*.*?\*/', re.S)


def _blank_strings(expr):
    """Blank string literals and comments (same length), keeping #"quoted identifiers"."""
    return M_TOKEN.sub(lambda m: m.group(0) if m.group(0).startswith('#"') else " " * len(m.group(0)), expr)


def query_references(expr, table_names, self_name):
    """Other queries this M expression reads: loaded tables, plus staging queries
    that exist only in Power Query (not loaded to the model)."""
    code = _blank_strings(expr)
    steps = {s[2:-1] if s.startswith('#"') else s
             for s in re.findall(r'^\s*(#"[^"]+"|[A-Za-z_]\w*)\s*=(?!>)', code, re.M)}
    params = {p.strip().split(" ")[0] for group in re.findall(r"\(([^()]*)\)\s*(?:as\s+\w+\s*)?=>", code)
              for p in group.split(",") if p.strip()}
    refs = set()
    for t in table_names:
        if t == self_name:
            continue
        if re.search(rf'(?<!\[)#"{re.escape(t)}"', code) or (
                re.match(r"^[A-Za-z_]\w*$", t)
                and re.search(rf'(?<![\w."#\[]){re.escape(t)}(?![\w".\]])', code)):
            refs.add(t)
    # Quoted names that are read (not field access, not a step/record definition).
    hidden = set(re.findall(r'(?<!\[)#"([^"]+)"(?!\s*=)', code))
    # Bare names used as a whole step value (Source = Foo,) or inside Table.Combine({...}).
    hidden |= set(re.findall(r'=\s*([A-Za-z_]\w*)\s*(?:,|\n\s*in\b)', code))
    for group in re.findall(r"Table\.Combine\(\s*\{([^}]*)\}", code):
        hidden |= {g.strip() for g in group.split(",") if re.match(r"^[A-Za-z_]\w*$", g.strip())}
    hidden = {h for h in hidden if h not in steps and h not in params and not h.startswith("_")
              and h.lower() not in M_KEYWORDS}
    hidden -= refs | {self_name}
    return sorted(refs), sorted(hidden)


def parse_source(expr, table_names, self_name):
    """Classify where a table's data comes from and which other queries it reads."""
    refs, staging = query_references(expr, table_names, self_name)
    src = {"references": refs}
    if staging:
        src["staging_queries"] = staging
    sp = re.search(r'SharePoint\.Tables\(\s*"([^"]+)"', expr)
    files = re.search(r'SharePoint\.(?:Files|Contents)\(\s*"([^"]+)"', expr)
    refs = refs + staging
    if sp:
        src["kind"] = "SharePoint list"
        src["site"] = sp.group(1).rstrip("/")
        src["list_ids"] = sorted(set(re.findall(rf'Source\{{\[Id="({GUID})"\]\}}', expr)))
        titles = re.findall(r'Source\{\[Title="([^"]+)"\]\}', expr)
        if titles:
            src["list_titles"] = sorted(set(titles))
    elif files:
        src["kind"] = "Excel file on SharePoint" if "Excel.Workbook" in expr else "File on SharePoint"
        src["site"] = files.group(1).rstrip("/")
        src["files"] = sorted(set(re.findall(r'"([^"/\\]+\.(?:xlsx|xlsm|xls|csv))"', expr)))
        folder = re.search(r'\[Folder Path\],\s*"([^"]+)"', expr)
        if folder:
            src["folder"] = folder.group(1)
    elif "Web.Contents" in expr:
        url = re.search(r'Web\.Contents\(\s*"([^"?]+)', expr)
        src["kind"] = "Web API"
        src["url"] = url.group(1) if url else None
    elif "Binary.Decompress" in expr and "Table.FromRows" in expr:
        src["kind"] = "Manual table (Enter Data)"
    elif "Table.Combine" in expr and refs:
        src["kind"] = "Append of other queries"
    elif refs:
        src["kind"] = "Derived from other queries"
    elif re.search(r"DateTime(Zone)?\.(Fixed)?(Local|Utc)?Now", expr):
        src["kind"] = "Power Query calculation (refresh timestamp)"
    else:
        src["kind"] = "Power Query calculation"
    if sp and "Table.Combine" in expr and refs:
        src["kind"] = "SharePoint list + append"
    return src


# --- report (visual) field usage --------------------------------------------

def _read_text(z, name):
    raw = z.read(name)
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff") or raw[1:2] == b"\x00":
        return raw.decode("utf-16")
    return raw.decode("utf-8-sig")


def _walk(node, aliases, out):
    if isinstance(node, str):
        s = node.strip()
        if s[:1] in "{[" and len(s) > 1:  # legacy Layout nests JSON in strings
            try:
                _walk(json.loads(s), aliases, out)
            except ValueError:
                pass
        return
    if isinstance(node, list):
        for item in node:
            _walk(item, aliases, out)
        return
    if not isinstance(node, dict):
        return
    for item in node.get("From") or []:
        if isinstance(item, dict) and "Name" in item and "Entity" in item:
            aliases[item["Name"]] = item["Entity"]
    for kind in ("Column", "Measure", "PropertyVariationSource"):
        ref = node.get(kind)
        if isinstance(ref, dict) and "Property" in ref:
            sref = (ref.get("Expression") or {}).get("SourceRef") or {}
            entity = sref.get("Entity") or aliases.get(sref.get("Source"))
            if entity:
                out.add((entity, ref["Property"], "measure" if kind == "Measure" else "column"))
    for value in node.values():
        _walk(value, aliases, out)


def report_usage(pbix_path):
    """Pages, connection ids and the (table, field, kind) set used by the report."""
    with zipfile.ZipFile(pbix_path) as z:
        names = z.namelist()
        used, pages = set(), []
        if "Report/definition/report.json" in names:  # PBIR
            fmt = "PBIR"
            for n in names:
                if n.startswith("Report/definition/") and n.endswith(".json"):
                    doc = json.loads(_read_text(z, n))
                    _walk(doc, {}, used)
                    if n.endswith("/page.json"):
                        pages.append({"name": doc.get("displayName"),
                                      "hidden": doc.get("visibility") == "HiddenInViewMode",
                                      "visuals": sum(1 for m in names if m.startswith(n[:-9] + "visuals/")
                                                     and m.endswith("/visual.json"))})
        else:  # legacy single-file Layout
            fmt = "Legacy Layout"
            layout = json.loads(_read_text(z, "Report/Layout"))
            _walk(layout.get("config"), {}, used)
            _walk(layout.get("filters"), {}, used)
            for sec in layout.get("sections", []):
                containers = sec.get("visualContainers", [])
                for vc in containers:
                    _walk(vc, {}, used)
                _walk(sec.get("filters"), {}, used)
                pages.append({"name": sec.get("displayName"), "hidden": sec.get("visibility") == 1,
                              "visuals": len(containers)})
        conn = json.loads(_read_text(z, "Connections")) if "Connections" in names else {}
    remote = (conn.get("RemoteArtifacts") or [{}])[0]
    return {
        "format": fmt,
        "pages": pages,
        "dataset_id": remote.get("DatasetId"),
        "report_id": remote.get("ReportId"),
        "workspace_id": conn.get("OriginalWorkspaceObjectId"),
        "used_fields": sorted([list(u) for u in used]),
    }


# --- model extraction ---------------------------------------------------------

def _str(v):
    return None if v is None or (isinstance(v, float) and v != v) or v == "" else str(v)


def extract(pbix_path):
    model = PBIXRay(str(pbix_path))
    usage = report_usage(pbix_path)
    used = {(t, f) for t, f, _ in usage["used_fields"]}

    table_names = [t for t in model.tables if not AUTO_TABLE.match(t)]
    meta = dict(zip(model.metadata.Name, model.metadata.Value))
    order = json.loads(meta.get("PBI_QueryOrder", "[]"))
    table_names.sort(key=lambda t: (order.index(t) if t in order else len(order), t))

    m_expr = dict(zip(model.power_query.TableName, model.power_query.Expression))
    dax_tables = dict(zip(model.dax_tables.TableName, model.dax_tables.Expression))
    dax_cols = {(r.TableName, r.ColumnName): r.Expression for r in model.dax_columns.itertuples()}
    tm_tables = {r.Name: r for r in model.tmschema_tables.itertuples()}
    tm_cols = {(r.TableName, r.Name): r for r in model.tmschema_columns.itertuples()
               if r.TableName in table_names}

    tables = []
    for t in table_names:
        tm = tm_tables.get(t)
        if t in dax_tables:
            src = {"kind": "DAX calculated table", "references": []}
            expr = dax_tables[t]
        else:
            expr = m_expr.get(t, "")
            src = parse_source(expr, table_names, t) if expr else {"kind": "Unknown", "references": []}
        columns = []
        for r in model.schema[model.schema.TableName == t].itertuples():
            c = tm_cols.get((t, r.ColumnName))
            tom_type = TOM_DATA_TYPES.get(int(c.DataType)) if c is not None else None
            calc = _str(dax_cols.get((t, r.ColumnName)))
            columns.append({
                "name": r.ColumnName,
                "data_type": tom_type or PANDAS_DATA_TYPES.get(str(r.PandasDataType), str(r.PandasDataType)),
                "column_type": "Calculated" if calc else (
                    TOM_COLUMN_TYPES.get(int(c.Type), "Data") if c is not None else
                    ("CalculatedTableColumn" if t in dax_tables else "Data")),
                "hidden": bool(c.IsHidden) if c is not None else False,
                "format": _str(c.FormatString) if c is not None else None,
                "description": _str(c.Description) if c is not None else None,
                "source_column": _str(c.SourceColumn) if c is not None else None,
                "expression": calc,
                "used_in_report": (t, r.ColumnName) in used,
            })
        tables.append({
            "name": t,
            "hidden": bool(tm.IsHidden) if tm is not None else False,
            "description": _str(tm.Description) if tm is not None else None,
            "source": src,
            "expression": redact_m(expr) if t not in dax_tables else expr,
            "expression_language": "DAX" if t in dax_tables else "M",
            "columns": columns,
            "used_in_report": any(u[0] == t for u in used),
        })

    measures = [{
        "table": r.TableName, "name": r.Name, "expression": _str(r.Expression),
        "display_folder": _str(r.DisplayFolder), "description": _str(r.Description),
        "used_in_report": (r.TableName, r.Name) in used,
    } for r in model.dax_measures.itertuples()]

    rel = model.relationships
    relationships = [{
        "from_table": r.FromTableName, "from_column": r.FromColumnName,
        "to_table": r.ToTableName, "to_column": r.ToColumnName,
        "cardinality": r.Cardinality, "cross_filter": r.CrossFilteringBehavior,
        "active": bool(r.IsActive),
    } for r in rel.itertuples()
        if not AUTO_TABLE.match(r.FromTableName) and not AUTO_TABLE.match(r.ToTableName)]

    known = set(table_names)
    measure_keys = {(m["table"], m["name"]) for m in measures}
    column_keys = {(t["name"], c["name"]) for t in tables for c in t["columns"]}
    unresolved = sorted({f"{t}[{f}]" for t, f, _ in usage["used_fields"]
                         if t not in known or ((t, f) not in measure_keys and (t, f) not in column_keys)})

    return {
        "report": pbix_path.stem,
        "file": pbix_path.relative_to(ROOT).as_posix(),
        "status": "ok",
        "auto_date_tables": sum(1 for t in model.tables if AUTO_TABLE.match(t)),
        "report_meta": {k: v for k, v in usage.items() if k != "used_fields"},
        "tables": tables,
        "measures": measures,
        "relationships": relationships,
        "used_fields": usage["used_fields"],
        "unresolved_report_fields": unresolved,
    }


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="*", type=Path, help="pbix files (default: pbix/*.pbix)")
    args = ap.parse_args()
    files = [f.resolve() for f in args.files] or sorted((ROOT / "pbix").glob("*.pbix"))
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    failed = 0
    for f in files:
        try:
            data = extract(f)
        except Exception as e:  # keep going; record the failure for gaps.md
            failed += 1
            data = {"report": f.stem, "file": f.name, "status": "failed", "error": f"{type(e).__name__}: {e}"}
            print(f"FAIL {f.stem}: {data['error']}")
        else:
            ncols = sum(len(t["columns"]) for t in data["tables"])
            print(f"OK   {f.stem:32} tables={len(data['tables']):3} columns={ncols:4} "
                  f"measures={len(data['measures']):3} relationships={len(data['relationships']):3} "
                  f"unresolved={len(data['unresolved_report_fields'])}")
        (OUT_DIR / f"{f.stem}.json").write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
