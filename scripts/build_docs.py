"""Build schema docs from the extracted model JSON and the xlsx registers.

Inputs:  docs/models/_extracted/*.json (written by extract_pbix.py)
         xlsx/Dashboards Master File.xlsx
         xlsx/Power Automate Flow Master Register.xlsx
Outputs: docs/README.md, docs/models/<Report>.md, docs/lineage.md, docs/gaps.md,
         docs/Model Schema Register v0.xlsx (the generated register; docs/Model Schema Register.xlsx is
         maintained by hand and never written here)

The workbooks are read as cell values only; hyperlinks in them are never followed.

Usage:
    python scripts/build_docs.py
"""

import difflib
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path
from urllib.parse import quote

import openpyxl
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font

from extract_pbix import redact_m

ROOT = Path(__file__).resolve().parents[1]
DOCS = ROOT / "docs"
EXTRACTED = DOCS / "models" / "_extracted"
MASTER = ROOT / "xlsx" / "Dashboards Master File.xlsx"
FLOW_REGISTER = ROOT / "xlsx" / "Power Automate Flow Master Register.xlsx"

# Master File sheet name -> pbix file stem.
SHEET_TO_REPORT = {
    "PEPA Exec Dash": "PEPA Exec Dash",
    "Permit SPE Dash": "Permit SPE Dash",
    "Survey Dash": "Survey Dash",
    "ELT Permit Timeliness Dashboard": "ELT Permit Timeliness Report",
    "Permit Forecast": "Permit Forecast",
    "PEPA Management Dash": "PEPA Management Dash",
    "ELT FrontPage": "ELT Frontpage",
    "Compliance Monitoring Dashboard": "Compliance Monitoring Dash",
    "Programme Risk": "Programme Risk",
}

# Master File column headers vary by sheet.
MASTER_COLUMNS = {"tables": "table", "list": "is_list", "location": "location", "team": "team",
                  "maintained by": "maintained_by", "contact info": "contact",
                  "associated flow?": "flow", "flow": "flow", "app?": "app", "app": "app",
                  "notes": "notes", "comments - mohsin": "comments"}

# Words in the master "flow" column that describe a flow without naming it.
GENERIC_FLOW_WORDS = {"yes", "no", "none", "na", "scheduled", "solution", "inthesolution"}

CARDINALITY = {"M:1": "}o--||", "1:M": "||--o{", "1:1": "||--||", "M:M": "}o--o{"}


# --- small helpers -----------------------------------------------------------

def clean(v):
    if v is None:
        return None
    s = str(v).replace("\xa0", " ").replace("\ufffd", "'").strip()
    return s or None


def tkey(name):
    """Table-name key: trimmed, whitespace-collapsed, case-folded."""
    return " ".join(str(name).split()).casefold()


def norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def slug(heading):
    """GitHub-style heading anchor."""
    s = re.sub(r"[^\w\- ]", "", heading.strip().lower())
    return s.replace(" ", "-")


def cell(s):
    """Escape text for a Markdown table cell."""
    if s is None or s == "":
        return ""
    return str(s).replace("|", "\\|").replace("\r", "").replace("\n", "<br>")


def code(s):
    return f"`{s}`" if s else ""


def model_link(report, anchor=None, prefix="models/"):
    return f"{prefix}{quote(report)}.md" + (f"#{anchor}" if anchor else "")


def flow_link(name, prefix=""):
    return f"[{name}]({prefix}lineage.md#{slug(name)})"


def site_name(url):
    return url.rstrip("/").rsplit("/", 1)[-1] if url else None


def redact_text(s):
    return redact_m(s) if s else s


# --- loaders -----------------------------------------------------------------

def load_models():
    models = []
    for f in sorted(EXTRACTED.glob("*.json")):
        models.append(json.loads(f.read_text(encoding="utf-8")))
    if not models:
        sys.exit(f"No extracted models in {EXTRACTED}; run scripts/extract_pbix.py first.")
    return models


def load_master():
    wb = openpyxl.load_workbook(MASTER, read_only=True, data_only=True, keep_links=False)
    out = {}
    for ws in wb.worksheets:
        report = SHEET_TO_REPORT.get(ws.title)
        if not report:
            print(f"warning: Master File sheet {ws.title!r} has no matching pbix; skipped")
            continue
        info, rows, header = {}, [], None
        for row in ws.iter_rows(values_only=True):
            cells = [clean(c) for c in row]
            first = cells[0] if cells else None
            if header is None:
                if first == "Tables":
                    header = [MASTER_COLUMNS.get((c or "").strip().lower()) for c in cells]
                elif first and first.lower().startswith("sme"):
                    info["SMEs"] = "; ".join(x.strip() for x in first.splitlines()[1:] if x.strip())
                elif first:
                    info[first] = ", ".join(c for c in cells[1:] if c)
            elif first:
                rec = {k: redact_text(v) for k, v in zip(header, cells) if k and v}
                rows.append(rec)
        out[report] = {"sheet": ws.title, "info": info, "tables": rows}
    return out


def load_flows():
    wb = openpyxl.load_workbook(FLOW_REGISTER, read_only=True, data_only=True, keep_links=False)
    flows = []
    for ws in wb.worksheets:
        header = None
        for row in ws.iter_rows(values_only=True):
            cells = [clean(c) for c in row]
            if header is None:
                if "Name" in cells:
                    header = [(c or "").lower() for c in cells]
                continue
            rec = dict(zip(header, cells))
            if rec.get("name"):
                flows.append({"sheet": ws.title, "no": rec.get("no"), "name": rec["name"],
                              "type": rec.get("type"), "purpose": rec.get("purpose"),
                              "detail": rec.get("detailed process summary")})
    return flows


def master_list_name(location):
    """'Site - List - View' -> 'List'."""
    parts = [p.strip() for p in (location or "").split(" - ")]
    return parts[1] if len(parts) >= 3 else None


# --- matching ----------------------------------------------------------------

def match_documented_flows(row, flows):
    """Flows named in a master-file row. Returns ({flow: evidence}, [unmatched names])."""
    by_norm = {norm(f["name"]): f["name"] for f in flows}
    all_text = " ".join(row.get(k, "") for k in ("flow", "maintained_by", "notes"))
    n = norm(all_text)
    matched = {name: "named in Master File" for k, name in by_norm.items() if k and k in n}

    unmatched = []
    for phrase in re.split(r"[,;\n\[\]()]|\band\b|\s-|-\s", row.get("flow", "")):
        p = re.sub(r"(?i)\bflows?\b|^\s*(yes|no)\b", "", phrase).strip(" .:-")
        k = norm(p)
        if len(k) < 6 or k in GENERIC_FLOW_WORDS or any(norm(m) in k for m in matched):
            continue
        close = difflib.get_close_matches(k, list(by_norm), n=1, cutoff=0.8)
        if close:
            matched.setdefault(by_norm[close[0]], f'similar name in Master File: "{p}"')
        elif len(p.split()) <= 4 and (re.search(r"[a-z][A-Z]", p) or all(w[:1].isupper() for w in p.split())):
            unmatched.append(p)
    return matched, unmatched


def list_patterns(names):
    """Regexes for '<name> list' / '<name> SharePoint list' mentions, longest name first."""
    pats = []
    for name in sorted(set(names), key=len, reverse=True):
        words = [re.escape(w) for w in re.split(r"\s+|(?<=&)|(?=&)", name) if w]
        if words:
            body = r"\s*".join(words)
            pats.append((name, re.compile(rf"(?i)(?<![\w]){body}[\"”’']?\s+(?:SharePoint\s+)?list\b")))
    return pats


def infer_flow_lists(flows, list_names):
    """{flow name: {list name: evidence}} from flow descriptions and flow names."""
    pats = list_patterns(list_names)
    out = {}
    for f in flows:
        text = " ".join(x for x in (f["purpose"], f["detail"]) if x)
        found = {}
        for name, pat in pats:  # longest first, blanking matches so "Applications" can't re-match "Asbestos Applications"
            m = pat.search(text)
            if m:
                found[name] = f'description mentions "{m.group(0).strip()}"'
                text = pat.sub(" ", text)
        fname = norm(f["name"])
        for name in sorted(set(list_names), key=len, reverse=True):
            k = norm(name)
            if len(k) >= 8 and k in fname:
                found.setdefault(name, "flow name contains the list name")
                fname = fname.replace(k, "")
        out[f["name"]] = found
    return out


# --- index -------------------------------------------------------------------

def build_index(models, master, flows):
    flow_by_name = {f["name"]: f for f in flows}
    lists = {}  # (site, list_id) -> info
    table_rows = []  # one per (report, table)

    for m in models:
        if m["status"] != "ok":
            continue
        mrows = defaultdict(list)
        for r in master.get(m["report"], {}).get("tables", []):
            mrows[tkey(r["table"])].append(r)
        for t in m["tables"]:
            src = t["source"]
            rows = mrows.get(tkey(t["name"]), [])
            keys = [(site_name(src.get("site")), lid) for lid in src.get("list_ids", [])]
            for key in keys:
                li = lists.setdefault(key, {"site": src.get("site"), "list_id": key[1], "names": Counter(),
                                            "master_names": Counter(), "tables": [], "documented": {},
                                            "inferred": {}})
                li["names"][t["name"]] += 1
                for r in rows:
                    ln = master_list_name(r.get("location"))
                    if ln:
                        li["master_names"][ln] += 1
                li["tables"].append((m["report"], t["name"]))
            documented, unmatched = {}, []
            for r in rows:
                d, u = match_documented_flows(r, flows)
                for k, v in d.items():
                    documented.setdefault(k, v)
                unmatched += u
            for key in keys:
                for k, v in documented.items():
                    lists[key]["documented"].setdefault(k, f"{v} ({m['report']}: {t['name']})")
            table_rows.append({"report": m["report"], "table": t, "master": rows, "list_keys": keys,
                               "documented": documented, "unmatched_flows": unmatched, "inferred": {}})

    for li in lists.values():
        li["name"] = (li["master_names"].most_common(1) or li["names"].most_common(1))[0][0]
        li["all_names"] = set(li["names"]) | set(li["master_names"])

    name_to_keys = defaultdict(set)
    for key, li in lists.items():
        for n in li["all_names"]:
            name_to_keys[n].add(key)
    inferred = infer_flow_lists(flows, list(name_to_keys))
    for flow, found in inferred.items():
        for name, evidence in found.items():
            for key in name_to_keys[name]:
                lists[key]["inferred"].setdefault(flow, evidence)
    for tr in table_rows:
        for key in tr["list_keys"]:
            for flow, ev in lists[key]["inferred"].items():
                if flow not in tr["documented"]:
                    tr["inferred"][flow] = ev

    return {"lists": lists, "tables": table_rows, "flows": flows, "flow_by_name": flow_by_name}


def dax_refs(m):
    """Bracketed names referenced anywhere in the model's DAX (approximate)."""
    texts = [x["expression"] or "" for x in m["measures"]]
    for t in m["tables"]:
        if t["expression_language"] == "DAX":
            texts.append(t["expression"] or "")
        texts += [c["expression"] or "" for c in t["columns"]]
    return set(re.findall(r"\[([^\]]+)\]", "\n".join(texts)))


def column_usage(m):
    refs = dax_refs(m)
    rel_cols = {(r["from_table"], r["from_column"]) for r in m["relationships"]} | \
               {(r["to_table"], r["to_column"]) for r in m["relationships"]}
    usage = {}
    for t in m["tables"]:
        for c in t["columns"]:
            how = [w for w, ok in (("report", c["used_in_report"]),
                                   ("relationship", (t["name"], c["name"]) in rel_cols),
                                   ("DAX", c["name"] in refs)) if ok]
            usage[(t["name"], c["name"])] = how
    return usage


# --- gaps --------------------------------------------------------------------

def find_gaps(models, master, index):
    gaps = []

    def add(model, category, item, detail=""):
        gaps.append({"model": model, "category": category, "item": item, "detail": detail})

    for m in models:
        if m["status"] != "ok":
            add(m["report"], "Extraction failed", m["file"], m.get("error", ""))
    ok = [m for m in models if m["status"] == "ok"]

    for m in ok:
        rep = m["report"]
        mt = master.get(rep, {"tables": []})["tables"]
        model_keys = {tkey(t["name"]) for t in m["tables"]}
        staging = {tkey(q) for t in m["tables"] for q in t["source"].get("staging_queries", [])}
        master_keys = {tkey(r["table"]) for r in mt}
        if rep not in master:
            add(rep, "Not in Master File", rep, "No sheet for this dashboard.")
        for t in m["tables"]:
            if tkey(t["name"]) not in master_keys:
                add(rep, "Model table not in Master File", t["name"], t["source"]["kind"])
        seen = set()
        for r in mt:
            k = tkey(r["table"])
            if k in model_keys or k in seen:
                continue
            seen.add(k)
            detail = ("Power Query-only (staging) query, not loaded to the model" if k in staging
                      else "Not found in the model")
            add(rep, "Master File table not in model", r["table"], detail)
        for tr in (x for x in index["tables"] if x["report"] == rep):
            for r in tr["master"]:
                loc = r.get("location", "")
                if master_list_name(loc) and tr["table"]["source"]["kind"] == "Manual table (Enter Data)":
                    add(rep, "Source differs from Master File", tr["table"]["name"],
                        f'Master File says "{loc}"; the model uses an Enter Data table.')
            for u in tr["unmatched_flows"]:
                add(rep, "Flow name not in Flow Register", tr["table"]["name"], u)
            if tr["documented"]:
                for flow, ev in tr["documented"].items():
                    if ev.startswith("similar name"):
                        add(rep, "Flow matched by similar name", tr["table"]["name"], f"{flow} ← {ev}")
            elif any((r.get("flow") or "").lower().startswith("yes") for r in tr["master"]):
                txt = "; ".join(r["flow"] for r in tr["master"] if r.get("flow"))
                if tr["inferred"]:
                    txt += f" → likely: {', '.join(sorted(tr['inferred']))} (inferred)"
                add(rep, "Flow documented but not named", tr["table"]["name"], txt)
        for f in m["unresolved_report_fields"]:
            add(rep, "Report field not in model", f, "Referenced by a visual, filter or formatting rule")
        rel_tables = {r["from_table"] for r in m["relationships"]} | {r["to_table"] for r in m["relationships"]}
        lonely = [t["name"] for t in m["tables"] if t["name"] not in rel_tables]
        if lonely and m["relationships"]:
            add(rep, "Tables without relationships", f"{len(lonely)} table{'s' if len(lonely) > 1 else ''}",
                ", ".join(lonely))
        usage = column_usage(m)
        for t in m["tables"]:
            unused = [c["name"] for c in t["columns"] if not usage[(t["name"], c["name"])]]
            if unused:
                add(rep, "Columns not referenced", f"{t['name']} ({len(unused)}/{len(t['columns'])})",
                    ", ".join(unused))
        for p in m["report_meta"]["pages"]:
            name = p["name"] or ""
            if re.match(r"(?i)^duplicate of", name) or (p["hidden"] and re.match(r"(?i)^page \d+$", name)):
                add(rep, "Possibly leftover page", p["name"], f"{p['visuals']} visuals"
                    + (", hidden" if p["hidden"] else ""))

    linked = {f for li in index["lists"].values() for f in (*li["documented"], *li["inferred"])}
    linked |= {f for tr in index["tables"] for f in tr["documented"]}
    for f in index["flows"]:
        if f["name"] not in linked:
            add("(all)", "Flow not linked to any model table", f["name"], f"{f['sheet']} · {f['type'] or ''}")
    return gaps


# --- Markdown: per model -----------------------------------------------------

def er_diagram(m):
    rels = m["relationships"]
    types = {(t["name"], c["name"]): c["data_type"] for t in m["tables"] for c in t["columns"]}
    names = sorted({r["from_table"] for r in rels} | {r["to_table"] for r in rels})
    ids, used = {}, set()
    for n in names:
        base = re.sub(r"\W+", "_", n).strip("_") or "T"
        base = base if not base[0].isdigit() else f"T_{base}"
        i, cand = 1, base
        while cand in used:
            i += 1
            cand = f"{base}_{i}"
        used.add(cand)
        ids[n] = cand
    keys = defaultdict(dict)
    for r in rels:
        keys[r["to_table"]].setdefault(r["to_column"], set()).add("PK")
        keys[r["from_table"]].setdefault(r["from_column"], set()).add("FK")
    lines = ["```mermaid", "erDiagram"]
    for n in names:
        lines.append(f'    {ids[n]} {{' if ids[n] == n else f'    {ids[n]}["{n}"] {{')
        for col, marks in sorted(keys[n].items()):
            typ = re.sub(r"\W", "", (types.get((n, col)) or "unknown").lower()) or "unknown"
            attr = re.sub(r"\W", "_", col)
            attr = attr if not attr[0].isdigit() else f"_{attr}"
            lines.append(f"        {typ} {attr} {', '.join(sorted(marks, reverse=True))}")
        lines.append("    }")
    for r in rels:
        card = CARDINALITY.get(r["cardinality"], "}o--o{")
        if not r["active"]:
            card = card.replace("--", "..")
        label = f'{r["from_column"]} → {r["to_column"]}' + ("" if r["active"] else " (inactive)") \
            + (" ⇄" if r["cross_filter"] == "Both" else "")
        lines.append(f'    {ids[r["from_table"]]} {card} {ids[r["to_table"]]} : "{label.replace(chr(34), chr(39))}"')
    lines.append("```")
    return "\n".join(lines)


def source_text(src, master_rows, lists):
    kind = src["kind"]
    if src.get("list_ids"):
        parts = []
        for lid in src["list_ids"]:
            li = lists.get((site_name(src.get("site")), lid))
            parts.append(f"`{li['name'] if li else lid}`")
        return f"{kind}: {', '.join(parts)} on `{site_name(src.get('site'))}`"
    if src.get("files"):
        return f"{kind}: {', '.join(code(f) for f in src['files'])} on `{site_name(src.get('site'))}`"
    if src.get("url"):
        return f"{kind}: `{src['url']}`"
    return kind


def flows_text(tr, prefix="../"):
    parts = []
    if tr["documented"]:
        parts.append("documented: " + ", ".join(flow_link(f, prefix) for f in sorted(tr["documented"])))
    if tr["inferred"]:
        parts.append("inferred: " + ", ".join(flow_link(f, prefix) for f in sorted(tr["inferred"])))
    return "; ".join(parts)


def write_model_doc(m, master, index):
    rep = m["report"]
    ms = master.get(rep, {"sheet": None, "info": {}, "tables": []})
    trs = {tr["table"]["name"]: tr for tr in index["tables"] if tr["report"] == rep}
    lists = index["lists"]
    usage = column_usage(m)
    meta = m["report_meta"]
    ncols = sum(len(t["columns"]) for t in m["tables"])
    hidden_pages = sum(1 for p in meta["pages"] if p["hidden"])
    purpose_file = DOCS / "models" / "_purpose" / f"{rep}.md"  # optional hand-written section
    purpose = purpose_file.read_text(encoding="utf-8").strip() if purpose_file.exists() else ""

    out = [f"# {rep}", "",
           f"*Generated by `scripts/build_docs.py` from `{m['file']}`"
           + (f', the Master File sheet "{ms["sheet"]}"' if ms["sheet"] else "")
           + " and the Power Automate Flow Master Register. Regenerate it rather than editing by hand.*", "",
           "| | |", "|---|---|"]
    for label, key in (("Used by", "Used By"), ("Scheduled refresh", "Scheduled Refresh"),
                       ("Dashboard owner", "Dashboard Owner"), ("SMEs", "SMEs")):
        out.append(f"| {label} | {cell(ms['info'].get(key))} |")
    out += [f"| Report | {meta['format']}, {len(meta['pages'])} pages ({hidden_pages} hidden) |",
            f"| Power BI IDs | dataset `{meta['dataset_id']}` · report `{meta['report_id']}` |", "",
            f"**Model:** {len(m['tables'])} tables · {ncols} columns · {len(m['measures'])} measures · "
            f"{len(m['relationships'])} relationships. "
            + (f"{m['auto_date_tables']} hidden auto date/time tables (`LocalDateTable_*`) are left out."
               if m["auto_date_tables"] else ""), ""]
    out += (["- [Business purpose](#business-purpose)"] if purpose else []) + [
        "- [Relationship diagram](#relationship-diagram)", "- [Tables](#tables)",
        "- [Table details](#table-details)", "- [Measures](#measures)", "- [Report pages](#report-pages)", ""]

    if purpose:
        out += ["## Business purpose", "",
                f"*Hand-written: edit `docs/models/_purpose/{rep}.md`, not this section.*", "", purpose, ""]

    # Relationships
    out += ["## Relationship diagram", ""]
    if m["relationships"]:
        out += [er_diagram(m), "",
                "`}o--||` many-to-one · `||--||` one-to-one · dotted line = inactive · ⇄ = filters in both directions. "
                "Only key columns are shown.", ""]
        rel_tables = {r["from_table"] for r in m["relationships"]} | {r["to_table"] for r in m["relationships"]}
        lonely = [t["name"] for t in m["tables"] if t["name"] not in rel_tables]
        if lonely:
            out += ["Not related to any other table: " + ", ".join(code(n) for n in lonely) + ".", ""]
        out += ["| From | To | Cardinality | Cross-filter | Active |", "|---|---|---|---|---|"]
        for r in m["relationships"]:
            out.append(f"| {cell(r['from_table'])}[{cell(r['from_column'])}] | {cell(r['to_table'])}[{cell(r['to_column'])}] "
                       f"| {r['cardinality']} | {r['cross_filter']} | {'yes' if r['active'] else 'no'} |")
        out.append("")
    else:
        out += ["This model has no relationships; every table stands alone and visuals combine them through measures.", ""]

    # Table summary
    out += ["## Tables", "", "| Table | Source | Columns | In report | In Master File | Flows |",
            "|---|---|---|---|---|---|"]
    for t in m["tables"]:
        tr = trs[t["name"]]
        out.append(f"| [{cell(t['name'])}](#{slug(t['name'])}) | {cell(source_text(t['source'], tr['master'], lists))} "
                   f"| {len(t['columns'])} | {'yes' if t['used_in_report'] else 'no'} "
                   f"| {'yes' if tr['master'] else '**no**'} | {cell(flows_text(tr))} |")
    out.append("")

    # Table details
    out += ["## Table details", ""]
    for t in m["tables"]:
        tr = trs[t["name"]]
        src = t["source"]
        out += [f"### {t['name']}", "", f"**Source:** {source_text(src, tr['master'], lists)}"
                + (f" (list id `{', '.join(src['list_ids'])}`)" if src.get("list_ids") else "")
                + (f", folder `{src['folder']}`" if src.get("folder") else ""), ""]
        reads = [code(r) for r in src.get("references", [])] + \
                [f"{code(r)} (Power Query only)" for r in src.get("staging_queries", [])]
        if reads:
            out += [f"**Reads from:** {', '.join(reads)}", ""]
        if t.get("description"):
            out += [f"**Description:** {t['description']}", ""]
        for r in tr["master"]:
            bits = [f"{label}: {r[k]}" for label, k in (("Location", "location"), ("Team", "team"),
                    ("Maintained by", "maintained_by"), ("Contact", "contact"), ("Flow", "flow"),
                    ("App", "app"), ("Notes", "notes"), ("Comments", "comments")) if r.get(k)]
            out += ["**Master File:** " + " · ".join(cell(b) for b in bits), ""]
        if not tr["master"]:
            out += ["**Master File:** not documented.", ""]
        if tr["documented"] or tr["inferred"]:
            out.append("**Power Automate flows:**")
            out.append("")
            for f, ev in sorted(tr["documented"].items()):
                out.append(f"- {flow_link(f, '../')}: documented ({ev})")
            for f, ev in sorted(tr["inferred"].items()):
                out.append(f"- {flow_link(f, '../')}: inferred ({ev})")
            out.append("")
        out += ["| Column | Type | Kind | Hidden | Referenced by |", "|---|---|---|---|---|"]
        for c in t["columns"]:
            how = usage[(t["name"], c["name"])]
            out.append(f"| {cell(c['name'])} | {c['data_type']} | {c['column_type']} | {'yes' if c['hidden'] else ''} "
                       f"| {', '.join(how) if how else '—'} |")
        out.append("")
        calc = [c for c in t["columns"] if c["expression"]]
        if calc:
            out += ["<details><summary>Calculated columns (DAX)</summary>", ""]
            for c in calc:
                out += [f"**{c['name']}**", "", "```dax", c["expression"].strip(), "```", ""]
            out += ["</details>", ""]
        if t["expression"]:
            lang = "dax" if t["expression_language"] == "DAX" else "m"
            title = "Table expression (DAX)" if lang == "dax" else "Power Query (M), redacted"
            out += [f"<details><summary>{title}</summary>", "", f"```{lang}", t["expression"].strip(), "```", "",
                    "</details>", ""]

    # Measures
    out += ["## Measures", ""]
    if m["measures"]:
        out += ["| Measure | Home table | In report | Display folder | Description |", "|---|---|---|---|---|"]
        for x in sorted(m["measures"], key=lambda x: (x["table"], x["name"])):
            out.append(f"| {cell(x['name'])} | {cell(x['table'])} | {'yes' if x['used_in_report'] else 'no'} "
                       f"| {cell(x['display_folder'])} | {cell(x['description'])} |")
        out += ["", "<details><summary>Measure expressions (DAX)</summary>", ""]
        for x in sorted(m["measures"], key=lambda x: (x["table"], x["name"])):
            out += [f"**{x['table']}[{x['name']}]**", "", "```dax", (x["expression"] or "").strip(), "```", ""]
        out += ["</details>", ""]
    else:
        out += ["No measures.", ""]

    # Pages
    out += ["## Report pages", "", "| Page | Hidden | Visuals |", "|---|---|---|"]
    for p in meta["pages"]:
        out.append(f"| {cell(p['name'])} | {'yes' if p['hidden'] else ''} | {p['visuals']} |")
    out.append("")

    path = DOCS / "models" / f"{rep}.md"
    path.write_text("\n".join(out), encoding="utf-8")
    return path


# --- Markdown: lineage, gaps, index -----------------------------------------

def mermaid_id(prefix, i):
    return f"{prefix}{i}"


def mlabel(s):
    return str(s).replace('"', "#quot;")


def write_lineage(models, index):
    lists = index["lists"]
    keys = sorted(lists, key=lambda k: (k[0] or "", lists[k]["name"].lower()))
    reports = [m["report"] for m in models if m["status"] == "ok"]

    out = ["# Data lineage: flows → SharePoint lists → Power BI models", "",
           "*Generated by `scripts/build_docs.py`. A flow is **documented** when the Dashboards Master File names it "
           "for a table, and **inferred** when the Flow Register's description mentions the list "
           '(for example "the Applications list") or the flow is named after it. Inferred links are leads to '
           "confirm, not facts.*", "",
           "- [SharePoint lists](#sharepoint-lists)", "- [Lists → models](#lists--models)",
           "- [Flows → lists](#flows--lists)", "- [Other sources](#other-sources)", "- [Flows](#flows)", "",
           "## SharePoint lists", "",
           "| List | Site | List ID | Used by (model: table) | Flows: documented | Flows: inferred |",
           "|---|---|---|---|---|---|"]
    for k in keys:
        li = lists[k]
        used = "<br>".join(f"[{cell(r)}]({model_link(r, slug(t))}): {cell(t)}" for r, t in sorted(set(li["tables"])))
        out.append(f"| **{cell(li['name'])}** | {cell(k[0])} | `{k[1]}` | {used} "
                   f"| {', '.join(flow_link(f) for f in sorted(li['documented']))} "
                   f"| {', '.join(flow_link(f) for f in sorted(set(li['inferred']) - set(li['documented'])))} |")
    out.append("")

    out += ["## Lists → models", "", "```mermaid", "flowchart LR"]
    by_site = defaultdict(list)
    for i, k in enumerate(keys):
        by_site[k[0]].append((i, k))
    for s, (site, items) in enumerate(sorted(by_site.items(), key=lambda x: x[0] or "")):
        out.append(f'    subgraph S{s}["{mlabel(site)}"]')
        for i, k in items:
            out.append(f'        {mermaid_id("L", i)}["{mlabel(lists[k]["name"])}"]')
        out.append("    end")
    for j, r in enumerate(reports):
        out.append(f'    {mermaid_id("M", j)}(["{mlabel(r)}"])')
    for i, k in enumerate(keys):
        for r in sorted({r for r, _ in lists[k]["tables"]}):
            out.append(f'    {mermaid_id("L", i)} --> {mermaid_id("M", reports.index(r))}')
    out += ["```", ""]

    flow_names = sorted({f for k in keys for f in (*lists[k]["documented"], *lists[k]["inferred"])})
    out += ["## Flows → lists", "", "Solid arrow = documented in the Master File; dotted = inferred.", "",
            "```mermaid", "flowchart LR"]
    for i, f in enumerate(flow_names):
        out.append(f'    {mermaid_id("F", i)}["{mlabel(f)}"]')
    for i, k in enumerate(keys):
        if lists[k]["documented"] or lists[k]["inferred"]:
            out.append(f'    {mermaid_id("L", i)}[("{mlabel(lists[k]["name"])}")]')
    for i, k in enumerate(keys):
        for f in sorted(lists[k]["documented"]):
            out.append(f'    {mermaid_id("F", flow_names.index(f))} --> {mermaid_id("L", i)}')
        for f in sorted(set(lists[k]["inferred"]) - set(lists[k]["documented"])):
            out.append(f'    {mermaid_id("F", flow_names.index(f))} -.-> {mermaid_id("L", i)}')
    out += ["```", ""]

    out += ["## Other sources", "", "| Source | Kind | Used by (model: table) |", "|---|---|---|"]
    other = defaultdict(set)
    for m in models:
        if m["status"] != "ok":
            continue
        for t in m["tables"]:
            s = t["source"]
            if s.get("list_ids"):
                continue
            if s.get("files"):
                for f in s["files"]:
                    other[(f"{f} ({site_name(s.get('site'))})", s["kind"])].add((m["report"], t["name"]))
            elif s.get("url"):
                other[(s["url"], s["kind"])].add((m["report"], t["name"]))
            elif s["kind"] == "Manual table (Enter Data)":
                other[(t["name"], s["kind"])].add((m["report"], t["name"]))
    for (src, kind), used in sorted(other.items(), key=lambda x: (x[0][1], x[0][0].lower())):
        out.append(f"| {cell(src)} | {kind} | "
                   + "<br>".join(f"[{cell(r)}]({model_link(r, slug(t))}): {cell(t)}" for r, t in sorted(used)) + " |")
    out.append("")

    out += ["## Flows", ""]
    for f in index["flows"]:
        doc = [k for k in keys if f["name"] in lists[k]["documented"]]
        inf = [k for k in keys if f["name"] in lists[k]["inferred"] and k not in doc]
        tables_doc = sorted({(tr["report"], tr["table"]["name"]) for tr in index["tables"]
                             if f["name"] in tr["documented"]})
        out += [f"### {f['name']}", "",
                f"*{f['sheet']} #{f['no']} · {f['type'] or 'type not recorded'}*", ""]
        if f["purpose"]:
            out += [cell(f["purpose"]).replace("<br>", " "), ""]
        if doc or tables_doc:
            out.append("- **Documented for:** " + ", ".join(
                f"[{cell(r)}]({model_link(r, slug(t))}): {cell(t)}" for r, t in tables_doc))
        if inf:
            out.append("- **Inferred lists:** " + "; ".join(
                f"{lists[k]['name']} ({lists[k]['inferred'][f['name']]})" for k in inf))
        if not (doc or inf or tables_doc):
            out.append("- Not linked to any table in the 9 models.")
        out.append("")
    path = DOCS / "lineage.md"
    path.write_text("\n".join(out), encoding="utf-8")
    return path


def write_gaps(gaps):
    order = ["Extraction failed", "Not in Master File", "Model table not in Master File",
             "Master File table not in model", "Source differs from Master File", "Flow name not in Flow Register",
             "Flow matched by similar name", "Flow documented but not named", "Flow not linked to any model table",
             "Report field not in model", "Tables without relationships", "Possibly leftover page",
             "Columns not referenced"]
    notes = {
        "Master File table not in model": "Tables the Master File lists for a dashboard that its model doesn't load.",
        "Flow matched by similar name": "The Master File's spelling differs from the register; check the match is right.",
        "Flow not linked to any model table": "Register flows that no Master File entry names and no list mention links to a model table.",
        "Report field not in model": "Report JSON points at a table or field the model doesn't have. These are usually "
                                     "leftovers from copying pages between reports, and can show as broken visuals "
                                     "or ignored formatting rules.",
        "Columns not referenced": "Columns not used by any visual, relationship or DAX expression. Approximate: use "
                                  "inside other Power Query steps isn't checked, so confirm before removing.",
    }
    by_cat = defaultdict(list)
    for g in gaps:
        by_cat[g["category"]].append(g)
    out = ["# Documentation and model gaps", "",
           "*Generated by `scripts/build_docs.py` by comparing the 9 models with the Dashboards Master File and the "
           "Power Automate Flow Master Register.*", "", "| Check | Count |", "|---|---|"]
    for c in order:
        out.append(f"| [{c}](#{slug(c)}) | {len(by_cat.get(c, []))} |")
    out.append("")
    for c in order:
        items = by_cat.get(c, [])
        out += [f"## {c}", ""]
        if c in notes:
            out += [notes[c], ""]
        if not items:
            out += ["None.", ""]
            continue
        if c == "Columns not referenced":
            for g in items:
                out.append(f"- **{cell(g['model'])} · {cell(g['item'])}**: {cell(g['detail'])}")
            out.append("")
            continue
        out += ["| Model | Item | Detail |", "|---|---|---|"]
        for g in items:
            out.append(f"| {cell(g['model'])} | {cell(g['item'])} | {cell(g['detail'])} |")
        out.append("")
    path = DOCS / "gaps.md"
    path.write_text("\n".join(out), encoding="utf-8")
    return path


def write_index(models, gaps):
    out = ["# Power BI model documentation", "",
           "*Generated by `scripts/build_docs.py` from `pbix/` and `xlsx/`. To refresh, see CLAUDE.md.*", "",
           "| Model | Tables | Columns | Measures | Relationships | Pages |", "|---|---|---|---|---|---|"]
    for m in models:
        if m["status"] != "ok":
            out.append(f"| {m['report']} | extraction failed | | | | |")
            continue
        out.append(f"| [{m['report']}]({model_link(m['report'])}) | {len(m['tables'])} "
                   f"| {sum(len(t['columns']) for t in m['tables'])} | {len(m['measures'])} "
                   f"| {len(m['relationships'])} | {len(m['report_meta']['pages'])} |")
    counts = Counter(g["category"] for g in gaps)
    out += ["", "- [lineage.md](lineage.md): SharePoint lists, other sources and Power Automate flows across all models.",
            f"- [gaps.md](gaps.md): {sum(counts.values())} findings, such as tables missing from the Master File "
            "and stale report fields.",
            "- [Model Schema Register.xlsx](Model%20Schema%20Register.xlsx): the same data as filterable sheets, "
            "plus review findings, fix steps, report pages and business purpose. Maintained by hand; "
            "`build_docs.py` doesn't write it.",
            "- [Model Schema Register v0.xlsx](Model%20Schema%20Register%20v0.xlsx): the generated register only, "
            "rewritten on each `build_docs.py` run.",
            "- `models/_extracted/*.json`: raw extraction output (no row data; M code redacted).", ""]
    path = DOCS / "README.md"
    path.write_text("\n".join(out), encoding="utf-8")
    return path


# --- Excel -------------------------------------------------------------------

def _xl(v):
    if v is None:
        return None
    if isinstance(v, bool):
        return "yes" if v else "no"
    if isinstance(v, (list, tuple, set)):
        v = ", ".join(str(x) for x in v)
    s = ILLEGAL_CHARACTERS_RE.sub("", str(v))
    return s[:32000]


def write_excel(models, index, gaps):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    lists = index["lists"]
    trs = {(tr["report"], tr["table"]["name"]): tr for tr in index["tables"]}

    def sheet(title, header, rows, widths):
        ws = wb.create_sheet(title)
        ws.append(header)
        for c in ws[1]:
            c.font = Font(bold=True)
        for r in rows:
            ws.append([_xl(v) for v in r])
        for row in ws.iter_rows(min_row=2):
            for c in row:
                if isinstance(c.value, str):
                    c.data_type = "s"  # keep DAX/M text that starts with "=" from becoming a formula
                    c.alignment = Alignment(vertical="top", wrap_text=len(c.value) > 60)
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
        ws.freeze_panes = "C2"
        ws.auto_filter.ref = ws.dimensions

    ok = [m for m in models if m["status"] == "ok"]
    rows = []
    for m in ok:
        for t in m["tables"]:
            tr = trs[(m["report"], t["name"])]
            s = t["source"]
            mr = tr["master"][0] if tr["master"] else {}
            rows.append([m["report"], t["name"], s["kind"], site_name(s.get("site")),
                         s.get("list_ids"), [lists[k]["name"] for k in tr["list_keys"]],
                         s.get("files") or s.get("url"), s.get("references"), s.get("staging_queries"),
                         len(t["columns"]), t["used_in_report"], bool(tr["master"]), mr.get("location"),
                         mr.get("team"), mr.get("maintained_by"), mr.get("contact"), mr.get("flow"),
                         mr.get("app"), mr.get("notes"), sorted(tr["documented"]), sorted(tr["inferred"]),
                         t["expression"]])
    sheet("Tables", ["Model", "Table", "Source kind", "Site", "List ID", "List name", "File / URL", "Reads from",
                     "Power Query-only inputs", "Columns", "Used in report", "In Master File",
                     "Master: location", "Master: team", "Master: maintained by", "Master: contact",
                     "Master: flow", "Master: app", "Master: notes", "Flows (documented)", "Flows (inferred)",
                     "Power Query (M, redacted) / DAX"],
          rows, [28, 30, 24, 30, 38, 26, 30, 30, 26, 9, 9, 9, 40, 24, 30, 22, 30, 20, 40, 30, 30, 60])

    rows = []
    for m in ok:
        usage = column_usage(m)
        for t in m["tables"]:
            for c in t["columns"]:
                rows.append([m["report"], t["name"], c["name"], c["data_type"], c["column_type"], c["hidden"],
                             usage[(t["name"], c["name"])] or "none", c["format"], c["description"],
                             c["source_column"], c["expression"]])
    sheet("Columns", ["Model", "Table", "Column", "Data type", "Kind", "Hidden", "Referenced by", "Format",
                      "Description", "Source column", "DAX expression"],
          rows, [28, 30, 30, 11, 20, 8, 22, 16, 30, 26, 60])

    rows = [[m["report"], x["table"], x["name"], x["used_in_report"], x["display_folder"], x["description"],
             x["expression"]] for m in ok for x in m["measures"]]
    sheet("Measures", ["Model", "Home table", "Measure", "Used in report", "Display folder", "Description", "DAX"],
          rows, [28, 26, 34, 10, 18, 30, 80])

    rows = [[m["report"], r["from_table"], r["from_column"], r["to_table"], r["to_column"], r["cardinality"],
             r["cross_filter"], r["active"]] for m in ok for r in m["relationships"]]
    sheet("Relationships", ["Model", "From table", "From column", "To table", "To column", "Cardinality",
                            "Cross filter", "Active"], rows, [28, 28, 24, 28, 24, 11, 12, 8])

    fb = index["flow_by_name"]
    rows = []
    for tr in index["tables"]:
        for f, ev in sorted(tr["documented"].items()):
            rows.append([tr["report"], tr["table"]["name"], f, "documented", ev, fb[f]["type"], fb[f]["sheet"]])
        for f, ev in sorted(tr["inferred"].items()):
            rows.append([tr["report"], tr["table"]["name"], f, "inferred", ev, fb[f]["type"], fb[f]["sheet"]])
    sheet("TableFlows", ["Model", "Table", "Flow", "Link", "Evidence", "Flow type", "Register sheet"],
          rows, [28, 30, 32, 11, 60, 14, 16])

    rows = []
    for f in index["flows"]:
        linked = sorted({f"{tr['report']}: {tr['table']['name']}" for tr in index["tables"]
                         if f["name"] in tr["documented"] or f["name"] in tr["inferred"]})
        rows.append([f["sheet"], f["no"], f["name"], f["type"],
                     sorted({lists[k]["name"] for k in lists
                             if f["name"] in lists[k]["documented"] or f["name"] in lists[k]["inferred"]}),
                     linked, f["purpose"]])
    sheet("Flows", ["Register sheet", "No", "Flow", "Type", "Linked lists", "Linked model tables", "Purpose"],
          rows, [16, 6, 34, 14, 36, 60, 80])

    sheet("Gaps", ["Model", "Check", "Item", "Detail"],
          [[g["model"], g["category"], g["item"], g["detail"]] for g in gaps], [28, 34, 40, 90])

    path = DOCS / "Model Schema Register v0.xlsx"  # not the hand-maintained "Model Schema Register.xlsx"
    wb.save(path)
    return path


def main():
    models = load_models()
    master = load_master()
    flows = load_flows()
    index = build_index(models, master, flows)
    gaps = find_gaps(models, master, index)
    (DOCS / "models").mkdir(parents=True, exist_ok=True)
    written = [write_model_doc(m, master, index) for m in models if m["status"] == "ok"]
    written += [write_lineage(models, index), write_gaps(gaps), write_index(models, gaps),
                write_excel(models, index, gaps)]
    for p in written:
        print("wrote", p.relative_to(ROOT).as_posix())
    counts = Counter(g["category"] for g in gaps)
    print("gaps:", ", ".join(f"{k}={v}" for k, v in counts.most_common()))


if __name__ == "__main__":
    main()
