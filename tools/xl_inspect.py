#!/usr/bin/env python3
"""Read-only inspectors for parts the other xlkit tools do not expose.

  xl_inspect.py pivots <file>   pivot table definitions + slicers, as JSON
  xl_inspect.py forms  <file>   UserForm control layout (position, size, caption), as JSON

Both are meant for "extract, then rebuild" work: the rebuild tools (create_pivot_table,
xlide_edit_form) take exactly these fields. Nothing is written.
"""
from __future__ import annotations

import json
import posixpath
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
R_ID = "{%s}id" % NS["r"]
HIMETRIC_PER_POINT = 2540 / 72


def _rels(z: zipfile.ZipFile, part: str) -> dict[str, str]:
    """Relationship id -> absolute part name for a part's .rels."""
    folder, name = posixpath.split(part)
    rels_part = posixpath.join(folder, "_rels", name + ".rels")
    if rels_part not in z.namelist():
        return {}
    out = {}
    for rel in ET.fromstring(z.read(rels_part)).findall("rel:Relationship", NS):
        target = rel.get("Target", "")
        out[rel.get("Id", "")] = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))
    return out


def _sheets(z: zipfile.ZipFile) -> list[tuple[str, str]]:
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = _rels(z, "xl/workbook.xml")
    return [(s.get("name", ""), rels.get(s.get(R_ID, ""), "")) for s in wb.findall("m:sheets/m:sheet", NS)]


def _cache_defs(z: zipfile.ZipFile) -> dict[str, str]:
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = _rels(z, "xl/workbook.xml")
    return {c.get("cacheId", ""): rels.get(c.get(R_ID, ""), "") for c in wb.findall("m:pivotCaches/m:pivotCache", NS)}


def _cache_info(z: zipfile.ZipFile, part: str) -> dict:
    root = ET.fromstring(z.read(part))
    fields = []
    for cf in root.findall("m:cacheFields/m:cacheField", NS):
        items = []
        for it in cf.findall("m:sharedItems/*", NS):
            items.append(it.get("v") if it.get("v") is not None else "")
        fields.append({"name": cf.get("name", ""), "items": items, "formula": cf.get("formula"),
                       "grouped": cf.find("m:fieldGroup", NS) is not None})
    ws = root.find("m:cacheSource/m:worksheetSource", NS)
    for cf, spec in zip(root.findall("m:cacheFields/m:cacheField", NS), fields):
        base = cf.find("m:fieldGroup", NS).get("base") if cf.find("m:fieldGroup", NS) is not None else None
        if base is not None and int(base) < len(fields) and int(base) != fields.index(spec):
            spec["source_name"] = fields[int(base)]["name"]  # a date/number group field: rebuild from its source column
    source = None
    if ws is not None:
        source = f"{ws.get('sheet')}!{ws.get('ref')}" if ws.get("ref") else ws.get("name")
    return {"fields": fields, "source": source, "source_type": (root.find("m:cacheSource", NS).get("type") if root.find("m:cacheSource", NS) is not None else None)}


def pivots(z: zipfile.ZipFile) -> dict:
    caches = {cid: _cache_info(z, part) for cid, part in _cache_defs(z).items() if part in z.namelist()}
    out, slicers = [], []
    for sheet_name, sheet_part in _sheets(z):
        for target in _rels(z, sheet_part).values():
            if not re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", target):
                continue
            root = ET.fromstring(z.read(target))
            cache = caches.get(root.get("cacheId", ""), {"fields": [], "source": None})
            fields = cache["fields"]
            loc = root.find("m:location", NS)
            warnings = []

            def label(fld: int, item_idx: int) -> str:
                items = root.findall("m:pivotFields/m:pivotField", NS)[fld].findall("m:items/m:item", NS)
                x = items[item_idx].get("x") if item_idx < len(items) else None
                shared = fields[fld]["items"]
                return shared[int(x)] if x is not None and int(x) < len(shared) else ""

            def names(tag: str, attr: str = "x") -> list[str]:
                return [fields[int(f.get(attr))].get("source_name") or fields[int(f.get(attr))]["name"]
                        for f in root.findall(f"m:{tag}/m:field", NS) if int(f.get(attr, "-2")) >= 0]

            hidden, sort = {}, {}
            for i, pf in enumerate(root.findall("m:pivotFields/m:pivotField", NS)):
                hid = [label(i, k) for k, it in enumerate(pf.findall("m:items/m:item", NS)) if it.get("h") == "1" and it.get("t") in (None, "data")]
                if hid:
                    hidden[fields[i]["name"]] = hid
                if pf.get("sortType") and pf.get("sortType") != "manual":
                    sort[fields[i]["name"]] = pf.get("sortType")
                if pf.find("m:items/m:item[@m]", NS) is not None or pf.get("multipleItemSelectionAllowed") == "1":
                    warnings.append(f"field '{fields[i]['name']}': multi-select filter")
            pages = []
            for p in root.findall("m:pageFields/m:pageField", NS):
                fld = int(p.get("fld"))
                sel = label(fld, int(p.get("item"))) if p.get("item") not in (None, "") else "(All)"
                pages.append({"field": fields[fld]["name"], "selected": sel})
            datas = [{"field": fields[int(d.get("fld"))]["name"], "function": d.get("subtotal", "sum"),
                      "caption": d.get("name"), "num_fmt_id": d.get("numFmtId")} for d in root.findall("m:dataFields/m:dataField", NS)]
            sort_by = {}
            for i, pf in enumerate(root.findall("m:pivotFields/m:pivotField", NS)):
                ref = pf.find("m:autoSortScope/m:pivotArea/m:references/m:reference/m:x", NS)
                if ref is not None and int(ref.get("v", "-1")) < len(datas):
                    sort_by[fields[i]["name"]] = datas[int(ref.get("v"))]["caption"]
            if any(f["formula"] for f in fields):
                warnings.append("calculated fields: " + ", ".join(f["name"] for f in fields if f["formula"]))
            grouping = {}
            for cache_el in [ET.fromstring(z.read(_cache_defs(z)[root.get("cacheId", "")]))]:
                for cf in cache_el.findall("m:cacheFields/m:cacheField", NS):
                    rp = cf.find("m:fieldGroup/m:rangePr", NS)
                    if rp is not None:
                        base = cf.find("m:fieldGroup", NS).get("base")
                        key = fields[int(base)]["name"] if base is not None and int(base) < len(fields) else cf.get("name", "")
                        grouping[key] = (
                            {"group_dates": [rp.get("groupBy")]} if rp.get("groupBy") and rp.get("groupBy") != "range"
                            else {"group_numbers": {"start": rp.get("startNum"), "end": rp.get("endNum"), "interval": rp.get("groupInterval")}})
            if any(f["grouped"] for f in fields) and not grouping:
                warnings.append("a grouped field uses a custom (manual) group that create_pivot_table cannot rebuild")
            if root.find("m:calculatedItems", NS) is not None:
                warnings.append("calculated items present")
            out.append({
                "sheet": sheet_name, "name": root.get("name"), "location": loc.get("ref") if loc is not None else None,
                "source": cache.get("source"), "source_type": cache.get("source_type"),
                "rows": names("rowFields"), "columns": names("colFields"), "filters": pages, "values": datas,
                "hidden_items": hidden, "sort": sort, "sort_by": sort_by, "grouping": grouping,
                "layout": "compact" if root.get("compact") != "0" else ("outline" if root.get("outline") == "1" else "tabular"),
                "row_grand_totals": root.get("rowGrandTotals") != "0", "col_grand_totals": root.get("colGrandTotals") != "0",
                "style": (root.find("m:pivotTableStyleInfo", NS).get("name") if root.find("m:pivotTableStyleInfo", NS) is not None else None),
                "calculated_fields": [{"name": f["name"], "formula": f["formula"]} for f in fields if f["formula"]],
                "warnings": warnings,
            })
    return {"pivot_tables": out, "slicers": _slicers(z)}


def _col_letters(n: int) -> str:
    s = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _anchors(z: zipfile.ZipFile) -> dict[str, tuple[str, str]]:
    """Drawing object name -> (sheet name, top-left cell) for every drawing on every sheet."""
    out = {}
    for sheet_name, sheet_part in _sheets(z):
        for target in _rels(z, sheet_part).values():
            if not re.fullmatch(r"xl/drawings/drawing\d+\.xml", target):
                continue
            xml = z.read(target).decode("utf8", "ignore")
            for block in re.findall(r"<xdr:(?:twoCellAnchor|oneCellAnchor|absoluteAnchor)\b.*?</xdr:(?:twoCellAnchor|oneCellAnchor|absoluteAnchor)>", xml, re.S):
                nm = re.search(r'<xdr:cNvPr [^>]*?name="([^"]*)"', block)
                col = re.search(r"<xdr:from>\s*<xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>", block, re.S)
                if nm and col:
                    out[nm.group(1)] = (sheet_name, f"{_col_letters(int(col.group(1)))}{int(col.group(2)) + 1}")
    return out


def _slicers(z: zipfile.ZipFile) -> list[dict]:
    anchors = _anchors(z)
    caches = {}
    for part in z.namelist():
        if re.fullmatch(r"xl/slicerCaches/slicerCache\d+\.xml", part):
            root = ET.fromstring(z.read(part))
            caches[root.get("name")] = {"field": root.get("sourceName"), "pivot_tables": [p.get("name") for p in root.iter() if p.tag.endswith("}pivotTable")]}
        elif re.fullmatch(r"xl/timelineCaches/timelineCache\d+\.xml", part):
            root = ET.fromstring(z.read(part))
            caches[root.get("name")] = {"field": root.get("sourceName"), "pivot_tables": [p.get("name") for p in root.iter() if p.tag.endswith("}pivotTable")], "timeline": True}
    out = []
    for part in sorted(z.namelist()):
        kind = "slicer" if re.fullmatch(r"xl/slicers/slicer\d+\.xml", part) else "timeline" if re.fullmatch(r"xl/timelines/timeline\d+\.xml", part) else None
        if not kind:
            continue
        for el in ET.fromstring(z.read(part)):
            cache = caches.get(el.get("cache"), {})
            sheet, at = anchors.get(el.get("name"), (None, None))
            out.append({"kind": kind, "name": el.get("name"), "caption": el.get("caption"), "cache": el.get("cache"),
                        "field": cache.get("field"), "pivot_tables": cache.get("pivot_tables", []),
                        "columns": int(el.get("columnCount", "1")) if kind == "slicer" else None, "sheet": sheet, "at": at,
                        "level": (["years", "quarters", "months", "days"][int(el.get("level", "2"))] if kind == "timeline" and el.get("level", "2").isdigit() and int(el.get("level", "2")) < 4 else None)})
    return out


def forms(path: str) -> dict:
    import pyopenvba

    out = []
    with pyopenvba.ExcelFile(path) as book:
        if not book.has_vba_project():
            return {"forms": [], "note": "no VBA project"}
        for form in book.forms():
            level = form._levels[0]
            sites = {s.name: s for s in level.stream.sites}
            controls = []
            for ctl in form.controls:
                site = sites.get(ctl.name)
                pos = site.position if site is not None else None
                size = ctl.record.sizes.get("Size")
                caption = ctl.record.strings.get("Caption")
                controls.append({
                    "name": ctl.name, "type": ctl.kind.split(".")[-1], "container": ctl.is_container,
                    "left": round(pos[0] / HIMETRIC_PER_POINT, 1) if pos else None,
                    "top": round(pos[1] / HIMETRIC_PER_POINT, 1) if pos else None,
                    "width": round(size.width / HIMETRIC_PER_POINT, 1) if size else None,
                    "height": round(size.height / HIMETRIC_PER_POINT, 1) if size else None,
                    "caption": caption.text if caption is not None else None,
                    "tab_index": ctl.tab_index,
                })
            src = form.designer_source or ""

            def twips(key: str):
                m = re.search(rf"{key}\s*=\s*(\d+)", src)
                return round(int(m.group(1)) / 20, 1) if m else None

            cap = re.search(r'Caption\s*=\s*"((?:[^"]|"")*)"', src)
            out.append({"name": form.name, "caption": cap.group(1).replace('""', '"') if cap else None,
                        "width": twips("ClientWidth"), "height": twips("ClientHeight"), "unit": "points", "controls": controls,
                        "note": "nested containers (Frame/MultiPage) list only the container itself" if any(c["container"] for c in controls) else None})
    return {"forms": out}


def main() -> int:
    if len(sys.argv) != 3 or sys.argv[1] not in ("pivots", "forms"):
        print(__doc__)
        return 2
    kind, path = sys.argv[1], sys.argv[2]
    try:
        if kind == "pivots":
            with zipfile.ZipFile(path) as z:
                result = pivots(z)
        else:
            result = forms(path)
    except Exception as exc:  # noqa: BLE001
        print(json.dumps({"error": f"{type(exc).__name__}: {exc}"}, ensure_ascii=False))
        return 4
    print(json.dumps(result, ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
