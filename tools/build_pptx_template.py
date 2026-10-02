#!/usr/bin/env python3
"""Build the lean PowerPoint theme used by "Export PowerPoint".

Takes a presentation that carries the client look (masters, layouts, theme, logos), removes every slide, chart,
embedded workbook, note, tag, custom XML part and personal/classification metadata, and writes the result to
assets/js/pptx-template.js as base64 so the dashboard can use it offline (file://).

usage:  python3 tools/build_pptx_template.py "/path/to/ATS - Monthly Council.pptx"
Re-run whenever the corporate template changes. Only masters/layouts/theme/media are kept - no slide content.
"""
import base64
import io
import os
import posixpath
import re
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "assets", "js", "pptx-template.js")

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
T_SLIDE = "/relationships/slide"
DROP_TYPES = ("/slide", "/customXml", "/thumbnail", "/custom-properties", "/extended-properties", "/metadata/", "/handoutMaster", "/authors", "/commentAuthors")  # matched as exact suffix below


def rels_path(part):
    d, f = posixpath.split(part)
    return posixpath.join(d, "_rels", f + ".rels")


def parse_rels(xml):
    out = []
    for m in re.finditer(r"<Relationship\b([^>]*?)/?>", xml):
        a = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1)))
        out.append(a)
    return out


def main(src):
    zin = zipfile.ZipFile(src)
    names = set(zin.namelist())
    read = lambda n: zin.read(n).decode("utf8")

    # ---- presentation.xml: no slides, no sections, no handout master
    pres = read("ppt/presentation.xml")
    pres = re.sub(r"<p:sldIdLst>.*?</p:sldIdLst>", "", pres, flags=re.S)
    pres = re.sub(r"<p:handoutMasterIdLst>.*?</p:handoutMasterIdLst>", "", pres, flags=re.S)
    pres = re.sub(r"<p:custShowLst>.*?</p:custShowLst>", "", pres, flags=re.S)
    pres = re.sub(r"<p:extLst>.*?</p:extLst>\s*(?=</p:presentation>)", "", pres, flags=re.S)
    prels = read("ppt/_rels/presentation.xml.rels")

    def keep_rel(a):
        t = a.get("Type", "")
        return not any(t.endswith(x) or (x.endswith("/") and x in t) for x in DROP_TYPES)

    kept = [a for a in parse_rels(prels) if keep_rel(a)]
    prels_new = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="%s">%s</Relationships>' % (
        REL_NS, "".join('<Relationship Id="%s" Type="%s" Target="%s"/>' % (a["Id"], a["Type"], a["Target"]) for a in kept))

    # ---- reachability from the (trimmed) presentation
    keep = {"ppt/presentation.xml", "ppt/_rels/presentation.xml.rels"}
    stack = []
    for a in kept:
        stack.append(posixpath.normpath(posixpath.join("ppt", a["Target"])))
    while stack:
        p = stack.pop()
        if p in keep or p not in names:
            continue
        keep.add(p)
        rp = rels_path(p)
        if rp in names:
            keep.add(rp)
            for a in parse_rels(read(rp)):
                if a.get("TargetMode") == "External":
                    continue
                stack.append(posixpath.normpath(posixpath.join(posixpath.dirname(p), a["Target"])))
    keep.add("[Content_Types].xml")

    # ---- root rels: package + core properties only
    root_rels = [a for a in parse_rels(read("_rels/.rels")) if a["Type"].endswith("/officeDocument") or a["Type"].endswith("/core-properties")]
    root_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="%s">%s</Relationships>' % (
        REL_NS, "".join('<Relationship Id="%s" Type="%s" Target="%s"/>' % (a["Id"], a["Type"], a["Target"]) for a in root_rels))
    keep.add("docProps/core.xml")

    core = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>ATS report</dc:title><dc:creator>ATS Dashboard</dc:creator></cp:coreProperties>')

    # ---- content types: drop overrides of removed parts
    ct = read("[Content_Types].xml")

    def ct_fix(m):
        part = m.group(1).lstrip("/")
        return m.group(0) if part in keep else ""

    ct = re.sub(r'<Override PartName="([^"]+)"[^>]*/>', ct_fix, ct)
    ct = re.sub(r'<Override PartName="/docProps/(app|custom|thumbnail)[^"]*"[^>]*/>', "", ct)
    if "/docProps/core.xml" not in ct:
        ct = ct.replace("</Types>", '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>')

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zout:
        zout.writestr("[Content_Types].xml", ct)
        zout.writestr("_rels/.rels", root_xml)
        zout.writestr("docProps/core.xml", core)
        zout.writestr("ppt/presentation.xml", pres)
        zout.writestr("ppt/_rels/presentation.xml.rels", prels_new)
        for n in sorted(keep):
            if n in ("[Content_Types].xml", "ppt/presentation.xml", "ppt/_rels/presentation.xml.rels", "docProps/core.xml") or n not in names:
                continue
            blob = zin.read(n)
            if n.startswith(("ppt/slideMasters/", "ppt/slideLayouts/")) and n.endswith(".xml"):
                # neutralise the sensitivity-label marking metadata (the visible "OFFICIAL" text stays)
                blob = re.sub(rb'(name="MSIPCMContentMarking"\s+descr=")[^"]*(")', rb"\1\2", blob)
            zout.writestr(n, blob)
    data = buf.getvalue()
    with open(OUT, "w") as f:
        f.write("/* generated by tools/build_pptx_template.py - client PowerPoint theme (masters, layouts, logos); contains no slide content */\n")
        f.write('window.ATS_PPTX_TEMPLATE = "' + base64.b64encode(data).decode("ascii") + '";\n')
    media = sum(zin.getinfo(n).file_size for n in keep if n.startswith("ppt/media/") and n in names)
    print("kept %d of %d parts; media %.2f MB; template zip %.2f MB -> %s" % (len(keep), len(names), media / 1e6, len(data) / 1e6, OUT))


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    main(sys.argv[1])
