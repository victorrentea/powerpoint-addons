#!/usr/bin/env python3
"""📌 Make PowerPoint open the Victor Tools pane by itself whenever this deck opens.

    ./pin-deck.py DECK.pptx [DECK2.pptx …]

A sideloaded add-in has no ribbon button of its own: it has to be opened from Home › Add-ins
› Developer Add-ins in every deck, every session. Office's answer is a tag in the file, but the
one Office.js can write (Office.AutoShowTaskpaneWithDocument with visibility="0") did nothing
on PowerPoint for Mac 16.114 (2026-11-03): the same deck reopened without the pane. Only
visibility="1" in ppt/webextensions/taskpanes.xml brought it up, and Office.js can't write that,
so this writes it into the file directly, adding the add-in's parts if the deck has none yet.

The deck must be closed in PowerPoint (it would overwrite the file on its next save). The
original is copied to ~/Library/Caches/powerpoint-addons/pinned-backups/ first.
"""
import re
import shutil
import subprocess
import sys
import time
import zipfile
from pathlib import Path

ADDIN_ID = "8eeda3fa-1a86-4d27-b218-48e137aba64c"
MANIFEST = Path(__file__).with_name("manifest.xml")
BACKUPS = Path.home() / "Library/Caches/powerpoint-addons/pinned-backups"

WE_NS = "http://schemas.microsoft.com/office/webextensions/webextension/2010/11"
TP_NS = "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_WE = "http://schemas.microsoft.com/office/2011/relationships/webextension"
REL_TP = "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes"
CT_WE = "application/vnd.ms-office.webextension+xml"
CT_TP = "application/vnd.ms-office.webextensiontaskpanes+xml"
HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
AUTO = '<we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>'


def manifest_version():
    return re.search(r"<Version>([^<]+)</Version>", MANIFEST.read_text()).group(1)


def open_in_powerpoint(deck):
    if subprocess.run(["/usr/bin/pgrep", "-x", "Microsoft PowerPoint"], capture_output=True).returncode:
        return False
    names = subprocess.run(["osascript", "-e", 'tell application "Microsoft PowerPoint" to get full name of every presentation'],
                           capture_output=True, text=True).stdout
    return str(deck.resolve()) in names or deck.name in names


def add_override(ct, part, content_type):
    if f'PartName="{part}"' in ct:
        return ct
    return ct.replace("</Types>", f'<Override PartName="{part}" ContentType="{content_type}"/></Types>')


def next_rid(rels):
    return f"rId{max([0] + [int(n) for n in re.findall(r'Id="rId(\d+)"', rels)]) + 1}"


def pin(parts, version):
    """Edits the {name: bytes} parts in place; returns what it did."""
    text = lambda name: parts[name].decode("utf-8")
    ct = text("[Content_Types].xml")
    root_rels = text("_rels/.rels")
    tp_name, tp_rels_name = "ppt/webextensions/taskpanes.xml", "ppt/webextensions/_rels/taskpanes.xml.rels"
    if tp_name not in parts:
        parts[tp_name] = (HEAD + f'<wetp:taskpanes xmlns:wetp="{TP_NS}"></wetp:taskpanes>').encode()
        parts[tp_rels_name] = (HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                               "</Relationships>").encode()
        ct = add_override(ct, "/" + tp_name, CT_TP)
    if REL_TP not in root_rels:
        root_rels = root_rels.replace("</Relationships>", f'<Relationship Id="{next_rid(root_rels)}" Type="{REL_TP}" '
                                                          f'Target="{tp_name}"/></Relationships>')
    tp, tp_rels = text(tp_name), text(tp_rels_name)

    ours = next((n for n in parts if re.fullmatch(r"ppt/webextensions/webextension\d+\.xml", n)
                 and f'id="{ADDIN_ID}"' in text(n)), None)
    if ours:
        we = text(ours)
        we = re.sub(rf'(<we:reference id="{ADDIN_ID}" version=")[^"]*', rf"\g<1>{version}", we)
        if "Office.AutoShowTaskpaneWithDocument" in we:
            we = re.sub(r'(name="Office\.AutoShowTaskpaneWithDocument" value=")[^"]*', r"\g<1>true", we)
        elif "<we:properties/>" in we:
            we = we.replace("<we:properties/>", f"<we:properties>{AUTO}</we:properties>")
        else:
            we = we.replace("<we:properties>", f"<we:properties>{AUTO}", 1)
        parts[ours] = we.encode()
        done = "re-pinned"
    else:
        n = 1 + max([0] + [int(m) for m in re.findall(r"webextension(\d+)\.xml", " ".join(parts))])
        ours = f"ppt/webextensions/webextension{n}.xml"
        parts[ours] = (HEAD + f'<we:webextension xmlns:we="{WE_NS}" id="{{{ADDIN_ID.upper()}}}">'
                       f'<we:reference id="{ADDIN_ID}" version="{version}" store="developer" storeType="Registry"/>'
                       f"<we:alternateReferences/><we:properties>{AUTO}</we:properties><we:bindings/>"
                       f'<we:snapshot xmlns:r="{R_NS}"/></we:webextension>').encode()
        ct = add_override(ct, "/" + ours, CT_WE)
        done = "pinned"
    target = ours.rsplit("/", 1)[1]
    rid = re.search(rf'Id="(rId\d+)"[^>]*Target="{re.escape(target)}"', tp_rels) or \
        re.search(rf'Target="{re.escape(target)}"[^>]*Id="(rId\d+)"', tp_rels)
    if rid:
        rid = rid.group(1)
    else:
        rid = next_rid(tp_rels)
        tp_rels = tp_rels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="{REL_WE}" Target="{target}"/></Relationships>')
    pane = re.search(rf'<wetp:taskpane\b[^>]*>(?:(?!</wetp:taskpane>).)*r:id="{rid}"(?:(?!</wetp:taskpane>).)*</wetp:taskpane>', tp, re.S)
    if pane:
        opening = re.match(r"<wetp:taskpane\b[^>]*>", pane.group(0)).group(0)
        fixed = re.sub(r'visibility="\d"', 'visibility="1"', opening) if "visibility=" in opening \
            else opening.replace("<wetp:taskpane", '<wetp:taskpane visibility="1"', 1)
        tp = tp.replace(opening + pane.group(0)[len(opening):], fixed + pane.group(0)[len(opening):], 1)
    else:
        tp = tp.replace("</wetp:taskpanes>", f'<wetp:taskpane dockstate="right" visibility="1" width="350" row="0">'
                                             f'<wetp:webextensionref xmlns:r="{R_NS}" r:id="{rid}"/></wetp:taskpane></wetp:taskpanes>')
    parts.update({"[Content_Types].xml": ct.encode(), "_rels/.rels": root_rels.encode(),
                  tp_name: tp.encode(), tp_rels_name: tp_rels.encode()})
    return done


def main(decks):
    version = manifest_version()
    BACKUPS.mkdir(parents=True, exist_ok=True)
    for deck in map(Path, decks):
        if open_in_powerpoint(deck):
            print(f"✋ {deck.name}: open in PowerPoint — close it first (it would overwrite this on save)")
            continue
        with zipfile.ZipFile(deck) as z:
            infos = z.infolist()
            parts = {i.filename: z.read(i.filename) for i in infos}
        done = pin(parts, version)
        backup = BACKUPS / f"{deck.stem}-{time.strftime('%Y%m%d-%H%M%S')}{deck.suffix}"
        shutil.copy2(deck, backup)
        tmp = deck.with_name(f".{deck.name}.pinning")
        order = [i.filename for i in infos] + [n for n in parts if n not in {i.filename for i in infos}]
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for name in order:   # [Content_Types].xml stays first, as Office writes it
                z.writestr(name, parts[name])
        tmp.replace(deck)
        print(f"📌 {deck.name}: {done} (backup: {backup})")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    main(sys.argv[1:])
