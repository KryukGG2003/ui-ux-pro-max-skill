#!/usr/bin/env python3
"""
Дописывает в готовый .pptx то, чего не умеет pptxgenjs:

0. Чистка дублей <a:pPr>: pptxgenjs пишет по одному на каждый run внутри
   абзаца, а схема допускает только один. Они идентичны, лишние удаляются.
1. Дедупликация фоновых картинок (pptxgenjs кладёт свою копию на каждый слайд).
2. Переход между слайдами  — <p:transition>.
3. Анимации появления       — <p:timing> с каскадом по элементам слайда.

Элементы группируются по «волнам»: всё, что лежит внутри карточки,
появляется вместе с ней, а не по отдельности.

Запуск:  python3 add-animations.py struktura-konflikta.pptx
"""
import hashlib
import re
import shutil
import sys
import zipfile
from collections import OrderedDict

EMU_IN = 914400
FADE_MS = 480          # длительность проявления одного элемента
STAGGER_MS = 130       # задержка между волнами
CONTAINER_MIN_AREA = 1.6   # кв. дюймы: с этого размера фигура считается карточкой


# --------------------------------------------------------------------------
# Разбор фигур слайда
# --------------------------------------------------------------------------
def parse_shapes(xml):
    """Возвращает список фигур в порядке документа: id, геометрия, тип."""
    shapes = []
    # p:sp — фигуры и надписи; p:pic — картинки
    for m in re.finditer(r"<p:(sp|pic)>(.*?)</p:\1>", xml, re.S):
        block = m.group(2)
        mid = re.search(r'<p:cNvPr id="(\d+)"', block)
        if not mid:
            continue
        sid = int(mid.group(1))
        off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"', block)
        ext = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"', block)
        if not off or not ext:
            continue
        x, y = int(off.group(1)) / EMU_IN, int(off.group(2)) / EMU_IN
        cx, cy = int(ext.group(1)) / EMU_IN, int(ext.group(2)) / EMU_IN
        prst = re.search(r'<a:prstGeom prst="([^"]+)"', block)
        shapes.append({
            "id": sid, "x": x, "y": y, "w": cx, "h": cy,
            "prst": prst.group(1) if prst else "",
            "has_text": "<a:t>" in block,
        })
    return shapes


def assign_waves(shapes):
    """Фигуры внутри карточки получают её номер волны."""
    containers = []          # (x, y, w, h, wave)
    waves = []
    next_wave = 0
    for sh in shapes:
        cxc, cyc = sh["x"] + sh["w"] / 2, sh["y"] + sh["h"] / 2
        found = None
        for (bx, by, bw, bh, wv) in containers:
            if bx - 0.02 <= cxc <= bx + bw + 0.02 and by - 0.02 <= cyc <= by + bh + 0.02:
                found = wv
                break
        if found is None:
            found = next_wave
            next_wave += 1
            if sh["w"] * sh["h"] >= CONTAINER_MIN_AREA and sh["prst"] in ("roundRect", "rect"):
                containers.append((sh["x"], sh["y"], sh["w"], sh["h"], found))
        waves.append(found)
    return waves


# --------------------------------------------------------------------------
# Сборка <p:timing>
# --------------------------------------------------------------------------
def build_timing(shapes, waves):
    if not shapes:
        return ""

    groups = OrderedDict()
    for sh, wv in zip(shapes, waves):
        groups.setdefault(wv, []).append(sh["id"])

    uid = [4]                      # id 1..3 заняты каркасом
    def nid():
        uid[0] += 1
        return uid[0]

    effects = []
    for order, (_, ids) in enumerate(groups.items()):
        delay = order * STAGGER_MS
        for spid in ids:
            eff_id, set_id, anim_id = nid(), nid(), nid()
            effects.append(
                f'<p:par><p:cTn id="{eff_id}" presetID="10" presetClass="entr" presetSubtype="0"'
                f' fill="hold" grpId="0" nodeType="withEffect">'
                f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>'
                f'<p:childTnLst>'
                f'<p:set><p:cBhvr>'
                f'<p:cTn id="{set_id}" dur="1" fill="hold">'
                f'<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
                f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
                f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
                f'</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
                f'<p:animEffect transition="in" filter="fade"><p:cBhvr>'
                f'<p:cTn id="{anim_id}" dur="{FADE_MS}"/>'
                f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
                f'</p:cBhvr></p:animEffect>'
                f'</p:childTnLst></p:cTn></p:par>'
            )

    return (
        '<p:timing><p:tnLst><p:par>'
        '<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
        '<p:seq concurrent="1" nextAc="seek">'
        '<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
        '<p:par><p:cTn id="3" fill="hold">'
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst>'     # старт без клика
        '<p:childTnLst><p:par><p:cTn id="4" fill="hold">'
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
        '<p:childTnLst>' + "".join(effects) + '</p:childTnLst>'
        '</p:cTn></p:par></p:childTnLst>'
        '</p:cTn></p:par>'
        '</p:childTnLst></p:cTn>'
        '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
        '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
    )


TRANSITION = '<p:transition spd="med" advClick="1"><p:fade/></p:transition>'

PPR_RE = re.compile(r"<a:pPr[^>]*/>|<a:pPr[^>]*>.*?</a:pPr>", re.S)


def dedupe_ppr(xml):
    """В одном <a:p> допустим только один <a:pPr>; pptxgenjs дублирует его на каждый run."""
    def fix_para(m):
        body = m.group(1)
        found = PPR_RE.findall(body)
        if len(found) < 2:
            return m.group(0)
        seen = [False]

        def keep_first(mm):
            if not seen[0]:
                seen[0] = True
                return mm.group(0)
            return ""

        return "<a:p>" + PPR_RE.sub(keep_first, body) + "</a:p>"

    return re.sub(r"<a:p>(.*?)</a:p>", fix_para, xml, flags=re.S)


# --------------------------------------------------------------------------
def main(path):
    src = zipfile.ZipFile(path)
    names = src.namelist()
    items = {n: src.read(n) for n in names}
    src.close()

    # ---- 1. дедупликация media ----
    media = {n: v for n, v in items.items() if n.startswith("ppt/media/") and v}
    canon = {}
    remap = {}
    for n in sorted(media):
        h = hashlib.sha1(media[n]).hexdigest()
        if h in canon:
            remap[n.split("/")[-1]] = canon[h].split("/")[-1]
            del items[n]
        else:
            canon[h] = n
    if remap:
        for n in list(items):
            if n.startswith("ppt/slides/_rels/") or n.endswith(".rels"):
                t = items[n].decode("utf-8")
                for old, new in remap.items():
                    t = t.replace(f'"../media/{old}"', f'"../media/{new}"')
                items[n] = t.encode("utf-8")

    # ---- 2. переходы и анимации ----
    animated = 0
    for n in sorted(items):
        if not re.fullmatch(r"ppt/slides/slide\d+\.xml", n):
            continue
        xml = dedupe_ppr(items[n].decode("utf-8"))
        if "<p:timing>" in xml:
            items[n] = xml.encode("utf-8")
            continue
        shapes = parse_shapes(xml)
        waves = assign_waves(shapes)
        timing = build_timing(shapes, waves)
        xml = xml.replace("</p:sld>", TRANSITION + timing + "</p:sld>")
        items[n] = xml.encode("utf-8")
        animated += 1

    # ---- 3. перезапись архива ----
    shutil.copy(path, path + ".bak")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for n in names:
            if n in items:
                z.writestr(n, items[n])

    print(f"слайдов обработано: {animated}")
    print(f"дублей media убрано: {len(remap)}")


if __name__ == "__main__":
    main(sys.argv[1])
