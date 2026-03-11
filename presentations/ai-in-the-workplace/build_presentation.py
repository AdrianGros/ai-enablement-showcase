from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from textwrap import dedent, wrap
from zipfile import ZIP_DEFLATED, ZipFile


EMU = 914400
SLIDE_W = 12192000
SLIDE_H = 6858000


@dataclass
class SlideSpec:
    title: str
    kind: str
    kicker: str | None
    statement: str | None
    items: list[str]


SLIDES = [
    SlideSpec(
        title="Kuenstliche Intelligenz im Arbeitsalltag",
        kind="title",
        kicker="AI Enablement Showcase",
        statement="Moeglichkeiten, Grenzen und verantwortliche Nutzung",
        items=[
            "KI ist kein Zukunftsthema mehr.",
            "Die entscheidende Frage ist: Wie nutzen wir sie sinnvoll?",
        ],
    ),
    SlideSpec(
        title="Warum das Thema jetzt wichtig ist",
        kind="statement",
        kicker="Relevanz",
        statement="Generative KI wird fuer Entwuerfe, Zusammenfassungen und Informationsarbeit genutzt.",
        items=[
            "Ohne Methode sinkt die Qualitaet schnell.",
            "Deshalb braucht es Nutzungskompetenz, nicht nur Toolzugang.",
        ],
    ),
    SlideSpec(
        title="Was moderne KI heute gut kann",
        kind="cards",
        kicker="Nutzen",
        statement="Aktuelle generative KI-Systeme eignen sich fuer Entwuerfe, Zusammenfassungen, Umformulierungen und erste Ideen in textnaher Arbeit.",
        items=[
            "Texte entwerfen",
            "Inhalte zusammenfassen",
            "Informationen umformulieren",
            "Ideen strukturieren",
            "Erste Denkansaetze liefern",
        ],
    ),
    SlideSpec(
        title="Was moderne KI nicht gut kann",
        kind="cards",
        kicker="Grenzen",
        statement="Generative KI kann flussig klingende, aber unzutreffende Inhalte erzeugen. Ergebnisse sollten deshalb nicht ungeprueft uebernommen werden.",
        items=[
            "Korrektheit nicht garantieren",
            "Fehlenden Kontext kaum erraten",
            "Fachurteile nicht ersetzen",
            "Verantwortung nicht uebernehmen",
        ],
    ),
    SlideSpec(
        title="Wie moderne KI grob funktioniert",
        kind="process",
        kicker="Funktionsprinzip",
        statement="Grosse Sprachmodelle erzeugen Text, indem sie auf Basis gelernter Muster wahrscheinliche Tokenfolgen vorhersagen.",
        items=[
            "Trainiert auf vielen Beispielen",
            "Erkennt Muster in Sprache und Struktur",
            "Erzeugt Antworten ueber Wahrscheinlichkeiten",
            "Qualitaet haengt stark von Eingabe und Kontext ab",
        ],
    ),
    SlideSpec(
        title="Wo KI im Arbeitsalltag hilft",
        kind="cards",
        kicker="Praxis",
        statement="Ein sinnvoller Einstieg liegt in Unterstuetzungsaufgaben wie Entwuerfen, Zusammenfassungen, Vorbereitung und Strukturierung in der Wissensarbeit.",
        items=[
            "Vorbereitung von Meetings",
            "Zusammenfassungen und Notizen",
            "Erste Textentwuerfe",
            "Strukturierung von Informationen",
            "Recherche und Kommunikation",
        ],
    ),
    SlideSpec(
        title="Gute Nutzung ist eine Methode",
        kind="steps",
        kicker="Arbeitsmethode",
        statement="Expliziter Kontext, Rolle, Ziel, Beispiele und Review verbessern die Ergebnisqualitaet oft deutlich.",
        items=[
            "Kontext",
            "Rolle",
            "Ziel",
            "Beispiele",
            "Pruefen",
        ],
    ),
    SlideSpec(
        title="Was in der Praxis schieflaeuft",
        kind="cards",
        kicker="Fehlerquellen",
        statement="Schwache Anweisungen und fehlende Pruefung erhoehen die Wahrscheinlichkeit fuer generische, irrelevante oder falsche Ergebnisse.",
        items=[
            "Zu vage Anfragen",
            "Kein fachlicher Kontext",
            "Keine Beispiele",
            "Keine Qualitaetskontrolle",
            "Blinde Uebernahme von Ergebnissen",
        ],
    ),
    SlideSpec(
        title="Warum Agenten eine andere Risikoklasse sind",
        kind="ladder",
        kicker="Risikologik",
        statement="KI-Systeme, die Tools oder externe Aktionen ausloesen koennen, brauchen staerkere Schutzmechanismen, Aufsicht und Grenzen als reine Chat-Nutzung.",
        items=[
            "Chatbots antworten",
            "Agenten koennen handeln",
            "Mehr Wirkung bedeutet mehr Kontrollbedarf",
            "Mehr Autonomie braucht Leitplanken",
        ],
    ),
    SlideSpec(
        title="Leitplanken fuer verantwortliche Nutzung",
        kind="cards",
        kicker="Guardrails",
        statement="Verantwortliche Nutzung profitiert von klaren Grenzen, menschlicher Pruefung und Regeln fuer Aufgaben mit hoeherer Wirkung.",
        items=[
            "Keine sensiblen Daten eingeben",
            "Wichtige Ergebnisse pruefen",
            "Klare Regeln fuer Nutzung",
            "Mehr Kontrolle bei hoeherem Risiko",
        ],
    ),
    SlideSpec(
        title="Drei Regeln fuer den Alltag",
        kind="rules",
        kicker="Merksatz",
        statement="Nuetzliche KI-Nutzung im Alltag braucht Kontext, Verifikation und menschliche Verantwortung.",
        items=[
            "KI ist Werkzeug, nicht Wahrheit",
            "Qualitaet braucht Kontext",
            "Verantwortung bleibt beim Menschen",
        ],
    ),
    SlideSpec(
        title="Fazit und Fragen",
        kind="closing",
        kicker="Abschluss",
        statement="KI ist am nuetzlichsten, wenn sie mit klaren Prompts, Pruefung und passenden Guardrails eingesetzt wird.",
        items=[
            "Klare Prompts",
            "Pruefung",
            "Angemessene Guardrails",
            "Fragen?",
        ],
    ),
]


BG = "F7F4EE"
INK = "1E293B"
MUTED = "667085"
ACCENT = "0F766E"
ACCENT_2 = "0B3B45"
CARD = "FFFFFF"
WARM = "E7D8C9"
CORAL = "D97757"
LINE = "D6D3D1"


def inches(value: float) -> int:
    return int(value * EMU)


def escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def xml_textbox(shape_id: int, name: str, x: float, y: float, w: float, h: float, paragraphs: list[tuple[str, int, str, bool]], fill: str | None = None, line: str | None = None, radius: str = "rect", inset: int = 0) -> str:
    fill_xml = f"<a:solidFill><a:srgbClr val=\"{fill}\"/></a:solidFill>" if fill else "<a:noFill/>"
    line_xml = f"<a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"{line}\"/></a:solidFill></a:ln>" if line else "<a:ln><a:noFill/></a:ln>"
    paras = []
    for text, size, color, bold in paragraphs:
        weight = ' b="1"' if bold else ""
        paras.append(
            f"<a:p><a:r><a:rPr lang=\"en-US\" sz=\"{size}\"{weight}><a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill></a:rPr><a:t>{escape(text)}</a:t></a:r><a:endParaRPr lang=\"en-US\" sz=\"{size}\"/></a:p>"
        )
    return (
        f"<p:sp>"
        f"<p:nvSpPr><p:cNvPr id=\"{shape_id}\" name=\"{escape(name)}\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr>"
        f"<p:spPr><a:xfrm><a:off x=\"{inches(x)}\" y=\"{inches(y)}\"/><a:ext cx=\"{inches(w)}\" cy=\"{inches(h)}\"/></a:xfrm>"
        f"<a:prstGeom prst=\"{radius}\"><a:avLst/></a:prstGeom>{fill_xml}{line_xml}</p:spPr>"
        f"<p:txBody><a:bodyPr wrap=\"square\" lIns=\"{inset}\" tIns=\"{inset}\" rIns=\"{inset}\" bIns=\"{inset}\"/><a:lstStyle/>"
        + "".join(paras)
        + "</p:txBody></p:sp>"
    )


def xml_rect(shape_id: int, name: str, x: float, y: float, w: float, h: float, fill: str, radius: str = "rect") -> str:
    return (
        f"<p:sp>"
        f"<p:nvSpPr><p:cNvPr id=\"{shape_id}\" name=\"{escape(name)}\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"
        f"<p:spPr><a:xfrm><a:off x=\"{inches(x)}\" y=\"{inches(y)}\"/><a:ext cx=\"{inches(w)}\" cy=\"{inches(h)}\"/></a:xfrm>"
        f"<a:prstGeom prst=\"{radius}\"><a:avLst/></a:prstGeom>"
        f"<a:solidFill><a:srgbClr val=\"{fill}\"/></a:solidFill>"
        f"<a:ln><a:noFill/></a:ln></p:spPr>"
        f"<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"
    )


def slide_shell(shapes: list[str]) -> str:
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
        "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"
        "<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val=\"" + BG + "\"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>"
        "<p:spTree>"
        "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"
        "<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>"
        + "".join(shapes)
        + "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
    )


def common_chrome(title: str, kicker: str | None, statement: str | None) -> list[str]:
    shapes = [
        xml_rect(2, "Top Bar", 0, 0, 13.333, 0.28, ACCENT),
    ]
    if kicker:
        shapes.append(
            xml_textbox(
                3,
                "Kicker",
                0.62,
                0.45,
                3.2,
                0.32,
                [(kicker.upper(), 1200, ACCENT, True)],
            )
        )
    shapes.append(
        xml_textbox(
            4,
            "Title",
            0.62,
            0.82,
            11.7,
            0.8,
            [(title, 2400, INK, True)],
        )
    )
    if statement:
        wrapped = wrap(statement, width=62)
        if len(wrapped) > 2:
            wrapped = [wrapped[0], " ".join(wrapped[1:])]
        shapes.append(xml_rect(5, "Statement Box", 0.62, 1.5, 11.8, 1.02, "FDF8F3", "roundRect"))
        shapes.append(
            xml_textbox(
                6,
                "Statement",
                0.9,
                1.77,
                11.0,
                0.62,
                [(line, 1040, MUTED, False) for line in wrapped],
            )
        )
    return shapes


def build_title_slide(spec: SlideSpec) -> str:
    shapes = [
        xml_rect(2, "Top Bar", 0, 0, 13.333, 0.42, ACCENT),
        xml_rect(3, "Accent Panel", 0.62, 1.55, 5.1, 3.75, ACCENT_2, "roundRect"),
        xml_rect(4, "Warm Panel", 6.08, 1.55, 6.62, 3.75, WARM, "roundRect"),
        xml_textbox(5, "Kicker", 0.84, 1.88, 3.0, 0.4, [(spec.kicker or "", 1100, "D9F2EF", True)]),
        xml_textbox(6, "Title", 0.84, 2.35, 4.45, 1.55, [(spec.title, 2600, "FFFFFF", True)]),
        xml_textbox(7, "Subtitle", 6.48, 2.05, 5.4, 0.9, [(spec.statement or "", 1800, INK, True)]),
        xml_textbox(
            8,
            "Prompt",
            6.48,
            3.18,
            5.25,
            1.3,
            [(item, 1160, INK, False) for item in spec.items],
        ),
    ]
    return slide_shell(shapes)


def build_statement_slide(spec: SlideSpec) -> str:
    shapes = common_chrome(spec.title, spec.kicker, None)
    shapes.append(xml_rect(6, "Main Statement Box", 0.62, 1.7, 12.05, 1.25, CARD, "roundRect"))
    shapes.append(
        xml_textbox(
            7,
            "Main Statement",
            0.92,
            2.0,
            11.5,
            0.7,
            [(spec.statement or "", 1800, INK, True)],
        )
    )
    x_positions = [0.62, 4.32, 8.02]
    for idx, item in enumerate(spec.items):
        shapes.append(xml_rect(8 + idx * 2, f"Card {idx + 1}", x_positions[idx], 3.35, 3.45, 2.0, CARD, "roundRect"))
        shapes.append(
            xml_textbox(
                9 + idx * 2,
                f"Card Text {idx + 1}",
                x_positions[idx] + 0.28,
                3.85,
                2.9,
                1.05,
                [(item, 1120, INK, True)],
            )
        )
    return slide_shell(shapes)


def build_cards_slide(spec: SlideSpec) -> str:
    shapes = common_chrome(spec.title, spec.kicker, spec.statement)
    cards = spec.items
    positions = [
        (0.62, 2.55),
        (4.95, 2.55),
        (9.28, 2.55),
        (2.82, 4.45),
        (7.15, 4.45),
    ]
    for idx, (item, (x, y)) in enumerate(zip(cards, positions, strict=False)):
        fill = CARD if idx % 2 == 0 else "FDF8F3"
        shapes.append(xml_rect(10 + idx * 2, f"Card {idx + 1}", x, y, 3.45, 1.45, fill, "roundRect"))
        shapes.append(
            xml_textbox(
                11 + idx * 2,
                f"Card Label {idx + 1}",
                x + 0.28,
                y + 0.38,
                2.9,
                0.72,
                [(line, 1040, INK, True) for line in wrap(item, width=19)],
            )
        )
    return slide_shell(shapes)


def build_process_slide(spec: SlideSpec) -> str:
    shapes = common_chrome(spec.title, spec.kicker, spec.statement)
    xs = [0.62, 3.65, 6.68, 9.71]
    labels = ["1", "2", "3", "4"]
    colors = [ACCENT, CORAL, ACCENT, CORAL]
    for idx, item in enumerate(spec.items):
        shapes.append(xml_rect(10 + idx * 3, f"Step Box {idx + 1}", xs[idx], 3.2, 2.6, 2.2, CARD, "roundRect"))
        shapes.append(xml_rect(11 + idx * 3, f"Step Header {idx + 1}", xs[idx], 3.2, 2.6, 0.42, colors[idx]))
        shapes.append(
            xml_textbox(
                12 + idx * 3,
                f"Step Text {idx + 1}",
                xs[idx] + 0.24,
                3.85,
                2.1,
                1.1,
                [(labels[idx], 1800, colors[idx], True), (item, 1040, INK, True)],
            )
        )
    return slide_shell(shapes)


def build_steps_slide(spec: SlideSpec) -> str:
    shapes = common_chrome(spec.title, spec.kicker, spec.statement)
    widths = [1.8, 1.8, 1.8, 1.8, 1.8]
    x = 0.62
    colors = [ACCENT, CORAL, ACCENT, CORAL, ACCENT]
    for idx, item in enumerate(spec.items):
        shapes.append(xml_rect(20 + idx * 2, f"Step {idx + 1}", x, 3.45, widths[idx], 2.0, colors[idx], "roundRect"))
        shapes.append(
            xml_textbox(
                21 + idx * 2,
                f"Step Label {idx + 1}",
                x + 0.18,
                3.95,
                1.44,
                1.0,
                [(str(idx + 1), 1500, "FFFFFF", True), (item, 1120, "FFFFFF", True)],
            )
        )
        x += 2.35
    return slide_shell(shapes)


def build_ladder_slide(spec: SlideSpec) -> str:
    shapes = common_chrome(spec.title, spec.kicker, spec.statement)
    sizes = [(0.62, 4.75, 4.25), (1.45, 3.65, 5.9), (2.28, 2.55, 7.55), (3.11, 1.45, 9.2)]
    colors = [CARD, "FFF2EB", "F4FBFA", "E6F4F1"]
    for idx, item in enumerate(spec.items):
        x, y, w = sizes[idx]
        shapes.append(xml_rect(30 + idx * 2, f"Ladder {idx + 1}", x, y, w, 0.85, colors[idx], "roundRect"))
        shapes.append(
            xml_textbox(
                31 + idx * 2,
                f"Ladder Label {idx + 1}",
                x + 0.28,
                y + 0.2,
                w - 0.5,
                0.4,
                [(item, 1100, INK, True)],
            )
        )
    return slide_shell(shapes)


def build_rules_slide(spec: SlideSpec) -> str:
    shapes = common_chrome(spec.title, spec.kicker, spec.statement)
    boxes = [(0.95, ACCENT), (4.45, CORAL), (7.95, ACCENT_2)]
    for idx, (item, (x, color)) in enumerate(zip(spec.items, boxes, strict=False)):
        shapes.append(xml_rect(40 + idx * 2, f"Rule {idx + 1}", x, 3.1, 2.95, 2.45, color, "roundRect"))
        shapes.append(
            xml_textbox(
                41 + idx * 2,
                f"Rule Text {idx + 1}",
                x + 0.24,
                3.82,
                2.45,
                1.15,
                [(item, 1160, "FFFFFF", True)],
            )
        )
    return slide_shell(shapes)


def build_closing_slide(spec: SlideSpec) -> str:
    shapes = common_chrome(spec.title, spec.kicker, None)
    shapes.append(xml_rect(50, "Summary Panel", 0.62, 1.75, 8.65, 3.95, ACCENT_2, "roundRect"))
    shapes.append(
        xml_textbox(
            51,
            "Summary Statement",
            0.98,
            2.35,
            7.8,
            1.35,
            [(spec.statement or "", 1800, "FFFFFF", True)],
        )
    )
    x = 9.58
    for idx, item in enumerate(spec.items[:3]):
        shapes.append(xml_rect(52 + idx * 2, f"Close Card {idx + 1}", x, 2.05 + idx * 1.1, 2.45, 0.78, WARM if idx != 1 else "F9E0D4", "roundRect"))
        shapes.append(
            xml_textbox(
                53 + idx * 2,
                f"Close Text {idx + 1}",
                x + 0.18,
                2.28 + idx * 1.1,
                2.0,
                0.3,
                [(item, 980, INK, True)],
            )
        )
    shapes.append(
        xml_textbox(
            60,
            "Question",
            9.58,
            5.45,
            2.45,
            0.55,
            [("Fragen?", 1400, ACCENT, True)],
        )
    )
    return slide_shell(shapes)


def build_slide(spec: SlideSpec) -> str:
    if spec.kind == "title":
        return build_title_slide(spec)
    if spec.kind == "statement":
        return build_statement_slide(spec)
    if spec.kind == "process":
        return build_process_slide(spec)
    if spec.kind == "steps":
        return build_steps_slide(spec)
    if spec.kind == "ladder":
        return build_ladder_slide(spec)
    if spec.kind == "rules":
        return build_rules_slide(spec)
    if spec.kind == "closing":
        return build_closing_slide(spec)
    return build_cards_slide(spec)


def slide_rel() -> str:
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/>"
        "</Relationships>"
    )


def write_pptx(out_path: Path) -> None:
    created = datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    with ZipFile(out_path, "w", compression=ZIP_DEFLATED) as zf:
        zf.writestr(
            "[Content_Types].xml",
            content_types_xml(),
        )
        zf.writestr("_rels/.rels", root_rels_xml())
        zf.writestr("docProps/core.xml", core_xml(created))
        zf.writestr("docProps/app.xml", app_xml())
        zf.writestr("ppt/presentation.xml", presentation_xml())
        zf.writestr("ppt/_rels/presentation.xml.rels", presentation_rels_xml())
        zf.writestr("ppt/slideMasters/slideMaster1.xml", slide_master_xml())
        zf.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", slide_master_rels_xml())
        zf.writestr("ppt/slideLayouts/slideLayout1.xml", slide_layout_xml())
        zf.writestr("ppt/theme/theme1.xml", theme_xml())
        zf.writestr("ppt/presProps.xml", pres_props_xml())
        zf.writestr("ppt/viewProps.xml", view_props_xml())
        zf.writestr("ppt/tableStyles.xml", table_styles_xml())
        for idx, slide in enumerate(SLIDES, start=1):
            zf.writestr(f"ppt/slides/slide{idx}.xml", build_slide(slide))
            zf.writestr(f"ppt/slides/_rels/slide{idx}.xml.rels", slide_rel())


def content_types_xml() -> str:
    overrides = [
        "<Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/>",
        "<Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml\"/>",
        "<Override PartName=\"/ppt/slideLayouts/slideLayout1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"/>",
        "<Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>",
        "<Override PartName=\"/ppt/presProps.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presProps+xml\"/>",
        "<Override PartName=\"/ppt/viewProps.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml\"/>",
        "<Override PartName=\"/ppt/tableStyles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml\"/>",
        "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>",
        "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>",
    ]
    for idx in range(1, len(SLIDES) + 1):
        overrides.append(
            f"<Override PartName=\"/ppt/slides/slide{idx}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>"
        )
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        + "".join(overrides)
        + "</Types>"
    )


def root_rels_xml() -> str:
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"ppt/presentation.xml\"/>"
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>"
        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>"
        "</Relationships>"
    )


def core_xml(created: str) -> str:
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" "
        "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" "
        "xmlns:dcterms=\"http://purl.org/dc/terms/\" "
        "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" "
        "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"
        "<dc:title>AI Enablement Showcase Deck</dc:title>"
        "<dc:creator>Codex</dc:creator>"
        "<cp:lastModifiedBy>Codex</cp:lastModifiedBy>"
        "<dcterms:created xsi:type=\"dcterms:W3CDTF\">" + created + "</dcterms:created>"
        "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">" + created + "</dcterms:modified>"
        "</cp:coreProperties>"
    )


def app_xml() -> str:
    return dedent(
        """\
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
                    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
          <Application>Codex PPTX Builder</Application>
          <Slides>12</Slides>
          <Notes>0</Notes>
          <HiddenSlides>0</HiddenSlides>
          <MMClips>0</MMClips>
          <ScaleCrop>false</ScaleCrop>
          <HeadingPairs>
            <vt:vector size="2" baseType="variant">
              <vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant>
              <vt:variant><vt:i4>1</vt:i4></vt:variant>
            </vt:vector>
          </HeadingPairs>
          <TitlesOfParts>
            <vt:vector size="1" baseType="lpstr">
              <vt:lpstr>Office Theme</vt:lpstr>
            </vt:vector>
          </TitlesOfParts>
          <Company></Company>
          <LinksUpToDate>false</LinksUpToDate>
          <SharedDoc>false</SharedDoc>
          <HyperlinksChanged>false</HyperlinksChanged>
          <AppVersion>16.0000</AppVersion>
        </Properties>
        """
    )


def presentation_xml() -> str:
    sld_ids = "".join(
        f"<p:sldId id=\"{255 + idx}\" r:id=\"rId{idx + 1}\"/>" for idx in range(1, len(SLIDES) + 1)
    )
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<p:presentation xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
        "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"
        "<p:sldMasterIdLst><p:sldMasterId id=\"2147483648\" r:id=\"rId13\"/></p:sldMasterIdLst>"
        f"<p:sldIdLst>{sld_ids}</p:sldIdLst>"
        f"<p:sldSz cx=\"{SLIDE_W}\" cy=\"{SLIDE_H}\"/>"
        "<p:notesSz cx=\"6858000\" cy=\"9144000\"/>"
        "<p:defaultTextStyle/>"
        "</p:presentation>"
    )


def presentation_rels_xml() -> str:
    rels = []
    for idx in range(1, len(SLIDES) + 1):
        rels.append(
            f"<Relationship Id=\"rId{idx}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide{idx}.xml\"/>"
        )
    rels.append(
        "<Relationship Id=\"rId13\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"slideMasters/slideMaster1.xml\"/>"
    )
    rels.append(
        "<Relationship Id=\"rId14\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps\" Target=\"presProps.xml\"/>"
    )
    rels.append(
        "<Relationship Id=\"rId15\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps\" Target=\"viewProps.xml\"/>"
    )
    rels.append(
        "<Relationship Id=\"rId16\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles\" Target=\"tableStyles.xml\"/>"
    )
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        + "".join(rels)
        + "</Relationships>"
    )


def slide_master_xml() -> str:
    return dedent(
        f"""\
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                     xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cSld name="AI Enablement Master">
            <p:bg>
              <p:bgPr>
                <a:solidFill><a:srgbClr val="{BG}"/></a:solidFill>
                <a:effectLst/>
              </p:bgPr>
            </p:bg>
            <p:spTree>
              <p:nvGrpSpPr>
                <p:cNvPr id="1" name=""/>
                <p:cNvGrpSpPr/>
                <p:nvPr/>
              </p:nvGrpSpPr>
              <p:grpSpPr>
                <a:xfrm>
                  <a:off x="0" y="0"/>
                  <a:ext cx="0" cy="0"/>
                  <a:chOff x="0" y="0"/>
                  <a:chExt cx="0" cy="0"/>
                </a:xfrm>
              </p:grpSpPr>
            </p:spTree>
          </p:cSld>
          <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
          <p:sldLayoutIdLst>
            <p:sldLayoutId id="1" r:id="rId1"/>
          </p:sldLayoutIdLst>
          <p:txStyles>
            <p:titleStyle/>
            <p:bodyStyle/>
            <p:otherStyle/>
          </p:txStyles>
        </p:sldMaster>
        """
    )


def slide_master_rels_xml() -> str:
    return dedent(
        """\
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
        </Relationships>
        """
    )


def slide_layout_xml() -> str:
    return dedent(
        """\
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                     xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                     type="blank" preserve="1">
          <p:cSld name="Blank Layout">
            <p:spTree>
              <p:nvGrpSpPr>
                <p:cNvPr id="1" name=""/>
                <p:cNvGrpSpPr/>
                <p:nvPr/>
              </p:nvGrpSpPr>
              <p:grpSpPr>
                <a:xfrm>
                  <a:off x="0" y="0"/>
                  <a:ext cx="0" cy="0"/>
                  <a:chOff x="0" y="0"/>
                  <a:chExt cx="0" cy="0"/>
                </a:xfrm>
              </p:grpSpPr>
            </p:spTree>
          </p:cSld>
          <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
        </p:sldLayout>
        """
    )


def theme_xml() -> str:
    return dedent(
        """\
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="AI Enablement Theme">
          <a:themeElements>
            <a:clrScheme name="AI Enablement Colors">
              <a:dk1><a:srgbClr val="1E293B"/></a:dk1>
              <a:lt1><a:srgbClr val="F7F4EE"/></a:lt1>
              <a:dk2><a:srgbClr val="0B3B45"/></a:dk2>
              <a:lt2><a:srgbClr val="FFFFFF"/></a:lt2>
              <a:accent1><a:srgbClr val="0F766E"/></a:accent1>
              <a:accent2><a:srgbClr val="D97757"/></a:accent2>
              <a:accent3><a:srgbClr val="E7D8C9"/></a:accent3>
              <a:accent4><a:srgbClr val="667085"/></a:accent4>
              <a:accent5><a:srgbClr val="D6D3D1"/></a:accent5>
              <a:accent6><a:srgbClr val="FDF8F3"/></a:accent6>
              <a:hlink><a:srgbClr val="0F766E"/></a:hlink>
              <a:folHlink><a:srgbClr val="667085"/></a:folHlink>
            </a:clrScheme>
            <a:fontScheme name="AI Enablement Fonts">
              <a:majorFont>
                <a:latin typeface="Aptos Display"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
              </a:majorFont>
              <a:minorFont>
                <a:latin typeface="Aptos"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
              </a:minorFont>
            </a:fontScheme>
            <a:fmtScheme name="Office">
              <a:fillStyleLst>
                <a:solidFill><a:schemeClr val="lt1"/></a:solidFill>
                <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
                <a:solidFill><a:schemeClr val="accent2"/></a:solidFill>
              </a:fillStyleLst>
              <a:lnStyleLst>
                <a:ln w="9525"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln>
                <a:ln w="25400"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:ln>
                <a:ln w="38100"><a:solidFill><a:schemeClr val="accent3"/></a:solidFill></a:ln>
              </a:lnStyleLst>
              <a:effectStyleLst>
                <a:effectStyle><a:effectLst/></a:effectStyle>
                <a:effectStyle><a:effectLst/></a:effectStyle>
                <a:effectStyle><a:effectLst/></a:effectStyle>
              </a:effectStyleLst>
              <a:bgFillStyleLst>
                <a:solidFill><a:schemeClr val="lt1"/></a:solidFill>
                <a:solidFill><a:schemeClr val="lt2"/></a:solidFill>
                <a:solidFill><a:schemeClr val="accent3"/></a:solidFill>
              </a:bgFillStyleLst>
            </a:fmtScheme>
          </a:themeElements>
          <a:objectDefaults/>
          <a:extraClrSchemeLst/>
        </a:theme>
        """
    )


def pres_props_xml() -> str:
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<p:presentationPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
        "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>"
    )


def view_props_xml() -> str:
    return dedent(
        """\
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:normalViewPr showOutlineIcons="0"/>
          <p:slideViewPr>
            <p:cSldViewPr snapToGrid="1" snapToObjects="1"/>
          </p:slideViewPr>
          <p:notesTextViewPr>
            <p:cViewPr/>
          </p:notesTextViewPr>
          <p:gridSpacing cx="72008" cy="72008"/>
        </p:viewPr>
        """
    )


def table_styles_xml() -> str:
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<a:tblStyleLst xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" def=\"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}\"/>"
    )


def main() -> None:
    out = Path(__file__).with_name("ai-in-the-workplace.pptx")
    write_pptx(out)
    print(out)


if __name__ == "__main__":
    main()
