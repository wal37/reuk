from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape
import zipfile


OUT = Path("/Users/wale/Desktop/resuk/rweb/RESEARCHAI_PITCH_DECK.pptx")

EMU = 914400
SLIDE_W = 13.333 * EMU
SLIDE_H = 7.5 * EMU


def e(value: str) -> str:
    return escape(value)


def pt(value: float) -> int:
    return int(value * 100)


def inch(value: float) -> int:
    return int(value * EMU)


class SlideBuilder:
    def __init__(self) -> None:
        self._shape_id = 1
        self.parts: list[str] = []
        self.images: list[Path] = []

    def _next_id(self) -> int:
        self._shape_id += 1
        return self._shape_id

    def rect(
        self,
        x: float,
        y: float,
        w: float,
        h: float,
        fill: str,
        *,
        line: str | None = None,
        rounded: bool = False,
        alpha: int | None = None,
    ) -> None:
        sid = self._next_id()
        line_xml = (
            f"<a:ln w='12700'><a:solidFill><a:srgbClr val='{line}'/></a:solidFill></a:ln>"
            if line
            else "<a:ln><a:noFill/></a:ln>"
        )
        fill_inner = (
            f"<a:srgbClr val='{fill}'><a:alpha val='{alpha}'/></a:srgbClr>"
            if alpha is not None
            else f"<a:srgbClr val='{fill}'/>"
        )
        geom = "roundRect" if rounded else "rect"
        self.parts.append(
            f"""
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{sid}" name="Shape {sid}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="{inch(x)}" y="{inch(y)}"/>
      <a:ext cx="{inch(w)}" cy="{inch(h)}"/>
    </a:xfrm>
    <a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>
    <a:solidFill>{fill_inner}</a:solidFill>
    {line_xml}
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p/>
  </p:txBody>
</p:sp>"""
        )

    def line(self, x1: float, y1: float, x2: float, y2: float, color: str, width_pt: float = 1.5) -> None:
        thickness = max(width_pt / 72, 0.02)
        x = min(x1, x2)
        y = min(y1, y2)
        w = max(abs(x2 - x1), thickness)
        h = max(abs(y2 - y1), thickness)
        self.rect(x, y, w, h, color)

    def text(
        self,
        x: float,
        y: float,
        w: float,
        h: float,
        text: str,
        *,
        size: float = 20,
        color: str = "1F2D1B",
        bold: bool = False,
        font: str = "Aptos",
        align: str = "l",
        valign: str = "t",
        all_caps: bool = False,
    ) -> None:
        sid = self._next_id()
        paragraphs = []
        for line in text.split("\n"):
            run_text = e(line)
            cap = " cap='all'" if all_caps else ""
            paragraphs.append(
                f"""
    <a:p>
      <a:pPr algn="{align}"/>
      <a:r>
        <a:rPr lang="en-US" sz="{pt(size)}" b="{'1' if bold else '0'}"{cap}>
          <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
          <a:latin typeface="{font}"/>
        </a:rPr>
        <a:t>{run_text}</a:t>
      </a:r>
      <a:endParaRPr lang="en-US" sz="{pt(size)}">
        <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
        <a:latin typeface="{font}"/>
      </a:endParaRPr>
    </a:p>"""
            )
        self.parts.append(
            f"""
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{sid}" name="Text {sid}"/>
    <p:cNvSpPr txBox="1"/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="{inch(x)}" y="{inch(y)}"/>
      <a:ext cx="{inch(w)}" cy="{inch(h)}"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="{valign}" wrap="square"/>
    <a:lstStyle/>
    {''.join(paragraphs)}
  </p:txBody>
</p:sp>"""
        )

    def image(self, x: float, y: float, w: float, h: float, path: str | Path, *, rounded: bool = False) -> None:
        sid = self._next_id()
        image_path = Path(path)
        self.images.append(image_path)
        rid = f"rId{len(self.images) + 1}"
        geom = "roundRect" if rounded else "rect"
        self.parts.append(
            f"""
<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="{sid}" name="Picture {sid}" descr="{e(image_path.name)}"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="{rid}"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>
      <a:off x="{inch(x)}" y="{inch(y)}"/>
      <a:ext cx="{inch(w)}" cy="{inch(h)}"/>
    </a:xfrm>
    <a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>"""
        )

    def slide_xml(self) -> str:
        return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
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
      {''.join(self.parts)}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>"""


def slide(
    bg: str,
    accent: str,
    title: str,
    subtitle: str,
    number: str | None = None,
    body: list[tuple[float, float, float, float, str, dict]] | None = None,
) -> tuple[str, list[Path]]:
    s = SlideBuilder()
    s.rect(0, 0, 13.333, 7.5, bg)
    s.rect(0.52, 0.42, 12.25, 0.08, accent)
    if number:
        s.text(0.72, 0.55, 1.1, 0.55, number, size=18, color=accent, bold=True, all_caps=True)
    s.text(0.72, 0.95, 5.8, 1.2, title, size=25, color="FFFFFF", bold=True, font="Georgia")
    s.text(0.72, 2.0, 6.0, 0.9, subtitle, size=12, color="E7EFE2")
    for block in body or []:
        x, y, w, h, txt, opts = block
        if opts.get("box"):
            s.rect(x, y, w, h, opts.get("fill", "F5F8F0"), line=opts.get("line"), rounded=opts.get("rounded", True))
        s.text(
            x + opts.get("tx", 0.16),
            y + opts.get("ty", 0.14),
            w - opts.get("tw", 0.32),
            h - opts.get("th", 0.28),
            txt,
            size=opts.get("size", 14),
            color=opts.get("color", "1F2D1B"),
            bold=opts.get("bold", False),
            font=opts.get("font", "Aptos"),
            align=opts.get("align", "l"),
        )
    return s.slide_xml(), s.images


def build_slides() -> list[tuple[str, list[Path]]]:
    slides: list[tuple[str, list[Path]]] = []
    ucl_logo = Path("/Users/wale/Desktop/resuk/rweb/ucllogo.jpeg")
    edin_logo = Path("/Users/wale/Desktop/resuk/rweb/edinlogo.png")

    # 1 cover
    s = SlideBuilder()
    s.rect(0, 0, 13.333, 7.5, "324826")
    s.rect(0.75, 0.75, 4.8, 6.0, "5D7442", rounded=True)
    s.rect(5.85, 0.75, 6.75, 6.0, "F5F8F0", rounded=True)
    s.rect(10.7, 1.15, 1.15, 1.15, "FFFFFF", rounded=True)
    s.image(10.92, 1.38, 0.72, 0.6, ucl_logo)
    s.rect(10.7, 2.55, 1.15, 1.15, "FFFFFF", rounded=True)
    s.image(10.92, 2.8, 0.72, 0.56, edin_logo)
    s.text(1.1, 1.1, 4.0, 0.5, "RESEARCHAI", size=18, color="EAF3DF", bold=True, all_caps=True)
    s.text(1.1, 1.8, 3.8, 2.1, "Participant research infrastructure for modern AI programs.", size=25, color="FFFFFF", bold=True, font="Georgia")
    s.text(1.1, 4.35, 3.7, 1.2, "A research operating layer for intake, screening, study flow, contributor memory, and repeat-ready cohorts.", size=12, color="EAF3DF")
    s.text(6.35, 1.2, 5.2, 0.45, "14-SLIDE COMPANY DECK", size=13, color="6D8751", bold=True, all_caps=True)
    s.text(6.35, 2.0, 5.0, 2.1, "ResearchAI brings participant systems, study operations, and AI-facing workflows into one structured environment.", size=23, color="1F2D1B", bold=True, font="Georgia")
    s.rect(6.35, 4.55, 4.1, 1.35, "E6EDDC", rounded=True)
    s.text(6.65, 4.88, 3.4, 0.7, "UK AI Sector Study 2025\n5,862 companies | £23.9B revenue", size=12.5, color="324826", bold=True)
    s.text(10.9, 4.95, 1.1, 0.45, "UCL", size=10.5, color="324826", bold=True, align="c")
    s.text(10.9, 3.95, 1.1, 0.45, "Edinburgh", size=10, color="324826", bold=True, align="c")
    s.text(6.35, 6.08, 5.1, 0.35, "Contact: info@researchai.co.uk", size=11.5, color="5F7059", bold=True)
    slides.append((s.slide_xml(), s.images))

    # 2 thesis
    slides.append(
        slide(
            "5D7442",
            "DCE8CF",
            "Research is operational work, not just project work.",
            "The platform is designed around the day-to-day mechanics of sourcing, qualifying, routing, reviewing, and retaining contributors for serious research programs.",
            "01",
            [
                (7.15, 1.0, 2.0, 2.0, "01\nIntake", {"box": True, "fill": "F5F8F0", "size": 20, "bold": True, "color": "324826"}),
                (9.4, 1.0, 2.0, 2.0, "02\nFit", {"box": True, "fill": "DCE8CF", "size": 20, "bold": True, "color": "324826"}),
                (7.15, 3.3, 2.0, 2.0, "03\nStudies", {"box": True, "fill": "E9F1E0", "size": 20, "bold": True, "color": "324826"}),
                (9.4, 3.3, 2.0, 2.0, "04\nMemory", {"box": True, "fill": "C9D8B5", "size": 20, "bold": True, "color": "324826"}),
            ],
        )
    )

    # 3 problem
    slides.append(
        slide(
            "F5F8F0",
            "6D8751",
            "Where participant research breaks down.",
            "Most teams still patch together sourcing, qualification, study execution, notes, and re-contact workflows across separate tools.",
            "02",
            [
                (0.8, 3.0, 2.8, 1.5, "Fragmented intake\nProfiles, forms, and study history live in different places.", {"box": True, "fill": "FFFFFF", "line": "DCE8CF", "size": 14, "bold": True}),
                (3.9, 3.0, 2.8, 1.5, "Weak screening\nFit decisions are inconsistent and difficult to audit later.", {"box": True, "fill": "FFFFFF", "line": "DCE8CF", "size": 14, "bold": True}),
                (7.0, 3.0, 2.8, 1.5, "Operational drift\nAssignments, notes, and payout readiness separate over time.", {"box": True, "fill": "FFFFFF", "line": "DCE8CF", "size": 14, "bold": True}),
                (10.1, 3.0, 2.4, 1.5, "Lost cohorts\nRepeat-ready contributors are hard to preserve and reuse.", {"box": True, "fill": "FFFFFF", "line": "DCE8CF", "size": 14, "bold": True}),
            ],
        )
    )

    # 4 solution
    slides.append(
        slide(
            "324826",
            "C9D8B5",
            "ResearchAI centralizes participant workflow.",
            "The product aligns contributor records, screening logic, study routing, review states, and trust signals inside one operating system.",
            "03",
            [
                (0.9, 3.0, 11.5, 0.85, "Register → Screen → Assign → Review → Reuse", {"box": True, "fill": "5D7442", "size": 20, "bold": True, "color": "FFFFFF", "align": "c"}),
                (0.9, 4.25, 3.6, 1.5, "One record per contributor\nIdentity, devices, expertise, notes, and past participation stay connected.", {"box": True, "fill": "F5F8F0", "size": 14, "bold": True}),
                (4.85, 4.25, 3.6, 1.5, "One operating flow per study\nInvitations, review, completion, and payout readiness stay visible.", {"box": True, "fill": "E6EDDC", "size": 14, "bold": True}),
                (8.8, 4.25, 3.6, 1.5, "One memory layer across programs\nTeams can re-contact stronger cohorts instead of starting from zero.", {"box": True, "fill": "DCE8CF", "size": 14, "bold": True}),
            ],
        )
    )

    # 5 participant infrastructure
    slides.append(
        slide(
            "F5F8F0",
            "5D7442",
            "Participant infrastructure, not just recruitment.",
            "The platform is designed for broad research pools, specialist cohorts, and recurring evaluation communities that need stronger operational memory.",
            "04",
            [
                (0.9, 2.95, 3.8, 2.2, "General participant pools\nSurveys, concept tests, interviews, and product feedback programs.", {"box": True, "fill": "FFFFFF", "line": "C9D8B5", "size": 16, "bold": True}),
                (4.78, 2.95, 3.8, 2.2, "Expert cohorts\nHealthcare, policy, technical, financial, and domain-specific contributors.", {"box": True, "fill": "E9F1E0", "line": "C9D8B5", "size": 16, "bold": True}),
                (8.66, 2.95, 3.8, 2.2, "Evaluation networks\nRepeat reviewers for annotation, red teaming, scoring, and model feedback.", {"box": True, "fill": "DCE8CF", "line": "C9D8B5", "size": 16, "bold": True}),
            ],
        )
    )

    # 6 study operations timeline
    s = SlideBuilder()
    s.rect(0, 0, 13.333, 7.5, "324826")
    s.text(0.9, 0.95, 5.3, 0.5, "05", size=18, color="C9D8B5", bold=True, all_caps=True)
    s.text(0.9, 1.45, 6.0, 1.1, "Study operations become visible.", size=25, color="FFFFFF", bold=True, font="Georgia")
    s.text(0.9, 2.35, 5.8, 0.9, "ResearchAI gives teams a single operating sequence from recruiting participants to closing work cleanly.", size=12, color="E7EFE2")
    s.line(2.0, 4.15, 11.4, 4.15, "C9D8B5", 2.0)
    steps = [
        (1.6, 4.4, "Recruit", "Broad pools and expert cohorts"),
        (4.15, 3.1, "Qualify", "Fit logic and human review"),
        (6.7, 4.4, "Assign", "Studies, tasks, interviews"),
        (9.25, 3.1, "Close", "Notes, outcomes, payout readiness"),
    ]
    for x, y, label, body in steps:
        s.rect(x, y, 2.15, 1.35, "F5F8F0", rounded=True)
        s.text(x + 0.18, y + 0.12, 1.8, 0.35, label, size=15, color="324826", bold=True)
        s.text(x + 0.18, y + 0.55, 1.8, 0.55, body, size=10.5, color="5F7059")
    slides.append((s.slide_xml(), s.images))

    # 7 AI workflow
    slides.append(
        slide(
            "5D7442",
            "E4EED8",
            "Built for participant, study, and AI workflows.",
            "The operating model can support human evaluation loops, annotation programs, benchmarking tasks, and repeatable feedback programs tied to AI product work.",
            "06",
            [
                (0.9, 3.15, 3.7, 2.45, "Model evaluations\nRoute trained reviewers into structured scoring and assessment programs.", {"box": True, "fill": "F5F8F0", "size": 15, "bold": True}),
                (4.82, 3.15, 3.7, 2.45, "Human feedback loops\nTrack who reviewed what, with what quality, and under which eligibility rules.", {"box": True, "fill": "E6EDDC", "size": 15, "bold": True}),
                (8.74, 3.15, 3.7, 2.45, "Repeat programs\nPreserve strong contributor networks for ongoing AI studies instead of rebuilding every cycle.", {"box": True, "fill": "C9D8B5", "size": 15, "bold": True}),
            ],
        )
    )

    # 8 institutions
    s = SlideBuilder()
    s.rect(0, 0, 13.333, 7.5, "F5F8F0")
    s.rect(0.52, 0.42, 12.25, 0.08, "6D8751")
    s.text(0.72, 0.55, 1.1, 0.55, "07", size=18, color="6D8751", bold=True, all_caps=True)
    s.text(0.72, 0.95, 6.4, 1.05, "Trusted institution sourcing is a demand-side advantage.", size=25, color="1F2D1B", bold=True, font="Georgia")
    s.text(0.72, 2.0, 6.3, 0.9, "Higher-trust participant pools matter more as AI studies move into regulated, academic, and expert-reviewed settings.", size=12, color="5F7059")
    s.rect(0.95, 3.0, 5.1, 2.7, "FFFFFF", line="DCE8CF", rounded=True)
    s.image(1.25, 3.55, 1.65, 1.15, ucl_logo)
    s.image(3.45, 3.48, 1.65, 1.28, edin_logo)
    s.text(1.15, 5.0, 1.9, 0.35, "UCL", size=11.5, color="324826", bold=True, align="c")
    s.text(3.25, 5.0, 2.1, 0.35, "Edinburgh", size=11.5, color="324826", bold=True, align="c")
    s.rect(6.45, 3.0, 5.95, 2.7, "E9F1E0", line="DCE8CF", rounded=True)
    s.text(6.8, 3.32, 5.2, 1.85, "Target contributor channels\nUniversities | health systems | policy groups | specialist reviewers | AI evaluation communities", size=16, color="324826", bold=True)
    s.text(6.8, 5.0, 5.0, 0.45, "Stronger sources improve response quality, auditability, and repeat-study readiness.", size=11.5, color="5F7059")
    slides.append((s.slide_xml(), s.images))

    # 9 use cases
    slides.append(
        slide(
            "324826",
            "DCE8CF",
            "Who ResearchAI is built for.",
            "The platform fits research groups that need operating discipline, auditability, and repeat-ready contributors.",
            "08",
            [
                (0.9, 3.0, 2.75, 1.85, "Product research teams\nUser interviews, concept testing, recurring feedback.", {"box": True, "fill": "F5F8F0", "size": 14, "bold": True}),
                (3.95, 3.0, 2.75, 1.85, "Policy and public interest research\nStructured contributor screening and documentation.", {"box": True, "fill": "E6EDDC", "size": 14, "bold": True}),
                (7.0, 3.0, 2.75, 1.85, "Healthcare and expert studies\nHigher-trust cohorts and specialist routing.", {"box": True, "fill": "DCE8CF", "size": 14, "bold": True}),
                (10.05, 3.0, 2.35, 1.85, "AI labs and evaluation programs\nOngoing human-in-the-loop research.", {"box": True, "fill": "F5F8F0", "size": 14, "bold": True}),
            ],
        )
    )

    # 10 visibility
    slides.append(
        slide(
            "F5F8F0",
            "5D7442",
            "The system surfaces research health, not just task lists.",
            "Teams need visibility across coverage, quality, and retention to know whether their participant infrastructure is actually getting stronger.",
            "09",
            [
                (1.0, 3.0, 3.6, 2.2, "Coverage\nSpot gaps in language, geography, devices, expertise, or availability.", {"box": True, "fill": "FFFFFF", "line": "DCE8CF", "size": 16, "bold": True}),
                (4.87, 3.0, 3.6, 2.2, "Quality\nStrengthen assignment decisions with notes, outcomes, and history.", {"box": True, "fill": "E9F1E0", "line": "DCE8CF", "size": 16, "bold": True}),
                (8.74, 3.0, 3.6, 2.2, "Retention\nPreserve trusted contributor cohorts for repeat programs.", {"box": True, "fill": "DCE8CF", "line": "DCE8CF", "size": 16, "bold": True}),
            ],
        )
    )

    # 11 governance
    slides.append(
        slide(
            "5D7442",
            "E4EED8",
            "Governance is part of the workflow, not a separate layer.",
            "Screening, review rules, contributor memory, and readiness signals are designed to stay attached to the same records as the work itself.",
            "10",
            [
                (0.9, 3.1, 3.75, 2.3, "Review controls\nScreen before assignment and keep eligibility logic attached to the study.", {"box": True, "fill": "F5F8F0", "size": 15, "bold": True}),
                (4.8, 3.1, 3.75, 2.3, "Contributor memory\nPreserve notes, trust signals, and history across repeated programs.", {"box": True, "fill": "E6EDDC", "size": 15, "bold": True}),
                (8.7, 3.1, 3.75, 2.3, "Program integrity\nReduce drift by keeping records, notes, and operational states together.", {"box": True, "fill": "DCE8CF", "size": 15, "bold": True}),
            ],
        )
    )

    # 12 market frame
    slides.append(
        slide(
            "F5F8F0",
            "6D8751",
            "Demand is large, fast-growing, and already segmented by revenue share.",
            "ResearchAI sits where AI infrastructure, human evaluation, and regulated research operations increasingly overlap.",
            "11",
            [
                (0.95, 3.0, 2.85, 2.15, "$390.9B\nAI market, 2025", {"box": True, "fill": "FFFFFF", "line": "DCE8CF", "size": 18, "bold": True, "align": "c"}),
                (4.0, 3.0, 2.85, 2.15, "35.5%\nNorth America share", {"box": True, "fill": "E9F1E0", "line": "DCE8CF", "size": 18, "bold": True, "align": "c"}),
                (7.05, 3.0, 2.85, 2.15, "$36.7B\nAI in healthcare, 2025", {"box": True, "fill": "DCE8CF", "line": "DCE8CF", "size": 18, "bold": True, "align": "c"}),
                (10.1, 3.0, 2.35, 2.15, "36.2%\nAnnotation tools NA share", {"box": True, "fill": "FFFFFF", "line": "DCE8CF", "size": 16.5, "bold": True, "align": "c"}),
                (0.95, 5.55, 11.4, 0.55, "Sources: Precedence Research (AI market, 2025), Precedence Research (AI in healthcare, 2025), Grand View Research (data annotation tools, 2024 share).", {"box": False, "size": 9.2, "color": "5F7059"}),
            ],
        )
    )

    # 13 roadmap
    s = SlideBuilder()
    s.rect(0, 0, 13.333, 7.5, "324826")
    s.text(0.9, 0.9, 1.0, 0.4, "12", size=18, color="C9D8B5", bold=True, all_caps=True)
    s.text(0.9, 1.35, 5.4, 0.9, "A staged rollout is straightforward.", size=25, color="FFFFFF", bold=True, font="Georgia")
    s.text(0.9, 2.2, 5.7, 0.8, "The platform can be introduced as a focused workflow system before expanding into repeat research infrastructure.", size=12, color="E7EFE2")
    phases = [
        ("Phase 1", "Core intake and screener workflows"),
        ("Phase 2", "Study routing, review, and operational visibility"),
        ("Phase 3", "Repeat-ready AI and contributor programs"),
    ]
    x = 1.0
    for i, (phase, desc) in enumerate(phases):
        s.rect(x, 3.5, 3.7, 2.1, "F5F8F0", rounded=True)
        s.text(x + 0.22, 3.75, 3.1, 0.4, phase, size=17, color="324826", bold=True)
        s.text(x + 0.22, 4.25, 3.1, 0.85, desc, size=13, color="5F7059")
        if i < 2:
            s.text(x + 3.85, 4.35, 0.3, 0.3, "→", size=26, color="DCE8CF", bold=True, align="c")
        x += 4.15
    slides.append((s.slide_xml(), s.images))

    # 14 contact close
    s = SlideBuilder()
    s.rect(0, 0, 13.333, 7.5, "F5F8F0")
    s.rect(0.85, 0.85, 11.65, 5.8, "324826", rounded=True)
    s.text(1.3, 1.35, 2.0, 0.35, "13", size=18, color="DCE8CF", bold=True, all_caps=True)
    s.text(1.3, 1.85, 5.3, 1.15, "ResearchAI", size=30, color="FFFFFF", bold=True, font="Georgia")
    s.text(1.3, 3.05, 5.3, 1.0, "Research infrastructure for modern participant, study, and AI workflows.", size=13, color="E7EFE2")
    s.rect(7.0, 1.45, 4.4, 3.8, "5D7442", rounded=True)
    s.text(7.35, 1.9, 3.6, 0.4, "Contact", size=18, color="FFFFFF", bold=True)
    s.text(7.35, 2.45, 3.4, 0.7, "info@researchai.co.uk", size=15, color="E4EED8", bold=True)
    s.text(7.35, 3.25, 3.4, 0.85, "75 Shelton Street\nLondon, England", size=13, color="E4EED8")
    s.text(1.3, 5.6, 4.9, 0.6, "Built to run calmer, stronger research systems.", size=18, color="FFFFFF", bold=True)
    slides.append((s.slide_xml(), s.images))

    return slides


def slide_rels(image_targets: list[str]) -> str:
    image_rels = "".join(
        f'\n  <Relationship Id="rId{i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{target}"/>'
        for i, target in enumerate(image_targets)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout7.xml"/>
{image_rels}
</Relationships>"""

def core_xml() -> str:
    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>ResearchAI Pitch Deck</dc:title>
  <dc:creator>OpenAI Codex</dc:creator>
  <cp:lastModifiedBy>OpenAI Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>"""


def app_xml(slide_count: int) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TotalTime>1</TotalTime><Words>0</Words><Application>Microsoft Macintosh PowerPoint</Application><PresentationFormat>On-screen Show (16:9)</PresentationFormat><Paragraphs>0</Paragraphs><Slides>{slide_count}</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides><MMClips>0</MMClips><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant><vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant><vt:variant><vt:i4>0</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Office Theme</vt:lpstr></vt:vector></TitlesOfParts><Manager></Manager><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinkBase></HyperlinkBase><HyperlinksChanged>false</HyperlinksChanged><AppVersion>14.0000</AppVersion></Properties>"""


def presentation_xml_from_template(slide_count: int) -> str:
    slide_ids = "".join(
        f'<p:sldId id="{255 + i}" r:id="rId{i + 6}"/>' for i in range(1, slide_count + 1)
    )
    return f"""<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1" autoCompressPictures="0"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>{slide_ids}</p:sldIdLst><p:sldSz cx="12191695" cy="6858000" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr><a:lvl1pPr marL="0" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="457200" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr><a:lvl3pPr marL="914400" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr><a:lvl4pPr marL="1371600" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr><a:lvl5pPr marL="1828800" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr><a:lvl6pPr marL="2286000" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr><a:lvl7pPr marL="2743200" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr><a:lvl8pPr marL="3200400" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr><a:lvl9pPr marL="3657600" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr></p:defaultTextStyle></p:presentation>"""


def presentation_rels_from_template(slide_count: int) -> str:
    fixed = [
        ('rId1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster', 'slideMasters/slideMaster1.xml'),
        ('rId2', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings', 'printerSettings/printerSettings1.bin'),
        ('rId3', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps', 'presProps.xml'),
        ('rId4', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps', 'viewProps.xml'),
        ('rId5', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', 'theme/theme1.xml'),
        ('rId6', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles', 'tableStyles.xml'),
    ]
    slides = [
        (f"rId{i + 6}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide", f"slides/slide{i}.xml")
        for i in range(1, slide_count + 1)
    ]
    rels = "".join(f'<Relationship Id="{rid}" Type="{typ}" Target="{target}"/>' for rid, typ, target in fixed + slides)
    return f"""<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>"""


def content_types_from_template(slide_count: int) -> str:
    overrides = "".join(
        f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(1, slide_count + 1)
    )
    return f"""<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.presentationml.printerSettings"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout10.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout11.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout4.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout5.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout6.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout7.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout8.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout9.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>{overrides}<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/></Types>"""


def write_pptx() -> None:
    slides = build_slides()
    template = Path("/Users/wale/Desktop/scifi/sweb/LOOPNODE_PITCH_DECK.pptx")
    media_map: dict[Path, str] = {}
    media_index = 1
    for _, image_paths in slides:
        for image_path in image_paths:
            if image_path not in media_map:
                suffix = image_path.suffix.lower()
                ext = ".jpg" if suffix == ".jpg" else ".jpeg" if suffix == ".jpeg" else ".png"
                media_map[image_path] = f"image{media_index}{ext}"
                media_index += 1
    with zipfile.ZipFile(template) as src, zipfile.ZipFile(OUT, "w", compression=zipfile.ZIP_DEFLATED) as z:
        keep = []
        for name in src.namelist():
            if name.startswith("ppt/slides/slide") or name.startswith("ppt/slides/_rels/slide"):
                continue
            if name.startswith("ppt/media/"):
                continue
            if name in {"ppt/presentation.xml", "ppt/_rels/presentation.xml.rels", "[Content_Types].xml", "docProps/core.xml", "docProps/app.xml"}:
                continue
            keep.append(name)
        for name in keep:
            z.writestr(name, src.read(name))
        for image_path, media_name in media_map.items():
            z.writestr(f"ppt/media/{media_name}", image_path.read_bytes())
        z.writestr("[Content_Types].xml", content_types_from_template(len(slides)))
        z.writestr("docProps/core.xml", core_xml())
        z.writestr("docProps/app.xml", app_xml(len(slides)))
        z.writestr("ppt/presentation.xml", presentation_xml_from_template(len(slides)))
        z.writestr("ppt/_rels/presentation.xml.rels", presentation_rels_from_template(len(slides)))
        for i, (xml, image_paths) in enumerate(slides, start=1):
            z.writestr(f"ppt/slides/slide{i}.xml", xml)
            image_targets = [media_map[path] for path in image_paths]
            z.writestr(f"ppt/slides/_rels/slide{i}.xml.rels", slide_rels(image_targets))


if __name__ == "__main__":
    write_pptx()
    print(OUT)
