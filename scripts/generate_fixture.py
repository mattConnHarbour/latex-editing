#!/usr/bin/env python3
"""Generate a DOCX fixture containing native Office Math (OMML) equations."""

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "public" / "calculus-equations.docx"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def math_run(text: str, *, normal: bool = False) -> str:
    props = "<m:rPr><m:nor/></m:rPr>" if normal else ""
    preserve = ' xml:space="preserve"' if text[:1].isspace() or text[-1:].isspace() else ""
    return f"<m:r>{props}<m:t{preserve}>{escape(text)}</m:t></m:r>"


def seq(*parts: str) -> str:
    return "".join(parts)


def frac(numerator: str, denominator: str) -> str:
    return f"<m:f><m:fPr/><m:num>{numerator}</m:num><m:den>{denominator}</m:den></m:f>"


def superscript(base: str, exponent: str) -> str:
    return f"<m:sSup><m:sSupPr/><m:e>{base}</m:e><m:sup>{exponent}</m:sup></m:sSup>"


def subscript(base: str, sub: str) -> str:
    return f"<m:sSub><m:sSubPr/><m:e>{base}</m:e><m:sub>{sub}</m:sub></m:sSub>"


def lower_limit(base: str, limit: str) -> str:
    return f"<m:limLow><m:limLowPr/><m:e>{base}</m:e><m:lim>{limit}</m:lim></m:limLow>"


def sub_sup(base: str, sub: str, sup: str) -> str:
    return f"<m:sSubSup><m:sSubSupPr/><m:e>{base}</m:e><m:sub>{sub}</m:sub><m:sup>{sup}</m:sup></m:sSubSup>"


def radical(value: str, degree: str | None = None) -> str:
    deg = degree or ""
    hide = '<m:degHide m:val="1"/>' if degree is None else ""
    return f"<m:rad><m:radPr>{hide}</m:radPr><m:deg>{deg}</m:deg><m:e>{value}</m:e></m:rad>"


def delimiter(value: str, left: str = "(", right: str = ")") -> str:
    return (
        f'<m:d><m:dPr><m:begChr m:val="{escape(left)}"/><m:endChr m:val="{escape(right)}"/></m:dPr>'
        f"<m:e>{value}</m:e></m:d>"
    )


def nary(symbol: str, lower: str, upper: str, expression: str) -> str:
    empty_run = math_run("")
    lower_hidden = lower == empty_run
    upper_hidden = upper == empty_run
    hidden = ('<m:subHide m:val="1"/>' if lower_hidden else "") + ('<m:supHide m:val="1"/>' if upper_hidden else "")
    return (
        f'<m:nary><m:naryPr><m:chr m:val="{symbol}"/><m:limLoc m:val="undOvr"/>{hidden}</m:naryPr>'
        f"<m:sub>{lower}</m:sub><m:sup>{upper}</m:sup><m:e>{expression}</m:e></m:nary>"
    )


def equation(parts: str) -> str:
    return (
        '<w:p><w:pPr><w:spacing w:before="0" w:after="140"/></w:pPr>'
        f'<m:oMathPara><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr><m:oMath>{parts}</m:oMath></m:oMathPara>'
        '</w:p>'
    )


def paragraph(text: str, style: str | None = None) -> str:
    props = f'<w:pStyle w:val="{style}"/>' if style else ""
    if style in {"FormulaLabel", "Heading1"}:
        props += "<w:keepNext/><w:keepLines/>"
    return f"<w:p><w:pPr>{props}</w:pPr><w:r><w:t>{escape(text)}</w:t></w:r></w:p>"


def formula(label: str, parts: str) -> str:
    return paragraph(label, "FormulaLabel") + equation(parts)


def build_body() -> str:
    r = math_run
    sections: list[tuple[str, list[tuple[str, str]]]] = [
        (
            "Limits and derivatives",
            [
                ("Derivative from first principles", seq(r("f′(x) = "), lower_limit(r("lim"), r("h→0")), frac(seq(r("f(x+h) − f(x)")), r("h")))),
                ("Power rule", seq(frac(r("d"), r("dx")), superscript(r("x"), r("n")), r(" = n"), superscript(r("x"), r("n−1")))),
                ("Product rule", seq(frac(r("d"), r("dx")), delimiter(r("uv")), r(" = u′v + uv′"))),
                ("Quotient rule", seq(frac(r("d"), r("dx")), delimiter(frac(r("u"), r("v"))), r(" = "), frac(r("vu′ − uv′"), superscript(r("v"), r("2"))))),
                ("Chain rule", seq(frac(r("dy"), r("dx")), r(" = "), frac(r("dy"), r("du")), frac(r("du"), r("dx")))),
                ("L'Hôpital's rule", seq(lower_limit(r("lim"), r("x→a")), frac(r("f(x)"), r("g(x)")), r(" = "), lower_limit(r("lim"), r("x→a")), frac(r("f′(x)"), r("g′(x)")))),
            ],
        ),
        (
            "Integrals and the fundamental theorem",
            [
                ("Fundamental theorem of calculus", seq(frac(r("d"), r("dx")), nary("∫", r("a"), r("x"), r("f(t) dt")), r(" = f(x)"))),
                ("Antiderivative power rule", seq(nary("∫", r(""), r(""), superscript(r("x"), r("n"))), r(" dx = "), frac(superscript(r("x"), r("n+1")), r("n+1")), r(" + C"))),
                ("Integration by parts", seq(nary("∫", r(""), r(""), r("u dv")), r(" = uv − "), nary("∫", r(""), r(""), r("v du")))),
                ("Substitution rule", seq(nary("∫", r(""), r(""), r("f(g(x))g′(x) dx")), r(" = "), nary("∫", r(""), r(""), r("f(u) du")))),
                ("Improper Gaussian integral", seq(nary("∫", r("−∞"), r("∞"), superscript(r("e"), r("−x²"))), r(" dx = "), radical(r("π")))),
                ("Average value on an interval", seq(subscript(r("f"), r("avg")), r(" = "), frac(r("1"), r("b−a")), nary("∫", r("a"), r("b"), r("f(x) dx")))),
            ],
        ),
        (
            "Sums, products, and series",
            [
                ("Arithmetic sum", seq(nary("∑", r("k=1"), r("n"), r("k")), r(" = "), frac(r("n(n+1)"), r("2")))),
                ("Sum of squares", seq(nary("∑", r("k=1"), r("n"), superscript(r("k"), r("2"))), r(" = "), frac(r("n(n+1)(2n+1)"), r("6")))),
                ("Finite geometric series", seq(nary("∑", r("k=0"), r("n"), superscript(r("r"), r("k"))), r(" = "), frac(seq(r("1 − "), superscript(r("r"), r("n+1"))), r("1 − r")))),
                ("Infinite geometric series", seq(nary("∑", r("k=0"), r("∞"), superscript(r("r"), r("k"))), r(" = "), frac(r("1"), r("1 − r")), r(",  −1 < r < 1"))),
                ("Riemann sum", seq(nary("∫", r("a"), r("b"), r("f(x) dx")), r(" = "), lower_limit(r("lim"), r("n→∞")), nary("∑", r("i=1"), r("n"), r("f(ξ) Δx")))),
                ("Taylor series", seq(r("f(x) = "), nary("∑", r("n=0"), r("∞"), frac(seq(superscript(r("f"), r("(n)")), r("(a)"), superscript(seq(r("(x−a)")), r("n"))), r("n!"))))),
                ("Wallis product", seq(frac(r("π"), r("2")), r(" = "), nary("∏", r("n=1"), r("∞"), frac(superscript(r("(2n)"), r("2")), r("(2n−1)(2n+1)"))))),
            ],
        ),
        (
            "Multivariable and vector calculus",
            [
                ("Gradient", seq(r("∇f = "), delimiter(seq(frac(r("∂f"), r("∂x")), r(", "), frac(r("∂f"), r("∂y")), r(", "), frac(r("∂f"), r("∂z")))))),
                ("Mixed partial derivative", seq(frac(seq(superscript(r("∂"), r("2")), r("f")), r("∂x ∂y")), r(" = "), frac(seq(superscript(r("∂"), r("2")), r("f")), r("∂y ∂x")))),
                ("Laplacian", seq(superscript(r("∇"), r("2")), r("f = "), frac(seq(superscript(r("∂"), r("2")), r("f")), superscript(r("∂x"), r("2"))), r(" + "), frac(seq(superscript(r("∂"), r("2")), r("f")), superscript(r("∂y"), r("2"))), r(" + "), frac(seq(superscript(r("∂"), r("2")), r("f")), superscript(r("∂z"), r("2"))))),
                ("Double integral", seq(nary("∫", r(""), r(""), nary("∫", r("D"), r(""), r("f(x,y) dA"))))),
                ("Divergence theorem", seq(nary("∯", r("∂V"), r(""), r("F · n dS")), r(" = "), nary("∭", r("V"), r(""), r("∇ · F dV")))),
                ("Green's theorem", seq(nary("∮", r("C"), r(""), r("P dx + Q dy")), r(" = "), nary("∬", r("D"), r(""), delimiter(seq(frac(r("∂Q"), r("∂x")), r(" − "), frac(r("∂P"), r("∂y"))))), r(" dA"))),
            ],
        ),
        (
            "Differential equations and transforms",
            [
                ("Exponential growth", seq(frac(r("dy"), r("dt")), r(" = ky  ⇒  y(t) = "), subscript(r("y"), r("0")), superscript(r("e"), r("kt")))),
                ("Second-order linear equation", seq(frac(seq(superscript(r("d"), r("2")), r("y")), superscript(r("dx"), r("2"))), r(" + p(x)"), frac(r("dy"), r("dx")), r(" + q(x)y = g(x)"))),
                ("Logistic equation", seq(frac(r("dP"), r("dt")), r(" = rP"), delimiter(seq(r("1 − "), frac(r("P"), r("K")))))),
                ("Laplace transform", seq(r("L{f(t)} = F(s) = "), nary("∫", r("0"), r("∞"), superscript(r("e"), r("−st"))), r("f(t) dt"))),
                ("Fourier transform", seq(r("F(ω) = "), nary("∫", r("−∞"), r("∞"), r("f(t)")), superscript(r("e"), r("−iωt")), r(" dt"))),
            ],
        ),
    ]

    body = [
        paragraph("CALCULUS REFERENCE", "Kicker"),
        paragraph("Equation Rendering Fixture", "Title"),
        paragraph("A stress document for native Word math layout, import, editing, and export", "Subtitle"),
        paragraph("This fixture intentionally combines stacked fractions, large operators, upper and lower limits, radicals, scripts, delimiters, Greek symbols, and nested expressions.", "Intro"),
    ]
    for heading, items in sections:
        body.append(paragraph(heading, "Heading1"))
        for label, parts in items:
            if label == "Taylor series":
                body.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
            body.append(formula(label, parts))
    return "".join(body)


def build_document() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}" xmlns:m="{M}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>{build_body()}
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId3"/>
      <w:footerReference w:type="default" r:id="rId4"/>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
      <w:cols w:space="708"/><w:docGrid w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>'''


STYLES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W}">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:eastAsia="Aptos"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="300" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="120" w:line="300" w:lineRule="auto"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Subtitle"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="0" w:after="100"/></w:pPr><w:rPr><w:rFonts w:ascii="Aptos Display" w:hAnsi="Aptos Display"/><w:b/><w:color w:val="172033"/><w:sz w:val="52"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Intro"/><w:pPr><w:spacing w:after="280"/></w:pPr><w:rPr><w:color w:val="526078"/><w:sz w:val="25"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Kicker"><w:name w:val="Kicker"/><w:basedOn w:val="Normal"/><w:next w:val="Title"/><w:pPr><w:spacing w:after="100"/></w:pPr><w:rPr><w:b/><w:color w:val="3157D5"/><w:sz w:val="18"/><w:caps/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Intro"><w:name w:val="Intro"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="180" w:line="300" w:lineRule="auto"/></w:pPr><w:rPr><w:color w:val="3F4A5E"/><w:sz w:val="22"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="FormulaLabel"/><w:qFormat/><w:uiPriority w:val="9"/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="200"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:ascii="Aptos Display" w:hAnsi="Aptos Display"/><w:b/><w:color w:val="2E5AAC"/><w:sz w:val="32"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="FormulaLabel"><w:name w:val="Formula label"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="120" w:after="20"/></w:pPr><w:rPr><w:b/><w:color w:val="43506A"/><w:sz w:val="20"/></w:rPr></w:style>
</w:styles>'''


CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''


def write_fixture() -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    parts = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>''',
        "word/document.xml": build_document(),
        "word/styles.xml": STYLES,
        "word/settings.xml": f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="{W}" xmlns:m="{M}"><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><m:mathPr><m:mathFont m:val="Cambria Math"/><m:brkBin m:val="before"/><m:smallFrac m:val="0"/><m:dispDef/></m:mathPr></w:settings>''',
        "word/fontTable.xml": f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts xmlns:w="{W}"><w:font w:name="Aptos"/><w:font w:name="Aptos Display"/><w:font w:name="Cambria Math"/></w:fonts>''',
        "word/_rels/document.xml.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/></Relationships>''',
        "word/header1.xml": f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="{W}"><w:p><w:pPr><w:spacing w:after="0"/><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="6" w:color="D8DEEA"/></w:pBdr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:color w:val="687386"/><w:sz w:val="17"/><w:caps/></w:rPr><w:t>Equation rendering fixture</w:t></w:r></w:p></w:hdr>''',
        "word/footer1.xml": f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="{W}"><w:p><w:pPr><w:jc w:val="right"/><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:rPr><w:color w:val="687386"/><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">Page </w:t></w:r><w:fldSimple w:instr="PAGE"><w:r><w:rPr><w:color w:val="687386"/><w:sz w:val="18"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>''',
        "docProps/core.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Equation Rendering Fixture</dc:title><dc:subject>Common calculus equations in native Word math</dc:subject><dc:creator>SuperDoc</dc:creator><cp:keywords>calculus, equations, OMML, rendering</cp:keywords><dcterms:created xsi:type="dcterms:W3CDTF">2026-10-06T00:00:00Z</dcterms:created></cp:coreProperties>''',
        "docProps/app.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>SuperDoc fixture generator</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><Company>SuperDoc</Company><AppVersion>1.0</AppVersion></Properties>''',
    }
    with ZipFile(OUTPUT, "w", ZIP_DEFLATED) as archive:
        for name, content in parts.items():
            archive.writestr(name, content.encode("utf-8"))
    print(f"Wrote {OUTPUT}")


if __name__ == "__main__":
    write_fixture()
