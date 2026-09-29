"""
Shared fixtures for the document_processor test suite.

Each builder constructs a real .docx on disk exercising one structure that the
scan/replace code has to get right: hyperlinks, merged cells, linked headers,
inline images, nested tables, text boxes and tracked insertions.
"""

import base64
import os
import sys

import pytest
from docx import Document
from docx.enum.section import WD_SECTION
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.shared import Inches

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Smallest valid PNG, used wherever a test needs a real inline image.
_PNG_1PX = base64.b64decode(
    b"iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


# ---------------------------------------------------------------- helpers

def add_run(paragraph, text):
    """Append a plain <w:r><w:t> run, bypassing python-docx's text setter."""
    xml = (
        '<w:r xmlns:w="%s"><w:t xml:space="preserve">%s</w:t></w:r>'
        % (NS_W, _escape(text))
    )
    paragraph._p.append(parse_xml(xml))


def add_runs(paragraph, *texts):
    """Append several runs, so a phrase can be split across run boundaries."""
    for text in texts:
        add_run(paragraph, text)


def add_hyperlink(paragraph, text, url="https://example.com/spec"):
    """Append a real external hyperlink containing `text`."""
    r_id = paragraph.part.relate_to(
        url, "%s/hyperlink" % NS_R, is_external=True
    )
    xml = (
        '<w:hyperlink xmlns:w="%s" xmlns:r="%s" r:id="%s">'
        '<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:hyperlink>'
        % (NS_W, NS_R, r_id, _escape(text))
    )
    paragraph._p.append(parse_xml(xml))


def add_tracked_insertion(paragraph, text, author="Reviewer"):
    """Append a <w:ins> tracked insertion containing `text`."""
    xml = (
        '<w:ins xmlns:w="%s" w:id="900" w:author="%s" w:date="2026-01-01T00:00:00Z">'
        '<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:ins>'
        % (NS_W, author, _escape(text))
    )
    paragraph._p.append(parse_xml(xml))


def add_deletion(paragraph, text, author="Reviewer"):
    """Append a <w:del> tracked deletion (uses w:delText, must stay untouched)."""
    xml = (
        '<w:del xmlns:w="%s" w:id="901" w:author="%s" w:date="2026-01-01T00:00:00Z">'
        '<w:r><w:delText xml:space="preserve">%s</w:delText></w:r></w:del>'
        % (NS_W, author, _escape(text))
    )
    paragraph._p.append(parse_xml(xml))


def add_formatted_run(paragraph, text, bold=False, color=None):
    """Append a run carrying explicit formatting (w:rPr with w:b / w:color)."""
    props = ""
    if bold:
        props += "<w:b/>"
    if color:
        props += '<w:color w:val="%s"/>' % color
    xml = (
        '<w:r xmlns:w="%s"><w:rPr>%s</w:rPr>'
        '<w:t xml:space="preserve">%s</w:t></w:r>'
        % (NS_W, props, _escape(text))
    )
    paragraph._p.append(parse_xml(xml))


def add_run_with_rpr_change(paragraph, text, change_id=902):
    """Append a bold run whose formatting is itself a tracked change."""
    xml = (
        '<w:r xmlns:w="%s"><w:rPr><w:b/>'
        '<w:rPrChange w:id="%d" w:author="Reviewer" w:date="2026-01-01T00:00:00Z">'
        '<w:rPr/></w:rPrChange></w:rPr>'
        '<w:t xml:space="preserve">%s</w:t></w:r>'
        % (NS_W, change_id, _escape(text))
    )
    paragraph._p.append(parse_xml(xml))


def add_tab_run(paragraph):
    """Append a run holding only a <w:tab/>."""
    paragraph._p.append(parse_xml('<w:r xmlns:w="%s"><w:tab/></w:r>' % NS_W))


def add_paragraph_with_outline_level(doc, text, level):
    """A Normal-styled paragraph carrying its own w:outlineLvl."""
    p = doc.add_paragraph(text)
    ppr = p._p.get_or_add_pPr()
    ppr.append(parse_xml('<w:outlineLvl xmlns:w="%s" w:val="%d"/>' % (NS_W, level)))
    return p


def add_custom_paragraph_style(doc, style_id, name, based_on=None, outline_level=None):
    """Append a paragraph style to the styles part. Test code may use
    doc.styles; only the scan path must not."""
    parts = ['<w:style xmlns:w="%s" w:type="paragraph" w:styleId="%s">' % (NS_W, style_id),
             '<w:name w:val="%s"/>' % _escape(name)]
    if based_on:
        parts.append('<w:basedOn w:val="%s"/>' % based_on)
    if outline_level is not None:
        parts.append('<w:pPr><w:outlineLvl w:val="%d"/></w:pPr>' % outline_level)
    parts.append('</w:style>')
    doc.styles.element.append(parse_xml("".join(parts)))


def add_styled_paragraph(doc, text, style_id):
    """A paragraph whose w:pStyle names `style_id` directly (no style lookup)."""
    p = doc.add_paragraph(text)
    ppr = p._p.get_or_add_pPr()
    ppr.insert(0, parse_xml('<w:pStyle xmlns:w="%s" w:val="%s"/>' % (NS_W, style_id)))
    return p


def drop_styles_part(doc):
    """Detach the styles part so the saved package has none."""
    for r_id, rel in list(doc.part.rels.items()):
        if rel.reltype.endswith("/styles"):
            doc.part.drop_rel(r_id)


def has_styles_part(path):
    from zipfile import ZipFile
    with ZipFile(path) as z:
        return "word/styles.xml" in z.namelist()


def add_complex_field(paragraph, instruction, cached_result):
    """
    Append a complex field: fldChar begin / instrText / separate / cached
    result / fldChar end.

    This is how Word stores a TOC entry or a cross-reference. The cached result
    lives in an ordinary <w:t>, which is why it is scanned like any other text.
    """
    parts = [
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
        '<w:r><w:instrText xml:space="preserve">%s</w:instrText></w:r>' % _escape(instruction),
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
        '<w:r><w:t xml:space="preserve">%s</w:t></w:r>' % _escape(cached_result),
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    ]
    for part in parts:
        xml = part.replace('<w:r>', '<w:r xmlns:w="%s">' % NS_W, 1)
        paragraph._p.append(parse_xml(xml))


def add_simple_field(paragraph, instruction, cached_result):
    """Append a <w:fldSimple> whose cached result sits in a child run."""
    xml = (
        '<w:fldSimple xmlns:w="%s" w:instr="%s">'
        '<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:fldSimple>'
        % (NS_W, _escape(instruction), _escape(cached_result))
    )
    paragraph._p.append(parse_xml(xml))


def add_vml_textbox(paragraph, text):
    """Append a legacy VML text box containing a paragraph of `text`."""
    xml = (
        '<w:r xmlns:w="%s" xmlns:v="urn:schemas-microsoft-com:vml">'
        '<w:pict><v:shape id="tb1" style="width:100pt;height:50pt">'
        '<v:textbox><w:txbxContent>'
        '<w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>'
        '</w:txbxContent></v:textbox></v:shape></w:pict></w:r>'
        % (NS_W, _escape(text))
    )
    paragraph._p.append(parse_xml(xml))


def add_alternate_content_textbox(paragraph, choice_text, fallback_text):
    """
    Append a text box in mc:AlternateContent form.

    Word writes both a modern mc:Choice copy and a legacy mc:Fallback copy of
    the same text box. Only the Choice copy should be surfaced, otherwise one
    visible text box would report two matches.
    """
    xml = (
        '<mc:AlternateContent xmlns:mc="%s" xmlns:w="%s" '
        'xmlns:v="urn:schemas-microsoft-com:vml">'
        '<mc:Choice Requires="wps">'
        '<w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>'
        '</mc:Choice>'
        '<mc:Fallback>'
        '<w:p><w:r><w:pict><v:shape id="tb2" style="width:1pt;height:1pt">'
        '<v:textbox><w:txbxContent>'
        '<w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>'
        '</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>'
        '</mc:Fallback>'
        '</mc:AlternateContent>'
        % (NS_MC, NS_W, _escape(choice_text), _escape(fallback_text))
    )
    paragraph._p.addnext(parse_xml(xml))


def add_notes_part(doc, kind, texts):
    """
    Attach a footnotes / endnotes / comments part containing one note per text.

    `kind` is "footnote", "endnote" or "comment". A separator entry is included
    for footnotes, mirroring what Word writes, so the traversal has to skip it.
    """
    from docx.opc.constants import CONTENT_TYPE as CT
    from docx.opc.packuri import PackURI
    from docx.opc.part import XmlPart

    spec = {
        "footnote": (CT.WML_FOOTNOTES, "footnotes", "/word/footnotes.xml"),
        "endnote": (CT.WML_ENDNOTES, "endnotes", "/word/endnotes.xml"),
        "comment": (CT.WML_COMMENTS, "comments", "/word/comments.xml"),
    }[kind]
    content_type, root_tag, partname = spec
    reltype = "%s/%s" % (NS_R, root_tag)

    entries = []
    if kind == "footnote":
        entries.append(
            '<w:footnote w:type="separator" w:id="-1">'
            '<w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
        )
    for i, text in enumerate(texts, start=2):
        entries.append(
            '<w:%s w:id="%d"><w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p></w:%s>'
            % (kind, i, _escape(text), kind)
        )

    xml = '<w:%s xmlns:w="%s">%s</w:%s>' % (root_tag, NS_W, "".join(entries), root_tag)
    part = XmlPart(PackURI(partname), content_type, parse_xml(xml), doc.part.package)
    doc.part.relate_to(part, reltype)
    return part


def _escape(text):
    return (text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def count_drawings(path):
    """Number of <w:drawing> elements anywhere in the document body."""
    doc = Document(path)
    return len(doc.element.body.findall(".//" + qn("w:drawing")))


def all_text(path):
    """Every visible paragraph's text, using the processor's own traversal."""
    import document_processor as dp
    doc = Document(path)
    return [dp.paragraph_text(p) for p, _t, _d, _k in dp._collect_paragraphs(doc)]


def revisions(path):
    """Every w:ins / w:del in the body, in document order, as
    (local tag, id, author, text)."""
    doc = Document(path)
    out = []
    for el in doc.element.body.iter(qn("w:ins"), qn("w:del")):
        texts = [t.text or "" for t in el.iter(qn("w:t"), qn("w:delText"))]
        out.append((el.tag.split("}")[1], el.get(qn("w:id")),
                    el.get(qn("w:author")), "".join(texts)))
    return out


def all_revision_ids(path):
    """Every w:ins / w:del id across every XML part of the package."""
    doc = Document(path)
    ids = []
    for part in doc.part.package.iter_parts():
        root = getattr(part, "element", None)
        if root is None:
            continue
        for el in root.iter(qn("w:ins"), qn("w:del"), qn("w:rPrChange")):
            ids.append(el.get(qn("w:id")))
    return ids


def _resolved_texts(path, keep):
    """Paragraph texts after simulating Word's Accept All (keep='ins') or
    Reject All (keep='del') on the whole package."""
    import document_processor as dp
    doc = Document(path)
    drop, unwrap = (qn("w:del"), qn("w:ins")) if keep == "ins" else (qn("w:ins"), qn("w:del"))
    for part in doc.part.package.iter_parts():
        root = getattr(part, "element", None)
        if root is None:
            continue
        for el in list(root.iter(drop)):
            el.getparent().remove(el)
        for el in list(root.iter(unwrap)):
            parent = el.getparent()
            for child in list(el):
                if keep == "del":
                    for dt in child.iter(qn("w:delText")):
                        t = parse_xml('<w:t xmlns:w="%s" xml:space="preserve"/>' % NS_W)
                        t.text = dt.text
                        dt.getparent().replace(dt, t)
                el.addprevious(child)
            parent.remove(el)
    return [dp.paragraph_text(p) for p, _t, _d, _k in dp._collect_paragraphs(doc)]


def accept_all(path):
    return _resolved_texts(path, "ins")


def reject_all(path):
    return _resolved_texts(path, "del")


# ---------------------------------------------------------------- fixtures

@pytest.fixture
def png_file(tmp_path):
    path = tmp_path / "dot.png"
    path.write_bytes(_PNG_1PX)
    return str(path)


@pytest.fixture
def docdir(tmp_path):
    """An empty folder for .docx fixtures, matching the tool's folder-scan API."""
    d = tmp_path / "docs"
    d.mkdir()
    return d


@pytest.fixture
def hyperlink_doc(docdir):
    """'See <hyperlink>2022 CBC link</hyperlink> and done.'"""
    doc = Document()
    p = doc.add_paragraph()
    add_run(p, "See ")
    add_hyperlink(p, "2022 CBC link")
    add_run(p, " and done.")
    path = docdir / "hyperlink.docx"
    doc.save(str(path))
    return str(path)


@pytest.fixture
def merged_and_linked_doc(docdir):
    """A merged cell across 3 grid columns plus a header shared by 3 sections."""
    doc = Document()
    doc.add_paragraph("Body 2022 CBC baseline")

    table = doc.add_table(rows=2, cols=3)
    table.cell(0, 0).merge(table.cell(0, 2))
    table.cell(0, 0).paragraphs[0].text = "Merged: 2022 CBC note"
    table.cell(1, 0).text = "Plain 2022 CBC cell"

    doc.sections[0].header.paragraphs[0].text = "HEADER 2022 CBC"
    doc.add_section(WD_SECTION.NEW_PAGE)
    doc.add_section(WD_SECTION.NEW_PAGE)

    path = docdir / "merged_linked.docx"
    doc.save(str(path))
    return str(path)


@pytest.fixture
def image_in_run_doc(docdir, png_file):
    """A match spanning three runs, the middle of which carries an inline image."""
    doc = Document()
    p = doc.add_paragraph()
    add_run(p, "See Figure 20")
    run = p.add_run()
    run.add_picture(png_file, width=Inches(0.2))
    run.add_text("22")
    add_run(p, " CBC ref")
    path = docdir / "image.docx"
    doc.save(str(path))
    return str(path)


@pytest.fixture
def nested_table_doc(docdir):
    doc = Document()
    outer = doc.add_table(rows=1, cols=1)
    inner = outer.cell(0, 0).add_table(rows=1, cols=1)
    inner.cell(0, 0).text = "Nested 2022 CBC"
    path = docdir / "nested.docx"
    doc.save(str(path))
    return str(path)


@pytest.fixture
def textbox_doc(docdir):
    doc = Document()
    p = doc.add_paragraph()
    add_run(p, "Anchor paragraph.")
    add_vml_textbox(p, "Textbox 2022 CBC")
    path = docdir / "textbox.docx"
    doc.save(str(path))
    return str(path)


@pytest.fixture
def alt_content_doc(docdir):
    doc = Document()
    p = doc.add_paragraph()
    add_run(p, "Anchor.")
    add_alternate_content_textbox(p, "Choice 2022 CBC", "Choice 2022 CBC")
    path = docdir / "altcontent.docx"
    doc.save(str(path))
    return str(path)


@pytest.fixture
def field_doc(docdir):
    """A complex field (cross-reference) and a w:fldSimple, both citing 2022 CBC."""
    doc = Document()
    p = doc.add_paragraph()
    add_run(p, "Refer to ")
    add_complex_field(p, r' REF _Ref1 \\h ', "2022 CBC")
    add_run(p, " and ")
    add_simple_field(p, r' REF _Ref2 \\h ', "2022 CBC simple")
    add_run(p, " end.")
    path = docdir / "fields.docx"
    doc.save(str(path))
    return str(path)


@pytest.fixture
def tracked_changes_doc(docdir):
    doc = Document()
    p = doc.add_paragraph()
    add_run(p, "Kept ")
    add_tracked_insertion(p, "2022 CBC inserted")
    add_deletion(p, " 2022 CBC deleted")
    path = docdir / "tracked.docx"
    doc.save(str(path))
    return str(path)


@pytest.fixture
def notes_doc(docdir):
    """Body plus a footnote, an endnote and a comment, all citing 2022 CBC."""
    doc = Document()
    doc.add_paragraph("Body cites 2022 CBC here.")
    add_notes_part(doc, "footnote", ["Footnote cites 2022 CBC too."])
    add_notes_part(doc, "endnote", ["Endnote cites 2022 CBC as well."])
    add_notes_part(doc, "comment", ["Reviewer asks about 2022 CBC."])
    path = docdir / "notes.docx"
    doc.save(str(path))
    return str(path)


@pytest.fixture
def split_run_doc(docdir):
    """'Reference the 2022 CBC for compliance.' split across run boundaries."""
    doc = Document()
    p = doc.add_paragraph()
    add_runs(p, "Reference the 20", "22 ", "CBC for compliance.")
    path = docdir / "split.docx"
    doc.save(str(path))
    return str(path)


@pytest.fixture
def repeated_doc(docdir):
    doc = Document()
    doc.add_paragraph("aaaa and 2022 2022 2022")
    path = docdir / "repeated.docx"
    doc.save(str(path))
    return str(path)


@pytest.fixture
def formatted_split_doc(docdir):
    """'20' bold, '22 CBC' bold+red, ' ref' plain - a phrase spanning formats."""
    doc = Document()
    p = doc.add_paragraph()
    add_formatted_run(p, "20", bold=True)
    add_formatted_run(p, "22 CBC", bold=True, color="FF0000")
    add_run(p, " ref")
    path = docdir / "formatted.docx"
    doc.save(str(path))
    return str(path)


@pytest.fixture
def nested_folder_docs(docdir):
    """a.docx at the top, sub/b.docx, sub/deeper/c.docx, plus a Word lock file."""
    (docdir / "sub" / "deeper").mkdir(parents=True)
    paths = []
    for rel in ("a.docx", os.path.join("sub", "b.docx"),
                os.path.join("sub", "deeper", "c.docx")):
        doc = Document()
        doc.add_paragraph("Nested cites 2022 CBC.")
        path = docdir / rel
        doc.save(str(path))
        paths.append(str(path))
    (docdir / "sub" / "~$lock.docx").write_bytes(b"lock")
    return paths


@pytest.fixture
def heading_doc(docdir):
    """Heading 1 / body / Heading 2 / body / table / Heading 1 / body."""
    doc = Document()
    doc.add_heading("GENERAL", level=1)
    doc.add_paragraph("Intro cites 2022 CBC.")
    doc.add_heading("REFERENCES", level=2)
    doc.add_paragraph("Body cites 2022 CBC.")
    table = doc.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "Cell cites 2022 CBC."
    doc.add_heading("PRODUCTS", level=1)
    doc.add_paragraph("Later cites 2022 CBC.")
    path = docdir / "headings.docx"
    doc.save(str(path))
    return str(path)
