# SPDX-FileCopyrightText: The Docling Contributors
# SPDX-License-Identifier: MIT

"""Tests for blank paragraphs in a PowerPoint text shape."""

from io import BytesIO

from docling_core.types.doc import DoclingDocument, GroupLabel, ListItem, TextItem
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Inches

from docling.backend.mspowerpoint_backend import MsPowerpointDocumentBackend
from docling.datamodel.base_models import InputFormat
from docling.datamodel.document import InputDocument


def _convert(slide_setup) -> DoclingDocument:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Title"
    slide_setup(slide)
    stream = BytesIO()
    prs.save(stream)
    stream.seek(0)
    in_doc = InputDocument(
        path_or_stream=stream,
        format=InputFormat.PPTX,
        backend=MsPowerpointDocumentBackend,
        filename="blank.pptx",
    )
    return in_doc._backend.convert()


def test_blank_bullet_paragraphs_add_no_list_items() -> None:
    """A blank paragraph in a bulleted body is spacing, not an empty bullet."""

    def setup(slide) -> None:
        frame = slide.placeholders[1].text_frame
        frame.text = "a"
        for text in ["", "b", "   ", "c", "\u00a0"]:
            frame.add_paragraph().text = text

    doc = _convert(setup)

    assert doc.export_to_markdown() == "# Title\n\n- a\n- b\n- c"


def test_blank_paragraph_does_not_use_up_a_list_number() -> None:
    def setup(slide) -> None:
        frame = slide.shapes.add_textbox(
            Inches(1), Inches(4), Inches(4), Inches(1)
        ).text_frame
        frame.text = "a"
        for text in ["", "b"]:
            frame.add_paragraph().text = text
        for paragraph in frame.paragraphs:
            paragraph._p.get_or_add_pPr().append(
                paragraph._p.makeelement(qn("a:buAutoNum"), {"type": "arabicPeriod"})
            )

    doc = _convert(setup)

    items = [i for i, _ in doc.iterate_items() if isinstance(i, ListItem)]
    assert [(i.text, i.marker) for i in items] == [("a", "1."), ("b", "2.")]


def test_blank_plain_paragraph_still_splits_a_list() -> None:
    def setup(slide) -> None:
        frame = slide.shapes.add_textbox(
            Inches(1), Inches(4), Inches(4), Inches(1)
        ).text_frame
        frame.text = "a"
        for text in ["", "b"]:
            frame.add_paragraph().text = text
        for paragraph in (frame.paragraphs[0], frame.paragraphs[2]):
            paragraph._p.get_or_add_pPr().append(
                paragraph._p.makeelement(qn("a:buChar"), {"char": "-"})
            )

    doc = _convert(setup)

    assert len([g for g in doc.groups if g.label == GroupLabel.LIST]) == 2


def test_blank_plain_paragraphs_add_no_text_items() -> None:
    """A blank paragraph in a text box adds no empty text item."""

    def setup(slide) -> None:
        frame = slide.shapes.add_textbox(
            Inches(1), Inches(4), Inches(4), Inches(1)
        ).text_frame
        frame.text = "first"
        for text in ["", "second", "  "]:
            frame.add_paragraph().text = text

    doc = _convert(setup)

    texts = [
        i.text
        for i, _ in doc.iterate_items()
        if isinstance(i, TextItem) and i.text != "Title"
    ]
    assert texts == ["first", "second"]
