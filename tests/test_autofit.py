"""Keep PowerPoint text layout intact until it reaches the renderer.

No Telegram connection or LibreOffice installation is needed for these tests.
Only the external renderer and Telegram calls are replaced; splitting, image
compression, PDF merging and the conversion handler run normally.
"""
import asyncio
import io
import tempfile
import unittest
import zipfile
from contextlib import ExitStack
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, patch

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
from pypdf import PdfWriter

import bot


def set_autofit(frame, scale, reduction):
    body = frame._txBody.find(qn('a:bodyPr'))
    for child in list(body):
        if child.tag in (qn('a:noAutofit'), qn('a:spAutoFit'), qn('a:normAutofit')):
            body.remove(child)
    etree.SubElement(body, qn('a:normAutofit'),
                     fontScale=scale, lnSpcReduction=reduction)


def make_presentation(path):
    prs = Presentation()
    bitmap = io.BytesIO()
    Image.new('RGB', (400, 400), 'navy').save(bitmap, format='BMP')
    for index in range(2):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = f'Slide {index + 1}'
        frame = slide.placeholders[1].text_frame
        frame.text = 'Inherited size: long text that must remain inside its box.'
        # The reported deck inherits its size from the master, not the runs.
        assert frame.paragraphs[0].font.size is None
        assert frame.paragraphs[0].runs[0].font.size is None
        set_autofit(frame, '85000', '20000')

        mixed = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1)).text_frame
        mixed.text = 'Explicit paragraph size'
        mixed.paragraphs[0].font.size = Pt(28)
        run = mixed.paragraphs[0].add_run()
        run.text = ' and a larger bold run'
        run.font.size = Pt(32)
        run.font.bold = True
        mixed.paragraphs[0].line_spacing = 0.9
        set_autofit(mixed, '92500', '10000')

        group = slide.shapes.add_group_shape()
        grouped = group.shapes.add_textbox(Inches(1), Inches(2), Inches(2), Inches(1)).text_frame
        grouped.text = 'Grouped text'
        grouped.paragraphs[0].font.size = Pt(24)
        set_autofit(grouped, '92500', '10000')

        table = slide.shapes.add_table(1, 1, Inches(1), Inches(3), Inches(2), Inches(1)).table
        cell = table.cell(0, 0).text_frame
        cell.text = 'Line-spacing-only shrink'
        cell.paragraphs[0].font.size = Pt(20)
        set_autofit(cell, '100000', '20000')
        bitmap.seek(0)
        slide.shapes.add_picture(bitmap, Inches(4), Inches(1), Inches(1), Inches(1))
    prs.save(path)


def text_bodies(path):
    """Include run sizes, line spacing, end styles and autofit in the oracle."""
    result = []
    with zipfile.ZipFile(path) as z:
        names = sorted(n for n in z.namelist()
                       if n.startswith('ppt/slides/slide') and n.endswith('.xml'))
        for name in names:
            root = etree.fromstring(z.read(name))
            result.append(tuple(etree.tostring(e, method='c14n') for e in root.iter()
                                if e.tag in (qn('p:txBody'), qn('a:txBody'))))
    return result


class AutofitPreservationTests(unittest.TestCase):
    def run_handler(self, *, chunked=False, streaming=False):
        with tempfile.TemporaryDirectory() as temp:
            # The handler owns and deletes its work directory.
            work = Path(temp) / 'job'
            work.mkdir()
            source = work / 'presentation.pptx'
            make_presentation(source)
            expected = text_bodies(source)
            observed = []

            def render(input_path, output_dir):
                bodies = text_bodies(input_path)
                observed.extend(bodies)
                output = Path(output_dir) / (Path(input_path).stem + '.pdf')
                writer = PdfWriter()
                for _ in bodies:
                    writer.add_blank_page(width=720, height=540)
                writer.write(str(output))
                return str(output)

            update = SimpleNamespace(effective_chat=SimpleNamespace(id=1),
                                     message=SimpleNamespace(reply_document=AsyncMock()))
            context = SimpleNamespace(bot=SimpleNamespace(send_chat_action=AsyncMock()))
            status = SimpleNamespace(edit_text=AsyncMock(), delete=AsyncMock())
            with ExitStack() as stack:
                stack.enter_context(patch.object(bot, 'ensure_fonts_available'))
                renderer = stack.enter_context(patch.object(bot, 'convert_pptx_to_pdf', side_effect=render))
                compressor = stack.enter_context(patch.object(bot, 'compress_pptx_images', wraps=bot.compress_pptx_images))
                shrinker = stack.enter_context(patch.object(bot, 'shrink_pptx_streaming', wraps=bot.shrink_pptx_streaming))
                stack.enter_context(patch.object(bot, 'CHUNK_SLIDE_THRESHOLD', 0 if chunked else 40))
                stack.enter_context(patch.object(bot, 'CHUNK_SIZE', 1))
                if streaming:
                    stack.enter_context(patch.object(bot, 'IMAGE_COMPRESS_THRESHOLD_MB', 0))
                asyncio.run(bot.convert_and_reply(update, context, str(source), str(work), status))

                self.assertEqual(observed, expected, 'Text layout changed before rendering')
                self.assertEqual(renderer.call_count, 2 if chunked else 1)
                self.assertEqual(shrinker.call_count, 1 if streaming else 0)
                expected_compressions = (0 if streaming else 2) if chunked else int(streaming)
                self.assertEqual(compressor.call_count, expected_compressions)
                update.message.reply_document.assert_awaited_once()
                status.delete.assert_awaited_once()

    def test_small_presentation_preserves_inherited_and_explicit_autofit(self):
        self.run_handler()

    def test_chunked_presentation_preserves_autofit_after_image_compression(self):
        self.run_handler(chunked=True)

    def test_streaming_compression_preserves_autofit(self):
        self.run_handler(streaming=True)

    def test_streaming_then_chunking_preserves_autofit(self):
        self.run_handler(chunked=True, streaming=True)


if __name__ == '__main__':
    unittest.main()
