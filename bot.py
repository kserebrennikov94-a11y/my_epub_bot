import asyncio
import html
import io
import logging
import os
import re
import tempfile
import threading
import traceback
import uuid
import zipfile
from http.server import BaseHTTPRequestHandler, HTTPServer
from typing import Dict, List, Optional, Tuple

from aiogram import Bot, Dispatcher, F
from aiogram.types import BufferedInputFile, Message
from docx import Document
from docx.document import Document as DocxDocument
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from ebooklib import epub


# ============================================================
# 1. Render Web Service HTTP stub
# ============================================================
class HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-type", "text/plain; charset=utf-8")
        self.end_headers()
        self.wfile.write(b"OK")

    def log_message(self, format, *args):
        return



def run_dummy_server():
    port = int(os.environ.get("PORT", "10000"))
    print(f"HTTP stub starting on port {port}", flush=True)
    server = HTTPServer(("0.0.0.0", port), HealthHandler)
    server.serve_forever()


threading.Thread(target=run_dummy_server, daemon=True).start()


# ============================================================
# 2. Config
# ============================================================
TOKEN = os.environ.get("BOT_TOKEN")
if not TOKEN:
    raise RuntimeError("Переменная окружения BOT_TOKEN не задана")

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

bot = Bot(token=TOKEN)
dp = Dispatcher()

# Временное хранилище обложек
user_data: Dict[int, bytes] = {}

STYLE = """
@namespace epub "http://www.idpf.org/2007/ops";
body {
    font-family: Georgia, serif;
    margin: 5%;
    line-height: 1.45;
    text-align: justify;
}
h1, h2, h3 {
    text-align: center;
    margin-top: 1em;
    margin-bottom: 0.7em;
    line-height: 1.2;
}
h1 { font-size: 1.45em; }
h2 { font-size: 1.20em; }
h3 { font-size: 1.05em; }
p {
    text-indent: 1.5em;
    margin: 0 0 0.45em 0;
}
p.center {
    text-indent: 0;
    text-align: center;
}
p.caption {
    text-indent: 0;
    text-align: center;
    font-style: italic;
    margin-top: 0.35em;
    margin-bottom: 0.7em;
}
img {
    display: block;
    margin: 1em auto;
    max-width: 100%;
    height: auto;
}
table {
    width: 100%;
    border-collapse: collapse;
    margin: 1em 0;
    font-size: 0.92em;
}
th, td {
    border: 1px solid #666;
    padding: 6px;
    vertical-align: top;
}
.pagebreak {
    page-break-before: always;
}
"""

INVALID_XML_RE = re.compile("[" "\x00-\x08" "\x0B\x0C" "\x0E-\x1F" "]")
MULTISPACE_RE = re.compile(r"[ \t]{2,}")
PAGE_MARK_RE = re.compile(r"^\[?\s*Стр\.?\s*\d+\s*\]?$", re.IGNORECASE)
SUBSECTION_RE = re.compile(r"^\d+\.\d+([\.]|\s)")
LECTURE_RE = re.compile(r"^(Лекция|Lecture|Chapter)\s*\d+", re.IGNORECASE)
FIGURE_RE = re.compile(r"^(Рис\.|Рисунок|Схема|Таблица|Figure|Fig\.|Table)\s*\d+", re.IGNORECASE)


# ============================================================
# 3. Helpers
# ============================================================
def sanitize_xml_text(text: str) -> str:
    if not text:
        return ""
    return INVALID_XML_RE.sub("", text)



def normalize_whitespace(text: str) -> str:
    text = text.replace("\xa0", " ")
    text = text.replace("\u00ad", "")
    text = text.replace("\u200b", "")
    text = sanitize_xml_text(text)
    text = MULTISPACE_RE.sub(" ", text)
    return text.strip()



def iter_block_items(parent):
    if isinstance(parent, DocxDocument):
        parent_elm = parent.element.body
    else:
        parent_elm = parent._element

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)



def detect_image_type(blob: bytes) -> Tuple[str, str]:
    if blob.startswith(b"\xff\xd8\xff"):
        return "jpg", "image/jpeg"
    if blob.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png", "image/png"
    if blob[:6] in (b"GIF87a", b"GIF89a"):
        return "gif", "image/gif"
    if blob[:2] == b"BM":
        return "bmp", "image/bmp"
    return "bin", "application/octet-stream"



def collect_images(book: epub.EpubBook, doc: Document) -> Dict[str, Tuple[str, str]]:
    image_map: Dict[str, Tuple[str, str]] = {}
    img_counter = 0

    for rel_id, rel in doc.part.rels.items():
        if "image" not in rel.target_ref:
            continue

        blob = rel.target_part.blob
        ext, media_type = detect_image_type(blob)
        img_counter += 1
        img_name = f"images/img_{img_counter}.{ext}"

        book.add_item(
            epub.EpubItem(
                uid=f"img_{img_counter}",
                file_name=img_name,
                media_type=media_type,
                content=blob,
            )
        )
        image_map[rel_id] = (img_name, media_type)

    return image_map



def extract_inline_images(paragraph: Paragraph, image_map: Dict[str, Tuple[str, str]]) -> List[str]:
    xml = paragraph._p.xml
    found: List[str] = []
    for rel_id, (img_name, _) in image_map.items():
        if rel_id in xml:
            found.append(img_name)
    return found



def has_page_break(paragraph: Paragraph) -> bool:
    xml = paragraph._p.xml
    return ('w:type="page"' in xml) or ("<w:br" in xml and 'type="page"' in xml)



def is_heading_paragraph(paragraph: Paragraph, text: str) -> Optional[int]:
    style_name = (paragraph.style.name or "").strip().lower() if paragraph.style else ""

    style_map = {
        "heading 1": 1,
        "heading 2": 2,
        "heading 3": 3,
        "заголовок 1": 1,
        "заголовок 2": 2,
        "заголовок 3": 3,
        "title": 1,
        "subtitle": 2,
        "название": 1,
        "подзаголовок": 2,
    }
    if style_name in style_map:
        return style_map[style_name]

    if not text:
        return None

    # Heuristics for generic DOCX
    if LECTURE_RE.match(text):
        return 1
    if SUBSECTION_RE.match(text):
        return 2
    if FIGURE_RE.match(text):
        return 3

    # Short centered bold lines often act as headings
    is_center = paragraph.alignment == 1 or str(paragraph.alignment).endswith("CENTER (1)")
    bold_ratio = 0
    total = 0
    for run in paragraph.runs:
        rtxt = run.text or ""
        total += len(rtxt)
        if run.bold:
            bold_ratio += len(rtxt)
    if total and len(text) <= 120 and (bold_ratio / total) > 0.6:
        if is_center:
            return 1
        return 2

    return None



def render_runs(paragraph: Paragraph) -> str:
    parts: List[str] = []
    for run in paragraph.runs:
        txt = normalize_whitespace(run.text or "")
        if not txt:
            continue
        txt = html.escape(txt)
        if run.bold:
            txt = f"<strong>{txt}</strong>"
        if run.italic:
            txt = f"<em>{txt}</em>"
        if run.underline:
            txt = f"<u>{txt}</u>"
        parts.append(txt)
    return "".join(parts).strip()



def paragraph_alignment_class(paragraph: Paragraph) -> str:
    align = paragraph.alignment
    if str(align).endswith("CENTER (1)") or align == 1:
        return "center"
    return ""



def is_noise_paragraph(paragraph: Paragraph, text: str) -> bool:
    if not text:
        return True

    # Generic removal of imported page markers
    if PAGE_MARK_RE.match(text):
        return True

    # Repeated tiny gray header/footer-like lines
    if len(text) < 100:
        style_name = (paragraph.style.name or "").lower() if paragraph.style else ""
        if "header" in style_name or "footer" in style_name or "колонтитул" in style_name:
            return True

        small_font_count = 0
        run_count = 0
        for run in paragraph.runs:
            if run.text and run.text.strip():
                run_count += 1
                size = run.font.size.pt if run.font.size else None
                color = run.font.color.rgb if run.font.color and run.font.color.rgb else None
                if size is not None and size <= 9:
                    small_font_count += 1
                if color is not None and str(color).upper() in {"808080", "7F7F7F", "999999", "A6A6A6"}:
                    return True
        if run_count and small_font_count == run_count and LECTURE_RE.match(text):
            return True

    return False



def table_to_html(table: Table) -> str:
    rows_html: List[str] = []
    for row in table.rows:
        cell_html: List[str] = []
        for cell in row.cells:
            paras: List[str] = []
            for p in cell.paragraphs:
                text = render_runs(p) or html.escape(normalize_whitespace(p.text or ""))
                text = sanitize_xml_text(text).strip()
                if text:
                    paras.append(text)
            value = "<br/>".join(paras) if paras else "&nbsp;"
            cell_html.append(f"<td>{value}</td>")
        rows_html.append("<tr>" + "".join(cell_html) + "</tr>")
    return "<table>" + "".join(rows_html) + "</table>"


# ============================================================
# 4. Generic DOCX -> EPUB
# ============================================================
def build_book_html(doc: Document, image_map: Dict[str, Tuple[str, str]], title: str) -> str:
    body_parts: List[str] = []
    title_added = False

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = normalize_whitespace(block.text or "")
            inline_images = extract_inline_images(block, image_map)

            if is_noise_paragraph(block, text) and not inline_images:
                continue

            if has_page_break(block):
                body_parts.append('<div class="pagebreak"></div>')

            if not text and not inline_images:
                continue

            heading_level = is_heading_paragraph(block, text)

            if heading_level == 1:
                body_parts.append(f"<h1>{html.escape(text)}</h1>")
                title_added = True
            elif heading_level == 2:
                body_parts.append(f"<h2>{html.escape(text)}</h2>")
            elif heading_level == 3:
                body_parts.append(f'<p class="caption">{html.escape(text)}</p>')
            else:
                rendered = render_runs(block)
                if not rendered and text:
                    rendered = html.escape(text)
                if rendered:
                    extra_class = paragraph_alignment_class(block)
                    class_attr = f' class="{extra_class}"' if extra_class else ""
                    body_parts.append(f"<p{class_attr}>{rendered}</p>")

            for img_name in inline_images:
                body_parts.append(f'<img src="{html.escape(img_name)}" alt="Иллюстрация"/>')

        elif isinstance(block, Table):
            body_parts.append(table_to_html(block))

    if not title_added:
        body_parts.insert(0, f"<h1>{html.escape(title)}</h1>")

    if not body_parts:
        body_parts = ["<p>Документ пуст.</p>"]

    return "".join(body_parts)



def create_epub_from_docx_path(docx_path: str, filename: str, cover_image: Optional[bytes] = None) -> bytes:
    if not os.path.exists(docx_path):
        raise RuntimeError("Временный DOCX-файл не найден")
    if os.path.getsize(docx_path) == 0:
        raise RuntimeError("DOCX-файл пустой")
    if not zipfile.is_zipfile(docx_path):
        raise RuntimeError("Файл не является корректным DOCX (ZIP-архивом)")

    doc = Document(docx_path)
    book = epub.EpubBook()
    title = filename.rsplit('.', 1)[0]

    book.set_identifier(str(uuid.uuid4()))
    book.set_title(title)
    book.set_language("ru")

    if cover_image:
        book.set_cover("cover.jpg", cover_image)

    nav_css = epub.EpubItem(
        uid="style_nav",
        file_name="style/nav.css",
        media_type="text/css",
        content=STYLE.encode("utf-8"),
    )
    book.add_item(nav_css)

    image_map = collect_images(book, doc)
    body_html = build_book_html(doc, image_map, title)

    chapter = epub.EpubHtml(title=title, file_name="book.xhtml", lang="ru")
    chapter.set_content(
        f"""<html>
<head><title>{html.escape(title)}</title></head>
<body>{body_html}</body>
</html>""".encode("utf-8")
    )
    chapter.add_item(nav_css)

    book.add_item(chapter)
    book.toc = (chapter,)
    book.add_item(epub.EpubNcx())
    book.spine = [chapter]

    with tempfile.NamedTemporaryFile(delete=False, suffix=".epub") as tmp_epub:
        epub_path = tmp_epub.name

    try:
        epub.write_epub(epub_path, book, {})
        with open(epub_path, "rb") as f:
            epub_bytes = f.read()
        if not epub_bytes:
            raise RuntimeError("EPUB получился пустым")
        return epub_bytes
    finally:
        if os.path.exists(epub_path):
            os.remove(epub_path)


# ============================================================
# 5. Telegram handlers
# ============================================================
@dp.message(F.command("start"))
async def start_cmd(message: Message):
    await message.answer(
        "👋 Привет!\n\n"
        "Пришли обложку, если нужна, а затем DOCX.\n"
        "Я соберу EPUB для обычного DOCX: с текстом, таблицами и изображениями."
    )


@dp.message(F.photo)
async def handle_photo(message: Message):
    photo = message.photo[-1]
    file_info = await bot.get_file(photo.file_id)
    downloaded_file = await bot.download_file(file_info.file_path)
    user_data[message.from_user.id] = downloaded_file.read()
    await message.reply("🖼 Обложка сохранена. Теперь пришли DOCX.")


@dp.message(F.document)
async def handle_docx(message: Message):
    if not message.document.file_name.lower().endswith(".docx"):
        await message.reply("Я принимаю только файлы .docx.")
        return

    status_msg = await message.answer("🚀 Принял файл, начинаю сборку EPUB...")
    temp_path = None

    try:
        buffer = io.BytesIO()
        await bot.download(message.document, destination=buffer)
        buffer.seek(0)
        docx_bytes = buffer.getvalue()

        logger.info("Downloaded file size: %s bytes", len(docx_bytes))
        if not docx_bytes:
            raise RuntimeError("Telegram вернул пустой файл")

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(docx_bytes)
            temp_path = tmp.name

        cover = user_data.get(message.from_user.id)
        epub_data = create_epub_from_docx_path(temp_path, message.document.file_name, cover)

        new_name = message.document.file_name.rsplit('.', 1)[0] + ".epub"
        await message.answer_document(
            BufferedInputFile(epub_data, filename=new_name),
            caption=f"📚 Готово: {new_name}",
        )

        user_data.pop(message.from_user.id, None)
        await status_msg.delete()

    except Exception as exc:
        tb = traceback.format_exc()
        logger.error(tb)
        short_tb = tb[-3500:]
        await message.answer(
            "❌ Ошибка конвертации:\n"
            f"{exc}\n\n"
            "Последние строки traceback:\n"
            f"<pre>{html.escape(short_tb)}</pre>",
            parse_mode="HTML"
        )

    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                logger.warning("Could not remove temporary file: %s", temp_path)


async def main() -> None:
    logger.info("Bot is starting...")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
