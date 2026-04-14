import asyncio
import html
import io
import logging
import os
import re
import threading
import uuid
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
.toc-title {
    text-align: center;
    font-size: 1.4em;
    margin-bottom: 1em;
}
.toc-entry {
    text-indent: 0;
    margin: 0.15em 0;
}
"""


# ============================================================
# 3. DOCX traversal helpers
# ============================================================
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


# ============================================================
# 4. Cleanup helpers
# ============================================================
JUNK_CHARS_RE = re.compile(r"[□■◆◊▪¤�]")
MULTISPACE_RE = re.compile(r"[ \t]{2,}")

BROKEN_REPLACEMENTS = {
    "Bыготского": "Выготского",
    "Bопросы": "Вопросы",
    "Cписок": "Список",
    "Cодержание": "Содержание",
    "Пpeдисловие": "Предисловие",
    "Лeкция": "Лекция",
    "З. Фрейда": "3. Фрейда",
    "Лекция 12 Младенческий возраст": "Лекция 12. Младенческий возраст",
    "Лекция 13 Ранний возраст": "Лекция 13. Ранний возраст",
    "Лекция 14Дошкольный возраст": "Лекция 14. Дошкольный возраст",
    "Лекция 16 Подростковый возраст": "Лекция 16. Подростковый возраст",
    "Лекция 17 Зрелые возрасты": "Лекция 17. Зрелые возрасты",
    "Вопросы для самопроверкиСписок литературы": "Вопросы для самопроверки\nСписок литературы",
}


def normalize_whitespace(text: str) -> str:
    text = text.replace("\xa0", " ")
    text = text.replace("\u00ad", "")
    text = text.replace("\u200b", "")
    text = MULTISPACE_RE.sub(" ", text)
    return text.strip()


def clean_common_ocr_noise(text: str) -> str:
    if not text:
        return ""

    text = JUNK_CHARS_RE.sub("", text)
    text = text.replace(" ,", ",")
    text = text.replace(" .", ".")
    text = text.replace(" :", ":")
    text = text.replace(" ;", ";")
    text = text.replace("( ", "(")
    text = text.replace(" )", ")")

    for bad, good in BROKEN_REPLACEMENTS.items():
        text = text.replace(bad, good)

    return normalize_whitespace(text)


def looks_like_garbage(text: str) -> bool:
    stripped = re.sub(r"\s+", "", text or "")
    if len(stripped) < 4:
        return False
    if re.search(r"(.)\1{4,}", stripped):
        return True

    letters = [ch for ch in stripped if ch.isalpha()]
    if not letters:
        return False

    unique_ratio = len(set(letters)) / max(1, len(letters))
    return len(letters) >= 8 and unique_ratio < 0.25


def clean_heading_text(text: str) -> str:
    text = clean_common_ocr_noise(text)
    text = re.sub(r"^[\W_]+", "", text)
    text = re.sub(r"[\W_]+$", "", text)
    text = re.sub(r"\s+[A-Za-zА-Яа-я]$", "", text)
    text = re.sub(r"^Лекция\s+(\d+)\s*[\.-]?\s*", lambda m: f"Лекция {m.group(1)}. ", text)
    text = re.sub(r"\s{2,}", " ", text).strip()
    return text


def is_probable_main_heading(text: str) -> bool:
    text = (text or "").strip()
    if not text:
        return False
    if text in {"Содержание", "Предисловие", "Список литературы", "Вопросы для самопроверки"}:
        return True
    return bool(re.match(r"^Лекция\s+\d+", text))


def is_probable_subheading(text: str) -> bool:
    text = (text or "").strip()
    return bool(re.match(r"^\d+\.\d+([\.]|\s)", text))


# ============================================================
# 5. Images and tables
# ============================================================
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


def render_runs(paragraph: Paragraph) -> str:
    parts: List[str] = []

    for run in paragraph.runs:
        txt = run.text or ""
        if not txt:
            continue

        txt = clean_common_ocr_noise(txt)
        txt = html.escape(txt)
        if not txt:
            continue

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


def table_to_html(table: Table) -> str:
    rows_html: List[str] = []

    for row in table.rows:
        cell_html: List[str] = []

        for cell in row.cells:
            paras: List[str] = []

            for p in cell.paragraphs:
                text = render_runs(p) or html.escape(clean_common_ocr_noise(p.text or ""))
                text = text.strip()
                if text:
                    paras.append(text)

            value = "<br/>".join(paras) if paras else "&nbsp;"
            cell_html.append(f"<td>{value}</td>")

        rows_html.append("<tr>" + "".join(cell_html) + "</tr>")

    return "<table>" + "".join(rows_html) + "</table>"


# ============================================================
# 6. EPUB builder
# ============================================================
class ChapterBuffer:
    def __init__(self, title: str):
        self.title = title
        self.body: List[str] = []
        self.subitems: List[Tuple[str, str]] = []


def make_xhtml(title: str, body_html: str, css_item: epub.EpubItem) -> epub.EpubHtml:
    ch = epub.EpubHtml(
        title=title,
        file_name=f"text/{uuid.uuid4().hex}.xhtml",
        lang="ru",
    )
    ch.content = f"""<?xml version='1.0' encoding='utf-8'?>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>{html.escape(title)}</title>
</head>
<body>
{body_html}
</body>
</html>
"""
    ch.add_item(css_item)
    return ch


def build_clean_toc_page(doc: Document) -> str:
    entries: List[str] = []
    in_toc = False

    for block in iter_block_items(doc):
        if not isinstance(block, Paragraph):
            continue

        text = normalize_whitespace(clean_common_ocr_noise(block.text or ""))
        if not text:
            continue

        if text == "Содержание":
            in_toc = True
            continue

        if not in_toc:
            continue

        if text.startswith("Лекция "):
            entries.append(f'<div class="toc-entry"><strong>{html.escape(clean_heading_text(text))}</strong></div>')
            continue

        if text == "Предисловие" or is_probable_subheading(text) or text in {"Вопросы для самопроверки", "Список литературы"}:
            entries.append(f'<div class="toc-entry">{html.escape(text)}</div>')
            continue

        if len(entries) > 20 and text == "Предисловие":
            break

        if len(entries) > 120:
            break

    if not entries:
        return ""

    return '<h1 class="toc-title">Содержание</h1>' + "\n".join(entries)


def create_epub_for_karabanova(docx_bytes: bytes, filename: str, cover_image: Optional[bytes] = None) -> bytes:
    if not docx_bytes:
        raise RuntimeError("DOCX-файл пустой")

    try:
        doc = Document(io.BytesIO(docx_bytes))
    except Exception as e:
        raise RuntimeError(f"Не удалось прочитать DOCX: {e}")

    book = epub.EpubBook()

    title = filename.rsplit(".", 1)[0]
    book.set_identifier(str(uuid.uuid4()))
    book.set_title(title)
    book.set_language("ru")
    book.add_author("О. А. Карабанова")

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

    chapters: List[epub.EpubHtml] = []
    toc_items: List = []

    toc_page_html = build_clean_toc_page(doc)
    if toc_page_html:
        toc_page = make_xhtml("Содержание", toc_page_html, nav_css)
        book.add_item(toc_page)
        chapters.append(toc_page)
        toc_items.append(toc_page)

    current = ChapterBuffer(title)
    skipping_raw_toc = False

    def flush_current() -> None:
        nonlocal current

        if not current.body:
            return

        chapter = make_xhtml(current.title, "\n".join(current.body), nav_css)
        book.add_item(chapter)
        chapters.append(chapter)

        if current.subitems:
            sublinks = []
            for subtitle, anchor in current.subitems:
                sublinks.append(epub.Link(chapter.file_name + anchor, subtitle, anchor.lstrip("#")))
            toc_items.append((epub.Section(current.title), tuple(sublinks)))
        else:
            toc_items.append(chapter)

        current = ChapterBuffer(title)

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            raw_text = block.text or ""
            text = normalize_whitespace(clean_common_ocr_noise(raw_text))
            heading_text = clean_heading_text(text)
            inline_images = extract_inline_images(block, image_map)

            if text == "Содержание":
                skipping_raw_toc = True
                continue

            if skipping_raw_toc:
                if text == "Предисловие":
                    skipping_raw_toc = False
                else:
                    continue

            if has_page_break(block):
                current.body.append('<div class="pagebreak"></div>')

            if not text and not inline_images:
                continue

            if is_probable_main_heading(heading_text):
                if current.body:
                    flush_current()
                current.title = heading_text
                current.body.append(f"<h1>{html.escape(heading_text)}</h1>")
                continue

            if is_probable_subheading(heading_text):
                anchor_id = re.sub(r"[^a-zA-Z0-9а-яА-Я]+", "_", heading_text).strip("_")
                current.body.append(f'<h2 id="{html.escape(anchor_id)}">{html.escape(heading_text)}</h2>')
                current.subitems.append((heading_text, f"#{anchor_id}"))
                continue

            if re.match(r"^(Рис\.|Рисунок|Схема|Таблица)\s*\d+", heading_text):
                current.body.append(f'<p class="caption">{html.escape(heading_text)}</p>')
            else:
                rendered = render_runs(block)
                if not rendered and heading_text and not looks_like_garbage(heading_text):
                    rendered = html.escape(heading_text)

                if rendered:
                    extra_class = paragraph_alignment_class(block)
                    class_attr = f' class="{extra_class}"' if extra_class else ""
                    current.body.append(f"<p{class_attr}>{rendered}</p>")

            for img_name in inline_images:
                current.body.append(f'<img src="{html.escape(img_name)}" alt="Иллюстрация"/>')

        elif isinstance(block, Table):
            current.body.append(table_to_html(block))

    flush_current()

    if not chapters:
        fallback = make_xhtml(title, "<p>Документ пуст.</p>", nav_css)
        book.add_item(fallback)
        chapters.append(fallback)
        toc_items.append(fallback)

    book.toc = tuple(toc_items)
    book.add_item(epub.EpubNav())
    book.add_item(epub.EpubNcx())
    book.spine = ["nav"] + chapters

    out = io.BytesIO()
    epub.write_epub(out, book, {})
    return out.getvalue()


# ============================================================
# 7. Telegram handlers
# ============================================================
@dp.message(F.command("start"))
async def start_cmd(message: Message):
    await message.answer(
        "👋 Привет!\n\n"
        "Пришли обложку, если нужна, а затем DOCX.\n"
        "Я соберу EPUB под твою восстановленную книгу: "
        "с более чистыми заголовками, содержанием, таблицами и картинками."
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

    try:
        file_info = await bot.get_file(message.document.file_id)

        buffer = io.BytesIO()
        await bot.download_file(file_info.file_path, destination=buffer)
        buffer.seek(0)
        docx_bytes = buffer.getvalue()

        if not docx_bytes:
            raise RuntimeError("Telegram вернул пустой файл")

        cover = user_data.get(message.from_user.id)
        epub_data = create_epub_for_karabanova(docx_bytes, message.document.file_name, cover)

        new_name = message.document.file_name.rsplit(".", 1)[0] + ".epub"
        await message.answer_document(
            BufferedInputFile(epub_data, filename=new_name),
            caption=f"📚 Готово: {new_name}",
        )

        user_data.pop(message.from_user.id, None)
        await status_msg.delete()

    except Exception as exc:
        logger.exception("Conversion error")
        await message.answer(f"❌ Ошибка конвертации: {exc}")


async def main() -> None:
    logger.info("Bot is starting...")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
