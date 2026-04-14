import asyncio
import html
import imghdr
import io
import logging
import os
import re
import socketserver
import threading
import uuid
import http.server
from typing import Dict, List, Optional, Tuple

from aiogram import Bot, Dispatcher, F
from aiogram.types import BufferedInputFile, Message
from docx import Document
from docx.document import Document as DocxDocument
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from ebooklib import epub


# ============================================================
# 1. Dummy server for Render
# ============================================================
def run_dummy_server() -> None:
    port = int(os.environ.get("PORT", 10000))
    handler = http.server.SimpleHTTPRequestHandler
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", port), handler) as httpd:
        httpd.serve_forever()


threading.Thread(target=run_dummy_server, daemon=True).start()


# ============================================================
# 2. Config
# ============================================================
TOKEN = os.environ.get("BOT_TOKEN", "8320222564:AAHJ7gvgHGyj8ZBrGsF6d9L-1hvRby0XxXo")

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

bot = Bot(token=TOKEN)
dp = Dispatcher()

# temporary storage for covers
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
    margin-top: 1.0em;
    margin-bottom: 0.7em;
    line-height: 1.2;
}
h1 { font-size: 1.45em; }
h2 { font-size: 1.2em; }
h3 { font-size: 1.05em; }
p {
    text-indent: 1.5em;
    margin: 0 0 0.45em 0;
}
p.no-indent {
    text-indent: 0;
}
p.center {
    text-indent: 0;
    text-align: center;
}
.small-gap {
    margin-top: 0.2em;
}
.large-gap {
    margin-top: 1em;
}
.pagebreak {
    page-break-before: always;
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
.caption {
    text-align: center;
    text-indent: 0;
    font-style: italic;
    margin-top: 0.35em;
    margin-bottom: 0.7em;
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
hr {
    border: none;
    border-top: 1px solid #999;
    margin: 1em 0;
}
"""


# ============================================================
# 3. Helpers for reading DOCX in order
# ============================================================
def iter_block_items(parent):
    """Yield Paragraph and Table objects in document order."""
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
# 4. Text cleanup tuned for the reconstructed book
# ============================================================
JUNK_CHARS_RE = re.compile(r"[□■◆◊▪¤�]")
MULTISPACE_RE = re.compile(r"[ \t]{2,}")
BROKEN_LATIN_REPLACEMENTS = {
    "Bыготского": "Выготского",
    "Bопросы": "Вопросы",
    "Cписок": "Список",
    "Cодержание": "Содержание",
    "Пpeдисловие": "Предисловие",
    "Лeкция": "Лекция",
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

    for bad, good in BROKEN_LATIN_REPLACEMENTS.items():
        text = text.replace(bad, good)

    # specific fixes for the reconstructed book
    replacements = {
        "З. Фрейда": "3. Фрейда",
        "Лекция 12 Младенческий возраст": "Лекция 12. Младенческий возраст",
        "Лекция 13 Ранний возраст": "Лекция 13. Ранний возраст",
        "Лекция 14Дошкольный возраст": "Лекция 14. Дошкольный возраст",
        "Лекция 16 Подростковый возраст": "Лекция 16. Подростковый возраст",
        "Лекция 17 Зрелые возрасты": "Лекция 17. Зрелые возрасты",
        "Вопросы для самопроверкиСписок литературы": "Вопросы для самопроверки\nСписок литературы",
    }
    for bad, good in replacements.items():
        text = text.replace(bad, good)

    return normalize_whitespace(text)



def looks_like_garbage(text: str) -> bool:
    if not text:
        return False

    stripped = re.sub(r"\s+", "", text)
    if len(stripped) < 4:
        return False

    # too many repeated same characters, like жжжжжжх
    if re.search(r"(.)\1{4,}", stripped):
        return True

    letters = [ch for ch in stripped if ch.isalpha()]
    if not letters:
        return False

    unique_ratio = len(set(letters)) / max(1, len(letters))
    if len(letters) >= 8 and unique_ratio < 0.25:
        return True

    return False



def clean_heading_text(text: str) -> str:
    text = clean_common_ocr_noise(text)

    # Remove random leading/trailing garbage symbols or single orphan letters
    text = re.sub(r"^[\W_]+", "", text)
    text = re.sub(r"[\W_]+$", "", text)
    text = re.sub(r"\s+[A-Za-zА-Яа-я]$", "", text)

    # Normalize lecture headings
    text = re.sub(r"^Лекция\s+(\d+)\s*[\.-]?\s*", lambda m: f"Лекция {m.group(1)}. ", text)
    text = re.sub(r"\s{2,}", " ", text).strip()

    return text



def is_probable_main_heading(text: str) -> bool:
    text = text.strip()
    if not text:
        return False
    if text in {"Содержание", "Предисловие", "Список литературы", "Вопросы для самопроверки"}:
        return True
    if re.match(r"^Лекция\s+\d+", text):
        return True
    return False



def is_probable_subheading(text: str) -> bool:
    text = text.strip()
    if re.match(r"^\d+\.\d+\.", text):
        return True
    if re.match(r"^\d+\.\d+\s", text):
        return True
    return False


# ============================================================
# 5. Formatting + images
# ============================================================
def detect_image_type(blob: bytes) -> Tuple[str, str]:
    kind = imghdr.what(None, h=blob)
    if kind == "jpeg":
        return "jpg", "image/jpeg"
    if kind == "png":
        return "png", "image/png"
    if kind == "gif":
        return "gif", "image/gif"
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
    return 'w:type="page"' in xml or "<w:br" in xml and 'type="page"' in xml



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
    try:
        align = paragraph.alignment
    except Exception:
        align = None

    # 1 == center in python-docx enum, but keeping it tolerant
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
# 6. EPUB creation specifically tuned for your reconstructed DOCX
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



def build_toc_page(doc: Document) -> str:
    entries: List[str] = []
    in_toc = False

    for block in iter_block_items(doc):
        if not isinstance(block, Paragraph):
            continue

        raw = clean_common_ocr_noise(block.text or "")
        text = normalize_whitespace(raw)
        if not text:
            continue

        if text == "Содержание":
            in_toc = True
            continue

        if in_toc and text == "Предисловие":
            entries.append('<div class="toc-entry">Предисловие</div>')
            continue

        if in_toc and text.startswith("Лекция "):
            entries.append('<div class="toc-entry large-gap"><strong>' + html.escape(clean_heading_text(text)) + '</strong></div>')
            continue

        if in_toc and (is_probable_subheading(text) or text in {"Вопросы для самопроверки", "Список литературы"}):
            entries.append('<div class="toc-entry">' + html.escape(text) + '</div>')
            continue

        # stop when body clearly begins
        if in_toc and text == "Предисловие" and len(entries) > 5:
            break

        if in_toc and len(entries) > 80:
            break

    if not entries:
        return ""

    return '<h1 class="toc-title">Содержание</h1>' + "\n".join(entries)



def create_epub_for_karabanova(docx_bytes: bytes, filename: str, cover_image: Optional[bytes] = None) -> bytes:
    doc = Document(io.BytesIO(docx_bytes))
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

    toc_page = build_toc_page(doc)
    if toc_page:
        toc_ch = make_xhtml("Содержание", toc_page, nav_css)
        book.add_item(toc_ch)
        chapters.append(toc_ch)
        toc_items.append(toc_ch)

    current = ChapterBuffer(title)

    def flush_current() -> None:
        nonlocal current
        if not current.body:
            return

        chapter = make_xhtml(current.title, "\n".join(current.body), nav_css)
        book.add_item(chapter)
        chapters.append(chapter)

        if current.subitems:
            subchapters = []
            for subtitle, anchor in current.subitems:
                sec = epub.Section(current.title)
                # not used directly; keeping subtitle links in tuple form
                subchapters.append(epub.Link(chapter.file_name + anchor, subtitle, anchor.lstrip("#")))
            toc_items.append((epub.Section(current.title), tuple(subchapters)))
        else:
            toc_items.append(chapter)

        current = ChapterBuffer(title)

    started_main_text = False
    seen_toc_heading = False

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            raw_text = block.text or ""
            cleaned_text = clean_common_ocr_noise(raw_text)
            text = normalize_whitespace(cleaned_text)
            heading_text = clean_heading_text(text)
            inline_images = extract_inline_images(block, image_map)

            if text == "Содержание":
                seen_toc_heading = True
                continue

            # skip raw TOC area because we generate our own clean TOC page
            if seen_toc_heading and not started_main_text:
                if text == "Предисловие":
                    seen_toc_heading = False
                else:
                    continue

            if text:
                started_main_text = True

            if has_page_break(block):
                current.body.append('<div class="pagebreak"></div>')

            if not text and not inline_images:
                continue

            # Main chapter headings
            if is_probable_main_heading(heading_text):
                if current.body:
                    flush_current()
                current.title = heading_text
                current.body.append(f"<h1>{html.escape(heading_text)}</h1>")
                continue

            # Subheadings like 1.1., 2.3., etc.
            if is_probable_subheading(heading_text):
                anchor_id = re.sub(r"[^a-zA-Z0-9а-яА-Я]+", "_", heading_text).strip("_")
                current.body.append(f'<h2 id="{html.escape(anchor_id)}">{html.escape(heading_text)}</h2>')
                current.subitems.append((heading_text, f"#{anchor_id}"))
                continue

            # Figure/table captions
            if re.match(r"^(Рис\.|Рисунок|Схема|Таблица)\s*\d+", heading_text):
                current.body.append(f'<p class="caption">{html.escape(heading_text)}</p>')
            else:
                rendered = render_runs(block)
                if not rendered:
                    fallback = heading_text
                    if fallback and not looks_like_garbage(fallback):
                        rendered = html.escape(fallback)

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
        fallback_ch = make_xhtml(title, "<p>Документ пуст.</p>", nav_css)
        book.add_item(fallback_ch)
        chapters.append(fallback_ch)
        toc_items.append(fallback_ch)

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
        "Пришли сначала обложку, если она нужна, а потом файл .docx.\n"
        "Я соберу EPUB, заточенный под твою восстановленную книгу: с более чистыми заголовками, содержанием и нормальной вставкой таблиц."
    )


@dp.message(F.photo)
async def handle_photo(message: Message):
    photo = message.photo[-1]
    file_info = await bot.get_file(photo.file_id)
    downloaded_file = await bot.download_file(file_info.file_path)
    user_data[message.from_user.id] = downloaded_file.read()
    await message.reply("🖼 Обложка сохранена. Теперь отправь DOCX.")


@dp.message(F.document)
async def handle_docx(message: Message):
    if not message.document.file_name.lower().endswith(".docx"):
        await message.reply("Я принимаю только файлы .docx.")
        return

    status_msg = await message.answer("🚀 Принял файл, начинаю сборку EPUB...")

    try:
        file_io = await bot.download(message.document.file_id)
        cover = user_data.get(message.from_user.id)

        epub_data = create_epub_for_karabanova(
            file_io.read(),
            message.document.file_name,
            cover,
        )

        new_name = message.document.file_name.rsplit('.', 1)[0] + ".epub"
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
    if TOKEN == "8320222564:AAHJ7gvgHGyj8ZBrGsF6d9L-1hvRby0XxXo":
        raise RuntimeError("Укажи BOT_TOKEN в переменной окружения или в коде.")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
