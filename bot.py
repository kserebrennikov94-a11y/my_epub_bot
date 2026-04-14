import asyncio
import logging
import io
import uuid
import threading
import http.server
import socketserver
import os
from aiogram import Bot, Dispatcher, F, types
from aiogram.types import Message, BufferedInputFile
from docx import Document
from ebooklib import epub

# --- СЕРВЕР-ЗАГЛУШКА ДЛЯ RENDER ---
def run_dummy_server():
    port = int(os.environ.get("PORT", 10000))
    handler = http.server.SimpleHTTPRequestHandler
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", port), handler) as httpd:
        httpd.serve_forever()

threading.Thread(target=run_dummy_server, daemon=True).start()

# --- НАСТРОЙКИ ---
TOKEN = "8320222564:AAHJ7gvgHGyj8ZBrGsF6d9L-1hvRby0XxXo"
logging.basicConfig(level=logging.INFO)
bot = Bot(token=TOKEN)
dp = Dispatcher()

# Хранилище для временных данных (в продакшене лучше использовать БД или Redis)
user_data = {}

# Красивый CSS для книги
STYLE = '''
@namespace epub "http://www.idpf.org/2007/ops";
body {
    font-family: "Georgia", serif;
    margin: 5%;
    line-height: 1.5;
    text-align: justify;
}
h1, h2 {
    text-align: center;
    color: #333;
    margin-top: 1em;
}
p {
    text-indent: 1.5em; /* Красная строка */
    margin-bottom: 0.5em;
}
img {
    display: block;
    margin: 1em auto;
    max-width: 100%;
}
'''

def create_epub(docx_bytes, filename, cover_image=None):
    doc = Document(io.BytesIO(docx_bytes))
    book = epub.EpubBook()
    
    title = filename.replace(".docx", "")
    book.set_identifier(str(uuid.uuid4()))
    book.set_title(title)
    book.set_language('ru')

    if cover_image:
        book.set_cover("cover.jpg", cover_image)

    # Ресурс со стилями
    nav_css = epub.EpubItem(uid="style_nav", file_name="style/nav.css", media_type="text/css", content=STYLE)
    book.add_item(nav_css)

    chapters = []
    current_chapter_title = "Начало"
    current_chapter_content = "<html><body>"
    
    # Обработка изображений (собираем их заранее)
    img_counter = 0
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            img_counter += 1
            img_name = f"img_{img_counter}.png"
            epub_img = epub.EpubItem(uid=f"img_{img_counter}", file_name=img_name, 
                                     media_type="image/png", content=rel.target_part.blob)
            book.add_item(epub_img)

    # Проходим по параграфам и разбиваем на главы
    for para in doc.paragraphs:
        # Проверяем, является ли параграф заголовком ( Heading 1 или Heading 2)
        is_heading = para.style.name.startswith('Heading')
        
        if is_heading and len(current_chapter_content) > 20: # Если встретили заголовок и старая глава не пуста
            # Закрываем старую главу
            current_chapter_content += "</body></html>"
            ch = epub.EpubHtml(title=current_chapter_title, file_name=f'chap_{len(chapters)}.xhtml')
            ch.content = current_chapter_content
            ch.add_item(nav_css)
            book.add_item(ch)
            chapters.append(ch)
            
            # Начинаем новую главу
            current_chapter_title = para.text
            current_chapter_content = f"<html><body><h1>{para.text}</h1>"
        else:
            # Просто добавляем текст
            if para.text.strip():
                current_chapter_content += f"<p>{para.text}</p>"

    # Добавляем последнюю главу
    current_chapter_content += "</body></html>"
    ch = epub.EpubHtml(title=current_chapter_title, file_name=f'chap_{len(chapters)}.xhtml')
    ch.content = current_chapter_content
    ch.add_item(nav_css)
    book.add_item(ch)
    chapters.append(ch)

    # Настраиваем структуру
    book.toc = tuple(chapters)
    book.add_item(epub.EpubNav())
    book.add_item(epub.EpubNcx())
    book.spine = ['nav'] + chapters

    out = io.BytesIO()
    epub.write_epub(out, book, {})
    return out.getvalue()

    # Добавляем обложку, если она есть
    if cover_image:
        book.set_cover("cover.jpg", cover_image)

    # Контент
    content_html = "<html><body>"
    
    # Обработка изображений из Word
    img_counter = 0
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            img_counter += 1
            img_name = f"img_{img_counter}.png"
            epub_img = epub.EpubItem(uid=f"img_{img_counter}", file_name=img_name, 
                                     media_type="image/png", content=rel.target_part.blob)
            book.add_item(epub_img)
            content_html += f'<img src="{img_name}"/>'

    # Текст
    for para in doc.paragraphs:
        if para.text.strip():
            if para.style.name.startswith('Heading'):
                content_html += f"<h2>{para.text}</h2>"
            else:
                content_html += f"<p>{para.text}</p>"

    content_html += "</body></html>"

    chapter = epub.EpubHtml(title=title, file_name='chapter.xhtml')
    chapter.content = content_html
    book.add_item(chapter)

    # Добавляем CSS
    nav_css = epub.EpubItem(uid="style_nav", file_name="style/nav.css", media_type="text/css", content=STYLE)
    book.add_item(nav_css)
    chapter.add_item(nav_css)

    book.spine = ['nav', chapter]
    book.add_item(epub.EpubNav())
    book.add_item(epub.EpubNcx())

    out = io.BytesIO()
    epub.write_epub(out, book, {})
    return out.getvalue()

@dp.message(F.photo)
async def handle_photo(message: Message):
    # Сохраняем последнее присланное фото как обложку
    photo = message.photo[-1]
    file_info = await bot.get_file(photo.file_id)
    downloaded_file = await bot.download_file(file_info.file_path)
    user_data[message.from_user.id] = downloaded_file.read()
    await message.reply("🖼 Обложка сохранена! Теперь пришлите файл .docx")

@dp.message(F.document)
async def handle_docx(message: Message):
    if not message.document.file_name.lower().endswith('.docx'):
        return

    status_msg = await message.answer("⌛ Магия форматирования...")
    
    try:
        file_io = await bot.download(message.document.file_id)
        cover = user_data.get(message.from_user.id) # Берем обложку, если пользователь её прислал ранее
        
        epub_data = create_epub(file_io.read(), message.document.file_name, cover)
        
        new_name = message.document.file_name.rsplit('.', 1)[0] + ".epub"
        await message.answer_document(BufferedInputFile(epub_data, filename=new_name), caption="✨ Книга готова!")
        
        # Очищаем обложку после использования
        if message.from_user.id in user_data:
            del user_data[message.from_user.id]
            
        await status_msg.delete()
    except Exception as e:
        logging.error(e)
        await message.answer("❌ Ошибка при создании книги.")

@dp.message()
async def start(message: Message):
    await message.answer("📚 **Привет! Я сделаю твою книгу красивой.**\n\n"
                         "1. Пришли картинку (это будет обложка).\n"
                         "2. Пришли файл .docx.\n\n"
                         "Если картинку не прислать, книга будет без обложки.")

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
