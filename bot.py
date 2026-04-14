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

# --- 1. СЕРВЕР-ЗАГЛУШКА ДЛЯ RENDER ---
# Это нужно, чтобы бесплатный Render не отключал бота
def run_dummy_server():
    port = int(os.environ.get("PORT", 10000))
    handler = http.server.SimpleHTTPRequestHandler
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", port), handler) as httpd:
        httpd.serve_forever()

threading.Thread(target=run_dummy_server, daemon=True).start()

# --- 2. НАСТРОЙКИ (ВАШ ТОКЕН ЗДЕСЬ) ---
TOKEN = "8320222564:AAHJ7gvgHGyj8ZBrGsF6d9L-1hvRby0XxXo" # <-- Просто вставьте ваш токен между кавычек

logging.basicConfig(level=logging.INFO)
bot = Bot(token=TOKEN)
dp = Dispatcher()

# Временное хранилище для обложек Владимира
user_data = {}

# Красивое оформление книги
STYLE = '''
@namespace epub "http://www.idpf.org/2007/ops";
body { font-family: "Georgia", serif; margin: 5%; line-height: 1.5; text-align: justify; }
h1, h2 { text-align: center; color: #333; margin-top: 1em; }
p { text-indent: 1.5em; margin-bottom: 0.5em; }
img { display: block; margin: 1em auto; max-width: 100%; }
'''

# --- 3. ЛОГИКА СОЗДАНИЯ КНИГИ ---
def create_epub(docx_bytes, filename, cover_image=None):
    doc = Document(io.BytesIO(docx_bytes))
    book = epub.EpubBook()
    
    title = filename.replace(".docx", "")
    book.set_identifier(str(uuid.uuid4()))
    book.set_title(title)
    book.set_language('ru')

    if cover_image:
        book.set_cover("cover.jpg", cover_image)

    # Добавляем стили
    nav_css = epub.EpubItem(uid="style_nav", file_name="style/nav.css", media_type="text/css", content=STYLE)
    book.add_item(nav_css)

    chapters = []
    current_chapter_title = "Начало"
    current_chapter_content = "<html><body>"
    
    # Обработка картинок внутри текста
    img_counter = 0
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            img_counter += 1
            img_name = f"img_{img_counter}.png"
            epub_img = epub.EpubItem(uid=f"img_{img_counter}", file_name=img_name, 
                                     media_type="image/png", content=rel.target_part.blob)
            book.add_item(epub_img)

    # Разрезаем текст на главы по заголовкам
    for para in doc.paragraphs:
        is_heading = para.style.name.startswith('Heading')
        if is_heading and len(current_chapter_content) > 50:
            current_chapter_content += "</body></html>"
            ch = epub.EpubHtml(title=current_chapter_title, file_name=f'chap_{len(chapters)}.xhtml')
            ch.content = current_chapter_content
            ch.add_item(nav_css)
            book.add_item(ch)
            chapters.append(ch)
            current_chapter_title = para.text
            current_chapter_content = f"<html><body><h1>{para.text}</h1>"
        else:
            if para.text.strip():
                current_chapter_content += f"<p>{para.text}</p>"

    # Последняя глава
    current_chapter_content += "</body></html>"
    ch = epub.EpubHtml(title=current_chapter_title, file_name=f'chap_{len(chapters)}.xhtml')
    ch.content = current_chapter_content
    ch.add_item(nav_css)
    book.add_item(ch)
    chapters.append(ch)

    book.toc = tuple(chapters)
    book.add_item(epub.EpubNav())
    book.add_item(epub.EpubNcx())
    book.spine = ['nav'] + chapters

    out = io.BytesIO()
    epub.write_epub(out, book, {})
    return out.getvalue()

# --- 4. ОБРАБОТЧИКИ ДЛЯ ВЛАДИМИРА ---

@dp.message(F.command("start"))
async def start_cmd(message: Message):
    welcome_text = (
        f"👋 **Здравствуйте, Владимир!**\n\n"
        "Я — ваш персональный ассистент по подготовке книг.\n\n"
        "**Как мы будем работать:**\n"
        "1️⃣ Если нужна **обложка**, просто пришлите мне любую картинку.\n"
        "2️⃣ Затем пришлите файл **.docx**.\n"
        "3️⃣ Я создам красивый .epub с оглавлением по вашим заголовкам.\n\n"
        "✨ *Жду ваш файл или фото.*"
    )
    await message.answer(welcome_text, parse_mode="Markdown")

@dp.message(F.photo)
async def handle_photo(message: Message):
    photo = message.photo[-1]
    file_info = await bot.get_file(photo.file_id)
    downloaded_file = await bot.download_file(file_info.file_path)
    user_data[message.from_user.id] = downloaded_file.read()
    await message.reply("🖼 **Владимир, обложка принята!** Теперь отправляйте основной файл .docx.")

@dp.message(F.document)
async def handle_docx(message: Message):
    if not message.document.file_name.lower().endswith('.docx'):
        await message.reply("Владимир, я работаю только с форматом **.docx**.")
        return

    status_msg = await message.answer("🚀 **Владимир, принял файл!** Начинаю конвертацию...")
    
    try:
        file_io = await bot.download(message.document.file_id)
        cover = user_data.get(message.from_user.id)
        
        epub_data = create_epub(file_io.read(), message.document.file_name, cover)
        
        new_name = message.document.file_name.rsplit('.', 1)[0] + ".epub"
        await message.answer_document(
            BufferedInputFile(epub_data, filename=new_name), 
            caption=f"📚 **Готово!**\nКнига «{new_name.replace('.epub', '')}» упакована."
        )
        
        # Очищаем данные обложки после завершения
        if message.from_user.id in user_data:
            del user_data[message.from_user.id]
        await status_msg.delete()
    except Exception as e:
        logging.error(e)
        await message.answer("❌ Произошла ошибка. Проверьте структуру файла.")

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
