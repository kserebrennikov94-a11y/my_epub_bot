import asyncio
import logging
import io
import uuid
from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, BufferedInputFile
from docx import Document
from ebooklib import epub

# --- НАСТРОЙКИ ---
TOKEN = "ВАШ_ТОКЕН_ИЗ_BOTFATHER"

logging.basicConfig(level=logging.INFO)
bot = Bot(token=TOKEN)
dp = Dispatcher()

def convert_docx_to_epub(docx_bytes, filename):
    """Логика конвертации документа в книгу"""
    doc = Document(io.BytesIO(docx_bytes))
    book = epub.EpubBook()
    
    # Метаданные книги
    title = filename.replace(".docx", "")
    book.set_identifier(str(uuid.uuid4()))
    book.set_title(title)
    book.set_language('ru')
    book.add_author('Конвертер Бот')

    # Собираем текст и картинки
    content_html = "<html><body>"
    
    # 1. Извлекаем картинки
    img_counter = 0
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            img_counter += 1
            img_name = f"img_{img_counter}.png"
            epub_img = epub.EpubItem(
                uid=f"img_{img_counter}", 
                file_name=img_name, 
                media_type="image/png", 
                content=rel.target_part.blob
            )
            book.add_item(epub_img)
            content_html += f'<div style="text-align:center;"><img src="{img_name}" style="max-width:100%"/></div>'

    # 2. Извлекаем текст с базовым форматированием
    for para in doc.paragraphs:
        if para.text.strip():
            if para.style.name.startswith('Heading'):
                content_html += f"<h2>{para.text}</h2>"
            else:
                content_html += f"<p>{para.text}</p>"

    content_html += "</body></html>"

    # Создаем главу
    chapter = epub.EpubHtml(title=title, file_name='chapter.xhtml')
    chapter.content = content_html
    book.add_item(chapter)

    # Настройка структуры EPUB
    book.spine = ['nav', chapter]
    book.add_item(epub.EpubNav())
    book.add_item(epub.EpubNcx())

    # Записываем результат в память
    out = io.BytesIO()
    epub.write_epub(out, book, {})
    return out.getvalue()

@dp.message(F.document)
async def handle_docx(message: Message):
    if not message.document.file_name.lower().endswith('.docx'):
        await message.answer("❌ Ошибка: я принимаю только файлы .docx")
        return

    status_msg = await message.answer("⌛ Читаю файл и создаю книгу...")
    
    try:
        # Скачиваем файл во временный буфер
        file_io = await bot.download(message.document.file_id)
        docx_bytes = file_io.read()
        
        # Конвертируем
        epub_data = convert_docx_to_epub(docx_bytes, message.document.file_name)
        
        # Подготавливаем файл для отправки
        new_name = message.document.file_name.rsplit('.', 1)[0] + ".epub"
        final_file = BufferedInputFile(epub_data, filename=new_name)
        
        await message.answer_document(document=final_file, caption="✅ Готово! Приятного чтения.")
        await status_msg.delete()
        
    except Exception as e:
        logging.error(f"Error: {e}")
        await message.answer("🚀 Произошла ошибка при конвертации. Попробуйте другой файл.")

@dp.message()
async def welcome(message: Message):
    await message.answer("Привет! Пришли мне файл в формате **.docx**, и я превращу его в **.epub** для твоей читалки.")

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        pass
