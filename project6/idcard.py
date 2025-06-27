import os
import re
import csv
import fitz  # PyMuPDF
import cv2
import numpy as np
import pandas as pd
import time
import asyncio
import shutil
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime
from dateutil.relativedelta import relativedelta
from openpyxl import Workbook
import asyncio
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from ethiopian_date import EthiopianDateConverter
from telegram.request import HTTPXRequest
from telegram import Update, InputFile
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
from telegram.error import TimedOut, NetworkError
from telegram import InputFile
# from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
# -------------------- Config --------------------
ALLOWED_USERS = [7684557853]
BASE_FOLDER = "project6"
EXTRACTED_FOLDER = os.path.join(BASE_FOLDER, "extracted")
OUTPUT_DIR = os.path.join(BASE_FOLDER, "psd")
TOKEN = "7994714384:AAEPZkHJXrzbPr-wrdHnO-qCXddqAqrqq88"
TIF_TEMPLATE = r"E:/ID/project6/template.tif"
FONT_PATH = r"E:\ID\project6\NotoSansEthiopic_Condensed-Bold.ttf"
BORDER_TEMPLATE_PATH = r"E:/ID/project6/border PNG.png"
FONT_SIZE = 72
TEXT_COLOR = (0, 0, 0)
PHOTO_SIZE = (885, 1171)
QR_SIZE = (1501, 1357)

POSITIONS = {
    "No": (2000, 500), 
    "File": (3184, 1380), 
    "Name": (2000, 510), 
    "ID": (1389, 1490),
    "Gender": (2000, 550), 
    "PhoneNumber": (2989, 600),
    "PhotoPath": (186, 320), 
    "QRCodePath": (3832, 118),
    "IssueDateGC": (3000, 270), 
    "ExpiryDateGC": (3000, 470),
    "IssueDateEC": (3000, 190), 
    "ExpiryDateEC": (3000, 390),
    "Amharic1": (1162, 501), 
    "English1": (2000, 570), 
    "Amharic2": (1162, 581), 
    "English2": (3000, 710),
    "Amharic3": (1158, 748), 
    "English3": (3000, 795), 
    "Amharic4": (1158, 819), 
    "English4": (3000, 918),
    "Amharic5": (1171, 987), 
    "English5": (3000, 992), 
    "Amharic6": (1167, 1063), 
    "English6": (3000, 1128),
    "Amharic7": (2000, 600), 
    "English7": (3000, 1214)
}
FIELD_FONT_SIZES = {
    "No": 72, 
    "File": 72, 
    "Name": 72, 
    "ID": 72, 
    "Gender": 72, 
    "PhoneNumber": 72,
    "IssueDateGC": 72,
    "ExpiryDateGC": 72,
    "IssueDateEC": 72,
    "ExpiryDateEC": 72,
    "Amharic1": 72,
    "English1": 72,
    "Amharic2": 72,
    "English2": 72,
    "Amharic3": 72,
    "English3": 72,
    "Amharic4": 72,
    "English4": 72,
    "Amharic5": 72,
    "English5": 72,
    "Amharic6": 72,
    "English6": 72,
    "Amharic7": 72,
    "English7": 72,
}
NUMERIC_FIELDS_WITH_SPACING = {"File", "ID"}
HIDDEN_FIELDS = ["No", "Name", "Gender", "English1", "Amharic7"]

os.makedirs(BASE_FOLDER, exist_ok=True)
os.makedirs(EXTRACTED_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# -------------------- Drawing Functions --------------------

def paste_image(base, image_path, position, size):
    if pd.isna(image_path) or not os.path.exists(str(image_path)):
        return
    try:
        img = Image.open(image_path).convert("RGBA").resize(size)
        base.paste(img, position, img)
    except Exception as e:
        print(f"❌ Error pasting image {image_path}: {e}")

def draw_text(draw, text, position, field_name="ID", fill=TEXT_COLOR):
    if not text or str(text).lower() in {"nan", ""}:
        return
    size = FIELD_FONT_SIZES.get(field_name, 48)
    font = ImageFont.truetype(FONT_PATH, size)
    if field_name in NUMERIC_FIELDS_WITH_SPACING:
        draw_spaced_text(draw, str(text), position, font, fill, spacing=4)
    else:
        draw.text(position, str(text), font=font, fill=fill)

def draw_spaced_text(draw, text, position, font, fill, spacing=4):
    x, y = position
    for char in text:
        draw.text((x, y), char, font=font, fill=fill)
        bbox = draw.textbbox((0, 0), char, font=font)
        char_width = bbox[2] - bbox[0]
        x += char_width + spacing

def overlay_generated_png_on_border(generated_png_path, border_template_path, output_path, position=(233, 38), size=(2086, 678), flip=False):
    if not os.path.exists(generated_png_path):
        print(f"❌ Generated PNG not found: {generated_png_path}")
        return
    if not os.path.exists(border_template_path):
        print(f"❌ Border template not found: {border_template_path}")
        return
    generated_img = Image.open(generated_png_path).convert("RGBA")
    border_img = Image.open(border_template_path).convert("RGBA")
    if flip:
        generated_img = generated_img.transpose(Image.FLIP_LEFT_RIGHT)
    if size:
        generated_img = generated_img.resize(size, resample=Image.LANCZOS)
    combined = border_img.copy()
    combined.paste(generated_img, position, generated_img)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    combined.save(output_path, format="PNG")
    print(f"✅ Final PNG with border saved at: {output_path}")

def get_resized_font(text, font_path, max_width, initial_size=72, min_size=20):
    """
    Shrinks font size until the text fits within max_width.
    Requires Pillow ≥ 10 for getlength().
    """
    font_size = initial_size
    font = ImageFont.truetype(font_path, font_size)
    try:
        while font.getlength(text) > max_width and font_size > min_size:
            font_size -= 1
            font = ImageFont.truetype(font_path, font_size)
    except AttributeError:
        # For Pillow < 10: fallback using draw.textbbox
        dummy_img = Image.new("RGB", (1, 1))
        dummy_draw = ImageDraw.Draw(dummy_img)
        while dummy_draw.textbbox((0, 0), text, font=font)[2] > max_width and font_size > min_size:
            font_size -= 1
            font = ImageFont.truetype(font_path, font_size)
    return font

# -------------------- ID Card Generation --------------------
def generate_id_card(data, txt_data):
    headers = [
        "No", "File", "Name", "ID", "Gender", "PhoneNumber",
        "PhotoPath", "QRCodePath", "IssueDateGC", "ExpiryDateGC",
        "IssueDateEC", "ExpiryDateEC"
    ] + [f"Amharic{i}" for i in range(1, 8)] + [f"English{i}" for i in range(1, 8)]

    row = dict(zip(headers, data + txt_data))

    base_img = Image.open(TIF_TEMPLATE).convert("RGBA")
    draw = ImageDraw.Draw(base_img)

    # Draw text and place original photo/QR
    for field, pos in POSITIONS.items():
        if field in HIDDEN_FIELDS:
            continue
        if "PhotoPath" in field:
            paste_image(base_img, row.get(field), pos, PHOTO_SIZE)
        elif "QRCodePath" in field:
            paste_image(base_img, row.get(field), pos, QR_SIZE)
        else:
            draw_text(draw, row.get(field), pos, field_name=field, fill=TEXT_COLOR)

    # 🧍 ID name horizontal
    # 🧍 ID name horizontal with auto-resizing
    name = row.get("ID", "")
    if name and str(name).lower() != "nan":
        max_width = 500  # Adjust this to fit your layout
        resized_font = get_resized_font(name, FONT_PATH, max_width, initial_size=72)
        draw.text((316, 1600), name, font=resized_font, fill=(0, 0, 0))

        # ➕ Clone and rotate IssueDateEC text
        issue_date_ec = row.get("IssueDateEC")
        if issue_date_ec and str(issue_date_ec).lower() not in {"nan", ""}:
            try:
                rotated_text_img = Image.new("RGBA", (500, 80), (255, 255, 255, 0))  # canvas for rotated text
                rotate_draw = ImageDraw.Draw(rotated_text_img)
                rotate_font = ImageFont.truetype(FONT_PATH, FIELD_FONT_SIZES.get("IssueDateEC", 72))
                rotate_draw.text((0, 0), str(issue_date_ec), font=rotate_font, fill=TEXT_COLOR)
                rotated_result = rotated_text_img.rotate(90, expand=1)
                base_img.paste(rotated_result, (75, 460), rotated_result)
            except Exception as e:
                print(f"❌ Error drawing rotated IssueDateEC: {e}")

        # 👯 Reuse & paste cloned photo with rounded edges + blur
        if row.get("PhotoPath") and os.path.exists(row["PhotoPath"]):
            try:
                # Open and resize original
                original_photo = Image.open(row["PhotoPath"]).convert("RGBA").resize(PHOTO_SIZE)

                # Clone and resize
                clone_size = (230, 280)  # Width x Height of clone
                clone_position = (2350, 1397)  # Target position on ID card
                clone_photo = original_photo.resize(clone_size, resample=Image.LANCZOS)

                # Add blur to the clone
                from PIL import ImageFilter
                clone_photo = clone_photo.filter(ImageFilter.GaussianBlur(radius=2))  # Slight blur

                # Create rounded mask
                mask = Image.new("L", clone_size, 0)
                draw_mask = ImageDraw.Draw(mask)
                draw_mask.rounded_rectangle([(0, 0), clone_size], radius=60, fill=255)

                # Apply mask to clone
                clone_photo.putalpha(mask)

                # Paste the clone on card
                base_img.paste(clone_photo, clone_position, clone_photo)
            except Exception as e:
                print(f"❌ Error cloning and processing photo: {e}")

        # 📤 Save PNG and border version
        name_clean = str(row.get("Name", "")).strip()
        name_value = name_clean.replace(" ", "_").replace("/", "_").replace("\\", "_") or f"id_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

        png_out = os.path.join(OUTPUT_DIR, f"{name_value}.png")
        flipped_png = base_img.transpose(Image.FLIP_LEFT_RIGHT)
        flipped_png.save(png_out, format="PNG")
        print(f"✅ Saved PNG: {png_out}")

        final_out = os.path.join(OUTPUT_DIR, f"{name_value}_with_border.png")
        overlay_generated_png_on_border(
            generated_png_path=png_out,
            border_template_path=BORDER_TEMPLATE_PATH,
            output_path=final_out,
            position=(229, 25),
            size=(2090, 685),
            flip=False
        )
        print(f"✅ Final PNG with border saved at: {final_out}")

        return png_out, final_out


# -------------------- CSV and Excel Handling --------------------

# -------------------- PDF Extraction --------------------
def is_name(text):
    return re.fullmatch(r'([A-Z][a-z]+(?: [A-Z][a-z]+){0,3})', text) is not None

def extract_txt_row_data(txt_path, max_pairs=7):
    row_data = []
    with open(txt_path, 'r', encoding='utf-8') as f:
        for line in f:
            if "|" in line:
                am, en = map(str.strip, line.strip().split("|", 1))
                row_data.extend([am, en])
            if len(row_data) >= max_pairs * 2:
                break
    while len(row_data) < max_pairs * 2:
        row_data.append("")
    return row_data

def extract_from_pdf(pdf_path, output_dir, number, original_filename):
    doc = fitz.open(pdf_path)
    page = doc[0]
    base_name = os.path.splitext(original_filename)[0]
    text = page.get_text()
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    id_number = next((m for m in re.findall(r'\b\d{10,18}\b', text)), "N/A")
    name = next((line.split(":")[-1].strip() for line in lines if 'name' in line.lower()), "N/A")
    if name == "N/A":
        name = next((line.strip() for line in lines if is_name(line)), "N/A")

    gender = re.search(r'\b(Male|Female|M|F)\b', text, re.IGNORECASE)
    gender = gender.group(1).capitalize() if gender else "N/A"
    phone = re.search(r'\b\d{10,15}\b', text)
    phone = phone.group(0) if phone else "N/A"

    text_path = os.path.join(output_dir, f"{base_name}_info.txt")
    with open(text_path, "w", encoding="utf-8") as f:
        f.write(f"ID Number: {id_number}\n\n{text}")

    photo_path = qr_path = "Not Found"
    for img_index, img in enumerate(page.get_images(full=True)):
        xref = img[0]
        img_data = doc.extract_image(xref)
        img_np = np.frombuffer(img_data["image"], np.uint8)
        img_cv = cv2.imdecode(img_np, cv2.IMREAD_COLOR)
        img_pil = Image.fromarray(cv2.cvtColor(img_cv, cv2.COLOR_BGR2RGB))
        img_path = os.path.join(output_dir, f"{base_name}_img_{img_index}.jpg")
        img_pil.save(img_path)
        if img_index == 1:
            photo_path = img_path
        elif img_index == 2:
            qr_path = img_path

    today_gc = datetime.today()
    expiry_gc = today_gc + relativedelta(years=5)
    today_ec = EthiopianDateConverter.to_ethiopian(today_gc.year, today_gc.month, today_gc.day)
    expiry_ec = EthiopianDateConverter.to_ethiopian(expiry_gc.year, expiry_gc.month, expiry_gc.day)

    return [
        number, base_name, name, id_number, gender, phone,
        photo_path, qr_path,
        today_gc.strftime("%Y/%m/%d"), expiry_gc.strftime("%Y/%m/%d"),
        f"{today_ec[0]:04d}/{today_ec[1]:02d}/{today_ec[2]:02d}",
        f"{expiry_ec[0]:04d}/{expiry_ec[1]:02d}/{expiry_ec[2]:02d}"
    ], text_path

def process_single_pdf(pdf_path, output_dir, number, original_filename):
    os.makedirs(output_dir, exist_ok=True)
    pdf_data, txt_path = extract_from_pdf(pdf_path, output_dir, number, original_filename)
    txt_data = extract_txt_row_data(txt_path) if os.path.exists(txt_path) else [""] * 14
    flipped_png_path, final_border_path = generate_id_card(pdf_data, txt_data)
    
    return flipped_png_path, final_border_path

# -------------------- Telegram Bot --------------------


# -------------------- UTILITY --------------------
def is_allowed(user_id: int) -> bool:
    return user_id in ALLOWED_USERS

def compress_png(path):
    try:
        img = Image.open(path)
        img.save(path, optimize=True)
        print(f"✅ Compressed PNG: {path}")
    except Exception as e:
        print(f"❌ Failed to compress PNG {path}: {e}")


# -------------------- BOT COMMANDS --------------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id):
        await update.message.reply_text("⛔ Access denied.")
        return
    await update.message.reply_text("Send your ID (as text), then send your PDF.")

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        await update.message.reply_text("⛔ Access denied.")
        return
    folder = os.path.join(BASE_FOLDER, str(user_id))
    os.makedirs(folder, exist_ok=True)
    with open(os.path.join(folder, "id.txt"), "w", encoding="utf-8") as f:
        f.write(update.message.text.strip())
    await update.message.reply_text("✅ ID received. Now send the PDF.")

async def reject_octet_stream(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id):
        await update.message.reply_text("⛔ Access denied.")
        return
    await update.message.reply_text("❌ Unsupported file type. Send a PDF.")

import asyncio
import shutil

async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        await update.message.reply_text("⛔ Access denied.")
        return

    doc = update.message.document
    if not doc.file_name.lower().endswith(".pdf"):
        await update.message.reply_text("❌ Only PDF files are supported.")
        return

    user_folder = os.path.join(BASE_FOLDER, str(user_id))
    os.makedirs(user_folder, exist_ok=True)
    uploaded_path = os.path.join(user_folder, doc.file_name)
    telegram_file = await doc.get_file()
    await telegram_file.download_to_drive(uploaded_path)

    msg = await update.message.reply_text("📄 Processing your ID...")

    try:
        output_dir = os.path.join(EXTRACTED_FOLDER, str(user_id))
        os.makedirs(output_dir, exist_ok=True)

        # Process the PDF and get PNG file paths
        flipped_png, border_png = process_single_pdf(uploaded_path, output_dir, 2, doc.file_name)

        # Function to send files with retries and delay
        async def send_file_with_retry(file_path):
            if file_path and os.path.exists(file_path):
                try:
                    with open(file_path, "rb") as f:
                        await context.bot.send_document(chat_id=update.effective_chat.id, document=f)
                    await asyncio.sleep(1)  # Short delay after sending
                except Exception as e:
                    print(f"❌ Failed to send {file_path}: {e}")
                    await update.message.reply_text(f"❌ Could not send {os.path.basename(file_path)}")
            else:
                await update.message.reply_text(f"❌ File not found: {os.path.basename(file_path) if file_path else 'Unknown'}")

        # Send both PNG files
        await send_file_with_retry(flipped_png)
        await send_file_with_retry(border_png)

    except Exception as e:
        print(f"❌ Processing error: {e}")
        await update.message.reply_text("❌ Failed to process the PDF.")
        return  # stop here if error

    await msg.delete()

    # Wait 2 minutes before cleanup
    await asyncio.sleep(120)

    # Cleanup folders: user_folder (PDF), output_dir (extracted), and PNG folder (OUTPUT_DIR)
    folders_to_delete = [user_folder, output_dir, OUTPUT_DIR]

    for folder in folders_to_delete:
        if os.path.exists(folder):
            try:
                shutil.rmtree(folder)
                print(f"🧹 Deleted folder: {folder}")
            except Exception as e:
                print(f"❌ Could not delete {folder}: {e}")


# -------------------- MAIN --------------------
def main():
   

    request = HTTPXRequest(
        connect_timeout=120.0,
        read_timeout=120.0,
        write_timeout=180.0,
        pool_timeout=120.0
    )
    app = ApplicationBuilder().token(TOKEN).request(request).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))
    app.add_handler(MessageHandler(filters.Document.ALL, reject_octet_stream))

    print("Bot is running...")
    app.run_polling()

if __name__ == "__main__":
    main()
