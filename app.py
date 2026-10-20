import streamlit as st
import os
import zipfile
import rarfile
import re
import uuid
import shutil
from pathlib import Path
from datetime import datetime
from PIL import Image, ImageEnhance, ImageFilter
import pytesseract
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Настройка страницы
st.set_page_config(page_title="Анализатор табличек с адресами", page_icon="", layout="wide")

# Создаем папки
os.makedirs('uploads', exist_ok=True)
os.makedirs('results', exist_ok=True)

def preprocess_image(img):
    """Предобработка изображения для OCR"""
    img = ImageEnhance.Contrast(img).enhance(2.0)
    img = img.filter(ImageFilter.SHARPEN).convert('L')
    return img.resize((img.size[0]*2, img.size[1]*2), Image.Resampling.LANCZOS)

def create_thumbnail(img_path, output_path, size=(300, 300)):
    """Создание миниатюры"""
    try:
        with Image.open(img_path) as img:
            img.thumbnail(size, Image.Resampling.LANCZOS)
            img.save(output_path, 'JPEG', quality=80)
        return True
    except:
        return False

def analyze_image(image_path):
    """Анализ изображения на наличие адреса"""
    try:
        img = preprocess_image(Image.open(image_path))
        text = pytesseract.image_to_string(img, config='--oem 3 --psm 6 -l rus+eng').lower()
        
        patterns = [
            r'ул\.\s*[\w\-]+',
            r'улица\s+[\w\-]+',
            r'пр\.\s*[\w\-]+',
            r'проспект\s+[\w\-]+',
            r'пер\.\s*[\w\-]+',
            r'переулок\s+[\w\-]+',
            r'д\.\s*\d+[\w\-/]*',
            r'дом\s+\d+[\w\-/]*'
        ]
        
        found = [m for p in patterns for m in re.findall(p, text)]
        
        if found:
            unique_found = list(dict.fromkeys(found))
            return True, f"Найдено: {', '.join(unique_found[:5])}"
        return False, "Адрес не распознан или недостаточно чёткий"
    except Exception as e:
        return False, f"Ошибка: {str(e)}"

def make_excel_pretty(excel_path):
    """Создание красивого Excel файла"""
    wb = load_workbook(excel_path)
    ws = wb.active
    ws.title = "Результаты"
    
    # Стили
    hf = Font(bold=True, color='FFFFFF', size=11)
    hfill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    yfill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    yfont = Font(color='006100', bold=True)
    nfill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    nfont = Font(color='9C0006', bold=True)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Заголовки
    for cell in ws[1]:
        cell.font = hf
        cell.fill = hfill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border
    
    # Данные
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical='center', wrap_text=True)
        
        if row[1].value == 'Да':
            row[1].fill = yfill
            row[1].font = yfont
        elif row[1].value == 'Нет':
            row[1].fill = nfill
            row[1].font = nfont
    
    # Ширина колонок
    for col_idx, column in enumerate(ws.columns, 1):
        max_length = max((len(str(c.value)) if c.value else 0) for c in column)
        ws.column_dimensions[chr(64+col_idx)].width = min(max_length + 2, 50)
    
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    wb.save(excel_path)

# Интерфейс
st.title(" Анализатор табличек с адресами")
st.markdown("Загрузите ZIP или RAR архив с фотографиями — программа проверит читаемость адреса")

st.warning("️ Обработка 1000 фото может занять до 20 минут")

# Загрузка файла
uploaded_file = st.file_uploader(
    "Выберите архив",
    type=['zip', 'rar'],
    help="Поддерживается до 1000 фотографий"
)

if uploaded_file:
    task_id = str(uuid.uuid4())[:8]
    filename = uploaded_file.name
    ext = filename.rsplit('.', 1)[1].lower()
    
    # Сохраняем архив
    archive_path = f"uploads/{task_id}_{filename}"
    with open(archive_path, 'wb') as f:
        f.write(uploaded_file.getbuffer())
    
    # Распаковка
    extract_dir = f"uploads/{task_id}"
    previews_dir = f"results/{task_id}/previews"
    os.makedirs(previews_dir, exist_ok=True)
    
    with st.spinner('Распаковка архива...'):
        try:
            if ext == 'zip':
                with zipfile.ZipFile(archive_path, 'r') as z:
                    z.extractall(extract_dir)
            elif ext == 'rar':
                with rarfile.RarFile(archive_path, 'r') as r:
                    r.extractall(extract_dir)
        except Exception as e:
            st.error(f"Ошибка распаковки: {e}")
            st.stop()
    
    # Сбор изображений
    images = []
    for root, dirs, files in os.walk(extract_dir):
        for f in files:
            if Path(f).suffix.lower() in {'.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.webp'}:
                images.append(os.path.join(root, f))
    
    if not images:
        st.error("В архиве не найдено изображений")
        st.stop()
    
    if len(images) > 1000:
        st.error(f"Слишком много файлов: {len(images)}. Максимум 1000.")
        st.stop()
    
    st.success(f"Найдено {len(images)} изображений. Начинаем анализ...")
    
    # Анализ с прогресс-баром
    results = []
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for idx, img_path in enumerate(images):
        name = Path(img_path).name
        is_valid, comment = analyze_image(img_path)
        create_thumbnail(img_path, f"{previews_dir}/{Path(name).stem}.jpg")
        results.append({
            'filename': name,
            'preview': f"{previews_dir}/{Path(name).stem}.jpg",
            'valid': is_valid,
            'comment': comment
        })
        
        progress_bar.progress((idx + 1) / len(images))
        status_text.text(f"Обработано {idx + 1} из {len(images)}")
    
    # Создание DataFrame
    df = pd.DataFrame([
        {
            'Название файла': r['filename'],
            'Соответствует': 'Да' if r['valid'] else 'Нет',
            'Комментарий': r['comment']
        }
        for r in results
    ])
    
    # Сохранение Excel
    out_path = f"results/{task_id}_Результат.xlsx"
    df.to_excel(out_path, index=False, engine='openpyxl')
    make_excel_pretty(out_path)
    
    # Статистика
    valid_count = sum(1 for r in results if r['valid'])
    invalid_count = len(results) - valid_count
    
    st.success("✅ Анализ завершен!")
    
    # Метрики
    col1, col2, col3 = st.columns(3)
    col1.metric("Всего фото", len(results))
    col2.metric("Читаемых", valid_count)
    col3.metric("Нечитаемых", invalid_count)
    
    # Фильтры
    filter_option = st.radio(
        "Фильтр результатов:",
        ["Все", "Только читаемые", "Только нечитаемые"],
        horizontal=True
    )
    
    if filter_option == "Только читаемые":
        df_show = df[df['Соответствует'] == 'Да']
    elif filter_option == "Только нечитаемые":
        df_show = df[df['Соответствует'] == 'Нет']
    else:
        df_show = df
    
    # Таблица с превью
    st.subheader(f"Результаты ({len(df_show)} из {len(results)})")
    
    for _, row in df_show.iterrows():
        cols = st.columns([1, 3, 1, 4])
        
        # Превью
        preview_path = next((r['preview'] for r in results if r['filename'] == row['Название файла']), None)
        if preview_path and os.path.exists(preview_path):
            cols[0].image(preview_path, width=100)
        
        cols[1].text(row['Название файла'])
        
        # Статус
        if row['Соответствует'] == 'Да':
            cols[2].success("Да")
        else:
            cols[2].error("Нет")
        
        cols[3].text(row['Комментарий'])
    
    # Кнопка скачивания
    with open(out_path, 'rb') as f:
        st.download_button(
            label="⬇️ Скачать Excel-отчёт",
            data=f,
            file_name=f"Результат_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    # Очистка временных файлов
    if os.path.exists(archive_path):
        os.remove(archive_path)
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir, ignore_errors=True)
