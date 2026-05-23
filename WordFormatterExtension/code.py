import os
import re
import win32com.client
from win32com.client import constants

def run(input_path, output_path=None, font_name="Calibri", font_size=12, replacements=None):
    """
    Основная функция блока.
    Выполняет форматирование Word-документа.
    """
    if replacements is None:
        replacements = {}

    # Проверка существования входного файла
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Файл не найден: {input_path}")

    if not input_path.lower().endswith('.docx'):
        raise ValueError("Поддерживаются только файлы формата .docx")

    # Запускаем Word
    word = None
    doc = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # Работаем в фоне
        word.DisplayAlerts = False

        # Открываем документ
        doc = word.Documents.Open(os.path.abspath(input_path))

        # 1. Удаляем пустые абзацы
        _remove_empty_paragraphs(doc)

        # 2. Удаляем множественные пробелы и табуляции
        _clean_whitespace(doc)

        # 3. Очищаем прямое форматирование
        _clear_direct_formatting(doc)

        # 4. Применяем единый шрифт ко всему документу
        _apply_unified_font(doc, font_name, font_size)

        # 5. Определяем заголовки по нумерации и ключевым словам
        _apply_headings(doc)

        # 6. Форматируем таблицы
        _format_tables(doc)

        # 7. Заменяем плейсхолдеры
        if replacements:
            _replace_placeholders(doc, replacements)

        # Сохраняем результат
        if output_path:
            doc.SaveAs(os.path.abspath(output_path))
        else:
            doc.Save()  # Перезаписываем исходный

        return {"success": True, "message": f"Документ успешно отформатирован и сохранён"}

    except Exception as e:
        raise RuntimeError(f"Ошибка при обработке документа: {str(e)}")

    finally:
        # Закрываем документ и Word
        if doc:
            doc.Close(False)
        if word:
            word.Quit()


def _remove_empty_paragraphs(doc):
    """Удаляет пустые абзацы"""
    paragraphs = doc.Paragraphs
    # Идём с конца, чтобы не сбивать индексы
    for i in range(paragraphs.Count, 0, -1):
        para = paragraphs(i)
        if len(para.Range.Text.strip()) == 0:
            para.Range.Delete()


def _clean_whitespace(doc):
    """Заменяет множественные пробелы и табуляции на один пробел"""
    # Включаем поиск по всему документу
    find = doc.Content.Find
    find.ClearFormatting()
    find.Replacement.ClearFormatting()

    # Замена множественных пробелов
    find.Text = "  +"  # два и более пробелов
    find.Replacement.Text = " "
    find.MatchWildcards = True
    find.Execute(Replace=2)  # wdReplaceAll = 2

    # Замена табуляций на пробел
    find.Text = "\t+"
    find.Replacement.Text = " "
    find.Execute(Replace=2)


def _clear_direct_formatting(doc):
    """Очищает прямое форматирование всего документа"""
    doc.Content.Font.Reset()
    # Сбрасываем также форматирование абзацев
    doc.Content.ParagraphFormat.Reset()


def _apply_unified_font(doc, font_name, font_size):
    """Применяет единый шрифт ко всему документу"""
    content = doc.Content
    content.Font.Name = font_name
    content.Font.Size = font_size
    content.Font.Bold = False
    content.Font.Italic = False


def _apply_headings(doc):
    """
    Автоматически определяет заголовки:
    - По нумерации: 1., 1.1., 2.3.4. и т.д.
    - По ключевым словам: "Введение", "Заключение", "Глава", "Раздел"
    """
    paragraphs = doc.Paragraphs

    for i in range(1, paragraphs.Count + 1):
        para = paragraphs(i)
        text = para.Range.Text.strip()
        if not text:
            continue

        # Шаблоны нумерации
        heading_patterns = [
            r"^\d+\.\s+",           # 1.
            r"^\d+\.\d+\.\s+",      # 1.1.
            r"^\d+\.\d+\.\d+\.\s+"  # 1.1.1.
        ]

        # Ключевые слова для заголовков
        keywords = ["Введение", "Заключение", "Глава", "Раздел", "Список", "Приложение"]

        is_heading = False

        # Проверяем нумерацию
        for pattern in heading_patterns:
            if re.match(pattern, text, re.UNICODE):
                is_heading = True
                level = pattern.count("\\.")  # количество точек = уровень
                break

        # Проверяем ключевые слова
        if not is_heading:
            for kw in keywords:
                if text.startswith(kw):
                    is_heading = True
                    level = 1
                    break

        if is_heading:
            # Применяем стиль заголовка
            if level == 1:
                para.Range.Style = doc.Styles("Заголовок 1")
            else:
                para.Range.Style = doc.Styles("Заголовок 2")


def _format_tables(doc):
    """Форматирует все таблицы в документе"""
    for table in doc.Tables:
        # Применяем границы ко всей таблице
        table.Borders.Enable = True
        table.Borders.InsideLineStyle = constants.wdLineStyleSingle
        table.Borders.OutsideLineStyle = constants.wdLineStyleSingle

        # Выравнивание содержимого по левому краю (кроме заголовков)
        for row in table.Rows:
            for cell in row.Cells:
                cell.Range.ParagraphFormat.Alignment = constants.wdAlignParagraphLeft

        # Форматируем первую строку как заголовок
        if table.Rows.Count > 0:
            header_row = table.Rows(1)
            for cell in header_row.Cells:
                cell.Range.Font.Bold = True
                # Светло-серая заливка
                cell.Shading.BackgroundPatternColor = 15773696  # RGB(240,240,240)
                cell.Range.ParagraphFormat.Alignment = constants.wdAlignParagraphCenter


def _replace_placeholders(doc, replacements):
    """
    Заменяет плейсхолдеры вида {{ключ}} на значения из словаря
    """
    find = doc.Content.Find
    find.ClearFormatting()
    find.Replacement.ClearFormatting()

    for key, value in replacements.items():
        find.Text = f"{{{{{key}}}}}"
        find.Replacement.Text = str(value)
        find.MatchCase = False
        find.MatchWholeWord = False
        find.Execute(Replace=2)  # wdReplaceAll