# KP Generator (проект «КП финал»)

**Весь проект расположен в этой папке** («КП финал» / KP_Final). Код, данные, шаблоны и инструкции не выходят за её пределы — можно копировать или архивировать только эту папку.

Генератор коммерческих предложений из Excel/PDF с поддержкой нескольких вариантов (компания, наценка, округление) и вывода в Excel или Word (DOCX).

## Требования

- Python 3.9+
- Зависимости: см. `requirements.txt`
- Для распознавания текста в PDF-сканах: [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) (опционально)

## Установка и запуск

Перейдите в папку проекта (та папка, где лежат `kp_generator`, `requirements.txt` и этот README) и выполните:

```bash
cd <путь к папке КП финал>
pip install -r requirements.txt
python -m kp_generator.app
```

## Сборка в .exe

```bash
pip install pyinstaller
build_exe.bat
```

Или вручную — см. раздел «Сборка в один .exe» в **ИНСТРУКЦИЯ.md**.

## Документация

- **ИНСТРУКЦИЯ.md** — как подготовить шаблоны, выбрать файлы, настроить параметры и где искать результаты.
- **kp_generator/TESSERACT_README.txt** — установка Tesseract для OCR PDF.

## Структура проекта

```
kp_generator/
├── app.py              # точка входа
├── gui.py              # главное окно
├── create_templates.py # создание примеров шаблонов
├── ...
├── extract/            # чтение Excel, PDF, разбор таблиц
├── render/             # генерация Excel и DOCX
└── assets/
    ├── companies.json  # список компаний
    └── templates/      # примеры шаблонов (DOCX и Excel)
```

Примеры шаблонов создаются командой:  
`python -m kp_generator.create_templates`  
(появятся в `kp_generator/assets/templates/`).

## Релиз

Вся папка **«КП финал»** — это и есть полный проект. Для передачи заказчику или выкладки на GitHub достаточно заархивировать именно эту папку (в ней уже есть README, ИНСТРУКЦИЯ, код, assets, шаблоны). При распространении exe добавьте в архив папку `dist/` с `KP_Generator.exe` и при необходимости папку `tesseract` с Tesseract для OCR.
