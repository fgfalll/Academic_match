# Технічна документація

## Архітектура проєкту

Проєкт складається з трьох основних модулів:

| Файл | Призначення |
|------|-------------|
| `academ_back.py` | Головний застосунок, Tkinter-інтерфейс, аналіз кандидатів |
| `ai_advisor.py` | AI-консультант з чат-інтерфейсом та веб-пошуком |
| `crypto_utils.py` | Шифрування PIN-коду та даних сесії |

## Залежності

| Бібліотека | Версія | Призначення |
|------------|-------|-------------|
| tkinter | stdlib | GUI-фреймворк |
| requests | latest | HTTP-запити до API |
| scholarly | latest | Google Scholar API |
| fake-useragent | latest | Імітація браузера |
| cryptography | latest | AES-шифрування |
| litellm | latest | Уніфікований AI API |
| markdown | latest | Рендеринг Markdown |
| tkinterweb | latest | HTML-рендеринг у Tkinter |
| beautifulsoup4 | latest | Парсинг HTML |
| ddgs | latest | DuckDuckGo пошук |

## Структура даних

### Кандидат

```python
{
    "name": str,                    # ПІБ
    "orcid": str,                   # ORCID ідентифікатор
    "gs_id": str,                   # Google Scholar ID
    "sources_fetched": set,          # Зібрані джерела: {"orcid", "openalex", "scholar"}
    "papers": list[Paper],          # Список публікацій
    "keywords": list[str],           # Ключові слова публікацій
    "concepts": list[str],          # Концепти з OpenAlex
    "manual_keywords": list[str],   # Введені вручну ключові слова
    "relevance_score": int,         # Сумарний бал релевантності
    "conflict": bool,               # Конфлікт інтересів
}
```

### Публікація

```python
{
    "title": str,
    "year": int,
    "authors": list[str],
    "journal": str,
    "abstract": str,
    "doi": str,
    "url": str,
    "source": str,                   # "orcid" | "openalex" | "scholar" | "manual"
    "author_keywords": list[str],
    "concepts": list[str],
    "manual_keywords": list[str],
    "score": int,
    "matched_keywords": list[str],
}
```

## Модуль аналізу (academ_back.py)

### Джерела даних

#### ORCID API
- Безкоштовний публічний API без автентифікації
- Ліміт: ~300 запитів/день для одного ORCID
- Дані: публікації, DOI, назви, рік

#### OpenAlex API
- Політичний пул: 10 req/сек
- Дані: work-id, abstract_inverted_index, concepts, author-position
- URL: `https://api.openalex.org/`

#### Google Scholar (scholarly)
- LLM-семплювання для обходу rate limit
- Затримки: 15-45с між запитами
- Мінімальний режим: лише назви публікацій

### Алгоритм оцінки релевантності

```
heuristic_score(title, concepts, author_keywords, manual_keywords, target_keywords, banned_keywords, abstract)

Параметри оцінки:
- Назва статті:        +5 балів
- Ключове слово:      +4 бали
- Концепт:            +3 бали
- Анотація:           +2 бали

Формула:
1. Нормалізація тексту (нижній регістр, заміна апострофів)
2. Пошук збігів target_keywords у title, keywords, concepts, abstract
3. Віднімання banned_keywords
4. Повернення (score, matched_keywords)
```

### Tkinter-віджет структура

```
root (Tk)
├── menubar
├── notebook (ttk.Notebook)
│   ├── tab_main  "1. Налаштування"
│   │   ├── settings_frame (LabelFrame)
│   │   ├── candidates_frame (LabelFrame)
│   │   └── log_frame (LabelFrame + Text)
│   ├── tab_edit  "2. Результати"
│   │   ├── candidate_tree (Treeview)
│   │   ├── paper_tree (Treeview)
│   │   └── details_frame
│   └── tab_advice "3. Аналіз термінів"
│       ├── candidate_listbox (Listbox)
│       ├── banned_words_frame
│       └── report_text (Text)
└── statusbar
```

### Збереження сесії

Формат файлу: `.acmp` (ZIP-архів)

```
session.acmp
├── session.json          # Дані сесії
├── encrypted_data.bin    # API ключі (AES-256-CBC, PBKDF2)
└── metadata.json         # Метадані
```

### Ключові функції

| Функція | Опис |
|---------|------|
| `get_author_info_openalex(orcid)` | Отримання імені автора з OpenAlex |
| `heuristic_score(...)` | Розрахунок релевантності публікації |
| `fetch_orcid_papers(orcid, cutoff)` | Збір публікацій з ORCID |
| `fetch_openalex_works(orcid)` | Збір публікацій з OpenAlex |
| `fetch_scholar_papers(gs_id, minimal)` | Збір публікацій з Google Scholar |
| `analyze_keywords(candidates)` | Аналіз термінів для кандидатів |

## AI-консультант (ai_advisor.py)

### LiteLLM-провайдери

Підтримується 10+ провайдерів:
- OpenAI (GPT-4, GPT-3.5)
- Anthropic (Claude)
- DeepSeek
- Gemini
- Azure OpenAI
- та інші

### Веб-пошук

```
web_search(query, num_results)
├── DuckDuckGo (ddgs) - основний
└── Tavily - fallback
```

### Артефакти

Система відображення результатів AI у структурованих блоках:

| Тип | Колір | Призначення |
|-----|-------|-------------|
| recommendation | зелений | Рекомендації щодо кандидата |
| summary | синій | Підсумки публікацій |
| comparison | червоний | Порівняння кандидатів |
| search_result | оранжевий | Результати веб-пошуку |

### Markdown-рендеринг

```
1. Отримання Markdown-тексту від LiteLLM
2. Рендеринг у HTML (markdown.markdown)
3. Обробка артефактних маркерів
4. Відображення у tkinterweb.HtmlFrame
```

### Системні промпти

Файли промптів містять:
- Регуляторна база: КМУ №44, №502, оновлення 2026
- Структура ради: 5 членів, ролі, альтернативні формати
- Наукометричні критерії: 3+ публікації, 5-річне вікно
- Конфлікт інтересів: 5-рівнева система обмежень
- Цифрова інфраструктура: NAQA.Svr, 30-денне розкриття

## Криптографія (crypto_utils.py)

### Алгоритми

| Параметр | Значення |
|----------|----------|
| Хешування PIN | SHA-256 |
| Ключ AES | PBKDF2-HMAC-SHA256 |
| Режим AES | CBC |
| Ітерації PBKDF2 | 100 000 |
| Сіль | `academic_match_salt_v1` |
| Вивід ключа | 256 біт (32 байти) |

### Функції

| Функція | Опис |
|---------|------|
| `hash_pin(pin)` | Хешування PIN для перевірки |
| `derive_aes_key(pin)` | Генерація AES-ключа з PIN |
| `encrypt_with_pin(data, pin)` | Шифрування даних |
| `decrypt_with_pin(encrypted, pin)` | Розшифрування даних |
| `encrypt_with_embedded_pin_hash(data, pin)` | Шифрування з вбудованим хешем PIN |
| `decrypt_with_embedded_pin_hash(encrypted, pin)` | Розшифрування з перевіркою хешу |

### Схема шифрування API-ключів

```
1. PIN → SHA-256 hash → збереження у .pin_hash
2. PIN → PBKDF2 → AES-ключ
3. Дані → AES-256-CBC → base64
4. PINHASH:{hash}:{data} → шифрування → encrypted.bin
```

## Збірка у exe

### PyInstaller

```bash
pyinstaller --onefile --noconsole academ_back.py
```

### setup.bat

Автоматичний скрипт:
1. Створення віртуального середовища
2. Встановлення залежностей
3. Збірка через PyInstaller
4. Результат: `dist/AcademicMatch.exe`

## API-обмеження

| Джерело | Ліміт | Затримка |
|---------|-------|----------|
| ORCID | 300/день | 1 req/сек |
| OpenAlex | 10/сек | 100 мс |
| Google Scholar | ~100/год | 15-45 сек |

## Файлова структура

```
academic-match/
├── academ_back.py          # Головний застосунок
├── ai_advisor.py           # AI-консультант
├── crypto_utils.py         # Криптографія
├── requirements.txt        # Залежності
├── setup.bat              # Скрипт збірки
├── README.md              # Документація
├── docs/
│   └── screenshots/       # Скріншоти інтерфейсу
├── Old/                   # Попередні версії
├── .pin_hash              # Хеш PIN (створення)
└── sessions/              # Збережені сесії
```
