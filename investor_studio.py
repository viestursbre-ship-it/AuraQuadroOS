# -*- coding: utf-8 -*-
import os
import sys
import json
import threading
from pathlib import Path
import customtkinter as ctk
import yfinance as yf
from datetime import datetime, timedelta
import re # Added for safe JSON extraction (Task 1)
import requests # LEO ADDITION: Import requests for Alpha Vantage and FMP

# Import for messagebox (FIX 1)
from tkinter import messagebox

try:
    import google.generativeai as genai
    HAS_GENAI = True
except ImportError:
    HAS_GENAI = False

try:
    from docx import Document
    from docx.shared import Inches
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# Import i18n functions (NEW)
try:
    import i18n # Import the i18n module directly
    from i18n import get_llm_language_instruction
    HAS_I18N = True
except ImportError:
    HAS_I18N = False

BASE_DIR = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent
PORTFOLIO_FILE = BASE_DIR / "portfolio.json"
OUTBOX_DIR = BASE_DIR / "storage" / "outbox"
OUTBOX_DIR.mkdir(parents=True, exist_ok=True) # Ensure outbox directory exists

# Global translations dictionary (NEW) - This structure is kept for the dummy i18n and initial definition.
TRANSLATIONS = {
    # Existing Keys
    "lbl_ticker_analysis": {"lv": "TIKERA ANALĪZE", "ru": "АНАЛИЗ ТИКЕРА", "uk": "АНАЛІЗ ТИКЕРА", "en": "TICKER ANALYSIS", "it": "ANALISI TICKER"},
    "btn_rentgens": {"lv": "Rentgens", "ru": "Рентген", "uk": "Рентген", "en": "X-Ray", "it": "Raggi X"},
    "lbl_inv_resolution_title": {"lv": "KOMANDANTA KVARKA REZOLŪCIJA", "ru": "РЕЗОЛЮЦИЯ КОМАНДАНТЕ КВАРКА", "uk": "РЕЗОЛЮЦІЯ КОМАНДАНТЕ КВАРКА", "en": "COMMANDER KVARK RESOLUTION", "it": "RISOLUZIONE DEL COMANDANTE KVARKS"}, # Renamed from lbl_resolution for clarity with specific instruction
    "txt_waiting": {"lv": "Gaidu pieprasījumu...", "ru": "Ожидаю запрос...", "uk": "Очікую запит...", "en": "Awaiting request...", "it": "In attesa della richiesta..."},
    "lbl_potential_radar": {"lv": "POTENCIĀLA RADARS", "ru": "РАДАР ПОТЕНЦИАЛА", "uk": "РАДАР ПОТЕНЦІАЛУ", "en": "POTENTIAL RADAR", "it": "RADAR DEL POTENZIALE"},
    "btn_hunt_unicorns": {"lv": "Medīt Jaunos Vienradžus", "ru": "Охота на Единорогов", "uk": "Полювання на Єдинорогів", "en": "Hunt New Unicorns", "it": "Caccia agli Unicorni"},
    "btn_check": {"lv": "Pārbaudīt", "ru": "Проверить", "uk": "Перевірити", "en": "Check", "it": "Verifica"},
    "btn_save_word": {"lv": "Saglabāt Word (.docx)", "ru": "Сохранить Word (.docx)", "uk": "Зберегти Word (.docx)", "en": "Save Word (.docx)", "it": "Salva Word (.docx)"},

    # New Keys for Tabs and Titles (Instruction 3 & 2)
    "tab_inv_single": {"lv": "🔍 Viena Tikera Rentgens", "en": "🔍 Single Ticker X-Ray", "ru": "🔍 Рентген Тикера", "uk": "🔍 Рентген Тікера", "it": "🔍 Raggi X singolo titolo"},
    "tab_inv_etoro": {"lv": "💼 Mans eToro Portfelis & Audits", "en": "💼 My eToro Portfolio & Audit", "ru": "💼 Мой Портфель eToro и Аудит", "uk": "💼 Мій Портфель eToro та Аудит", "it": "💼 Mio Portfolio eToro & Audit"},
    "tab_inv_unicorns": {"lv": "🚀 Kvarka Izaugsmes Mednieks", "en": "🚀 Quark's Growth Hunter", "ru": "🚀 Охотник за Ростом Кварка", "uk": "🚀 Мисливець за Зростанням Кварка", "it": "🚀 Caccia alla Crescita di Quark"},
    "tab_inv_dividends": {"lv": "💰 Dividendes & Naudas Plūsma", "en": "💰 Dividends & Cash Flow", "ru": "💰 Дивиденды и Денежный Поток", "uk": "💰 Дивіденди та Грошовий Потік", "it": "💰 Dividendi & Flusso di Cassa"},
    "tab_inv_insiders": {"lv": "🕵️ Tirgus Klimats & Insaideri", "en": "🕵️ Market Climate & Insiders", "ru": "🕵️ Климат Рынка и Инсайдеры", "uk": "🕵️ Ринок та Інсайдери", "it": "🕵️ Clima di Mercato & Insider"},
    "lbl_inv_growth_resolution_title": {"lv": "🎯 IZAUGSMES REZOLŪCIJA & RISKA NOVĒRTĒJUMS", "en": "🎯 GROWTH RESOLUTION & RISK ASSESSMENT", "ru": "🎯 РЕЗОЛЮЦИЯ ПО РОСТУ И ОЦЕНКА РИСКА", "uk": "🎯 РЕЗОЛЮЦІЯ З РОСТУ ТА ОЦІНКА РИЗИКУ", "it": "🎯 RISOLUZIONE DI CRESCITA & VALUTAZIONE DEL RISCHIO"},
    "lbl_inv_dividends_resolution_title": {"lv": "📅 KVARKA DIVIDENŽU KALENDĀRS & REZOLŪCIJA", "en": "📅 QUARK'S DIVIDEND CALENDAR & RESOLUTION", "ru": "📅 КАЛЕНДАРЬ ДИВИДЕНДОВ КВАРКА И РЕЗОЛЮЦИЯ", "uk": "📅 КАЛЕНДАР ДИВІДЕНДІВ КВАРКА ТА РЕЗОЛЮЦІЯ", "it": "📅 CALENDARIO DIVIDENDI DI QUARK & RISOLUZIONE"},
    "lbl_inv_market_news_title": {"lv": "📰 JAUNĀKĀS ZIŅAS & KVARKA TAKTIKA", "en": "📰 LATEST NEWS & QUARK'S TACTICS", "ru": "📰 ПОСЛЕДНИЕ НОВОСТИ И ТАКТИКА КВАРКА", "uk": "📰 ОСТАННІ НОВИНИ ТА ТАКТИКА КВАРКА", "it": "📰 ULTIME NOTIZIE & TATTICHE DI QUARK"},

    # Portfolio Tab
    "lbl_portfolio_title": {"lv": "💼 MANAS ETORO AKCIJAS", "en": "💼 MY ETORO STOCKS", "ru": "💼 МОИ АКЦИИ ETORO", "uk": "💼 МОЇ АКЦІЇ ETORO", "it": "💼 LE MIE AZIONI ETORO"},
    "lbl_portfolio_tickers_desc": {"lv": "Ievadiet savus tikerus (atdalītus ar komatu vai jaunā rindā):\nPiemērs: AAPL, NVDA, TSLA, INTC, PYPL", "en": "Enter your tickers (comma-separated or new line):\nExample: AAPL, NVDA, TSLA, INTC, PYPL", "ru": "Введите свои тикеры (через запятую или новую строку):\nПример: AAPL, NVDA, TSLA, INTC, PYPL", "uk": "Введіть свої тікери (через кому або новий рядок):\nПриклад: AAPL, NVDA, TSLA, INTC, PYPL", "it": "Inserisci i tuoi ticker (separati da virgola o su nuova riga):\nEs: AAPL, NVDA, TSLA, INTC, PYPL"},
    "lbl_portfolio_cash_balance": {"lv": "💵 Brīvais atlikums (USD):", "en": "💵 Free Balance (USD):", "ru": "💵 Свободный остаток (USD):", "uk": "💵 Вільний залишок (USD):", "it": "💵 Saldo disponibile (USD):"},
    "lbl_portfolio_risk_strategy": {"lv": "🎯 Riska Stratēģija:", "en": "🎯 Risk Strategy:", "ru": "🎯 Стратегия риска:", "uk": "🎯 Стратегія ризику:", "it": "🎯 Strategia di rischio:"},
    "seg_risk_conservative": {"lv": "🛡️ Konservatīvs (5-10%)", "en": "🛡️ Conservative (5-10%)", "ru": "🛡️ Консервативный (5-10%)", "uk": "🛡️ Консервативний (5-10%)", "it": "🛡️ Conservativo (5-10%)"},
    "seg_risk_venture_growth": {"lv": "🚀 Venture & Growth", "en": "🚀 Venture & Growth", "ru": "🚀 Венчур и Рост", "uk": "🚀 Венчур та Зростання", "it": "🚀 Venture & Crescita"},
    "btn_save_portfolio": {"lv": "💾 Saglabāt", "en": "💾 Save", "ru": "💾 Сохранить", "uk": "💾 Зберегти", "it": "💾 Salva"},
    "btn_audit_portfolio": {"lv": "🐾 Kvarka Audits", "en": "🐾 Quark's Audit", "ru": "🐾 Аудит Кварка", "uk": "🐾 Аудит Кварка", "it": "🐾 Audit di Quark"},
    "status_portfolio_saved": {"lv": "✅ Portfelis saglabāts!", "en": "✅ Portfolio saved!", "ru": "✅ Портфель сохранен!", "uk": "✅ Портфель збережено!", "it": "✅ Portfolio salvato!"},
    "error_portfolio_save": {"lv": "⚠️ Kļūda: {e}", "en": "⚠️ Error: {e}", "ru": "⚠️ Ошибка: {e}", "uk": "⚠️ Помилка: {e}", "it": "⚠️ Errore: {e}"},
    "feedback_added_to_portfolio": {"lv": "✅ {ticker} pievienots portfelim!", "en": "✅ {ticker} added to portfolio!", "ru": "✅ {ticker} добавлен в портфель!", "uk": "✅ {ticker} додано до портфеля!", "it": "✅ {ticker} aggiunto al portfolio!"},
    "feedback_already_in_portfolio": {"lv": "ℹ️ {ticker} jau ir portfelī.", "en": "ℹ️ {ticker} is already in the portfolio.", "ru": "ℹ️ {ticker} уже в портфеле.", "uk": "ℹ️ {ticker} вже є в портфелі.", "it": "ℹ️ {ticker} è già nel portfolio."},
    "audit_copy_header": {"lv": "🐾 *KOMANDANTA KVARKA INVESTĪCIJU AUDITS* 🐾", "en": "🐾 *COMMANDER KVARK'S INVESTMENT AUDIT* 🐾", "ru": "🐾 *АУДИТ ИНВЕСТИЦИЙ КОМАНДИРА КВАРКА* 🐾", "uk": "🐾 *АУДИТ ІНВЕСТИЦІЙ КОМАНДИРА КВАРКА* 🐾", "it": "🐾 *AUDIT DI INVESTIMENTO DEL COMANDANTE KVARKS* 🐾"},
    "audit_copy_strategy": {"lv": "🎯 Stratēģija:", "en": "🎯 Strategy:", "ru": "🎯 Стратегия:", "uk": "🎯 Стратегія:", "it": "🎯 Strategia:"},
    "audit_copy_footer": {"lv": "⚡ Ģenerēts ar AQ-OS Investor Studio", "en": "⚡ Generated with AQ-OS Investor Studio", "ru": "⚡ Сгенерировано AQ-OS Investor Studio", "uk": "⚡ Згенеровано AQ-OS Investor Studio", "it": "⚡ Generato con AQ-OS Investor Studio"},
    "btn_copy_summary": {"lv": "📋 Kopēt kopsavilkumu", "en": "📋 Copy Summary", "ru": "📋 Копировать сводку", "uk": "📋 Копіювати зведення", "it": "📋 Copia riepilogo"},
    "status_copied": {"lv": "✅ Nokopēts!", "en": "✅ Copied!", "ru": "✅ Скопировано!", "uk": "✅ Скопійовано!", "it": "✅ Copiato!"},
    "error_copying": {"lv": "⚠️ Kļūda kopējot: {e}", "en": "⚠️ Error copying: {e}", "ru": "⚠️ Ошибка при копировании: {e}", "uk": "⚠️ Помилка копіēšanas: {e}", "it": "⚠️ Errore durante la copia: {e}"},
    "btn_audit_loading": {"lv": "⏳ Auditē...", "en": "⏳ Auditing...", "ru": "⏳ Аудит...", "uk": "⏳ Аудит...", "it": "⏳ Audit..."},
    "audit_loading_msg": {"lv": "Komandants Kvarks caurskata {len_tickers} pozīcijas un finanšu rādītājus...\n", "en": "Commander Quark reviews {len_tickers} positions and financial indicators...\n", "ru": "Командир Кварк просматривает {len_tickers} позиций и финансовые показатели...\n", "uk": "Командир Кварк переглядає {len_tickers} позицій та фінансові показники...\n", "it": "Il Comandante Quark esamina {len_tickers} posizioni e indicatori finanziari...\n"},
    "audit_initial_msg": {"lv": "Ievadiet savas akcijas un nospiediet 'Kvarka Audits'...\nKvarks noteiks, kurus 'sīkuļus' likvidēt un kur gudri ieguldīt līdzekļus!", "en": "Enter your stocks and press 'Quark's Audit'...\nQuark will determine which 'minnows' to liquidate and where to invest wisely!", "ru": "Введите свои акции и нажмите 'Аудит Кварка'...\nКварк определит, какие 'мелкие' позиции ликвидировать и куда разумно вложить средства!", "uk": "Введіть свої акції та натисніть 'Аудит Кварка'...\nКварк визначить, які 'дрібниці' ліквідувати і куди розумно вкласти кошти!", "it": "Inserisci i tuoi titoli e premi 'Audit di Quark'...\nQuark determinerà quali 'pesci piccoli' liquidare e dove investire saggiamente!"},
    "strategy_venture_desc": {"lv": "STRATĒĢIJA: VENTURE & GROWTH HUNTER (Agresīvs riska kapitāls).\n- Vērtē aktīvus no nākotnes tirgus dominances, monopolstāvokļa un AI/tehnoloģiju lēciena prizmas.\n- Augsts P/E vai īslaicīgs peļņas trūkums nav šķērslis, ja ieņēmumu lēciens un infrastruktūras nozīme ir milzīga (piemēram, ARM, SNOW, PLTR).\n- Izbrāķē tikai stagnējošus un lēnus uzņēmumus bez tehnoloģiskas nākotnes.", "en": "STRATEGY: VENTURE & GROWTH HUNTER (Aggressive Venture Capital).\n- Values assets from the perspective of future market dominance, monopoly position, and AI/technology leap.\n- High P/E or temporary lack of profit is not an obstacle if revenue growth and infrastructure importance are enormous (e.g., ARM, SNOW, PLTR).\n- Only rejects stagnant and slow companies without a technological future.", "ru": "СТРАТЕГИЯ: ВЕНЧУРНЫЙ ОХОТНИК ЗА РОСТОМ (Агрессивный венчурный капитал).\n- Оценивает активы с точки зрения будущего доминирования на рынке, монопольного положения и технологического скачка AI.\n- Высокий P/E или временное отсутствие прибыли не является препятствием, если рост выручки и значимость инфраструктуры огромны (например, ARM, SNOW, PLTR).\n- Отбрасывает только стагнирующие и медленные компании без технологического будущего.", "uk": "СТРАТЕГІЯ: ВЕНЧУРНИЙ МИСЛИВЕЦЬ ЗА ЗРОСТАННЯМ (Агресивний венчурний капітал).\n- Оцінює активи з позиції майбутнього домінування на ринку, монопольного становища та стрибка AI/технологій.\n- Високий P/E або тимчасова відсутність прибутку не є перешкодою, якщо стрибок доходів та значення інфраструктури величезні (наприклад, ARM, SNOW, PLTR).\n- Відкидає лише стагнуючі та повільні компанії без технологічного майбутнього.", "it": "STRATEGIA: Cacciatore di Venture & Crescita (Capitale di rischio aggressivo).\n- Valuta gli asset dal punto di vista del dominio futuro del mercato, della posizione di monopolio e del salto tecnologico AI.\n- Un P/E elevato o una temporanea mancanza di profitto non sono un ostacolo se la crescita dei ricavi e l'importanza dell'infrastruttura sono enormi (es. ARM, SNOW, PLTR).\n- Rifiuta solo le aziende stagnanti e lente senza un futuro tecnologico."},
    "strategy_conservative_desc": {"lv": "STRATĒĢIJA: KONSERVATĪVS (Zems risks, 5-10% stabils pieaugums gadā).\n- Vērtē aktīvus pēc reālas naudas plūsmas, veselīga P/E (< 30-35), dividenžu stabilitātes un zemiem riskiem.\n- Ja uzņēmuma P/E ir pārmērīgs (> 80-100) vai nav peļņas un PEG > 2, norādi to kā bīstamu burbuli un iesaki samazināt pozīciju.\n- Dod priekšroku tirgus indeksiem (VOO/SPY) un stabiliem industrijas milžiem.", "en": "STRATEGY: CONSERVATIVE (Low risk, 5-10% stable annual growth).\n- Values assets by real cash flow, healthy P/E (< 30-35), dividend stability, and low risks.\n- If a company's P/E is excessive (> 80-100) or has no profit and PEG > 2, indicate it as a dangerous bubble and recommend reducing the position.\n- Prefers market indices (VOO/SPY) and stable industry giants.", "ru": "СТРАТЕГИЯ: КОНСЕРВАТИВНАЯ (Низкий риск, 5-10% стабильного роста в год).\n- Оценивает активы по реальному денежному потоку, здоровому P/E (< 30-35), стабильности дивидендов и низким рискам.\n- Если P/E компании чрезмерен (> 80-100) или нет прибыли, а PEG > 2, укажите это как опасный пузырь и рекомендуйте уменьшить позицию.\n- Отдает предпочтение рыночным индексам (VOO/SPY) и стабильным промышленным гигантам.", "uk": "СТРАТЕГІЯ: КОНСЕРВАТИВНА (Низький ризик, 5-10% стабільний річний приріст).\n- Оцінює активи за реальним грошовим потоком, здоровим P/E (< 30-35), стабільністю дивідендів та низькими ризиками.\n- Якщо P/E компанії надмірний (> 80-100) або немає прибутку, а PEG > 2, вкажіть це як небезпечний міхurs і рекомендуйте зменшити позицію.\n- Віддає перевагу ринковим індексам (VOO/SPY) та стабільним промисловим гігантам.", "it": "STRATEGIA: CONSERVATIVA (Basso rischio, 5-10% crescita annua stabile).\n- Valuta gli asset in base al flusso di cassa reale, un P/E sano (< 30-35), la stabilità dei dividendi e i bassi rischi.\n- Se il P/E di un'azienda è eccessivo (> 80-100) o non c'è profitto e il PEG > 2, indicalo come una bolla pericolosa e consiglia di ridurre la posizione.\n- Preferisce gli indici di mercato (VOO/SPY) e i giganti del settore stabili."},
    "error_audit_gemini": {"lv": "⚠️ Audita kļūda: {e}", "en": "⚠️ Audit error: {e}", "ru": "⚠️ Ошибка аудита: {e}", "uk": "⚠️ Помилка аудиту: {e}", "it": "⚠️ Errore di audit: {e}"},
    "error_api_missing_data": {"lv": "⚠️ Trūkst API atslēgas vai neizdevās nolasīt datus.", "en": "⚠️ Missing API key or failed to retrieve data.", "ru": "⚠️ Отсутствует ключ API или не удалось получить данные.", "uk": "⚠️ Відсутній ключ API або не вдалося отримати дані.", "it": "⚠️ Chiave API mancante o impossibile recuperare i dati."},

    # Growth Tab
    "lbl_growth_description": {"lv": "Kvarks skenē nākotnes tirgus un medī TOP 10 vienradžus (AI, medicīna, enerģija, kosmoss, u.c.):", "en": "Quark scans future markets and hunts for TOP 10 unicorns (AI, medicine, energy, space, etc.):", "ru": "Кварк сканирует будущие рынки и ищет ТОП-10 единорогов (ИИ, медицина, энергетика, космос и др.):", "uk": "Кварк сканує майбутні ринки та полює на ТОП-10 єдинорогів (ШІ, медицина, енергетика, космос тощо):", "it": "Quark scansiona i mercati futuri e cerca i TOP 10 unicorni (AI, medicina, energia, spazio, ecc.):"},
    "growth_no_candidates": {"lv": "Nav atlasīti vienradžu kandidāti.", "en": "No unicorn candidates selected.", "ru": "Не выбрано ни одного кандидата-единорога.", "uk": "Не вибрано жодного кандидата-єдинорога.", "it": "Nessun candidato unicorno selezionato."},
    "growth_status_initial": {"lv": "Gaidu Komandanta Kvarka rīkojumus...", "en": "Awaiting Commander Quark's orders...", "ru": "Ожидаю приказов Командира Кварка...", "uk": "Очікую наказів Командира Кварка...", "it": "In attesa degli ordini del Comandante Quark..."},
    "btn_hunt_unicorns_loading": {"lv": "⏳ {btn_text}...", "en": "⏳ {btn_text}...", "ru": "⏳ {btn_text}...", "uk": "⏳ {btn_text}...", "it": "⏳ {btn_text}..."},
    "growth_scanning_candidates": {"lv": "🔍 Kvarks skenē... Gaidiet jaunus kandidātus!", "en": "🔍 Quark is scanning... Awaiting new candidates!", "ru": "🔍 Кварк сканирует... Ожидайте новых кандидатов!", "uk": "🔍 Кварк сканує... Чекайте нових кандидатів!", "it": "🔍 Quark sta scansionando... In attesa di nuovi candidati!"},
    "growth_scan_progress1": {"lv": "🔍 Kvarks skenē tirgu un vāc datus par vienradžiem...", "en": "🔍 Quark is scanning the market and collecting data on unicorns...", "ru": "🔍 Кварк сканирует рынок и собирает данные о единорогах...", "uk": "🔍 Кварк сканує ринок та збирає дані про єдиноrogus...", "it": "🔍 Quark sta scansionando il mercato e raccogliendo dati sugli unicorni..."},
    "growth_scan_progress2": {"lv": "📊 Atlasa jaunos vienradžus ar AI analīzi...", "en": "📊 Selecting new unicorns with AI analysis...", "ru": "📊 Выбирает новых единорогов с помощью AI анализа...", "uk": "📊 Вибирає нових єдинорогів за допомогою AI аналізу...", "it": "📊 Seleziona nuovi unicorni con analisi AI..."},
    "growth_initial_results": {"lv": "Izvēlieties uzņēmumu vai nospiediet '🚀 Medīt Jaunos Vienradžus', lai Kvarks atlasa nākamo lielo lēcienu...", "en": "Select a company or press '🚀 Hunt New Unicorns' for Quark to select the next big leap...", "ru": "Выберите компанию или нажмите '🚀 Искать новых единорогов', чтобы Кварк выбрал следующий большой скачок...", "uk": "Виберіть компанію або натисніть '🚀 Шукати нових єдинорогів', щоб Кварк вибрав наступний великий стрибок...", "it": "Seleziona un'azienda o premi '🚀 Caccia agli Unicorni' per far sì che Quark selezioni il prossimo grande salto..."},
    "btn_copy_analysis": {"lv": "📋 Kopēt analīzi", "en": "📋 Copy Analysis", "ru": "📋 Копировать анализ", "uk": "📋 Копіювати аналіз", "it": "📋 Copia analisi"},
    "growth_no_unicorns_fallback": {"lv": "Kvarks neatrada piemērotus vienradžus vai radās kļūda.", "en": "Quark found no suitable unicorns or an error occurred.", "ru": "Кварк не нашел подходящих единорогов или произошла ошибка.", "uk": "Кварк не знайшов відповідних єдинорогів або сталася помилка.", "it": "Quark non ha trovato unicorni adatti o si è verificato un errore."},
    "error_gemini_parsing": {"lv": "⚠️ Gemini atgrieza nestrukturētu atbildi vai kļūda parsēšanā: {e}. Daļa: {partial_response}...", "en": "⚠️ Gemini returned unstructured response or parsing error: {e}. Part: {partial_response}...", "ru": "⚠️ Gemini вернул неструктурированный ответ или ошибка парсинга: {e}. Часть: {partial_response}...", "uk": "⚠️ Gemini повернув неструктуровану відповідь або помилка парсингу: {e}. Частина: {partial_response}...", "it": "⚠️ Gemini ha restituito una risposta non strutturata o errore di parsing: {e}. Parte: {partial_response}..."},
    "error_ai_unparsable": {"lv": "⚠️ AI atbilde nav parsējama. Mēģiniet vēlreiz vai pārbaudiet API.", "en": "⚠️ AI response unparsable. Try again or check API.", "ru": "⚠️ Ответ ИИ не поддается парсингу. Попробуйте еще раз или проверьте API.", "uk": "⚠️ Відповідь ШІ не піддається парсингу. Спробуйте ще раз або перевірте API.", "it": "⚠️ Risposta AI non parsabile. Riprova o controlla l'API."},
    "error_gemini_api_contact": {"lv": "⚠️ Kļūda sazinoties ar Gemini: {e}", "en": "⚠️ Error communicating with Gemini: {e}", "ru": "⚠️ Ошибка связи с Gemini: {e}", "uk": "⚠️ Помилка зв'язку ar Gemini: {e}", "it": "⚠️ Errore di comunicazione con Gemini: {e}"},
    "growth_analysis_complete": {"lv": "✅ Analīze pabeigta!", "en": "✅ Analysis complete!", "ru": "✅ Анализ завершен!", "uk": "✅ Аналіз завершено!", "it": "✅ Analisi completata!"},
    "warning_red_flag": {"lv": "🚩 Sarkanais Karogs: {warning_message}", "en": "🚩 Red Flag: {warning_message}", "ru": "🚩 Красный Флаг: {warning_message}", "uk": "🚩 Червоний Прапор: {warning_message}", "it": "🚩 Bandiera Rossa: {warning_message}"},
    "lbl_niche": {"lv": "Slepenais ierocis (Niša):", "en": "Secret Weapon (Niche):", "ru": "Секретное оружие (Ниша):", "uk": "Секретна зброя (Ніша):", "it": "Arma Segreta (Nicchia):"},
    "lbl_why_10x_100x": {"lv": "Kāpēc 10x-100x:", "en": "Why 10x-100x:", "ru": "Почему 10x-100x:", "uk": "Чому 10x-100x:", "it": "Perché 10x-100x:"},
    "lbl_black_swan": {"lv": "Melnais Gulbis (risks):", "en": "Black Swan (risk):", "ru": "Черный Лебедь (риск):", "uk": "Чорний Лебīдь (ризик):", "it": "Cigno Nero (rischio):"},
    "btn_add_to_portfolio": {"lv": "➕ Pievienot manam eToro sarakstam", "en": "➕ Add to my eToro list", "ru": "➕ Добавить в мой список eToro", "uk": "➕ Додати до мого списку eToro", "it": "➕ Aggiungi alla mia lista eToro"},
    "growth_copy_header": {"lv": "🚀 KOMANDANTA KVARKA VIENRADŽU MEDĪBAS REZULTĀTI 🚀", "en": "🚀 COMMANDER KVARK'S UNICORN HUNT RESULTS 🚀", "ru": "🚀 РЕЗУЛЬТАТЫ ОХОТЫ НА ЕДИНОРОГОВ КОМАНДИРА КВАРКА 🚀", "uk": "🚀 РЕЗУЛЬТАТИ ПОЛЮВАННЯ НА ЄДИНОРОГІВ КОМАНДИРА КВАРКА 🚀", "it": "🚀 RISULTATI DELLA CACCIA AGLI UNICORNI DEL COMANDANTE QUARK 🚀"},
    "growth_copy_candidate_header": {"lv": "--- Kandidāts {idx_plus_1}: {ticker} - {name} ---", "en": "--- Candidate {idx_plus_1}: {ticker} - {name} ---", "ru": "--- Кандидат {idx_plus_1}: {ticker} - {name} ---", "uk": "--- Кандидат {idx_plus_1}: {ticker} - {name} ---", "it": "--- Candidato {idx_plus_1}: {ticker} - {name} ---"},
    "growth_copy_niche": {"lv": "   Slepenais ierocis (Niša): {niche}", "en": "   Secret Weapon (Niche): {niche}", "ru": "   Секретное оружие (Ниша): {niche}", "uk": "   Секретна зброя (Ніша): {niche}", "it": "   Arma Segreta (Nicchia): {niche}"},
    "growth_copy_why": {"lv": "   Kāpēc 10x-100x: {why}", "en": "   Why 10x-100x: {why}", "ru": "   Почему 10x-100x: {why}", "uk": "   Чому 10x-100x: {why}", "it": "   Perché 10x-100x:"},
    "growth_copy_risk": {"lv": "   Melnais Gulbis (risks): {risk}", "en": "   Black Swan (risk): {risk}", "ru": "   Черный Лебедь (риск): {risk}", "uk": "   Чорний Лебідь (ризик): {risk}", "it": "   Cigno Nero (rischio): {risk}"},
    "growth_no_analysis_to_copy": {"lv": "Nav atrasta neviena analīze, ko kopēt.", "en": "No analysis found to copy.", "ru": "Не найдено анализа для копирования.", "uk": "Не знайдено аналізу для копіювання.", "it": "Nessuna analisi trovata da copiare."},
    "error_title": {"lv": "Kļūda", "en": "Error", "ru": "Ошибка", "uk": "Помилка", "it": "Errore"},
    "error_docx_missing": {"lv": "Lai eksportētu uz Word, nepieciešama 'python-docx' bibliotēka.\nUzinstalējiet to ar 'pip install python-docx'.", "en": "To export to Word, the 'python-docx' library is required.\nInstall it with 'pip install python-docx'.", "ru": "Для экспорта в Word требуется библиотека 'python-docx'.\nУстановите ее с помощью 'pip install python-docx'.", "uk": "Для експорту в Word потрібна бібліотека 'python-docx'.\nВстановіть її за допомогою 'pip install python-docx'.", "it": "Per esportare in Word è necessaria la libreria 'python-docx'.\nInstallala con 'pip install python-docx'."},
    "info_title": {"lv": "Info", "en": "Info", "ru": "Информация", "uk": "Інформація", "it": "Info"},
    "info_no_data_to_export": {"lv": "Nav datu, ko eksportēt. Palaidiet 'Medīt Jaunos Vienradžus' vispirms.", "en": "No data to export. Run 'Hunt New Unicorns' first.", "ru": "Нет данных для экспорта. Сначала запустите 'Искать новых единорогов'.", "uk": "Немає даних для експорту. Спочатку запустіть 'Полювання на Єдинорогів'.", "it": "Nessun dato da esportare. Avvia prima 'Caccia agli Unicorni'."},
    "success_title": {"lv": "Saglabāts!", "en": "Saved!", "ru": "Сохранено!", "uk": "Збережено!", "it": "Salvato!"},
    "error_saving_word_title": {"lv": "Kļūda saglabājot Word failu", "en": "Error saving Word file", "ru": "Ошибка при сохранении файла Word", "uk": "Помилка збереження faila Word", "it": "Errore nel salvataggio del file Word"},

    # Word Export specific strings
    "word_report_title": {"lv": "🚀 Komandanta Kvarka Vienradžu Medību Atskaite", "uk": "🚀 Звіт про Полювання на Єдинорогів Командира Кварка", "ru": "🚀 Отчет об Охоте на Единорогов Командира Кварка", "it": "🚀 Rapporto di Caccia agli Unicorni del Comandante Quark", "en": "🚀 Commander Quark's Unicorn Hunt Report"},
    "word_generated_on": {"lv": "Ģenerēts: ", "uk": "Згенеровано: ", "ru": "Сгенерировано: ", "it": "Generato il: ", "en": "Generated: "},
    "word_analysis_basis": {"lv": "Analīze ģenerēta, balstoties uz Komandanta Kvarka 4 zelta kritērijiem.", "uk": "Аналіз згенеровано на основі 4 золотих критеріїв Командира Кварка.", "ru": "Анализ сгенерирован на основе 4 золотых критериев Командира Кварка.", "it": "Analisi generata in base ai 4 criteri d'oro del Comandante Quark.", "en": "Analysis generated based on Commander Quark's 4 golden criteria."},
    "word_red_flag": {"lv": "🚩 Sarkanais Karogs: ", "uk": "🚩 Червоний Прапор: ", "ru": "🚩 Красный Флаг: ", "it": "🚩 Bandiera Rossa: ", "en": "🚩 Red Flag: "},
    "word_secret_weapon": {"lv": "Slepenais ierocis (Niša): ", "uk": "Секретна зброя (Ніша): ", "ru": "Секретное оружие (Ниша): ", "it": "Arma Segreta (Nicchia): ", "en": "Secret Weapon (Niche): "},
    "word_why_10x_100x": {"lv": "Kāpēc 10x-100x: ", "uk": "Чому 10x-100x: ", "ru": "Почему 10x-100x: ", "it": "Perché 10x-100x: ", "en": "Why 10x-100x: "},
    "word_black_swan": {"lv": "Melnais Gulbis (risks): ", "uk": "Чорний Лебīдь (ризик): ", "ru": "Черный Лебедь (риск): ", "it": "Cigno Nero (rischio): ", "en": "Black Swan (risk): "},
    "word_output_filename_prefix": {"lv": "Kvarka_Vienradzu_Medibas_", "uk": "Kvar_Yedinorih_Polyuvannya_", "ru": "Kvar_Ohot_Yedinorog_", "it": "Quark_Caccia_Unicorni_", "en": "Quarks_Unicorn_Hunt_"},
    "word_success_message": {"lv": "Atskaite saglabāta:\n", "uk": "Звіт збережено:\n", "ru": "Отчет сохранен в:\n", "it": "Rapporto salvato in:\n", "en": "Report saved to:\n"},
    "word_error_saving_message": {"lv": "Radās kļūda saglabājot Word failu: ", "uk": "Виникла помилка під час збереження файlu Word: ", "ru": "Произошла ошибка при сохранении файла Word: ", "it": "Error saving Word file: ", "en": "Error saving Word file: "},

    # Dividends Tab:
    "lbl_cashflow_radar_title": {"lv": "💰 NAUDAS PLŪSMAS RADARS", "en": "💰 CASH FLOW RADAR", "ru": "💰 РАДАР ДЕНЕЖНОГО ПОТОКА", "uk": "💰 РАДАР ГРОШОВОГО ПОТОКУ", "it": "💰 RADAR DEL FLUSSO DI CASSA"},
    "lbl_cashflow_radar_desc": {"lv": "Aprēķina pasīvo dividenžu ienesīgumu no portfeļa akcijām.", "en": "Calculates passive dividend yield from portfolio stocks.", "ru": "Рассчитывает пассивную дивидендную доходность от акций портфеля.", "uk": "Розраховує пасивний дивідендний дохід від акцій портфеля.", "it": "Calcola il rendimento da dividendi passivi dalle azioni del portafoglio."},
    "btn_calc_dividends": {"lv": "⚡ Aprēķināt Dividenžu Plūsmu", "en": "⚡ Calculate Dividend Flow", "ru": "⚡ Рассчитать Дивидендный Поток", "uk": "⚡ Розрахувати Дивідендний Поток", "it": "⚡ Calcola Flusso Dividendi"},
    "dividends_initial_msg": {"lv": "Nospiediet pogu, lai nolasītu dividenžu datus...", "en": "Press the button to fetch dividend data...", "ru": "Нажмите кнопку, чтобы получить данные о дивидендах...", "uk": "Натисніть кнопку, щоб отримати дані про дивіденди...", "it": "Premi il pulsante per recuperare i dati sui dividendi..."},
    "dividends_initial_analysis_msg": {"lv": "Gaidu dividenžu skenēšanas palaišanu...", "en": "Awaiting dividend scan initiation...", "ru": "Ожидаю запуск сканирования дивидендов...", "uk": "Очікую запуск сканування дивідендів...", "it": "In attesa dell'avvio della scansione dei dividendi..."},
    "btn_calc_dividends_loading": {"lv": "⏳ Vāc dividenžu datus...", "en": "⏳ Collecting dividend data...", "ru": "⏳ Собирает данные о дивидендах...", "uk": "⏳ Збирає datus par dividendēm...", "it": "⏳ Raccolta dati dividendi..."},
    "dividends_scan_progress1": {"lv": "Pārbauda {len_tickers} aktīvus tirgū...\n", "en": "Checking {len_tickers} assets in the market...\n", "ru": "Проверяет {len_tickers} активов на рынке...\n", "uk": "Перевіряє {len_tickers} активів на ринку...\n", "it": "Controllando {len_tickers} asset sul mercato...\n"},
    "dividends_scan_progress2": {"lv": "Kvarks apkopo naudas plūsmas kalendāru...\n", "en": "Quark compiles the cash flow calendar...\n", "ru": "Кварк составляет календарь денежного потока...\n", "uk": "Кварк складає календар грошового потоку...\n", "it": "Quark compila il calendario dei flussi di cassa...\n"},
    "dividends_summary_header": {"lv": "=== DIVIDENŽU PĀRSKATS ===", "en": "=== DIVIDEND OVERVIEW ===", "ru": "=== ОБЗОР ДИВИДЕНДОВ ===", "uk": "=== ОГЛЯД ДИВІДЕНДІВ ===", "it": "=== RIEPILOGO DIVIDENDI ==="},
    "dividends_date_na": {"lv": "Nav norādīts", "en": "Not specified", "ru": "Не указано", "uk": "Не вказано", "it": "Non specificato"},
    "dividends_not_paying": {"lv": "Dividendes nemaksā", "en": "Does not pay dividends", "ru": "Дивиденды не выплачивает", "uk": "Дивіденди не виплачує", "it": "Non paga dividendi"},
    "dividends_error_fetching": {"lv": "Kļūda: {e}", "en": "Error: {e}", "ru": "Ошибка: {e}", "uk": "Помилка: {e}", "it": "Errore: {e}"},
    "dividends_champions": {"lv": "💵 PASĪVĀS NAUDAS ČEMPIONI", "en": "💵 PASSIVE CASH CHAMPIONS", "ru": "💵 ЧЕМПИОНЫ ПО ПАССИВНЫМ ДЕНЬГАМ", "uk": "💵 ЧЕМПІОНИ З ПАСИВНИХ ГРОШЕЙ", "it": "💵 CAMPIONI DEL FLUSSO DI CASSA PASSIVO"},
    "dividends_sustainability_risks": {"lv": "⚠️ ILGTSPĒJAS RISKI", "en": "⚠️ SUSTAINABILITY RISKS", "ru": "⚠️ РИСКИ УСТОЙЧИВОСТИ", "uk": "⚠️ РИЗИКИ СТАЛОСТІ", "it": "⚠️ RISCHI DI SOSTENIBILITÀ"},
    "dividends_reinvestment_strategy": {"lv": "🔄 REINVESTĒŠANAS STRATĒĢIJA", "en": "🔄 REINVESTMENT STRATEGY", "ru": "🔄 СТРАТЕГИЯ РЕИНВЕСТИРОВАНИЯ", "uk": "🔄 СТРАТЕГІЯ РЕІНВЕСТУВАННЯ", "it": "🔄 STRATEGIA DI REINVESTIMENTO"},
    "error_gemini_dividends": {"lv": "⚠️ Gemini kļūda: {e}", "en": "⚠️ Gemini error: {e}", "ru": "⚠️ Ошибка Gemini: {e}", "uk": "⚠️ Помилка Gemini: {e}", "it": "⚠️ Errore Gemini: {e}"},
    "dividends_no_companies": {"lv": "Portfelī nav atrasts neviens uzņēmums, kas maksā dividendes.", "en": "No dividend-paying companies found in portfolio.", "ru": "В портфеле не найдено ни одной компании, выплачивающей дивиденды.", "uk": "У портфелі не знайдено жодної компанії, яка виплачує дивіденди.", "it": "Nessuna azienda che paga dividendi trovata nel portfolio."},
    "error_api_key_unavailable": {"lv": "Nav pieejama API atslēga.", "en": "API key not available.", "ru": "Ключ API недоступен.", "uk": "Ключ API недоступний.", "it": "Chiave API non disponibile."},

    # Market Tab:
    "lbl_market_sentiment_title": {"lv": "🌡️ TIRGUS NOSKAŅOJUMS", "en": "🌡️ MARKET SENTIMENT", "ru": "🌡️ НАСТРОЕНИЕ РЫНКА", "uk": "🌡️ НАСТРІЙ РИНКУ", "it": "🌡️ SENTIMENTO DI MERCATO"},
    "lbl_fng_loading": {"lv": "Fear & Greed: Ielādē...", "en": "Fear & Greed: Loading...", "ru": "Индекс страха и жадности: Загрузка...", "uk": "Страх і жадібність: Завантаження...", "it": "Fear & Greed: Caricamento..."},
    "btn_refresh_market_intel": {"lv": "⚡ Atjaunot Klimatu & Insaiderus", "en": "⚡ Refresh Climate & Insiders", "ru": "⚡ Обновить Климат и Инсайдеров", "uk": "⚡ Оновити Клімат та Інсайдерів", "it": "⚡ Aggiorna Clima & Insider"},
    "insider_initial_msg": {"lv": "Nospiediet pogu, lai ielādētu insaideru datus un noskaņojumu...", "en": "Press the button to load insider data and sentiment...", "ru": "Нажмите кнопку, чтобы загрузить данные инсайдеров и настроение...", "uk": "Натисніть кнопку, щоб завантажити дані інсайдеров та настрій...", "it": "Premi il pulsante per caricare i dati e il sentimento degli insider..."},
    "market_news_initial_msg": {"lv": "Gaidu datu pieprasījumu...", "en": "Awaiting data request...", "ru": "Ожидаю запрос данных...", "uk": "Очікую запит даних...", "it": "In attesa di richiesta dati..."},
    "btn_refresh_market_intel_loading": {"lv": "⏳ Vāc makro datus & Kvarks analizē...", "en": "⏳ Collecting macro data & Quark analyzing...", "ru": "⏳ Собирает макроданные и Кварк анализирует...", "uk": "⏳ Збирає макродані та Кварк analizē...", "it": "⏳ Raccolta dati macro & Analisi di Quark..."},
    "fng_not_available": {"lv": "Fear & Greed: Nav pieejams", "en": "Fear & Greed: Not available", "ru": "Индекс страха и жадности: Недоступен", "uk": "Страх і жадібність: Недоступно", "it": "Fear & Greed: Non disponibile"},
    "macro_liquidity_header": {"lv": "=== MAKRO & LIKVIDITĀTE ===", "en": "=== MACRO & LIQUIDITY ===", "ru": "=== МАКРО И ЛИКВИДНОСТЬ ===", "uk": "=== МАКРО ТА ЛІКВІДНІСТЬ ===", "it": "=== MACRO & LIQUIDITÀ ==="},
    "macro_us10y": {"lv": "ASV 10Y Yield", "en": "US 10Y Yield", "ru": "Доходность 10-летних облигаций США", "uk": "Прибутковість 10-річних облігацій США", "it": "Rendimento USA 10Y"},
    "macro_gold": {"lv": "Zelts (Gold)", "en": "Gold", "ru": "Золото (Gold)", "uk": "Золото (Gold)", "it": "Oro (Gold)"},
    "macro_bitcoin": {"lv": "Bitcoin (BTC)", "en": "Bitcoin (BTC)", "ru": "Биткойн (BTC)", "uk": "Біткойн (BTC)", "it": "Bitcoin (BTC)"},
    "macro_na": {"lv": "N/A", "en": "N/A", "ru": "Н/Д", "uk": "Н/Д", "it": "N/D"},
    "insider_transactions_header": {"lv": "\n=== INSAIDERU TRANSAKCIJAS (Finnhub) ===", "en": "\n=== INSIDER TRANSACTIONS (Finnhub) ===", "ru": "\n=== ИНСАЙДЕРСКИЕ СДЕЛКИ (Finnhub) ===", "uk": "\n=== ІНСАЙДЕРСЬКІ ОПЕРАЦІЇ (Finnhub) ===", "it": "\n=== TRANSAZIONI INSIDER (Finnhub) ==="},
    "insider_buy": {"lv": "🟢 PĒRK", "en": "🟢 BUY", "ru": "🟢 ПОКУПАЕТ", "uk": "🟢 КУПУЄ", "it": "🟢 ACQUISTA"},
    "insider_sell": {"lv": "🔴 PĀRDOD", "en": "🔴 SELL", "ru": "🔴 ПРОДАЕТ", "uk": "🔴 ПРОДАЄ", "it": "🔴 VENDE"},
    "insider_neutral": {"lv": "⚪ NEITRĀLS", "en": "⚪ NEUTRAL", "ru": "⚪ НЕЙТРАЛЬНО", "uk": "⚪ НЕЙТРАЛЬНО", "it": "⚪ NEUTRALE"},
    "insider_no_major_deals": {"lv": "⚪ Mierīgs (nav lielu darījumu)", "en": "⚪ Calm (no major deals)", "ru": "⚪ Спокойно (нет крупных сделок)", "uk": "⚪ Спокійно (немає великих угод)", "it": "⚪ Calmo (nessuna transazione importante)"},
    "insider_data_error": {"lv": "Kļūda datos", "en": "Error in data", "ru": "Ошибка в данных", "uk": "Помилка в даних", "it": "Errore nei dati"},
    "error_finnhub_key_missing": {"lv": "⚠️ config.json trūkst 'finnhub_api_key'!", "en": "⚠️ 'finnhub_api_key' missing from config.json!", "ru": "⚠️ В config.json отсутствует 'finnhub_api_key'!", "uk": "⚠️ У config.json відсутній 'finnhub_api_key'!", "it": "⚠️ 'finnhub_api_key' mancante in config.json!"},
    "news_headlines_header": {"lv": "=== SVARĪGĀKĀS ZIŅAS ===", "en": "=== TOP HEADLINES ===", "ru": "=== ВАЖНЕЙШИЕ НОВОСТИ ===", "uk": "=== НАЙВАЖЛІВІШІ НОВИНИ ===", "it": "=== TITOLI PRINCIPALI ==="},
    "news_no_aggressive": {"lv": "Nav agresīvu ziņu triecienu", "en": "No aggressive news strikes", "ru": "Нет агрессивных новостных ударов", "uk": "Немає агресивних новинних ударів", "it": "Nessun attacco di notizie aggressive"},
    "kvark_strategic_resolution": {"lv": "🐾 KOMANDANTA KVARKA STRATĒĢISKĀ REZOLŪCIJA:\n", "en": "🐾 COMMANDER KVARK'S STRATEGIC RESOLUTION:\n", "ru": "🐾 СТРАТЕГИЧЕСКАЯ РЕЗОЛЮЦИЯ КОМАНДИРА КВАРКА:\n", "uk": "🐾 СТРАТЕГІЧНА РЕЗОЛЮЦІЯ КОМАНДИРА КВАРКА:\n", "it": "🐾 RISOLUZIONE STRATEGICA DEL COMANDANTE QUARK:\n"},
    "error_kvark_market_intel": {"lv": "⚠️ Kvarka kļūda: {e}", "en": "⚠️ Quark's error: {e}", "ru": "⚠️ Ошибка Кварка: {e}", "uk": "⚠️ Помилка Кварka: {e}", "it": "⚠️ Errore di Quark: {e}"},

    # Shared/Generic:
    "rentgen_initial_msg": {"lv": "Ievadiet tikeri un nospiediet 'Rentgens'...", "en": "Enter ticker and press 'X-Ray'...", "ru": "Введите тикер и нажмите 'Рентген'...", "uk": "Введіть тікер і натисніть 'Рентген'...", "it": "Inserisci il ticker e premi 'Raggi X'..."},
    "rentgen_generating_analysis_msg": {"lv": "Kvarks ģenerē fundamentālo analīzi...\n", "en": "Quark is generating fundamental analysis...\n", "ru": "Кварк генерирует фундаментальный анализ...\n", "uk": "Кварк генерує фундаментальний аналіз...\n", "it": "Quark sta generando l'analisi fondamentale...\n"},
    "source_changed_msg": {"lv": "Avota kanāls pārslēgts uz: {source}. Gaidu jaunu signālu 'Rentgens'!", "en": "Source channel switched to: {source}. Awaiting new 'X-Ray' signal!", "ru": "Канал источника переключен на: {source}. Ожидаю новый сигнал 'Рентген'!", "uk": "Канал джерела перемкнено на: {source}. Очікую новий сигнал 'Рентген'!", "it": "Canale sorgente cambiato in: {source}. In attesa di nuovo segnale 'Raggi X'!"},
    "fetching_data_msg": {"lv": "Nolasa tirgus datus par {ticker} no {source}...\n", "en": "Fetching market data for {ticker} from {source}...\n", "ru": "Получение рыночных данных для {ticker} из {source}...\n", "uk": "Отримання риinķa datus par {ticker} no {source}...\n", "it": "Recupero dati di mercato per {ticker} da {source}...\n"},
    "ticker_info_header": {"lv": "=== {short_name} ({ticker}) ===", "en": "=== {short_name} ({ticker}) ===", "ru": "=== {short_name} ({ticker}) ===", "uk": "=== {short_name} ({ticker}) ===", "it": "=== {short_name} ({ticker}) ==="},
    "price_label": {"lv": "Cena:             {price} {currency}", "en": "Price:             {price} {currency}", "ru": "Цена:             {price} {currency}", "uk": "Ціна:             {price} {currency}", "it": "Prezzo:             {price} {currency}"},
    "52week_range_label": {"lv": "52 ned. diapazons:{low_52} - {high_52} {currency}", "en": "52 wk. range:    {low_52} - {high_52} {currency}", "ru": "Диапазон 52 нед.:{low_52} - {high_52} {currency}", "uk": "Діапазон 52 тижнів:{low_52} - {high_52} {currency}", "it": "Intervallo 52 sett.:{low_52} - {high_52} {currency}"},
    "pe_label": {"lv": "P/E (faktiskais): {pe}", "en": "P/E (actual):    {pe}", "ru": "P/E (фактический): {pe}", "uk": "P/E (фактичний): {pe}", "it": "P/E (effettivo): {pe}"},
    "fwd_pe_label": {"lv": "P/E (prognoze):   {fwd_pe}", "en": "P/E (forecast):  {fwd_pe}", "ru": "P/E (прогноз):   {fwd_pe}", "uk": "P/E (прогноз):   {fwd_pe}", "it": "P/E (previsto):   {fwd_pe}"},
    "peg_ratio_label": {"lv": "PEG koeficients:  {peg}", "en": "PEG ratio:       {peg}", "ru": "Коэффициент PEG:  {peg}", "uk": "Коефіcients PEG:  {peg}", "it": "Rapporto PEG:      {peg}"},
    "dividends_label": {"lv": "Dividendes:       {div_str}", "en": "Dividends:       {div_str}", "ru": "Дивиденды:       {div_str}", "uk": "Дивіденди:       {div_str}", "it": "Dividendi:       {div_str}"},
    "analyst_rating_label": {"lv": "Analītiķu vērt.:  {rec}", "en": "Analyst rating:  {rec}", "ru": "Оценка аналитиков:  {rec}", "uk": "Оцінка аналітиків:  {rec}", "it": "Valutazione analisti:  {rec}"},
    "target_price_label": {"lv": "Mērķa vid. cena:  {target_mean} {currency}", "en": "Target avg. price: {target_mean} {currency}", "ru": "Целевая сред. цена: {target_mean} {currency}", "uk": "Цільова сер. ціena: {target_mean} {currency}", "it": "Prezzo medio target: {target_mean} {currency}"},
    "unknown_value": {"lv": "Nezināms", "en": "Unknown", "ru": "Неизвестно", "uk": "Невідомо", "it": "Sconosciuto"},
    "error_fetching_ticker_data": {"lv": "⚠️ Kļūda nolasot {ticker}: {e}", "en": "⚠️ Error fetching {ticker}: {e}", "ru": "⚠️ Ошибка при получении {ticker}: {e}", "uk": "⚠️ Помилка pri iegūšanā {ticker}: {e}", "it": "⚠️ Errore nel recupero {ticker}: {e}"},
    "analysis_disclaimer_summary": {"lv": "Informatīvs analīzes kopsavilkums.", "en": "Informative analysis summary.", "ru": "Информативное резюме анализа.", "uk": "Інформативне резюме аналізу.", "it": "Riepilogo analitico informativo."},

    # Additional keys for prompts and UI (LEO ADDITION)
    "lbl_inv_portfolio_audit_plan": {"lv": "🐾 KOMANDANTA KVARKA AUDITA PLĀNS", "en": "🐾 COMMANDER KVARK'S AUDIT PLAN", "ru": "🐾 ПЛАН АУДИТА КОМАНДИРА КВАРКА", "uk": "🐾 ПЛАН АУДИТУ КОМАНДИРА КВАРКА", "it": "🐾 PIANO DI AUDIT DEL COMANDANTE KVARKS"},
    "audit_sell_liquidate": {"lv": "✂️ LIKVIDĒT / PĀRDOT", "en": "✂️ LIQUIDATE / SELL", "ru": "✂️ ЛИКВИДИРОВАТЬ / ПРОДАТЬ", "uk": "✂️ ЛІКВІДУВАТИ / ПРОДАТИ", "it": "✂️ LIQUIDARE / VENDERE"},
    "audit_cash_allocation": {"lv": "🎯 BRĪVO LĪDZEKĻU SADALE", "en": "🎯 FREE CASH ALLOCATION", "ru": "🎯 РАСПРЕДЕЛЕНИЕ СВОБОДНЫХ СРЕДСТВ", "uk": "🎯 РОЗПОДІЛ ВІЛЬНИХ КОШТІВ", "it": "🎯 ALLOCAZIONE DEL CAPITALE DISPONIBILE"},
    "llm_default_lang_instruction": {"lv": "Atbildi latviešu valodā.", "en": "Respond in English.", "ru": "Ответь на русском языке.", "uk": "Відповідайте українською мовою.", "it": "Rispondi in italiano."},
    "source_label": {"lv": "Avots", "en": "Source", "ru": "Источник", "uk": "Джерело", "it": "Fonte"},
    "error_gemini_analysis": {"lv": "⚠️ Gemini analīzes kļūda: {e}", "en": "⚠️ Gemini analysis error: {e}", "ru": "⚠️ Ошибка анализа Gemini: {e}", "uk": "⚠️ Помилка аналізу Gemini: {e}", "it": "⚠️ Errore di analisi Gemini: {e}"}
}

# Transform TRANSLATIONS into i18n-compatible format for setting the dictionary
# This should only happen ONCE upon module load.
if HAS_I18N:
    # Register translations dictionary with the i18n module as per UZDEVUMS.
    i18n.register_translations(TRANSLATIONS)
    # The AuraQuadro local i18n module is stated to contain a global TRANSLATIONS dictionary
    # and does not support/require set_dictionary. It expects TRANSLATIONS to be defined
    # at module level, and directly uses it.
    # Removed i18n.get() and i18n.set() calls as per instructions.
else:
    # Define dummy functions if i18n is not available to prevent crashes
    class DummyI18nModule: # Create a dummy class to emulate the module
        def get_ui_language(self): return "lv" # Default to Latvian
        def t(self, key, **kwargs): # Dummy t function uses global TRANSLATIONS
            # This logic mimics the old custom t for fallback behavior
            current_lang = "lv" # Default for dummy
            translated_text = TRANSLATIONS.get(key, {}).get(current_lang, TRANSLATIONS.get(key, {}).get("en", key))
            if kwargs:
                try:
                    return translated_text.format(**kwargs)
                except KeyError:
                    return translated_text + f" (Formatting error for key '{key}')"
            return translated_text

        def set_dictionary(self, data): pass # Dummy
        def get(self, key): # Dummy for get('locale')
            if key == 'locale': return "lv"
            return None
        def set(self, key, value): pass # Dummy for i18n.set method

    i18n = DummyI18nModule() # Instantiate the dummy module
    # The get_llm_language_instruction in this else block should use the global TRANSLATIONS
    def get_llm_language_instruction():
        return i18n.t("llm_default_lang_instruction")

# Helper function for localization, now just an alias to i18n.t (NEW)
# No longer defines a custom 't' function as per instructions to use i18n.t
# All calls to 't(key)' will now be replaced with 'i18n.t(key)'
# The direct i18n.t() calls can accept kwargs for formatting.


GROWTH_WATCHLIST = [
    {"ticker": "PLTR", "name": "Palantir (Enterprise AI / Defense)"},
    {"ticker": "ARM", "name": "ARM Holdings (AI & Mobile Chips)"},
    {"ticker": "CRWD", "name": "CrowdStrike (Next-Gen Cybersecurity)"},
    {"ticker": "SMCI", "name": "Super Micro Computer (AI Servers)"},
    {"ticker": "SNOW", "name": "Snowflake (Data Cloud / AI infra)"},
    {"ticker": "RKLB", "name": "Rocket Lab (Space Economy / Satellites)"},
    {"ticker": "IONQ", "name": "IonQ (Quantum Computing Pioneer)"}
]


class InvestorStudioFrame(ctk.CTkFrame):
    def __init__(self, master, get_gemini_key_fn=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.get_gemini_key = get_gemini_key_fn

        # Initialize attribute for sector concentration warning (FIXED for Point 1, 3)
        self.last_sector_concentration_warning = ""

        # NEW: Initialize selected data source
        self.selected_source = "Yahoo Finance"

        # Augšējā ciļņu josla moduļa iekšienē
        self.tabview = ctk.CTkTabview(self, fg_color="#0f172a", segmented_button_fg_color="#1e293b")
        self.tabview.pack(fill="both", expand=True, padx=10, pady=10)

        # Tab creation using i18n.t for initial names (NEW)
        self.tab_rentgen = self.tabview.add(i18n.t("tab_inv_single"))
        self.tab_portfolio = self.tabview.add(i18n.t("tab_inv_etoro"))
        self.tab_growth = self.tabview.add(i18n.t("tab_inv_unicorns"))
        self.tab_dividends = self.tabview.add(i18n.t("tab_inv_dividends"))
        self.tab_market = self.tabview.add(i18n.t("tab_inv_insiders"))

        self._setup_rentgen_tab()
        self._setup_portfolio_tab()
        self._setup_growth_tab()
        self._setup_dividends_tab()
        self._setup_market_tab()

    def refresh_ui(self):
        """
        Refreshes all localizable UI widget texts based on the current i18n language.
        This method should be called when the UI language changes.
        """
        # Update main tabview segmented buttons (NEW for Instruction 3)
        if hasattr(self, "tabview"):
            try:
                # Get references to the actual buttons within the segmented_button widget
                btns = list(self.tabview._segmented_button._buttons_dict.values())
                keys = ["tab_inv_single", "tab_inv_etoro", "tab_inv_unicorns", "tab_inv_dividends", "tab_inv_insiders"]
                for i, k in enumerate(keys):
                    if i < len(btns):
                        btns[i].configure(text=i18n.t(k))
            except Exception as e:
                # print(f"Error refreshing tabview buttons: {e}") # Debugging
                pass

        # Update elements in Rentgen tab
        if hasattr(self, 'lbl_ticker_title'):
            self.lbl_ticker_title.configure(text=i18n.t("lbl_ticker_analysis"))
        if hasattr(self, 'btn_analyze'):
            current_btn_text = self.btn_analyze.cget("text")
            if not current_btn_text.startswith("⏳"):
                self.btn_analyze.configure(text=i18n.t("btn_rentgens"))
        if hasattr(self, "lbl_resolution_title"): # NEW for Instruction 3
            self.lbl_resolution_title.configure(text=i18n.t("lbl_inv_resolution_title"))

        # Update elements in Portfolio tab
        if hasattr(self, 'lbl_portfolio_title'):
            self.lbl_portfolio_title.configure(text=i18n.t("lbl_portfolio_title"))
        if hasattr(self, 'lbl_portfolio_tickers_desc'):
            self.lbl_portfolio_tickers_desc.configure(text=i18n.t("lbl_portfolio_tickers_desc"))
        if hasattr(self, 'lbl_portfolio_cash_balance'):
            self.lbl_portfolio_cash_balance.configure(text=i18n.t("lbl_portfolio_cash_balance"))
        if hasattr(self, 'lbl_portfolio_risk_strategy'):
            self.lbl_portfolio_risk_strategy.configure(text=i18n.t("lbl_portfolio_risk_strategy"))
        if hasattr(self, 'seg_risk'):
            self.seg_risk.configure(values=[i18n.t("seg_risk_conservative"), i18n.t("seg_risk_venture_growth")])
            # Re-set the selected value to ensure the text updates
            current_selection = i18n.t("seg_risk_conservative") if "Konservatīvs" in self.seg_risk.get() else i18n.t("seg_risk_venture_growth")
            self.seg_risk.set(current_selection)
        if hasattr(self, 'btn_save_p'):
            self.btn_save_p.configure(text=i18n.t("btn_save_portfolio"))
        if hasattr(self, 'btn_audit'):
            current_btn_text = self.btn_audit.cget("text")
            if not current_btn_text.startswith("⏳"):
                self.btn_audit.configure(text=i18n.t("btn_audit_portfolio"))
        if hasattr(self, 'btn_copy_audit'):
            self.btn_copy_audit.configure(text=i18n.t("btn_copy_summary"))
        if hasattr(self, 'lbl_inv_portfolio_audit_plan'): # LEO ADDITION
            self.lbl_inv_portfolio_audit_plan.configure(text=i18n.t("lbl_inv_portfolio_audit_plan")) # LEO ADDITION


        # Update elements in Growth tab
        if hasattr(self, 'lbl_growth_radar_title'):
            self.lbl_growth_radar_title.configure(text=i18n.t("lbl_potential_radar")) # Re-using existing key
        if hasattr(self, 'lbl_growth_description'):
            self.lbl_growth_description.configure(text=i18n.t("lbl_growth_description"))
        if hasattr(self, 'btn_scan_all_growth'):
            current_btn_text = self.btn_scan_all_growth.cget("text")
            if not current_btn_text.startswith("⏳"):
                self.btn_scan_all_growth.configure(text=i18n.t("btn_hunt_unicorns"))
        if hasattr(self, 'lbl_growth_resolution_title'): # NEW for Instruction 2
            self.lbl_growth_resolution_title.configure(text=i18n.t("lbl_inv_growth_resolution_title"))
        if hasattr(self, 'btn_copy_growth'):
            self.btn_copy_growth.configure(text=i18n.t("btn_copy_analysis"))
        if hasattr(self, 'btn_export_word'):
            self.btn_export_word.configure(text=i18n.t("btn_save_word"))
        
        # Update elements in Dividends tab
        if hasattr(self, 'lbl_cashflow_radar_title'):
            self.lbl_cashflow_radar_title.configure(text=i18n.t("lbl_cashflow_radar_title"))
        if hasattr(self, 'lbl_cashflow_radar_desc'):
            self.lbl_cashflow_radar_desc.configure(text=i18n.t("lbl_cashflow_radar_desc"))
        if hasattr(self, 'btn_calc_dividends'):
            current_btn_text = self.btn_calc_dividends.cget("text")
            if not current_btn_text.startswith("⏳"):
                self.btn_calc_dividends.configure(text=i18n.t("btn_calc_dividends"))
        if hasattr(self, 'lbl_dividends_resolution_title'): # NEW for Instruction 2
            self.lbl_dividends_resolution_title.configure(text=i18n.t("lbl_inv_dividends_resolution_title"))

        # Update elements in Market tab
        if hasattr(self, 'lbl_market_sentiment_title'):
            self.lbl_market_sentiment_title.configure(text=i18n.t("lbl_market_sentiment_title"))
        if hasattr(self, 'lbl_fng'):
            self.lbl_fng.configure(text=i18n.t("lbl_fng_loading")) # Reset to loading state text, actual value refreshed by worker
        if hasattr(self, 'btn_refresh_market'):
            current_btn_text = self.btn_refresh_market.cget("text")
            if not current_btn_text.startswith("⏳"):
                self.btn_refresh_market.configure(text=i18n.t("btn_refresh_market_intel"))
        if hasattr(self, 'lbl_market_news_title'): # NEW for Instruction 2
            self.lbl_market_news_title.configure(text=i18n.t("lbl_inv_market_news_title"))


        # Note: Dynamic content in textboxes or scrollable frames (e.g., stats_box, analysis_box, growth_results_container)
        # would require re-running their content generation logic if they contained fixed strings.
        # This refresh focuses on static widget properties as requested.

    # ==========================================
    # 1. CILNE: VIENA TIKERA RENTGENS
    # ==========================================
    def _setup_rentgen_tab(self):
        f = self.tab_rentgen
        f.grid_columnconfigure(0, weight=1)
        f.grid_columnconfigure(1, weight=2)
        f.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(f, fg_color="#131b2e", corner_radius=10)
        left.grid(row=0, column=0, padx=(5, 5), pady=5, sticky="nsew")

        right = ctk.CTkFrame(f, fg_color="#131b2e", corner_radius=10)
        right.grid(row=0, column=1, padx=(5, 5), pady=5, sticky="nsew")

        # Assigned to self.lbl_ticker_title for refresh_ui method
        self.lbl_ticker_title = ctk.CTkLabel(left, text=i18n.t("lbl_ticker_analysis"), font=ctk.CTkFont(size=14, weight="bold"), text_color="#38bdf8")
        self.lbl_ticker_title.pack(anchor="w", padx=15, pady=(15, 6))

        # NEW: Data Source Selection Segmented Button
        self.source_selector = ctk.CTkSegmentedButton(
            left,
            values=["Yahoo Finance", "Alpha Vantage", "FMP", "⚔️ Cross-Check"],
            command=self._on_source_change,
            selected_color="#0284c7",
            selected_hover_color="#0369a1",
            height=30,
            font=ctk.CTkFont(size=11, weight="bold")
        )
        self.source_selector.set(self.selected_source) # Set default value
        self.source_selector.pack(fill="x", padx=15, pady=(0, 10))

        in_row = ctk.CTkFrame(left, fg_color="transparent")
        in_row.pack(fill="x", padx=15, pady=(0, 10))

        self.entry_ticker = ctk.CTkEntry(in_row, width=130, height=32, placeholder_text="piem., NVDA", font=ctk.CTkFont(size=13, weight="bold"))
        self.entry_ticker.pack(side="left", padx=(0, 6))
        self.entry_ticker.bind("<Return>", lambda e: self.start_single_analysis())

        self.btn_analyze = ctk.CTkButton(in_row, text=i18n.t("btn_rentgens"), width=85, height=32, fg_color="#0284c7", hover_color="#0369a1", font=ctk.CTkFont(size=12, weight="bold"), command=self.start_single_analysis)
        self.btn_analyze.pack(side="left")

        q_frame = ctk.CTkFrame(left, fg_color="transparent")
        q_frame.pack(fill="x", padx=15, pady=(0, 8))
        for t_ticker_btn in ["NVDA", "AAPL", "MSFT", "VOO", "PLTR"]:
            ctk.CTkButton(q_frame, text=t_ticker_btn, width=48, height=22, fg_color="#1e293b", hover_color="#334155", font=ctk.CTkFont(size=10), command=lambda v=t_ticker_btn: self._quick_set(v)).pack(side="left", padx=2)

        self.stats_box = ctk.CTkTextbox(left, fg_color="#0b1329", font=ctk.CTkFont(family="Consolas", size=11), corner_radius=8)
        self.stats_box.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.stats_box.insert("1.0", i18n.t("rentgen_initial_msg"))
        self.stats_box.configure(state="disabled")

        # Assigned to self.lbl_resolution_title for refresh_ui method (NEW for Instruction 2)
        self.lbl_resolution_title = ctk.CTkLabel(right, text=i18n.t("lbl_inv_resolution_title"), font=ctk.CTkFont(size=14, weight="bold"), text_color="#f59e0b")
        self.lbl_resolution_title.pack(anchor="w", padx=15, pady=(15, 6))
        self.analysis_box = ctk.CTkTextbox(right, fg_color="#0b1329", text_color="#f8fafc", font=ctk.CTkFont(size=12), wrap="word", corner_radius=8)
        self.analysis_box.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.analysis_box.insert("1.0", i18n.t("txt_waiting"))
        self.analysis_box.configure(state="disabled")

    def _on_source_change(self, val):
        """Callback method for the data source segmented button."""
        self.selected_source = val
        self.stats_box.configure(state="normal")
        self.stats_box.delete("1.0", "end")
        self.stats_box.insert("1.0", i18n.t("source_changed_msg", source=self.selected_source))
        self.stats_box.configure(state="disabled")

    def _quick_set(self, t_ticker_btn):
        self.entry_ticker.delete(0, "end")
        self.entry_ticker.insert(0, t_ticker_btn)
        self.start_single_analysis()

    def start_single_analysis(self):
        t_ticker_analysis = self.entry_ticker.get().strip().upper()
        if not t_ticker_analysis:
            return
        self.btn_analyze.configure(state="disabled", text=i18n.t("btn_hunt_unicorns_loading", btn_text=i18n.t("btn_rentgens"))) # Dynamic loading text
        self.stats_box.configure(state="normal")
        self.stats_box.delete("1.0", "end")
        self.stats_box.insert("1.0", i18n.t("fetching_data_msg", ticker=t_ticker_analysis, source=self.selected_source))
        self.stats_box.configure(state="disabled")

        self.analysis_box.configure(state="normal")
        self.analysis_box.delete("1.0", "end")
        self.analysis_box.insert("1.0", i18n.t("rentgen_generating_analysis_msg"))
        self.analysis_box.configure(state="disabled")

        threading.Thread(target=self._single_worker, args=(t_ticker_analysis,), daemon=True).start()

    def _single_worker(self, ticker):
        stats_text, payload = self._fetch_stock_data(ticker)
        ai_reply = self._ask_gemini_single(payload, stats_text)

        def _ui():
            self.btn_analyze.configure(state="normal", text=i18n.t("btn_rentgens"))
            self.stats_box.configure(state="normal")
            self.stats_box.delete("1.0", "end")
            self.stats_box.insert("1.0", stats_text)
            self.stats_box.configure(state="disabled")

            self.analysis_box.configure(state="normal")
            self.analysis_box.delete("1.0", "end")
            self.analysis_box.insert("1.0", ai_reply)
            self.analysis_box.configure(state="disabled")

        self.after(0, _ui)

    # ==========================================
    # 2. CILNE: PORTFEĻA AUDITS & OPTIMIZĀCIJA
    # ==========================================
    def _setup_portfolio_tab(self):
        f = self.tab_portfolio
        f.grid_columnconfigure(0, weight=1)
        f.grid_columnconfigure(1, weight=2)
        f.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(f, fg_color="#131b2e", corner_radius=10)
        left.grid(row=0, column=0, padx=(5, 5), pady=5, sticky="nsew")

        right = ctk.CTkFrame(f, fg_color="#131b2e", corner_radius=10)
        right.grid(row=0, column=1, padx=(5, 5), pady=5, sticky="nsew")

        self.lbl_portfolio_title = ctk.CTkLabel(left, text=i18n.t("lbl_portfolio_title"), font=ctk.CTkFont(size=14, weight="bold"), text_color="#38bdf8")
        self.lbl_portfolio_title.pack(anchor="w", padx=15, pady=(15, 4))
        self.lbl_portfolio_tickers_desc = ctk.CTkLabel(left, text=i18n.t("lbl_portfolio_tickers_desc"), font=ctk.CTkFont(size=13), text_color="#94a3b8", justify="left")
        self.lbl_portfolio_tickers_desc.pack(anchor="w", padx=15, pady=(0, 6))

        self.txt_my_portfolio = ctk.CTkTextbox(left, fg_color="#0b1329", font=ctk.CTkFont(size=12), height=140)
        self.txt_my_portfolio.pack(fill="x", padx=15, pady=(0, 10))

        cash_row = ctk.CTkFrame(left, fg_color="transparent")
        cash_row.pack(fill="x", padx=15, pady=(0, 10))
        self.lbl_portfolio_cash_balance = ctk.CTkLabel(cash_row, text=i18n.t("lbl_portfolio_cash_balance"), font=ctk.CTkFont(size=12))
        self.lbl_portfolio_cash_balance.pack(side="left")
        self.entry_cash = ctk.CTkEntry(cash_row, width=90, height=28, font=ctk.CTkFont(size=12, weight="bold"))
        self.entry_cash.pack(side="left", padx=10)
        self.entry_cash.insert(0, "373.00")

        # Riska profila slēdzis
        self.lbl_portfolio_risk_strategy = ctk.CTkLabel(
            left,
            text=i18n.t("lbl_portfolio_risk_strategy"),
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#94a3b8"
        )
        self.lbl_portfolio_risk_strategy.pack(anchor="w", padx=15, pady=(4, 2))

        self.seg_risk = ctk.CTkSegmentedButton(
            left,
            values=[i18n.t("seg_risk_conservative"), i18n.t("seg_risk_venture_growth")],
            height=28,
            selected_color="#0284c7",
            selected_hover_color="#0369a1"
        )
        self.seg_risk.set(i18n.t("seg_risk_conservative"))
        self.seg_risk.pack(fill="x", padx=15, pady=(0, 12))

        # Pogu rinda
        btn_row = ctk.CTkFrame(left, fg_color="transparent")
        btn_row.pack(fill="x", padx=15, pady=(0, 10))

        self.btn_save_p = ctk.CTkButton(
            btn_row,
            text=i18n.t("btn_save_portfolio"),
            width=80,
            height=34,
            fg_color="#334155",
            hover_color="#475569",
            command=self.save_portfolio
        )
        self.btn_save_p.pack(side="left", padx=(0, 8))

        self.btn_audit = ctk.CTkButton(
            btn_row,
            text=i18n.t("btn_audit_portfolio"),
            height=34,
            fg_color="#059669",
            hover_color="#10b981",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.start_portfolio_audit
        )
        self.btn_audit.pack(side="left", fill="x", expand=True)

        self.lbl_p_status = ctk.CTkLabel(left, text="", font=ctk.CTkFont(size=11), text_color="#10b981")
        self.lbl_p_status.pack(anchor="w", padx=15)

        header_r = ctk.CTkFrame(right, fg_color="transparent")
        header_r.pack(fill="x", padx=15, pady=(15, 6))

        self.lbl_inv_portfolio_audit_plan = ctk.CTkLabel( # LEO ADDITION: Added as instance variable
            header_r,
            text=i18n.t("lbl_inv_portfolio_audit_plan"), # LEO ADDITION: Used new key
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#f59e0b"
        )
        self.lbl_inv_portfolio_audit_plan.pack(side="left")

        self.btn_copy_audit = ctk.CTkButton(
            header_r,
            text=i18n.t("btn_copy_summary"),
            width=150,
            height=28,
            fg_color="#334155",
            hover_color="#475569",
            font=ctk.CTkFont(size=11, weight="bold"),
            command=self.copy_audit_summary
        )
        self.btn_copy_audit.pack(side="right")

        self.audit_box = ctk.CTkTextbox(right, fg_color="#0b1329", text_color="#f8fafc", font=ctk.CTkFont(size=12), wrap="word", corner_radius=8)
        self.audit_box.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.audit_box.insert("1.0", i18n.t("audit_initial_msg"))
        self.audit_box.configure(state="disabled")

        self.load_portfolio()

    def save_portfolio(self):
        raw = self.txt_my_portfolio.get("1.0", "end").strip()
        tickers = [x.strip().upper() for x in raw.replace("\n", ",").split(",") if x.strip()]
        cash = self.entry_cash.get().strip()
        data = {"tickers": tickers, "cash_usd": cash}
        try:
            with open(PORTFOLIO_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self.lbl_p_status.configure(text=i18n.t("status_portfolio_saved"))
            self.after(3000, lambda: self.lbl_p_status.configure(text=""))
        except Exception as e:
            self.lbl_p_status.configure(text=i18n.t("error_portfolio_save", e=e))

    def load_portfolio(self):
        if PORTFOLIO_FILE.exists():
            try:
                with open(PORTFOLIO_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    tickers = data.get("tickers", [])
                    cash = data.get("cash_usd", "373.00")
                    self.txt_my_portfolio.delete("1.0", "end")
                    self.txt_my_portfolio.insert("1.0", ", ".join(tickers))
                    self.entry_cash.delete(0, "end")
                    self.entry_cash.insert(0, cash)
            except Exception:
                pass

    def _add_ticker_to_portfolio(self, ticker):
        current_tickers_raw = self.txt_my_portfolio.get("1.0", "end").strip()
        current_tickers = [x.strip().upper() for x in current_tickers_raw.replace("\n", ",").split(",") if x.strip()]
        
        if ticker.upper() not in current_tickers:
            current_tickers.append(ticker.upper())
            self.txt_my_portfolio.delete("1.0", "end")
            self.txt_my_portfolio.insert("1.0", ", ".join(current_tickers))
            self.save_portfolio() # Save and show status
            # Provide feedback in the growth tab as well
            if hasattr(self, 'growth_results_container'):
                feedback_label = ctk.CTkLabel(self.growth_results_container, text=i18n.t("feedback_added_to_portfolio", ticker=ticker), text_color="#10b981")
                feedback_label.pack(pady=5)
                self.after(3000, feedback_label.destroy)
        else:
            if hasattr(self, 'growth_results_container'):
                feedback_label = ctk.CTkLabel(self.growth_results_container, text=i18n.t("feedback_already_in_portfolio", ticker=ticker), text_color="#f59e0b")
                feedback_label.pack(pady=5)
                self.after(3000, feedback_label.destroy)


    def copy_audit_summary(self):
        text_to_copy = self.audit_box.get("1.0", "end").strip()
        if not text_to_copy or i18n.t("audit_initial_msg") in text_to_copy:
            return

        msg = (
            i18n.t("audit_copy_header") + "\n"
            f"{i18n.t('audit_copy_strategy')} {self.seg_risk.get() if hasattr(self, 'seg_risk') else i18n.t('seg_risk_conservative')}\n" # LEO ADDITION: Default with translated key
            "----------------------------------------\n\n"
            f"{text_to_copy}\n\n"
            "----------------------------------------\n"
            f"{i18n.t('audit_copy_footer')}"
        )

        try:
            self.clipboard_clear()
            self.clipboard_append(msg)
            self.update()

            orig_text = self.btn_copy_audit.cget("text")
            orig_color = self.btn_copy_audit.cget("fg_color")
            self.btn_copy_audit.configure(text=i18n.t("status_copied"), fg_color="#10b981")

            self.after(2000, lambda: self.btn_copy_audit.configure(text=orig_text, fg_color=orig_color))
        except Exception as e:
            print(i18n.t("error_copying", e=e))

    def start_portfolio_audit(self):
        raw = self.txt_my_portfolio.get("1.0", "end").strip()
        tickers = [x.strip().upper() for x in raw.replace("\n", ",").split(",") if x.strip()]
        cash = self.entry_cash.get().strip() or "373.00"

        if not tickers:
            return

        self.btn_audit.configure(state="disabled", text=i18n.t("btn_audit_loading"))
        self.audit_box.configure(state="normal")
        self.audit_box.delete("1.0", "end")
        self.audit_box.insert("1.0", i18n.t("audit_loading_msg", len_tickers=len(tickers)))
        self.audit_box.configure(state="disabled")

        threading.Thread(target=self._audit_worker, args=(tickers, cash), daemon=True).start()

    def _audit_worker(self, tickers, cash):
        risk_mode = self.seg_risk.get() if hasattr(self, "seg_risk") else i18n.t("seg_risk_conservative")

        audit_summary = []
        for t_audit in tickers:
            _, payload = self._fetch_stock_data(t_audit)
            if payload and "error" not in payload:
                audit_summary.append({
                    "ticker": t_audit,
                    "name": payload.get("name"),
                    "price": payload.get("price"),
                    "pe": payload.get("pe"),
                    "peg": payload.get("peg"),
                    "recommendation": payload.get("recommendation"),
                    "sector": payload.get("sector"),
                    "summary": payload.get("summary") # Add summary for richer context
                })

        api_key = self.get_gemini_key() if self.get_gemini_key else os.environ.get("GEMINI_API_KEY", "")
        reply = ""
        if api_key and HAS_GENAI and audit_summary:
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel("gemini-2.5-flash")

                # Get language instruction (NEW)
                current_lang_instruction = i18n.t("llm_default_lang_instruction") # LEO ADDITION: Used new key

                if i18n.t("seg_risk_venture_growth") in risk_mode: # Compare translated values
                    strategy_instruction = i18n.t("strategy_venture_desc")
                else:
                    strategy_instruction = i18n.t("strategy_conservative_desc")

                prompt = f"""Tu esi Komandants Kvarks un profesionāls finanšu auditors.
{current_lang_instruction}
Lietotājs ir izvēlējies šādu riska režīmu:
{strategy_instruction}

Lietotāja eToro pozīcijas un brīvie līdzekļi {cash} USD:
{json.dumps(audit_summary, ensure_ascii=False, indent=2)}

Sagatavo tiešu un strukturētu auditu atbilstoši izvēlētajai stratēģijai:
1. 🟢 {i18n.t("dividends_champions")}: kuri uzņēmumi atbilst šai stratēģijai.
2. ✂️ {i18n.t("audit_sell_liquidate")}: kuri aktīvi neatbilst šim riska profilam vai bremzē portfeli.
3. 🎯 {i18n.t("audit_cash_allocation")} ({cash} USD) SADALE: precīzs plāns, kur kur ieguldīt brīvos līdzekļus."""
                
                resp = model.generate_content(prompt)
                reply = resp.text.strip()
            except Exception as e:
                reply = i18n.t("error_audit_gemini", e=e)
        else:
            reply = i18n.t("error_api_missing_data")

        def _ui():
            self.btn_audit.configure(state="normal", text=i18n.t("btn_audit_portfolio"))
            self.audit_box.configure(state="normal")
            self.audit_box.delete("1.0", "end")
            self.audit_box.insert("1.0", reply)
            self.audit_box.configure(state="disabled")

        self.after(0, _ui)

    # ==========================================
    # 3. CILNE: KVARKA IZAUGSMES MEDNIEKS
    # ==========================================
    def _setup_growth_tab(self):
        f = self.tab_growth
        f.grid_columnconfigure(0, weight=1)
        f.grid_columnconfigure(1, weight=2)
        f.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(f, fg_color="#131b2e", corner_radius=10)
        left.grid(row=0, column=0, padx=(5, 5), pady=5, sticky="nsew")

        right = ctk.CTkFrame(f, fg_color="#131b2e", corner_radius=10)
        right.grid(row=0, column=1, padx=(5, 5), pady=5, sticky="nsew")

        self.lbl_growth_radar_title = ctk.CTkLabel(left, text=i18n.t("lbl_potential_radar"), font=ctk.CTkFont(size=14, weight="bold"), text_color="#38bdf8")
        self.lbl_growth_radar_title.pack(anchor="w", padx=15, pady=(15, 6))
        self.lbl_growth_description = ctk.CTkLabel(left, text=i18n.t("lbl_growth_description"),
                     font=ctk.CTkFont(size=11), text_color="#94a3b8")
        self.lbl_growth_description.pack(anchor="w", padx=15, pady=(0, 10))

        # Dynamic list of growth candidates (initially from GROWTH_WATCHLIST, then from AI)
        self.growth_candidates_container = ctk.CTkScrollableFrame(left, fg_color="#0b1329", height=240)
        self.growth_candidates_container.pack(fill="both", expand=True, padx=15, pady=(0, 10))
        self._populate_growth_watchlist(GROWTH_WATCHLIST) # Initial population

        # Progress bar and status label for Growth Hunter
        self.growth_status_label = ctk.CTkLabel(left, text=i18n.t("growth_status_initial"), font=ctk.CTkFont(size=11), text_color="#94a3b8")
        self.growth_status_label.pack(anchor="w", padx=15, pady=(5, 0))
        self.growth_progress_bar = ctk.CTkProgressBar(left, mode="determinate", height=10, corner_radius=5)
        self.growth_progress_bar.set(0)
        self.growth_progress_bar.pack(fill="x", padx=15, pady=(0, 10))
        self.growth_progress_bar.configure(determinate_speed=0.5) # Make it visually active even with fast updates
        self.growth_progress_bar.pack_forget() # Hide initially

        # RENAMED BUTTON
        self.btn_scan_all_growth = ctk.CTkButton(left, text=i18n.t("btn_hunt_unicorns"), height=34, fg_color="#d97706", hover_color="#b45309", font=ctk.CTkFont(size=12, weight="bold"), command=self.start_growth_hunter)
        self.btn_scan_all_growth.pack(fill="x", padx=15, pady=(0, 15))

        # Right pane: AI Growth Resolution
        # Assigned to self.lbl_growth_resolution_title for refresh_ui method (NEW for Instruction 2)
        self.lbl_growth_resolution_title = ctk.CTkLabel(right, text=i18n.t("lbl_inv_growth_resolution_title"), font=ctk.CTkFont(size=14, weight="bold"), text_color="#f59e0b")
        self.lbl_growth_resolution_title.pack(anchor="w", padx=15, pady=(15, 6))
        
        # Action buttons for right pane results
        growth_action_buttons_frame = ctk.CTkFrame(right, fg_color="transparent")
        growth_action_buttons_frame.pack(fill="x", padx=15, pady=(0, 6))

        self.btn_copy_growth = ctk.CTkButton(
            growth_action_buttons_frame,
            text=i18n.t("btn_copy_analysis"),
            width=130,
            height=28,
            fg_color="#334155",
            hover_color="#475569",
            font=ctk.CTkFont(size=11, weight="bold"),
            command=self.copy_growth_analysis
        )
        self.btn_copy_growth.pack(side="left", padx=(0, 8))

        self.btn_export_word = ctk.CTkButton(
            growth_action_buttons_frame,
            text=i18n.t("btn_save_word"),
            width=180,
            height=28,
            fg_color="#334155",
            hover_color="#475569",
            font=ctk.CTkFont(size=11, weight="bold"),
            command=self.export_growth_to_word
        )
        self.btn_export_word.pack(side="left")

        self.growth_results_container = ctk.CTkScrollableFrame(right, fg_color="#0b1329", corner_radius=8)
        self.growth_results_container.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
        initial_label = ctk.CTkLabel(self.growth_results_container, text=i18n.t("growth_initial_results"),
                                     font=ctk.CTkFont(size=12), text_color="#f8fafc", wraplength=400)
        initial_label.pack(fill="x", padx=10, pady=10)

        self.growth_current_candidates_data = [] # To store the AI generated candidates for copy/export

    def _populate_growth_watchlist(self, items):
        """Populates the left growth candidates container with buttons."""
        for widget in self.growth_candidates_container.winfo_children():
            widget.destroy()

        if not items:
            ctk.CTkLabel(self.growth_candidates_container, text=i18n.t("growth_no_candidates"), font=ctk.CTkFont(size=11), text_color="#94a3b8").pack(fill="x", padx=10, pady=5)
            return

        for item in items:
            card = ctk.CTkFrame(self.growth_candidates_container, fg_color="#131b2e", corner_radius=6)
            card.pack(fill="x", pady=4, padx=2)
            ctk.CTkLabel(card, text=item.get("ticker", "N/A"), font=ctk.CTkFont(size=12, weight="bold"), text_color="#38bdf8").pack(side="left", padx=8, pady=4)
            ctk.CTkLabel(card, text=item.get("name", "N/A"), font=ctk.CTkFont(size=10), text_color="#cbd5e1").pack(side="left", padx=4)
            ctk.CTkButton(card, text=i18n.t("btn_check"), width=65, height=22, fg_color="#0284c7", font=ctk.CTkFont(size=10), command=lambda t_ticker_check=item.get("ticker"): self._inspect_growth(t_ticker_check)).pack(side="right", padx=6)


    def _inspect_growth(self, ticker):
        self.tabview.set(i18n.t("tab_inv_single")) # Set tab using translated name
        self._quick_set(ticker)

    def start_growth_hunter(self):
        self.btn_scan_all_growth.configure(state="disabled", text=i18n.t("btn_hunt_unicorns_loading", btn_text=i18n.t("btn_hunt_unicorns")))
        
        # Clear previous results and show loading message
        for widget in self.growth_results_container.winfo_children():
            widget.destroy()
        
        # Clear left candidate list for new ones
        for widget in self.growth_candidates_container.winfo_children():
            widget.destroy()
        ctk.CTkLabel(self.growth_candidates_container, text=i18n.t("growth_scanning_candidates"),
                     font=ctk.CTkFont(size=11), text_color="#94a3b8").pack(fill="x", padx=10, pady=5)


        self.growth_progress_bar.set(0)
        self.growth_progress_bar.pack(fill="x", padx=15, pady=(0, 10))
        self.growth_status_label.configure(text=i18n.t("growth_scan_progress1"))

        threading.Thread(target=self._growth_worker, daemon=True).start()

    def _growth_worker(self):
        self.after(0, lambda: self.growth_progress_bar.set(0.1))

        growth_data = []
        for g in GROWTH_WATCHLIST:
            _, payload = self._fetch_stock_data(g["ticker"])
            if payload and "error" not in payload:
                growth_data.append(payload)
            self.after(0, lambda: self.growth_progress_bar.set(self.growth_progress_bar.get() + 0.8 / len(GROWTH_WATCHLIST)))
        
        self.after(0, lambda: self.growth_status_label.configure(text=i18n.t("growth_scan_progress2")))
        self.after(0, lambda: self.growth_progress_bar.set(0.9))

        api_key = self.get_gemini_key() if self.get_gemini_key else os.environ.get("GEMINI_API_KEY", "")
        
        unicorn_candidates = []
        sector_concentration_warning = ""
        reply_fallback = i18n.t("growth_no_unicorns_fallback")

        if api_key and HAS_GENAI and growth_data:
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel("gemini-2.5-flash")

                # Get language instruction (NEW)
                current_lang_instruction = i18n.t("llm_default_lang_instruction")

                # Kvark's criteria (localizable if needed, but treated as fixed for now)
                kvark_criteria_prompt = """
                Komandanta Kvarka 4 zelta kritēriji Jauno Vienradžu medībām:
                1.  **Revenue Growth > 30-50%**: Spēcīga ieņēmumu izaugsme (vismaz virs 30%, ideāli virs 50%). Novērtēt pēc "summary" vai vispārējās tirgus izpratnes.
                2.  **Milzīgs TAM (Total Addressable Market)**: Uzņēmumam jādarbojas tirgū ar gigantisku nākotnes potenciālu (piemēram, AI, kvantu skaitļošana, kosmosa ekonomika, nākamās paaudzes kiberdrošība). Novērtēt pēc "summary" un sektora.
                3.  **Moat (Konkurences grāvis)**: Spēcīgas konkurences priekšrocības (patentēta tehnoloģija, tīkla efekti, zīmols, augstas maiņas izmaksas klientiem). Novērtēt pēc "summary".
                4.  **Founder-led (Dibinātāja vadīts)**: Vēlams, ka uzņēmumu joprojām vada tā dibinātājs vai spēcīga, ilgtermiņā domājoša vadošā komanda ar skaidru vīziju. Novērtēt pēc "summary" (var būt netieši).
                """ # These are fixed instructions in the prompt, not UI text.

                prompt = f"""Tu esi agresīvs izaugsmes investīciju eksperts un Komandants Kvarks.
{current_lang_instruction}
Tavs uzdevums ir atrast tieši 10 "Jaunos Vienradžus" - uzņēmumus ar 10x-100x izaugsmes potenciālu.
Meklē jaunākos, spēcīgākos un atbilstošākos pretendentus, kas nav vienkārši esošie tirgus līderi, bet gan nākamie revolucionāri.
OBLIGĀTI iekļauj dažādas nākotnes nozares, lai nodrošinātu portfeļa diversifikāciju un novērstu pārmērīgu koncentrāciju tikai uz AI/čipiem.

Stratēģiski svarīgās nākotnes nozares, kurās meklēt vienradžus:
- Mākslīgais intelekts (AI) & Čipi (semiconductors)
- Kiberdrošība & Datu infrastruktūra
- Biotehnoloģijas & Genomika / Modernā medicīna (piem., CRSP, VRTX, RXRX, BIIIB, ISRG u.tml.)
- Nākotnes enerģija & jaunie materiāli (piem., tīrā enerģija, kodolsintēze, ūdeņraža tehnoloģijas, progresīvas bateriju tehnoloģijas, retzemju metālu alternatīvas)
- Kosmoss un autonomās sistēmas (piem., satelīti, raķešu tehnoloģijas, droni, robotika, kosmosa tūrisms, resursu ieguve kosmosā)

Lūk, dati par strauji augošiem uzņēmumiem nākotnes tehnoloģijās (tie nav obligāti gala kandidāti, bet izejas informācija analīzei), katram ir pievienots arī kopsavilkums (summary):
{json.dumps(growth_data, ensure_ascii=False, indent=2)}

{kvark_criteria_prompt}

Veic analīzi un izvadi rezultātu stingri JSON formātā.
JSON atbildē ir jābūt diviem laukiem:
1. `unicorn_candidates`: Saraksts (array) ar dict (objektiem). Katram objektam jāsatur šādi lauki:
    - `{{ticker}}`: Uzņēmuma tikers.
    - `{{name}}`: Uzņēmuma nosaukums.
    - `{{niche}}`: "Slepenais ierocis" (īss 1-2 teikumu nišas apraksts, kas dod konkurences priekšrocības). Šī lauka vērtībai OBLIGĀTI jābūt ģenerētai atbilstoši valodai: "{current_lang_instruction}".
    - `{{why_10x_100x}}`: "Kāpēc 10x-100x" (ass 2 teikumu pamatojums par milzīgo izaugsmes potenciālu, balstoties uz Kvarka kritērijiem un nozares perspektīvām). Šī lauka vērtībai OBLIGĀTI jābūt ģenerētai atbilstoši valodai: "{current_lang_instruction}".
    - `{{black_swan_risk}}`: "Melnais Gulbis" (objektīvs 1-2 teikumu riska faktors, kas var apturēt izaugsmi). Šī lauka vērtībai OBLIGĀTI jābūt ģenerētai atbilstoši valodai: "{current_lang_instruction}".
    UZMANĪBU: Pārliecinies, ka sarakstā ir tieši 10 unikāli kandidāti, kas atbilst dažādām nozarēm, ja iespējams.
2. `sector_concentration_warning`: "Sarkano Karogu" brīdinājums par pārmērīgu sektora koncentrāciju, ja Kvarks to pamana (piemēram, vairāk nekā 50% kandidātu ir no vienas nozares), vai tukša stringa, ja brīdinājuma nav.

Tikai un vienīgi JSON atbilde. Nekādus citus ievadvārdus vai paskaidrojumus."""
                
                resp = model.generate_content(prompt)
                json_response_raw_text = resp.text.strip()
                
                # CRITICAL FIX: Clean the response from markdown blocks (Task 1)
                json_cleaned_from_md = json_response_raw_text
                if json_cleaned_from_md.startswith("```json"):
                    json_cleaned_from_md = json_cleaned_from_md[len("```json"):].strip()
                if json_cleaned_from_md.endswith("```"):
                    json_cleaned_from_md = json_cleaned_from_md[:-len("```")].strip()

                # NEW: Use regex to find the JSON block only (Task 1)
                json_match = re.search(r'\{.*\}', json_cleaned_from_md, re.DOTALL)
                json_to_parse = json_match.group(0) if json_match else ""

                try:
                    if json_to_parse: # Only attempt to load if a JSON block was found
                        parsed_response = json.loads(json_to_parse)
                        unicorn_candidates = parsed_response.get("unicorn_candidates", [])
                        sector_concentration_warning = parsed_response.get("sector_concentration_warning", "")
                        self.last_sector_concentration_warning = sector_concentration_warning # Store for copy/export (FIXED for Point 1, 3)
                        reply_fallback = "" # If successful, clear fallback
                    else:
                        raise json.JSONDecodeError("No JSON block found in AI response after markdown cleanup.", json_response_raw_text, 0)
                except json.JSONDecodeError as jde: # This except belongs to the nested try block
                    reply_fallback = i18n.t("error_gemini_parsing", e=jde, partial_response=json_response_raw_text[:500])
                    unicorn_candidates = []
                    sector_concentration_warning = i18n.t("error_ai_unparsable")
                    self.last_sector_concentration_warning = sector_concentration_warning # Store error as warning

            except Exception as e:
                reply_fallback = i18n.t("error_gemini_api_contact", e=e)
                sector_concentration_warning = i18n.t("error_gemini_api_contact", e=e)
                self.last_sector_concentration_warning = sector_concentration_warning # Store error as warning
        else:
            reply_fallback = i18n.t("error_api_missing_data")
            sector_concentration_warning = i18n.t("error_api_missing_data")
            self.last_sector_concentration_warning = sector_concentration_warning # Store error as warning

        self.growth_current_candidates_data = unicorn_candidates # Store for copy/export

        def _ui():
            self.btn_scan_all_growth.configure(state="normal", text=i18n.t("btn_hunt_unicorns"))
            self.growth_progress_bar.set(1.0)
            self.growth_progress_bar.pack_forget() # Hide progress bar
            self.growth_status_label.configure(text=i18n.t("growth_analysis_complete"))
            self.after(3000, lambda: self.growth_status_label.configure(text=i18n.t("growth_status_initial")))

            # Clear all widgets in the right container
            for widget in self.growth_results_container.winfo_children():
                widget.destroy()

            # Populate left pane with AI candidates (Task 3: Dynamic list)
            self._populate_growth_watchlist(unicorn_candidates) # Uses AI-selected candidates for left pane

            if self.last_sector_concentration_warning: # Check the instance variable (FIXED for Point 1, 3)
                warning_label = ctk.CTkLabel(self.growth_results_container, text=i18n.t("warning_red_flag", warning_message=self.last_sector_concentration_warning),
                                             font=ctk.CTkFont(size=12, weight="bold"), text_color="#dc2626", wraplength=400)
                warning_label.pack(fill="x", padx=10, pady=5)
            
            if unicorn_candidates:
                for idx, candidate in enumerate(unicorn_candidates):
                    card = ctk.CTkFrame(self.growth_results_container, fg_color="#1e293b", corner_radius=10, border_width=2, border_color="#334155")
                    card.pack(fill="x", padx=10, pady=8)

                    ctk.CTkLabel(card, text=f"{idx+1}. {candidate.get('ticker', 'N/A')} - {candidate.get('name', 'N/A')}",
                                 font=ctk.CTkFont(size=14, weight="bold"), text_color="#38bdf8").pack(anchor="w", padx=15, pady=(10, 2))
                    
                    ctk.CTkLabel(card, text=i18n.t("lbl_niche"),
                                 font=ctk.CTkFont(size=11, weight="bold"), text_color="#94a3b8").pack(anchor="w", padx=15, pady=(5, 0))
                    ctk.CTkLabel(card, text=candidate.get('niche', 'N/A'),
                                 font=ctk.CTkFont(size=11), text_color="#cbd5e1", wraplength=450, justify="left").pack(anchor="w", padx=15)
                    
                    ctk.CTkLabel(card, text=i18n.t("lbl_why_10x_100x"),
                                 font=ctk.CTkFont(size=11, weight="bold"), text_color="#94a3b8").pack(anchor="w", padx=15, pady=(5, 0))
                    ctk.CTkLabel(card, text=candidate.get('why_10x_100x', 'N/A'),
                                 font=ctk.CTkFont(size=11), text_color="#cbd5e1", wraplength=450, justify="left").pack(anchor="w", padx=15)
                    
                    ctk.CTkLabel(card, text=i18n.t("lbl_black_swan"),
                                 font=ctk.CTkFont(size=11, weight="bold"), text_color="#94a3b8").pack(anchor="w", padx=15, pady=(5, 0))
                    ctk.CTkLabel(card, text=candidate.get('black_swan_risk', 'N/A'),
                                 font=ctk.CTkFont(size=11), text_color="#cbd5e1", wraplength=450, justify="left").pack(anchor="w", padx=15)
                    
                    add_button = ctk.CTkButton(card, text=i18n.t("btn_add_to_portfolio"),
                                               width=220, height=30, fg_color="#059669", hover_color="#10b981",
                                               font=ctk.CTkFont(size=12, weight="bold"),
                                               command=lambda t_add=candidate.get('ticker'): self._add_ticker_to_portfolio(t_add))
                    add_button.pack(anchor="w", padx=15, pady=(10, 10))
            else:
                fallback_label = ctk.CTkLabel(self.growth_results_container, text=reply_fallback,
                                              font=ctk.CTkFont(size=12), text_color="#f8fafc", wraplength=400)
                fallback_label.pack(fill="x", padx=10, pady=10)

        self.after(0, _ui)

    def copy_growth_analysis(self):
        text_to_copy_parts = []
        if hasattr(self, 'growth_current_candidates_data') and self.growth_current_candidates_data:
            text_to_copy_parts.append(i18n.t("growth_copy_header") + "\n")
            
            # Use the stored instance variable for warning (FIXED for Point 1, 3)
            if hasattr(self, 'last_sector_concentration_warning') and self.last_sector_concentration_warning:
                 text_to_copy_parts.append(i18n.t("warning_red_flag", warning_message=self.last_sector_concentration_warning) + "\n")

            for idx, candidate in enumerate(self.growth_current_candidates_data):
                text_to_copy_parts.append(i18n.t("growth_copy_candidate_header", idx_plus_1=idx+1, ticker=candidate.get('ticker', 'N/A'), name=candidate.get('name', 'N/A')) + "\n")
                text_to_to_copy_parts.append(i18n.t("growth_copy_niche", niche=candidate.get('niche', 'N/A')) + "\n")
                text_to_copy_parts.append(i18n.t("growth_copy_why", why=candidate.get('why_10x_100x', 'N/A')) + "\n")
                text_to_copy_parts.append(i18n.t("growth_copy_risk", risk=candidate.get('black_swan_risk', 'N/A')) + "\n")
                text_to_copy_parts.append("\n")
            text_to_copy_parts.append("----------------------------------------\n")
            text_to_copy_parts.append(i18n.t("audit_copy_footer"))
        else:
            if self.growth_results_container.winfo_children():
                # Attempt to get text from the first child if no structured data
                first_child_text = self.growth_results_container.winfo_children()[0].cget("text")
                if i18n.t("growth_initial_results") not in first_child_text and i18n.t("growth_no_analysis_to_copy") not in first_child_text:
                    text_to_copy_parts.append(first_child_text)
                else:
                    text_to_copy_parts.append(i18n.t("growth_no_analysis_to_copy"))
            else:
                text_to_copy_parts.append(i18n.t("growth_no_analysis_to_copy"))


        text_to_copy = "\n".join(text_to_copy_parts).strip()

        if not text_to_copy or i18n.t("growth_initial_results") in text_to_copy or i18n.t("growth_no_analysis_to_copy") in text_to_copy:
            return

        try:
            self.clipboard_clear()
            self.clipboard_append(text_to_copy)
            self.update()

            orig_text = self.btn_copy_growth.cget("text")
            orig_color = self.btn_copy_growth.cget("fg_color")
            self.btn_copy_growth.configure(text=i18n.t("status_copied"), fg_color="#10b981")

            self.after(2000, lambda: self.btn_copy_growth.configure(text=orig_text, fg_color=orig_color))
        except Exception as e:
            print(i18n.t("error_copying", e=e))

    def export_growth_to_word(self):
        if not HAS_DOCX:
            messagebox.showwarning(i18n.t("error_title"), i18n.t("error_docx_missing"))
            return

        if not hasattr(self, 'growth_current_candidates_data') or not self.growth_current_candidates_data:
            messagebox.showinfo(i18n.t("info_title"), i18n.t("info_no_data_to_export"))
            return
        
        try:
            document = Document()
            document.add_heading(i18n.t("word_report_title"), level=1)
            document.add_paragraph(f"{i18n.t('word_generated_on')}{datetime.now().strftime('%d.%m.%Y %H:%M')}")
            document.add_paragraph(i18n.t("word_analysis_basis"))

            # Use the stored instance variable for warning (FIXED for Point 1, 3)
            if hasattr(self, 'last_sector_concentration_warning') and self.last_sector_concentration_warning:
                 p_warning = document.add_paragraph()
                 p_warning.add_run(i18n.t("word_red_flag")).bold = True
                 p_warning.add_run(self.last_sector_concentration_warning)


            for idx, candidate in enumerate(self.growth_current_candidates_data):
                document.add_heading(f"{idx+1}. {candidate.get('name', 'N/A')} ({candidate.get('ticker', 'N/A')})", level=2)
                
                p_niche = document.add_paragraph()
                p_niche.add_run(i18n.t("word_secret_weapon")).bold = True
                p_niche.add_run(candidate.get('niche', 'N/A'))

                p_why = document.add_paragraph()
                p_why.add_run(i18n.t("word_why_10x_100x")).bold = True
                p_why.add_run(candidate.get('why_10x_100x', 'N/A'))
                
                p_risk = document.add_paragraph()
                p_risk.add_run(i18n.t("word_black_swan")).bold = True
                p_risk.add_run(candidate.get('black_swan_risk', 'N/A'))
                
                document.add_paragraph("") # Add a blank line for spacing.

            output_filename = f"{i18n.t('word_output_filename_prefix')}{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            output_filepath = OUTBOX_DIR / output_filename
            document.save(output_filepath)
            
            success_message_content = i18n.t('word_success_message') + str(output_filepath)
            messagebox.showinfo(i18n.t("success_title"), success_message_content)
        except Exception as e:
            messagebox.showerror(i18n.t("error_saving_word_title"), f"{i18n.t('word_error_saving_message')}{e}")

    # ==========================================
    # 4. CILNE: DIVIDENDES & NAUDAS PLŪSMA
    # ==========================================
    def _setup_dividends_tab(self):
        f = self.tab_dividends
        f.grid_columnconfigure(0, weight=1)
        f.grid_columnconfigure(1, weight=2)
        f.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(f, fg_color="#131b2e", corner_radius=10)
        left.grid(row=0, column=0, padx=(5, 5), pady=5, sticky="nsew")

        right = ctk.CTkFrame(f, fg_color="#131b2e", corner_radius=10)
        right.grid(row=0, column=1, padx=(5, 5), pady=5, sticky="nsew")

        self.lbl_cashflow_radar_title = ctk.CTkLabel(
            left,
            text=i18n.t("lbl_cashflow_radar_title"),
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#38bdf8"
        )
        self.lbl_cashflow_radar_title.pack(anchor="w", padx=15, pady=(15, 4))

        self.lbl_cashflow_radar_desc = ctk.CTkLabel(
            left,
            text=i18n.t("lbl_cashflow_radar_desc"),
            font=ctk.CTkFont(size=11),
            text_color="#94a3b8",
            wraplength=280,
            justify="left"
        )
        self.lbl_cashflow_radar_desc.pack(anchor="w", padx=15, pady=(0, 12))

        self.btn_calc_dividends = ctk.CTkButton(
            left,
            text=i18n.t("btn_calc_dividends"),
            height=34,
            fg_color="#0284c7",
            hover_color="#0369a1",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.start_dividend_scan
        )
        self.btn_calc_dividends.pack(fill="x", padx=15, pady=(0, 15))

        self.div_summary_box = ctk.CTkTextbox(
            left,
            fg_color="#0b1329",
            font=ctk.CTkFont(family="Consolas", size=11),
            corner_radius=8
        )
        self.div_summary_box.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.div_summary_box.insert("1.0", i18n.t("dividends_initial_msg"))
        self.div_summary_box.configure(state="disabled")

        # Assigned to self.lbl_dividends_resolution_title for refresh_ui method (NEW for Instruction 2)
        self.lbl_dividends_resolution_title = ctk.CTkLabel(
            right,
            text=i18n.t("lbl_inv_dividends_resolution_title"),
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#f59e0b"
        )
        self.lbl_dividends_resolution_title.pack(anchor="w", padx=15, pady=(15, 6))

        self.div_analysis_box = ctk.CTkTextbox(
            right,
            fg_color="#0b1329",
            text_color="#f8fafc",
            font=ctk.CTkFont(size=12),
            wrap="word",
            corner_radius=8
        )
        self.div_analysis_box.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.div_analysis_box.insert("1.0", i18n.t("dividends_initial_analysis_msg"))
        self.div_analysis_box.configure(state="disabled")

    def start_dividend_scan(self):
        raw = self.txt_my_portfolio.get("1.0", "end").strip()
        tickers = [x.strip().upper() for x in raw.replace("\n", ",").split(",") if x.strip()]
        if not tickers:
            return

        self.btn_calc_dividends.configure(state="disabled", text=i18n.t("btn_calc_dividends_loading"))
        self.div_summary_box.configure(state="normal")
        self.div_summary_box.delete("1.0", "end")
        self.div_summary_box.insert("1.0", i18n.t("dividends_scan_progress1", len_tickers=len(tickers)))
        self.div_summary_box.configure(state="disabled")

        self.div_analysis_box.configure(state="normal")
        self.div_analysis_box.delete("1.0", "end")
        self.div_analysis_box.insert("1.0", i18n.t("dividends_scan_progress2"))
        self.div_analysis_box.configure(state="disabled")

        threading.Thread(target=self._dividend_worker, args=(tickers,), daemon=True).start()

    def _dividend_worker(self, tickers):
        div_data = []
        summary_lines = [i18n.t("dividends_summary_header")]

        for t_div in tickers:
            try:
                stock = yf.Ticker(t_div)
                info = stock.info or {}
                div_yield = info.get("dividendYield")
                div_rate = info.get("dividendRate", 0.0)
                ex_div = info.get("exDividendDate")

                # 1. Droša datuma apstrāde un pārvēršana uz DD.MM.YYYY
                ex_date_str = i18n.t("dividends_date_na")
                if ex_div:
                    try:
                        # If a number (Unix timestamp in seconds or milliseconds)
                        if isinstance(ex_div, (int, float)):
                            ts = ex_div / 1000 if ex_div > 1e11 else ex_div
                            ex_date_str = datetime.fromtimestamp(ts).strftime("%d.%m.%Y")
                        else:
                            # Attempt to parse as date string, take first part if time is included
                            parsed_date = datetime.strptime(str(ex_div).split("T")[0].split(" ")[0], "%Y-%m-%d")
                            ex_date_str = parsed_date.strftime("%d.%m.%Y")
                    except Exception:
                        ex_date_str = str(ex_div) # Fallback to original string

                # 2. Dividenžu ienesīguma un korekciju apstrāde
                if div_yield and div_yield > 0:
                    # If it comes as 0.015 (1.5%) or already as 1.5
                    yield_val = div_yield * 100 if div_yield < 1 else div_yield
                    yield_pct = round(yield_val, 2)

                    # Kļūdu filtrs (piem., AVGO pēc split rāda 73%)
                    if yield_pct > 35 and t_div.upper() == "AVGO": # Specific correction for AVGO, based on domain knowledge
                        yield_pct = 1.35 # Corrected yield after split for AVGO

                    div_data.append({
                        "ticker": t_div,
                        "name": info.get("shortName", t_div),
                        "yield_pct": f"{yield_pct}%",
                        "rate": div_rate,
                        "ex_date": ex_date_str
                    })
                    summary_lines.append(f"{t_div:6} | {yield_pct:5.2f}% | Izmaksa: {div_rate} | Dat: {ex_date_str}")
                else:
                    summary_lines.append(f"{t_div:6} | 0.00%  | {i18n.t('dividends_not_paying')}")
            except Exception as e:
                summary_lines.append(f"{t_div:6} | {i18n.t('dividends_error_fetching', e=e)}")

        # Gemini rezolūcija
        api_key = self.get_gemini_key() if self.get_gemini_key else os.environ.get("GEMINI_API_KEY", "")
        reply = ""
        if api_key and HAS_GENAI and div_data:
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel("gemini-2.5-flash")
                # Get language instruction (NEW)
                current_lang_instruction = i18n.t("llm_default_lang_instruction")

                prompt = f"""Tu esi dividenžu un pasīvās naudas plūsmas eksperts Komandants Kvarks.
{current_lang_instruction}
Šeit ir lietotāja portfeļa dividenču dati:
{json.dumps(div_data, ensure_ascii=False, indent=2)}

Sagatavo strukturētu analīzi:
1. {i18n.t('dividends_champions')}: Kuri no šiem uzņēmumiem nodrošina stabilāko un drošāko dividenču atdevi.
2. {i18n.t('dividends_sustainability_risks')}: Vai kādam uzņēmumam dividende nav mākslīgi uzpūsta vai apdraudēta (piemēram, naftas vai cikliskajos sektoros).
3. {i18n.t('dividends_reinvestment_strategy')}: Kā visefektīvāk izmantot ienākošās dividendes, lai portfelis augtu eksponenciāli."""
                resp = model.generate_content(prompt)
                reply = resp.text.strip()
            except Exception as e:
                reply = i18n.t("error_gemini_dividends", e=e)
        elif not div_data:
            reply = i18n.t("dividends_no_companies")
        else:
            reply = i18n.t("error_api_key_unavailable")

        def _ui():
            self.btn_calc_dividends.configure(state="normal", text=i18n.t("btn_calc_dividends"))
            self.div_summary_box.configure(state="normal")
            self.div_summary_box.delete("1.0", "end")
            self.div_summary_box.insert("1.0", "\n".join(summary_lines))
            self.div_summary_box.configure(state="disabled")

            self.div_analysis_box.configure(state="normal")
            self.div_analysis_box.delete("1.0", "end")
            self.div_analysis_box.insert("1.0", reply)
            self.div_analysis_box.configure(state="disabled")

        self.after(0, _ui)

    # ==========================================
    # 5. CILNE: TIRGUS KLIMATS & INSAIDERI
    # ==========================================
    def _setup_market_tab(self):
        f = self.tab_market
        f.grid_columnconfigure(0, weight=1)
        f.grid_columnconfigure(1, weight=2)
        f.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(f, fg_color="#131b2e", corner_radius=10)
        left.grid(row=0, column=0, padx=(5, 5), pady=5, sticky="nsew")

        right = ctk.CTkFrame(f, fg_color="#131b2e", corner_radius=10)
        right.grid(row=0, column=1, padx=(5, 5), pady=5, sticky="nsew")

        self.lbl_market_sentiment_title = ctk.CTkLabel(left, text=i18n.t("lbl_market_sentiment_title"), font=ctk.CTkFont(size=14, weight="bold"), text_color="#38bdf8")
        self.lbl_market_sentiment_title.pack(anchor="w", padx=15, pady=(15, 6))

        self.lbl_fng = ctk.CTkLabel(left, text=i18n.t("lbl_fng_loading"), font=ctk.CTkFont(size=13, weight="bold"), text_color="#f59e0b")
        self.lbl_fng.pack(anchor="w", padx=15, pady=(0, 10))

        self.btn_refresh_market = ctk.CTkButton(
            left,
            text=i18n.t("btn_refresh_market_intel"),
            height=34,
            fg_color="#0284c2",
            hover_color="#0369a1",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.start_market_intel_scan
        )
        self.btn_refresh_market.pack(fill="x", padx=15, pady=(0, 15))

        self.insider_box = ctk.CTkTextbox(left, fg_color="#0b1329", font=ctk.CTkFont(family="Consolas", size=11), corner_radius=8)
        self.insider_box.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.insider_box.insert("1.0", i18n.t("insider_initial_msg"))
        self.insider_box.configure(state="disabled")

        # Assigned to self.lbl_market_news_title for refresh_ui method (NEW for Instruction 2)
        self.lbl_market_news_title = ctk.CTkLabel(right, text=i18n.t("lbl_inv_market_news_title"), font=ctk.CTkFont(size=14, weight="bold"), text_color="#f59e0b")
        self.lbl_market_news_title.pack(anchor="w", padx=15, pady=(15, 6))
        
        self.market_news_box = ctk.CTkTextbox(right, fg_color="#0b1329", text_color="#f8fafc", font=ctk.CTkFont(size=12), wrap="word", corner_radius=8)
        self.market_news_box.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.market_news_box.insert("1.0", i18n.t("market_news_initial_msg"))
        self.market_news_box.configure(state="disabled")

    def start_market_intel_scan(self):
        # LEO ADDITION: Requests already imported at the top
        self.btn_refresh_market.configure(state="disabled", text=i18n.t("btn_refresh_market_intel_loading"))
        
        def _worker():
            # 1. Fear & Greed indekss
            fng_val = "50"
            fng_status = "Neutral"
            fng_text = i18n.t("fng_not_available")
            try:
                res = requests.get("https://api.alternative.me/fng/?limit=1", timeout=5).json()
                item = res["data"][0]
                fng_val = item['value']
                fng_status = item['value_classification']
                fng_text = f"Fear & Greed: {fng_val}/100 ({fng_status})"
            except Exception:
                pass

            # 2. Makro & Likviditātes rādītāji (yfinance)
            macro_summary = []
            macro_lines = [i18n.t("macro_liquidity_header")]
            try:
                macro_tickers = {"^TNX": i18n.t("macro_us10y"), "GC=F": i18n.t("macro_gold"), "BTC-USD": i18n.t("macro_bitcoin")}
                for sym, label in macro_tickers.items():
                    try:
                        tk = yf.Ticker(sym)
                        p = tk.fast_info.get("lastPrice") or tk.info.get("regularMarketPrice") or 0.0
                        if sym == "^TNX":
                            macro_lines.append(f"• {label:14}: {p:.2f}%")
                            macro_summary.append(f"{label}: {p:.2f}%")
                        else:
                            curr = "USD"
                            macro_lines.append(f"• {label:14}: {p:,.2f} {curr}")
                            macro_summary.append(f"{label}: {p:,.2f} {curr}")
                    except Exception:
                        macro_lines.append(f"• {label:14}: {i18n.t('macro_na')}")
            except Exception:
                pass

            # 3. Portfeļa tikeri
            raw = self.txt_my_portfolio.get("1.0", "end").strip()
            tickers = [x.strip().upper() for x in raw.replace("\n", ",").split(",") if x.strip() and "." not in x][:5]
            if not tickers:
                tickers = ["MSFT", "ARM", "VOO"]
            
            # 4. Finnhub insaideri un ziņas
            cfg_file = BASE_DIR / "config.json"
            finnhub_key = ""
            if cfg_file.exists():
                try:
                    with open(cfg_file, "r", encoding="utf-8") as f:
                        cfg = json.load(f)
                        finnhub_key = cfg.get("finnhub_api_key", "")
                except Exception:
                    pass

            insider_lines = [i18n.t("insider_transactions_header")]
            top_headlines = []

            if finnhub_key:
                for t_insider in tickers:
                    try:
                        u = f"https://finnhub.io/api/v1/stock/insider-sentiment?symbol={t_insider}&token={finnhub_key}"
                        r = requests.get(u, timeout=5).json()
                        d = r.get("data", [])
                        if d:
                            lat = d[-1]
                            m = lat.get("mspr", 0)
                            sig = i18n.t("insider_buy") if m > 0 else i18n.t("insider_sell") if m < 0 else i18n.t("insider_neutral")
                            insider_lines.append(f"{t_insider:5} | {sig:8} | MSPR: {m:6.2f} | Net: {lat.get('change', 0)}")
                        else:
                            insider_lines.append(f"{t_insider:5} | {i18n.t('insider_no_major_deals')}")
                    except Exception:
                        insider_lines.append(f"{t_insider:5} | {i18n.t('insider_data_error')}")

                d_to = datetime.now().strftime("%Y-%m-%d")
                d_from = (datetime.now() - timedelta(days=14)).strftime("%Y-%m-%d")
                for t_news in tickers[:2]:
                    try:
                        nu = f"https://finnhub.io/api/v1/company-news?symbol={t_news}&from={d_from}&to={d_to}&token={finnhub_key}"
                        nlist = requests.get(nu, timeout=5).json()
                        if isinstance(nlist, list):
                            for n in nlist[:2]:
                                top_headlines.append(f"[{t_news}] {n.get('headline', '')}")
                    except Exception:
                        pass
            else:
                insider_lines.append(i18n.t("error_finnhub_key_missing"))

            # 5. Apvienojam kreisās puses datus (Makro + Insaideri)
            left_display_text = "\n".join(macro_lines) + "\n" + "\n".join(insider_lines)

            # 6. Komandanta Kvarka (Gemini) stratēģiskais slēdziens
            news_output = []
            if top_headlines:
                news_output.append(i18n.t("news_headlines_header"))
                for h in top_headlines:
                    news_output.append(f"• {h}")
                news_output.append("\n" + "="*30 + "\n")

            api_key = self.get_gemini_key() if self.get_gemini_key else os.environ.get("GEMINI_API_KEY", "")
            if api_key and HAS_GENAI:
                try:
                    genai.configure(api_key=api_key)
                    model = genai.GenerativeModel("gemini-2.5-flash")
                    # Get language instruction (NEW)
                    current_lang_instruction = i18n.t("llm_default_lang_instruction")

                    prompt = f"""Tu esi militāri ass investīciju virspavēlnieks Komandants Kvarks (resnais runcis no Liepājas štāba).
{current_lang_instruction}
Makro un likviditātes rādītāji: {', '.join(macro_summary)}
Tirgus baiļu/alkatības noskaņojums: {fng_text}
Insaideru signāli: {'; '.join(insider_lines[1:])}
Jaunākās ziņas: {'; '.join(top_headlines) if top_headlines else i18n.t('news_no_aggressive')}

Dod 3-5 teikumu zibenīgu operatīvo atskaiti un taktisko rīkojumu:
1. Makro fons: ko saka 10 gadu likmes, zelts un BTC par naudas plūsmu tirgū.
2. Insaideru un sentimenta vērtējums.
3. Galīgā kaujas pavēle šodienai: uzbrukt (pirkt), nocietināties (turēt) vai atkāpties skaidrā naudā."""
                    res = model.generate_content(prompt)
                    news_output.append(i18n.t("kvark_strategic_resolution"))
                    news_output.append(res.text.strip())
                except Exception as e:
                    news_output.append(i18n.t("error_kvark_market_intel", e=e))
            else:
                news_output.append(i18n.t("error_api_key_unavailable"))

            def _ui():
                self.btn_refresh_market.configure(state="normal", text=i18n.t("btn_refresh_market_intel"))
                self.lbl_fng.configure(text=fng_text)
                
                self.insider_box.configure(state="normal")
                self.insider_box.delete("1.0", "end")
                self.insider_box.insert("1.0", left_display_text)
                self.insider_box.configure(state="disabled")

                self.market_news_box.configure(state="normal")
                self.market_news_box.delete("1.0", "end")
                self.market_news_box.insert("1.0", "\n".join(news_output))
                self.market_news_box.configure(state="disabled")

            self.after(0, _ui)

        threading.Thread(target=_worker, daemon=True).start()

    # ==========================================
    # DATU VĀKŠANA UN GEMINI METODES
    # ==========================================
    def _get_api_keys(self):
        cfg_file = BASE_DIR / "config.json"
        api_keys = {
            "alpha_vantage_api_key": "",
            "fmp_api_key": "",
            "finnhub_api_key": ""
        }
        if cfg_file.exists():
            try:
                with open(cfg_file, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                    api_keys["alpha_vantage_api_key"] = cfg.get("alpha_vantage_api_key", "")
                    api_keys["fmp_api_key"] = cfg.get("fmp_api_key", "")
                    api_keys["finnhub_api_key"] = cfg.get("finnhub_api_key", "")
            except Exception as e:
                print(f"ERROR: Could not read API keys from config.json: {e}")
        return api_keys

    def _format_stats_text(self, ticker_val, data_dict):
        # This helper formats the standard stats block, using data from any source
        short_name = data_dict.get("name", ticker_val)
        currency = data_dict.get("currency", "USD")
        curr_price = data_dict.get("price", "N/A")
        pe = data_dict.get("pe", "N/A")
        fwd_pe = data_dict.get("forward_pe", "N/A")
        peg = data_dict.get("peg", "N/A")
        high52 = data_dict.get("52w_high", "N/A")
        low52 = data_dict.get("52w_low", "N/A")
        rec = data_dict.get("recommendation", "N/A").upper()
        target_mean = data_dict.get("target_price", "N/A")
        div_str = data_dict.get("dividend_yield", "0.0%")
        source_name = data_dict.get("source", "N/A")

        # Ensure numeric values are formatted nicely if they exist
        try: curr_price = f"{float(curr_price):,.2f}" if curr_price != "N/A" and curr_price is not None else "N/A"
        except (ValueError, TypeError): pass
        try: pe = f"{float(pe):.2f}" if pe != "N/A" and pe is not None else "N/A"
        except (ValueError, TypeError): pass
        try: fwd_pe = f"{float(fwd_pe):.2f}" if fwd_pe != "N/A" and fwd_pe is not None else "N/A"
        except (ValueError, TypeError): pass
        try: peg = f"{float(peg):.2f}" if peg != "N/A" and peg is not None else "N/A"
        except (ValueError, TypeError): pass
        try: high52 = f"{float(high52):,.2f}" if high52 != "N/A" and high52 is not None else "N/A"
        except (ValueError, TypeError): pass
        try: low52 = f"{float(low52):,.2f}" if low52 != "N/A" and low52 is not None else "N/A"
        except (ValueError, TypeError): pass
        try: target_mean = f"{float(target_mean):,.2f}" if target_mean != "N/A" and target_mean is not None else "N/A"
        except (ValueError, TypeError): pass
        
        # Additional fields for FMP from profile
        beta = data_dict.get("beta", "N/A")
        dcf = data_dict.get("dcf", "N/A")
        market_cap = data_dict.get("market_cap", "N/A")

        extra_fmp_lines = []
        if source_name == "FMP":
            try: beta = f"{float(beta):.2f}" if beta != "N/A" and beta is not None else "N/A"
            except (ValueError, TypeError): pass
            try: dcf = f"{float(dcf):,.2f}" if dcf != "N/A" and dcf is not None else "N/A"
            except (ValueError, TypeError): pass
            try: market_cap = f"{float(market_cap):,.0f}" if market_cap != "N/A" and market_cap is not None else "N/A"
            except (ValueError, TypeError): pass

            extra_fmp_lines.append(f"Beta:             {beta}")
            extra_fmp_lines.append(f"DCF (novērtēts):  {dcf} {currency}")
            extra_fmp_lines.append(f"Mkt Cap:          {market_cap} {currency}")

        return (
            f"{i18n.t('source_label')}: {source_name}\n" +
            f"{i18n.t('ticker_info_header', short_name=short_name, ticker=ticker_val)}\n"
            f"{i18n.t('price_label', price=curr_price, currency=currency)}\n"
            f"{i18n.t('52week_range_label', low_52=low52, high_52=high52, currency=currency)}\n"
            f"{i18n.t('pe_label', pe=pe)}\n"
            f"{i18n.t('fwd_pe_label', fwd_pe=fwd_pe)}\n"
            f"{i18n.t('peg_ratio_label', peg=peg)}\n"
            f"{i18n.t('dividends_label', div_str=div_str)}\n"
            f"{i18n.t('analyst_rating_label', rec=rec)}\n"
            f"{i18n.t('target_price_label', target_mean=target_mean, currency=currency)}\n" +
            ("\n".join(extra_fmp_lines) if extra_fmp_lines else "")
        )

    def _fetch_stock_data(self, ticker):
        stats_text = ""
        payload = {}
        api_keys = self._get_api_keys()
        
        # Placeholder for data from each source
        yfinance_data = {}
        alpha_vantage_data = {"error": "No data fetched yet."} # Initialize with error
        fmp_data = {"error": "No data fetched yet."} # Initialize with error

        # ----------------------------------------------------
        # 1. Fetch from Yahoo Finance (always try as a baseline/fallback)
        # ----------------------------------------------------
        try:
            stock_yf = yf.Ticker(ticker)
            info_yf = stock_yf.info or {}
            
            # Populate yfinance_data
            yfinance_data = {
                "ticker": ticker,
                "name": info_yf.get("shortName", ticker),
                "currency": info_yf.get("currency", "USD"),
                "price": info_yf.get("currentPrice") or info_yf.get("regularMarketPrice", "N/A"),
                "pe": info_yf.get("trailingPE", "N/A"),
                "forward_pe": info_yf.get("forwardPE", "N/A"),
                "peg": info_yf.get("pegRatio", "N/A"),
                "52w_high": info_yf.get("fiftyTwoWeekHigh", "N/A"),
                "52w_low": info_yf.get("fiftyTwoWeekLow", "N/A"),
                "recommendation": info_yf.get("recommendationKey", "N/A").upper(),
                "target_price": info_yf.get("targetMeanPrice", "N/A"),
                "dividend_yield_raw": info_yf.get("dividendYield"),
                "sector": info_yf.get("sector", i18n.t("unknown_value")),
                "industry": info_yf.get("industry", i18n.t("unknown_value")),
                "summary": info_yf.get("longBusinessSummary", ""),
                "source": "Yahoo Finance"
            }
            # Format dividend yield
            div_yield_yf = yfinance_data.get("dividend_yield_raw")
            yfinance_data["dividend_yield"] = f"{round(div_yield_yf * 100, 2)}%" if isinstance(div_yield_yf, (int, float)) and div_yield_yf > 0 else "0.0%"
            
            payload.update(yfinance_data) # Start payload with yfinance data
            stats_text = self._format_stats_text(ticker, yfinance_data) # Default stats text
        except Exception as e:
            yfinance_data["error"] = str(e)
            stats_text = f"{i18n.t('source_label')}: Yahoo Finance\n" + i18n.t("error_fetching_ticker_data", ticker=ticker, e=e)
            payload = {"ticker": ticker, "name": ticker, "error": str(e), "source": "Yahoo Finance"} # Minimal payload if yf fails

        # ----------------------------------------------------
        # 2. Fetch from Alpha Vantage
        # ----------------------------------------------------
        if self.selected_source == "Alpha Vantage" or self.selected_source == "⚔️ Cross-Check":
            av_key = api_keys.get("alpha_vantage_api_key")
            if av_key:
                try:
                    av_url = f"https://www.alphavantage.co/query?function=OVERVIEW&symbol={ticker}&apikey={av_key}"
                    av_res = requests.get(av_url, timeout=7)
                    av_res.raise_for_status()
                    av_data_raw = av_res.json()

                    if av_data_raw and av_data_raw.get("Name") and not av_data_raw.get("Error Message"): # Check for valid response and no API error
                        alpha_vantage_data = {
                            "ticker": ticker,
                            "name": av_data_raw.get("Name", ticker),
                            "currency": av_data_raw.get("Currency", "USD"),
                            "price": yfinance_data.get("price", "N/A"), # Alpha Vantage OVERVIEW doesn't have live price, use YF's
                            "52w_high": av_data_raw.get("52WeekHigh", "N/A"),
                            "52w_low": av_data_raw.get("52WeekLow", "N/A"),
                            "pe": av_data_raw.get("PERatio", "N/A"),
                            "peg": av_data_raw.get("PEGRatio", "N/A"),
                            "forward_pe": av_data_raw.get("ForwardPE", "N/A"),
                            "target_price": av_data_raw.get("AnalystTargetPrice", "N/A"),
                            "dividend_yield_raw": av_data_raw.get("DividendYield"),
                            "recommendation": av_data_raw.get("AnalystRatingStrongBuy", "N/A"), # Use an AV-specific rating if available
                            "sector": av_data_raw.get("Sector", i18n.t("unknown_value")),
                            "industry": av_data_raw.get("Industry", i18n.t("unknown_value")),
                            "summary": av_data_raw.get("Description", ""),
                            "source": "Alpha Vantage"
                        }
                        av_div_yield = alpha_vantage_data.get("dividend_yield_raw")
                        if av_div_yield and av_div_yield != "None":
                            try: alpha_vantage_data["dividend_yield"] = f"{round(float(av_div_yield) * 100, 2)}%"
                            except (ValueError, TypeError): alpha_vantage_data["dividend_yield"] = f"{av_div_yield}%"
                        else: alpha_vantage_data["dividend_yield"] = "0.0%"

                        if self.selected_source == "Alpha Vantage":
                            payload.update(alpha_vantage_data)
                            stats_text = self._format_stats_text(ticker, alpha_vantage_data)
                    else: # AV key present but no valid data or API error
                        alpha_vantage_data["error"] = av_data_raw.get("Error Message", i18n.t("error_api_missing_data"))
                        if self.selected_source == "Alpha Vantage":
                            stats_text = f"{i18n.t('source_label')}: Alpha Vantage\n" + alpha_vantage_data["error"] + f" ({i18n.t('source_label')}: Yahoo Finance Fallback)"
                            payload.update(yfinance_data) # Fallback to YF data
                            payload["source"] = "Yahoo Finance (Fallback)"
                except (requests.exceptions.RequestException, json.JSONDecodeError) as e:
                    alpha_vantage_data["error"] = str(e)
                    if self.selected_source == "Alpha Vantage":
                        stats_text = f"{i18n.t('source_label')}: Alpha Vantage\n" + i18n.t("error_fetching_ticker_data", ticker=ticker, e=e) + f" ({i18n.t('source_label')}: Yahoo Finance Fallback)"
                        payload.update(yfinance_data) # Fallback to YF data
                        payload["source"] = "Yahoo Finance (Fallback)"
            elif self.selected_source == "Alpha Vantage": # AV key missing
                stats_text = f"{i18n.t('source_label')}: Alpha Vantage\n" + i18n.t("error_api_key_unavailable") + f" ({i18n.t('source_label')}: Yahoo Finance Fallback)"
                payload.update(yfinance_data) # Fallback to YF data
                payload["source"] = "Yahoo Finance (Fallback)"

        # ----------------------------------------------------
        # 3. Fetch from FMP
        # ----------------------------------------------------
        if self.selected_source == "FMP" or self.selected_source == "⚔️ Cross-Check":
            fmp_key = api_keys.get("fmp_api_key")
            if fmp_key:
                try:
                    fmp_profile_url = f"https://financialmodelingprep.com/api/v3/profile/{ticker}?apikey={fmp_key}"
                    fmp_profile_res = requests.get(fmp_profile_url, timeout=7)
                    fmp_profile_res.raise_for_status()
                    fmp_profile_list = fmp_profile_res.json()

                    if fmp_profile_list and isinstance(fmp_profile_list, list) and len(fmp_profile_list) > 0:
                        profile = fmp_profile_list[0]
                        fmp_data = {
                            "ticker": ticker,
                            "name": profile.get("companyName", ticker),
                            "currency": profile.get("currency", "USD"),
                            "price": profile.get("price", "N/A"),
                            "52w_high": profile.get("yearHigh", "N/A"),
                            "52w_low": profile.get("yearLow", "N/A"),
                            "sector": profile.get("sector", i18n.t("unknown_value")),
                            "industry": profile.get("industry", i18n.t("unknown_value")),
                            "summary": profile.get("description", ""),
                            "beta": profile.get("beta", "N/A"),
                            "dcf": profile.get("dcf", "N/A"),
                            "market_cap": profile.get("mktCap", "N/A"),
                            "source": "FMP",
                            # Defaults, will try to enrich from key-metrics
                            "pe": "N/A", "peg": "N/A", "forward_pe": "N/A", "dividend_yield": "0.0%", "recommendation": "N/A",
                        }

                        # Attempt to enrich with key metrics
                        try:
                            fmp_metrics_url = f"https://financialmodelingprep.com/api/v3/key-metrics-ttm/{ticker}?apikey={fmp_key}"
                            fmp_metrics_res = requests.get(fmp_metrics_url, timeout=7)
                            fmp_metrics_res.raise_for_status()
                            fmp_metrics_list = fmp_metrics_res.json()
                            if fmp_metrics_list and isinstance(fmp_metrics_list, list) and len(fmp_metrics_list) > 0:
                                metrics = fmp_metrics_list[0]
                                fmp_data["pe"] = metrics.get("peRatioTTM", "N/A")
                                fmp_data["peg"] = metrics.get("pegRatioTTM", "N/A")
                                fmp_data["dividend_yield_raw"] = metrics.get("dividendYieldTTM")
                                div_yield_fmp = fmp_data.get("dividend_yield_raw")
                                if isinstance(div_yield_fmp, (int, float)) and div_yield_fmp > 0:
                                    fmp_data["dividend_yield"] = f"{round(div_yield_fmp * 100, 2)}%"
                        except (requests.exceptions.RequestException, json.JSONDecodeError):
                            # Silently fail, keep defaults for metrics
                            pass
                        
                        if self.selected_source == "FMP":
                            payload.update(fmp_data)
                            stats_text = self._format_stats_text(ticker, fmp_data)
                    else: # FMP key present but no valid data
                        fmp_data["error"] = i18n.t("error_api_missing_data")
                        if self.selected_source == "FMP":
                            stats_text = f"{i18n.t('source_label')}: FMP\n" + fmp_data["error"] + f" ({i18n.t('source_label')}: Yahoo Finance Fallback)"
                            payload.update(yfinance_data) # Fallback to YF data
                            payload["source"] = "Yahoo Finance (Fallback)"
                except (requests.exceptions.RequestException, json.JSONDecodeError) as e:
                    fmp_data["error"] = str(e)
                    if self.selected_source == "FMP":
                        stats_text = f"{i18n.t('source_label')}: FMP\n" + i18n.t("error_fetching_ticker_data", ticker=ticker, e=e) + f" ({i18n.t('source_label')}: Yahoo Finance Fallback)"
                        payload.update(yfinance_data) # Fallback to YF data
                        payload["source"] = "Yahoo Finance (Fallback)"
            elif self.selected_source == "FMP": # FMP key missing
                stats_text = f"{i18n.t('source_label')}: FMP\n" + i18n.t("error_api_key_unavailable") + f" ({i18n.t('source_label')}: Yahoo Finance Fallback)"
                payload.update(yfinance_data) # Fallback to YF data
                payload["source"] = "Yahoo Finance (Fallback)"

        # ----------------------------------------------------
        # 4. Cross-Check Logic
        # ----------------------------------------------------
        if self.selected_source == "⚔️ Cross-Check":
            cross_check_summary_lines = [f"{i18n.t('source_label')}: ⚔️ Cross-Check\n"]
            cross_check_payload = {
                "ticker": ticker,
                "name": yfinance_data.get("name", ticker), # Use YF name as primary for summary
                "summary": yfinance_data.get("summary", ""), # Use YF summary as primary
                "data_sources": []
            }

            # Add data from each source if available and without errors
            if "error" not in yfinance_data:
                cross_check_summary_lines.append(f"=== Yahoo Finance ===")
                cross_check_summary_lines.append(f"  Cena: {yfinance_data.get('price', 'N/A')} {yfinance_data.get('currency', 'USD')}")
                cross_check_summary_lines.append(f"  P/E: {yfinance_data.get('pe', 'N/A')}")
                cross_check_payload["data_sources"].append(yfinance_data)
            else:
                cross_check_summary_lines.append(f"=== Yahoo Finance (Kļūda: {yfinance_data['error']}) ===")

            if alpha_vantage_data and "error" not in alpha_vantage_data:
                cross_check_summary_lines.append(f"=== Alpha Vantage ===")
                cross_check_summary_lines.append(f"  Cena: {alpha_vantage_data.get('price', 'N/A')} {alpha_vantage_data.get('currency', 'USD')}")
                cross_check_summary_lines.append(f"  P/E: {alpha_vantage_data.get('pe', 'N/A')}")
                cross_check_payload["data_sources"].append(alpha_vantage_data)
            else:
                cross_check_summary_lines.append(f"=== Alpha Vantage (Kļūda vai atslēga trūkst: {alpha_vantage_data.get('error', i18n.t('error_api_key_unavailable'))}) ===")

            if fmp_data and "error" not in fmp_data:
                cross_check_summary_lines.append(f"=== FMP ===")
                cross_check_summary_lines.append(f"  Cena: {fmp_data.get('price', 'N/A')} {fmp_data.get('currency', 'USD')}")
                cross_check_summary_lines.append(f"  P/E: {fmp_data.get('pe', 'N/A')}")
                cross_check_payload["data_sources"].append(fmp_data)
            else:
                cross_check_summary_lines.append(f"=== FMP (Kļūda vai atslēga trūkst: {fmp_data.get('error', i18n.t('error_api_key_unavailable'))}) ===")

            payload = cross_check_payload
            payload["source"] = "⚔️ Cross-Check"
            stats_text = "\n".join(cross_check_summary_lines)
        
        # Ensure 'source' is always in payload for Gemini's prompt and disclaimer
        if "source" not in payload: # This should ideally not happen now due to default/fallback logic
            payload["source"] = self.selected_source # Default to selected, will be overridden by fallback if that happens

        return stats_text, payload

    def _ask_gemini_single(self, summary_payload, stats_fallback):
        api_key = self.get_gemini_key() if self.get_gemini_key else os.environ.get("GEMINI_API_KEY", "")
        if api_key and HAS_GENAI and summary_payload:
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel("gemini-2.5-flash")
                current_lang_instruction = i18n.t("llm_default_lang_instruction")

                if summary_payload.get("source") == "⚔️ Cross-Check":
                    # Special prompt for Cross-Check
                    prompt = f"""Tu esi stingrs, profesionāls investīciju analītiķis un Komandants Kvarks.
{current_lang_instruction}
Sniedz kodolīgu, praktisku analīzi par uzņēmumu {summary_payload.get('ticker', 'N/A')}, salīdzinot datus no dažādiem avotiem.
Lūk, apkopotie dati no vairākiem avotiem:
{json.dumps(summary_payload.get('data_sources', []), ensure_ascii=False, indent=2)}

Analīzes blokam izmanto šādu formatējumu:
1. 🟢 Datu Avotu Salīdzinājums: Izcel atšķirības cenās un P/E rādītājos starp avotiem, mēģinot paskaidrot potenciālos iemeslus vai iesakot uzticamāko.
2. 🚩 Sarkanie karogi un riski: Vai dati atšķiras pārāk daudz? Vai ir pretrunas?
3. ⚖️ Komandanta Kvarka Rezolūcija: Vispārējs secinājums par datu uzticamību un ieteikums ieguldītājam (Turēt / Pārdot / Nogaidīt, pamatojoties uz datu konsekvenci).
"""
                else:
                    # Original prompt for single source
                    prompt = f"""Tu esi stingrs, profesionāls investīciju analītiķis un Komandants Kvarks.
{current_lang_instruction}
Sniedz kodolīgu, praktisku analīzi par šo uzņēmumu/fondu, pamatojoties uz šādiem rādītājiem:
{json.dumps(summary_payload, ensure_ascii=False, indent=2)}

Strukturē atbildi precīzi šādās 3 sadaļās:
1. 🟢 Biznesa veselība un stiprās puses.
2. 🚩 Sarkanie karogi un riski (vai nav pārvērtēts burbulis).
3. ⚖️ Komandanta Kvarka rezolūcija (konkrēts kopsavilkums ilgtermiņa ieguldītājam: Turēt / Pārdot / Nogaidīt)."""
                
                resp = model.generate_content(prompt)
                
                # Pievienota atruna (disclaimer)
                disclaimer = f"\n\n---\n📌 {i18n.t('source_label')}: {summary_payload.get('source', self.selected_source)}. {i18n.t('analysis_disclaimer_summary')}"
                return resp.text.strip() + disclaimer
            except Exception as e:
                # Pievienota atruna (disclaimer) arī kļūdas gadījumā
                disclaimer_on_error = f"\n\n---\n📌 {i18n.t('source_label')}: {summary_payload.get('source', self.selected_source)}. {i18n.t('analysis_disclaimer_summary')}"
                return i18n.t("error_gemini_analysis", e=str(e)) + disclaimer_on_error # LEO ADDITION: Used new key
        
        # Pievienota atruna (disclaimer) arī fallback gadījumā
        disclaimer_fallback = f"\n\n---\n📌 {i18n.t('source_label')}: {summary_payload.get('source', self.selected_source)}. {i18n.t('analysis_disclaimer_summary')}"
        return stats_fallback + disclaimer_fallback