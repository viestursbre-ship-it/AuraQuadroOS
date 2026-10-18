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
    "lbl_inv_growth_resolution_title": {"lv": "🎯 IZAUGSMES REZOLŪCIJA & RISKA NOVĒRTĒJUMS", "en": "🎯 GROWTH RESOLUTION & RISK ASSESSMENT", "ru": "🎯 РЕЗОЛЮЦИЯ ПО РОСТУ И ОЦЕНКА РИСКА", "uk": "🎯 РЕЗОЛЮЦІЯ З РОSTU ТА ОЦІНКА РИЗИКУ", "it": "🎯 RISOLUZIONE DI CRESCITA & VALUTAZIONE DEL RISCHIO"},
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
    "error_copying": {"lv": "⚠️ Kļūda kopējot: {e}", "en": "⚠️ Error copying: {e}", "ru": "⚠️ Ошибка при копировании: {e}", "uk": "⚠️ Помилка kopēšanas: {e}", "it": "⚠️ Errore durante la copia: {e}"},
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
    "growth_copy_niche": {"lv": "   Slepenais ierocis (Niša): {niche}", "en": "   Secret Weapon (Niche): {niche}", "ru": "   Секретное оружие (Ниша): {niche}", "uk": "   Секретна зброя (Ніша):", "it": "   Arma Segreta (Nicchia): {niche}"},
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
    "btn_calc_dividends": {"lv": "⚡ Aprēķināt Dividenžu Plūsmu", "en": "⚡ Calculate Dividend Flow", "ru": "⚡ Рассчитать Дивидендный Поток", "uk": "⚡ Розрахувати Дивідендний Potok", "it": "⚡ Calcola Flusso Dividendi"},
    "dividends_initial_msg": {"lv": "Nospiediet pogu, lai nolasītu dividenžu datus...", "en": "Press the button to fetch dividend data...", "ru": "Нажмите кнопку, чтобы получить данные о дивидендах...", "uk": "Натисніть кнопку, щоб отримати дані про дивіденди...", "it": "Premi il pulsante per recuperare i dati sui dividendi..."},
    "dividends_initial_analysis_msg": {"lv": "Gaidu dividenžu skenēšanas palaišanu...", "en": "Awaiting dividend scan initiation...", "ru": "Ожидаю запуск сканирования дивидендов...", "uk": "Очікую запуск сканування дивідендів...", "it": "In attesa dell'avvio della scansione dei dividendi..."},
    "btn_calc_dividends_loading": {"lv": "⏳ Vāc dividenžu datus...", "en": "⏳ Collecting dividend data...", "ru": "⏳ Собирает данные о дивидендах...", "uk": "⏳ Збирає datus par dividendēm...", "it": "⏳ Raccolta dati dividendi..."},
    "dividends_scan_progress1": {"lv": "Pārbauda {len_tickers} aktīvus tirgū...\n", "en": "Checking {len_tickers} assets in the market...\n", "ru": "Проверяет {len_tickers} активов на рынке...\n", "uk": "Перевіряє {len_tickers} активів на ринку...\n", "it": "Controllando {len_tickers} asset sul mercato...\n"},
    "dividends_scan_progress2": {"lv": "Kvarks apkopo naudas plūsmas kalendāru...\n", "en": "Quark compiles the cash flow calendar...\n", "ru": "Кварк составляет календарь денежного потока...\n", "uk": "Кварк складає календар grošas plūsmas...\n", "it": "Quark compila il calendario dei flussi di cassa...\n"},
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
    "btn_refresh_market_intel_loading": {"lv": "⏳ Vāc makro datus & Kvarks analizē...", "en": "⏳ Collecting macro data & Quark analyzing...", "ru": "⏳ Собирает макроданные и Кварк анализирует...", "uk": "⏳ Збирає makrodati un Kvarks analizē...", "it": "⏳ Raccolta dati macro & Analisi di Quark..."},
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
    "error_kvark_market_intel": {"lv": "⚠️ Kvarka kļūda: {e}", "en": "⚠️ Quark's error: {e}", "ru": "⚠️ Ошибка Кварка: {e}", "uk": "⚠️ Помилка Kvarka: {e}", "it": "⚠️ Errore di Quark: {e}"},

    # Shared/Generic:
    "rentgen_initial_msg": {"lv": "Ievadiet tikeri un nospiediet 'Rentgens'...", "en": "Enter ticker and press 'X-Ray'...", "ru": "Введите тикер и нажмите 'Рентген'...", "uk": "Введіть тікер і натисніть 'Рентген'...", "it": "Inserisci il ticker e premi 'Raggi X'..."},
    "rentgen_generating_analysis_msg": {"lv": "Kvarks ģenerē fundamentālo analīzi...\n", "en": "Quark is generating fundamental analysis...\n", "ru": "Кварк генерирует фундаментальный анализ...\n", "uk": "Кварк генерує фундаментальний аналіз...\n", "it": "Quark sta generando l'analisi fondamentale...\n"},
    "source_changed_msg": {"lv": "Avota kanāls pārslēgts uz: {source}. Gaidu jaunu signālu 'Rentgens'!", "en": "Source channel switched to: {source}. Awaiting new 'X-Ray' signal!", "ru": "Канал источника переключен на: {source}. Ожидаю новый сигнал 'Рентген'!", "uk": "Канал джерела перемкнено на: {source}. Очікую новий сигнал 'Рентген'!", "it": "Canale sorgente cambiato in: {source}. In attesa di nuovo segnale 'Raggi X'!"},
    "fetching_data_msg": {"lv": "Nolasa tirgus datus par {ticker} no {source}...\n", "en": "Fetching market data for {ticker} from {source}...\n", "ru": "Получение рыночных данных для {ticker} из {source}...\n", "uk": "Отримання rinka datus par {ticker} no {source}...\n", "it": "Recupero dati di mercato per {ticker} da {source}...\n"},
    "ticker_info_header": {"lv": "=== {short_name} ({ticker}) ===", "en": "=== {short_name} ({ticker}) ===", "ru": "=== {short_name} ({ticker}) ===", "uk": "=== {short_name} ({ticker}) ===", "it": "=== {short_name} ({ticker}) ==="},
    "price_label": {"lv": "Cena:             {price} {currency}", "en": "Price:             {price} {currency}", "ru": "Цена:             {price} {currency}", "uk": "Ціна:             {price} {currency}", "it": "Prezzo:             {price} {currency}"},
    "52week_range_label": {"lv": "52 ned. diapazons:{low_52} - {high_52} {currency}", "en": "52 wk. range:    {low_52} - {high_52} {currency}", "ru": "Диапазон 52 нед.:{low_52} - {high_52} {currency}", "uk": "Діапазон 52 тижнів:{low_52} - {high_52} {currency}", "it": "Intervallo 52 sett.:{low_52} - {high_52} {currency}"},
    "pe_label": {"lv": "P/E (faktiskais): {pe}", "en": "P/E (actual):    {pe}", "ru": "P/E (фактический): {pe}", "uk": "P/E (фактичний): {pe}", "it": "P/E (effettivo): {pe}"},
    "fwd_pe_label": {"lv": "P/E (prognoze):   {fwd_pe}", "en": "P/E (forecast):  {fwd_pe}", "ru": "P/E (прогноз):   {fwd_pe}", "uk": "P/E (прогноз):   {fwd_pe}", "it": "P/E (previsto):   {fwd_pe}"},
    "peg_ratio_label": {"lv": "PEG koeficients:  {peg}", "en": "PEG ratio:       {peg}", "ru": "Коэффициент PEG:  {peg}", "uk": "Коефіcients PEG:  {peg}", "it": "Rapporto PEG:      {peg}"},
    "dividends_label": {"lv": "Dividendes:       {div_str}", "en": "Dividends:       {div_str}", "ru": "Дивиденды:       {div_str}", "uk": "Дивіденди:       {div_str}", "it": "Dividendi:       {div_str}"},
    "analyst_rating_label": {"lv": "Analītiķu vērt.:  {rec}", "en": "Analyst rating:  {rec}", "ru": "Оценка аналитиков:  {rec}", "uk": "Оцінка аналітиків:  {rec}", "it": "Valutazione analisti:  {rec}"},
    "target_price_label": {"lv": "Mērķa vid. cena:  {target_mean} {currency}", "en": "Target avg. price: {target_mean} {currency}", "ru": "Целевая сред. цена: {target_mean} {currency}", "uk": "Цільова ser. cena: {target_mean} {currency}", "it": "Prezzo medio target: {target_mean} {currency}"},
    "unknown_value": {"lv": "Nezināms", "en": "Unknown", "ru": "Неизвестно", "uk": "Невідомо", "it": "Sconosciuto"},
    "error_fetching_ticker_data": {"lv": "⚠️ Kļūda nolasot {ticker}: {e}", "en": "⚠️ Error fetching {ticker}: {e}", "ru": "⚠️ Ошибка при получении {ticker}: {e}", "uk": "⚠️ Помилка pri iegūšanā {ticker}: {e}", "it": "⚠️ Errore nel recupero {ticker}: {e}"},
    "analysis_disclaimer_summary": {"lv": "Informatīvs analīzes kopsavilkums.", "en": "Informative analysis summary.", "ru": "Информативное резюме анализа.", "uk": "Інформативне резюме аналізу.", "it": "Riepilogo analitico informativo."},

    # Additional keys for prompts and UI (LEO ADDITION)
    "lbl_inv_portfolio_audit_plan": {"lv": "🐾 KOMANDANTA KVARKA AUDITA PLĀNS", "en": "🐾 COMMANDER KVARK'S AUDIT PLAN", "ru": "🐾 ПЛАН АУДИТА КОМАНДИРА КВАРКА", "uk": "🐾 ПЛАН АУДИТУ КОМАНДИРА КВАРКА", "it": "🐾 PIANO DI AUDIT DEL COMANDANTE KVARKS"},
    "audit_sell_liquidate": {"lv": "✂️ LIKVIDĒT / PĀRDOT", "en": "✂️ LIQUIDATE / SELL", "ru": "✂️ ЛИКВИДИРОВАТЬ / ПРОДАТЬ", "uk": "✂️ ЛІКВІДУВАТИ / ПРОДАТИ", "it": "✂️ LIQUIDARE / VENDERE"},
    "audit_cash_allocation": {"lv": "🎯 BRĪVO LĪDZEKĻU SADALE", "en": "🎯 FREE CASH ALLOCATION", "ru": "🎯 РАСПРЕДЕЛЕНИЕ СВОБОДНЫХ СРЕДСТВ", "uk": "🎯 РОЗПОДІЛ ВІЛЬНИХ КОШТІВ", "it": "🎯 ALLOCAZIONE DEL CAPITALE DISPONIBILE"},
    "llm_default_lang_instruction": {"lv": "Atbildi latviešu valodā.", "en": "Respond in English.", "ru": "Ответь на русском языке.", "uk": "Відповідайте українською мовою.", "it": "Rispondi in italiano."},
    "source_label": {"lv": "Avots", "en": "Source", "ru": "Источник", "uk": "Джерело", "it": "Fonte"},
    "error_gemini_analysis": {"lv": "⚠️ Gemini analīzes kļūda: {e}", "en": "⚠️ Gemini analysis error: {e}", "ru": "⚠️ Oшибка анализа Gemini: {e}", "uk": "⚠️ Помилка аналізу Gemini: {e}", "it": "⚠️ Errore di analisi Gemini: {e}"}
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
        f = self.tab_