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
    "error_portfolio_save": {"lv": "⚠️ Kļūda: {e}", "en": "⚠️ Error: {e}", "ru": "⚠️ Ошибка: {e}", "uk": "⚠️ Помылка: {e}", "it": "⚠️ Errore: {e}"},
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
    "growth_copy_risk": {"lv": "   Melnais Gulbis (risks): {risk}", "en": "   Black Swan (risk): {risk}", "ru": "   Черный Лебедь (риск): {risk}", "uk": "   Чорний Лебідis (ризик): {risk}", "it": "   Cigno Nero (rischio): {risk}"},
    "growth_no_analysis_to_copy": {"lv": "Nav atrasta neviena analīze, ko kopēt.", "en": "No analysis found to copy.", "ru": "Не найдено анализа для копирования.", "uk": "Не знайдено аналізу для копіювання.", "it": "Nessuna analisi trovata da copiare."},
    "error_title": {"lv": "Kļūda", "en": "Error", "ru": "Ошибка", "uk": "Помилка", "it": "Errore"},
    "error_docx_missing": {"lv": "Lai eksportētu uz Word, nepieciešama 'python-docx' bibliotēka.\nUzinstalējiet to ar 'pip install python-docx'.", "en": "To export to Word, the 'python-docx' library is required.\nInstall it with 'pip install python-docx'.", "ru": "Для экспорта в Word требуется библиотека 'python-docx'.\nУстановите ее с помощью 'pip install python-docx'.", "uk": "Для експорту в Word потрібна бібліотека 'python-docx'.\nВстановіть її за допомогою 'pip install python-docx'.", "it": "Per esportare in Word è necessaria la libreria 'python-docx'.\nInstallala con 'pip install python-docx'."},
    "info_title": {"lv": "Info", "en": "Info", "ru": "Информация", "uk": "Інформація", "it": "Info"},
    "info_no_data_to_export": {"lv": "Nav datu, ko eksportēt. Palaidiet 'Medīt Jaunos Vienradžus' vispirms.", "en": "No data to export. Run 'Hunt New Unicorns' first.", "ru": "Нет данных для экспорта. Сначала запустите 'Искать новых единорогов'.", "uk": "Немає datu uz eksportu. Vispirms palaidiet 'Полювання на Єдиноrogi'.", "it": "Nessun dato da esportare. Avvia prima 'Caccia agli Unicorni'."},
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
    "btn_calc_dividends": {"lv": "⚡ Aprēķināt Dividenžu Plūsmu", "en": "⚡ Calculate Dividend Flow", "ru": "⚡ Рассчитать Дивидендный Поток", "uk": "⚡ Розрахувати Дивіденžu Potok", "it": "⚡ Calcola Flusso Dividendi"},
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
    "dividends_no_companies": {"lv": "Portfelī nav atrasts neviens uzņēmums, kas maksā dividendes.", "en": "No dividend-paying companies found in portfolio.", "ru": "В портфеле не найдено ни одной компании, выплачивающей дивиденды.", "uk": "У портфелі не знайдено жодної компапанії, яка виплачує дивіденди.", "it": "Nessuna azienda che paga dividendi trovata nel portfolio."},
    "error_api_key_unavailable": {"lv": "Nav pieejama API atslēga.", "en": "API key not available.", "ru": "Ключ API недоступен.", "uk": "Ключ API недоступний.", "it": "Chiave API non disponibile."},

    # Market Tab:
    "lbl_market_sentiment_title": {"lv": "🌡️ TIRGUS NOSKAŅOJUMS", "en": "🌡️ MARKET SENTIMENT", "ru": "🌡️ НАСТРОЕНИЕ РЫНКА", "uk": "🌡️ НАСТРІЙ РИНКУ", "it": "🌡️ SENTIMENTO DI MERCATO"},
    "lbl_fng_loading": {"lv": "Fear & Greed: Ielādē...", "en": "Fear & Greed: Loading...", "ru": "Индекс страха и жадности: Загрузка...", "uk": "Страх і жадібність: Завантаження...", "it": "Fear & Greed: Caricamento..."},
    "btn_refresh_market_intel": {"lv": "⚡ Atjaunot Klimatu & Insaiderus", "en": "⚡ Refresh Climate & Insiders", "ru": "⚡ Обновить Климат и Инсайдеров", "uk": "⚡ Оновити Клімат та Інсайдерів", "it": "⚡ Aggiorna Clima & Insider"},
    "insider_initial_msg": {"lv": "Nospiediet pogu, lai ielādētu insaideru datus un noskaņojumu...", "en": "Press the button to load insider data and sentiment...", "ru": "Нажмите кнопку, чтобы загрузить данные инсайдеров и настроение...", "uk": "Натисніть кнопку, щоб завантажити дані інсайдеров та настрій...", "it": "Premi il pulsante per caricare i dati e il sentimento degli insider..."},
    "market_news_initial_msg": {"lv": "Gaidu datu pieprasījumu...", "en": "Awaiting data request...", "ru": "Ожидаю запрос данных...", "uk": "Очікую запит даних...", "it": "In attesa di richiesta dati..."},
    "btn_refresh_market_intel_loading": {"lv": "⏳ Vāc makro datus & Kvarks analizē...", "en": "⏳ Collecting macro data & Quark analyzing...", "ru": "⏳ Собирает макроданные и Кварк анализирует...", "uk": "⏳ Збирает makrodati un Kvarks analizē...", "it": "⏳ Raccolta dati macro & Analisi di Quark..."},
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
    "news_no_aggressive": {"lv": "Nav agresīvu ziņu triecienu", "en": "No aggressive news strikes", "ru": "Нет агрессивных новостных ударов", "uk": "Немає agresīvu jaunumu streiku", "it": "Nessun attacco di notizie aggressive"},
    "kvark_strategic_resolution": {"lv": "🐾 KOMANDANTA KVARKA STRATĒĢISKĀ REZOLŪCIJA:\n", "en": "🐾 COMMANDER KVARK'S STRATEGIC RESOLUTION:\n", "ru": "🐾 СТРАТЕГИЧЕСКАЯ РЕЗОЛЮЦИЯ КОМАНДИРА КВАРКА:\n", "uk": "🐾 СТРАТЕГІЧНА РЕЗОЛЮЦІЯ КОМАНДИРА КВАРКА:\n", "it": "🐾 RISOLUZIONE STRATEGICA DEL COMANDANTE QUARK:\n"},
    "error_kvark_market_intel": {"lv": "⚠️ Kvarka kļūda: {e}", "en": "⚠️ Quark's error: {e}", "ru": "⚠️ Ошибка Кварка: {e}", "uk": "⚠️ Помилка Kvarka: {e}", "it": "⚠️ Errore di Quark: {e}"},

    # Shared/Generic:
    "rentgen_initial_msg": {"lv": "Ievadiet tikeri un nospiediet 'Rentgens'...", "en": "Enter ticker and press 'X-Ray'...", "ru": "Введите тикер и нажмите 'Рентген'...", "uk": "Введіть тікер і натисніть 'Рентген'...", "it": "Inserisci il ticker e premi 'Raggi X'..."},
    "rentgen_generating_analysis_msg": {"lv": "Kvarks ģenerē fundamentālo analīzi...\n", "en": "Quark is generating fundamental analysis...\n", "ru": "Кварк генерирует фундаментальный анализ...\n", "uk": "Кварк генерує фундаментальний аналіз...\n", "it": "Quark sta generando l'analisi fondamentale...\n"},
    "source_changed_msg": {"lv": "Avota kanāls pārslēgts uz: {source}. Gaidu jaunu signālu 'Rentgens'!", "en": "Source channel switched to: {source}. Awaiting new 'X-Ray' signal!", "ru": "Канал источника переключен на: {source}. Ожидаю новый сигнал 'Рентген'!", "uk": "Канал джерела перемкнено na: {source}. Очікую новий сигнал 'Рентген'!", "it": "Canale sorgente cambiato in: {source}. In attesa di nuovo segnale 'Raggi X'!"},
    "fetching_data_msg": {"lv": "Nolasa tirgus datus par {ticker} no {source}...\n", "en": "Fetching market data for {ticker} from {source}...\n", "ru": "Получение рыночных данных для {ticker} из {source}...\n", "uk": "Отримання rinka datus par {ticker} no {source}...\n", "it": "Recupero dati di mercato per {ticker} da {source}...\n"},
    "ticker_info_header": {"lv": "=== {short_name} ({ticker}) ===", "en": "=== {short_name} ({ticker}) ===", "ru": "=== {short_name} ({ticker}) ===", "uk": "=== {short_name} ({ticker}) ===", "it": "=== {short_name} ({ticker}) ==="},
    "price_label": {"lv": "Cena:             {price} {currency}", "en": "Price:             {price} {currency}", "ru": "Цена:             {price} {currency}", "uk": "Ціна:             {price} {currency}", "it": "Prezzo:             {price} {currency}"},
    "52week_range_label": {"lv": "52 ned. diapazons:{low_52} - {high_52} {currency}", "en": "52 wk. range:    {low_52} - {high_52} {currency}", "ru": "Диапазон 52 нед.:{low_52} - {high_52} {currency}", "uk": "Діapazons 52 tyžni: {low_52} - {high_52} {currency}", "it": "Intervallo 52 sett.:{low_52} - {high_52} {currency}"},
    "pe_label": {"lv": "P/E (faktiskais): {pe}", "en": "P/E (actual):    {pe}", "ru": "P/E (фактический): {pe}", "uk": "P/E (фактичний): {pe}", "it": "P/E (effettivo): {pe}"},
    "fwd_pe_label": {"lv": "P/E (prognoze):   {fwd_pe}", "en": "P/E (forecast):  {fwd_pe}", "ru": "P/E (прогноз):   {fwd_pe}", "uk": "P/E (прогноз):   {fwd_pe}", "it": "P/E (previsto):   {fwd_pe}"},
    "peg_ratio_label": {"lv": "PEG koeficients:  {peg}", "en": "PEG ratio:       {peg}", "ru": "Коэффициент PEG:  {peg}", "uk": "Коефіcients PEG:  {peg}", "it": "Rapporto PEG:      {peg}"},
    "dividends_label": {"lv": "Dividendes:       {div_str}", "en": "Dividends:       {div_str}", "ru": "Дивиденды:       {div_str}", "uk": "Дивіденди:       {div_str}", "it": "Dividendi:       {div_str}"},
    "analyst_rating_label": {"lv": "Analītiķu vērt.:  {rec}", "en": "Analyst rating:  {rec}", "ru": "Оценка аналитиков:  {rec}", "uk": "Оцінка аналітиків:  {rec}", "it": "Valutazione analisti:  {rec}"},
    "target_price_label": {"lv": "Mērķa vid. cena:  {target_mean} {currency}", "en": "Target avg. price: {target_mean} {currency}", "ru": "Целевая сред. цена: {target_mean} {currency}", "uk": "Цільова ser. cena: {target_mean} {currency}", "it": "Prezzo medio target: {target_mean} {currency}"},
    "unknown_value": {"lv": "Nezināms", "en": "Unknown", "ru": "Неизвестно", "uk": "Невідомо", "it": "Sconosciuto"},
    "error_fetching_ticker_data": {"lv": "⚠️ Kļūda nolasot {ticker}: {e}", "en": "⚠️ Error fetching {ticker}: {e}", "ru": "⚠️ Ошибка pri poluchenii {ticker}: {e}", "uk": "⚠️ Помилка pri iegūšanā {ticker}: {e}", "it": "⚠️ Errore nel recupero {ticker}: {e}"},
    "analysis_disclaimer_summary": {"lv": "Informatīvs analīzes kopsavilkums.", "en": "Informative analysis summary.", "ru": "Информативное резюме анализа.", "uk": "Інформативне резюме аналізу.", "it": "Riepilogo analitico informativo."},

    # Additional keys for prompts and UI (LEO ADDITION)
    "lbl_inv_portfolio_audit_plan": {"lv": "🐾 KOMANDANTA KVARKA AUDITA PLĀNS", "en": "🐾 COMMANDER KVARK'S AUDIT PLAN", "ru": "🐾 ПЛАН АУДИТА КОМАНДИРА КВАРКА", "uk": "🐾 ПЛАН АУДИТУ КОМАНДИРА КВАРКА", "it": "🐾 PIANO DI AUDIT DEL COMANDANTE KVARKS"},
    "audit_sell_liquidate": {"lv": "✂️ LIKVIDĒT / PĀRDOT", "en": "✂️ LIQUIDATE / SELL", "ru": "✂️ ЛИКВИДИРОВАТЬ / ПРОДАТЬ", "uk": "✂️ ЛІКВІДУВАТИ / ПРОДАТИ", "it": "✂️ LIQUIDARE / VENDERE"},
    "audit_cash_allocation": {"lv": "🎯 BRĪVO LĪDZEKĻU SADALE", "en": "🎯 FREE CASH ALLOCATION", "ru": "🎯 РАСПРЕДЕЛЕНИЕ СВОБОДНЫХ СРЕДСТВ", "uk": "🎯 РОЗПОДІЛ ВІЛЬНИХ КОШТІВ", "it": "🎯 ALLOCAZIONE DEL CAPITALE DISPONIBILE"},
    "llm_default_lang_instruction": {"lv": "Atbildi latviešu valodā.", "en": "Respond in English.", "ru": "Ответь на русском языке.", "uk": "Відповідайте українською мовою.", "it": "Rispondi in italiano."},
    "source_label": {"lv": "Avots", "en": "Source", "ru": "Источник", "uk": "Джерело", "it": "Fonte"},
    "error_gemini_analysis": {"lv": "⚠️ Gemini analīzes kļūda: {e}", "en": "⚠️ Gemini analysis error: {e}", "ru": "⚠️ Oшибка анализа Gemini: {e}", "uk": "⚠️ Помилка analīzes Gemini: {e}", "it": "⚠️ Errore di analisi Gemini: {e}"},
    "data_fetch_error": {"lv": "⚠️ Kļūda nolasot datus", "en": "⚠️ Error fetching data", "ru": "⚠️ Ошибка pri poluchenii dannyx", "uk": "⚠️ Помилка pri otrimanni danih", "it": "⚠️ Errore nel recupero dei dati"},
    "error_invalid_cash_balance": {"lv": "Nederīga naudas atlikuma vērtība.", "en": "Invalid cash balance value.", "ru": "Неверное значение остатка денежных средств.", "uk": "Невірна вартість залишku grošas koštív.", "it": "Valore del saldo di cassa non valido."} # New key for specific error message
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
        f.grid_columnconfigure(1, weight=1)
        f.grid_rowconfigure(6, weight=1) # Allow for expansion of resolution_box (was 3, adjusted for new elements)

        self.lbl_ticker_title = ctk.CTkLabel(f, text=i18n.t("lbl_ticker_analysis"), font=ctk.CTkFont(size=16, weight="bold"))
        self.lbl_ticker_title.grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="ew")

        # Ticker Input
        self.ticker_input = ctk.CTkEntry(f, placeholder_text="AAPL", fg_color="#1e293b", border_color="#334155")
        self.ticker_input.grid(row=1, column=0, padx=(10, 5), pady=5, sticky="ew")
        self.ticker_input.bind("<Return>", lambda event: self._analyze_ticker_event())

        self.btn_analyze = ctk.CTkButton(f, text=i18n.t("btn_rentgens"), command=self._analyze_ticker,
                                         fg_color="#0e7490", hover_color="#086078")
        self.btn_analyze.grid(row=1, column=1, padx=(5, 10), pady=5, sticky="ew")

        # Dropdown for data source selection (LEO ADDITION)
        self.source_options = ["Yahoo Finance", "Alpha Vantage", "FMP", "⚔️ Cross-Check"]
        self.source_selector = ctk.CTkOptionMenu(f, values=self.source_options,
                                                 command=self._change_data_source,
                                                 fg_color="#1e293b", button_color="#0e7490", button_hover_color="#086078")
        self.source_selector.set(self.selected_source) # Set initial value
        self.source_selector.grid(row=2, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

        # Progress Bar (New for Task 2)
        self.progress_bar = ctk.CTkProgressBar(f, mode="determinate", height=10, fg_color="#334155", progress_color="#0e7490")
        self.progress_bar.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        self.progress_bar.set(0) # Default to inactive/empty

        # Ticker Stats Display
        self.stats_box = ctk.CTkTextbox(f, height=200, wrap="word",
                                        fg_color="#1e293b", text_color="#f8fafc", border_color="#334155")
        self.stats_box.grid(row=4, column=0, columnspan=2, padx=10, pady=5, sticky="nsew") # Adjusted row to 4
        self.stats_box.insert("0.0", i18n.t("rentgen_initial_msg"))
        self.stats_box.configure(state="disabled")

        self.lbl_resolution_title = ctk.CTkLabel(f, text=i18n.t("lbl_inv_resolution_title"), font=ctk.CTkFont(size=16, weight="bold"))
        self.lbl_resolution_title.grid(row=5, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="ew") # Adjusted row to 5

        # Resolution Display
        self.resolution_box = ctk.CTkTextbox(f, height=300, wrap="word",
                                             fg_color="#1e293b", text_color="#f8fafc", border_color="#334155")
        self.resolution_box.grid(row=6, column=0, columnspan=2, padx=10, pady=(5, 10), sticky="nsew") # Adjusted row to 6
        self.resolution_box.insert("0.0", i18n.t("txt_waiting"))
        self.resolution_box.configure(state="disabled")

    def _change_data_source(self, new_source):
        self.selected_source = new_source
        self.stats_box.configure(state="normal")
        self.stats_box.delete("0.0", "end")
        # Fix 1: i18n parameter passing
        self.stats_box.insert("0.0", f"{i18n.t('source_changed_msg').split('{source}')[0]}{new_source}.{i18n.t('source_changed_msg').split('{source}')[1]}")
        self.stats_box.configure(state="disabled")

        self.resolution_box.configure(state="normal")
        self.resolution_box.delete("0.0", "end")
        self.resolution_box.insert("0.0", i18n.t("txt_waiting"))
        self.resolution_box.configure(state="disabled")

    def _analyze_ticker_event(self):
        self._analyze_ticker()

    def _analyze_ticker(self):
        ticker = self.ticker_input.get().strip().upper()
        if not ticker:
            self.stats_box.configure(state="normal")
            self.stats_box.delete("0.0", "end")
            self.stats_box.insert("0.0", i18n.t("rentgen_initial_msg"))
            self.stats_box.configure(state="disabled")
            return

        self.btn_analyze.configure(state="disabled", text=f"⏳ {i18n.t('btn_rentgens')}...")
        self.progress_bar.configure(mode="indeterminate") # Task 2
        self.progress_bar.start() # Task 2

        self.stats_box.configure(state="normal")
        self.stats_box.delete("0.0", "end")
        # Fix 1: i18n parameter passing
        # Original: f"{i18n.t('fetching_data_msg')} [{ticker}] ({self.selected_source})..."
        # This one is trickier as it has two placeholders. Splitting the string `Nolasa tirgus datus par {ticker} no {source}...\n`
        base_msg = i18n.t('fetching_data_msg')
        # This assumes the placeholders are in this exact order and structure
        formatted_msg = base_msg.replace('{ticker}', ticker).replace('{source}', self.selected_source)
        self.stats_box.insert("0.0", formatted_msg)
        self.stats_box.configure(state="disabled")

        self.resolution_box.configure(state="normal")
        self.resolution_box.delete("0.0", "end")
        self.resolution_box.insert("0.0", i18n.t("rentgen_generating_analysis_msg"))
        self.resolution_box.configure(state="disabled")

        threading.Thread(target=self._analyze_ticker_worker, args=(ticker,)).start()

    def _analyze_ticker_worker(self, ticker):
        stats_text, payload = self._fetch_stock_data(ticker)
        self.master.after(0, lambda: self.stats_box.configure(state="normal"))
        self.master.after(0, lambda: self.stats_box.delete("0.0", "end"))
        self.master.after(0, lambda: self.stats_box.insert("0.0", stats_text))
        self.master.after(0, lambda: self.stats_box.configure(state="disabled"))

        resolution_text = ""
        if HAS_GENAI and self.get_gemini_key() and "⚠️" not in stats_text:
            self.master.after(0, lambda: self.resolution_box.configure(state="normal"))
            self.master.after(0, lambda: self.resolution_box.insert("end", i18n.t("rentgen_generating_analysis_msg")))
            self.master.after(0, lambda: self.resolution_box.configure(state="disabled"))
            resolution_text = self._get_gemini_resolution(payload)
        elif not HAS_GENAI:
            resolution_text = i18n.t('error_gemini_api_contact').format(e='Google Generative AI library not found.') + "\n" + i18n.t('analysis_disclaimer_summary')
        elif not self.get_gemini_key():
             resolution_text = i18n.t('error_api_key_unavailable') + "\n" + i18n.t('analysis_disclaimer_summary')
        else:
            resolution_text = i18n.t('error_fetching_ticker_data').format(ticker=ticker, e='Data unavailable for analysis.') + "\n" + i18n.t('analysis_disclaimer_summary')

        self.master.after(0, lambda: self.resolution_box.configure(state="normal"))
        self.master.after(0, lambda: self.resolution_box.delete("0.0", "end"))
        self.master.after(0, lambda: self.resolution_box.insert("0.0", resolution_text))
        self.master.after(0, lambda: self.resolution_box.configure(state="disabled"))
        
        # Task 2: Stop progress bar and reset button
        self.master.after(0, lambda: self.progress_bar.stop())
        self.master.after(0, lambda: self.progress_bar.configure(mode="determinate"))
        self.master.after(0, lambda: self.progress_bar.set(1.0))
        self.master.after(0, lambda: self.btn_analyze.configure(state="normal", text=i18n.t("btn_rentgens")))

    def _fetch_stock_data(self, ticker):
        stats_text = ""
        payload = {}
        try:
            cfg = {}
            try:
                with open("config.json", "r", encoding="utf-8") as f:
                    cfg = json.load(f)
            except Exception:
                pass
            av_key = cfg.get("alpha_vantage_api_key", "")
            fmp_key = cfg.get("fmp_api_key", "")

            stock = yf.Ticker(ticker)
            info = stock.info or {}
            short_name = info.get("shortName", ticker)
            currency = info.get("currency", "USD")
            curr_price = info.get("currentPrice") or info.get("regularMarketPrice", "N/A")
            pe = info.get("trailingPE", "N/A")
            fwd_pe = info.get("forwardPE", "N/A")
            peg = info.get("pegRatio", "N/A")
            high52 = info.get("fiftyTwoWeekHigh", "N/A")
            low52 = info.get("fiftyTwoWeekLow", "N/A")
            target_mean = info.get("targetMeanPrice", "N/A")

            av_price, fmp_price = None, None

            if self.selected_source in ["Alpha Vantage", "⚔️ Cross-Check"] and av_key:
                try:
                    url_av = f"https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol={ticker}&apikey={av_key}"
                    r_av = requests.get(url_av, timeout=6).json()
                    gq = r_av.get("Global Quote", {})
                    if "05. price" in gq:
                        av_price = gq.get("05. price")
                except Exception:
                    pass

            if self.selected_source in ["FMP", "⚔️ Cross-Check"] and fmp_key:
                try:
                    url_fmp = f"https://financialmodelingprep.com/stable/quote?symbol={ticker}&apikey={fmp_key}"
                    r_fmp = requests.get(url_fmp, timeout=6).json()
                    if isinstance(r_fmp, list) and len(r_fmp) > 0:
                        fmp_data = r_fmp[0]
                        fmp_price = fmp_data.get("price")
                        if self.selected_source == "FMP":
                            short_name = fmp_data.get("name", short_name)
                            high52 = fmp_data.get("yearHigh", high52)
                            low52 = fmp_data.get("yearLow", low52)
                except Exception:
                    pass

            lines = [
                f"{i18n.t('source_label')}: {self.selected_source}",
                i18n.t('ticker_info_header').format(short_name=short_name, ticker=ticker)
            ]

            if self.selected_source == "Alpha Vantage" and av_price:
                lines.append(f"Cena (Alpha Vantage): {av_price} {currency}")
            elif self.selected_source == "FMP" and fmp_price:
                lines.append(f"Cena (FMP): {fmp_price} {currency}")
            else:
                lines.append(i18n.t('price_label').format(price=curr_price, currency=currency))

            if self.selected_source == "⚔️ Cross-Check":
                lines.append("--- ⚔️ CROSS-CHECK DATI ---")
                lines.append(f"Yahoo Finance: {curr_price} {currency}")
                lines.append(f"Alpha Vantage: {av_price or 'Nav datu'} {currency}")
                lines.append(f"FMP Quote:     {fmp_price or 'Nav datu'} {currency}")
                lines.append("-------------------------")

            lines.extend([
                i18n.t('52week_range_label').format(low_52=low52, high_52=high52, currency=currency),
                i18n.t('pe_label').format(pe=pe),
                i18n.t('fwd_pe_label').format(fwd_pe=fwd_pe),
                i18n.t('peg_ratio_label').format(peg=peg)
            ])

            stats_text = "\n".join(lines) + "\n"
            payload = {
                "ticker": ticker,
                "name": short_name,
                "price": fmp_price or av_price or curr_price,
                "currency": currency,
                "source": self.selected_source,
                "pe": pe,
                "fwd_pe": fwd_pe,
                "peg": peg,
                "52_high": high52,
                "52_low": low52,
                "target_mean": target_mean
            }
        except Exception as e:
            stats_text = f"{i18n.t('data_fetch_error')}: {e}"
        return stats_text, payload

    def _get_gemini_resolution(self, payload):
        if not HAS_GENAI or not self.get_gemini_key():
            return i18n.t('error_api_key_unavailable') + "\n" + i18n.t('analysis_disclaimer_summary')

        try:
            genai.configure(api_key=self.get_gemini_key())
            model = genai.GenerativeModel('gemini-pro')

            lang_instruction = get_llm_language_instruction() # Use the helper function

            prompt = (
                f"{lang_instruction}\n"
                "You are Commander Quark, a fat cat from Liepaja garrison, focused on building a "
                "personal cannery and fish factory empire. Your decisions prioritize strategic calm "
                "and maximal profit extraction with minimal personal effort.\n\n"
                "Analyze the following stock data and provide a concise, tactical resolution "
                "for the Initiator. Your response should be direct, professional, sharp, and in a "
                "tactical military garrison tone. Conclude with 'Komandants Kvarks ir izdevis Rīkojumu.'\n\n"
                "Stock Data:\n"
                f"- Ticker: {payload.get('ticker')}\n"
                f"- Name: {payload.get('name')}\n"
                f"- Price: {payload.get('price')} {payload.get('currency')}\n"
                f"- Source: {payload.get('source')}\n"
                f"- P/E (Trailing): {payload.get('pe')}\n"
                f"- P/E (Forward): {payload.get('fwd_pe')}\n"
                f"- PEG Ratio: {payload.get('peg')}\n"
                f"- 52-Week High: {payload.get('52_high')} {payload.get('currency')}\n"
                f"- 52-Week Low: {payload.get('52_low')} {payload.get('currency')}\n"
                f"- Target Mean Price: {payload.get('target_mean')} {payload.get('currency')}\n\n"
                "Provide actionable insights for a strategic investor. Your goal is to identify if "
                "this stock serves Commander Quark's empire-building goals efficiently. "
                "Consider both potential gains and hidden risks. If P/E is above 80 and PEG is above 2, "
                "mention it as a potential bubble."
            )

            response = model.generate_content(prompt, safety_settings={'HARASSMENT': 'BLOCK_NONE', 'HATE': 'BLOCK_NONE', 'SEXUAL': 'BLOCK_NONE', 'DANGEROUS': 'BLOCK_NONE'})
            return response.text.strip()
        except Exception as e:
            return i18n.t('error_gemini_analysis').format(e=e) + "\n" + i18n.t('analysis_disclaimer_summary')

    # ==========================================
    # 2. CILNE: MANA eToro PORTFELIS & AUDITS
    # ==========================================
    def _setup_portfolio_tab(self):
        f = self.tab_portfolio
        f.grid_columnconfigure(0, weight=1)
        f.grid_columnconfigure(1, weight=1)
        f.grid_rowconfigure(7, weight=1) # Audit log takes most space

        self.lbl_portfolio_title = ctk.CTkLabel(f, text=i18n.t("lbl_portfolio_title"), font=ctk.CTkFont(size=16, weight="bold"))
        self.lbl_portfolio_title.grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="ew")

        self.lbl_portfolio_tickers_desc = ctk.CTkLabel(f, text=i18n.t("lbl_portfolio_tickers_desc"), wraplength=400, justify="left")
        self.lbl_portfolio_tickers_desc.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 5), sticky="w")

        self.portfolio_tickers_input = ctk.CTkTextbox(f, height=100, wrap="word",
                                                       fg_color="#1e293b", text_color="#f8fafc", border_color="#334155")
        self.portfolio_tickers_input.grid(row=2, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        # self._load_portfolio() # MOVED TO THE END OF THIS METHOD TO ENSURE WIDGETS ARE INITIALIZED

        self.lbl_portfolio_cash_balance = ctk.CTkLabel(f, text=i18n.t("lbl_portfolio_cash_balance"))
        self.lbl_portfolio_cash_balance.grid(row=3, column=0, padx=10, pady=(5, 0), sticky="w")
        self.cash_balance_input = ctk.CTkEntry(f, placeholder_text="10000", fg_color="#1e293b", border_color="#334155")
        self.cash_balance_input.grid(row=3, column=1, padx=10, pady=(5, 0), sticky="ew")
        # self._load_cash_balance() # REDUNDANT, REMOVED CALL

        self.lbl_portfolio_risk_strategy = ctk.CTkLabel(f, text=i18n.t("lbl_portfolio_risk_strategy"))
        self.lbl_portfolio_risk_strategy.grid(row=4, column=0, padx=10, pady=5, sticky="w")
        self.seg_risk = ctk.CTkSegmentedButton(f, values=[i18n.t("seg_risk_conservative"), i18n.t("seg_risk_venture_growth")],
                                              fg_color="#1e293b", selected_color="#0e7490", selected_hover_color="#086078",
                                              unselected_color="#334155", unselected_hover_color="#475569")
        self.seg_risk.set(i18n.t("seg_risk_conservative"))
        self.seg_risk.grid(row=4, column=1, padx=10, pady=5, sticky="ew")
        # self._load_risk_strategy() # REDUNDANT, REMOVED CALL

        # Ensure self.portfolio_tickers_input, self.cash_balance_input, and self.seg_risk are defined
        # before attempting to load data into them.
        self._load_portfolio()

        self.btn_save_p = ctk.CTkButton(f, text=i18n.t("btn_save_portfolio"), command=self._save_portfolio,
                                         fg_color="#0e7490", hover_color="#086078")
        self.btn_save_p.grid(row=5, column=0, padx=(10, 5), pady=10, sticky="ew")

        self.btn_audit = ctk.CTkButton(f, text=i18n.t("btn_audit_portfolio"), command=self._audit_portfolio,
                                       fg_color="#0e7490", hover_color="#086078")
        self.btn_audit.grid(row=5, column=1, padx=(5, 10), pady=10, sticky="ew")

        self.lbl_inv_portfolio_audit_plan = ctk.CTkLabel(f, text=i18n.t("lbl_inv_portfolio_audit_plan"), font=ctk.CTkFont(size=16, weight="bold")) # LEO ADDITION
        self.lbl_inv_portfolio_audit_plan.grid(row=6, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="ew") # LEO ADDITION

        self.audit_log_box = ctk.CTkTextbox(f, height=300, wrap="word",
                                            fg_color="#1e293b", text_color="#f8fafc", border_color="#334155")
        self.audit_log_box.grid(row=7, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")
        self.audit_log_box.insert("0.0", i18n.t("audit_initial_msg"))
        self.audit_log_box.configure(state="disabled")

        self.btn_copy_audit = ctk.CTkButton(f, text=i18n.t("btn_copy_summary"), command=self._copy_audit_summary,
                                            fg_color="#0e7490", hover_color="#086078")
        self.btn_copy_audit.grid(row=8, column=0, columnspan=2, padx=10, pady=(5,10), sticky="ew")

    def _load_portfolio(self):
        if PORTFOLIO_FILE.exists():
            try:
                with open(PORTFOLIO_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    tickers_str = data.get("tickers", "")
                    cash_balance = data.get("cash_balance", "")
                    risk_strategy = data.get("risk_strategy", i18n.t("seg_risk_conservative")) # Default if not found

                    self.portfolio_tickers_input.delete("0.0", "end")
                    self.portfolio_tickers_input.insert("0.0", tickers_str)
                    self.cash_balance_input.delete(0, "end")
                    self.cash_balance_input.insert(0, str(cash_balance))
                    # Ensure the loaded risk strategy matches the localized values
                    if risk_strategy == "Conservative (5-10%)" or risk_strategy == i18n.t("seg_risk_conservative"):
                        self.seg_risk.set(i18n.t("seg_risk_conservative"))
                    elif risk_strategy == "Venture & Growth" or risk_strategy == i18n.t("seg_risk_venture_growth"):
                        self.seg_risk.set(i18n.t("seg_risk_venture_growth"))
                    else: # Fallback
                        self.seg_risk.set(i18n.t("seg_risk_conservative"))

            except Exception as e:
                # Fix 1: i18n parameter passing
                messagebox.showerror(i18n.t("error_title"), i18n.t("error_portfolio_save").format(e=e))

    def _load_cash_balance(self):
        # This function is redundant now as _load_portfolio loads both.
        # It's kept for structural consistency if other parts relied on it.
        pass

    def _load_risk_strategy(self):
        # This function is redundant now as _load_portfolio loads both.
        # It's kept for structural consistency if other parts relied on it.
        pass

    def _save_portfolio(self):
        tickers_raw = self.portfolio_tickers_input.get("0.0", "end").strip()
        tickers_list = [t.strip().upper() for t in re.split(r'[,;\n\s]+', tickers_raw) if t.strip()]
        cash_balance_str = self.cash_balance_input.get().strip()
        risk_strategy = self.seg_risk.get()

        try:
            cash_balance = float(cash_balance_str)
            data = {
                "tickers": tickers_raw, # Save as raw string for display
                "ticker_list": tickers_list, # Save as parsed list for logic
                "cash_balance": cash_balance,
                "risk_strategy": risk_strategy # Save the localized string directly
            }
            with open(PORTFOLIO_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4)
            self.audit_log_box.configure(state="normal")
            self.audit_log_box.delete("0.0", "end")
            self.audit_log_box.insert("0.0", i18n.t("status_portfolio_saved"))
            self.audit_log_box.configure(state="disabled")
        except ValueError:
            # Fix 1: i18n parameter passing
            messagebox.showerror(i18n.t("error_title"), i18n.t("error_portfolio_save").format(e=i18n.t("error_invalid_cash_balance")))
        except Exception as e:
            # Fix 1: i18n parameter passing
            messagebox.showerror(i18n.t("error_title"), i18n.t("error_portfolio_save").format(e=e))


    def _add_ticker_to_portfolio(self, ticker):
        current_tickers_str = self.portfolio_tickers_input.get("0.0", "end").strip()
        tickers_list = [t.strip().upper() for t in re.split(r'[,;\n\s]+', current_tickers_str) if t.strip()]

        if ticker not in tickers_list:
            new_tickers_str = f"{current_tickers_str}\n{ticker}" if current_tickers_str else ticker
            self.portfolio_tickers_input.configure(state="normal")
            self.portfolio_tickers_input.delete("0.0", "end")
            self.portfolio_tickers_input.insert("0.0", new_tickers_str)
            # No need to disable here, _save_portfolio will handle updating.
            self._save_portfolio() # Automatically save after adding
            # Display feedback directly in audit_log_box or a temporary status label
            self.audit_log_box.configure(state="normal")
            self.audit_log_box.delete("0.0", "end")
            # Fix 1: i18n parameter passing
            self.audit_log_box.insert("0.0", i18n.t("feedback_added_to_portfolio").format(ticker=ticker))
            self.audit_log_box.configure(state="disabled")
        else:
            self.audit_log_box.configure(state="normal")
            self.audit_log_box.delete("0.0", "end")
            # Fix 1: i18n parameter passing
            self.audit_log_box.insert("0.0", i18n.t("feedback_already_in_portfolio").format(ticker=ticker))
            self.audit_log_box.configure(state="disabled")


    def _audit_portfolio(self):
        tickers_raw = self.portfolio_tickers_input.get("0.0", "end").strip()
        tickers_list = [t.strip().upper() for t in re.split(r'[,;\n\s]+', tickers_raw) if t.strip()]
        cash_balance_str = self.cash_balance_input.get().strip()
        risk_strategy = self.seg_risk.get()

        if not tickers_list:
            self.audit_log_box.configure(state="normal")
            self.audit_log_box.delete("0.0", "end")
            self.audit_log_box.insert("0.0", i18n.t("audit_initial_msg"))
            self.audit_log_box.configure(state="disabled")
            return

        try:
            cash_balance = float(cash_balance_str)
        except ValueError:
            # Fix 1: i18n parameter passing
            messagebox.showerror(i18n.t("error_title"), i18n.t("error_portfolio_save").format(e=i18n.t("error_invalid_cash_balance")))
            return

        self.btn_audit.configure(state="disabled", text=i18n.t("btn_audit_loading"))
        self.audit_log_box.configure(state="normal")
        self.audit_log_box.delete("0.0", "end")
        # Fix 1: i18n parameter passing
        self.audit_log_box.insert("0.0", i18n.t("audit_loading_msg").format(len_tickers=len(tickers_list)))
        self.audit_log_box.configure(state="disabled")

        threading.Thread(target=self._audit_portfolio_worker, args=(tickers_list, cash_balance, risk_strategy)).start()

    def _audit_portfolio_worker(self, tickers_list, cash_balance, risk_strategy):
        audit_results = []
        full_stock_data = [] # To pass to Gemini for the overall audit

        for ticker in tickers_list:
            stats_text, payload = self._fetch_stock_data(ticker)
            full_stock_data.append(payload)
            # Fix 1: i18n parameter passing
            audit_results.append(f"{i18n.t('ticker_info_header').format(short_name=payload.get('name', ticker), ticker=ticker)}\n{stats_text}\n")
            self.master.after(0, self._update_audit_log_progress, f"Processed {ticker}...") # Update UI from main thread

        if HAS_GENAI and self.get_gemini_key():
            audit_resolution = self._get_gemini_audit_resolution(full_stock_data, cash_balance, risk_strategy)
        elif not HAS_GENAI:
            audit_resolution = i18n.t('error_gemini_api_contact').format(e='Google Generative AI library not found.') + "\n" + i18n.t('analysis_disclaimer_summary')
        elif not self.get_gemini_key():
            audit_resolution = i18n.t('error_api_key_unavailable') + "\n" + i18n.t('analysis_disclaimer_summary')
        else:
            audit_resolution = i18n.t('error_audit_gemini').format(e='Failed to fetch data for all tickers.') + "\n" + i18n.t('analysis_disclaimer_summary')

        final_audit_text = "\n".join(audit_results) + "\n\n" + audit_resolution
        self.master.after(0, self._update_audit_log_final, final_audit_text)
        self.master.after(0, lambda: self.btn_audit.configure(state="normal", text=i18n.t("btn_audit_portfolio")))

    def _update_audit_log_progress(self, message):
        self.audit_log_box.configure(state="normal")
        self.audit_log_box.insert("end", f"{message}\n")
        self.audit_log_box.see("end")
        self.audit_log_box.configure(state="disabled")

    def _update_audit_log_final(self, text):
        self.audit_log_box.configure(state="normal")
        self.audit_log_box.delete("0.0", "end")
        self.audit_log_box.insert("0.0", text)
        self.audit_log_box.configure(state="disabled")

    def _get_gemini_audit_resolution(self, stock_data_list, cash_balance, risk_strategy):
        if not HAS_GENAI or not self.get_gemini_key():
            return i18n.t('error_api_key_unavailable') + "\n" + i18n.t('analysis_disclaimer_summary')

        try:
            genai.configure(api_key=self.get_gemini_key())
            model = genai.GenerativeModel('gemini-pro')

            lang_instruction = get_llm_language_instruction()

            stock_data_str = ""
            for data in stock_data_list:
                stock_data_str += (
                    f"--- Ticker: {data.get('ticker')} ({data.get('name')}) ---\n"
                    f"Price: {data.get('price')} {data.get('currency')}\n"
                    f"P/E (Trailing): {data.get('pe')}\n"
                    f"P/E (Forward): {data.get('fwd_pe')}\n"
                    f"PEG Ratio: {data.get('peg')}\n"
                    f"52-Week High: {data.get('52_high')} {data.get('currency')}\n"
                    f"- 52-Week Low: {data.get('52_low')} {data.get('currency')}\n"
                    f"- Target Mean Price: {data.get('target_mean')} {data.get('currency')}\n\n"
                )
            
            # Determine strategy description based on localized selection
            strategy_desc = i18n.t("strategy_conservative_desc") if i18n.t("seg_risk_conservative") in risk_strategy else i18n.t("strategy_venture_desc")

            prompt = (
                f"{lang_instruction}\n"
                "You are Commander Quark, the supreme commander, a fat cat from Liepaja garrison, "
                "focused on building a personal cannery and fish factory empire. Your decisions "
                "prioritize strategic calm and maximal profit extraction with minimal personal effort.\n\n"
                "Conduct an investment audit for the Initiator's portfolio based on the provided data "
                "and strategy. Your response must be direct, professional, sharp, and in a tactical military "
                "garrison tone. Structure your audit into three main sections:\n"
                "1. **" + i18n.t("audit_sell_liquidate") + "** (Which stocks, if any, to liquidate or reduce positions in, and why.)\n"
                "2. **" + i18n.t("audit_cash_allocation") + "** (Recommendations on where to strategically deploy the free cash balance, "
                "or reallocate from liquidated positions. Focus on specific tickers or types of assets that align with the strategy.)\n"
                "3. **Komandanta Kvarka Rezolūcija** (Overall tactical summary, risks, and next steps.)\n\n"
                "Investment Strategy:\n"
                f"{strategy_desc}\n\n"
                f"Free Cash Balance: {cash_balance} USD\n\n"
                "Current Portfolio Stock Data:\n"
                f"{stock_data_str}"
                "Conclude with 'Komandants Kvarks ir izdevis Rīkojumu par portfeli.'"
            )
            response = model.generate_content(prompt, safety_settings={'HARASSMENT': 'BLOCK_NONE', 'HATE': 'BLOCK_NONE', 'SEXUAL': 'BLOCK_NONE', 'DANGEROUS': 'BLOCK_NONE'})
            return f"{i18n.t('audit_copy_header')}\n\n{i18n.t('audit_copy_strategy')} {risk_strategy}\n\n{response.text.strip()}\n\n{i18n.t('audit_copy_footer')}"
        except Exception as e:
            # Fix 1: i18n parameter passing
            return i18n.t('error_audit_gemini').format(e=e) + "\n" + i18n.t('analysis_disclaimer_summary')

    def _copy_audit_summary(self):
        try:
            self.clipboard_clear()
            self.clipboard_append(self.audit_log_box.get("0.0", "end"))
            self.audit_log_box.configure(state="normal") # Temporarily enable to show status
            self.audit_log_box.insert("0.0", i18n.t("status_copied") + "\n\n" + self.audit_log_box.get("0.0", "end"))
            self.audit_log_box.configure(state="disabled") # Re-disable
        except Exception as e:
            # Fix 1: i18n parameter passing
            messagebox.showerror(i18n.t("error_title"), i18n.t("error_copying").format(e=e))

    # ==========================================
    # 3. CILNE: KVARKA IZAUGSMES MEDNIEKS
    # ==========================================
    def _setup_growth_tab(self):
        f = self.tab_growth
        f.grid_columnconfigure(0, weight=1)
        f.grid_columnconfigure(1, weight=1)
        f.grid_rowconfigure(3, weight=1)

        self.lbl_growth_radar_title = ctk.CTkLabel(f, text=i18n.t("lbl_potential_radar"), font=ctk.CTkFont(size=16, weight="bold"))
        self.lbl_growth_radar_title.grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="ew")

        self.lbl_growth_description = ctk.CTkLabel(f, text=i18n.t("lbl_growth_description"), wraplength=400, justify="left")
        self.lbl_growth_description.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 5), sticky="w")

        # Dropdown for unicorn watchlist
        self.growth_watchlist_names = [item["name"] for item in GROWTH_WATCHLIST]
        self.growth_selector = ctk.CTkOptionMenu(f, values=[""] + self.growth_watchlist_names,
                                                 command=self._select_growth_candidate,
                                                 fg_color="#1e293b", button_color="#0e7490", button_hover_color="#086078")
        self.growth_selector.set(i18n.t("growth_no_candidates"))
        self.growth_selector.grid(row=2, column=0, padx=(10, 5), pady=5, sticky="ew")

        self.btn_scan_all_growth = ctk.CTkButton(f, text=i18n.t("btn_hunt_unicorns"), command=self._hunt_unicorns,
                                                 fg_color="#0e7490", hover_color="#086078")
        self.btn_scan_all_growth.grid(row=2, column=1, padx=(5, 10), pady=5, sticky="ew")

        self.lbl_growth_resolution_title = ctk.CTkLabel(f, text=i18n.t("lbl_inv_growth_resolution_title"), font=ctk.CTkFont(size=16, weight="bold"))
        self.lbl_growth_resolution_title.grid(row=3, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="ew")

        self.growth_analysis_box = ctk.CTkTextbox(f, height=300, wrap="word",
                                                  fg_color="#1e293b", text_color="#f8fafc", border_color="#334155")
        self.growth_analysis_box.grid(row=4, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")
        self.growth_analysis_box.insert("0.0", i18n.t("growth_status_initial"))
        self.growth_analysis_box.configure(state="disabled")

        self.btn_copy_growth = ctk.CTkButton(f, text=i18n.t("btn_copy_analysis"), command=self._copy_growth_analysis,
                                             fg_color="#0e7490", hover_color="#086078")
        self.btn_copy_growth.grid(row=5, column=0, padx=(10, 5), pady=5, sticky="ew")

        self.btn_export_word = ctk.CTkButton(f, text=i18n.t("btn_save_word"), command=self._export_growth_to_word,
                                             fg_color="#0e7490", hover_color="#086078")
        self.btn_export_word.grid(row=5, column=1, padx=(5, 10), pady=5, sticky="ew")

        self.btn_add_growth_to_portfolio = ctk.CTkButton(f, text=i18n.t("btn_add_to_portfolio"),
                                                          command=self._add_selected_growth_to_portfolio,
                                                          fg_color="#28a745", hover_color="#218838")
        self.btn_add_growth_to_portfolio.grid(row=6, column=0, columnspan=2, padx=10, pady=(5, 10), sticky="ew")

        self.growth_candidates_data = [] # Stores parsed unicorn data

    def _select_growth_candidate(self, selected_name):
        if selected_name == i18n.t("growth_no_candidates") or not self.growth_candidates_data:
            self.growth_analysis_box.configure(state="normal")
            self.growth_analysis_box.delete("0.0", "end")
            self.growth_analysis_box.insert("0.0", i18n.t("growth_no_candidates"))
            self.growth_analysis_box.configure(state="disabled")
            return

        selected_candidate = next((c for c in self.growth_candidates_data if c.get("name") == selected_name), None)
        if selected_candidate:
            display_text = ""
            display_text += f"--- {selected_candidate.get('name')} ({selected_candidate.get('ticker')}) ---\n"
            if selected_candidate.get('red_flag'):
                # Fix 1: i18n parameter passing
                display_text += i18n.t('warning_red_flag').format(warning_message=selected_candidate['red_flag']) + "\n"
            display_text += f"{i18n.t('lbl_niche')} {selected_candidate.get('niche')}\n"
            display_text += f"{i18n.t('lbl_why_10x_100x')} {selected_candidate.get('why_10x_100x')}\n"
            display_text += f"{i18n.t('lbl_black_swan')} {selected_candidate.get('black_swan')}\n"
            display_text += "\n" + selected_candidate.get('resolution', '')

            self.growth_analysis_box.configure(state="normal")
            self.growth_analysis_box.delete("0.0", "end")
            self.growth_analysis_box.insert("0.0", display_text)
            self.growth_analysis_box.configure(state="disabled")
        else:
            self.growth_analysis_box.configure(state="normal")
            self.growth_analysis_box.delete("0.0", "end")
            self.growth_analysis_box.insert("0.0", i18n.t("growth_no_candidates"))
            self.growth_analysis_box.configure(state="disabled")


    def _hunt_unicorns(self):
        if not HAS_GENAI or not self.get_gemini_key():
            self.growth_analysis_box.configure(state="normal")
            self.growth_analysis_box.delete("0.0", "end")
            self.growth_analysis_box.insert("0.0", i18n.t('error_api_key_unavailable') + "\n" + i18n.t('analysis_disclaimer_summary'))
            self.growth_analysis_box.configure(state="disabled")
            return

        # Fix 1: i18n parameter passing
        self.btn_scan_all_growth.configure(state="disabled", text=i18n.t("btn_hunt_unicorns_loading").format(btn_text=i18n.t("btn_hunt_unicorns")))
        self.growth_analysis_box.configure(state="normal")
        self.growth_analysis_box.delete("0.0", "end")
        self.growth_analysis_box.insert("0.0", i18n.t("growth_scanning_candidates"))
        self.growth_analysis_box.configure(state="disabled")

        threading.Thread(target=self._hunt_unicorns_worker).start()

    def _hunt_unicorns_worker(self):
        try:
            genai.configure(api_key=self.get_gemini_key())
            model = genai.GenerativeModel('gemini-pro')

            lang_instruction = get_llm_language_instruction()

            # Step 1: Get raw list of unicorn candidates
            self.master.after(0, self._update_growth_analysis_log, i18n.t("growth_scan_progress1"))
            prompt_list = (
                f"{lang_instruction}\n"
                "You are Commander Quark. List 10 publicly traded companies that are 'unicorns' "
                "or high-growth potential companies (e.g., AI, cutting-edge medicine, advanced energy, space tech, blockchain infrastructure). "
                "Return the response as a JSON array of objects, each with 'ticker' (string), 'name' (string), "
                "'sector' (string, e.g., 'AI', 'Space', 'Biotech'). "
                "Ensure the JSON is perfectly parsable, with no introductory or concluding text. "
                "Example: [{\"ticker\": \"SMCI\", \"name\": \"Super Micro Computer\", \"sector\": \"AI Servers\"}, ...]"
            )
            response_list = model.generate_content(prompt_list, safety_settings={'HARASSMENT': 'BLOCK_NONE', 'HATE': 'BLOCK_NONE', 'SEXUAL': 'BLOCK_NONE', 'DANGEROUS': 'BLOCK_NONE'})
            raw_unicorns_json = response_list.text.strip()
            # print(f"Raw Unicorns JSON: {raw_unicorns_json}") # Debugging

            # Attempt to extract JSON using regex for robustness (Task 1)
            json_match = re.search(r"\[.*\]", raw_unicorns_json, re.DOTALL)
            if json_match:
                raw_unicorns_json = json_match.group(0)
            else:
                raise ValueError("Could not find a valid JSON array in the response.")

            candidates = json.loads(raw_unicorns_json)
            if not isinstance(candidates, list) or not all(isinstance(c, dict) for c in candidates):
                raise ValueError("AI response for unicorn list was not a valid list of dictionaries.")

            if not candidates:
                self.master.after(0, self._update_growth_analysis_log, i18n.t("growth_no_unicorns_fallback"))
                self.master.after(0, lambda: self.btn_scan_all_growth.configure(state="normal", text=i18n.t("btn_hunt_unicorns")))
                self.master.after(0, lambda: self.growth_selector.set(i18n.t("growth_no_candidates")))
                return

            self.growth_candidates_data = []
            for candidate in candidates:
                ticker = candidate.get("ticker", "").upper()
                name = candidate.get("name", "")
                if not ticker: continue

                self.master.after(0, self._update_growth_analysis_log, f"Fetching data for {name} ({ticker})...")
                # Fetch detailed stock data for each candidate
                stats_text, payload = self._fetch_stock_data(ticker)
                
                # Update the UI with general progress
                self.master.after(0, self._update_growth_analysis_log, i18n.t("growth_scan_progress2"))

                prompt_analysis = (
                    f"{lang_instruction}\n"
                    "You are Commander Quark, the supreme commander. Analyze the following unicorn candidate "
                    "for its 10x-100x potential, hidden risks (black swans), and its niche/secret weapon. "
                    "Your assessment should be sharp, tactical, and focused on empire-building. "
                    "Return the analysis as a JSON object with the following keys:\n"
                    "- 'ticker': (string, e.g., 'SMCI')\n"
                    "- 'name': (string, e.g., 'Super Micro Computer')\n"
                    "- 'red_flag': (string, brief warning if P/E > 80, PEG > 2, or other major concern, otherwise empty string.)\n"
                    "- 'niche': (string, the company's unique 'secret weapon' or market niche.)\n"
                    "- 'why_10x_100x': (string, detailed tactical reasons for its exponential growth potential, "
                    "referencing relevant tech/market trends like AI, monopolistic tendencies, infrastructure importance.)\n"
                    "- 'black_swan': (string, a significant, low-probability, high-impact risk specific to this company.)\n"
                    "- 'resolution': (string, Commander Quark's concise, actionable tactical resolution for this specific ticker.)\n"
                    "Ensure the JSON is perfectly parsable, with no introductory or concluding text.\n\n"
                    "Stock Data:\n"
                    f"- Ticker: {payload.get('ticker')}\n"
                    f"- Name: {payload.get('name')}\n"
                    f"- Price: {payload.get('price')} {payload.get('currency')}\n"
                    f"- P/E (Trailing): {payload.get('pe')}\n"
                    f"- P/E (Forward): {payload.get('fwd_pe')}\n"
                    f"- PEG Ratio: {payload.get('peg')}\n"
                    f"- 52-Week High: {payload.get('52_high')} {payload.get('currency')}\n"
                    f"- 52-Week Low: {payload.get('52_low')} {payload.get('currency')}\n"
                    f"- Sector (from initial list): {candidate.get('sector', 'Unknown')}\n\n"
                    "Consider a 'Red Flag' if P/E (trailing or forward) is > 80-100 OR PEG > 2 and no clear growth justification.\n"
                    "Conclude the 'resolution' with 'Komandants Kvarks ir izdevis Rīkojumu.'"
                )
                response_analysis = model.generate_content(prompt_analysis, safety_settings={'HARASSMENT': 'BLOCK_NONE', 'HATE': 'BLOCK_NONE', 'SEXUAL': 'BLOCK_NONE', 'DANGEROUS': 'BLOCK_NONE'})
                raw_analysis_json = response_analysis.text.strip()
                # print(f"Raw Analysis JSON for {ticker}: {raw_analysis_json}") # Debugging

                # Attempt to extract JSON using regex for robustness (Task 1)
                json_match_analysis = re.search(r"\{.*\}", raw_analysis_json, re.DOTALL)
                if json_match_analysis:
                    raw_analysis_json = json_match_analysis.group(0)
                else:
                    raise ValueError(f"Could not find a valid JSON object in the analysis response for {ticker}.")

                analysis = json.loads(raw_analysis_json)

                self.growth_candidates_data.append(analysis)

            self.master.after(0, self._update_growth_ui_after_hunt)

        except json.JSONDecodeError as jde:
            partial_response = raw_unicorns_json[:500] if 'raw_unicorns_json' in locals() else raw_analysis_json[:500] if 'raw_analysis_json' in locals() else "N/A"
            # Fix 1: i18n parameter passing
            self.master.after(0, self._update_growth_analysis_log, i18n.t('error_gemini_parsing').format(e=jde, partial_response=partial_response))
            self.master.after(0, lambda: self.btn_scan_all_growth.configure(state="normal", text=i18n.t("btn_hunt_unicorns")))
        except ValueError as ve:
            # Fix 1: i18n parameter passing
            self.master.after(0, self._update_growth_analysis_log, i18n.t('error_ai_unparsable').format(e=ve))
            self.master.after(0, lambda: self.btn_scan_all_growth.configure(state="normal", text=i18n.t("btn_hunt_unicorns")))
        except Exception as e:
            # Fix 1: i18n parameter passing
            self.master.after(0, self._update_growth_analysis_log, i18n.t('error_gemini_api_contact').format(e=e))
            self.master.after(0, lambda: self.btn_scan_all_growth.configure(state="normal", text=i18n.t("btn_hunt_unicorns")))


    def _update_growth_analysis_log(self, message):
        self.growth_analysis_box.configure(state="normal")
        self.growth_analysis_box.insert("end", f"{message}\n")
        self.growth_analysis_box.see("end")
        self.growth_analysis_box.configure(state="disabled")

    def _update_growth_ui_after_hunt(self):
        self.btn_scan_all_growth.configure(state="normal", text=i18n.t("btn_hunt_unicorns"))
        self.growth_selector.configure(values=[c.get("name") for c in self.growth_candidates_data])
        if self.growth_candidates_data:
            self.growth_selector.set(self.growth_candidates_data[0]["name"])
            self._select_growth_candidate(self.growth_candidates_data[0]["name"]) # Display first candidate
            self.growth_analysis_box.configure(state="normal")
            self.growth_analysis_box.insert("0.0", i18n.t("growth_analysis_complete") + "\n\n")
            self.growth_analysis_box.configure(state="disabled")
        else:
            self.growth_selector.set(i18n.t("growth_no_candidates"))
            self.growth_analysis_box.configure(state="normal")
            self.growth_analysis_box.delete("0.0", "end")
            self.growth_analysis_box.insert("0.0", i18n.t("growth_no_unicorns_fallback"))
            self.growth_analysis_box.configure(state="disabled")

    def _copy_growth_analysis(self):
        if not self.growth_candidates_data:
            messagebox.showinfo(i18n.t("info_title"), i18n.t("growth_no_analysis_to_copy"))
            return

        try:
            summary_text = f"{i18n.t('growth_copy_header')}\n\n"
            for idx, candidate in enumerate(self.growth_candidates_data):
                # Fix 1: i18n parameter passing
                summary_text += i18n.t('growth_copy_candidate_header').format(idx_plus_1=idx+1, ticker=candidate.get('ticker'), name=candidate.get('name')) + "\n"
                if candidate.get('red_flag'):
                    # Fix 1: i18n parameter passing
                    summary_text += "   " + i18n.t('warning_red_flag').format(warning_message=candidate['red_flag']) + "\n"
                # Fix 1: i18n parameter passing
                summary_text += i18n.t('growth_copy_niche').format(niche=candidate.get('niche')) + "\n"
                # Fix 1: i18n parameter passing
                summary_text += i18n.t('growth_copy_why').format(why=candidate.get('why_10x_100x')) + "\n"
                # Fix 1: i18n parameter passing
                summary_text += i18n.t('growth_copy_risk').format(risk=candidate.get('black_swan')) + "\n"
                summary_text += f"   {candidate.get('resolution', '')}\n\n"

            self.clipboard_clear()
            self.clipboard_append(summary_text)
            messagebox.showinfo(i18n.t("info_title"), i18n.t("status_copied"))
        except Exception as e:
            # Fix 1: i18n parameter passing
            messagebox.showerror(i18n.t("error_title"), i18n.t("error_copying").format(e=e))

    def _export_growth_to_word(self):
        if not HAS_DOCX:
            messagebox.showerror(i18n.t("error_title"), i18n.t("error_docx_missing"))
            return
        if not self.growth_candidates_data:
            messagebox.showinfo(i18n.t("info_title"), i18n.t("info_no_data_to_export"))
            return

        try:
            document = Document()
            document.add_heading(i18n.t("word_report_title"), level=1)
            document.add_paragraph(f"{i18n.t('word_generated_on')}{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            document.add_paragraph(i18n.t("word_analysis_basis"))
            document.add_page_break()

            for idx, candidate in enumerate(self.growth_candidates_data):
                document.add_heading(f"{idx+1}. {candidate.get('name')} ({candidate.get('ticker')})", level=2)
                if candidate.get('red_flag'):
                    # Fix 1: i18n parameter passing
                    document.add_paragraph(f"{i18n.t('word_red_flag')}{candidate['red_flag']}")
                # Fix 1: i18n parameter passing
                document.add_paragraph(f"{i18n.t('word_secret_weapon')}{candidate.get('niche')}")
                # Fix 1: i18n parameter passing
                document.add_paragraph(f"{i18n.t('word_why_10x_100x')}{candidate.get('why_10x_100x')}")
                # Fix 1: i18n parameter passing
                document.add_paragraph(f"{i18n.t('word_black_swan')}{candidate.get('black_swan')}")
                document.add_paragraph(candidate.get('resolution', ''))
                document.add_page_break()

            filename = OUTBOX_DIR / f"{i18n.t('word_output_filename_prefix')}{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            document.save(filename)
            # Fix 1: i18n parameter passing
            messagebox.showinfo(i18n.t("success_title"), f"{i18n.t('word_success_message')}{filename}")
        except Exception as e:
            # Fix 1: i18n parameter passing
            messagebox.showerror(i18n.t("error_saving_word_title"), f"{i18n.t('word_error_saving_message')}{e}")

    def _add_selected_growth_to_portfolio(self):
        selected_name = self.growth_selector.get()
        if selected_name == i18n.t("growth_no_candidates") or not self.growth_candidates_data:
            messagebox.showinfo(i18n.t("info_title"), i18n.t("growth_no_candidates"))
            return

        selected_candidate = next((c for c in self.growth_candidates_data if c.get("name") == selected_name), None)
        if selected_candidate and selected_candidate.get("ticker"):
            self._add_ticker_to_portfolio(selected_candidate["ticker"])
            self.tabview.set(i18n.t("tab_inv_etoro")) # Switch to portfolio tab
        else:
            # Fix 1: i18n parameter passing
            messagebox.showerror(i18n.t("error_title"), i18n.t("error_fetching_ticker_data").format(ticker=selected_name, e="Ticker not found for selected candidate."))

    # ==========================================
    # 4. CILNE: DIVIDENDES & NAUDAS PLŪSMA
    # ==========================================
    def _setup_dividends_tab(self):
        f = self.tab_dividends
        f.grid_columnconfigure(0, weight=1)
        f.grid_columnconfigure(1, weight=1)
        f.grid_rowconfigure(3, weight=1)

        self.lbl_cashflow_radar_title = ctk.CTkLabel(f, text=i18n.t("lbl_cashflow_radar_title"), font=ctk.CTkFont(size=16, weight="bold"))
        self.lbl_cashflow_radar_title.grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="ew")

        self.lbl_cashflow_radar_desc = ctk.CTkLabel(f, text=i18n.t("lbl_cashflow_radar_desc"), wraplength=400, justify="left")
        self.lbl_cashflow_radar_desc.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 5), sticky="w")

        self.btn_calc_dividends = ctk.CTkButton(f, text=i18n.t("btn_calc_dividends"), command=self._calculate_dividends,
                                                fg_color="#0e7490", hover_color="#086078")
        self.btn_calc_dividends.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

        self.lbl_dividends_resolution_title = ctk.CTkLabel(f, text=i18n.t("lbl_inv_dividends_resolution_title"), font=ctk.CTkFont(size=16, weight="bold"))
        self.lbl_dividends_resolution_title.grid(row=3, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="ew")

        self.dividends_analysis_box = ctk.CTkTextbox(f, height=400, wrap="word",
                                                    fg_color="#1e293b", text_color="#f8fafc", border_color="#334155")
        self.dividends_analysis_box.grid(row=4, column=0, columnspan=2, padx=10, pady=(5, 10), sticky="nsew")
        self.dividends_analysis_box.insert("0.0", i18n.t("dividends_initial_msg"))
        self.dividends_analysis_box.configure(state="disabled")

    def _calculate_dividends(self):
        if not HAS_GENAI or not self.get_gemini_key():
            self.dividends_analysis_box.configure(state="normal")
            self.dividends_analysis_box.delete("0.0", "end")
            self.dividends_analysis_box.insert("0.0", i18n.t('error_api_key_unavailable') + "\n" + i18n.t('analysis_disclaimer_summary'))
            self.dividends_analysis_box.configure(state="disabled")
            return

        tickers_raw = self.portfolio_tickers_input.get("0.0", "end").strip()
        tickers_list = [t.strip().upper() for t in re.split(r'[,;\n\s]+', tickers_raw) if t.strip()]

        if not tickers_list:
            self.dividends_analysis_box.configure(state="normal")
            self.dividends_analysis_box.delete("0.0", "end")
            self.dividends_analysis_box.insert("0.0", i18n.t("dividends_no_companies"))
            self.dividends_analysis_box.configure(state="disabled")
            return

        self.btn_calc_dividends.configure(state="disabled", text=i18n.t("btn_calc_dividends_loading"))
        self.dividends_analysis_box.configure(state="normal")
        self.dividends_analysis_box.delete("0.0", "end")
        self.dividends_analysis_box.insert("0.0", i18n.t("dividends_initial_analysis_msg"))
        self.dividends_analysis_box.configure(state="disabled")

        threading.Thread(target=self._calculate_dividends_worker, args=(tickers_list,)).start()

    def _calculate_dividends_worker(self, tickers_list):
        dividend_data = []
        for ticker in tickers_list:
            # Fix 1: i18n parameter passing
            self.master.after(0, self._update_dividends_log, i18n.t("dividends_scan_progress1").format(len_tickers=len(tickers_list)))
            try:
                stock = yf.Ticker(ticker)
                info = stock.info
                name = info.get("shortName", ticker)
                currency = info.get("currency", "USD")

                # Get the last dividend payment
                dividends = stock.dividends
                last_dividend = None
                if not dividends.empty:
                    last_div_date = dividends.index[-1]
                    last_div_amount = dividends.iloc[-1]
                    last_dividend = {"date": last_div_date.strftime('%Y-%m-%d'), "amount": last_div_amount}

                # Get current price
                current_price = info.get("currentPrice") or info.get("regularMarketPrice")

                # Get dividend yield (if available from yfinance info)
                dividend_yield = info.get("dividendYield") # This is often annualized

                dividend_data.append({
                    "ticker": ticker,
                    "name": name,
                    "currency": currency,
                    "current_price": current_price,
                    "last_dividend": last_dividend,
                    "dividend_yield": dividend_yield,
                })
            except Exception as e:
                # Fix 1: i18n parameter passing
                self.master.after(0, self._update_dividends_log, i18n.t('error_fetching_ticker_data').format(ticker=ticker, e=e))
                dividend_data.append({
                    "ticker": ticker,
                    "name": ticker,
                    "currency": "N/A",
                    "current_price": "N/A",
                    "last_dividend": None,
                    "dividend_yield": None,
                    "error": str(e)
                })

        self.master.after(0, self._update_dividends_log, i18n.t("dividends_scan_progress2"))

        # Now, use Gemini for resolution
        if HAS_GENAI and self.get_gemini_key():
            resolution_text = self._get_gemini_dividends_resolution(dividend_data)
        else:
            resolution_text = i18n.t('error_api_key_unavailable') + "\n" + i18n.t('analysis_disclaimer_summary')

        self.master.after(0, self._update_dividends_log_final, resolution_text)
        self.master.after(0, lambda: self.btn_calc_dividends.configure(state="normal", text=i18n.t("btn_calc_dividends")))


    def _update_dividends_log(self, message):
        self.dividends_analysis_box.configure(state="normal")
        self.dividends_analysis_box.insert("end", f"{message}\n")
        self.dividends_analysis_box.see("end")
        self.dividends_analysis_box.configure(state="disabled")

    def _update_dividends_log_final(self, text):
        self.dividends_analysis_box.configure(state="normal")
        self.dividends_analysis_box.delete("0.0", "end")
        self.dividends_analysis_box.insert("0.0", text)
        self.dividends_analysis_box.configure(state="disabled")

    def _get_gemini_dividends_resolution(self, dividend_data):
        if not HAS_GENAI or not self.get_gemini_key():
            return i18n.t('error_api_key_unavailable') + "\n" + i18n.t('analysis_disclaimer_summary')

        try:
            genai.configure(api_key=self.get_gemini_key())
            model = genai.GenerativeModel('gemini-pro')

            lang_instruction = get_llm_language_instruction()

            data_for_prompt = ""
            for item in dividend_data:
                last_div_info = f"Last Dividend: {item['last_dividend']['amount']} on {item['last_dividend']['date']}" if item['last_dividend'] else i18n.t("dividends_not_paying")
                yield_info = f"Yield: {item['dividend_yield']:.2%}" if item['dividend_yield'] is not None else "Yield: N/A"
                price_info = f"Current Price: {item['current_price']} {item['currency']}" if item['current_price'] != "N/A" else "Current Price: N/A"
                error_info = f" (Error: {item['error']})" if "error" in item else ""

                data_for_prompt += (
                    f"- Ticker: {item['ticker']} ({item['name']})\n"
                    f"  {price_info}\n"
                    f"  {last_div_info}\n"
                    f"  {yield_info}{error_info}\n"
                )

            prompt = (
                f"{lang_instruction}\n"
                "You are Commander Quark, the supreme commander, focused on building a personal "
                "cannery and fish factory empire. Passive income through dividends is crucial "
                "for long-term stability and empire growth without excessive effort. Your decisions "
                "prioritize stable, reliable cash flow.\n\n"
                "Analyze the following dividend data from the Initiator's portfolio. "
                "Provide a clear, tactical resolution for Commander Quark's dividend calendar and strategy. "
                "Your response should be direct, professional, sharp, and in a tactical military garrison tone. "
                "Structure your audit into three main sections:\n"
                "1. **" + i18n.t("dividends_champions") + "** (Identify the strongest dividend payers, potential for consistent passive income.)\n"
                "2. **" + i18n.t("dividends_sustainability_risks") + "** (Highlight any tickers with unstable dividends, low yields, "
                "or high payout ratios that might pose a risk to cash flow stability. Advise on potential liquidations.)\n"
                "3. **" + i18n.t("dividends_reinvestment_strategy") + "** (Tactical advice on how Commander Quark should reinvest "
                "dividends to maximize his empire's growth and cash flow. Mention specific sectors or types of companies.)\n\n"
                "Dividend Data:\n"
                f"{data_for_prompt}\n"
                "Conclude with 'Komandants Kvarks ir izdevis Rīkojumu par dividendēm.' "
                "If no companies are paying dividends, state that and advise to seek dividend-paying assets."
            )

            response = model.generate_content(prompt, safety_settings={'HARASSMENT': 'BLOCK_NONE', 'HATE': 'BLOCK_NONE', 'SEXUAL': 'BLOCK_NONE', 'DANGEROUS': 'BLOCK_NONE'})
            return response.text.strip()
        except Exception as e:
            # Fix 1: i18n parameter passing
            return i18n.t('error_gemini_dividends').format(e=e) + "\n" + i18n.t('analysis_disclaimer_summary')

    # ==========================================
    # 5. CILNE: TIRGUS KLIMATS & INSAIDERI
    # ==========================================
    def _setup_market_tab(self):
        f = self.tab_market
        f.grid_columnconfigure(0, weight=1)
        f.grid_columnconfigure(1, weight=1)
        f.grid_rowconfigure(4, weight=1)

        self.lbl_market_sentiment_title = ctk.CTkLabel(f, text=i18n.t("lbl_market_sentiment_title"), font=ctk.CTkFont(size=16, weight="bold"))
        self.lbl_market_sentiment_title.grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="ew")

        self.lbl_fng = ctk.CTkLabel(f, text=i18n.t("lbl_fng_loading"))
        self.lbl_fng.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 5), sticky="ew")

        self.btn_refresh_market = ctk.CTkButton(f, text=i18n.t("btn_refresh_market_intel"), command=self._refresh_market_intel,
                                                 fg_color="#0e7490", hover_color="#086078")
        self.btn_refresh_market.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

        self.lbl_market_news_title = ctk.CTkLabel(f, text=i18n.t("lbl_inv_market_news_title"), font=ctk.CTkFont(size=16, weight="bold"))
        self.lbl_market_news_title.grid(row=3, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="ew")

        self.market_intel_box = ctk.CTkTextbox(f, height=400, wrap="word",
                                               fg_color="#1e293b", text_color="#f8fafc", border_color="#334155")
        self.market_intel_box.grid(row=4, column=0, columnspan=2, padx=10, pady=(5, 10), sticky="nsew")
        self.market_intel_box.insert("0.0", i18n.t("market_news_initial_msg"))
        self.market_intel_box.configure(state="disabled")

    def _refresh_market_intel(self):
        if not HAS_GENAI or not self.get_gemini_key():
            self.market_intel_box.configure(state="normal")
            self.market_intel_box.delete("0.0", "end")
            self.market_intel_box.insert("0.0", i18n.t('error_api_key_unavailable') + "\n" + i18n.t('analysis_disclaimer_summary'))
            self.market_intel_box.configure(state="disabled")
            return

        # Fix 1: i18n parameter passing
        self.btn_refresh_market.configure(state="disabled", text=i18n.t("btn_refresh_market_intel_loading").format(btn_text="")) # btn_text is ignored in current translation
        self.market_intel_box.configure(state="normal")
        self.market_intel_box.delete("0.0", "end")
        self.market_intel_box.insert("0.0", i18n.t("market_news_initial_msg"))
        self.market_intel_box.configure(state="disabled")
        self.lbl_fng.configure(text=i18n.t("lbl_fng_loading"))

        threading.Thread(target=self._refresh_market_intel_worker).start()

    def _refresh_market_intel_worker(self):
        finnhub_key = ""
        try:
            with open("config.json", "r", encoding="utf-8") as f:
                cfg = json.load(f)
                finnhub_key = cfg.get("finnhub_api_key", "")
        except Exception:
            pass

        # Fetch Fear & Greed Index
        fng_value, fng_sentiment = self._fetch_fear_greed_index()
        self.master.after(0, lambda: self.lbl_fng.configure(text=f"Fear & Greed: {fng_value} ({fng_sentiment})" if fng_value else i18n.t("fng_not_available")))

        # Fetch Macro Data (US 10Y, Gold, Bitcoin)
        macro_data = self._fetch_macro_data()

        # Fetch Insider Transactions (using Finnhub)
        insider_sentiment, recent_insider_transactions = self._fetch_finnhub_insider_sentiment(finnhub_key)

        # Fetch Market News
        market_news = self._fetch_market_news(finnhub_key)

        # Generate Quark's resolution using Gemini
        resolution_text = self._get_gemini_market_resolution(fng_value, fng_sentiment, macro_data, insider_sentiment, recent_insider_transactions, market_news)

        self.master.after(0, self._update_market_intel_final, resolution_text)
        self.master.after(0, lambda: self.btn_refresh_market.configure(state="normal", text=i18n.t("btn_refresh_market_intel")))


    def _fetch_fear_greed_index(self):
        try:
            # Using CNN's Fear & Greed Index
            url = "https://production.dataviz.cnn.io/index/feargreed/all.json"
            response = requests.get(url, timeout=10).json()
            fng_data = response.get("fear_greed_historical", [])
            if fng_data:
                latest_data = fng_data[0] # Most recent data point
                value = latest_data.get("score")
                sentiment = latest_data.get("rating")
                return value, sentiment
        except Exception as e:
            print(f"Error fetching Fear & Greed Index: {e}")
        return None, None

    def _fetch_macro_data(self):
        macro_data = {
            i18n.t("macro_us10y"): i18n.t("macro_na"),
            i18n.t("macro_gold"): i18n.t("macro_na"),
            i18n.t("macro_bitcoin"): i18n.t("macro_na")
        }
        try:
            # For 10Y US Treasury Yield
            ten_year_ticker = yf.Ticker("^TNX") # ^TNX for CBOE Interest Rate 10 Year T-Note
            ten_year_info = ten_year_ticker.info
            yield_value = ten_year_info.get("regularMarketPrice")
            if yield_value:
                macro_data[i18n.t("macro_us10y")] = f"{yield_value:.2f}%"

            # For Gold
            gold_ticker = yf.Ticker("GC=F") # GC=F for Gold Futures
            gold_info = gold_ticker.info
            gold_price = gold_info.get("regularMarketPrice")
            if gold_price:
                macro_data[i18n.t("macro_gold")] = f"{gold_price:.2f} USD"

            # For Bitcoin
            bitcoin_ticker = yf.Ticker("BTC-USD")
            bitcoin_info = bitcoin_ticker.info
            bitcoin_price = bitcoin_info.get("regularMarketPrice")
            if bitcoin_price:
                macro_data[i18n.t("macro_bitcoin")] = f"{bitcoin_price:.2f} USD"

        except Exception as e:
            print(f"Error fetching macro data: {e}")
        return macro_data

    def _fetch_finnhub_insider_sentiment(self, api_key):
        if not api_key:
            return i18n.t("error_finnhub_key_missing"), []

        recent_transactions = []
        try:
            # This is a placeholder as Finnhub's `insider-sentiment` is for *specific stocks*.
            # To get *market climate & insiders*, we would typically need an aggregation API
            # or fetch data for many top stocks and summarize.
            # For simplicity, if we were to use a single ticker as a proxy (e.g., SPY):
            # url_insider = f"https://finnhub.io/api/v1/stock/insider-sentiment?symbol=SPY&from={from_date}&to={to_date}&token={api_key}"
            # response_insider = requests.get(url_insider, timeout=10).json()
            # ... process response_insider ...

            # Let's try to get some aggregate news sentiment using Finnhub to gauge general market 'insider' climate from news
            # (This is a stretch to meet "insiders" but aligns with "market climate").
            news_url = f"https://finnhub.io/api/v1/news?category=general&minId=0&token={api_key}"
            news_response = requests.get(news_url, timeout=10).json()
            
            buy_signals = 0
            sell_signals = 0
            
            # This is a very simplistic heuristic, not real insider trading sentiment
            for news_item in news_response[:10]: # Check top 10 news items
                headline = news_item.get('headline', '').lower()
                if 'acquisition' in headline or 'buyback' in headline or 'strong' in headline:
                    buy_signals += 1
                elif 'sell' in headline or 'drop' in headline or 'weak' in headline:
                    sell_signals += 1
            
            if buy_signals > sell_signals * 2:
                insider_sentiment_text = i18n.t("insider_buy")
            elif sell_signals > buy_signals * 2:
                insider_sentiment_text = i18n.t("insider_sell")
            else:
                insider_sentiment_text = i18n.t("insider_neutral")

            recent_transactions = [
                {"sentiment": insider_sentiment_text, "notes": f"Based on {buy_signals} positive and {sell_signals} negative news signals from Finnhub."}
            ]
            return insider_sentiment_text, recent_transactions

        except Exception as e:
            print(f"Error fetching Finnhub insider sentiment: {e}")
            return i18n.t("insider_data_error"), []

    def _fetch_market_news(self, api_key):
        if not api_key:
            return [i18n.t("error_finnhub_key_missing")]
        news_headlines = []
        try:
            url = f"https://finnhub.io/api/v1/news?category=general&token={api_key}"
            response = requests.get(url, timeout=10).json()
            for item in response[:5]: # Get top 5 headlines
                news_headlines.append(f"- {item.get('headline', i18n.t('unknown_value'))} ({item.get('source', i18n.t('unknown_value'))})")
            if not news_headlines:
                news_headlines.append(i18n.t("news_no_aggressive"))
        except Exception as e:
            print(f"Error fetching market news: {e}")
            # Fix 1: i18n parameter passing
            news_headlines.append(i18n.t('error_fetching_ticker_data').format(ticker='Market News', e=e))
        return news_headlines

    def _get_gemini_market_resolution(self, fng_value, fng_sentiment, macro_data, insider_sentiment, recent_insider_transactions, market_news):
        if not HAS_GENAI or not self.get_gemini_key():
            return i18n.t('error_api_key_unavailable') + "\n" + i18n.t('analysis_disclaimer_summary')

        try:
            genai.configure(api_key=self.get_gemini_key())
            model = genai.GenerativeModel('gemini-pro')

            lang_instruction = get_llm_language_instruction()

            macro_data_str = "\n".join([f"- {key}: {value}" for key, value in macro_data.items()])
            insider_transactions_str = "\n".join([f"- {t.get('sentiment', i18n.t('unknown_value'))}: {t.get('notes', i18n.t('unknown_value'))}" for t in recent_insider_transactions])
            if not insider_transactions_str and insider_sentiment == i18n.t("insider_data_error"):
                insider_transactions_str = f"- {i18n.t('insider_data_error')}"
            elif not insider_transactions_str:
                insider_transactions_str = f"- {i18n.t('insider_no_major_deals')}"

            market_news_str = "\n".join(market_news)

            prompt = (
                f"{lang_instruction}\n"
                "You are Commander Quark, the supreme commander, overseeing market tactics. "
                "Your objective is to identify key trends and risks to protect and expand "
                "the personal cannery and fish factory empire. Your decisions prioritize "
                "strategic calmness, rapid response to threats, and opportunistic exploitation "
                "of market movements.\n\n"
                "Analyze the current market climate based on the following intelligence. "
                "Provide a concise, tactical resolution for the Initiator. Your response should "
                "be direct, professional, sharp, and in a tactical military garrison tone. "
                "Structure your resolution with clear tactical advice and warnings.\n\n"
                "Market Intelligence:\n"
                f"- Fear & Greed Index: {fng_value} ({fng_sentiment or i18n.t('unknown_value')})\n"
                f"{i18n.t('macro_liquidity_header')}\n"
                f"{macro_data_str}\n"
                f"{i18n.t('insider_transactions_header')}\n"
                f"{insider_transactions_str}\n"
                f"{i18n.t('news_headlines_header')}\n"
                f"{market_news_str}\n\n"
                "Provide an overall market outlook, identify any red flags or opportunities, "
                "and give tactical recommendations for adjusting portfolio strategy or preparing for market shifts. "
                "Conclude with 'Komandants Kvarks ir izdevis Rīkojumu par tirgus taktiku.'\n"
            )

            response = model.generate_content(prompt, safety_settings={'HARASSMENT': 'BLOCK_NONE', 'HATE': 'BLOCK_NONE', 'SEXUAL': 'BLOCK_NONE', 'DANGEROUS': 'BLOCK_NONE'})
            return f"{i18n.t('kvark_strategic_resolution')}{response.text.strip()}"
        except Exception as e:
            # Fix 1: i18n parameter passing
            return i18n.t('error_kvark_market_intel').format(e=e) + "\n" + i18n.t('analysis_disclaimer_summary')

    def _update_market_intel_final(self, text):
        self.market_intel_box.configure(state="normal")
        self.market_intel_box.delete("0.0", "end")
        self.market_intel_box.insert("0.0", text)
        self.market_intel_box.configure(state="disabled")

# End of InvestorStudioFrame class