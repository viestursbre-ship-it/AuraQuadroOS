```markdown

# AQ-OS Sistēmas Dziļā Specifikācija: Bruno Vīzija (Versija 2.5)



## 1. Sistēmas Arhitektūra: AQ-OS Dzinējs



AQ-OS (AuraQuadro Operating System) ir inteliģenta dokumentu apstrādes sistēma, kas paredzēta datu ieguvei, analīzei un strukturēšanai no dažādiem dokumentu formātiem. Tās arhitektūra ir veidota modulāra un izturīga, nodrošinot augstu efektivitāti un uzticamību.



**Galvenie komponenti:**



* **FolderWatcherService:** Sistēmas acis un ausis. Tā nepārtraukti uzrauga ienākošo dokumentu mapi (`INBOX_DIR`), gaidot jaunus failus, kas prasa apstrādi. Darbojas atsevišķā pavedienā, lai nodrošinātu nepārtrauktu monitoringu bez galvenā dzinēja bloķēšanas.

* **DocDigestService:** Sistēmas smadzenes un rokas. Šis ir centrālais modulis, kas pārvalda visu dokumenta dzīves ciklu:

* Failu lasīšana no dažādiem formātiem (`.docx`, `.txt`, `.pdf`, `.pptx`).

* Satura analīze, izmantojot **Gemini AI moduli**.

* Kopsavilkuma dokumenta ģenerēšana (`.docx`).

* Datu reģistrēšana Excel reģistrā.

* Apstrādāto failu droša arhivēšana.

* **Gemini AI modulis:** Mūsu inteliģence, ko nodrošina Google Gemini ģeneratīvā AI. Tas veic padziļinātu dokumentu satura analīzi, pielāgojoties dokumenta tipam (juridiskie līgumi vai akadēmiskas prezentācijas).

* **Excel Reģistrs:** Mūsu grāmatvedis, kas precīzi reģistrē visus apstrādātos dokumentus, to galveno informāciju un statusu, nodrošinot caurspīdību un pārskatāmību.

* **Failu Sistēmas Integrācija:** Sistēma aktīvi mijiedarbojas ar lokālo failu sistēmu, izmantojot dedikētas mapes ienākošajiem, izejošajiem un arhivētajiem dokumentiem.



## 2. Dzinēja Mapju Struktūra (Storage)



Visi AQ-OS dati tiek centralizēti glabāti `storage` mapē, nodrošinot skaidru un organizētu failu plūsmu.



```

storage/

├── inbox/ # Ienākošie dokumenti, gaida apstrādi.

├── outbox/ # Apstrādāto dokumentu kopsavilkumi (DIGEST_*.docx).

├── archive/ # Arhivētie, apstrādātie oriģinālie dokumenti.

└── Ligumu_Registrs.xlsx # Centralizēts Excel reģistrs visiem dokumentiem.

```



Šī struktūra ir kā *perfekti organizēta virtuves atvilktne* – katrai lietai sava vieta, lai darbs ritētu gludi.



## 3. Gemini AI Analīzes Modulis (`_analyze_with_gemini`)



Šis ir AQ-OS sirds, kur notiek *maģija*. Gemini AI modulis ir atbildīgs par dokumenta satura izpratni un strukturēšanu.



* **API atslēga:** Darbībai nepieciešama `AQ_AI_API_KEY` vides mainīgā iestatīšana. Bez tās AI klusēs kā *izslēgts radio*.

* **Dinamiskā Promptu Izvēle:** Modulis ir *gudrs* un pielāgo savu pieeju atkarībā no dokumenta tipa:

* **Juridiskajiem dokumentiem (.docx, .txt, .pdf):** Tiek izmantots detalizēts prompts, lai izvilktu būtisku informāciju par līguma pusēm, finanšu noteikumiem, īpašumtiesībām, termiņiem, riskiem un citiem juridiski svarīgiem aspektiem, strukturējot to stingrā JSON formātā.

* **Prezentācijām (.pptx):** Tiek aktivizēts īpašs prompts, lai izveidotu *akadēmisku kopsavilkumu*. AI fokusējas uz galvenajām tēzēm, jēdzieniem, tēmu kopsavilkumiem un prezentācijas mērķi, ignorējot juridiskās "blēņas".

* **Izejas formāts:** AI vienmēr atgriež tīru JSON objektu, kas ir viegli parsējams un apstrādājams tālāk. Nav nekādu lieku *čupu* vai *garlaicīgu Markdown bloku*!

* **Modelis un temperatūra:** Tiek izmantots `gemini-2.5-flash` modelis ar zemu temperatūru (`0.2`), nodrošinot precīzas un konsekventas atbildes, nevis *radošas fantāzijas*.



## 4. Drošā Arhīva Loģika (`safe_archive_file`)



Mūsu arhīvs ir kā *labi nostiprināts vīna pagrabs* – nekas nepazūd un nekas netiek nejauši aizstāts!



* **Mērķis:** Novērst failu pārrakstīšanas kļūdas un dublikātu problēmas, kad apstrādātais dokuments tiek pārvietots uz `archive` mapi.

* **Darbība:**

1. Tiek izveidots galamērķa arhīva direktorijs, ja tas vēl nepastāv.

2. Pirms faila pārvietošanas, sistēma pārbauda, vai arhīva mapē jau nav fails ar tādu pašu nosaukumu.

3. **Laika zīmogu pievienošana:** Ja dublikāts tiek atrasts, oriģinālajam faila nosaukumam tiek pievienots unikāls laika zīmogs (piemēram, `dokumenta_nosaukums_20231027_143501.docx`). Tas garantē, ka katra faila versija tiek saglabāta, un nekad nav divu vienādu nosaukumu.

4. Fails tiek droši pārvietots, izmantojot `shutil.move`, kas ir efektīvāka par kopēšanu un dzēšanu.

* **Ieguvumi:** Nodrošina failu integritāti, novērš datu zudumu un sistēmas apstāšanos neparedzētu failu nosaukumu konfliktu dēļ. *Perfetto!*



## 5. Excel Reģistra Integrācija (`_append_to_excel_registry`)



Lai neviena informācija netiktu pazaudēta, mēs uzturam precīzu Excel reģistru:



* **Inicializācija:** Ja reģistra fails (`Ligumu_Registrs.xlsx`) nepastāv, tas tiek automātiski izveidots ar standarta galvenēm un treknraksta formatējumu.

* **Datu Robustums:** Funkcija ir izstrādāta, lai izturētu nestandarta vai trūkstošus laukus AI atbildē. Katrs lauks tiek pārbaudīts, un, ja informācija nav pieejama, tiek izmantots "N/A", lai nodrošinātu, ka ieraksts nekad netiek izlaists. Datu tipi tiek droši konvertēti uz virknēm.

* **Automātiska Kolonnu Platums:** Pēc katra ieraksta Excel kolonnu platums tiek automātiski pielāgots, lai nodrošinātu optimālu lasāmību.

* **Logošana:** Katrs veiksmīgs ieraksts tiek reģistrēts sistēmas žurnālā.



## 6. DocX Kopsavilkumu Ģenerēšana (`_create_digest_docx`)



Pēc AI analīzes rezultāti tiek apkopoti glītā, cilvēkiem lasāmā `.docx` formātā, kas ir kā *izcili pagatavots ēdiens* – viegli patērējams un bagāts ar informāciju.



* **Strukturēti virsraksti:** Kopsavilkums tiek organizēts ar hierarhiskiem virsrakstiem (1. un 2. līmeņa), lai atvieglotu navigāciju un informācijas meklēšanu.

* **Sadaļas:** Automātiski tiek veidotas sadaļas "Vispārīgā Informācija", "Puses un Lomas", "Darījuma Priekšmets", "Finanšu Noteikumi", "Īpašumtiesības un Līguma Izbeigšana" un "Riski un Brīdinājumi".

* **Datu atainošana:** AI iegūtie dati tiek dinamiski ievietoti atbilstošajās sadaļās, attiecīgi apstrādājot vārdnīcas, sarakstus un atsevišķas vērtības. Īpaša uzmanība pievērsta partiju vārdu un lomu formatēšanai, kā arī akadēmisko tēžu izvadei no PPTX analīzes.

* **Viegla Pārskatāmība:** Katra būtiskā detaļa tiek atainota punktos, padarot kopsavilkumu ātri pārskatāmu un saprotamu.



AQ-OS ir izveidots, lai strādātu *ātri, precīzi un bez kļūdām*, atbrīvojot cilvēkus no garlaicīgā dokumentu darba. Tas ir *dzīvesveids*, nevis tikai programma!

``` 