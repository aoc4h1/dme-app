# Doktori Minimum Ellenőrző (DME) App

A **Doktori Minimum Ellenőrző** egy R Shiny alapú webes alkalmazás, amely az MTMT (Magyar Tudományos Művek Tára) nyilvános API-jára építve automatikusan összesíti és kategorizálja a jelöltek publikációs és hivatkozási teljesítményét. 

A kalkulációk a **Magyar Tudományos Akadémia (MTA) IX. Osztályának Szociológiai Tudományos Bizottsága** (2019-es szabályzat) doktori minimumkövetelményeinek megfelelően készülnek.

---

## Főbb funkciók

- **Automatikus MTMT profil- és adatlekérés:** Elég megadni a jelölt 8 jegyű MTMT Szerzői ID-ját, az alkalmazás letölti az összes publikációt és a hozzájuk kapcsolódó hivatkozásokat.
- **Folyóirat-kategorizálás:** Összeveti a megjelent cikkeket a Bizottság hivatalos folyóiratlistájával (`SzocTB_20230601.xlsx`), elkülönítve a hazai és nemzetközi listás lapokat.
- **Önhivatkozások automatikus szűrése:** Az `authorships` adatok alapján kiszűri és külön kezeli a saját szerzői hivatkozásokat.
- **Intelligens pontozási logika:** Figyelembe veszi a társszerzők számát, a könyvek oldalszámát (ívszámítás), a publikáció nyelvét, valamint azt, hogy az adott közlemény vagy hivatkozás a PhD fokozat megszerzése előtt vagy után keletkezett.
- **Interaktív, védett Excel export:** Egy képletekkel, legördülő listákkal és feltételes formázásokkal (hibák piros kiemelése) ellátott, professzionális `.xlsx` munkafüzetet generál, ahol a jelölt manuálisan kiegészítheti az MTMT-ből hiányzó adatokat (pl. ívszámok).

---

## A projekt felépítése

dme-app/
│
├── app.R                  # A teljes Shiny UI és Server logika
├── dme-app.Rproj          # RStudio projekt fájl
│
├── SzocTB_20230601.xlsx   # Folyóirat-adatbázis (Hazai és Nemzetközi lapok)
├── mtmt_output.xlsx       # Excel kimeneti template (formázások, struktúra)
├── dme_utmutato.docx      # Letölthető felhasználói útmutató
│
└── README.md              # Ez a leírás

## Jogi nyilatkozat (Disclaimer)
A kalkulátor **nem hivatalos program**. A pontszámok ellenőrzése és a hiányzó adatok pontos kitöltése a jelölt saját felelőssége. A program esetleges számítási hibáiért az alkotók semmilyen felelősséget nem vállalnak.
