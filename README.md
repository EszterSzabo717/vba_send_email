# vba_send_email
macro that sends emails according to client status

# Energetikai Projekt Email Értesítő Makró

## Leírás
Ez az Excel `.xlsm` fájl egy **automatizált ügyfél értesítő rendszert** tartalmaz VBA segítségével, kifejezetten nagy adatbázisok kezelésére (akár 8–10 ezer sor).  

A makró főbb funkciói:  
- Tömbös feldolgozás → gyors, nagy mennyiségű adat kezelés  
- Email értesítések küldése Microsoft Outlook Desktop alkalmazáson keresztül  
- Hibakezelés, duplikáció ellenőrzés, már értesített ügyfelek adatainak megőrzése  
- Minden kiküldés **azonnali logolása szöveges fájlba**, a makró leállása esetén is megmarad a futás előzménye  
- Excel oszlopok frissítése: **Státusz / Eredmény / Küldés dátuma**  
- Felugró összesítés a futás végén: sikeres és hibás küldések, futás ideje, hivatkozás a log fájlra

---

## Tesztadatok
A mellékelt `.xlsm` fájl tartalmaz **dummy ügyfél listát**:  
| Név           | Email                   | Státusz        |  
|---------------|------------------------|----------------|  
| Teszt User1   | test1@example.com       | értesítendő    |  
| Teszt User2   | test2@example.com       | értesítendő    |  

- Amennyiben cím **formailag érvényes**, a makró „sikeresnek” könyveli  
- **Fontos:** a makró **nem ellenőrzi, hogy a cím ténylegesen létezik-e**. Tesztcímek (`example.com`) esetén az Outlook / levelezőszerver adhat visszapattanást (NDR) később  

---

## Használat – GitHub / Saját gép
1. Töltsd le a `.xlsm` fájlt a GitHub-ról a kívánt mappába  
2. Nyisd meg Excelben  
3. Engedélyezd a makrókat (ha le vannak tiltva)  
4. Nyomd meg az **„értesítések elküldése**, vagy futtasd a `Ertesites_Batch_Log_File4` makrót a VBA editorból  
5. A makró:  
   - Küldi az emaileket Outlookon keresztül  
   - Excelben frissíti az értesítendő ügyfelek státuszát és a küldés eredményét  
   - A **log fájl** (`Ertesites_Log.txt`) **automatikusan létrejön abban a mappában, ahol az `.xlsm` található**  
   - A **végén felugró ablakban** megjelenik a sikeres és hibás küldések száma, futás ideje és a log fájl elérési útja  

---

## Log fájl formátum
A `Ertesites_Log.txt` **áttekinthető, soronkénti logolással**, például:
2,Teszt User1,test1@example.com
,értesítve,sikeres,2026-03-04 10:15:25
3,Teszt User2,test2@example.com
,értesítve,sikeres,2026-03-04 10:15:26
---- ÖSSZESÍTÉS ----
Futás ideje: 2026-03-04 10:30:25
Sikeres: 2, Hibás: 0

- **Első 6 oszlop** → sor, név, email, státusz, eredmény, küldés dátuma  
- **Összesítő blokk a végén** → futás ideje, sikeres és hibás küldések száma  

---

## Portfólió megjegyzések
- A makró **biztonságos**, a már „értesítve” státuszú sorokat nem módosítja  
- Hibás címek, duplikált emailek, Outlook hibák **részletesen logolódnak**  
- A felugró ablak segít a felhasználónak gyors áttekintést kapni a futásról  
- A szöveges log biztosítja, hogy **még a makró leállása esetén is nyoma marad minden sor feldolgozásának**  

---

## Követelmények
- Microsoft Excel `.xlsm` formátummal (Office 365 vagy Classic Excel)  
- Microsoft Outlook Desktop telepítve és bejelentkezve  
- Makrók engedélyezve az Excel biztonsági beállításaiban  



