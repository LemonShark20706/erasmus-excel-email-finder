# Erasmus Excel Email Finder

Erasmus Excel Email Finder egy Python alapú konzolos eszköz, ami a `sources` mappában lévő Erasmus/jellegű Excel fájlokból kinyeri a szervezeti rekordokat, e-mailt keres hozzájuk, majd fájlonként külön output Excel fájlt készít a `done` mappába.

## Fő funkciók

- Változó szerkezetű Excel fájlok feldolgozása (nem fix oszlopszámra épít).
- Több munkalap kezelése ugyanazon fájlon belül.
- Releváns sorok felismerése projektazonosító minták alapján.
- E-mail keresés webes találatokból (DDGS + oldalforrás ellenőrzés).
- Erősített e-mail validáció és országkód alapú domain végződés-szűrés.
- Interaktív konzol UI (`questionary`) menüvel.
- Automatikus csomag- és mappaellenőrzés induláskor.

## Mappa struktúra

- `sources/`: ide kerülnek a feldolgozandó `.xlsx` fájlok.
- `done/`: ide menti a program az elkészült outputokat.

Ha a mappák nem léteznek, a program létrehozza őket.

## Futtatás

```bash
python main.py
```

## Konzol menü

A program induláskor menüt ad 4 opcióval:

1. `Project Config`

- Ellenőrzi és telepíti a szükséges csomagokat.
- Ellenőrzi/létrehozza a `sources` és `done` mappákat.
- A végén Enterre lép vissza a menübe.

2. `Show README`
- Kiírja a `README.md` tartalmát a konzolra.
- Enterre lép vissza a menübe.

3. `Start Processing`

- Rákérdez: készen áll-e a feldolgozás.
- Jóváhagyás után elindítja a fájlok feldolgozását.

4. `Exit`

- Kilép a programból.

## Feldolgozási logika röviden

1. Beolvassa a `sources` mappa összes `.xlsx` fájlját.
2. Minden munkalapon végigmegy, és a projektazonosítót tartalmazó sorokat keresi.
3. A közelben lévő fejlécek és heurisztikák alapján kinyeri:

- `ProjectNumber`
- `OrgName`
- `OrgCity` (címből is próbálja kinyerni, ha szükséges)

4. Deduplikál projekt/szervezet/város alapján.
5. Szervezetenként e-mailt keres.
6. Fájlonként output Excel-t készít.

## E-mail validáció és szűrés

A rendszer több szinten szűri az e-mail címeket:

- Formátum és alap validáció (`@`, domain szerkezet, TLD ellenőrzés).
- Kizárt végződések: pl. `png`, `jpg`, `svg`, `pdf`, `css`, `js`, stb.
- Placeholder domainek kizárása: pl. `mysite`, `example`, `domain`, `localhost`.
- Gyenge/robot címek kizárása: pl. `noreply`, `support`, `abuse`, stb.
- Országkód kompatibilitás:
- A projektazonosítóból olvas országkódot (pl. `HU`, `SK`, `SI`).
- Elfogadott TLD: országkód + `com` + `net`.
- Szlovén kódra (`SI`) `si` és `sl` végződés is elfogadott.

## Pontozás

A találatokat pontozza (`score_email`) és a jobb jelölteket preferálja.
Például az `info`, `office`, `contact`, `school` típusú local-part plusz pontot ad.

## Output formátum

Minden input fájlhoz külön output készül a `done` mappába:

- fájlnév minta: `<input_fajlnev>_emails_output.xlsx`

Output oszlopok:

- `ProjectNumber`
- `OrgName`
- `Email`
- `OrgCity`
- `Verified`

Fontos:

- Ha az e-mail **nem verified**, akkor az `Email` mező üresen marad.
- A forrásfájl és sheet neve nem kerül bele az outputba.

## Progress megjelenítés

A feldolgozás során:

- Minden fájl indulásakor törli a konzolt.
- Az első sorban fájlszintű progress jelenik meg:
- progress bar + százalék + `X/Y` + aktuális fájlnév.
- Alatta rekordszintű élő progress fut (`Found`, `Verified` számlálókkal).

Ez VS Code terminálban is olvashatóbb, mint a soronkénti teljes logolás.

## Szükséges csomagok

A program automatikusan telepíti, ha hiányoznak:

- `pandas`
- `requests`
- `xlsxwriter`
- `ddgs`
- `questionary`
- `openpyxl`

## Belső felépítés (osztályok)

- `StartupFolderValidator`: csomag- és mappaellenőrzés.
- `ExcelStructureDetector`: változó Excel-szerkezet felismerése és rekordkinyerés.
- `EmailFinder`: e-mail keresés, validálás, pontozás.
- `OutputExcelWriter`: output Excel mentés.
- `SourseFolderProcessor`: teljes feldolgozási folyamat.
- `ConsoleMenuApp`: interaktív menükezelés.

## Gyors használat

1. Tedd az `.xlsx` fájlokat a `sources` mappába.
2. Futtasd: `python help.py`
3. Menüben:

- `Project Config`
- `Start Processing`

4. Kész outputok a `done` mappában.
