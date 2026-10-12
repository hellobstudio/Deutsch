# Szógyakorló – német szókincs gyakorlása Excel-táblázatból

Android-app, amely egy Excel-fájlból veszi a szavakat, és ezekből készít feladatokat:

- **Kvíz**: válaszd ki a szó jelentését (németről magyarra és magyarról németre).
- **Névelő**: der, die vagy das? (csak főneveknél)
- **Lyukas szöveg, listából**: a példamondatba legördülő listából választod ki a hiányzó szót.
- **Lyukas szöveg, beírással**: beírod a hiányzó szót (az ä, ö, ü, ß betűkhöz külön gombok vannak).
- **Mondatrendezés**: az összekevert szavakból rakod össze a mondatot, a magyar fordítás segít.
- **Vegyes**: az előzőek összekeverve.

Az app megjegyzi, melyik szót hányszor rontottad el, és ha bekapcsolod, ezeket gyakrabban kérdezi. Egy kör végén a hibás szavakat azonnal újra is gyakorolhatod. Témakörre szűrni is lehet.

Az app internet nélkül működik. Egy beépített, 48 szavas mintaszókinccsel indul, ezt cserélheted le a saját Excel-fájlodra.

## Az Excel-fájl felépítése

A `nemet_szokincs_sablon.xlsx` (ugyanez a fájl: `www/szokincs.xlsx`) első munkalapja a **Szavak**:

| Oszlop | Kötelező? | Példa |
|---|---|---|
| Német | igen | Kaffee |
| Névelő | főnévnél | der |
| Többes szám | nem | Kaffees |
| Magyar | igen | kávé |
| Szófaj | nem | főnév |
| Téma | nem | Étel és ital |
| Példamondat (német) | nem | Ich trinke jeden Morgen einen \*Kaffee\*. |
| Mondat magyarul | nem | Minden reggel megiszom egy kávét. |

A példamondatban **csillagok közé** tedd a gyakorolt szót. Ez lesz a lyukas szövegben a hiányzó rész, ezért ragozott alak is lehet: `Ich *verstehe* die Frage nicht.`
Példamondat nélkül a szó csak a kvízben és a névelős feladatban szerepel. A részletes leírás a fájl **Útmutató** munkalapján van.

## APK készítése GitHubon (telepítés nélkül, a böngészőből)

1. **Hozz létre egy új repositoryt** a GitHubon (jobb felül **+** → **New repository**), pl. `szogyakorlo` néven. Lehet privát is.
2. **Töltsd fel a fájlokat.** Csomagold ki a zip-et, majd a repository oldalán: **Add file** → **Upload files**, és húzd be a kicsomagolt `szogyakorlo` mappa *tartalmát* (nem magát a mappát). Végül **Commit changes**.
   - Figyelj a `.github` mappára! A ponttal kezdődő mappák rejtettek lehetnek: Windowson *Nézet → Rejtett elemek*, macOS-en a Finderben `Cmd + Shift + .`.
   - Ha a `.github` mappát nem sikerül feltölteni: **Add file** → **Create new file**, a fájl neve legyen `.github/workflows/build-apk.yml`, és másold bele a fájl tartalmát.
3. **Várd meg az építést.** Az **Actions** fülön elindul az „Android APK építése” folyamat (kb. 5–10 perc). Ha nem indul el magától, kattints rá, majd **Run workflow**.
4. **Töltsd le az APK-t.** A zöld pipával végzett futás oldalán lent, az **Artifacts** résznél van a `szogyakorlo-apk`. Ez egy zip, benne a `szogyakorlo.apk`.
5. **Telepítsd a telefonra.** Másold át az APK-t, és nyisd meg. Az Android ilyenkor engedélyt kér az „ismeretlen forrásból” történő telepítéshez.

Minden további feltöltés vagy módosítás után a GitHub automatikusan új APK-t készít.

### Kiadás (Release) letölthető APK-val

Ha azt szeretnéd, hogy az APK állandó linken legyen elérhető: a repository oldalán **Releases** → **Draft a new release** → *Choose a tag* mezőbe írd be pl. `v1.0.0` → **Create new tag** → **Publish release**. Pár perc múlva az építés a kiadáshoz csatolja a `szogyakorlo.apk` fájlt.

## Saját szókincs betöltése a telefonon

1. Töltsd ki a sablont Excelben vagy Google Táblázatokban, és mentsd `.xlsx` formátumban.
2. Tedd a fájlt a telefonra (Letöltések mappa, Google Drive stb.).
3. Az appban: **Szókincs** → **Excel-fájl betöltése**, és válaszd ki a fájlt.

Ha valamelyik sorral gond van (pl. hiányzik a jelentés, vagy a példamondatban nincs megjelölve a szó), az app kiírja, melyik sorról van szó. A **Mintaszókincs** gombbal bármikor visszatérhetsz a beépített listához.

**Ha a beépített szókincset is cserélnéd**, cseréld le a repositoryban a `www/szokincs.xlsx` fájlt (ugyanezen a néven), és a GitHub új APK-t készít.

## Jó tudni

- Az APK egy állandó, a repositoryban lévő teszt-aláírókulccsal (`keystore/debug.keystore`) készül, így az új verziók a régi fölé telepíthetők, és megmaradnak a betöltött szavak és az eredmények. Ez a kulcs saját használatra és osztályon belüli terjesztésre jó; Google Play-áruházba feltöltéshez saját, titkos kulcs kellene.
- Az app ikonja egyelőre a Capacitor alapértelmezett ikonja.
- A program felépítése: a `www` mappában van maga az app (`index.html`, `style.css`, `app.js`), a [Capacitor](https://capacitorjs.com) csomagolja Android-appá, az Excel-fájlokat a [SheetJS](https://sheetjs.com) olvassa be.

## Kipróbálás számítógépen (nem kötelező)

Ha van a gépeden [Node.js](https://nodejs.org) (20-as vagy újabb verzió):

```bash
npm install
npm run serve
```

Utána nyisd meg böngészőben: http://localhost:8080
