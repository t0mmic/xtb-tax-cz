# xtb-tax-cz

Nasazení: https://xtb-tax-czech.vercel.app/

# XTB Fiscal Report Builder - pro průběžný přehled zisku/ztrát během roku 
přepočtený kurzem ČNB. 
Netestováno se všemi typy investičních instrumentů (Nemám vše k dispozici 
ale lze upravit)

Lokální webová aplikace pro zpracování exportů XTB a vyplnění šablony
`data/Fiscal Report.xlsx`.

## Spuštění

1. Nainstalujte Node.js 18 nebo novější.
2. Ve složce projektu spusťte `npm install`.
3. Spusťte `npm start`.
4. Otevřete <http://127.0.0.1:3000> (nebo port nastavený v proměnné `PORT`).
5. Vyberte uložené exporty nebo nahrajte XLSX přímo ve webovém rozhraní.
   Nahrané soubory se zpracují v paměti, nejvýše 8 souborů / 20 MB na soubor /
   64 MB celkem, a na disk se neukládají.
6. Volitelně zadejte EODHD klíč v sekci „EODHD – pro dohledání ISIN“ a vytvořte
   report. Klíč zůstane pouze v paměti serveru do restartu. Registrace na eohd 
   je zdarma

Pokud je nastaven EODHD klíč, odešle se EODHD tento klíč a symboly instrumentů, 
pro které zatím není ISIN v lokální cache.
Zdrojové XLSX ani vygenerovaný report se EODHD neposílají.

## Vstupy a kurzy

- Zdrojové soubory musí být XTB XLSX exporty s listy `Closed Positions`,
  `Cash Operations` a `Open Positions`. Měna účtu se určuje z prefixu názvu
  souboru (např. `USD_...xlsx`).
- Denní kurzy ČNB se čtou z `data/cnb-rates.json`. Pro víkend nebo svátek se
  použije poslední dostupný kurz před datem transakce.
- Roční kurz MF je potřeba zadat ručně ve formuláři „Kurzy“. Aplikace jej
  neodvozuje z denních kurzů. Stejným formulářem lze doplnit chybějící denní
  kurz ČNB nebo novou měnu.
- Kurz je vyjádřen v CZK za jednu jednotku cizí měny.

## Výpočty a omezení

- Uzavřené obchody se zařadí podle data uzavření; otevřené obchody se přenesou
  ze zdrojového listu bez ohledu na datum otevření.
- Horní souhrn odděluje účty vedené v CZK od účtů v cizí měně; oba sloupce
  jsou vyjádřeny v CZK.
- U akcií a frakčních práv se hodnoty v měně instrumentu přepočítají kurzem
  nákupu a prodeje zvlášť; k výsledku se přičte exportovaný poplatek přepočtený
  kurzem účtu.
- Zlomkové nákupy akcií a ETF se sčítají podle tickeru a účtu v celé dostupné
  historii, včetně otevřených pozic. Dokud jejich součet nedosáhne jednoho kusu,
  jsou vedeny jako Frakční práva; po dosažení celého kusu se mezi akcie/ETF
  přeřadí i dřívější zlomkové transakce daného tickeru a účtu.
- U CFD se použije exportovaný `Profit/Loss`, který již zahrnuje náklady
  uvedené v XTB exportu. Související řádky `Close trade` a již zahrnuté swapy
  se proto nepřičítají podruhé.
- Dividendy používají měnu a částku na akcii z komentáře XTB (např.
  `USD 0.75/ SHR`). Počet akcií se rekonstruuje z otevřených a uzavřených pozic
  podle ID pozice, je-li k dispozici, jinak podle tickeru, k datu výplaty nebo
  až 45 dní před ní; použitý počet a případný časový odstup jsou uvedeny v
  popisu kurzu. Pokud pozici nelze dohledat, počet kusů se
  odhadne z částky v účtu a denního kurzu ČNB a report jej označí jako odhad.
  Srážková daň s komentářem `WHT` se počítá z odpovídající dividendy.
- Dividendy, srážková daň a poplatky ke konkrétnímu instrumentu používají měnu
  podle tickeru a zvolený kurz k datu transakce; u ročního MF se odvozená
  původní částka přepočte ročním kurzem. Ostatní peněžní operace používají
  měnu účtu. Vklady a převody jsou uvedeny v úpravách a nevstupují do souhrnu
  zisku a ztráty. Sekce „Profit Loss Summary with adjustments“ uvádí tyto
  úpravy také v souhrnu v měně účtu a v přepočtu do CZK.
- Modře jsou označené použité měny a kurzy; zeleně částky po přepočtu do CZK.
- XTB export nedodává ISIN ani spolehlivý daňový domicil emitenta. ISIN lze
  doplňovat v lokálním souboru `data/isin-cache.json` jako mapu tickerů na ISIN
  (např. `"AAPL.US": "US0378331005"`). Je-li nastaven `EODHD_API_KEY`,
  chybějící ISIN se dohledají přes EODHD Search API (s ID Mapping jako záložním
  zdrojem) a uloží do této cache; bez klíče se použijí jen již uložené hodnoty.
  Pro ticker bez záznamu se
  zachovává dosavadní hodnota `US`; před daňovým použitím ji zkontrolujte.

Výstup je předvyplnění šablony, nikoli daňové stanovisko. Zkontrolujte
zařazení a výsledné částky před použitím.
