# Vánoční program 2026 – prezentace

Projekt generuje PowerPointovou prezentaci s programem na Vánoční svátky
(23. 12. 2026 – 1. 1. 2027) pro rodinu s dětmi.

## Soubory

- `build.js` – generátor prezentace (pptxgenjs). Veškerý obsah i vzhled se mění zde.
- `Vanocni_program_2026.pptx` – výstup. Needitovat ručně, vždy přegenerovat z `build.js`.

## Sestavení

```bash
node build.js
```

Skript zapíše `.pptx` a následně na něj aplikuje barevné téma (`applyTheme`).
Balíčky `pptxgenjs`, `react`, `react-dom`, `react-icons` a `sharp` jsou předinstalované – nic neinstaluj.

## Kontrola po každé změně

```bash
S=<adresář pptx skillu>
python $S/scripts/office/validate.py Vanocni_program_2026.pptx
python $S/scripts/office/soffice.py --headless --convert-to pdf Vanocni_program_2026.pptx
rm -f slide-*.jpg && pdftoppm -jpeg -r 60 Vanocni_program_2026.pdf slide
```

Prohlédni vyrenderované snímky: přetékající text, překryvy, okraje.

## Konvence

- **Jazyk:** veškerý text česky, s diakritikou a českými uvozovkami („…“).
- **Data:** formát `24. 12.` (s mezerami). Dny v týdnu musí sedět na rok 2026
  (24. 12. = čtvrtek, 31. 12. = čtvrtek, 1. 1. 2027 = pátek).
- **Plátno:** `LAYOUT_WIDE` (13,33" × 7,5"), okraje min. 0,5".
- **Písma:** nadpisy Cambria, text Calibri. Tělo textu min. 14 pt, nadpisy 36 pt.
- **Barvy (téma „Vanoce“):** tmavě zelená `1B4332` (dominantní), vánoční červená `B3192B`,
  zlatá `D4A017` / `E9C46A`, světlá `F1F5F2`. V kódu používej `pres.SchemeColor`, ne hex.
- **Layouty:** `Titul` (tmavý, úvod a závěr) a `Obsah` (světlý, nadpis vlevo + podtitul kurzívou).
  Patička a číslo snímku jsou na layoutu, ne na snímcích.
- **Motiv:** ikony z `react-icons/fa` v barevném kruhu (`circleIcon`), karty se zaoblenými rohy (`card`).
- **Nepoužívat:** dekorativní pruhy, linky pod nadpisy, emoji, `#` v hex barvách.

## Struktura prezentace

1. Titulní snímek
2. Svátky v kostce (časová osa)
3. Přípravy do 23. 12.
4. Štědrý den – harmonogram a zvyky
5. Boží hod a Štěpán
6. Mezi svátky (27.–30. 12.)
7. Silvestr a Nový rok
8. Kontrolní seznam
9. Závěr – Veselé Vánoce
