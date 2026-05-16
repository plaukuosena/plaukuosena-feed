# Plaukuosena.lt Kaina24 feed generator

Šis paketas skirtas automatizuoti `plaukuosena.lt` produktų feed'ą Kaina24 platformai.

## Ką daro sistema

1. Nuskaito `https://plaukuosena.lt/sitemap.xml`
2. Atrenka tikėtinus produktų URL
3. Nuskaito produktų puslapius
4. Paima `Product` structured data / JSON-LD, jei ji yra
5. Jei reikia, naudoja meta/HTML fallback
6. Sugeneruoja:
   - `kaina24.xml`
   - `products_snapshot.csv`
   - `products_snapshot.xlsx`

## Kodėl taip

Hostinger Website Builder neturi patogaus produktų eksporto, todėl sistema apeina ribojimą naudodama viešą svetainę ir sitemap.

Kaina24 pardavėjų puslapyje nurodo, kad prekių informaciją reikia pateikti XML formatu. Jei naudojamas Google Merchant ar Meta feed'as, jis dažnai taip pat gali tikti, bet šiame pakete ruošiame atskirą Kaina24 XML.

## Paleidimas kompiuteryje

```bash
pip install -r requirements.txt
python generate_feed.py
```

Po paleidimo atsiras:

```text
kaina24.xml
products_snapshot.csv
products_snapshot.xlsx
```

## Automatinis paleidimas per GitHub Actions

1. Sukurti privatų arba viešą GitHub repository.
2. Įkelti visus šio aplanko failus.
3. Įjungti GitHub Actions.
4. Workflow `.github/workflows/update-feed.yml` automatiškai paleis generatorių kasdien.
5. Kaina24 galima pateikti viešą `kaina24.xml` nuorodą.

## Vieša XML nuoroda

Paprasčiausias variantas:

- naudoti GitHub Pages
- arba laikyti `kaina24.xml` Hostinger faile
- arba paprašyti programuotojo, kad XML būtų prieinamas pvz. `https://plaukuosena.lt/kaina24.xml`

## Svarbu dėl nuolaidų kodų

Kaina24 matys tik kainą, kuri viešai rodoma produkto puslapyje ir feed'e.

Jeigu nuolaida yra tik kupono kodas, pvz. `-10% su kodu`, Kaina24 jos nematys kaip realios produkto kainos.

Kad Kaina24 matytų akciją, akcinė kaina turi būti rodoma pačiame produkto puslapyje.

## Ką reikės patikrinti po pirmo paleidimo

- Ar visi produktai pateko į feed'ą
- Ar kainos teisingos
- Ar nuotraukos atsidaro
- Ar SKU nesidubliuoja
- Ar Kaina24 priima XML struktūrą

## Jei trūksta produktų

Taisyk `config.json` lauką:

```json
"product_hint_words": []
```

Įrašyk papildomus žodžius, kurie pasikartoja produktų URL.

## SKU logika

SKU kuriamas pagal principą:

```text
BRAND-MODELIS-TIPAS-TALPA
```

Pvz.:

```text
WEL-US-SHA-250
OLX-N4-SHA-250
DAV-OI-OIL-135
```

Brand ir tipo kodus galima koreguoti `config.json` faile.
