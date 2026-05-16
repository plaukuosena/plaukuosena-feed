# Plaukuosena.lt Kaina24 feed generator

Šitas ZIP paruoštas taip, kad failus reikia kelti tiesiai į GitHub repo pagrindą, ne į papildomą aplanką.

Repo pagrindiniame lygyje turi matytis:

- `.github/workflows/update-feed.yml`
- `generate_feed.py`
- `config.json`
- `requirements.txt`
- `README.md`
- `KODEE_UZDUOTIS.txt`

## Kaip paleisti

1. GitHub repo turi būti Public.
2. Įkelk VISĄ šio ZIP turinį į repo pagrindą.
3. Eik į Settings → Pages.
4. Source pasirink: GitHub Actions.
5. Eik į Actions → Update and publish Kaina24 feed → Run workflow.
6. Kai workflow taps žalias, XML bus pasiekiamas adresu:

`https://<github-vartotojas>.github.io/<repo-pavadinimas>/kaina24.xml`

Pvz.:

`https://plaukuosena.github.io/plaukuosena-feed/kaina24.xml`

## Ką daro workflow

- Kasdien paleidžia `generate_feed.py`
- Sugeneruoja `kaina24.xml`
- Sugeneruoja `products_snapshot.csv`
- Sugeneruoja `products_snapshot.xlsx`
- Publikuoja `kaina24.xml` per GitHub Pages
