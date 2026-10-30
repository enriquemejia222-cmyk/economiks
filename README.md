# Economiks

Riesgo país, crecimiento, inflación, desempleo, deuda y dólares de Latinoamérica, con Estados Unidos como referencia.
Datos de J.P. Morgan (vía los bancos centrales de República Dominicana y Perú), el FMI y el Banco Mundial.

## Cómo funciona
- `index.html` es la página. Lee todo de `data.json`.
- `update.py` baja todas las fuentes y regenera `data.json`. También va sumando a `historial_riesgo.json` el riesgo país de cada día.
- `.github/workflows/update.yml` corre el script de lunes a viernes, revisa los datos con `check.py` y publica la página en GitHub Pages. Si los datos se ven mal, no publica y la página sigue con los del día anterior.

## Ponerlo en línea (una sola vez)
1. Crea un repositorio público llamado `economiks` en GitHub y sube estos archivos a la rama `main`.
2. En el repositorio: **Settings → Pages → Build and deployment → Source: GitHub Actions**.
3. **Actions → Actualizar datos y publicar → Run workflow** para la primera actualización.
La página queda en `https://TU-USUARIO.github.io/economiks/`. Para un dominio propio (por ejemplo economiks.do), ponlo en Settings → Pages → Custom domain.

## Correrlo en tu computadora
```
pip install pymupdf openpyxl numpy      # y poppler (pdftotext)
python3 -I update.py --out . --hist historial_riesgo.json
python3 -m http.server                   # abre http://localhost:8000
```
