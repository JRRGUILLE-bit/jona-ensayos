# Jona — disponibilidad de ensayos

Formulario web para relevar disponibilidad de actores para los ensayos de **Jona**.

Sitio publicado:

https://jrrguille-bit.github.io/jona-ensayos/

## Archivos

- `index.html`: formulario público para actores.
- `apps-script.gs`: código para pegar en Google Apps Script y guardar respuestas en Google Sheets.
- `.nojekyll`: evita procesamiento innecesario de Jekyll en GitHub Pages.

## Fechas cargadas

Las fechas fueron tomadas del Excel de dirección y convertidas a opciones de formulario:

- Lunes 13 de julio — Uno a uno con Nico
- Miércoles 15 de julio — Ensayo grupal con Nico, Jefe, compañero de trabajo, amigo de Jona y Negro
- Jueves 16 de julio — Uno a uno con Lorena
- Viernes 17 de julio — Escena de la llamada — Nico + Jona
- Lunes 20 de julio — Uno a uno con Jona
- Martes 21 de julio — Uno a uno con Nico niño
- Miércoles 22 de julio — Uno a uno con Jona niño
- Lunes 27 de julio — Uno a uno con Abuela
- Martes 28 de julio — Ensayo Nico niño + Jona niño
- Miércoles 29 de julio — Ensayo de protagonistas — Nico, Jona, Lorena y Abuela
- Viernes 31 de julio, noche en locación — Ensayo Nico + Jona

## Cómo conectar Google Sheets

1. Crear una Google Sheet nueva.
2. Ir a **Extensiones → Apps Script**.
3. Pegar el contenido de `apps-script.gs`.
4. Guardar.
5. Ir a **Implementar → Nueva implementación**.
6. Tipo: **Aplicación web**.
7. Ejecutar como: **Yo**.
8. Quién tiene acceso: **Cualquier persona**.
9. Copiar la URL de la aplicación web.
10. Pegar esa URL en `index.html`, en esta línea:

```js
const GOOGLE_SCRIPT_URL = "";
```

Cuando la URL esté cargada, las respuestas del formulario entran directo a la hoja `Respuestas`.

## Estado actual

El formulario ya se puede ver y probar. Hasta conectar la URL del Apps Script, el botón de enviar avisa que falta conectar la Google Sheet y permite copiar un resumen de la disponibilidad.
