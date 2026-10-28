# Herramienta web · Taller de codiseño

Página estática para el taller (celular, computador de mesa y proyector). Las respuestas se guardan en una hoja de Google mediante una aplicación web de Google Apps Script. Este repositorio **no contiene datos del cliente**: ideas, ejemplos, áreas y clave del facilitador viven en `Datos.gs`, dentro del proyecto de Apps Script, y la página los carga al abrir.

## Vistas

| Vista | Enlace | Uso |
| --- | --- | --- |
| Celular | `https://<usuario>.github.io/<repo>/` | Momentos 1 y 2: para qué, escribir preguntas, votar |
| Mesa | `…/?vista=mesa` | Momentos 3 y 4, un computador por mesa |
| Resultados | `…/?vista=resultados` | Proyector; "Modo facilitador" pide la clave de `Datos.gs` |

## Puesta en marcha

1. **Servidor (Apps Script).** Abrir la hoja de respuestas > Extensiones > Apps Script. Pegar `apps-script/Code.gs` en `Code.gs` y crear `Datos.gs` (botón + > Secuencia de comandos) con el archivo que está en la carpeta de Drive del taller (no está en este repositorio).
2. Implementar > Nueva implementación > Aplicación web. Ejecutar como: **Yo**. Quién tiene acceso: **Cualquier persona** (no "Cualquier persona con cuenta de Google"). Autorizar y copiar la URL que termina en `/exec`.
3. **Página (GitHub Pages).** Pegar esa URL en `config.js`. En el repositorio: Settings > Pages > Deploy from a branch > `main` / root.
4. Probar desde un celular con datos móviles y sin sesión de Google: marcar un para qué y revisar que aparezca en la hoja (pestaña ParaQue). Borrar las filas de prueba.

## Datos personales

No se piden nombres ni correos: solo área y mesa. No escribir datos de personas atendidas (Ley 1581 de 2012).
