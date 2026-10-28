# Herramienta web · Taller de codiseño

Página estática para el taller. Las respuestas se guardan en una hoja de Google mediante una aplicación web de Google Apps Script. Este repositorio **no contiene datos del cliente**: ideas, ejemplos, áreas y clave del facilitador viven en `Datos.gs`, dentro del proyecto de Apps Script, y la página los carga al abrir.

## Un solo enlace

| Vista | Enlace | Uso |
| --- | --- | --- |
| Participantes | `https://<usuario>.github.io/<repo>/` | Celular o computador: 1 Para qué · 2 Escribir preguntas · Votar · 3 Qué tenemos · 4 Cómo se ve |
| Resultados | `…/?vista=resultados` | Proyector: pestañas por momento y resumen de marco lógico. "Modo facilitador" pide la clave de `Datos.gs` |

En el modo facilitador se fijan los para qué acordados, los actores nuevos, las preguntas priorizadas (por defecto, las 4 más votadas) y, si se quiere, grupos de trabajo (sin grupos, o de 2 a 6). El enlace antiguo `?vista=mesa` abre "Qué tenemos".

## Puesta en marcha

1. **Servidor (Apps Script).** Pegar `apps-script/Code.gs` en `Code.gs` y crear `Datos.gs` con el archivo de la carpeta de Drive del taller (no está en este repositorio).
2. Implementar > Aplicación web. Ejecutar como: **Yo**. Quién tiene acceso: **Cualquier persona**. Copiar la URL que termina en `/exec`.
3. **Página (GitHub Pages).** Pegar esa URL en `config.js`. Settings > Pages > Deploy from a branch > `main` / root.
4. Probar desde un celular sin sesión de Google y borrar las filas de prueba (o "Borrar datos simulados" en el modo facilitador).

## Datos personales

No se piden nombres ni correos: solo área y, si se activan, grupo. No escribir datos de personas atendidas (Ley 1581 de 2012).
