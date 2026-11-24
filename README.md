# Calculadora de honorarios por temporalidad

Para ejecutarla localmente con Node.js:

```sh
node scripts/serve-local.cjs
```

Abra http://127.0.0.1:4173. Para usar otro puerto, configure `PORT` antes de iniciar el servidor.

Capture folios en los niveles que necesite. **Descargar PDF** genera directamente una página tamaño Carta con los niveles que tienen folios. **Descargar Excel** genera un `.xlsx` con encabezados, formatos monetarios, fórmulas y totales. Mostrar los niveles en cero en pantalla no los agrega a las descargas. También se conserva la impresión mediante Ctrl+P.

El nombre del PDF incluye el acrónimo del organismo (columna C del catálogo), o su nombre si no tiene acrónimo, la cantidad total de folios, el tipo de financiamiento, el importe total y las fechas del periodo. Por ejemplo: `Honorarios_JGCDMX_10_Folios_FISCALES_$5,365,853.33_del_2026-01-01_al_2026-12-29.pdf`. Para financiamiento "OTRO" se usa el texto capturado. Se eliminan caracteres incompatibles con nombres de archivo y se limita la longitud del organismo y del financiamiento para permitir guardar nombres extensos.

Las bibliotecas de exportación están incluidas en `vendor/`. La interfaz utiliza Tailwind, fuentes e iconos externos, y el catálogo de organismos se consulta en la hoja de Google configurada. Si no se puede consultar el catálogo, se puede escribir el organismo manualmente.

Para verificar las exportaciones y generar archivos de ejemplo:

```sh
node scripts/verify-exports.cjs --samples
```

Los ejemplos se guardan en `output/previos/`. Las pruebas comprueban la selección de niveles, los totales, los metadatos y el número de páginas del PDF; también reabren el Excel para verificar sus valores y fórmulas.
