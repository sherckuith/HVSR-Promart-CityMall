# HVSR Promart-CityMall

Trabajo de análisis de respuesta de sitio con mediciones Nakamura para el proyecto `Trabajos ARS`. Este directorio contiene el material de campo, la clasificación local de MiniSEED, las coordenadas de los puntos y la bitácora preparada para el análisis HVSR.

Repositorio: <https://github.com/sherckuith/HVSR-Promart-CityMall>

Skill utilizada: <https://github.com/sherckuith/infonak>

## Fuente de campo

La fuente temporal y operativa fue la exportación completa de WhatsApp:

```text
Informe/Registro Fotografico/+593 96 953 1673.zip
```

El ZIP contiene `chat.txt` y `chat.md`, con la conversación desde el lunes hasta el 2026-10-08. La exportación se conserva localmente como evidencia, pero no se publica en GitHub por contener la conversación de WhatsApp. El Excel y este README documentan el método y los resultados extraídos.

El trabajo utiliza dos relojes:

- `reloj`: tiempo de medición sincronizado con UTC; se usó para intersectar los MiniSEED.
- `inicio`/`final`: tiempo mostrado por el equipo DAS, conservado para documentar el desfase por falta de NTP en campo.

## Estructura

```text
RAW/Blue/<día>/                 MiniSEED originales del Das-Blue / RS3D
RAW/Yellow/<día>/               MiniSEED originales del Das-Yellow / RS4D
Sitios Nakamura/N1 ... N15/     registros copiados y clasificados
Sitios Nakamura/N#/Registros/Descartados/
Informe/Bitacora.xlsx           coordenadas, tiempos y nombres de registros
Informe/Registro Fotografico/   fotos organizadas por sitio
```

RAW no se modifica. Los registros de los sitios son copias, no movimientos de los originales.

## Clasificación realizada

La regla usada para cada bloque MiniSEED de 300 s fue:

```text
inicio_archivo < final_reloj
fin_archivo    > inicio_reloj
```

El nombre y el encabezado MiniSEED se revisaron contra el intervalo UTC del mensaje. Las variantes `N2 actualizado`, `N3 actualizado` y `N9 actualizado` se conservaron como mediciones adicionales del mismo número de sitio y no como puntos nuevos.

## Datos extraídos de la conversación

Los tiempos siguientes son los datos de la bitácora completa. `Reloj UTC` es el intervalo utilizado para buscar los MiniSEED; `equipo` es el intervalo mostrado por el DAS.

| Sitio | Latitud | Longitud | Reloj UTC | Tiempo del equipo | Equipo |
|---|---:|---:|---|---|---|
| N1 | -2.1389506 | -79.9109948 | 2026-10-05 18:20–18:45 | 18:20–18:45 | — |
| N2 | -2.1387868 | -79.9103833 | 2026-10-07 13:00–13:35 | 08:06–08:40 | Blue |
| N3 | -2.1388312 | -79.9103810 | 2026-10-07 13:06–13:42 | 08:19–08:55 | Yellow |
| N4 | -2.1391645 | -79.9102661 | 2026-10-07 15:16–15:47 | 08:19–08:50 | Blue |
| N5 | -2.1398310 | -79.9095325 | 2026-10-07 15:26–16:00 | 08:19–08:55 | Yellow |
| N6 | -2.1398780 | -79.9097595 | 2026-10-07 15:56–16:30 | 08:59–09:32 | Blue |
| N7 | -2.1397941 | -79.9103107 | 2026-10-07 16:07–16:43 | 09:01–09:36 | Yellow |
| N8 | -2.1387928 | -79.9104797 | 2026-10-07 16:35–17:08 | 09:38–10:11 | Blue |
| N9 | -2.1394408 | -79.9114019 | 2026-10-07 16:51–sin final | 09:44–sin final | Yellow; batería descargada |
| N10 | -2.1395091 | -79.9115469 | 2026-10-07 17:15–17:48 | 10:18–10:50 | Blue |
| N11 | -2.1394707 | -79.9115146 | 2026-10-07 17:54–18:27 | 10:56–11:30 | Blue |
| N12 | -2.1391612 | -79.9112014 | 2026-10-07 18:31–19:02 | 11:34–12:05 | Blue |
| N13 | -2.1399517 | -79.9107717 | 2026-10-07 19:06–19:37 | 12:09–12:40 | Blue |
| N14 | -2.1402488 | -79.9103401 | 2026-10-07 19:40–20:07 | 12:43–13:10 | Blue |
| N15 | -2.1402582 | -79.9101212 | 2026-10-07 20:11–20:43 | 13:14–13:46 | Blue |
| N9 actualizado | -2.1395529 | -79.9108640 | 2026-10-07 20:51–21:25 | 13:53–14:27 | Blue |
| N3 actualizado | -2.1392937 | -79.9104110 | 2026-10-07 21:29–22:03 | 14:31–15:05 | Blue |
| N2 actualizado | -2.1391298 | -79.9107470 | 2026-10-07 22:06–22:40 | 15:08–15:42 | Blue |

| Sitio | Registros base | Variante | Equipo / observación |
|---|---:|---:|---|
| N1 | 5 | — | clasificación previa verificada; incluye capturas de TremorForge |
| N2 | 5 | N2 actualizado: 7 | Blue |
| N3 | 4 | N3 actualizado: 8 | Yellow en base; Blue en actualización |
| N4 | 7 | — | Blue |
| N5 | 0 | — | Yellow; sin intersección RAW comprobada |
| N6 | 7 | — | Blue |
| N7 | 0 | — | Yellow; sin intersección RAW comprobada |
| N8 | 7 | — | Blue |
| N9 | 0 | N9 actualizado: 7 | base Yellow sin final por descarga de batería; actualización Blue |
| N10 | 7 | — | Blue |
| N11 | 8 | — | Blue |
| N12 | 7 | — | Blue |
| N13 | 7 | — | Blue |
| N14 | 6 | — | Blue |
| N15 | 7 | — | Blue |

Los sitios N5, N7 y N9 base tienen intervalos de reloj en la conversación, pero no se encontró un MiniSEED Yellow de RAW que intersecte esos intervalos. No se forzaron registros desde otro sitio. En N9 base, además, la conversación indica que la batería se descargó y no se conoce el tiempo final; la actualización nocturna sí tiene siete registros Blue.

## Bitácora

`Informe/Bitacora.xlsx` contiene las ocho columnas requeridas:

1. Número de N
2. Latitud
3. Longitud
4. Tiempo de inicio (reloj UTC)
5. Tiempo final (reloj UTC)
6. Tiempo de inicio en equipo
7. Tiempo final del equipo
8. Registros mseed correspondientes

Las coordenadas de las variantes se mantienen en las carpetas de cada variante y se anotan en la bitácora. Los registros no asignables siguen en RAW y se consideran pendientes de revisión, no descartes.

## Validación y estado

- Estructura de carpetas: preparada para 15 puntos Nakamura.
- Excel: generado, recalculado, inspeccionado y renderizado visualmente.
- Clasificación: realizada localmente con copia desde RAW.
- RAW: preservado sin modificación.
- Informe final: todavía no emitido.
- Release: no creado; se hará únicamente cuando el informe esté listo y se solicite.

## Política Git

El repositorio se publica en la rama `main`. Se incluyen RAW, registros clasificados, fotografías, mapas, capturas, Excel y documentación. El `.gitignore` excluye únicamente la exportación/conversación de WhatsApp y los MiniSEED ubicados dentro de `Registros/Descartados/`. Los MiniSEED de RAW y de los sitios clasificados sí forman parte del repositorio.
