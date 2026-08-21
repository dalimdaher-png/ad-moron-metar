# Cruce METAR/SPECI -> IMC/VMC — Torre Morón (SADM)

Agrega a la planilla de movimientos, **hora por hora**, la condición meteorológica
del aeródromo tomada del METAR/SPECI oficial.

---

## Qué hay acá

| Archivo | Para qué |
|---|---|
| `metar_moron.py` | Parser de METAR + clasificador IMC/VMC. Es el motor, no se corre solo. |
| `cruzar_planilla.py` | Inyecta las columnas dentro de la PLANILLA MOV ACFT. Se puede correr solo, o vía `actualizar_mensual.py`. |
| `actualizar_mensual.py` | **Automatización mensual completa**: baja METAR, cruza y manda el correo. Ver más abajo. |
| `configurar_email.py` | Setup de una sola vez: guarda la contraseña de Gmail. Se corre a mano. |
| `actualizar_ahora.bat` | Doble click para correr `actualizar_mensual.py` sin abrir una terminal. |
| `config_email.json` | Remitente / destinatario del correo mensual. |
| `PowerQuery_METAR_SADM.m` | Alternativa sin Python: se pega dentro del Excel. |
| `datos/sadm_agosto.csv` | METAR + SPECI de SADM, 1 al 21 de agosto de 2026. |
| `datos/original_YYYY-MM.xlsx` | Planilla original de cada mes, la deja quien la usa antes de que corra la tarea. |
| `datos/PLANILLA ... original.xlsx` | Tu planilla sin tocar, por si hay que volver atrás. |
| `PLANILLA ... con meteo por hora.xlsx` | El resultado. |

## Requisitos

```
python  3.9 o superior
pip install openpyxl
```

## Uso

**Con el CSV que ya está bajado:**

```
python cruzar_planilla.py "datos/PLANILLA MOV ACFT AGOSTO 2026 - original.xlsx" "salida.xlsx" datos/sadm_agosto.csv
```

**Bajando METAR nuevos** (necesita internet; en la PC del trabajo debería andar):

```
python metar_moron.py 2026-09-01 2026-09-30
```

Genera `meteo_SADM_2026-09-01_2026-09-30.xlsx` y, de paso, se puede guardar el CSV
crudo para reusar con `cruzar_planilla.py`.

## Automatización mensual (bajar + cruzar + mandar por correo)

Corre solo el **día 1 de cada mes a las 08:00** (tarea programada de Windows
"AD Moron - actualizacion mensual"), procesando el mes que acaba de cerrar.
También se puede disparar a mano.

**Setup de una sola vez** (ya hecho en esta PC, dejar documentado por si hay que
repetirlo en otra):

```
pip install keyring
python configurar_email.py
```

Pide la cuenta de Gmail que envía, el email del jefe (destinatario), y la
contraseña de aplicación de Gmail (**no** la contraseña normal — se genera en
https://myaccount.google.com/apppasswords con verificación en 2 pasos activada).
La contraseña queda en el almacén de credenciales de Windows, no en ningún
archivo de texto. El remitente y el destinatario quedan en `config_email.json`
(no es secreto, se puede editar a mano si cambia el destinatario).

**Antes del día 1 de cada mes**, dejar la planilla original de ese mes en:

```
datos/original_YYYY-MM.xlsx      (ej: datos/original_2026-09.xlsx)
```

Si no está, la tarea no rompe: manda un correo de aviso al remitente con el
error en vez de fallar en silencio (queda todo en `logs/actualizar.log`
también).

**Para correr a mano** (probar, reprocesar un mes, o generar sin mandar el
correo):

```
python actualizar_mensual.py              # mes anterior al actual
python actualizar_mensual.py 2026-08      # mes puntual
python actualizar_mensual.py 2026-08 --sin-correo   # no manda el mail
```

O doble click en `actualizar_ahora.bat`.

**Mientras el destinatario en `config_email.json` diga "PENDIENTE"**, el correo
se manda igual pero a la propia cuenta remitente, como prueba — así se puede
validar todo el flujo antes de tener la dirección del jefe.

## Qué queda en la planilla

En cada hoja diaria (1 a 31), columnas **R a V**, alineadas con las franjas de las
filas 10 a 33:

| Col | Contenido |
|---|---|
| R | CONDICION — IMC / VMC / SIN DATO |
| S | VIS (m) — la peor visibilidad de esa hora |
| T | TECHO (ft) — la capa BKN/OVC/VV más baja |
| U | SPECI — si hubo un reporte especial dentro de la hora |
| V | METAR OBSERVADO — el texto crudo |

Más dos hojas nuevas:

- **METEO_SADM** — la tabla fuente. Las 31 hojas la leen con BUSCARV usando la fecha de `I6`.
- **CONTROL** — comparación entre la marca "IMC" escrita a mano y la clasificación
  automática. En naranja, arriba de todo, las horas que fueron IMC **y tuvieron
  movimientos**: eso es lo que la planilla actual no puede mostrar.

El área de impresión sigue siendo `B1:M46`, así que el formulario firmado sale igual.

---

## Criterio aplicado

```
IMC   si   visibilidad < 5000 m   O   techo < 1500 ft
VMC   en cualquier otro caso
```

- **Techo** = capa más baja BKN, OVC o VV. **SCT y FEW no son techo.**
- **CAVOK** = visibilidad 10 km o más y sin nubes significativas -> VMC.
- Se usa "menor que": un techo de exactamente 1500 ft da **VMC**.
- **SPECI incluidos**: la condición de la hora es la **peor** reportada dentro de
  esa hora, no la del METAR de la hora en punto.
- Todo en **UTC (Z)**, igual que la planilla.

### PENDIENTE — el umbral no está confirmado

5000 m / 1500 ft es el criterio general (RAAC 91 / OACI Anexo 2), **no** el
publicado para Morón. Si Morón usa otro, se cambian las dos constantes arriba de
`metar_moron.py` y se vuelve a correr:

```python
VIS_MIN_M    = 5000
TECHO_MIN_FT = 1500
```

Esto no es cosmético: **20 de las 22 horas IMC-con-movimientos son por techo entre
1000 y 1400 ft.** Con un umbral de techo más bajo, casi toda esa tabla desaparece.

---

## Fuente de datos

Iowa Environmental Mesonet — archivo público de METAR/SPECI, cobertura mundial.

```
https://mesonet.agron.iastate.edu/cgi-bin/request/asos.py
  ?station=SADM&data=metar&year1=..&month1=..&day1=..
  &year2=..&month2=..&day2=..&tz=Etc/UTC&format=onlycomma
```

Es el **mismo texto crudo** que muestra Ogimet — se puede verificar reporte por
reporte. Ogimet no se usa como fuente automatizada porque su `robots.txt` prohíbe
el acceso automático y bloquea scrapers; sirve para consulta manual.

**Ojo con el rate limit:** si se piden rangos grandes muy seguidos devuelve HTTP
429. `metar_moron.py` corta en tramos de 7 días con pausa de 3 segundos.

---

## Problemas detectados en la planilla de agosto (no son del script)

1. **Hoja "23": la fecha en `I6` dice `2026-06-23`** — junio, no agosto. Por eso esa
   hoja muestra `s/d` en todas las columnas nuevas. Corregir `I6` y se llena sola.
2. **Hoja 19, celda `F31`**: tiene pegado un NOTAM completo de FIR Ezeiza dentro de
   una columna de conteo de movimientos.
3. **Hoja 2**: el `imc` está escrito tres veces por fila, en las columnas E, F y G.
4. Los formatos de la marca manual varían: `imc`, `IMC`, `AD IMC`.
5. **METAR con error de tipeo en el origen**: `05/08 05:00Z` dice `BRK008` en vez de
   `BKN008`. No se corrige automáticamente. Si era BKN, esa hora es IMC (techo
   800 ft) y el archivo la da VMC. Está marcada en `METEO_SADM`, columna
   `TOKENS_RAROS`.

## Hallazgo principal

Las **55 horas marcadas "IMC" a mano en agosto tienen 0 movimientos. Todas.**
Esa columna no es un registro meteorológico: es la explicación de un cero. Por
construcción nunca puede mostrar "volamos menos por el clima", solo "no volamos".

El cruce automático encuentra **22 horas que fueron IMC y tuvieron tráfico igual —
334 movimientos** que hoy la planilla no puede exhibir. Están en la hoja CONTROL.

De las 55 marcas manuales, **54 dan IMC** con este criterio: la marca manual no está
equivocada, está incompleta.

