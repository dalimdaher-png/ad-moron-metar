#!/usr/bin/env python3
"""Automatiza el proceso completo: baja METAR del mes, cruza con la planilla
original de ese mes y manda el resultado por correo.

Uso normal (mes anterior al actual, para correr el dia 1 de cada mes):
    python actualizar_mensual.py

Forzar un mes puntual (util para pruebas o para reprocesar):
    python actualizar_mensual.py 2026-08

Generar el archivo sin mandar el correo:
    python actualizar_mensual.py 2026-08 --sin-correo

Requiere que "configurar_email.py" ya se haya corrido una vez (guarda la
contrasena de aplicacion de Gmail en el almacen de credenciales de Windows).
"""
import csv, json, smtplib, subprocess, sys, traceback
from calendar import monthrange
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from metar_moron import descargar

DIR = Path(__file__).resolve().parent
DATOS = DIR / 'datos'
LOGS = DIR / 'logs'
CONFIG = DIR / 'config_email.json'
ESTACION = 'SADM'

MESES = {1: 'ENERO', 2: 'FEBRERO', 3: 'MARZO', 4: 'ABRIL', 5: 'MAYO', 6: 'JUNIO',
         7: 'JULIO', 8: 'AGOSTO', 9: 'SEPTIEMBRE', 10: 'OCTUBRE', 11: 'NOVIEMBRE', 12: 'DICIEMBRE'}


def mes_objetivo(arg):
    if arg:
        y, m = (int(x) for x in arg.split('-'))
        return y, m
    hoy = date.today()
    primero_actual = hoy.replace(day=1)
    ultimo_anterior = primero_actual - timedelta(days=1)
    return ultimo_anterior.year, ultimo_anterior.month


def log(msg):
    LOGS.mkdir(exist_ok=True)
    linea = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    print(linea)
    with open(LOGS / 'actualizar.log', 'a', encoding='utf-8') as f:
        f.write(linea + '\n')


def bajar_csv(d1, d2, destino):
    filas = descargar(ESTACION, d1, d2)
    with open(destino, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(['station', 'valid', 'metar'])
        for valid, raw in filas:
            w.writerow([ESTACION, valid, raw])
    return len(filas)


def cruzar(original, salida, csv_crudo):
    r = subprocess.run(
        [sys.executable, str(DIR / 'cruzar_planilla.py'), str(original), str(salida), str(csv_crudo)],
        capture_output=True, text=True, cwd=str(DIR))
    if r.returncode != 0:
        raise RuntimeError(f"cruzar_planilla.py fallo:\n{r.stdout}\n{r.stderr}")
    return r.stdout.strip()


def mandar_correo(asunto, cuerpo, adjunto, destinatario_override=None):
    import keyring
    cfg = json.loads(CONFIG.read_text(encoding='utf-8'))
    remitente = cfg['remitente']
    destinatario = destinatario_override or cfg['destinatario']
    clave = keyring.get_password('moron_meteo_email', remitente)
    if not clave:
        raise RuntimeError(
            f"No hay contrasena guardada para {remitente}. Corre 'python configurar_email.py' primero.")

    msg = EmailMessage()
    msg['Subject'] = asunto
    msg['From'] = remitente
    msg['To'] = destinatario
    msg.set_content(cuerpo)
    if adjunto:
        data = Path(adjunto).read_bytes()
        msg.add_attachment(data, maintype='application',
                            subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            filename=Path(adjunto).name)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
        s.login(remitente, clave)
        s.send_message(msg)
    return destinatario


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    sin_correo = '--sin-correo' in sys.argv

    y, m = mes_objetivo(args[0] if args else None)
    d1 = date(y, m, 1)
    d2 = date(y, m, monthrange(y, m)[1])
    mes_txt = f"{MESES[m]} {y}"
    ymd = f"{y:04d}-{m:02d}"

    log(f"=== Iniciando actualizacion para {mes_txt} ===")

    try:
        original = DATOS / f"original_{ymd}.xlsx"
        if not original.exists():
            raise FileNotFoundError(
                f"No encontre '{original}'. Dejala ahi con ese nombre antes de que corra la tarea, "
                f"con las {monthrange(y, m)[1]} hojas diarias de {mes_txt}.")

        DATOS.mkdir(exist_ok=True)
        csv_crudo = DATOS / f"sadm_{ymd}.csv"
        n = bajar_csv(d1, d2, csv_crudo)
        log(f"METAR/SPECI bajados: {n} reportes -> {csv_crudo.name}")

        salida = DIR / f"PLANILLA MOV ACFT {MESES[m]} {y} - con meteo por hora.xlsx"
        resumen = cruzar(original, salida, csv_crudo)
        log(f"Cruce hecho -> {salida.name}\n{resumen}")

        if sin_correo:
            log("(--sin-correo) no se manda el mail.")
            return

        cfg = json.loads(CONFIG.read_text(encoding='utf-8'))
        placeholder = 'PENDIENTE' in cfg['destinatario'].upper()
        destinatario_override = cfg['remitente'] if placeholder else None
        cuerpo = (
            f"Cruce METAR/SPECI -> IMC/VMC de {mes_txt}, adjunto.\n\n{resumen}\n\n"
            f"Generado automaticamente el {datetime.now():%Y-%m-%d %H:%M}."
        )
        if placeholder:
            cuerpo += ("\n\n(Destinatario del jefe todavia no configurado en config_email.json: "
                       "este correo se mando a la propia cuenta como prueba.)")
        destino = mandar_correo(f"AD Morón {mes_txt} - METAR IMC/VMC por hora", cuerpo, salida,
                                 destinatario_override)
        log(f"Correo enviado a {destino}")

    except Exception as e:
        log(f"ERROR: {e}")
        log(traceback.format_exc())
        try:
            cfg = json.loads(CONFIG.read_text(encoding='utf-8'))
            mandar_correo(f"[ERROR] AD Morón {mes_txt} no se pudo generar",
                          f"Fallo la actualizacion automatica de {mes_txt}:\n\n{e}\n\n"
                          f"Revisa logs/actualizar.log para el detalle.",
                          None, destinatario_override=cfg['remitente'])
        except Exception:
            log("(tampoco se pudo mandar el aviso de error por correo)")
        sys.exit(1)


if __name__ == '__main__':
    main()
