#!/usr/bin/env python3
"""Configuracion de una sola vez: guarda la contrasena de aplicacion de Gmail
en el almacen de credenciales de Windows (no queda en ningun archivo de texto).

Antes de correr esto necesitas una "contrasena de aplicacion" de Gmail:
  1. Activa la verificacion en 2 pasos en https://myaccount.google.com/security
  2. Anda a https://myaccount.google.com/apppasswords
  3. Genera una para "Correo" / nombre libre (ej: "meteo moron"), 16 caracteres.

Corre esto en una terminal (no lo puede correr Claude porque necesita que
tipees la contrasena vos, oculta):

    python configurar_email.py
"""
import getpass
import json
from pathlib import Path

import keyring

DIR = Path(__file__).resolve().parent
CONFIG = DIR / 'config_email.json'


def main():
    cfg = json.loads(CONFIG.read_text(encoding='utf-8')) if CONFIG.exists() else {
        'remitente': '', 'destinatario': '', 'estacion': 'SADM'}

    remitente = input(f"Cuenta de Gmail que envia [{cfg.get('remitente') or 'dalimdaher@gmail.com'}]: ").strip()
    remitente = remitente or cfg.get('remitente') or 'dalimdaher@gmail.com'

    destinatario = input(f"Email del jefe (destinatario) [{cfg.get('destinatario') or 'sin definir'}]: ").strip()
    destinatario = destinatario or cfg.get('destinatario') or 'PENDIENTE - completar con el email del jefe'

    clave = getpass.getpass("Contrasena de aplicacion de Gmail (16 caracteres, no se muestra en pantalla): ").strip()
    if not clave:
        print("No se ingreso contrasena, no se guardo nada.")
        return

    keyring.set_password('moron_meteo_email', remitente, clave)
    cfg['remitente'] = remitente
    cfg['destinatario'] = destinatario
    CONFIG.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding='utf-8')

    print(f"\nListo. Guardado en el almacen de credenciales de Windows para {remitente}.")
    print(f"Destinatario configurado: {destinatario}")
    print("Podes probar todo con: python actualizar_mensual.py 2026-08")


if __name__ == '__main__':
    main()
