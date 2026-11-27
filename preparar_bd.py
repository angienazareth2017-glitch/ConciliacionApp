"""Prepara PostgreSQL desde una terminal privada, fuera de Streamlit."""
from __future__ import annotations

import argparse
import os
import sqlite3
from contextlib import closing
from getpass import getpass
from pathlib import Path

import psycopg

from seguridad import conectar, crear_usuario


def migrar_sqlite(conn, ruta: Path) -> None:
    with closing(sqlite3.connect(ruta.resolve().as_uri() + "?mode=ro", uri=True)) as origen:
        origen.execute("BEGIN")
        usuarios = origen.execute(
            "SELECT id, username, nombre, password_hash, salt, rol, activo, creado_en FROM usuarios"
        ).fetchall()
        historial = origen.execute(
            "SELECT id, usuario_id, id_cierre, payload, creado_en FROM historial"
        ).fetchall()
    if not any(u[5] == "admin" and u[6] == 1 for u in usuarios):
        raise ValueError("La copia SQLite debe contener un administrador activo.")
    with conn.cursor() as cur:
        cur.executemany(
            """INSERT INTO tesoriq.usuarios(id, username, nombre, password_hash, salt, rol, activo, creado_en)
               VALUES (%s, %s, %s, %s, %s, %s, %s, %s)""",
            [(*u[:6], bool(u[6]), u[7]) for u in usuarios],
        )
        cur.executemany(
            """INSERT INTO tesoriq.historial(id, usuario_id, id_cierre, payload, creado_en)
               VALUES (%s, %s, %s, %s, %s)""",
            historial,
        )
    siguiente_usuario = max(u[0] for u in usuarios) + 1
    siguiente_historial = max((h[0] for h in historial), default=0) + 1
    conn.execute(f"ALTER TABLE tesoriq.usuarios ALTER COLUMN id RESTART WITH {siguiente_usuario:d}")
    conn.execute(f"ALTER TABLE tesoriq.historial ALTER COLUMN id RESTART WITH {siguiente_historial:d}")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    acciones = parser.add_mutually_exclusive_group(required=True)
    acciones.add_argument("--crear-admin", metavar="USUARIO")
    acciones.add_argument("--migrar-sqlite", type=Path, metavar="COPIA_SQLITE")
    args = parser.parse_args()
    dsn = os.environ.get("TESORIQ_DATABASE_URL_ADMIN") or getpass("Conexión PostgreSQL de mantenimiento: ")
    password = nombre = ""
    if args.crear_admin:
        nombre = input("Nombre del administrador: ")
        password = getpass("Contraseña: ")
        if password != getpass("Confirma la contraseña: "):
            parser.exit(1, "Las contraseñas no coinciden.\n")

    esquema = Path(__file__).parent / "migraciones" / "001_persistencia.sql"
    try:
        with conectar(dsn) as conn:
            conn.execute("SELECT pg_advisory_xact_lock(176312141, 1)")
            conn.execute(esquema.read_text(encoding="utf-8"))
            conn.execute("LOCK TABLE tesoriq.usuarios, tesoriq.historial IN SHARE ROW EXCLUSIVE MODE")
            if conn.execute("SELECT EXISTS (SELECT 1 FROM tesoriq.usuarios) AS ocupado").fetchone()["ocupado"]:
                raise ValueError("La base ya tiene cuentas. La preparación inicial se ejecuta una sola vez.")
            if args.migrar_sqlite:
                migrar_sqlite(conn, args.migrar_sqlite)
            else:
                ok, mensaje = crear_usuario(conn, args.crear_admin, nombre, password, "admin")
                if not ok:
                    raise ValueError(mensaje)
    except ValueError as error:
        parser.exit(1, f"{error}\n")
    except (OSError, psycopg.Error, sqlite3.Error):
        parser.exit(1, "La preparación falló. Revisa la conexión, los permisos y la copia de origen.\n")
    print("Preparación completada.")


if __name__ == "__main__":
    main()
