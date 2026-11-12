import asyncio
import os
import sqlite3
import time
from datetime import date

import flet as ft
from openpyxl import Workbook
from openpyxl.styles import Font

# ---------------------------------------------------------------- almacenamiento
DATA_DIR = os.getenv("FLET_APP_STORAGE_DATA", os.path.abspath("."))
EXPORT_DIR = os.path.join(DATA_DIR, "exports")
os.makedirs(EXPORT_DIR, exist_ok=True)

db = sqlite3.connect(os.path.join(DATA_DIR, "cronoclase.db"), check_same_thread=False)
db.executescript(
    """
    CREATE TABLE IF NOT EXISTS students(
        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS exercises(
        id INTEGER PRIMARY KEY AUTOINCREMENT, day TEXT NOT NULL,
        name TEXT NOT NULL, passes INTEGER NOT NULL);
    CREATE TABLE IF NOT EXISTS results(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        exercise_id INTEGER NOT NULL, student_id INTEGER NOT NULL,
        pass_no INTEGER NOT NULL, seconds REAL NOT NULL);
    """
)
db.commit()


def q(sql, args=()):
    return db.execute(sql, args).fetchall()


def run(sql, args=()):
    db.execute(sql, args)
    db.commit()


def fmt(sec):
    m, s = divmod(sec, 60)
    return f"{int(m):02d}:{s:05.2f}"


# ---------------------------------------------------------------- exportacion
def write_day_blocks(ws, day):
    """Escribe un bloque por ejercicio: alumno x pasadas + mejor + promedio."""
    row = 1
    for ex_id, name, passes in q(
        "SELECT id, name, passes FROM exercises WHERE day=? ORDER BY id", (day,)
    ):
        ws.cell(row, 1, f"{name}  ({day})").font = Font(bold=True, size=12)
        row += 1
        headers = ["Alumno"] + [f"Pasada {i}" for i in range(1, passes + 1)] + ["Mejor", "Promedio"]
        for c, h in enumerate(headers, 1):
            ws.cell(row, c, h).font = Font(bold=True)
        row += 1
        for sid, sname in q(
            "SELECT DISTINCT s.id, s.name FROM students s "
            "JOIN results r ON r.student_id=s.id WHERE r.exercise_id=? ORDER BY s.id",
            (ex_id,),
        ):
            times = {
                p: sec
                for p, sec in q(
                    "SELECT pass_no, seconds FROM results WHERE exercise_id=? AND student_id=?",
                    (ex_id, sid),
                )
            }
            ws.cell(row, 1, sname)
            for p in range(1, passes + 1):
                if p in times:
                    ws.cell(row, 1 + p, round(times[p], 2))
            if times:
                ws.cell(row, passes + 2, round(min(times.values()), 2))
                ws.cell(row, passes + 3, round(sum(times.values()) / len(times), 2))
            row += 1
        row += 1
    ws.column_dimensions["A"].width = 28


def export_day(day):
    wb = Workbook()
    ws = wb.active
    ws.title = day
    write_day_blocks(ws, day)
    path = os.path.join(EXPORT_DIR, f"clase_{day}.xlsx")
    wb.save(path)
    return path


def export_all():
    wb = Workbook()
    det = wb.active
    det.title = "Detalle"
    for c, h in enumerate(["Fecha", "Ejercicio", "Alumno", "Pasada", "Segundos"], 1):
        det.cell(1, c, h).font = Font(bold=True)
    rows = q(
        "SELECT e.day, e.name, s.name, r.pass_no, r.seconds FROM results r "
        "JOIN exercises e ON e.id=r.exercise_id JOIN students s ON s.id=r.student_id "
        "ORDER BY e.day, e.id, s.id, r.pass_no"
    )
    for i, (d, ex, st, p, sec) in enumerate(rows, 2):
        for c, v in enumerate([d, ex, st, p, round(sec, 2)], 1):
            det.cell(i, c, v)
    det.column_dimensions["B"].width = 24
    det.column_dimensions["C"].width = 28
    for (day,) in q("SELECT DISTINCT day FROM exercises ORDER BY day"):
        write_day_blocks(wb.create_sheet(day), day)
    path = os.path.join(EXPORT_DIR, "clase_todos_los_dias.xlsx")
    wb.save(path)
    return path


# ---------------------------------------------------------------- interfaz
def main(page: ft.Page):
    page.title = "CronoClase"
    page.scroll = ft.ScrollMode.AUTO
    page.padding = 16
    settings = {"round_robin": False}

    def today():
        return date.today().isoformat()

    def render(*controls):
        page.clean()
        page.add(*controls)
        page.update()

    def back_btn(target):
        return ft.TextButton("< Volver", on_click=lambda e: target())

    # ---------------- inicio
    def show_home(msg=""):
        render(
            ft.Text("CronoClase", size=30, weight=ft.FontWeight.BOLD),
            ft.Text(f"Hoy: {today()}"),
            ft.ElevatedButton("Alumnos", on_click=lambda e: show_students(), width=300),
            ft.ElevatedButton("Entrenamiento de hoy", on_click=lambda e: show_today(), width=300),
            ft.ElevatedButton("Exportar", on_click=lambda e: show_export(), width=300),
            ft.Text(msg),
        )

    # ---------------- alumnos
    def show_students():
        name_in = ft.TextField(label="Nombre del alumno", expand=True)

        def add(e):
            if name_in.value.strip():
                run("INSERT INTO students(name) VALUES (?)", (name_in.value.strip(),))
                show_students()

        def delete(sid):
            run("DELETE FROM results WHERE student_id=?", (sid,))
            run("DELETE FROM students WHERE id=?", (sid,))
            show_students()

        rows = [
            ft.Row(
                [
                    ft.Text(n, expand=True),
                    ft.TextButton("Borrar", on_click=lambda e, i=i: delete(i)),
                ]
            )
            for i, n in q("SELECT id, name FROM students ORDER BY id")
        ]
        render(
            back_btn(show_home),
            ft.Text("Alumnos", size=24, weight=ft.FontWeight.BOLD),
            ft.Row([name_in, ft.ElevatedButton("Agregar", on_click=add)]),
            *rows,
        )

    # ---------------- ejercicios del dia
    def show_today():
        name_in = ft.TextField(label="Ejercicio", expand=True)
        passes_in = ft.TextField(
            label="Pasadas", value="3", width=100, keyboard_type=ft.KeyboardType.NUMBER
        )

        def add(e):
            if name_in.value.strip() and passes_in.value.strip().isdigit():
                run(
                    "INSERT INTO exercises(day, name, passes) VALUES (?,?,?)",
                    (today(), name_in.value.strip(), int(passes_in.value)),
                )
                show_today()

        def delete(eid):
            run("DELETE FROM results WHERE exercise_id=?", (eid,))
            run("DELETE FROM exercises WHERE id=?", (eid,))
            show_today()

        rows = [
            ft.Row(
                [
                    ft.ElevatedButton(
                        f"{n} ({p} pasadas)",
                        expand=True,
                        on_click=lambda e, i=i: show_run(i),
                    ),
                    ft.TextButton("Borrar", on_click=lambda e, i=i: delete(i)),
                ]
            )
            for i, n, p in q("SELECT id, name, passes FROM exercises WHERE day=? ORDER BY id", (today(),))
        ]
        render(
            back_btn(show_home),
            ft.Text("Ejercicios de hoy", size=24, weight=ft.FontWeight.BOLD),
            ft.Row([name_in, passes_in]),
            ft.ElevatedButton("Agregar ejercicio", on_click=add),
            ft.Divider(),
            ft.Text("Toca un ejercicio para empezar a cronometrar:"),
            *rows,
        )

    # ---------------- cronometro por ejercicio
    def show_run(ex_id):
        ex_name, passes = q("SELECT name, passes FROM exercises WHERE id=?", (ex_id,))[0]
        students = q("SELECT id, name FROM students ORDER BY id")
        counts = {
            sid: q(
                "SELECT COUNT(*) FROM results WHERE exercise_id=? AND student_id=?", (ex_id, sid)
            )[0][0]
            for sid, _ in students
        }

        # determinar quien sigue
        current = None
        if students:
            if settings["round_robin"]:
                lowest = min(counts.values())
                if lowest < passes:
                    for sid, sname in students:
                        if counts[sid] == lowest:
                            current = (sid, sname, lowest + 1)
                            break
            else:
                for sid, sname in students:
                    if counts[sid] < passes:
                        current = (sid, sname, counts[sid] + 1)
                        break

        state = {"running": False, "t0": 0.0}
        label = ft.Text("00:00.00", size=64, weight=ft.FontWeight.BOLD)
        btn = ft.ElevatedButton("INICIAR", width=300, height=70)

        async def tick():
            while state["running"]:
                label.value = fmt(time.perf_counter() - state["t0"])
                label.update()
                await asyncio.sleep(0.03)

        def toggle(e):
            if not state["running"]:
                state["running"] = True
                state["t0"] = time.perf_counter()
                btn.text = "PARAR"
                btn.update()
                page.run_task(tick)
            else:
                elapsed = time.perf_counter() - state["t0"]
                state["running"] = False
                sid, _, pno = current[0], current[1], current[2]
                run(
                    "INSERT INTO results(exercise_id, student_id, pass_no, seconds) VALUES (?,?,?,?)",
                    (ex_id, sid, pno, elapsed),
                )
                show_run(ex_id)

        btn.on_click = toggle

        def undo(e):
            last = q("SELECT id FROM results WHERE exercise_id=? ORDER BY id DESC LIMIT 1", (ex_id,))
            if last:
                run("DELETE FROM results WHERE id=?", (last[0][0],))
            show_run(ex_id)

        def set_mode(e):
            settings["round_robin"] = e.control.value
            show_run(ex_id)

        def leave(e):
            state["running"] = False
            show_today()

        if current:
            head = [
                ft.Text(f"Alumno: {current[1]}", size=26, weight=ft.FontWeight.BOLD),
                ft.Text(f"Pasada {current[2]} de {passes}", size=18),
                label,
                btn,
            ]
        elif not students:
            head = [ft.Text("No hay alumnos cargados.")]
        else:
            head = [ft.Text("Ejercicio completo", size=24, weight=ft.FontWeight.BOLD)]

        summary = []
        for sid, sname in students:
            times = q(
                "SELECT seconds FROM results WHERE exercise_id=? AND student_id=? ORDER BY pass_no",
                (ex_id, sid),
            )
            summary.append(ft.Text(f"{sname}: " + "  |  ".join(fmt(t[0]) for t in times)))

        render(
            ft.TextButton("< Volver", on_click=leave),
            ft.Text(ex_name, size=22, weight=ft.FontWeight.BOLD),
            ft.Switch(
                label="Rotar alumnos en cada pasada",
                value=settings["round_robin"],
                on_change=set_mode,
            ),
            *head,
            ft.TextButton("Deshacer ultimo tiempo", on_click=undo),
            ft.Divider(),
            *summary,
        )

    # ---------------- exportar
    def show_export():
        out = ft.Text("")

        def do_day(e):
            out.value = f"Guardado en:\n{export_day(today())}"
            out.update()

        def do_all(e):
            out.value = f"Guardado en:\n{export_all()}"
            out.update()

        render(
            back_btn(show_home),
            ft.Text("Exportar a Excel", size=24, weight=ft.FontWeight.BOLD),
            ft.ElevatedButton("Exportar el dia de hoy", on_click=do_day, width=300),
            ft.ElevatedButton("Exportar todos los dias", on_click=do_all, width=300),
            out,
        )

    show_home()


ft.app(main)
