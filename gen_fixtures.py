"""
Genera los fixtures Excel usados por los tests:
  - src/test/resources/fixtures/extraccion.xlsx
  - src/test/resources/fixtures/cierre.xlsx

Se deja este script para poder regenerar los fixtures si cambian las
cabeceras esperadas por los perfiles. Los .xlsx resultantes se
versionan en src/test/resources/fixtures/ junto con los tests.

Uso (desde la raiz del proyecto):
    pip install openpyxl --break-system-packages
    python3 gen_fixtures.py
"""

from openpyxl import Workbook

OUT_DIR = "src/test/resources/fixtures"

# =========================================================================
# extraccion.xlsx
# =========================================================================
# Perfil Extraccion (config.properties): cabeceras en fila 1. Debe tener
# al menos 4 cabeceras de las del profile para que se detecte.
# 14 filas con Peticion + 1 con Peticion vacia (skip por ancla).
EXTRACCION_HEADERS = [
    "Peticion", "Titulo", "Aplicaci_Activi", "Estado", "Usuario_Resp_Tecnico",
    "Horas_AutoriTotPeticion", "Horas_RealizadoTot", "Estado_Distribucion",
    "Planificacion", "Objeto_EstudioPeticion", "Codigo_Facturacion",
    "Recurso", "Funcion", "UltimaPrevision_Horas_Mes", "Realizadas_Horas_Mes",
    "Total_Horas_Autorizadas_Recurso", "Total_Horas_Realizadas_Recurso",
]

def build_extraccion():
    wb = Workbook()
    ws = wb.active
    ws.title = "Extraccion"
    ws.append(EXTRACCION_HEADERS)

    # 14 filas validas
    rows = [
        ["P-001", "Migracion BBDD",          "DF", "Abierta",  "tresp1@x", 40, 12, "OK", "P1", "Obj A", "CTR-100", "M-1001", "Dev",    10, 5,  40, 20],
        ["P-002", "Refactor UI",             "HE", "Abierta",  "tresp1@x", 60, 30, "OK", "P1", "Obj B", "CTR-100", "M-1002", "Dev",    20, 15, 60, 35],
        ["P-003", "API v2",                  "EW", "Cerrada",  "tresp2@x", 25, 25, "OK", "P2", "Obj C", "CTR-200", "M-1003", "Dev",     0,  0, 25, 25],
        ["P-004", "Dashboards",              "J6", "Abierta",  "tresp2@x", 80, 10, "OK", "P2", "Obj D", "CTR-200", "M-1004", "Dev",    30, 10, 80, 20],
        ["P-005", "Integracion SSO",         "JX", "Abierta",  "tresp1@x", 45,  0, "OK", "P1", "Obj E", "CTR-100", "M-1001", "Dev",    15,  0, 45,  5],
        ["P-006", "Alertas",                 "HW", "Abierta",  "tresp3@x", 20,  8, "OK", "P3", "Obj F", "CTR-300", "M-1005", "Dev",     5,  2, 20, 10],
        ["P-007", "Migracion cloud",         "DF", "Abierta",  "tresp1@x", 100, 40,"OK", "P1", "Obj G", "CTR-100", "M-1001", "Dev",    50, 15,100, 60],
        ["P-008", "Seguridad",               "HE", "Cerrada",  "tresp2@x", 30, 30, "OK", "P2", "Obj H", "CTR-200", "M-1002", "Dev",     0,  0, 30, 30],
        ["P-009", "Perf mejora",             "EW", "Abierta",  "tresp2@x", 35, 20, "OK", "P2", "Obj I", "CTR-200", "M-1003", "Dev",    10,  5, 35, 25],
        ["P-010", "Tooling",                 "PE", "Abierta",  "tresp3@x", 12,  4, "OK", "P3", "Obj J", "CTR-300", "M-1006", "Dev",     4,  1, 12,  4],
        ["P-011", "ZZ-App nueva (sin mapeo)","ZZ", "Abierta",  "tresp3@x", 18,  6, "OK", "P3", "Obj K", "CTR-300", "M-1007", "Dev",     6,  2, 18,  7],
        ["P-012", "Refactor tests",          "WH", "Abierta",  "tresp1@x", 22, 10, "OK", "P1", "Obj L", "CTR-100", "M-1002", "Dev",     8,  3, 22, 11],
        ["P-013", "Import CSV",              "PF", "Abierta",  "tresp2@x", 15,  5, "OK", "P2", "Obj M", "CTR-200", "M-1008", "Dev",     5,  1, 15,  6],
        ["P-014", "Export Parquet",          "KE", "Abierta",  "tresp3@x", 28, 12, "OK", "P3", "Obj N", "CTR-300", "M-1006", "Dev",     7,  3, 28, 13],
    ]
    for r in rows:
        ws.append(r)

    # Fila 16: Peticion vacia → debe ser saltada por MesSheetBuilder (ancla vacia)
    skip_row = [""] + ["-"] * (len(EXTRACCION_HEADERS) - 1)
    ws.append(skip_row)

    wb.save(f"{OUT_DIR}/extraccion.xlsx")
    print("Generated extraccion.xlsx: 1 header row + 14 data rows + 1 skip row (Peticion vacia)")


# =========================================================================
# cierre.xlsx
# =========================================================================
# Perfil Cierre (config.properties): cabeceras en fila 2 (la 1 es titulo).
# 16 imputaciones con totales conocidos que los SUMIFS puedan cruzar.
CIERRE_HEADERS = [
    "Project Key", "Issue Key", "Aplicación BFA", "Labels",
    "Parent Issue Summary", "Summary", "Description", "Issue Type",
    "Time entry: User", "Time entry: Date", "Component Name",
    "Time entry: Description", "Matricula", "Funcion", "Account", "Hours",
]

def build_cierre():
    wb = Workbook()
    ws = wb.active
    ws.title = "Cierre"

    # Fila 1: metadata / titulo
    ws.append(["EXPORT JIRA - Closure Report"] + [None] * (len(CIERRE_HEADERS) - 1))
    # Fila 2: cabeceras
    ws.append(CIERRE_HEADERS)

    # 16 filas de imputaciones. Totales conocidos tras añadir la condición Funcion
    # (SUMIFS con match Component Name:Peticion, Matricula:Recurso, Funcion:Funcion):
    # Todas las filas de Extraccion tienen Funcion="Dev", asi que solo suman las
    # imputaciones de Cierre con Funcion="Dev". Las "Sup" quedan filtradas.
    #   P-001 + M-1001 + Dev -> PROJ-1 (3) + PROJ-2 (2) = 5 horas   (PROJ-3 con Sup queda fuera)
    #   P-002 + M-1002 + Dev -> PROJ-4 (5) = 5 horas
    #   P-003 + M-1003 + Dev -> PROJ-5 (10) + PROJ-6 (5) = 15 horas
    rows = [
        ["PROJ", "PROJ-1",  "DF", "back", "epic-1",  "sum1",  "desc",  "Task", "u1", "2026-01-02", "P-001", "t1", "M-1001", "Dev", "acc1",  3],
        ["PROJ", "PROJ-2",  "DF", "back", "epic-1",  "sum2",  "desc",  "Task", "u1", "2026-01-03", "P-001", "t2", "M-1001", "Dev", "acc1",  2],
        ["PROJ", "PROJ-3",  "DF", "back", "epic-1",  "sum3",  "desc",  "Task", "u1", "2026-01-04", "P-001", "t3", "M-1001", "Sup", "acc1",  4],
        ["PROJ", "PROJ-4",  "HE", "ui",   "epic-2",  "sum4",  "desc",  "Task", "u2", "2026-01-05", "P-002", "t4", "M-1002", "Dev", "acc1",  5],
        ["PROJ", "PROJ-5",  "EW", "api",  "epic-3",  "sum5",  "desc",  "Task", "u3", "2026-01-06", "P-003", "t5", "M-1003", "Dev", "acc2", 10],
        ["PROJ", "PROJ-6",  "EW", "api",  "epic-3",  "sum6",  "desc",  "Task", "u3", "2026-01-07", "P-003", "t6", "M-1003", "Dev", "acc2",  5],
        ["PROJ", "PROJ-7",  "J6", "dash", "epic-4",  "sum7",  "desc",  "Task", "u4", "2026-01-08", "P-004", "t7", "M-1004", "Dev", "acc2",  7],
        ["PROJ", "PROJ-8",  "JX", "sso",  "epic-5",  "sum8",  "desc",  "Task", "u1", "2026-01-09", "P-005", "t8", "M-1001", "Dev", "acc1",  2],
        ["PROJ", "PROJ-9",  "HW", "alrt", "epic-6",  "sum9",  "desc",  "Task", "u5", "2026-01-10", "P-006", "t9", "M-1005", "Dev", "acc3",  6],
        ["PROJ", "PROJ-10", "DF", "back", "epic-7",  "sum10", "desc",  "Task", "u1", "2026-01-11", "P-007", "t10","M-1001", "Dev", "acc1", 20],
        ["PROJ", "PROJ-11", "HE", "sec",  "epic-8",  "sum11", "desc",  "Task", "u2", "2026-01-12", "P-008", "t11","M-1002", "Dev", "acc1", 15],
        ["PROJ", "PROJ-12", "EW", "perf", "epic-9",  "sum12", "desc",  "Task", "u3", "2026-01-13", "P-009", "t12","M-1003", "Dev", "acc2",  8],
        ["PROJ", "PROJ-13", "PE", "tool", "epic-10", "sum13", "desc",  "Task", "u6", "2026-01-14", "P-010", "t13","M-1006", "Dev", "acc3",  4],
        ["PROJ", "PROJ-14", "WH", "test", "epic-12", "sum14", "desc",  "Task", "u2", "2026-01-15", "P-012", "t14","M-1002", "Dev", "acc1",  3],
        ["PROJ", "PROJ-15", "PF", "imp",  "epic-13", "sum15", "desc",  "Task", "u7", "2026-01-16", "P-013", "t15","M-1008", "Dev", "acc2",  1],
        ["PROJ", "PROJ-16", "KE", "exp",  "epic-14", "sum16", "desc",  "Task", "u6", "2026-01-17", "P-014", "t16","M-1006", "Dev", "acc3",  9],
    ]
    for r in rows:
        ws.append(r)

    wb.save(f"{OUT_DIR}/cierre.xlsx")
    print("Generated cierre.xlsx: 1 metadata row + 1 header row + 16 imputation rows")


if __name__ == "__main__":
    import os
    os.makedirs(OUT_DIR, exist_ok=True)
    build_extraccion()
    build_cierre()
