from __future__ import annotations

from pathlib import Path
from typing import Iterable

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
LOGO = ROOT / "media" / "logo-upt.png"
PY = ROOT

PROJECT = "Database Nexus: replicacion de datos entre bases de datos"
SYSTEM = "Database Nexus"
COURSE = "SI783 - Sistemas de Informacion"
TEACHER = "Docente del curso SI783"
TEAM = "Equipo Database Nexus - EPIS (SI783-2026-I)"
CITY_YEAR = "Tacna - Peru"
YEAR = "2026"
DATE = "19/06/2026"
VERSION = "1.0"

BLUE = RGBColor(31, 77, 120)
LIGHT_BLUE = RGBColor(46, 116, 181)
GRAY = "F2F4F7"


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_width(cell, width) -> None:
    cell.width = width
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.find(qn("w:tcW"))
    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)
    tc_w.set(qn("w:w"), str(int(width.inches * 1440)))
    tc_w.set(qn("w:type"), "dxa")


def apply_cell_text(cell, text: str, bold: bool = False) -> None:
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if len(text) < 25 else WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(9)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def table_style(table, widths: Iterable[float] | None = None, header: bool = True) -> None:
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    if widths:
        widths = list(widths)
        for row in table.rows:
            for idx, width in enumerate(widths):
                if idx < len(row.cells):
                    set_cell_width(row.cells[idx], Inches(width))
    if header and table.rows:
        for cell in table.rows[0].cells:
            set_cell_shading(cell, GRAY)
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.bold = True


def setup_doc(doc: Document) -> None:
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.header_distance = Inches(0.492)
    section.footer_distance = Inches(0.492)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Calibri"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    normal.font.size = Pt(11)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.1

    for name, size, color in [
        ("Heading 1", 16, LIGHT_BLUE),
        ("Heading 2", 13, LIGHT_BLUE),
        ("Heading 3", 12, BLUE),
    ]:
        style = styles[name]
        style.font.name = "Calibri"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
        style.font.size = Pt(size)
        style.font.color.rgb = color
        style.font.bold = True
        style.paragraph_format.space_before = Pt(12 if name != "Heading 1" else 16)
        style.paragraph_format.space_after = Pt(6)


def add_footer(doc: Document, label: str) -> None:
    for section in doc.sections:
        footer = section.footer.paragraphs[0]
        footer.text = f"{SYSTEM} | {label} | Version {VERSION}"
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in footer.runs:
            run.font.size = Pt(8)
            run.font.color.rgb = RGBColor(100, 100, 100)


def cover(doc: Document, title: str, subtitle: str = "") -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if LOGO.exists():
        p.add_run().add_picture(str(LOGO), width=Cm(2.6))
    for text, bold, size in [
        ("UNIVERSIDAD PRIVADA DE TACNA", True, 12),
        ("FACULTAD DE INGENIERIA", True, 12),
        ("Escuela Profesional de Ingenieria de Sistemas", True, 11),
        (title, True, 12),
        (f"Curso: {COURSE}", False, 11),
        (f"Docente: {TEACHER}", False, 11),
        ("Integrantes:", True, 11),
        (TEAM, True, 11),
        (CITY_YEAR, True, 11),
        (YEAR, True, 11),
    ]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(text)
        r.bold = bold
        r.font.size = Pt(size)
    if subtitle:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(subtitle)
        r.italic = True
    doc.add_page_break()


def version_table(doc: Document) -> None:
    add_h(doc, "CONTROL DE VERSIONES", 1)
    table = doc.add_table(rows=3, cols=6)
    rows = [
        ["Version", "Hecha por", "Revisada por", "Aprobada por", "Fecha", "Motivo"],
        [VERSION, "Equipo Database Nexus", "Docente SI783", "EPIS", DATE, "Version completa para entrega academica"],
        ["0.1", "Equipo Database Nexus", "Equipo", "Equipo", "12/06/2026", "Estructura inicial del documento"],
    ]
    for row, data in zip(table.rows, rows):
        for cell, text in zip(row.cells, data):
            apply_cell_text(cell, text, bold=row is table.rows[0])
    table_style(table, [0.7, 1.3, 1.3, 1.1, 0.9, 1.2])
    doc.add_page_break()


def add_h(doc: Document, text: str, level: int = 1) -> None:
    doc.add_heading(text, level=level)


def add_p(doc: Document, text: str) -> None:
    p = doc.add_paragraph(text)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def add_bullets(doc: Document, items: Iterable[str]) -> None:
    for item in items:
        doc.add_paragraph(item, style="List Bullet")


def add_numbered(doc: Document, items: Iterable[str]) -> None:
    for item in items:
        doc.add_paragraph(item, style="List Number")


def add_key_table(doc: Document, rows: list[tuple[str, str]]) -> None:
    table = doc.add_table(rows=1, cols=2)
    apply_cell_text(table.rows[0].cells[0], "Campo", True)
    apply_cell_text(table.rows[0].cells[1], "Descripcion", True)
    for key, value in rows:
        cells = table.add_row().cells
        apply_cell_text(cells[0], key, True)
        apply_cell_text(cells[1], value)
    table_style(table, [1.8, 4.7])


def add_matrix(doc: Document, headers: list[str], rows: list[list[str]], widths: list[float] | None = None) -> None:
    table = doc.add_table(rows=1, cols=len(headers))
    for cell, h in zip(table.rows[0].cells, headers):
        apply_cell_text(cell, h, True)
    for row in rows:
        cells = table.add_row().cells
        for cell, value in zip(cells, row):
            apply_cell_text(cell, value)
    table_style(table, widths)


def add_index(doc: Document, sections: list[str]) -> None:
    add_h(doc, "INDICE GENERAL", 1)
    for idx, title in enumerate(sections, start=1):
        doc.add_paragraph(f"{idx}. {title}")
    doc.add_page_break()


def common_context(doc: Document) -> None:
    add_h(doc, "Contexto del proyecto", 2)
    add_p(
        doc,
        "Database Nexus es una aplicacion web SaaS multiusuario orientada a configurar, "
        "supervisar y ejecutar replicaciones por lotes entre motores como PostgreSQL, "
        "MySQL, MariaDB, SQL Server, SQLite y MongoDB. El proyecto responde a la necesidad "
        "de pequenas y medianas organizaciones peruanas que operan con sistemas dispersos, "
        "archivos historicos y bases externas sin una integracion confiable.",
    )
    add_p(
        doc,
        "El escenario de referencia es una empresa comercial y de servicios de Tacna que "
        "necesita consolidar datos de ventas, inventario y clientes para reportes diarios, "
        "evitando reprocesos manuales en hojas de calculo y reduciendo errores en cierres "
        "operativos.",
    )


def save_doc(doc: Document, filename: str, label: str) -> None:
    add_footer(doc, label)
    doc.save(ROOT / filename)


def factibilidad() -> None:
    doc = Document()
    setup_doc(doc)
    cover(doc, f"Proyecto {PROJECT}")
    add_h(doc, f"Sistema {SYSTEM}", 1)
    add_p(doc, f"Informe de Factibilidad - Version {VERSION}")
    version_table(doc)
    sections = [
        "Descripcion del Proyecto",
        "Riesgos",
        "Analisis de la Situacion actual",
        "Estudio de Factibilidad",
        "Analisis Financiero",
        "Conclusiones",
    ]
    add_index(doc, sections)
    add_h(doc, "Informe de Factibilidad", 1)
    add_h(doc, "1. Descripcion del Proyecto", 1)
    add_key_table(doc, [
        ("Nombre del proyecto", PROJECT),
        ("Duracion", "4 meses academicos, de marzo a junio de 2026"),
        ("Lugar de referencia", "Tacna, Peru"),
        ("Modalidad", "Aplicacion web desplegable en Render/Supabase o infraestructura equivalente"),
    ])
    common_context(doc)
    add_h(doc, "Objetivo general", 2)
    add_p(doc, "Implementar una plataforma web que permita replicar datos entre bases heterogeneas de forma controlada, auditable y segura para mejorar la disponibilidad de informacion operativa.")
    add_h(doc, "Objetivos especificos", 2)
    add_bullets(doc, [
        "Configurar conexiones aisladas por usuario con credenciales cifradas.",
        "Permitir seleccion de tablas, mapeo de columnas y modos insertar, upsert, reemplazar o recargar.",
        "Registrar avance, errores y reanudacion de trabajos interrumpidos.",
        "Generar monitoreo de salud y documentacion basica de esquemas para apoyo tecnico.",
    ])
    add_h(doc, "2. Riesgos", 1)
    add_matrix(doc, ["Riesgo", "Impacto", "Mitigacion"], [
        ["Credenciales externas mal gestionadas", "Alto", "Cifrado AES-256-GCM, variables de entorno y no retorno de secretos al cliente."],
        ["Cambios de esquema durante la replicacion", "Medio", "Validacion previa, comparacion de esquemas y registro de errores por tabla."],
        ["Limitaciones de red en servicios cloud", "Medio", "Timeouts, reintentos y recomendacion de VPN o listas blancas cuando aplique."],
        ["Volumen de datos mayor al esperado", "Medio", "Lotes de 5000 filas, limite de carga configurable y pruebas con muestras representativas."],
        ["Uso de datos personales", "Alto", "Consentimiento, minimizacion de datos y alineamiento con Ley 29733."],
    ], [1.7, 1.1, 3.7])
    add_h(doc, "3. Analisis de la Situacion actual", 1)
    add_h(doc, "Planteamiento del problema", 2)
    add_p(doc, "Muchas pymes peruanas mantienen informacion en aplicaciones separadas: ventas en MySQL, inventario en SQL Server, reportes historicos en SQLite o archivos importados y datos de clientes en sistemas externos. La integracion suele realizarse de forma manual, con copias CSV y macros, generando retrasos, duplicidad de registros y baja trazabilidad.")
    add_h(doc, "Consideraciones de hardware y software", 2)
    add_bullets(doc, [
        "Clientes: navegadores modernos en laptops de gama media con conexion estable a internet.",
        "Backend: Node.js 20 con Fastify, API REST y autenticacion JWT.",
        "Persistencia: PostgreSQL/Supabase para metadatos, trabajos y catalogos comprimidos.",
        "Motores soportados: PostgreSQL, MySQL, MariaDB, SQL Server, SQLite y MongoDB.",
        "Despliegue viable: Render, VPS nacional o nube con cumplimiento de politicas internas.",
    ])
    add_h(doc, "4. Estudio de Factibilidad", 1)
    add_p(doc, "La evaluacion se realizo considerando tecnologia disponible, costos realistas de desarrollo academico-profesional, capacidad operativa de una pyme tacnena y marco normativo peruano.")
    add_h(doc, "Factibilidad Tecnica", 2)
    add_p(doc, "El proyecto es tecnicamente factible porque utiliza tecnologias maduras, documentadas y con disponibilidad local de talento: React, Node.js, PostgreSQL, control de acceso JWT, cifrado de credenciales y despliegue cloud. La infraestructura requerida no exige servidores propios; puede operar con servicios gestionados y conexiones salientes autorizadas.")
    add_h(doc, "Factibilidad Economica", 2)
    add_matrix(doc, ["Categoria", "Detalle", "Costo estimado S/"], [
        ["Costos generales", "Internet, energia, utiles y pruebas de conectividad", "480"],
        ["Operativos de desarrollo", "Ambiente cloud, dominio, backups y monitoreo basico", "1,300"],
        ["Ambiente tecnico", "Supabase/Render, pruebas con motores externos y almacenamiento temporal", "1,200"],
        ["Personal", "Analista, desarrollador backend, frontend, QA y lider tecnico", "10,000"],
        ["Total", "Costo directo de desarrollo", "12,980"],
    ], [1.4, 3.6, 1.5])
    add_h(doc, "Costos de personal", 3)
    add_matrix(doc, ["Rol", "Horas", "Tarifa S/", "Subtotal S/"], [
        ["Analista funcional", "35", "35", "1,225"],
        ["Desarrollador backend", "90", "40", "3,600"],
        ["Desarrollador frontend", "65", "35", "2,275"],
        ["QA/Documentacion", "45", "30", "1,350"],
        ["Gestion tecnica", "45", "35", "1,575"],
    ], [2.1, 1.0, 1.0, 1.2])
    add_p(doc, "Forma de pago referencial: 40% al aprobar alcance, 40% al entregar version funcional y 20% tras pruebas de aceptacion.")
    add_h(doc, "Factibilidad Operativa", 2)
    add_p(doc, "La operacion es viable porque el sistema concentra las tareas en un panel web con perfiles de usuario, configuraciones aisladas y monitoreo de trabajos. El personal tecnico del cliente requiere capacitacion breve de 4 horas para administrar conexiones, revisar logs y ejecutar respaldos.")
    add_h(doc, "Factibilidad Legal", 2)
    add_p(doc, "El proyecto no presenta conflicto legal si se aplican contratos de confidencialidad, autorizacion de tratamiento de datos y medidas de seguridad conforme a la Ley N. 29733, Ley de Proteccion de Datos Personales, y su reglamento. Para datos tributarios o laborales se recomienda limitar acceso por rol y mantener trazabilidad.")
    add_h(doc, "Factibilidad Social", 2)
    add_p(doc, "El impacto social es favorable porque reduce trabajo repetitivo y permite que el personal se enfoque en analisis y atencion al cliente. Debe comunicarse que la plataforma apoya la operacion y no reemplaza decisiones humanas ni responsabilidades de control interno.")
    add_h(doc, "Factibilidad Ambiental", 2)
    add_p(doc, "El impacto ambiental es bajo. Al disminuir impresiones, traslados para consolidar informacion y uso de equipos encendidos para tareas manuales, el proyecto contribuye a una operacion digital mas eficiente.")
    add_h(doc, "5. Analisis Financiero", 1)
    add_matrix(doc, ["Indicador", "Valor", "Interpretacion"], [
        ["Inversion inicial", "S/ 12,980", "Desarrollo, pruebas y despliegue inicial."],
        ["Beneficio anual estimado", "S/ 18,600", "Ahorro por reduccion de reprocesos, errores y horas de consolidacion."],
        ["Relacion B/C", "1.43", "Mayor a 1, el proyecto es aceptable."],
        ["VAN referencial", "S/ 3,120", "Calculado con COK anual de 12% y beneficios netos conservadores."],
        ["TIR referencial", "24%", "Superior al costo de oportunidad estimado."],
    ], [1.6, 1.5, 3.4])
    add_h(doc, "Beneficios", 2)
    add_bullets(doc, [
        "Reduccion de 8 a 2 horas semanales en consolidacion manual.",
        "Menor riesgo de errores al copiar datos entre sistemas.",
        "Mejor disponibilidad de informacion para reportes diarios.",
        "Mayor trazabilidad ante auditorias internas o requerimientos de gerencia.",
    ])
    add_h(doc, "6. Conclusiones", 1)
    add_p(doc, "El proyecto Database Nexus es viable tecnica, economica, operativa, legal, social y ambientalmente para el escenario peruano evaluado. La inversion es moderada, la tecnologia es alcanzable y los beneficios esperados superan los costos bajo supuestos conservadores.")
    save_doc(doc, "FD01-EPIS-Informe de Factibilidad.docx", "Informe de Factibilidad")


def vision() -> None:
    doc = Document()
    setup_doc(doc)
    cover(doc, f"Proyecto {PROJECT}")
    add_h(doc, f"Sistema {SYSTEM}", 1)
    add_p(doc, f"Documento de Vision - Version {VERSION}")
    version_table(doc)
    sections = ["Introduccion", "Posicionamiento", "Descripcion de interesados y usuarios", "Vista General del Producto", "Caracteristicas", "Restricciones", "Rangos de calidad", "Precedencia y Prioridad", "Otros requerimientos", "Conclusiones"]
    add_index(doc, sections)
    add_h(doc, "Informe de Vision", 1)
    add_h(doc, "1. Introduccion", 1)
    add_h(doc, "Proposito", 2)
    add_p(doc, "Definir la vision del producto Database Nexus, sus usuarios, capacidades principales y restricciones para orientar el desarrollo durante el ciclo academico 2026-I.")
    add_h(doc, "Alcance", 2)
    add_p(doc, "Incluye autenticacion, gestion de configuraciones, importacion limitada de archivos de base de datos, replicacion por lotes, monitoreo, documentacion de esquemas y administracion basica de usuarios.")
    add_h(doc, "Definiciones, siglas y abreviaturas", 2)
    add_key_table(doc, [("SaaS", "Software como servicio accesible por navegador."), ("ETL", "Extraccion, transformacion y carga de datos."), ("Upsert", "Operacion que inserta o actualiza registros segun clave."), ("JWT", "Token de autenticacion para sesiones web."), ("COK", "Costo de oportunidad de capital.")])
    add_h(doc, "Referencias", 2)
    add_bullets(doc, ["Ley N. 29733 - Proteccion de Datos Personales.", "Buenas practicas OWASP para aplicaciones web.", "Documentacion tecnica de PostgreSQL, MySQL, SQL Server y MongoDB."])
    add_h(doc, "Vision general", 2)
    common_context(doc)
    add_h(doc, "2. Posicionamiento", 1)
    add_h(doc, "Oportunidad de negocio", 2)
    add_p(doc, "En Peru, muchas pymes adoptan soluciones cloud, pero aun conviven con sistemas legados y bases locales. Database Nexus ofrece una alternativa ligera para integrar informacion sin adquirir plataformas empresariales costosas.")
    add_h(doc, "Definicion del problema", 2)
    add_matrix(doc, ["Elemento", "Descripcion"], [
        ["Problema", "Consolidacion manual de datos entre sistemas heterogeneos."],
        ["Afecta a", "Analistas, administradores, gerencia y personal operativo."],
        ["Impacto", "Retrasos, duplicidad, errores y baja trazabilidad."],
        ["Solucion", "Plataforma web con replicacion configurada, monitoreo y logs."],
    ], [1.7, 4.8])
    add_h(doc, "3. Descripcion de los interesados y usuarios", 1)
    add_matrix(doc, ["Interesado", "Necesidad", "Beneficio esperado"], [
        ["Gerencia", "Reportes confiables y oportunos", "Decisiones con datos consistentes."],
        ["Administrador TI", "Controlar conexiones y ejecuciones", "Menor soporte manual y mayor trazabilidad."],
        ["Analista de datos", "Preparar informacion para reportes", "Menos tiempo de limpieza y copiado."],
        ["Usuario operativo", "Consultar datos actualizados", "Menos duplicidad en registros diarios."],
    ], [1.4, 2.5, 2.6])
    add_h(doc, "Entorno de usuario", 2)
    add_p(doc, "Los usuarios acceden desde navegadores Chrome, Edge o Firefox. El administrador trabaja con paneles protegidos, mientras los usuarios comunes crean configuraciones, seleccionan tablas y revisan ejecuciones propias.")
    add_h(doc, "Necesidades de interesados y usuarios", 2)
    add_bullets(doc, ["Aislamiento de informacion por usuario.", "Configuracion simple de origen y destino.", "Ejecuciones reanudables y con progreso visible.", "Mensajes de error entendibles.", "Exportacion o consulta de documentacion de esquemas."])
    add_h(doc, "4. Vista General del Producto", 1)
    add_h(doc, "Perspectiva del producto", 2)
    add_p(doc, "Database Nexus funciona como una aplicacion web centralizada que orquesta conexiones temporales hacia bases externas y conserva metadatos en PostgreSQL/Supabase.")
    add_h(doc, "Resumen de capacidades", 2)
    add_matrix(doc, ["Capacidad", "Descripcion"], [
        ["Autenticacion", "Registro, inicio de sesion y JWT de 7 dias."],
        ["Configuraciones", "Alta, validacion y cifrado de credenciales por usuario."],
        ["Replicacion", "Seleccion multiple de tablas, mapeo y modos insertar/upsert/reemplazar/recargar."],
        ["Monitoreo", "Health monitor con disponibilidad y latencia."],
        ["Documentador", "Perfil de columnas, comparacion de esquemas y ocultamiento de campos sensibles."],
    ], [1.5, 5.0])
    add_h(doc, "Suposiciones y dependencias", 2)
    add_bullets(doc, ["Disponibilidad de internet estable.", "Credenciales autorizadas por el cliente.", "Bases externas accesibles desde el entorno de despliegue.", "Volumen inicial dentro del limite configurado de carga."])
    add_h(doc, "Costos y precios", 2)
    add_p(doc, "Para el piloto academico se estima una inversion de S/ 12,980. En un escenario comercial, podria ofrecerse como servicio mensual desde S/ 180 para pymes, mas costos cloud segun volumen.")
    add_h(doc, "Licenciamiento e instalacion", 2)
    add_p(doc, "El sistema puede desplegarse en infraestructura cloud con dependencias open source y servicios gestionados. El cliente debe aceptar politicas de uso, confidencialidad y tratamiento de datos.")
    add_h(doc, "5. Caracteristicas del producto", 1)
    add_bullets(doc, ["Panel web responsivo.", "Cifrado de credenciales.", "Soporte multi-motor.", "Replicacion por lotes con reintentos.", "Logs persistentes.", "Administracion de usuarios.", "Limpieza automatica de configuraciones temporales."])
    add_h(doc, "6. Restricciones", 1)
    add_bullets(doc, ["No implementa CDC exact-once en esta version.", "La replicacion incremental por offset requiere origen estable.", "Los respaldos .bak requieren SQL Server externo configurado.", "El limite inicial de archivo es configurable y sugerido en 500 MB."])
    add_h(doc, "7. Rangos de Calidad", 1)
    add_matrix(doc, ["Atributo", "Meta"], [["Seguridad", "Credenciales cifradas y acceso por rol."], ["Disponibilidad", "99% en horario laboral para piloto."], ["Rendimiento", "Lotes de 5000 filas sin bloquear la interfaz."], ["Usabilidad", "Usuario capacitado debe crear configuracion basica en menos de 10 minutos."], ["Mantenibilidad", "Servicios separados por rutas y modulos backend."]], [1.6, 4.9])
    add_h(doc, "8. Precedencia y Prioridad", 1)
    add_numbered(doc, ["Autenticacion y aislamiento por usuario.", "Gestion segura de configuraciones.", "Replicacion basica entre motores relacionales.", "Monitoreo y reanudacion.", "Documentador y mejoras de experiencia."])
    add_h(doc, "9. Otros requerimientos del producto", 1)
    add_bullets(doc, ["Cumplimiento de Ley 29733 cuando se traten datos personales.", "Comunicacion HTTPS en produccion.", "Registro de auditoria minimo para ejecuciones.", "Politica de retencion de configuraciones temporales por 24 horas."])
    add_h(doc, "Conclusiones", 1)
    add_p(doc, "La vision del producto es coherente con una necesidad real del mercado peruano y con el alcance academico: una plataforma practica, segura y extensible para replicacion de datos heterogeneos.")
    save_doc(doc, "FD02-EPIS-Informe Vision.docx", "Documento de Vision")


def srs() -> None:
    doc = Document()
    setup_doc(doc)
    cover(doc, f"Proyecto {PROJECT}")
    add_h(doc, f"Sistema {SYSTEM}", 1)
    add_p(doc, f"Documento de Especificacion de Requerimientos de Software - Version {VERSION}")
    version_table(doc)
    sections = ["Introduccion", "Descripcion general", "Requerimientos funcionales", "Requerimientos no funcionales", "Interfaces externas", "Modelo de datos", "Criterios de aceptacion"]
    add_index(doc, sections)
    add_h(doc, "1. Introduccion", 1)
    add_p(doc, "Este documento especifica los requerimientos de software de Database Nexus para orientar la construccion, pruebas y validacion de la aplicacion.")
    add_h(doc, "2. Descripcion general", 1)
    common_context(doc)
    add_h(doc, "Perspectiva del producto", 2)
    add_p(doc, "El sistema se concibe como una plataforma independiente que no reemplaza a los sistemas transaccionales existentes. Su funcion es actuar como capa de integracion controlada, permitiendo que un usuario autorizado defina origen, destino, tablas, mapeos y modo de carga sin modificar directamente la logica de los sistemas fuente.")
    add_h(doc, "Funciones principales", 2)
    add_bullets(doc, ["Administracion de cuentas y perfiles.", "Registro seguro de conexiones externas.", "Lectura de catalogos, tablas y columnas.", "Configuracion de reglas de replicacion.", "Ejecucion de trabajos con seguimiento de progreso.", "Consulta de errores, historial y estado de salud.", "Documentacion tecnica de esquemas importados."])
    add_h(doc, "Caracteristicas de los usuarios", 2)
    add_p(doc, "Los usuarios esperados tienen conocimiento basico de bases de datos, pero no necesariamente experiencia en administracion avanzada. Por ello, la interfaz debe guiar la configuracion, validar entradas y presentar mensajes de error accionables, especialmente en casos de credenciales incorrectas, red no disponible o incompatibilidad de tipos.")
    add_h(doc, "Actores", 2)
    add_matrix(doc, ["Actor", "Responsabilidad"], [["Administrador", "Gestiona usuarios y revisa estado general sin acceder a credenciales."], ["Usuario registrado", "Crea configuraciones, ejecuta replicaciones y consulta resultados."], ["Servicio de replicacion", "Procesa trabajos por lotes, reintenta y registra progreso."], ["Base externa", "Origen o destino autorizado por el usuario."]], [1.8, 4.7])
    add_h(doc, "3. Requerimientos funcionales", 1)
    rf = [
        ["RF-01", "Registrar e iniciar sesion de usuarios mediante credenciales validas.", "Alta"],
        ["RF-02", "Cifrar credenciales de conexiones externas antes de almacenarlas.", "Alta"],
        ["RF-03", "Crear, listar, editar y eliminar configuraciones por usuario.", "Alta"],
        ["RF-04", "Probar conectividad con bases soportadas antes de guardar configuracion.", "Alta"],
        ["RF-05", "Seleccionar tablas y columnas origen/destino para replicacion.", "Alta"],
        ["RF-06", "Ejecutar replicacion en modos insertar, upsert, reemplazar y recargar.", "Alta"],
        ["RF-07", "Mostrar progreso, registros procesados, errores y estado final.", "Alta"],
        ["RF-08", "Reanudar trabajos pendientes tras reinicio del servidor.", "Media"],
        ["RF-09", "Programar ejecuciones incrementales por offset.", "Media"],
        ["RF-10", "Generar documentacion de esquema y ocultar campos sensibles.", "Media"],
        ["RF-11", "Administrar usuarios desde un panel de administrador.", "Media"],
        ["RF-12", "Eliminar configuraciones temporales vencidas despues de 24 horas.", "Alta"],
    ]
    add_matrix(doc, ["ID", "Descripcion", "Prioridad"], rf, [0.8, 4.8, 0.9])
    add_h(doc, "4. Requerimientos no funcionales", 1)
    rnf = [
        ["RNF-01", "Seguridad", "JWT, rate limit, Helmet, cifrado AES-256-GCM y CORS controlado."],
        ["RNF-02", "Rendimiento", "Procesamiento por lotes de 5000 filas con timeout de conexion de 5 segundos."],
        ["RNF-03", "Disponibilidad", "Despliegue cloud con health check /api/status."],
        ["RNF-04", "Usabilidad", "Interfaz web clara para configurar una replicacion basica en menos de 10 minutos."],
        ["RNF-05", "Mantenibilidad", "Backend modular por rutas y servicios; frontend con componentes reutilizables."],
        ["RNF-06", "Legal", "Tratamiento de datos personales conforme a Ley 29733 cuando corresponda."],
    ]
    add_matrix(doc, ["ID", "Atributo", "Especificacion"], rnf, [0.8, 1.2, 4.5])
    add_h(doc, "5. Interfaces externas", 1)
    add_bullets(doc, ["API REST bajo /api para autenticacion, usuarios, configuraciones, replicaciones, salud y esquema.", "Interfaz web React para usuarios y administradores.", "Conectores hacia PostgreSQL, MySQL, MariaDB, SQL Server, SQLite y MongoDB.", "Variables de entorno para DATABASE_URL, JWT_SECRET, ENCRYPTION_KEY y parametros de restauracion SQL Server."])
    add_h(doc, "Supuestos y dependencias", 2)
    add_bullets(doc, ["El cliente proporciona credenciales validas y autorizadas para cada base externa.", "Las bases origen y destino son accesibles desde el entorno donde se despliega el backend.", "El volumen inicial de datos se mantiene dentro de los limites configurados para el piloto.", "La organizacion define politicas de retencion y uso de datos antes de operar en produccion."])
    add_h(doc, "Reglas de negocio", 2)
    add_bullets(doc, ["Una configuracion pertenece a un unico usuario y no puede compartirse sin autorizacion explicita.", "Las credenciales almacenadas nunca se muestran nuevamente en texto claro.", "Las ejecuciones fallidas deben conservar el detalle suficiente para diagnostico tecnico.", "Una configuracion temporal vencida se elimina automaticamente despues de 24 horas.", "El administrador puede ver indicadores generales, pero no secretos de conexiones."])
    add_h(doc, "6. Modelo de datos", 1)
    add_matrix(doc, ["Entidad", "Descripcion"], [["users", "Usuarios, roles y datos de acceso."], ["configurations", "Conexiones cifradas, motor, catalogos y propietario."], ["replication_jobs", "Estado, progreso, modo de ejecucion y errores."], ["health_checks", "Disponibilidad y latencia por configuracion."], ["schema_snapshots", "Metadatos comprimidos para documentacion y comparacion."]], [1.7, 4.8])
    add_h(doc, "7. Criterios de aceptacion", 1)
    add_bullets(doc, ["Un usuario no puede ver configuraciones de otro usuario.", "Las credenciales no se exponen en respuestas del API.", "Una replicacion de prueba entre dos tablas finaliza con conteo consistente.", "Los errores de conectividad se muestran de forma comprensible.", "El administrador puede visualizar usuarios sin acceder a secretos."])
    add_h(doc, "Trazabilidad de requerimientos", 2)
    add_matrix(doc, ["Necesidad", "Requerimientos asociados", "Prueba sugerida"], [["Seguridad de credenciales", "RF-02, RNF-01", "Inspeccionar respuesta API y verificar que no retorna contrasenas."], ["Replicacion confiable", "RF-05, RF-06, RF-07", "Ejecutar carga de tabla de muestra y comparar conteos."], ["Operacion recuperable", "RF-08, RNF-03", "Interrumpir servicio y validar reanudacion del trabajo."], ["Cumplimiento peruano", "RNF-06", "Revisar consentimiento, roles y minimizacion de datos personales."]], [2.0, 2.1, 2.4])
    save_doc(doc, "FD03-EPIS-Informe Especificación Requerimientos.docx", "SRS")


def arquitectura() -> None:
    doc = Document()
    setup_doc(doc)
    cover(doc, f"Proyecto {PROJECT}")
    add_h(doc, f"Sistema {SYSTEM}", 1)
    add_p(doc, f"Documento de Arquitectura de Software - Version {VERSION}")
    version_table(doc)
    sections = ["Introduccion", "Objetivos y restricciones arquitectonicas", "Representacion 4+1", "Vista logica", "Vista de procesos", "Vista de desarrollo", "Vista fisica", "Escenarios", "Decisiones arquitectonicas"]
    add_index(doc, sections)
    add_h(doc, "1. Introduccion", 1)
    add_h(doc, "Proposito (Diagrama 4+1)", 2)
    add_p(doc, "Presentar la arquitectura de Database Nexus usando el modelo 4+1, relacionando vistas con requerimientos funcionales, atributos de calidad y decisiones de diseno.")
    add_h(doc, "Alcance", 2)
    add_p(doc, "El documento cubre frontend, backend, persistencia, conectores de bases de datos, seguridad, despliegue cloud y escenarios relevantes de replicacion.")
    add_h(doc, "Definiciones", 2)
    add_key_table(doc, [("Vista logica", "Modulos y responsabilidades funcionales."), ("Vista de procesos", "Flujos de ejecucion y concurrencia."), ("Vista de desarrollo", "Organizacion del codigo y dependencias."), ("Vista fisica", "Nodos de despliegue y conectividad."), ("Escenarios", "Casos de uso que validan la arquitectura.")])
    add_h(doc, "2. Objetivos y restricciones arquitectonicas", 1)
    add_matrix(doc, ["Objetivo", "Restriccion asociada"], [["Seguridad de credenciales", "Uso obligatorio de ENCRYPTION_KEY y no exposicion de secretos."], ["Aislamiento multiusuario", "Todas las configuraciones se filtran por user_id."], ["Portabilidad", "Node.js 20 y despliegue en Render/VPS con variables de entorno."], ["Rendimiento controlado", "Lotes y limites de carga configurables."], ["Mantenibilidad", "Separacion por rutas, servicios y utilidades."]], [2.5, 4.0])
    add_h(doc, "Priorizacion de requerimientos", 2)
    add_matrix(doc, ["ID", "Descripcion", "Prioridad"], [["RF-02", "Cifrado de credenciales", "Alta"], ["RF-06", "Ejecucion de replicacion", "Alta"], ["RNF-01", "Seguridad", "Alta"], ["RNF-02", "Rendimiento por lotes", "Alta"], ["RF-10", "Documentador de esquemas", "Media"]], [0.8, 4.8, 0.9])
    add_h(doc, "3. Representacion de la arquitectura del sistema", 1)
    add_p(doc, "La arquitectura se organiza alrededor de una API central que controla identidad, reglas de acceso y orquestacion de trabajos. El frontend no accede directamente a bases externas; todas las operaciones sensibles pasan por el backend, donde se aplican validaciones, cifrado, limites de carga y registro de eventos.")
    add_h(doc, "Vista de casos de uso", 2)
    add_matrix(doc, ["Caso de uso", "Actor", "Resultado"], [["Autenticarse", "Usuario", "Sesion JWT valida."], ["Configurar conexion", "Usuario", "Credenciales cifradas y motor validado."], ["Ejecutar replicacion", "Usuario", "Trabajo por lotes con progreso persistente."], ["Monitorear salud", "Usuario", "Disponibilidad y latencia recientes."], ["Administrar usuarios", "Administrador", "Vista de usuarios y estado general."]], [2.0, 1.4, 3.1])
    add_h(doc, "Vista logica", 2)
    add_matrix(doc, ["Subsistema", "Responsabilidad"], [["Frontend React", "Rutas, formularios, paneles y experiencia de usuario."], ["API Fastify", "Autenticacion, validacion, seguridad HTTP y exposicion REST."], ["Servicio de configuraciones", "Gestiona conexiones, cifrado y catalogos."], ["Servicio de replicacion", "Orquesta lectura, transformacion, escritura, reintentos y progreso."], ["Servicio de salud", "Ejecuta comprobaciones de disponibilidad y latencia."], ["Persistencia PostgreSQL", "Usuarios, metadatos, trabajos y snapshots."]], [2.0, 4.5])
    add_h(doc, "Vista de procesos", 2)
    add_numbered(doc, ["El usuario inicia sesion y obtiene JWT.", "Crea una configuracion; el backend valida conectividad y cifra secretos.", "Selecciona tablas y modo de replicacion.", "El servicio de replicacion procesa lotes, actualiza progreso y registra errores.", "Si el servidor reinicia, los trabajos pendientes se reanudan al arrancar."])
    add_p(doc, "El procesamiento por lotes evita mantener transacciones largas y permite reportar avance incremental. Cuando ocurre un error, el trabajo se marca con detalle tecnico y se conserva la informacion necesaria para repetir la ejecucion sin perder el diagnostico.")
    add_h(doc, "Vista de desarrollo", 2)
    add_bullets(doc, ["backend/src/routes: endpoints de dominio.", "backend/src/services: logica de replicacion, expiracion y salud.", "backend/src/db: inicializacion y persistencia.", "frontend/src/pages: vistas principales.", "frontend/src/components: rutas privadas y componentes reutilizables."])
    add_h(doc, "Vista fisica", 2)
    add_matrix(doc, ["Nodo", "Tecnologia", "Funcion"], [["Cliente", "Navegador web", "Uso de interfaz React."], ["Servicio web", "Node.js/Fastify en Render o VPS", "API y servicio estatico frontend."], ["Base de metadatos", "PostgreSQL/Supabase", "Usuarios, trabajos y catalogos."], ["Bases externas", "PostgreSQL/MySQL/SQL Server/SQLite/MongoDB", "Origen o destino de datos."], ["SQL Server auxiliar", "Instancia externa opcional", "Restauracion temporal de .bak."]], [1.4, 2.2, 2.9])
    add_h(doc, "Escenarios arquitectonicos", 2)
    add_bullets(doc, ["Replicar clientes desde MySQL a PostgreSQL para reportes.", "Reanudar una ejecucion interrumpida por reinicio del servidor.", "Bloquear acceso de un usuario a configuraciones de otro.", "Detectar cambio de esquema antes de ejecutar una carga completa."])
    add_h(doc, "Seguridad", 2)
    add_p(doc, "La seguridad se aborda por capas: Helmet endurece cabeceras HTTP, CORS limita origenes permitidos, JWT controla sesiones, rate limit reduce abuso por IP y AES-256-GCM protege credenciales externas. En produccion se requiere HTTPS, rotacion de secretos y segregacion de ambientes.")
    add_h(doc, "Persistencia y datos", 2)
    add_p(doc, "PostgreSQL/Supabase almacena metadatos, no busca convertirse en repositorio universal de datos replicados. Los catalogos y snapshots se comprimen para reducir consumo; los archivos cargados se usan de forma temporal y se eliminan al finalizar o vencer el periodo de retencion.")
    add_h(doc, "Escalabilidad y evolucion", 2)
    add_p(doc, "La primera version prioriza lotes y reanudacion. Una evolucion natural seria incorporar colas persistentes, workers separados, CDC por WAL/binlog/change streams y alertas por correo o mensajeria para operaciones criticas.")
    add_h(doc, "Decisiones arquitectonicas", 1)
    add_matrix(doc, ["Decision", "Justificacion"], [["Fastify sobre Express", "Menor sobrecarga, buen soporte TypeScript y plugins de seguridad."], ["PostgreSQL/Supabase para metadatos", "Servicio gestionado, backups y compatibilidad cloud."], ["Cifrado AES-256-GCM", "Protege credenciales externas con autenticidad e integridad."], ["Replicacion por lotes", "Adecuada para alcance academico y pymes sin CDC avanzado."], ["Frontend React", "Rapidez de desarrollo y componentes mantenibles."]], [2.2, 4.3])
    save_doc(doc, "FD04-EPIS-Informe Arquitectura de Software.docx", "Arquitectura")


def propuesta() -> None:
    doc = Document()
    setup_doc(doc)
    cover(doc, f"Propuesta del Proyecto {PROJECT}")
    version_table(doc)
    add_h(doc, "Proyecto", 1)
    add_key_table(doc, [("Nombre", f"{PROJECT}, Tacna, {YEAR}"), ("Presentado por", TEAM), ("Cargo", "Equipo de analisis y desarrollo"), ("Fecha", DATE)])
    add_h(doc, "Resumen Ejecutivo", 1)
    add_p(doc, "La propuesta plantea desarrollar Database Nexus, una plataforma web para replicar datos entre bases heterogeneas usadas por pymes peruanas. El sistema permitira reducir reprocesos manuales, mejorar la trazabilidad y disponer de informacion actualizada para reportes operativos.")
    add_h(doc, "Ficha del proyecto", 1)
    add_matrix(doc, ["Campo", "Valor"], [["Nombre del Proyecto propuesto", f"{PROJECT}, Tacna, {YEAR}"], ["Proposito", "Automatizar la replicacion controlada entre bases de datos para consolidar informacion operativa."], ["Resultados esperados", "Aplicacion web funcional, API segura, replicacion por lotes, monitoreo y documentacion de esquemas."], ["Poblacion objetivo", "Pymes comerciales, areas TI, analistas y estudiantes que requieren integracion de datos."], ["Monto de inversion", "S/ 12,980"], ["Duracion", "4 meses"]], [2.2, 4.3])
    add_h(doc, "Justificacion", 1)
    add_p(doc, "En Tacna y otras regiones del Peru, empresas comerciales y de servicios suelen operar con sistemas independientes. La falta de integracion genera costos ocultos por horas de consolidacion, errores de digitacion y baja oportunidad en reportes. Una solucion web de alcance moderado es una alternativa realista frente a herramientas empresariales costosas.")
    add_p(doc, "El proyecto tambien tiene valor academico porque integra analisis de requerimientos, seguridad, arquitectura web, bases de datos y evaluacion de factibilidad. La propuesta permite demostrar una solucion aplicable al contexto regional sin sobredimensionar infraestructura ni costos.")
    add_h(doc, "Objetivos", 1)
    add_bullets(doc, ["Implementar autenticacion y aislamiento multiusuario.", "Permitir conexiones seguras hacia motores de bases de datos comunes.", "Ejecutar replicaciones configurables y monitoreadas.", "Documentar esquemas y facilitar comparaciones basicas.", "Validar el sistema con escenarios representativos de una pyme peruana."])
    add_h(doc, "Cronograma referencial", 1)
    add_matrix(doc, ["Fase", "Periodo", "Entregable"], [["Analisis", "Marzo 2026", "Vision, factibilidad y SRS."], ["Diseno", "Abril 2026", "Arquitectura y prototipo navegable."], ["Construccion", "Mayo 2026", "Backend, frontend y replicacion basica."], ["Pruebas", "Junio 2026", "QA, documentacion e informe final."]], [1.6, 1.4, 3.5])
    add_h(doc, "Presupuesto", 1)
    add_matrix(doc, ["Concepto", "Monto S/"], [["Personal de desarrollo", "10,000"], ["Servicios cloud y dominios", "1,300"], ["Pruebas e infraestructura temporal", "1,200"], ["Costos generales", "480"], ["Total", "12,980"]], [4.5, 2.0])
    add_h(doc, "Resultados esperados", 1)
    add_bullets(doc, ["Sistema web desplegable.", "Documentos academicos completos.", "Pruebas de replicacion con datos de muestra.", "Manual breve para usuario tecnico.", "Recomendaciones para continuidad del proyecto."])
    add_h(doc, "Alcance y limites", 1)
    add_p(doc, "El alcance cubre un piloto funcional con autenticacion, configuracion, replicacion por lotes y monitoreo. No se compromete alta disponibilidad empresarial, replicacion en tiempo real exact-once ni conectividad con sistemas que no puedan exponerse de forma segura al backend.")
    add_h(doc, "Criterios de exito", 1)
    add_bullets(doc, ["Replicar una tabla de prueba entre dos motores soportados.", "Mantener credenciales cifradas y ocultas en respuestas del API.", "Registrar progreso y errores durante una ejecucion.", "Permitir que un usuario tecnico complete una configuracion basica con capacitacion breve."])
    save_doc(doc, "FD06-EPIS-PropuestaProyecto.docx", "Propuesta")


def final_report() -> None:
    doc = Document()
    setup_doc(doc)
    cover(doc, "Informe Final", f"Proyecto {PROJECT}")
    version_table(doc)
    sections = ["Antecedentes", "Planteamiento del problema", "Justificacion", "Alcance", "Objetivos", "Marco teorico", "Desarrollo de la solucion", "Factibilidad", "Tecnologia", "Metodologia", "Cronograma", "Presupuesto", "Conclusiones", "Recomendaciones", "Bibliografia", "Anexos"]
    add_index(doc, sections)
    add_h(doc, "Antecedentes", 1)
    add_p(doc, "La integracion de datos es una necesidad recurrente en organizaciones que crecen incorporando sistemas por etapas. En pymes peruanas es frecuente encontrar bases de datos distintas para ventas, inventario, clientes y reportes, lo que obliga a consolidaciones manuales.")
    add_h(doc, "Planteamiento del Problema", 1)
    add_p(doc, "El problema central es la falta de un mecanismo accesible, seguro y trazable para replicar informacion entre bases heterogeneas. Esto ocasiona demoras en reportes, inconsistencias y dependencia de personal tecnico para tareas repetitivas.")
    add_h(doc, "Justificacion", 1)
    add_p(doc, "Database Nexus se justifica porque ofrece una solucion de costo moderado, basada en tecnologias abiertas y desplegable en cloud, alineada con la realidad de empresas regionales que requieren mejorar su gestion de informacion sin adoptar plataformas empresariales complejas.")
    add_h(doc, "Alcance", 1)
    add_bullets(doc, ["Incluye autenticacion, configuraciones, replicacion por lotes, monitoreo y documentacion de esquemas.", "No incluye CDC exact-once ni gobierno de datos corporativo avanzado.", "El piloto se orienta a escenarios con volumen moderado y bases accesibles por red."])
    add_h(doc, "Objetivos", 1)
    add_p(doc, "Desarrollar una plataforma web que permita replicar datos entre bases de datos heterogeneas de manera segura, auditable y usable para una pyme peruana.")
    add_bullets(doc, ["Reducir tareas manuales de consolidacion.", "Proteger credenciales y aislar datos por usuario.", "Registrar progreso y errores de las ejecuciones.", "Validar factibilidad tecnica, economica y legal."])
    add_h(doc, "Marco Teorico", 1)
    add_p(doc, "La replicacion por lotes consiste en trasladar conjuntos de registros desde un origen hacia un destino bajo reglas definidas. A diferencia de CDC, no captura necesariamente cada cambio en tiempo real, pero resulta adecuada para reportes periodicos y consolidaciones operativas. La seguridad se apoya en autenticacion, cifrado, control de acceso y minimizacion de datos.")
    add_h(doc, "Desarrollo de la Solucion", 1)
    add_p(doc, "La solucion se estructuro como una aplicacion React con backend Fastify. El backend expone rutas REST para autenticacion, configuraciones, replicaciones, monitoreo y esquema. PostgreSQL/Supabase almacena usuarios, metadatos, trabajos y catalogos comprimidos.")
    add_h(doc, "Analisis de Factibilidad", 1)
    add_p(doc, "El analisis concluye que el proyecto es viable. La tecnologia se encuentra disponible, los costos son razonables para un piloto y el marco legal peruano permite su operacion siempre que se respeten medidas de proteccion de datos.")
    add_h(doc, "Tecnologia de Desarrollo", 1)
    add_bullets(doc, ["Frontend: React, Vite y Tailwind CSS.", "Backend: Node.js 20, Fastify, JWT, Helmet y multipart.", "Datos: PostgreSQL/Supabase y conectores hacia motores externos.", "Despliegue: Render o VPS con variables de entorno y health check."])
    add_h(doc, "Metodologia de implementacion", 1)
    add_p(doc, "Se aplico una metodologia incremental: primero alcance y requerimientos, luego arquitectura, construccion de funcionalidades criticas, pruebas con datos de muestra y documentacion final.")
    add_h(doc, "Cronograma", 1)
    add_matrix(doc, ["Actividad", "Mar", "Abr", "May", "Jun"], [["Analisis y factibilidad", "X", "", "", ""], ["Vision y SRS", "X", "X", "", ""], ["Arquitectura", "", "X", "", ""], ["Construccion", "", "X", "X", ""], ["Pruebas y ajustes", "", "", "X", "X"], ["Informe final", "", "", "", "X"]], [2.5, 1.0, 1.0, 1.0, 1.0])
    add_h(doc, "Presupuesto", 1)
    add_matrix(doc, ["Categoria", "Monto S/"], [["Personal", "10,000"], ["Cloud y dominio", "1,300"], ["Ambiente de pruebas", "1,200"], ["Costos generales", "480"], ["Total", "12,980"]], [4.5, 2.0])
    add_h(doc, "Conclusiones", 1)
    add_bullets(doc, ["Database Nexus responde a una necesidad real de integracion de datos en pymes.", "La arquitectura propuesta es suficiente para el alcance por lotes y permite evolucion futura.", "La inversion estimada es recuperable por ahorro operativo y mejora de calidad de informacion.", "El cumplimiento legal depende de politicas claras de tratamiento y seguridad de datos."])
    add_h(doc, "Recomendaciones", 1)
    add_bullets(doc, ["Implementar CDC en una version futura para escenarios criticos.", "Agregar auditoria avanzada y reportes descargables.", "Formalizar politicas de respaldo y retencion por cliente.", "Realizar pruebas de carga con volumen superior al piloto."])
    add_h(doc, "Bibliografia", 1)
    add_bullets(doc, ["Congreso de la Republica del Peru. Ley N. 29733, Ley de Proteccion de Datos Personales.", "OWASP Foundation. Web Security Testing Guide.", "PostgreSQL Global Development Group. PostgreSQL Documentation.", "MongoDB Inc. MongoDB Manual."])
    add_h(doc, "Anexos", 1)
    add_bullets(doc, ["Anexo 01: Informe de Factibilidad.", "Anexo 02: Documento de Vision.", "Anexo 03: Documento SRS.", "Anexo 04: Documento SAD.", "Anexo 05: Manuales y documentos tecnicos complementarios."])
    save_doc(doc, "FD05-EPIS-Informe ProyectoFinal.docx", "Informe Final")


def markdown_files() -> None:
    fact = """# Informe de Factibilidad - Database Nexus

**Proyecto:** Database Nexus: replicacion de datos entre bases de datos  
**Curso:** SI783 - Sistemas de Informacion  
**Lugar:** Tacna, Peru  
**Version:** 1.0  

## Descripcion del Proyecto

Database Nexus es una aplicacion web SaaS multiusuario para configurar, supervisar y ejecutar replicaciones por lotes entre PostgreSQL, MySQL, MariaDB, SQL Server, SQLite y MongoDB. El escenario de referencia es una pyme de Tacna que consolida ventas, inventario y clientes desde sistemas dispersos.

## Riesgos

| Riesgo | Impacto | Mitigacion |
|---|---|---|
| Credenciales externas mal gestionadas | Alto | Cifrado AES-256-GCM y variables de entorno |
| Cambios de esquema | Medio | Validacion previa y comparacion de esquemas |
| Volumen elevado | Medio | Lotes de 5000 filas y limites configurables |
| Datos personales | Alto | Cumplimiento de Ley 29733 |

## Factibilidad

El proyecto es tecnicamente viable por usar tecnologias maduras como React, Node.js, Fastify y PostgreSQL/Supabase. Es economicamente viable con una inversion estimada de S/ 12,980 y beneficio anual conservador de S/ 18,600. Operativamente requiere capacitacion breve para usuarios tecnicos y legalmente debe aplicar politicas de proteccion de datos personales.

## Analisis financiero

| Indicador | Valor |
|---|---:|
| Inversion inicial | S/ 12,980 |
| Beneficio anual estimado | S/ 18,600 |
| Relacion B/C | 1.43 |
| VAN referencial | S/ 3,120 |
| TIR referencial | 24% |

## Conclusion

Database Nexus es viable para un piloto peruano de integracion de datos. La solucion reduce reprocesos manuales, mejora trazabilidad y permite escalar hacia mecanismos de replicacion mas avanzados en versiones futuras.
"""
    vis = """# Documento de Vision - Database Nexus

**Sistema:** Database Nexus  
**Proyecto:** replicacion de datos entre bases de datos  
**Version:** 1.0  
**Lugar:** Tacna, Peru  

## Proposito

Definir la vision del producto Database Nexus, una plataforma web para replicar datos entre bases heterogeneas de forma segura, trazable y accesible para pymes peruanas.

## Alcance

Incluye autenticacion, configuraciones cifradas, replicacion por lotes, monitoreo de salud, documentacion de esquemas y administracion basica de usuarios. No incluye CDC exact-once en la primera version.

## Posicionamiento

El producto atiende a empresas que conviven con sistemas legados y bases diferentes. Frente a procesos manuales con CSV u hojas de calculo, Database Nexus permite configurar flujos repetibles y monitoreados.

## Usuarios

| Usuario | Necesidad |
|---|---|
| Gerencia | Reportes oportunos y confiables |
| Administrador TI | Control de conexiones y trabajos |
| Analista de datos | Consolidacion rapida de informacion |
| Usuario operativo | Datos actualizados para tareas diarias |

## Capacidades

- Autenticacion JWT.
- Credenciales cifradas.
- Conectores PostgreSQL, MySQL, MariaDB, SQL Server, SQLite y MongoDB.
- Modos insertar, upsert, reemplazar y recargar.
- Progreso persistente y reanudacion.
- Health monitor y documentador de esquemas.

## Conclusion

La vision es construir una herramienta realista, de costo moderado y alineada con necesidades peruanas de integracion de datos en organizaciones pequenas y medianas.
"""
    (ROOT / "FD01-Informe-Factibilidad.md").write_text(fact, encoding="utf-8")
    (ROOT / "FD02-Informe-Vision.md").write_text(vis, encoding="utf-8")


def main() -> None:
    propuesta()
    factibilidad()
    vision()
    srs()
    arquitectura()
    final_report()
    markdown_files()


if __name__ == "__main__":
    main()
