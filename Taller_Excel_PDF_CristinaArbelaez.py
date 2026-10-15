import openpyxl
from tabulate import tabulate
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

encabezados = ["ID", "Producto", "Categoría", "Precio ($)", "Ventas (Unid)"]
alineacion = ("center", "left", "center", "right", "right")


def cargar_datos():
    wb = openpyxl.load_workbook('ventas_tecnologia_v2.xlsx', data_only=True)
    hoja = wb.active
    datos = []
    for fila in hoja.iter_rows(min_row=2, values_only=True):
        datos.append(list(fila))
    print(f"Total de registros leídos en memoria: {len(datos)}")
    return datos


def mostrar_catalogo(datos):
    print(f"\n=== REPORTE DE VENTAS EN CONSOLA ({len(datos)} PRODUCTOS) ===")
    print(tabulate(datos, headers=encabezados, tablefmt="fancy_grid", colalign=alineacion))

    total_unidades = sum(fila[4] for fila in datos)
    precio_promedio = sum(fila[3] for fila in datos) / len(datos)
    print(f"\nTotal de unidades vendidas: {total_unidades}")
    print(f"Precio promedio de los productos: ${precio_promedio:,.2f}")


def filtrar_por_precio(datos):
    datos_filtrados = [fila for fila in datos if fila[3] > 100]

    print(f"\n=== PRODUCTOS CON PRECIO MAYOR A $100 ({len(datos_filtrados)} PRODUCTOS) ===")
    print(tabulate(datos_filtrados, headers=encabezados, tablefmt="fancy_grid", colalign=alineacion))


def buscar_por_id(datos):
    # Reto 3- Buscar por ID
    try:
        id_buscado = int(input("Ingrese el ID del producto a buscar: "))
        resultado = [fila for fila in datos if fila[0] == id_buscado]

        if resultado:
            print("\n=== PRODUCTO ENCONTRADO ===")
            print(tabulate(resultado, headers=encabezados, tablefmt="fancy_grid", colalign=alineacion))
        else:
            print(f"\n[ADVERTENCIA] No existe ningún producto con el ID {id_buscado}.")
    except ValueError:
        print("\n[ERROR] Debe ingresar un número entero, por ejemplo: 105")


def exportar_pdf(datos):
    total_unidades = sum(fila[4] for fila in datos)
    precio_promedio = sum(fila[3] for fila in datos) / len(datos)

    # Crear el documento base
    doc = SimpleDocTemplate("reporte_ejecutivo_ventas.pdf", pagesize=letter)
    estilos = getSampleStyleSheet()

    # Crear un estilo personalizado 
    estilo_titulo = ParagraphStyle(
        "EstiloTitulo", parent=estilos["Title"],
        fontName="Helvetica-Bold", fontSize=18, 
        textColor=colors.HexColor("#4B8BBE"), spaceAfter=20)

    estilo_cuerpo = ParagraphStyle(
        "EstiloCuerpo", parent=estilos["Normal"],
        fontSize=9)
    estilo_encabezados = ParagraphStyle(
        "EstiloEncabezados", parent=estilos["Normal"],
        fontName="Helvetica-Bold", fontSize=10,
        textColor=colors.white, alignment=1)

    # La cajita donde voy poniendo cosas
    story = [
        Paragraph("Informe Ejecutivo de Ventas Tecno", estilo_titulo),
        Paragraph(f'Total de productos procesados: {len(datos)}', estilo_cuerpo),
    ]

    encabezados_pdf = [Paragraph(f"<b>{h}</b>", estilo_encabezados) for h in encabezados]
    filas_pdf = [
        [
            Paragraph(str(r[0]), ParagraphStyle('C', parent=estilo_cuerpo, alignment=1)),
            Paragraph(str(r[1]), ParagraphStyle('L', parent=estilo_cuerpo, alignment=0)),
            Paragraph(str(r[2]), ParagraphStyle('C', parent=estilo_cuerpo, alignment=1)),
            Paragraph(f"${r[3]:,.2f}", ParagraphStyle('R', parent=estilo_cuerpo, alignment=2)),
            Paragraph(str(r[4]), ParagraphStyle('R', parent=estilo_cuerpo, alignment=2))
        ]
        for r in datos
    ]

    #
    fila_resumen = [
        Paragraph("", ParagraphStyle('C', parent=estilo_cuerpo, alignment=1)),
        Paragraph("<b>RESUMEN</b>", ParagraphStyle('L', parent=estilo_cuerpo, alignment=0)),
        Paragraph("<b>Promedio / Total</b>", ParagraphStyle('C', parent=estilo_cuerpo, alignment=1)),
        Paragraph(f"<b>${precio_promedio:,.2f}</b>", ParagraphStyle('R', parent=estilo_cuerpo, alignment=2)),
        Paragraph(f"<b>{total_unidades:,}</b>", ParagraphStyle('R', parent=estilo_cuerpo, alignment=2)),
    ]

    #configurar los anchos y columnas
    anchos_cols=[40, 160, 100, 80, 80]
    table_pdf = Table([encabezados_pdf] + filas_pdf + [fila_resumen], colWidths=anchos_cols, repeatRows=1)

    #estilo con table style
    tabla_pdf_style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#4B8BBE")),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey])
    ])

    table_pdf.setStyle(tabla_pdf_style)
    table_pdf.setStyle(TableStyle([
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#CBD5E0")),
        ('LINEABOVE', (0, -1), (-1, -1), 1.5, colors.HexColor("#4B8BBE")),
    ]))

    story.append(table_pdf)

    datos_filtrados = [fila for fila in datos if fila[3] > 100]

    # --- PDF filtrado (Reto 1) ---
    filas_pdf_filtradas = [
        [
            Paragraph(str(r[0]), ParagraphStyle('C', parent=estilo_cuerpo, alignment=1)),
            Paragraph(str(r[1]), ParagraphStyle('L', parent=estilo_cuerpo, alignment=0)),
            Paragraph(str(r[2]), ParagraphStyle('C', parent=estilo_cuerpo, alignment=1)),
            Paragraph(f"${r[3]:,.2f}", ParagraphStyle('R', parent=estilo_cuerpo, alignment=2)),
            Paragraph(str(r[4]), ParagraphStyle('R', parent=estilo_cuerpo, alignment=2))
        ]
        for r in datos_filtrados
    ]
    tabla_filtrada = Table([encabezados_pdf] + filas_pdf_filtradas, colWidths=anchos_cols, repeatRows=1)
    tabla_filtrada.setStyle(tabla_pdf_style)

    story.append(PageBreak())
    story.append(Paragraph("Productos con precio mayor a $100", estilo_titulo))
    story.append(Paragraph(f"Total de productos: {len(datos_filtrados)}", estilo_cuerpo))
    story.append(Spacer(1, 12))
    story.append(tabla_filtrada)

    try:
        doc.build(story)
        print("\n[ÉXITO] Archivo PDF generado: reporte_ejecutivo_ventas.pdf")
    except PermissionError:
        print("\n[ERROR] No se pudo guardar el PDF. Ciérrelo si lo tiene abierto e intente de nuevo.")


def menu():
    datos = cargar_datos()

    while True:
        print("\n===== MENÚ TECNOVENTAS =====")
        print("1) Ver catálogo completo")
        print("2) Filtrar por precio (> $100)")
        print("3) Buscar por ID")
        print("4) Exportar informe PDF")
        print("5) Salir")
        opcion = input("Elija una opción: ")

        if opcion == "1":
            mostrar_catalogo(datos)
        elif opcion == "2":
            filtrar_por_precio(datos)
        elif opcion == "3":
            buscar_por_id(datos)
        elif opcion == "4":
            exportar_pdf(datos)
        elif opcion == "5":
            print("\nHasta luego")
            break
        else:
            print("\n[ADVERTENCIA] Opción no válida, elija un número del 1 al 5")


menu()