from lxml import objectify
import xlsxwriter

workbook = xlsxwriter.Workbook('../finanzas.xlsx')
worksheet = workbook.add_worksheet()

# formatting
cell_styling = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
currency_styling = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'num_format': '#,##0.00'})

worksheet.write('A1', 'CDC')
worksheet.write('B1', 'FECHA EMISION')
worksheet.write('C1', 'PRODUCTO/SERVICIO')
worksheet.write('D1', 'CANTIDAD')
worksheet.write('E1', 'UNIDAD DE MEDIDA')
worksheet.write('F1', 'PRECIO UNITARIO')
worksheet.write('G1', 'PRECIO BRUTO')
worksheet.write('H1', 'TOTAL PAGADO')
worksheet.write('I1', 'NOMBRE O RAZON SOCIAL')
worksheet.write('J1', 'COMERCIO')

ticket_row = 0
product_row = 0
for file_path in ['../file1.xml', '../file2.xml']:
    tree = objectify.parse(file_path)
    root = tree.getroot()

    # Campos
    cdc = root.DE.attrib['Id']  # Codigo de Control
    fecha_emision_DE = root.DE.gDatGralOpe.dFeEmiDE.text  # DE -> Documento Electronico
    lista_de_productos = [{
                            "nombre": p.dDesProSer.text,
                            "cantidad": p.dCantProSer.text,
                            "unidad_de_medida": p.dDesUniMed.text,
                            "precio_unitario": p.gValorItem.dPUniProSer,
                            "precio_bruto": p.gValorItem.dTotBruOpeItem
                            } \
                          for p in root.DE.gDtipDE.gCamItem]
    total_pagado = root.DE.gTotSub.dTotOpe.text
    nombre_razon_social = root.DE.gDatGralOpe.gDatRec.dNomRec.text
    comercio = root.DE.gDatGralOpe.gEmis.dNomEmi.text

    for producto in lista_de_productos:
        worksheet.write(product_row + 1, 2, producto['nombre'])
        worksheet.write(product_row + 1, 3, producto['cantidad'])
        worksheet.write(product_row + 1, 4, producto['unidad_de_medida'])
        worksheet.write(product_row + 1, 5, producto['precio_unitario'])
        worksheet.write(product_row + 1, 6, producto['precio_bruto'])
        product_row += 1

    if len(lista_de_productos) > 1:
        worksheet.merge_range(ticket_row + 1, 0, ticket_row + len(lista_de_productos), 0, cdc, cell_styling)
        worksheet.merge_range(ticket_row + 1, 1, ticket_row + len(lista_de_productos), 1, fecha_emision_DE, cell_styling)
        worksheet.merge_range(ticket_row + 1, 7, ticket_row + len(lista_de_productos), 7, total_pagado)
        worksheet.merge_range(ticket_row + 1, 8, ticket_row + len(lista_de_productos), 8, nombre_razon_social, cell_styling)
        worksheet.merge_range(ticket_row + 1, 9, ticket_row + len(lista_de_productos), 9, comercio, cell_styling)
    else:
        worksheet.write(ticket_row + 1, 0, cdc)
        worksheet.write(ticket_row + 1, 1, fecha_emision_DE)
        worksheet.write(ticket_row + 1, 7, total_pagado)
        worksheet.write(ticket_row + 1, 8, nombre_razon_social)
        worksheet.write(ticket_row + 1, 9, comercio)
    ticket_row = len(lista_de_productos)


workbook.close()