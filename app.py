from flask import Flask, render_template, request
from lxml import objectify
import xlsxwriter

app = Flask(__name__)

@app.route("/")
def hello_world():
    return render_template("index.html")

@app.post("/upload")
def upload_file():
    workbook = xlsxwriter.Workbook('./finanzaass.xlsx')
    worksheet = workbook.add_worksheet()

    # formatting
    cell_styling = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    currency_styling = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'num_format': '#,##0.00'})

    # headers
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

    # width
    worksheet.set_column(0, 0, 10)
    worksheet.set_column(2, 2, 50)
    if request.method == "POST":
        
        ticket_row = 1
        product_row = 1
        for file in request.files.getlist("xmlFile"):
            tree = objectify.fromstring(file.read())
            # Fields
            cdc = tree.DE.attrib['Id']  # Codigo de Control
            fecha_emision_DE = tree.DE.gDatGralOpe.dFeEmiDE.text  # DE -> Documento Electronico
            product_services_list = [{
                "name": p.dDesProSer.text,
                "quantity": round(float(p.dCantProSer.text), 3),
                "unit_of_measure": p.dDesUniMed.text,
                "unit_price": p.gValorItem.dPUniProSer,
                "gross_total": p.gValorItem.dTotBruOpeItem
            } \
                for p in tree.DE.gDtipDE.gCamItem]
            amount_paid_total = float(tree.DE.gTotSub.dTotOpe.text)
            company_name = tree.DE.gDatGralOpe.gDatRec.dNomRec.text  # nombre o razon social
            seller = tree.DE.gDatGralOpe.gEmis.dNomEmi.text

            for product_service in product_services_list:
                worksheet.write(product_row, 2, product_service['name'], cell_styling)
                worksheet.write(product_row, 3, product_service['quantity'], currency_styling)
                worksheet.write(product_row, 4, product_service['unit_of_measure'], cell_styling)
                worksheet.write(product_row, 5, product_service['unit_price'], currency_styling)
                worksheet.write(product_row, 6, product_service['gross_total'], currency_styling)
                product_row += 1

            if len(product_services_list) > 1:
                worksheet.merge_range(ticket_row, 0, ticket_row + len(product_services_list) - 1, 0, cdc, cell_styling)
                worksheet.merge_range(ticket_row, 1, ticket_row + len(product_services_list) - 1, 1, fecha_emision_DE,
                                      cell_styling)
                worksheet.merge_range(ticket_row, 7, ticket_row + len(product_services_list) - 1, 7, amount_paid_total,
                                      currency_styling)
                worksheet.merge_range(ticket_row, 8, ticket_row + len(product_services_list) - 1, 8, company_name,
                                      cell_styling)
                worksheet.merge_range(ticket_row, 9, ticket_row + len(product_services_list) - 1, 9, seller,
                                      cell_styling)
            else:
                worksheet.write(ticket_row, 0, cdc, cell_styling)
                worksheet.write(ticket_row, 1, fecha_emision_DE, cell_styling)
                worksheet.write(ticket_row, 7, amount_paid_total, currency_styling)
                worksheet.write(ticket_row, 8, company_name, cell_styling)
                worksheet.write(ticket_row, 9, seller, cell_styling)
            ticket_row = product_row

        workbook.close()

    return "success"