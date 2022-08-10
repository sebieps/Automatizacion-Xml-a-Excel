import xml.etree.cElementTree as ET
from xml.dom import minidom
import openpyxl
from openpyxl.styles import Font

wb = openpyxl.Workbook()
sheet = wb.active

# VARIABLES
i = 0
_serviceidITEM = ''
_uuidSERVICE = ''
_serviceidITEM1 = ''
_uuidSERVICE1 = ''
_pais = ''
_routerid = ''
_empresa = ''
_paraExcel = 2


def dividirServerIp(palabra):
    indice = palabra.find(':')
    _paranumero = indice + 3
    longitud = len(palabra)
    ip = palabra[0:indice]
    numero = palabra[_paranumero:longitud]
    sheet[f'E{_paraExcel}'] = ip
    sheet[f'F{_paraExcel}'] = numero


archivo_xml = ET.parse('ABRIR.xml')

raiz = archivo_xml.getroot()

Workspaces1 = archivo_xml.find('Workspaces')
Workspace1 = archivo_xml.find('Workspace')

Services1 = archivo_xml.find('Services')

Routers1 = archivo_xml.find('Routers')

Messageservers1 = archivo_xml.find('Messageservers')


sheet['A1'] = 'Pais'
sheet['A1'].font = Font(color='FF0000', bold=True, size=14)
sheet['B1'] = 'Empresa'
sheet['B1'].font = Font(color='FF0000', bold=True, size=14)
sheet['C1'] = 'Hostname + Descripcion'
sheet['C1'].font = Font(color='FF0000', bold=True, size=14)
sheet['D1'] = 'Systemid'
sheet['D1'].font = Font(color='FF0000', bold=True, size=14)
sheet['E1'] = 'IP'
sheet['E1'].font = Font(color='FF0000', bold=True, size=14)
sheet['F1'] = 'No'
sheet['F1'].font = Font(color='FF0000', bold=True, size=14)
sheet['G1'] = 'SapRouter'
sheet['G1'].font = Font(color='FF0000', bold=True, size=14)
sheet['H1'] = 'Host'
sheet['H1'].font = Font(color='FF0000', bold=True, size=14)
sheet['I1'] = 'Comentario'
sheet['I1'].font = Font(color='FF0000', bold=True, size=14)


for hijo in Workspaces1:
    Workspace1 = hijo

for hijo in Workspace1:
    if(hijo.tag == 'Item'):
        _serviceidITEM = hijo.attrib['serviceid']
        for service in Services1:
            _uuidSERVICE = service.attrib['uuid']
            if(_serviceidITEM == _uuidSERVICE):
                sheet[f'A{_paraExcel}'] = ' '
                sheet[f'B{_paraExcel}'] = ' '
                sheet[f'C{_paraExcel}'] = service.attrib['name']
                sheet[f'D{_paraExcel}'] = service.attrib['systemid']
                if(service.attrib.keys().mapping.__contains__('server')):
                    dividirServerIp(service.attrib['server'])
                for comentario in service:
                    sheet[f'I{_paraExcel}'] = comentario.text
                for router in Routers1:
                    if(service.attrib.keys().mapping.__contains__('routerid')):
                        _routerid = service.attrib['routerid']
                        if(_routerid == router.attrib['uuid']):
                            sheet[f'G{_paraExcel}'] = router.attrib['router']
                _paraExcel += 1
    if(hijo.tag == 'Node'):
        _pais = hijo.attrib['name']
        for nodeEmpresa in hijo:
            if(nodeEmpresa.attrib.keys().mapping.__contains__('name') and nodeEmpresa.attrib['name'] != 'Argentina' and nodeEmpresa.attrib['name'] != 'Brasil' and nodeEmpresa.attrib['name'] != 'Chile' and nodeEmpresa.attrib['name'] != 'Clientes' and nodeEmpresa.attrib['name'] != 'Colombia' and nodeEmpresa.attrib['name'] != 'Ecuador' and nodeEmpresa.attrib['name'] != 'Peru'):
                _empresa = nodeEmpresa.attrib['name']
            for item in nodeEmpresa:
                if(item.attrib.keys().mapping.__contains__('serviceid')):
                    _serviceidITEM1 = item.attrib['serviceid']
                for service in Services1:
                    _uuidSERVICE1 = service.attrib['uuid']
                    if(_serviceidITEM1 == _uuidSERVICE1):
                        sheet[f'A{_paraExcel}'] = _pais
                        sheet[f'B{_paraExcel}'] = _empresa
                        sheet[f'C{_paraExcel}'] = service.attrib['name']
                        sheet[f'D{_paraExcel}'] = service.attrib['systemid']
                        if(service.attrib.keys().mapping.__contains__('server')):
                            dividirServerIp(service.attrib['server'])
                        for comentario in service:
                            sheet[f'I{_paraExcel}'] = comentario.text
                        for router in Routers1:
                            if(service.attrib.keys().mapping.__contains__('routerid')):
                                _routerid = service.attrib['routerid']
                                if(_routerid == router.attrib['uuid']):
                                    sheet[f'G{_paraExcel}'] = router.attrib['router']
                        _paraExcel += 1
    for item in hijo:
        if(hijo.attrib['name'] == 'Chile'):
            _pais = 'Chile'
            _empresa = ''
        else:
            _pais = ''
            _empresa = hijo.attrib['name']
        if(item.tag == 'Item'):
            _serviceidITEM1 = item.attrib['serviceid']
            for service in Services1:
                _uuidSERVICE1 = service.attrib['uuid']
                if(_serviceidITEM1 == _uuidSERVICE1):
                    sheet[f'A{_paraExcel}'] = _pais
                    sheet[f'B{_paraExcel}'] = _empresa
                    sheet[f'C{_paraExcel}'] = service.attrib['name']
                    sheet[f'D{_paraExcel}'] = service.attrib['systemid']
                    if(service.attrib.keys().mapping.__contains__('server')):
                        dividirServerIp(service.attrib['server'])
                    for comentario in service:
                        sheet[f'I{_paraExcel}'] = comentario.text
                        print(comentario.text)
                    for router in Routers1:
                        if(service.attrib.keys().mapping.__contains__('routerid')):
                            _routerid = service.attrib['routerid']
                            if(_routerid == router.attrib['uuid']):
                                sheet[f'G{_paraExcel}'] = router.attrib['router']
                    _paraExcel += 1


print('Excel creado con EXITO')
wb.save('EXCELXML.xlsx')
