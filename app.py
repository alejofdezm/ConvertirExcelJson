import pandas as pd
import json
from flask import Flask, request, jsonify
from uuid import uuid4  # Para generar IDs únicos
from datetime import datetime

app = Flask(__name__)

listDay = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
listbankBranch = ['BLCR', 'BLNI', 'BLPA', 'BLHN', 'BLRD']

const_LunesAViernes = 'Lunes a Viernes'
const_LunesASabado = 'Lunes a Sábado'
const_LunesADomingo = 'Lunes a Domingo'
const_SabadoADomingo = 'Sábado a Domingo'


def getDays(days):
    if days in listDay:
        return [days]
    elif days == const_LunesAViernes:
        return listDay[:5]
    elif days == const_LunesASabado:
        return listDay[:6]
    elif days == const_LunesADomingo:
        return listDay
    elif days == const_SabadoADomingo:
        return listDay[5:]


def convertir_formato(hora_militar):
    hora_str = f"{hora_militar:04d}"
    hora_obj = datetime.strptime(hora_str, "%H%M")
    return hora_obj.strftime("%I:%M %p")


def getDaysText(days, start_hour, end_hour):
    if days in listDay:
        return days + " de " + convertir_formato(start_hour) + " a " + convertir_formato(end_hour)
    elif days == const_LunesAViernes:
        return const_LunesAViernes + " de " + convertir_formato(start_hour) + " a " + convertir_formato(end_hour)
    elif days == const_LunesASabado:
        return const_LunesASabado + " de " + convertir_formato(start_hour) + " a " + convertir_formato(end_hour)
    elif days == const_LunesADomingo:
        return const_LunesADomingo + " de " + convertir_formato(start_hour) + " a " + convertir_formato(end_hour)
    elif days == const_SabadoADomingo:
        return const_SabadoADomingo + " de " + convertir_formato(start_hour) + " a " + convertir_formato(end_hour)


def convert_schedule_text_to_json(days, start_hour, end_hour):
    schedule = {
        "text": getDaysText(days, start_hour, end_hour),
        "days": getDays(days),
        "hours": {
            "min": start_hour,
            "max": end_hour
        }
    }
    return schedule


def convert_excel_row_to_json(group):
    locations_dict = {}
    for _, grp_row in group.iterrows():
        location_key = (
            grp_row['Nombre'],
            grp_row['Descripcion'],
            grp_row['Tipo'],
            grp_row['Detalle'],
            "ubicacionCarritoCompra-LAFISECR.png",
            grp_row['Lat'],
            grp_row['Lng'],
        )
        schedule = convert_schedule_text_to_json(grp_row['Dias'], grp_row['HoraDesde'], grp_row['HoraHasta'])
        if location_key in locations_dict:
            locations_dict[location_key]['schedules'].append(schedule)
        else:
            location_json = {
                "id": str(uuid4()),
                "name": grp_row['Nombre'],
                "description": grp_row['Descripcion'],
                "type": grp_row['Tipo'],
                "detail": grp_row['Detalle'],
                "icon": "ubicacionCarritoCompra-LAFISECR.png",
                "lat": grp_row['Lat'],
                "lng": grp_row['Lng'],
                "schedules": [schedule]
            }
            locations_dict[location_key] = location_json
    return list(locations_dict.values())


def convert_excel_to_json(excel_file):
    xls = pd.ExcelFile(excel_file)

    bank_branch = xls.sheet_names[0]

    if bank_branch not in listbankBranch:
        return {"mensaje": "El nombre de la hoja no esta en la lista de sucursales", "sucursales": listbankBranch}

    df = pd.read_excel(excel_file, sheet_name=bank_branch)
    
    json_template = {
        "bankBranch": bank_branch,
        "configs": []
    }

    grouped = df.groupby(['Ubicacion'])
    
    for place, group in grouped: 
        place_config = {
            "place": place[0],
            "locations": convert_excel_row_to_json(group)
        }
        json_template['configs'].append(place_config)

    return json_template


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify(error="No file part"), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify(error="No selected file"), 400
    if file:
        json_data = convert_excel_to_json(file)
        return jsonify(json_data), 200


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')
