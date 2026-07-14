from flask import Flask, render_template, request, send_file
import pandas as pd
from datetime import datetime
import os
import io

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

COLUMN_RENAMES = {
    'UNNAMED: 3': 'FECHA DE RECEPCION',
    'UNNAMED: 7': 'NOMBRE',
    'UNNAMED: 9': 'FECHA DE ELABORACION',
    'UNNAMED: 10': 'FECHA DE ENTREGA AL MP O TRIBUNAL',
    'NUMERO DE OFICIO ORE GUARICO': 'NUMERO DE OFICIO',
}
HIDDEN_COLUMNS = {
    'UNNAMED: 12',
    'RESPUESTA ENVIADA EN OFICIO NO. Y FECHA',
}


def normalize_column_name(col):
    normalized = str(col).strip().replace('\n', ' ').upper()
    return COLUMN_RENAMES.get(normalized, normalized)


def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    normalized_names = [normalize_column_name(col) for col in df.columns]
    seen = {}
    unique_names = []

    for name in normalized_names:
        count = seen.get(name, 0)
        seen[name] = count + 1
        if count == 0:
            unique_names.append(name)
        else:
            unique_names.append(f'{name}_{count}')

    df.columns = unique_names
    return df


def is_hidden_column(col: str) -> bool:
    return any(col == hidden or col.startswith(f'{hidden}_') for hidden in HIDDEN_COLUMNS)


def find_fecha_col(columns):
    return next(
        (col for col in columns if 'FECHA DE RECEPCION' in col),
        None,
    )


def find_oficio_col(columns):
    oficio_candidates = [
        col for col in columns
        if 'NUMERO' in col and 'OFICIO' in col
    ]
    if oficio_candidates:
        return oficio_candidates[0]

    oficio_candidates = [
        col for col in columns
        if 'OFICIO' in col and 'RESPUESTA' not in col
    ]
    if oficio_candidates:
        return oficio_candidates[0]

    return next((col for col in columns if 'OFICIO' in col), None)


def is_blank_value(value):
    if pd.isna(value):
        return True
    if isinstance(value, str) and value.strip().lower() in ['', 'nan', 'nat', 'none']:
        return True
    return False


def get_status_class(status: str) -> str:
    if status == 'Respondido':
        return 'row-respondido'
    if status == 'No Respondido':
        return 'row-no-respondido'
    if status == 'No Detectado':
        return 'row-no-detectado'
    return ''


@app.route('/', methods=['GET'])
def index():
    today = datetime.today().date()
    return render_template(
        'index.html',
        title='Análisis de Solicitudes 2026',
        subtitle='Carga tu archivo y analiza rápidamente el estado de los trámites por oficio.',
        min_date=today.replace(month=1, day=1),
        max_date=today.replace(month=12, day=31),
        table_headers=None,
        table_rows=None,
        row_classes=None,
        error=None,
        success=None,
    )


@app.route('/procesar', methods=['POST'])
def procesar():
    uploaded_file = request.files.get('file')
    fecha_inicio = request.form.get('fecha_inicio')
    fecha_fin = request.form.get('fecha_fin')
    is_export = request.form.get('export') == 'xlsx'

    if (uploaded_file is None or uploaded_file.filename == '') and not is_export:
        return render_template(
            'index.html',
            title='Análisis de Solicitudes 2026',
            subtitle='Carga tu archivo y analiza rápidamente el estado de los trámites por oficio.',
            error='No se seleccionó ningún archivo. Por favor, sube un archivo CSV o XLSX.',
            min_date=datetime.today().date().replace(month=1, day=1),
            max_date=datetime.today().date().replace(month=12, day=31),
            table_headers=None,
            table_rows=None,
            row_classes=None,
        )

    try:
        filename = uploaded_file.filename.lower()
        if filename.endswith('.csv'):
            df = pd.read_csv(uploaded_file, skiprows=2)
        else:
            # 1. Cargar el archivo Excel completo en memoria para evaluar sus hojas
            excel_file = pd.ExcelFile(uploaded_file)
            
            # 2. Buscar dinámicamente cualquier pestaña que contenga el texto clave
            target_sheet = None
            for sheet in excel_file.sheet_names:
                if '2026 SOLICITUDES' in sheet.upper():
                    target_sheet = sheet
                    break
            
            # 3. Si no encuentra ninguna coincidencia válida, lanzar excepción detallada
            if not target_sheet:
                raise ValueError("No se encontró ninguna pestaña que contenga '2026 SOLICITUDES' dentro del archivo Excel.")
                
            # 4. Leer la pestaña detectada de forma automática
            df = pd.read_excel(excel_file, sheet_name=target_sheet, skiprows=2)

        df = clean_columns(df)
        fecha_columna = find_fecha_col(df.columns)

        if fecha_columna is None:
            raise ValueError("No se encontró la columna 'FECHA DE RECEPCION'.")

        df[fecha_columna] = pd.to_datetime(df[fecha_columna], errors='coerce')
        df = df.dropna(subset=[fecha_columna])
        df = df[df[fecha_columna].dt.year == 2026]

        min_date = df[fecha_columna].min().date() if not df.empty else datetime(2026, 1, 1).date()
        max_date = df[fecha_columna].max().date() if not df.empty else datetime(2026, 12, 31).date()

        fecha_inicio_val = datetime.fromisoformat(fecha_inicio).date() if fecha_inicio else min_date
        fecha_fin_val = datetime.fromisoformat(fecha_fin).date() if fecha_fin else max_date

        fecha_inicio_val = max(min_date, fecha_inicio_val)
        fecha_fin_val = min(max_date, fecha_fin_val)

        mask = (
            (df[fecha_columna].dt.date >= fecha_inicio_val)
            & (df[fecha_columna].dt.date <= fecha_fin_val)
        )
        df_filtrado = df.loc[mask].copy()

        oficio_col = find_oficio_col(df.columns)
        if oficio_col:
            df_filtrado['Estado_Respuesta'] = df_filtrado[oficio_col].apply(
                lambda x: 'Respondido' if not is_blank_value(x) else 'No Respondido'
            )
        else:
            df_filtrado['Estado_Respuesta'] = 'No Detectado'

        if is_export:
            visible_cols = [col for col in df_filtrado.columns if not is_hidden_column(col)]
            df_export = df_filtrado[visible_cols].copy()
            
            for col in df_export.columns:
                if pd.api.types.is_datetime64_any_dtype(df_export[col]):
                    df_export[col] = df_export[col].dt.strftime('%Y-%m-%d')

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_export.to_excel(writer, index=False, sheet_name='Resultados')
            output.seek(0)
            
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name=f"reporte_filtrado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )

        respondidos = int((df_filtrado['Estado_Respuesta'] == 'Respondido').sum())
        no_respondidos = int((df_filtrado['Estado_Respuesta'] == 'No Respondido').sum())
        no_detectado = int((df_filtrado['Estado_Respuesta'] == 'No Detectado').sum())

        visible_columns = [col for col in df_filtrado.columns if not is_hidden_column(col)]
        table_headers = visible_columns
        table_rows = []
        row_classes = []

        for _, row in df_filtrado.iterrows():
            row_values = []
            for col in visible_columns:
                value = row[col]
                if pd.isna(value):
                    row_values.append('')
                elif isinstance(value, (pd.Timestamp, datetime)):
                    row_values.append(value.strftime('%Y-%m-%d'))
                else:
                    row_values.append(str(value))
            table_rows.append(row_values)
            row_classes.append(get_status_class(row['Estado_Respuesta']))

        return render_template(
            'index.html',
            title='Análisis de Solicitudes 2026',
            subtitle='Carga tu archivo y analiza rápidamente el estado de los trámites por oficio.',
            success=f'Archivo procesado correctamente: {uploaded_file.filename}',
            filename=uploaded_file.filename,
            columns=', '.join(df.columns),
            fecha_inicio=fecha_inicio_val,
            fecha_fin=fecha_fin_val,
            total=len(df_filtrado),
            respondidos=respondidos,
            no_respondidos=no_respondidos,
            no_detectado=no_detectado,
            table_headers=table_headers,
            table_rows=table_rows,
            row_classes=row_classes,
            min_date=min_date,
            max_date=max_date,
            selected_inicio=fecha_inicio_val,
            selected_fin=fecha_fin_val,
        )
    except Exception as e:
        return render_template(
            'index.html',
            title='Análisis de Solicitudes 2026',
            subtitle='Carga tu archivo y analiza rápidamente el estado de los trámites por oficio.',
            error=f'Error al procesar la pestaña requerida del archivo: {e}',
            min_date=datetime.today().date().replace(month=1, day=1),
            max_date=datetime.today().date().replace(month=12, day=31),
            table_headers=None,
            table_rows=None,
            row_classes=None,
        )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)