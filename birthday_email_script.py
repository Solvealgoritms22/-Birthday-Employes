import pandas as pd
import datetime
import pytz
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from string import Template
import base64
import locale
from PIL import Image, ImageOps
import io

# Define la zona horaria de Santo Domingo
santo_domingo_zone = pytz.timezone('America/Santo_Domingo')

# Carga la data de empleados
df = pd.read_excel('data.xlsx')

# Filtrar empleados que cumplen años hoy y en los próximos 7 días
today = datetime.datetime.now(santo_domingo_zone).date()
in_seven_days = today + datetime.timedelta(days=7)

# Listas para empleados que cumplen años hoy y en los próximos 7 días con consentimiento
employees_birthday_today = []
employees_birthday_next_week = {}

# Iterar sobre los empleados para clasificarlos y verificar el consentimiento
for index, row in df.iterrows():
    if row['Consentimiento'].lower() == 'si':  # Verificar el consentimiento
        birthday = row['Fecha_nacimiento'].date()
        # Ajustar el año de cumpleaños al año actual para comparación
        birthday_this_year = birthday.replace(year=today.year)

        if birthday_this_year == today:
            employees_birthday_today.append(row)
        elif today < birthday_this_year <= in_seven_days:
            encargado = row['Encargado']
            correo_encargado = row['Correo_encargado']  # Obtener el correo del encargado
            # Si el encargado es nuevo en el diccionario, inicializar con su correo y una lista vacía de empleados
            if encargado not in employees_birthday_next_week:
                employees_birthday_next_week[encargado] = {'correo': correo_encargado, 'empleados': []}
            # Añadir el empleado a la lista del encargado
            employees_birthday_next_week[encargado]['empleados'].append(row)

# Función para redimensionar y codificar la imagen en base64
def resize_image(image_path, base_width=500):
    with open(image_path, "rb") as image_file:
        img = Image.open(image_file)
        w_percent = (base_width / float(img.size[0]))
        h_size = int((float(img.size[1]) * float(w_percent)))
        img = img.resize((base_width, h_size), Image.Resampling.LANCZOS)

        # Guardar la imagen redimensionada en un buffer
        buffer = io.BytesIO()
        img.save(buffer, format="JPEG")
        buffer.seek(0)

        # Codificar la imagen redimensionada en base64
        return base64.b64encode(buffer.read()).decode('utf-8')
    
# Función para verificar la existencia de la imagen del perfil
def get_image_path(cedula):
    image_path = f'img/perfiles/{cedula}.jpg'
    if not os.path.exists(image_path):
        image_path = 'img/perfiles/user.jpg'
    return image_path

# Función para codificar la imagen en base64
def get_image_as_base64(url):
    with open(url, 'rb') as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')

# Función para insertar empleados en la plantilla HTML
def generate_birthday_list_html(employees, template_path, is_today=True):
    list_items_html = ""
    # Abrir y leer la plantilla una vez fuera del bucle
    with open(template_path, 'r', encoding='utf-8') as file:
        template_content = file.read()

    # Intenta establecer la localización al español
    try:
        locale.setlocale(locale.LC_TIME, 'es_ES' if os.name != 'nt' else 'Spanish_Spain')
    except locale.Error:
        print("La localización española no está disponible en este sistema.")

    for employee in employees:
        # Obtén la ruta de la imagen y luego redimensiónala y codifícala en base64
        image_path = get_image_path(employee['Cedula'])
        image_base64 = resize_image(image_path)

        # Crear un objeto Template con la plantilla leída
        template = Template(template_content)

        # Formatear la fecha de cumpleaños
        birthday_format = '%-d de %B' if os.name != 'nt' else '%#d de %B'  # Formato para Unix y Windows
        formatted_birthday = employee['Fecha_nacimiento'].strftime(birthday_format)

        # Sustituir los marcadores de posición con los datos del empleado
        list_item_html = template.substitute(
            foto_base64=image_base64,
            nombre_completo=employee['Nombre'],
            posicion=employee['Cargo'],
            departamento_fecha=employee['Departamento'] if is_today else formatted_birthday
        )
        # Concatenar cada elemento de la lista
        list_items_html += list_item_html

    return list_items_html

# Función para enviar correo electrónico
def send_email(subject, body, to_addresses, images=None):
    message = MIMEMultipart()
    message['Subject'] = subject
    message['To'] = ', '.join(to_addresses)
    message.attach(MIMEText(body, 'html'))

    # Adjuntar imágenes si las hay
    if images:
        for image_path in images:
            with open(image_path, 'rb') as image_file:
                img = MIMEImage(image_file.read())
                img.add_header('Content-Disposition', 'attachment', filename=os.path.basename(image_path))
                message.attach(img)

    # Iniciar conexión con el servidor SMTP y enviar el correo
    with smtplib.SMTP("smtp.office365.com", 587) as server:
        server.starttls()
        server.login("cnetransfer@cne.gob.do", "Cne.2022+")
        server.sendmail("cnetransfer@cne.gob.do", to_addresses, message.as_string())

# Verificar si hay empleados que cumplen años hoy
if employees_birthday_today:
    # Generar y enviar correo de cumpleaños general
    birthday_today_html = generate_birthday_list_html(employees_birthday_today, 'Cumpleaños_general.html', is_today=True)
    # Distinguir entre hombre y mujer para el mensaje de felicitación
    if len(employees_birthday_today) == 1:
        gender_prefix = "La colaboradora" if employees_birthday_today[0]['Sexo'] == 'F' else "El colaborador"
        general_subject = f"{gender_prefix} {employees_birthday_today[0]['Nombre']} está de cumpleaños hoy!"
    else:
        general_subject = "Los siguientes colaboradores están de cumpleaños hoy!"
    send_email(general_subject, birthday_today_html, ["dfajardo@cne.gob.do"])
else:
    print("No hay empleados que cumplan años hoy.")


# Obtener la fecha actual
today = datetime.datetime.now()

# Verificar si hoy es lunes (weekday() devuelve 0 para lunes, 1 para martes, etc.)
if today.weekday() == 3:
    for encargado, info in employees_birthday_next_week.items():
        employees = info['empleados']
        if employees:  # Si hay empleados que cumplen años en los próximos 7 días
            # Generar y enviar correo al encargado
            birthday_next_week_html = generate_birthday_list_html(employees, 'Cumpleaños_supervisor.html', is_today=False)
            supervisor_subject = "Estos son tus colaboradores que cumplen año durante la semana !"
            supervisor_email = info['correo']
            send_email(supervisor_subject, birthday_next_week_html, [supervisor_email])
else:
    print("Hoy no es lunes. No se enviarán correos a los encargados.")

# Función para codificar la imagen en base64
def encode_image_for_email(image_path):
    with open(image_path, 'rb') as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')

# Función para enviar correo electrónico con imagen incrustada en el cuerpo
def send_birthday_email_with_image(employee_email, image_path):
    # Codificar la imagen en base64
    encoded_image = encode_image_for_email(image_path)

    # Crear el HTML para el cuerpo del correo electrónico con la imagen incrustada
    email_html = f"""
    <html>
        <body>
            <img src="data:image/jpeg;base64,{encoded_image}" alt="Felicidades">
        </body>
    </html>
    """

    # Crear el mensaje de correo electrónico
    message = MIMEMultipart("alternative")
    message["Subject"] = f"Feliz Cumpleaños {employee['Nombre']} !"
    message["From"] = "cnetransfer@cne.gob.do"
    message["To"] = employee_email

    # Adjuntar el HTML al mensaje
    message.attach(MIMEText(email_html, "html"))

    # Iniciar conexión con el servidor SMTP y enviar el correo
    with smtplib.SMTP("smtp.office365.com", 587) as server:
        server.starttls()
        server.login("cnetransfer@cne.gob.do", "Cne.2022+")
        server.sendmail("cnetransfer@cne.gob.do", employee_email, message.as_string())

# Enviar imagen de felicitación a cada empleado que cumple años hoy
for employee in employees_birthday_today:
    send_birthday_email_with_image(f"{employee['Correo']}", "felicidades.jpeg")
