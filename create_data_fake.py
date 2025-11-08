import pandas as pd
import random
from faker import Faker

# Inicializa Faker en español
fake = Faker('es_CO')

# Configuración de número de hojas y registros
num_hojas = 5
registros_por_hoja = 30
nombre_archivo = r"D:\dev\Control_vacunas\CUNAS_DATOS_FICTICIOS.xlsx"

# Función para generar datos ficticios
def generar_datos_ficticios(n):
    datos = []
    for _ in range(n):
        nombre = fake.first_name()
        apellido = fake.last_name()
        identificacion = random.randint(10000000, 99999999)
        fecha_nacimiento = fake.date_of_birth(minimum_age=0, maximum_age=10)
        direccion = fake.address().replace("\n", ", ")
        correo = fake.email()
        telefono = fake.phone_number()
        acudiente = fake.name()
        datos.append({
            "NOMBRE": f"{nombre} {apellido}",
            "IDENTIFICACION": identificacion,
            "FECHA_NACIMIENTO": fecha_nacimiento,
            "DIRECCION": direccion,
            "CORREO": correo,
            "TELEFONO": telefono,
            "ACUDIENTE": acudiente
        })
    return pd.DataFrame(datos)

# Crear un Excel con 5 hojas diferentes
with pd.ExcelWriter(nombre_archivo, engine="openpyxl") as writer:
    for i in range(1, num_hojas + 1):
        hoja = f"CUNAS {i}"
        df = generar_datos_ficticios(registros_por_hoja)
        df.to_excel(writer, index=False, sheet_name=hoja)

print(f"✅ Archivo '{nombre_archivo}' creado correctamente con {num_hojas} hojas.")
