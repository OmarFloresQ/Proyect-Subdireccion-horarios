from json import loads
from random import randint, choice
from datetime import datetime
import string
file = "/home/daniel/PycharmProjects/Horarios/teacher_availability.json"
with open(file, "r") as f:
    teachers_info = loads(f.read())

csv_formated_text = ""
for k, v in teachers_info.items():
    n = v["nombre"].split(" ")
    apellido_paterno = n.pop(0)
    apellidi_materno = n.pop(0)
    nombre = " ".join(n)
    username = (nombre.replace(" ", ".") + "".join([str(randint(0,9)) for _ in range(3)])).lower()
    current_datetime = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    remember_token = "".join([choice(string.digits + string.ascii_lowercase) for _ in range(15)])

    # name - last_name - last_materno - email - address - campus - uacademica -  username - password - level - created_at - updated_at - remember_token - unidad - ua - programaedu
    txt = f'"{k}","{nombre}","{apellido_paterno}","{apellidi_materno}","{v["email"]}","Calle Tamarido 123, Colonia Valentin","1","1","{username}","supersecurepass123","5","{current_datetime}","{current_datetime}","{remember_token}","1","1","1"\n'
    csv_formated_text += txt

print(csv_formated_text)
with open("/home/daniel/PycharmProjects/Horarios/users.csv", "w", encoding="utf-8") as csv_file:
    csv_file.write(csv_formated_text)
#LOAD DATA INFILE '/tmp/users.csv' INTO TABLE users FIELDS TERMINATED BY ',' ENCLOSED BY '"' LINES TERMINATED BY '\n' (id, name, last_name, last_materno, email, address, campus, uacademica, username, password, level, created_at, updated_at, remember_token, unidad, ua, programaedu);