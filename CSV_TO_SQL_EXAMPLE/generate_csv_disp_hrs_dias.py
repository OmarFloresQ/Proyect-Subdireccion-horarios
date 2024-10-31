from json import loads
from random import randint, choice
from datetime import datetime
import string
from re import search
file = "/home/daniel/PycharmProjects/Horarios/teacher_availability.json"
with open(file, "r") as f:
    teachers_info = loads(f.read())

csv_formated_text = ""
days_map = {
    "Monday": 1,
    "Tuesday": 2,
    "Wednesday": 3,
    "Thursday": 4,
    "Friday": 5,
    "Saturday": 6
}
# table scheme -> periodo, id, dia, hora, status, observaciones, users_id
for userid, v in teachers_info.items():
    for hour_number, dv in v["disponibilidad"].items():
        for day_number, availability_status in dv.items():
            if availability_status != "Not Available":
                #num_detected = search(r'\d', availability_status)
                #if num_detected:
                csv_formated_text += f'"20241","{userid}","{day_number}","{hour_number}","1","{userid}"\n'

with open("disp_hrs_dias.csv", "w") as csv_file:
    csv_file.write(csv_formated_text)

#LOAD DATA INFILE '/tmp/disp_hrs_dias.csv' INTO TABLE disp_hrs_dias FIELDS TERMINATED BY ',' ENCLOSED BY '"' LINES TERMINATED BY '\n' (periodo, id, dia, hora, status, users_id);