from vk_api.bot_longpoll import VkBotLongPoll, VkBotEventType
import vk_api
import datetime
import os
import urllib
import wget
import xlrd
import time

vk = vk_api.VkApi(token="fcc1b4e4e1a351c3e185047a885b13d0dcc0819a8f188d78a068f333d92fe372ff45cb4ccbf7cbad031ea")

vk._auth_token()

vk.get_api()

longpoll = VkBotLongPoll(vk, 192023320)

while True:
    for event in longpoll.listen():
        if event.type == VkBotEventType.MESSAGE_NEW:
            if event.object.peer_id != event.object.from_id:
                userText = event.object.text.lower()
                if userText == "!#":

                    date = datetime.datetime.now()

                    day = date.day
                    dayPlus1 = date.day + 1
                    dayPlus2 = date.day + 2
                    month = date.month
                    year = str(date.year)

                    d = datetime.date(int(year), int(month), int(day))
                    dayOfWeek = d.isoweekday()

                    dayNikita = ""
                    coupleFile = ""

                    # Скачивание .xls файла для дальнейшего чтения
                    if dayOfWeek == 6:
                        if month < 10:
                            month = "0" + str(month)

                        if dayPlus1 < 10:
                            day = "0" + str(dayPlus1)

                        link = "http://www.urtk-mephi.ru/Raspisan/" + str(day) + "." + str(month) + ".20" + ".xls"
                        linkPlus1 = "http://www.urtk-mephi.ru/Raspisan/" + str(dayPlus2) + "." + str(
                            month) + ".20" + ".xls"
                    else:
                        if month < 10:
                            month = "0" + str(month)

                        if dayPlus1 < 10:
                            day = "0" + str(dayPlus1)

                        link = "http://www.urtk-mephi.ru/Raspisan/" + str(day) + "." + str(month) + ".20" + ".xls"
                        linkPlus1 = "http://www.urtk-mephi.ru/Raspisan/" + str(dayPlus1) + "." + str(
                            month) + ".20" + ".xls"

                    try:
                        file = wget.download(linkPlus1, "day_temp.xls")

                        availabilityFile = os.path.isfile("day.xls")
                        if availabilityFile == True:
                            os.remove("day.xls")

                        os.rename(file, "day.xls")

                        if dayOfWeek == 6:
                            dayNikita = dayPlus2
                        else:
                            dayNikita = dayPlus1
                    except urllib.error.HTTPError:

                        if dayOfWeek == 6:
                            coupleFile = "Расписание 1С1 на " + str(dayPlus2) + "." + str(month) + "." + str(
                                year) + " не найдено, есть только на " + str(day) + "." + str(
                                month) + "." + str(year)
                        else:
                            coupleFile = "Расписание 1С1 на " + str(dayPlus1) + "." + str(month) + "." + str(
                                year) + " не найдено, есть только на " + str(day) + "." + str(
                                month) + "." + str(year)

                        file = wget.download(link, "day_temp.xls")

                        availabilityFile = os.path.isfile("day.xls")
                        if availabilityFile == True:
                            os.remove("day.xls")

                        os.rename(file, "day.xls")
                        dayNikita = day

                    # Проверка на чётность недели
                    today = datetime.datetime.today()
                    w = int(today.strftime("%U"))

                    # Если сегодня суббота нечётной (или чётной) недели, то стандартное расписание на понедельник
                    # будет выводится НЕЧЁТНОЕ (или чётное), а понедельник уже чётное (или нечётное) число.

                    # В субботу к переменной w прибавляется 1, дабы на понедельник выводилось правильное расписание.
                    if dayOfWeek == 6:
                        w += 1

                    if w % 2 == 0:
                        ParityOfTheWeek = True
                    else:
                        ParityOfTheWeek = False

                    # Чтение .xls файла
                    try:
                        xls = xlrd.open_workbook("day.xls")
                        sheet = xls.sheet_by_index(0)
                    except FileNotFoundError:
                        print("Изменения нету:")

                    group1 = sheet.cell(2, 0).value
                    group2 = sheet.cell(2, 3).value
                    group3 = sheet.cell(2, 6).value
                    group4 = sheet.cell(2, 9).value

                    if group1 == 'Группа 1C1':
                        couple1 = "1. " + sheet.cell(3, 1).value
                        couple2 = "2. " + sheet.cell(4, 1).value
                        couple3 = "3. " + sheet.cell(5, 1).value
                        couple4 = "4. " + sheet.cell(6, 1).value
                    elif group2 == 'Группа 1C1':
                        couple1 = "1. " + sheet.cell(3, 4).value
                        couple2 = "2. " + sheet.cell(4, 4).value
                        couple3 = "3. " + sheet.cell(5, 4).value
                        couple4 = "4. " + sheet.cell(6, 4).value
                    elif group3 == 'Группа 1C1':
                        couple1 = "1. " + sheet.cell(3, 7).value
                        couple2 = "2. " + sheet.cell(4, 7).value
                        couple3 = "3. " + sheet.cell(5, 7).value
                        couple4 = "4. " + sheet.cell(6, 7).value
                    elif group4 == 'Группа 1C1':
                        couple1 = "1. " + sheet.cell(3, 10).value
                        couple2 = "2. " + sheet.cell(4, 10).value
                        couple3 = "3. " + sheet.cell(5, 10).value
                        couple4 = "4. " + sheet.cell(6, 10).value
                    else:

                        '''
                        1 - Понедельник
                        2 - Вторник
                        3 - Среда
                        4 - Четверг
                        5 - Пятница
                        6 - Суббота
                        7 - Воскресенье
                        '''

                        time = date.time().hour

                        if time < 11:
                            if ParityOfTheWeek:
                                if dayOfWeek == 1:
                                    couple1 = "1. Обществознание"
                                    couple2 = "2. Математика"
                                    couple3 = "3. История"
                                    couple4 = "4."
                                elif dayOfWeek == 2:
                                    couple1 = "1. Физкультура"
                                    couple2 = "2. Обществознание"
                                    couple3 = "3. Ин.яз"
                                    couple4 = "4. "
                                elif dayOfWeek == 3:
                                    couple1 = "1. Математика"
                                    couple2 = "2. Физика"
                                    couple3 = "3. Обж"
                                    couple4 = "4."
                                elif dayOfWeek == 4:
                                    couple1 = "1. Математика"
                                    couple2 = "2. Русский яз."
                                    couple3 = "3. Информатика"
                                    couple4 = "4."
                                elif dayOfWeek == 5:
                                    couple1 = "1. Литература"
                                    couple2 = "2. Математика"
                                    couple3 = "3. Физика"
                                    couple4 = "4. ОБЖ"
                                elif dayOfWeek == 6:
                                    couple1 = "1. Астрономия"
                                    couple2 = "2. Химия"
                                    couple3 = ""
                                    couple4 = "4."
                                elif dayOfWeek == 1:
                                    couple1 = "1. Обществознание"
                                    couple2 = "2. Математика"
                                    couple3 = "3. История"
                                    couple4 = "4."
                            elif not ParityOfTheWeek:
                                if dayOfWeek == 1:
                                    couple1 = "1. Физика"
                                    couple2 = "2. Математика"
                                    couple3 = "3. Астрономия"
                                    couple4 = "4."
                                elif dayOfWeek == 2:
                                    couple1 = "1. Физкультура"
                                    couple2 = "2. Физика"
                                    couple3 = "3. Обществознание"
                                    couple4 = "4. Ин.яз"
                                elif dayOfWeek == 3:
                                    couple1 = "1. Математика"
                                    couple2 = "2. История"
                                    couple3 = "3. Химия"
                                    couple4 = "4."
                                elif dayOfWeek == 4:
                                    couple1 = "1. Математика"
                                    couple2 = "2. Русский яз."
                                    couple3 = "3. Физкультура"
                                    couple4 = "4."
                                elif dayOfWeek == 5:
                                    couple1 = "1. Литература"
                                    couple2 = "2. Математика"
                                    couple3 = "3. Информатика"
                                    couple4 = "4. "
                                elif dayOfWeek == 6:
                                    couple1 = "1. История"
                                    couple2 = "2. ОБЖ"
                                    couple3 = ""
                                    couple4 = ""
                                elif dayOfWeek == 1:
                                    couple1 = "1. Обществознание"
                                    couple2 = "2. Математика"
                                    couple3 = "3. История"
                                    couple4 = ""
                        else:
                            if ParityOfTheWeek:
                                if dayOfWeek == 7:
                                    couple1 = "1. Обществознание"
                                    couple2 = "2. Математика"
                                    couple3 = "3. История"
                                    couple4 = ""
                                elif dayOfWeek == 1:
                                    couple1 = "1. Физкультура"
                                    couple2 = "2. Обществознание"
                                    couple3 = "3. Ин.ях"
                                    couple4 = "4."
                                elif dayOfWeek == 2:
                                    couple1 = "1. Математика"
                                    couple2 = "2. Физика"
                                    couple3 = "3. ОБЖ"
                                    couple4 = ""
                                elif dayOfWeek == 3:
                                    couple1 = "1. Математика"
                                    couple2 = "2. Русский яз."
                                    couple3 = "3. Информатика"
                                    couple4 = ""
                                elif dayOfWeek == 4:
                                    couple1 = "1. Литература"
                                    couple2 = "2. Математика"
                                    couple3 = "3. Физика"
                                    couple4 = "4. ОБЖ"
                                elif dayOfWeek == 5:
                                    couple1 = "1. Астрономия"
                                    couple2 = "2. Химия"
                                    couple3 = ""
                                    couple4 = ""
                                elif dayOfWeek == 6:
                                    couple1 = "1. Обществознание"
                                    couple2 = "2. Математика"
                                    couple3 = "3. История"
                                    couple4 = ""
                            elif not ParityOfTheWeek:
                                if dayOfWeek == 7:
                                    couple1 = "1. Физика"
                                    couple2 = "2. Математика"
                                    couple3 = "3. Астрономия"
                                    couple4 = ""
                                elif dayOfWeek == 1:
                                    couple1 = "1. Физкультура"
                                    couple2 = "2. Физика"
                                    couple3 = "3. Обществознание"
                                    couple4 = "4. Ин.яз"
                                elif dayOfWeek == 2:
                                    couple1 = "1. Математика"
                                    couple2 = "2. История"
                                    couple3 = "3. Химия"
                                    couple4 = "4."
                                elif dayOfWeek == 3:
                                    couple1 = "1. Математика"
                                    couple2 = "2. Русский яз."
                                    couple3 = "3. Физкультура"
                                    couple4 = "4."
                                elif dayOfWeek == 4:
                                    couple1 = "1. Литература"
                                    couple2 = "2. Математика"
                                    couple3 = "3. Информатика"
                                    couple4 = "4. "
                                elif dayOfWeek == 5:
                                    couple1 = "1. История"
                                    couple2 = "2. ОБЖ"
                                    couple3 = "3."
                                    couple4 = "4."
                                elif dayOfWeek == 6:
                                    couple1 = "1. Обществознание"
                                    couple2 = "2. Математика"
                                    couple3 = "3. История"
                                    couple4 = "4."
                    if coupleFile == "":
                        vk.method("messages.send", {"peer_id": event.object.peer_id,
                                                    "message": "Расписание 1С1 на " + str(dayNikita) + "." + str(month) + "." + str(year) + "\n" + "\n\n" + couple1 + "\n" + couple2 + "\n" + couple3 + "\n" + couple4,
                                                    "random_id": 0})

                    elif coupleFile != "":
                        vk.method("messages.send", {"peer_id": event.object.peer_id,
                                                    "message": coupleFile + "\n\n" + couple1 + "\n" + couple2 + "\n" + couple3 + "\n" + couple4,
                                                    "random_id": 0})
