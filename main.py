#TODO
# Запрос "Соревнования", вывод информации о соревнованиях
# Запрос "Мероприятия", вывод информации о соревнованиях
# Запрос "Занятие", вывод темы
# Запрос "Теория", вывод ссылок на теорию
# Запрос "Справка", вывод информации по боту
# Запрос "Конференция", добавление участника в нужную конференцию
# Рассылки всем, кто когда-либо пользовался ботом
# Создание аттестационных листов

import vk_api
from vk_api.utils import get_random_id
from vk_api.bot_longpoll import VkBotLongPoll, VkBotEventType
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import re

GOOGLE_API_ACCESS_TOKEN = 'project-pocus-c50cb81abfc0.json'
VK_API_ACCESS_TOKEN = "62055792611d923155246f1fcbb4d757cf12739ccaa6fca122cf01bbfdc0f79a2d7c7102c3c2ef9f3c22f"
GROUP_ID = "190125423"

#Реализация google sheets api
class GoogleSheets:

    def __init__(self):
        self.authorize()
        self.gspread_auth_start_time = datetime.datetime.now()

    def authorize(self):
        self.scope = ['https://spreadsheets.google.com/feeds',
                 'https://www.googleapis.com/auth/drive']
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_API_ACCESS_TOKEN, self.scope)
        self.gc = gspread.authorize(self.credentials)
        self.sh = self.gc.open("Рейтинг Кванториум 2019-2020")

    def req_rating(self, name):
        if datetime.datetime.now() > self.gspread_auth_start_time + datetime.timedelta(minutes=10):
            self.authorize()
            self.gspread_auth_start_time = datetime.datetime.now()
        for wks in self.sh.worksheets():
            if wks.title.find('sorted') != -1:
                try:
                    name_re = re.compile(name + r'.*')
                    cell = wks.find(name_re)
                    return 'Группа: %s \nУченик: %s' % (wks.title[:wks.title.find('_')], name) + \
                        '\n\nМесто в рейтинге: ' + str(wks.cell(cell.row, 1).value) + \
                        '\nОбщее количество баллов: ' + str(wks.cell(cell.row, 7).value) + \
                        '\n\nУчастие в соревнованиях:' + \
                        '\n\nМуниципального уровня - ' + str(wks.cell(cell.row, 3).value) + \
                        '\nРегионального уровня - ' + str(wks.cell(cell.row, 4).value) + \
                        '\nВсероссийского уровня - ' + str(wks.cell(cell.row, 5).value) + \
                        '\nМеждународного уровня - ' + str(wks.cell(cell.row, 6).value)
                except Exception:
                    continue
        return 'Пользователя %s нет в рейтинговой таблице' % name
            

    def req_mark(self, name):
        if datetime.datetime.now() > self.gspread_auth_start_time + datetime.timedelta(minutes=10):
            self.authorize()
            self.gspread_auth_start_time = datetime.datetime.now()
        for wks in self.sh.worksheets():
            if wks.title.find('sorted') != -1:
                try:
                    name_re = re.compile(name + r'.*')
                    cell_name = wks.find(name_re)
                    date = datetime.datetime.now()
                    for i in range(7):
                        try:
                            date_find = date - datetime.timedelta(i)
                            date_format = date_find.strftime('%d/%m/%Y')
                            date_format = date_format.replace('/', '.')
                            cell_date = wks.find(date_format)
                            return 'Группа: %s \nУченик: %s' % (wks.title[:wks.title.find('_')], name) + \
                                   '\nОценка за занятие %s: %s' % (str(cell_date.value),
                                                                   wks.cell(cell_name.row, cell_date.col).value)
                        except Exception:
                            continue
                    return 'Занятий нет'
                except Exception:
                    continue
        return 'Пользователя %s нет в рейтинговой таблице' % name


#Реализация бота
class BotApi:

    vk_session = vk_api.VkApi(
        token=VK_API_ACCESS_TOKEN
    )
    vk = vk_session.get_api()
    longpoll = VkBotLongPoll(vk_session, GROUP_ID)
    gs = GoogleSheets()

    def run(self):
        for event in self.longpoll.listen():
            if event.type == VkBotEventType.MESSAGE_NEW:
                msg_low = event.obj.message['text'].lower()
                profile = self.vk.users.get(user_ids=event.obj.message['from_id'])
                if msg_low.startswith('!рейтинг'):
                    if event.from_user: # проверяем пришло сообщение от пользователя или нет
                        self.vk.messages.send(
                            user_id=event.obj.message['from_id'],
                            random_id=get_random_id(),
                            message=self.gs.req_rating(profile[0]['last_name'] + ' ' + profile[0]['first_name'])
                        )
                    if event.from_chat: # проверяем пришло сообщение из чата
                        self.vk.messages.send(
                            chat_id=event.chat_id,
                            random_id=get_random_id(),
                            message=self.gs.req_rating(profile[0]['last_name'] + ' ' + profile[0]['first_name'])
                        )
                elif msg_low.startswith('!оценка'):
                    if event.from_user: # проверяем пришло сообщение от пользователя или нет
                        self.vk.messages.send(
                            user_id=event.obj.message['from_id'],
                            random_id=get_random_id(),
                            message=self.gs.req_mark(profile[0]['last_name'] + ' ' + profile[0]['first_name'])
                        )
                    if event.from_chat: # проверяем пришло сообщение из чата
                        self.vk.messages.send(
                            chat_id=event.chat_id,
                            random_id=get_random_id(),
                            message=self.gs.req_mark(profile[0]['last_name'] + ' ' + profile[0]['first_name'])
                        )

pocus = BotApi()
pocus.run()
