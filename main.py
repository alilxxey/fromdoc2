import telebot
import datetime
import requests
import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import smtplib
import os
from email import encoders
import errno
import re
import functools
import collections
import string


TOKEN = '1784219747:AAEad0gIWff6xDK6xVs28P3rnErr5PAk_eY'
TOMSG = 'dogovor@abc-analitika.ru'
bot = telebot.TeleBot(TOKEN)

ur = "https://script.google.com/macros/s/AKfycbzihEHTJg_LZCAhkwZJHxXLWx5U_aADnhICu0U0U-eVKtPiA6uqNIOV8aebvjU9HYM7/exec?"


class Client:
    def __init__(self):
        self.content = {'0': {
            'inn': None,
            'fullname': None,
            'registration': None,
            'series': None,
            'number': None,
            'snils': None,
            'cardnumber': None,
            'codep': None,
            'phonenumber': None,
            'given': None,
            'date': None,
            'email': None,
            'status': 0,
            'ph': 0
        }}

    def get(self, _id):
        return self.content[str(_id)]

    def create(self, _id):
        self.content[str(_id)] = self.content['0'].copy()

    def setcontent(self, _id, cont):
        _id = str(_id)
        cont = str(cont)
        if self.content[_id]['status'] == 0:
            self.content[_id]['fullname'] = cont
        elif self.content[_id]['status'] == 1:
            self.content[_id]['registration'] = cont
        elif self.content[_id]['status'] == 2:
            self.content[_id]['inn'] = cont
        elif self.content[_id]['status'] == 3:
            self.content[_id]['snils'] = cont
        elif self.content[_id]['status'] == 4:
            if len(cont.split()) == 2:
                self.content[_id]['series'] = cont.split()[0]
                self.content[_id]['number'] = cont.split()[1]
            if len(cont.split()) == 3:
                self.content[_id]['series'] = cont.split()[0] + cont.split()[1]
                self.content[_id]['number'] = cont.split()[2]
        elif self.content[_id]['status'] == 5:
            self.content[_id]['given'] = cont
        elif self.content[_id]['status'] == 6:
            self.content[_id]['date'] = cont
        elif self.content[_id]['status'] == 7:
            self.content[_id]['codep'] = cont
        elif self.content[_id]['status'] == 8:
            self.content[_id]['email'] = cont
        elif self.content[_id]['status'] == 9:
            self.content[_id]['phonenumber'] = cont
        elif self.content[_id]['status'] == 10:
            self.content[_id]['cardnumber'] = cont


client = Client()


def sendmsg(name, _id, tag):
    _id = str(_id)
    msg = MIMEMultipart()
    password = "PostBot530!"
    msg['From'] = "postbot@abc-analitika.ru"
    msg['To'] = TOMSG
    msg['Subject'] = client.content[str(_id)]['fullname']
    filepath = f'document{_id}.docx'
    filename = os.path.basename(filepath)
    ctype, encoding = mimetypes.guess_type(filepath)
    if ctype is None or encoding is not None:
        ctype = 'application/octet-stream'
    maintype, subtype = ctype.split('/', 1)
    with open(filepath, 'rb') as fp:
        file = MIMEBase(maintype, subtype)
        file.set_payload(fp.read())
        encoders.encode_base64(file)
    msg.attach(file)
    file.add_header('Content-Disposition', 'attachment', filename=filename)

    filepath = f'{_id}0.jpg'
    filename = os.path.basename(filepath)
    ctype, encoding = mimetypes.guess_type(filepath)
    if ctype is None or encoding is not None:
        ctype = 'application/octet-stream'
    maintype, subtype = ctype.split('/', 1)
    with open(filepath, 'rb') as fp:
        file = MIMEImage(fp.read(), _subtype=subtype)
    msg.attach(file)
    file.add_header('Content-Disposition', 'attachment', filename=filename)
    filepath = f'{_id}1.jpg'
    filename = os.path.basename(filepath)
    ctype, encoding = mimetypes.guess_type(filepath)
    if ctype is None or encoding is not None:
        ctype = 'application/octet-stream'
    maintype, subtype = ctype.split('/', 1)
    with open(filepath, 'rb') as fp:
        file = MIMEImage(fp.read(), _subtype=subtype)
    msg.attach(file)
    file.add_header('Content-Disposition', 'attachment', filename=filename)
    msg.attach(MIMEText(f"telegram: {tag}    phone: {client.content[_id]['phonenumber']}"
                        f"     mail: {client.content[_id]['email']}", 'plain'))
    server = smtplib.SMTP_SSL('smtp.yandex.ru: 465')
    print(1)
    server.login(msg['From'], password)
    print(2)
    server.send_message(msg)
    print(3)
    server.quit()
    pass


def renderdoc(_id, tag):
    params = ''
    print(1)
    for i, j in client.content.get(_id).items():
        params += f'{i}={j}&'
    dt = datetime.datetime.now().timetuple()
    yr, mnth, day = str(dt[0])[2:], dt[1], dt[2]
    params += f'yr={str(yr)}&mnth={str(mnth) if len(str(mnth)) == 2 else "0" + str(mnth)}&day={str(day)}'
    url = ur + params
    response = requests.get(url)
    print("file generated")
    response = requests.get(response.content)
    print("file downloaded")

    with open(f"document{_id}.docx", "wb") as f:
        f.write(response.content)
    print('sending')
    sendmsg(f"document{_id}.docx", _id, tag)
    print('sent')


@bot.message_handler(commands=['start'])
def start(message):
    print('started')
    bot.send_message(message.chat.id, 'Здравствуйте! Введите следующие данные:\nВаше ФИО')
    client.create(message.chat.id)


def genmess(status):
    if status == 1:
        return 'Введите адрес регистрации'
    elif status == 2:
        return 'Введите ИНН'
    elif status == 3:
        return 'Введите СНИЛС'
    elif status == 4:
        return 'Введите серию и номер паспорта через пробел (пример: 1234 123456)'
    elif status == 5:
        return 'Введите кем выдан паспорт'
    elif status == 6:
        return 'Введите дату выдачи паспорта'
    elif status == 7:
        return 'Введите код подразделения паспорта'
    elif status == 8:
        return 'Введите свой адрес эл.почты'
    elif status == 9:
        return 'Введите свой номер телефона'
    elif status == 10:
        return 'Введите номер своей банковской карты'


@bot.message_handler()
def getinf(message):
    print(message.text)
    client.setcontent(message.chat.id, message.text)
    _id = message.chat.id
    client.content[str(_id)]['status'] += 1
    if client.content[str(_id)]['status'] >= 11:
        bot.send_message(message.chat.id, 'Отправьте фото своего паспорта (только разворот с фотографией)')
    else:
        bot.send_message(message.chat.id, genmess(client.content[str(_id)]['status']))


@bot.message_handler(content_types=['photo'])
def handle_docs_photo(message):
    try:
        _id = message.chat.id
        file_info = bot.get_file(message.photo[len(message.photo)-1].file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        with open(str(message.chat.id) + str(client.content[str(_id)]['ph']) + '.' + file_info.file_path.split('.')[-1], 'wb') as new_file:
            new_file.write(downloaded_file)
        if client.content[str(_id)]['ph'] == 0:
            client.content[str(_id)]['ph'] = client.content[str(_id)]['ph'] + 1
            bot.send_message(message.chat.id, 'Отправьте фото своего паспорта (только разворот с пропиской)')
        else:
            bot.reply_to(message, "Спасибо, договор сформирован и отправлен территориальному менеджеру. "
                                  "В ближайшее время он с вами свяжется.")
            renderdoc(str(message.chat.id), tag=str(message.from_user.username))
    except Exception as e:
        pass


if __name__ == '__main__':
    bot.polling(none_stop=True, interval=3)

