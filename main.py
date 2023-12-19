import imaplib
import email
import time
import datetime as dt
import os
from email.header import decode_header
from send_to_tg import send_to_tg
import config

# Параметры для доступа к почтовому ящику


#Почта Акимовой
akimova_mail, my_mail = 'ms.akimova.61@mail.ru', "kirillzg@yandex.ru"

# Период проверки новых писем (в секундах)
check_interval = 50  # 1 минута

def check_email():
    try:
        # Подключение к серверу IMAP
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(config.login, config.password)
        mail.select('inbox')

        # Поиск новых писем
        result, data = mail.search(None, "UNSEEN")
        if result == "OK":
            new_email_ids = data[0].split()
            for email_id in new_email_ids:
                # Получение данных письма
                result, email_data = mail.fetch(email_id, "(RFC822)")
                if result == "OK":
                    raw_email = email_data[0][1]
                    msg = email.message_from_bytes(raw_email)

                    # Проверка отправителя
                    sender = msg["From"]
                    if akimova_mail or my_mail in sender:

                        # Проверка вложений
                        if msg.get_content_maintype() == "multipart":
                            for part in msg.walk():
                                if part.get_content_maintype() == "multipart" or part.get_content_maintype() == "text":
                                    continue
                                filename = part.get_filename()
                                if filename:
                                    # Декодирование имени файла
                                    filename, charset = decode_header(filename)[0]
                                    if charset:
                                        filename = filename.decode(charset)
                                    file_path = os.path.join('', filename)

                                    if "Изменения в расписании - 3.xlsx" in filename:
                                        file_path = os.path.join('', filename)

                                        # Скачивание вложения
                                        with open(file_path, "wb") as file:
                                            file.write(part.get_payload(decode=True))
                                        print(f"Скачан файл: {filename}")
                                        send_to_tg()

    except Exception as e:
        print(f"Ошибка при проверке почты: {e}")


if __name__ == "__main__":
    while True:
        if dt.datetime.now().hour >= 17 and config.count == 0:
            config.after_17 = 1
            config.count = 1
            send_to_tg()
        else:
            if dt.datetime.now().hour >= 17:
                config.after_17 = 1
                check_email()
                time.sleep(check_interval)
            else:
                config.count = 0
                config.after_17 = 0
                check_email()
                time.sleep(check_interval)

