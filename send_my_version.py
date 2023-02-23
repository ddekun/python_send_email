import smtplib
import os
import time
import pandas as pd
import mimetypes
from tqdm import tqdm
from email import encoders
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase


def send_email(text=None, subject=None):
    # Адрес отправителя
    sender = os.getenv("sender")
    # Пароль отправителя. Мы получаем его из переменной окружения EMAIL_PASSWORD.
    # Если переменная не установлена, то вызываем исключение.
    password = os.getenv("EMAIL_PASSWORD")
    if not password:
        raise ValueError("EMAIL_PASSWORD не установлено в переменных окружения")

    # Читаем список email-адресов из Excel-файла.
    df = pd.read_excel("emails.xlsx")
    emails = df["emails"].tolist()
    # print(df.columns)
    # Читаем шаблон письма из HTML-файла.
    template_path = "Page_Template.html"
    try:
        with open(template_path, encoding="utf-8") as file:
            template = file.read()
    except FileNotFoundError:
        raise ValueError(f"Файл шаблона '{template_path}' не найден!")

    # Настраиваем подключение к SMTP-серверу Yandex
    server = smtplib.SMTP("smtp.yandex.ru", 587)
    server.starttls()

    try:
        # Авторизуемся на сервере.
        server.login(sender, password)

        # Отправляем сообщение каждому адресату из списка.
        for email in emails:
            msg = MIMEMultipart()
            msg["From"] = sender
            msg["To"] = email
            msg["Subject"] = subject

            if text:
                msg.attach(MIMEText(text))

            msg.attach(MIMEText(template, "html"))

            print("Collecting...")
            for file in tqdm(os.listdir("attachments")):
                time.sleep(0.4)
                filename = os.path.basename(file)
                ftype, encoding = mimetypes.guess_type(file)
                file_type, subtype = ftype.split("/")

                if file_type == "text":
                    with open(f"attachments/{file}") as f:
                        file = MIMEText(f.read())
                elif file_type == "image":
                    with open(f"attachments/{file}", "rb") as f:
                        file = MIMEImage(f.read(), subtype)
                elif file_type == "audio":
                    with open(f"attachments/{file}", "rb") as f:
                        file = MIMEAudio(f.read(), subtype)
                elif file_type == "application":
                    with open(f"attachments/{file}", "rb") as f:
                        file = MIMEApplication(f.read(), subtype)
                else:
                    with open(f"attachments/{file}", "rb") as f:
                        file = MIMEBase(file_type, subtype)
                        file.set_payload(f.read())
                        encoders.encode_base64(file)

                file.add_header('content-disposition', 'attachment', filename=filename)
                msg.attach(file)

            server.sendmail(sender, email, msg.as_string())

        # Закрываем соединение с SMTP-сервером и возвращаем сообщение об успешной отправке.
        server.quit()
        return "Сообщения успешно отправлены!"

    except smtplib.SMTPAuthenticationError as ex:
        raise ValueError(f"Аутентификация не удалась: {ex}")
    except Exception as ex:
        raise ValueError(f"Произошла ошибка при отправке сообщения: {ex}")


def main():
    subject = input("Введите тему сообщения: ")
    text = input("Введите текст сообщения: ")
    print(send_email(text=text, subject=subject))


if __name__ == "__main__":
    main()
input()