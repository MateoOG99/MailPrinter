import os
import sys
import json
import uuid
import time
import queue
import shutil
import base64
import imaplib
import smtplib
import threading
import subprocess

from pathlib import Path

import pymupdf as fitz

from email import policy
from email.header import decode_header, make_header
from email.parser import BytesParser
from email.message import EmailMessage
from email.utils import parseaddr

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText


CONFIG_FILE = "mail_printer_config.json"


# =========================================================
# СЛУЖЕБНЫЕ ФУНКЦИИ
# =========================================================

os.environ.setdefault(
    "PYTHONIOENCODING",
    "utf-8"
)

if sys.platform.startswith("linux"):
    os.environ.setdefault(
        "LANG",
        "C.UTF-8"
    )

    os.environ.setdefault(
        "LC_ALL",
        "C.UTF-8"
    )


def command_environment():
    """
    Окружение для LibreOffice и CUPS.
    """

    environment = os.environ.copy()

    if sys.platform.startswith("linux"):
        environment["LANG"] = "C.UTF-8"
        environment["LC_ALL"] = "C.UTF-8"

    return environment


def run_command(command):
    """
    Запускает внешнюю команду с поддержкой UTF-8.
    """

    return subprocess.run(
        command,
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=command_environment()
    )


def decode_text(value):
    """
    Декодирует тему письма или имя вложения.
    """

    if not value:
        return ""

    try:
        return str(
            make_header(
                decode_header(value)
            )
        )
    except Exception:
        return str(value)


def safe_filename(filename):
    """
    Убирает опасные символы из имени файла.
    """

    filename = os.path.basename(filename)

    for symbol in '<>:"/\\|?*':
        filename = filename.replace(
            symbol,
            "_"
        )

    if not filename:
        filename = "attachment.bin"

    return filename


def encode_imap_utf7(text):
    """
    Кодирует название русской папки
    в modified UTF-7 для IMAP.
    """

    result = []
    unicode_buffer = []

    def flush_unicode_buffer():
        if not unicode_buffer:
            return

        source_text = "".join(
            unicode_buffer
        )

        encoded = source_text.encode(
            "utf-16-be"
        )

        encoded = base64.b64encode(
            encoded
        ).decode("ascii")

        encoded = encoded.rstrip("=")
        encoded = encoded.replace(
            "/",
            ","
        )

        result.append(
            "&" + encoded + "-"
        )

        unicode_buffer.clear()

    for character in text:
        code = ord(character)

        if (
            0x20 <= code <= 0x7E
            and character != "&"
        ):
            flush_unicode_buffer()
            result.append(character)

        elif character == "&":
            flush_unicode_buffer()
            result.append("&-")

        else:
            unicode_buffer.append(character)

    flush_unicode_buffer()

    return "".join(result)


class MailPrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MailPrinter")
        self.root.geometry("950x760")

        self.stop_event = threading.Event()
        self.worker_thread = None
        self.log_queue = queue.Queue()

        self.last_mailbox_message = None

        self.settings = {
            "imap_server": tk.StringVar(
                value="imap.yandex.ru"
            ),

            "imap_port": tk.StringVar(
                value="993"
            ),

            "mailbox": tk.StringVar(
                value="АвтоАльфа"
            ),

            "smtp_server": tk.StringVar(
                value="smtp.yandex.ru"
            ),

            "smtp_port": tk.StringVar(
                value="465"
            ),

            "email": tk.StringVar(),

            "password": tk.StringVar(),

            "subject_prefix": tk.StringVar(
                value=(
                    'Автоматическое уведомление от ООО ГК '
                    '"АВТОАЛЬФА" Заказ № Заказ поставщику'
                )
            ),

            "allowed_senders": tk.StringVar(),

            "save_folder": tk.StringVar(
                value=str(
                    Path.cwd() / "attachments"
                )
            ),

            "printer": tk.StringVar(
                value="HP-LaserJet-M402dn"
            ),

            "check_interval": tk.StringVar(
                value="120"
            ),

            "reply_text": tk.StringVar(
                value=(
                    "Здравствуйте!\n\n"
                    "Ваш файл получен и отправлен на печать."
                )
            ),
        }

        self.create_interface()
        self.load_settings()

        self.root.after(
            200,
            self.process_log_queue
        )

    # =========================================================
    # ИНТЕРФЕЙС
    # =========================================================

    def create_interface(self):
        main_frame = ttk.Frame(
            self.root,
            padding=10
        )

        main_frame.pack(
            fill="both",
            expand=True
        )

        settings_frame = ttk.LabelFrame(
            main_frame,
            text="Настройки"
        )

        settings_frame.pack(
            fill="x",
            padx=5,
            pady=5
        )

        row = 0

        self.add_entry(
            settings_frame,
            row,
            "IMAP-сервер:",
            self.settings["imap_server"]
        )
        row += 1

        self.add_entry(
            settings_frame,
            row,
            "IMAP-порт:",
            self.settings["imap_port"]
        )
        row += 1

        self.add_entry(
            settings_frame,
            row,
            "Папка почты:",
            self.settings["mailbox"]
        )
        row += 1

        ttk.Label(
            settings_frame,
            text=(
                "Например: АвтоАльфа. "
                "Для входящих используйте INBOX."
            )
        ).grid(
            row=row,
            column=1,
            sticky="w",
            padx=5,
            pady=(0, 5)
        )
        row += 1

        self.add_entry(
            settings_frame,
            row,
            "SMTP-сервер:",
            self.settings["smtp_server"]
        )
        row += 1

        self.add_entry(
            settings_frame,
            row,
            "SMTP-порт:",
            self.settings["smtp_port"]
        )
        row += 1

        self.add_entry(
            settings_frame,
            row,
            "Email аккаунта:",
            self.settings["email"]
        )
        row += 1

        ttk.Label(
            settings_frame,
            text="Пароль приложения:"
        ).grid(
            row=row,
            column=0,
            sticky="w",
            padx=5,
            pady=3
        )

        ttk.Entry(
            settings_frame,
            textvariable=self.settings["password"],
            show="*",
            width=70
        ).grid(
            row=row,
            column=1,
            sticky="ew",
            padx=5,
            pady=3
        )
        row += 1

        self.add_entry(
            settings_frame,
            row,
            "Постоянная фраза в теме:",
            self.settings["subject_prefix"]
        )
        row += 1

        ttk.Label(
            settings_frame,
            text=(
                "Обрабатываются темы, которые начинаются "
                "с этой фразы."
            )
        ).grid(
            row=row,
            column=1,
            sticky="w",
            padx=5,
            pady=(0, 5)
        )
        row += 1

        self.add_entry(
            settings_frame,
            row,
            "Разрешенные отправители:",
            self.settings["allowed_senders"]
        )
        row += 1

        ttk.Label(
            settings_frame,
            text=(
                "Несколько адресов укажите через запятую. "
                "Пустое поле — разрешены все отправители."
            )
        ).grid(
            row=row,
            column=1,
            sticky="w",
            padx=5,
            pady=(0, 5)
        )
        row += 1

        ttk.Label(
            settings_frame,
            text="Папка временных файлов:"
        ).grid(
            row=row,
            column=0,
            sticky="w",
            padx=5,
            pady=3
        )

        folder_frame = ttk.Frame(
            settings_frame
        )

        folder_frame.grid(
            row=row,
            column=1,
            sticky="ew",
            padx=5,
            pady=3
        )

        ttk.Entry(
            folder_frame,
            textvariable=self.settings[
                "save_folder"
            ],
            width=70
        ).pack(
            side="left",
            fill="x",
            expand=True
        )

        ttk.Button(
            folder_frame,
            text="Выбрать",
            command=self.select_folder
        ).pack(
            side="left",
            padx=5
        )
        row += 1

        self.add_entry(
            settings_frame,
            row,
            "Имя принтера:",
            self.settings["printer"]
        )
        row += 1

        ttk.Label(
            settings_frame,
            text=(
                "Имя можно узнать командой lpstat -p. "
                "Пустое поле — принтер по умолчанию."
            )
        ).grid(
            row=row,
            column=1,
            sticky="w",
            padx=5,
            pady=(0, 5)
        )
        row += 1

        self.add_entry(
            settings_frame,
            row,
            "Интервал проверки, секунд:",
            self.settings["check_interval"]
        )

        settings_frame.columnconfigure(
            1,
            weight=1
        )

        reply_frame = ttk.LabelFrame(
            main_frame,
            text="Текст ответа клиенту"
        )

        reply_frame.pack(
            fill="x",
            padx=5,
            pady=5
        )

        ttk.Entry(
            reply_frame,
            textvariable=self.settings[
                "reply_text"
            ],
            width=120
        ).pack(
            fill="x",
            padx=5,
            pady=5
        )

        buttons_frame = ttk.Frame(
            main_frame
        )

        buttons_frame.pack(
            fill="x",
            padx=5,
            pady=5
        )

        ttk.Button(
            buttons_frame,
            text="Сохранить настройки",
            command=self.save_settings
        ).pack(
            side="left",
            padx=5
        )

        ttk.Button(
            buttons_frame,
            text="Запустить",
            command=self.start_monitoring
        ).pack(
            side="left",
            padx=5
        )

        ttk.Button(
            buttons_frame,
            text="Остановить",
            command=self.stop_monitoring
        ).pack(
            side="left",
            padx=5
        )

        ttk.Button(
            buttons_frame,
            text="Выход",
            command=self.close_program
        ).pack(
            side="right",
            padx=5
        )

        log_frame = ttk.LabelFrame(
            main_frame,
            text="Журнал работы"
        )

        log_frame.pack(
            fill="both",
            expand=True,
            padx=5,
            pady=5
        )

        self.log_text = ScrolledText(
            log_frame,
            state="disabled"
        )

        self.log_text.pack(
            fill="both",
            expand=True
        )

    def add_entry(
        self,
        parent,
        row,
        label,
        variable
    ):
        ttk.Label(
            parent,
            text=label
        ).grid(
            row=row,
            column=0,
            sticky="w",
            padx=5,
            pady=3
        )

        ttk.Entry(
            parent,
            textvariable=variable,
            width=70
        ).grid(
            row=row,
            column=1,
            sticky="ew",
            padx=5,
            pady=3
        )

    # =========================================================
    # ЖУРНАЛ
    # =========================================================

    def write_log(self, text):
        self.log_queue.put(
            str(text)
        )

    def process_log_queue(self):
        while not self.log_queue.empty():
            text = self.log_queue.get()

            self.log_text.configure(
                state="normal"
            )

            self.log_text.insert(
                "end",
                text + "\n"
            )

            self.log_text.see(
                "end"
            )

            self.log_text.configure(
                state="disabled"
            )

        self.root.after(
            200,
            self.process_log_queue
        )

    # =========================================================
    # НАСТРОЙКИ
    # =========================================================

    def save_settings(self):
        data = {}

        for name, variable in self.settings.items():
            if name != "password":
                data[name] = variable.get()

        try:
            with open(
                CONFIG_FILE,
                "w",
                encoding="utf-8"
            ) as file:
                json.dump(
                    data,
                    file,
                    ensure_ascii=False,
                    indent=4
                )

            self.write_log(
                "Настройки сохранены."
            )

        except Exception as error:
            messagebox.showerror(
                "Ошибка",
                f"Не удалось сохранить настройки:\n{error}"
            )

    def load_settings(self):
        if not os.path.exists(
            CONFIG_FILE
        ):
            return

        try:
            with open(
                CONFIG_FILE,
                "r",
                encoding="utf-8"
            ) as file:
                data = json.load(file)

            for name, value in data.items():
                if name in self.settings:
                    self.settings[
                        name
                    ].set(value)

            self.write_log(
                "Настройки загружены."
            )

        except Exception as error:
            self.write_log(
                f"Не удалось загрузить настройки: "
                f"{error}"
            )

    def select_folder(self):
        folder = filedialog.askdirectory()

        if folder:
            self.settings[
                "save_folder"
            ].set(folder)

    # =========================================================
    # ЗАПУСК И ОСТАНОВКА
    # =========================================================

    def start_monitoring(self):
        if (
            self.worker_thread
            and self.worker_thread.is_alive()
        ):
            self.write_log(
                "Мониторинг уже запущен."
            )
            return

        if not self.settings[
            "email"
        ].get().strip():
            messagebox.showwarning(
                "Проверка",
                "Введите адрес электронной почты."
            )
            return

        if not self.settings[
            "password"
        ].get():
            messagebox.showwarning(
                "Проверка",
                "Введите пароль приложения."
            )
            return

        if not self.settings[
            "subject_prefix"
        ].get().strip():
            messagebox.showwarning(
                "Проверка",
                "Введите постоянную фразу из темы письма."
            )
            return

        try:
            interval = int(
                self.settings[
                    "check_interval"
                ].get()
            )

            if interval < 30:
                raise ValueError

        except ValueError:
            messagebox.showwarning(
                "Проверка",
                "Интервал должен быть числом "
                "не меньше 30 секунд."
            )
            return

        folder = Path(
            self.settings[
                "save_folder"
            ].get()
        )

        folder.mkdir(
            parents=True,
            exist_ok=True
        )

        self.stop_event.clear()

        self.worker_thread = threading.Thread(
            target=self.monitoring_loop,
            daemon=True
        )

        self.worker_thread.start()

        self.write_log(
            "Мониторинг почты запущен."
        )

    def stop_monitoring(self):
        self.stop_event.set()

        self.write_log(
            "Получена команда остановки."
        )

    def close_program(self):
        self.stop_event.set()
        self.root.destroy()

    # =========================================================
    # МОНИТОРИНГ ПОЧТЫ
    # =========================================================

    def monitoring_loop(self):
        try:
            interval = int(
                self.settings[
                    "check_interval"
                ].get()
            )
        except ValueError:
            interval = 120

        while not self.stop_event.is_set():
            try:
                self.check_mail()

                self.write_log(
                    f"Следующая проверка через "
                    f"{interval} секунд."
                )

                self.stop_event.wait(
                    interval
                )

            except Exception as error:
                self.write_log(
                    f"Ошибка проверки почты: {error}"
                )

                self.write_log(
                    "Следующая попытка подключения "
                    "через 60 секунд."
                )

                self.stop_event.wait(
                    60
                )

        self.write_log(
            "Мониторинг остановлен."
        )

    def connect_to_mailbox(self):
        imap_server = self.settings[
            "imap_server"
        ].get().strip()

        imap_port = int(
            self.settings[
                "imap_port"
            ].get()
        )

        mail = None

        try:
            mail = imaplib.IMAP4_SSL(
                imap_server,
                imap_port,
                timeout=30
            )

            mail.login(
                self.settings[
                    "email"
                ].get().strip(),
                self.settings[
                    "password"
                ].get()
            )

            mailbox_name = self.settings[
                "mailbox"
            ].get().strip()

            if not mailbox_name:
                mailbox_name = "INBOX"

            encoded_mailbox = encode_imap_utf7(
                mailbox_name
            )

            status, response = mail.select(
                encoded_mailbox
            )

            if status != "OK":
                raise RuntimeError(
                    f"Не удалось открыть папку почты: "
                    f"{mailbox_name}"
                )

            if (
                self.last_mailbox_message
                != mailbox_name
            ):
                self.write_log(
                    f"Проверяется папка почты: "
                    f"{mailbox_name}"
                )

                self.last_mailbox_message = (
                    mailbox_name
                )

            return mail

        except Exception:
            if mail is not None:
                try:
                    mail.logout()
                except Exception:
                    pass

            raise

    def check_mail(self):
        mail = None

        try:
            mail = self.connect_to_mailbox()

            status, data = mail.uid(
                "search",
                None,
                "UNSEEN"
            )

            if status != "OK":
                self.write_log(
                    "Не удалось получить список писем."
                )
                return

            message_uids = data[0].split()

            if not message_uids:
                return

            self.write_log(
                f"Найдено непрочитанных писем: "
                f"{len(message_uids)}"
            )

            for uid in message_uids:
                if self.stop_event.is_set():
                    break

                self.process_message(
                    mail,
                    uid
                )

        finally:
            if mail is not None:
                try:
                    mail.close()
                except Exception:
                    pass

                try:
                    mail.logout()
                except Exception:
                    pass

    def process_message(self, mail, uid):
        uid_text = uid.decode(
            errors="replace"
        )

        self.write_log(
            f"Начинаю обработку письма UID "
            f"{uid_text}"
        )

        # BODY.PEEK[] скачивает письмо,
        # но не помечает его прочитанным.
        status, data = mail.uid(
            "fetch",
            uid,
            "(BODY.PEEK[])"
        )

        if status != "OK":
            self.write_log(
                "Не удалось получить содержимое письма."
            )
            return

        raw_message = None

        for item in data:
            if isinstance(item, tuple):
                raw_message = item[1]
                break

        if not raw_message:
            self.write_log(
                "Письмо не содержит данных."
            )
            return

        message = BytesParser(
            policy=policy.default
        ).parsebytes(raw_message)

        subject = decode_text(
            message.get(
                "Subject",
                ""
            )
        )

        sender = parseaddr(
            message.get(
                "From",
                ""
            )
        )[1].lower().strip()

        self.write_log(
            f"Получена тема письма: {subject}"
        )

        self.write_log(
            f"Отправитель письма: {sender}"
        )

        required_prefix = self.settings[
            "subject_prefix"
        ].get().strip().casefold()

        actual_subject = subject.strip().casefold()

        if not actual_subject.startswith(
            required_prefix
        ):
            self.write_log(
                "Тема письма не подходит. "
                "Письмо оставлено непрочитанным."
            )
            return

        allowed_senders = [
            address.strip().lower()
            for address in self.settings[
                "allowed_senders"
            ].get().split(",")
            if address.strip()
        ]

        if (
            allowed_senders
            and sender not in allowed_senders
        ):
            self.write_log(
                "Отправитель не разрешен. "
                "Письмо оставлено непрочитанным."
            )
            return

        self.write_log(
            "Тема и отправитель подходят."
        )

        attachment = self.download_attachment(
            message
        )

        if attachment is None:
            self.write_log(
                "В письме нет вложения. "
                "Письмо оставлено непрочитанным."
            )
            return

        file_path, original_filename = attachment

        try:
            self.write_log(
                "Начинаю отправку ответа клиенту."
            )

            self.send_reply(
                recipient=sender,
                subject=subject,
                file_path=file_path,
                original_filename=original_filename
            )

            self.write_log(
                "Ответ клиенту успешно отправлен."
            )

            self.write_log(
                "Начинаю подготовку файла к печати."
            )

            self.print_file(
                file_path
            )

            self.write_log(
                "Файл успешно передан "
                "в очередь печати."
            )

            mark_status, mark_data = mail.uid(
                "store",
                uid,
                "+FLAGS",
                "\\Seen"
            )

            if mark_status != "OK":
                raise RuntimeError(
                    "Файл обработан, но письмо "
                    "не удалось пометить прочитанным."
                )

            self.write_log(
                "Письмо обработано и помечено "
                "как прочитанное."
            )

        except Exception as error:
            self.write_log(
                f"Ошибка обработки письма: {error}"
            )

            self.write_log(
                "Письмо осталось непрочитанным."
            )

        finally:
            try:
                if file_path.exists():
                    file_path.unlink()

                self.write_log(
                    f"Временный файл удален: "
                    f"{original_filename}"
                )

            except Exception as error:
                self.write_log(
                    f"Не удалось удалить временный файл: "
                    f"{error}"
                )

    # =========================================================
    # ВЛОЖЕНИЯ
    # =========================================================

    def download_attachment(self, message):
        folder = Path(
            self.settings[
                "save_folder"
            ].get()
        )

        folder.mkdir(
            parents=True,
            exist_ok=True
        )

        for part in message.walk():
            filename = part.get_filename()

            if not filename:
                continue

            original_filename = safe_filename(
                decode_text(filename)
            )

            temporary_filename = (
                uuid.uuid4().hex
                + "_"
                + original_filename
            )

            file_path = folder / temporary_filename

            content = part.get_payload(
                decode=True
            )

            if content is None:
                continue

            with open(
                file_path,
                "wb"
            ) as file:
                file.write(content)

            self.write_log(
                f"Вложение скачано: "
                f"{original_filename}"
            )

            return (
                file_path,
                original_filename
            )

        return None

    # =========================================================
    # ОТПРАВКА ОТВЕТА
    # =========================================================

    def send_reply(
        self,
        recipient,
        subject,
        file_path,
        original_filename
    ):
        email_address = self.settings[
            "email"
        ].get().strip()

        password = self.settings[
            "password"
        ].get()

        smtp_server = self.settings[
            "smtp_server"
        ].get().strip()

        smtp_port = int(
            self.settings[
                "smtp_port"
            ].get()
        )

        reply = EmailMessage()

        reply["From"] = email_address
        reply["To"] = recipient
        reply["Subject"] = "Re: " + subject

        reply.set_content(
            self.settings[
                "reply_text"
            ].get()
        )

        with open(
            file_path,
            "rb"
        ) as file:
            file_data = file.read()

        extension = Path(
            original_filename
        ).suffix.lower()

        if extension == ".xlsx":
            maintype = "application"
            subtype = (
                "vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )

        elif extension == ".xls":
            maintype = "application"
            subtype = "vnd.ms-excel"

        elif extension == ".xlsm":
            maintype = "application"
            subtype = (
                "vnd.ms-excel.sheet.macroEnabled.12"
            )

        else:
            maintype = "application"
            subtype = "octet-stream"

        reply.add_attachment(
            file_data,
            maintype=maintype,
            subtype=subtype,
            filename=original_filename
        )

        if smtp_port == 465:
            with smtplib.SMTP_SSL(
                smtp_server,
                smtp_port,
                timeout=30
            ) as smtp:
                smtp.login(
                    email_address,
                    password
                )

                smtp.send_message(
                    reply
                )

        else:
            with smtplib.SMTP(
                smtp_server,
                smtp_port,
                timeout=30
            ) as smtp:
                smtp.ehlo()
                smtp.starttls()
                smtp.ehlo()

                smtp.login(
                    email_address,
                    password
                )

                smtp.send_message(
                    reply
                )

    # =========================================================
    # LIBREOFFICE
    # =========================================================

    def find_libreoffice(self):
        if sys.platform.startswith("win"):
            possible_paths = [
                (
                    r"C:\Program Files\LibreOffice"
                    r"\program\soffice.exe"
                ),
                (
                    r"C:\Program Files (x86)"
                    r"\LibreOffice\program\soffice.exe"
                ),
            ]

            for possible_path in possible_paths:
                if os.path.exists(
                    possible_path
                ):
                    return possible_path

            return shutil.which(
                "soffice"
            )

        return (
            shutil.which(
                "libreoffice"
            )
            or shutil.which(
                "soffice"
            )
        )

    def convert_excel_to_pdf(self, excel_path):
        excel_path = Path(
            excel_path
        )

        if not excel_path.exists():
            raise FileNotFoundError(
                f"Excel-файл не найден: "
                f"{excel_path}"
            )

        office_program = self.find_libreoffice()

        if not office_program:
            raise RuntimeError(
                "LibreOffice не найден."
            )

        pdf_path = excel_path.with_suffix(
            ".pdf"
        )

        if pdf_path.exists():
            pdf_path.unlink()

        pdf_filter = (
            "pdf:calc_pdf_Export:"
            '{"SinglePageSheets":'
            '{"type":"boolean","value":"true"}}'
        )

        command = [
            office_program,
            "--headless",
            "--convert-to",
            pdf_filter,
            "--outdir",
            str(excel_path.parent),
            str(excel_path)
        ]

        result = run_command(
            command
        )

        if result.returncode != 0:
            error_text = (
                result.stderr.strip()
                or result.stdout.strip()
                or "Неизвестная ошибка LibreOffice"
            )

            raise RuntimeError(
                "Ошибка преобразования Excel в PDF:\n"
                + error_text
            )

        if not pdf_path.exists():
            raise RuntimeError(
                "LibreOffice не создал PDF-файл."
            )

        self.write_log(
            "Excel преобразован в PDF."
        )

        return pdf_path

    # =========================================================
    # КНИЖНЫЙ PDF
    # =========================================================

    def make_portrait_pdf(self, source_pdf):
        source_pdf = Path(
            source_pdf
        )

        if not source_pdf.exists():
            raise FileNotFoundError(
                f"PDF-файл не найден: "
                f"{source_pdf}"
            )

        target_pdf = source_pdf.with_name(
            source_pdf.stem
            + "_portrait.pdf"
        )

        source_document = fitz.open(
            str(source_pdf)
        )

        result_document = fitz.open()

        # Размер A4 в пунктах PDF.
        page_width = 595
        page_height = 842

        margin_left = 20
        margin_top = 20
        margin_right = 20
        margin_bottom = 20

        available_width = (
            page_width
            - margin_left
            - margin_right
        )

        available_height = (
            page_height
            - margin_top
            - margin_bottom
        )

        for source_page in source_document:
            source_rect = source_page.rect

            scale_x = (
                available_width
                / source_rect.width
            )

            scale_y = (
                available_height
                / source_rect.height
            )

            scale = min(
                scale_x,
                scale_y
            )

            new_width = (
                source_rect.width
                * scale
            )

            new_height = (
                source_rect.height
                * scale
            )

            # Содержимое сверху и слева.
            x = margin_left
            y = margin_top

            target_page = result_document.new_page(
                width=page_width,
                height=page_height
            )

            target_rect = fitz.Rect(
                x,
                y,
                x + new_width,
                y + new_height
            )

            target_page.show_pdf_page(
                target_rect,
                source_document,
                source_page.number
            )

        result_document.save(
            str(target_pdf)
        )

        result_document.close()
        source_document.close()

        self.write_log(
            "Создан книжный PDF формата A4."
        )

        return target_pdf

    # =========================================================
    # ПЕЧАТЬ
    # =========================================================

    def print_file(self, file_path):
        file_path = Path(
            file_path
        )

        if not file_path.exists():
            raise FileNotFoundError(
                f"Файл для печати не найден: "
                f"{file_path}"
            )

        pdf_file = None
        portrait_pdf = None

        try:
            extension = file_path.suffix.lower()

            if extension in [
                ".xls",
                ".xlsx",
                ".xlsm",
                ".ods"
            ]:
                pdf_file = self.convert_excel_to_pdf(
                    file_path
                )

                portrait_pdf = self.make_portrait_pdf(
                    pdf_file
                )

                file_to_print = portrait_pdf

            elif extension == ".pdf":
                portrait_pdf = self.make_portrait_pdf(
                    file_path
                )

                file_to_print = portrait_pdf

            else:
                raise RuntimeError(
                    "Поддерживаются только "
                    "Excel и PDF."
                )

            printer_name = self.settings[
                "printer"
            ].get().strip()

            if sys.platform.startswith("linux"):
                command = [
                    "lp",
                    "-o",
                    "media=A4",
                    "-o",
                    "fit-to-page"
                ]

                if printer_name:
                    command.extend([
                        "-d",
                        printer_name
                    ])

                command.append(
                    str(file_to_print)
                )

                result = run_command(
                    command
                )

                if result.returncode != 0:
                    error_text = (
                        result.stderr.strip()
                        or result.stdout.strip()
                        or "Неизвестная ошибка CUPS"
                    )

                    raise RuntimeError(
                        "Ошибка печати:\n"
                        + error_text
                    )

                self.write_log(
                    "Один книжный PDF отправлен "
                    "в очередь печати."
                )

                if result.stdout.strip():
                    self.write_log(
                        result.stdout.strip()
                    )

            elif sys.platform.startswith("win"):
                os.startfile(
                    str(file_to_print),
                    "print"
                )

                time.sleep(10)

                self.write_log(
                    "Книжный PDF отправлен "
                    "на печать Windows."
                )

            else:
                raise RuntimeError(
                    "Операционная система "
                    "не поддерживается."
                )

        finally:
            if (
                portrait_pdf is not None
                and portrait_pdf.exists()
            ):
                try:
                    portrait_pdf.unlink()

                    self.write_log(
                        "Временный книжный PDF удален."
                    )

                except Exception as error:
                    self.write_log(
                        f"Не удалось удалить "
                        f"книжный PDF: {error}"
                    )

            if (
                pdf_file is not None
                and pdf_file.exists()
            ):
                try:
                    pdf_file.unlink()

                    self.write_log(
                        "Временный PDF удален."
                    )

                except Exception as error:
                    self.write_log(
                        f"Не удалось удалить PDF: "
                        f"{error}"
                    )


if __name__ == "__main__":
    root = tk.Tk()

    application = MailPrinterApp(
        root
    )

    root.mainloop()