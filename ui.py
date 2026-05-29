import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from config import SMTP_SERVICES
from draft_manager import save_draft, load_draft
from email_parser import parse_emails, validate_emails, find_cross_duplicates
from mail_sender import send_bulk_emails
from settings_manager import load_settings, save_settings
from excel_import import load_emails_from_excel


class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Рабочая email-рассылка")
        self.root.geometry("980x760")

        self.attachments = []
        self.settings = load_settings()

        self._build_ui()
        self._load_default_values()

    def _build_ui(self):
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        smtp_frame = ttk.LabelFrame(frame, text="SMTP-настройки", padding=10)
        smtp_frame.pack(fill=tk.X)

        ttk.Label(smtp_frame, text="Сервис:").grid(row=0, column=0, sticky=tk.W)

        self.smtp_service_var = tk.StringVar()
        self.smtp_combo = ttk.Combobox(
            smtp_frame,
            textvariable=self.smtp_service_var,
            values=list(SMTP_SERVICES.keys()),
            state="readonly",
            width=20
        )
        self.smtp_combo.grid(row=0, column=1, padx=5, sticky=tk.W)
        self.smtp_combo.bind("<<ComboboxSelected>>", self._on_smtp_changed)

        ttk.Label(smtp_frame, text="Custom server:").grid(row=0, column=2, sticky=tk.W)

        self.custom_server_var = tk.StringVar()
        self.custom_server_entry = ttk.Entry(
            smtp_frame,
            textvariable=self.custom_server_var,
            width=24
        )
        self.custom_server_entry.grid(row=0, column=3, padx=5)

        ttk.Label(smtp_frame, text="Port:").grid(row=0, column=4, sticky=tk.W)

        self.custom_port_var = tk.StringVar(value="587")
        self.custom_port_entry = ttk.Entry(
            smtp_frame,
            textvariable=self.custom_port_var,
            width=8
        )
        self.custom_port_entry.grid(row=0, column=5, padx=5)

        ttk.Label(smtp_frame, text="Отправитель:").grid(row=1, column=0, sticky=tk.W, pady=5)

        self.sender_var = tk.StringVar()
        ttk.Entry(
            smtp_frame,
            textvariable=self.sender_var,
            width=35
        ).grid(row=1, column=1, padx=5, sticky=tk.W)

        ttk.Label(smtp_frame, text="Пароль приложения:").grid(row=1, column=2, sticky=tk.W, pady=5)

        self.password_var = tk.StringVar()
        ttk.Entry(
            smtp_frame,
            textvariable=self.password_var,
            show="*",
            width=30
        ).grid(row=1, column=3, padx=5, sticky=tk.W)

        ttk.Label(smtp_frame, text="Задержка, сек:").grid(row=1, column=4, sticky=tk.W, pady=5)

        self.delay_var = tk.StringVar(value="1")
        ttk.Entry(
            smtp_frame,
            textvariable=self.delay_var,
            width=8
        ).grid(row=1, column=5, padx=5)

        recipients_frame = ttk.LabelFrame(frame, text="Получатели", padding=10)
        recipients_frame.pack(fill=tk.BOTH, expand=True, pady=8)

        ttk.Label(recipients_frame, text="Кому:").grid(row=0, column=0, sticky=tk.NW)

        self.to_text = tk.Text(recipients_frame, height=5)
        self.to_text.bind("<KeyRelease>", self._update_recipients_count)
        self.to_text.bind("<FocusOut>", self._update_recipients_count)
        self.to_text.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=4)

        self.recipients_count_var = tk.StringVar(value="Получателей: 0")

        ttk.Label(
            recipients_frame,
            textvariable=self.recipients_count_var
        ).grid(row=0, column=2, sticky=tk.NW, padx=5)

        ttk.Label(recipients_frame, text="Копия CC:").grid(row=1, column=0, sticky=tk.NW)

        self.cc_text = tk.Text(recipients_frame, height=3)
        self.cc_text.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=4)

        self.save_cc_var = tk.BooleanVar()
        ttk.Checkbutton(
            recipients_frame,
            text="Сохранить CC по умолчанию",
            variable=self.save_cc_var
        ).grid(row=1, column=2, sticky=tk.W)

        ttk.Label(recipients_frame, text="Скрытая копия BCC:").grid(row=2, column=0, sticky=tk.NW)

        self.bcc_text = tk.Text(recipients_frame, height=3)
        self.bcc_text.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=4)

        self.save_bcc_var = tk.BooleanVar()
        ttk.Checkbutton(
            recipients_frame,
            text="Сохранить BCC по умолчанию",
            variable=self.save_bcc_var
        ).grid(row=2, column=2, sticky=tk.W)

        recipients_frame.columnconfigure(1, weight=1)

        message_frame = ttk.LabelFrame(frame, text="Письмо", padding=10)
        message_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(message_frame, text="Тема:").grid(row=0, column=0, sticky=tk.W)

        self.subject_var = tk.StringVar()
        ttk.Entry(
            message_frame,
            textvariable=self.subject_var
        ).grid(row=0, column=1, sticky=tk.EW, padx=5, pady=4)

        ttk.Label(message_frame, text="Текст:").grid(row=1, column=0, sticky=tk.NW)

        self.body_text = tk.Text(message_frame, height=8)
        self.body_text.grid(row=1, column=1, sticky=tk.NSEW, padx=5, pady=4)

        message_frame.columnconfigure(1, weight=1)
        message_frame.rowconfigure(1, weight=1)

        attach_frame = ttk.LabelFrame(frame, text="Вложения", padding=10)
        attach_frame.pack(fill=tk.X, pady=8)

        self.attachments_listbox = tk.Listbox(attach_frame, height=4)
        self.attachments_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)

        attach_buttons = ttk.Frame(attach_frame)
        attach_buttons.pack(side=tk.RIGHT, padx=5)

        ttk.Button(
            attach_buttons,
            text="Добавить файл",
            command=self._add_attachment
        ).pack(fill=tk.X, pady=2)

        ttk.Button(
            attach_buttons,
            text="Удалить файл",
            command=self._remove_attachment
        ).pack(fill=tk.X, pady=2)

        buttons_frame = ttk.Frame(frame)
        buttons_frame.pack(fill=tk.X, pady=5)

        ttk.Button(
            buttons_frame,
            text="Проверить адреса",
            command=self._check_emails
        ).pack(side=tk.LEFT, padx=3)

        ttk.Button(
            buttons_frame,
            text="Импорт Excel",
            command=self._import_excel
        ).pack(side=tk.LEFT, padx=3)

        ttk.Button(
            buttons_frame,
            text="Предпросмотр",
            command=self._preview_email
        ).pack(side=tk.LEFT, padx=3)

        ttk.Button(
            buttons_frame,
            text="Сохранить рассылку",
            command=self._save_draft
        ).pack(side=tk.LEFT, padx=3)

        ttk.Button(
            buttons_frame,
            text="Загрузить рассылку",
            command=self._load_draft_dialog
        ).pack(side=tk.LEFT, padx=3)

        ttk.Button(
            buttons_frame,
            text="Отправить рассылку",
            command=self._send_emails
        ).pack(side=tk.RIGHT, padx=3)

        progress_frame = ttk.Frame(frame)
        progress_frame.pack(fill=tk.X, pady=5)

        self.progress_var = tk.DoubleVar()

        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100
        )

        self.progress_bar.pack(fill=tk.X)

        self.progress_label_var = tk.StringVar(
            value="Готов к отправке"
        )

        ttk.Label(
            progress_frame,
            textvariable=self.progress_label_var
        ).pack(anchor="w")

        log_frame = ttk.LabelFrame(frame, text="Журнал выполнения", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(log_frame, height=8)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def _load_default_values(self):
        self.smtp_service_var.set(self.settings.get("last_smtp_service", "Yandex"))
        self.sender_var.set(self.settings.get("last_sender_email", ""))
        self.delay_var.set(str(self.settings.get("delay_seconds", 1)))

        self.cc_text.insert("1.0", self.settings.get("default_cc", ""))
        self.bcc_text.insert("1.0", self.settings.get("default_bcc", ""))

        self._on_smtp_changed()

    def _on_smtp_changed(self, event=None):
        is_custom = self.smtp_service_var.get() == "Custom SMTP"
        state = "normal" if is_custom else "disabled"

        self.custom_server_entry.configure(state=state)
        self.custom_port_entry.configure(state=state)

    def _add_attachment(self):
        paths = filedialog.askopenfilenames(title="Выберите файлы для вложения")

        for path in paths:
            if path not in self.attachments:
                self.attachments.append(path)
                self.attachments_listbox.insert(tk.END, path)

    def _remove_attachment(self):
        selection = self.attachments_listbox.curselection()

        if not selection:
            return

        index = selection[0]
        self.attachments_listbox.delete(index)
        del self.attachments[index]

    def _get_form_data(self):
        to_list = parse_emails(self.to_text.get("1.0", tk.END))
        cc_list = parse_emails(self.cc_text.get("1.0", tk.END))
        bcc_list = parse_emails(self.bcc_text.get("1.0", tk.END))

        return {
            "smtp_service": self.smtp_service_var.get(),
            "custom_server": self.custom_server_var.get().strip(),
            "custom_port": self.custom_port_var.get().strip(),
            "sender_email": self.sender_var.get().strip(),
            "to": to_list,
            "cc": cc_list,
            "bcc": bcc_list,
            "subject": self.subject_var.get().strip(),
            "body": self.body_text.get("1.0", tk.END).strip(),
            "attachments": self.attachments,
            "delay_seconds": self.delay_var.get().strip()
        }

    def _import_excel(self):

        file_path = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm")
            ]
        )

        if not file_path:
            return

        try:
            emails = load_emails_from_excel(file_path)

            self.to_text.delete("1.0", tk.END)

            self.to_text.insert(
                "1.0",
                "\n".join(emails)
            )

            self._update_recipients_count()

            messagebox.showinfo(
                "Импорт",
                f"Загружено {len(emails)} адресов"
            )

        except Exception as exc:
            messagebox.showerror(
                "Ошибка импорта",
                str(exc)
            )

    def _update_recipients_count(self, event=None):
        to_list = parse_emails(self.to_text.get("1.0", tk.END))
        self.recipients_count_var.set(f"Получателей: {len(to_list)}")
    def _check_emails(self):
        data = self._get_form_data()

        all_invalid = []

        for field_name in ["to", "cc", "bcc"]:
            _, invalid = validate_emails(data[field_name])
            all_invalid.extend(invalid)

        cross_duplicates = find_cross_duplicates(
            data["to"],
            data["cc"],
            data["bcc"]
        )

        if all_invalid:
            messagebox.showerror(
                "Ошибка",
                "Некорректные email:\n" + "\n".join(all_invalid)
            )
            return

        if cross_duplicates:
            messagebox.showwarning(
                "Предупреждение",
                "Адреса встречаются в нескольких полях:\n" + "\n".join(cross_duplicates)
            )
            return

        messagebox.showinfo(
            "Проверка",
            f"Адреса корректны.\nКому: {len(data['to'])}\nCC: {len(data['cc'])}\nBCC: {len(data['bcc'])}"
        )

    def _save_current_settings(self):
        if self.save_cc_var.get():
            self.settings["default_cc"] = self.cc_text.get("1.0", tk.END).strip()

        if self.save_bcc_var.get():
            self.settings["default_bcc"] = self.bcc_text.get("1.0", tk.END).strip()

        self.settings["last_sender_email"] = self.sender_var.get().strip()
        self.settings["last_smtp_service"] = self.smtp_service_var.get()

        try:
            self.settings["delay_seconds"] = int(self.delay_var.get() or 1)
        except ValueError:
            self.settings["delay_seconds"] = 1

        save_settings(self.settings)

    def _save_draft(self):
        data = self._get_form_data()

        file_path = filedialog.asksaveasfilename(
            title="Сохранить рассылку",
            initialdir="drafts",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")]
        )

        if not file_path:
            return

        path = save_draft(data, file_path)

        self._save_current_settings()

        messagebox.showinfo(
            "Черновик",
            f"Рассылка сохранена:\n{path}"
        )

    def _load_draft_dialog(self):
        path = filedialog.askopenfilename(
            title="Выберите черновик",
            initialdir="drafts",
            filetypes=[("JSON files", "*.json")]
        )

        if not path:
            return

        data = load_draft(path)
        self._apply_draft(data)

        messagebox.showinfo("Черновик", "Рассылка загружена.")

    def _apply_draft(self, data):
        self.smtp_service_var.set(data.get("smtp_service", "Yandex"))
        self.custom_server_var.set(data.get("custom_server", ""))
        self.custom_port_var.set(str(data.get("custom_port", "587")))
        self.sender_var.set(data.get("sender_email", ""))
        self.subject_var.set(data.get("subject", ""))
        self.delay_var.set(str(data.get("delay_seconds", 1)))

        self.to_text.delete("1.0", tk.END)
        self.to_text.insert("1.0", ", ".join(data.get("to", [])))

        self.cc_text.delete("1.0", tk.END)
        self.cc_text.insert("1.0", ", ".join(data.get("cc", [])))

        self.bcc_text.delete("1.0", tk.END)
        self.bcc_text.insert("1.0", ", ".join(data.get("bcc", [])))

        self.body_text.delete("1.0", tk.END)
        self.body_text.insert("1.0", data.get("body", ""))

        self.attachments = data.get("attachments", [])
        self.attachments_listbox.delete(0, tk.END)

        for path in self.attachments:
            self.attachments_listbox.insert(tk.END, path)

        self._update_recipients_count()
        self._on_smtp_changed()

    def _preview_email(self):
        data = self._get_form_data()

        preview_window = tk.Toplevel(self.root)
        preview_window.title("Предпросмотр письма")
        preview_window.geometry("700x600")

        text = tk.Text(preview_window, wrap=tk.WORD)
        text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        attachments_text = "\n".join(data["attachments"]) if data["attachments"] else "Нет вложений"

        preview_content = (
            f"SMTP-сервис: {data['smtp_service']}\n"
            f"Отправитель: {data['sender_email']}\n"
            f"Количество получателей: {len(data['to'])}\n\n"
            f"Кому будет отправлено отдельно:\n"
            f"{chr(10).join(data['to'])}\n\n"
            f"Копия CC:\n"
            f"{chr(10).join(data['cc']) if data['cc'] else 'Нет'}\n\n"
            f"Скрытая копия BCC:\n"
            f"{chr(10).join(data['bcc']) if data['bcc'] else 'Нет'}\n\n"
            f"Тема:\n"
            f"{data['subject']}\n\n"
            f"Текст письма:\n"
            f"{data['body']}\n\n"
            f"Вложения:\n"
            f"{attachments_text}"
        )

        text.insert("1.0", preview_content)
        text.config(state=tk.DISABLED)

    def _send_emails(self):
        data = self._get_form_data()

        if not data["sender_email"]:
            messagebox.showerror("Ошибка", "Укажите отправителя.")
            return

        if not self.password_var.get():
            messagebox.showerror("Ошибка", "Укажите пароль приложения.")
            return

        if not data["to"]:
            messagebox.showerror("Ошибка", "Укажите хотя бы одного получателя.")
            return

        if not data["subject"]:
            messagebox.showerror("Ошибка", "Укажите тему письма.")
            return

        if not data["body"]:
            messagebox.showerror("Ошибка", "Укажите текст письма.")
            return

        all_invalid = []

        for field_name in ["to", "cc", "bcc"]:
            _, invalid = validate_emails(data[field_name])
            all_invalid.extend(invalid)

        if all_invalid:
            messagebox.showerror(
                "Ошибка",
                "Некорректные email:\n" + "\n".join(all_invalid)
            )
            return

        confirm = messagebox.askyesno(
            "Подтверждение",
            f"Отправить {len(data['to'])} отдельных писем?"
        )

        if not confirm:
            return

        self._save_current_settings()

        try:
            delay = int(data["delay_seconds"] or 1)
        except ValueError:
            delay = 1

        self._log("Начата отправка...")

        try:
            send_bulk_emails(
                smtp_service=data["smtp_service"],
                custom_server=data["custom_server"],
                custom_port=int(data["custom_port"] or 587),
                sender_email=data["sender_email"],
                password=self.password_var.get(),
                to_list=data["to"],
                cc_list=data["cc"],
                bcc_list=data["bcc"],
                subject=data["subject"],
                body=data["body"],
                attachment_paths=data["attachments"],
                delay_seconds=delay,
                progress_callback=self._update_progress,
                log_callback=self._log
            )

            messagebox.showinfo(
                "Готово",
                "Рассылка завершена. История сохранена в Excel."
            )

        except Exception as exc:
            messagebox.showerror("Ошибка отправки", str(exc))
            self._log(f"Критическая ошибка: {exc}")

    def _update_progress(self, current, total):

        percent = (current / total) * 100

        self.progress_var.set(percent)

        self.progress_label_var.set(
            f"Отправлено {current} из {total}"
        )

        self.root.update_idletasks()

    def _log(self, text):
        self.log_text.insert(tk.END, text + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()