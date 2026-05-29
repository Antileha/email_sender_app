import os
import smtplib
import time
from email.message import EmailMessage
from email.headerregistry import Address
from email.header import Header

from config import SMTP_SERVICES
from history import write_history


def build_message(
    sender_email,
    recipient,
    cc_list,
    bcc_list,
    subject,
    body,
    attachment_paths
):
    msg = EmailMessage()

    msg["From"] = str(Header(sender_email, "utf-8"))
    msg["To"] = str(Header(recipient, "utf-8"))

    if cc_list:
        msg["Cc"] = ", ".join(cc_list)

    # BCC не добавляем в заголовки письма.
    # Эти адреса передаются только в SMTP-список получателей.
    msg["Subject"] = str(Header(subject, "utf-8"))
    msg.set_content(body)

    for path in attachment_paths:
        if not path or not os.path.exists(path):
            continue

        with open(path, "rb") as file:
            data = file.read()

        filename = os.path.basename(path)

        msg.add_attachment(
            data,
            maintype="application",
            subtype="octet-stream",
            filename=filename
        )

    return msg


def send_bulk_emails(
    smtp_service,
    custom_server,
    custom_port,
    sender_email,
    password,
    to_list,
    cc_list,
    bcc_list,
    subject,
    body,
    attachment_paths,
    delay_seconds=1,
    progress_callback=None,
    log_callback=None
):
    if smtp_service == "Custom SMTP":
        smtp_server = custom_server
        smtp_port = custom_port
        use_tls = True
    else:
        service_cfg = SMTP_SERVICES[smtp_service]
        smtp_server = service_cfg["server"]
        smtp_port = service_cfg["port"]
        use_tls = service_cfg["use_tls"]

    attachment_names = ", ".join(os.path.basename(x) for x in attachment_paths)

    success_count = 0
    error_count = 0
    failed_recipients = []

    with smtplib.SMTP(smtp_server, smtp_port, timeout=30) as server:
        if use_tls:
            server.starttls()

        server.login(sender_email, password)

        total = len(to_list)

        for index, recipient in enumerate(to_list, start=1):
            try:
                msg = build_message(
                    sender_email=sender_email,
                    recipient=recipient,
                    cc_list=cc_list,
                    bcc_list=bcc_list,
                    subject=subject,
                    body=body,
                    attachment_paths=attachment_paths
                )

                all_recipients = [recipient] + cc_list + bcc_list

                server.send_message(
                    msg,
                    from_addr=sender_email,
                    to_addrs=all_recipients
                )

                write_history(
                    smtp_service=smtp_service,
                    sender=sender_email,
                    recipient=recipient,
                    cc=", ".join(cc_list),
                    bcc=", ".join(bcc_list),
                    subject=subject,
                    attachments=attachment_names,
                    status="Успешно",
                    error=""
                )

                success_count += 1

                if log_callback:
                    log_callback(f"Успешно отправлено: {recipient}")

                if progress_callback:
                    progress_callback(index, total)

            except Exception as exc:
                write_history(
                    smtp_service=smtp_service,
                    sender=sender_email,
                    recipient=recipient,
                    cc=", ".join(cc_list),
                    bcc=", ".join(bcc_list),
                    subject=subject,
                    attachments=attachment_names,
                    status="Ошибка",
                    error=str(exc)
                )

                error_count += 1
                failed_recipients.append(recipient)

                if log_callback:
                    log_callback(f"Ошибка отправки {recipient}: {exc}")

            time.sleep(max(0, delay_seconds))

    return {
                "success_count": success_count,
                "error_count": error_count,
                "failed_recipients": failed_recipients
            }