# config/settings.py
# pylint: disable=C0116, C0114

import os
from dotenv import load_dotenv


def load_env() -> dict:
    """
    Carrega as variáveis de ambiente do arquivo .env e retorna um dicionário
    com configurações SMTP e do remetente.
    """
    load_dotenv(override=True)

    smtp_mail_host = os.getenv("SMTP_MAIL_HOST")
    smtp_mail_port = int(os.getenv("SMTP_MAIL_PORT", 587))
    smtp_mail_ssl_port = int(os.getenv("SMTP_MAIL_SSL_PORT", 465))
    smtp_mail_user = os.getenv("SMTP_MAIL_USER")
    smtp_mail_pass = os.getenv("SMTP_MAIL_PASS")
    mail_user_sender = os.getenv("MAIL_USER_SENDER")
    mail_user_name = os.getenv("MAIL_USER_NAME")
    mail_user_pass = os.getenv("MAIL_USER_PASS")

    # Validação mínima
    if not smtp_mail_host or not smtp_mail_user or not smtp_mail_pass:
        raise ValueError("Configuração SMTP incompleta! Verifique seu arquivo .env.")

    return {
        "smtp_mail_host": smtp_mail_host,
        "smtp_mail_port": smtp_mail_port,
        "smtp_mail_ssl_port": smtp_mail_ssl_port,
        "smtp_mail_user": smtp_mail_user,
        "smtp_mail_pass": smtp_mail_pass,
        "mail_user_sender": mail_user_sender,
        "mail_user_name": mail_user_name,
        "mail_user_pass": mail_user_pass,
    }
