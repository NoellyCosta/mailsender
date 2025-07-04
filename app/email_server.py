# app/email_server.py
# pylint: disable=C0116, C0114, C0103, W0718

import imaplib
import time
from config.settings import load_env


def save_email_inBox(usuario: str, senha: str, msg):
    """
    Salva uma cópia do e-mail enviado na pasta 'Enviadas' (ou equivalente)
    usando IMAP.
    """
    data_env = load_env()

    try:
        with imaplib.IMAP4_SSL(data_env['smtp_mail_host']) as imap:
            imap.login(usuario, senha)
            imap.append('INBOX.enviadas', '\\Seen', imaplib.Time2Internaldate(time.time()), msg.as_bytes())
            print("E-mail salvo na caixa de enviados com sucesso!")
            imap.logout()
    except imaplib.IMAP4.error as e:
        print(f"Erro de IMAP ao salvar na caixa de enviados: {e}")
    except Exception as e:
        print(f"Erro inesperado ao salvar na caixa de enviados: {e}")
