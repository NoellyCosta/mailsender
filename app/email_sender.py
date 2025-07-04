from email.header import Header
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.utils import formataddr
from config.settings import load_env
from app.email_server import save_email_inBox
import os
import sys
from PIL import Image
import io


def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
    return os.path.join(base_path, relative_path)


def resize_image_to_bytes(image_path: str, max_width: int, max_height: int) -> bytes:
    """
    Redimensiona a imagem mantendo proporção, limitando largura e altura máximas.
    Retorna os bytes da imagem em formato PNG.
    """
    img = Image.open(image_path)
    img.thumbnail((max_width, max_height), Image.LANCZOS)

    img_bytes = io.BytesIO()
    img.save(img_bytes, format="PNG")
    return img_bytes.getvalue()


def send_email(destinatario: str, nome_destinatario: str, tipo_vaga: str, unidade: str,
               tipo_email: str = "primeiro_contato",
               logo_path: str = None, header_path: str = None) -> bool:
    data_env = load_env()

    msg = MIMEMultipart('related')
    sender_name = str(Header(data_env['mail_user_name'], "utf-8"))
    msg['From'] = formataddr((sender_name, data_env['mail_user_sender']))
    msg['To'] = destinatario

    if tipo_email == "primeiro_contato":
        msg['Subject'] = f"Confirmação de Recebimento – Processo Seletivo {tipo_vaga}"
        body_message = set_format_message_primeiro_contato(nome_destinatario, tipo_vaga, unidade)
    elif tipo_email == "negativa":
        msg['Subject'] = f"Retorno do Processo Seletivo – {tipo_vaga}"
        body_message = set_format_message_negativa(nome_destinatario, tipo_vaga, unidade)
    else:
        print(f"Tipo de email desconhecido: {tipo_email}")
        return False

    msg_alternative = MIMEMultipart('alternative')
    msg.attach(msg_alternative)
    msg_alternative.attach(MIMEText(body_message, 'html', 'utf-8'))

    if logo_path is None:
        logo_path = resource_path('src/assets/assinaturaRodolfo.png')
    if header_path is None:
        header_path = resource_path('src/assets/obrigado.jpeg')

    if not os.path.isfile(logo_path):
        print(f"Logo não encontrado: {logo_path}")
    if not os.path.isfile(header_path):
        print(f"Header não encontrado: {header_path}")

    # Header com altura aumentada para 200 px
    try:
        header_bytes = resize_image_to_bytes(header_path, 600, 200)
        image_header = MIMEImage(header_bytes, 'png')
        image_header.add_header('Content-ID', '<headerimg>')
        msg.attach(image_header)
    except Exception as e:
        print(f"Erro ao anexar header: {e}")

    # Assinatura com altura aumentada para 160 px e largura 550 px
    try:
        logo_bytes = resize_image_to_bytes(logo_path, 550, 160)
        image_logo = MIMEImage(logo_bytes, 'png')
        image_logo.add_header('Content-ID', '<logo>')
        msg.attach(image_logo)
    except Exception as e:
        print(f"Erro ao anexar logo: {e}")

    try:
        with smtplib.SMTP_SSL(
            host=data_env['smtp_mail_host'],
            port=int(data_env['smtp_mail_port'])
        ) as server:
            server.login(
                data_env['mail_user_sender'],
                data_env['mail_user_pass'],
            )
            print("Conectado ao servidor de e-mail!")
            server.send_message(msg)
            print("E-mail enviado com sucesso!")

        save_email_inBox(data_env['mail_user_sender'], data_env['mail_user_pass'], msg.as_string())
        return True
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")
        return False


def set_format_message_primeiro_contato(nome_destinatario: str, tipo_vaga: str, unidade: str) -> str:
    return f"""
    <html>
        <body style="margin:0; padding:0; background-color:#f4f4f4; font-family: Arial, sans-serif; color: #333;">
            <div style="max-width: 700px; margin: 20px auto; background-color: #fff; padding: 30px 40px; box-sizing: border-box; border-radius: 6px; line-height: 1.5;">
                <div style="text-align: center; margin-bottom: 30px;">
                    <img src="cid:headerimg" alt="Header" style="width: 100%; max-width: 600px; max-height: 200px; height: auto; display: block; margin: 0 auto;" />
                </div>
                <p>Prezado(a) <span style="font-weight: bold;">{nome_destinatario}</span>,</p>
                <p>Agradecemos pelo seu interesse em fazer parte do nosso time na Walmir Valença Advocacia!</p>
                <p>Confirmamos o recebimento da sua inscrição para a vaga de <span style="font-weight: bold;">{tipo_vaga}</span> na unidade <span style="font-weight: bold;">{unidade}</span>. Estamos analisando todos os currículos com atenção e, em breve, retornaremos com atualizações.</p>
                <p>Acompanhe seu e-mail e verifique a caixa de spam para garantir que nossas comunicações não sejam perdidas.</p>
                <p>Agradecemos pela paciência e entusiasmo em integrar nossa equipe!</p>
                <div style="text-align: center; margin-top: 40px;">
                    <img src="cid:logo" alt="Assinatura" style="width: 100%; max-width: 550px; max-height: 160px; height: auto; display: block; margin: 0 auto;" />
                </div>
            </div>
        </body>
    </html>
    """


def set_format_message_negativa(nome_destinatario: str, tipo_vaga: str, unidade: str) -> str:
    return f"""
    <html>
        <body style="margin:0; padding:0; background-color:#f4f4f4; font-family: Arial, sans-serif; color: #333;">
            <div style="max-width: 700px; margin: 20px auto; background-color: #fff; padding: 30px 40px; box-sizing: border-box; border-radius: 6px; line-height: 1.5;">
                <div style="text-align: center; margin-bottom: 30px;">
                    <img src="cid:headerimg" alt="Header" style="width: 100%; max-width: 600px; max-height: 200px; height: auto; display: block; margin: 0 auto;" />
                </div>
                <p>Olá <span style="font-weight: bold;">{nome_destinatario}</span>,</p>
                <p>Primeiramente, pedimos desculpas pelo atraso no retorno.</p>
                <p>Agradecemos sinceramente o seu interesse em fazer parte da nossa equipe e por ter participado do processo seletivo para a vaga de <span style="font-weight: bold;">{tipo_vaga}</span> em <span style="font-weight: bold;">{unidade}</span>. Após uma análise cuidadosa do seu currículo, informamos que optamos por seguir com outro candidato para esta oportunidade.</p>
                <p>Apesar disso, ficamos muito impressionados com suas qualificações e gostaríamos de manter seu currículo em nosso banco de dados para futuras oportunidades.</p>
                <p>Agradecemos mais uma vez pelo seu interesse e esperamos que possamos nos conectar no futuro.</p>
                <p>Desejamos muito sucesso em sua trajetória profissional!</p>
                <div style="text-align: center; margin-top: 40px;">
                    <img src="cid:logo" alt="Assinatura" style="width: 100%; max-width: 550px; max-height: 160px; height: auto; display: block; margin: 0 auto;" />
                </div>
            </div>
        </body>
    </html>
    """
