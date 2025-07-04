from app.email_sender import send_email

if __name__ == "__main__":
    destinatario = "harryrobert1997@gmail.com"  # Coloque seu email real para teste
    nome_destinatario = "Fulano"
    tipo_vaga = "Advogado"
    unidade = "São Paulo"

    send_email(destinatario, nome_destinatario, tipo_vaga, unidade)
