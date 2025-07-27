import streamlit as st
import smtplib
from email.message import EmailMessage

def send_reset_email(recipient_email, token, app_url): # CORRECTION : ajout de app_url ici
    """Envoie l'e-mail de réinitialisation."""
    sender_email = st.secrets["SENDER_EMAIL"]
    sender_password = st.secrets["SENDER_PASSWORD"]
    
    reset_link = f"{app_url}?reset_token={token}"

    msg = EmailMessage()
    msg.set_content(f"""Bonjour,

Vous avez demandé une réinitialisation de votre mot de passe pour l'application PerformCheck.

Veuillez cliquer sur le lien ci-dessous pour choisir un nouveau mot de passe. Ce lien expirera dans une heure.

{reset_link}

Si vous n'avez pas demandé cette réinitialisation, vous pouvez ignorer cet e-mail.
""")
    msg['Subject'] = 'Réinitialisation de votre mot de passe - PerformCheck'
    msg['From'] = sender_email
    msg['To'] = recipient_email

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(sender_email, sender_password)
            smtp.send_message(msg)
        return True
    except Exception as e:
        # En cas d'erreur, il est utile de voir le message dans la console du terminal
        print(f"Erreur SMTP : {e}")
        return False