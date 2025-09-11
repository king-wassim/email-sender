import os
import csv
import smtplib
import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from jinja2 import Environment, FileSystemLoader

def load_config():
    """Charge la configuration depuis le fichier config.json"""
    with open('data/config.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def load_companies():
    companies = []
    with open('data/companies.csv', 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            companies.append(row)
    return companies

def render_email_template(company_info):
    env = Environment(loader=FileSystemLoader('templates'))
    template = env.get_template('email_template.j2')
    return template.render(company_info)

def send_email_brevo(config, company, email_body):
    # Création du message
    msg = MIMEMultipart()
    msg['Subject'] = config['subject']
    msg['From'] = f"{config['sender_name']} <{config['sender_email']}>"
    msg['To'] = company['contact_email']
    
    # Ajout du corps du message
    msg.attach(MIMEText(email_body, 'plain', 'utf-8'))
    
    # Ajout de la pièce jointe (CV)
    with open('attachments/cv.pdf', 'rb') as f:
        attach = MIMEApplication(f.read(), _subtype='pdf')
        attach.add_header('Content-Disposition', 'attachment', filename='cv.pdf')
        msg.attach(attach)
    
    # Envoi de l'email via Brevo SMTP
    try:
        with smtplib.SMTP(config['smtp_server'], config['smtp_port']) as server:
            server.starttls()
            server.login(config['login'], config['password'])
            server.send_message(msg)
        
        print(f"✓ Email envoyé à {company['company_name']}")
        return True
        
    except Exception as e:
        print(f"✗ Erreur avec {company['company_name']}: {str(e)}")
        return False

def main():
    config = load_config()
    companies = load_companies()
    
    print(f"Début de l'envoi des emails à {len(companies)} entreprises...")
    
    success_count = 0
    for company in companies:
        try:
            email_body = render_email_template(company)
            if send_email_brevo(config, company, email_body):
                success_count += 1
        except Exception as e:
            print(f"✗ Échec pour {company['company_name']}: {str(e)}")
    
    print(f"Envoi terminé. {success_count}/{len(companies)} emails envoyés avec succès.")

if __name__ == "__main__":
    main()