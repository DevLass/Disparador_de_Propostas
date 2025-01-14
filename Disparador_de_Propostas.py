import flet as ft
import requests
from bs4 import BeautifulSoup
import re
import openpyxl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def main(page: ft.Page):
    page.title = "Envio de E-mails Automático"
    page.bgcolor = "#1E1E2E"
    page.padding = 30
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    
    # Ajustando tamanho da janela
    page.window_width = 600  # Largura da janela
    page.window_height = 700  # Altura da janela
    page.window_resizable = False  # Para evitar redimensionamento
    
    titulo = ft.Text("📧 Envio de E-mails", size=28, weight=ft.FontWeight.BOLD, color="white")
    
    tipo_input = ft.TextField(label="Tipo de Estabelecimento", width=400, bgcolor="#282A36", color="white")
    assunto_input = ft.TextField(label="Assunto do E-mail", width=400, bgcolor="#282A36", color="white")
    corpo_input = ft.TextField(label="Corpo do E-mail (HTML)", width=400, multiline=True, min_lines=13, max_lines=18, bgcolor="#282A36", color="white")
    resultado_text = ft.Text(color="white", size=16)
    
    def buscar_urls(query, pages=5):
        all_urls = []
        for page in range(pages):
            start = page * 10
            search_url = f"https://www.google.com/search?q={query}&start={start}"
            headers = {"User-Agent": "Mozilla/5.0"}
            response = requests.get(search_url, headers=headers)
            soup = BeautifulSoup(response.text, 'html.parser')
            links = [a['href'] for a in soup.find_all('a', href=True)]
            urls = [link.split('&')[0].replace('/url?q=', '') for link in links if '/url?q=' in link]
            all_urls.extend(urls)
        return all_urls

    def extrair_emails_e_nome(url):
        estabelecimento_nome = "Nome não encontrado"
        emails = set()
        try:
            response = requests.get(url)
            soup = BeautifulSoup(response.content, 'html.parser')
            title = soup.find('title')
            if title:
                estabelecimento_nome = title.get_text()
            emails = set(re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}', soup.get_text()))
            return {'nome': estabelecimento_nome, 'emails': list(emails)}
        except requests.exceptions.RequestException:
            return {'nome': estabelecimento_nome, 'emails': []}

    def enviar_email(destinatario, assunto, corpo):
        smtp_server = "smtp.gmail.com"
        smtp_port = 587
        email_user = "Agillizaagency@gmail.com"
        email_password = "wgow tnfa enoq ynjn"
        msg = MIMEMultipart()
        msg["From"] = email_user
        msg["To"] = destinatario
        msg["Subject"] = assunto
        msg.attach(MIMEText(corpo, "html"))
        try:
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(email_user, email_password)
            server.sendmail(email_user, destinatario, msg.as_string())
            server.quit()
            return True
        except Exception:
            return False
    
    def processar(event):
        tipo_estabelecimento = tipo_input.value
        assunto = assunto_input.value
        corpo = corpo_input.value
        query = f"{tipo_estabelecimento} email contato site:.br"
        urls = buscar_urls(query, pages=2)
        dados_estabelecimentos = [extrair_emails_e_nome(url) for url in urls]
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Nome do Estabelecimento", "E-mails"])
        emails_enviados = 0
        for dado in dados_estabelecimentos:
            if dado['emails']:
                ws.append([dado['nome'], ", ".join(dado['emails'])])
                for email in dado['emails']:
                    if enviar_email(email, assunto, corpo):
                        emails_enviados += 1
        wb.save("emails.xlsx")
        resultado_text.value = f"✔ Processo concluído. {emails_enviados} e-mails enviados."
        page.update()
    
    botao_enviar = ft.ElevatedButton(
        "🚀 Iniciar Processo", on_click=processar, bgcolor="#FF5733", color="white",
        style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=10))
    )
    
    page.add(
        ft.Column([
            titulo,
            tipo_input,
            assunto_input,
            corpo_input,
            botao_enviar,
            resultado_text
        ], spacing=20, alignment=ft.MainAxisAlignment.CENTER)
    )
    
ft.app(target=main)
