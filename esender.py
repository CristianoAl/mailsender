import pandas as pd;
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

recipients = pd.read_excel('./recipients.xlsx')



for  index, item in recipients.iterrows():
    msn = MIMEMultipart()
    msn['Subject'] = item['assunto']
    message_html = f'''
    <html>
        <head></head>
        <body>
            <p>Prezado {item['nome']},
            você recebeu um email de servidor x!</p><br>
            texto texto texto<br>
            texto texto texto texto texto texto texto texto texto texto texto texto
            </p>
            <hr size="1" />
            <table>
            <tr>
                <td style="font-family:tahoma,arial,verdana; font-size:11px; text-align:center" valign="top">
                <a href="" target="_blank"><img src="https://scontent.fbsb9-1.fna.fbcdn.net/v/t39.30808-6/344574028_5786835281425621_4262162839943568702_n.jpg?_nc_cat=106&ccb=1-7&_nc_sid=6ee11a&_nc_ohc=F9ouqEYmtWMQ7kNvgFtzbbb&_nc_zt=23&_nc_ht=scontent.fbsb9-1.fna&_nc_gid=AOsCeUPQlkBwLpIBGLExgnm&oh=00_AYAF0NnQAUR7j-C2jztDXZ3Xy9K29EBkkjMn8piK9u4UJg&oe=6731AD2B" width="190px" height="150px" border="0" /></a>
                </td>
                <td style="font-family:tahoma,arial,verdana; font-size:12px; padding-left:10px">
                <strong>Cristiano Alves</strong>
                <br />
                <strong>Fone:</strong> (31) 3409-7110 | (31) 3409-7325
                <br />
                <strong>Site:</strong> <a href="https://ipead.face.ufmg.br/novo_site" target="_blank">https://ipead.face.ufmg.br/</a>
                <br />
                <strong>Instagram: </strong>@ipeadufmg<br />
                Fundação Instituto de Pesquisas Econômicas, Administrativas e Contábeis de Minas Gerais
                <br /></td> </tr></table>
        </body>
    </html>'''
    msn['From'] = 'info@ipead.org.br'
    msn['To'] = item['email']
    msn.attach(MIMEText(message_html, 'html'))

    server = smtplib.SMTP('smtp.gmail.com', port=587)
    server.starttls()
    server.login('info@ipead.org.br', 'qeic yjfb ivan hffo')
    server.sendmail(msn['From'], msn['To'], msn.as_string())
    server.quit()

