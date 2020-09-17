import smtplib, ssl, pyodbc
from email.message import EmailMessage
from datetime import datetime
import schedule
import time

# Função que converte o array 2d recebido numa estrutura HTML
def cursor2html(list2d):
    htable=u'<table border="1" bordercolor=000000 cellspacing="0" cellpadding="1" style="table-layout:fixed;vertical-align:bottom;font-size:13px;font-family:verdana,sans,sans-serif;border-collapse:collapse;border:1px solid rgb(130,130,130)" >'
    list2d[0] = [u'<b>' + i + u'</b>' for i in list2d[0]]
    for row in list2d:
        newrow = u'<tr>' 
        newrow += u'<td align="left" style="padding:1px 4px">'+ str(row[0]) + '</td>'
        row.remove(row[0])
        newrow = newrow + ''.join([u'<td align="left" style="padding:1px 4px">' + str(x) + u'</td>' for x in row])  
        newrow += '</tr>' 
        htable += newrow
    htable += '</table>'
    return htable

def sendit():
    # Ligação à base de dados e execução de uma query
    conn = pyodbc.connect('Driver={SQL Server};' 
                          'Server=DESKTOP-ILGEKKB\SQLEXPRESS;'
                          'Database=redinsp;'
                          'Trusted_Connection=yes;')

    cursor = conn.cursor()
    sql_query = "SELECT no AS [N°], nome AS [NOME], SUM(etotal) AS [Valor] FROM ft (NOLOCK) WHERE ft.ndoc = 1 AND pdata <= GETDATE() GROUP BY no, nome ORDER BY [Valor] DESC"
    cursor.execute(sql_query)

    # Converte o cursor num array 2d
    header = [i[0] for i in cursor.description]
    rows = [list(i) for i in cursor.fetchall()]
    rows.insert(0,header)
    cursor.close()
    conn.close()

    # Conversão da array em html
    mail_table = cursor2html(rows)

    # Dados ligação SMTP
    smtp_server = "smtp.gmail.com"
    port = 587
    password = "bifinhoscomchampignons"
    sender_email = "tsobral.i9biz@gmail.com"
    dest_email = "tiago.sobral@i9.business"
    msg_email = str(mail_table)
    context = ssl.create_default_context()

    # Dados do Email a ser enviado
    msg = EmailMessage() 
    msg['Subject'] = "Extrato dívida de clientes. Dia " + str(datetime.strftime(datetime.now(), '%d-%m-%Y'))
    msg['From'] = sender_email
    msg['To'] = dest_email
    msg.set_type('text/html')
    msg_email = "<b>Valores vencidos por cliente</b>" + msg_email
    msg.set_content(msg_email, subtype="html")
    server = smtplib.SMTP(smtp_server,port)

    # Envio do email
    try:    
        server.ehlo() # Can be omitted
        server.starttls(context=context) # Secure the connection
        server.ehlo() # Can be omitted
        server.login(sender_email, password)
        #server.sendmail(sender_email, dest_email, msg_email)
        server.send_message(msg)
        # Sucesso
        msg_ok = "Enviado para " + dest_email + " às " + str(datetime.now())
        print(msg_ok)
        # Extra: Log de envios
        log_file = open('log_alertas.txt', 'a')
        log_file.write("[" + str(datetime.now()) + "] OK: " + "Enviado para " + dest_email + " (" + str(msg['Subject']) + ")\n")
        log_file.close()
    except Exception as e:
        log_file = open('log_alertas.txt', 'a')
        log_file.write("[" + str(datetime.now()) + "] ERRO!: " + e + "\n")
        log_file.close()
        print(e)
    finally:
        server.quit()

# Frequência de envio do email
schedule.every(10).minutes.do(sendit)

while 1:
    schedule.run_pending()
    time.sleep(60)