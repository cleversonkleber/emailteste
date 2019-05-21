import smtplib
from email.mime.text import MIMEText

remetente = 'cleverson.cordeiro.cc1@vcimentos.com'
# Informações da mensagem
destinatario = 'cleversonkleber@gmail.com'
assunto      = 'Enviando email com python'
texto        = 'Esse email foi enviado usando python! :)'

# self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI").Folders(mailbox)
# self.inbox = self.outlook.Folders('Caixa de Entrada').Folders('a00-Interaxa')
# messages = self.inbox.Items
# message = messages.GetFirst()


# Preparando a mensagem
msg = '\r\n'.join([
  'From: %s' % remetente,
  'To: %s' % destinatario,
  'Subject: %s' % assunto,
  '',
  '%s' % texto
  ])



smtpobj = smtplib.SMTP('smtp.office365.com')
#smtpobj.ehlo()
smtpobj.starttls()
smtpobj.login('email@email.com','senha')

smtpobj.sendmail(remetente,destinatario,msg)
smtpobj.quit()
