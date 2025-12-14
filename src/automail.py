from tkinter.messagebox import *
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import email, smtplib, ssl
import openpyxl as xl
from pathlib import Path
from configparser import ConfigParser
import dialogs


class EmailSender:
    def reload_sender(self):
        parser = ConfigParser()
        parser.read('./config/intern.ini')
        self.sender_email = parser.get('mail', 'email')
        self.password_email = parser.get('mail', 'password')
        self.server = 'smtp.' + self.sender_email.split('@')[1]

    def send(self, to, subject, text, files = []):
        self.reload_sender()

        try:
            # Create a multipart message and set headers
            message = MIMEMultipart()
            message["From"] = self.sender_email
            message["To"] = to
            message["Subject"] = subject
            #message["Bcc"] = to  # Recommended for mass emails

            # Turn these into plain/html MIMEText objects
            part = MIMEText(text, "plain")
            message.attach(part)

            for filename in files:
                # Open PDF file in binary mode
                with open(filename, "rb") as attachment:
                    # Add file as application/octet-stream
                    # Email client can usually download this automatically as attachment
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())

                # Encode file in ASCII characters to send by email    
                encoders.encode_base64(part)

                # Add header as key/value pair to attachment part
                p = Path(filename)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {p.name}",
                )

                # Add attachment to message and convert message to string
                message.attach(part)

            text = message.as_string()

            # Log in to server using secure context and send email
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(self.server, 465, context=context) as server:
                server.login(self.sender_email, self.password_email)
                server.sendmail(self.sender_email, to, text)

            return True

        except Exception as e:
            return askyesno('Mail', 'Une erreur s\'est produite en envoyant un mail.\nAdresse de récéption prévue:\n' + str(to) + '\nErreur:\n' + str(e) + '\nVoulez vous continuer ?')


def importExcelFile(path):
    wb = xl.load_workbook(path)
    dialogs.text('Le fichier {} comporte plusieurs feuilles'.format(path))
    for i, name in enumerate(wb.sheetnames):
        star = ''
        if wb.active.title == name:
            act = i+1
            star = '*'

        dialogs.item(i+1, name, star)

    chx = dialogs.question('Quelle feuille voulez vous utiliser ?', act)
    sh = wb[wb.sheetnames[int(chx)-1]]

    ## Formatage :
    # Colonne A : groupe
    # Colonne B : nom
    # Colonne C : adresse Email

    row = 2
    table = {}
    groupe = None
    while sh.cell(row = row, column = 2).value:
        grp = sh.cell(row = row, column = 1).value
        if grp:
            groupe = grp
            table[groupe] = []

        nom = sh.cell(row = row, column = 2).value
        addr = sh.cell(row = row, column = 3).value
        table[groupe].append((nom, addr))
        row += 1

    return table

def AutoSendMail(table, files, semaine):
    ems = EmailSender()
    sujet = 'Emploi du temps semaine ' + str(semaine)
    content = '''Coucou {name},
Voici en pièce-jointe ton emploi du temps pour la semaine {semaine}.
Normalement il tourne sur la dernière version du colloscope, mais mieux vaut jeter un rapide coup d'oeuil pour vérifier.
S'il y a une erreur, il faut me le signaler !
A plus,
Benoit
PS: j'ai mis à jour la journée de mardi selon les informations de Mme Aubry !
Il faut juste vérifier pour certains d'entre vous qui ont colle avec elle.
'''

    n = 0
    l = 0
    pc = 0
    for _, eleves in table.items():
        l += len(eleves)

    import time
    for groupe, eleves in table.items():
        fichier = files[groupe]
        for nom, addr in eleves:
            n += 1
            pc = int(100 * n / l)
            vpc = int(20 * n / l)
            print(' [' + '='*vpc + ' '*(20-vpc) + '] ' + str(pc) + ' %', end = '\r')

            if not addr:
                continue

            ems.send(addr, sujet,
                     content.format(name = nom, semaine = semaine),
                     files = [fichier],)

    print()


if __name__ == '__main__':
    #es = EmailSender()
    #es.send('benoit.charreyron@orange.fr', 'Premier test !', 'Coucou, voici mon test pour envoyer par mail l\'EDT !', files = ['output/groupe-J.pdf'])
    table = importExcelFile('emails.xlsx')
    fichiers = {l: f'output/groupe-{l}.pdf' for l in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']}
    semaine = 16
    AutoSendMail(table, fichiers, semaine)









