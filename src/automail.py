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
import time

TEST_MODE = False

class EmailSender:
    def reload_sender(self):
        parser = ConfigParser()
        parser.read('./config/intern.ini')
        self.sender_email = parser.get('mail', 'email')
        self.password_email = parser.get('mail', 'password')
        self.server = 'smtp.' + self.sender_email.split('@')[1]

    def send(self, to, subject, text, files = [], test = True):
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
            if TEST_MODE or test:
                dialogs.info('Simulation envoi...', end = '')
                time.sleep(0.3)
                f = open(f'temp-emails/{to}-{time.time()}.eml', 'w', encoding = 'utf-8')
                f.write(text)
                f.close()
                return True
            else:
                dialogs.info('                   ', end = '')

            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(self.server, 465, context=context) as server:
                server.login(self.sender_email, self.password_email)
                server.sendmail(self.sender_email, to, text)

            return True

        except Exception as e:
            print()
            dialogs.warning('Une erreur s\'est produite en envoyant un mail.\nAdresse de récéption prévue:\n' + str(to) + '\nErreur:\n' + str(e) + '\nVoulez vous continuer ?')
            ask = dialogs.question('Continuer ?', default = 'Oui').lower()[0] == 'o'
            return ask


def importExcelFile(path):
    wb = xl.load_workbook(path)
    out = []
    for i, title in enumerate(['élèves', 'emplois du temps', 'appels']):
        dialogs.text('Email pour les', title)
        sh = dialogs.ask_feuille(wb, path, default = i)

        ## Formatage :
        # Colonne A : groupe
        # Colonne B : nom
        # Colonne C : adresse Email

        row = 2
        table = {}
        groupe = None
        while sh.cell(row = row, column = 2).value:
            if i == 0:
                grp = sh.cell(row = row, column = 1).value
                if grp:
                    groupe = grp
                    table[groupe] = []

                nom = sh.cell(row = row, column = 2).value
                addr = sh.cell(row = row, column = 3).value
                table[groupe].append((nom, addr))

            else:
                nom = sh.cell(row = row, column = 1).value
                addr = sh.cell(row = row, column = 2).value
                table[addr] = nom

            row += 1

        out.append(table)

    return out

def autoformat(message, variables):
    for k, v in variables.items():
        message = message.replace('{' + str(k) + '}', str(v))

    return message

def ligne_colle(infos):
    txt = ''
    for salle, heure, jour, prof, groupe in infos:
        txt += f' - {prof} {jour.lower()} à {heure} en {salle}\n'

    return txt

def send_edt(fichiers, table_out, template, semaine):
    ems = EmailSender()
    sujet = 'Emplois du temps semaine ' + str(semaine)
    f = open(template, 'r', encoding = 'utf-8')
    content = f.read()
    f.close()

    for addr, civilite in table_out.items():
        ems.send(addr, sujet,
                 autoformat(content, {'civilite': civilite,
                                      'semaine': semaine,}),
                 files = fichiers,
                 test = False,
                 )

def AutoSendMail(table, files, semaine, infos, template):
    ems = EmailSender()
    sujet = 'Emploi du temps semaine ' + str(semaine)
    f = open(template, 'r', encoding = 'utf-8')
    content = f.read()
    f.close()

    n = 0
    l = 0
    pc = 0
    for _, eleves in table.items():
        l += len(eleves)

    ask = dialogs.question('Confirmez la désactivation du mode de test ?', prompt = '[OUI/NON -> NON] >>> ', default = 'NON')
    if ask != 'OUI':
        TEST = True
    else:
        TEST = False

    for groupe, eleves in table.items():
        fichier = files[groupe]
        info = infos[groupe]
        for nom, addr in eleves:
            n += 1
            pc = int(100 * n / l)
            vpc = int(20 * n / l)
            dialogs.text(' [', '='*vpc, ' '*(20-vpc), '] ', str(pc), ' % ', nom, end = ' ', sep = '')

            this_test = TEST
            if not addr:
                this_test = True
                dialogs.warning('NO-MAIL! ', end = '')

            r = ems.send(addr, sujet,
                         autoformat(content, {'name': nom,
                                              'semaine': semaine,
                                              'colles': ligne_colle(info),
                                              'position': info[0][-1]}),

                         files = [fichier],
                         test = this_test,
                         )

            print('\r', end = '')

            if not r:
                return

    print()


if __name__ == '__main__':
    lettres = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    fichiers = {l: f'output/groupe-{l}.pdf' for l in lettres}
    print('premier mail')
    es = EmailSender()
    es.send('bravocharlie1273@orange.fr', 'Emplois du temps', 'Voici tous les emplois du temps.', files = list(fichiers.values()))

    print('second mail')
    es = EmailSender()
    es.send('bravocharlie1273@orange.fr', 'Premier test !', 'Coucou, voici mon test pour envoyer par mail l\'EDT !', files = ['output/groupe-J.pdf'])
    #table = importExcelFile('emails.xlsx')
    #fichiers = {l: f'output/groupe-{l}.pdf' for l in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']}
    #semaine = 16
    #AutoSendMail(table, fichiers, semaine)









