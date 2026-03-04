from tkinter.messagebox import *
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import email, smtplib, ssl
import openpyxl as xl
from pathlib import Path
from configparser import ConfigParser
import dialogs
import time

TEST_MODE = False # False # Double sécurité !
IDLE_MODE = '' # ' ' en console

class EmailSender:
    def __init__(self, pwd):
        self.password_email = pwd

    def reload_sender(self):
        """Chargement du server et de son nom
        à partir du fichier de configuration interne.
        """

        parser = ConfigParser()
        parser.read('./config/intern.ini')
        self.sender_email = parser.get('mail', 'email')
        self.server = 'smtp.' + self.sender_email.split('@')[1]

    def send(self, to, subject, text, files = [], test = True):
        """Permet d'envoyer un email automatiquement
        Arguments
        * to : destinataire
        * subject : sujet du message
        * text: corps du message
        * files : fichiers attachés au message
        * test : mode de test activé

        Retourne
        * booléen: erreur ou pas d'erreur
        """

        # On recharge le paramétrage du serveur
        self.reload_sender()

        try:
            # Création d'un mail et paramétrage des entêtes
            message = MIMEMultipart()
            message["From"] = self.sender_email
            message["To"] = to
            message["Subject"] = subject

            # Configuration du type d'affichage de text (plain vs html)
            part = MIMEText(text, "plain")
            message.attach(part)

            for filename in files:
                # Ouverture des fichiers joints et lecture pour encodage
                p = Path(filename)
                with open(filename, "rb") as attachment:
                    #part = MIMEBase("application", "octet-stream")
                    #part.set_payload(attachment.read())
                    part = MIMEApplication(attachment.read(), Name = str(p.name))

                # Encodage des fichiers en Base64 pour l'envoie par mail
                #encoders.encode_base64(part)

                # Ajout des descripteurs de pièce jointe
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {p.name}",
                )

                # Ajout de la pièce jointe
                message.attach(part)

            text = message.as_string()

            # Vérification du mode de test
            if TEST_MODE or test:
                # Simulation de l'envoie du mail, et enregistrement du fichier .eml dans le dossier temporaire
                dialogs.info('Simulation envoi...', end = '')
                time.sleep(0.3)
                f = open(f'temp-emails/{to}-{time.time()}.eml', 'w', encoding = 'utf-8')
                f.write(text)
                f.close()
                return True # Le mail factice a bien été envoyé

            else:
                dialogs.info('                   ', end = '')

            # Création de la connexion
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(self.server, 465, context=context) as server:
                # Ouverture du serveur mail
                server.login(self.sender_email, self.password_email)
                server.sendmail(self.sender_email, to, text) # Envoi

            # Mail bien envoyé
            return True

        except Exception as e: # En cas d'erreur
            print()
            dialogs.warning('Une erreur s\'est produite en envoyant un mail.\nAdresse de récéption prévue:\n' + str(to) + '\nErreur:\n' + str(e) + '\nVoulez vous continuer ?')
            ask = dialogs.question('Continuer ?', default = 'Oui').lower()[0] == 'o'
            return ask


def importExcelFile(path):
    """Importation des tables de excel.
    Arguments
    * path: le fichier excel contenant les adresses mails

    Retourne
    * les listes des liste des emails
    """

    wb = xl.load_workbook(path) # Ouverture du fichier excel
    out = []
    for i, title in enumerate(['élèves', 'emplois du temps', 'appels']):
        dialogs.text('Email pour les', title)
        sh = dialogs.ask_feuille(wb, path, default = i)

        ## Formatage :
        # Colonne A : groupe (si première feuille sinon on commence à la A)
        # Colonne B : nom
        # Colonne C : adresse Email
        # Colonne D : langue vivante de l'élève (facultative: None par défaut)

        row = 2
        table = {}
        groupe = None
        while sh.cell(row = row, column = 2).value: # Tant qu'il y a une valeur
            if i == 0:
                # Si première table, on regarde le groupe présent ou précédent
                grp = sh.cell(row = row, column = 1).value
                if grp:
                    groupe = grp
                    table[groupe] = []

                nom = sh.cell(row = row, column = 2).value
                family = sh.cell(row = row, column = 3).value
                addr = sh.cell(row = row, column = 4).value
                lang = sh.cell(row = row, column = 5).value
                ssgrp = sh.cell(row = row, column = 6).value

                # Ajout à la table en cours
                table[groupe].append((nom, family, addr, lang, ssgrp))

            else: # Sinon, on ne regarde pas le groupe
                nom = sh.cell(row = row, column = 1).value
                addr = sh.cell(row = row, column = 2).value
                table[addr] = nom

            row += 1

        out.append(table)

    return out

def autoformat(message, variables):
    """Fonction d'application de dictionnaire sur text formaté
    A partir d'un dictionnaire et d'une str, retourne le str en
    remplacant les clé de variables par valeur

    Arguments
    * message (str): le texte de template
    * variables (dict): les valeurs à remplacer

    Retourne
    * message: le nouveau message personnalisé
    """

    for k, v in variables.items():
        message = message.replace('{' + str(k) + '}', str(v))

    return message

def ligne_colle(infos):
    """Génère un recap des colles à introduire dans le message.
    Arguments
    * infos: une liste des infos des colles

    Retourne
    * message: un text recap
    """

    txt = ''
    for salle, heure, jour, prof, groupe in infos:
        txt += f' - {prof} {jour.lower()} à {heure} en {salle}\n'

    return txt

def send_edt(fichiers, table_out, template, semaine, pwd):
    """Envoi automatique des autres infos à qui de droit.
    Prend en charge l'envoi avec informations sur la civilité dans le
    modèle et le numéro de la semaine. Permet d'envoyer plusieurs fichier,
    à de nombreuses personne. Exemple: envoyer les listes d'appels aux profs, ...

    Arguments
    * fichiers : liste des fichiers à envoyer
    * table_out : la table des adresses mail
    * template : le fichier modèle de texte
    * semaine : le numéro de la semaine

    Retourne
    Rien
    """

    # Ouverture de l'envoyeur
    ems = EmailSender(pwd)
    sujet = 'Emplois du temps semaine ' + str(semaine)

    # Récupération du modèle
    f = open(template, 'r', encoding = 'utf-8')
    content = f.read()
    f.close()

    for addr, civilite in table_out.items():
        # Envoi du mail
        ems.send(addr, sujet,
                 autoformat(content, {'civilite': civilite,
                                      'semaine': semaine,}),
                 files = fichiers,
                 test = False,
                 )

def AutoSendMail(table, files, semaine, infos, template, pwd):
    """Fonction d'envoi automatique des mails pour les élèves
    Arguments
    * table : la table des adresses mail
    * files : les fichiers pdf des emplois du temps
    * semaine : le numéro de semaine actuel
    * infos : les informations sur les groupes (pour récupérer le nom de la position)
    * template : le fichier modèle de texte
    * pwd: le mot de passe de l'expéditeur

    Retourne
    Rien
    """

    # Ouverture du serveur
    ems = EmailSender(pwd)
    sujet = 'Emploi du temps semaine ' + str(semaine)

    # Lecture du modèle
    f = open(template, 'r', encoding = 'utf-8')
    content = f.read()
    f.close()

    n = 0
    l = 0
    pc = 0
    for _, eleves in table.items():
        l += len(eleves)

    # Sécurité automatique pour ne pas envoyer par erreur tous les mails automatiquement
    ask = dialogs.question('Confirmez la désactivation du mode de test ?', prompt = '[OUI/NON -> NON] >>> ', default = 'NON')
    if ask != 'OUI':
        TEST = True
    else:
        TEST = False

    for groupe, eleves in table.items():
        # Récupération du fichier et de la position du groupe
        #fichier = files[groupe]
        info = infos[groupe]
        for nom, family, addr, _, _ in eleves:
            fichier = files[(nom, family)]
            n += 1
            pc = int(100 * n / l)
            vpc = int(20 * n / l)
            dialogs.text(' [', '='*vpc, ' '*(20-vpc), '] ', str(pc), ' % ', nom, end = ' ', sep = '')

            this_test = TEST
            if not addr: # Si il n'y a pas d'adresse mail, on fait comme si il y avait simulation d'envoi
                this_test = True
                dialogs.warning('NO-MAIL! ', end = '')

            # Envoi automatique
            r = ems.send(addr, sujet,
                         autoformat(content, {'name': nom,
                                              'semaine': int(semaine),
                                              'ds': semaine.DS,
                                              'colles': ligne_colle(info),
                                              'position': info[0][-1]}),

                         files = [fichier],
                         test = this_test,
                         )

            print('\r', end = IDLE_MODE)

            if not r: # Si une erreur s'est produite et qu'on doit s'arrêter, on coupe
                return

    print()
