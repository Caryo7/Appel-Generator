import excelparser
import edtfiller
import dialogs
from configparser import *
import config as confr
import automail
import excelsaver

import os
import sys
from pathlib import Path

class Semaine:
    def __init__(self, numero):
        self.me = numero
        self.DS = 'A'

    def __repr__(self):
        return self.me

    def __str__(self):
        return str(self.me)

    def __int__(self):
        return int(self.me)

    def __eq__(self, nb):
        return nb == self.me

# Récupération du nom de la dernière semaine à partir du dernier fichier ZIP généré
def find_latest_week(folder):
    """Scanne le répertoire courant et retourne
    le plus grand chiffre contenu dans le nom d'un fichier.

    Entrée
    * folder: dossier à observer

    Retourne
    * chiffre le plus grand
    """

    # Scan du dossier des fichier ZIP
    p = Path(folder)
    zips = list(p.glob('**/*.zip'))
    numbers = []
    for z in zips:
        # Récupération du nom de la semaine
        z = str(z).split('-')[-1]
        z = z.replace('.zip', '')
        z = z.replace('S', '')
        z = int(z)
        numbers.append(z) # Ajout du nombre

    return max(numbers) + 1 # On prendra par défaut la semaine suivante

def create_appel():
    config = dialogs.ask_config()

    excel_file = dialogs.question('Lien vers le fichier Excel', default = config.input_file)
    output_file = dialogs.question('Fichier Excel de sortie', default = config.output_file)

    table = excelparser.read_colloscope(excel_file, config)
    weeks = excelparser.all_weeks(table)

    tables = {}
    for semaine in weeks:
        thistable = excelparser.selector(table, semaine)
        tables[semaine] = thistable

    app = excelparser.appel(tables, output_file, config)

    dialogs.text('La compilation est terminée')
    opn = dialogs.question('Voulez vous ouvrir le fichier créé ?', default = 'oui')
    if opn.lower().startswith('o'):
        os.popen(output_file)

    dialogs.end()
    return 1

def create_edts():
    config = dialogs.ask_config()

    excel_file = dialogs.question('Lien vers le fichier Excel', default = config.input_file)
    table = excelparser.read_colloscope(excel_file, config)
    semaine = Semaine(dialogs.question('Semaine à générer', type = int))
    thistable = excelparser.selector(table, semaine)
    groupes = excelparser.sort_groupes(thistable)
    excel_edt = dialogs.question('Lien vers le fichier Excel EDT', default = 'EDT-PT.xlsx')
    folder = dialogs.question('Lien vers le dossier de sortie', default = config.output_path)
    r = edtfiller.fill_edt(groupes, excel_edt, folder, semaine)
    edtfiller.clear()

    if not r:
        return 0

    return 1

    #edtfiller.excel.Quit()
    #del edtfiller.excel

def send_mail():
    config = dialogs.ask_config()

    emails_file = dialogs.question('Lien vers les adresses mails', default = config.emails_file)
    colloscope_file = dialogs.question('Lien vers le colloscope', default = config.input_file)
    output_folder = dialogs.question('Dossier de sortie', default = config.output_path)
    table = automail.importExcelFile(emails_file)
    semaine = Semaine(dialogs.question('Numéro de la semaine', type = int))
    table_colles = excelparser.read_colloscope(colloscope_file, config)
    thistable = excelparser.selector(table_colles, semaine)
    infos = {}
    for colle in thistable:
        if colle.groupe not in infos:
            infos[colle.groupe] = []

        infos[colle.groupe].append((colle.salle, colle.heure, colle.jour, colle.prof, colle.colle_id))

    groupes = list(infos.keys())
    fichiers = {l: output_folder + f'groupe-{l}.pdf' for l in groupes}
    automail.AutoSendMail(table, fichiers, semaine, infos)
    dialogs.text('Envoi automatique des fichiers à l\'adresse BC')
    es = automail.EmailSender()
    es.send('', 'Emplois du temps', 'Voici tous les emplois du temps.', files = list(fichiers.values()))

    return 1

def general(semaine, pwd, show_folder = True, config = None):
    if not config:
        config = dialogs.ask_config()

    output_folder = dialogs.question('Dossier de sortie', default = config.output_path)
    emails_file = dialogs.question('Adresses mails', default = config.emails_file)
    colloscope_file = dialogs.question('Colloscope', default = config.input_file)
    excel_edt = dialogs.question('Emplois du temps', default = config.edt_path)
    template = dialogs.question('Template mail (TXT)', default = config.template_file)
    template_edt = dialogs.question('Template mail (EDT)', default = config.template_edt)
    template_appels = dialogs.question('Template mail (Appels)', default = config.template_appels)
    appel_file = dialogs.question('Feuille d\'appel', default = config.output_file)
    modif_file = dialogs.question('Modification de dernière minute', default = config.modif_file)
    table_addr, table_edt, table_appels = automail.importExcelFile(emails_file)

    table = excelparser.read_colloscope(colloscope_file, config)
    semaine.DS = excelparser.get_this_ds(colloscope_file, config, semaine)
    thistable = excelparser.selector(table, semaine)
    thistable = excelparser.read_modifs(modif_file, thistable)
    if thistable is None:
        dialogs.warning(f"La semaine {semaine} n'existe pas ! Attention aux vacances !")
        return -1

    groupes = excelparser.sort_groupes(thistable)
    dialogs.text('Lancement de la génération')
    r = edtfiller.fill_edt(groupes, excel_edt, output_folder, semaine, table_addr, config)
    edtfiller.clear()

    weeks = excelparser.all_weeks(table)
    tables = {}
    for _semaine in weeks:
        _thistable = excelparser.selector(table, _semaine)
        tables[int(_semaine)] = _thistable

    app = excelsaver.appel(tables, appel_file, config, semaine)

    if show_folder:
        os.system(output_folder.replace('/', '\\') + f'output-S' + str(semaine) + '.zip')

    dialogs.question("\n\nFin de la génération. Validez pour continuer l'envoi des mails", default = '')

    infos = {}
    for colle in thistable:
        if colle.groupe not in infos:
            infos[colle.groupe] = []

        infos[colle.groupe].append((colle.salle, colle.heure, colle.jour, colle.prof, colle.colle_id))

    groupes_mail = list(infos.keys())
    fichiers = {}
    for grp, people in table_addr.items():
        for nom, family, _, _, _ in people:
            fichiers[(nom, family)] = output_folder + f'groupe-{grp}-{nom}.{family}.pdf'

    ask = dialogs.question("Envoi automatiques des mail administration ?", default = 'oui')
    if ask.lower() == 'oui':
        dialogs.text("\nEnvoi automatique de tous les emplois du temps")
        automail.send_edt(list(fichiers.values()), table_edt, template_edt, semaine, pwd)
        dialogs.text("\nEnvoi automatique de toutes les feuilles d'appel")
        automail.send_edt([appel_file.replace('.xlsx', '.pdf')], table_appels, template_appels, semaine, pwd)

    dialogs.text("\nEnvoi automatique de tous les emplois du temps")
    automail.AutoSendMail(table_addr, fichiers, semaine, infos, template, pwd)
    return -1


def quitter():
    return -1

def htest(pwd, semaine):
    # Programme principal
    dialogs.clear()
    dialogs.text('''Programme de gestion du colloscope''')
    actions = [ # Actions disponibles
        (lambda: general(pwd, semaine), 'Programme principal'),
        (create_appel, 'Créer une feuille d\'appel'),
        (create_edts, 'Créer les emplois du temps pour une semaine'),
        (send_mail, 'Envoyer les emplois du temps par mail'),
        (quitter, 'Quitter'),
        ]

    # On continue tant que la commande d'arrêt n'est pas déclenchée
    e = 0
    while e >= 0:
        if e == 1:
            dialogs.clear()

        dialogs.text('Menu principal :')

        for i, (_, title) in enumerate(actions):
            dialogs.item(i+1, title)

        dialogs.text()
        # Choix d'une action
        action_id = dialogs.question('Choix', default = 1)
        fnct = actions[int(action_id)-1][0]
        # Lancement de la fonction
        e = fnct()
        print()
        dialogs.question('Programme terminé', default = '')

if __name__ == '__main__':
    pwd = dialogs.ask_pwd("Mot de passe de l'expéditeur")

    parser = ConfigParser()
    parser.read('config/intern.ini', encoding = 'utf-8')
    run = parser.get('sequence', 'run').split(';')
    show_folder = parser.get('view', 'zip').lower() == 'yes'
    week = find_latest_week('output/')
    semaine = Semaine(dialogs.question('Numéro de la semaine (Attention aux vacances !)', type = int, default = week))

    if not run:
        htest(pwd, semaine)

    for config in run:
        general(semaine, pwd, show_folder, confr.Configuration(config))

    input('\x1b[0mProgramme terminé... ')
