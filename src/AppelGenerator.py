import excelparser
import edtfiller
import dialogs
import config as confr
import automail

import os
import sys
from pathlib import Path

def find_latest_week(folder):
    p = Path(folder)
    zips = list(p.glob('**/*.zip'))
    numbers = []
    for z in zips:
        z = str(z).split('-')[-1]
        z = z.replace('.zip', '')
        z = z.replace('S', '')
        z = int(z)
        numbers.append(z)

    return max(numbers) + 1

def ask_config():
    dialogs.clear()
    p = Path('config/')
    dialogs.text('Bienvenue, veuillez choisir une configuration')
    config_files = list(p.glob('**/*.ini'))
    for i, fp in enumerate(config_files):
        dialogs.item(i+1, fp.name)

    item = dialogs.question()
    config_file = config_files[int(item)-1]
    config = confr.Configuration(config_file)

    return config

def create_appel():
    config = ask_config()

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
    config = ask_config()

    excel_file = dialogs.question('Lien vers le fichier Excel', default = config.input_file)
    table = excelparser.read_colloscope(excel_file, config)
    semaine = dialogs.question('Semaine à générer', type = int)
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
    config = ask_config()

    emails_file = dialogs.question('Lien vers les adresses mails', default = config.emails_file)
    colloscope_file = dialogs.question('Lien vers le colloscope', default = config.input_file)
    output_folder = dialogs.question('Dossier de sortie', default = config.output_path)
    table = automail.importExcelFile(emails_file)
    semaine = dialogs.question('Numéro de la semaine', type = int)
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
    es.send('bravocharlie1273@orange.fr', 'Emplois du temps', 'Voici tous les emplois du temps.', files = list(fichiers.values()))

    return 1

def general():
    config = ask_config()
    output_folder = dialogs.question('Dossier de sortie', default = config.output_path)
    week = find_latest_week(output_folder)
    semaine = dialogs.question('Numéro de la semaine', type = int, default = week)
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
    thistable = excelparser.selector(table, semaine)
    thistable = excelparser.read_modifs(modif_file, thistable)

    groupes = excelparser.sort_groupes(thistable)
    r = edtfiller.fill_edt(groupes, excel_edt, output_folder, semaine)
    edtfiller.clear()

    weeks = excelparser.all_weeks(table)
    tables = {}
    for _semaine in weeks:
        _thistable = excelparser.selector(table, _semaine)
        tables[_semaine] = _thistable

    app = excelparser.appel(tables, appel_file, config, semaine)

    os.system(output_folder.replace('/', '\\') + f'output-S' + str(semaine) + '.zip')
    dialogs.question("\n\nFin de la génération. Validez pour continuer l'envoi des mails", default = '')

    infos = {}
    for colle in thistable:
        if colle.groupe not in infos:
            infos[colle.groupe] = []

        infos[colle.groupe].append((colle.salle, colle.heure, colle.jour, colle.prof, colle.colle_id))

    groupes_mail = list(infos.keys())
    fichiers = {l: output_folder + f'groupe-{l}.pdf' for l in groupes_mail}
    dialogs.text("\nEnvoi automatique de tous les emplois du temps")
    automail.send_edt(list(fichiers.values()), table_edt, template_edt, semaine)
    dialogs.text("\nEnvoi automatique de toutes les feuilles d'appel")
    automail.send_edt([appel_file.replace('.xlsx', '.pdf')], table_appels, template_appels, semaine)

    dialogs.text("\nEnvoi automatique de tous les emplois du temps")
    automail.AutoSendMail(table_addr, fichiers, semaine, infos, template)
    #es = automail.EmailSender()
    #es.send('bravocharlie1273@orange.fr', 'Emplois du temps', 'Voici tous les emplois du temps.', files = list(fichiers.values()), test = False)
    #es.send('bravocharlie1273@orange.fr', 'Feuilles d\'appel', 'Voici toutes les feuilles pour les appels.', files = [appel_file.replace('.xlsx', '.pdf')], test = False)

    return -1


def quitter():
    return -1

if __name__ == '__main__':
    dialogs.clear()
    dialogs.text('''Programme de gestion du colloscope''')
    actions = [
        (general, 'Programme principal'),
        (create_appel, 'Créer une feuille d\'appel'),
        (create_edts, 'Créer les emplois du temps pour une semaine'),
        (send_mail, 'Envoyer les emplois du temps par mail'),
        (quitter, 'Quitter'),
        ]

    e = 0
    while e >= 0:
        if e == 1:
            dialogs.clear()

        dialogs.text('Menu principal :')

        for i, (_, title) in enumerate(actions):
            dialogs.item(i+1, title)

        dialogs.text()
        action_id = dialogs.question('Choix', default = 1)
        fnct = actions[int(action_id)-1][0]
        e = fnct()
        print()
        dialogs.question('Programme terminé', default = '')
