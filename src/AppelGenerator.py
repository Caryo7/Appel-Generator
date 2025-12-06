import excelparser
import edtfiller
import dialogs
import config as confr

import os
import sys
from pathlib import Path

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
    folder = dialogs.question('Lien vers le dossier de sortie', default = 'output/')
    r = edtfiller.fill_edt(groupes, excel_edt, folder, semaine)
    edtfiller.clear()
    if not r:
        return 0

    return 1

    #edtfiller.excel.Quit()
    #del edtfiller.excel

def aide():
    dialogs.clear()
    dialogs.text('Aide')
    dialogs.text("Utilisation des feuilles d'appel")
    dialogs.text("""Le principe des feuilles d'appel est de permettre de créer
la liste des élèves dans chaque groupes.
Pour cela, démarrez l'outil de création de feuille d'appel depuis le menu
principal et suivez les instructions.
Le programme vous demande une configuration. Choisissez en une parmis celles
proposées. Les configurations doivent être enregistrées dans le dossier config.

Le programme charge la personnalisation et vous demande des informations.
Choisissez un fichier : par défaut il est enregistrer dans la configuration.
Si la proposition bleu foncée vous convient, ne tapez rien et faites Entrée.
Vous pouvez aussi écrire un autre nom (chemin absolu ou relatif au programme).

On vous demande ensuite le fichier de sortie des feuilles d'appel. Sur le
même principe, vous pouvez changer le fichier.

Choisissez ensuite la semaine en cours à analyser. Le programme lira uniquement
la colonne correspondante pour vous faire les listes d'appel de la semaine.

Le programme tourne alors seul et génère le fichier de sortie.
Il vous propose de l'ouvrir directement à la fin de la création. Répondez par
oui ou non.

Le fichier peut alors être ouvert. Le programme retourne au menu principal""")
    dialogs.end()
    return 1

def quitter():
    return -1

if __name__ == '__main__':
    dialogs.clear()
    dialogs.text('''Programme de gestion du colloscope''')
    actions = [
        (create_appel, 'Créer une feuille d\'appel'),
        (create_edts, 'Créer les emplois du temps pour une semaine'),
        (aide, 'Aide'),
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
        action_id = dialogs.question('Choix', default = len(actions))
        fnct = actions[int(action_id)-1][0]
        e = fnct()
