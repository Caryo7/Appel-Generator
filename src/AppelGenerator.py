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
import launchers.general
import box
import graph

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
    ems = box.askEmail() #dialogs.ask_pwd("Mot de passe de l'expéditeur")

    parser = ConfigParser()
    parser.read('config/intern.ini', encoding = 'utf-8')
    run = parser.get('sequence', 'run').split(';')
    show_folder = parser.get('view', 'zip').lower() == 'yes'
    week = find_latest_week('output/')
    semaine = Semaine(box.question('Numéro de la semaine (Attention aux vacances !)', type = int, default = week))

    if not run:
        htest(ems, semaine)

    for config in run:
        launchers.general.general(semaine, ems, show_folder, confr.Configuration(config))

    box.question('Programme terminé...')
    graph.clear()
    print(graph.R)
