from pathlib import Path
import os
import sys

import config as confr

NEVER_ASK = True

def clear():
    # Efface le contenu du terminal
    os.system('cls' if sys.platform == 'win32' else 'clear')

def info(*args, **kwargs):
    # Affiche les infos
    print('\x1b[33m', *args, **kwargs)

def text(*args, **kwargs):
    #Affiche du texte
    print('\x1b[0m', *args, **kwargs)

def item(i, name, flag = ''):
    # Affiche un élément de liste (avec drapeau si séléction par défautl par exemple
    print('  ', '\x1b[34m', i, '\x1b[32m', name, '\x1b[31;5m', flag, '\x1b[0m')

def warning(*data, sep = ' ', end = '\n'):
    # Affiche un avertissement (rouge)
    txt = ''
    for d in data:
        txt += str(d) + sep

    print('\x1b[31;5m', txt, '\x1b[0m', end = end)

def question(text = '', default = None, prompt = '>>>', type = None):
    # Pose une question et retourne la réponse.
    # Prise en compte des réponse par défaut

    if default is not None:
        df = '\x1b[34m[{}]'.format(default)
    else:
        df = ''

    print('\x1b[36m', text, df)

    while 1:
        if NEVER_ASK and default is not None:
            ask = ''
        else:
            ask = input('   \x1b[35;5m' + prompt + '\x1b[0m ')

        print()
        if ask == default:
            data = default
            break

        elif not ask and default:
            data = default
            break

        elif not ask:
            pass

        else:
            data = ask
            break

    if type is not None:
        data = type(data)

    return data

def end(action = 'revenir au menu principal'):
    # Texte de fin
    input(f' \x1b[34mAppuyez sur \x1b[37;5mEntrée\x1b[0m \x1b[34mpour {action}...')

def ask_feuille(wb, path, default = None):
    # Demande d'une feuille automatique d'un fichier excel
    act = 0
    text('Le fichier {} comporte plusieurs feuilles'.format(path))
    for i, name in enumerate(wb.sheetnames):
        star = ''
        if wb.active.title == name and default is None:
            act = i+1
            star = '*'

        if default == i:
            star = '*'
            act = i + 1

        item(i+1, name, star)

    chx = question('Quelle feuille voulez vous utiliser ?', act)
    return wb[wb.sheetnames[int(chx)-1]]

# Ouverture de la configuration pour une classe
def ask_config():
    """Demande quel fichier de configuration utiliser.
    Ces derniers sont stockés dans le dossier config/
    Ces derniers sont ensuite utilisés pour le reste du programme

    Entrée
    Rien

    Retourne
    * une class de configuration
    """

    clear()
    # On scanne tous les fichiers
    p = Path('config/')
    text('Bienvenue, veuillez choisir une configuration')
    config_files = list(p.glob('**/*.ini'))
    for i, fp in enumerate(config_files):
        item(i+1, fp.name)

    #  Choix du fichier de configuration
    _item = question()
    config_file = config_files[int(_item)-1]
    # Ouverture de la configuration
    config = confr.Configuration(config_file)

    return config

if __name__ == '__main__':
    text('bonjour !')
    warning('attention !')
    question('comment allez vous ?')
