import os
import sys

def clear():
    os.system('cls' if sys.platform == 'win32' else 'clear')

def info(*args, **kwargs):
    print('\x1b[33m', *args, **kwargs)

def text(*args, **kwargs):
    print('\x1b[0m', *args, **kwargs)

def item(i, name, flag = ''):
    print('  ', '\x1b[34m', i, '\x1b[32m', name, '\x1b[31;5m', flag, '\x1b[0m')

def warning(*data, sep = ' ', end = '\n'):
    txt = ''
    for d in data:
        txt += str(d) + sep

    print('\x1b[31;5m', txt, '\x1b[0m', end = end)

def question(text = '', default = None, prompt = '>>>', type = None):
    if default is not None:
        df = '\x1b[34m[{}]'.format(default)
    else:
        df = ''

    print('\x1b[36m', text, df)

    while 1:
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
    input(f' \x1b[34mAppuyez sur \x1b[37;5mEntrée\x1b[0m \x1b[34mpour {action}...')

def ask_feuille(wb, path, default = None):
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

if __name__ == '__main__':
    text('bonjour !')
    warning('attention !')
    question('comment allez vous ?')
