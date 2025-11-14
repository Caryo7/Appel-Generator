import os
import sys

def clear():
    os.system('cls' if sys.platform == 'win32' else 'clear')


def text(*args, **kwargs):
    print('\x1b[0m', *args, **kwargs)

def item(i, name, flag = ''):
    print('  ', '\x1b[34m', i, '\x1b[32m', name, '\x1b[31;5m', flag, '\x1b[0m')

def warning(*data, sep = ' '):
    txt = ''
    for d in data:
        txt += str(d) + sep

    print('\x1b[31;5m', txt, '\x1b[0m')

def question(text = '', default = None, prompt = '>>>', type = None):
    if default is not None:
        df = '\x1b[34m[{}]'.format(default)
    else:
        df = ''

    print('\x1b[36m', text, df)
    ask = input('   \x1b[35;5m' + prompt + '\x1b[0m ')
    print()
    if not ask:
        data = default
    else:
        data = ask

    if type is not None:
        data = type(data)

    return data

def end(action = 'revenir au menu principal'):
    input(f' \x1b[34mAppuyez sur \x1b[37;5mEntrée\x1b[0m \x1b[34mpour {action}...')

if __name__ == '__main__':
    text('bonjour !')
    warning('attention !')
    question('comment allez vous ?')
