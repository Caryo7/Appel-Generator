E = '\x1b[{}m'
R = '\x1b[0m'
import sys
import os
try:
    col, height = os.get_terminal_size()
except:
    exit()

COLORS_FG = {'red': 31,
             'green': 32,
             'black': 30,
             'blue': 34,
             'yellow': 33,
             'magenta': 35,
             'cyan': 36,
             'white': 37,
             }

COLORS_BG = {'red': 41,
             'green': 42,
             'black': 40,
             'blue': 44,
             'yellow': 43,
             'magenta': 45,
             'cyan': 46,
             'white': 47,
             }

STYLES = {'bold': 1,
          'italic': 3,
          'underline': 4,
          'blink': 5,
          'rapid': 6,
          'double': 21,
          }

class Theme:
    background = {'bg': 'blue', 'fg': 'red'}
    text = {'bg': 'white', 'fg': 'black'}
    cursor = {'bg': 'white', 'fg': 'red', 'style': 'blink'}
    question = {'bg': 'blue', 'fg': 'green'}
    soft = {'bg': 'blue', 'fg': 'red', 'style': 'blink'}

theme = Theme()

def clear():
    if sys.platform in ('win32', 'nt'):
        os.system('cls')
    else:
        os.system('clear')

def balise(style):
    cmd = ''
    if 'bg' in style:
        cmd += str(COLORS_BG[style['bg']]) + ';'
    if 'fg' in style:
        cmd += str(COLORS_FG[style['fg']]) + ';'
    if 'style' in style:
        cmd += str(STYLES[style['style']]) + ';'

    return E.format(cmd[:-1])

def length(txt):
    u = ''
    inb = False
    for c in txt:
        if c == '\x1b':
            inb = True

        elif inb and c == 'm':
            inb = False
            continue

        if not inb:
            u += c

    return len(u)

def c(*args, style):
    return balise(style) + ''.join(args)

def cnt(txt, ln = col):
    n = length(txt)
    a = (ln - n)//2
    return a*' ' + txt + (ln-n-a)*' '

def emptyLine():
    return c(' |' + ' '*(col-4) + '| ', style = theme.background) + R

def start_line():
    txt1 = ' '*(col)
    txt2 = ' +' + '-'*(col-4) + '+ '
    r1 = c(txt1, style=theme.background) + R
    r2 = c(txt2, style=theme.background) + R
    r3 = emptyLine()
    return '\n'.join([r1, r2, r3]) + '\n'

def end_line():
    txt1 = ' '*(col)
    txt2 = ' +' + '-'*(col-4) + '+ '
    r1 = c(txt1, style=theme.background) + R
    r2 = c(txt2, style=theme.background) + R
    r3 = emptyLine()
    return '\n'.join([r3, r2, r1]) + '\n'

def inFrame(*args):
    txt = ''
    for a in args:
        txt += a

    return c(' |', style=theme.background) + txt + c('| ', style=theme.background) + R

def autoWrap(txt, lines = []):
    buffer = ''
    for i, c in enumerate(txt):
        buffer += c
        if len(buffer) > col - 6:
            buffer = ' '.join(buffer.split(' ')[:-1])
            lines.append(buffer)
            return autoWrap(txt[len(buffer):], lines)

    return lines + [buffer]

def centerLine(*args, cursor = False):
    txt = c(*args, style = theme.text)
    if cursor:
        txt = txt.replace('_', c('_', style = theme.cursor) + balise(theme.text))
    else:
        txt = txt.replace('_', '')

    return inFrame(
        cnt(
            txt + R + balise(theme.background),
            ln = col-4,
            )
        )

def setLines(*args, cursor = False):
    lines = autoWrap(''.join(args))
    txt = ''
    for line in lines:
        txt += centerLine(line, cursor = cursor)

    return txt

def centerText(*lines):
    n = len(lines)
    nup = (height - n - 4)//2
    txt = start_line()
    txt += (nup-2)*emptyLine()
    for l in lines:
        txt += l

    txt += (height - len(lines) - nup - 7)*emptyLine()
    txt += end_line()
    return txt

def finalPrint(txt, asking = '', prompt = '>>> '):
    print(balise(theme.background))
    ans = None
    while not ans:
        clear()
        print(txt, end = '')
        ans = input(c(asking, style = theme.question) + ' ' + c(prompt, style = theme.soft) + R + balise(theme.background))

    return ans

class Prompt:
    def __init__(self, name, value = '', show = None):
        self.name = name
        self.value = value
        self.show = show

    def get(self):
        if self.show != '*':
            return self.name + ' : _' + self.value
        else:
            return self.name + ' : _' + '*'*len(self.value)

def askData(txt, prompts, nothing = True):
    title = setLines(txt)
    space = emptyLine()

    cursor = 0
    while 1:
        options = []
        default = ''
        for i, pro in enumerate(prompts):
            c = cursor == i
            if c:
                default = pro.value

            options.append(setLines(pro.get(), cursor = cursor == i))

        data = centerText(title, space, *options)
        out = finalPrint(data, asking = default)

        if out == 'next':
            return prompts, 'next'
        elif out == 'prev':
            return prompts, 'prev'
        elif out == 'exit':
            return None, 'exit'
        else:
            prompts[cursor].value = out
            cursor = (cursor + 1) % len(prompts)


print(askData('Veuillez entrer vos coordonnées pour l\'envoi autamatique des mails',
              prompts = [Prompt('Adresse mail'), Prompt('Mot de passe', show='*')],
              ))

print(R)
