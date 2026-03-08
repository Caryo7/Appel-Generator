from graph import *
import automail
import getpass

def askEmail():
    mail = Prompt('Adresse mail')
    pwd = Prompt('Mot de passe', show = '*', fnct = getpass.getpass)
    data = askData('Veuillez entrer vos coordonnées mail',
                         prompts = [mail, pwd])

    es = automail.EmailSender(data[0][0].value, data[0][1].value)
    return es

def question(text = '', default = '', prompt = '>>> ', type = lambda k: k):
    pro = Prompt(value = default)
    pro.splitter = ''
    a = askData(text, [pro], prompt)
    if a[1] == 'next':
        return type(a[0][0].value)

ATTENTION = '''   ^                
  /|\\               
 / | \\   ATTENTION !
/__.__\\             '''

def warning(title, lines = []):
    warn = ATTENTION.split('\n')
    ls = []
    for w in warn:
        ls.append(setLines(c(w, style = theme.warning), cursor = None))

    ls.append(emptyLine())
    ls.append(setLines(title))
    ls.append(emptyLine())
    for l in lines:
        ls.append(setLines(l))

    data = centerText(*ls)
    return finalPrint(data)

def ask_feuille(title, wb, path, default = None):
    act = 0
    ls = []
    ls.append(setLines(title))
    ls.append(emptyLine())
    ls.append(setLines('Le fichier {} comporte plusieurs feuilles'.format(path)))
    ls.append(emptyLine())
    lmax = max(list(map(lambda k: len(k), wb.sheetnames))) + 10
    for i, name in enumerate(wb.sheetnames):
        flag = ''
        if wb.active.title == name and default is None:
            act = i+1
            flag = '*'

        if default == i:
            flag = '*'
            act = i + 1

        txt = f'  {str(i+1)}. {str(name)} {str(flag)}'
        txt = txt + ' '*(lmax - len(txt))
        ls.append(setLines(txt))

    ls.append(emptyLine())
    ls.append(setLines('Quelle feuille voulez vous utiliser ?'))
    data = centerText(*ls)
    chx = finalPrint(data, asking = str(act), aloadempty = True)
    if not chx:
        chx = act

    return wb[wb.sheetnames[int(chx)-1]]
