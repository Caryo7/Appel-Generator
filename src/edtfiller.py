import openpyxl as xl
import os
from win32com.client import DispatchEx
from pathlib import Path

import dialogs

excel = DispatchEx('Excel.Application')
excel.Visible = 0

# A ce stade, on a de quoi extraire le colloscope du fichier excel.
# On peut récupérer la liste des colles de chaque groupes, et leur POSIX
# On peut en faire de même pour chaque prof.
# Il reste alors à faire un programme qui lit l'emploit du temps des élèves
# Et qui insert aux bons endroits les cours.
# Pour faire plus simple, on supposera que les horraires de tout le monde
# Sont les mêmes, et que seules les matières changent.
# Alors, on peut écrire

def dt(h1, h2):
    h1, m1 = h1.split('h')
    h2, m2 = h2.split('h')
    h1 = int(h1)
    h2 = int(h2)
    m1 = int(m1) if m1 else 0
    m2 = int(m2) if m2 else 0

    dh = abs(h2-h1)
    dm = abs(m2-m1)
    return dh == 0 and dm <= 15


class EDT:
    def __init__(self, path):
        self.wb = xl.load_workbook(path)
        self.colles = []
        
    def me(self, groupe, semaine):
        self.groupe = groupe
        for col in range(1, 40): # range du nombre de colonnes de l'EDT
            for row in range(1, 30): # range du nombre de lignes de l'EDT
                v = self.sh.cell(column=col, row = row).value
                if v is None:
                    continue

                if v.upper() == 'PT':
                    self.sh.cell(column=col, row=row).value = 'PT - GROUPE {} - SEMAINE {}'.format(groupe, semaine)

    def feed(self, groupe_id):
        groupe_id = groupe_id.replace('/', '-')
        self.sh = self.wb[groupe_id]
        for sh in list(self.wb):
            if sh.title == groupe_id:
                continue

            del self.wb[sh.title]
    
    def fill(self, colle):
        day_line = 3
        col_heure = 1
        for col in range(1, 40): # range du nombre de colonnes de l'EDT
            for row in range(1, 30): # range du nombre de lignes de l'EDT
                v = self.sh.cell(column=col, row = row).value
                if v is None:
                    continue

                if v.upper() == 'COLLE':
                    heure_edt = self.sh.cell(column=col_heure, row = row).value.lower().split('-')[0]
                    jour_edt = self.sh.cell(column = col, row = day_line).value.lower()
                    heure_cl = colle.heure.lower().replace('00', '').split('-')[0]
                    jour_cl = colle.jour.lower()

                    if jour_edt == jour_cl and dt(heure_cl, heure_edt):
                        self.sh.cell(column=col, row = row).value = '{colleur} {salle}'.format(
                            salle = colle.salle,
                            colleur = colle.prof)

                        return True

        return False

    def export(self, folder):
        for col in range(1, 40): # range du nombre de colonnes de l'EDT
            for row in range(1, 30): # range du nombre de lignes de l'EDT
                v = self.sh.cell(column=col, row = row).value
                if v is None:
                    continue

                if v.upper() == 'COLLE':
                    self.sh.cell(column=col, row = row).value = None

        fp = os.path.join(os.path.abspath('.'), folder, 'groupe-{}'.format(self.groupe))
        #try:
        self.wb.save(fp + '.xlsx')
        #except:
        #    dialogs.warning("Erreur d'enregistrement, fichier ouvert (cf processus arrière plan)")

        #try:
        wb = excel.Workbooks.Open(fp + '.xlsx')
        wb.application.displayalerts = False
        ws = wb.Worksheets[0]
        ws.SaveAs(fp + '.pdf', FileFormat=57)
        wb.Close()

        #except:
            #dialogs.warning("Erreur sur l'exportation, recommencez en fermant tout !")
            #excel.Quit()
            #del excel

            #return False

        return True

def import_edt(path):
    """Ouvre l'emploi du temps des élèves, et en fait une "copie"
    Arguments
    * path : le lien vers le fichier qui contient l'emploi du temps
             au format excel.
             
    Retourne un emploi du temps (class EDT)
    """

    e = EDT(path)
    return e

def fill_edt(groupes, path, folder):
    """Avec un dictionnaire des groupes de colle et un emploi du temps,
    vient remplir les trous volontairement laissés par le professeur
    en charge de la création des emplois du temps.
    Chaque groupe de colle à donc un emploi du temps qui lui est propre.
    Sous réserve bien sur des hypothèses simplificatrices susmentionnées
    plus haut à propos des cours qui ont lieux en même temps.
    
    Arguments
    * groupes : dictionnaires (groupe: colles)
    * path    : Fichier emploi du temps
    * folder  : dossier de sortie des emplois du temps générés
    """

    n = len(groupes)
    i = 0
    pc = 0.0
    opc = 0.0
    for groupe, semaine in groupes.items():
        groupe_id = semaine.groupe_id
        edt = EDT(path)
        edt.feed(groupe_id)
        edt.me(groupe, semaine.colles[0].semaine) # On récupère le numéro de la semaine via la première colle du groupe

        ce = []
        for colle in semaine.colles:
            r = edt.fill(colle)
            if not r:
                ce.append(str(colle))

        if ce:
            dialogs.warning('Une ou plusieurs colle indéterminées !', ', '.join(ce), '\n')

        r = edt.export(folder)
        if not r:
            return False

        i += 1
        pc = int(100*i/n)
        if pc != opc:
            opc = pc
            dialogs.text('\r [' + '='*(int(pc/5)) + ' '*(20-int(pc/5)) + '] {} %'.format(pc))

    return True

def clear():
    for fp in list(Path('output/').glob('**/*.xlsx')):
        os.remove(fp)


if __name__ == '__main__':
    import excelparser as ep
    import config as confr
    config = confr.Configuration('config/PT.ini')
    table = ep.read_colloscope('colloscope.xlsx', config)
    semaine = 6
    thistable = ep.selector(table, semaine)
    groupes = ep.sort_groupes(thistable)
    r = fill_edt(groupes, 'EDT-PT.xlsx', 'output/')
    clear()
    if not r:
        #return False
        quit()

    excel.Quit()
    del excel


