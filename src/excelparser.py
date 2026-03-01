## Générateur de liste d'appel
## Programme écrit par Benoit Charreyron
## Copyright 2025
## MIT License
## Fichier principal de traitement interne

from pathlib import Path
from openpyxl.styles import PatternFill
from openpyxl.styles.fonts import Font
import openpyxl as xl
import sys
import os

import config as confr # On utilise le fichier config du dossier
import dialogs
import edtfiller
import excelsaver

class Colle:
    # Attributs de la classe (fixe pour chaque colle)
    salle = None
    heure = None
    jour = None
    semaine = None
    
    # Attributs variables (dépend des colles)
    prof = None # Champs remplis plus tard !
    eleves = None
    groupe = None
    
    # Champ spéficique pour l'élève
    colle_id = None
    
    def __init__(self,
                 salle,
                 heure,
                 jour,
                 semaine,
                 prof,
                 groupe,
                 colle_id,
                 ):
        """Initialisation de la classe colle.
        Arguments
        * salle : la salle de la colle
        * heure : l'heure de la colle
        * jour : le jour dans la semaine de la colle
        * semaine : le numéro de semaine de la colle
        * prof : le professeur qui fait la colle
        * groupe : le groupe qui suit la colle
        * colle_id : l'identifiant de la colle (et du groupe cette semaine là)

        Retourne
        * Colle : la colle
        """

        self.salle = salle
        self.heure = heure
        self.jour = jour
        self.semaine = semaine
        self.prof = prof
        self.groupe = groupe
        self.colle_id = colle_id

    def __repr__(self):
        return '{groupe}-{colle_id}'.format(
            groupe = self.groupe,
            colle_id = self.colle_id
            )

def read_colloscope(path, config):
    """Lecture du fichier excel .xlsx du colloscope.
    Ce dernier sera traité selon la configuration, et retournera une
    liste de colles. Cette liste de colle sera mélangée, mais chaque colle
    aura sont propre identifiant pour se repérer dans le temps !
    
    Arguments
    * path : fichier excel à ouvrir
    
    Sortie
    * table : liste de Colle
    """

    # Section ouverture du fichier excel
    wb = xl.load_workbook(path)
    dialogs.text("Ouverture du colloscope")
    sh = dialogs.ask_feuille(wb, path, config.default_col_sheet)

    # Section lecture de la liste des élèves
    groupes = {}
    col_groupe = config.col_groupe
    col_student = config.col_student
    lignes_eleves = config.lignes_eleves

    for line in lignes_eleves:
        l = line
        while sh.cell(column = col_groupe, row = l).value is None:
            l -= 1

        groupe = sh.cell(column = col_groupe, row = l).value
        std = []
        for col_stud in col_student:
            std.append(sh.cell(column = col_stud, row = line).value)

        student = ' '.join(std)
        if groupe not in groupes:
            groupes[groupe] = []

        groupes[groupe].append(student)

    # Section lecture du colloscope

    # Paramétrage interne de la fonction. Ce dernier sera a terme
    # sauvegardé dans un fichier de configuration .ini.
    # Voir pour avoir les droits d'écriture dans le dossier de travail.
    col_prof = config.col_prof # Format int
    col_salle = config.col_salle # "
    col_heure = config.col_heure # "
    col_jour = config.col_jour # "
    col_id = config.col_id # " Colle identifiant pour savoir dans quel groupe on
                # est pour le reste de l'emploi du temps !

    col_groupes = config.col_groupes # Format list
    
    lignes_semaine = config.lignes_semaine # Format int
    lignes = config.lignes # Format list

    # Le choix est fait de parcourir toutes les lignes en premier,
    # On aurait aussi pu parcourir les colonnes en premier
    table = []
    for line in lignes:
        prof = sh.cell(column=col_prof, row=line).value
        salle = sh.cell(column=col_salle, row=line).value
        heure = sh.cell(column=col_heure, row=line).value
        jour = sh.cell(column=col_jour, row=line).value

        l = line
        while sh.cell(column = col_id, row=l).value is None:
            l -= 1
        cid = sh.cell(column=col_id, row=l).value

        #print(prof, salle, heure, jour, cid)

        for line_semaine in lignes_semaine:
            for col in col_groupes:
                semaine = sh.cell(column=col, row = line_semaine).value
                if semaine is None:
                    continue

                l = line
                while sh.cell(column = col, row=l).value is None:
                    l -= 1

                groupe = sh.cell(column = col, row=l).value
                #print('   ', semaine, groupe)
                Kh = Colle(
                    salle,
                    heure,
                    jour,
                    semaine,
                    prof,
                    groupe,
                    cid)

                try:
                    Kh.eleves = groupes[groupe]
                except:
                    dialogs.warning('Attention', prof, 'semaine', semaine, 'a le groupe', groupe, 'qui est inconnu !')

                table.append(Kh)

    wb.close() # Ne pas oublier pour ne pas bloquer excel !

    return table

def get_this_ds(path, config, semaine):
    """Récupération du DS de la semaine dans le colloscope.
    Ouvre le colloscope et vient lire la table spéciale des DS
    Arguments
    * path: le fichier excel du colloscope
    * config: le fichier de configuration
    * semaine: la semaine actuelle

    Retourne
    * un nom de DS
    """

    dialogs.text('Détection du DS')
    wb = xl.load_workbook(path)
    sh = dialogs.ask_feuille(wb, path, default = config.ds_sheet)
    col = config.col_sem_ds
    col_ds = config.col_ds
    lignes = config.lignes_ds
    for l in lignes:
        if sh.cell(column = col, row = l).value == semaine:
            return sh.cell(column = col_ds, row = l).value

def all_weeks(table):
    """Cette fonction prend en argument une table et sort toutes les semaines
    qui y sont présentés"""

    weeks = []
    for colle in table:
        if colle.semaine not in weeks:
            weeks.append(colle.semaine)

    return weeks

def selector(table, semaine):
    """Cette fonction sera à extraire de la liste complète des colles
    uniquement les colles de la semaine intéresée.
    Le principe est le parcours général de la liste, pour l'extraction.
    
    Arguments
    * table : liste de colle de la fonction read_colloscope
    * semaine : identifiant de la semaine (ATTENTION, en str)
     
    Retourne
    * table : liste de colle de la semaine !
    """
    
    extraction = [] # On pourrait le faire par compréhension, mais c'est 
                    # plus indigeste à écrire
    
    for colle in table:
        if colle.semaine == semaine:
            extraction.append(colle)
            extraction[-1].semaine = semaine
            
    return extraction

def find_colle_id(table, groupe):
    """Cette fonction retourne l'identifiant
    du groupe de colle en à partir d'un nom de groupe
    et d'une table de groupes pour une semaine

    Arguments
    * table : la table des colles de la semaine
    * groupe : le groupe dont on cherche l'identifiant

    Retourne
    * colle_id : identifiant de colle si trouvée, None sinon
    """

    for colle in table:
        if colle.groupe.lower() == groupe.lower():
            return colle.colle_id

def read_modifs(modif_file, table):
    """Cette fonction sert à modifier le colloscope principal
    en changeant des colles et/ou en ajoutant / supprimant.
    Cela permet de manipuler la dernière semaine

    Arguments
    * modif_file : le chemin du fichier contenant les modifications
                   (qui doit être sur une seule semaine (après selector)
    * table : la table principale des groupes

    Retourne
    * table : table principal qui a été modifiée
    """

    wb = xl.load_workbook(modif_file)
    sh = dialogs.ask_feuille(wb, modif_file)

    ## Formatage
    # Colonne 1 : Groupe
    # Colonne 2 : Jour
    # Colonne 3 : Heure
    # Colonne 4 : Professeur
    # Colonne 5 : Salle
    # Première ligne : 2

    modifs = []
    row = 1
    colle_id = None
    try:
        semaine = table[0].semaine
    except:
        return None

    while sh.cell(row = row+1, column = 1).value:
        row += 1
        groupe = sh.cell(row = row, column = 1).value
        jour = sh.cell(row = row, column = 2).value
        heure = sh.cell(row = row, column = 3).value
        prof = sh.cell(row = row, column = 4).value
        room = sh.cell(row = row, column = 5).value

        kolle = Colle(room, heure.lower(), jour, semaine, prof, groupe, colle_id)
        modifs.append(kolle)

    output_table = []
    used = []
    for colle_org in table:
        updated = False
        for colle_mod in modifs:
            if colle_mod in used:
                continue

            if colle_org.jour.lower() != colle_mod.jour.lower():
                continue

            if not edtfiller.dt(colle_org.heure.lower(), colle_mod.heure.lower()):
                continue

            if colle_org.groupe.lower() != colle_mod.groupe.lower():
                continue

            colle_org.salle = colle_mod.salle
            colle_org.prof = colle_mod.prof
            output_table.append(colle_org)
            used.append(colle_mod)
            updated = True

        if not updated:
            output_table.append(colle_org)


    for colle_mod in modifs:
        if colle_mod in used:
            continue

        colle_id = find_colle_id(table, colle_mod.groupe)
        colle_mod.colle_id = colle_id
        output_table.append(colle_mod)

    return output_table


class Week:
    def __init__(self, groupe_id):
        self.groupe_id = groupe_id # = colle_id
        self.colles = []
        
    def append(self, colle):
        self.colles.append(colle)

    def __repr__(self):
        return '{groupe_id}-{nbcolle}'.format(
            groupe_id = self.groupe_id,
            nbcolle = len(self.colles))


def sort_groupes(table):
    """Cette fonction prend une table sur une semaine uniquement !
    (Voir pour passer la table comùplète par selector(table, semaine)
     pour l'extraire) et retourne un dictionnaire par groupe de colle.
     Ce dernier dictionnaire contiendra la colle de la semaine.
     Si le groupe possède plusieurs colles, il faut voir si on met une 
     liste...
     
     Arguments
     * table : une liste de colle pour UNE semaine
     
     Retourne
     * dictionnaire: (groupe: colles (class Week))
     """

    groupes = {}
    for colle in table:
        if colle.groupe not in groupes:
            groupes[colle.groupe] = Week(colle.colle_id)

        groupes[colle.groupe].append(colle)

    return groupes

def sort_profs(table):
    """Cette fonction prend une table sur une semaine uniquement !
    (Voir pour passer la table comùplète par selector(table, semaine)
     pour l'extraire) et retourne un dictionnaire par colleur.
     Ce dernier dictionnaire contiendra la colle de la semaine.
     Si le colleurfait plusieurs colles, il faut voir si on met une 
     liste...
     
     Arguments
     * table : une liste de colle pour UNE semaine

     Retourne
     * dictionnaire: (prof: colles)
     """

    profs = {}
    for colle in table:
        if colle.prof in profs:
            profs[colle.prof].append(colle)
        else:
            profs[colle.prof] = [colle]
            
    return profs

def appel(tables, path, config, week = None):
    """Création d'une feuille d'appel pour chaque groupes pour chaque
    semaine. Cela utilise une table sur une semaine !

    Arguments
    * tables : un dictionnaire de table, une table par semaine
    * path  : fichier excel à générer
    * config : le fichier de configuration des feuilles
    * week : optionnel pour indiquer une semaine en particulier.
             Si semaine est donné, il n'y aura que la feuille en question !

    Retourne
    Rien
    """

    feuilles = config.feuilles
    wb = xl.Workbook()
    wb.remove(wb.active)
    font_student = Font()
    font_title = Font(bold = True)
    fill_student = PatternFill()#start_color='bad9ff', end_color='bad9ff', fill_type="solid")
    fill_title = PatternFill()#start_color='bad9ff', end_color='bad9ff', fill_type="solid")
    sheet_layout = config.layout

    for semaine, table in tables.items():
        if week != semaine and week is not None:
            continue

        sh = wb.create_sheet('Semaine-{}'.format(semaine))
        sh.cell(row = 1, column = 1).value = 'Appel semaine ' + str(semaine)
        sh.cell(row = 1, column = 1).font = font_title
        sh.cell(row = 1, column = 1).fill = fill_title
        sh.page_margins.left = 1
        sh.page_margins.right = 1
        sh.page_margins.top = 1
        sh.page_margins.bottom = 1
        sh.sheet_properties.pageSetUpPr.fitToPage = True

        data = {i: [] for i in range(len(feuilles))}
        for colle in table:
            for i, sheet in enumerate(feuilles):
                if colle.colle_id in sheet[0]:
                    data[i].append(colle)

        extract = {}
        eleves = []
        for i, colles in data.items():
            extract[i] = []
            for colle in colles:
                if colle.eleves is None:
                    continue

                for eleve in colle.eleves:
                    if eleve not in extract[i]:
                        extract[i].append(eleve)
                        if len(eleve) not in eleves:
                            eleves.append(len(eleve))

        for k in extract:
            extract[k].sort() # Tri alphabétique des noms dans la feuille d'appel

        for i, (_, title) in enumerate(feuilles):
            col, row_base = sheet_layout[i]
            sh.cell(row = row_base, column = col).value = title
            sh.cell(row = row_base, column = col).font = font_title
            sh.cell(row = row_base, column = col).fill = fill_title

            for line, student in enumerate(extract[i]):
                sh.cell(row = line + row_base + 1, column = col).value = student
                sh.cell(row = line + row_base + 1, column = col).font = font_student
                sh.cell(row = line + row_base + 1, column = col).fill = fill_student

            sh.column_dimensions[chr(col+64)].width = max(eleves + [len(title)]) * 1.2

    wb.save(path)

    if week is not None or 1:
        excelsaver.export_pdf(os.path.abspath(path))

