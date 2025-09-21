from configparser import *

class Configuration:
    def _import(self, section, option, type = int):
        data = self.parser.get(section, option)
        try:
            return type(data)
        except:
            return str(data)

    def _import_list(self, section, option, type = int):
        data = self.parser.get(section, option)
        data = data.replace(' ', '')
        data = data.split(';')
        data = list(map(type, data))
        return data

    def __init__(self, path):
        self.parser = ConfigParser()
        self.parser.read(path, encoding = 'utf-8')

        self.col_groupe = self._import('groupes', 'col_groupes')
        self.col_student = self._import_list('groupes', 'col_student')
        self.lignes_eleves = self._import_list('groupes', 'lignes_eleves')

        self.col_prof = self._import('colloscope', 'col_prof')
        self.col_salle = self._import('colloscope', 'col_salle')
        self.col_heure = self._import('colloscope', 'col_heure')
        self.col_jour = self._import('colloscope', 'col_jour')
        self.col_id = self._import('colloscope', 'col_id')
        
        self.col_groupes = self._import_list('colloscope', 'col_groupes')
        
        self.lignes_semaine = self._import_list('colloscope', 'lignes_semaine')
        self.lignes = self._import_list('colloscope', 'lignes')

        self.output_file = self._import('path', 'output_file')
        self.input_file = self._import('path', 'input_file')
        self.feuille_title = self._import('path', 'feuille_title')

        feuilles = []
        for section in self.parser.sections():
            if 'sheet-' not in section:
                continue

            title = self._import(section, 'title', str)
            data = self._import_list(section, 'data', str)
            feuilles.append([data, title])

        self.feuilles = feuilles
