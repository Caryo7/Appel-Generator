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
        self.emails_file = self._import('path', 'emails_file')
        self.output_path = self._import('path', 'output_path')
        self.edt_path = self._import('path', 'edt_path')
        self.modif_file = self._import('path', 'modification_file')

        self.template_file = self._import('mails', 'template_eleve')
        self.template_edt = self._import('mails', 'template_edt')
        self.template_appels = self._import('mails', 'template_appels')

        self.feuilles = []
        self.layout = {}
        for section in self.parser.sections():
            if 'sheet-' not in section:
                continue

            title = self._import(section, 'title', str)
            data = self._import_list(section, 'data', str)
            col = self._import(section, 'column', int)
            row_base = self._import(section, 'row_base', int)
            self.feuilles.append([data, title])
            self.layout[len(self.feuilles) - 1] = (col, row_base)
