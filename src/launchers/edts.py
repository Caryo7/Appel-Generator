import config as confr
import dialogs
import excelparser
import excelsaver
import edtfiller
import automail

def create_edts():
    config = dialogs.ask_config()

    excel_file = dialogs.question('Lien vers le fichier Excel', default = config.input_file)
    table = excelparser.read_colloscope(excel_file, config)
    semaine = Semaine(dialogs.question('Semaine à générer', type = int))
    thistable = excelparser.selector(table, semaine)
    groupes = excelparser.sort_groupes(thistable)
    excel_edt = dialogs.question('Lien vers le fichier Excel EDT', default = 'EDT-PT.xlsx')
    folder = dialogs.question('Lien vers le dossier de sortie', default = config.output_path)
    r = edtfiller.fill_edt(groupes, excel_edt, folder, semaine)
    edtfiller.clear()

    if not r:
        return 0

    return 1

    #edtfiller.excel.Quit()
    #del edtfiller.excel

