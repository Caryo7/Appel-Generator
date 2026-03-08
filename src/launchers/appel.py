import config as confr
import dialogs
import excelparser
import excelsaver
import edtfiller
import automail

def create_appel():
    config = dialogs.ask_config()

    excel_file = dialogs.question('Lien vers le fichier Excel', default = config.input_file)
    output_file = dialogs.question('Fichier Excel de sortie', default = config.output_file)

    table = excelparser.read_colloscope(excel_file, config)
    weeks = excelparser.all_weeks(table)

    tables = {}
    for semaine in weeks:
        thistable = excelparser.selector(table, semaine)
        tables[semaine] = thistable

    app = excelparser.appel(tables, output_file, config)

    dialogs.text('La compilation est terminée')
    opn = dialogs.question('Voulez vous ouvrir le fichier créé ?', default = 'oui')
    if opn.lower().startswith('o'):
        os.popen(output_file)

    dialogs.end()
    return 1
