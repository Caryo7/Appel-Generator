import config as confr
import dialogs
import excelparser
import excelsaver
import edtfiller
import automail

def send_mail():
    config = dialogs.ask_config()

    emails_file = dialogs.question('Lien vers les adresses mails', default = config.emails_file)
    colloscope_file = dialogs.question('Lien vers le colloscope', default = config.input_file)
    output_folder = dialogs.question('Dossier de sortie', default = config.output_path)
    table = automail.importExcelFile(emails_file)
    semaine = Semaine(dialogs.question('Numéro de la semaine', type = int))
    table_colles = excelparser.read_colloscope(colloscope_file, config)
    thistable = excelparser.selector(table_colles, semaine)
    infos = {}
    for colle in thistable:
        if colle.groupe not in infos:
            infos[colle.groupe] = []

        infos[colle.groupe].append((colle.salle, colle.heure, colle.jour, colle.prof, colle.colle_id))

    groupes = list(infos.keys())
    fichiers = {l: output_folder + f'groupe-{l}.pdf' for l in groupes}
    automail.AutoSendMail(table, fichiers, semaine, infos)
    dialogs.text('Envoi automatique des fichiers à l\'adresse BC')
    es = automail.EmailSender()
    es.send('', 'Emplois du temps', 'Voici tous les emplois du temps.', files = list(fichiers.values()))

    return 1
