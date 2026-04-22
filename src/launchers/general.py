import config as confr
import dialogs
import excelparser
import excelsaver
import edtfiller
import automail

import os
import box

def general(semaine, ems, prof_manager, show_folder = True, config = None):
    if not config:
        config = dialogs.ask_config()

    classe = config.classe
    output_folder = box.question('Dossier de sortie', default = config.output_path)
    emails_file = box.question('Adresses mails', default = config.emails_file)
    colloscope_file = box.question('Colloscope', default = config.input_file)
    excel_edt = box.question('Emplois du temps', default = config.edt_path)
    template = box.question('Template mail (TXT)', default = config.template_file)
    template_edt = box.question('Template mail (EDT)', default = config.template_edt)
    template_appels = box.question('Template mail (Appels)', default = config.template_appels)
    appel_file = box.question('Feuille d\'appel', default = config.output_file)
    modif_file = box.question('Modification de dernière minute', default = config.modif_file)
    table_addr, table_edt, table_appels = automail.importExcelFile(emails_file)

    table = excelparser.read_colloscope(colloscope_file, config, table_addr)
    semaine.DS = excelparser.get_this_ds(colloscope_file, config, semaine)
    thistable = excelparser.selector(table, semaine)
    thistable = excelparser.read_modifs(modif_file, thistable)
    if thistable is None:
        box.warning("Ouverture", [f"La semaine {semaine} n'existe pas ! Attention aux vacances !",])
        return 1

    iscolles = box.question('Colles', default = 'oui' if semaine.DS is not None else 'non') == 'oui'
    groupes = excelparser.sort_groupes(thistable)
    profs = excelparser.sort_profs(thistable, classe)
    prof_manager.feed(profs)
    #prof_manager.start() ##### DEBUG !!

    r = edtfiller.fill_edt(groupes, excel_edt, output_folder, semaine, table_addr, config, iscolle = iscolles)
    edtfiller.clear()

    weeks = excelparser.all_weeks(table)
    tables = {}
    for _semaine in weeks:
        _thistable = excelparser.selector(table, _semaine)
        tables[int(_semaine)] = _thistable

    app = excelsaver.appel(tables, appel_file, config, semaine)

    if show_folder:
        os.system(os.path.join(output_folder, config.output_zip.format(semaine)))

    #box.question("Fin de la génération. Validez pour continuer l'envoi des mails", default = '')
    os.system(f'edit "{template}"')

    infos = {}
    for colle in thistable:
        if colle.groupe not in infos:
            infos[colle.groupe] = []

        infos[colle.groupe].append((colle.salle, colle.heure, colle.jour, colle.prof, colle.colle_id))

    groupes_mail = list(infos.keys())
    fichiers = {}
    for grp, people in table_addr.items():
        for nom, family, _, _, _ in people:
            fichiers[(nom, family)] = output_folder + f'groupe-{grp}-{nom}.{family}.pdf'

    automail.AutoSendMail(table_addr, fichiers, semaine, infos, template, ems)
    ask = box.question("Envoi automatiques des mail administration ?", default = 'oui')
    if ask.lower() == 'oui':
        #dialogs.text("\nEnvoi automatique de tous les emplois du temps")
        automail.send_edt(list(fichiers.values()), table_edt, template_edt, semaine, ems)
        #dialogs.text("\nEnvoi automatique de toutes les feuilles d'appel")
        automail.send_edt([appel_file.replace('.xlsx', '.pdf')], table_appels, template_appels, semaine, ems)

    return -1
