from win32com.client import DispatchEx
import dialogs

excel = DispatchEx('Excel.Application')
excel.Visible = 0

def export_pdf(path_from, path_to = None):
    """Enregistrement du fichier excel. Utilise une
    bibliothèque de control de Excel pour demander à excel de
    faire le travail.

    Arguments
    * path_from: le fichier excel à exporter
    * path_to: le fichier PDF de sortie

    Retourne
    * path_to
    """

    if path_to is None:
        # Si on ne donne pas de destination, on prendra le même
        # nom que le fichier excel
        path_to = path_from.replace('.xlsx', '.pdf')

    try:
        wb = excel.Workbooks.Open(path_from)
        wb.application.displayalerts = False
        ws = wb.Worksheets[0]
        ws.SaveAs(path_to, FileFormat=57)
        wb.Close()

    except Exception as e:
        # Si erreur d'exportation on retournera faux, pour erreur
        dialogs.warning("Erreur d'exportation ! Fermez les taches Excel déjà démarrées")
        excel.Quit()
        return False

    return path_to
