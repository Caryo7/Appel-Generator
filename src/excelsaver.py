from win32com.client import DispatchEx
import dialogs

excel = DispatchEx('Excel.Application')
excel.Visible = 0

def export_pdf(path_from, path_to = None):
    if path_to is None:
        path_to = path_from.replace('.xlsx', '.pdf')

    try:
        wb = excel.Workbooks.Open(path_from)
        wb.application.displayalerts = False
        ws = wb.Worksheets[0]
        ws.SaveAs(path_to, FileFormat=57)
        wb.Close()

    except Exception as e:
        dialogs.warning("Erreur d'exportation ! Fermez les taches Excel déjà démarrées")
        excel.Quit()
        return False

    return path_to
