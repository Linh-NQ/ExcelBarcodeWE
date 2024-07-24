import pandas as pd
import glob
import os
import xlwings as xw
from xlwings.constants import LineStyle
from xlwings import utils
from xlwings import Range

import datetime

import shutil
import tempfile
import pyperclip


# Funktion, um Zwischenablage zu löschen
def clear_windows_clipboard():
    #os.system("echo off | clip")
    pyperclip.copy('')
    
# Makro erstellen für Feedback
def run_excel_macro_with_text(spendenankunft, text):
    spendenankunft.macro('ShowPopup')(text)

# Funktion, welches Makro ausführt, um Dropdown zu erstellen
def create_dropdown(wb, dropdown_liste, zelle):
    wb.macro("CreateDropdown")(dropdown_liste, zelle)


# red_flag initialisieren (gibt an, wenn Master ID überschritten wurde)
red_flag = False
red_flag_middle = False


        
def go_dodo_macro(path):
    """ Funktion für Makro, um Excel auszufüllen
    """
    
    global red_flag_middle, red_flag, flag_barcode, flag_id_vergeben, flag_id_überprüfen, flag_barcode_scannen, flag_barcode_gescannt, flag_id_überschritten, flag_id_neu_vergeben, flag_datensatz, flag_id, flag_master_error
    
    # Identifizieren der Datenbank
    data_path = r'O:\Serumbank\Wareneingang\Vorlagen'
    files = glob.glob(os.path.join(data_path, '*.xlsx*'))
    for file in files:
        if ('Datenbanktool' in file) & ('Kopie' not in file):
            data_path = file
            break
            
    # Kopie der Datenbank erstellen
    datenbank_copy = "O:\\Serumbank\\Wareneingang\\Vorlagen\\Datenbank_Copy.xlsx"
    shutil.copyfile(data_path, datenbank_copy)
            
    datenbank_ok = True
    try:
        datenbank_pd = pd.read_excel("O:\\Serumbank\\Wareneingang\\Vorlagen\\Datenbank_Copy.xlsx", sheet_name = 'Datenbank-Spendeneingang', header = 1)
        # gleichzeitig Produktnummern ablesen
        produkt_nr = pd.read_excel("O:\\Serumbank\\Wareneingang\\Vorlagen\\Datenbank_Copy.xlsx", sheet_name = 'Produktnummern')
    except:
        import traceback
        traceback.print_exc()
        datenbank_ok = False
    
    if datenbank_ok:

        # flags für Fehlermeldungen
        flag_barcode = False
        flag_id_vergeben = False
        flag_id_überprüfen = False
        flag_barcode_scannen = False
        flag_barcode_gescannt = False
        flag_id_überschritten = False
        flag_id_neu_vergeben = False
        flag_datensatz = False
        flag_id = False
        flag_master_error = False


        data_prim_id = []
        for i in range(len(datenbank_pd)):
            data_prim_id.append(datenbank_pd.iloc[i,0])
        # Datenbank nach Primär ID splitten und Tabellen in dictionary speichern
        # Liste mit Start-Index für jede Tabelle
        start_indices = []
        for i in range(len(data_prim_id)):
            if (str(data_prim_id[i]).lower() != 'x') & (data_prim_id[i] != 'None'):
                start_indices.append(i+3)
        table_ranges = []
        for i in range(len(start_indices)-1):
            table_ranges.append('C{}:AE{}'.format(start_indices[i], start_indices[i+1]-1))

        # zuerst Spendenankunft einlesen
        spendenankunft = xw.Book(path)
        sht1 = spendenankunft.sheets['Spendenankunft']
        sht_drop = spendenankunft.sheets['Dropdown']
            
        # Datenbank mit xlwings einlesen
        app = xw.App(visible = False)
        
        try:
            datenbank = xw.Book("O:\\Serumbank\\Wareneingang\\Vorlagen\\Datenbank_Copy.xlsx")
            data_sheet = datenbank.sheets['Datenbank-Spendeneingang']    

            # Ablesen der vergebenen Master IDs
            column = 'B'
            last_row = sht1.range(column + str(sht1.cells.last_cell.row)).end('up').row
            id_vergabe_range = sht1.range(column + '1:' + column + str(last_row))
            id_vergabe = id_vergabe_range.value
            id_vergabe = [entry for entry in id_vergabe if entry is not None]

            # nach Zeile mit roter Primär ID suchen
            # den Schritt skippen, wenn keine Einträge vorhanden sind
            if sht1.range('B3').value is not None:
                last_row_excel = sht1.range('B3').end('down').row
                for i in range(3, last_row_excel):
                    cell = sht1.range('B{}'.format(i))
                    if (cell.api.Font.Color == xw.utils.rgb_to_int((255, 0, 0))) | (cell.api.Font.Color == xw.utils.rgb_to_int((255, 165, 0))):
                        start_row_flag = i
                        red_flag_middle = True
                        break
            
            def fill_excel():
                """ Füllt Excel-Tabelle aus
                """           
                # aktuelle Start-Zeile identifizieren
                # (wird daran erkannt, dass Zelle in Spalte 'Primär ID' leer ist)
                global red_flag, last_row, start_row, data_index, data_index_end, red_flag_middle, red_flag, flag_barcode, flag_id_vergeben, flag_id_überprüfen, flag_barcode_scannen, flag_barcode_gescannt, flag_id_überschritten, flag_id_neu_vergeben, flag_datensatz, flag_id, flag_master_error     
        
                if red_flag:
                    # Wenn im untersuchten Block alle IDs rot sind, wird der gesamte Block ersetzt
                    if start_row == data_index:
                        start_row = last_row
                elif red_flag_middle:
                    start_row = start_row_flag        
                else:
                    start_row = 3
                    while True:        
                        if sht1.range('B{}'.format(start_row)).value is None:
                            break
                        else:
                            start_row +=1

                prim_id_total = sht1.range('B1').value
                # für den Check, ob End-Master ID bereits vergeben wurde
                master_id_given_col = sht1.range('B3:B{}'.format(last_row_excel)).value
                master_id_given = [master for master in master_id_given_col if master is not None]
                # Liste mit vergebenen Master IDs, die nicht rot sind
                column_range = sht1.range('B3:B{}'.format(last_row_excel))
                entries_in_black = []
                for cell in column_range:
                    # falls es sich um den ersten Eintrag handelt:
                    if str(cell.value) == 'None':
                        break
                    # Check if the font color is black (RGB value for black: 0)
                    font_color = cell.api.Font.Color
                    if font_color == 0:
                        entries_in_black.append(cell.value)
                master_id_given = entries_in_black

                if red_flag_middle:
                    new_end_id = sht1.range('E1').value

                    # Check, ob Barcode-Nr. bereits gescannt wurde            
                    scanned_barcodes = sht1.range('C3:C{}'.format(last_row_excel)).value    
                    scanned_barcodes = list(set([barcode for barcode in scanned_barcodes if barcode is not None]))

                     # falls Master ID überschritten und nicht verändert wurde, poppt User Warnung auf
                    if new_end_id in master_id_given:
                        # falls Master ID überschritten und neue Probe eingescannt wurde
                        if prim_id_total not in scanned_barcodes:
                            flag_id_überschritten = True
                        else:
                            flag_id_neu_vergeben = True                

                    else:
                        if prim_id_total in scanned_barcodes:
                            # es werden dann nur die roten Master IDs mit den neuen ersetzt
                            # neue end ID ablesen
                            old_id_list = sht1.range('B{}:B{}'.format(start_row, last_row_excel)).value
                            # Finden der neuen start id
                            new_start_id = sht1.range('C1').value
                            while True:
                                if new_start_id not in master_id_given:
                                    break
                                new_start_id += 1

                            for i in range(len(old_id_list)):
                                cell_update = sht1.range('B{}'.format(start_row+i))
                                if i == 0:
                                    old_id = sht1.range('B{}'.format(start_row+i)).value                       
                                    if cell_update.value == old_id:
                                        if cell_update.api.Font.Color == xw.utils.rgb_to_int((255, 0, 0)):
                                            cell_update.api.Font.Color = xw.utils.rgb_to_int((0, 0, 0))

                                        if new_start_id > sht1.range('C1').value:
                                            cell_update.value = new_start_id
                                        else:
                                            cell_update.value = sht1.range('C1').value
                                            
                                else:
                                    if sht1.range('C{}'.format(start_row+i)).value is not None:
                                        new_start_id += 1
                                    if cell_update.api.Font.Color == xw.utils.rgb_to_int((255, 0, 0)):
                                        
                                        if new_start_id > sht1.range('C1').value:
                                            cell_update.value = new_start_id
                                        else:
                                            cell_update.value = sht1.range('C1').value

                                        old_id = sht1.range('B{}'.format(start_row+i)).value
                                        cell_update.api.Font.Color = xw.utils.rgb_to_int((0, 0, 0))
                                # falls ID wieder überschritten wird:
                                if new_start_id > new_end_id:
                                    cell_update.api.Font.Color = xw.utils.rgb_to_int((255, 0, 0))
                else:            
                    # Suchfunktion
                    # dafür Primär ID bis zum zweiten Bindestrich splitten
                    flag_yellow = False
                    if len(prim_id_total.split('-')) == 3:
                        projektnr = prim_id_total.split('-')[1]
                    else:
                        projektnr = prim_id_total.split('-')[-1]
                        flag_yellow = True

                    # Suche in Datenbank und Index des entsprechenden Eintrags ablesen
                    data_index = None
                    for i in range(len(data_prim_id)):
                        if projektnr in data_prim_id[i]:
                            data_index = i
                            break
                    for i in range(data_index+1,len(data_prim_id)):
                        if data_prim_id[i].upper() == 'X':
                            data_index_end = i
                        else:
                            break

                    # Check, ob Primär ID-Feld orange ist                
                    if (data_sheet.range('A{}'.format(3+data_index)).color == (255, 153, 0)) | (data_sheet.range('A{}'.format(3+data_index)).color == (255, 192, 0)):
                        flag_datensatz = True
                    else:
                        data_table = data_sheet.range('C{}:AE{}'.format(data_index+3, data_index_end+3))
                        start_cell = sht1.range('D{}'.format(start_row))
                        # Copy and paste the table data and formatting
                        data_table.api.Copy()
                        start_cell.api.PasteSpecial(Paste=-4163)  # -4163 corresponds to Paste All             
                        clear_windows_clipboard()

                        # Primär ID in die erste Zeile eintragen
                        cell_id = sht1.range('C{}'.format(start_row))
                        cell_id.value = prim_id_total
                        cell_id.api.Font.Bold = True
                        # restl Zeilen leer und grau
                        empty_id_cells = sht1.range('C{}:C{}'.format(start_row+1, data_index_end-data_index+start_row))
                        empty_id_cells.value = ''
                        empty_id_cells.color = (174, 170, 170)       

                        # Eintragen von Datum
                        heute = datetime.date.today()
                        heute = heute.strftime("%d.%m.%Y")    
                        datum_index = []
                        for j in range(data_index_end-data_index+1):
                            cell_date = sht1.range('F' + str(start_row+j))
                            if (cell_date.color == (255, 255, 0)) | (cell_date.color != (174, 170, 170)):
                                cell_date.value = heute
                                datum_index.append(j)
                        einlagerung_cols = ['L', 'O', 'Q', 'S']
                        for i in range(len(einlagerung_cols)):
                            for j in datum_index:
                                cell_date = sht1.range(einlagerung_cols[i] + str(start_row+j))
                                if cell_date.value != '/':
                                    cell_date.value = heute

                        # Spalte mit Projektnummer ausfüllen
                        # beim Start einer neuen Master ID, trägt die Zeile eine Projektnr.
                        # restl. Zellen darunter werden mit '/' gefüllt

                        # Überprüfen, ob Produktnr vorhanden ist
                        list_projektnr = list(produkt_nr['Projektnummer'])
                        list_produktnr = list(produkt_nr['Produktnummer'])

                        for i in range(len(produkt_nr)):
                            if projektnr in list_projektnr[i]:
                                projektnr = projektnr + '\n' + list_produktnr[i]
                                break


                        ziffer_flag = False
                        for i_nr in range(len(projektnr)):
                            if projektnr[i_nr] in ['1', '2', '3', '4', '5', '6', '7', '8', '9']:
                                ziffer_flag = True
                                break
                        if ziffer_flag:
                            # Projektnr. wird in der gesamten Spalte eingetragen
                            cell_nr = sht1.range('Z{}:Z{}'.format(start_row, start_row+data_index_end-data_index))
                            cell_nr.number_format = 'Standard'
                            cell_nr.value = str(projektnr)
                            cell_nr.color = (255, 255, 255)

                        # Operator-Spalte ausfüllen
                        operator = sht1.range('F1').value
                        cell_ops = [sht1.range('W{}'.format(i + start_row)) for i in datum_index]
                        cell_prims = [sht1.range('C{}'.format(i + start_row)) for i in datum_index]

                        for cell_op in cell_ops:
                            cell_op.value = operator
                            cell_op.color = (255, 255, 255)
                            cell_op.api.Font.ColorIndex = 1

                        for cell_prim in cell_prims:
                            cell_prim.value = prim_id_total
                            if flag_yellow:
                                cell_prim.color = (255, 255, 0)
                            else:
                                cell_prim.color = (255, 255, 255)
                            cell_prim.api.Font.Bold = True


                        # LID-Spalte ergänzen, falls Zelle nicht grau ist
                        for i in range(data_index_end-data_index+1):
                            lid_cell = sht1.cells(start_row+i, 4)
                            lid = prim_id_total.split('-')[0]    
                            try:
                                lid = int(lid)
                            except:
                                lid = ''        
                            if (lid_cell.color != (174, 170, 170)) & (lid_cell.value != lid):
                                lid_cell.value = lid
                                # in weiße Farbe ändern
                                if lid_cell.color != (255, 255, 255):
                                    lid_cell.color = (255, 255, 255)
                                if lid == '': # gelbes Feld
                                    lid_cell.color = (255, 255, 0)

                        # Eintragen der Master ID
                        # Immer wieder neues Ablesen der Start- und End-ID
                        start_id_new = sht1.range('C1').value
                        end_id_new = sht1.range('E1').value
                        # aktuelle ID ablesen und in entsprechende Spalte eintragen
                        if start_row == 3:
                            akt_id = start_id_new
                        else:           
                            if start_id_new in master_id_given:
                                akt_id = sht1.cells(start_row-1, 2).value + 1
                            else:
                                akt_id = start_id_new

                        red_flag = False
                        for i in range(data_index_end-data_index+1):
                            # Master ID weiter hochzählen
                            cell_lid = sht1.range('D'+ str(start_row+i))
                            if cell_lid.color != (174, 170, 170):
                                if i != 0:
                                    akt_id += 1        
                            cell = sht1.cells(start_row+i, 2)
                            cell.value = akt_id
                            cell.api.Font.Color = xw.utils.rgb_to_int((0, 0, 0))
                            # Check, ob tagesaktuelle ID überschritten wird
                            if akt_id > end_id_new:
                                cell.api.Font.Color = xw.utils.rgb_to_int((255, 0, 0))
                                red_flag = True
                                flag_master_error = True
                                last_row = start_row
                            # Check, Master ID bereits vergeben wurde
                            if akt_id in master_id_given:
                                flag_id = True
                                cell.api.Font.Color = xw.utils.rgb_to_int((255,165,0))


                        # Spalte 'Bemerkungen zum Lager' ausfüllen
                        heute = datetime.date.today()
                        kw = heute.strftime("%U")
                        for i in range(data_index_end-data_index+1):
                            cell = sht1.range('AF' + str(start_row+i))
                            cell_value = cell.value
                            if cell_value == '/':
                                cell.value = 'KW'+kw
                                cell.color = (255, 255, 0)

                        # Dropdowns einfügen für die Spalten F, G, H, T, U, AA, AB und AD
                        def return_dropdown_list(spalte):
                            dropdown_range = sht_drop.range('{}2'.format(spalte)).expand('down')
                            dropdown_values = list(filter(lambda x: pd.notna(x), dropdown_range.value))
                            return dropdown_values
                        
                        dropdown_g = return_dropdown_list('A')
                        dropdown_h = return_dropdown_list('B')
                        dropdown_i = return_dropdown_list('C')
                        dropdown_u = return_dropdown_list('C')
                        dropdown_m = return_dropdown_list('D')
                        dropdown_v = return_dropdown_list('E')
                        dropdown_ab = return_dropdown_list('B')
                        dropdown_ac = return_dropdown_list('F')
                        dropdown_ae = return_dropdown_list('G')
                        
                        # Zustand nach Prozessierung anpassen
                        matrix = sht1.range('I' + str(start_row)).value
                        if 'faeces' in matrix:
                            stuhl_index = dropdown_v.index('Stuhl')+1
                            dropdown_v = dropdown_v[stuhl_index:]
                        elif 'urin' in matrix:
                            stuhl_index = dropdown_v.index('Stuhl')
                            urin_index = dropdown_v.index('Urin')+1
                            dropdown_v = dropdown_v[urin_index:stuhl_index]
                        elif ('swab' in matrix) | ('saliv' in matrix):
                            urin_index = dropdown_v.index('Urin')
                            speichel_index = dropdown_v.index('Speichel/Swaps')+1
                            dropdown_v = dropdown_v[speichel_index:urin_index]
                        else:
                            speichel_index = dropdown_v.index('Speichel/Swaps')
                            dropdown_v = dropdown_v[1:speichel_index]

                        # Zeilen, wo Dropdowns erstellt werden sollen aus datum_index ablesen
                        for index_drop in datum_index:
                            index_drop = index_drop + start_row
                            create_dropdown(spendenankunft, dropdown_g, 'G{}'.format(index_drop))
                            create_dropdown(spendenankunft, dropdown_h, 'H{}'.format(index_drop))
                            create_dropdown(spendenankunft, dropdown_i, 'I{}'.format(index_drop))
                            create_dropdown(spendenankunft, dropdown_m, 'M{}'.format(index_drop))
                            create_dropdown(spendenankunft, dropdown_u, 'U{}'.format(index_drop))
                            create_dropdown(spendenankunft, dropdown_v, 'V{}'.format(index_drop))
                            create_dropdown(spendenankunft, dropdown_ab, 'AB{}:AB{}'.format(start_row, data_index_end-data_index+start_row))
                            create_dropdown(spendenankunft, dropdown_ac, 'AC{}:AC{}'.format(start_row, data_index_end-data_index+start_row))
                            create_dropdown(spendenankunft, dropdown_ae, 'AE{}:AE{}'.format(start_row, data_index_end-data_index+start_row))

                        # Formatierung: Umrandung
                        range_line = sht1.range('B{}:AF{}'.format(start_row, data_index_end-data_index+start_row))
                        range_line.api.Borders.LineStyle = LineStyle.xlContinuous
                        range_line.api.Borders.Weight = 2
                        # Apply thick lines as an outer border around the table
                        bottom_border = range_line.api.Borders(9)  # 9 represents the bottom border
                        bottom_border.LineStyle = LineStyle.xlContinuous
                        bottom_border.Weight = 3
                        left_border = range_line.api.Borders(7)  # 7 represents the left border
                        left_border.LineStyle = LineStyle.xlContinuous
                        left_border.Weight = 3
                        right_border = range_line.api.Borders(10)  # 10 represents the right border
                        right_border.LineStyle = LineStyle.xlContinuous
                        right_border.Weight = 3           


            # Primär ID ablesen (befindet sich im Scannfeld)
            sht1 = spendenankunft.sheets['Spendenankunft']
            global prim_id_total
            prim_id_total = sht1.range('B1').value
            if prim_id_total is None:
                flag_barcode_scannen = True
            else:
                # Check, falls Barcode schon mal gescannt wurde
                last_row_excel = sht1.range('B3').end('down').row
                scanned_barcodes = list(sht1.range('C3:C{}'.format(last_row_excel)).value)
                scanned_barcodes = [barcode for barcode in scanned_barcodes if barcode is not None]

                if (prim_id_total in scanned_barcodes) & (red_flag_middle == False):
                    flag_barcode_gescannt = True
                else:
                    # Check, ob zweite Master ID größer ist als die erste
                    start_id_new = sht1.range('C1').value
                    end_id_new = sht1.range('E1').value
                    try:
                        if start_id_new >= end_id_new:
                            flag_id_überprüfen = True         
                        else:
                            fill_excel()
                    except:
                        # Meldung, wenn keine Start und End Master IDs eingegeben wurden
                        if (start_id_new is None) | (end_id_new is None):
                            flag_id_vergeben = True
                        # Meldung, wenn Barcodenr. nicht gefunden wurde
                        else:
                            flag_barcode = True
                             
        
        finally:
            # Datenbank schließen
            datenbank.close()
            app.quit()

        # Scannfeld löschen
        scanfield = sht1.range('B1')
        scanfield.value = ''
        # Scannfeld auswählen (über Makro)
        spendenankunft.macro("SelectCellB1")()


        # Meldungen werden zum Schluss angezeigt, damit Programm komplett ausgeführt wird
        if flag_barcode:
            run_excel_macro_with_text(spendenankunft, 'Barcodenr. nicht in Datenbank gefunden')
        if flag_id_vergeben:
            run_excel_macro_with_text(spendenankunft, 'Bitte zu vergebene Master ID eintragen')
        if flag_id_überprüfen:
            run_excel_macro_with_text(spendenankunft, 'Bitte Master ID überprüfen')
        if flag_barcode_scannen:
            run_excel_macro_with_text(spendenankunft, 'Bitte Barcode scannen')  
        if flag_barcode_gescannt:
            run_excel_macro_with_text(spendenankunft, 'Dieser Barcode wurde bereits gescannt')
        if flag_id_überschritten:
            run_excel_macro_with_text(spendenankunft, 'Master ID überschritten und neue Probe eingescannt')
        if flag_id_neu_vergeben:
            run_excel_macro_with_text(spendenankunft, 'Master ID überschritten, bitte neue vergeben')
        if flag_datensatz:
            run_excel_macro_with_text(spendenankunft, "Spendendatensatz noch nicht freigegeben")
        if flag_master_error:
            run_excel_macro_with_text(spendenankunft, 'Master ID wurde überschritten.')
        if flag_id:
            run_excel_macro_with_text(spendenankunft, 'Master ID wurde schon vergeben.')


        
    else:
        spendenankunft = xw.Book(path)
        run_excel_macro_with_text(spendenankunft, 'Datenbank von jemandem geöffnet') 

    
# Read the variable from the file
file_path = 'O:\\Serumbank\\Wareneingang\\Vorlagen\\Spendenankunft_run\\path.txt'
with open(file_path, 'r') as f:
    path = f.read()
path = path.strip('\n')
    
go_dodo_macro(path)


# In[ ]:




