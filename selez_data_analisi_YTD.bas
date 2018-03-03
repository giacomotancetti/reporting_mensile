Attribute VB_Name = "selez_data_analisi_YTD"
Public par_data As Integer 'la variabile par_data indica il numero d'ordine della data selezionata per le analisi fra
                           ' le date disponibili

'costruzione vettore date analisi
Function costr_vett_date()
    
    Dim rig_date_pdc As Integer

    rig_date_pdc = 2    'riga date nel foglio "PdC_Generale"
    'calcolo n° colonne piene nel foglio "PdC_Generale"
    n_col_piene_pdc = Worksheets("PdC_Generale").Rows(rig_date_pdc).Cells.SpecialCells(xlCellTypeConstants).Count
    'calcolo n° celle contenenti date nella riga "rig_date_pdc" nel foglio "PdC_Generale"
    j = 0
    For i = 1 To n_col_piene_pdc
        If IsDate(Worksheets("PdC_Generale").Cells(rig_date_pdc, i)) = True Then
            j = j + 1
            ReDim date_cons(j)
        End If
    Next i
            
    'creazione del vettore "date_cons"
    j = 0
    For i = 1 To n_col_piene_pdc
        If IsDate(Worksheets("PdC_Generale").Cells(rig_date_pdc, i)) = True Then
            j = j + 1
            date_cons(j) = Worksheets("PdC_Generale").Cells(rig_date_pdc, i)
        
        End If
    Next i

    costr_vett_date = date_cons()

End Function

' la funzione legge l'input da tastiera corrispondente al parametro associato ad una delle date disponibili e ricava il mese
' dell'analisi

Function lettura_data_an_YTD(date_cons)

    Dim txt As String

    txt = "Selezionare la data alla quale si vuole eseguire l'analisi Year To Date (YTD): " & "(Es: 1 per data xx/yy/zzzz)" & vbCrLf
    Title = "Reporting - Analisi Year To Date"
    For i = 1 To UBound(date_cons)
        txt = txt & i & ") " & date_cons(i) & vbCrLf
    Next i

    par_data_str = InputBox(txt, Title)
    
    While IsNumeric(par_data_str) = False Or CInt(par_data_str) > (i - 1) Or CInt(par_data_str) <= 0
        MsgBox ("Errore! Inserire un numero intero (primo carattere della riga) indicato nella finestra successiva")
        par_data_str = InputBox(txt)
    Wend

    par_data = CInt(par_data_str)
    ' data dell'analisi selezionata dall'utente
    data_an_YTD = date_cons(par_data)
    
    lettura_data_an_YTD = data_an_YTD

            
End Function


