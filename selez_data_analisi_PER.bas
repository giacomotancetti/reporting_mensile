Attribute VB_Name = "selez_data_analisi_PER"
' la funzione legge l'input da tastiera corrispondente al parametro associato ad una delle date disponibili e ricava il mese
' dell'analisi DI PERIOD0
Function lettura_data_an_PER(date_cons)

    Dim txt As String

    txt = "Selezionare la data alla quale si vuole eseguire l'analisi DI PERIODO (PER): " & "(Es: 1 per data xx/yy/zzzz)" & vbCrLf
    Title = "Reporting - Analisi Di Periodo"
    For i = 1 To UBound(date_cons)
        txt = txt & i & ") " & date_cons(i) & vbCrLf
    Next i

    par_data_str = InputBox(txt, Title)
    
    While IsNumeric(par_data_str) = False Or CInt(par_data_str) > (i - 1) Or CInt(par_data_str) <= 0
        MsgBox ("Errore! Inserire un numero intero (primo carattere della riga) indicato nella finestra successiva")
        par_data_str = InputBox(txt)
    Wend

    par_data_PER = CInt(par_data_str)
    ' data dell'analisi selezionata dall'utente
    data_an_PER = date_cons(par_data_PER)
    
    lettura_data_an_PER = data_an_PER
           
End Function


