Attribute VB_Name = "SP_funzioni_dati"
'La funzione costruisce la matrice "matr_dati_bdgt" che contiene i dati del foglio "CE_bdgt_carica"
Function calc_matr_dati_bdgt_SP(cod_rag)

    Dim col_cc_bdgt_SP As Integer
    Dim col_cr_bdgt_SP As Integer
    Dim n_rig_piene_bdgt_SP As Integer
    Dim num_date As Integer
    Dim col_val_bdgt_SP As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    'Dichiarazione del numero di colonna colonne che contegono i dati
    col_cc_bdgt_SP = 40  'colonna codici di conto nel foglio "SP_bdgt_carica" !!!NON UTILIZZATO!!!
    col_cr_bdgt_SP = 17   'colonna codici di raggruppamento nel foglio "SP_bdgt_carica"
    
    'calcolo n° righe piene nel foglio "SP_bdgt_carica"
    n_rig_piene_bdgt_SP = Worksheets("SP_bdgt_carica").Cells(Rows.Count, col_cr_bdgt_SP).End(xlUp).Row
    
    'calcolo del numero delle date di analisi
    num_date = 12
    
    'costruzione della matrice "matr_dati_sp_bdgt"
    col_val_bdgt_SP = 18  'colonna di partenza per lettura valori nel foglio "SP_bdgt_carica"

    ReDim matr_dati_bdgt_SP(UBound(cod_rag, 1), num_date, 2, n_rig_piene_bdgt_SP) As String

    'Struttura della matrice "matr_dati_sp_bdgt"
    'matr_dati_sp_bdgt[codice raggruppamento, data analisi, codice di conto, valori]
    For i = 1 To UBound(cod_rag, 1)
        cod_i = cod_rag(i, 1)
        For j = 1 To num_date
            For k = 2 To n_rig_piene_bdgt_SP
                If Worksheets("SP_bdgt_carica").Cells(k, col_cr_bdgt_SP) = cod_i Then
                    matr_dati_bdgt_SP(i, j, 1, k) = Worksheets("SP_bdgt_carica").Cells(k, col_cc_bdgt_SP)  'nella colonna 1 è riportato il codice conto
                    matr_dati_bdgt_SP(i, j, 2, k) = Worksheets("SP_bdgt_carica").Cells(k, col_val_bdgt_SP + j)  'nella colonna 2 è riportato il valore [€]
                End If
            Next k
        Next j
    Next i
  
    'Assegnazione del valore 0 agli elementi vuoti della matrice
    For i = 1 To UBound(matr_dati_bdgt_SP, 1)
            For j = 1 To UBound(matr_dati_bdgt_SP, 2)
                For k = 1 To UBound(matr_dati_bdgt_SP, 4)
                    If matr_dati_bdgt_SP(i, j, 2, k) = "" Then
                        matr_dati_bdgt_SP(i, j, 2, k) = 0
                    End If
                Next k
            Next j
    Next i
    
    'Assegnazione del valore da restituire dalla funzione
    calc_matr_dati_bdgt_SP = matr_dati_bdgt_SP

End Function

'La funzione calcola il vettore delle somme dei valori a budget per ogni codice di raggruppamento
Function calc_somme_bdgt_SP(matr_dati_bdgt_SP, cod_rag)

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    'calcolo del numero delle date di analisi
    num_date = 12

    ReDim somme_bdgt_SP(UBound(cod_rag), num_date) As Double

    For i = 1 To UBound(cod_rag)
        For j = 1 To num_date
            For k = 1 To UBound(matr_dati_bdgt_SP, 4)
                somme_bdgt_SP(i, j) = somme_bdgt_SP(i, j) + matr_dati_bdgt_SP(i, j, 2, k)
            Next k
        Next j
    Next i
    
    'Assegnazione del valore da restituire dalla funzione
    calc_somme_bdgt_SP = somme_bdgt_SP

End Function

'La funzione calcola il vettore delle somme dei valori a budget per ogni codice di raggruppamento per analisi PER
Function calc_somme_bdgt_SP_PER(somme_bdgt_SP, mese_PER)

    'calcolo della matrice "somme_bdgt_SP_PER" relativa all'analisi di periodo
    ReDim somme_bdgt_SP_PER(UBound(somme_bdgt_SP, 1))
    For i = 1 To UBound(somme_bdgt_SP, 1)
        somme_bdgt_SP_PER(i) = somme_bdgt_SP(i, mese_PER + 1) - somme_bdgt_SP(i, mese_PER)
    Next i
    
    calc_somme_bdgt_SP_PER = somme_bdgt_SP_PER
    
End Function
