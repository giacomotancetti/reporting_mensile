Attribute VB_Name = "CE_funzioni_dati"
'La funzione costruisce la matrice "cod_rag" che contiene le carateristiche di ogni codice di raggruppamento definito nel foglio
'"codifiche"
Function calc_cod_rag()
    
    Dim col_cr As Integer
    Dim col_descr As Integer
    Dim col_segno As Integer
    Dim n_rig_piene As Integer
    Dim i As Integer
            
    'Dichiarazione del numero di colonna colonne che contegono i dati
    col_cr = 1
    col_descr = 2
    col_segno = 3
    
    'calcolo n° righe piene nel foglio "codifiche"
    n_rig_piene = Worksheets("codifiche").Cells(Rows.Count, col_cr).End(xlUp).Row

    ReDim cod_rag(n_rig_piene, 3) As String
    
    'Costruzione della matrice "cod_rag_loc"
    For i = 1 To n_rig_piene
        cod_rag(i, 1) = Worksheets("codifiche").Cells(i, col_cr)
        cod_rag(i, 2) = Worksheets("codifiche").Cells(i, col_descr)
        cod_rag(i, 3) = Worksheets("codifiche").Cells(i, col_segno)
    Next i
    
    'Assegnazione del valore da restituire dalla funzione
    'cod_rag(codice raggruppamento, descrizione, segno)
    calc_cod_rag = cod_rag
       
End Function

'La funzione costruisce la matrice "matr_dati_cons" che contiene i dati del foglio "PdC_Generale"
Function calc_matr_dati_cons(date_cons, cod_rag)

    Dim col_cc_pdc As Integer
    Dim col_cr_pdc As Integer
    Dim col_val_pdc As Integer
    Dim n_rig_piene_cc As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    'Dichiarazione del numero di colonna colonne che contegono i dati
    col_cc_pdc = 1      'colonna codici di conto nel foglio "PdC_Generale"
    col_cr_pdc = 5      'colonna codici di raggruppamento nel foglio "PdC_Generale"
    col_val_pdc = 7     'colonna valore nel foglio "PdC_Generale"
    
    'calcolo n° righe piene nel foglio "PdC_Generale"
    n_rig_piene_cc = Worksheets("PdC_Generale").Cells(Rows.Count, col_cc_pdc).End(xlUp).Row
    
    'calcolo del numero delle date di analisi
    num_date = UBound(date_cons)
    
    ReDim matr_dati_cons(UBound(cod_rag, 1), num_date, 2, n_rig_piene_cc) As String

    'Costruzione della matrice "matr_dati_cons"
    'matr_dati_cons[codice raggruppamento, data analisi, codice di conto, valori]
    For i = 1 To UBound(cod_rag, 1)
        cod_i = cod_rag(i, 1)
            For j = 1 To num_date
                For k = 1 To n_rig_piene_cc
                    If Worksheets("PdC_Generale").Cells(k, col_cr_pdc) = cod_i Then
                        matr_dati_cons(i, j, 1, k) = Worksheets("PdC_Generale").Cells(k, col_cc_pdc)      'nella colonna 1 è riportato il codice conto
                        matr_dati_cons(i, j, 2, k) = Worksheets("PdC_Generale").Cells(k, col_val_pdc + j) 'nella colonna 2 è riportato il valore [euro]
                    End If
                Next k
            Next j
    Next i
    
    'Assegnazione del valore 0 agli elementi vuoti della matrice
    For i = 1 To UBound(matr_dati_cons, 1)
            For j = 1 To UBound(matr_dati_cons, 2)
                For k = 1 To UBound(matr_dati_cons, 4)
                    If matr_dati_cons(i, j, 2, k) = "" Then
                        matr_dati_cons(i, j, 2, k) = 0
                    End If
                Next k
            Next j
    Next i
    
    'Assegnazione del valore da restituire dalla funzione
    calc_matr_dati_cons = matr_dati_cons

End Function

'La funzione calcola il vettore delle somme dei valori a consuntivo per ogni codice di raggruppamento
Function calc_somme_cons(matr_dati_cons, cod_rag, date_cons)

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    'calcolo del numero delle date di analisi
    num_date = UBound(date_cons)

    ReDim somme_cons(UBound(cod_rag), num_date) As Double

    For i = 1 To UBound(cod_rag)
        For j = 1 To num_date
            For k = 1 To UBound(matr_dati_cons, 4)
                somme_cons(i, j) = somme_cons(i, j) + matr_dati_cons(i, j, 2, k)
            Next k
        Next j
    Next i
    
    'Assegnazione del valore da restituire dalla funzione
    calc_somme_cons = somme_cons

End Function

'La funzione costruisce la matrice "matr_dati_bdgt" che contiene i dati del foglio "CE_bdgt_carica"
Function calc_matr_dati_bdgt_CE(cod_rag)

    Dim col_cc_cebdgt As Integer
    Dim col_cr_cebdgt As Integer
    Dim n_rig_piene_cebdgt As Integer
    Dim num_date As Integer
    Dim col_val_cebdgt As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    'Dichiarazione del numero di colonna colonne che contegono i dati
    col_cc_cebdgt = 20  'colonna codici di conto nel foglio "CE_bdgt_carica" !!!NON UTILIZZATO!!!
    col_cr_cebdgt = 1   'colonna codici di raggruppamento nel foglio "CE_bdgt_carica"
    
    'calcolo n° righe piene nel foglio "CE_bdgt_carica"
    n_rig_piene_cebdgt = Worksheets("CE_bdgt_carica").Cells(Rows.Count, col_cr_cebdgt).End(xlUp).Row
    
    'calcolo del numero delle date di analisi
    num_date = 12
    
    'costruzione della matrice "matr_dati_bdgt"
    col_val_cebdgt = 2  'colonna di partenza per lettura valori nel foglio "CE_bdgt_carica"

    ReDim matr_dati_bdgt(UBound(cod_rag, 1), num_date, 2, n_rig_piene_cebdgt) As String

    'Struttura della matrice "matr_dati_bdgt"
    'matr_dati_bdgt[codice raggruppamento, data analisi, codice di conto, valori]
    For i = 1 To UBound(cod_rag, 1)
        cod_i = cod_rag(i, 1)
        For j = 1 To num_date
            For k = 2 To n_rig_piene_cebdgt
                If Worksheets("CE_bdgt_carica").Cells(k, col_cr_cebdgt) = cod_i Then
                    matr_dati_bdgt(i, j, 1, k) = Worksheets("CE_bdgt_carica").Cells(k, col_cc_cebdgt)  'nella colonna 1 è riportato il codice conto
                    matr_dati_bdgt(i, j, 2, k) = Worksheets("CE_bdgt_carica").Cells(k, col_val_cebdgt + j)  'nella colonna 2 è riportato il valore [€]
                End If
            Next k
        Next j
    Next i
  
    'Assegnazione del valore 0 agli elementi vuoti della matrice
    For i = 1 To UBound(matr_dati_bdgt, 1)
            For j = 1 To UBound(matr_dati_bdgt, 2)
                For k = 1 To UBound(matr_dati_bdgt, 4)
                    If matr_dati_bdgt(i, j, 2, k) = "" Then
                        matr_dati_bdgt(i, j, 2, k) = 0
                    End If
                Next k
            Next j
    Next i
    
    'Assegnazione del valore da restituire dalla funzione
    calc_matr_dati_bdgt_CE = matr_dati_bdgt

    

End Function

'La funzione calcola il vettore delle somme dei valori a budget per ogni codice di raggruppamento
Function calc_somme_bdgt(matr_dati_bdgt_CE, cod_rag)

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    'calcolo del numero delle date di analisi
    num_date = 12

    ReDim somme_bdgt(UBound(cod_rag), num_date) As Double

    For i = 1 To UBound(cod_rag)
        For j = 1 To num_date
            For k = 1 To UBound(matr_dati_bdgt_CE, 4)
                somme_bdgt(i, j) = somme_bdgt(i, j) + matr_dati_bdgt_CE(i, j, 2, k)
            Next k
        Next j
    Next i
    
    'Assegnazione del valore da restituire dalla funzione
    calc_somme_bdgt = somme_bdgt

End Function

'La funzione calcola il vettore delle somme dei valori a consuntivo per ogni codice di raggruppamento per l'analisi di periodo PER
Function calc_somme_cons_PER(cod_rag, somme_cons, indice_data_PER)
    'calcolo della matrice "somme_cons_PER" relativa all'analisi di periodo
    ReDim somme_cons_PER(UBound(somme_cons, 1))
    For i = 1 To UBound(somme_cons, 1)
        somme_cons_PER(i) = somme_cons(i, indice_data_PER + 1) - somme_cons(i, indice_data_PER)
    Next i

    pos_RIMP = trova_riga_cdr(cod_rag, "rimp")
    pos_RFMP = trova_riga_cdr(cod_rag, "rfmp")
    If indice_data_PER = 1 Then
        somme_cons_PER(pos_RIMP) = somme_cons(pos_RIMP, indice_data_PER)
        ElseIf indice_data_PER > 1 Then
        somme_cons_PER(pos_RIMP) = somme_cons(pos_RFMP, indice_data_PER - 1)
    End If
    somme_cons_PER(pos_RFMP) = somme_cons(pos_RFMP, indice_data_PER + 1)

    pos_RISEM = trova_riga_cdr(cod_rag, "risem")
    pos_RFSEM = trova_riga_cdr(cod_rag, "rfsem")
    If indice_data_PER = 1 Then
        somme_cons_PER(pos_RISEM) = somme_cons(pos_RISEM, indice_data_PER)
        ElseIf indice_data_PER > 1 Then
        somme_cons_PER(pos_RISEM) = somme_cons(pos_RFSEM, indice_data_PER - 1)
    End If
    somme_cons_PER(pos_RFSEM) = somme_cons(pos_RFSEM, indice_data_PER + 1)

    pos_RIW = trova_riga_cdr(cod_rag, "riw")
    pos_RFW = trova_riga_cdr(cod_rag, "riw")
    If indice_data_PER = 1 Then
        somme_cons_PER(pos_RIW) = somme_cons(pos_RIW, indice_data_PER)
        ElseIf indice_data_PER > 1 Then
        somme_cons_PER(pos_RIW) = somme_cons(pos_RFW, indice_data_PER - 1)
    End If
    somme_cons_PER(pos_RFW) = somme_cons(pos_RFW, indice_data_PER + 1)

    pos_RIPF = trova_riga_cdr(cod_rag, "ripf")
    pos_RFPF = trova_riga_cdr(cod_rag, "rfpf")
    If indice_data_PER = 1 Then
        somme_cons_PER(pos_RIPF) = somme_cons(pos_RIPF, indice_data_PER)
    ElseIf indice_data_PER > 1 Then
        somme_cons_PER(pos_RIPF) = somme_cons(pos_RFPF, indice_data_PER - 1)
    End If
        somme_cons_PER(pos_RFPF) = somme_cons(pos_RFPF, indice_data_PER + 1)

    ' calcolo valore VENDITE
    vendite_cons = somme_cons_PER(trova_riga_cdr(cod_rag, "RI")) + somme_cons_PER(trova_riga_cdr(cod_rag, "RE")) + somme_cons_PER(trova_riga_cdr(cod_rag, "RR")) + somme_cons_PER(trova_riga_cdr(cod_rag, "RS")) - somme_cons_PER(trova_riga_cdr(cod_rag, "resi"))
    somme_cons_PER(trova_riga_cdr(cod_rag, "vendite_cons")) = vendite_cons

    ' calcolo valore VALORE DELLA PRODUZIONE
    valore_prod_cons = vendite_cons + somme_cons_PER(trova_riga_cdr(cod_rag, "capitalizz"))
    somme_cons_PER(trova_riga_cdr(cod_rag, "valore_prod_cons")) = valore_prod_cons

    costo_mp_imp_cons = somme_cons_PER(pos_RIMP) + somme_cons_PER(trova_riga_cdr(cod_rag, "acq")) + somme_cons_PER(trova_riga_cdr(cod_rag, "acqfilos")) + somme_cons_PER(trova_riga_cdr(cod_rag, "trasmp")) + somme_cons_PER(trova_riga_cdr(cod_rag, "mr")) + somme_cons_PER(trova_riga_cdr(cod_rag, "imb")) - somme_cons_PER(pos_RFMP)
    somme_cons_PER(trova_riga_cdr(cod_rag, "costo_mp_imp_cons")) = costo_mp_imp_cons
   
    costo_sl_imp_cons = somme_cons_PER(pos_RISEM) + somme_cons_PER(trova_riga_cdr(cod_rag, "acqsemil")) - somme_cons_PER(pos_RFSEM)
    somme_cons_PER(trova_riga_cdr(cod_rag, "costo_sl_imp_cons")) = costo_sl_imp_cons

    ' calcolo valore COSTO LAVORO DIRETTO
    costo_lav_dir_cons = somme_cons_PER(trova_riga_cdr(cod_rag, "mod")) + somme_cons_PER(trova_riga_cdr(cod_rag, "modtemp"))
    somme_cons_PER(trova_riga_cdr(cod_rag, "costo_lav_dir_cons")) = costo_lav_dir_cons
  
    ' calcolo valore TOT COSTI VARIABILI
    tot_costi_var_cons = costo_mp_imp_cons + costo_sl_imp_cons + costo_lav_dir_cons + somme_cons_PER(trova_riga_cdr(cod_rag, "altricons")) + somme_cons_PER(trova_riga_cdr(cod_rag, "traspf")) + somme_cons_PER(trova_riga_cdr(cod_rag, "ener")) + somme_cons_PER(trova_riga_cdr(cod_rag, "lavest"))
    somme_cons_PER(trova_riga_cdr(cod_rag, "tot_costi_var_cons")) = tot_costi_var_cons
 
    ' calcolo valore MARGINE DI CONTRIBUZIONE
    margine_contr_cons = valore_prod_cons - tot_costi_var_cons
    somme_cons_PER(trova_riga_cdr(cod_rag, "margine_contr_cons")) = margine_contr_cons
  
    ' calcolo valore TOTALE SPESE DI FABBRICA
    tot_spese_fab_cons = somme_cons_PER(trova_riga_cdr(cod_rag, "modin")) + somme_cons_PER(trova_riga_cdr(cod_rag, "modR&S")) + somme_cons_PER(trova_riga_cdr(cod_rag, "amtind")) + somme_cons_PER(trova_riga_cdr(cod_rag, "ass")) + somme_cons_PER(trova_riga_cdr(cod_rag, "man")) + somme_cons_PER(trova_riga_cdr(cod_rag, "altri"))
    somme_cons_PER(trova_riga_cdr(cod_rag, "tot_spese_fab_cons")) = tot_spese_fab_cons
  
    ' calcolo valore TOTALE COSTI DI FABBRICAZIONE
    tot_costi_fab_cons = tot_spese_fab_cons + tot_costi_var_cons
    somme_cons_PER(trova_riga_cdr(cod_rag, "tot_costi_fab_cons")) = tot_costi_fab_cons

    ' calcolo valore COSTO DEI PRODOTTI FABBRICATI
    costo_prod_fab_cons = tot_costi_fab_cons + somme_cons_PER(trova_riga_cdr(cod_rag, "riw")) - somme_cons_PER(trova_riga_cdr(cod_rag, "rfw"))
    somme_cons_PER(trova_riga_cdr(cod_rag, "costo_prod_fab_cons")) = costo_prod_fab_cons

    ' calcolo valore COSTO DEI PRODOTTI VENDUTI
    costo_prod_ven_cons = costo_prod_fab_cons + somme_cons_PER(trova_riga_cdr(cod_rag, "ripf")) - somme_cons_PER(trova_riga_cdr(cod_rag, "rfpf"))
    somme_cons_PER(trova_riga_cdr(cod_rag, "costo_prod_ven_cons")) = costo_prod_ven_cons

    ' calcolo valore UTILE LORDO SULLE VENDITE
    utile_lor_ven_cons = valore_prod_cons - costo_prod_ven_cons
    somme_cons_PER(trova_riga_cdr(cod_rag, "utile_lor_ven_cons")) = utile_lor_ven_cons

    ' calcolo valore TOTALE COSTI COMM.LI
    tot_costi_comm_cons = (somme_cons_PER(trova_riga_cdr(cod_rag, "provv")) + somme_cons_PER(trova_riga_cdr(cod_rag, "vvtt")) + somme_cons_PER(trova_riga_cdr(cod_rag, "stipcom")) + somme_cons_PER(trova_riga_cdr(cod_rag, "asscom")) + somme_cons_PER(trova_riga_cdr(cod_rag, "amtcom")) + somme_cons_PER(trova_riga_cdr(cod_rag, "altrcom")))
    somme_cons_PER(trova_riga_cdr(cod_rag, "tot_costi_comm_cons")) = tot_costi_comm_cons
 
    ' calcolo valore TOTALE COSTI GEN.LI E AMM.VI
    tot_costi_gen_amm_cons = (somme_cons_PER(trova_riga_cdr(cod_rag, "stipamv")) + somme_cons_PER(trova_riga_cdr(cod_rag, "leg")) + somme_cons_PER(trova_riga_cdr(cod_rag, "consamv")) + somme_cons_PER(trova_riga_cdr(cod_rag, "cda")) + somme_cons_PER(trova_riga_cdr(cod_rag, "vvamv")) + somme_cons_PER(trova_riga_cdr(cod_rag, "vvtamv")) + somme_cons_PER(trova_riga_cdr(cod_rag, "amtamv")))
    somme_cons_PER(trova_riga_cdr(cod_rag, "tot_costi_gen_amm_cons")) = tot_costi_gen_amm_cons

    ' calcolo valore TOTALE COSTI OPERATIVI
    tot_costi_op_cons = tot_costi_comm_cons + tot_costi_gen_amm_cons
    somme_cons_PER(trova_riga_cdr(cod_rag, "tot_costi_op_cons")) = tot_costi_op_cons

    ' calcolo valore UTILE OPERATIVO NETTO
    utile_op_netto_cons = utile_lor_ven_cons - tot_costi_op_cons
    somme_cons_PER(trova_riga_cdr(cod_rag, "utile_op_netto_cons")) = utile_op_netto_cons

    ' calcolo valore SALDO GESTIONE FINANZIARIA
    If (-somme_cons_PER(trova_riga_cdr(cod_rag, "onfin")) < 0) Then
        saldo_gest_fin_cons = -(-somme_cons_PER(trova_riga_cdr(cod_rag, "onfin")) + somme_cons_PER(trova_riga_cdr(cod_rag, "serfin")) + somme_cons_PER(trova_riga_cdr(cod_rag, "profin")))
        somme_cons_PER(trova_riga_cdr(cod_rag, "saldo_gest_fin_cons")) = saldo_gest_fin_cons
        saldo_gest_fin_cons = (-somme_cons_PER(trova_riga_cdr(cod_rag, "onfin")) + somme_cons_PER(trova_riga_cdr(cod_rag, "serfin")) + somme_cons_PER(trova_riga_cdr(cod_rag, "profin")))
        somme_cons_PER(trova_riga_cdr(cod_rag, "saldo_gest_fin_cons")) = saldo_gest_fin_cons
    End If

    ' calcolo valore SALDO GESTIONE STRAORDINARIA
    If somme_cons_PER(trova_riga_cdr(cod_rag, "onstr")) > 0 Then
        saldo_gest_str_cons = (somme_cons_PER(trova_riga_cdr(cod_rag, "prostr")) + somme_cons_PER(trova_riga_cdr(cod_rag, "onstr")))
        somme_cons_PER(trova_riga_cdr(cod_rag, "saldo_gest_str_cons")) = saldo_gest_str_cons
        saldo_gest_str_cons = (somme_cons_PER(trova_riga_cdr(cod_rag, "prostr")) - somme_cons_PER(trova_riga_cdr(cod_rag, "onstr")))
        somme_cons_PER(trova_riga_cdr(cod_rag, "saldo_gest_str_cons")) = saldo_gest_str_cons
    End If
  
    ' calcolo valore UTILE PRIMA DELLE IMPOSTE
    utile_pre_imp_cons = utile_op_netto_cons + saldo_gest_fin_cons + saldo_gest_str_cons
    somme_cons_PER(trova_riga_cdr(cod_rag, "utile_pre_imp_cons")) = utile_pre_imp_cons

    ' calcolo valore UTILE NETTO
    utile_netto_cons = utile_pre_imp_cons - somme_cons_PER(trova_riga_cdr(cod_rag, "td"))
    somme_cons_PER(trova_riga_cdr(cod_rag, "utile_netto_cons")) = utile_netto_cons

    calc_somme_cons_PER = somme_cons_PER

End Function


'La funzione calcola il vettore delle somme dei valori a budget per ogni codice di raggruppamento per l'analisi di periodo PER
Function calc_somme_bdgt_PER(somme_bdgt_CE, mese_PER, cod_rag)

    'calcolo della matrice "somme_bdgt_CE_PER" relativa all'analisi di periodo
    ReDim somme_bdgt_CE_PER(UBound(somme_bdgt_CE, 1))
    For i = 1 To UBound(somme_bdgt_CE, 1)
        somme_bdgt_CE_PER(i) = somme_bdgt_CE(i, mese_PER + 1) - somme_bdgt_CE(i, mese_PER)
    Next i

    pos_RIMP = trova_riga_cdr(cod_rag, "rimp")
    pos_RFMP = trova_riga_cdr(cod_rag, "rfmp")
    
    If mese_PER = 1 Then
        somme_bdgt_CE_PER(pos_RIMP) = somme_bdgt_CE(pos_RIMP, mese_PER)
        ElseIf mese_PER > 1 Then
        somme_bdgt_CE_PER(pos_RIMP) = somme_bdgt_CE(pos_RFMP, mese_PER - 1)
    End If
    
    somme_bdgt_CE_PER(pos_RFMP) = somme_bdgt_CE(pos_RFMP, mese_PER + 1)

    pos_RISEM = trova_riga_cdr(cod_rag, "risem")
    pos_RFSEM = trova_riga_cdr(cod_rag, "rfsem")
    If mese_PER = 1 Then
        somme_bdgt_CE_PER(pos_RISEM) = somme_bdgt_CE(pos_RISEM, mese_PER)
        ElseIf mese_PER > 1 Then
        somme_bdgt_CE_PER(pos_RISEM) = somme_bdgt_CE(pos_RFSEM, mese_PER - 1)
    End If
    somme_bdgt_CE_PER(pos_RFSEM) = somme_bdgt_CE(pos_RFSEM, mese_PER + 1)

    pos_RIW = trova_riga_cdr(cod_rag, "riw")
    pos_RFW = trova_riga_cdr(cod_rag, "rfw")
    If mese_PER = 1 Then
        somme_bdgt_CE_PER(pos_RIW) = somme_bdgt_CE(pos_RIW, mese_PER)
        ElseIf mese_PER > 1 Then
        somme_bdgt_CE_PER(pos_RIW) = somme_bdgt_CE(pos_RFW, mese_PER - 1)
    End If
    somme_bdgt_CE_PER(pos_RFW) = somme_bdgt_CE(pos_RFW, mese_PER + 1)

    pos_RIPF = trova_riga_cdr(cod_rag, "ripf")
    pos_RFPF = trova_riga_cdr(cod_rag, "rfpf")
    If mese_PER = 1 Then
        somme_bdgt_CE_PER(pos_RIPF) = somme_bdgt_CE(pos_RIPF, mese_PER)
        ElseIf mese_PER > 1 Then
        somme_bdgt_CE_PER(pos_RIPF) = somme_bdgt_CE(pos_RFPF, mese_PER - 1)
    End If
    somme_bdgt_CE_PER(pos_RFPF) = somme_bdgt_CE(pos_RFPF, mese_PER + 1)

    ' calcolo valore VENDITE
    vendite_bdgt = somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "RI")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "RE")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "RR")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "RS")) - somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "resi"))
    somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "vendite_bdgt")) = vendite_bdgt

    ' calcolo valore VALORE DELLA PRODUZIONE
    valore_prod_bdgt = vendite_bdgt + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "capitalizz"))
    somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "valore_prod_bdgt")) = valore_prod_bdgt

    costo_mp_imp_bdgt = somme_bdgt_CE_PER(pos_RIMP) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "acq")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "acqfilos")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "trasmp")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "mr")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "imb")) - somme_bdgt_CE_PER(pos_RFMP)
    somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "costo_mp_imp_bdgt")) = costo_mp_imp_bdgt
   
    costo_sl_imp_bdgt = somme_bdgt_CE_PER(pos_RISEM) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "acqsemil")) - somme_bdgt_CE_PER(pos_RFSEM)
    somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "costo_sl_imp_bdgt")) = costo_sl_imp_bdgt

    ' calcolo valore COSTO LAVORO DIRETTO
    costo_lav_dir_bdgt = somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "mod")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "modtemp"))
    somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "costo_lav_dir_bdgt")) = costo_lav_dir_bdgt
  
    ' calcolo valore TOT COSTI VARIABILI
    tot_costi_var_bdgt = costo_mp_imp_bdgt + costo_sl_imp_bdgt + costo_lav_dir_bdgt + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "altricons")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "traspf")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "ener")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "lavest"))
    somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "tot_costi_var_bdgt")) = tot_costi_var_bdgt
 
    ' calcolo valore MARGINE DI CONTRIBUZIONE
    margine_contr_bdgt = valore_prod_bdgt - tot_costi_var_bdgt
    somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "margine_contr_bdgt")) = margine_contr_bdgt
  
    ' calcolo valore TOTALE SPESE DI FABBRICA
    tot_spese_fab_bdgt = somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "modin")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "modR&S")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "amtind")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "ass")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "man")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "altri"))
    somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "tot_spese_fab_bdgt")) = tot_spese_fab_bdgt
  
    ' calcolo valore TOTALE COSTI DI FABBRICAZIONE
    tot_costi_fab_bdgt = tot_spese_fab_bdgt + tot_costi_var_bdgt
    somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "tot_costi_fab_bdgt")) = tot_costi_fab_bdgt

    ' calcolo valore COSTO DEI PRODOTTI FABBRICATI
    costo_prod_fab_bdgt = tot_costi_fab_bdgt + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "riw")) - somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "rfw"))
    somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "costo_prod_fab_bdgt")) = costo_prod_fab_bdgt

    ' calcolo valore COSTO DEI PRODOTTI VENDUTI
    costo_prod_ven_bdgt = costo_prod_fab_bdgt + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "ripf")) - somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "rfpf"))
    somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "costo_prod_ven_bdgt")) = costo_prod_ven_bdgt

    ' calcolo valore UTILE LORDO SULLE VENDITE
    utile_lor_ven_bdgt = valore_prod_bdgt - costo_prod_ven_bdgt
    somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "utile_lor_ven_bdgt")) = utile_lor_ven_bdgt

    ' calcolo valore TOTALE COSTI COMM.LI
    tot_costi_comm_bdgt = (somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "provv")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "vvtt")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "stipcom")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "asscom")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "amtcom")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "altrcom")))
    somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "tot_costi_comm_bdgt")) = tot_costi_comm_bdgt
 
    ' calcolo valore TOTALE COSTI GEN.LI E AMM.VI
    tot_costi_gen_amm_bdgt = (somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "stipamv")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "leg")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "consamv")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "cda")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "vvamv")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "vvtamv")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "amtamv")))
    somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "tot_costi_gen_amm_bdgt")) = tot_costi_gen_amm_bdgt

    ' calcolo valore TOTALE COSTI OPERATIVI
    tot_costi_op_bdgt = tot_costi_comm_bdgt + tot_costi_gen_amm_bdgt
    somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "tot_costi_op_bdgt")) = tot_costi_op_bdgt

    ' calcolo valore UTILE OPERATIVO NETTO
    utile_op_netto_bdgt = utile_lor_ven_bdgt - tot_costi_op_bdgt
    somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "utile_op_netto_bdgt")) = utile_op_netto_bdgt

    ' calcolo valore SALDO GESTIONE FINANZIARIA
    If (-somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "onfin")) < 0) Then
        saldo_gest_fin_bdgt = -(-somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "onfin")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "serfin")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "profin")))
        somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "saldo_gest_fin_bdgt")) = saldo_gest_fin_bdgt
        saldo_gest_fin_bdgt = (-somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "onfin")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "serfin")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "profin")))
        somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "saldo_gest_fin_bdgt")) = saldo_gest_fin_bdgt
    End If

    ' calcolo valore SALDO GESTIONE STRAORDINARIA
    If somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "onstr")) > 0 Then
        saldo_gest_str_bdgt = (somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "prostr")) + somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "onstr")))
        somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "saldo_gest_str_bdgt")) = saldo_gest_str_bdgt
        saldo_gest_str_bdgt = (somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "prostr")) - somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "onstr")))
        somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "saldo_gest_str_bdgt")) = saldo_gest_str_bdgt
    End If
  
    ' calcolo valore UTILE PRIMA DELLE IMPOSTE
    utile_pre_imp_bdgt = utile_op_netto_bdgt + saldo_gest_fin_bdgt + saldo_gest_str_bdgt
    somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "utile_pre_imp_bdgt")) = utile_pre_imp_bdgt

    ' calcolo valore UTILE NETTO
    utile_netto_bdgt = utile_pre_imp_bdgt - somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "td"))
    somme_bdgt_CE_PER(trova_riga_cdr(cod_rag, "utile_netto_bdgt")) = utile_netto_bdgt

    calc_somme_bdgt_PER = somme_bdgt_CE_PER

End Function

' Funzione trova riga elemento in matrice
Function trova_riga_cdr(arr, val) As Integer
    Dim r As Integer, c As Integer
    For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            If arr(r, c) = val Then
                trova_riga_cdr = r
                Exit Function
            End If
        Next c
    Next r
End Function
