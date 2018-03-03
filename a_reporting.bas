Attribute VB_Name = "a_reporting"
Sub reporting()

' la macro crea la tabella di analisi del conto economico nel foglio "CE_tab"
'--------------------------------------------------------------------------------------------------------------------------

'pulizia del foglio "CE_tab"
Worksheets("CE_tab").Cells.Clear

'pulizia del foglio "SP_tab"
Worksheets("SP_tab").Cells.Clear

'1) lettura delle date dei dati a consuntivo disponibili
date_cons = costr_vett_date()

'2) selezione data analisi
data_an_YTD = lettura_data_an_YTD(date_cons)
data_an_PER = lettura_data_an_PER(date_cons)

' lettura del mese analisi
mese_YTD = Month(data_an_YTD)
mese_PER = Month(data_an_PER)

' determinazione del numero di ordine della data PER nel vettore "date_cons"
indice_data_PER = trova_ind_data_PER(date_cons, data_an_PER)

'3) costruzione matrice codici di raggruppamento cod_rag(codice raggruppamento, descrizione, segno)
cod_rag = calc_cod_rag()

'4) costruzione matrice matr_dati_cons(codici di raggruppamento, codici di conto, valori) consuntivo
matr_dati_cons = calc_matr_dati_cons(date_cons, cod_rag)
       
'5) costruzione del vettore delle somme dei valori per ogni codice di raggruppamento
somme_cons = calc_somme_cons(matr_dati_cons, cod_rag, date_cons)

'5b) costruzione del vettore delle somme dei valori per ogni codice di raggruppamento analisi PER
somme_cons_PER = calc_somme_cons_PER(cod_rag, somme_cons, indice_data_PER)

'6) costruzione matrice matr_dati_bdgt(codici di raggruppamento, codici di conto, valori) budget CE
matr_dati_bdgt_CE = calc_matr_dati_bdgt_CE(cod_rag)

'7) costruzione matrice matr_dati_bdgt(codici di raggruppamento, codici di conto, valori) budget SP
matr_dati_bdgt_SP = calc_matr_dati_bdgt_SP(cod_rag)
 
'8) costruzione vettore somme_bdgt_CE budget
somme_bdgt_CE = calc_somme_bdgt(matr_dati_bdgt_CE, cod_rag)

'8b) costruzione vettore somme_bdgt_CE_PER budget per analisi PER
somme_bdgt_CE_PER = calc_somme_bdgt_PER(somme_bdgt_CE, mese_PER, cod_rag)

'9) costruzione vettore somme_bdgt_SP budget
somme_bdgt_SP = calc_somme_bdgt_SP(matr_dati_bdgt_SP, cod_rag)

'9b) costruzione vettore somme_bdgt_SP budget
somme_bdgt_SP_PER = calc_somme_bdgt_SP_PER(somme_bdgt_SP, mese_PER)
    
'10) calcolo somme parziali
Call somme_parziali_CE(somme_cons, somme_bdgt_CE, cod_rag)
Call somme_parziali_SP(somme_cons, somme_bdgt_SP, cod_rag)

'11) creazione tabella CE analisi YTD
Call tabella_YTD_CE(cod_rag, somme_cons, somme_bdgt_CE, data_an_YTD, mese_YTD)

'12) creazione tabella CE analisi PER
Call tabella_PER_CE(cod_rag, somme_cons, somme_cons_PER, somme_bdgt_CE, somme_bdgt_CE_PER, date_cons, indice_data_PER)

'13) creazione tabella SP analisi YTD
Call tabella_YTD_SP(cod_rag, somme_cons, somme_bdgt_SP, data_an_YTD, mese_YTD)

'14) creazione tabella SP analisi PER
Call tabella_PER_SP(cod_rag, somme_cons_PER, somme_bdgt_SP_PER, date_cons, indice_data_PER)

End Sub
