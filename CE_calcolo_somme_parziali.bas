Attribute VB_Name = "CE_calcolo_somme_parziali"
' le variabili vendite_cons e vendite_bdg vengono dichiarate pubbliche per essere usate nel calcolo delle colonne %
' del foglio 'CE_tab'
Public vendite_cons As Double
Public vendite_bdgt As Double

Sub somme_parziali_CE(somme_cons, somme_bdgt_CE, cod_rag)

'0) dichiarazione variabili

'Dim vendite_cons As Double
Dim valore_prod_cons As Double
Dim costo_mp_imp_cons As Double
Dim costo_sl_imp_cons As Double
Dim costo_lav_dir_cons As Double
Dim tot_costi_var_cons As Double
Dim margine_contr_cons As Double
Dim tot_spese_fab_cons As Double
Dim tot_costi_fab_cons As Double
Dim costo_prod_fab_cons As Double
Dim costo_prod_ven_cons As Double
Dim utile_lor_ven_cons As Double
Dim tot_costi_comm_cons As Double
Dim tot_costi_gen_amm_cons As Double
Dim tot_costi_op_cons As Double
Dim utile_op_netto_cons As Double
Dim saldo_gest_fin_cons As Double
Dim saldo_gest_str_cons As Double
Dim utile_pre_imp_cons As Double
Dim utile_netto_cons As Double

'Dim vendite_bdgt As Double
Dim valore_prod_bdgt As Double
Dim costo_mp_imp_bdgt As Double
Dim costo_sl_imp_bdgt As Double
Dim costo_lav_dir_bdgt As Double
Dim tot_costi_var_bdgt As Double
Dim margine_contr_bdgt As Double
Dim tot_spese_fab_bdgt As Double
Dim tot_costi_fab_bdgt As Double
Dim costo_prod_fab_bdgt As Double
Dim costo_prod_ven_bdgt As Double
Dim utile_lor_ven_bdgt As Double
Dim tot_costi_comm_bdgt As Double
Dim tot_costi_gen_amm_bdgt As Double
Dim tot_costi_op_bdgt As Double
Dim utile_op_netto_bdgt As Double
Dim saldo_gest_fin_bdgt As Double
Dim saldo_gest_str_bdgt As Double
Dim utile_pre_imp_bdgt As Double
Dim utile_netto_bdgt As Double

'1) calcolo dei termini mancanti della matrice somme_cons ottenuti mediante somme dei termini già calcolati
For i = 1 To UBound(somme_cons, 2)
    ' calcolo valore VENDITE
    vendite_cons = somme_cons(trova_riga_cdr(cod_rag, "RI"), i) + somme_cons(trova_riga_cdr(cod_rag, "RE"), i) + somme_cons(trova_riga_cdr(cod_rag, "RR"), i) + somme_cons(trova_riga_cdr(cod_rag, "RS"), i) - somme_cons(trova_riga_cdr(cod_rag, "resi"), i)
    somme_cons(trova_riga_cdr(cod_rag, "vendite_cons"), i) = vendite_cons

    ' calcolo valore VALORE DELLA PRODUZIONE
    valore_prod_cons = vendite_cons + somme_cons(trova_riga_cdr(cod_rag, "capitalizz"), i)
    somme_cons(trova_riga_cdr(cod_rag, "valore_prod_cons"), i) = valore_prod_cons

    ' calcolo valore COSTO MATERIE PRIME IMPIEGATE
    costo_mp_imp_cons = somme_cons(trova_riga_cdr(cod_rag, "rimp"), i) + somme_cons(trova_riga_cdr(cod_rag, "acq"), i) + somme_cons(trova_riga_cdr(cod_rag, "acqfilos"), i) + somme_cons(trova_riga_cdr(cod_rag, "trasmp"), i) + somme_cons(trova_riga_cdr(cod_rag, "mr"), i) + somme_cons(trova_riga_cdr(cod_rag, "imb"), i) - somme_cons(trova_riga_cdr(cod_rag, "rfmp"), i)
    somme_cons(trova_riga_cdr(cod_rag, "costo_mp_imp_cons"), i) = costo_mp_imp_cons

    ' calcolo valore COSTO SEMILAVORATI IMPIEGATI
    costo_sl_imp_cons = somme_cons(trova_riga_cdr(cod_rag, "risem"), i) + somme_cons(trova_riga_cdr(cod_rag, "acqsemil"), i) - somme_cons(trova_riga_cdr(cod_rag, "rfsem"), i)
    somme_cons(trova_riga_cdr(cod_rag, "costo_sl_imp_cons"), i) = costo_sl_imp_cons

    ' calcolo valore COSTO LAVORO DIRETTO
    costo_lav_dir_cons = somme_cons(trova_riga_cdr(cod_rag, "mod"), i) + somme_cons(trova_riga_cdr(cod_rag, "modtemp"), i)
    somme_cons(trova_riga_cdr(cod_rag, "costo_lav_dir_cons"), i) = costo_lav_dir_cons

    ' calcolo valore TOT COSTI VARIABILI
    tot_costi_var_cons = costo_mp_imp_cons + costo_sl_imp_cons + costo_lav_dir_cons + somme_cons(trova_riga_cdr(cod_rag, "altricons"), i) + somme_cons(trova_riga_cdr(cod_rag, "traspf"), i) + somme_cons(trova_riga_cdr(cod_rag, "ener"), i) + somme_cons(trova_riga_cdr(cod_rag, "lavest"), i)
    somme_cons(trova_riga_cdr(cod_rag, "tot_costi_var_cons"), i) = tot_costi_var_cons

    ' calcolo valore MARGINE DI CONTRIBUZIONE
    margine_contr_cons = valore_prod_cons - tot_costi_var_cons
    somme_cons(trova_riga_cdr(cod_rag, "margine_contr_cons"), i) = margine_contr_cons

    ' calcolo valore TOTALE SPESE DI FABBRICA
    tot_spese_fab_cons = somme_cons(trova_riga_cdr(cod_rag, "modin"), i) + somme_cons(trova_riga_cdr(cod_rag, "modR&S"), i) + somme_cons(trova_riga_cdr(cod_rag, "amtind"), i) + somme_cons(trova_riga_cdr(cod_rag, "ass"), i) + somme_cons(trova_riga_cdr(cod_rag, "man"), i) + somme_cons(trova_riga_cdr(cod_rag, "altri"), i)
    somme_cons(trova_riga_cdr(cod_rag, "tot_spese_fab_cons"), i) = tot_spese_fab_cons

    ' calcolo valore TOTALE COSTI DI FABBRICAZIONE
    tot_costi_fab_cons = tot_spese_fab_cons + tot_costi_var_cons
    somme_cons(trova_riga_cdr(cod_rag, "tot_costi_fab_cons"), i) = tot_costi_fab_cons

    ' calcolo valore COSTO DEI PRODOTTI FABBRICATI
    costo_prod_fab_cons = tot_costi_fab_cons + somme_cons(trova_riga_cdr(cod_rag, "riw"), i) - somme_cons(trova_riga_cdr(cod_rag, "rfw"), i)
    somme_cons(trova_riga_cdr(cod_rag, "costo_prod_fab_cons"), i) = costo_prod_fab_cons

    ' calcolo valore COSTO DEI PRODOTTI VENDUTI
    costo_prod_ven_cons = costo_prod_fab_cons + somme_cons(trova_riga_cdr(cod_rag, "ripf"), i) - somme_cons(trova_riga_cdr(cod_rag, "rfpf"), i)
    somme_cons(trova_riga_cdr(cod_rag, "costo_prod_ven_cons"), i) = costo_prod_ven_cons

    ' calcolo valore UTILE LORDO SULLE VENDITE
    utile_lor_ven_cons = valore_prod_cons - costo_prod_ven_cons
    somme_cons(trova_riga_cdr(cod_rag, "utile_lor_ven_cons"), i) = utile_lor_ven_cons

    ' calcolo valore TOTALE COSTI COMM.LI
    tot_costi_comm_cons = (somme_cons(trova_riga_cdr(cod_rag, "provv"), i) + somme_cons(trova_riga_cdr(cod_rag, "vvtt"), i) + somme_cons(trova_riga_cdr(cod_rag, "stipcom"), i) + somme_cons(trova_riga_cdr(cod_rag, "asscom"), i) + somme_cons(trova_riga_cdr(cod_rag, "amtcom"), i) + somme_cons(trova_riga_cdr(cod_rag, "altrcom"), i))
    somme_cons(trova_riga_cdr(cod_rag, "tot_costi_comm_cons"), i) = tot_costi_comm_cons

    ' calcolo valore TOTALE COSTI GEN.LI E AMM.VI
    tot_costi_gen_amm_cons = (somme_cons(trova_riga_cdr(cod_rag, "stipamv"), i) + somme_cons(trova_riga_cdr(cod_rag, "leg"), i) + somme_cons(trova_riga_cdr(cod_rag, "consamv"), i) + somme_cons(trova_riga_cdr(cod_rag, "cda"), i) + somme_cons(trova_riga_cdr(cod_rag, "vvamv"), i) + somme_cons(trova_riga_cdr(cod_rag, "vvtamv"), i) + somme_cons(trova_riga_cdr(cod_rag, "amtamv"), i))
    somme_cons(trova_riga_cdr(cod_rag, "tot_costi_gen_amm_cons"), i) = tot_costi_gen_amm_cons

    ' calcolo valore TOTALE COSTI OPERATIVI
    tot_costi_op_cons = tot_costi_comm_cons + tot_costi_gen_amm_cons
    somme_cons(trova_riga_cdr(cod_rag, "tot_costi_op_cons"), i) = tot_costi_op_cons

    ' calcolo valore UTILE OPERATIVO NETTO
    utile_op_netto_cons = utile_lor_ven_cons - tot_costi_op_cons
    somme_cons(trova_riga_cdr(cod_rag, "utile_op_netto_cons"), i) = utile_op_netto_cons

    ' calcolo valore SALDO GESTIONE FINANZIARIA
    saldo_gest_fin_cons = (-somme_cons(trova_riga_cdr(cod_rag, "onfin"), i) + somme_cons(trova_riga_cdr(cod_rag, "serfin"), i) + somme_cons(trova_riga_cdr(cod_rag, "profin"), i))
    somme_cons(trova_riga_cdr(cod_rag, "saldo_gest_fin_cons"), i) = saldo_gest_fin_cons

    ' calcolo valore SALDO GESTIONE STRAORDINARIA
    saldo_gest_str_cons = (somme_cons(trova_riga_cdr(cod_rag, "prostr"), i) - somme_cons(trova_riga_cdr(cod_rag, "onstr"), i))
    somme_cons(trova_riga_cdr(cod_rag, "saldo_gest_str_cons"), i) = saldo_gest_str_cons

    ' calcolo valore UTILE PRIMA DELLE IMPOSTE
    utile_pre_imp_cons = utile_op_netto_cons + saldo_gest_fin_cons + saldo_gest_str_cons
    somme_cons(trova_riga_cdr(cod_rag, "utile_pre_imp_cons"), i) = utile_pre_imp_cons

    ' calcolo valore UTILE NETTO
    utile_netto_cons = utile_pre_imp_cons - somme_cons(trova_riga_cdr(cod_rag, "td"), i)
    somme_cons(trova_riga_cdr(cod_rag, "utile_netto_cons"), i) = utile_netto_cons
    
Next i

'*****************************************************************************************
'2) calcolo dei termini mancanti della matrice somme_bdgt ottenuti mediante somme dei termini già calcolati
For i = 1 To 12   ' n. mesi per il quale è stato realizzato il budget
    ' calcolo valore VENDITE
    vendite_bdgt = somme_bdgt_CE(trova_riga_cdr(cod_rag, "RI"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "RE"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "RR"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "RS"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "resi"), i)
    somme_bdgt_CE(trova_riga_cdr(cod_rag, "vendite_bdgt"), i) = vendite_bdgt

    ' calcolo valore VALORE DELLA PRODUZIONE
    valore_prod_bdgt = vendite_bdgt + somme_bdgt_CE(trova_riga_cdr(cod_rag, "capitalizz"), i)
    somme_bdgt_CE(trova_riga_cdr(cod_rag, "valore_prod_bdgt"), i) = valore_prod_bdgt

    ' calcolo valore COSTO MATERIE PRIME IMPIEGATE
    costo_mp_imp_bdgt = somme_bdgt_CE(trova_riga_cdr(cod_rag, "rimp"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "acq"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "acqfilos"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "trasmp"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "mr"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "imb"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "rfmp"), i)
    somme_bdgt_CE(trova_riga_cdr(cod_rag, "costo_mp_imp_bdgt"), i) = costo_mp_imp_bdgt

    ' calcolo valore COSTO SEMILAVORATI IMPIEGATI
    costo_sl_imp_bdgt = somme_bdgt_CE(trova_riga_cdr(cod_rag, "risem"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "acqsemil"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "rfsem"), i)
    somme_bdgt_CE(trova_riga_cdr(cod_rag, "costo_sl_imp_bdgt"), i) = costo_sl_imp_bdgt

    ' calcolo valore COSTO LAVORO DIRETTO
    costo_lav_dir_bdgt = somme_bdgt_CE(trova_riga_cdr(cod_rag, "mod"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "modtemp"), i)
    somme_bdgt_CE(trova_riga_cdr(cod_rag, "costo_lav_dir_bdgt"), i) = costo_lav_dir_bdgt

    ' calcolo valore TOT COSTI VARIABILI
    tot_costi_var_bdgt = costo_mp_imp_bdgt + costo_sl_imp_bdgt + costo_lav_dir_bdgt + somme_bdgt_CE(trova_riga_cdr(cod_rag, "altricons"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "traspf"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "ener"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "lavest"), i)
    somme_bdgt_CE(trova_riga_cdr(cod_rag, "tot_costi_var_bdgt"), i) = tot_costi_var_bdgt

    ' calcolo valore MARGINE DI CONTRIBUZIONE
    margine_contr_bdgt = valore_prod_bdgt - tot_costi_var_bdgt
    somme_bdgt_CE(trova_riga_cdr(cod_rag, "margine_contr_bdgt"), i) = margine_contr_bdgt

    ' calcolo valore TOTALE SPESE DI FABBRICA
    tot_spese_fab_bdgt = somme_bdgt_CE(trova_riga_cdr(cod_rag, "modin"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "modR&S"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "amtind"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "ass"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "man"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "altri"), i)
    somme_bdgt_CE(trova_riga_cdr(cod_rag, "tot_spese_fab_bdgt"), i) = tot_spese_fab_bdgt

    ' calcolo valore TOTALE COSTI DI FABBRICAZIONE
    tot_costi_fab_bdgt = tot_spese_fab_bdgt + tot_costi_var_bdgt
    somme_bdgt_CE(trova_riga_cdr(cod_rag, "tot_costi_fab_bdgt"), i) = tot_costi_fab_bdgt

    ' calcolo valore COSTO DEI PRODOTTI FABBRICATI
    costo_prod_fab_bdgt = tot_costi_fab_bdgt + somme_bdgt_CE(trova_riga_cdr(cod_rag, "riw"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "rfw"), i)
    somme_bdgt_CE(trova_riga_cdr(cod_rag, "costo_prod_fab_bdgt"), i) = costo_prod_fab_bdgt

    ' calcolo valore COSTO DEI PRODOTTI VENDUTI
    costo_prod_ven_bdgt = costo_prod_fab_bdgt + somme_bdgt_CE(trova_riga_cdr(cod_rag, "ripf"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "rfpf"), i)
    somme_bdgt_CE(trova_riga_cdr(cod_rag, "costo_prod_ven_bdgt"), i) = costo_prod_ven_bdgt

    ' calcolo valore UTILE LORDO SULLE VENDITE
    utile_lor_ven_bdgt = valore_prod_bdgt - costo_prod_ven_bdgt
    somme_bdgt_CE(trova_riga_cdr(cod_rag, "utile_lor_ven_bdgt"), i) = utile_lor_ven_bdgt

    ' calcolo valore TOTALE COSTI COMM.LI
    tot_costi_comm_bdgt = (somme_bdgt_CE(trova_riga_cdr(cod_rag, "provv"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "vvtt"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "stipcom"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "asscom"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "amtcom"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "altrcom"), i))
    somme_bdgt_CE(trova_riga_cdr(cod_rag, "tot_costi_comm_bdgt"), i) = tot_costi_comm_bdgt

    ' calcolo valore TOTALE COSTI GEN.LI E AMM.VI
    tot_costi_gen_amm_bdgt = (somme_bdgt_CE(trova_riga_cdr(cod_rag, "stipamv"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "leg"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "consamv"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "cda"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "vvamv"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "vvtamv"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "amtamv"), i))
    somme_bdgt_CE(trova_riga_cdr(cod_rag, "tot_costi_gen_amm_bdgt"), i) = tot_costi_gen_amm_bdgt

    ' calcolo valore TOTALE COSTI OPERATIVI
    tot_costi_op_bdgt = tot_costi_comm_bdgt + tot_costi_gen_amm_bdgt
    somme_bdgt_CE(trova_riga_cdr(cod_rag, "tot_costi_op_bdgt"), i) = tot_costi_op_bdgt

    ' calcolo valore UTILE OPERATIVO NETTO
    utile_op_netto_bdgt = utile_lor_ven_bdgt - tot_costi_op_bdgt
    somme_bdgt_CE(trova_riga_cdr(cod_rag, "utile_op_netto_bdgt"), i) = utile_op_netto_bdgt

    ' calcolo valore SALDO GESTIONE FINANZIARIA
    saldo_gest_fin_bdgt = (-somme_bdgt_CE(trova_riga_cdr(cod_rag, "onfin"), i) - somme_bdgt_CE(trova_riga_cdr(cod_rag, "serfin"), i) + somme_bdgt_CE(trova_riga_cdr(cod_rag, "profin"), i))
    somme_bdgt_CE(trova_riga_cdr(cod_rag, "saldo_gest_fin_bdgt"), i) = saldo_gest_fin_bdgt

    ' calcolo valore SALDO GESTIONE STRAORDINARIA
    saldo_gest_str_bdgt = (somme_bdgt_CE(trova_riga_cdr(cod_rag, "prostr"), i) - somme_bdgt_CE(trova_riga_cdr(cod_rag, "onstr"), i))
    somme_bdgt_CE(trova_riga_cdr(cod_rag, "saldo_gest_str_bdgt"), i) = saldo_gest_str_bdgt

    ' calcolo valore UTILE PRIMA DELLE IMPOSTE
    utile_pre_imp_bdgt = utile_op_netto_bdgt + saldo_gest_fin_bdgt + saldo_gest_str_bdgt
    somme_bdgt_CE(trova_riga_cdr(cod_rag, "utile_pre_imp_bdgt"), i) = utile_pre_imp_bdgt

    ' calcolo valore UTILE NETTO
    utile_netto_bdgt = utile_pre_imp_bdgt - somme_bdgt_CE(trova_riga_cdr(cod_rag, "td"), i)
    somme_bdgt_CE(trova_riga_cdr(cod_rag, "utile_netto_bdgt"), i) = utile_netto_bdgt

Next i

End Sub

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
