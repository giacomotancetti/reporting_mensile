Attribute VB_Name = "SP_calcolo_somme_parziali"
Sub somme_parziali_SP(somme_cons, somme_bdgt_SP, cod_rag)

'0) dichiarazione variabili

'Dim vendite_cons As Double
Dim imm_tecn_cons As Double

'1) compilazione del foglio "SP" colonna ACTUAL somme parziali
For i = 1 To UBound(somme_cons, 2)

    ' calcolo valore IMMOBILIZZAZIONI TECNICHE
    imm_tecn_cons = somme_cons(trova_riga_cdr(cod_rag, "im"), i) + somme_cons(trova_riga_cdr(cod_rag, "imp"), i) + somme_cons(trova_riga_cdr(cod_rag, "attr"), i) + somme_cons(trova_riga_cdr(cod_rag, "ser"), i) + somme_cons(trova_riga_cdr(cod_rag, "rete"), i) + somme_cons(trova_riga_cdr(cod_rag, "pc"), i) + somme_cons(trova_riga_cdr(cod_rag, "m&a"), i) + somme_cons(trova_riga_cdr(cod_rag, "auto"), i) + somme_cons(trova_riga_cdr(cod_rag, "fam"), i)
    somme_cons(trova_riga_cdr(cod_rag, "imm_tecn_cons"), i) = imm_tecn_cons

    ' calcolo valore IMMOBILIZZAZIONI IMMATERIALI
    imm_imm_cons = somme_cons(trova_riga_cdr(cod_rag, "pimp"), i) + somme_cons(trova_riga_cdr(cod_rag, "R&S"), i) + somme_cons(trova_riga_cdr(cod_rag, "soft"), i) + somme_cons(trova_riga_cdr(cod_rag, "L&C"), i) + somme_cons(trova_riga_cdr(cod_rag, "aaimm"), i)
    somme_cons(trova_riga_cdr(cod_rag, "imm_imm_cons"), i) = imm_imm_cons

    ' calcolo valore IMMOBILIZZAZIONI FINANZIARIE
    imm_fin_cons = somme_cons(trova_riga_cdr(cod_rag, "P&T"), i) + somme_cons(trova_riga_cdr(cod_rag, "Cred"), i)
    somme_cons(trova_riga_cdr(cod_rag, "imm_fin_cons"), i) = imm_fin_cons

    ' calcolo valore ATTIVO IMMOBILIZZATO
    att_imm_cons = somme_cons(trova_riga_cdr(cod_rag, "imm_tecn_cons"), i) + somme_cons(trova_riga_cdr(cod_rag, "imm_imm_cons"), i) + somme_cons(trova_riga_cdr(cod_rag, "imm_fin_cons"), i)
    somme_cons(trova_riga_cdr(cod_rag, "att_imm_cons"), i) = att_imm_cons

    ' calcolo valore ATTIVO IMMOBILIZZATO
    att_imm_cons = somme_cons(trova_riga_cdr(cod_rag, "imm_tecn_cons"), i) + somme_cons(trova_riga_cdr(cod_rag, "imm_imm_cons"), i) + somme_cons(trova_riga_cdr(cod_rag, "imm_fin_cons"), i)
    somme_cons(trova_riga_cdr(cod_rag, "att_imm_cons"), i) = att_imm_cons

    ' calcolo valore GIACENZE MAGAZZINO
    giac_mag_cons = somme_cons(trova_riga_cdr(cod_rag, "rpf"), i) + somme_cons(trova_riga_cdr(cod_rag, "rsem"), i) + somme_cons(trova_riga_cdr(cod_rag, "rimb"), i) + somme_cons(trova_riga_cdr(cod_rag, "rmp"), i)
    somme_cons(trova_riga_cdr(cod_rag, "giac_mag_cons"), i) = giac_mag_cons

    ' calcolo valore CREDITI
    crediti_cons = somme_cons(trova_riga_cdr(cod_rag, "credvscl"), i) + somme_cons(trova_riga_cdr(cod_rag, "FSC"), i) + somme_cons(trova_riga_cdr(cod_rag, "rratt"), i) + somme_cons(trova_riga_cdr(cod_rag, "aacc"), i)
    somme_cons(trova_riga_cdr(cod_rag, "crediti_cons"), i) = crediti_cons

    ' calcolo valore ATTIVITA' CORRENTI
    att_corr_cons = somme_cons(trova_riga_cdr(cod_rag, "giac_mag_cons"), i) + somme_cons(trova_riga_cdr(cod_rag, "crediti_cons"), i)
    somme_cons(trova_riga_cdr(cod_rag, "att_corr_cons"), i) = att_corr_cons

    ' calcolo valore PASSIVITA' CORRENTI
    pass_corr_cons = somme_cons(trova_riga_cdr(cod_rag, "forn"), i) + somme_cons(trova_riga_cdr(cod_rag, "dip"), i) + somme_cons(trova_riga_cdr(cod_rag, "eriva"), i) + somme_cons(trova_riga_cdr(cod_rag, "fimp"), i) + somme_cons(trova_riga_cdr(cod_rag, "iipp"), i) + somme_cons(trova_riga_cdr(cod_rag, "impdd"), i) + somme_cons(trova_riga_cdr(cod_rag, "debErateiz"), i) + somme_cons(trova_riga_cdr(cod_rag, "aadeb"), i) + somme_cons(trova_riga_cdr(cod_rag, "rrpass"), i)
    somme_cons(trova_riga_cdr(cod_rag, "pass_corr_cons"), i) = pass_corr_cons

    ' calcolo valore CAPITALE CIRCOLANTE OPERATIVO
    cap_circ_op_cons = somme_cons(trova_riga_cdr(cod_rag, "att_corr_cons"), i) + somme_cons(trova_riga_cdr(cod_rag, "pass_corr_cons"), i)
    somme_cons(trova_riga_cdr(cod_rag, "cap_circ_op_cons"), i) = cap_circ_op_cons

    ' calcolo valore LIQUIDITA'
    liq_cons = somme_cons(trova_riga_cdr(cod_rag, "cash"), i)
    somme_cons(trova_riga_cdr(cod_rag, "liq_cons"), i) = liq_cons

    ' calcolo valore CAPITALE INVESTITO
    cap_inv_cons = somme_cons(trova_riga_cdr(cod_rag, "att_imm_cons"), i) + somme_cons(trova_riga_cdr(cod_rag, "cap_circ_op_cons"), i) + somme_cons(trova_riga_cdr(cod_rag, "liq_cons"), i)
    somme_cons(trova_riga_cdr(cod_rag, "cap_inv_cons"), i) = cap_inv_cons

    ' calcolo valore PATRIMONIO NETTO
    patr_netto_cons = somme_cons(trova_riga_cdr(cod_rag, "CS"), i) + somme_cons(trova_riga_cdr(cod_rag, "ris"), i) + somme_cons(trova_riga_cdr(cod_rag, "ut"), i)
    somme_cons(trova_riga_cdr(cod_rag, "patr_netto_cons"), i) = patr_netto_cons

    ' calcolo valore PASSIVITA' A MEDIO/LUNGO
    pass_ml_cons = somme_cons(trova_riga_cdr(cod_rag, "deblt"), i) + somme_cons(trova_riga_cdr(cod_rag, "TFR"), i) + somme_cons(trova_riga_cdr(cod_rag, "aaff"), i)
    somme_cons(trova_riga_cdr(cod_rag, "pass_ml_cons"), i) = pass_ml_cons

    ' calcolo valore PASSIVITA' FINANZ A BREVE
    pass_b_cons = somme_cons(trova_riga_cdr(cod_rag, "debbt"), i)
    somme_cons(trova_riga_cdr(cod_rag, "pass_b_cons"), i) = pass_b_cons

    ' calcolo valore FONTI DI FINANZIAMENTO
    fonti_fin_cons = somme_cons(trova_riga_cdr(cod_rag, "patr_netto_cons"), i) + somme_cons(trova_riga_cdr(cod_rag, "pass_ml_cons"), i) + somme_cons(trova_riga_cdr(cod_rag, "pass_b_cons"), i)
    somme_cons(trova_riga_cdr(cod_rag, "fonti_fin_cons"), i) = fonti_fin_cons
    
Next i

'2) compilazione del foglio "SP" colonna BUDGET somme parziali
For i = 1 To 12

    ' calcolo valore IMMOBILIZZAZIONI TECNICHE
    imm_tecn_bdgt = somme_bdgt_SP(trova_riga_cdr(cod_rag, "im"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "imp"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "attr"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "ser"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "rete"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "pc"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "m&a"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "auto"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "fam"), i)
    somme_bdgt_SP(trova_riga_cdr(cod_rag, "imm_tecn_bdgt"), i) = imm_tecn_bdgt

    ' calcolo valore IMMOBILIZZAZIONI IMMATERIALI
    imm_imm_bdgt = somme_bdgt_SP(trova_riga_cdr(cod_rag, "pimp"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "R&S"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "soft"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "L&C"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "aaimm"), i)
    somme_bdgt_SP(trova_riga_cdr(cod_rag, "imm_imm_bdgt"), i) = imm_imm_bdgt

    ' calcolo valore IMMOBILIZZAZIONI FINANZIARIE
    imm_fin_bdgt = somme_bdgt_SP(trova_riga_cdr(cod_rag, "P&T"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "Cred"), i)
    somme_bdgt_SP(trova_riga_cdr(cod_rag, "imm_fin_bdgt"), i) = imm_fin_bdgt

    ' calcolo valore ATTIVO IMMOBILIZZATO
    att_imm_bdgt = somme_bdgt_SP(trova_riga_cdr(cod_rag, "imm_tecn_bdgt"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "imm_imm_bdgt"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "imm_fin_bdgt"), i)
    somme_bdgt_SP(trova_riga_cdr(cod_rag, "att_imm_bdgt"), i) = att_imm_bdgt

    ' calcolo valore ATTIVO IMMOBILIZZATO
    att_imm_bdgt = somme_bdgt_SP(trova_riga_cdr(cod_rag, "imm_tecn_bdgt"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "imm_imm_bdgt"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "imm_fin_bdgt"), i)
    somme_bdgt_SP(trova_riga_cdr(cod_rag, "att_imm_bdgt"), i) = att_imm_bdgt

    ' calcolo valore GIACENZE MAGAZZINO
    giac_mag_bdgt = somme_bdgt_SP(trova_riga_cdr(cod_rag, "rpf"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "rsem"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "rimb"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "rmp"), i)
    somme_bdgt_SP(trova_riga_cdr(cod_rag, "giac_mag_bdgt"), i) = giac_mag_bdgt

    ' calcolo valore CREDITI
    crediti_bdgt = somme_bdgt_SP(trova_riga_cdr(cod_rag, "credvscl"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "FSC"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "rratt"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "aacc"), i)
    somme_bdgt_SP(trova_riga_cdr(cod_rag, "crediti_bdgt"), i) = crediti_bdgt

    ' calcolo valore ATTIVITA' CORRENTI
    att_corr_bdgt = somme_bdgt_SP(trova_riga_cdr(cod_rag, "giac_mag_bdgt"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "crediti_bdgt"), i)
    somme_bdgt_SP(trova_riga_cdr(cod_rag, "att_corr_bdgt"), i) = att_corr_bdgt

    ' calcolo valore PASSIVITA' CORRENTI
    pass_corr_bdgt = somme_bdgt_SP(trova_riga_cdr(cod_rag, "forn"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "dip"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "eriva"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "fimp"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "iipp"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "impdd"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "debErateiz"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "aadeb"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "rrpass"), i)
    somme_bdgt_SP(trova_riga_cdr(cod_rag, "pass_corr_bdgt"), i) = pass_corr_bdgt

    ' calcolo valore CAPITALE CIRCOLANTE OPERATIVO
    cap_circ_op_bdgt = somme_bdgt_SP(trova_riga_cdr(cod_rag, "att_corr_bdgt"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "pass_corr_bdgt"), i)
    somme_bdgt_SP(trova_riga_cdr(cod_rag, "cap_circ_op_bdgt"), i) = cap_circ_op_bdgt

    ' calcolo valore LIQUIDITA'
    liq_bdgt = somme_bdgt_SP(trova_riga_cdr(cod_rag, "cash"), i)
    somme_bdgt_SP(trova_riga_cdr(cod_rag, "liq_bdgt"), i) = liq_bdgt

    ' calcolo valore CAPITALE INVESTITO
    cap_inv_bdgt = somme_bdgt_SP(trova_riga_cdr(cod_rag, "att_imm_bdgt"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "cap_circ_op_bdgt"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "liq_bdgt"), i)
    somme_bdgt_SP(trova_riga_cdr(cod_rag, "cap_inv_bdgt"), i) = cap_inv_bdgt

    ' calcolo valore PATRIMONIO NETTO
    patr_netto_bdgt = somme_bdgt_SP(trova_riga_cdr(cod_rag, "CS"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "ris"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "ut"), i)
    somme_bdgt_SP(trova_riga_cdr(cod_rag, "patr_netto_bdgt"), i) = patr_netto_bdgt

    ' calcolo valore PASSIVITA' A MEDIO/LUNGO
    pass_ml_bdgt = somme_bdgt_SP(trova_riga_cdr(cod_rag, "deblt"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "TFR"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "aaff"), i)
    somme_bdgt_SP(trova_riga_cdr(cod_rag, "pass_ml_bdgt"), i) = pass_ml_bdgt

    ' calcolo valore PASSIVITA' FINANZ A BREVE
    pass_b_bdgt = somme_bdgt_SP(trova_riga_cdr(cod_rag, "debbt"), i)
    somme_bdgt_SP(trova_riga_cdr(cod_rag, "pass_b_bdgt"), i) = pass_b_bdgt

    ' calcolo valore FONTI DI FINANZIAMENTO
    fonti_fin_bdgt = somme_bdgt_SP(trova_riga_cdr(cod_rag, "patr_netto_bdgt"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "pass_ml_bdgt"), i) + somme_bdgt_SP(trova_riga_cdr(cod_rag, "pass_b_bdgt"), i)
    somme_bdgt_SP(trova_riga_cdr(cod_rag, "fonti_fin_bdgt"), i) = fonti_fin_bdgt

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
