Attribute VB_Name = "trova_indice_data_PER"
Function trova_ind_data_PER(date_cons, data_an_PER)

    ' trova la posizione della data di analisi periodica nel vettore date_cons
    pos = trova_riga_vet(date_cons, data_an_PER)
    
    trova_ind_data_PER = pos
    
End Function


' Funzione trova riga elemento in vettore
Function trova_riga_vet(vet, val) As Integer
    Dim r As Integer
    For r = 1 To UBound(vet, 1)
            If vet(r) = val Then
                trova_riga_vet = r
                Exit Function
            End If
    Next r
End Function
