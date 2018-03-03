Attribute VB_Name = "SP_tabella_PER_dati"
Sub tabella_PER_SP(cod_rag, somme_cons_PER, somme_bdgt_SP_PER, date_cons, indice_data_PER)

Dim n_rig_piene_str_tab_SP As Integer
Dim cod_rag_tab As String

'0) creazione matrice struttura tabella matr_str_tab_SP(codice raggruppamento,caratteristiche fomrattazione)
n_rig_piene_str_tab_SP = Worksheets("str_tab_SP").Columns(1).Cells.SpecialCells(xlCellTypeConstants).Count

ReDim matr_str_tab_SP(n_rig_piene_str_tab_SP, 3)

For i = 1 To n_rig_piene_str_tab_SP
    matr_str_tab_SP(i, 1) = Worksheets("str_tab_SP").Cells(i + 1, 1)
    matr_str_tab_SP(i, 2) = Worksheets("str_tab_SP").Cells(i + 1, 2)
    matr_str_tab_SP(i, 3) = Worksheets("str_tab_SP").Cells(i + 1, 3)
Next i

'1) formattazione intestazione tabella

With Worksheets("SP_tab").Range("M6:R6")
    .Merge
    .NumberFormat = "@"
    .HorizontalAlignment = xlCenter
    .Font.Name = "Trebuchet MS"
    .Font.FontStyle = "Bold"
    .Font.Size = 10
    .Borders.Weight = xlMedium
    .Interior.Color = RGB(165, 165, 165)
End With

Worksheets("SP_tab").Range("M7:R7").Merge
With Worksheets("SP_tab").Range("M7:R7")
    .NumberFormat = "@"
    .HorizontalAlignment = xlCenter
    .Font.Name = "Trebuchet MS"
    .Font.FontStyle = "Bold"
    .Font.Size = 10
    .Borders.Weight = xlMedium
End With

Worksheets("SP_tab").Range("M8:N8").Merge
With Worksheets("SP_tab").Range("M8:N8")
    .NumberFormat = "@"
    .HorizontalAlignment = xlCenter
    .Font.Name = "Trebuchet MS"
    .Font.FontStyle = "Bold"
    .Font.Size = 10
    .Borders.Weight = xlMedium
End With

Worksheets("SP_tab").Range("O8:P8").Merge
With Worksheets("SP_tab").Range("O8:P8")
    .NumberFormat = "@"
    .HorizontalAlignment = xlCenter
    .Font.Name = "Trebuchet MS"
    .Font.FontStyle = "Bold"
    .Font.Size = 10
    .Borders.Weight = xlMedium
End With
   
Worksheets("SP_tab").Range("Q8:R8").Merge
With Worksheets("SP_tab").Range("Q8:R8")
    .NumberFormat = "@"
    .HorizontalAlignment = xlCenter
    .Font.Name = "Trebuchet MS"
    .Font.FontStyle = "Bold"
    .Font.Size = 10
    .Borders.Weight = xlMedium
End With

With Worksheets("SP_tab").Range("M9:R9")
    .NumberFormat = "@"
    .HorizontalAlignment = xlCenter
    .Font.Name = "Trebuchet MS"
    .Font.FontStyle = "Bold"
    .Font.Size = 10
    .Borders.Weight = xlMedium
End With

With Worksheets("SP_tab").Range("M9:R9")
    .NumberFormat = "@"
    .HorizontalAlignment = xlCenter
    .Font.Name = "Trebuchet MS"
    .Font.FontStyle = "Bold"
    .Font.Size = 10
    .Borders.Weight = xlMedium
End With

Worksheets("SP_tab").Rows(9).RowHeight = 26
Worksheets("SP_tab").Columns("E").ColumnWidth = 34
Worksheets("SP_tab").Columns("M").ColumnWidth = 19
Worksheets("SP_tab").Columns("N").ColumnWidth = 10
Worksheets("SP_tab").Columns("O").ColumnWidth = 19
Worksheets("SP_tab").Columns("P").ColumnWidth = 10
Worksheets("SP_tab").Columns("Q").ColumnWidth = 19
Worksheets("SP_tab").Columns("R").ColumnWidth = 10

'2) Formattazione colonne numeri come "CONTABILITA'" e allineamento testo

Worksheets("SP_tab").Activate

With Worksheets("SP_tab").Range("M10", Range("M10").End(xlDown))
    .NumberFormat = "_( €* #,##0.00_);_(-€* #,##0.00;_( €* ""-""??_);_(@_)"
    .HorizontalAlignment = xlRight
End With

With Worksheets("SP_tab").Range("O10", Range("O10").End(xlDown))
    .NumberFormat = "_( €* #,##0.00_);_(-€* #,##0.00;_( €* ""-""??_);_(@_)"
    .HorizontalAlignment = xlRight
End With

With Worksheets("SP_tab").Range("Q10", Range("Q10").End(xlDown))
    .NumberFormat = "_( €* #,##0.00_);_(-€* #,##0.00;_( €* ""-""??_);_(@_)"
    .HorizontalAlignment = xlRight
End With

With Worksheets("SP_tab").Range("N10", Range("N10").End(xlDown))
    .NumberFormat = "0.0%;[Red](0.0%)"
    .HorizontalAlignment = xlRight
End With
    
With Worksheets("SP_tab").Range("P10", Range("P10").End(xlDown))
    .NumberFormat = "0.0%;[Red](0.0%)"
    .HorizontalAlignment = xlRight
End With
   
With Worksheets("SP_tab").Range("R10", Range("R10").End(xlDown))
    .NumberFormat = "0.0%;[Red](0.0%)"
    .HorizontalAlignment = xlRight
End With

'3) scrittura testi intestazione
Worksheets("SP_tab").Range("M6") = "STATO PATRIMONIALE"
Worksheets("SP_tab").Range("M7") = "ANALISI DI PERIODO (PER): " & "DAL " & date_cons(indice_data_PER) & " " & "AL " & date_cons(indice_data_PER + 1)
Worksheets("SP_tab").Range("M8") = "ACTUAL"
Worksheets("SP_tab").Range("O8") = "BUDGET"
Worksheets("SP_tab").Range("Q8") = "VARIANCE"
Worksheets("SP_tab").Range("M9") = "VALUE"
Worksheets("SP_tab").Range("N9") = "%"
Worksheets("SP_tab").Range("O9") = "VALUE"
Worksheets("SP_tab").Range("P9") = "%"
Worksheets("SP_tab").Range("Q9") = "VALUE"
Worksheets("SP_tab").Range("R9") = "%"


'6) compilazione tabella colonna dati consuntivo
For i = 1 To n_rig_piene_str_tab_SP - 1
    cod_rag_i = matr_str_tab_SP(i, 1)    'codice raggruppamento i-esimo
    n = trova_riga_cdr(cod_rag, cod_rag_i)
    Worksheets("SP_tab").Cells(9 + i, 5) = cod_rag(n, 2)
    Worksheets("SP_tab").Cells(9 + i, 13) = somme_cons_PER(n)
Next i



'8) compilazione tabella colonna dati budget
For j = 1 To n_rig_piene_str_tab_SP - 1
    cod_rag_tab = matr_str_tab_SP(j, 2)
    m = trova_riga_cdr(cod_rag, cod_rag_tab)
    Worksheets("SP_tab").Cells(9 + j, 15) = somme_bdgt_SP_PER(m)
Next j

'7) compilazione tabella colonna "VARIANCE"
For k = 1 To n_rig_piene_str_tab_SP - 1
    Worksheets("SP_tab").Cells(9 + k, 17) = Worksheets("SP_tab").Cells(9 + k, 13) - Worksheets("SP_tab").Cells(9 + k, 15)
Next k

For l = 1 To n_rig_piene_str_tab_SP - 1
    If Worksheets("SP_tab").Cells(9 + l, 15) <> 0 Then
        Worksheets("SP_tab").Cells(9 + l, 18) = (Worksheets("SP_tab").Cells(9 + l, 17)) / (Worksheets("SP_tab").Cells(9 + l, 15))
    Else
        Worksheets("SP_tab").Cells(9 + l, 18) = "-"
    End If
        
Next l

'8) compilazione tabella colonna "%" ACTUAL
For m = 1 To n_rig_piene_str_tab_SP - 1
    If Worksheets("SP_tab").Cells(9 + m, 13) <> 0 Then
        Worksheets("SP_tab").Cells(9 + m, 14) = Worksheets("SP_tab").Cells(9 + m, 13) / vendite_cons
    Else
        Worksheets("SP_tab").Cells(9 + m, 14) = "-"
    End If
        
Next m

'9) compilazione tabella colonna "%" BUDGET
For n = 1 To n_rig_piene_str_tab_SP - 1
    If Worksheets("SP_tab").Cells(9 + n, 15) <> 0 Then
        Worksheets("SP_tab").Cells(9 + n, 16) = Worksheets("SP_tab").Cells(9 + n, 15) / vendite_bdgt
    Else
        Worksheets("SP_tab").Cells(9 + n, 16) = "-"
    End If
        
Next n

'4) Formattazione corpo della tabella
 For i = 1 To n_rig_piene_str_tab_SP - 1
    ' Formattazione bordi tabella
    Worksheets("SP_tab").Cells(9 + i, 13).Borders.Weight = xlThin
    Worksheets("SP_tab").Cells(9 + i, 14).Borders.Weight = xlThin
    Worksheets("SP_tab").Cells(9 + i, 15).Borders.Weight = xlThin
    Worksheets("SP_tab").Cells(9 + i, 16).Borders.Weight = xlThin
    Worksheets("SP_tab").Cells(9 + i, 17).Borders.Weight = xlThin
    Worksheets("SP_tab").Cells(9 + i, 18).Borders.Weight = xlThin
    ' Formattazione carattere tabella
    If matr_str_tab_SP(i, 3) = "g" Then
        Worksheets("SP_tab").Rows(9 + i).Font.FontStyle = "Bold"
    End If
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
