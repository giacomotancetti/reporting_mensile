Attribute VB_Name = "SP_tabella_YTD_dati"
Sub tabella_YTD_SP(cod_rag, somme_cons, somme_bdgt_SP, data_an_YTD, mese_YTD)

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

With Worksheets("SP_tab").Range("F6:K6")
    .Merge
    .NumberFormat = "@"
    .HorizontalAlignment = xlCenter
    .Font.Name = "Trebuchet MS"
    .Font.FontStyle = "Bold"
    .Font.Size = 10
    .Borders.Weight = xlMedium
    .Interior.Color = RGB(165, 165, 165)
End With

Worksheets("SP_tab").Range("F7:K7").Merge
With Worksheets("SP_tab").Range("F7:K7")
    .NumberFormat = "@"
    .HorizontalAlignment = xlCenter
    .Font.Name = "Trebuchet MS"
    .Font.FontStyle = "Bold"
    .Font.Size = 10
    .Borders.Weight = xlMedium
End With

Worksheets("SP_tab").Range("F8:G8").Merge
With Worksheets("SP_tab").Range("F8:G8")
    .NumberFormat = "@"
    .HorizontalAlignment = xlCenter
    .Font.Name = "Trebuchet MS"
    .Font.FontStyle = "Bold"
    .Font.Size = 10
    .Borders.Weight = xlMedium
End With

Worksheets("SP_tab").Range("H8:I8").Merge
With Worksheets("SP_tab").Range("H8:I8")
    .NumberFormat = "@"
    .HorizontalAlignment = xlCenter
    .Font.Name = "Trebuchet MS"
    .Font.FontStyle = "Bold"
    .Font.Size = 10
    .Borders.Weight = xlMedium
End With
   
Worksheets("SP_tab").Range("J8:K8").Merge
With Worksheets("SP_tab").Range("J8:K8")
    .NumberFormat = "@"
    .HorizontalAlignment = xlCenter
    .Font.Name = "Trebuchet MS"
    .Font.FontStyle = "Bold"
    .Font.Size = 10
    .Borders.Weight = xlMedium
End With


With Worksheets("SP_tab").Range("F9:K9")
    .NumberFormat = "@"
    .HorizontalAlignment = xlCenter
    .Font.Name = "Trebuchet MS"
    .Font.FontStyle = "Bold"
    .Font.Size = 10
    .Borders.Weight = xlMedium
End With

With Worksheets("SP_tab").Range("F9:K9")
    .NumberFormat = "@"
    .HorizontalAlignment = xlCenter
    .Font.Name = "Trebuchet MS"
    .Font.FontStyle = "Bold"
    .Font.Size = 10
    .Borders.Weight = xlMedium
End With

Worksheets("SP_tab").Rows(9).RowHeight = 26
Worksheets("SP_tab").Columns("E").ColumnWidth = 34
Worksheets("SP_tab").Columns("F").ColumnWidth = 19
Worksheets("SP_tab").Columns("G").ColumnWidth = 10
Worksheets("SP_tab").Columns("H").ColumnWidth = 19
Worksheets("SP_tab").Columns("I").ColumnWidth = 10
Worksheets("SP_tab").Columns("J").ColumnWidth = 19
Worksheets("SP_tab").Columns("K").ColumnWidth = 10

'2) Formattazione colonne numeri come "CONTABILITA'" e allineamento testo

Worksheets("SP_tab").Activate

With Worksheets("SP_tab").Range("F10", Range("F10").End(xlDown))
    .NumberFormat = "_( €* #,##0.00_);_(-€* #,##0.00;_( €* ""-""??_);_(@_)"
    .HorizontalAlignment = xlRight
End With

With Worksheets("SP_tab").Range("H10", Range("H10").End(xlDown))
    .NumberFormat = "_( €* #,##0.00_);_(-€* #,##0.00;_( €* ""-""??_);_(@_)"
    .HorizontalAlignment = xlRight
End With

With Worksheets("SP_tab").Range("J10", Range("J10").End(xlDown))
    .NumberFormat = "_( €* #,##0.00_);_(-€* #,##0.00;_( €* ""-""??_);_(@_)"
    .HorizontalAlignment = xlRight
End With

With Worksheets("SP_tab").Range("K10", Range("K10").End(xlDown))
    .NumberFormat = "0.0%;[Red](0.0%)"
    .HorizontalAlignment = xlRight
End With
    
With Worksheets("SP_tab").Range("G10", Range("G10").End(xlDown))
    .NumberFormat = "0.0%;[Red](0.0%)"
    .HorizontalAlignment = xlRight
End With
   
With Worksheets("SP_tab").Range("I10", Range("I10").End(xlDown))
    .NumberFormat = "0.0%;[Red](0.0%)"
    .HorizontalAlignment = xlRight
End With

'3) scrittura testi intestazione
Worksheets("SP_tab").Range("F6") = "STATO PATRIMONIALE"
Worksheets("SP_tab").Range("F7") = "DATA ANALISI YEAR TO DATE (YTD): " & data_an_YTD
Worksheets("SP_tab").Range("F8") = "ACTUAL"
Worksheets("SP_tab").Range("H8") = "BUDGET"
Worksheets("SP_tab").Range("J8") = "VARIANCE"
Worksheets("SP_tab").Range("F9") = "VALUE"
Worksheets("SP_tab").Range("G9") = "%"
Worksheets("SP_tab").Range("H9") = "VALUE"
Worksheets("SP_tab").Range("I9") = "%"
Worksheets("SP_tab").Range("J9") = "VALUE"
Worksheets("SP_tab").Range("K9") = "%"

'5) compilazione tabella colonna dati consuntivo

For i = 1 To n_rig_piene_str_tab_SP - 1
    cod_rag_tab = matr_str_tab_SP(i, 1)
    n = trova_riga_cdr(cod_rag, cod_rag_tab)
    Worksheets("SP_tab").Cells(9 + i, 5) = cod_rag(n, 2)
    Worksheets("SP_tab").Cells(9 + i, 6) = somme_cons(n, par_data)
Next i

'6) compilazione tabella colonna dati budget

For j = 1 To n_rig_piene_str_tab_SP - 1
    cod_rag_tab = matr_str_tab_SP(j, 2)
    m = trova_riga_cdr(cod_rag, cod_rag_tab)
    Worksheets("SP_tab").Cells(9 + j, 8) = somme_bdgt_SP(m, mese_YTD)
Next j

'7) compilazione tabella colonna "VARIANCE"
For k = 1 To n_rig_piene_str_tab_SP - 1
    Worksheets("SP_tab").Cells(9 + k, 10) = Worksheets("SP_tab").Cells(9 + k, 6) - Worksheets("SP_tab").Cells(9 + k, 8)
Next k

For l = 1 To n_rig_piene_str_tab_SP - 1
    If Worksheets("SP_tab").Cells(9 + l, 8) <> 0 Then
        Worksheets("SP_tab").Cells(9 + l, 11) = (Worksheets("SP_tab").Cells(9 + l, 10)) / (Worksheets("SP_tab").Cells(9 + l, 8))
    Else
        Worksheets("SP_tab").Cells(9 + l, 11) = "-"
    End If
        
Next l

'8) compilazione tabella colonna "%" ACTUAL

For m = 1 To n_rig_piene_str_tab_SP - 1
    If Worksheets("SP_tab").Cells(9 + m, 6) <> 0 Then
        Worksheets("SP_tab").Cells(9 + m, 7) = Worksheets("SP_tab").Cells(9 + m, 6) / vendite_cons
    Else
        Worksheets("SP_tab").Cells(9 + m, 7) = "-"
    End If
        
Next m

'9) compilazione tabella colonna "%" BUDGET

For n = 1 To n_rig_piene_str_tab_SP - 1
    If Worksheets("SP_tab").Cells(9 + n, 8) <> 0 Then
        Worksheets("SP_tab").Cells(9 + n, 9) = Worksheets("SP_tab").Cells(9 + n, 8) / vendite_bdgt
    Else
        Worksheets("SP_tab").Cells(9 + n, 9) = "-"
    End If
Next n

'4) Formattazione corpo della tabella
 For i = 1 To n_rig_piene_str_tab_SP - 1
    ' Formattazione bordi tabella
    Worksheets("SP_tab").Cells(9 + i, 5).Borders.Weight = xlThin
    Worksheets("SP_tab").Cells(9 + i, 6).Borders.Weight = xlThin
    Worksheets("SP_tab").Cells(9 + i, 7).Borders.Weight = xlThin
    Worksheets("SP_tab").Cells(9 + i, 8).Borders.Weight = xlThin
    Worksheets("SP_tab").Cells(9 + i, 9).Borders.Weight = xlThin
    Worksheets("SP_tab").Cells(9 + i, 10).Borders.Weight = xlThin
    Worksheets("SP_tab").Cells(9 + i, 11).Borders.Weight = xlThin
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
