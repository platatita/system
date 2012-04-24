Attribute VB_Name = "Wyszukaj_Pionowo"
Option Explicit

'Funkcji tej nalezy przeaza� 4 parametry; 1- to zaznaczony obszar w arkuszu; 2- to szukana warto��; 3- to nr kolumny w kt�rej ma szuka� tej warto�ci
'4- to nr kolumny z kt�rej ma pobiera� warto�ci i sumowa� je w przypadku znalezienia szukanej warto�ci w wskazanej wcze�nie kolumnie.
Public Function Wyszukaj_Pionowo(ByVal SelectedRange As Range, ByVal SearchVal As Range, _
                                    ByVal NrColumn_SearchVal As Integer, ByVal NrColumn_WithVal As Integer) As Double

    '
    Dim rg As Range
    Dim i1, i2 As Integer
    Dim suma As Double
    
    '
    For i1 = 1 To SelectedRange.Rows.Count
        '
        For i2 = 1 To SelectedRange.Columns.Count
        
            '
            If i2 = NrColumn_SearchVal And SelectedRange.Cells(i1, i2).Value = SearchVal.Value Then
                '
                suma = suma + SelectedRange.Cells(i1, NrColumn_WithVal).Value
                
            End If
        
        Next i2
    
    Next i1
    
    
    '
    Wyszukaj_Pionowo = suma


End Function

