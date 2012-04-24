Attribute VB_Name = "FormatowanieKomorekData"
Option Explicit

Dim Tytul As String
Dim Obszar As Range
Dim LiWierszy As Long, LiKolumn As Long
Dim PWr As Long, PKl As Long
Dim Dzien As Integer, Dzien1 As Integer



Sub FormatKomorData()

' Program ten formatuje wiersze na ca�ej jego d�ugo�ci. W tym przypadku program
' ten ma za zadanie pogrubi� g�rn� linie kom�rki, kt�ra spe�nia postawione przez
' urzytkownika kryteria. Urzytkownik musi zaznaczy� kom�rki, w kt�rych s� wpisane
' daty poszczeg�lnych losowa�, a program bogrubi g�rn� linie danego wiersza
' je�eli rozpozna, �e data w tym wierszu jest pierwszym dniem ka�dego miesi�ca.
' Daty musz� by� wpisane w wierszach.


Dim kom As String


Tytul = "Formatowanie kom�rek ze wzgl�du na dat�"

kom = "Zaznacz obszar kom�rek, w kt�rych znajduj� si� daty " & Chr(13)
kom = kom & "poszczeg�lnych losowa�. Kom�rki te powinny znajdowa�" & Chr(13)
kom = kom & "si� w poszczeg�lnych wierszach" & Chr(13)

'
Set Obszar = Application.InputBox(prompt:=kom, Title:=Tytul, Type:=8)



'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP

    '
    Call Przeszukiwanie
    
'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP

MsgBox "Koniec dzia�ania tego programu", vbExclamation, Tytul

End Sub

Private Sub Przeszukiwanie()

'
Dim i As Long, i1 As Long
Dim Arkusz As String


'
Arkusz = Obszar.Worksheet.Name
'
Worksheets(Arkusz).Activate


LiWierszy = Obszar.Rows.Count
LiKolumn = Obszar.Columns.Count

PWr = Obszar.Row - 1
PKl = Obszar.Column - 1


For i = 1 To LiWierszy
    
    For i1 = 1 To LiKolumn
    
        If IsDate(Obszar(i, i1)) = True Then
        
            'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
                
                '
                Call Formatowanie(i, i1)
            
            'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
        
        End If
    
    Next i1

Next i


End Sub

Private Sub Formatowanie(i, i1)

'

Dzien = Day(Obszar(i, i1))

If Dzien < Dzien1 Or Day(Obszar(i, i1)) = 1 Then
   
        Cells(i + PWr, i1 + PKl).EntireRow.Select
        
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = xlAutomatic
        End With

   
End If

Dzien1 = Dzien

End Sub
