Attribute VB_Name = "Usuwanie_Wierszy"
Option Explicit
Option Base 1


' w�asny typ, kt�ry ma przechowywa� nr. kolumn, nr. wiersza, liczb� wierszy
' i liczb� kolumn z zaznaczonych obszar�w.
Private Type NewType1

    NrWier As Long
    NrKolum As Long
    LiWier As Long
    LiKolum As Long
    
End Type


' ObszarGl - zmienna ta przechowuje wszystkie w�a�ciwo�ci i parametry
' zaznaczonego g��wnego obszaru w kt�rym b�dzie si� odbywa�o przenoszenie warto�ci
' kom�rek i usuwanie pustych wierszy.
Dim ObszarGl As Range
'parametr(1) - obszar g��wny, parametr(2)- parametry wiersza pierwszego(rozpoczynaj�cego)
Dim Parametry() As NewType1


'
Dim tytul As String



Public Sub Usuwanie_Wierszy()

' Procedura ta usuwa puste wiersze w zaznaczonym obszarze po przeniesieniu
' zawarto�ci wierszy do jednego g��wnego.

Dim suma As Long, i As Long, i2 As Long
Dim usun As String
Dim PustaKom As Boolean
' Zmienna ta przechwuj� liczbe usuni�tych wierszy z zaznaczonego obszaru.
Dim LUW As Long

'
tytul = "Usuwanie wierszy"


    '
    Call Obszar_glowny


'
PustaKom = False
suma = 0
usun = ""
i = 0
i2 = 0
LUW = 0


' G��wna p�tla usuwaj�ca puste wiersze.
Do While (i < Parametry(1).LiWier)
    
    '
    i = i + 1
    suma = suma + 1
    ' zerowanie zmiennych wskazuj�cych na brak zawarto�ci w kom�rkach.
    PustaKom = False

    '
    For i2 = 1 To Parametry(1).LiKolum 'liczba kolumn w obszarze g��wnym
    
            If IsEmpty(ObszarGl.Cells(suma, i2)) = True Then
                PustaKom = True
            Else
                PustaKom = False
            End If

    Next i2

    '
    If PustaKom = True Then
    
        usun = CStr((suma + Parametry(1).NrWier - 1)) & ":" & CStr((suma + Parametry(1).NrWier - 1))
        Range(usun).EntireRow.Delete
        '
        suma = suma - 1
        LUW = LUW + 1
    End If
    
    '
    DoEvents
    
Loop


    '
    MsgBox "Usuni�to - " & LUW & " - wierszy z zaznaczonego obszaru - " & Parametry(1).LiWier & " - wierszy.", vbInformation, tytul


End Sub


Private Sub Obszar_glowny()

' W tej procedurze zostaje zaznaczony g��wny obszar przez u�ytkownika
' i zostaj� pobrane jego parametry do tablicy Parametry(), kt�ra to przechowuje
' w�a�ciwo�ci wszystkich zaznaczonych obszar�w.

Dim kom As String

'
kom = "Zaznacz obszar, w kt�rym chcesz usun�� puste wiersze?"
'
Set ObszarGl = Application.InputBox(prompt:=kom, Title:=tytul, Left:=450, _
                                    Top:=450, Type:=8)



' ustawianie wst�pnego rozmiaru tablicy Parametry().
ReDim Parametry(1)


' pobieranie parametr�w obszaru g��wnego.
Parametry(1).LiKolum = ObszarGl.Columns.Count
Parametry(1).LiWier = ObszarGl.Rows.Count
Parametry(1).NrKolum = ObszarGl.Column
Parametry(1).NrWier = ObszarGl.Row



End Sub

