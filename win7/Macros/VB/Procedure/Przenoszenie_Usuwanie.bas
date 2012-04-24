Attribute VB_Name = "Przenoszenie_Usuwanie"
Option Explicit
Option Base 1


' Funkcja WinApi - odlicza milisekundy od otwarcia systemu.
Public Declare Function GetTickCount Lib "kernel32" () As Long


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
' obszar(1)- zaznaczony wiersz 1 i tak az do ostatniego, a ostatni obszar(n) w tablicy
' zawiera zawsze obszar ostatnie zaznaczonej kolumny w wierszu g��wnym.
Dim Obszar() As Range
'parametr(1) - obszar g��wny, parametr(2)- parametry wiersza pierwszego(rozpoczynaj�cego)
Dim Parametry() As NewType1


'
Dim tytul As String
' OdsWierszy - jest zmienn� przechowuj�c� liczb� wierszy dziel�cych wiersze g��wne.
'   Np: co 4 - wiersz jest wierszem g��wnym, do kt�ego przenoszone s� warto�ci.
' LiWiPrzen - jest zmienn� przechowuj�c� liczb� wierszy kt�r� u�ytkownik podaje,
'   �e chce je przeni��. Nast�pnie musi je zaznaczy� i w takiej kolejno�ci jak je zaznaczy
'   b�d� one przenoszone.
Dim PunktX As Integer, PunktY As Integer, OdsWierszy As Integer, LiWiPrzen As Integer
' odpowidz dotycz�ca usuni�cia pustych wierszy.
Dim odp As Long
' czas2 - zmienna ta przechwytuje czas zako�czenia dzia�ania programu.
Dim czas2 As Long
'
Dim szybko As Boolean


'***************************************************************************************

Public Sub Przenoszenie_Usuwanie()

' Program ten ma za zadanie przenie�� zawarto�� wierszy znajduj�cych si� pod
' wierszem g��wnym. u�ytkownik musi najpierw zaznaczy� obszar g��wny w jaki ma by�
' wykonywany progarm, a nast�pnie musi poda� co ile wierszy znajduje si� wiersz
' g��wny(wiersz do kt�rego s� przenoszone zawarto�ci kom�rek z wierszy znajduj�cych
' si� poni�ej). Nast�pnie podaje ile wierszy ma by� za ka�dym razem przeniesionych.
' W czwartym kroku u�ytkownik zaznacza zadeklarowan� wcze�niej liczb� wierszy pojedynczo.
' Ostatni krok s�u�y do wskazania ostatnij kolumny wiersza g��wnego, po kt�rej to kolumnie
' nast�pi wstawianie zawarto�ci kom�rek. Odst�py pomi�dzy wierszami g��wnymi musz�
' by� takie same aby program zadzia�a� poprawnie.

Dim czas1 As Long
Dim WynikCzas As Long
Dim CzasDzialania As Variant


'
tytul = "Przenoszeni i usuwanie pustych wierszy."
PunktX = 450
PunktY = 100


'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
    
    '  wywo�ywanie procedury
    Call Obszar_glowny

'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP

'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP

    ' wywo�ywanie procedury
    Call Obszary
    
'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP


                      ' wy��czenie funkcji wy�wietlaj�cej informacje post�pie wykonywania
                      ' programu, w celu przy�pieszenia dzia�ania programu.
If szybko = True Then Application.ScreenUpdating = False


' pobranie czasu wej�cia do g��wnego �r�d�a programu.
czas1 = GetTickCount

'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
    
    '  wywo�ywanie procedury
    Call Przenoszenie

'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP


' obliczenie r�nicy w czasie dzia�ania programu.
WynikCzas = ((czas2 - czas1) / 1000)

' przekszta�cenie wyniku w czas mgodzin, minut, sekund.
CzasDzialania = TimeSerial(0, 0, WynikCzas)


                      ' ponowne w��czenie funkcji.
If szybko = True Then Application.ScreenUpdating = True



    ' wy�wietlanie komunikatu zako�czenia pracy
    MsgBox "Koniec dzia�ania programu. Czas dzia�ania wynosi - " & CzasDzialania, vbInformation, tytul
    

End Sub

Private Sub Przenoszenie()

' Procedura ta przenosi zawarto�� zaznaczonych wierszy w zaznaczonym
' g��wnym obszarze do g��wnegwiersza znajduj�cego si� co zmienna "OdsWierszy"

'
Dim i1 As Integer, i2 As Integer, Dodatek As Integer, Przes As Integer, w As Integer
Dim p As Long, LKwstaw As Integer
Dim k1 As Integer, k2 As Integer
Dim kp1 As Integer, kp2 As Integer, w1 As Long


'
Dodatek = 0
Przes = 0
w = 0
p = 1
LKwstaw = 0


'
For i1 = 1 To Parametry(1).LiWier ' liczba wierszy w obszarze g��wnym.


    '
    If i1 = (Parametry(LiWiPrzen + 2).NrWier - Parametry(1).NrWier) + Dodatek + 1 Then
        
        Dodatek = Dodatek + 1
        Przes = Przes + 1
        
        ' ustawianie parametr�w wierszy i kolumn do zaznaczenie
        w = Parametry(Przes + 1).NrWier - Parametry(1).NrWier + p
        k1 = Parametry(Przes + 1).NrKolum - Parametry(1).NrKolum + 1
        k2 = Parametry(Przes + 1).LiKolum + Parametry(Przes + 1).NrKolum - Parametry(1).NrKolum
        
        ' zaznaczanie wierszy do przeniesienie wed�ug zaznaczonego g��wnego obszaru
        ObszarGl.Range(Cells(w, k1), Cells(w, k2)).Select
       
       '------------------------------------------------------------------------------------
       
       ' ustawianie parametr�w wierszy i kolumn zaznaczonych przeznaczonych
       ' do przeniesienia.
       w1 = Parametry(LiWiPrzen + 2).NrWier + p - 1
       kp1 = Parametry(LiWiPrzen + 2).NrKolum + LKwstaw + 1
       kp2 = Parametry(LiWiPrzen + 2).NrKolum + Parametry(Przes + 1).LiKolum + LKwstaw
       
        ' przenoszenie zaznaczonego wiersza do wskazanego miejsca.
        Selection.Cut Destination:=Range(Cells(w1, kp1), Cells(w1, kp2))
                
       '------------------------------------------------------------------------------------
        
        ' ustawianie nowej kolumny do wstawiania zawarto�ci wierszy. Kolumna pierwsza
        ' i ostatnia zwi�ksza si� za ka�dym razem o liczb� kolumn, kt�re zosta�y wstawione
        ' z wierszy zaznaczonych przez u�ytkownika.
        LKwstaw = LKwstaw + Parametry(Przes + 1).LiKolum
        
        '
        If Przes = LiWiPrzen Then
            Dodatek = (CInt(OdsWierszy) + Dodatek) - LiWiPrzen
            Przes = 0
            p = p + CInt(OdsWierszy)
            LKwstaw = 0
        End If
        
    
    End If

    '
    DoEvents
    
Next i1

'---------------------------------------------------------------------------------

' zapisanie czasu zako�czenia dzia�ania programu.
czas2 = GetTickCount

If odp = 6 Then

    'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
        
        ' wywo�ywanie procedury usuwaj�cej puste wiersze w zaznaczonym obszarze.
        Call UsuwanieWierszy
    
    'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP

End If


End Sub

Private Sub UsuwanieWierszy()

' Procedura ta usuwa puste wiersze w zaznaczonym obszarze po przeniesieniu
' zawarto�ci wierszy do jednego g��wnego.

Dim suma As Long, i As Long, i2 As Long
Dim usun As String
Dim PustaKom As Boolean
' Zmienna ta przechwuj� liczbe usuni�tych wierszy z zaznaczonego obszaru.
Dim LUW As Long


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
        'ObszarGl.Range(usun).EntireRow.Delete
        suma = suma - 1
        LUW = LUW + 1
    End If
    
    '
    DoEvents
    
Loop

' zapisanie czasu zako�czenia dzia�ania programu.
czas2 = GetTickCount

    '
    MsgBox "Usuni�to - " & LUW & " - wierszy z zaznaczonego obszaru - " & Parametry(1).LiWier & " - wierszy.", vbInformation, tytul


End Sub

Private Sub Obszar_glowny()

' W tej procedurze zostaje zaznaczony g��wny obszar przez u�ytkownika
' i zostaj� pobrane jego parametry do tablicy Parametry(), kt�ra to przechowuje
' w�a�ciwo�ci wszystkich zaznaczonych obszar�w.

Dim kom As String

'
kom = "Zaznacz obszar, w kt�rym chcesz uszeregowa� warto�ci kom�rek?"
'
Set ObszarGl = Application.InputBox(prompt:=kom, Title:=tytul, Left:=PunktX, _
                                    Top:=PunktY, Type:=8)



' ustawianie wst�pnego rozmiaru tablicy Parametry().
ReDim Parametry(1)


' pobieranie parametr�w obszaru g��wnego.
Parametry(1).LiKolum = ObszarGl.Columns.Count
Parametry(1).LiWier = ObszarGl.Rows.Count
Parametry(1).NrKolum = ObszarGl.Column
Parametry(1).NrWier = ObszarGl.Row


' ustawianie rozmiaru tablicy na liczb� zaznaczonych wierszy w obszarze g��wnym.
ReDim TabPustWier(Parametry(1).LiWier)


End Sub

Private Sub Obszary()

' Zaznaczanie wierszy zadeklarowanych przez u�ytkownika i pobranie ich w�a�ciwo�ci
' do tablicy Parametry().

Dim kom1 As String, kom2 As String
Dim i As Integer


'
LiWiPrzen = 0
OdsWierszy = 0

'
kom2 = "Podaj co ile wierszy znajduje si� wiersz g��wny?"
' pobieranie od u�ytkownika liczby wierszy do przeniesienia.
OdsWierszy = InputBox(kom2, tytul, 5, PunktX + 8500, PunktY + 3500)

'
kom2 = "Podaj ile wierszy chcesz przenie��?"
' pobieranie od u�ytkownika liczby wierszy do przeniesienia.
LiWiPrzen = InputBox(kom2, tytul, 3, PunktX + 8500, PunktY + 3500)


' ustawianie rozmiaru tablicy.
ReDim Obszar(LiWiPrzen + 1)
ReDim Preserve Parametry(LiWiPrzen + 2)

'
For i = 1 To LiWiPrzen + 1
    
    If i = (LiWiPrzen + 1) Then
        '
        kom1 = "Zaznacz ostatni� kolumn� wiersza, po kt�rego ma si� " & Chr(10)
        kom1 = kom1 & "rozpocz�� wstawienie zawarto�ci kom�rek z wierszy " & Chr(10)
        kom1 = kom1 & "znajduj�cych si� poni�ej?" & Chr(10)
    Else
        '
        kom1 = "Zaznacz " & i & " wiersz, kt�ry ma by� przeniesiony."
    End If
    
    
    Set Obszar(i) = Application.InputBox(prompt:=kom1, Title:=tytul, _
                                        Left:=PunktX, Top:=PunktY + 10 + (i * 10), Type:=8)
    
    ' pobieranie do tablicy parametr�w poszczeg�lnych obszar�w zaznaczonych
    ' przez u�ytkownika.
    Parametry(i + 1).LiKolum = Obszar(i).Columns.Count
    Parametry(i + 1).LiWier = Obszar(i).Rows.Count
    Parametry(i + 1).NrKolum = Obszar(i).Column
    Parametry(i + 1).NrWier = Obszar(i).Row
        
Next i

odp = MsgBox("Chcesz przy�pieszy� dzia�anie programu?", _
                    vbInformation + vbYesNo + vbDefaultButton1, tytul)
If odp = 6 Then
    szybko = True
Else
    szybko = False
End If

'
odp = MsgBox("Czy chcesz usun�� puste wiersze z zaznaczonego obszaru?", _
                    vbInformation + vbYesNo + vbDefaultButton1, tytul)


End Sub

'***************************************************************************************
