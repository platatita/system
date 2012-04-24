Attribute VB_Name = "Parse_Text_Lotto"
Option Explicit


' Funkcja WinApi - odlicza milisekundy od otwarcia systemu.
Public Declare Function GetTickCount Lib "kernel32" () As Long

'
'deklaracja globalnej zmiennej
Private tytul As String


Public Sub Parse_Text_Lotto()
    '
    'Program ten s�u�y do parsowania tekstu, czyli rozbijania na poszczeg�lne litery lub ci�gi znak�w.
    'W obecnej postaci jest on przystosowany do wyszukiwania liczb, numer�w losowa� i dat losowa�
    'w pobranych z internetu tekstach. Program ten dzia�a na zasadzie rozpoznawania, kt�ra litera w
    'przeszukiwanym ci�gu znak�w jest typu numerycznego. Je�eli mi�dzy znalezionymi liczbami nie
    'wystepuje �aden separator(czyli inny znak od cyfry) opr�cz separator�w zadeklarowanych traktuje ten
    'ci�g tekstowy jako jedn� liczb�. Nastepnie ka�d� rozpoznan� liczb� czy te� dat� w tekscie dodaje
    'do listy, kt�ra to w jedny cyklu(jedny obrocie p�tli) zawiera rozbity tekst. Tak wype�niona list jest
    'przekazywana do wpisania w kom�rki od miejsca wcze�niej wskazanego. I tak wk�ko a� do przeszukania
    'ca�ego zaznaczonego tekstu.
    '
    
    'deklaracja lokalnych zmiennych
    Dim znaki
    Dim i As Long, j As Long
    Dim parse_text As String
    Dim past_range As Range
    Dim search_range As Range
    Dim al As New ArrayList
    Dim odp As Byte
    Dim czas1 As Long, czas2 As Long, czas As Single
    
    
    '
    tytul = "Parse Text"
    '
    'zapisywanie do zmiennej separator�w akceptowanych przez program
    znaki = Array("-")
    
    
    '
    'pobieranie do zmiennej zaznaczonego obszaru przeszukiwa�
    Set search_range = Application.InputBox( _
        prompt:="Zaznacz obszar w kt�rym maj� zosta� podzielone i posortowane dane:", _
        Title:=tytul, _
        Type:=8)
    '
    'pobieranie do zmiennej obszaru od kt�rego ma nast�pi� wpisywanie rozbitego tekstu
    Set past_range = Application.InputBox( _
        prompt:="Zaznacz kom�rk� od kt�rej zostan� wstawione dane:", _
        Title:=tytul, _
        Type:=8)
        
    '
    odp = MsgBox("Chcesz przy�pieszy� dzia�anie programu?", _
                    vbInformation + vbYesNo + vbDefaultButton1, tytul)
    '
    If odp = vbYes Then
        Application.ScreenUpdating = False 'wy��czanie od�wie�ania ekranu na czas dzia�ania programu
    End If
    
    
    
    
    'pobranie czasu rozpocz�cia dzia�ania programu
    czas1 = GetTickCount
    
    
    '
    'g��wna p�tla programu
    For i = 1 To search_range.Rows.Count 'liczba wierszy w zaznaczonym obszarze
        '
        For j = 1 To search_range.Columns.Count 'liczba kolumn w zaznaczonym obszarze
        
            'pobieranie do zmiennej zawarto�ci kom�rki wskazanej przez zmienne i, j
            'z zaznczonego obszaru przeszukiwa�
            parse_text = search_range(i, j)
            
            'sprawdzanie czy kom�rk jest pusta, je�eli tak to j� pomija w przeszukiwaniu
            If parse_text <> "" Then
                'wywo�ywanie funkcij rozbijaj�cej tekst i zapisuj�cej poszczeg�lne liczby do list
                Call ParseText(znaki, parse_text, al)
                'wywo�ywanie procedury sortuj�cej zawarto�� listy
                Call al.Sort1(2, al.Get_Count)
                'wywo�. procedury wpisuj�cej zawarto�� list w poszczeg�lne kom�rki
                Call PastText(i, al, past_range)
                'czyszczenie listy
                al.Clear
            End If
            
        Next j
        
        '
        DoEvents

    Next i
    
    
    ''pobranie czasu rozpocz�cia dzia�ania programu
    czas2 = GetTickCount
    'obliczenia r�nicy. R�nic� nale�y podzieli� przez 1000 poniewa� czas jest mierzony w milisekundach
    czas = ((czas2 - czas1) / 1000)
    
    
    '
    If odp = vbYes Then
        Application.ScreenUpdating = True 'w��czenie od�wie�ania ekranu
    End If
    
    
    
    'wy�wietlenie ko�cowego komunikatu i czasu dzia�ania programu
    MsgBox "Koniec. Czas dzia�ania - " & czas & " s", vbInformation + vbOKOnly + vbDefaultButton1, tytul
    
    
    '
    'koniec
    '
        
End Sub

Private Function ParseText(ByVal zns, ByVal text As String, ByRef all As ArrayList) As Boolean
    '
    Dim nr As String, z As String
    Dim isnr As Boolean
    Dim i As Integer
    
    '
    isnr = False
    
    '
    For i = 1 To Len(text)
        '
        z = Mid(text, i, 1)
        
        '
        If IsNumeric(z) Or TestChars(zns, z) Then
            '
            isnr = True
            '
            nr = nr + z
        Else
            '
            If isnr Then
                '
                all.Add = nr
                '
                nr = ""
                isnr = False
                
            End If
            
        End If
        
    Next i

    
End Function

Private Function TestChars(ByVal znss, ByVal zn As String) As Boolean
    '
    Dim i As Integer
    
    '
    For i = 0 To UBound(znss)
        '
        If zn = znss(i) Then
            '
            TestChars = True
            '
            Exit Function
            
        End If
    
    Next i
    
    '
    TestChars = False

End Function

Private Sub PastText(ByVal row As Long, ByVal all As ArrayList, ByRef prg As Range)
    '
    Dim i As Integer
    
    '
    For i = 0 To all.Get_Count
        '
        prg(row, i + 1) = all.Get_Data(i)
    
    Next i
    
End Sub
