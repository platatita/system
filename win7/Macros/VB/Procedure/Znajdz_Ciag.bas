Attribute VB_Name = "Znajdz_Ciag"
Option Explicit
Option Base 1


Dim tytul As String
Dim WejscieZPetli As Boolean
Dim P As Long
Dim zakres As Range
Dim Wiersze As Long, Kolumny As Long
Dim Razy As Integer
Dim SciezkaZapisu As String


' DlugoscCiagu - liczba cyfr, z kt�rych b�dzie sk�ada� si� ci�g liczbowy
' np: 10 oznacza, �e ci�g b�dzie sk�ada� si� z 10 liczb(1,2,3,4,5,6,7,8,9,10).
Const DlugoscCiagu = 3
' maxNumer - zmienna ta zawiera liczb� cyfr do kombinacji ci�g�w
' np: 15 oznacza, �e kombinacja b�dzie tworzona z 15 liczb(1-15).
Const maxNumer = 80


' Tablica zawieraj�ca liczb� cyfr z kt�rych b�dzie sk�ada� si� ci�g liczbowy.
Dim InTablica() As Integer
' Tablica flagi ma sygnalizowa� kt�ra liczba osi�gne�a maxNumer i przej�� do
' pozycji wcze�niejszej, aby jej warto�� zwi�kszy� o jeden.
Dim Flagi() As Integer


Public Sub Znajdz_Ciag()

' Procedur ta ma za zadanie przeszuka� zaznaczony fragm�t tabeli i zlicy� ile razy
' wystepuje w niej dany ci�g liczbowy.

'
tytul = "Inicjalizacja"


'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP

        ' wywo�uje procedur� SprawdzenieZaznaczenia()
        Call SprawdzenieZaznaczenia

'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP


'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP

        ' wywo�uje procedur� TworzPlik()
        Call TworzPlik

'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP


Load UserForm1
UserForm1.Show


'--------------------------------------------------------------------------------------

Dim i As Integer, pozycja As Integer
Dim EndPetli As Integer


'
WejscieZPetli = False


' Ustawianie rozmiaru tablicy
ReDim InTablica(DlugoscCiagu) As Integer
' Ustawianie rozmiaru tablicy
ReDim Flagi(DlugoscCiagu) As Integer


' Pierwsze wype�nianie tablic Flagi i InTablica. Elementy tablicy Flagi s� ustawiane
' na "0" a elementy tablicy InTablica na 1,2...
    For i = 1 To DlugoscCiagu
               
        InTablica(i) = i
        Flagi(i) = 0
    
    Next i
    
    
            ' Tu wstaw odwo�anie do szukania SzukajPodanegoCiagu
            ' wywo�uje procedur� SzukajPodanegoCiagu()
            Call SzukajPodanegoCiagu

    ' pozycja - zmienna ta ustawia w kt�rym miejscu(na kt�rym elemencie tablicy "InTablica")
    ' ma si� rozpocz�� praca p�tli.
    pozycja = DlugoscCiagu
    
    ' EndPetli - zmienna ta sugnalizuje kiedy ma si� zakonczy� wykonywanie p�tli.
    EndPetli = maxNumer - pozycja + 1
    
    
petla:

    Do While InTablica(1) < EndPetli
    
        ' zwi�ksza warto�c danego elementu tablicy o jeden "1".
        InTablica(pozycja) = InTablica(pozycja) + 1
                
        If InTablica(pozycja) = (maxNumer + pozycja - (DlugoscCiagu)) Then Flagi(pozycja) = 1: _
                pozycja = pozycja - 1: WejscieZPetli = True: Call SzukajPodanegoCiagu: GoTo petla ': Call SzukajPodanegoCiagu
        
        
        
        If WejscieZPetli = True Then
        
            For i = 1 To DlugoscCiagu
                         
              If Flagi(i) = 1 Then Flagi(i) = 0: pozycja = pozycja + 1: _
                      InTablica(i) = InTablica(i - 1) + 1
                      
            Next i
        
            WejscieZPetli = False
        
        End If
        
            ' Tu wstaw odwo�anie do szukania SzukajPodanegoCiagu
            ' wywo�uje procedur� SzukajPodanegoCiagu()
            Call SzukajPodanegoCiagu
            
            DoEvents

    Loop

' Zatrzymanie wykonywania programu.
Stop


Dim Odp As Integer

Odp = msgbox("Koniec programu" & "  " & P, vbInformation, tytul)


End Sub

Private Sub SzukajPodanegoCiagu()

' Przeszukiwanie tablicy zaznaczon� tablic� szukaj�� danego ci�gu.

Dim i1 As Long, i2 As Long, i As Byte
Dim ciag As Byte


P = P + 1: UserForm1.Label1.Caption = P

For i1 = 1 To Wiersze

    For i2 = 1 To Kolumny


            For i = 1 To DlugoscCiagu
                
                If zakres(i1, i2) = InTablica(i) Then ciag = ciag + 1
                If ciag = 0 And i2 > (Kolumny - DlugoscCiagu) Then GoTo line1
                If zakres(i1, 1) > InTablica(1) Then GoTo line1
                
            Next i
                        
    
    Next i2

line1:

    
    If ciag = DlugoscCiagu Then Razy = Razy + 1
    
    ciag = 0
    
Next i1


    'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
    
            ' wywo�uje procedur� Zapisywanie()
            Call Zapisywanie
    
    'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP

End Sub

Private Sub Zapisywanie()

' Procedura ta zapisuje wynik do pliku tekstowego.

Dim LiczbaIni As String, LiczbaIni1 As String

Dim FSO As New FileSystemObject
Dim texktStream As TextStream
Dim TextFile As File
Dim i As Byte

Set TextFile = FSO.GetFile(SciezkaZapisu)

Set texktStream = TextFile.OpenAsTextStream(ForAppending)


For i = 1 To DlugoscCiagu

    LiczbaIni = LiczbaIni & CStr(InTablica(i)) & ", "
    
Next i


If Razy < DlugoscCiagu Then Razy = 0

LiczbaIni1 = LiczbaIni & "  " & " =  " & Razy & " razy"

' zapisywanie do pliku tekstowego wyniku szukania.
texktStream.WriteLine CStr(LiczbaIni1)


' zamykanie pliku *.txt
texktStream.Close

End Sub

Private Sub SprawdzenieZaznaczenia()

' Procedura ta sprawdza, czy zosta� zaznaczony obszar do przeszukania,
' je�eli tak to pobiera jego rozmiary.

Dim Odp As Integer
Dim komunikat As String, KomunikatInput As String



    ' Sprawdzanie, czy zosta� zaznaczony jaki� obszar. Je�eli nie to ostrze�eni
    ' albo koniec programu.
    If IsArray(Selection) Then
        
        ' Przypisanie zmiennej zakres w�a�ciwo�ci Range, czyli zaznaczonego obszaru.
        Set zakres = Selection
        
    Else
        
        
        komunikat = "Zaznacz obszar, kt�ry chcesz przeszukiwa�." & Chr(13) & Chr(13)
        komunikat = komunikat & "Chcesz teraz zaznaczy� ten obszar?" & Chr(13) & Chr(13)

        Odp = msgbox(komunikat, vbExclamation + vbYesNo, tytul)
        
            If Odp = vbYes Then
                
                KomunikatInput = "Zaznacz obszar, kt�ry ma zosta� przeszukany"

                ' Przypisanie zmiennej zakres w�a�ciwo�ci Range, czyli zaznaczonego obszaru.
                Set zakres = Application.InputBox(prompt:=KomunikatInput, Title:=tytul, Type:=8)
                
            ElseIf Odp = vbNo Then
                
                End
                
            End If
        
    
    End If



Wiersze = zakres.Rows.Count
Kolumny = zakres.Columns.Count


End Sub

Private Sub TworzPlik()

' Procedura ta ma zadanie utworzy� plik tekstowy je�eli jeszcze  nie istnieje.

Dim FSO As New FileSystemObject
Dim txtStream As TextStream
Dim CzyIstniejPlik  As Boolean, CzyIstniejPlikRand As Boolean
Dim LiczbaRandomize  As Integer, SciezkaZapisuRand As String

Dim DlSciezki As Integer
Dim i As Integer


' Ustawienie warto�� pocz�tkowych
CzyIstniejPlikRand = False
CzyIstniejPlik = False


' Scie�ka dost�pu do pliku i jego zapisu.
SciezkaZapisu = "C:\M�j szukaj\M�j Szukaj.txt"

    
    CzyIstniejPlik = FSO.FileExists(SciezkaZapisu)

    If CzyIstniejPlik = False Then
    
            ' Tworzenie pliku tekstowego o podanej nazwie i w podanym miejscu.
            Set txtStream = FSO.CreateTextFile(SciezkaZapisu, False)
            
            
            ' zamykanie pliku tekstowego
            txtStream.Close
            
    ElseIf CzyIstniejPlik = True Then

            Dim NrPliku As Integer
            NrPliku = 1000
            
            DlSciezki = Len(SciezkaZapisu) - 4
            
            SciezkaZapisuRand = Left(SciezkaZapisu, DlSciezki)
            
            
            For i = 1 To NrPliku
                
                
                    LiczbaRandomize = i
                    
                    SciezkaZapisuRand = SciezkaZapisuRand & CStr(LiczbaRandomize) & ".txt"
                    
                    ' Sprawdzanie, czy jstnieje nast�puj�cy plik.
                    CzyIstniejPlikRand = FSO.FileExists(SciezkaZapisuRand)
                    
                    
                    ' Je�eli tak!
                    If CzyIstniejPlikRand = True Then
                        
                        SciezkaZapisuRand = Left(SciezkaZapisu, DlSciezki)

                        GoTo line1
                    ' Je�eli nie!
                    Else
                        
                        SciezkaZapisu = SciezkaZapisuRand
                        
                        Exit For
                    
                    End If
line1:
                    
            Next i
                 
            ' Je�eli wyst�pi b��d to program przechodzi do nast�pnej lini.
            On Error Resume Next
            
            ' Tworzenie pliku tekstowego o podanej nazwie i w podanym miejscu.
            Set txtStream = FSO.CreateTextFile(SciezkaZapisu, False)
            
            ' Odczytywanie kodu b��du.
            If Err.Number = 58 Then
                
                Dim KomunikatError As String
                
                    KomunikatError = "Nie mo�na utworzy� nowego pliku tekstowego, poniewa�" & Chr(13)
                    KomunikatError = KomunikatError & "utworzy�e� ich ju� ponad  " & NrPliku & "  w tym folderze" & Chr(13)
                    KomunikatError = KomunikatError & "i musisz zmieni� folder lub usun�� nikt�re pliki tekstowe," & Chr(13)
                    KomunikatError = KomunikatError & "aby rozpocz�� zapisywanie danych do pliku. " & Chr(13) & Chr(13) & Chr(13)
                
                    msgbox KomunikatError, vbExclamation, tytul
                    
                    End
            End If
            
            
            ' zamykanie pliku tekstowego
            txtStream.Close
   
    
    End If


End Sub
