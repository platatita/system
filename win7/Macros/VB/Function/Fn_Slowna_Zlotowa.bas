Attribute VB_Name = "Fu_Slowna_Zlotowa"
Option Explicit

Dim Tlumaczenie As String, Tlumaczenie_GR As String
Dim Grosze As Integer, Zlotowki As Long, Zlotowa As String


Function Slowna_Zlotowa(Liczba As Range, Optional Tytul1 As String = " Liczba na s�owa") As String


' G��wnym i jedynym zadaniem tej funkcji jest przet�umaczenie ��danej przez urzytkownika
' liczby na ci�g s��w. Np. 1563 na "tysi�� pi��set sze��dziesi�t trzy z�ote".
' Funkcja ta jest stworzona w celu wsp�pracowanie z programami ksi�gowymi, poniewa�
' ka�d� pobran� liczb� traktuje je sume pieni�n� i dodaje do ka�dej liczby s�owa
' okre�laj�ce wysoko�� liczby w z�ot�wkach i groszach. Zaokr�gla ona pobran� liczbe
' do setnych grosza. Powy�ej lub r�wne 5 do g�ry poni�ej 5 na d�.


Dim LiczbaDOTl As Long, LiczbaDOTl1 As Double
Dim ZnakLiczby As String


'
LiczbaDOTl1 = Liczba.Value
LiczbaDOTl = 0
Grosze = 0
Tlumaczenie = ""
Tlumaczenie_GR = ""
Zlotowa = ""
ZnakLiczby = ""

'
If (LiczbaDOTl1 > -0.005 And LiczbaDOTl1 < 0) Then


    Tlumaczenie = "zero z�oty "
    Tlumaczenie_GR = "zero groszy"

Else


    'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
        
        '
        Call Sprawdzanie_Znaku_Liczby(LiczbaDOTl1, ZnakLiczby)
    
    'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
    
    
    'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
        
        '
        Call Zaokraglanie_Liczby(LiczbaDOTl1)
        
    'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
    
    
    'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
        
        '
        Call Grosze_(LiczbaDOTl1, LiczbaDOTl, Grosze)
    
    'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
    
    
    
    Select Case LiczbaDOTl1
        Case Is < 1
             
            If ZnakLiczby = "" Then
                
                Zlotowa = "zero z�oty "
            
            ElseIf ZnakLiczby = "minus " Then
                
                Zlotowa = ""
                
            End If
            
        Case 1 To 19
        
            'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
                
                '
                Call Nastki(LiczbaDOTl)
                
            'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
        
        Case 20 To 99
        
            'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
            
                '
                Call Dziesiatki(LiczbaDOTl)
            
            'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
    
        Case 100 To 999
        
            'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
            
                '
                Call Setki(LiczbaDOTl)
            
            'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
        
        Case 1000 To 9999
            
            'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
            
                '
                Call Tysiace(LiczbaDOTl)
                
            'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
            
        Case 10000 To 99999
        
            'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
            
                '
                Call DziesiatkiTysiecy(LiczbaDOTl)
                
            'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
            
        Case 100000 To 999999
            
            'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
            
                '
                Call SetkiTysiecy(LiczbaDOTl)
            
            'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
        
        Case 1000000 To 9999999
        
            'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
            
                '
                Call Miliony(LiczbaDOTl)
            
            'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
        
        Case 10000000 To 99999999
        
          'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
          
              '
              Call DziesiatkiMilionow(LiczbaDOTl)
        
          'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
        
        Case Is >= 100000000
            
            '
            Slowna_Zlotowa = "Sory kolego(kole�anko) ale mam za ma�y m�d�ek do takich liczba!!!"
            '
            GoTo koniec
            
    End Select
    
    
    
    'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
        
        '
        Call Zlotowki_(ZnakLiczby, LiczbaDOTl, LiczbaDOTl1, Grosze)
    
    'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP

End If



'GGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGG
'GGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGG

        '
        Slowna_Zlotowa = ZnakLiczby & Tlumaczenie & Zlotowa & Tlumaczenie_GR

'GGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGG
'GGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGG


koniec:


End Function

Private Sub Sprawdzanie_Znaku_Liczby(LiczbaDOTl1, ZnakLiczby)


'
Dim Wartosc As Integer


'
Wartosc = Sgn(LiczbaDOTl1)

    '
    If Wartosc = 1 Then
        
        ZnakLiczby = ""
        
    ElseIf Wartosc = -1 Then
        
        ZnakLiczby = "minus "
        
        'Abs()- ta funkcja zamienia znak liczby z ujemnego na dodatni.
        LiczbaDOTl1 = Abs(LiczbaDOTl1)
        
    End If


End Sub

Private Sub Zaokraglanie_Liczby(LiczbaDOTl1)

'
Dim Znak2 As Long
Dim wynik3 As String, Ulamki As String
Dim DL2 As Long, Wynik4 As Single, Wynik5 As Byte


'
Znak2 = 0
wynik3 = ""
DL2 = 0
Wynik4 = 0



'
Wynik4 = LiczbaDOTl1 - Fix(LiczbaDOTl1)
'
If Wynik4 = 0 Then
    
    Exit Sub

Else
    
    Znak2 = InStr(1, LiczbaDOTl1, ",")
    wynik3 = Mid(LiczbaDOTl1, Znak2 + 1, 3)
    
End If


'
DL2 = Len(wynik3)


If DL2 > 2 Then

    '
    If Val(Right(wynik3, 1)) < 5 Then
    
        Ulamki = "0," & Left(wynik3, 2)
        
    ElseIf Val(Right(wynik3, 1)) >= 5 Then
        
        '
        Wynik5 = Val(Mid(wynik3, 2, 1))
        
        Wynik5 = Wynik5 + 1
        
        Ulamki = "0," & Left(wynik3, 1) & CStr(Wynik5)
        
    End If

    '
    LiczbaDOTl1 = Fix(LiczbaDOTl1) + CDbl(Ulamki)

End If



End Sub

Private Sub Grosze_(LiczbaDOTl1, LiczbaDOTl, Grosze)

'
Dim DL As Long
Dim Znak1 As Integer
Dim wynik1 As Integer, grosze1 As String


'
Znak1 = 0
wynik1 = 0
DL = 0
grosze1 = ""


'
Znak1 = InStr(1, LiczbaDOTl1, ",")
DL = Len(CStr(LiczbaDOTl1))

'
If Znak1 = 0 Then
    
    Tlumaczenie_GR = "zero groszy "
    LiczbaDOTl = LiczbaDOTl1
    '
    Tlumaczenie = ""
    Exit Sub
    
End If


'
wynik1 = DL - Znak1
'
If wynik1 = 1 Then

    grosze1 = Right(LiczbaDOTl1, wynik1) & 0
    LiczbaDOTl = CInt(grosze1)

ElseIf wynik1 >= 2 Then

    grosze1 = Right(LiczbaDOTl1, wynik1)
    LiczbaDOTl = CInt(grosze1)

End If



'
Grosze = Val(grosze1)


If Grosze = 0 Then

    Tlumaczenie_GR = "zero groszy "
    GoTo koniec
    
ElseIf Grosze < 20 Then
    
    '
    LiczbaDOTl = Grosze

    '
    Call Nastki(LiczbaDOTl)
    '
    Tlumaczenie_GR = Tlumaczenie
    
ElseIf Grosze >= 20 And Grosze < 100 Then
    
    '
    LiczbaDOTl = Grosze

    '
    Call Dziesiatki(LiczbaDOTl)
    '
    Tlumaczenie_GR = Tlumaczenie
    
End If

'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
    '
    If Grosze = 1 Then
        
        Tlumaczenie_GR = Tlumaczenie_GR & "grosz"
        GoTo koniec
        
    End If
    
    

            If Len(CStr(Grosze)) >= 2 Then
                
                Grosze = Right(Grosze, 2)
                
                    If Grosze >= 12 And Grosze <= 14 Then
                        Grosze = 15
                    Else
                        Grosze = Right(Grosze, 1)
                    End If
                
            Else
                    Grosze = Right(Grosze, 1)
            End If


    '
    If Grosze >= 2 And Grosze <= 4 Then
        
        Tlumaczenie_GR = Tlumaczenie_GR & "grosze"
    
    Else
    
        Tlumaczenie_GR = Tlumaczenie_GR & "groszy"
        
    End If
'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW


koniec:

'
LiczbaDOTl = CLng(Left(LiczbaDOTl1, Znak1 - 1))
'
Tlumaczenie = ""


End Sub

Private Sub Zlotowki_(ZnakLiczby, LiczbaDOTl, LiczbaDOTl1, Grosze)

'

'
Zlotowki = Fix(LiczbaDOTl1)

If Zlotowki = 0 Then Exit Sub

'
If Len(CStr(Zlotowki)) >= 2 Then
    
    Zlotowki = Right(Zlotowki, 2)
    
        If Zlotowki >= 12 And Zlotowki <= 14 Then
            Zlotowki = 15
        Else
            Zlotowki = Right(Zlotowki, 1)
        End If
    
Else

    Zlotowki = Right(Zlotowki, 1)
    
End If


'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
    '
    If Zlotowki = 0 Then
    
            If Zlotowa = "" Then
            
                Zlotowa = "z�oty "
                
            End If
                   
    ElseIf Zlotowki = 1 Then
    
        Zlotowa = Zlotowa & "z�oty "
        
    ElseIf Zlotowki >= 2 And Zlotowki <= 4 Then
        
        Zlotowa = Zlotowa & "z�ote "
    
    Else
    
        Zlotowa = Zlotowa & "z�oty "
        
    End If
'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW



End Sub

Private Sub Nastki(LiczbaDOTl)

'

Select Case LiczbaDOTl
    Case 1
        Tlumaczenie = Tlumaczenie & "jeden "
    Case 2
        Tlumaczenie = Tlumaczenie & "dwa "
    Case 3
        Tlumaczenie = Tlumaczenie & "trzy "
    Case 4
        Tlumaczenie = Tlumaczenie & "cztery "
    Case 5
        Tlumaczenie = Tlumaczenie & "pi�� "
    Case 6
        Tlumaczenie = Tlumaczenie & "sze�� "
    Case 7
        Tlumaczenie = Tlumaczenie & "siedem "
    Case 8
        Tlumaczenie = Tlumaczenie & "osiem "
    Case 9
        Tlumaczenie = Tlumaczenie & "dziewi�� "
    Case 10
        Tlumaczenie = Tlumaczenie & "dziesi�� "
    Case 11
        Tlumaczenie = Tlumaczenie & "jedena�cie "
    Case 12
        Tlumaczenie = Tlumaczenie & "dwana�cie "
    Case 13
        Tlumaczenie = Tlumaczenie & "trzyna�cie "
    Case 14
        Tlumaczenie = Tlumaczenie & "czterna�cie "
    Case 15
        Tlumaczenie = Tlumaczenie & "pi�tna�cie "
    Case 16
        Tlumaczenie = Tlumaczenie & "szesna�cie "
    Case 17
        Tlumaczenie = Tlumaczenie & "siedemna�cie "
    Case 18
        Tlumaczenie = Tlumaczenie & "osiemna�cie "
    Case 19
        Tlumaczenie = Tlumaczenie & "dziewi�tna�cie "
End Select


End Sub

Private Sub Dziesiatki(LiczbaDOTl)

'
Dim Liczba As Integer


'
Liczba = Left(LiczbaDOTl, 1) & 0

Select Case Liczba
    Case 10
        Tlumaczenie = Tlumaczenie & "dziesi�� "
    Case 20
        Tlumaczenie = Tlumaczenie & "dwadzie�cia "
    Case 30
        Tlumaczenie = Tlumaczenie & "trzydzie�ci "
    Case 40
        Tlumaczenie = Tlumaczenie & "czterdzie�ci "
    Case 50
        Tlumaczenie = Tlumaczenie & "pi��dziesi�t "
    Case 60
        Tlumaczenie = Tlumaczenie & "sze��dziedzi�t "
    Case 70
        Tlumaczenie = Tlumaczenie & "siedemdziesi�t "
    Case 80
        Tlumaczenie = Tlumaczenie & "osiemdziesi�t "
    Case 90
        Tlumaczenie = Tlumaczenie & "dziewi��dziesi�t "
End Select


'
LiczbaDOTl = Right(LiczbaDOTl, 1)

If LiczbaDOTl = 0 Then
    Exit Sub
Else
    '
    Call Nastki(LiczbaDOTl)
End If



End Sub

Private Sub Setki(LiczbaDOTl)

'
Dim Liczba As Integer, DL As Integer


'
Liczba = LiczbaDOTl

'
LiczbaDOTl = Left(Liczba, 1) & "00"


Select Case LiczbaDOTl
    Case 100
        Tlumaczenie = Tlumaczenie & "sto "
    Case 200
        Tlumaczenie = Tlumaczenie & "dwie�cie "
    Case 300
        Tlumaczenie = Tlumaczenie & "trzysta "
    Case 400
        Tlumaczenie = Tlumaczenie & "czterysta "
    Case 500 To 900
                
        '
        LiczbaDOTl = Left(Liczba, 1)
        
        '
        Call Nastki(LiczbaDOTl)
        
        DL = Len(Tlumaczenie)
        Tlumaczenie = Left(Tlumaczenie, DL - 1)
        Tlumaczenie = Tlumaczenie & "set "
        
End Select
    
  
'
LiczbaDOTl = Right(Liczba, 2)

'
If LiczbaDOTl = 0 Then

    Exit Sub
    
ElseIf LiczbaDOTl < 20 Then
    '
    Call Nastki(LiczbaDOTl)
Else
    '
    Call Dziesiatki(LiczbaDOTl)
End If


End Sub

Private Sub Tysiace(LiczbaDOTl)

'
Dim Liczba As Integer, DL As Integer

'
Liczba = LiczbaDOTl

'
LiczbaDOTl = Left(Liczba, 1) & "000"


Select Case LiczbaDOTl
    Case 1000
        Tlumaczenie = Tlumaczenie & "tysi�� "
    Case 2000
        Tlumaczenie = Tlumaczenie & "dwa tysi�ce "
    Case 3000
        Tlumaczenie = Tlumaczenie & "trzy tysi�ce "
    Case 4000
        Tlumaczenie = Tlumaczenie & "cztery tysi�ce "
    Case 5000 To 9000
                
        '
        LiczbaDOTl = Left(Liczba, 1)
        
        '
        Call Nastki(LiczbaDOTl)
        
        '
        Tlumaczenie = Tlumaczenie & "tysi�cy "
        
        
End Select
    

'
DL = Len(CStr(Liczba))
LiczbaDOTl = Right(Liczba, DL - 1)

If LiczbaDOTl = 0 Then
    Exit Sub
ElseIf LiczbaDOTl < 20 Then
    '
    Call Nastki(LiczbaDOTl)
ElseIf LiczbaDOTl >= 20 And LiczbaDOTl < 100 Then
    '
    Call Dziesiatki(LiczbaDOTl)
Else
    '
    Call Setki(LiczbaDOTl)
End If


End Sub

Private Sub DziesiatkiTysiecy(LiczbaDOTl)

'
Dim Liczba As Long, DL As Integer

'
Liczba = LiczbaDOTl

'
LiczbaDOTl = Left(Liczba, 2)

If LiczbaDOTl < 20 Then
    '
    Call Nastki(LiczbaDOTl)
Else
    '
    Call Dziesiatki(LiczbaDOTl)
End If

'
Tlumaczenie = Tlumaczenie & "tysi�cy "

 
 
'
DL = Len(CStr(Liczba))
LiczbaDOTl = Right(Liczba, DL - 2)


If LiczbaDOTl = 0 Then
    Exit Sub
ElseIf LiczbaDOTl < 20 Then
    '
    Call Nastki(LiczbaDOTl)
ElseIf LiczbaDOTl >= 20 And LiczbaDOTl < 100 Then
    '
    Call Dziesiatki(LiczbaDOTl)
Else
    '
    Call Setki(LiczbaDOTl)
    
End If


End Sub

Private Sub SetkiTysiecy(LiczbaDOTl)

'
Dim Liczba As Long, DL As Integer

'
Liczba = LiczbaDOTl

'
LiczbaDOTl = Left(Liczba, 3)


'
Call Setki(LiczbaDOTl)

'
Tlumaczenie = Tlumaczenie & "tysi�cy "

       
'
DL = Len(CStr(Liczba))
LiczbaDOTl = Right(Liczba, DL - 3)

If LiczbaDOTl = 0 Then
    Exit Sub
ElseIf LiczbaDOTl < 20 Then
    '
    Call Nastki(LiczbaDOTl)
ElseIf LiczbaDOTl >= 20 And LiczbaDOTl < 100 Then
    '
    Call Dziesiatki(LiczbaDOTl)
ElseIf LiczbaDOTl >= 100 And LiczbaDOTl < 1000 Then
    '
    Call Setki(LiczbaDOTl)
Else
    '
    DziesiatkiTysiecy (LiczbaDOTl)
End If


End Sub

Private Sub Miliony(LiczbaDOTl)

'
'
Dim Liczba As Long, DL As Integer

'
Liczba = LiczbaDOTl

'
LiczbaDOTl = Left(Liczba, 1)

Select Case LiczbaDOTl
    Case 1
        Tlumaczenie = Tlumaczenie & "jeden milion "
    Case 2 To 4
        '
        Call Nastki(LiczbaDOTl)
        '
        Tlumaczenie = Tlumaczenie & "miliony "
    Case 5 To 9
        '
        Call Nastki(LiczbaDOTl)
        '
        Tlumaczenie = Tlumaczenie & "milion�w "
End Select

DL = Len(CStr(Liczba))
LiczbaDOTl = Right(Liczba, DL - 1)

If LiczbaDOTl = 0 Then
    Exit Sub
ElseIf LiczbaDOTl < 20 Then
    '
    Call Nastki(LiczbaDOTl)
ElseIf LiczbaDOTl >= 20 And LiczbaDOTl < 100 Then
    '
    Call Dziesiatki(LiczbaDOTl)
ElseIf LiczbaDOTl >= 100 And LiczbaDOTl < 1000 Then
    '
    Call Setki(LiczbaDOTl)
ElseIf LiczbaDOTl >= 1000 And LiczbaDOTl < 10000 Then
    '
    Call Tysiace(LiczbaDOTl)
ElseIf LiczbaDOTl >= 10000 And LiczbaDOTl < 100000 Then
    '
    Call DziesiatkiTysiecy(LiczbaDOTl)
Else
    '
    Call SetkiTysiecy(LiczbaDOTl)
End If


End Sub

Private Sub DziesiatkiMilionow(LiczbaDOTl)

'
'
Dim Liczba As Long, DL As Integer

'
Liczba = LiczbaDOTl

'
LiczbaDOTl = Left(Liczba, 2)

If LiczbaDOTl < 20 Then
    '
    Call Nastki(LiczbaDOTl)
ElseIf LiczbaDOTl >= 20 And LiczbaDOTl < 100 Then
    '
    Call Dziesiatki(LiczbaDOTl)
End If

'
Tlumaczenie = Tlumaczenie & "milion�w "



DL = Len(CStr(Liczba))
LiczbaDOTl = Right(Liczba, DL - 2)

If LiczbaDOTl = 0 Then
    Exit Sub
ElseIf LiczbaDOTl < 20 Then
    '
    Call Nastki(LiczbaDOTl)
ElseIf LiczbaDOTl >= 20 And LiczbaDOTl < 100 Then
    '
    Call Dziesiatki(LiczbaDOTl)
ElseIf LiczbaDOTl >= 100 And LiczbaDOTl < 1000 Then
    '
    Call Setki(LiczbaDOTl)
ElseIf LiczbaDOTl >= 1000 And LiczbaDOTl < 10000 Then
    '
    Call Tysiace(LiczbaDOTl)
ElseIf LiczbaDOTl >= 10000 And LiczbaDOTl < 100000 Then
    '
    Call DziesiatkiTysiecy(LiczbaDOTl)
Else
    '
    Call SetkiTysiecy(LiczbaDOTl)
End If


End Sub

