Attribute VB_Name = "Fun_Kombinacje"
Option Explicit

Function Funkcja_Kombinacje(Zakres_liczbowy As Integer, D�ugo��_ci�gu As Integer) As Variant

' Funkcja ta oblicza liczb� kombinacji dla danego zakresu liczbowego o danej
' d�ugo�ci ci�gu. Np dla liczb od 1 do 80 dla ci�gu d�ugo�ci 3 lub 10 liczb.
' Wystarczy poda� zakres liczbowy i d�ugo�� ci�gu.


Dim mianownik1 As Integer
Dim wynikdul As Variant
Dim wynikgora As Variant

Dim nd As Integer
Dim ng As Integer

  
  mianownik1 = Zakres_liczbowy - D�ugo��_ci�gu + 1
  wynikdul = 1
  wynikgora = 1
  
    For nd = 1 To D�ugo��_ci�gu
      wynikdul = wynikdul * nd
    Next nd
    
    For ng = mianownik1 To Zakres_liczbowy
      wynikgora = wynikgora * ng
    Next ng
          
      
     Funkcja_Kombinacje = wynikgora / wynikdul

End Function

