Attribute VB_Name = "HorasDec"
Option Explicit


'Dada una hora devuelve un numero en formato Single
Public Function DevuelveValorHora(vHora As Date) As Single
Dim Aux As Single
Dim C

    C = Minute(vHora)
    Aux = C / 60
    DevuelveValorHora = Hour(vHora) + Round(Aux, 2)
End Function





'2017
'  0. Normal
'  1- Hora menor que 0
'  2- Hora mayor= que 24
Public Function DevuelveValorHora3(FueraIntervalo As Byte, vHora As String) As Single
Dim Aux As Single
Dim C
Dim H As Date
Dim Incre As Integer
        
        
    If FueraIntervalo = 1 Then
        'negativo
        H = vHora
        C = Minute(H)
        Aux = C / 60
        DevuelveValorHora3 = Hour(H) + Round(Aux, 2)
        
        DevuelveValorHora3 = -DevuelveValorHora3
           
    Else
        If FueraIntervalo = 2 Then
            Aux = Mid(vHora, 1, 2)
            Aux = Aux - 24
            Incre = 24
            H = CDate(Format(Val(Aux), "00") & Mid(vHora, 3))
        Else
            'Normal
            H = CDate(vHora)
            Incre = 0
        End If
         C = Minute(H)
        Aux = C / 60
        DevuelveValorHora3 = Hour(H) + Incre + Round(Aux, 2)
    End If
    
   
End Function

'Una hora formateada en formato "hh:mm" nos dira si hour() es menor que cero o mayor que 23
Public Function HoraFueraInterval(CadenaHora As String) As Byte
    Dim k As Integer
Dim H As Integer
    
    k = InStr(1, CadenaHora, ":")
    H = Mid(CadenaHora, 1, k - 1)
    HoraFueraInterval = 0
    If H < 0 Then
        HoraFueraInterval = 1
    Else
        If H > 23 Then HoraFueraInterval = 2
    End If
            
End Function



'Dada una hora en centesimal la pasamos a formato hora
Public Function DevuelveHora(vHora As Single) As Date
Dim X
Dim Y
Dim cad As String

    vHora = Abs(Round(vHora, 2))
    X = Int(vHora)
    Y = vHora - X
    'En y esta la parte centesimal de una hora
    Y = Round(Y * 60, 0)
    X = X Mod 24
    cad = X & ":" & Y & ":00"
    DevuelveHora = CDate(cad)
End Function




Public Function DiasMes(m As Integer, anyo As Integer) As Integer
    
    Select Case m
    Case 1, 3, 5, 7, 8, 10, 12
        DiasMes = 31
    Case 2
        If (anyo Mod 4) = 0 Then
            DiasMes = 29
        Else
            DiasMes = 28
        End If
    Case Else
        DiasMes = 30
    End Select
End Function




'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaPuntosHoras(CADENA As String) As String
Dim i As Integer
Do
    i = InStr(1, CADENA, ".")
    If i > 0 Then
        CADENA = Mid(CADENA, 1, i - 1) & ":" & Mid(CADENA, i + 1)
    End If
    Loop Until i = 0
TransformaPuntosHoras = CADENA
End Function



Public Function Horas_Quitar24(Hora As Date, QuitarSigno As Boolean) As String
Dim Resultado As String
Dim Num As Integer
Dim R As Integer
    
    

    Resultado = ""
    'Minutos
    If Second(Hora) = 0 Then
        Resultado = "00"
        R = 0
    Else
        Resultado = Format(60 - Second(Hora), "00")
        R = 1
    End If
    Resultado = ":" & Resultado
    
    Num = Minute(Hora) + R
    If Num = 0 Then
        R = 0
    Else
        R = 60 - Num
    End If
    Resultado = ":" & Format(R, "00") & Resultado
    If R > 0 Then R = 1
    Num = Hour(Hora) + R
    If R > 23 Then Err.Raise 513, , "Campo hora mayor que 24"
        
    Num = 24 - Num
    Horas_Quitar24 = ""
    If Not QuitarSigno Then Horas_Quitar24 = "-"
    Horas_Quitar24 = Horas_Quitar24 & Format(Num, "00") & Resultado
    
End Function
