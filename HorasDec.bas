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





'Enero 2015
'Dada una hora devuelve un numero en formato Single
Public Function DevuelveValorHora2(FueraIntervalo As Boolean, vHora As String) As Single
Dim Aux As Single
Dim C
Dim H As Date
Dim Incre As Integer
    
    If FueraIntervalo Then
        'Demomento es +24
        Aux = Mid(vHora, 1, 2)
        If Aux > 23 Then
            Aux = Aux - 24
            Incre = 24
        Else
            Stop
        End If
        H = CDate(Format(Val(Aux), "00") & Mid(vHora, 3))
    Else
        H = CDate(vHora)
        Incre = 0
    End If
    
    C = Minute(H)
    Aux = C / 60
    DevuelveValorHora2 = Hour(H) + Incre + Round(Aux, 2)
End Function

'Una hora formateada en formato "hh:mm" nos dira si hour() es menor que cero o mayor que 23
Public Function HoraFueraIntervalo(CadenaHora As String) As Boolean
Dim k As Integer
Dim H As Integer
    
    k = InStr(1, CadenaHora, ":")
    H = Mid(CadenaHora, 1, k - 1)
    HoraFueraIntervalo = H < 0 Or H > 23
End Function



'Dada una hora en centesimal la pasamos a formato hora
Public Function DevuelveHora(vHora As Single) As Date
Dim X
Dim Y
Dim Cad As String

    vHora = Abs(Round(vHora, 2))
    X = Int(vHora)
    Y = vHora - X
    'En y esta la parte centesimal de una hora
    Y = Round(Y * 60, 0)
    X = X Mod 24
    Cad = X & ":" & Y & ":00"
    DevuelveHora = CDate(Cad)
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
Dim I As Integer
Do
    I = InStr(1, CADENA, ".")
    If I > 0 Then
        CADENA = Mid(CADENA, 1, I - 1) & ":" & Mid(CADENA, I + 1)
    End If
    Loop Until I = 0
TransformaPuntosHoras = CADENA
End Function
