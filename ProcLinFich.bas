Attribute VB_Name = "ProcLinFich"
Option Explicit

'Esta libreria es para procesar las lineas tal y como vienen del
'reloj.
'Por lo tanto este modulo depende, totalmente, del reloj
'
'LINEA= linea del fichero de texto

Public Sub ProcesarLinea(Linea As String, Contador As Long, anyo As Integer, segundos As Integer)
Dim I As Integer
Dim vector(4) As String
Dim Rs As ADODB.Recordset
Dim LError As String

On Error GoTo ErrorProcesandoLinea
        For I = 1 To 4
            vector(I) = ""
        Next I
        'Separamos los campos segun presencia en TCP 3
        'Ejemplo de linea tcp3
        ' tar  mes dia hora minut nada inci nada
        '01234,11,23,08,20,0000,0000,18411
        'FECHA
        vector(0) = Mid(Linea, 10, 2) & "/" & Mid(Linea, 7, 2) & "/" & anyo
        'Hora
        vector(1) = Mid(Linea, 13, 2) & ":" & Mid(Linea, 16, 2) & ":" & Format(segundos, "00")
        'operario
        vector(2) = Mid(Linea, 1, 5)
        'seccion
        'vector(3) =   Mid(Linea, 26, 3)
        'tecla
        vector(4) = Mid(Linea, 24, 4)
        
        'Ahora insertamos en la BD
        Set Rs = New ADODB.Recordset
        Rs.CursorType = adOpenKeyset
        Rs.LockType = adLockOptimistic
        Rs.Open "TemporalFichajes", conn, , , adCmdTable
        Rs.AddNew
        Rs!Secuencia = Contador
        Rs!Numtarjeta = vector(2)
        Rs!Fecha = vector(0)
        Rs!Hora = vector(1)
        Rs!idInci = vector(4)
        Rs.Update
        Rs.Close
    
        Exit Sub
ErrorProcesandoLinea:
    EscribeErrorLinea "Error: " & vbCrLf & Linea & vbCrLf & Err.Number & " - " & Err.Description
End Sub


'Estamos tratando ahora marcajes del tipo de
'la cooperativa de ALZIRA.
Public Sub ProcesarLineaALZ(Linea As String, Contador As Long, PuntoInicio As Integer)
Dim I As Integer
Dim vector(4) As String
Dim Rs As ADODB.Recordset
Dim LError As String
Dim Longitud As Integer

On Error GoTo ErrorProcesandoLinea
For I = 1 To 4
    vector(I) = ""
Next I

'Separamos los campos segun presencia en ficheros produccion
'Ejemplo de antes de NOVIEMBRE DE 2002
'02 2001/11/01 06:24:13 0030 233 079 000

'Ejemplo actual
'    011600211071619140000021ILO
'   de donde
'   > Numero de empleado  5                 --> 01160
'   > Ano 2                                 --> 02
'   > Mes 2                                 --> 11
'   > Dia 2                                 --> 07
'   > Hora 6                                --> 161914
'   > Numero de reloj/terminal 6
'   > Datos control (s/significado)      16
    
''''''''''''''------------------------------ ANTES
''''''''''''''FECHA
'''''''''''''vector(0) = Mid(Linea, 4, 10)
''''''''''''''Hora
'''''''''''''vector(1) = Mid(Linea, 15, 8)
''''''''''''''tarjeta
'''''''''''''vector(2) = Mid(Linea, 24, 4)
''''''''''''''seccion
'''''''''''''vector(3) = Mid(Linea, 29, 3)
''''''''''''''tecla
'''''''''''''vector(4) = Mid(Linea, 33, 3)
Longitud = 6 - PuntoInicio
'------------------------------ AHORA
'tarjeta
vector(2) = Mid(Linea, PuntoInicio, Longitud)
'FECHA
vector(0) = "20" & Mid(Linea, 6, 2) & "/" & Mid(Linea, 8, 2) & "/" & Mid(Linea, 10, 2)     'Le añadimos el 20 para que sea 2002
'Hora
vector(1) = Mid(Linea, 12, 2) & ":" & Mid(Linea, 14, 2) & ":" & Mid(Linea, 16, 2)
'seccion
vector(3) = 0
'tecla
vector(4) = 0



'ANTIGUOS
'Segun los parametros, si las fechas van con asteriscos hay
'que despreciarlas
'i = InStr(1, vector(0), "*")

'AHora
I = 0
If I = 0 Then
    'La fecha es correcta.
    'Los parametros dicen que cuando el codig de operario es
    '9001,9002,9003,9004,9005 se desprecia
    I = DespreciarMarcaje(vector(2))
    If I = 1 Then Exit Sub
    
    'llegados a este punto insertamos en la BD
    Set Rs = New ADODB.Recordset
    Rs.CursorType = adOpenKeyset
    Rs.LockType = adLockOptimistic
    Rs.Open "TipoAlzicoop", conn, , , adCmdTable
    Rs.AddNew
    Rs!Secuencia = Contador
    Rs!Tarjeta = vector(2)
    Rs!Fecha = Format(vector(0), "dd/mm/yyyy")
    
    Rs!Hora = vector(1)
    Rs!Seccion = vector(3)
    Rs!tecla = vector(4)
    
    'Hora real
    'Modificacion del 22 Julio 2004
    Rs!HoraReal = Rs!Hora
    Rs.Update
    Rs.Close
    Set Rs = Nothing
End If
Exit Sub
ErrorProcesandoLinea:
    EscribeErrorLinea "Error: " & vbCrLf & Linea & vbCrLf & Err.Number & " - " & Err.Description
End Sub



Private Function DespreciarMarcaje(CadenaOperario As String) As Integer
'Esto es pq antes, en ALZIRA, los marcajes llegaban desde produccion, con lo cual
'habia que despreciar los ticajes de una deteriminada forma
'Y eran aquellos que los operarios eran 9000 y demas
'Select Case CadenaOperario
'Case "9001", "9002", "9003", "9004", "9005"
'    DespreciarMarcaje = 1
'Case Else
'    DespreciarMarcaje = 0
'End Select
DespreciarMarcaje = 0
End Function



Public Function TransformaLineaRobotics(CADENA As String, ByRef ElAnyo As Integer) As String
Dim C As String

    'Se trata de a partir de la cadena de ROBOTICS
    'GENERO LA CADENA DE TCP que es la que trabajaremos
    
    'TCP3:      01234,11,23,08,20,0000,0000,18411
    
    '           1   5    0    5    0    5
    'ROBOTICS:  " 1110401OUT2000000003118.34
    '             ^             Terminal
    '              ddmmyy
    '                    In/OUT para nosotros irrelevante
    '                               ttt ---> Trabajador  (Sera nuestra tarjeta)
    '                                  hh.mm
    
    
    'SI tiene incidencia viene asin
    '           "  1020801SF22000000001811.5411
    '              ddmmyy                    codinci (11)
    

    
    '                     SF  Inci manual
    '                               ttt ---> Trabajador  (Sera nuestra tarjeta)
    '                                  hh.mm
    

       
        
        
    C = Mid(CADENA, 18, 5) & ","
    C = C & Mid(CADENA, 5, 2) & "," & Mid(CADENA, 3, 2) & ","
    ElAnyo = CInt("20" & Mid(CADENA, 7, 2))
    C = C & Mid(CADENA, 23, 2) & "," & Mid(CADENA, 26, 2)  'HORA
    C = C & ",0000,"
    
    If InStr(1, CADENA, "F") > 0 Then
        'Lleva INCIDENCIA MANUAL
        
        C = C & Format(Val(Mid(CADENA, 28)), "0000")
    Else
        'NO LLEVA inci
        C = C & "0000"
    End If
    C = C & ",12345"
    TransformaLineaRobotics = C
End Function



Public Function TransformaLineaCoopic(CADENA As String, ByRef ElAnyo As Integer, ByRef Seg As Integer) As String
Dim C As String

    'Se trata de a partir de la cadena de ROBOTICS
    'GENERO LA CADENA DE TCP que es la que trabajaremos
    
    'TCP3:      01234,11,23,08,20,0000,0000,18411
    
    '           1  5    0    5    0    5
    'COOPIC:  000471610271127520000011ILOC010
    '         ttttt  trabajador
    '                                Terminal
    '              yymmdd
    '                    hhmmss
        
        
    C = Mid(CADENA, 1, 5) & ","
    C = C & Mid(CADENA, 8, 2) & "," & Mid(CADENA, 10, 2) & ","
    ElAnyo = CInt("20" & Mid(CADENA, 6, 2))
    C = C & Mid(CADENA, 12, 2) & "," & Mid(CADENA, 14, 2)  'HORA
    Seg = CInt(Mid(CADENA, 16, 2))
    C = C & ",0000,"
    'NO LLEVA inci
    C = C & "0000"
    C = C & ",12345"
    TransformaLineaCoopic = C
End Function


