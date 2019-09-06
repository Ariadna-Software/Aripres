Attribute VB_Name = "ExpresionBusqueda"

Public Function SeparaCampoBusqueda(tipo As String, Campo As String, CADENA As String, ByRef DevSQL As String) As Byte
Dim cad As String
Dim Aux As String
Dim Ch As String
Dim Fin As Boolean
Dim i, J As String

On Error GoTo ErrSepara
SeparaCampoBusqueda = 1
DevSQL = ""
cad = ""
Select Case tipo
Case "N"
    '----------------  NUMERICO  ---------------------
    i = CararacteresCorrectos(CADENA, "N")
    If i > 0 Then Exit Function  'Ha habido un error y salimos
    'Comprobamos si hay intervalo ':'
    i = InStr(1, CADENA, ":")
    If i > 0 Then
        'Intervalo numerico
        cad = Mid(CADENA, 1, i - 1)
        Aux = Mid(CADENA, i + 1)
        If Not IsNumeric(cad) Or Not IsNumeric(Aux) Then Exit Function  'No son numeros
        'Intervalo correcto
        'Construimos la cadena
        
        '+-+-+- 23/05/2005 Canvi de Cèsar: per a que DevSQL tinga este aspecte: (taula.camp) >= 5 AND (taula.camp) <= 7 -+-+-+-+-
        'DevSQL = campo & " >= " & cad & " AND " & campo & " <= " & Aux
        DevSQL = Campo & " >= " & cad & ") AND (" & Campo & " <= " & Aux
        '+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
        
        'ELSE
        Else
            'Prueba
            'Comprobamos que no es el mayor
            If CADENA = ">>" Or CADENA = "<<" Then
                DevSQL = "1=1"
             Else
                    Fin = False
                    i = 1
                    cad = ""
                    Aux = "NO ES NUMERO"
                    While Not Fin
                        Ch = Mid(CADENA, i, 1)
                        If Ch = ">" Or Ch = "<" Or Ch = "=" Then
                            cad = cad & Ch
                            Else
                                Aux = Mid(CADENA, i)
                                Fin = True
                        End If
                        i = i + 1
                        If i > Len(CADENA) Then Fin = True
                    Wend
                    'En aux debemos tener el numero
                    If Not IsNumeric(Aux) Then Exit Function
                    'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                    If cad = "" Then cad = " = "
                    DevSQL = Campo & " " & cad & " " & Aux
            End If
        End If
Case "F"
     '---------------- FECHAS ------------------
     
    i = CararacteresCorrectos(CADENA, "F")
    If i = 1 Then Exit Function
    'Comprobamos si hay intervalo ':'
    i = InStr(1, CADENA, ":")
    If i > 0 Then
        'Intervalo de fechas
        cad = Mid(CADENA, 1, i - 1)
        Aux = Mid(CADENA, i + 1)
        If Not EsFechaOKString(cad) Or Not EsFechaOKString(Aux) Then Exit Function  'Fechas incorrectas
        'Intervalo correcto
        'Construimos la cadena
        cad = Format(cad, FormatoFecha)
        Aux = Format(Aux, FormatoFecha)
        'En my sql es la ' no el #
        'DevSQL = Campo & " >=#" & Cad & "# AND " & Campo & " <= #" & AUX & "#"
        DevSQL = Campo & " >='" & cad & "' AND " & Campo & " <= '" & Aux & "'"
        '----
        'ELSE
        Else
            'Comprobamos que no es el mayor
            If CADENA = ">>" Or CADENA = "<<" Then
                  DevSQL = "1=1"
            Else
                Fin = False
                i = 1
                cad = ""
                Aux = "NO ES FECHA"
                While Not Fin
                    Ch = Mid(CADENA, i, 1)
                    If Ch = ">" Or Ch = "<" Or Ch = "=" Then
                        cad = cad & Ch
                        Else
                            Aux = Mid(CADENA, i)
                            Fin = True
                    End If
                    i = i + 1
                    If i > Len(CADENA) Then Fin = True
                Wend
                'En aux debemos tener el numero
                If Not EsFechaOKString(Aux) Then Exit Function
                'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                If Not Left(Campo, 1) = "{" Then
                    Aux = "'" & Format(Aux, FormatoFecha) & "'"
                Else
                    Aux = "Date(" & Year(Aux) & "," & Month(Aux) & "," & Day(Aux) & ")"
                End If
                If cad = "" Then cad = " = "
                DevSQL = Campo & " " & cad & " " & Aux
            End If
        End If
    
    
Case "T"
    '---------------- TEXTO ------------------
    i = CararacteresCorrectos(CADENA, "T")
    If i = 1 Then Exit Function
    
    'Comprobamos que no es el mayor
    If CADENA = ">>" Or CADENA = "<<" Then
        DevSQL = "1=1"
        Exit Function
    End If
    
    'Comprobamos si hay intervalo ':'
    i = InStr(1, CADENA, ":")
    If i > 0 Then
        'Intervalo numerico
        cad = Mid(CADENA, 1, i - 1)
        Aux = Mid(CADENA, i + 1)
'        If Not IsNumeric(cad) Or Not IsNumeric(Aux) Then Exit Function  'No son numeros
        'Intervalo correcto
        'Construimos la cadena
        DevSQL = Campo & " >= '" & cad & "' AND " & Campo & " <= '" & Aux & "'"
    Else
    
        'Comprobamos si es LIKE o NOT LIKE
        cad = Mid(CADENA, 1, 2)
        If cad = "<>" Then
            CADENA = Mid(CADENA, 3)
            If Left(Campo, 1) <> "{" Then
                'No es consulta seleccion para Report.
                DevSQL = Campo & " NOT LIKE '"
            Else
                'Consulta de seleccion para Crystal Report
                DevSQL = "NOT (" & Campo & " LIKE """ & CADENA & """)"
            End If
        Else
            If Left(Campo, 1) <> "{" Then
            'NO es para report
                DevSQL = Campo & " LIKE '"
            Else  'Es para report
                i = InStr(1, CADENA, "*")
                'Poner Consulta de seleccion para Crystal Report
                If i > 0 Then
                    DevSQL = Campo & " LIKE """ & CADENA & """"
                Else
                    DevSQL = Campo & " = """ & CADENA & """"
                End If
            End If
        End If
        
    
        'Cambiamos el * por % puesto que en ADO es el caraacter para like
        i = 1
        Aux = CADENA
        If Not Left(Campo, 1) = "{" Then
          'No es para report
           While i <> 0
               i = InStr(1, Aux, "*")
               If i > 0 Then
                    Aux = Mid(Aux, 1, i - 1) & "%" & Mid(Aux, i + 1)
                End If
            Wend
        End If
        
        'Cambiamos el ? por la _ pue es su omonimo
        i = 1
        While i <> 0
            i = InStr(1, Aux, "?")
            If i > 0 Then Aux = Mid(Aux, 1, i - 1) & "_" & Mid(Aux, i + 1)
        Wend
    
        'Poner el valor de la expresion
        If Left(Campo, 1) <> "{" Then
            'No es consulta seleccion para Report.
            DevSQL = DevSQL & Aux & "'"
        'Else
            'Consulta de seleccion para Crystal Report
            'DevSQL = DevSQL & CADENA & """)"
        End If
    End If
    '=========
    'ANTES
'    If cad = "<>" Then
'        '====David
'        'Aux = Mid(CADENA, 3)
'        'LAura
'        Aux = Mid(Aux, 3)
'        '====
'        If Left(Campo, 1) <> "{" Then
'            'Mo es consulta seleccion para Report.
'            DevSQL = Campo & " NOT LIKE '" & Aux & "'"
'        Else
'            'Consulta de seleccion para Crystal Report
'            DevSQL = Campo & " <> " & Aux & ""
'        End If
'    Else
'        If Left(Campo, 1) <> "{" Then
'            DevSQL = Campo & " LIKE '" & Aux & "'"
'        ElseIf Left(Aux, 4) = "like" Then
'            'Consulta de seleccion para Crystal Report
'            DevSQL = Campo & " " & Aux
'        Else
'            'Consulta de seleccion para Crystal Report
'            DevSQL = Campo & " = """ & Aux & """"
'        End If
'    End If
    
    
Case "B"
    'Como vienen de check box o del option box
    'los escribimos nosotros luego siempre sera correcta la
    'sintaxis
    'Los booleanos. Valores buenos son
    'Verdadero , Falso, True, False, = , <>
    'Igual o distinto
    i = InStr(1, CADENA, "<>")
    If i = 0 Then
        'IGUAL A valor
        cad = " = "
        Else
            'Distinto a valor
        cad = " <> "
    End If
    'Verdadero o falso
    i = InStr(1, CADENA, "V")
    If i > 0 Then
            Aux = "True"
            Else
            Aux = "False"
    End If
    'Ponemos la cadena
    DevSQL = Campo & " " & cad & " " & Aux
    
Case Else
    'No hacemos nada
        Exit Function
End Select
SeparaCampoBusqueda = 0
ErrSepara:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function


Private Function CararacteresCorrectos(vcad As String, tipo As String) As Byte
Dim i As Integer
Dim Ch As String
Dim Error As Boolean

CararacteresCorrectos = 1
Error = False
Select Case tipo
Case "N"
    'Numero. Aceptamos numeros, >,< = :
    For i = 1 To Len(vcad)
        Ch = Mid(vcad, i, 1)
        Select Case Ch
            Case "0" To "9"
            Case "<", ">", ":", "=", ".", " ", "-"
            Case Else
                Error = True
                Exit For
        End Select
    Next i
Case "T"
    'Texto aceptamos numeros, letras y el interrogante y el asterisco
    For i = 1 To Len(vcad)
        Ch = Mid(vcad, i, 1)
        Select Case Ch
            Case "a" To "z"
            Case "è", "é", "í" 'Añade Laura: 16/03/06
            Case "A" To "Z"
            Case "0" To "9"
            Case "*", "%", "?", "_", "\", "/", ":", ".", " " ' estos son para un caracter sol no esta demostrado , "%", "&"
            'Esta es opcional
            Case "<", ">"
            Case "Ñ", "ñ"
            Case "-" 'Añade Laura
            Case Else
                Error = True
                Exit For
        End Select
    Next i
    
Case "F"
    'Tipo Fecha. Aceptamos Numeros , "/" ,":"
    For i = 1 To Len(vcad)
        Ch = Mid(vcad, i, 1)
        Select Case Ch
            Case "0" To "9"
            Case "<", ">", ":", "/", "="
            Case Else
                Error = True
                Exit For
        End Select
    Next i

Case "B"
    'Numeros , "/" ,":"
    For i = 1 To Len(vcad)
        Ch = Mid(vcad, i, 1)
        Select Case Ch
            Case "0" To "9"
            Case "<", ">", ":", "/", "=", " "
            Case Else
                Error = True
                Exit For
        End Select
    Next i
End Select
'Si no ha habido error cambiamos el retorno
If Not Error Then CararacteresCorrectos = 0
End Function


Public Function QuitarCaracterEnter(vcad As String) As String
'Dim i As Integer
'
'    Do
'        i = InStr(1, vcad, Chr(13))
'        If i > 0 Then 'Hay ENTER
'            vcad = Mid(vcad, 1, i - 1) & Mid(vcad, i + 2)
'        End If
'    Loop Until i = 0
'    QuitarCaracterEnter = vcad
End Function



Public Function ContieneCaracterBusqueda(CADENA As String) As Boolean
'Comprueba si la cadena contiene algun caracter especial de busqueda
' >,>,>=,: , ....
'si encuentra algun caracter de busqueda devuelve TRUE y sale
Dim B As Boolean
Dim i As Integer

    
    
    'NULL
    If UCase(CADENA) = "NULL" Then
        ContieneCaracterBusqueda = True
        Exit Function
    End If
    
    
    
    i = 1
    B = False
    Do
        Ch = Mid(CADENA, i, 1)
        Select Case Ch
            Case "<", ">", ":", "="
                B = True
            Case "*", "%", "?", "_", "\", ":" ', "."
                B = True
            Case Else
                B = False
        End Select
    'Next i
        i = i + 1
    Loop Until (B = True) Or (i > Len(CADENA))
    ContieneCaracterBusqueda = B
End Function

