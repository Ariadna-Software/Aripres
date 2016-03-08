Attribute VB_Name = "GestorLib"
'***************************************************************
'***************************************************************
'
'
'   Traido de GESALB
'
'
'***************************************************************
'***************************************************************
'***************************************************************






Public Sub AbrirBaseDatos()
    If GesHuellaDB Is Nothing Then
        '-- Abrimos la base datos
        Set GesHuellaDB = New BaseDatos2
        'ANTES 2013. BD en access.   Ahora la llevamos tb a MYSQL
        'GesHuellaDB.abrir "accGestorHuella", "", ""
        'GesHuellaDB.Tipo = "ACCESS"
        
        GesHuellaDB.abrir "accGestorHuella2", "", ""
        GesHuellaDB.Tipo = "MYSQL"
        
        
    End If
End Sub

Public Function hex2(B As Byte) As String
    Dim s As String
    s = Hex(B)
    Do While Len(s) < 2
      s = "0" + s
    Loop
    hex2 = s
End Function

Public Function CalcCRC(sTrama As String) As String
  Dim i As Integer
  Dim iCRC As Integer
  
  iCRC = 0
  For i = 1 To Len(sTrama)
    iCRC = (iCRC + Asc(Mid(sTrama, i, 1))) Mod 256
  Next i
  CalcCRC = hex2(iCRC Mod 256)
End Function


'CATADAU TENDRA EL SUYO PROPIO

Public Function GrabaFichajeGesLabALZIRA(registro As String)
    Dim SQL As String
    '-- Graba en la tabla EntradaFichajes
    Dim Secuencia As Long
    Dim idTrabajador As Long
    Dim Fecha As Date
    Dim Hora As Date
    Dim idInci As Integer
    Dim HoraReal As Date
    Dim FecHoraLeida As String
    Dim CualEsIncidencia As String
    Dim InsertamosEnBdAccess As Boolean
    
    '---
    Dim usu As UsuarioHuella
    Set usu = New UsuarioHuella
    
    
    
    'Incidencia
    CualEsIncidencia = Right(registro, 2)  'Los dos ultimos son las incidencias
    
    
    'YA NO ESTA CATADU AQUI DENTRO
    'A aripres solo entran las 00,01,02
    'En alzira SOLO entran estas dos
    'If MiEmpresa.QueEmpresa = 4 Then
    '    InsertamosEnBdAccess = True
    'Else
        'ALZIRA
        InsertamosEnBdAccess = CualEsIncidencia = "00" Or CualEsIncidencia = "01" Or CualEsIncidencia = "02"
    'End If
    
    
    If InsertamosEnBdAccess Then
        '-- Si el usuario no está dado de alta despeciamos la información
        If usu.Leer(Mid(registro, 1, 10)) Then
            
        
        
            idTrabajador = usu.GesLabID
            FecHoraLeida = Mid(registro, 11, 12)
            Fecha = CDate(Mid(FecHoraLeida, 5, 2) & "/" & _
                            Mid(FecHoraLeida, 3, 2) & "/" & _
                            "20" & Mid(FecHoraLeida, 1, 2))
            HoraReal = CDate(Mid(FecHoraLeida, 7, 2) & ":" & _
                            Mid(FecHoraLeida, 9, 2) & ":" & _
                            Mid(FecHoraLeida, 11, 2))
            Hora = HoraReal
            idInci = 0

            

     '       'Catadau.
     '       'Es obligado marcar la salida
     '       'Ya que las tareas tb entran al proceso,y luego las tengo que quitar
     '       If MiEmpresa.QueEmpresa = 4 Then
     '           'Todas las tareas son e menos las salidas que grabara la incidencia 2
     '           If CualEsIncidencia = "02" Then idInci = 2   'SALIDA
     '       End If
            
            
            Secuencia = ObtenerSecuencia()
            SQL = "insert into EntradaFichajes(Secuencia, idTrabajador, Fecha, Hora, idInci, HoraReal) " & _
                        " values("
            SQL = SQL & Secuencia & ","
            SQL = SQL & idTrabajador & ","
            SQL = SQL & DBSet(Fecha, "F") & ","
            SQL = SQL & DBSet(Hora, "H") & ","
            SQL = SQL & idInci & ","
            SQL = SQL & DBSet(HoraReal, "H") & ")"
            'db.ejecutar SQL
            conn.Execute SQL
            
        End If
    Else
        'NO es 00,01,02
        'Stop
        
    End If
End Function


'DOS PROCESOS , meter sea lo que sea en marcajeskimaldi
' y luego meterlo tb en fichajeactual
Public Function GrabaFichajeGesLabCATADAU(registro As String, Nodo As Byte)
    Dim SQL As String
    '-- Graba en la tabla EntradaFichajes
    Dim Secuencia As Long
    Dim idTrabajador As Long
    Dim Fecha As Date
    Dim Hora As Date
    Dim idInci As Integer
    Dim HoraReal As Date
    Dim FecHoraLeida As String

    
    
    '---
    Dim usu As UsuarioHuella
    Set usu = New UsuarioHuella
    
    
    
  
  
        'Trozo comun
        FecHoraLeida = Mid(registro, 11, 12)
        Fecha = CDate(Mid(FecHoraLeida, 5, 2) & "/" & _
                        Mid(FecHoraLeida, 3, 2) & "/" & _
                        "20" & Mid(FecHoraLeida, 1, 2))
        HoraReal = CDate(Mid(FecHoraLeida, 7, 2) & ":" & _
                        Mid(FecHoraLeida, 9, 2) & ":" & _
                        Mid(FecHoraLeida, 11, 2))
  
  
        'Primero,sea lo que sea, insertamos en marcajkes kimaldi
        'MarcajesKimaldi
        SQL = "INSERT INTO MarcajesKimaldi (Nodo,Fecha,Hora,TipoMens,Marcaje) VALUES "
        SQL = SQL & "(" & Nodo & "," & db.Fecha(Fecha) & "," & db.Hora(HoraReal) & "," & db.Texto(Right(registro, 2))
        'COJO SOLO las ultimas 4 posciones
        SQL = SQL & "," & db.Texto(Mid(registro, 7, 4)) & ")"
        db.ejecutar SQL
        
        '-- Si el usuario no está dado de alta despeciamos la información
        If usu.Leer(Mid(registro, 1, 10)) Then
            
            
        
            idTrabajador = usu.GesLabID
            
            Hora = HoraReal
            idInci = 0

            

            'Catadau.
            'Es obligado marcar la salida
            'Ya que las tareas tb entran al proceso,y luego las tengo que quitar
            If MiEmpresa.QueEmpresa = 4 Then
                'Todas las tareas son e menos las salidas que grabara la incidencia 2
                If Right(registro, 2) = "02" Then idInci = 2   'SALIDA
            End If
            
            
            Secuencia = ObtenerSecuencia()  'ObtenerSecuencia(db)
            SQL = "insert into EntradaFichajes(Secuencia, idTrabajador, Fecha, Hora, idInci, HoraReal) " & _
                        " values("
            SQL = SQL & db.Numero(Secuencia) & ","
            SQL = SQL & db.Numero(idTrabajador) & ","
            SQL = SQL & db.Fecha(Fecha) & ","
            SQL = SQL & db.Hora(Hora) & ","
            SQL = SQL & db.Numero(idInci) & ","
            SQL = SQL & db.Hora(HoraReal) & ")"
            db.ejecutar SQL
        End If

End Function


'Public Function ObtenerSecuencia(db As BaseDatos) As Long
Public Function ObtenerSecuencia() As Long
    Dim SQL As String
    Dim Rs As ADODB.Recordset
    SQL = "select Max(Secuencia) from EntradaFichajes"
    'Set Rs = db.cursor(SQL)
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not IsNull(Rs.Fields(0)) Then
        ObtenerSecuencia = Rs.Fields(0) + 1
    Else
        ObtenerSecuencia = 1
    End If
    Rs.Close
    Set Rs = Nothing
End Function







'Esto estaba en GESALB, en otro modulo
Public Sub CargaComboSecciones(ByRef CBO As ComboBox, AñadirTodas As Boolean)
Dim SQL As String
Dim Rs As ADODB.Recordset

    CBO.Clear
    SQL = "select IdSeccion,nombre from secciones order by NOMBRE"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If AñadirTodas Then
        CBO.AddItem "Todas las secciones"
        CBO.ItemData(CBO.NewIndex) = -1
    End If
    
    While Not Rs.EOF
        CBO.AddItem Rs!Nombre & " (" & Rs!idSeccion & ")"
        CBO.ItemData(CBO.NewIndex) = Rs!idSeccion
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    If AñadirTodas Then CBO.ListIndex = 0
    
End Sub
