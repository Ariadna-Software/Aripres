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
        GesHuellaDB.tipo = "MYSQL"
        
        
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

Public Function GrabaFichajeGesLabALZIRA(registro As String, RelojAuxiliar As Boolean)
    Dim SQL As String
    '-- Graba en la tabla EntradaFichajes
    Dim Secuencia As Long
    Dim idTrabajador As Long
    Dim Fecha As Date
    Dim Hora As Date
    Dim IdInci As Integer
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
        '-- Si el usuario no est� dado de alta despeciamos la informaci�n
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
            IdInci = 0

            

     '       'Catadau.
     '       'Es obligado marcar la salida
     '       'Ya que las tareas tb entran al proceso,y luego las tengo que quitar
     '       If MiEmpresa.QueEmpresa = 4 Then
     '           'Todas las tareas son e menos las salidas que grabara la incidencia 2
     '           If CualEsIncidencia = "02" Then idInci = 2   'SALIDA
     '       End If
            
            
            Secuencia = ObtenerSecuencia(RelojAuxiliar)
            SQL = "insert into " & IIf(RelojAuxiliar, "entradafichajAuxliares", "EntradaFichajes")
            SQL = SQL & "(Secuencia, idTrabajador, Fecha, Hora, idInci, HoraReal)  values("
            SQL = SQL & Secuencia & ","
            SQL = SQL & idTrabajador & ","
            SQL = SQL & DBSet(Fecha, "F") & ","
            SQL = SQL & DBSet(Hora, "H") & ","
            SQL = SQL & IdInci & ","
            SQL = SQL & DBSet(HoraReal, "H") & ")"
            'db.ejecutar SQL
            conn.Execute SQL
            
        End If
    Else
        'NO es 00,01,02
        '
        
    End If
End Function


'DOS PROCESOS , meter sea lo que sea en marcajeskimaldi
' y luego meterlo tb en fichajeactual
Public Function GrabaFichajeGesLabCATADAU(registro As String, Nodo_ As Byte, RelojAuxiliar As Boolean)
    Dim SQL As String
    '-- Graba en la tabla EntradaFichajes
    Dim Secuencia As Long
    Dim idTrabajador As Long
    Dim Fecha As Date
    Dim Hora As Date
    Dim IdInci As Integer
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
        SQL = SQL & "(" & Nodo_ & "," & DBSet(Fecha, "F") & "," & DBSet(HoraReal, "H") & "," & DBSet(Right(registro, 2), "T")
        'COJO SOLO las ultimas 4 posciones
        SQL = SQL & "," & DBSet(Mid(registro, 7, 4), "T") & ")"
        conn.Execute SQL
        
        '-- Si el usuario no est� dado de alta despeciamos la informaci�n
        If usu.Leer(Mid(registro, 1, 10)) Then
            
            
        
            idTrabajador = usu.GesLabID
            
            Hora = HoraReal
            IdInci = 0

            

            'Catadau.
            'Es obligado marcar la salida
            'Ya que las tareas tb entran al proceso,y luego las tengo que quitar
            If vEmpresa.QueEmpresa = 4 Then
                'Todas las tareas son e menos las salidas que grabara la incidencia 2
                If Right(registro, 2) = "02" Then IdInci = 2   'SALIDA
            End If
            
            
            Secuencia = ObtenerSecuencia(RelojAuxiliar)  'ObtenerSecuencia(db)
            SQL = "insert into EntradaFichajes(Secuencia, idTrabajador, Fecha, Hora, idInci, HoraReal) " & _
                        " values("
            SQL = SQL & Secuencia & ","
            SQL = SQL & idTrabajador & ","
            SQL = SQL & DBSet(Fecha, "F") & ","
            SQL = SQL & DBSet(Hora, "H") & ","
            SQL = SQL & IdInci & ","
            SQL = SQL & DBSet(HoraReal, "H") & ")"
            Debug.Print SQL
            conn.Execute SQL
        End If

End Function


'Public Function ObtenerSecuencia(db As BaseDatos) As Long
Public Function ObtenerSecuencia(EsTablaRelojAuxiliar As Boolean) As Long
    Dim SQL As String
    Dim Rs As ADODB.Recordset
    SQL = "select Max(Secuencia) from " & IIf(EsTablaRelojAuxiliar, "entradafichajAuxliares", "EntradaFichajes")
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







