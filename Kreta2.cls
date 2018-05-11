VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Kreta2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvarNumero As Integer 'copia local
Private mvarSocket As Winsock
Private mvarRecibido As String


Private mvarDeshabilitado As Boolean
Private mvarRelojAuxiliar As Boolean
    'Si es reloj auxiliar, No contabiliza para horas  (ALZIRA Terminales baño)


'Para activar este evento, use RaiseEvent con la siguiente sintaxis:
'RaiseEvent LanzaMensaje[(arg1, arg2, ... , argn)]
Public Event LanzaMensaje(Mensaje As String)


Public Property Let Numero(ByVal vData As Integer)
    mvarNumero = vData
End Property

Public Property Get Numero() As Integer
    Numero = mvarNumero
End Property

Public Property Set Socket(ByVal vData As Winsock)
    Set mvarSocket = vData
End Property

Public Property Get Socket() As Winsock
    Set Socket = mvarSocket
End Property

Public Property Let Recibido(ByVal vData As String)
    Let mvarRecibido = vData
End Property


Public Property Get Deshabilitado() As Boolean
    Deshabilitado = mvarDeshabilitado
End Property

Public Property Let Deshabilitado(ByVal vData As Boolean)
    mvarDeshabilitado = vData
End Property


Public Property Get RelojAuxiliar() As Boolean
    RelojAuxiliar = mvarRelojAuxiliar
End Property


Public Property Let RelojAuxiliar(ByVal vData As Boolean)
    mvarRelojAuxiliar = vData
End Property




Public Function ComprobarConexion() As Boolean
    Dim TiempoEspera As Long
    Dim Conectado As Boolean
    TiempoEspera = 0
    ComprobarConexion = False
    If Socket Is Nothing Then Exit Function
    If Socket.State = sckConnected Then
        ComprobarConexion = True
        Exit Function
    End If
    '-- Comprobamos si está conectado y si no lo está lo conectamos
    Socket.Close
    Socket.Connect
    '-- Realizamos un bucle de espera de conexión
    While Not (Conectado)
        TiempoEspera = TiempoEspera + 1
        If TiempoEspera > 10000 Then Conectado = True ' mentira pero salimos del bucle
        If Socket.State = sckConnected Then Conectado = True ' verdad
        DoEvents
    Wend
    If TiempoEspera > 10000 Then Exit Function
    '-- Ahora ya tenemos conexión
    ComprobarConexion = True
End Function

Public Function CargarConfiguracion() As Boolean
    Dim Sql As String
    Dim RS As ADODB.Recordset
    Dim Envio As String
    Dim Espero As String
    '-- CargarConfiguración: Carga la configuración según esté en la base de datos
    '   en un terminal determinado
    If Not ComprobarConexion() Then Exit Function
    CargarConfiguracion = True
    Sql = "select * from parametros"
    Set RS = GesHuellaDB.cursor(Sql)
    If Not RS.EOF Then
        RS.MoveFirst
        While Not RS.EOF
            '-- Montamos la cadena para grabar la configuración
            Envio = "311" & Right("00" & Hex(RS!codparam), 2) & Right("00" & Hex(RS!valor), 2)
            Espero = "4111"
            RaiseEvent LanzaMensaje("Terminal: " & mvarNumero & " Mandando Parámetro: " & RS!codparam)
            If EnviarRecibir(Envio, Espero) <> 0 Then CargarConfiguracion = False
            RS.MoveNext
        Wend
    End If
End Function

Public Function CargarMensajes() As Boolean
    Dim Sql As String
    Dim RS As ADODB.Recordset
    Dim Envio As String
    Dim Espero As String
    '-- CargarConfiguración: Carga la configuración según esté en la base de datos
    '   en un terminal determinado
    If Not ComprobarConexion() Then Exit Function
    CargarMensajes = True
    Sql = "select * from mensajes"
    Set RS = GesHuellaDB.cursor(Sql)
    If Not RS.EOF Then
        RS.MoveFirst
        While Not RS.EOF
            '-- Montamos la cadena para grabar la configuración
            Envio = "321" & Right("00" & Hex(RS!codmens), 2) & Left(RS!valor & String(20, " "), 20)
            Espero = "4211"
            RaiseEvent LanzaMensaje("Terminal: " & mvarNumero & " Mandando mensaje: " & RS!codmens)
            If EnviarRecibir(Envio, Espero) <> 0 Then CargarMensajes = False
            RS.MoveNext
        Wend
    End If
End Function

Public Function CargarDias() As Boolean
    Dim Sql As String
    Dim RS As ADODB.Recordset
    Dim Envio As String
    Dim Espero As String
    '-- CargarConfiguración: Carga la configuración según esté en la base de datos
    '   en un terminal determinado
    If Not ComprobarConexion() Then Exit Function
    CargarDias = True
    Sql = "select * from dias"
    Set RS = GesHuellaDB.cursor(Sql)
    If Not RS.EOF Then
        RS.MoveFirst
        While Not RS.EOF
            '-- Montamos la cadena para grabar la configuración
            Envio = "331" & Right("00" & Hex(RS!coddia), 2) & Left(RS!valor & String(3, " "), 3)
            Espero = "4311"
            RaiseEvent LanzaMensaje("Terminal: " & mvarNumero & " cargando dia: " & RS!coddia)
            If EnviarRecibir(Envio, Espero) <> 0 Then CargarDias = False
            RS.MoveNext
        Wend
    End If
End Function

Public Function CargarMeses() As Boolean
    Dim Sql As String
    Dim RS As ADODB.Recordset
    Dim Envio As String
    Dim Espero As String
    '-- CargarConfiguración: Carga la configuración según esté en la base de datos
    '   en un terminal determinado
    If Not ComprobarConexion() Then Exit Function
    CargarMeses = True
    Sql = "select * from meses"
    Set RS = GesHuellaDB.cursor(Sql)
    If Not RS.EOF Then
        RS.MoveFirst
        While Not RS.EOF
            '-- Montamos la cadena para grabar la configuración
            Envio = "341" & Right("00" & Hex(RS!codmes), 2) & Left(RS!valor & String(3, " "), 3)
            Espero = "4411"
            RaiseEvent LanzaMensaje("Terminal: " & mvarNumero & " cargando mes: " & RS!codmes)
            If EnviarRecibir(Envio, Espero) <> 0 Then CargarMeses = False
            RS.MoveNext
        Wend
    End If
End Function
Public Function CargarIncidencias() As Boolean
    Dim Sql As String
    Dim RS As ADODB.Recordset
    Dim Envio As String
    Dim Espero As String
    '-- CargarConfiguración: Carga la configuración según esté en la base de datos
    '   en un terminal determinado
    If Not ComprobarConexion() Then Exit Function
    CargarIncidencias = True
    Sql = "select * from incidencias"
    Set RS = GesHuellaDB.cursor(Sql)
    If Not RS.EOF Then
        RS.MoveFirst
        While Not RS.EOF
            '-- Montamos la cadena para grabar la configuración
            Envio = "544" & Right("0000" & Hex(RS!codinc), 4) & Left(RS!valor & String(20, " "), 20)
            Espero = "64411"
            RaiseEvent LanzaMensaje("Terminal: " & mvarNumero & " Mandando incidencia: " & RS!codinc)
            If EnviarRecibir(Envio, Espero) <> 0 Then CargarIncidencias = False
            RS.MoveNext
        Wend
    End If
End Function

Public Function CargarHSPorDefecto() As Boolean
    Dim Envio As String
    Dim Espero As String
    If Not ComprobarConexion() Then Exit Function
    '-- Mandamos un horario por defecto con la idea de poder cargar usuarios
    Envio = "5240001060014001400220022000600"
    Espero = "62411"
    EnviarRecibir Envio, Espero
    RaiseEvent LanzaMensaje("Terminal: " & mvarNumero & " Cargar horario")
    '-- Mandamos semana
    Envio = "534000101010101010101"
    Espero = "63411"
    RaiseEvent LanzaMensaje("Terminal: " & mvarNumero & " Cargar semana")
    EnviarRecibir Envio, Espero
    CargarHSPorDefecto = True
End Function

Public Function BorrarTodosLosUsuarios() As Boolean
    Dim Envio As String
    Dim Espero As String
    BorrarTodosLosUsuarios = False
    If Not ComprobarConexion() Then Exit Function
    '-- Mandamos un horario por defecto con la idea de poder cargar usuarios
    Envio = "561"
    Espero = "66111"
    If EnviarRecibir(Envio, Espero) Then BorrarTodosLosUsuarios = True
    RaiseEvent LanzaMensaje("Terminal: " & mvarNumero & " borrar todos los usuarios")
End Function

Public Function BorrarTodasLasIncidencias() As Boolean
    Dim Envio As String
    Dim Espero As String
    BorrarTodasLasIncidencias = False
    If Not ComprobarConexion() Then Exit Function
    '-- Mandamos un horario por defecto con la idea de poder cargar usuarios
    Envio = "541"
    Espero = "64111"
    If EnviarRecibir(Envio, Espero) Then BorrarTodasLasIncidencias = True
    RaiseEvent LanzaMensaje("Terminal: " & mvarNumero & " borrar todas las incidencias")
End Function

Public Function BorrarUsuario(usu As UsuarioHuella) As Boolean
    Dim Envio As String
    Dim Espero As String
    BorrarUsuario = False
    If Not ComprobarConexion() Then Exit Function
    '-- Mandamos un horario por defecto con la idea de poder cargar usuarios
    Envio = "562" & usu.CodUsuario
    Espero = "66211"
    If EnviarRecibir(Envio, Espero) Then BorrarUsuario = True
    RaiseEvent LanzaMensaje("Terminal: " & mvarNumero & " borrar el los usuario:" & usu.CodUsuario)
End Function

'La primera vez, para la empresa 3 (BLEGIDA) , si hay en entradafichajes los pasa tb  || lbl puede ser nothing
Public Function LeerMarcajes(Directorio As String, PrimeraVez As Boolean, ByRef lbl As Label) As Boolean
    '-- Lee los marcajes del terminal y los guarda en un fichero en el directorio indicado
    Dim NF As Integer
    Dim Fichero As String
    Dim Final As Boolean
    Dim Envio As String
    Dim Espero As String
    Dim Espero2 As String
    Dim Aux As String
    Dim TieneDatos As Boolean
    Dim RS As ADODB.Recordset
    Dim FicheroCatadu As String
    
    
    
    
    If Not ComprobarConexion() Then Exit Function
    NF = FreeFile


        
    'Si es reloj auxuliar , el fichero se llamará \AUX, si no conserva su nombre
    
    If mvarRelojAuxiliar Then
        Fichero = Directorio & "\AUX" & Format(Now, "yyyyMMddhhmmss") & _
                "T" & Format(mvarNumero, "00") & ".txt"
    Else
        Fichero = Directorio & "\HU" & Format(Now, "yyyyMMddhhmmss") & _
                "T" & Format(mvarNumero, "00") & ".txt"
    End If
    espera 0.3
    Final = False
    TieneDatos = False
    

    
    Open Fichero For Output As NF

    
    While Not Final
    
        'El lbl
        If Not lbl Is Nothing Then
            If Val(Timer) - Val(lbl.Tag) > 2 Then
                lbl.Visible = Not lbl.Visible
                lbl.Refresh
                lbl.Tag = Val(Timer)
            End If
            
        End If
    
    
    
    
        '-- Leemos el marcaje
        Envio = "587"
        Espero = "687"
        Espero2 = "68700000000000000000000000000"
        If EnviarRecibir(Envio, Espero, Espero2) = 2 Then
            '-- Se acabó
            Final = True
        Else
            If mvarRecibido <> "" Then TieneDatos = True
            '-- Grabamos lo recibido en el fichero
            mvarRecibido = Mid(mvarRecibido, 5, 26)
            
            ''esto era para belgida
            ''pOR SI ACASO VUELVEN
            'If vEmpresa.QueEmpresa = 2 Then
            '    'Transformamos el fichero en formato TCP3  00000000081106291327163302(esto es la linea del kreta)
            '    '00001,11,27,12,59,0001,0002,32843
            '
            '    ' tra   m d   h  m  0    0    0
            '
            '    Aux = Mid(mvarRecibido, 1, 10)
            '    If Not IsNumeric(Aux) Then
            '        'ERROR no es fichada valida
            '        mvarRecibido = ""
            '    Else
            '        Aux = Mid(mvarRecibido, 6, 5)  'trab
            '        Aux = Aux & "," & Mid(mvarRecibido, 13, 2)  'mes
            '        Aux = Aux & "," & Mid(mvarRecibido, 15, 2)  'dia
            '        Aux = Aux & "," & Mid(mvarRecibido, 17, 2)  'hora
            '        Aux = Aux & "," & Mid(mvarRecibido, 19, 2)  'min
            '        Aux = Aux & ",0000,0000,00000"
            '        mvarRecibido = Aux
            '    End If
            'Else
            '
            '    'ALZIRA. Marzo 2012
            '    'Guardaremos todo en el fichero. Luego, procesando, discriminaremos las que son de incidencias de las que no
            '
            
            '                'ALZIRA. Veremos si la fichada es correcta Y si tiene incidencia o no
            '                If Len(mvarRecibido) = 26 Then
            '                    'Tienen longitud correcta
            '                    Aux = Right(mvarRecibido, 2)
            '                    If Val(Aux) > 2 Then
            '                        'OKkkkkk, tiene incidencia. NO la guardare en el fichero
            '                        TieneIncidencias = True
            '                        Print #FicheroIncidencias, mvarRecibido
            '                        mvarRecibido = "" 'para que no lo guarde en el de fichajes
            '                    End If
            '                End If
            '            End If
            If mvarRecibido <> "" Then Print #NF, mvarRecibido
            '-- Borramos el marcaje
            Envio = "582"
            Espero = "682"
            EnviarRecibir Envio, Espero
        End If
    Wend
    
    
    'SOLO ERA PARA BELGIDA. Si entra en el 4 habra que ajustar esto
    If PrimeraVez And vEmpresa.QueEmpresa = 2222 Then
        '1ª y BELGIç
        lbl.Caption = "Leyendo aritick"
        lbl.Refresh
        Aux = "Select * from entradafichajes order by Secuencia desc"
        Set RS = New ADODB.Recordset
        RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Espero = ""
        If Not RS.EOF Then
            Espero = "OK"
            Envio = CStr(RS!Secuencia)  'MAX secuencia
            
            While Not RS.EOF
                    Aux = Format(RS!idTrabajador, "00000")  'trab
                    Aux = Aux & "," & Format(RS!Fecha, "mm")  'mes
                    Aux = Aux & "," & Format(RS!Fecha, "dd")  'dia
                    Aux = Aux & "," & Format(RS!Hora, "hh")   'hora
                    Aux = Aux & "," & Format(RS!Hora, "nn")   'min
                    ''lo mando aunque no es importante
                    'Aux = Aux & "," & Format(rs!Fecha, "yyyy")   'año
                    'Aux = Aux & ",0000,"
                    ''mandare la secuencia
                    'Aux = Aux & Format(rs!Secuencia, "00000")
                    Aux = Aux & ",0000,0000,00000"
                    
                    
                    Print #NF, Aux
                    
                    
                    RS.MoveNext
            Wend
            
            Aux = "DELETE from entradafichajes where secuencia <= " & Envio
            conn.Execute Aux
            
            TieneDatos = True 'Tiene datos
        End If 'de belgida
        RS.Close
        Set RS = Nothing
        
    End If
    
    
    
                        
    Close #NF
   
   On Error Resume Next
    If Not TieneDatos Then
        
        lbl.Caption = "Fichero vacio"
        lbl.Refresh
        espera 0.5
        
        Err.Clear
    
    End If
    
    
    
    
    
    LeerMarcajes = True
End Function

Public Function EnviarRecibir(Envio As String, Espero As String, Optional Espero2 As String) As Integer
    Dim TiempoEspera2 As Single
    Dim FueraTiempo As Boolean
    Dim SegundaEspera As Boolean
    Dim FinBucle As Boolean ' se usa para lo que su nombre indica
    
    FueraTiempo = False
    SegundaEspera = False
    FinBucle = False
    EnviarRecibir = 0
    Socket.SendData Chr(2) & Envio & Chr(3)
    TiempoEspera2 = Timer
    mvarRecibido = ""
    While Not FinBucle
        'Debug.Print Format(Now, "nn:ss ") & mvarRecibido
        'TiempoEspera = TiempoEspera + 1
        'If TiempoEspera > 100000 Then FinBucle = True: FueraTiempo = True
        If Espero2 <> "" And InStr(1, mvarRecibido, Espero2) Then
                SegundaEspera = True
                FinBucle = True
        Else
            If InStr(1, mvarRecibido, Espero) Then
                FinBucle = True
            Else
                If InStr(1, mvarRecibido, "TIMEOUT") > 0 Then
                    FueraTiempo = True
                    FinBucle = True
                Else
                    If Timer - TiempoEspera2 > 10 Then
                        'Si es mayor que 10 segundos GRABO UN MAL
                        FueraTiempo = True
                        FinBucle = True
                    Else
                        'Dejamos que corra el tiempo
                    End If
                End If
            End If
        End If
        DoEvents
    Wend
    If FueraTiempo Then EnviarRecibir = 1
    If SegundaEspera Then EnviarRecibir = 2
End Function



Public Function GrabaHoraTerminal()
Dim sAux As String
Dim Espero As String
    If Not ComprobarConexion() Then Exit Function

    
    sAux = GetTimeInfo
    sAux = "504" + sAux
    Espero = "60411"
    If EnviarRecibir(sAux, Espero, "NO") <> 0 Then
        MsgBox "Se han producido errores", vbExclamation
    Else
        MsgBox "Terminales actualizados", vbInformation
    End If
        
End Function


Private Function GetTimeInfo() As String
  Dim iPos As Integer
  Dim sAuxDate As String
  Dim sAuxTime As String
  Dim sVal As String
  Dim vAux As Variant
  
  GetTimeInfo = ""
  
  'Agafem la data del sistema
  vAux = Now
   
  'Agafem el any.
  GetTimeInfo = GetTimeInfo + Right("00" + LTrim(Str(Year(vAux))), 2)
  
  'Agafem el mes.
  GetTimeInfo = GetTimeInfo + Right("00" + LTrim(Str(Month(vAux))), 2)
  
  'Agafem l'any.
  GetTimeInfo = GetTimeInfo + Right("00" + LTrim(Str(Day(vAux))), 2)
  
  'Agafem la hora.
  GetTimeInfo = GetTimeInfo + Right("00" + LTrim(Str(Hour(vAux))), 2)
  
  'Agafem els minuts.
  GetTimeInfo = GetTimeInfo + Right("00" + LTrim(Str(Minute(vAux))), 2)
  
  'Agafem els segons.
  GetTimeInfo = GetTimeInfo + Right("00" + LTrim(Str(Second(vAux))), 2)
  
  'Posem el dia de de setmana
  GetTimeInfo = GetTimeInfo + Right("00" + LTrim(Str(Weekday(vAux, vbMonday))), 2)
End Function





