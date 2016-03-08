Attribute VB_Name = "bus"
 'NOTA: en este m�dul, adem�s, n'hi han funcions generals que no siguen de formularis
Option Explicit

Public vEmpresa As Cempresas
Public vUsu As Usuario



'Definicion Conexi�n a BASE DE DATOS
'---------------------------------------------------
'Conexi�n a la BD PlannerTours de la empresa
Public conn As ADODB.Connection

'Para cargar datos en trozos que no hay llamadas a ningun sitio
Public miRs As ADODB.Recordset
Public miSQL As String

'Definicion de FORMATOS
'---------------------------------------------------
Public FormatoFecha As String
Public FormatoImporte As String 'Decimal(12,2)
Public FormatoPrecio As String 'Decimal(10,4)
'Public FormatoCantidad As String 'Decimal(10,2)
Public FormatoPorcen As String 'Decimal(5,2) 'Porcentajes
Public FormatoExp As String  'Expedientes

Public FormatoDec10d2 As String 'Decimal(10,2)

'Public FormatoKms As String 'Decimal(8,4)


Public teclaBuscar As Integer 'llamada desde prismaticos



Public CadenaDesdeOtroForm As String

'Global para n� de registro eliminado
Public NumRegElim  As Long

'Para algunos campos de texto suletos controlarlos
'Public miTag As CTag

'Variable para saber si se ha actualizado algun asiento
'Public AlgunAsientoActualizado As Boolean
'Public TieneIntegracionesPendientes As Boolean

Public miRsAux As ADODB.Recordset

Public AnchoLogin As String  'Para fijar los anchos de columna


'Metemos KRETA dentro de este proyecto
Public ColK2 As ColKreta2
Public GesHuellaDB As BaseDatos2

'FALTA###
'meterlas dentro de la configuracion, o
Public Servidor As String
Public ValorBD As String


Public Function AbrirConnParaUsuarios() As Boolean
Dim cad As String

    On Error GoTo EAbrirConexionU

    AbrirConnParaUsuarios = False
    Set conn = Nothing
    Set conn = New Connection
    conn.CursorLocation = adUseServer

    
    
    
    
    cad = "DSN=Aripres4;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=Usuarios;"
    
    'Server
    cad = cad & "SERVER=" & RecuperaValor(vEmpresa.Server, 1)
    
    'Usuarios
    cad = cad & ";UID=" & RecuperaValor(vEmpresa.Server, 2)
    
    'Password
    cad = cad & ";PASSWORD=" & RecuperaValor(vEmpresa.Server, 3)
    
    cad = cad & ";PORT=;OPTION=3;STMT=;"
    
    
    
     'Cad = "Provider=MSDASQL.1;Extended Properties=""DATABASE=Usuarios;DESCRIPTION=MySQL ODBC 3.51 Driver DSN;DSN=Aripres4;OPTION=3;;PORT=3306;;"
   
    
    
    
    
    conn.ConnectionString = cad
    conn.Open
    AbrirConnParaUsuarios = True
    Exit Function
    
EAbrirConexionU:
    MuestraError Err.Number, "Abrir conexi�n presencia.", Err.Description

End Function

' **** DATOS DEL LOGIN ****
'Public CodEmple As Integer
'Public codAgenc As Integer
'Public codEmpre As Integer
'Public codGrupo As Integer
'Public claEmpre As Integer
'Public TipEmple As Integer
'Public areEmple As Integer
' *************************


'Inicio Aplicaci�n
Public Sub Main()


     
     'obric la conexio
    If AbrirConexion() = False Then
        MsgBox "La aplicaci�n no puede continuar sin acceso a los datos. ", vbCritical
        End
    End If


    


    FormatoFecha = "yyyy-mm-dd"
    teclaBuscar = 43
    
    'HAY QUE PASAR LA CARGA DE ICONOS AL frmMAIN
    'De momento sigue en ppal
    Load frmPpal
    
    

    
    
    'Cargamos la empresa
    Set vEmpresa = New Cempresas
    If vEmpresa.Leer(1) = 1 Then
        Set vEmpresa = Nothing
        frmEmpresa.Show vbModal
        End
    End If
    
    
    'AHora vamos con el usuario
    If vEmpresa.Server <> "" Then
        conn.Close
        'Abrir otra conexion
        If Not AbrirConnParaUsuarios() Then
            AbrirConexion
            frmEmpresa.Show vbModal
            End
        End If
    End If
    
    
    frmIdentifica.Show vbModal
    If vUsu Is Nothing Then End
    
    If vEmpresa.Server <> "" Then
        conn.Close
        'Abrir otra conexion
        AbrirConexion
    End If
    
    
    'PARA IMPRIMIR
    'Servidor
    
    NumRegElim = InStr(1, conn.ConnectionString, ";SERVER=")
    Servidor = Mid(conn.ConnectionString, NumRegElim + 8)
    NumRegElim = InStr(1, Servidor, ";")
    Servidor = Mid(Servidor, 1, NumRegElim - 1)
    
    'BAse datos
    NumRegElim = InStr(1, conn.ConnectionString, "DATABASE=")
    ValorBD = Mid(conn.ConnectionString, NumRegElim + 9)
    NumRegElim = InStr(1, ValorBD, ";")
    ValorBD = Mid(ValorBD, 1, NumRegElim - 1)
    

    
    
    'Borramos ciertos datos temporales y demas
    conn.Execute "DELETE FROM tmpdatosmes"
    
    InicializarFormatos
    
    frmMain.Show
    
End Sub

'espera els segon que li digam
Public Function espera(Segundos As Single)
    Dim T1
    T1 = Timer
    Do
    Loop Until Timer - T1 > Segundos
End Function



Public Function AbrirConexion() As Boolean
Dim cad As String
On Error GoTo EAbrirConexion
    
    AbrirConexion = False
    Set conn = Nothing
    Set conn = New Connection
    'Conn.CursorLocation = adUseClient
    conn.CursorLocation = adUseServer

    
    
    'Este es el que hay que dejar
    cad = "DSN=Aripres4;DESC=MySQL ODBC 3.51 Driver DSN;;;;OPTION=3;STMT=;"
    cad = cad & ";Persist Security Info=true"
    
    
    'O esta
    'Cad = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=Aripres4"
    
     
       
   
    'local.
    'cad = "DSN=Aripres4;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=Aripres1;SERVER=localhost;UID=root;PASSWORD=aritel;PORT=3306;OPTION=3;STMT=;"
    'servidor ARIADNA
    'cad = "DSN=Aripres4;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=Aripres;SERVER=ARIADNA;UID=root;PASSWORD=aritel;PORT=3306;OPTION=3;STMT=;"
    
'
'    'FALTA### QUITAR###
'    'VALORES puestos A CAPON
'
'    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=aripres"
'    Cad = Cad & ";SERVER=SRV2008;"
'    Cad = Cad & ";UID=root"
'    Cad = Cad & ";PWD=aritel"
'    '---- Laura: 29/09/2006
'    Cad = Cad & ";PORT=3306;OPTION=3;STMT="
'    '----
'    Cad = Cad & ";Persist Security Info=true"
    
    
    
    
    
    conn.ConnectionString = cad
    conn.Open
    AbrirConexion = True
    Exit Function
    
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n presencia.", Err.Description
End Function


Public Function EjecutaSQL(CadenaSQL As String) As Boolean
    On Error Resume Next
    conn.Execute CadenaSQL
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
        EjecutaSQL = False
    Else
        EjecutaSQL = True
    End If
End Function


'Public Function AbrirConexionConta(Usuario As String, Pass As String) As Boolean
'Dim cad As String
'Dim nomConta As String 'nombre de la BD de la contabilidad
'Dim serConta As String 'servidor donde esta la BD de la contabilidad
'On Error GoTo EAbrirConexion
'
'    AbrirConexionConta = False
'
'    Set ConnConta = Nothing
'    Set ConnConta = New Connection
''    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
'    ConnConta.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
'
'    'Obtener la BD de contabilidad
''    serConta = "serconta"
''    nomConta = DevuelveDesdeBDNew(cPTours, "paramcon", "bdaconta", "codempre", CStr(codEmpre), "N", serConta)
'
''    nomConta = DevuelveDesdeBDNew(cPTours, "empresas", "bdaconta", "codempre", CStr(vSesion.Empresa), "N")
'
'
'
''    vEmpresa.BDConta = nomConta
'    If vEmpresa.BDConta <> "" Then
'    '    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=conta" & vParamConta.NumeroConta & ";SERVER=" & vParamConta.ServidorConta & ";"
'    '    cad = cad & ";UID=" & vParamConta.UsuarioConta
'    '    cad = cad & ";PWD=" & vParamConta.PasswordConta
'    '    cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=conta2;SERVER=david;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        If serConta <> "" Then 'especificamos servidor
'            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & vEmpresa.BDConta & ";SERVER=" & serConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        Else 'por defecto cogera la BD del servidor que haya en el ODBC
'            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & vEmpresa.BDConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        End If
'        ConnConta.ConnectionString = cad
'        ConnConta.Open
'        ConnConta.Execute "Set AUTOCOMMIT = 1"
'        AbrirConexionConta = True
'    Else
'        AbrirConexionConta = False
'    End If
'    Exit Function
'EAbrirConexion:
'    MuestraError Err.Number, "Abrir conexi�n Contabilidad.", Err.Description
'End Function
'
'
'Public Function AbrirConexionAuxCon(Empresa As String, Usuario As String, Pass As String) As Boolean
'Dim cad As String
'Dim nomConta As String 'nombre de la BD de la contabilidad
'Dim serConta As String 'servidor donde esta la BD de la contabilidad
'On Error GoTo EAbrirConexion
'
'    AbrirConexionAuxCon = False
'
'    Set ConnAuxCon = Nothing
'    Set ConnAuxCon = New Connection
''    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
'    ConnAuxCon.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
'
'    'Obtener la BD de contabilidad
''    SQL = "select bdaconta FROM paramcon WHERE codempre=" & codEmpre
'    serConta = "serconta"
'    nomConta = DevuelveDesdeBDNew(cPTours, "paramcon", "bdaconta", "codempre", Empresa, "N", serConta)
''    vEmpresa.BDConta = nomConta
'    If nomConta <> "" Then
'    '    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=conta" & vParamConta.NumeroConta & ";SERVER=" & vParamConta.ServidorConta & ";"
'    '    cad = cad & ";UID=" & vParamConta.UsuarioConta
'    '    cad = cad & ";PWD=" & vParamConta.PasswordConta
'    '    cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=conta2;SERVER=david;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        If serConta <> "" Then 'especificamos servidor
'            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";SERVER=" & serConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        Else 'por defecto cogera la BD del servidor que haya en el ODBC
'            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        End If
'        ConnAuxCon.ConnectionString = cad
'        ConnAuxCon.Open
'        ConnAuxCon.Execute "Set AUTOCOMMIT = 1"
'        AbrirConexionAuxCon = True
'    Else
'        AbrirConexionAuxCon = False
'    End If
'    Exit Function
'EAbrirConexion:
'    MuestraError Err.Number, "Abrir conexi�n Contabilidad.", Err.Description
'End Function
'
'
'
'
'
'
'
'Public Function LeerDatosEmpresa()
''Crea instancia de la clase Cempresa con los valores en
''Tabla: Empresas
''BDatos: PTours y Conta
'
'    Set vEmpresa = New Cempresa
'    If vEmpresa.LeerDatos(vSesion.Empresa) = False Then  'De PlannerTours
'        MsgBox "No se han podido cargar los datos de la empresa. Debe configurar la aplicaci�n.", vbExclamation
'        Set vEmpresa = Nothing
'        Set vSesion = Nothing
'        Set Conn = Nothing
'        End
'    End If
'
'    'Abrir conexi�n a la BDatos de Contabilidad para acceder a
'    'Tablas: Cuentas, Tipos IVA,...
'    If AbrirConexionConta("root", "aritel") = False Then
'        MsgBox "La aplicaci�n no puede continuar sin acceso a los datos de Contabilidad. ", vbCritical
'        AccionesCerrar
'        End
'    End If
'
'    If vEmpresa.LeerNiveles(vSesion.Empresa) = False Then 'De Contabilidad
'        MsgBox "No se han podido cargar los niveles de la contabilidad de la empresa. Debe configurar la aplicaci�n.", vbExclamation
'        AccionesCerrar
'        End
'    End If
'End Function
'
'


Public Function PonerDatosPpal()
Dim cad As String

  '  cad = DevuelveDesdeBDNew(cPTours, "agencias", "desagenc", "codagenc", vSesion.Agencia, "N")
'    If cad <> "" Then cad = "   -  Agencia: " & cad
'
'    Select Case vSesion.AreaEmple
'        Case 1 'Administracion
''            frmPal.Caption = "PlannerTours" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  Empresa: " & vEmpresa.nomEmpre & cad
'            MDIppal.Caption = "PlannerTours" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  Empresa: " & vEmpresa.nomEmpre & cad
'        Case 2 'Minorista
''            frmPal3.Caption = "PlannerTours" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  Empresa: " & vEmpresa.nomEmpre & cad
'        Case 3 'Mayorista
''            frmPal2.Caption = "PlannerTours" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  Empresa: " & vEmpresa.nomEmpre & cad
'    End Select
'    If Err.Number <> 0 Then MuestraError Err.Description, "Poniendo datos de la pantalla principal", Err.Description
End Function

    

Public Sub MuestraError(Numero As Long, Optional CADENA As String, Optional Desc As String)
    Dim cad As String
    Dim Aux As String
    
    'Con este sub pretendemos unificar el msgbox para todos los errores
    'que se produzcan
    On Error Resume Next
    cad = "Se ha producido un error: " & vbCrLf
    If CADENA <> "" Then
        cad = cad & vbCrLf & CADENA & vbCrLf & vbCrLf
    End If
    'Numeros de errores que contolamos
    If conn.Errors.Count > 0 Then
        'ControlamosError Aux
        conn.Errors.Clear
    Else
        Aux = ""
    End If
    If Aux <> "" Then Desc = Aux
    If Desc <> "" Then cad = cad & vbCrLf & Desc & vbCrLf & vbCrLf
    If Aux = "" Then cad = cad & "N�mero: " & Numero & vbCrLf & "Descripci�n: " & Error(Numero)
    MsgBox cad, vbExclamation
End Sub

Public Function DBSet(vData As Variant, Tipo As String, Optional EsNulo As String) As Variant
'Establece el valor del dato correcto antes de Insertar en la BD
Dim cad As String

        If IsNull(vData) Then
            DBSet = ValorNulo
            Exit Function
        End If

        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    If vData = "" Then
                        If EsNulo = "N" Then
                            DBSet = "''"
                        Else
                            DBSet = ValorNulo
                        End If
                    Else
                        cad = (CStr(vData))
                        NombreSQL cad
                        DBSet = "'" & cad & "'"
                    End If
                    
                Case "N"    'Numero
                    If vData = "" Or vData = 0 Then
                        If EsNulo <> "" Then
                            If EsNulo = "S" Then
                                DBSet = ValorNulo
                            Else
                                DBSet = 0
                            End If
                        Else
                            DBSet = 0
                        End If
                    Else
                        cad = CStr(ImporteFormateado(CStr(vData)))
                        DBSet = TransformaComasPuntos(cad)
                    End If
                    
                Case "F"    'Fecha
'                     '==David
''                    DBLet = "0:00:00"
'                     '==Laura
                    If vData = "" Then
                        If EsNulo = "S" Then
                            DBSet = ValorNulo
                        Else
                            DBSet = "'1900-01-01'"
                        End If
                    Else
                        DBSet = "'" & Format(vData, FormatoFecha) & "'"
                    End If
                    
                Case "FH" 'Fecha/Hora
                    If vData = "" Then
                        If EsNulo = "S" Then DBSet = ValorNulo
                    Else
                        DBSet = "'" & Format(vData, "yyyy-mm-dd hh:mm:ss") & "'"
                    End If
                    
                Case "H" 'Hora
                    If vData = "" Then
                    Else
                        DBSet = "'" & Format(vData, "hh:mm:ss") & "'"
                    End If
                    
                Case "B"  'Boolean
                    If vData Then
                        DBSet = 1
                    Else
                        DBSet = 0
                    End If
            End Select
        End If
End Function

Public Function DBLetMemo(vData As Variant) As Variant
    On Error Resume Next
    
    DBLetMemo = vData
    
    
    
    If Err.Number <> 0 Then
        Err.Clear
        DBLetMemo = ""
    End If
End Function



Public Function DBLet(vData As Variant, Optional Tipo As String) As Variant
'Para cuando recupera Datos de la BD
    If IsNull(vData) Then
        DBLet = ""
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    DBLet = ""
                Case "N"    'Numero
                    DBLet = 0
                Case "F"    'Fecha
                     '==David
                    DBLet = "0:00:00"
                     '==Laura
'                     DBLet = "0000-00-00"
'                      DBLet = ""
                Case "D"
                    DBLet = 0
                Case "B"  'Boolean
                    DBLet = False
                Case Else
                    DBLet = ""
            End Select
        End If
    Else
        DBLet = vData
    End If
End Function

'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo antes de bloquear
'   Prepara la conexion para bloquear
Public Sub PreparaBloquear()
    conn.Execute "commit"
    conn.Execute "set autocommit=0"
End Sub

'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo despues de un bloque
'   Prepara la conexion para bloquear
Public Sub TerminaBloquear()
    conn.Execute "commit"
    conn.Execute "set autocommit=1"
End Sub

'///////////////////////////////////////////////////////////////
'
'   Cogemos un numero formateado: 1.256.256,98  y deevolvemos 1256256,98
'   Tiene que venir num�rico
Public Function ImporteFormateado(Importe As String) As Currency
Dim i As Integer

    If Importe = "" Then
        ImporteFormateado = 0
    Else
        'Primero quitamos los puntos
        Do
            i = InStr(1, Importe, ".")
            If i > 0 Then Importe = Mid(Importe, 1, i - 1) & Mid(Importe, i + 1)
        Loop Until i = 0
        ImporteFormateado = Importe
    End If
End Function

'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaComasPuntos(CADENA As String) As String
Dim i As Integer
    Do
        i = InStr(1, CADENA, ",")
        If i > 0 Then
            CADENA = Mid(CADENA, 1, i - 1) & "." & Mid(CADENA, i + 1)
        End If
    Loop Until i = 0
    TransformaComasPuntos = CADENA
End Function

'Para los nombre que pueden tener ' . Para las comillas habra que hacer dentro otro INSTR
Public Sub NombreSQL(ByRef CADENA As String)
Dim J As Integer
Dim i As Integer
Dim Aux As String
    
    
    'Buscamos los \
    J = 1
    Do
        i = InStr(J, CADENA, "\")
        If i > 0 Then
            CADENA = Mid(CADENA, 1, i) & "\" & Mid(CADENA, i + 1)
            J = i + 2
        End If
    Loop Until i = 0
    
    
    
    J = 1
    Do
        i = InStr(J, CADENA, "'")
        If i > 0 Then
            Aux = Mid(CADENA, 1, i - 1) & "\"
            CADENA = Aux & Mid(CADENA, i)
            J = i + 2
        End If
    Loop Until i = 0
    
    

    
End Sub

Public Function EsFechaOKString(ByRef T As String) As Boolean
Dim cad As String
    
    cad = T
    If InStr(1, cad, "/") = 0 Then
        If Len(T) = 8 Then
            cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        Else
            If Len(T) = 6 Then cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        End If
    End If
    If IsDate(cad) Then
        EsFechaOKString = True
        T = Format(cad, "dd/mm/yyyy")
    Else
        EsFechaOKString = False
    End If
End Function


Public Function EsFechaOK(ByRef T As TextBox) As Boolean
Dim cad As String
    
    cad = T
    If InStr(1, cad, "/") = 0 Then
        If Len(T) = 8 Then
            cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        Else
            If Len(T) = 6 Then cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        End If
    End If
    If IsDate(cad) Then
        EsFechaOK = True
        T.Text = Format(cad, "dd/mm/yyyy")
    Else
        EsFechaOK = False
    End If
End Function



Public Function DevuelveDesdeBD(kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional ByRef otroCampo As String) As String
    Dim rs As Recordset
    Dim cad As String
    Dim Aux As String
    
    On Error GoTo EDevuelveDesdeBD
    DevuelveDesdeBD = ""
    cad = "Select " & kCampo
    If otroCampo <> "" Then cad = cad & ", " & otroCampo
    cad = cad & " FROM " & Ktabla
    cad = cad & " WHERE " & Kcodigo & " = "
    If Tipo = "" Then Tipo = "N"
    Select Case Tipo
    Case "N"
        'No hacemos nada
        cad = cad & ValorCodigo
    Case "T", "F"
        cad = cad & "'" & ValorCodigo & "'"
    Case Else
        MsgBox "Tipo : " & Tipo & " no definido", vbExclamation
        Exit Function
    End Select
    
    
    
    'Creamos el sql
    Set rs = New ADODB.Recordset
    rs.Open cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not rs.EOF Then
        DevuelveDesdeBD = DBLet(rs.Fields(0))
        If otroCampo <> "" Then otroCampo = DBLet(rs.Fields(1))
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
EDevuelveDesdeBD:
        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
End Function



''Este metodo sustituye a DevuelveDesdeBD
''Funciona para claves primarias formadas por 2 campos
'Public Function DevuelveDesdeBDnew(vBD As Byte, Ktabla As String, kCampo As String, Kcodigo1 As String, valorCodigo1 As String, Optional tipo1 As String, Optional ByRef otroCampo As String, Optional KCodigo2 As String, Optional ValorCodigo2 As String, Optional tipo2 As String) As String
''IN: vBD --> Base de Datos a la que se accede
'Dim RS As Recordset
'Dim cad As String
'Dim Aux As String
'
'On Error GoTo EDevuelveDesdeBDnew
'    DevuelveDesdeBDnew = ""
'    If valorCodigo1 = "" And ValorCodigo2 = "" Then Exit Function
'    cad = "Select " & kCampo
'    If otroCampo <> "" Then cad = cad & ", " & otroCampo
'    cad = cad & " FROM " & Ktabla
'    cad = cad & " WHERE " & Kcodigo1 & " = "
'    If tipo1 = "" Then tipo1 = "N"
'    Select Case tipo1
'        Case "N"
'            'No hacemos nada
'            If IsNumeric(valorCodigo1) Then
'                cad = cad & Val(valorCodigo1)
'            Else
'                MsgBox "El campo debe ser num�rico.", vbExclamation
'                DevuelveDesdeBDnew = "Error"
'                Exit Function
'            End If
'        Case "T", "F"
'            cad = cad & "'" & valorCodigo1 & "'"
'        Case Else
'            MsgBox "Tipo : " & tipo1 & " no definido", vbExclamation
'            Exit Function
'    End Select
'
'    If KCodigo2 <> "" Then
'        cad = cad & " AND " & KCodigo2 & " = "
'        If tipo2 = "" Then tipo2 = "N"
'        Select Case tipo2
'        Case "N"
'            'No hacemos nada
'            If ValorCodigo2 = "" Then
'                cad = cad & "-1"
'            Else
'                cad = cad & Val(ValorCodigo2)
'            End If
'        Case "T"
'            cad = cad & "'" & ValorCodigo2 & "'"
'        Case "F"
'            cad = cad & "'" & Format(ValorCodigo2, FormatoFecha) & "'"
'        Case Else
'            MsgBox "Tipo : " & tipo2 & " no definido", vbExclamation
'            Exit Function
'        End Select
'    End If
'
'
'    'Creamos el sql
'    Set RS = New ADODB.Recordset
'
'    Select Case vBD
'        Case cPTours 'vBD=1: PlannerTours
'            RS.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'        Case cConta 'BD 2: Contabilidad
'            RS.Open cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
'        Case 3 'vBD=3: contabilidad distinta a la de la empresa conectada
'            RS.Open cad, ConnAuxCon, adOpenForwardOnly, adLockOptimistic, adCmdText
'    End Select
''    If vBD = cPTours Then 'vBD=1: PlannerTours
''        RS.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
''    ElseIf vBD = cConta Then  'BD 2: Contabilidad
''        RS.Open cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
''    End If
'
'    If Not RS.EOF Then
'        DevuelveDesdeBDnew = DBLet(RS.Fields(0))
'        If otroCampo <> "" Then otroCampo = DBLet(RS.Fields(1))
'    End If
'    RS.Close
'    Set RS = Nothing
'    Exit Function
'
'EDevuelveDesdeBDnew:
'        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
'End Function


'LAURA
'Este metodo sustituye a DevuelveDesdeBD
'Funciona para claves primarias formadas por 3 campos
Public Function DevuelveDesdeBDNew(vBD As Byte, Ktabla As String, kCampo As String, Kcodigo1 As String, valorCodigo1 As String, Optional tipo1 As String, Optional ByRef otroCampo As String, Optional KCodigo2 As String, Optional ValorCodigo2 As String, Optional tipo2 As String, Optional KCodigo3 As String, Optional ValorCodigo3 As String, Optional tipo3 As String) As String
'IN: vBD --> Base de Datos a la que se accede
Dim rs As Recordset
Dim cad As String
Dim Aux As String
    
On Error GoTo EDevuelveDesdeBDnew
    DevuelveDesdeBDNew = ""
'    If valorCodigo1 = "" And ValorCodigo2 = "" Then Exit Function
    cad = "Select " & kCampo
    If otroCampo <> "" Then cad = cad & ", " & otroCampo
    cad = cad & " FROM " & Ktabla
    If Kcodigo1 <> "" Then
        cad = cad & " WHERE " & Kcodigo1 & " = "
        If tipo1 = "" Then tipo1 = "N"
    Select Case tipo1
        Case "N"
            'No hacemos nada
            cad = cad & Val(valorCodigo1)
        Case "T"
            cad = cad & DBSet(valorCodigo1, "T")
        Case "F"
            cad = cad & "'" & valorCodigo1 & "'"
        Case Else
            MsgBox "Tipo : " & tipo1 & " no definido", vbExclamation
            Exit Function
    End Select
    End If
    
    If KCodigo2 <> "" Then
        cad = cad & " AND " & KCodigo2 & " = "
        If tipo2 = "" Then tipo2 = "N"
        Select Case tipo2
        Case "N"
            'No hacemos nada
            If ValorCodigo2 = "" Then
                cad = cad & "-1"
            Else
                cad = cad & Val(ValorCodigo2)
            End If
        Case "T"
'            cad = cad & "'" & ValorCodigo2 & "'"
            cad = cad & DBSet(ValorCodigo2, "T")
        Case "F"
            cad = cad & "'" & Format(ValorCodigo2, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo2 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    If KCodigo3 <> "" Then
        cad = cad & " AND " & KCodigo3 & " = "
        If tipo3 = "" Then tipo3 = "N"
        Select Case tipo3
        Case "N"
            'No hacemos nada
            If ValorCodigo3 = "" Then
                cad = cad & "-1"
            Else
                cad = cad & Val(ValorCodigo3)
            End If
        Case "T"
            cad = cad & "'" & ValorCodigo3 & "'"
        Case "F"
            cad = cad & "'" & Format(ValorCodigo3, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo3 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    
    'Creamos el sql
    Set rs = New ADODB.Recordset
    
   ' If vBD = 1 Then 'BD 1: Ariges
        rs.Open cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
   ' Else    'BD 2: Conta
      '  RS.Open cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
   ' End If
    
    If Not rs.EOF Then
        DevuelveDesdeBDNew = DBLet(rs.Fields(0))
        If otroCampo <> "" Then otroCampo = DBLet(rs.Fields(1))
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
    
EDevuelveDesdeBDnew:
        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
End Function




'CESAR
Public Function DevuelveDesdeBDnew2(kBD As Integer, kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional num As Byte, Optional ByRef otroCampo As String) As String
Dim rs As Recordset
Dim cad As String
Dim Aux As String
Dim v_aux As Integer
Dim Campo As String
Dim valor As String
Dim tip As String

On Error GoTo EDevuelveDesdeBDnew2
DevuelveDesdeBDnew2 = ""

cad = "Select " & kCampo
If otroCampo <> "" Then cad = cad & ", " & otroCampo
cad = cad & " FROM " & Ktabla

If Kcodigo <> "" Then cad = cad & " where "

For v_aux = 1 To num
    Campo = RecuperaValor(Kcodigo, v_aux)
    valor = RecuperaValor(ValorCodigo, v_aux)
    tip = RecuperaValor(Tipo, v_aux)
        
    cad = cad & Campo & "="
    If tip = "" Then Tipo = "N"
    
    Select Case tip
            Case "N"
                'No hacemos nada
                cad = cad & valor
            Case "T", "F"
                cad = cad & "'" & valor & "'"
            Case Else
                MsgBox "Tipo : " & tip & " no definido", vbExclamation
            Exit Function
    End Select
    
    If v_aux < num Then cad = cad & " AND "
  Next v_aux

'Creamos el sql
Set rs = New ADODB.Recordset
Select Case kBD
    Case 1
        rs.Open cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
End Select

If Not rs.EOF Then
    DevuelveDesdeBDnew2 = DBLet(rs.Fields(0))
    If otroCampo <> "" Then otroCampo = DBLet(rs.Fields(1))
Else
     If otroCampo <> "" Then otroCampo = ""
End If
rs.Close
Set rs = Nothing
Exit Function
EDevuelveDesdeBDnew2:
    MuestraError Err.Number, "Devuelve DesdeBDnew2.", Err.Description
End Function


Public Function EsEntero(Texto As String) As Boolean
Dim i As Integer
Dim C As Integer
Dim L As Integer
Dim res As Boolean

    res = True
    EsEntero = False

    If Not IsNumeric(Texto) Then
        res = False
    Else
        'Vemos si ha puesto mas de un punto
        C = 0
        L = 1
        Do
            i = InStr(L, Texto, ".")
            If i > 0 Then
                L = i + 1
                C = C + 1
            End If
        Loop Until i = 0
        If C > 1 Then res = False
        
        'Si ha puesto mas de una coma y no tiene puntos
        If C = 0 Then
            L = 1
            Do
                i = InStr(L, Texto, ",")
                If i > 0 Then
                    L = i + 1
                    C = C + 1
                End If
            Loop Until i = 0
            If C > 1 Then res = False
        End If
        
    End If
        EsEntero = res
End Function

Public Function TransformaPuntosComas(CADENA As String) As String
    Dim i As Integer
    Do
        i = InStr(1, CADENA, ".")
        If i > 0 Then
            CADENA = Mid(CADENA, 1, i - 1) & "," & Mid(CADENA, i + 1)
        End If
        Loop Until i = 0
    TransformaPuntosComas = CADENA
End Function



Public Function TransformaPunto2Puntos(CADENA As String) As String
    Dim i As Integer
    Do
        i = InStr(1, CADENA, ".")
        If i > 0 Then
            CADENA = Mid(CADENA, 1, i - 1) & ":" & Mid(CADENA, i + 1)
        End If
        Loop Until i = 0
    TransformaPunto2Puntos = CADENA
End Function


Public Sub InicializarFormatos()
    FormatoFecha = "yyyy-mm-dd"
'    FormatoFechaHora = "yyyy-mm-dd hh:mm:ss"
    FormatoImporte = "#,###,###,##0.00"  'Decimal(12,2)
    FormatoPrecio = "###,##0.00##"  'Decimal(10,4)
'    FormatoCantidad = "##,###,##0.00"   'Decimal(10,2)
    FormatoPorcen = "##0.00" 'Decima(5,2) para porcentajes
    
    FormatoDec10d2 = "##,###,##0.00"   'Decimal(10,2)
    
    FormatoExp = "0000000000"
'    FormatoKms = "#,##0.00##" 'Decimal(8,4)
End Sub


'Public Sub AccionesCerrar()
''cosas que se deben hacen cuando finaliza la aplicacion
'    On Error Resume Next
'
'    'cerrar clases q estan abiertas durante la ejecucion
'    Set vEmpresa = Nothing
'    Set vSesion = Nothing
'
''    Set vParam = Nothing
''    Set vParamAplic = Nothing
''    Set vParamConta = Nothing
'
'
'    'Cerrar Conexiones a bases de datos
'    Conn.Close
'    ConnConta.Close
'    Set Conn = Nothing
'    Set ConnConta = Nothing
'
'    If Err.Number <> 0 Then Err.Clear
'End Sub


Public Function DevuelveTextoIncidencia(vId As Integer, Optional ByRef vSigno As Single) As String
Dim rs As ADODB.Recordset
Dim SQL As String

    DevuelveTextoIncidencia = ""
    vSigno = 0
    Set rs = New ADODB.Recordset
    SQL = "SELECT * From Incidencias " & _
        " WHERE Incidencias.IdInci =" & vId
    rs.Open SQL, conn, , , adCmdText
    If Not rs.EOF Then
            DevuelveTextoIncidencia = rs.Fields(1)
            If rs!excesodefecto Then
                vSigno = -1
                Else
                vSigno = 1
            End If
    End If
    rs.Close
    Set rs = Nothing
End Function





Public Function DevuelveCodigo(vNUmTar) As Long
Dim rs As ADODB.Recordset
Dim SQL As String
    DevuelveCodigo = -1
    Set rs = New ADODB.Recordset
    SQL = "SELECT idTrabajador From Trabajadores " & _
        " WHERE NumTarjeta ='" & vNUmTar & "'"
    rs.Open SQL, conn, , , adCmdText
    If Not rs.EOF Then
            DevuelveCodigo = rs.Fields(0)
    End If
    rs.Close
    Set rs = Nothing
End Function







Public Function MarcajesCorrectos(Correctos As Boolean, vSQL As String) As Boolean
Dim rs As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    If vSQL <> "" Then
        vSQL = vSQL & " AND "
    Else
        vSQL = vSQL & " WHERE "
    End If
    vSQL = vSQL & " correcto = " & Abs(Correctos)
    vSQL = "Select count(*) from marcajes " & vSQL
    rs.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    MarcajesCorrectos = False
    If Not rs.EOF Then
        If DBLet(rs.Fields(0), "N") > 0 Then MarcajesCorrectos = True
    End If
    rs.Close
    Set rs = Nothing
End Function



Public Function ComprobarMarcajesCorrectos(FI As Date, FF As Date, Correctos As Boolean) As Byte
Dim rs As ADODB.Recordset
Dim cad As String
Dim C As Long
    
    ComprobarMarcajesCorrectos = 127
    C = 0
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    'SQL. Marcajes incorrectos entre las dos fechas
    cad = "Select count(Entrada) "
    cad = cad & " From Secciones, Trabajadores, Marcajes"
    cad = cad & " WHERE  Secciones.IdSeccion = Trabajadores.Seccion AND"
    cad = cad & " Trabajadores.idTrabajador = Marcajes.idTrabajador"
    
    cad = cad & " AND Fecha>='" & Format(FI, FormatoFecha) & "'"
    cad = cad & " AND Fecha<='" & Format(FF, FormatoFecha) & "'"
    cad = cad & " AND Correcto="
    If Correctos Then
        cad = cad & " True"
    Else
        cad = cad & " False"
    End If
    rs.Open cad, conn, , , adCmdText
    If Not rs.EOF Then
        If Not IsNull(rs.Fields(0)) Then C = rs.Fields(0)
    End If
    rs.Close
    'Si c>0 entonces tiene marcajes incorrectos
    If C > 0 Then
        ComprobarMarcajesCorrectos = 1
        Else
            ComprobarMarcajesCorrectos = 0
    End If
End Function

