VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BaseDatos2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mtipo As String


'ANTES ESTABA LA variable private. Ahora no la toco
Private conn As ADODB.Connection

Private mrs As ADODB.Recordset

Public Property Let Tipo(ByVal vData As String)
    mtipo = vData
End Property

Public Property Get Tipo() As String
    Tipo = mtipo
End Property


Public Function abrir(vODBC As String, vUsu As String, vPass As String) As Boolean
    '-- Abrimos la conexión en función del ODBC pasado
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;" & _
                            "Data Source=" & vODBC & _
                            ";UID=" & vUsu & _
                            ";PWD=" & vPass
    conn.Open
End Function

Public Function AbrirConexionDavid(StringConexion As String)
    'Para dejar este trozo como esta...
    'en lugar de abri conexion de arriba hare asignar la variable de la CONN
    Set conn = New ADODB.Connection
    conn.ConnectionString = StringConexion
    conn.Open
    
End Function

'antes Public Function abrir_MYSQL(vServidor As String, vBaseDatos As String, vUsu As String, vPass As String) As Boolean
Public Function abrir_MYSQL(cadenaconexion As String) As Boolean
    '-- Abrimos la conexión en función de los datos pasados
    Set conn = New ADODB.Connection
    On Error GoTo abrir_MYSQL_error
    'conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};DESC=" & _
                        ";DATABASE=" & vBaseDatos & _
                        ";SERVER=" & vServidor & _
                        ";UID=" & vUsu & _
                        ";PWD=" & vPass & _
                        ";PORT=3306;OPTION=3;STMT="
                        
    conn.ConnectionString = cadenaconexion
    conn.Open
    mtipo = "MYSQL"
    abrir_MYSQL = True
    Exit Function
abrir_MYSQL_error:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "BaseDatos.abrir_MYSQL"
End Function

Public Function abrirL(vUsu As String, vPass As String, vBaseDatos As String, vServer As String) As Boolean
    '-- Abrimos la conexión en función del ODBC pasado
    On Error GoTo err_abrirL
    Set conn = New ADODB.Connection
    conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & vBaseDatos & ";SERVER=" & vServer & ";" _
                & ";UID=" & vUsu _
                 & ";PWD=" & vPass
    conn.Open
    abrirL = True
    Exit Function
err_abrirL:
    MsgBox Err.Description & "(Err:" & Err.Number & ")", vbCritical
    abrirL = False
End Function



Public Function ejecutar(vSQL As String) As Boolean
    If Not (conn Is Nothing) Then
        conn.Execute vSQL
    Else
        '-- Falta la conexión y no podemos continuar
    End If
End Function

Public Function cursor(vSQL As String) As ADODB.Recordset
    '-- Devuelve un recordset con los valores obtenidos por SQL
    Set mrs = New ADODB.Recordset
    mrs.Open vSQL, conn, adOpenForwardOnly
    Set cursor = mrs
End Function

Public Function Fecha(vData) As String
    '-- Devuelve un string con la fecha correspondiente
    If Not IsNull(vData) And IsDate(vData) Then
        Select Case mtipo
            Case "ACCESS"
                Fecha = "#" & Format(vData, "yyyy/mm/dd") & "#"
            Case "MYSQL"
                Fecha = "'" & Format(vData, "yyyy-mm-dd") & "'"
            Case "MULTIBASE"
                Fecha = "'" & Format(vData, "dd/mm/yyyy") & "'"
            Case Default
                Fecha = "'" & Format(vData, "yyyy-mm-dd") & "'"
        End Select
    Else
        Fecha = "NULL"
    End If
End Function

Public Function Hora(vData) As String
    '-- Devuelve un string con la fecha correspondiente
    If Not IsNull(vData) And IsDate(vData) Then
        Select Case mtipo
            Case "ACCESS"
                Hora = "#" & Format(vData, "hh:mm:ss") & "#"
            Case "MYSQL"
                Hora = "'" & Format(vData, "hh:mm:ss") & "'"
            Case "MULTIBASE"
                Hora = "'" & Format(vData, "hh:mm:ss") & "'"
            Case Default
                Hora = "'" & Format(vData, "hh:mm:ss") & "'"
        End Select
    Else
        Hora = "NULL"
    End If
End Function

Public Function Numero(vData, Optional decimales As Integer) As String
    If Not IsNull(vData) Then
        Numero = CStr(vData)
        If decimales Then
            Numero = TransformaComasPuntos(CStr(vData / (10 ^ decimales)))
        Else
            Numero = TransformaComasPuntos(CStr(vData))
        End If
    Else
        Numero = "0"
    End If
    If Numero = "" Then Numero = 0
End Function

Public Function Texto(vData) As String
    If Not IsNull(vData) Then
        Texto = "'" & vData & "'"
    Else
        Texto = "''"
    End If
End Function

