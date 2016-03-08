Attribute VB_Name = "l_ppal"
Option Explicit

Public Conn As Connection
Public RS As ADODB.Recordset
Public Fecha As Date

Public Sub Main()
    If Not App.PrevInstance Then

        If AbrirConexion() Then
            Set RS = New ADODB.Recordset
            
            
            
            RS.Open "Select curdate(),curtime()", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If RS.EOF Then
                MsgBox "Error leyendo fecha servidor MYSQL", vbExclamation
                
            Else
                
                Fecha = CDate(Format(RS.Fields(0), "dd/mm/yyyy") & " " & Format(RS.Fields(1), "hh:mm:ss"))
                RS.Close
                frmPpalMar.Show vbModal
                
            End If
        
        
        End If  'De abrir conexion
        
    End If
End Sub


Public Function AbrirConexion() As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion
    
    AbrirConexion = False
    Set Conn = Nothing
    Set Conn = New Connection
    'Conn.CursorLocation = adUseClient
    Conn.CursorLocation = adUseServer
    Cad = "DSN=Aripres4;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=aripres1;;PORT=3306;OPTION=3;STMT=;"
    Cad = "DSN=Aripres4;DESC=MySQL ODBC 3.51 Driver DSN;;;;OPTION=3;STMT=;"
    Conn.ConnectionString = Cad
    Conn.Open
    AbrirConexion = True
    Exit Function
    
EAbrirConexion:
    MsgBox "Abrir conexión presencia." & vbCrLf & vbCrLf & Err.Description, vbCritical
End Function

