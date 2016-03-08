VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmHuella 
   Caption         =   "Datos reloj huella"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "NGAC_LOG"
      Top             =   960
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   6735
      Begin VB.Label lblIndicador 
         Caption         =   "Label2"
         Height          =   495
         Left            =   600
         TabIndex        =   6
         Top             =   480
         Width           =   5655
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   3360
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdLeer 
      Caption         =   "Leer "
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdBuscar 
      Height          =   375
      Left            =   6600
      Picture         =   "frmHuella.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label Label2 
      Caption         =   "Tabla"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Base datos"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmHuella.frx":6852
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmHuella"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Cnn As Connection
Dim Cad As String

Private Sub cmdBuscar_Click()
    cd1.InitDir = App.Path
    cd1.Filter = "mdb|*.mdb"
    cd1.FileName = ""
    cd1.ShowOpen
    If cd1.FileName <> "" Then
        Text1.Text = cd1.FileName
        LeerGuardar False
    End If
End Sub

Private Sub cmdCancel_Click()

    Unload Me
End Sub

Private Sub cmdLeer_Click()
Dim Ok As Boolean

    If Text1.Text = "" Then Exit Sub
    If Dir(Text1.Text, vbArchive) = "" Then
        MsgBox "DB no encontrada", vbExclamation
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Botones False
    Me.lblIndicador.Caption = "Proceso inciado.."
    Me.Frame1.Visible = True
    Me.Refresh
    espera 1
    Ok = LeerDatosBD
    
    
    Botones True
    Me.lblIndicador.Caption = ""
    Me.Frame1.Visible = False
    Screen.MousePointer = vbDefault
    If Ok Then Unload Me
End Sub

Private Sub Botones(Habilitar As Boolean)
    Me.cmdBuscar.Enabled = Habilitar
    Me.cmdCancel.Enabled = Habilitar
    Me.cmdLeer.Enabled = Habilitar
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal1.Icon
    LeerGuardar True
    Frame1.Visible = False
    Screen.MousePointer = vbDefault
End Sub


Private Sub LeerGuardar(Leer As Boolean)
Dim NF As Integer
Dim C As String
On Error GoTo EL
    NF = FreeFile
    C = App.Path & "\huell.dat"
    Cad = ""
    If Leer Then
        If Dir(C, vbArchive) <> "" Then
            
            Open C For Input As #NF
            Line Input #NF, Cad
            Close #NF
        End If
        Text1.Text = Cad
    Else
        
            Open C For Output As #NF
            Print #NF, Text1.Text
            Close #NF
    End If
    
    Exit Sub
EL:
    MuestraError Err.Number
End Sub



Private Function LeerDatosBD() As Boolean
Dim rOrigen As ADODB.Recordset
Dim RT As ADODB.Recordset
Dim TodoBien As Boolean
Dim Usuario As String
Dim Sec As Long
Dim Seguir As Boolean

Dim SecuenciaIncial As Long

On Error GoTo ELee

    LeerDatosBD = False
    
    
    'Trabajadores
    Set RT = New ADODB.Recordset
    RT.Open "Select idTrabajador,Numtarjeta from trabajadores", Conn, adOpenKeyset, adLockPessimistic, adCmdText
    If RT.EOF Then
        RT.Close
        MsgBox "Error leyendo datos trabajadores", vbExclamation
        Exit Function
    End If
    
    
    'preparar DAOTS
    PreparaLogDatos
    
    Cad = DevuelveDesdeBD("max(secuencia)", "entradafichajes", "1", "1")
    If Cad = "" Then Cad = "0"
    Sec = Val(Cad) + 1
    SecuenciaIncial = Sec
    
    'Preparar estructura
    
    Set Cnn = New ADODB.Connection
    
    Cad = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Text1.Text & ";"
    Cad = Cad & "User Id=Admin;Password=123456;"
    Cad = Cad & "Persist Security Info=Fal"

    'Ahora
    Cad = "driver={Microsoft Access Driver (*.mdb)};dbq=" & Text1.Text & ";uid=admin;pwd=123456"


    
    Cnn.ConnectionString = Cad
    
    Cnn.CursorLocation = adUseServer
    Cnn.Open

    
        
    

    Cnn.BeginTrans
    Conn.BeginTrans
    Cad = "Select * from " & Text2.Text  ' "
    Set rOrigen = New ADODB.Recordset
    rOrigen.Open Cad, Cnn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TodoBien = True
    
    While Not rOrigen.EOF
        'Labales
        lblIndicador.Caption = "secuencia: " & rOrigen!logindex
        lblIndicador.Refresh
        
        
        'Accion
        'añadimos en el log
        AnyadirLog rOrigen
        
        Usuario = DBLet(rOrigen!userid, "T")
        If Usuario <> "" Then
            'Usuario correcto con datos. Vemos su datos
            Cad = DBLet(rOrigen!authtype, "T")
            If Cad = "128" Then
                'Usuario autenticado
                 Usuario = Mid(Usuario, 1, 4) 'los 4 primeros sera la tarjeta
                 Usuario = mConfig.DigitoTrabajadores & Usuario
                
                'Localizo el trabajador

                 RT.MoveFirst
                 Seguir = True
                 While Seguir
                    If RT!Numtarjeta = Usuario Then
                        Seguir = False
                    Else
                        RT.MoveNext
                        Seguir = Not RT.EOF
                    End If
                Wend
                 
                 If RT.EOF Then
                    'NO se encuentra el trabajador
                    Cad = "No se encuentra el trabajador con tarjeta: " & Usuario
                    MsgBox Cad, vbExclamation
                    'Raise event
                   ' Err.Raise 516, Cad, Cad
                Else
                    'AQUI ESTA
                    Cad = "INSERT INTO entradafichajes(Secuencia, idTrabajador, Fecha ,Hora ,idInci ,HoraReal)"
                    Cad = Cad & " VALUES (" & Sec & "," & RT!idTrabajador & ",#" & Format(rOrigen!logtime, FormatoFecha) & "#,#"
                    Cad = Cad & Format(rOrigen!logtime, "hh:mm:ss") & "#,0,#"
                    Cad = Cad & Format(rOrigen!logtime, "hh:mm:ss") & "#)"
                    Conn.Execute Cad
                    Sec = Sec + 1
                End If
                
            End If
            
        End If
        'Dberiamos borrar
        rOrigen.Delete
        rOrigen.MoveNext
        
    
    Wend
    rOrigen.Close
    

    LeerDatosBD = True
    
    'Adelante
    lblIndicador.Caption = "Actualizando"
    lblIndicador.Refresh
    Cnn.CommitTrans
    Conn.CommitTrans
    
    
    
    
    'Repetidos
    If Sec <> SecuenciaIncial Then
        'Si ha metido alguno
        lblIndicador.Caption = "Repetidos"
        lblIndicador.Refresh
        Repetidos SecuenciaIncial
    Else
        lblIndicador.Caption = "No hay datos"
        lblIndicador.Refresh
        espera 1
    End If
ELee:
    If Err.Number <> 0 Then
        MuestraError Err.Number
        Cnn.RollbackTrans
        Conn.RollbackTrans
    
    End If
    Set Cnn = Nothing
End Function





Private Sub PreparaLogDatos()
Dim C As String

    On Error Resume Next
'    C = "drop table logHuella"
'    Conn.Execute C
    C = "select 0 as logindex, #2009-01-01 01:01:01# as logtime , """" as userid ,0 as authtype into logHuella from trabajadores where idtrabajador=-1"
    Conn.Execute C
    
    
    If Err.Number <> 0 Then Err.Clear 'YA existe
    
    
    
    
End Sub

Private Sub AnyadirLog(ByRef R As Recordset)
Dim C As String
    C = "INSERT INTO loghuella(logindex,logtime,userid,authtype) VALUES (" & R!logindex & ",#"
    C = C & Format(R!logtime, "yyyy-mm-dd hh:nn:ss") & "#,'" & R!userid & "'," & R!authtype & ")"
    Conn.Execute C
End Sub

Private Sub Repetidos(PrimeraInsercion As Long)
Dim KReg As Integer
Dim CadInci As String
Dim RFin As ADODB.Recordset
Dim Fecha, Hora
Dim ContFich As Long

    Cad = DevuelveDesdeBD("repeticion", "Empresas", "idEmpresa", "1", "N")
    KReg = Val(Cad)
    If KReg > 0 Then
        espera 1
        Me.Refresh
        'Obtenemos la fecha mas baja
        Set RFin = New ADODB.Recordset
     
            CadInci = "Select min(fecha) from EntradaFichajes WHERE Secuencia >= " & PrimeraInsercion
            RFin.Open CadInci, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            CadInci = "#1900/01/01#"
            If Not RFin.EOF Then
                If Not IsNull(RFin.Fields(0)) Then CadInci = "#" & Format(RFin.Fields(0), FormatoFecha) & "#"
            End If
            RFin.Close
            CadInci = " Fecha >= " & CadInci

        
        'Ya tenemos a partir de k fecha, y con k cadencia vamos a eliminar repetidos
        CadInci = "Select * from Entradafichajes WHERE " & CadInci
        CadInci = CadInci & " ORDER BY idTrabajador,Fecha,Hora"
        RFin.Open CadInci, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        PrimeraInsercion = 0 'Tendremos el codigo del trabajador
        CadInci = "DELETE from EntradaFichajes WHERE Secuencia = "
        While Not RFin.EOF
            If RFin!idTrabajador <> PrimeraInsercion Then
                lblIndicador.Caption = "Trabajador: " & RFin!idTrabajador
                lblIndicador.Refresh
                'Nuevo trabajador
                PrimeraInsercion = RFin!idTrabajador
                Fecha = RFin!Fecha
                Hora = RFin!Hora
            Else
                'Es el mismo trabajador.
                'Veamos la fecha
                If RFin!Fecha <> Fecha Then
                    Fecha = RFin!Fecha
                    Hora = RFin!Hora
                Else
                    'MISMO TRABAJADOR , MISMA FECHA
                    ContFich = DateDiff("n", Hora, RFin!Hora)
                    If ContFich > KReg Then
                        'Las horas se diferencian. NO elimino
                        Hora = RFin!Hora
                    Else
                        'SI elimino
                        Conn.Execute CadInci & RFin!Secuencia
                    End If
                End If
            End If
            'Siguiente
            RFin.MoveNext
        Wend
        RFin.Close
    
    End If  'Eliminacion marcajes repetidos


End Sub

Private Sub Text2_DblClick()
    Text2.Locked = False
End Sub
