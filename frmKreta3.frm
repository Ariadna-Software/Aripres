VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmKreta3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comunicador"
   ClientHeight    =   4215
   ClientLeft      =   105
   ClientTop       =   405
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5741
      _Version        =   393216
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "Operaciones"
      TabPicture(0)   =   "frmKreta3.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdMarcajes"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdHora"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Configurar terminales"
      TabPicture(1)   =   "frmKreta3.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdGrabar"
      Tab(1).Control(1)=   "chkConfig(3)"
      Tab(1).Control(2)=   "chkConfig(2)"
      Tab(1).Control(3)=   "chkConfig(1)"
      Tab(1).Control(4)=   "chkConfig(0)"
      Tab(1).Control(5)=   "Label1"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Grabar trabajador"
      TabPicture(2)   =   "frmKreta3.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdSeccion"
      Tab(2).Control(1)=   "chkSeccionBorrar"
      Tab(2).Control(2)=   "cboSeccion"
      Tab(2).Control(3)=   "Command1"
      Tab(2).Control(4)=   "Text5(1)"
      Tab(2).Control(5)=   "Text5(0)"
      Tab(2).Control(6)=   "Label2(1)"
      Tab(2).Control(7)=   "Line1"
      Tab(2).Control(8)=   "Image2(0)"
      Tab(2).Control(9)=   "Label2(0)"
      Tab(2).ControlCount=   10
      Begin VB.CommandButton cmdSeccion 
         Caption         =   "grabar seccion"
         Height          =   495
         Left            =   -69480
         TabIndex        =   19
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CheckBox chkSeccionBorrar 
         Caption         =   "Borrar todos los datos terminal"
         Height          =   195
         Left            =   -74760
         TabIndex        =   18
         Top             =   2760
         Width           =   2655
      End
      Begin VB.ComboBox cboSeccion 
         Height          =   315
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2160
         Width           =   3495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "grabar trabajador"
         Height          =   495
         Left            =   -69480
         TabIndex        =   15
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   0
         Left            =   -74760
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdHora 
         Caption         =   "Poner en hora"
         Height          =   495
         Left            =   4200
         TabIndex        =   11
         Top             =   1560
         Width           =   2895
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   495
         Left            =   -70080
         TabIndex        =   9
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox chkConfig 
         Caption         =   "Trabajadores"
         Height          =   255
         Index           =   3
         Left            =   -71880
         TabIndex        =   8
         Top             =   2160
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chkConfig 
         Caption         =   "Incidencias"
         Height          =   255
         Index           =   2
         Left            =   -74160
         TabIndex        =   7
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CheckBox chkConfig 
         Caption         =   "Mensajes"
         Height          =   255
         Index           =   1
         Left            =   -71880
         TabIndex        =   6
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CheckBox chkConfig 
         Caption         =   "Configuración base"
         Height          =   255
         Index           =   0
         Left            =   -74160
         TabIndex        =   5
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton cmdMarcajes 
         Caption         =   "Leer marcajes"
         Height          =   495
         Left            =   600
         TabIndex        =   4
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Sección"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   17
         Top             =   2160
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   -74760
         X2              =   -67800
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   -73920
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Trabajador"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   14
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Proceso de configuracion de los terminales:"
         Height          =   375
         Left            =   -74640
         TabIndex        =   10
         Top             =   600
         Width           =   5055
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   3600
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock tcpCliente 
      Index           =   0
      Left            =   120
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "192.168.123.10"
      RemotePort      =   1001
   End
   Begin MSWinsockLib.Winsock tcpCliente 
      Index           =   1
      Left            =   600
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "192.168.123.10"
      RemotePort      =   1001
   End
   Begin MSWinsockLib.Winsock tcpCliente 
      Index           =   2
      Left            =   1080
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "192.168.123.10"
      RemotePort      =   1001
   End
   Begin MSWinsockLib.Winsock tcpCliente 
      Index           =   3
      Left            =   1560
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "192.168.123.10"
      RemotePort      =   1001
   End
   Begin VB.CommandButton cmdProbar 
      Caption         =   "Pruebas"
      Height          =   615
      Left            =   4800
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSWinsockLib.Winsock tcpCliente 
      Index           =   4
      Left            =   2040
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "192.168.123.10"
      RemotePort      =   1001
   End
   Begin MSWinsockLib.Winsock tcpCliente 
      Index           =   5
      Left            =   2520
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "192.168.123.10"
      RemotePort      =   1001
   End
   Begin MSWinsockLib.Winsock tcpCliente 
      Index           =   6
      Left            =   3000
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "192.168.123.10"
      RemotePort      =   1001
   End
   Begin MSWinsockLib.Winsock tcpCliente 
      Index           =   7
      Left            =   360
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "192.168.123.10"
      RemotePort      =   1001
   End
   Begin MSWinsockLib.Winsock tcpCliente 
      Index           =   8
      Left            =   840
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "192.168.123.10"
      RemotePort      =   1001
   End
   Begin MSWinsockLib.Winsock tcpCliente 
      Index           =   9
      Left            =   1320
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "192.168.123.10"
      RemotePort      =   1001
   End
   Begin VB.Label lblInf 
      Alignment       =   2  'Center
      Caption         =   "Información de proceso..."
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   3600
      Width           =   4695
   End
End
Attribute VB_Name = "frmKreta3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents k2 As Kreta2
Attribute k2.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1

Private Conectado As Boolean
Private SeVe As Boolean

Dim Rs As ADODB.Recordset


Private Sub cmdGuardarMarcajes_Click()

    
    'CargarFichajesGeslab2 mConfig.DirMarcajes
    CargarFichajesGeslab2
    MsgBox "Los marcajes han sido guardados"
End Sub



Private Sub cmdGrabar_Click()
Dim T1 As Single

    Dim I As Integer
    For I = 0 To Me.chkConfig.Count - 1
        If Me.chkConfig(I).Value = 1 Then Exit For
    Next
    
    If I = Me.chkConfig.Count Then
        MsgBox "Seleccione alguna opcion de configuracion de los terminales", vbExclamation
        Exit Sub
    End If
        
    Screen.MousePointer = vbHourglass
    cmdGrabar.Enabled = False
    DoEvents
    
    lblInf.Caption = "Comienza proceso"
    lblInf.Refresh
    If Me.chkConfig(0).Value Then
        lblInf.Caption = "Comienza proceso"
        lblInf.Refresh
        CargarConfiguracion
        espera 0.5
    End If
    If Me.chkConfig(1).Value Then
        lblInf.Caption = "Mensajes"
        lblInf.Refresh
        CargarMensajes
        espera 0.5
    End If
    If Me.chkConfig(2).Value Then
        lblInf.Caption = "Incidencias"
        lblInf.Refresh
        CargarIncidencias
        espera 0.5
    End If
    
    Me.Refresh
    T1 = 0
    If Me.chkConfig(3).Value Then
        T1 = Timer
        lblInf.Caption = "Carga usuarios"
        lblInf.Refresh
        CargarUsuariosTodosTerminales2 -1, True
        espera 0.5
    End If
    T1 = Timer - T1
    If T1 < 5 And T1 > 0 Then espera T1
    
        
        
    lblInf.Caption = "Proceso finalizado"
    lblInf.Refresh
    espera 0.5
    lblInf.Caption = ""
    cmdGrabar.Enabled = True
    Screen.MousePointer = vbDefault
    
    
    
End Sub

Private Sub cmdHora_Click()
Dim I As Integer
    If ColK2 Is Nothing Then CargarTerminales
    For I = 1 To ColK2.Count
        Set k2 = ColK2(I)
        k2.GrabaHoraTerminal
    Next
End Sub

Private Sub cmdMarcajes_Click()
    
    
    If vEmpresa.QueEmpresa = 4 Then
        'Catadau
        lblInf.Caption = "Leyendo carpeta srv"
        lblInf.Refresh
        If Not AccedeFicherosServidor Then
            lblInf.Caption = ""
            Exit Sub
        End If
        

    End If
    
    
    
    
    Screen.MousePointer = vbHourglass
    
    
    LeerMarcajes vEmpresa.DirMarcajes
    
    
    
    'Procesar fichero huella, solo para alzira o catadau
    If vEmpresa.QueEmpresa = 2 Then CargarFichajesGeslab2
    
    
    'Enero 2015.   Proceso NOCTURNO
    If vEmpresa.HorarioNocturno2 Then
        Screen.MousePointer = vbHourglass
        ProcesoHorasAcabalgadas
        Me.lblInf.Caption = ""
    End If
    
    
    
    
    MsgBox "Proceso lectura finalizado", vbInformation


    Screen.MousePointer = vbDefault
    Unload Me  'me piro
End Sub


'Esto era para KATADAU
Private Function AccedeFicherosServidor() As Boolean
    On Error Resume Next
    MsgBox "FALTA"
''''    AccedeFicherosServidor = False
''''
''''    If MiEmpresa.QueEmpresa = 4 Then
''''        If MiEmpresa.pathCostesServer = "" Then
''''            MsgBox "No existe carpeta en el servidor (pathcosteserver)", vbExclamation
''''            Exit Function
''''        End If
''''    End If
''''
''''    If Dir(MiEmpresa.pathCostesServer & "\*.dbz") = "" Then
''''        'NADA
''''    End If
''''    If Err.Number <> 0 Then
''''        MsgBox "Error accediendo a: " & MiEmpresa.pathCostesServer, vbExclamation
''''    Else
''''        AccedeFicherosServidor = True
''''    End If
End Function


Private Sub cmdProbar2_Click()

    Dim I As Integer
    Dim usu As UsuarioHuella




    
    
    
    
     '-- Primero cargamos los terminales
    If ColK2 Is Nothing Then CargarTerminales
    '-- Ahora los usuarios
    
    

        
       
        
            Set usu = New UsuarioHuella
            If usu.Leer(3) Then
                lblInf.Caption = "Grabar usuario SIN"
                lblInf.Refresh
                '-- Ahora hay que cargarlo en todos los terminales
                For I = 1 To ColK2.Count
                    Set k2 = ColK2(I)
                    
                    'Primero borro el usuario(por si acaso existe)
                    k2.BorrarUsuario usu
                    espera 0.5
                    If Not usu.CargarEnTerminalSINHUELLA(k2) Then
                        
                    Else
                        lblInf.Caption = "Ok"
                        lblInf.Refresh
                        espera 0.8
                    End If
                    DoEvents
                    
                Next
            End If
 
            
                



End Sub



Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSeccion_Click()

    If Me.cboSeccion.ListIndex < 0 Then Exit Sub
    
    If MsgBox("Desea continuar con la seccion: " & cboSeccion.List(cboSeccion.ListIndex) & " ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    cmdSeccion.Enabled = False
    Me.Command1.Enabled = False
    
    lblInf.Caption = "Carga usuarios seccion " & cboSeccion.List(cboSeccion.ListIndex)
    lblInf.Refresh
    Me.Refresh
    CargarUsuariosTodosTerminales2 cboSeccion.ItemData(cboSeccion.ListIndex), chkSeccionBorrar.Value = 1
    espera 0.5
    cmdSeccion.Enabled = True
    Me.Command1.Enabled = True
    lblInf.Caption = ""
End Sub

Private Sub Command1_Click()

    If Text5(0).Text = "" Then Exit Sub
    Screen.MousePointer = vbHourglass
    Command1.Enabled = False
    Grabar1Trabajador
    Command1.Enabled = True
    lblInf.Caption = ""
    Screen.MousePointer = vbDefault
End Sub

Private Sub Grabar1Trabajador()
Dim usu As UsuarioHuella
Dim I As Integer
Dim B As Boolean

Dim SQL As String
Dim Rs As ADODB.Recordset


    
    
    
    
     '-- Primero cargamos los terminales
    If ColK2 Is Nothing Then CargarTerminales
    '-- Ahora los usuarios
    
    SQL = "select * from usuarios WHERE GesLabID = " & Text5(0).Text
    Set Rs = GesHuellaDB.cursor(SQL)
    If Rs.EOF Then
      MsgBox "No tiene ID huella asociado", vbExclamation
      
    Else

        
        SQL = ""
        
            Set usu = New UsuarioHuella
            If usu.Leer(Rs!CodUsuario) Then
                lblInf.Caption = "Grabar usuario " & Rs!CodUsuario
                lblInf.Refresh
                '-- Ahora hay que cargarlo en todos los terminales
                For I = 1 To ColK2.Count
                    Set k2 = ColK2(I)
                    
                    'Primero borro el usuario(por si acaso existe)
                    k2.BorrarUsuario usu
                    espera 0.5
                    
                    If usu.FIR = "" Then
                        'USUARIO SIN HUELLA
                        B = usu.CargarEnTerminalSINHUELLA(k2)
                    Else
                        B = usu.CargarEnTerminal(k2)
                    End If
                    If Not B Then
                        SQL = SQL & "Terminal: " & k2.Numero & "   " & usu.GesLabID & " - " & usu.Mensaje & vbCrLf
                    Else
                        lblInf.Caption = "Ok"
                        lblInf.Refresh
                        espera 0.8
                    End If
                    DoEvents
                    
                Next
            End If
 
            If SQL <> "" Then MsgBox SQL, vbExclamation
                
        
        
        
    End If
    Rs.Close
    Set Rs = Nothing
    
End Sub

Private Sub Form_Activate()
    SeVe = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    
    
    
    cmdProbar.Visible = False
    CargarTerminales
    lblInf.Caption = ""
    CargaSecciones
    Me.SSTab1.Tab = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SeVe = False
    Set ColK2 = Nothing
    CerrarPuertos
    espera 0.5
End Sub

Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
    Text5(0).Text = vCodigo
    Text5(1).Text = vCadena
End Sub

Private Sub k2_LanzaMensaje(Mensaje As String)
    If SeVe Then
        lblInf.Caption = Mensaje
        lblInf.Refresh
        'DoEvents
    End If
End Sub

Private Sub Image2_Click(Index As Integer)

    Set frmB = New frmBuscaGrid
        frmB.vTabla = "Trabajadores"
        frmB.vCampos = "NomTrabajador"
        'frmB.CampoCodigo = "IdTrabajador"
        'frmB.TipoDatos = 3
        frmB.vTitulo = "EMPLEADOS"
        frmB.Show vbModal
        Set frmB = Nothing
End Sub

Private Sub tcpCliente_Close(Index As Integer)
    'Stop
End Sub

Private Sub tcpCliente_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim Recibido As String
    tcpCliente(Index).GetData Recibido, vbString
    ColK2.Item(CStr(Index)).Recibido = Recibido
End Sub

Public Sub CargarConfiguracion()
    '-- Cargamos lo que toca
    Dim SQL As String
    Dim Rs As ADODB.Recordset
    Dim I As Integer
    If ColK2 Is Nothing Then CargarTerminales
    For I = 1 To ColK2.Count
        Set k2 = ColK2(I)
        k2.CargarConfiguracion
        k2.CargarHSPorDefecto
        k2.CargarDias
        k2.CargarMeses
    Next
End Sub

Public Sub CargarTerminales()
   
  
    '-- En la carga montamos todos os terminales posibles
    Set ColK2 = New ColKreta2
    Dim SQL As String

    Dim NumTerm As Integer
    SQL = " select * from terminales"
    Set Rs = GesHuellaDB.cursor(SQL)
    If Not Rs.EOF Then
        Rs.MoveFirst
        While Not Rs.EOF
            NumTerm = Rs!codterm
            lblInf.Caption = "Cargando terminal " & CStr(NumTerm)
            lblInf.Refresh
           ' If tcpCliente.LBound <= NumTerm Then
           ' Debug.Print tcpCliente(0).Index
                tcpCliente(NumTerm).Close
                tcpCliente(NumTerm).Protocol = sckTCPProtocol
                tcpCliente(NumTerm).RemoteHost = Rs!IP
                tcpCliente(NumTerm).RemotePort = 1001
          '  End If
            Set k2 = New Kreta2
            Set k2.Socket = tcpCliente(NumTerm)
            k2.Numero = NumTerm
            If Not k2.ComprobarConexion() Then
                MsgBox "No hay conexión con el terminal: " & k2.Numero & _
                        " IP:" & k2.Socket.RemoteHost, vbExclamation
            End If
            ColK2.Add k2.Socket, NumTerm, CStr(NumTerm)
            Rs.MoveNext
        Wend
    End If
    Set k2 = Nothing
    Set Rs = Nothing
End Sub

Public Function CargarUsuariosTodosTerminales2(Seccion As Integer, BorrarTodos As Boolean) As Boolean
    Dim usu As UsuarioHuella
    Dim I As Integer
    Dim Col2 As Collection
    Dim TraSeccion As String
    Dim SinHuella As Boolean
    Dim B As Boolean
    Dim Cuantos As Integer
    Dim J As Integer
    '-- Primero cargamos los terminales
    If ColK2 Is Nothing Then CargarTerminales
    '-- Ahora los usuarios
    Dim SQL As String
    Dim Rs As ADODB.Recordset
    SQL = "select * from usuarios"
    TraSeccion = ""
    If Seccion >= 0 Then
        'Veremos que trabadores son de esa seccion
        Set Rs = New ADODB.Recordset
        TraSeccion = "Select IdTrabajador from trabajadores WHERE seccion = " & CStr(Seccion)
        Rs.Open TraSeccion, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        TraSeccion = ""
        While Not Rs.EOF
            TraSeccion = TraSeccion & ", " & Rs!idTrabajador
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        If TraSeccion <> "" Then
            TraSeccion = Mid(TraSeccion, 2)
            TraSeccion = " WHERE GeslabID IN (" & TraSeccion & ")"
        End If
    End If
    If TraSeccion <> "" Then SQL = SQL & TraSeccion
    Set Rs = GesHuellaDB.cursor(SQL)
    
    If Not Rs.EOF Then
        '-- Primero borramos los usuarios de los diferentes terminales
        
        If BorrarTodos Then
            lblInf.Caption = "Borrar usuarios"
            lblInf.Refresh
        
            For I = 1 To ColK2.Count
                Set k2 = ColK2(I)
                k2.BorrarTodosLosUsuarios
            Next
            lblInf.Caption = "Fin borre"
            lblInf.Refresh
            DoEvents
            espera 1
        End If
        
        lblInf.Caption = "Leer registros"
        lblInf.Refresh
        Cuantos = 0
        Rs.MoveFirst
        While Not Rs.EOF
            Cuantos = Cuantos + 1
            Rs.MoveNext
        Wend
        Rs.MoveFirst
        
        
        Set Col2 = New Collection
        While Not Rs.EOF
            Set usu = New UsuarioHuella
            J = J + 1
            If usu.Leer(Rs!CodUsuario) Then
                lblInf.Caption = "Grabar usuario " & Rs!CodUsuario & "  (" & J & " / " & Cuantos & ")"
                lblInf.Refresh
                

                
                
                
                '-- Ahora hay que cargarlo en todos los terminales
                For I = 1 To ColK2.Count
                    Set k2 = ColK2(I)
                    
                    If Not BorrarTodos Then
                        k2.BorrarUsuario usu
                        espera 0.5
                    End If
                    
                    
                    lblInf.Caption = "Grabar usuario " & Rs!CodUsuario & "  (" & J & " / " & Cuantos & ")"
                    lblInf.Refresh
                    
                    If usu.FIR = "" Then
                        'USUARIO SIN HUELLA
                        B = usu.CargarEnTerminalSINHUELLA(k2)
                    Else
                        B = usu.CargarEnTerminal(k2)
                    End If
                    If Not B Then
                        Col2.Add "T: " & k2.Numero & "   " & usu.GesLabID & " - " & usu.Mensaje
                    End If
                    
                    DoEvents
                    espera 0.05
                Next
            End If
            Rs.MoveNext
        Wend
        
        
        If Not Col2 Is Nothing Then
            If Col2.Count > 0 Then
                'FALTA###
                'FALTA###
                'FALTA###
                SQL = "Error grabando: " & vbCrLf & vbCrLf
                For I = 1 To Col2.Count
                    SQL = SQL & vbCrLf & Col2.Item(I)
                Next
                'frmVarios2.Text1 = SQL
                'frmVarios2.Show vbModal
                
                
                
            End If
            Set Col2 = Nothing
        End If
        
    End If
    Rs.Close
End Function




Private Function CargarMensajes() As Boolean
    '-- Cargamos lo que toca
    Dim SQL As String
    'Dim rs As ADODB.Recordset
    Dim I As Integer
    If ColK2 Is Nothing Then CargarTerminales
    For I = 1 To ColK2.Count
        Set k2 = ColK2(I)
        k2.CargarMensajes
    Next
    CargarMensajes = True
End Function

Private Function CargarIncidencias() As Boolean
    '-- Cargamos lo que toca
    Dim SQL As String
    'Dim rs As ADODB.Recordset
    Dim I As Integer
    If ColK2 Is Nothing Then CargarTerminales
    For I = 1 To ColK2.Count
        Set k2 = ColK2(I)
        k2.BorrarTodasLasIncidencias
        k2.CargarIncidencias
    Next
    CargarIncidencias = True
End Function

Public Function LeerMarcajes(Directorio As String) As Boolean
    '-- Cargamos lo que toca
    Dim SQL As String
    'Dim rs As ADODB.Recordset
    Dim I As Integer
    
    lblInf.Caption = "Inicio proceso lectura"
    lblInf.Refresh
    
    
    If ColK2 Is Nothing Then CargarTerminales
    
    
    Me.SSTab1.Enabled = False
    Me.cmdSalir.Enabled = False
    lblInf.Tag = Val(Timer)
    
    For I = 1 To ColK2.Count
        
        Set k2 = ColK2(I)
        lblInf.Caption = "lectura reloj: " & k2.Numero
        lblInf.Refresh
        k2.LeerMarcajes Directorio, I = 1, lblInf
    Next
    LeerMarcajes = True
    lblInf.Caption = ""
    Me.SSTab1.Enabled = True
    Me.cmdSalir.Enabled = True
    
End Function

Public Function CargarFichajesGeslab2() As Boolean
    '-- CargarFichajesGeslab:
    '   Se encarga de mirar en el directorio indicado si hay ficheros de fichajes
    '   y los actualiza en GesLab
    Dim Fichero As String
    Dim Leido As String
    Dim NF As Integer
'    Dim db As BaseDatos
    
    Dim tam As Long
    Dim llev As Long
    Dim Nodo As Byte  'Para catadu
    Dim Cole As Collection
    Dim jj As Integer
 
    
    'NO ABRIMOS LA BD
    lblInf.Caption = "Preparando datos"
    lblInf.Refresh
    
    If Dir(vEmpresa.DirProcesados, vbDirectory) = "" Then MkDir vEmpresa.DirProcesados
    
    '   Set db = New BaseDatos
    'db.AbrirConexionDavid conn.ConnectionString
    'db.Tipo = "ACCESS"
    'db.abrir_MYSQL
    
    Fichero = Dir(vEmpresa.DirMarcajes & "\HU*")
    Set Cole = New Collection
    Do While Fichero <> ""
        Cole.Add Fichero
        Fichero = Dir
    Loop
        
    For jj = 1 To Cole.Count
        DoEvents
        Me.Refresh
        Screen.MousePointer = vbHourglass
        
        NF = FreeFile
        Fichero = Cole.Item(jj)
        tam = FileLen(vEmpresa.DirMarcajes & "\" & Fichero)
        
        lblInf.Caption = "Fichero"
        lblInf.Refresh
        
        If vEmpresa.QueEmpresa = 4 Then
            lblInf.Caption = "Fichero"
            lblInf.Refresh
            
            'Copiamos al SERVIDOR EL FICHERO
            'FileCopy Directorio & "\" & Fichero, vempresa.c .pathCostesServer & "\" & Fichero
            llev = InStr(1, Fichero, ".")
            
            If llev = 0 Then
                Nodo = 10
            Else
                Leido = Mid(Fichero, llev - 2, 2) 'los dos ultimos antes del punto
                Nodo = CByte(Val(Leido))
            End If
            
        End If
        llev = 0
    
        Open vEmpresa.DirMarcajes & "\" & Fichero For Input As #NF
        Do While Not EOF(NF)
            Line Input #1, Leido
            llev = llev + Len(Leido)
            lblInf.Caption = Fichero & "  " & llev & " de " & tam
            lblInf.Refresh
            
            If vEmpresa.QueEmpresa = 2 Then
            
                'ALZIRA
                GrabaFichajeGesLabALZIRA Leido
            Else
                'CATADU
                'GrabaFichajeGesLabCATADAU Leido, db, Nodo
               
            End If
        Loop
        Close #NF
        lblInf.Caption = "Mover a procesados"
        lblInf.Refresh
    
        'FALTA###
        If CopiarEnProcesados(Fichero) Then Kill vEmpresa.DirMarcajes & "\" & Fichero
        
    Next
    Set Cole = Nothing
    
    
    'Ver entradas repetidas
    EntradasRepetidas
    
    
   ' Set db = Nothing
    lblInf.Caption = ""
    lblInf.Refresh
End Function


Private Sub EntradasRepetidas()
Dim RFin As ADODB.Recordset
Dim idTrabajador As Long
Dim CadInci As String
Dim Fecha As Date
Dim Hora As Date
Dim Diferencia As Long

    If vEmpresa.Repeticion <= 0 Then Exit Sub
        
    lblInf.Caption = "Entradas duplicadas"
    lblInf.Refresh
    Set RFin = New ADODB.Recordset
        

        
        'Ya tenemos a partir de k fecha, y con k cadencia vamos a eliminar repetidos
        CadInci = "Select * from Entradafichajes WHERE hora <='23:59:59'"
        CadInci = CadInci & " ORDER BY idTrabajador,Fecha,Hora"
        RFin.Open CadInci, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        idTrabajador = 0 'Tendremos el codigo del trabajador
        CadInci = "DELETE from EntradaFichajes WHERE Secuencia = "
        While Not RFin.EOF
           
        
            If RFin!idTrabajador <> idTrabajador Then
                
                lblInf.Caption = "Trabajador: " & RFin!idTrabajador
                lblInf.Refresh
                
                'Nuevo trabajador
                idTrabajador = RFin!idTrabajador
                Fecha = RFin!Fecha
                Hora = Format(RFin!Hora, "hh:mm:ss")
            Else
                'Es el mismo trabajador.
                'Veamos la fecha
                If RFin!Fecha <> Fecha Then
                    Fecha = RFin!Fecha
                    Hora = Format(RFin!Hora, "hh:mm:ss")
                Else
                    'MISMO TRABAJADOR , MISMA FECHA
                    Diferencia = DateDiff("n", Hora, Format(RFin!Hora, "hh:mm:ss"))
                    If Diferencia >= vEmpresa.Repeticion Then
                        'Las horas se diferencian. NO elimino
                        Hora = Format(RFin!Hora, "hh:mm:ss")
                    Else
                        'SI elimino
                        conn.Execute CadInci & RFin!Secuencia
                    End If
                End If
            End If
            'Siguiente
            RFin.MoveNext
        Wend
        RFin.Close
    
    
    Set RFin = Nothing




End Sub

Private Function CopiarEnProcesados(ByVal Fichero As String) As Boolean
Dim Cad As String

    
    On Error GoTo eCopiarEnProcesados
    CopiarEnProcesados = False
    
    If Len(Fichero) >= 9 Then
        Cad = Mid(Fichero, 3, 6)
        If Dir(vEmpresa.DirProcesados & "\" & Cad, vbDirectory) = "" Then
            If Not CrearSubCarpeta(Cad) Then Cad = ""
        End If
        If Cad <> "" Then Cad = Cad & "\"
    Else
        Cad = ""
    End If
    
    
    FileCopy vEmpresa.DirMarcajes & "\" & Fichero, vEmpresa.DirProcesados & "\" & Cad & Fichero
    
    
    CopiarEnProcesados = True
    Exit Function
eCopiarEnProcesados:
    MsgBox "Avise soporte tecnico" & vbCrLf & "Error moviendo a procesados" & vbCrLf & Fichero & vbCrLf & Err.Description, vbExclamation
    
End Function


Private Function CrearSubCarpeta(Carpeta As String) As Boolean
    On Error Resume Next
    CrearSubCarpeta = False
    MkDir vEmpresa.DirProcesados & "\" & Carpeta
    If Err.Number <> 0 Then
        MuestraError Err.Number, Carpeta, Err.Description
    Else
        CrearSubCarpeta = True
    End If
End Function

Private Sub PonerEmpleadoVacio()
            Text5(0).Text = ""
            Text5(1).Text = ""
'            Text2(0).Text = ""
'            Text2(0).Tag = ""
End Sub
Private Sub PonerEmpleado(Cod As String, Campo As String)
Dim RT As ADODB.Recordset
Dim SQL As String
    
    SQL = "Select * from Trabajadores where "
    SQL = SQL & Campo & " = " & Cod
    Set RT = New ADODB.Recordset
    RT.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RT.EOF Then
        'ponerempleadovacio
        PonerEmpleadoVacio
    Else
        'Ponemos los datos del empleado
        If IsNull(RT!Numtarjeta) Then
            MsgBox "No tiene codigo HUELLA asociado", vbExclamation
            PonerEmpleadoVacio
        Else
            Text5(0).Text = RT!idTrabajador
            Text5(1).Text = RT!nomtrabajador
            
        End If
    End If
    RT.Close
    Set RT = Nothing
End Sub

Private Sub Text5_LostFocus(Index As Integer)
    If Index = 1 Then Exit Sub
    Text5(Index).Text = Trim(Text5(Index).Text)
    If Text5(Index).Text <> "" Then
        If Not IsNumeric(Text5(Index).Text) Then
            MsgBox "Codigo incorrecto: " & Text5(Index).Text, vbExclamation
            Text5(Index).Text = ""
        End If
    End If
    If Text5(Index).Text = "" Then
        PonerEmpleadoVacio
    Else
        If Index = 0 Then
            PonerEmpleado Text5(Index).Text, "idTrabajador"
        Else
            PonerEmpleado "'" & Text5(Index).Text & "'", "NumTarjeta"
        End If
    End If

End Sub


Private Sub CargaSecciones()
    CargaComboSecciones Me.cboSeccion, False

End Sub


Private Sub CerrarPuertos()
Dim J As Byte
    On Error Resume Next
        For J = 0 To tcpCliente.Count - 1
            tcpCliente(J).Close
            If Err.Number <> 0 Then
                MsgBox "Cerrando puerto tcpIP terminal " & J, vbCritical
                Err.Clear
            End If
        Next
        Err.Clear
End Sub





'***************************************************************************************************************
'***************************************************************************************************************
'***************************************************************************************************************
'
'
'   Proceso nocturno. Horas acabalagadas
'
'
'***************************************************************************************************************
'***************************************************************************************************************
'***************************************************************************************************************
'  Ira dia a dia, desde el ultimo dia procesado hasta el siguiente que pueda procesar
'  Vera si en el dia N hay algun marcaje superior a las HoraParametros
'  Si. Ese marcae puede ser el incial de los que ficharan al dia seguiente
'   o puede ser el ultimo del dia actual
'    ¿Como lo sabremos?  Pq no habra ningun fichaje entre [ HoraParametros -6   y  horaparametros]
'    es decir, por ejemplo, Dia 1. si Horaparametros son las 22:30, si no tiene ningun marcaje entre
'    las 16:30 y las 22:30 significara que ESE es el primero de los que vendran al dia siguiente
'    Es decir, cuando procesemos el dia 2, los marcajes anteriores a HoraParametros -6  seran del mismo dia que el 1
'    Si diaPrimeraFichada= dia trabajado
'       -updaearemos los del dia siguiente a hh:mm + 24:00 y dia=dia -1
'    Si no
'       -updaearemos el primero a hh:mm  - 24:00    y dia=dia +1
'
Private Sub ProcesoHorasAcabalgadas()
Dim primeraFechaProcesar As Date
Dim Cad As String
Dim RegistrosTratar As Collection
Dim FE As Date
Dim Hora As Date
Dim H1 As Date
Dim I As Long

Dim QueDia As Integer
Dim DiaTraba As Collection

Dim UltimaDiaHoraTraidoMaquina2 As String
Dim UltimoDiaProcesado As Date
Dim FechaParaActualizarEnParametros As Date


Dim DiasATratar As Collection
Dim J As Integer

    On Error GoTo eProcesoHorasAcabalgadas
    
    lblInf.Caption = "horas acabalg."
    lblInf.Refresh
    
    
    
    
    'Ver si hay que entrar en el proceso
    
    
    
    
    Set Rs = New ADODB.Recordset
    
    Cad = DevuelveDesdeBD("AcabalUltimoDiaProcesado", "empresas", "1", "1")
    If Cad = "" Then
        'No esta grababado todavia. voy a ver la primera fecha de entradafichajes
        'select min(fecha) from entradafichajes
        Cad = DevuelveDesdeBD("min(fecha)", "entradafichajes", "1", "1")
        If Cad = "" Then Cad = "02/01/1900"
        Cad = DateAdd("d", -1, CDate(Cad))  'Para que el primer dia trabajado se el primero de entradafichakes
    End If
    primeraFechaProcesar = CDate(Cad)
    UltimoDiaProcesado = primeraFechaProcesar
    
    'Voy a ver ultimo dia -hora que hemos traido desde la maquina
    Cad = "Select fecha , concat(horareal,'') h1 from entradafichajes ORDER BY 1 desc,2 desc"
    Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
        Err.Raise 513, , "Entrada fichejes vacia"
    Else
        Cad = "23:59:59"
        If Not HoraFueraIntervalo(Rs!H1) Then Cad = Format(Rs!H1, "hh:nn:ss")
        Cad = Format(Rs!Fecha, "dd/mm/yyyy") & " " & Cad
    End If
    Rs.Close
    UltimaDiaHoraTraidoMaquina2 = Cad
    
    
    
    'Vamos a ver los dias a tratar
    Set DiasATratar = New Collection
    
    I = Round(vEmpresa.MaxRetraso * 60, 0)
    Cad = DateAdd("n", -I, vEmpresa.AcabalgadoHora)   'para ver si puedo procesar el dia
    If CDate(Format(UltimaDiaHoraTraidoMaquina2, "hh:nn:ss")) >= CDate(Cad) Then
        Cad = ""
    Else
        Cad = " AND fecha <" & DBSet(UltimaDiaHoraTraidoMaquina2, "F")
    End If
    
    
    
    Cad = "Select distinct fecha from entradafichajes where fecha> " & DBSet(primeraFechaProcesar, "F") & Cad
    Cad = Cad & "  order by 1"
    
 
    Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        DiasATratar.Add CStr(Format(Rs.Fields(0), "dd/mm/yyyy"))
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    If DiasATratar.Count = 0 Then GoTo eProcesoHorasAcabalgadas
    
    
    'Vamos a estudiar cada dia
    For QueDia = 1 To DiasATratar.Count
            'Voy a ver que dias tienen fichajes superiror a las 22:30 (parametros)
            ' Y luego estudiare esos dias
            I = Round(vEmpresa.MaxRetraso * 60, 0)
            Cad = DateAdd("n", -I, vEmpresa.AcabalgadoHora)   'para ver si puedo procesar el dia
            Hora = Cad
            Cad = "fecha = " & DBSet(DiasATratar.Item(QueDia), "F") & " AND hora > " & DBSet(Hora, "H") & " and hora <= '23:59:59'"
            'Select  from entradafichajes where fecha> '2001-01-10' AND hora > '22:00:00' and hora <= '23:59:59' ORDER BY fecha,idtrabajador
            Cad = "Select distinct idtrabajador from entradafichajes where " & Cad & " ORDER BY idtrabajador"
            Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Set RegistrosTratar = New Collection
            I = -1
            While Not Rs.EOF
                Cad = Rs!idTrabajador
                RegistrosTratar.Add Cad
                Rs.MoveNext
            Wend
            Rs.Close
            
            'Para que haga los nothing
            If RegistrosTratar.Count > 0 Then
                    lblInf.Caption = "Dia " & DiasATratar.Item(QueDia)
                    lblInf.Refresh
                    
                    'Procesar el dia
                    ' Para ello vere fichajes mas alla de las 22:30 del dia D
                    ' vere si el trabajador para ese dia NO tiene ninguna entrada entre las 22:30 y las 14:30, es decir,
                    ' si realmente la 22:30 es la entrada del dia
                    '
                    ' SACARE UN FRM previo, donde podra decir si el trabajador ese dia es acabalglado  o no
                    
                    '
                    conn.Execute "Delete from tmpnotrabajo"
                    espera 0.5
                    Cad = ""
                    For I = 1 To RegistrosTratar.Count
                        Cad = Cad & ", (" & RegistrosTratar(I) & ")"
                    Next
                    Cad = Mid(Cad, 2)
                    Cad = "INSERT INTO tmpnotrabajo(idTra) VALUES " & Cad
                    conn.Execute Cad
                    espera 0.5
                    
                    CadenaDesdeOtroForm = ""
                    frmAcabalgados.Fecha = CDate(DiasATratar.Item(QueDia))
                    frmAcabalgados.Show vbModal
                    
                    'Ha pulsado aceptar
                    If CadenaDesdeOtroForm = "" Then
                        'Cancelado proceso
                         GoTo eProcesoHorasAcabalgadas
                    Else
                        lblInf.Caption = "Actualizando BD..."
                        Me.Refresh
                        DoEvents
                        Screen.MousePointer = vbHourglass
                    
                        Cad = "Select * from tmpnotrabajo ORDER by idtra"
                        Set DiaTraba = New Collection
                        
                        Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        While Not Rs.EOF
                            DiaTraba.Add CStr(Rs!IdTra)
                            Rs.MoveNext
                        Wend
                        Rs.Close
                        
                        
                        'Para cada trabajador
                        For J = 1 To DiaTraba.Count
                        
                         
                            Debug.Print DiaTraba(J)
                        
                            
                            lblInf.Caption = "Trabajador: " & DiaTraba(J)
                            lblInf.Refresh
                        
                            
                            'Abrimos un frm para que nos diga si son acabalgados o normales y se l
                            
                            
                        
                            
                            FE = CDate(DiasATratar.Item(QueDia))
                            
                            If vEmpresa.AcabalgadoDiaInicio Then
                                'La primera fichada marca el dia de inicio
                                '       -updaearemos los del dia siguiente a hh:mm + 24:00 y dia=dia -1
                                Cad = "fecha = " & DBSet(FE, "F")
                                Cad = Cad & ",hora = ADDTIME(hora , '24:00:00' ) "
                                Cad = Cad & ",horareal = ADDTIME(horareal , '24:00:00' ) "
                                Cad = Cad & " WHERE fecha = " & DBSet(DateAdd("d", 1, FE), "F")
                                Cad = Cad & " AND hora < " & DBSet(Hora, "H")
                            Else
                                'La fichada es la primera del dia siguiente
                                '       -updaearemos el primero a hh:mm  - 24:00    y dia=dia +1
                                Cad = "fecha = " & DBSet(DateAdd("d", 1, FE), "F")
                                Cad = Cad & ",hora = ADDTIME(hora , '-24:00:00' ) "
                                Cad = Cad & ",horareal = ADDTIME(horareal , '-24:00:00' ) "
                                Cad = Cad & " WHERE fecha = " & DBSet(FE, "F")
                                Cad = Cad & " AND hora >= " & DBSet(Hora, "H")
                                                  
                            End If
                            Cad = Cad & " AND idtrabajador = " & DiaTraba(J)
                            Cad = "UPDATE entradafichajes set " & Cad
                            conn.Execute Cad
                            espera 0.1
                                
                        Next J
                                
                        Cad = "UPDATE empresas set AcabalUltimoDiaProcesado = " & DBSet(FE, "F")
                        conn.Execute Cad
                        lblInf.Caption = "Actualizando..."
                        lblInf.Refresh
                        espera 1.5
                        
                    End If
            Else
                'Hemos pulsado tratar dia pero no hay trabajadores para mirar
                Cad = "UPDATE empresas set AcabalUltimoDiaProcesado = " & DBSet(DiasATratar.Item(QueDia), "F")
                conn.Execute Cad
            End If
            
            Set RegistrosTratar = Nothing
    Next QueDia
eProcesoHorasAcabalgadas:
    lblInf.Caption = "Proceso acab. finalizado"
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set Rs = Nothing
    Set RegistrosTratar = Nothing
    Set DiaTraba = Nothing
    Set DiasATratar = Nothing
    
End Sub

    
