VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRelojBiostar2 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Suprema BIOSTAR 2"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameError 
      Height          =   6855
      Left            =   6240
      TabIndex        =   5
      Top             =   120
      Width           =   7695
      Begin VB.CommandButton Command3 
         Caption         =   "&Copiar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         ToolTipText     =   "Copiar en portapapeles"
         Top             =   6360
         Width           =   1170
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   8
         Top             =   6360
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5535
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   9763
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   9596
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Trabajadores en el reloj sin asignar en ARipres"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   4695
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4680
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   6120
      Width           =   1170
   End
   Begin VB.CommandButton cmdLeer 
      Caption         =   "&Leer"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   6120
      Width           =   1170
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ultima lectura realizada"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   3
      Top             =   960
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   7500
      Left            =   -1560
      Picture         =   "frmRelojBiostar2.frx":0000
      Top             =   -360
      Width           =   7500
   End
End
Attribute VB_Name = "frmRelojBiostar2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim Cn As Connection
Dim cad As String


Private Sub cmdLeer_Click()
    Screen.MousePointer = vbHourglass
    ListView1.ListItems.Clear
    If LeerDatos Then
        Label11.Caption = "Lectura OK"
        Me.Command1.Caption = "Cerrar"
    Else
        Label11.Caption = ""
    End If
    
    Set miRsAux = Nothing
    If ListView1.ListItems.Count > 0 Then
        FrameError.Visible = True
        cmdLeer.Enabled = False
    
    End If
    Screen.MousePointer = vbDefault
    
End Sub

'tmpcombinada(IdTrabajador,Fecha,H1,codusu,idinci)
Private Function LeerDatos() As Boolean
Dim ColTablas As Collection
Dim i As Integer
Dim Inicial  As Long
Dim MaxFecha As Date
Dim B As Boolean

    Label11.Caption = "Buscando tablas"
    Label11.Refresh
    
    'Borramos de la temporal en aripres
    conn.Execute "DELETE FROM tmppresencia where codusu =" & vUsu.Codigo
    
    cad = Format(Text2.Text, "yyyymm")
    Inicial = Val(cad)
    Set ColTablas = New Collection
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "show tables", Cn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cad = miRsAux.Fields(0)
        
        If LCase(Mid(cad, 1, 5)) = "t_lg2" Then   'Todas son del año 2000
            'TABLA de marcaje. Veamos si esta donde nos quedamos
           
            cad = Mid(cad, 5)
            If Val(cad) >= Inicial Then
                'ZIIIIIII
                'Hay que tratarla
                ColTablas.Add CStr(miRsAux.Fields(0))
        
            End If
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    MaxFecha = Text2.Text
    Label11.Tag = DBSet(MaxFecha, "FH")
    
    'Ya tenemos las tablas
    
    For i = 1 To ColTablas.Count
        Label11.Caption = "Leyendo " & ColTablas.Item(i) & "  (" & i & "/" & ColTablas.Count & ")"
        Label11.Refresh
                                                                    'huell pin
        cad = "SELECT * FROM " & ColTablas.Item(i) & " WHERE evt in (4865,4097)"
        cad = cad & " AND srvdt >" & Label11.Tag
        cad = cad & " order by srvdt"
        miRsAux.Open cad, Cn, adOpenKeyset, adLockPessimistic, adCmdText
        cad = ""
        If Not miRsAux.EOF Then
            
            
            While Not miRsAux.EOF
                NumRegElim = NumRegElim + 1
                '                       tmppresencia(Id,idtra,Fecha,H1,codusu,Incidencias)
                cad = cad & ", (" & NumRegElim & "," & miRsAux!usrid & "," & DBSet(miRsAux!srvdt, "F") & ",'" & Format(miRsAux!srvdt, "hh:nn:ss")
                cad = cad & "'," & vUsu.Codigo & ",0)"
                
                miRsAux.MoveNext
            Wend
            miRsAux.MovePrevious
            If miRsAux!srvdt > MaxFecha Then MaxFecha = miRsAux!srvdt
        End If
        miRsAux.Close
        
        'tmpcombinada(IdTrabajador,Fecha,H1,codusu,idinci)
        If cad <> "" Then
            cad = Mid(cad, 2)
            cad = "INSERT INTO tmppresencia(Id,idtra,Fecha,H1,codusu,Incidencias) VALUES " & cad
            conn.Execute cad
        End If
    Next i
    Set ColTablas = Nothing
    
    
    
    
    
    'Veamos si todos los trabajadores existen
    
    Label11.Caption = "Comprobando trabajadores"
    Label11.Refresh
    espera 0.5
    cad = "Select distinct idtra from tmppresencia where codusu =" & vUsu.Codigo & " ORDER BY 1"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not miRsAux.EOF
        Label11.Caption = "Biostar ID" & miRsAux!idTRa
        Label11.Refresh
        cad = DevuelveDesdeBD("idtrabajador", "trabajadores", "idTraReloj2", CStr(miRsAux!idTRa), "T")
        If cad = "" Then
            cad = DevuelveNombreBiostar(CStr(miRsAux!idTRa))
            NumRegElim = NumRegElim + 1
            ListView1.ListItems.Add , , CStr(miRsAux!idTRa)
            ListView1.ListItems(NumRegElim).SubItems(1) = cad
        Else
            'Para no volver a leer , pondremos el trabajador de Aripres en el campo seccion
            cad = "UPDATE tmppresencia set seccion=" & cad & " WHERE idtra=" & miRsAux!idTRa
            conn.Execute cad
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    If ListView1.ListItems.Count = 0 Then
        
        Label11.Caption = "Insertando datos "
        Label11.Refresh
            
        
        cad = DevuelveDesdeBD("max(secuencia)", "entradafichajes", "1", "1")
        cad = Val(cad) + 1
        
        'Estan todos los trabajadores en la BD
        cad = " tmppresencia ,(SELECT @rownum:=" & cad & " ) r where codusu = " & vUsu.Codigo
        
        '                                                                            1:  REloj secundario
        cad = "SELECT @rownum:=@rownum+1 AS rownum ,seccion,fecha ,h1,incidencias,h1,1 FROM " & cad & " ORDER BY fecha,h1"
        
        cad = "INSERT INTO entradafichajes(Secuencia,idTrabajador,Fecha,Hora,idInci,HoraReal,reloj) " & cad
        
        conn.Execute cad
    
    
        
        EntradasRepetidasProceso Me.Label11
    
        'Nov 2018  --> Esta dentro la caribale para que lo ejecute o no If Not vEmpresa.AcabaJornadaDiaSiguiente
        espera 0.25
        HorasNocturnas Me.Label11
        
        
        cad = "UPDATE  biostar2 set ultimaFechaLeida =" & DBSet(MaxFecha, "FH")
        conn.Execute cad
        
        Text2.Text = Format(MaxFecha, "dd/mm/yyyy hh:nn:ss")
        LeerDatos = True
        
        
        Label11.Caption = "Lectura OK"
        
        
        cmdLeer.Enabled = False
        
    End If
    
End Function




Private Sub Command1_Click()
    If Not Cn Is Nothing Then Cn.Close
    Set Cn = Nothing
    Unload Me
   
End Sub

Private Function DevuelveNombreBiostar(Id As String) As String
Dim R As ADODB.Recordset

    On Error GoTo eDevuelveNombreBiostar

    Set R = New ADODB.Recordset
    R.Open "select nm from t_usr where USRID =" & Id, Cn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If R.EOF Then
        DevuelveNombreBiostar = "No encontrado"
    Else
        If IsNull(R!nm) Then
            DevuelveNombreBiostar = "Vacio. "
        Else
            DevuelveNombreBiostar = R!nm
        End If
    End If
    R.Close
    
    
eDevuelveNombreBiostar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
        DevuelveNombreBiostar = "ERROR"
    End If
    Set R = Nothing
    
End Function


Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    
    cad = ""
    For NumRegElim = 1 To ListView1.ListItems.Count
        cad = cad & ListView1.ListItems(NumRegElim).Text & Chr(9) & ListView1.ListItems(NumRegElim).SubItems(1) & vbCrLf
    Next NumRegElim
    Clipboard.SetText cad
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
        
End Sub

Private Sub Form_Activate()
    If cmdLeer.Tag = 1 Then
        cmdLeer.Tag = 0
        DoEvents
        Screen.MousePointer = vbHourglass
        If AbrirConexion Then
            'Vamos a veer cual fue la ultima fecha
            If UltimaLecturaBD Then cmdLeer.Enabled = True
        End If
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    Text2.Text = "Leyendo"
    cmdLeer.Enabled = False
    cmdLeer.Tag = 1
    FrameError.Visible = False
    FrameError.Top = 30
    FrameError.Left = 60
End Sub


Private Function AbrirConexion() As Boolean
    
    On Error GoTo EAbrirConexion
    AbrirConexion = False
    
    cad = DevuelveDesdeBD("configreloj2", "empresas", "1", "1")
    If cad = "" Then
        MsgBox "Mal configurado BIOSTAR2", vbExclamation
        Exit Function
    End If
    Set Cn = New Connection
    Cn.CursorLocation = adUseServer

    
    
    'Este es el que hay que dejar
    cad = "DSN=" & cad & ""   'Aripres4;DESC=MySQL ODBC 3.51 Driver DSN;;;;OPTION=3;STMT=;
    '    Cad = Cad & ";Persist Security Info=true"
    Cn.Open cad
    
    AbrirConexion = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, Err.Description
    Text2.Text = "ERROR"
    Set Cn = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set Cn = Nothing
End Sub


Private Function UltimaLecturaBD() As Boolean
    On Error GoTo eUltimaLecturaBD
    
    UltimaLecturaBD = False
    
    cad = "Select ultimaFechaLeida FROM biostar2"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        Me.Text2.Text = "01/01/2000 00:00:00"
    Else
        Me.Text2.Text = Format(miRsAux.Fields(0), "dd/mm/yyyy hh:mm:ss")
    End If
    miRsAux.Close
    UltimaLecturaBD = True
eUltimaLecturaBD:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
End Function






