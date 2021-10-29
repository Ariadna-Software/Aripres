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
   Begin VB.CommandButton cmdCambiar 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4200
      Picture         =   "frmRelojBiostar2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Cambiar ultima fecha leida"
      Top             =   1200
      Width           =   375
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
      Picture         =   "frmRelojBiostar2.frx":0A02
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
Dim Cad As String


Private Sub cmdCambiar_Click()
Dim C1 As String
    
    If Text2.Text = "ERROR" Then Exit Sub
    Cad = ""
    
    C1 = InputBox("Ultima fecha leida", "Biostar", Format(Text2.Text, "dd/mm/yyyy"))
    If C1 = "" Then C1 = "N"
    If Not IsDate(C1) Then Exit Sub
    Cad = Format(C1, "dd/mm/yyyy")
    
    C1 = InputBox("Ultima HORA leida", "Biostar", Format(Text2.Text, "hh:mm:ss"))
    If C1 = "" Then C1 = "N"
    If Not IsDate(C1) Then Exit Sub
    If InStr(1, C1, ":") = 0 Then Exit Sub
    
    
    Cad = Cad & " " & C1
    If Not IsDate(Cad) Then
        MsgBox "Fecha incorrecta " & Cad, vbExclamation
        Exit Sub
    End If
    
    C1 = DevuelveDesdeBD("max(fecha)", "marcajes", "1", "1")
    If C1 = "" Then C1 = "01/01/2001"
    
    If CDate(C1) >= CDate(Format(Cad, "dd/mm/yyyy")) Then
        MsgBox "Fecha procesada", vbExclamation
        If vUsu.Codigo > 0 Then Exit Sub
    End If
    
    C1 = "Desea guardar como ultima fecha leida : " & Cad
    If MsgBox(C1, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    C1 = "UPDATE biostar2 SET ultimaFechaLeida = " & DBSet(Cad, "FH")
    If EjecutaSQL(C1) Then
        Text2.Text = Cad
    End If
    
End Sub

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
Dim I As Integer
Dim Inicial  As Long
Dim MaxFecha As Date
Dim B As Boolean
Dim NF As Integer

    Label11.Caption = "Buscando tablas"
    Label11.Refresh
    
    'Borramos de la temporal en aripres
    conn.Execute "DELETE FROM tmppresencia where codusu =" & vUsu.Codigo
    
    Cad = Format(Text2.Text, "yyyymm")
    Inicial = Val(Cad)
    Set ColTablas = New Collection
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "show tables", Cn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = miRsAux.Fields(0)
        
        If LCase(Mid(Cad, 1, 5)) = "t_lg2" Then   'Todas son del año 2000
            'TABLA de marcaje. Veamos si esta donde nos quedamos
           
            Cad = Mid(Cad, 5)
            If Val(Cad) >= Inicial Then
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
    
    'Octubre 2020
    'Añadimos identificador termianal.  Hay una tabla
    
    
    For I = 1 To ColTablas.Count
        Label11.Caption = "Leyendo " & ColTablas.Item(I) & "  (" & I & "/" & ColTablas.Count & ")"
        Label11.Refresh
                                                                    'huell cara pin   tarjeta (20/05/20(
        Cad = "SELECT * ,from_unixtime(devdt, '%Y-%m-%d %H:%i:%s') horaRealUnix "
        Cad = Cad & " FROM " & ColTablas.Item(I) & " WHERE evt in (4865,4867,4097,4102)"
        Cad = Cad & " AND from_unixtime(devdt, '%Y-%m-%d %H:%i:%s') >" & Label11.Tag
        Cad = Cad & " order by devdt"
        miRsAux.Open Cad, Cn, adOpenKeyset, adLockPessimistic, adCmdText
        Cad = ""
        If Not miRsAux.EOF Then
            
            
            While Not miRsAux.EOF
                NumRegElim = NumRegElim + 1
                '                       tmppresencia(Id,idtra,Fecha,H1,codusu,Incidencias)
                'cad = cad & ", (" & NumRegElim & "," & miRsAux!usrid & "," & DBSet(miRsAux!srvdt, "F") & ",'" & Format(miRsAux!srvdt, "hh:nn:ss")
                Cad = Cad & ", (" & NumRegElim & "," & miRsAux!usrid & "," & DBSet(miRsAux!horaRealUnix, "F") & ",'" & Format(miRsAux!horaRealUnix, "hh:nn:ss")
                Cad = Cad & "'," & vUsu.Codigo & ",0," & DBSet(miRsAux!DEVUID, "T")
                Cad = Cad & ")"

                miRsAux.MoveNext
            Wend
            miRsAux.MovePrevious
            If miRsAux!horaRealUnix > MaxFecha Then MaxFecha = miRsAux!horaRealUnix
        End If
        miRsAux.Close
        
        'tmpcombinada(IdTrabajador,Fecha,H1,codusu,idinci)
        ' -->  NomTrabajador: devUI    semana: luego pondre que id es el devUI en aripres.terminales
        If Cad <> "" Then
            Cad = Mid(Cad, 2)
            Cad = "INSERT INTO tmppresencia(Id,idtra,Fecha,H1,codusu,Incidencias,NomTrabajador) VALUES " & Cad
            conn.Execute Cad
        End If
    Next I
    Set ColTablas = Nothing
    
    
    
    
    
    'Veamos si todos los trabajadores existen
    
    Label11.Caption = "Comprobando trabajadores"
    Label11.Refresh
    espera 0.5
    Cad = "Select distinct idtra from tmppresencia where codusu =" & vUsu.Codigo & " ORDER BY 1"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not miRsAux.EOF
        Label11.Caption = "Biostar ID" & miRsAux!idTRa
        Label11.Refresh
        Cad = DevuelveDesdeBD("idtrabajador", "trabajadores", "idTraReloj2", CStr(miRsAux!idTRa), "T")
        If Cad = "" Then
            Cad = DevuelveNombreBiostar(CStr(miRsAux!idTRa))
            NumRegElim = NumRegElim + 1
            ListView1.ListItems.Add , , CStr(miRsAux!idTRa)
            ListView1.ListItems(NumRegElim).SubItems(1) = Cad
        Else
            'Para no volver a leer , pondremos el trabajador de Aripres en el campo seccion
            Cad = "UPDATE tmppresencia set seccion=" & Cad & " WHERE idtra=" & miRsAux!idTRa
            conn.Execute Cad
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    
    Label11.Caption = "Comprobando terminales"
    Label11.Refresh
    espera 0.5
    Cad = "Select distinct NomTrabajador from tmppresencia where codusu =" & vUsu.Codigo & " ORDER BY 1"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        Label11.Caption = "Biostar ID" & miRsAux!nomtrabajador
        Label11.Refresh
        Cad = "tipo='BIOSTAR' AND idterminal"
        Cad = DevuelveDesdeBD("id", "terminales", Cad, miRsAux!nomtrabajador, "T")
        If Cad = "" Then
            Cad = "TERMINAL Sin dar de alta:   " & miRsAux!nomtrabajador
            NumRegElim = NumRegElim + 1
            ListView1.ListItems.Add , , CStr(miRsAux!nomtrabajador)
            ListView1.ListItems(NumRegElim).SubItems(1) = Cad
        Else
            'Para no volver a leer , pondremos el trabajador de Aripres en el campo seccion
            Cad = "UPDATE tmppresencia set semana=" & Cad & " WHERE NomTrabajador=" & DBSet(miRsAux!nomtrabajador, "T")
            conn.Execute Cad
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    
    
    
    
    
    If ListView1.ListItems.Count = 0 Then
        
        Label11.Caption = "Insertando datos "
        Label11.Refresh
            
        
        Cad = DevuelveDesdeBD("max(secuencia)", "entradafichajes", "1", "1")
        Cad = Val(Cad) + 1
        
        'Estan todos los trabajadores en la BD
        Cad = " tmppresencia ,(SELECT @rownum:=" & Cad & " ) r where codusu = " & vUsu.Codigo
        
        '
        Cad = "SELECT @rownum:=@rownum+1 AS rownum ,seccion,fecha ,h1,incidencias,h1,semana FROM " & Cad & " ORDER BY fecha,h1"
        
        Cad = "INSERT INTO entradafichajes(Secuencia,idTrabajador,Fecha,Hora,idInci,HoraReal,reloj) " & Cad
        
        conn.Execute Cad
    
    
        
        EntradasRepetidasProceso Me.Label11
    
        'Nov 2018  --> Esta dentro la caribale para que lo ejecute o no If Not vEmpresa.AcabaJornadaDiaSiguiente
        espera 0.25
        HorasNocturnas Me.Label11
        
        
        Cad = "UPDATE  biostar2 set ultimaFechaLeida =" & DBSet(MaxFecha, "FH")
        conn.Execute Cad
        
        Text2.Text = Format(MaxFecha, "dd/mm/yyyy hh:nn:ss")
        
        
        
         'Octubre 2021
         ' Blanca: Grabamos el año, justo despues del trabajador
         If vEmpresa.QueEmpresa = vbBelgida Then
            '1ª y BELGIç
            Label11.Caption = "Generando fichero txt"
            Label11.Refresh
            
            If AbrirFicheroExportar(NF) Then
            
                Cad = "Select * from entradafichajes WHERE revisado =0 order by fecha,hora"
                Set miRsAux = New ADODB.Recordset
                miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If Not miRsAux.EOF Then
                   
                    
                    While Not miRsAux.EOF
                            Cad = Format(miRsAux!idTrabajador, "00000")  'trab
                            'OCt 2021
                            Cad = Cad & "," & Format(miRsAux!Fecha, "yyyy")  'año
                            '
                            Cad = Cad & "," & Format(miRsAux!Fecha, "mm")  'mes
                            Cad = Cad & "," & Format(miRsAux!Fecha, "dd")  'dia
                            Cad = Cad & "," & Format(miRsAux!Hora, "hh")   'hora
                            Cad = Cad & "," & Format(miRsAux!Hora, "nn")   'min
                            ''lo mando aunque no es importante
                            'Aux = Aux & "," & Format(rs!Fecha, "yyyy")   'año
                            'Aux = Aux & ",0000,"
                            ''mandare la secuencia
                            'Aux = Aux & Format(rs!Secuencia, "00000")
                            Cad = Cad & ",0000,0000,00000"
                            
                            
                            Print #NF, Cad
                            
                            
                            miRsAux.MoveNext
                    Wend
                    
                    Cad = "UPDATE  entradafichajes SET revisado =1 where revisado=0"
                    conn.Execute Cad
                    
                    Label11.Caption = "Cerrando fichero"
                    Label11.Refresh
                    espera 1
                    
                    
                End If 'de belgida
                miRsAux.Close
                Set miRsAux = Nothing
                Close #NF
            End If
   
        End If
        
        
        
        
        
        
        
        
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
    
    Cad = ""
    For NumRegElim = 1 To ListView1.ListItems.Count
        Cad = Cad & ListView1.ListItems(NumRegElim).Text & Chr(9) & ListView1.ListItems(NumRegElim).SubItems(1) & vbCrLf
    Next NumRegElim
    Clipboard.SetText Cad
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
        
End Sub

Private Sub Form_Activate()
    If cmdLeer.Tag = 1 Then
        cmdLeer.Tag = 0
        DoEvents
        Screen.MousePointer = vbHourglass
        If AbrirConexion Then
            'Vamos a veer cual fue la ultima fecha
            If UltimaLecturaBD Then
                cmdLeer.Enabled = True
                Me.cmdCambiar.Enabled = True
            End If
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
    
    On Error GoTo eAbrirConexion
    AbrirConexion = False
    
    Cad = DevuelveDesdeBD("configreloj2", "empresas", "1", "1")
    If Cad = "" Then
        MsgBox "Mal configurado BIOSTAR2", vbExclamation
        Exit Function
    End If
    Set Cn = New Connection
    Cn.CursorLocation = adUseServer

    
    
    'Este es el que hay que dejar
    Cad = "DSN=" & Cad & ""   'Aripres4;DESC=MySQL ODBC 3.51 Driver DSN;;;;OPTION=3;STMT=;
    '    Cad = Cad & ";Persist Security Info=true"
    Cn.Open Cad
    
    AbrirConexion = True
    Exit Function
eAbrirConexion:
    MuestraError Err.Number, Err.Description, Cad
    Text2.Text = "ERROR"
    Set Cn = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set Cn = Nothing
End Sub


Private Function UltimaLecturaBD() As Boolean
    On Error GoTo eUltimaLecturaBD
    
    UltimaLecturaBD = False
    
    Cad = "Select ultimaFechaLeida FROM biostar2"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then Cad = Format(miRsAux.Fields(0), "dd/mm/yyyy hh:mm:ss")
    End If
    If Cad = "" Then Cad = "01/01/2020 00:00:00"
    
    Me.Text2.Text = Cad
    
    miRsAux.Close
    UltimaLecturaBD = True
eUltimaLecturaBD:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
End Function






Private Function AbrirFicheroExportar(ByRef NF As Integer) As Boolean
Dim Fichero As String
    On Error GoTo eAbrirFicheroExportar
    
    AbrirFicheroExportar = False
    
    Fichero = vEmpresa.DirMarcajes & "\HU" & Format(Now, "yyyyMMddhhnnss") & "T" & Format(99, "00") & ".txt"
    NF = FreeFile
    Open Fichero For Output As NF
    
    AbrirFicheroExportar = True
    
    Exit Function
eAbrirFicheroExportar:
        MuestraError Err.Number, , Err.Description
        
        
End Function

