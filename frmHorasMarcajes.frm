VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHorasMarcajes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Marcajes m�quina"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "frmHorasMarcajes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Volver"
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   0
      Top             =   6240
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Aceptar"
      Height          =   375
      Index           =   0
      Left            =   5520
      TabIndex        =   15
      Top             =   6240
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Eliminar"
      Height          =   495
      Index           =   2
      Left            =   5580
      Picture         =   "frmHorasMarcajes.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4140
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Modificar"
      Height          =   495
      Index           =   1
      Left            =   5580
      Picture         =   "frmHorasMarcajes.frx":6954
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3540
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Nuevo"
      Height          =   495
      Index           =   0
      Left            =   5580
      Picture         =   "frmHorasMarcajes.frx":6A56
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2940
      Width           =   975
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3615
      Left            =   180
      TabIndex        =   8
      Top             =   3000
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Hora"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Incidencia"
         Object.Width           =   5636
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "idInci"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Acabalgado"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Reloj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Maps"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "map"
         Object.Width           =   1058
      EndProperty
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1995
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   6555
      Begin VB.TextBox txtHorario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   5040
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1500
         Width           =   1100
      End
      Begin VB.TextBox txtHorario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   1800
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1500
         Width           =   1100
      End
      Begin VB.TextBox txtHorario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   5040
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   960
         Width           =   1100
      End
      Begin VB.TextBox txtHorario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1800
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   960
         Width           =   1100
      End
      Begin VB.TextBox txtHorario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2220
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   420
         Width           =   2775
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Salida segunda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   3540
         TabIndex        =   13
         Top             =   1560
         Width           =   1425
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Entrada segunda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1545
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Salida primera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   3540
         TabIndex        =   11
         Top             =   1020
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Entrada primera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   1020
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Horario del empleado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1980
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Marcajes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   5475
   End
End
Attribute VB_Name = "frmHorasMarcajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmH As frmSoloHora
Attribute frmH.VB_VarHelpID = -1
Public Event HayModificacion(SiNo As Boolean, vOpcion As Byte)
Public Opcion As Byte
    '0.- Horas en entradamarcajes
    '1.- En HSITORICO

Public Nombre As String
Public vM As CMarcajes
Public vH As CHorarios


'Para saber a que item nos referimos
Private Secuencia As Long
Private SeHaModificado As Boolean
Private PuedeSalir As Boolean

Private Sub Command1_Click(Index As Integer)
Dim Cad As String
Dim RC As Byte

Select Case Index
Case 0
    'Nuevo
    Secuencia = -1
    Set frmH = New frmSoloHora
    frmH.Hora = ""
    frmH.Inci = 0
    frmH.CadInci = ""
    frmH.TipoAcabalgada = 2 'NORMAL
    frmH.Show vbModal
    Set frmH = Nothing
Case 1
    'Modificar
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    Secuencia = ListView1.SelectedItem.Tag
    Set frmH = New frmSoloHora
    frmH.Hora = ListView1.SelectedItem
    frmH.Inci = ListView1.SelectedItem.SubItems(2)
    frmH.CadInci = ListView1.SelectedItem.SubItems(1)
    frmH.TipoAcabalgada = ListView1.SelectedItem.SubItems(3)
    frmH.vReloj = ListView1.SelectedItem.SubItems(4)
    frmH.Show vbModal
    Set frmH = Nothing
Case 2
    'Eliminar
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    
    If Opcion = 0 Then
        Cad = "Seguro que desea eliminar el marcaje efectuado " & vbCrLf
        Cad = Cad & " a las " & ListView1.SelectedItem.Text & vbCrLf
        If ListView1.SelectedItem.SubItems(2) <> 0 Then _
            Cad = Cad & "y con la incidencia : " & ListView1.SelectedItem.SubItems(1)
        RC = MsgBox(Cad, vbQuestion + vbYesNo)
        If RC = vbYes Then
            'Eliminamos
            Cad = "Delete from EntradaMarcajes where secuencia=" & ListView1.SelectedItem.Tag
            conn.Execute Cad
            SeHaModificado = True
            ListView1.ListItems.Remove ListView1.SelectedItem.Index
         End If
         
    Else
        'Elminamos del HCO
        Cad = "Seguro que desea eliminar el marcaje efectuado " & vbCrLf
        Cad = Cad & " a las " & ListView1.SelectedItem.Text & vbCrLf
        If ListView1.SelectedItem.SubItems(2) <> 0 Then _
            Cad = Cad & "y con la incidencia : " & ListView1.SelectedItem.SubItems(1)
        RC = MsgBox(Cad, vbQuestion + vbYesNo)
        If RC = vbYes Then
            'Eliminamos
            
            Cad = "Delete from EntradaFichajes where secuencia=" & ListView1.SelectedItem.Tag
            conn.Execute Cad
            SeHaModificado = True
            ListView1.ListItems.Remove ListView1.SelectedItem.Index
        End If
    End If
End Select
End Sub

Private Sub Command2_Click(Index As Integer)
    If SeHaModificado Then
        RaiseEvent HayModificacion(True, Opcion)
    End If
    PuedeSalir = True
    Unload Me
End Sub

Private Sub Form_Load()
    PuedeSalir = False
    Label1.Caption = Nombre
    If Opcion = 0 Then
        Caption = "Modificar marcajes"
        Label1.ForeColor = vbBlack
    Else
        Caption = "Modificar ENTRADAS fichajes"
        Label1.ForeColor = vbRed
    End If
    
    'Ponemos el horario oficial
    SeHaModificado = False
    txtHorario(0).Text = vH.NomHorario
    txtHorario(1).Text = vH.HoraE1
    txtHorario(2).Text = vH.HoraS1
    txtHorario(3).Text = vH.HoraE2
    txtHorario(4).Text = vH.HoraS2
    
    Me.ListView1.SmallIcons = frmPpal.ImageListReloj
    
    Cargalistview
End Sub



Private Sub Cargalistview()
Dim RS As ADODB.Recordset
Dim SQL As String
Dim itm As ListItem

    ListView1.ListItems.Clear
    Set RS = New ADODB.Recordset
    
    'Enero 2015
    'Dias acabalgados
    'Una columna mas para decir las horas pertenen al dia en horas normales o en las acabalgadas
    ' Es decir, cuando
    '  AcabalgadoDiaInicio -->  true -> Siginifica que hay horas por encima de las 24:00
    '                           false -> siginifica que hay horas por debajo de las 00:00
    'ACabalgado:  0: normal     1: Hora inferior a 00:00    2: hora superior a 23:59:59
    If Opcion = 0 Then
        SQL = "SELECT EntradaMarcajes.Hora, Incidencias.NomInci, EntradaMarcajes.idInci,Secuencia"
        
        SQL = SQL & " ,if(Hora<'0:00:00',ADDTIME(hora , '24:00:00' ),if(hour(Hora)>=24,ADDTIME(hora , '-24:00:00' ),Hora)) HoraPintar  "
        
        SQL = SQL & " ,if(Hora<'0:00:00',1,if(hour(Hora)>=24,2,0)) Acabalgado "
        SQL = SQL & ", Reloj QueReloj"
        SQL = SQL & ",latitud,longitud , terminales.descripcion, terminales.tipo"
        SQL = SQL & " FROM Incidencias , EntradaMarcajes "
        SQL = SQL & " LEFT JOIN terminales ON reloj=id"
        SQL = SQL & " WHERE EntradaMarcajes.idInci = Incidencias.IdInci"
        SQL = SQL & " AND idMarcaje=" & vM.Entrada
        SQL = SQL & " Order by Hora"
    Else
        SQL = "    SELECT EntradaFichajes.Hora, Incidencias.NomInci, EntradaFichajes.idInci,Secuencia"
        
        SQL = SQL & " ,if(Hora<'0:00:00',ADDTIME(hora , '24:00:00' ),if(hour(Hora)>=24,ADDTIME(hora , '-24:00:00' ),Hora)) HoraPintar "
        
        SQL = SQL & " ,if(Hora<'0:00:00',1,if(hour(Hora)>=24,2,0)) Acabalgado "
        SQL = SQL & ", Reloj QueReloj"
        SQL = SQL & ",latitud,longitud , terminales.descripcion, terminales.tipo"
        SQL = SQL & " FROM EntradaFichajes INNER JOIN Incidencias ON EntradaFichajes.idInci = Incidencias.IdInci"
        SQL = SQL & " LEFT JOIN terminales ON reloj=id"
        SQL = SQL & " WHERE idTrabajador=" & vM.idTrabajador
        SQL = SQL & " AND Fecha = '" & Format(vM.Fecha, FormatoFecha) & "'"
        SQL = SQL & " Order by Hora"
    End If
    
    
    
    RS.Open SQL, conn, , , adCmdText
    While Not RS.EOF
        Set itm = ListView1.ListItems.Add(, , Format(RS!HoraPintar, "hh:mm:ss"))
        If RS!IdInci = 0 Then
            itm.SubItems(1) = ""
            Else
            itm.SubItems(1) = RS!NomInci
        End If
        itm.SubItems(2) = RS!IdInci
        
        itm.SubItems(3) = RS!acabalgado
        
        itm.SubItems(4) = DBLet(RS!quereloj, "N")
        
        itm.SubItems(5) = " "
        itm.SubItems(6) = " "
        If Not IsNull(RS!Longitud) And Not IsNull(RS!latitud) Then
            itm.SubItems(5) = TransformaComasPuntos(CStr(RS!latitud)) & "," & TransformaComasPuntos(CStr(RS!Longitud))
            itm.SubItems(6) = "*"
        End If
        
        itm.Tag = RS!Secuencia
        
        itm.SmallIcon = 0
       
        
        If itm.SubItems(4) <> "0" Then
            
            'SQL = CargaIconoTerminales(DBLet(RS!tipo, "T"))
            
            itm.SmallIcon = CInt(CargaIconoTerminalesZONA(RS!quereloj, SQL))
            If SQL = "" Then SQL = "IdTerminal: " & RS!quereloj
            itm.ToolTipText = DBLet("Area " & SQL, "T")
        End If
        
        
        
        
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not PuedeSalir Then Cancel = 1
End Sub

Private Sub frmH_Seleccionar(vHora As Date, vInci As Integer, Acabagado As Byte, kReloj As Integer)
Dim RS As ADODB.Recordset
Dim valor As Long
Dim Tabla As String
Dim Cad As String
Dim Hora As String
Dim LaHora As Integer

    On Error GoTo ErrFRMSelecionar
    Set RS = New ADODB.Recordset
    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    If Opcion = 0 Then
        Tabla = "EntradaMarcajes"
    Else
        Tabla = "EntradaFichajes"  'Los fichajes
    End If
    
    
    'Fijo la hora
    ' 0: normal     1: Hora inferior a 00:00    2: hora superior a 23:59:59
    If Acabagado = 0 Then
        Hora = Format(vHora, "hh:mm:ss")
    Else
        
        If Acabagado = 1 Then
            Hora = Horas_Quitar24(vHora, False)
            
        Else
            If vEmpresa.HorarioNocturno2 Then
                'LO QUE HACIA ANTES
                Hora = Format(vHora, ":nn:ss")
                LaHora = Hour(vHora)
        
                LaHora = LaHora - 24
                Hora = Format(LaHora, "00") & Hora
            Else
                Hora = Format(vHora, ":nn:ss")
                LaHora = Hour(vHora) + 24
                Hora = Format(LaHora, "00") & Hora
            End If
        End If
        
    End If
    
    
    If Secuencia < 0 Then
        'Nuevo marcaje
        valor = 1
    
          
        RS.Open "Select max(Secuencia) FROM " & Tabla, conn, , , adCmdText
        If Not RS.EOF Then
            If Not IsNull(RS.Fields(0)) Then valor = RS.Fields(0) + 1
        End If
        RS.Close
        
        'A�adimos
        Cad = "INSERT INTO " & Tabla & "(Secuencia,idTrabajador"
        If Opcion = 0 Then Cad = Cad & ", idMarcaje"
        Cad = Cad & ",Fecha , Hora, idInci, HoraReal,Reloj) VALUES ("
        
        Cad = Cad & valor & "," & vM.idTrabajador & ","
        If Opcion = 0 Then Cad = Cad & vM.Entrada & ","
        Cad = Cad & DBSet(vM.Fecha, "F") & ","
        
        
        Cad = Cad & "'" & Hora & "'," & vInci & ",'" & Hora & "'," & kReloj & ")"
       
    Else
        'MODIFICAR
        
        
        Cad = "UPDATE " & Tabla & " SET hora = '" & Hora & "' , idinci = " & vInci
        '
        If Opcion = 1 Then Cad = Cad & ",  HoraReal = '" & Hora & "'"
        Cad = Cad & ",  Reloj = " & kReloj
        
        Cad = Cad & " WHERE  secuencia=" & Secuencia
        
    End If
    conn.Execute Cad
    SeHaModificado = True
    
    espera 0.25
    Set RS = Nothing
Cargalistview
Exit Sub
ErrFRMSelecionar:
    MsgBox "Error: " & Err.Description, vbExclamation
End Sub

Private Sub ListView1_DblClick()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    
    If Trim(ListView1.SelectedItem.SubItems(6)) = "" Then
        Command1_Click (1)
    Else
        AbrirGeolocalizacion
    End If
End Sub


Private Sub AbrirGeolocalizacion()
Dim Cad As String
    Cad = "https://www.google.com/maps/?q=" & ListView1.SelectedItem.SubItems(5)
    LanzaVisorMimeDocumento Me.Hwnd, Cad
    
End Sub
