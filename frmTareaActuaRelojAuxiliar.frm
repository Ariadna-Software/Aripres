VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTareaActuaRelojAuxiliar 
   Caption         =   "Relojes auxiliares"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   12390
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11235
      Begin VB.CommandButton cmdVer 
         Caption         =   "Ver"
         Height          =   375
         Left            =   3000
         TabIndex        =   12
         ToolTipText     =   "Ver marcajes trabajador"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdResumen 
         Caption         =   "Resumen"
         Height          =   375
         Left            =   9720
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdfecha 
         Caption         =   "+"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   9
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdfecha 
         Caption         =   "-"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdImpimir 
         Caption         =   "Listado"
         Height          =   375
         Left            =   8400
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optTicaje 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   5040
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optTicaje 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Actualizar"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Informes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   7200
         TabIndex        =   11
         Top             =   480
         Width           =   900
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   600
         ToolTipText     =   "Buscar fecha"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   450
      End
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2955
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   5212
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "T1"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "T2"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "T3"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "T4"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "T5"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "T6"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "T7"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "T8"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "T9"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "T10"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "T11"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "T12"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "T13"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "T14"
         Object.Width           =   1147
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareaActuaRelojAuxiliar.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareaActuaRelojAuxiliar.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareaActuaRelojAuxiliar.frx":0634
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareaActuaRelojAuxiliar.frx":1046
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareaActuaRelojAuxiliar.frx":6C68
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTareaActuaRelojAuxiliar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MaxCol = 16

Public QueFecha As Date
        
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Dim PrimeraVez As Boolean
Dim Tamanyo As Long
Dim Contador As Long
Dim Cad As String
Dim Modifi As Boolean


Private Sub cmdfecha_Click(Index As Integer)
Dim F As Date
    If Text1(1).Text <> "" Then
        Screen.MousePointer = vbHourglass
        Frame1.Enabled = False
        If Index = 0 Then
            Contador = -1
        Else
            Contador = 1
        End If
                
        F = DateAdd("d", Contador, CDate(Text1(1).Text))
        Text1(1).Text = Format(F, "dd/mm/yyyy")
        Command2_Click 0
        Frame1.Enabled = True
        Me.Refresh
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdImpimir_Click()

    AbrirRpt 18
    
End Sub

Private Sub AbrirRpt(Indice As Integer)
    CadenaDesdeOtroForm = Text1(1).Text
    frmListado.Opcion = Indice
    frmListado.Show vbModal
End Sub



Private Sub cmdResumen_Click()
    AbrirRpt 19
End Sub

Private Sub cmdVer_Click()
    If ListView2.ListItems.Count = 0 Then Exit Sub
    If ListView2.SelectedItem Is Nothing Then Exit Sub
        
        
        
        
    
    Dim SQL As String
    
    SQL = "horareal"
    If vEmpresa.HorarioNocturno2 Then
        If vEmpresa.QueEmpresa = 2 Then
            SQL = "if(hour(horareal)<0,ADDTIME(hora , '24:00:00' ),''),if(hour(horareal)>24,ADDTIME(hora , '-24:00:00' ),horareal)"
        End If
    End If
    SQL = "Select entradafichajauxliares.*," & SQL
    SQL = SQL & " as acabalga FROM entradafichajauxliares WHERE Fecha = '" & Format(Text1(1).Text, FormatoFecha) & "'"
    SQL = SQL & " AND idTrabajador = " & Me.ListView2.SelectedItem.Text & " ORDER BY HoraReal"
   
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    frmVerMar.Show vbModal
    
End Sub

Private Sub Command2_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    PonMarcajes
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Me.Refresh
        Command2_Click 0
        Me.Text1(1).SetFocus
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    Me.Icon = frmMain.Icon
    ListView2.Visible = True
    

    'Imagenes

    Me.ListView2.SmallIcons = Me.ImageList1
    
    'Fecha
    
    Cad = DevuelveDesdeBD("max(fecha)", "entradafichajauxliares", "1", "1")
    If Cad = "" Then Cad = Format(Now, "dd/mm/yyyy")
    QueFecha = CDate(Cad)
    Text1(1).Text = Format(QueFecha, "dd/mm/yyyy")
    
    
    CargaColumnas

End Sub


Private Sub CargaColumnas()


    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , "Codigo", "Codigo", 900
    ListView2.ColumnHeaders.Add , "Nombre", "Nombre", 4000
    ListView2.ColumnHeaders.Add , "Excede", "Ex", 400
    For NumRegElim = 1 To MaxCol
        ListView2.ColumnHeaders.Add , "T" & NumRegElim, "T" & NumRegElim, 650
    Next
End Sub


Private Sub Form_Resize()

    If Me.WindowState = 1 Then Exit Sub
    If Me.Width < 5000 Then Me.Width = 5000
    If Me.Height < 3000 Then Me.Height = 3000
    



    'LIST2
    ListView2.Top = Me.Frame1.Top + Frame1.Height + 30
    ListView2.Width = Me.Width - 320
    ListView2.Height = Me.Height - ListView2.Top - 420
End Sub


Private Sub PonFoco(ByRef T As TextBox)
    T.SelStart = 0
    T.SelLength = Len(T.Text)
End Sub

Private Sub Focus(T As TextBox)
On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub


Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
    Contador = vCodigo
    Me.Tag = vCadena
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Contador = RecuperaValor(CadenaDevuelta, 1)
    Me.Tag = RecuperaValor(CadenaDevuelta, 2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    conn.Close
    Set conn = Nothing
    End
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text1(CInt(imgFec(1).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmHoras_HayModificacion(SiNo As Boolean, vOpcion As Byte)
    Modifi = SiNo
End Sub

Private Sub Image2_Click(Index As Integer)
'    Set frmC = New frmCal
'    frmC = Now
'    frmC.Tag = Index
'    If Text1(Index).Text <> "" Then
'        If IsDate(Text1(Index).Text) Then frmC.Fecha = CDate(Text1(Index).Text)
'    End If
'    frmC.Show vbModal
'    Set frmC = Nothing
End Sub

Private Sub imgFec_Click(Index As Integer)
   Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim Obj As Object

    Set frmC = New frmCal
    esq = imgFec(Index).Left
    dalt = imgFec(Index).Top
    
    
    Set Obj = imgFec(Index).Container
    
    While imgFec(Index).Parent.Name <> Obj.Name
        esq = esq + Obj.Left
        dalt = dalt + Obj.Top
        Set Obj = Obj.Container
    Wend
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
    
    
    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

    imgFec(1).Tag = Index '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If Text1(Index).Text <> "" Then frmC.NovaData = Text1(Index).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco Text1(CByte(imgFec(1).Tag)) '<===
    Command2_Click 0
    ' ********************************************
End Sub

Private Sub ListView2_DblClick()

        cmdVer_Click     'modificar

End Sub

Private Sub ListView2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ListView2_DblClick
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    PonFoco Text1(Index)
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).Text = "" Then Exit Sub
    If Not EsFechaOK(Text1(Index)) Then
        MsgBox "No es una fecha correcta: " & Text1(Index).Text, vbExclamation
        Text1(Index).Text = ""
        Focus Text1(Index)
    End If
End Sub










Private Function CodigoCorrecto(Trabajador As Boolean, Marcaje As String, valor As Long) As Boolean
Dim SQL As String
Dim RT As ADODB.Recordset

    Set RT = New ADODB.Recordset
    CodigoCorrecto = False
    If Trabajador Then
        SQL = "Select idTrabajador from Trabajadores where numtarjeta = '" & Marcaje & "';"
    Else
        SQL = "Select idTarea from Tareas where tarjeta = '" & Marcaje & "';"
    End If

        
    RT.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RT.EOF Then
        CodigoCorrecto = True
        valor = RT.Fields(0)
    Else
        valor = -1
    End If
    RT.Close
    Set RT = Nothing
    
End Function









'---------------------------------------------------------------------------
'----------  Ponemos los datos de la tarea en este momento
'----------------------------------------------------------------------------

Private Sub PonMarcajes()
    'Dos recordsets
    Dim i As Integer
    Dim RS As ADODB.Recordset
    Dim RT As ADODB.Recordset
    Dim SQL As String
    Dim Item As ListItem
    Dim HN As Date
    Dim HT As Currency
    Dim Difer As Currency
    Dim HoraPintar As Date
    
    ListView2.ListItems.Clear
    SQL = "SELECT entradafichajauxliares.idTrabajador, Trabajadores.NomTrabajador"
    SQL = SQL & " FROM entradafichajauxliares ,Trabajadores WHERE entradafichajauxliares.idTrabajador = Trabajadores.IdTrabajador"
    SQL = SQL & " AND Fecha = '" & Format(Text1(1).Text, FormatoFecha) & "' "
    If vUsu.Nivel > 2 Then SQL = SQL & " AND Trabajadores.controlnomina >0"
    SQL = SQL & " GROUP BY entradafichajauxliares.idTrabajador, Trabajadores.NomTrabajador"
    SQL = SQL & " ORDER BY "
    If optTicaje(0).Value Then
        SQL = SQL & " entradafichajauxliares.idTrabajador"
    Else
        SQL = SQL & "  Trabajadores.NomTrabajador"
    End If
    
    Set RS = New ADODB.Recordset
    Set RT = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    SQL = "horareal"
    If vEmpresa.HorarioNocturno2 Then
        If vEmpresa.QueEmpresa = 2 Then
            SQL = "if(hour(horareal)<0,ADDTIME(hora , '24:00:00' ),''),if(hour(horareal)>24,ADDTIME(hora , '-24:00:00' ),horareal)"
        End If
    End If
    SQL = "Select entradafichajauxliares.*," & SQL
    SQL = SQL & " as acabalga FROM entradafichajauxliares WHERE Fecha = '" & Format(Text1(1).Text, FormatoFecha) & "'"
    SQL = SQL & " AND idTrabajador = "
    While Not RS.EOF
        RT.Open SQL & RS.Fields(0) & " ORDER BY HoraReal", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
        Set Item = ListView2.ListItems.Add(, , RS.Fields(0))
        Item.SubItems(1) = RS.Fields(1)
        Item.SubItems(2) = " "
        i = 3
        HT = 0
        While Not RT.EOF
            
            If i < MaxCol + 3 Then
                
                'If RT!HoraReal > "23:59:59" Then
                '    HoraPintar = DateAdd("h", -24, RT!HoraReal)
                'ElseIf RT!HoraReal < "00:00:00" Then
                '    HoraPintar = DateAdd("h", 24, RT!HoraReal)
                'Else
                '    HoraPintar = RT!HoraReal
                '
                'End If
                Item.SubItems(i) = Format(RT!acabalga, "hh:mm")
            Else
                Item.SubItems(2) = "*"
            End If
            
            If i > 3 Then
                'Se calcula cada dos. Primero entrada y segundo salida
                If i Mod 2 = 0 Then
                    Difer = DateDiff("n", HN, RT!acabalga)
                    HT = HT + Difer
                End If
            End If
            HN = RT!acabalga
            i = i + 1
            RT.MoveNext
        Wend
        
        'El icono
        If i Mod 2 = 0 Then
            
                Item.SmallIcon = 4
                Item.ToolTipText = "Impares"
                
            
        Else
        
            If HT > 60 Then
                'Mas de una HORA de parada
                
                Item.SmallIcon = 5
                Item.ToolTipText = "> 1 hora.     " & HT & " min."
            Else
                Item.ToolTipText = ""
                Item.SmallIcon = 3
            End If
        End If
        RT.Close
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    Set RT = Nothing
End Sub




'-------------------------------------------------------------




'Private Sub ImprimirTicajeActual()
'Dim SQL As String
'Dim vC As CHorarios
'Dim HI As Date
'Dim HIAustada As Date
'Dim HF As Date
'Dim Horas As Currency
'Dim Ajustadas As Currency
'Dim QuitoAlmuerzo As Boolean
'Dim difer As Currency
'Dim Minutos As Integer
'
'    On Error GoTo eImprimirTicajeActual
'
'    SQL = "Delete from tmpCombinada where codusu = " & vUsu.Codigo
'    conn.Execute SQL
'
'    If vEmpresa.QueEmpresa = 2 Then
'        'Solo para alzira, de momento
'        Set vC = New CHorarios
'
'
'
'    End If
'
'
'
'    For Contador = 1 To ListView2.ListItems.Count
'        'Para los ticajes
'
'        If vEmpresa.QueEmpresa = 2 Then
'            'select * from calendariol,trabajadores where  calendariol.idcal=trabajadores.idcal and fecha='' and idtrabajador=1
'            Cad = "calendariol.idcal=trabajadores.idcal and fecha=" & DBSet(Text1(1).Text, "F") & " and idtrabajador"
'            SQL = "trabajadores.idcal"
'            Cad = DevuelveDesdeBD("idhorario", "calendariol,trabajadores", Cad, ListView2.ListItems(Contador).Text, "N", SQL)
'            If Val(Cad) = 0 Then Err.Raise 513, , "Error obteniendo horario trabajador: " & ListView2.ListItems(Contador).Text
'
'            If Val(Cad) <> vC.IdHorario Then
'                Minutos = 1
'                If vC.Leer(CInt(Cad), CDate(Text1(1).Text), CInt(SQL)) > 0 Then
'                    Err.Raise 513, , "Error obteniendo horario / calendario: " & Cad & " - " & SQL
'                Else
'
'                    If vC.Rectificar > 0 Then
'                      If vC.Rectificar = vbRecESCuarto Then
'                        Minutos = 15
'                      Else
'                          Minutos = 30   'Entradas salidas cada media hora
'                      End If
'                    End If
'                End If
'            End If
'        End If
'        Cad = ""
'        SQL = ""
'        Tamanyo = 2
'        Horas = 0
'        Ajustadas = 0
'        QuitoAlmuerzo = False
'        While Tamanyo < 15
'            If ListView2.ListItems(Contador).SubItems(Tamanyo) = "" Then
'
'                If Tamanyo >= 15 Then
'                    MsgBox "Incorrecto numero ticajes"
'                Else
'                    If vEmpresa.QueEmpresa = 2 Then
'                        If (Tamanyo Mod 2) = 0 Then
'                            'Correcto. Marcajes pares
'
'                        Else
'                            'MAL. Impares
'                            Horas = 0
'
'                        End If
'                    End If
'
'                End If
'
'
'                Tamanyo = 15 'Para salirnos
'
'
'            Else
'                If vEmpresa.QueEmpresa = 2 Then
'
'
'                    If (Tamanyo Mod 2) = 1 Then
'                        'Segundo ticaje. Calculo horas
'                        HF = CDate(ListView2.ListItems(Contador).SubItems(Tamanyo) & ":00")
'                        difer = DateDiff("n", HI, HF)
'                        Horas = Horas + difer
'
'
'                        'Ajustada
'                        HF = HoraRectificada(HF, vEmpresa.AjusteSalida, Minutos)
'                        difer = DateDiff("n", HIAustada, HF)
'
'                        Ajustadas = Ajustadas + difer
'
'
'                    Else
'
'                        HI = CDate(ListView2.ListItems(Contador).SubItems(Tamanyo) & ":00")
'                        HIAustada = HoraRectificada(HI, vEmpresa.AjusteEntrada, Minutos)
'
'                        difer = 0
'                        If vC.DtoAlm > 0 Then
'                            If HIAustada < vC.HoraDtoAlm Then QuitoAlmuerzo = True
'                        End If
'
'                    End If
'
'                End If
'
'              Cad = Cad & ",'" & ListView2.ListItems(Contador).SubItems(Tamanyo) & ":00'"
'              SQL = SQL & ",H" & Tamanyo - 1
'            End If
'            Tamanyo = Tamanyo + 1
'        Wend
'
'        SQL = "INSERT INTO tmpCombinada(codusu,idTrabajador,Fecha,HT,HE" & SQL & ") VALUES (" & vUsu.Codigo & ","
'        SQL = SQL & ListView2.ListItems(Contador).Text & ",'" & Format(Text1(1).Text, FormatoFecha) & "',"
'        If Horas = 0 Then
'            SQL = SQL & "0,0"
'        Else
'            'Pasamos las horas a sexagesimal
'            Horas = Round(Horas / 60, 2)
'            Ajustadas = Round(Ajustadas / 60, 2)
'            If QuitoAlmuerzo Then
'                Horas = Horas - vC.DtoAlm
'                Ajustadas = Ajustadas - vC.DtoAlm
'            End If
'            SQL = SQL & DBSet(Horas, "N") & ","
'            SQL = SQL & DBSet(Ajustadas, "N")
'        End If
'
'
'        Cad = Cad & ")"
'        conn.Execute SQL & Cad
'    Next Contador
'
'
'eImprimirTicajeActual:
'    If Err.Number <> 0 Then MuestraError Err.Number, "", Err.Description
'    Set vC = Nothing
'End Sub
