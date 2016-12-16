VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTareaActual 
   Caption         =   "Seleccionar posteriores"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15870
   Icon            =   "frmTareaActual.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   15870
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   240
      TabIndex        =   13
      Top             =   0
      Width           =   14835
      Begin VB.CommandButton cmdSelecc 
         Height          =   375
         Index           =   0
         Left            =   12000
         Picture         =   "frmTareaActual.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Seleccionar anteriores al filtro"
         Top             =   360
         Width           =   435
      End
      Begin VB.CommandButton cmdSelecc 
         Height          =   375
         Index           =   1
         Left            =   12480
         Picture         =   "frmTareaActual.frx":6C94
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Seleccionar posteriores al filtro"
         Top             =   360
         Width           =   435
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Multi-seleccion"
         Height          =   315
         Left            =   9360
         TabIndex        =   32
         Top             =   390
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   2
         Left            =   10800
         TabIndex        =   31
         Text            =   "Filtro hora"
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdfecha 
         Caption         =   "+"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   28
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdfecha 
         Caption         =   "-"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   27
         Top             =   240
         Width           =   255
      End
      Begin VB.Frame FrameTapa 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   1080
         TabIndex        =   26
         Top             =   840
         Width           =   5175
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Index           =   1
         Left            =   14160
         Picture         =   "frmTareaActual.frx":70D6
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Modificar marcaje"
         Top             =   360
         Width           =   435
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Index           =   0
         Left            =   13560
         Picture         =   "frmTareaActual.frx":7660
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Agregar marcaje"
         Top             =   360
         Width           =   435
      End
      Begin VB.CommandButton cmdImpimir 
         Height          =   375
         Left            =   6240
         Picture         =   "frmTareaActual.frx":7BEA
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Height          =   435
         Index           =   0
         Left            =   7560
         Picture         =   "frmTareaActual.frx":85EC
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Añadir"
         Top             =   330
         Width           =   435
      End
      Begin VB.OptionButton optTicaje 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   21
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton optTicaje 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Height          =   435
         Index           =   1
         Left            =   8040
         Picture         =   "frmTareaActual.frx":8FEE
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Modificar"
         Top             =   330
         Width           =   435
      End
      Begin VB.CommandButton Command1 
         Height          =   435
         Index           =   2
         Left            =   8520
         Picture         =   "frmTareaActual.frx":E7D0
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Borrar"
         Top             =   330
         Width           =   435
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Actualizar"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Line Line1 
         X1              =   9240
         X2              =   9240
         Y1              =   360
         Y2              =   720
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmTareaActual.frx":13C5A
         ToolTipText     =   "Buscar fecha"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Marcajes"
         Height          =   195
         Left            =   6810
         TabIndex        =   19
         Top             =   450
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   450
      End
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2955
      Left            =   60
      TabIndex        =   12
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
         Object.Width           =   1411
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
      Left            =   300
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareaActual.frx":13CE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareaActual.frx":13FFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareaActual.frx":14319
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareaActual.frx":14D2B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2655
      Left            =   3000
      TabIndex        =   5
      Top             =   1560
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2655
      Left            =   60
      TabIndex        =   4
      Top             =   1620
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4683
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   915
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9735
      Begin VB.CommandButton cmdImpTarea 
         Height          =   375
         Left            =   3720
         Picture         =   "frmTareaActual.frx":1A94D
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   300
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   420
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   420
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Actualizar"
         Height          =   375
         Index           =   0
         Left            =   2400
         TabIndex        =   2
         Top             =   300
         Width           =   1215
      End
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   315
         Left            =   4320
         TabIndex        =   6
         Top             =   360
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   8
         Left            =   0
         Picture         =   "frmTareaActual.frx":1C9BF
         ToolTipText     =   "Buscar fecha"
         Top             =   0
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   720
         Picture         =   "frmTareaActual.frx":1CA4A
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hora"
         Height          =   195
         Left            =   1560
         TabIndex        =   8
         Top             =   180
         Width           =   450
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   180
         Width           =   450
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Hora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   6900
      TabIndex        =   11
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Trabajador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Tarea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "frmTareaActual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
        ' 0.- Tarea actual. Para las tareas
        ' 1.- Marcajes de la fecha seleccionada.
        '       Es decir, de la tabla de entrada, sin procesar,
        '
Public QueFecha As Date
        
Private WithEvents frmc As frmCal
Attribute frmc.VB_VarHelpID = -1
Private WithEvents frmHoras As frmHorasMarcajes
Attribute frmHoras.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Dim PrimeraVez As Boolean
Dim Tamanyo As Long
Dim Contador As Long
Dim Cad As String
Dim Modifi As Boolean



Private Sub Check1_Click()
Dim I As Integer
        ListView2.MultiSelect = Check1.Value
        cmdSelecc(0).Visible = Check1.Value = 1
        cmdSelecc(1).Visible = Check1.Value = 1
        Command3(0).Visible = Check1.Value = 1
        Command3(1).Visible = Check1.Value = 1
        Me.Text1(2).Visible = Check1.Value = 1
        
        If Not Check1.Value Then
            For I = 1 To ListView2.ListItems.Count
                ListView2.ListItems(I).Selected = False
            Next I
            Set ListView2.SelectedItem = Nothing
        End If
End Sub

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


'    If ListView2.ListItems.Count = 0 Then
'        MsgBox "Ningun dato a imprimir", vbExclamation
'        Exit Sub
'    End If
'
    
    'Vamos a imprimir los datos
'    Screen.MousePointer = vbHourglass
'    ImprimirTicajeActual
'
'    If optTicaje(0).Value Then
'        Cad = "pOrden= {tmpcombinada.idtrabajador}|"
'    Else
'        Cad = "pOrden= {trabajadores.nomtrabajador}|"
'    End If
'
'    frmImprimir.Opcion = 32
'    frmImprimir.FormulaSeleccion = "{tmpcombinada.codusu} = " & vUsu.Codigo
'    frmImprimir.OtrosParametros = Cad
'    frmImprimir.NumeroParametros = 1
'    frmImprimir.Show vbModal
'    Screen.MousePointer = vbDefault
    
    
    CadenaDesdeOtroForm = Text1(1).Text
    frmListado.Opcion = 12
    frmListado.Show vbModal
End Sub






Private Sub cmdImpTarea_Click()
    'Imprime la tarea actual de producicion
    If TreeView1.Nodes.Count = 0 Then Exit Sub
'
'
'    frmImprimir.Opcion = 143
'    frmImprimir.NumeroParametros = 1
'    frmImprimir.OtrosParametros = "Texto1= ""Fecha / hora :     " & Text1(0).Text & "   -  " & Text2.Text & """|"
'    frmImprimir.Show vbModal
End Sub

Private Sub cmdSelecc_Click(Index As Integer)
Dim I As Integer
Dim B As Boolean
Dim Hora As Date
    If Me.ListView2.ListItems.Count = 0 Then Exit Sub
    
    
    If Text1(2).Text = "" Then Text1(2).Text = "Filtro hora"
    If Text1(2).Text = "Filtro hora" Then
        MsgBox "Escriba una hora en el campo de filtro", vbExclamation
        PonerFoco Text1(2)
    Else
        
        Hora = CDate(Text1(2).Text)
        For Contador = 1 To ListView2.ListItems.Count
            B = False
            For I = 2 To 17
                If ListView2.ListItems(Contador).SubItems(I) = "" Then
                    Exit For
                Else
                    If Index = 0 Then
                        'Seleccionamos los que tengan un marcaje anterior
                        If CDate(Me.ListView2.ListItems(Contador).SubItems(I)) <= Hora Then B = True
                        Exit For
                    Else
                        'Marcaje posterior
                        If CDate(Me.ListView2.ListItems(Contador).SubItems(I)) >= Hora Then
                            B = True
                            Exit For
                        End If
                    End If
                End If
            Next I
            Me.ListView2.ListItems(Contador).Selected = B
        Next Contador
        PonFocoLW
    End If
End Sub

Private Sub PonFocoLW()
    On Error Resume Next
    ListView2.SetFocus
    Err.Clear
End Sub

Private Sub Combo1_Click()
    If PrimeraVez Then Exit Sub
    PonMarcajes
End Sub

Private Sub Command1_Click(Index As Integer)
Dim valor As Long
    
    If Index > 0 Then
        If ListView2.SelectedItem Is Nothing Then
            MsgBox "Seleccione un trabajador", vbExclamation
            Exit Sub
        End If
    End If
    Modifi = False
    Select Case Index
    Case 0, 1
        'INSERTAR
        Contador = -1
        Me.Tag = ""
        
        If Index = 1 Then
            'Modificar
            Contador = Val(ListView2.SelectedItem.Text)
            Me.Tag = ListView2.SelectedItem.SubItems(1)
        Else
            ' INSERTAR
            If Text1(1).Text = "" Then Exit Sub
            If Not IsDate(Text1(1).Text) Then
                MsgBox "Campo fecha incorrecto", vbExclamation
                Exit Sub
            End If
            Cad = "Codigo|idTrabajador|N|00000|15·"
            Cad = Cad & "Nombre|nomtrabajador|T||60·"
            Cad = Cad & "Tarjeta|numtarjeta|T||20·"
            Set frmB = New frmBuscaGrid
            frmB.vTabla = "Trabajadores"
            frmB.vCampos = Cad
            frmB.vDevuelve = "0|1|"
            frmB.vSelElem = 0
            frmB.vTitulo = "TRABAJADORES"
            frmB.Show vbModal
                    
            
            
            If Contador > 0 Then
                Cad = "Va a crear marcajes para el trabajador: " & Me.Tag
                Cad = Cad & "   ¿Desea continuar?"
                If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Contador = -1
            End If
        End If
        If Contador < 1 Then Exit Sub
        valor = Contador
        InsertarModificar
                        
        
    Case 2
        'ELIMINAR
                Cad = "Va a eliminar ""TODOS"" los marcajes para el trabajador: " & ListView2.SelectedItem.SubItems(1) & " en la fecha: " & Text1(1).Text
                Cad = Cad & "   ¿Desea continuar?"
                If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
                    
                Cad = "DELETE from EntradaFichajes WHERE idTrabajador=" & ListView2.SelectedItem.Text
                Cad = Cad & " AND Fecha = '" & Format(CDate(Text1(1).Text), FormatoFecha) & "'"
                conn.Execute Cad
                Modifi = True
    End Select
    If Modifi Then
        Screen.MousePointer = vbHourglass
        PonMarcajes
        espera 0.5
        'Volvemos a situarlo en donde estaba
        Set ListView2.SelectedItem = Nothing
        For Tamanyo = 1 To ListView2.ListItems.Count
            If Val(ListView2.ListItems(Tamanyo).Text) = valor Then
                Set ListView2.SelectedItem = ListView2.ListItems(Tamanyo)
                ListView2.SelectedItem.EnsureVisible
            Else
                ListView2.ListItems(Tamanyo).Selected = False
            End If
        Next Tamanyo
        
        Screen.MousePointer = vbDefault
    End If
    Me.Tag = ""
End Sub


Private Function InsertarModificar() As Boolean
Dim idCal As String
    'QUITAR### esta var
    Dim vM As CMarcajes
    Dim vH As CHorarios
    
    InsertarModificar = False
    
    'El marcaje solo lo utilizare para poner el Codigo del trabajador
    Set vM = New CMarcajes
    vM.idTrabajador = Contador
    vM.Fecha = CDate(Text1(1).Text)


    'Horario
    If vEmpresa.CreaCalDiariaTra Then
        'Por ejemplo TEINSA.
        'Cad trabajador tienen una entrada en calendariot
        Cad = DevuelveDesdeBD("idhorario", "calendariot", "idTrabajador", vM.idTrabajador, "N")
    Else
        
        'En alzira los horarios no van POR trabajador, si no que lo tiene el calendario
        Cad = "trabajadores.idcal=calendariol.idcal AND idtrabajador"
        Cad = DevuelveDesdeBD("idhorario", "trabajadores,calendariol", Cad, vM.idTrabajador, "N")
    End If
    
    
    If Cad = "" Then
        Contador = 0
    Else
        Contador = Val(Cad)
    End If
    If Contador < 1 Then
        MsgBox "Error obteniendo horario", vbExclamation
        Exit Function
    End If
    Set vH = New CHorarios
    If vH.Leer(CInt(Contador), vM.Fecha, 0) = 0 Then

        Set frmHoras = New frmHorasMarcajes
        frmHoras.Nombre = Me.Tag
        Set frmHoras.vH = vH
        Set frmHoras.vM = vM
        frmHoras.Nombre = Me.Tag
        frmHoras.Opcion = Opcion  'Marcajes
        frmHoras.Show vbModal
        Set frmHoras = Nothing
    End If
    Set vM = Nothing
    Set vH = Nothing
    InsertarModificar = True
End Function


Private Sub Command2_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    pb1.Value = 0
    pb1.Visible = True
    If Opcion = 0 Then
        'Generatemporal
    Else
        PonMarcajes
    End If
    pb1.Visible = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command3_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Contador = 0
    'Añadiremos en tmpCambioHor
    Cad = "DELETE from tmpCambioHor where codusu = " & vUsu.Codigo
    conn.Execute Cad
    espera 0.2
    Cad = "INSERT INTO tmpCambioHor values ("
    For Tamanyo = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(Tamanyo).Selected Then
            conn.Execute Cad & ListView2.ListItems(Tamanyo).Text & "," & vUsu.Codigo & ")"
            Contador = Contador + 1
        End If
    Next Tamanyo
    If Contador > 0 Then
            
               
         'Segun index
        Tamanyo = 0
        If Text1(2).Visible Then
            If Text1(2).Text <> "Filtro hora" Then Tamanyo = 1
        End If
        frmCambioHorario2.Inserta2Ticajes = Tamanyo = 1
        frmCambioHorario2.Opcion = 2 + Index
        frmCambioHorario2.Fecha = CDate(Me.Text1(1).Text)
        frmCambioHorario2.Show vbModal
        PonMarcajes
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Me.Refresh
        Command2_Click 0
        Me.Text1(Opcion).SetFocus
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    
    ListView2.Visible = Opcion = 1
    Frame2.Visible = Opcion = 0
    If Opcion = 0 Then
       Caption = "Tarea actual"
    Else
        Caption = "Marcaje actual"
    End If
    Frame1.Visible = Opcion <> 0
    
    'Imagenes
    Me.TreeView1.ImageList = Me.ImageList1
    Me.ListView1.SmallIcons = Me.ImageList1
    Me.ListView2.SmallIcons = Me.ImageList1
    
    'Fecha
    
    Text1(0).Text = Format(QueFecha, "dd/mm/yyyy")
    Text1(1).Text = Format(QueFecha, "dd/mm/yyyy")
    'Hora
    Text2.Text = Format(Hour(Now), "00") & ":00"

    pb1.Visible = False
    Combo1.Clear
    Combo1.AddItem "Sección.."
    Combo1.ListIndex = 0
    Set miRsAux = New ADODB.Recordset
    Cad = "select idseccion,nombre from secciones"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Combo1.AddItem miRsAux!Nombre
        Combo1.ItemData(Combo1.NewIndex) = miRsAux!IdSeccion
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    FrameTapa.Visible = vUsu.Nivel > 2
End Sub

Private Sub Form_Resize()

    If Me.WindowState = 1 Then Exit Sub
    If Me.Width < 5000 Then Me.Width = 5000
    If Me.Height < 3000 Then Me.Height = 3000
    
    Me.Frame2.Width = Me.Width - Frame2.Left - 300
    
    Me.Label3.Top = Me.Frame2.Top + Frame2.Height + 30
    Me.Label4.Top = Me.Label3.Top
    Me.Label5.Top = Me.Label3.Top
    
    
    Me.TreeView1.Top = Me.Label3.Top + Label3.Height + 30
    ListView1.Top = Me.TreeView1.Top
    Me.TreeView1.Height = Me.Height - Me.TreeView1.Top - 500
    Me.ListView1.Height = Me.TreeView1.Height
    
    'la proporcionde ancho = 2/5 3/5
    Me.TreeView1.Width = (2 * (Me.Width \ 5)) - 100
    Me.ListView1.Left = Me.TreeView1.Left + Me.TreeView1.Width + 100
    Me.ListView1.Width = Me.Width - Me.ListView1.Left - 420
    
    
    'Dentro del listview, las columnas en proporcion 8 a 2
    Me.Label4.Left = Me.ListView1.Left + 60
    Me.ListView1.ColumnHeaders(1).Width = CInt((ListView1.Width / 10) * 8) - 100
    Me.Label5.Left = Me.Label4.Left + Me.ListView1.ColumnHeaders(1).Width
    Me.ListView1.ColumnHeaders(2).Width = Me.ListView1.Width - Me.ListView1.ColumnHeaders(1).Width - 500
    
    
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

Private Sub frmc_Selec(vFecha As Date)
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
    Dim obj As Object

    Set frmc = New frmCal
    esq = imgFec(Index).Left
    dalt = imgFec(Index).Top
    
    
    Set obj = imgFec(Index).Container
    
    While imgFec(Index).Parent.Name <> obj.Name
        esq = esq + obj.Left
        dalt = dalt + obj.Top
        Set obj = obj.Container
    Wend
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
    
    
    frmc.Left = esq + imgFec(Index).Parent.Left + 30
    frmc.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

    imgFec(1).Tag = Index '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If Text1(Index).Text <> "" Then frmc.NovaData = Text1(Index).Text
    ' ********************************************

    frmc.Show vbModal
    Set frmc = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco Text1(CByte(imgFec(1).Tag)) '<===
    Command2_Click 0
    ' ********************************************
End Sub

Private Sub ListView2_DblClick()
    If Not ListView2.SelectedItem Is Nothing Then
        Command1_Click 1  'modificar
    End If
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
    If Index = 2 Then
        'HORA
    If Text1(Index).Text = "" Then
            Text1(Index).Text = "Filtro hora"
        Else
            If Text1(Index).Text <> "Filtro hora" Then
                If Not PonerFormatoHora(Text1(Index)) Then Text1(Index).Text = "Filtro hora"
            End If
        End If
    Else
        If Not EsFechaOK(Text1(Index)) Then
            MsgBox "No es una fecha correcta: " & Text1(Index).Text, vbExclamation
            Text1(Index).Text = ""
            Focus Text1(Index)
        End If
    End If
End Sub

Private Sub Text2_GotFocus()
    PonFoco Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Text2_LostFocus()
    Text2.Text = Trim(Text2.Text)
    If Text2.Text = "" Then Exit Sub
    Text2.Text = TransformaPuntosHoras(Text2.Text)
    If Not IsDate(Text2.Text) Then
        MsgBox "No es una hora correcta: " & Text2.Text, vbExclamation
        Text2.Text = ""
        Focus Text2
        Exit Sub
    End If
    
    Text2.Text = Format(Text2.Text, "hh:mm")
End Sub






'---------------------------------------------------------------------------
'----------  Ponemos los datos de la tarea en este momento
'----------------------------------------------------------------------------

'ESTE TROZO ES PARA KIMALDI

'La siguiente funcion esta copiada de procesar marcajes
'Private Sub Generatemporal()
'Dim SQL As String
'Dim RS As ADODB.Recordset
'Dim AntTarea As Long
'Dim Procesar As Boolean
'Dim salida As Boolean
'Dim Insertar As Boolean
'Dim Trabajador As Long
'Dim Hora As Date
'Dim NOD As Node
'
'On Error GoTo ETemporal
'    'Borramos los nodos
'    TreeView1.Nodes.Clear
'    ListView1.ListItems.Clear
'    Me.Refresh
'
'    Me.Tag = "Obtener tmtTareasRealizadas"
'
'    'Obtenemos la anterior ultima tarea k estaban realizando
'    AntTarea = 0
'    Set RS = New ADODB.Recordset
'    SQL = "Select Tarea from TareasRealizadas order by Fecha,Horafin"
'    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Not RS.EOF Then
'        RS.MoveLast 'Vemos el ultimo registro
'        AntTarea = DBLet(RS!Tarea, "N")
'    End If
'    RS.Close
'
'    'Eliminamos datos temporales
'    Conn.Execute "delete from tmpTareasRealizadas"
'
'    'SQL
'    SQL = " from MarcajesKimaldi  where (Fecha = #" & Format(Text1(0).Text, "yyyy/mm/dd") & "#)"
'
'    'Progress bar
'    RS.Open "Select count(*) " & SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    Tamanyo = 0
'    If Not RS.EOF Then Tamanyo = DBLet(RS.Fields(0), "N")
'    RS.Close
'
'    If Tamanyo = 0 Then Exit Sub
'
'    Me.Tag = "Obtener desde KIMALDI"
'    'Recorremos la tabla Kimaldi entre las fechas seleccionadas
'    ' y para cada registro de trabajador le insertamos su tarea correspondiente
'    SQL = " from MarcajesKimaldi  where (Fecha = #" & Format(Text1(0).Text, "yyyy/mm/dd") & "#)"
'    SQL = SQL & " ORDER BY Nodo,Fecha,Hora"
'    RS.Open "Select * " & SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    'Progress
'    Contador = 0
'
'    While Not RS.EOF
'        'Progress
'        Contador = Contador + 1
'        pb1.Value = CInt((Contador / Tamanyo) * 1000)
'
'        Procesar = True
'        salida = False
'        If DBLet(RS!tipomens) <> "" Then
'            If RS!tipomens <> "S" Then
'                Procesar = False
'            Else
'                salida = True
'            End If
'        End If
'
'        If Procesar Then
'            Insertar = False
'            If Not salida Then
'                'Veremos si es marcaje de trabajador o tarea
'                If Mid(RS!Marcaje, 1, 1) = mConfig.DigitoTrabajadores Then
'                    'Trabajador
'                    Insertar = CodigoCorrecto(True, RS!Marcaje, Trabajador)
'                Else
'                    'Tarea
'                    CodigoCorrecto False, RS!Marcaje, AntTarea
'
'                End If
'            Else
'                AntTarea = -1
'                Insertar = True
'                'Hay k ver k trabajador
'                CodigoCorrecto True, RS!Marcaje, Trabajador
'            End If
'
'            If Insertar Then
'                SQL = "INSERT into tmpTareasRealizadas (Fecha,Hora,  Trabajador,Tarea) VALUES ("
'                SQL = SQL & "#" & Format(RS!Fecha, "yyyy/mm/dd") & "#"
'                SQL = SQL & ",#" & Format(RS!Hora, "hh:mm:ss") & "#,"
'                SQL = SQL & Trabajador & ","
'                SQL = SQL & AntTarea & ")"
'                Conn.Execute SQL
'            End If
'        End If
'
'
'
'
'        'Siguiente
'        RS.MoveNext
'    Wend
'    RS.Close
'
'
'
'    'Llegados aqui hacemos el resto
'    pb1.Value = 0
'    Me.Refresh
'
'    'Borramos la tabla temporal
'    Conn.Execute "Delete from tmpTareaActual"
'
'    'Desde tmptareasrealizadas para cada trabajador vamos buscando su ultima tarea
'
'
'    Me.Tag = "Desde tmpTareasRealizadas"
'    SQL = "SELECT Count(tmpTareasRealizadas.trabajador) AS CuentaDetrabajador"
'    SQL = SQL & " From tmpTareasRealizadas"
'    SQL = SQL & " WHERE Hora <= #" & Format(Text2.Text, "hh:mm") & "#"
'    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    Tamanyo = 0
'    If Not RS.EOF Then
'        Tamanyo = DBLet(RS.Fields(0), "N")
'    End If
'    RS.Close
'
'    If Tamanyo = 0 Then Exit Sub
'
'    SQL = " From tmpTareasRealizadas WHERE Hora <= #" & Format(Text2.Text, "hh:mm") & "#"
'    SQL = SQL & " GROUP BY tmpTareasRealizadas.trabajador"
'    RS.Open "Select trabajador " & SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'    Contador = 0
'
'    While Not RS.EOF
'        'Progress
'        Contador = Contador + 1
'        pb1.Value = CInt((Contador / Tamanyo) * 1000)
'
'        Trabajador = RS.Fields(0)
'
'        Insertar = DevuelveUltimo(Trabajador, Hora, AntTarea)
'        If Insertar Then
'            SQL = "INSERT INTO tmpTareaActual (Trabajador,Tarea,Hora) VALUES ("
'            SQL = SQL & Trabajador & "," & AntTarea & ",#" & Format(Hora, "hh:mm") & "#)"
'            Conn.Execute SQL
'        End If
'        'Siguiente
'        RS.MoveNext
'    Wend
'    RS.Close
'
'    'Para cargar el arbol
'    Me.Tag = "Cargar el arbol"
'
'    'Ahora cargamos el arbol de las tareas
'    SQL = "SELECT tmpTareaActual.Tarea, Tareas.Descripcion"
'    SQL = SQL & " FROM tmpTareaActual LEFT JOIN Tareas ON tmpTareaActual.Tarea = Tareas.idTarea"
'    SQL = SQL & " GROUP BY tmpTareaActual.tarea, Tareas.Descripcion;"
'    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    While Not RS.EOF
'        If IsNull(RS!Descripcion) Then
'            If RS!Tarea = -1 Then
'                SQL = "SALIDA"
'            Else
'                SQL = "TAREA desconocida"
'            End If
'        Else
'            SQL = RS!Descripcion
'        End If
'
'        Set NOD = TreeView1.Nodes.Add(, , "C" & CStr(RS!Tarea), SQL)
'        NOD.Tag = RS!Tarea
'        NOD.Image = 1
'        'Siguiente
'        RS.MoveNext
'    Wend
'    RS.Close
'
'    'Ponemos el primero de todos
'    If TreeView1.Nodes.Count > 0 Then
'        Set TreeView1.SelectedItem = TreeView1.Nodes(1)
'        Cargalistview
'        Me.Refresh
'    End If
'    Set RS = Nothing
'    Exit Sub
'ETemporal:
'    MuestraError Err.Number, Me.Tag & vbCrLf & Err.Description
'End Sub
'

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



'Devolvera si hay k insertar o no
Private Function DevuelveUltimo(Trabajador As Long, Hora As Date, Tarea As Long) As Boolean
Dim SQL As String
Dim RT As ADODB.Recordset

    Set RT = New ADODB.Recordset
    SQL = "Select * from tmpTareasRealizadas WHERE Trabajador = " & Trabajador
    SQL = SQL & " AND Hora<=#" & Format(Text2.Text, "hh:mm") & "# ORDER BY Hora"
    RT.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    DevuelveUltimo = False
    If Not RT.EOF Then
        Do
            Hora = RT!Hora
            Tarea = RT!Tarea
            RT.MoveNext
        Loop Until RT.EOF
        DevuelveUltimo = True
    End If
    RT.Close
    Set RT = Nothing
End Function




'Cargar el listview
Private Sub Cargalistview()
Dim Item As ListItem
Dim RS As ADODB.Recordset
Dim SQL As String
    On Error GoTo ECarga
    
    ListView1.ListItems.Clear
    'Para no recargar si no hace falta
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    ListView1.Tag = TreeView1.SelectedItem.Text
    
    SQL = "SELECT Trabajadores.NomTrabajador, tmpTareaActual.Hora"
    SQL = SQL & " FROM tmpTareaActual INNER JOIN Trabajadores ON tmpTareaActual.Trabajador = Trabajadores.IdTrabajador "
    SQL = SQL & " WHERE tarea =" & TreeView1.SelectedItem.Tag & " ORDER BY Hora"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Set Item = ListView1.ListItems.Add(, , RS.Fields(0))
        Item.SmallIcon = 2
        Item.SubItems(1) = Format(RS!Hora, "hh:mm")
    
        'Siguiente
        RS.MoveNext
    Wend
    RS.Close

ECarga:
    If Err.Number <> 0 Then _
        MuestraError Err.Number, "Carga LISTVIEW" & vbCrLf & Err.Description
    Set RS = Nothing
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    If ListView1.Tag <> Node.Text Then
        Screen.MousePointer = vbHourglass
        Cargalistview
        Screen.MousePointer = vbDefault
    End If
End Sub





'---------------------------------------------------------------------------
'----------  Ponemos los datos de la tarea en este momento
'----------------------------------------------------------------------------

Private Sub PonMarcajes()
    'Dos recordsets
    Dim I As Integer
    Dim RS As ADODB.Recordset
    Dim RT As ADODB.Recordset
    Dim SQL As String
    Dim Item As ListItem
    
    Dim HoraPintar As Date
    
    ListView2.ListItems.Clear
    SQL = "SELECT EntradaFichajes.idTrabajador, Trabajadores.NomTrabajador"
    SQL = SQL & " FROM EntradaFichajes ,Trabajadores WHERE EntradaFichajes.idTrabajador = Trabajadores.IdTrabajador"
    SQL = SQL & " AND Fecha = '" & Format(Text1(1).Text, FormatoFecha) & "' "
    If vUsu.Nivel > 2 Then SQL = SQL & " AND Trabajadores.controlnomina >0"
    If Combo1.ListIndex > 0 Then SQL = SQL & " AND trabajadores.seccion=" & Combo1.ItemData(Combo1.ListIndex)
        
    SQL = SQL & " GROUP BY EntradaFichajes.idTrabajador, Trabajadores.NomTrabajador"
    SQL = SQL & " ORDER BY "
    If optTicaje(0).Value Then
        SQL = SQL & " EntradaFichajes.idTrabajador"
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
    SQL = "Select EntradaFichajes.*," & SQL
    SQL = SQL & " as acabalga FROM EntradaFichajes WHERE Fecha = '" & Format(Text1(1).Text, FormatoFecha) & "'"
    SQL = SQL & " AND idTrabajador = "
    While Not RS.EOF
        RT.Open SQL & RS.Fields(0) & " ORDER BY HoraReal", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        I = 2
        
        Set Item = ListView2.ListItems.Add(, , RS.Fields(0))
        Item.SubItems(1) = RS.Fields(1)
        While Not RT.EOF
            'If i < 8 Then  He puesto 2 mas
            If I < 17 Then
                
                'If RT!HoraReal > "23:59:59" Then
                '    HoraPintar = DateAdd("h", -24, RT!HoraReal)
                'ElseIf RT!HoraReal < "00:00:00" Then
                '    HoraPintar = DateAdd("h", 24, RT!HoraReal)
                'Else
                '    HoraPintar = RT!HoraReal
                '
                'End If
                Item.SubItems(I) = Format(RT!acabalga, "hh:mm")
            End If
            I = I + 1
            RT.MoveNext
        Wend
        
        'El icono
        If I Mod 2 = 0 Then
            Item.SmallIcon = 3
        Else
            Item.SmallIcon = 4
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
