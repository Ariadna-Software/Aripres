VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCalculoHorasMesConEstrc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nominas mes"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   14355
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameTipoAlzira 
      Height          =   855
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   14295
      Begin VB.CommandButton Command2 
         Height          =   315
         Index           =   4
         Left            =   8520
         Picture         =   "frmCalculoHorasMesConEstrc.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Eliminar de nomina"
         Top             =   360
         Width           =   375
      End
      Begin VB.ComboBox cboSeccion 
         Height          =   315
         Left            =   5400
         TabIndex        =   33
         Text            =   "Combo1"
         Top             =   360
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Index           =   2
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Imprimir"
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Index           =   1
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Recuperar datos"
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdGenHorasAlzi 
         Caption         =   "Calcular horas"
         Height          =   315
         Left            =   2640
         TabIndex        =   24
         Top             =   300
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1680
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   300
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   300
         Width           =   1455
      End
      Begin VB.CommandButton cmdGeneraAlz 
         Caption         =   "Genera nominas"
         Height          =   315
         Left            =   12600
         TabIndex        =   27
         Top             =   240
         Width           =   1515
      End
      Begin VB.OptionButton Option2 
         Caption         =   "C�digo"
         Height          =   255
         Index           =   0
         Left            =   4080
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   26
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Index           =   0
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Guardar datos"
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   4080
      TabIndex        =   10
      Top             =   4560
      Width           =   4095
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   315
         Left            =   300
         TabIndex        =   11
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.Frame FrameMes 
      Height          =   855
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   11475
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   7560
         TabIndex        =   32
         ToolTipText     =   "Depuracion"
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdQuitar 
         Height          =   315
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Eliminar trabajador de la lista"
         Top             =   300
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Index           =   3
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "MODIFICAR DATOS TRABAJADOR"
         Top             =   300
         Width           =   375
      End
      Begin VB.CommandButton cmdImprimir 
         Height          =   315
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Imprimir datos actuales"
         Top             =   300
         Width           =   375
      End
      Begin VB.CommandButton cmdHPlus 
         Caption         =   "Quitar H+"
         Height          =   315
         Index           =   1
         Left            =   6240
         TabIndex        =   18
         Top             =   300
         Width           =   1035
      End
      Begin VB.CommandButton cmdHPlus 
         Caption         =   "A�adir H+"
         Height          =   315
         Index           =   0
         Left            =   5160
         TabIndex        =   17
         Top             =   300
         Width           =   1035
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "C�digo"
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdBaja 
         Caption         =   "Baja trabajador"
         Height          =   315
         Left            =   5160
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Genera nominas"
         Height          =   315
         Left            =   9720
         TabIndex        =   12
         Top             =   300
         Width           =   1755
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   300
         Width           =   855
      End
      Begin VB.CommandButton cmdGenHoras 
         Caption         =   "Calcular horas"
         Height          =   315
         Left            =   2700
         TabIndex        =   2
         Top             =   300
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView lw1 
      Height          =   4215
      Left            =   60
      TabIndex        =   3
      Top             =   1320
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   18
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cod"
         Object.Width           =   1500
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "D / H"
         Object.Width           =   1693
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "D"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "HN"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "est"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "HE"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "D"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Hor"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "D"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "H"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "H+"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Ant"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Post"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Bruto"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Imp. estru."
         Object.Width           =   1623
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Plus"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Neto"
         Object.Width           =   2364
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
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCalculoHorasMesConEstrc.frx":0A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCalculoHorasMesConEstrc.frx":0F9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCalculoHorasMesConEstrc.frx":1536
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCalculoHorasMesConEstrc.frx":1850
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCalculoHorasMesConEstrc.frx":1DEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCalculoHorasMesConEstrc.frx":223C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCalculoHorasMesConEstrc.frx":27D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCalculoHorasMesConEstrc.frx":9038
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Bolsa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Left            =   7380
      TabIndex        =   14
      Top             =   960
      Width           =   480
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   7260
      X2              =   8400
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00008080&
      BorderWidth     =   2
      X1              =   6000
      X2              =   7140
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   4740
      X2              =   5880
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   3420
      X2              =   4560
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   2100
      X2              =   3240
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   1860
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Nomina"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   195
      Left            =   6120
      TabIndex        =   8
      Top             =   960
      Width           =   645
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Saldo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   4980
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Trabajadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3600
      TabIndex        =   6
      Top             =   960
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Oficial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   2220
      TabIndex        =   5
      Top             =   960
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Trabajador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   930
   End
   Begin VB.Menu mnPopup 
      Caption         =   "mnPopup"
      Visible         =   0   'False
      Begin VB.Menu mnVerDatos 
         Caption         =   "Ver datos/dia"
      End
      Begin VB.Menu mnbarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnModificaHoras 
         Caption         =   "Modificar Horas"
      End
   End
End
Attribute VB_Name = "frmCalculoHorasMesConEstrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion2 As Byte
    '0.- Normal   . Es decir, como picassent. Autmaticamente compensa horas y demas

    

    
    
Private Sql As String
Dim Importe1 As Currency
Dim PrimeraVez As Boolean

Private Nod As ListItem
Private HorasxDia2 As Currency  'La leemos de parametros

Private RT As ADODB.Recordset

Dim CadParam As String


Private Sub cmdGeneraAlz_Click()
    
    ProcesoDeGeneracionNominas
End Sub

Private Sub cmdGenHoras_Click()
Dim d As Integer
Dim FI As Date
Dim FF As Date

    If Combo1.ListIndex < 0 Then
        MsgBox "Seleccione un mes", vbExclamation
        Exit Sub
    End If
    If Val(Text1.Text) = 0 Then
        MsgBox "A�o incorrecto.", vbExclamation
        Exit Sub
    End If
        
    If lw1.ListItems.Count > 0 Then
        Sql = "Ya ha generado datos. � Seguro que desea volverlos a generar ?"
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
        
    Sql = "/" & Combo1.ListIndex + 1 & "/" & Text1.Text
    FI = CDate("01" & Sql)
    d = DiasMes(Combo1.ListIndex + 1, CInt(Text1.Text))
    FF = CDate(d & "/" & Combo1.ListIndex + 1 & "/" & Text1.Text)
        
        
    If ComprobarMarcajesCorrectos(FI, FF, True) = 0 Then
        Sql = "No existe marcajes entre las fechas."
        MsgBox Sql, vbExclamation
        Exit Sub
    End If
        
    If ComprobarMarcajesCorrectos(FI, FF, False) <> 0 Then
        Sql = "Existen marcajes incorrectos entre las fechas. �Desea continuar?"
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
        
        

        
    Label1.Caption = "Comienzo proceso"
    Frame1.Visible = True
    Me.Refresh
    
    Screen.MousePointer = vbHourglass
    
    lw1.ListItems.Clear
    
    CalculaEntreFechas FI, FF
    Frame1.Visible = False
    Screen.MousePointer = vbDefault

End Sub


Private Sub CalculaEntreFechas(FI As Date, FF As Date)
Dim RS As Recordset
Dim Horas As Currency
Dim Dias As Integer
Dim AntiguaFormaProcesar As Boolean
Dim Aux As String
Dim idCal As Integer
Dim vSeccion As Integer

    conn.Execute "DELETE FROM tmpHorasMesHorario"

    'Para comprobar si estando de baja han trabajado
    'En tmpPresencia voy a guardar
    conn.Execute "DELETE FROM tmpCombinada"

    Set RS = New ADODB.Recordset
    
    Aux = "select idhorario,idcal from calendariol where fecha >='2018-01-01' and idcal in (select idcal from trabajadores)  group by 1,2"
    RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Label1.Caption = "Obtener horarios"
    Label1.Refresh
    
    idCal = 1
    While Not RS.EOF
        idCal = RS!idCal
        'Horas = CalculaHorasHorario(IdCal, RS.Fields(0), Dias, FI, FF, False )
        Horas = CalculaHorasHorarioALZ(idCal, Dias, FI, FF)
        
        
        If Horas > 0 Then
            'Insertamos en tmp HORAS
            'Antes febrero2017
            'conn.Execute "INSERT INTO tmpHorasMesHorario(idHorario,Horas,Dias) VALUES (" & RS.Fields(0) & "," & TransformaComasPuntos(CStr(Horas)) & "," & Dias & ")"
            'Ahora
            conn.Execute "INSERT INTO tmpHorasMesHorario(idHorario,Horas,Dias) VALUES (" & idCal & "," & TransformaComasPuntos(CStr(Horas)) & "," & Dias & ")"
            
        End If
        RS.MoveNext
    Wend
    RS.Close
    
    Label1.Caption = "Horas trabajadas"
    Label1.Refresh
    
    vSeccion = -1
    If cboSeccion.ListIndex > 0 Then vSeccion = cboSeccion.ItemData(cboSeccion.ListIndex)
    CalculaHorasTrabajadas FI, FF, 0, vSeccion
    Me.Refresh
    

    
    Label1.Caption = "Datos periodo"
    Label1.Refresh
    CalculaDatosMes FI, FF, 0, -1
    
    Me.Refresh
    
    Label1.Caption = "Combina datos"
    Label1.Refresh
    CombinaDatos FI, FF
    
    'AHora realizamos los calculos de horas k kedan y demas
    Label1.Caption = "Datos a compensar"
    Label1.Refresh
    CalculoDatosACompensar
    
    Me.Refresh
    
    'Hacemos las comensaciones por horas
    Label1.Caption = "Compensaciones"
    Label1.Refresh
    
    AntiguaFormaProcesar = Dir(App.Path & "\AntigFP.dat", vbArchive) <> ""
   ' AntiguaFormaProcesar = True
    
    
    '    Depuracion = (Check1.Value = 1)
    '    HacerCompensacionesPicassent FI, FF, Label1
    
    HacerCompensaciones FI, FF, Label1
    
    'Ajustamos los que no hayan trabakado nada
    AjustaDatosBajaMesEntero
    







    'Ajustamos los de jornadas semanales
    



    Label1.Caption = "Carga datos"
    Label1.Refresh
    CargaDatos



    'Ahora vamos a comprobar si alguno de los k ha estado de baja
    'En este periodo a trabajado
    If lw1.ListItems.Count > 0 Then
        Label1.Caption = "Comprobar bajas con dias Tra."
        Label1.Refresh
        RS.Open "Select * from tmpcombinada", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            If HaTrabajadoConBaja(RS) Then
                Dias = 0
                Do
                    Dias = Dias + 1
                    If Dias <= lw1.ListItems.Count Then
                        If lw1.ListItems(Dias).Text = RS!idTrabajador Then
                            'Pongo el icono distinto
                            lw1.ListItems(Dias).SmallIcon = 5
                            'Salgo
                            Dias = 32000
                        End If
                    End If
                Loop Until Dias > lw1.ListItems.Count
            End If   'De ha trabajado estando de baja
            'Siguiente caso
        RS.MoveNext
        Wend
        RS.Close
    End If
    
    
    If lw1.ListItems.Count > 0 Then
        Label1.Caption = "Comprobar trabajadores sin datos"
        Label1.Refresh
        
        
        For idCal = lw1.ListItems.Count To 1 Step -1
            If Me.lw1.ListItems(idCal).SmallIcon = 7 Then
                'En principioo, NO ha sido tratado
                For vSeccion = 2 To 14
                    Select Case vSeccion
                    Case 2
                        If lw1.ListItems(idCal).SubItems(vSeccion) <> "0/0,00" Then Exit For
                    Case 3, 7, 9
                        If lw1.ListItems(idCal).SubItems(vSeccion) <> "0" Then Exit For
                        
                    Case 10, 11
                        If lw1.ListItems(idCal).SubItems(vSeccion) <> " " Then Exit For
                        
                    Case Else
                        If lw1.ListItems(idCal).SubItems(vSeccion) <> "0,00" Then Exit For
                    End Select
                Next vSeccion
                If vSeccion > 13 Then
                    'Todo vacio
                    'Eliminamos
                
                    Sql = "DELETE FROM tmpdatosmes where trabajador =" & lw1.ListItems(idCal).Text
                    conn.Execute Sql
                    lw1.ListItems.Remove idCal
               
                    
                End If
            End If
        Next idCal
    End If
    
    
    Set RS = Nothing
End Sub






Private Sub cmdGenHorasAlzi_Click()
Dim d As Integer
Dim FI As Date
Dim FF As Date


    

    If Combo2.ListIndex < 0 Then
        MsgBox "Seleccione un mes", vbExclamation
        Exit Sub
    End If
    If Val(Text2.Text) = 0 Then
        MsgBox "A�o incorrecto.", vbExclamation
        Exit Sub
    End If
        
    If lw1.ListItems.Count > 0 Then
        Sql = "Ya ha generado datos. � Seguro que desea volverlos a generar ?"
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
        
    Sql = "/" & Combo2.ListIndex + 1 & "/" & Text2.Text
    FI = CDate("01" & Sql)
    d = DiasMes(Combo2.ListIndex + 1, CInt(Text2.Text))
    FF = CDate(d & "/" & Combo2.ListIndex + 1 & "/" & Text2.Text)
        
        
    If ComprobarMarcajesCorrectos(FI, FF, True) = 0 Then
        Sql = "No existe marcajes entre las fechas."
        MsgBox Sql, vbExclamation
        Exit Sub
    End If
        
    If ComprobarMarcajesCorrectos(FI, FF, False) <> 0 Then
        Sql = "Existen marcajes incorrectos entre las fechas. �Desea continuar?"
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
        
        

        
    Label1.Caption = "Comienzo proceso"
    Frame1.Visible = True
    Me.Refresh
    
    Screen.MousePointer = vbHourglass
    
    lw1.ListItems.Clear
    CalculaEntreFechas FI, FF
    
    Frame1.Visible = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdHPlus_Click(Index As Integer)
Dim Importe As Currency
Dim Imp1 As Currency
Dim RS As ADODB.Recordset
    
    
    
     
    
    
    'falta revisar,. No he cambiado nada de lo de aqui bajo
    
    
    
    If lw1.SelectedItem Is Nothing Then Exit Sub
    
    If Index = 1 Then
        Sql = "reestablecer horas plus"
    Else
        Sql = "a�adir horas plus"
    End If
    If MsgBox("Desea continuar con la opci�n " & Sql & " ?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        
    If Index = 0 Then
        'Si ya ha compensado le decimos k ya ha compensado
        If lw1.SelectedItem.SubItems(10) <> "0.00" Then
            MsgBox "Ya ha compensando horas. Quite la compensacion primero", vbExclamation
            Exit Sub
        End If
    Else
        If lw1.SelectedItem.SubItems(10) = "0.00" Then
            MsgBox "Ya ha compensando horas. Quite la compensacion primero", vbExclamation
            Exit Sub
        End If
    End If
    
    

    If Index = 0 Then
        'Cuando ponemos la baja calculamos si tiene horas en bolsa despues.
        'las tranformamos en euros de mas en anticpos
        Imp1 = -1
        Importe = ImporteFormateadoAmoneda(lw1.SelectedItem.SubItems(12))
             Do
                 Sql = "Introduzca las horas de PLUS para " & lw1.SelectedItem.SubItems(1) & "." & vbCrLf & "M�ximo:" & Format(Importe, "0.00")
                 Sql = InputBox(Sql, "Horas +")
                 If Sql <> "" Then
                     If IsNumeric(Sql) Then
                         Sql = TransformaPuntosComas(Sql)
                         Imp1 = CCur(Sql)
                         If Imp1 > 0 Then
                            If Imp1 > Importe Then
                                MsgBox "No puede poner mas horas de las que tiene", vbExclamation
                                Imp1 = 0
                            Else
                                Sql = ""
                            End If
                        End If
                     End If
                 End If
             Loop Until Sql = ""
                         
            If Sql = "" And Imp1 <= 0 Then Exit Sub
                    
                    
      '  Importe = ImporteFormateadoAmoneda(ListView1.SelectedItem.SubItems(12))
        
       
            Sql = "SELECT Categorias.Importe1, Categorias.Importe2, Trabajadores.IdTrabajador,PorcSS,PorcIRPF"
            Sql = Sql & " FROM Categorias INNER JOIN Trabajadores ON Categorias.IdCategoria = Trabajadores.idCategoria"
            Sql = Sql & " WHERE Trabajadores.IdTrabajador=" & lw1.SelectedItem.Text

            Set RS = New ADODB.Recordset
            RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If RS.EOF Then
                MsgBox "Error leyendo datos trabajador", vbExclamation
            Else
                'Le ponemos las horas de plus
                lw1.SelectedItem.SubItems(10) = Format(Imp1, FormatoImporte)
                'En la bolsa le dejo las k tenia menos las k lleva al plus
                Importe = Importe - Imp1
                lw1.SelectedItem.SubItems(12) = Format(Importe, FormatoImporte)
            
                Importe = Imp1 * RS.Fields(0) 'importe2    horas * importe
                
                'PLUS
                lw1.SelectedItem.SubItems(14) = Format(Importe, FormatoImporte)
                
                
                Imp1 = (Importe * RS!PorcSS) + (Importe * RS!PorcIRPF)
                Imp1 = Imp1 / 100
                Importe = Importe - Imp1
                Importe = Round(Importe, 2)
                

               
               
                'Importe origninal
                Imp1 = ImporteFormateadoAmoneda(lw1.SelectedItem.SubItems(13))
                Importe = Importe + Imp1
        
                'Ponemos las horas de plus
                lw1.SelectedItem.SubItems(13) = Format(Importe, FormatoImporte)

                lw1.SelectedItem.SmallIcon = 4 'Icono de h+

            End If
            RS.Close
 
    Else
    
        PonerBaja False
    End If
End Sub

Private Sub cmdImprimir_Click()
Dim i As Integer

    If lw1.ListItems.Count < 1 Then Exit Sub
    
    
    'FALTA###
    
    'Borramos las dos tablas k utiliza
    Sql = "DELETE FROM tmpPagosMes"
    conn.Execute Sql
    Sql = "DELETE FROM tmpHoras"
    conn.Execute Sql
    espera 0.1
    
    'Para cada list item vamos a ver lo k pagamos
    VariableCompartida = "INSERT INTO tmpPagosMes(idTrabajador,nombre,SS,IRPF,HT,HC,importe1,Importe2,"
    VariableCompartida = VariableCompartida & "NETO,preciohora1,Pagos,BRUTO,INGRESAR) VALUES ("
    'Son en realidad
    '  OFICIAL         TRABAJADA       NOMINA            BOLSA             IMPORTES
    '   D   H         D   HN  HC      D  H  H+      Antes   Despues    PAGOS  PLUS   ANTICPOS
    
    ' Dias Trabajados y duas nomina van en la tabla tmpHoras,, en campos Dias, HorasE
    
    For i = 1 To lw1.ListItems.Count
        With lw1.ListItems(i)
            Sql = .Text & ",'" & .SubItems(1) & "',"
            
            'OFICIALES
            Sql = Sql & Mid(.SubItems(2), 1, InStr(1, .SubItems(2), "/") - 1) & ",'"
            Sql = Sql & TransformaComasPuntos(Mid(.SubItems(2), InStr(1, .SubItems(2), "/") + 1)) & "',"
            
            
            'Horas Normales y compensables
            Sql = Sql & TransformaComasPuntos(.SubItems(4)) & "," & TransformaComasPuntos(.SubItems(5)) & ","
            
            'Horas Nomina y H+
            Sql = Sql & TransformaComasPuntos(.SubItems(9)) & "," & TransformaComasPuntos(.SubItems(10)) & ","
            
            'Bolsa antes y despues
            Sql = Sql & TransformaComasPuntos(.SubItems(11)) & "," & TransformaComasPuntos(.SubItems(12)) & ","
            
            'Importes: pagos, PULS y Anticpipos
            Sql = Sql & "0," & TransformaComasPuntos(CCur(.SubItems(13))) & ","
            Sql = Sql & TransformaComasPuntos(CCur(.SubItems(14))) & ")"
            conn.Execute VariableCompartida & Sql
            
            'Insertamos los dias en tmpHoras
            Sql = "INSERT INTO tmpHoras (trabajador,Dias,horasE) VALUES (" & .Text & ","
            Sql = Sql & .SubItems(3) & "," & .SubItems(8) & ")"
            conn.Execute Sql
        End With
    Next i
    
    If vEmpresa.NominaAutomatica Then
        Sql = "Mes= """ & UCase(Combo1.List(Combo1.ListIndex)) & " " & Text1.Text & """|"
    Else
        Sql = "Mes= """ & UCase(Combo2.List(Combo2.ListIndex)) & " " & Text2.Text & """|"
    End If
    frmImprimir.Opcion = 15
    frmImprimir.OtrosParametros = Sql
    frmImprimir.NumeroParametros = 1
    frmImprimir.Show vbModal
    
End Sub




Private Sub cmdQuitar_Click()
    'Eliminar los datos del trabjaodr
    On Error GoTo E1
            'Modificar datos trabajador
        If lw1.ListItems.Count = 0 Then Exit Sub
        If lw1.SelectedItem Is Nothing Then Exit Sub
        
        Sql = "�Desea eliminar de la nomina al trabajador: " & lw1.SelectedItem.Text & " - " & lw1.SelectedItem.SubItems(1) & "?"
        If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
            Sql = "DELETE FROM tmpDatosMes WHERE tmpDatosMes.trabajador =" & lw1.SelectedItem.Text
            conn.Execute Sql
            lw1.ListItems.Remove lw1.SelectedItem.Index
            
        End If


    Exit Sub
E1:
    MuestraError Err.Number, Err.Description
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Command1_Click()
    ProcesoDeGeneracionNominas
End Sub


Private Sub ProcesoDeGeneracionNominas()
Dim B As Boolean
Dim RS As ADODB.Recordset
Dim i As Integer


    

    If lw1.ListItems.Count < 1 Then Exit Sub
    
    'Preguntamos si desea continuar
    Sql = "Seguro que desea generar las n�minas con estos valores?"
    If MsgBox(Sql, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub
    
    

        i = DiasMes(Combo2.ListIndex + 1, CInt(Text2.Text))
        Sql = "'" & Text2.Text & "-" & Combo2.ListIndex + 1 & "-" & i & "'"
  
    Sql = "Select * from Nominas where Fecha = " & Sql
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    If Not RS.EOF Then Sql = "SI"
    RS.Close
    Set RS = Nothing
    
    If Sql <> "" Then
        
        MsgBox "Ya se han generado las nominas de este mes.", vbExclamation
        
    End If
    
    
    'pondremos un transaccion
    Screen.MousePointer = vbHourglass
    conn.BeginTrans
    
    B = genNominas
    
    If B Then
        conn.CommitTrans
        MsgBox "Proceso finalizado", vbInformation
        lw1.ListItems.Clear
        Unload Me
    Else
        conn.RollbackTrans
    End If
    Screen.MousePointer = vbDefault

End Sub



Private Sub Command2_Click(Index As Integer)
Dim i As Integer
Dim RS As ADODB.Recordset


    'De momento no hacemos nada
    If Index <> 4 Then Exit Sub

    Select Case Index
    Case 0
        'Guardar los datos
        Sql = "Desea guardar los cambios?"
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        ModificarRecuperar True
    Case 1
         'Recuperar datos
        Sql = "Desea recuperar los datos almacenados?"
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        Screen.MousePointer = vbHourglass
        ModificarRecuperar False
        Screen.MousePointer = vbDefault
        
    Case 2
        'Imprimir
        cmdImprimir_Click
        
        
    Case 3
        'Modificar datos trabajador
        If Me.lw1.ListItems.Count = 0 Then Exit Sub
        If Me.lw1.SelectedItem Is Nothing Then Exit Sub
        
        If lw1.SelectedItem.SubItems(10) <> "0.00" Then
            MsgBox "Quite primero los anticipos", vbExclamation
            Exit Sub
        End If
        
        frmCambiosDatosNomina.Opcion = 0
        Load frmCambiosDatosNomina
        
        VariableCompartida = "" 'Si guarda o no guarda
        With lw1.SelectedItem
            
            Sql = Combo1.ListIndex + 1 & "  / " & Combo1.Text
            frmCambiosDatosNomina.Caption = Sql
            frmCambiosDatosNomina.lblIdTra(0) = .Text             'Trabajador
            frmCambiosDatosNomina.lblTra(0) = " - " & .SubItems(1)              'Trabajador
            i = InStr(.SubItems(2), "/")
            'OFICIALES
            frmCambiosDatosNomina.txtDias(0).Text = Mid(.SubItems(2), 1, i - 1) '
            frmCambiosDatosNomina.txtHN(0).Text = Mid(.SubItems(2), i + 1)
            'TRABAJADAS
            frmCambiosDatosNomina.txtDias(1).Text = .SubItems(3)
            frmCambiosDatosNomina.txtHN(1).Text = .SubItems(4)
            frmCambiosDatosNomina.txtHC(1).Text = .SubItems(5)
            'Nomina
            frmCambiosDatosNomina.txtDias(2).Text = .SubItems(8)
            frmCambiosDatosNomina.txtHN(2).Text = .SubItems(4)
            frmCambiosDatosNomina.txtHC(2).Text = .SubItems(9)
            'Bolsa horas
            frmCambiosDatosNomina.txtBolsa(0).Text = .SubItems(11)
            frmCambiosDatosNomina.txtBolsa(1).Text = .SubItems(12)

            
            frmCambiosDatosNomina.Show vbModal
        End With
        If VariableCompartida <> "" Then
            'HA UPDATEADO LOS DATOS
            PonSQL lw1.SelectedItem.Text
            Set RS = New ADODB.Recordset
            RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            If Not RS.EOF Then
                
                PonLinea lw1.SelectedItem, RS
                lw1.SelectedItem.SmallIcon = 4
            End If
            RS.Close
        End If
        
        
    Case 4
        If lw1.ListItems.Count = 0 Then Exit Sub
        If lw1.SelectedItem Is Nothing Then Exit Sub
        
        
        'Eliminar la entrada
        Sql = ""
        For i = 1 To lw1.ListItems.Count
            If lw1.ListItems(i).Selected Then Sql = Sql & "X"
        Next i
        If Sql = "" Then Exit Sub
        
        
        Sql = "Va a eliminar " & Len(Sql) & " trabajadores de la nomina. Continuar?"
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        
        Sql = ""
        For i = lw1.ListItems.Count To 1 Step -1
            If lw1.ListItems(i).Selected Then
                
                Sql = "DELETE FROM tmpdatosmes where trabajador =" & lw1.ListItems(i).Text
                conn.Execute Sql
                lw1.ListItems.Remove i
            End If
        Next i
        If Sql = "" Then Exit Sub
        
    End Select
End Sub


Private Sub ModificarRecuperar(Guardar As Boolean)

    If Guardar Then
        'Borramos los datos de la 2
        Sql = "Delete from tmpDatosMes2"
        conn.Execute Sql
        
        'Insertamos tmp
        Sql = "INSERT INTo tmpDatosMes2 SELECT * from tmpDatosMES"
        conn.Execute Sql
        
        
        'UPDATEAMOS para guardar el a�o
        'Es decir, en la tabla tmpdatosmes2 habra en lugar del mes solo
        'habra yyyymm
        Sql = "UPDATE tmpDatosMEs2 SET mes=" & Text2.Text & Combo2.ListIndex + 1
        conn.Execute Sql
        
    Else
        'Borramos los datos de la 1
        Sql = "Delete from tmpDatosMes"
        conn.Execute Sql
        
        'truquito
        CadParam = "Error leyendo datos almacenados."
        Sql = DevuelveDesdeBD("mes", "tmpdatosmes2", "mes", "mes", "N")
        If Sql = "" Then
            MsgBox CadParam, vbExclamation
            Exit Sub
        End If
        
        Importe1 = Val(Mid(Sql, 1, 4))
        If Importe1 = 0 Then
            MsgBox CadParam, vbExclamation
            Exit Sub
        End If
        Text2.Text = Importe1
        
        
        Importe1 = Val(Mid(Sql, 5, 2))
        Importe1 = Importe1 - 1
        Combo2.ListIndex = CInt(Importe1)
        'UPDATEAMOS para dejar el mes solamente
        Sql = "UPDATE tmpDatosMEs2 SET mes=" & Importe1 + 1
        conn.Execute Sql
        
        
        'Insertamos tmp
        Sql = "INSERT INTo tmpDatosMes SELECT * from tmpDatosMES2"
        conn.Execute Sql
        
        
        
        'Volvemos a poner el a�o en el dato
        'UPDATEAMOS para guardar el a�o
        'Es decir, en la tabla tmpdatosmes2 habra en lugar del mes solo
        'habra yyyymm
        Sql = "UPDATE tmpDatosMEs2 SET mes=" & Text2.Text & Combo2.ListIndex + 1
        conn.Execute Sql
        
        'Cargamos datos
        CargaDatos
        
        
        
    End If
    
End Sub



Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Combo2.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    PrimeraVez = True

    Frame1.Visible = False
    FrameTipoAlzira.Visible = True 'Opcion = 1
    Me.FrameMes.Visible = True 'Opcion = 0
    'If Opcion = 0 Then
    '    CargaCombo Me.Combo1, Text1
    '    Command1.Enabled = vUsu.Nivel < 2 'Administrador
    'Else
        CargaCombo Me.Combo2, Text2
        cmdGeneraAlz.Enabled = vUsu.Nivel < 2
    'End If
    
    lw1.SmallIcons = Me.ImageList1

    Sql = DevuelveDesdeBD("HorasJornada", "empresas", "idempresa", 1, "N")
    If Sql <> "" Then
        HorasxDia2 = CCur(Sql)
    Else
        HorasxDia2 = 0
    End If
    


    CargaComboSecciones cboSeccion, True
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub CargaCombo(ByRef C As ComboBox, ByRef T As TextBox)
Dim i As Integer
Dim F As Date

    For i = 1 To 12
        C.AddItem Format(CDate("01/" & i & "/2000"), "mmmm")
    Next i
    F = Now
    If Day(Now) < 14 Then F = DateAdd("m", -1, Now)
    C.ListIndex = Month(F) - 1
    T.Text = Year(F)
End Sub


Private Sub CargaColumnas()
Dim Anch As Single
Dim clmX As ColumnHeader


'lw1.ColumnHeaders(17).Width = lw1.Width - 320 - lw1.ColumnHeaders(17).Left

For Each clmX In lw1.ColumnHeaders
    
    If clmX.Index > 3 Then clmX.Alignment = lvwColumnRight
Next

'Las lineas
With lw1


    Line1.X2 = .ColumnHeaders(3).Left - 30 + 160

    Label3.Left = .ColumnHeaders(3).Left + 160
    Line2.X1 = .ColumnHeaders(3).Left + 30 + 160
    Line2.X2 = .ColumnHeaders(4).Left - 30 + 160
    
    Label4.Left = .ColumnHeaders(4).Left + 160
    Line3.X1 = .ColumnHeaders(4).Left + 30 + 160
    Line3.X2 = .ColumnHeaders(8).Left - 30 + 160
    
    Label5.Left = .ColumnHeaders(8).Left + 160
    Line4.X1 = .ColumnHeaders(8).Left + 30 + 160
    Line4.X2 = .ColumnHeaders(10).Left - 30 + 160
    
    Label6.Left = .ColumnHeaders(10).Left + 160
    Line5.X1 = .ColumnHeaders(10).Left + 30 + 160
    Line5.X2 = .ColumnHeaders(13).Left - 30 + 160
    
    .ColumnHeaders(15).Width = 1000
    
    Label7.Left = .ColumnHeaders(13).Left + 160
    Line6.X1 = .ColumnHeaders(13).Left + 30 + 160
    Line6.X2 = .ColumnHeaders(15).Left - 30 + 160
    Line6.Visible = False
    Label7.Visible = False
    'Peque�o reajuste k borda las lineas
    .ColumnHeaders(3).Width = 1000
    
    'La ultima columna a 0
End With
    


End Sub

Private Sub PonLinea(ByRef i As ListItem, ByRef RS As ADODB.Recordset)
'Si tiene dias pendientes
Dim J As Integer
Dim Cantidad1 As Currency
Dim Cantidad2 As Currency
Dim ToolTip As String
    
        ToolTip = ""
        If vEmpresa.NominaAutomatica Then
            'Normal. Pica y cata
            If RS!diasTrabajados = 0 Then
                If RS!MesDias = 0 Then
                    'ESTA DE BAJA
                    J = 3
                    ToolTip = "De baja"
                Else
                    J = 3 '10
                    ToolTip = ""
                End If
            Else
                If RS!ControlNomina = 1 Then
                    'Normal
                    J = 1
                Else
                    If RS!ControlNomina = 3 Then
                        'Jorandas semanas
                        J = 8
                        ToolTip = "Jornadas semanales"
                    Else
                        'Tipo de liquidaciones
                        J = 6
                        ToolTip = "Liquidaciones"
                    End If
                End If
                If RS!saldodias <> 0 Then J = J + 1
                
            End If
        
        Else
            'Como alzira
            If RS!diasTrabajados = 0 Then
                'J = 3
                J = 7
            Else
                If RS!extras <> 0 Then
                    J = 4
                    ToolTip = "Extras"
                Else
                    If RS!saldodias > 0 Then
                        'Ya ha compensado
                        J = 2
                        ToolTip = "Puede compensar"
                    Else
                        J = 1
                        ToolTip = "Ya compensado"
                    End If
                End If
            End If
        End If
        
        If RS!importefijo > 0 Then
            J = 8  'importe fijo
            ToolTip = "Importe fijo mes"
        End If
        
        If J = 7 Then ToolTip = "Revisar"
        
        i.SmallIcon = J
        i.Text = RS!Trabajador
        i.SubItems(1) = RS!nomtrabajador
        i.ToolTipText = ToolTip
        
        'Horas oficiles
        i.SubItems(2) = RS!MesDias & "/" & Format(RS!meshoras, "0.00")
        
        'Trabajados
        
        'If RS!horasn > 0 Then S top
        
        i.SubItems(3) = RS!diasTrabajados
        i.SubItems(4) = Format(RS!horasn, "0.00")
        i.SubItems(5) = Format(RS!HorasC, "0.00")
        i.SubItems(6) = Format(RS!horase, "0.00")
        
        'Saldo
        i.SubItems(7) = RS!saldodias
        
        Cantidad1 = RS!saldoh
        If Cantidad1 < 0 Then
            Cantidad1 = 0
            'Veremos si ha utilizado bolsa de horas, si no, pintaremos cero igualmente
            Cantidad2 = RS!bolsaantes - RS!bolsadespues
            If Cantidad2 < 0 Then
                MsgBox "Debe horas y aunmenta bolsa. Comprobar trabajador " & RS!nomtrabajador, vbExclamation
            Else
                Cantidad1 = -Cantidad2
            End If
        End If
        i.SubItems(8) = Format(Cantidad1, "0.00")
        
   
        i.SubItems(9) = RS!diasperiodo
        i.SubItems(10) = " "  'Horas que lleva a nomina son las horasn
        i.SubItems(11) = " "  'EXTRAS
        If RS!HorasC > 0 Then i.SubItems(11) = Format(RS!horase, "0.00")
        
        
        '
        'Bolsa
        i.SubItems(12) = Format(RS!bolsaantes, "0.00")
        i.SubItems(13) = Format(RS!bolsadespues, "0.00")
        
        
        'Bruto NORMALES
        i.SubItems(14) = Format(DBLet(RS!Bruto, "N"), "0.00")
        
        
        i.SubItems(15) = " "
        If RS!ImporEstruc > 0 Then i.SubItems(15) = Format(RS!ImporEstruc, "0.00")
        
        
        i.SubItems(16) = " "
        Importe1 = 0
        If RS!extras > 0 Then Importe1 = RS!extras
        If DBLet(RS!LlevaPlus, "N") <> 0 Then Importe1 = Importe1 + RS!LlevaPlus
        If Importe1 <> 0 Then
            i.SubItems(16) = Format(Importe1, "0.00")
            If DBLet(RS!LlevaPlus, "N") <> 0 Then
                i.ListSubItems(16).ForeColor = vbBlue
                i.ListSubItems(16).ToolTipText = "Lleva plus horas"
            End If
        End If
        Importe1 = RS!Bruto
        If RS!Bruto > 0 Then
            
            'Importe1 = RS!Bruto * RS!d
            If DBLet(RS!PorcAntiguedad, "N") > 0 Then
                Importe1 = RS!Bruto * RS!PorcAntiguedad
                Importe1 = Round((Importe1 / 100) + RS!Bruto, 2)
            End If
        End If
        
        
        
        'Neto
        Importe1 = Importe1 + RS!extras + RS!ImporEstruc
        i.SubItems(17) = " "
        If Importe1 <> 0 Then
            i.SubItems(17) = Format(Importe1, "0.00")
            If DBLet(RS!PorcAntiguedad, "N") > 0 Then
                i.ListSubItems(17).ForeColor = vbBlue
                i.ListSubItems(17).ToolTipText = "Lleva antiguedad"
            End If
        End If
        i.Tag = RS!ControlNomina
End Sub


Private Sub CargaDatos()
Dim i As ListItem
Dim RS As ADODB.Recordset
Dim NParam As Byte

    Set RS = New ADODB.Recordset
    lw1.ListItems.Clear
    PonSQL ""
    Sql = Sql & " order by "

    NParam = 1
    If vEmpresa.NominaAutomatica Then
        If Option1(0).Value Then NParam = 0
    Else
        If Option2(0).Value Then NParam = 0
    End If
    
    
        If NParam = 0 Then
            Sql = Sql & "id"
        Else
            Sql = Sql & "nom"
        End If
        NParam = 0
    
    Sql = Sql & "Trabajador"
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        Set i = lw1.ListItems.Add
        'If RS!Trabajador = 9006 Then St op
        PonLinea i, RS
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
        
End Sub




Private Sub Form_Resize()
Dim H As Single
Dim W As Single

    If Me.WindowState = 1 Then Exit Sub
    If Me.Width < 7000 Then
        W = 7000
        Me.Width = W
    Else
        W = Me.Width
    End If
    If Me.Height < 3900 Then
        H = 3900
        Me.Height = H
    Else
        H = Me.Height
    End If
    Me.lw1.Width = W - lw1.Left - 210
    Me.lw1.Height = H - lw1.Top - 500
    CargaColumnas
End Sub

'Private Sub ListView1_Click()
'Dim i
'    SQL = ""
'    For i = 1 To ListView1.ColumnHeaders.Count
'        SQL = SQL & ListView1.ColumnHeaders(i).Text & ": " & ListView1.ColumnHeaders(i).Width & vbCrLf
'    Next i
'    MsgBox SQL
'End Sub


Private Sub PonerBaja(Baja As Boolean)
Dim Importe As Currency
Dim Imp1 As Currency
Dim RS As ADODB.Recordset


    Stop

    If Baja Then
    
    
        'Cuando ponemos la baja calculamos si tiene horas en bolsa despues.
        'las tranformamos en euros de mas en anticpos
        Importe = ImporteFormateadoAmoneda(lw1.SelectedItem.SubItems(12))
        
        If Importe > 0 Then
            Sql = "SELECT Categorias.Importe1, Categorias.Importe2, Trabajadores.IdTrabajador,PorcSS,PorcIRPF"
            Sql = Sql & " FROM Categorias INNER JOIN Trabajadores ON Categorias.IdCategoria = Trabajadores.idCategoria"
            Sql = Sql & " WHERE Trabajadores.IdTrabajador=" & lw1.SelectedItem.Text

            Set RS = New ADODB.Recordset
            RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If RS.EOF Then
                MsgBox "Error leyendo datos trabajador", vbExclamation
            Else
                Importe = Importe * RS.Fields(0) 'importe2
                Imp1 = (Importe * RS!PorcSS) + (Importe * RS!PorcIRPF)
                Imp1 = Imp1 / 100
                Importe = Importe + Imp1
                Importe = Round(Importe, 2)
                'PLUS
                lw1.SelectedItem.SubItems(14) = Format(Importe, FormatoImporte)
               
               
                'Importe origninal
                Imp1 = ImporteFormateadoAmoneda(lw1.SelectedItem.SubItems(13))
                Importe = Importe + Imp1
                lw1.SelectedItem.SubItems(12) = "0.00" 'Le quitamos la bolsa
                lw1.SelectedItem.SubItems(13) = Format(Importe, FormatoImporte)

            End If
            RS.Close
        End If
    Else
        'Reestablecemos los valores de tmpDatosmes
        PonSQL lw1.SelectedItem.Text
        
        Set RS = New ADODB.Recordset
        RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RS.EOF Then
            MsgBox "Error leyendo datos tmpDatosMES del trabajador : " & lw1.SelectedItem.Text, vbExclamation
        Else
            PonLinea lw1.SelectedItem, RS
        End If
        RS.Close
    End If
    Set RS = Nothing
End Sub


Private Sub PonSQL(Id As String)
    Sql = "Select tmpDatosMes.*,nomtrabajador,controlnomina,PorcAntiguedad,ImporteFijoNomina from tmpDatosMes,Trabajadores"
    Sql = Sql & " WHERE tmpDatosMes.trabajador = Trabajadores.idTrabajador "
    If Id <> "" Then Sql = Sql & " AND tmpDatosMes.trabajador =" & Id
End Sub





Private Function PuedeCompensarDias() As Integer
Dim i As Integer

    PuedeCompensarDias = 0
    Sql = DevuelveDesdeBD("idHorario", "Trabajadores", "idTrabajador", lw1.SelectedItem.Text, "N")
    i = Val(Sql)
    
    'En la tabla tmpHorasMesHorario, al cargar los datos
    'se han cargado las horas oficiales
    Sql = DevuelveDesdeBD("Dias", "tmpHorasMesHorario", "idHorario", CStr(i), "N")
    If Sql <> "" Then
        i = Val(Sql)
        i = i - Val(lw1.SelectedItem.SubItems(8))
        If i > 0 Then PuedeCompensarDias = i
    End If
    
    
    
End Function


Private Sub CompensarDias(Dias As Integer)
Dim i As Integer
Dim Lab As Integer
Dim H As Currency
Dim H1 As Currency
Dim D1 As Integer
Dim RS As ADODB.Recordset
    

    Sql = DevuelveDesdeBD("idHorario", "Trabajadores", "idTrabajador", lw1.SelectedItem.Text, "N")
    i = Val(Sql)

    Lab = DiasLaborablesSemana(i)
    If Lab < 1 Then Exit Sub

    If Dias < Lab Then
        'Nos salimos pq no tengo bastantes dias para compensar un semana
        Exit Sub
    End If

    

    'QUiero saber las horas a la semana k puedo compensar
    Set RS = New ADODB.Recordset
    Sql = "Select * from Horarios Where idHorario =" & i
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
    
        H = CCur(lw1.SelectedItem.SubItems(11)) 'Las horas k le van a quedar en bolsa
        'Ya tengo el horario y los dias a compensar
        
        'A compensar
        D1 = 0 'Dias
        H1 = 0 'Horas
        
        'Por lo tanto veo cuantas semanas mas voy a compensar
        Do
            i = Dias \ Lab
            If i > 0 Then
                'Una semana seguro k puedo compensar. Vamos palla
                If H >= RS!TotalHoras Then   'Horas semana
                    D1 = D1 + Lab
                    H1 = H1 + RS!TotalHoras
                    H = H - RS!TotalHoras
                End If
                Dias = Dias - Lab
            End If
        Loop Until i = 0
    End If
    RS.Close
    
    
    'Si a compensado lo reflejo en la listview
    If D1 > 0 Then
        'Dias nomina
        
        
        'Horas para la nomina
        Sql = "Select Importe1,importe2,porcSS,porcIRPF from Categorias,Trabajadores WHERE Trabajadores.IdCategoria = Categorias.IdCategoria"
        Sql = Sql & " AND Trabajadores.idTrabajador =" & lw1.SelectedItem.Text
        RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            'Ponemos ya las nuevas horas en horas normales
            H = CCur(lw1.SelectedItem.SubItems(4)) + H1
            lw1.SelectedItem.SubItems(4) = Format(H, FormatoImporte)
            'Bolsa
            H = CCur(lw1.SelectedItem.SubItems(11)) - H1
            lw1.SelectedItem.SubItems(11) = Format(H, FormatoImporte)
            
            'Precio bruto
            H = H1 * RS!Importe1
            
            'Precio neto
            H1 = ((H * RS!PorcSS) + (H * RS!PorcIRPF)) / 100
            
            H = Round(H - H1, 2)
            'Anticipos
            H1 = ImporteFormateadoAmoneda(lw1.SelectedItem.SubItems(12))
            H = H + H1
            lw1.SelectedItem.SubItems(12) = H
            
            'Dias nomina
            i = Val(lw1.SelectedItem.SubItems(8)) + D1
            lw1.SelectedItem.SubItems(8) = i
        End If
    End If
    Set RS = Nothing
End Sub



Private Sub Form_Unload(Cancel As Integer)
    
    'ALZIRA
    If Me.lw1.ListItems.Count > 0 Then
        If MsgBox("Seguro que desea salir de la edici�n de Nominas?", vbQuestion + vbYesNo) = vbNo Then Cancel = 1
    End If
End Sub



Private Sub lw1_DblClick()
Dim vH As CHorarios
Dim F As Date
Dim F2 As Date
Dim MEDIOS As String

    With lw1.SelectedItem
        If lw1.SelectedItem Is Nothing Then Exit Sub
        
        Sql = DevuelveDesdeBD("idcal", "Trabajadores", "idTrabajador", .Text, "N")
      
        Sql = "fecha>=" & DBSet(CDate("01/" & Combo2.ListIndex + 1 & "/" & Text2.Text), "F") & " AND idcal = " & Sql & " AND 1"
        Sql = DevuelveDesdeBD("idHorario", "calendariol", Sql, "1", "N")
        If Sql = "" Then
            MsgBox "Error leyendo datos trabajador", vbExclamation
        Else
                
            
            Set vH = New CHorarios
            vH.IdHorario = Val(Sql)
            Sql = ""
           
            F = CDate("01/" & Combo2.ListIndex + 1 & "/" & Text2.Text)
            F2 = F
            F = DateAdd("m", 1, F)
            F = DateAdd("d", -1, F)
            Sql = ""
            'Picassent
            'MEDIOS = vH.LeerMediosDias(vH.IdHorario, F2, F)
            
            Sql = vH.LeerDiasFestivos(vH.IdHorario, F2, F)
            frmVerDiasMesTrabajador3.DiasEnNomina = .SubItems(8)
            frmVerDiasMesTrabajador3.TodoElMEs = 0
            frmVerDiasMesTrabajador3.JornadasSemanales = False '(.Tag = 3)
            frmVerDiasMesTrabajador3.MediosDias = MEDIOS
            frmVerDiasMesTrabajador3.FESTIVOS = Sql
            frmVerDiasMesTrabajador3.Trabajador = .SubItems(1) & "|" & .Text & "|"
            frmVerDiasMesTrabajador3.FechaIni = F2
            frmVerDiasMesTrabajador3.HorasMinimoDia = HorasxDia2
            frmVerDiasMesTrabajador3.Show vbModal
            Set vH = Nothing
        End If
    End With
End Sub




Private Sub lw1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If Not vEmpresa.NominaAutomatica Then
            PopupMenu mnPopup
        End If
    End If
End Sub




Private Sub mnVerDatos_Click()
    lw1_DblClick
End Sub

Private Sub Text1_GotFocus()
    With Text1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Text1_LostFocus()
    Text1.Text = Trim(Text1.Text)
    If Text1.Text <> "" Then
        If Not IsNumeric(Text1.Text) Then
            MsgBox "A�o debe ser num�rico. (" & Text1.Text & ")", vbExclamation
            Text1.Text = ""
        End If
    End If
    If Text1.Text = "" Then Text1.Text = Year(Now)
End Sub




Private Sub InsertarEnTemporalTrabajador(ByRef dSQL As String, idTrabajador As Long)
    On Error Resume Next
    conn.Execute dSQL
    If Err.Number <> 0 Then
        dSQL = "Error insertando el trabajador : " & idTrabajador & " . Entrada duplicada"
        MsgBox dSQL, vbExclamation
    End If
End Sub











Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub





Private Function genNominas() As Boolean
Dim i As Integer
Dim cad As String
Dim Importe As Currency
Dim Horas As Currency
Dim RS As ADODB.Recordset
Dim Aux As String
Dim PlusHorasTr As Currency
On Error GoTo EGenerarNominas
    genNominas = False

    Sql = "INSERT INTO Nominas (Fecha,IdTrabajador,Dias,HN,HC,HP,BolsaDespues,BolsaAntes"
    Sql = Sql & ",Anticipos,Antiguedad,IRPF,SSEmpr,PrecioHN,PrecioHC,PrecioHE,LlevaPlus , ImporteFijo"
    Sql = Sql & ",Bruto,ImporEstruc,Plus,FechaAlta,CentroA3 ) VALUES ('"
    i = DiasMes(Combo2.ListIndex + 1, CInt(Text2.Text))
    Sql = Sql & Text2.Text & "-" & Combo2.ListIndex + 1 & "-" & i & "',"
    
    'Primero generamos la tabla de  nominas con los importes marcados aqui
    cad = "SELECT tmpDatosMEs.*, "
    cad = cad & " Trabajadores.PorcSS, Trabajadores.PorcIRPF,Trabajadores.PorcAntiguedad,Importe1,Importe2,Importe3 "
    cad = cad & " ,Trabajadores.FecAlta ,Trabajadores.idCentroA3"
    cad = cad & " FROM tmpDatosMEs , Trabajadores ,categorias WHERE  tmpDatosMEs.Trabajador = Trabajadores.IdTrabajador"
    cad = cad & " AND trabajadores.idcategoria=Categorias.idcategoria "
    cad = cad & " AND Mes = " & Combo2.ListIndex + 1
    cad = cad & " ORDER BY tmpDatosMes.Trabajador"
    Set RS = New ADODB.Recordset
    RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
      '  If RS!HorasC > 0 Then Stop
      '  If RS!extras > 0 Then Stop
    
        'IdTrabajador,Dias
        cad = RS!Trabajador & "," & RS!diasperiodo & ","
        
        'HN,HC,HP   -> Las horas compensables seran aquellas que superen las horas de las horas mensuales, las que iran a bolsa, pero ahora se pagan
        Horas = RS!horase
        cad = cad & TransformaComasPuntos(RS!horasn) & "," & TransformaComasPuntos(DBLet(RS!HorasC, "N")) & "," & DBSet(Horas, "N") & ","
        


        
        'BolsaDespues,BolsaAntes,brutodespues,netodespues,importedelbote,brutoantes,netoan
        cad = cad & DBSet(RS!bolsadespues, "N") & "," & DBSet(RS!bolsaantes, "N") & ","
        
        'Ahora nO anticipmaos como hacia Picassent
        Importe = 0
        cad = cad & TransformaComasPuntos(DBLet(Importe, "N"))
        
        'Antiguedad,IRPF,SSEmpr
        cad = cad & "," & DBSet(RS!PorcAntiguedad, "N") & "," & DBSet(RS!PorcIRPF, "N") & "," & DBSet(RS!PorcSS, "N")
        'PrecioHN,PrecioHC,PrecioHE
        cad = cad & "," & DBSet(RS!Importe1, "N") & "," & DBSet(RS!Importe2, "N") & "," & DBSet(RS!Importe3, "N")
        
        'LlevaPlus , ImporteFijo
        cad = cad & "," & DBSet(RS!LlevaPlus, "N") & "," & DBSet(RS!importefijo, "N")
        
        ',Bruto,ImporEstruc,Plus
        cad = cad & "," & DBSet(RS!Bruto, "N") & "," & DBSet(RS!ImporEstruc, "N") & "," & DBSet(RS!extras, "N")
        
        'cad = cad & " ,Trabajadores.FecAlta ,Trabajadores.idCentroA3"
        cad = cad & "," & DBSet(RS!FecAlta, "F", "S") & "," & DBSet(RS!idCentroA3, "N", "S")
        
        cad = cad & ")"
        cad = Sql & cad
        conn.Execute cad
        
        
        
        
        'Pondremos la bolsa de horas Y, hay bajas,
        'entonces actualizaremos la baja de cada trabajador
        'al ultimo dia trabajado
        
        
        
        cad = "REPLACE INTO trabajadoresbolsahoras(IdTrabajador,ParaEmpresa,TipoHora,HorasBolsa) VALUES ("
        cad = cad & RS!Trabajador & ",0,1," & DBSet(RS!bolsadespues, "N") & ")"
        conn.Execute cad
        
        
        

        
        'Sig
        RS.MoveNext
    Wend
    

    
    RS.Close
    
    
    
    
    genNominas = True
    Exit Function
EGenerarNominas:
    MuestraError Err.Number, Err.Description
    Set RS = Nothing
End Function





