VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVerDiasMesTrabajador3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   Icon            =   "frmVerDiasMesTrabajador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   6360
      TabIndex        =   46
      Text            =   "0"
      Top             =   7440
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Ajuste semanal"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   45
      Top             =   8280
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Real"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   44
      Top             =   8280
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5400
      TabIndex        =   43
      Text            =   "0"
      Top             =   7440
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   4440
      TabIndex        =   42
      Text            =   "0"
      Top             =   7440
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   12
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   39
      Tag             =   "0"
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   11
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   37
      Tag             =   "0"
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   10
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   35
      Tag             =   "0"
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "Text3"
      Top             =   6840
      Width           =   4815
   End
   Begin VB.Frame FrameJornadasSemanales 
      Height          =   3015
      Left            =   9480
      TabIndex        =   19
      Top             =   600
      Width           =   4935
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2640
         TabIndex        =   25
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   3000
         TabIndex        =   24
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   23
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   22
         Top             =   2640
         Width           =   375
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   2175
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   2382
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "DO"
            Object.Width           =   688
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "HO"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "DT"
            Object.Width           =   688
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "HT"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ANT"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "POS"
            Object.Width           =   1058
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Resumen jornadas del trabajador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame FrameMedias 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4440
      TabIndex        =   28
      Top             =   2640
      Width           =   5055
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   31
         Tag             =   "0"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   29
         Tag             =   "0"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Sabado"
         Height          =   195
         Index           =   11
         Left            =   2520
         TabIndex        =   32
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Miercoles"
         Height          =   195
         Index           =   10
         Left            =   0
         TabIndex        =   30
         Top             =   480
         Width           =   1005
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   0
         X2              =   4440
         Y1              =   240
         Y2              =   240
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   6
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   17
      Tag             =   "0"
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   13
      Tag             =   "0"
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   11
      Tag             =   "0"
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   9
      Tag             =   "0"
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   8
      Tag             =   "0"
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   495
      Left            =   8400
      TabIndex        =   6
      Top             =   7320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   5
      Tag             =   "0"
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   3
      Tag             =   "0"
      Top             =   1200
      Width           =   855
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   5640
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
            Picture         =   "frmVerDiasMesTrabajador.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVerDiasMesTrabajador.frx":0D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVerDiasMesTrabajador.frx":172E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVerDiasMesTrabajador.frx":1B80
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVerDiasMesTrabajador.frx":1FD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVerDiasMesTrabajador.frx":22EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVerDiasMesTrabajador.frx":2886
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVerDiasMesTrabajador.frx":2BA0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   14208
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "HN"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "HC"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Total"
         Object.Width           =   1411
      EndProperty
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1215
      Left            =   4560
      TabIndex        =   15
      Top             =   5280
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2143
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "BAJA"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ALTA"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descripcion"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   7
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   26
      Tag             =   "0"
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Sumas directas"
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
      Index           =   16
      Left            =   4440
      TabIndex        =   41
      Top             =   7200
      Width           =   1305
   End
   Begin VB.Label Label2 
      Caption         =   "H Dia baja"
      Height          =   255
      Index           =   15
      Left            =   6120
      TabIndex        =   40
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "H Min"
      Height          =   255
      Index           =   14
      Left            =   6240
      TabIndex        =   38
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Comp XyS"
      Height          =   255
      Index           =   13
      Left            =   7440
      TabIndex        =   36
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Ticajes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   4440
      TabIndex        =   34
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Dias en NOMINA"
      Height          =   195
      Index           =   9
      Left            =   4440
      TabIndex        =   27
      Top             =   2280
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   1335
      Left            =   4440
      Top             =   3600
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   "Dias para nomina"
      Height          =   195
      Index           =   7
      Left            =   4440
      TabIndex        =   18
      Top             =   1800
      Width           =   1245
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   4440
      X2              =   8880
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label2 
      Caption         =   "Bajas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   4560
      TabIndex        =   16
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Festivos"
      Height          =   255
      Index           =   5
      Left            =   6960
      TabIndex        =   14
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Total"
      Height          =   255
      Index           =   4
      Left            =   8520
      TabIndex        =   12
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "H.E"
      Height          =   255
      Index           =   3
      Left            =   5400
      TabIndex        =   10
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Sin revisar"
      Height          =   255
      Index           =   2
      Left            =   6960
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "H.N"
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   4
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Trabajados"
      Height          =   255
      Index           =   0
      Left            =   4440
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmVerDiasMesTrabajador3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Trabajador As String  'con recupera los ponemos RecuperaValor
Public FechaIni As Date
Public FESTIVOS As String
Public MediosDias As String
Public JornadasSemanales As Boolean
Public TodoElMEs As Integer
Public DiasEnNomina As Integer
Public HorasCompensablesMiercolesSabado As Currency
Public HorasMinimoDia As Currency

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
        Text4(2).Visible = False
        If vEmpresa.QueEmpresa = vbCatadau Then
            Text4(2).Visible = True
            ListView1.ColumnHeaders(4).Text = "extra"
            
        End If
        Text3.Text = ""
        
        
        
        CargaGrid
        Label1.Caption = RecuperaValor(Trabajador, 1)
        Caption = "Datos: " & Label1.Caption
        Me.FrameJornadasSemanales.Visible = JornadasSemanales
        Shape1.Visible = JornadasSemanales
        FrameMedias.Visible = Not JornadasSemanales And MediosDias <> ""
        FrameJornadasSemanales.Left = 4440
        Me.Command1.Cancel = True
End Sub



Private Sub CargaGrid()
Dim cad As String
Dim RS As ADODB.Recordset
Dim itmX As ListItem
Dim i As Integer
Dim Importe As Currency
Dim Icono As Integer
Dim FFin As Date
Dim Dias As Currency
Dim Semana As Integer
Dim ContadorMier As Byte
Dim ContadorSab As Byte
Dim k As Integer
Dim H As Currency

Dim IdInci As Integer
Dim ExcesoDefecto As Boolean

    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    If JornadasSemanales Then ListView3.ListItems.Clear
    
    
    If Me.Option1(0).Value Then
        'Coge los datos de marcajes
           
        cad = "Select *,Incidencias.ExcesoDefecto"
        cad = cad & " FROM Marcajes INNER JOIN Incidencias ON Marcajes.IncFinal = Incidencias.IdInci    "
        cad = cad & " Where idTrabajador = " & RecuperaValor(Trabajador, 2)
        cad = cad & " AND Fecha >= " & DBSet(FechaIni, "F")
        If TodoElMEs = 0 Then
            FFin = DateAdd("m", 1, FechaIni)
            FFin = DateAdd("d", -1, FFin)
        Else
            'Es una semana
            FFin = DateAdd("d", TodoElMEs, FechaIni)
        End If
        cad = cad & " AND Fecha <= " & DBSet(FFin, "F")
        cad = cad & " ORDER By Fecha"
    
    
    Else
        'Semanas procesadas. Alzira y coopic
        If vEmpresa.QueEmpresa = vbCatadau Then
            
            cad = "select fecha,sum(if (tipohoras=0,horastrabajadas,0)) norm"
            cad = cad & ",sum(if (tipohoras=1,horastrabajadas,0)) estr"
            cad = cad & ",sum(if (tipohoras=2,horastrabajadas,0)) extr"
            cad = cad & " "
            cad = cad & " , 1 Correcto From jornadassemanalesalz"
            
        
        Else
        
            cad = "select fecha,sum(if (tipohoras<2,horastrabajadas,0)) norm"
            cad = cad & ",sum(if (tipohoras=2,horastrabajadas,0)) extr"
            cad = cad & " "
            cad = cad & " , 1 Correcto From jornadassemanalesalz"
        
        End If
        cad = cad & " Where idTrabajador = " & RecuperaValor(Trabajador, 2)
        cad = cad & " AND Fecha >= " & DBSet(FechaIni, "F")
        If TodoElMEs = 0 Then
            FFin = DateAdd("m", 1, FechaIni)
            FFin = DateAdd("d", -1, FFin)
        Else
            'Es una semana
            FFin = DateAdd("d", TodoElMEs, FechaIni)
        End If
        cad = cad & " AND Fecha <= " & DBSet(FFin, "F")
        cad = cad & " group by 1"
        cad = cad & " ORDER By Fecha"
        
    End If
    
    
    Set RS = New ADODB.Recordset
 
    RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'a 0
    For i = 0 To Me.Text1.Count - 1
        Text1(i).Tag = 0
    Next i

    
    Semana = 0
    While Not RS.EOF
        'Dias trbajados
        Text1(0).Tag = Text1(0).Tag + 1
        Set itmX = ListView1.ListItems.Add
        itmX.ToolTipText = ""
        cad = Format(RS!Fecha, "dd/mm/yyyy")
        i = Val(Format(RS!Fecha, "ww"))
        
        If cad = "26/12/2019" Then
            Debug.Print ""
        End If
        
        If i <> Semana Then
            itmX.ForeColor = &H6A6C6A
            Semana = i
        End If
        'Medios dias o para nomina
        Dias = 1
        If InStr(1, FESTIVOS, cad) = 0 Then
            'Si no es festivo
            If InStr(1, MediosDias, cad) > 0 Then
                
                
                Dias = Weekday(CDate(cad), vbMonday)
                If Dias = 3 Then
                    ContadorMier = ContadorMier + 1
                Else
                    If Dias = 6 Then ContadorSab = ContadorSab + 1
                End If
                Dias = 0.5
                cad = cad & " *"
            End If
        Else
            Dias = 0
        End If
        Text1(6).Tag = CCur(Text1(6).Tag) + Dias

        itmX.Text = cad
        
        If RS!Correcto = 1 Then
            
           
            If Me.Option1(0).Value Then
                IdInci = RS!IdInci
                ExcesoDefecto = RS!ExcesoDefecto
            Else
                'Compensado
                
                If RS!extr > 0 Then
                    IdInci = 1
                    ExcesoDefecto = True
                Else
                    IdInci = 2
                    If RS!norm > vEmpresa.HorasJornada Then
                        ExcesoDefecto = True
                    Else
                        ExcesoDefecto = False
                    End If
                End If
            End If
                
            If IdInci = 0 Then
                'Normal
                Icono = 1
                itmX.SubItems(1) = Format(RS!HorasTrabajadas, FormatoImporte)
                
            Else

                If ExcesoDefecto Then
                
                    If Me.Option1(0).Value Then
                        Importe = RS!HorasTrabajadas - RS!HorasIncid
                        itmX.SubItems(1) = Format(Importe, FormatoImporte)
                        itmX.SubItems(2) = Format(RS!HorasIncid, FormatoImporte)
                        
                        
                    Else
                        If vEmpresa.QueEmpresa = 4 Then
                            'CATADAU
                            itmX.SubItems(1) = Format(RS!norm, FormatoImporte)
                            itmX.SubItems(2) = Format(RS!estr, FormatoImporte)
                            If IdInci = 1 Then
                                'Tiene horas extras
                                itmX.SubItems(3) = Format(RS!extr, FormatoImporte)
                            Else
                            
                                'Importe = RS!norm - vEmpresa.HorasJornada
                                'itmX.SubItems(1) = Format(vEmpresa.HorasJornada, FormatoImporte)
                                'itmX.SubItems(2) = Format(Importe, FormatoImporte)
                                Importe = 0
                                itmX.SubItems(3) = " "
                            End If
                            
                            
                            
                        Else
                            'COOPIC ALZIRA
                            If IdInci = 1 Then
                                'Tiene horas extras
                                itmX.SubItems(1) = Format(RS!norm, FormatoImporte)
                                itmX.SubItems(2) = Format(RS!extr, FormatoImporte)
                            Else
                            
                                'Importe = RS!norm - vEmpresa.HorasJornada
                                'itmX.SubItems(1) = Format(vEmpresa.HorasJornada, FormatoImporte)
                                'itmX.SubItems(2) = Format(Importe, FormatoImporte)
                                
                                Importe = 0
                                itmX.SubItems(1) = Format(RS!norm, FormatoImporte)
                                itmX.SubItems(2) = Format(Importe, FormatoImporte)
                                
                            End If
                        End If
                    End If
                    Icono = 1
                    Text1(3).Tag = Text1(3).Tag + ImporteFormateadoAmoneda(CStr(itmX.SubItems(2)))
                Else
                    Icono = 2
                    If Me.Option1(0).Value Then
                        Importe = RS!HorasTrabajadas
                    Else
                        Icono = 8
                        If vEmpresa.QueEmpresa = 4 Then
                            'CATADAU
                            itmX.SubItems(1) = Format(RS!norm, FormatoImporte)
                            itmX.SubItems(2) = Format(RS!estr, FormatoImporte)
                            If IdInci = 1 Then
                                'Tiene horas extras
                                itmX.SubItems(3) = Format(RS!extr, FormatoImporte)
                            Else
                            
                                'Importe = RS!norm - vEmpresa.HorasJornada
                                'itmX.SubItems(1) = Format(vEmpresa.HorasJornada, FormatoImporte)
                                'itmX.SubItems(2) = Format(Importe, FormatoImporte)
                                Importe = 0
                                itmX.SubItems(3) = " "
                            End If

                        Else
                            'Lo que habia
                    
                    
                            Importe = RS!norm
                    
                            itmX.SubItems(1) = Format(Importe, FormatoImporte)
                            If Me.HorasMinimoDia > 0 Then
                                                    
                            
                                                        
                                'Esto era para picassent
                                'k = Weekday(RS!Fecha, vbSunday)
                                'If k <> 3 And k <> 6 Then
                                '    'Miercoles SABADO
                                '
                                'Else
                                '    If RS!HorasTrabajadas < Me.HorasMinimoDia Then
                                '
                                '        Icono = 8
                                '        Text1(11).Tag = CCur(Text1(11).Tag) + Me.HorasMinimoDia - RS!HorasTrabajadas
                                '    End If
                                'End If
                            
                            End If
                        End If
                    End If
                End If
                
                If InStr(1, FESTIVOS, cad) > 0 Then
                    Icono = 5
                    Text1(5).Tag = Text1(5).Tag + 1
                End If
                
                
            End If
        Else
            'Incorrectos
            itmX.SubItems(1) = Format(RS!HorasTrabajadas, FormatoImporte)
            Text1(1).Tag = Text1(1).Tag + 1
            Icono = 4
        End If
        itmX.SmallIcon = Icono
        
        
        Text1(2).Tag = Text1(2).Tag + ImporteFormateadoAmoneda(CStr(itmX.SubItems(1)))
        
         
        If Option1(0).Value Then
            H = RS!HorasTrabajadas
            'Totales
            Text1(4).Tag = Text1(4).Tag + H
        
        
        Else
        
            H = RS!norm + RS!extr
        
        
            'Totales
            Text1(4).Tag = Text1(4).Tag + H
        
        End If
        
        
        RS.MoveNext
        
    Wend
    
    RS.Close
    
    
    
    
            cad = "SELECT Bajas.*, tipobaja.descbaja"
            cad = cad & " FROM tipobaja INNER JOIN Bajas ON tipobaja.idbaja = Bajas.idTipobaja"
            cad = cad & " WHERE Bajas.IdTrab = " & RecuperaValor(Trabajador, 2)
            Me.Tag = cad   'Este trozo sera comun para el resto de SQLs
            
            cad = cad & " AND Fechabaja<=" & DBSet(FFin, "F", "F")
            cad = cad & " AND FechaAlta is null ORDER BY fechabaja"
            RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RS.EOF
                'Añadimos en el list2
                Set itmX = ListView2.ListItems.Add
                cad = Format(RS!FechaBaja, "dd/mm/yyyy")
                itmX.Text = cad
                cad = ""
                itmX.SubItems(1) = cad
                itmX.SubItems(2) = RS!descbaja
                
                itmX.SmallIcon = 6
                
                'falta
                For i = 1 To ListView1.ListItems.Count
                    If ListView1.ListItems(i).Text = itmX.Text Then
                        'ESTE ES EL DIA DE LA BESTIA, o de la baja ;)
                        
                        Importe = 0
                        If ListView1.ListItems(i).SubItems(1) <> "" Then Importe = Importe + ImporteFormateadoAmoneda(CStr(ListView1.ListItems(i).SubItems(1)))
                        If ListView1.ListItems(i).SubItems(2) <> "" Then Importe = Importe + ImporteFormateadoAmoneda(CStr(ListView1.ListItems(i).SubItems(2)))
                        
                        Text1(2).Tag = CCur(Text1(2).Tag) - Importe
                        Text1(3).Tag = CCur(Text1(3).Tag) + Importe
                        Text1(12).Tag = Text1(12).Tag + Importe
                        ListView1.ListItems(i).SubItems(1) = " "
                        ListView1.ListItems(i).SubItems(2) = Format(Importe, FormatoImporte)
                        ListView1.ListItems(i).SmallIcon = 6 'baja
                    End If
                    
                Next
                
                
                RS.MoveNext
            Wend
            RS.Close
            
    
    
   
    
    
   
        cad = Me.Tag
        cad = cad & " AND FechaAlta>=" & DBSet(FechaIni, "F")
        cad = cad & " AND FechaAlta<=" & DBSet(FFin, "F")
        RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            'Añadimos en el list2
            Set itmX = ListView2.ListItems.Add
            cad = Format(RS!FechaBaja, "dd/mm/yyyy")
            itmX.Text = cad
            cad = Format(RS!fechaalta, "dd/mm/yyyy")
            itmX.SubItems(1) = cad
            itmX.SubItems(2) = RS!descbaja
            itmX.SmallIcon = 6
            
            
            
            
                        'falta
                For i = 1 To ListView1.ListItems.Count
                    If Mid(ListView1.ListItems(i).Text, 1, 10) = itmX.Text Then
                        'ESTE ES EL DIA DE LA BESTIA, o de la baja ;)
                        
                        Importe = 0
                        If ListView1.ListItems(i).SubItems(1) <> "" Then Importe = Importe + ImporteFormateadoAmoneda(CStr(ListView1.ListItems(i).SubItems(1)))
                        If ListView1.ListItems(i).SubItems(2) <> "" Then Importe = Importe + ImporteFormateadoAmoneda(CStr(ListView1.ListItems(i).SubItems(2)))
                        
                        Text1(2).Tag = CCur(Text1(2).Tag) - Importe
                        Text1(3).Tag = CCur(Text1(3).Tag) + Importe
                        Text1(12).Tag = Text1(12).Tag + Importe
                        ListView1.ListItems(i).SubItems(1) = ""
                        ListView1.ListItems(i).SubItems(2) = Format(Importe, FormatoImporte)
                        ListView1.ListItems(i).SmallIcon = 6 'baja
                    End If
                    
                Next
            
            
            
            
            RS.MoveNext
        Wend
        RS.Close
  
        cad = Me.Tag
        cad = cad & " AND FechaAlta>" & DBSet(FFin, "F")
        cad = cad & " AND Fechabaja<=" & DBSet(FFin, "F")
        RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            'Añadimos en el list2
            Set itmX = ListView2.ListItems.Add
            cad = Format(RS!FechaBaja, "dd/mm/yyyy")
            itmX.Text = cad
            cad = Format(RS!fechaalta, "dd/mm/yyyy")
            itmX.SubItems(1) = cad
            itmX.SubItems(2) = RS!descbaja
            itmX.SmallIcon = 6
            
            
                'falta
                For i = 1 To ListView1.ListItems.Count
                    If ListView1.ListItems(i).Text = itmX.Text Then
                        'ESTE ES EL DIA DE LA BESTIA, o de la baja ;)
                        
                        Importe = 0
                        If ListView1.ListItems(i).SubItems(1) <> "" Then Importe = Importe + ImporteFormateadoAmoneda(CStr(ListView1.ListItems(i).SubItems(1)))
                        If ListView1.ListItems(i).SubItems(2) <> "" Then Importe = Importe + ImporteFormateadoAmoneda(CStr(ListView1.ListItems(i).SubItems(2)))
                        
                        Text1(2).Tag = CCur(Text1(2).Tag) - Importe
                        Text1(3).Tag = CCur(Text1(3).Tag) + Importe
                        Text1(12).Tag = Text1(12).Tag + Importe
                        ListView1.ListItems(i).SubItems(1) = ""
                        ListView1.ListItems(i).SubItems(2) = Format(Importe, FormatoImporte)
                        ListView1.ListItems(i).SmallIcon = 6 'baja
                    End If
                    
                Next
            
            
            
            RS.MoveNext
        Wend
        RS.Close
   
    
    
    
    
    
    
    
     
    
    'Ajuste nomina
    Dias = CCur(Text1(6).Tag)
  '  If Dias > Int(Dias) Then
  '      Dias = Int(Dias) + 1
  '  Else
  '      Dias = Int(Dias)
  '  End If
    Text1(6).Text = Format(Dias, "0")
    Text1(0).Text = Format(Text1(0).Tag, "00")
    Text1(1).Text = Format(Text1(1).Tag, "00")
    Text1(5).Text = Format(Text1(5).Tag, "00")
    For i = 2 To 4   'Pq el 5 es de dias
        Text1(i).Text = Format(Text1(i).Tag, FormatoImporte)
    Next i
    i = 12
    Text1(i).Text = Format(Text1(i).Tag, FormatoImporte)
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    'Hago un conteo de miercoles y sabados trabajados
    If MediosDias <> "" Then
        Text1(8).Text = ContadorMier
        Text1(9).Text = ContadorSab
        Icono = 1
        ContadorSab = 0: ContadorMier = 0
        Do
            Semana = InStr(Icono, MediosDias, "|")
            If Semana > 0 Then
                cad = Mid(MediosDias, Icono, Semana - Icono)
                Dias = Weekday(CDate(cad), vbMonday)
                If Dias = 3 Then
                    ContadorMier = ContadorMier + 1
                Else
                    If Dias = 6 Then ContadorSab = ContadorSab + 1
                End If
                Icono = Semana + 1
            End If
        Loop Until Semana = 0
        Text1(8).Text = Text1(8).Text & " / " & ContadorMier
        Text1(9).Text = Text1(9).Text & " / " & ContadorSab
    End If
    'Si son jorandas semanles cargamos
    If JornadasSemanales Then
        cad = "Select * from JornadasSemanales where "
        cad = cad & "  Fecha>=" & DBSet(FechaIni, "F")
        cad = cad & " AND Fecha<=" & DBSet(FFin, "F")
        cad = cad & " AND idTrabajador = " & RecuperaValor(Trabajador, 2)
        cad = cad & " ORDER BY fecha"
        RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Text2(0).Tag = 0: Text2(1).Tag = 0: Text2(2).Tag = 0: Text2(3).Tag = 0
        While Not RS.EOF
            Set itmX = ListView3.ListItems.Add
            cad = Format(RS!Fecha, "dd/mm/yyyy")
            itmX.Text = cad
            itmX.SubItems(1) = RS!diasofi
            Text2(0).Tag = Text2(0).Tag + RS!diasofi
            
            itmX.SubItems(2) = RS!horasofi
            Text2(1).Tag = Text2(1).Tag + RS!horasofi
            
            itmX.SubItems(3) = RS!Dias
            Text2(2).Tag = Text2(2).Tag + RS!Dias
            
            itmX.SubItems(4) = Format(RS!HN, "0.00")
            Text2(3).Tag = Text2(3).Tag + RS!HN
            itmX.SmallIcon = 7
            
            
            
            itmX.SubItems(5) = Format(RS!bolsaantes, "0.00")
            itmX.SubItems(6) = Format(RS!bolsadespues, "0.00")
            
            
            RS.MoveNext
        Wend
        
        RS.Close
        
        For i = 0 To 3
            Text2(i).Text = Text2(i).Tag
        Next i
    End If
    Set RS = Nothing
    Text1(7).Text = DiasEnNomina
    Text1(11).Text = Format(Text1(11).Tag, "0.00")
    If Not (ListView1.SelectedItem Is Nothing) Then CargaTicajes ListView1.SelectedItem.Text
    
    'Solo para geslab en picasent
    'ESTO Sera para picassent cuando entre
    'If MiEmpresa.QueEmpresa = 0 Then ObtenerHorasNormalesQueEranCompensablesMyS
    
    Dim C1 As Currency
    Dim C2 As Currency
    Dim C3 As Currency
    C1 = 0
    C2 = 0
    C3 = 0
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).SubItems(1) <> "" Then C1 = C1 + ImporteFormateadoAmoneda(CStr(ListView1.ListItems(i).SubItems(1)))
        If ListView1.ListItems(i).SubItems(2) <> "" Then C2 = C2 + ImporteFormateadoAmoneda(CStr(ListView1.ListItems(i).SubItems(2)))
        If ListView1.ListItems(i).SubItems(3) <> "" Then C3 = C3 + ImporteFormateadoAmoneda(CStr(ListView1.ListItems(i).SubItems(3)))
    Next i
    Text4(0).Text = Format(C1, "0.00")
    If vEmpresa.QueEmpresa = 4 Then
        
        If Option1(1).Value Then
            Text4(1).Text = Format(C2, "0.00")
            Text4(2).Text = Format(C3, "0.00")
        Else
            Text4(1).Text = Format(C2, "0.00")
            Text4(2).Text = ""
        End If
    Else
        Text4(1).Text = Format(C2, "0.00")
    End If
End Sub




Private Sub CargaTicajes(Fecha As String)
Dim C As String
Dim RN As ADODB.Recordset
    Set RN = New ADODB.Recordset
    C = "Select * from entradamarcajes where idtrabajador=" & RecuperaValor(Trabajador, 2)
    C = C & " AND fecha = " & DBSet(Mid(Fecha, 1, 10), "F") & " ORDER BY hora"
    RN.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    C = ""
    While Not RN.EOF
        C = C & "    " & Format(RN!Hora, "hh:mm:ss")
        RN.MoveNext
    Wend
    RN.Close
    Set RN = Nothing
    Text3.Text = Trim(C)
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    CargaTicajes Item.Text
End Sub





'---------------------------------------------------
'Ajuste por cambio en compensables, miercoles y sabado
Private Sub ObtenerHorasNormalesQueEranCompensablesMyS()
Dim FechaFin As Date
    FechaFin = DateAdd("m", 1, FechaIni)
    FechaFin = DateAdd("d", -1, FechaFin)
    HorasCompensablesMiercolesSabado = 0
    RecalculoHorasMiercolesSabados FechaIni, FechaFin, True
    RecalculoHorasMiercolesSabados FechaIni, FechaFin, False
    
    If HorasCompensablesMiercolesSabado > 0 Then
        Text1(10).Text = HorasCompensablesMiercolesSabado
    Else
        Text1(10).Text = ""
    End If
    
End Sub

Private Sub RecalculoHorasMiercolesSabados(F1 As Date, F2 As Date, Miercoles As Boolean)
Dim cad As String
Dim RF As ADODB.Recordset
Dim HT As Currency
Dim Horas As Currency

    

    cad = "SELECT EntradaMarcajes.idTrabajador, EntradaMarcajes.Fecha, Weekday([Fecha]) AS Expr1"
    cad = cad & " From EntradaMarcajes"
    cad = cad & " Where EntradaMarcajes.Fecha >= #" & Format(F1, "yyyy/mm/dd") & "# And"
    cad = cad & " EntradaMarcajes.Fecha <= #" & Format(F2, "yyyy/mm/dd") & "# And "
    cad = cad & " Weekday([Fecha]) = "
    If Miercoles Then
        cad = cad & " 4"
    Else
        cad = cad & " 7"
    End If
    
    'Trabajador
    cad = cad & " AND idtrabajador = " & RecuperaValor(Trabajador, 2)
    cad = cad & " And Hora "
    If Miercoles Then
        cad = cad & " <"
    Else
        cad = cad & " >"
    End If
    cad = cad & " #14:00:00# group by  EntradaMarcajes.idTrabajador, EntradaMarcajes.Fecha,  Weekday([Fecha])"
    cad = cad & " ORDER BY EntradaMarcajes.idTrabajador, EntradaMarcajes.Fecha"
    Set RF = New ADODB.Recordset
    RF.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Horas = 0
    HT = 0
    While Not RF.EOF

        Horas = NuevoCalculoHorasDiaXS(RF!idTrabajador, RF!Fecha, Miercoles)
        HT = HT + Horas
        RF.MoveNext
        
    Wend
    RF.Close
    Set RF = Nothing
    HorasCompensablesMiercolesSabado = HorasCompensablesMiercolesSabado + HT



End Sub



Private Function NuevoCalculoHorasDiaXS_(Trabajador As Long, Fecha As Date, DeMiercoles As Boolean) As Currency
Dim RH As ADODB.Recordset
Dim C As String
Dim T1 As Currency
Dim T2 As Currency
Dim E As Boolean
Dim Seguir As Boolean
Dim NuevaHora As Currency
Dim HoraIntermediaMiercolesSabado As Date
Dim i As Integer

    HoraIntermediaMiercolesSabado = "14:00:00"
    NuevoCalculoHorasDiaXS_ = 0
    Set RH = New ADODB.Recordset
   
    C = "Select * from entradamarcajes where idtrabajador=" & Trabajador
    C = C & " AND fecha = #" & Format(Fecha, "yyyy/mm/dd") & "# "
    
    'C = C & " AND fecha = #" & Format(Fecha, "yyyy/mm/dd") & "# AND hora "
    'If DeMiercoles Then
    '    C = C & "<"
    'Else
    '    C = C & ">"
    'End If
    'C = C & "#14:00:00# ORDER BY hora"
    C = C & " ORDER BY hora"

    RH.Open C, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    Seguir = True
    'Si tiene algun --> Miercoles y es por la tarde ---->NORMAL, no hago nada
    '                     sabado   y es por la mañana ---> "
    If Not RH.EOF Then
        If DeMiercoles Then
       
            If RH!Hora > CDate(HoraIntermediaMiercolesSabado) Then
                Seguir = False
            Else
                'RH.MoveFirst
            End If
        Else
            RH.MoveLast
           'sabado
            If RH!Hora < CDate(HoraIntermediaMiercolesSabado) Then
                Seguir = False
            Else
                RH.MoveFirst
            End If
            
        End If
    End If
    
    
    If Not Seguir Then
        RH.Close
        Exit Function
    End If
    E = True
    NuevaHora = 0
    While Not RH.EOF

        If E Then
            If DeMiercoles Then
                If RH!Hora < CDate(HoraIntermediaMiercolesSabado) Then
                    T1 = CCur(DevuelveValorHora(RH!Hora))
                Else
                    'Ya es cuando le toca
                    RH.MoveLast
                End If
                
             Else
                If RH!Hora >= CDate(HoraIntermediaMiercolesSabado) Then
                    T1 = CCur(DevuelveValorHora(RH!Hora))
                Else
                    'Ya es cuando le toca
                  
                End If
             End If
                
        Else
            'Si tiene valor t1 calculamos dif
            If T1 > 0 Then
                T2 = CCur(DevuelveValorHora(RH!Hora))
                T1 = T2 - T1
            
                NuevaHora = NuevaHora + T1
                T1 = 0

            End If
        End If
        E = Not E
        RH.MoveNext
    Wend
        
    RH.Close    'para que no coja los 700 y 900
    If NuevaHora > 0 And Trabajador < 700 Then
        C = "Select marcajes.*,ExcesoDefecto from marcajes, Incidencias WHERE marcajes.IncFinal = Incidencias.IdInci AND "
        C = C & " idtrabajador=" & Trabajador & " AND fecha = #" & Format(Fecha, "yyyy/mm/dd") & "# "
        RH.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RH.EOF Then
            MsgBox "Mal: " & C
        Else
            
            
            
            
            'Si deberia tenre mas horas comepnsables
            If RH!IncFinal <> vEmpresa.IncRetraso Then
                If NuevaHora > RH!HorasIncid Then
                    T1 = NuevaHora - RH!HorasIncid
                    'Tenemos lo que incrementa y decrementa en comepnsables y en normales
                    'nunca puede aumentar mas que lo que ha trabajador
                    T2 = RH!HorasTrabajadas - RH!HorasIncid
                    If T1 > T2 Then T1 = T2
                    NuevaHora = T1
                Else
                   'No hacemos nada
                   NuevaHora = 0
                End If
            End If
            RH.Close
            NuevoCalculoHorasDiaXS_ = NuevaHora
            If NuevaHora > 0 Then
                MediosDias = Format(Fecha, "dd/mm/yyyy")
                For i = 1 To ListView1.ListItems.Count
                    C = Mid(ListView1.ListItems(i).Text, 1, 10)
                    If C = MediosDias Then
                        ListView1.ListItems(i).Text = MediosDias & " *C"
                        If ListView1.ListItems(i).SubItems(1) = "" Then
                            T1 = 0
                        Else
                            T1 = CCur(ListView1.ListItems(i).SubItems(1))
                        End If
                        If ListView1.ListItems(i).SubItems(2) = "" Then
                            T2 = 0
                        Else
                            T2 = CCur(ListView1.ListItems(i).SubItems(2))
                        End If
                        'Updateamos
                        T1 = T1 - NuevaHora
                        T2 = T2 + NuevaHora
                        ListView1.ListItems(i).SubItems(1) = Format(T1, FormatoImporte)
                        ListView1.ListItems(i).SubItems(2) = Format(T2, FormatoImporte)
                        'Los totales
                        T1 = CCur(Text1(2).Text)
                        T2 = CCur(Text1(3).Text)
                        T1 = T1 - NuevaHora
                        T2 = T2 + NuevaHora
                        Text1(2).Text = Format(T1, FormatoImporte)
                        Text1(3).Text = Format(T2, FormatoImporte)
                        Exit For
                    End If
                Next i
            End If
        End If
    End If
    
End Function




Private Sub Option1_Click(Index As Integer)
    Me.ListView1.ColumnHeaders(4).Text = "Total"
    If Option1(1).Value Then
        If vEmpresa.QueEmpresa = 4 Then Me.ListView1.ColumnHeaders(4).Text = "EXT"

    End If
    CargaGrid
End Sub
