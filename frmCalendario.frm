VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCalendario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calendarios"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   Icon            =   "frmCalendario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameAux 
      Height          =   5535
      Left            =   240
      TabIndex        =   14
      Top             =   1200
      Width           =   10815
      Begin VB.CheckBox chkActual 
         Caption         =   "Temporada actual"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   5160
         Value           =   1  'Checked
         Width           =   4335
      End
      Begin VB.Frame Frame4 
         Height          =   1335
         Left            =   4800
         TabIndex        =   21
         Top             =   3960
         Width           =   5895
      End
      Begin MSAdodcLib.Adodc AdoAux 
         Height          =   330
         Index           =   0
         Left            =   2040
         Top             =   5160
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   2
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "Descripcion|T|N|||calendariof|descripcion|||"
         Text            =   "Nº exped"
         Top             =   4680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   0
         Left            =   360
         MaxLength       =   11
         TabIndex        =   18
         Tag             =   "Cod|N|N|0|99999|calendariof|idcal||S|"
         Text            =   "Nº exped"
         Top             =   4680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   1
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Cod|F|N|||calendariof|Fecha|dd/mm/yyyy|S|"
         Text            =   "li"
         Top             =   4680
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSDataGridLib.DataGrid DataGridAux 
         Bindings        =   "frmCalendario.frx":000C
         Height          =   4455
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   7858
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   0
         Left            =   1440
         TabIndex        =   15
         Top             =   180
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Agregar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar "
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar "
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Copiar festivos"
               Object.Tag             =   "2"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   1
         Left            =   6120
         TabIndex        =   16
         Top             =   180
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar "
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Temporada siguiente"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3255
         Left            =   4800
         TabIndex        =   22
         Top             =   600
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   5741
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Inicio"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fin"
            Object.Width           =   2206
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Horario"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label Label4 
         Caption         =   "HORARIOS"
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "FESTIVOS"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   645
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   10095
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   600
         MaxLength       =   4
         TabIndex        =   0
         Tag             =   "Código|N|N|1|99999|calendario|idCal|0000|S|"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   3120
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "Descripcion|T|N|||calendario|descripcion|||"
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombre "
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Cód."
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   6960
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9180
      TabIndex        =   5
      Top             =   7200
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Top             =   7200
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3360
      Top             =   7080
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9180
      TabIndex        =   11
      Top             =   7200
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Index           =   0
         Left            =   7800
         TabIndex        =   13
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver Todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         HelpContextID   =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: CÈSAR +-+-
' +-+-+-+-+-+-+-+-+-+-+-
 
' **************** PER A QUE FUNCIONE EN UN ATRE MANTENIMENT ********************
' +-+-+-+- DISSENY +-+-+-+-
' 1. Posar tots els controls al formulari
' 2. Posar els TAGs
' 3. Posar els MAXLENGTHs
' 4. Posar els TABINDEXs
'
' +-+-+-+-+- CODIFICACIÓ +-+-+-+-+-
' 1. Definir variables per a cridar a atres formularis
' 2. En Form_Load() canviar el nom de la taula i la clau primaria de l'ORDER BY
' 3. En PonerModo Revisar lo que bloquejem, el nom de la clau primaria,
'    les imagens de buscar i les de dates
' 4. En PonerLongCampos() posar només els camps numérics
' 5. En PonerModoOpcionesMenu(Modo) comentar o descomentar depenent
'    de si n'hi ha menú desplegable o no
' 6. (SI N'HI HA BUSCAR DATA) En imgFec_Click(Index As Integer) i en
'    frmC_Selec(vFecha As Date), canviar l'index de imgFec pel 1r index de les
'    imagens de buscar data
' 7. (SI N'HI HAN CAMPS DE BUSCAR CODIS) En imgBuscar_Click(Index As Integer)
'    codificar tots els camps. Per a cada camp fer la funció, per eixemple,
'    frmPob_DatoSeleccionado(CadenaSeleccion As String)
' 8. En MandaBusquedaPrevia(CadB As String) arreglar-ho per a vore lo que es desije
' 9. En BotonAnyadir() canviar el nom de la taula i el nom de la clau primaria
' 10. En BotonEliminar() canviar els noms, els formats i el DELETE
' 11. Si alguna atra taula apunta a la actual, en SePuedeEliminar()
'     canviar els parametres de la funció; sino, comentar-ho tot
' 12. (SI N'HI HAN CAMPS DE BUSCAR CODIS) En PonerCampos() configurar els camps
'     de les descripcions
' 13. En DatosOk() canviar els arguments de DevuelveDesdeBD i el mensage
' 14. En Eliminar() canviar el nom de la clau primaria
' 15. (SI N'HI HAN COMBOS) En CargaCombo() configurar els distints Combos
'
'
' 30. Si el nom del camp que te la clau primaria NO es Text1(0), canviar-ho en:
'     BotonBuscar(), HacerBusqueda(), BotonAnyadir(), BotonModificar(), CmdCancelar_Click()
' *******************************************************************************

Option Explicit

Public DatosADevolverBusqueda As String    'Tindrà el nº de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)

' ****** Definir variables per a cridar a atres formularis *********
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
' *****************************************************************

Private Modo As Byte
'*************** MODOS ********************
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'   5.-  Manteniment Llinies

'***Variables comuns a tots els formularis*****
Dim PrimeraVez As Boolean
Private ModoLineas As Byte
Private CadenaConsulta As String 'SQL de la taula principal del formulari
Private Ordenacion As String
Private NombreTabla As String  'Nom de la taula

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
Dim Indice As Byte 'Indice del text1 donde se ponen los datos devueltos desde otros Formularios de Mtos
Dim cad_meua As String
Dim CadB As String


Private Sub chkActual_Click()
    If Modo = 2 Then
        CargaGrid 0, True
    End If
End Sub

Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    'situarnos en el registro que acabamos de insertar
                    Data1.RecordSource = "Select * from " & NombreTabla & " WHERE idCal=" & Text1(0).Text & Ordenacion
                    PosicionarData
                End If
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me) Then
                    TerminaBloquear
                    PosicionarData
                End If
            End If
        Case 5
            Select Case ModoLineas
                Case 1 'afegir llinia
                    InsertarLinea
                Case 2 'modificar llinies
                    ModificarLinea
                    PosicionarData
            End Select
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

' *** adrede per ad este manteniment ***
'Private Sub Combo1_Click(Index As Integer)
'    Text1_LostFocus (16) 'el camp que te el codi del banc
'End Sub
''***************************************+
'
'Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpress KeyAscii
'End Sub
'
'
'Private Sub Combo1_GotFocus(Index As Integer)
'  If Modo = 1 Then Combo1(Index).BackColor = vbYellow
'End Sub
'
'Private Sub Combo1_LostFocus(Index As Integer)
'    If Combo1(Index).BackColor = vbYellow Then Combo1(Index).BackColor = vbWhite
'End Sub
'

Private Sub Form_Activate()
    
    If PrimeraVez Then
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
End Sub

Private Sub Form_Load()
Dim I As Integer

    PrimeraVez = True

    ' ICONITOS DE LA BARRA
    btnPrimero = 16 'index del botó "primero"
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'el 1 i el 2 son separadors
        .Buttons(3).Image = 1   'Buscar
        .Buttons(4).Image = 2   'Todos
        'el 5 i el 6 son separadors
        .Buttons(7).Image = 3   'Insertar
        .Buttons(8).Image = 4   'Modificar
        .Buttons(9).Image = 5   'Borrar
        'el 10 i el 11 son separadors
        .Buttons(12).Image = 10  'Imprimir
        .Buttons(13).Image = 11  'Salir
        'el 13 i el 14 son separadors
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
      

      
    
    With Me.ToolAux(0)
        '.ImageList = frmPpal.imgListComun_VELL
        .HotImageList = frmPpal.imgListComun_OM16
        .DisabledImageList = frmPpal.imgListComun_BN16
        .ImageList = frmPpal.imgListComun16
        .Buttons(1).Image = 16   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        .Buttons(7).Image = 19   'Copiar horarios
    End With

    
    With Me.ToolAux(1)
        '.ImageList = frmPpal.imgListComun_VELL
        .HotImageList = frmPpal.imgListComun_OM16
        .DisabledImageList = frmPpal.imgListComun_BN16
        .ImageList = frmPpal.imgListComun16
        .Buttons(1).Image = 1  'Insertar
        .Buttons(2).Image = 15 'asignar
        .Buttons(4).Image = 16 'asignar
    End With

    
    
    ' *** canviar el nom de la taula i la clau primaria de l'ORDER BY ***
    NombreTabla = "calendario"
    Ordenacion = " ORDER BY idCal"
    ' **********************************************************
        
    'Vemos como esta guardado el valor del check
    'ckVistaPrevia(0).Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where idCal=-1"
    Data1.Refresh
          
'    For I = 0 To Combo1.Count - 1
'        CargaCombo I
'    Next I
    
    For I = 0 To DataGridAux.Count - 1
        CargaGrid I, False 'carregue els datagrids de llinies
        'DataGridAux(i).Enabled = False 'inicialment tots disabled
    Next I
    
    
    
    
    LimpiarCampos   'Limpia los campos TextBox
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
        Text1(0).BackColor = vbYellow 'codclien
    End If
End Sub


Private Sub LimpiarCampos()
Dim I As Integer
    
    On Error Resume Next

    Limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    ListView1.ListItems.Clear
    
'    Check1(0).Value = 0
'    For I = 0 To Me.Combo1.Count - 1
'        Me.Combo1(I).ListIndex = -1
'    Next I
'    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim I As Integer, NumReg As Byte
Dim B As Boolean

    On Error GoTo EPonerModo
 
    'Actualiza Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    'PonerIndicador lblIndicador, Modo, ModoLineas
    PonerIndicador lblIndicador, Modo
       
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.Visible = (Modo = 2)
    Else
        cmdRegresar.Visible = False
    End If
    
    '=======================================
    B = (Modo = 2)
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    
    '---------------------------------------------
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.Visible = B
    cmdAceptar.Visible = B
        
    'Bloquea los campos Text1 si no estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    
    ' **************************************************************************************************
    ' *** Revisar lo que bloquejem, el nom de la clau primaria, les imagens de buscar i les de dates ***
    BloquearText1 Me, Modo
    BloquearCheck1 Me, Modo
    BloquearCombo Me, Modo
    BloquearImgZoom Me, Modo, 0

'    PosicionarCombo Combo1(2), 724
'    For i = 0 To Combo1(2).ListCount - 1
'        If Combo1(2).ItemData(i) = 724 Then
'            Combo1(2).ListIndex = i
'            Exit For
'        End If
'    Next i
    
    If Modo = 4 Then _
        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
    
 '   BloquearImgBuscar Me, Modo
    BloquearImgFec Me, 11, Modo
    BloquearImgFec Me, 27, Modo
    BloquearImgFec Me, 28, Modo
 
    ' **************************************************************************************************
                
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor.
    PonerLongCampos
    PonerModoOpcionesMenu (Modo) 'Activar opciones menu según modo
    PonerOpcionesMenu 'Activar/Desact botones de menu segun permisos del usuario
    
    
    B = (Modo = 4) Or (Modo = 2)
    For I = 0 To DataGridAux.Count - 1
        DataGridAux(I).Enabled = B
    Next I
 
    
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los Text1
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub PonerOpcionesMenu()
'Activa/desact. las Opciones de Menu y Toolbar según permisos de usuario
    PonerOpcionesMenuGeneral Me
End Sub

Private Sub PonerModoOpcionesMenu(Modo)
'Activa/desact. las Opciones de Menu y Toolbar según el modo en que estemos
Dim B As Boolean
Dim Baux As Boolean
Dim I As Integer

' ******** comentar o descomentar depenent de si n'hi ha menú desplegable o no ****
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    B = (Modo = 2 Or Modo = 0)
    'Buscar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnBuscar.Enabled = B
    'Ver Todos
    Toolbar1.Buttons(4).Enabled = B
    Me.mnVerTodos.Enabled = B
    'Insertar
    Toolbar1.Buttons(7).Enabled = B
    Me.mnNuevo.Enabled = B
    
    B = (Modo = 2 And Data1.Recordset.RecordCount > 0) 'And Not DeConsulta
   
    'Modificar
    Toolbar1.Buttons(8).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(9).Enabled = B
    Me.mnEliminar.Enabled = B
    
    
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = b Or (Modo = 0)
    Toolbar1.Buttons(12).Enabled = B
    
    
    'LINEAS
    'b = (Modo = 3 Or Modo = 4 Or Modo = 2)
    B = (Modo = 2)
    For I = 0 To ToolAux.Count - 1
 
        ToolAux(I).Buttons(1).Enabled = B
        If I = 0 Then
            If B Then Baux = (B And Me.AdoAux(I).Recordset.RecordCount > 0)
            ToolAux(I).Buttons(2).Enabled = Baux
            ToolAux(I).Buttons(3).Enabled = Baux
            ToolAux(I).Buttons(7).Enabled = B
        Else
            ToolAux(I).Buttons(2).Enabled = B
            ToolAux(I).Buttons(4).Enabled = B
        End If
        

    Next I
    
    
    
    
' ********************************************************************************
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim Aux As String

    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        '   Como la clave principal es unica, con poner el sql apuntando
        '   al valor devuelto sobre la clave ppal es suficiente
        ' *** canviar o llevar el WHERE ***
        'CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

'Private Sub imgFec_Click(Index As Integer)
'    Dim esq As Long
'    Dim dalt As Long
'    Dim menu As Long
'    Dim obj As Object
'
'    Set frmC = New frmCal
'
'    esq = imgFec(Index).Left
'    dalt = imgFec(Index).Top
'
'    Set obj = imgFec(Index).Container
'
'    While imgFec(Index).Parent.Name <> obj.Name
'        esq = esq + obj.Left
'        dalt = dalt + obj.Top
'        Set obj = obj.Container
'    Wend
'
'    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
'
'    ' es desplega baix i cap a la dreta
'    'frmC.Left = esq + imgFec(Index).Parent.Left + 30
'    'frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40
'
'    ' es desplega dalt i cap a la esquerra
'    frmC.Left = esq + imgFec(Index).Parent.Left - frmC.Width + imgFec(Index).Width + 40
'    frmC.Top = dalt + imgFec(Index).Parent.Top - frmC.Height + menu - 25
'
'    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
'    imgFec(27).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
'    If Text1(Index).Text <> "" Then frmC.NovaData = Text1(Index).Text
'
'    frmC.Show vbModal
'    Set frmC = Nothing
'    PonerFoco Text1(CByte(imgFec(27).Tag))
    ' **************************************************************************
'End Sub

'Private Sub frmC_Selec(vFecha As Date)
'    Text1(CByte(imgFec(27).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
''    PonerFoco txtAux(CByte(imgFec(27).Tag))
'End Sub

'Private Sub imgBuscar_Click(Index As Integer)
'    'Screen.MousePointer = vbHourglass
'    TerminaBloquear
'    Select Case Index
'        Case 0, 8 'empresa
'            If Index = 8 Then
'                Indice = 33
'            Else
'                Indice = 3
'            End If
'            Set frmEmp = New frmEmpresas
'            frmEmp.DatosADevolverBusqueda = "0|1|"
'            If Not IsNumeric(Text1(Indice).Text) Then Text1(Indice).Text = ""
'            frmEmp.CodigoActual = Text1(Indice).Text
'            frmEmp.Show vbModal
'            Set frmEmp = Nothing
'            PonerFoco Text1(Indice)
'
'        Case 1 'agencia
'            Indice = 4
'            Set frmAge = New frmAgencias2
'            frmAge.DatosADevolverBusqueda = "0|1|"
'            If Not IsNumeric(Text1(Indice).Text) Then Text1(Indice).Text = ""
'            frmAge.DeConsulta = True
'            frmAge.Empresa = Text1(3).Text
'            frmAge.CodigoActual = Text1(4).Text
'            frmAge.Show vbModal
'            Set frmAge = Nothing
'            PonerFoco Text1(Indice)
'
'        Case 2, 5 'población i población banco
'            If Index = 2 Then
'                Indice = 6
'            ElseIf Index = 5 Then
'                Indice = 21
'            End If
'            Set frmPob = New frmPoblacio
'            frmPob.DatosADevolverBusqueda = "0|1|2|3|4|5|"
'            If Not IsNumeric(Text1(Indice).Text) Then Text1(Indice).Text = ""
'            frmPob.CodigoActual = Text1(Indice).Text
'            frmPob.Show vbModal
'            Set frmPob = Nothing
'            PonerFoco Text1(Indice)
'
'        Case 3 'Cuenta Contable
'            Set frmCtas = New frmCtasConta
'            frmCtas.DatosADevolverBusqueda = "0|1|"
'            If Not IsNumeric(Text1(15).Text) Then Text1(15).Text = ""
'            frmCtas.CodigoActual = Text1(15).Text
'            frmCtas.Show vbModal
'            Set frmCtas = Nothing
'            PonerFoco Text1(15)
'
'        Case 4 'Cuenta Bancaria
'            Set frmBan = New frmBancsofi
'            frmBan.DatosADevolverBusqueda = "4|1|3|"
'            frmBan.CodigoActual = Text1(16).Text
'            frmBan.NuevoPais = Me.Combo1(2).ItemData(Combo1(2).ListIndex)
'            frmBan.Show vbModal
'            Set frmBan = Nothing
'            PonerFoco Text1(16)
'
'        Case 6 'tipo de nómina
'            Set frmTiN = New frmTiponomi
'            frmTiN.DatosADevolverBusqueda = "0|1|"
'            If Not IsNumeric(Text1(24).Text) Then Text1(24).Text = ""
'            frmTiN.CodigoActual = Text1(24).Text
'            frmTiN.Show vbModal
'            Set frmTiN = Nothing
'            PonerFoco Text1(24)
'
'        Case 7 'tipo de empleado
'            Set frmTiE = New frmTiposemp
'            frmTiE.DatosADevolverBusqueda = "0|1|"
'            If Not IsNumeric(Text1(30).Text) Then Text1(30).Text = ""
'            frmTiE.CodigoActual = Text1(30).Text
'            frmTiE.Show vbModal
'            Set frmTiE = Nothing
'            PonerFoco Text1(30)
'    End Select

'    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
'End Sub



'Private Sub ImgModifHora_Click()
'Dim F As Date
'
'    miSQL = Text1(1).Text
'    F = DateAdd("m", -1, Now)
'    F = CDate("01/" & Month(F) & "/" & Year(F))
'    miSQL = miSQL & "|" & Format(F, "dd/mm/yyyy") & "|"
'    F = DateAdd("m", 1, Now)
'    F = CDate(DiasMes(Month(F), Year(F)) & "/" & Month(F) & "/" & Year(F))
'    miSQL = miSQL & Format(F, "dd/mm/yyyy") & "|"
'
'    'Parametros
'    ' Nombre | Fec ini | Fec Fin | Codigo trabjador
'    miSQL = miSQL & Text1(0).Text & "|"
'
'    FrmVarios.Opcion = 0
'    FrmVarios.Parametros = miSQL
'    FrmVarios.Show vbModal
'End Sub



'Private Sub imgZoom_Click(Index As Integer)
'    frmVerCalendario.CodigoTrab = Val(Text1(0).Text)
'    frmVerCalendario.idCal = Val(Text1(26).Text)
'    frmVerCalendario.Texto = Text1(1).Text
'    frmVerCalendario.Show vbModal
'End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub




Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)

    If Index = 0 Then
        Select Case Button.Index
            Case 1
    '            TerminaBloquear
                BotonAnyadirLinea Index
            Case 2
    '            TerminaBloquear
                BotonModificarLinea Index
            Case 3
    '            TerminaBloquear
                BotonEliminarLinea Index
    '            If Modo = 4 Then
    '                If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
    '            End If
            Case 7
                CadenaDesdeOtroForm = Text1(1).Text & "|" & Text1(0).Text & "|"
                frmListado.Opcion = 10
                frmListado.Show vbModal
                If CadenaDesdeOtroForm = "" Then
                    DoEvents
                    Screen.MousePointer = vbHourglass
                    CargaGrid 0, True
                    Screen.MousePointer = vbDefault
                End If
        End Select
        
    Else
        'Toolbar del horario
        Select Case Button.Index
        Case 1
            'VER
            frmVerCalendario.CodigoTrab = 0
            frmVerCalendario.FeIni = vEmpresa.FechaInicio
            frmVerCalendario.FeFin = vEmpresa.FechaFin
            frmVerCalendario.idCal = Val(Text1(0).Text)
            frmVerCalendario.Texto = Text1(1).Text
            frmVerCalendario.Show vbModal
        Case 2, 4
            'ASIGNAR
            frmAsignaHorario.Opcion = 0
            frmAsignaHorario.OtrosDatos = Text1(0).Text & "|" & Text1(1).Text & "|"
            If Button.Index = 2 Then
                frmAsignaHorario.FeFin = vEmpresa.FechaFin
                frmAsignaHorario.FeIni = vEmpresa.FechaInicio
            Else
                'Temporada siguiente
                frmAsignaHorario.FeFin = DateAdd("yyyy", 1, vEmpresa.FechaFin)
                frmAsignaHorario.FeIni = DateAdd("yyyy", 1, vEmpresa.FechaInicio)
            End If
            frmAsignaHorario.Show vbModal
            If Button.Index = 2 Then CargaCalendario
        End Select
    End If
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 3  'Buscar
           mnBuscar_Click
        Case 4  'Todos
            mnVerTodos_Click
        Case 7  'Nuevo
            mnNuevo_Click
        Case 8  'Modificar
            mnModificar_Click
        Case 9  'Borrar
            mnEliminar_Click
        Case 12 'Imprimir
        '    AbrirListado (10)
            'BotonImprimir
            printNou
'            MsgBox "Falta fer el llista d'empleats"
        Case 13    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
' *************** Si la clau primaria no es Text1(0), canviar-ho ***************
    If Modo <> 1 Then
        LimpiarCampos
        CargaGrid 0, False
        PonerModo 1
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
' ******************************************************************************
End Sub

Private Sub HacerBusqueda()

    CadB = ObtenerBusqueda(Me, False)
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' ******** Si la clau primaria no es Text1(0), canviar-ho ***********
        PonerFoco Text1(0)
        ' *******************************************************************
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
Dim Cad As String
        'Llamamos a al form
        ' **************** arreglar-ho per a vore lo que es desije ****************
        Cad = ""
        Cad = Cad & ParaGrid(Text1(0), 20, "Cód.")
        Cad = Cad & ParaGrid(Text1(1), 60, "Nombre")
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NombreTabla
            frmB.vSQL = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|"
            frmB.vTitulo = "Empleados"
            frmB.vSelElem = 0

            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                    cmdRegresar_Click
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
                PonerFoco Text1(kCampo)
            End If
        End If
        ' *************************************************************************
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim Aux As String
Dim I As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    Cad = ""
    I = 0
    Do
        J = I + 1
        I = InStr(J, DatosADevolverBusqueda, "|")
        If I > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, I - J)
            J = Val(Aux)
            Cad = Cad & Text1(J).Text & "|"
        End If
    Loop Until I = 0
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        
        'Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
        PonerCampos
        PonerModo 2
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub

EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonVerTodos()
    CadB = ""
    LimpiarCampos 'Limpia los Text1
    
    If chkVistaPrevia(0).Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub BotonAnyadir()
    LimpiarCampos 'Vacía los TextBox
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
    
    CargaGrid 0, False
    ' ******* Canviar el nom de la taula, el nom de la clau primaria, i el
    ' nom del camp que te la clau primaria si no es Text1(0) *************
    Text1(0).Text = SugerirCodigoSiguienteStr("calendario", "idcal")
    FormateaCampo Text1(0)
    
    kCampo = 0
    PonerFoco Text1(0)
    
    ' ********************************************************************
End Sub


Private Sub BotonModificar()
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then
        TerminaBloquear
        Exit Sub
    End If
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    ' *** Canviar el nom del camp que te la clau primaria si no es Text1(0) ***
'    BloquearTxt Text1(0), True
    PonerFoco Text1(1)
    ' *************************************************************************
End Sub

Private Sub BotonEliminar()
Dim Cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    'Comprobamos si se puede eliminar
'    If Not SePuedeEliminar Then Exit Sub

    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub

    ' *************** canviar els noms, els formats i el DELETE ****************                  "
    Cad = Cad & "¿Seguro que desea eliminar el Trabajador?"
    Cad = Cad & vbCrLf & "  Código: " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    Cad = Cad & vbCrLf & "  Nombre: " & Data1.Recordset.Fields(1)
    
    ' **************************************************************************
    
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Proveedor", Err.Description
End Sub


Private Sub PonerCampos()

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1


    CargaGrid 0, True
    
    CargaCalendario

    
    PonerModoOpcionesMenu Modo
    
'    Text2.Text = DevuelveDesdeBD("descripcion", "calendario", "idcal", Text1(26).Text, "N")
'    'Los combo
'    kCampo = Data1.Recordset!seccion
'    PosicionarCombo Combo1(0), kCampo
'
'     kCampo = Data1.Recordset!idCategoria
'    PosicionarCombo Combo1(1), kCampo
'
'    kCampo = Data1.Recordset!Control
'    PosicionarCombo Combo1(2), kCampo
'
'    kCampo = Data1.Recordset!tipocontrato
'    PosicionarCombo Combo1(3), kCampo
    
    'Cargamos el calendario laboral
    'Del trabajador
    'CargaCalendario
    
    'Text2(30).Text = PonerNombreDeCod(Text1(30), "tiposemp", "desemple", "tipemple", "N")
    ' *******************************************************************
    
    '-- Esto permanece para saber donde estamos
    kCampo = 0
    lblIndicador.Caption = PonerContRegistros(Me.Data1)
End Sub


Private Sub cmdCancelar_Click()
Dim v
    ' *** canviar el nom del camp que te la clau primaria si no es Text1(0) ***
    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            If Data1.Recordset.EOF Then
                PonerModo 0
            Else
                PonerCampos
                PonerModo 2
            End If
            PonerFoco Text1(0)

        Case 4  'Modificar
            TerminaBloquear
            PonerCampos
            PonerModo 2
            PonerFoco Text1(0)
            
            
            
            
        Case 5
            Select Case ModoLineas
            Case 1 'afegir llinia
                ModoLineas = 0
                DataGridAux(0).AllowAddNew = False
                
                'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                LLamaLineas 0, ModoLineas 'ocultar txtAux
                'If DataGridAux(NumTabMto).Enabled Then DataGridAux(NumTabMto).SetFocus
                DataGridAux(0).Enabled = True
                DataGridAux(0).SetFocus

                If Not AdoAux(0).Recordset.EOF Then
                    AdoAux(0).Recordset.MoveFirst
                End If

            Case 2 'modificar llinies
                ModoLineas = 0
                PonerModo 4
                If Not AdoAux(0).Recordset.EOF Then
                    v = AdoAux(0).Recordset.Fields(1) 'el 1 es el nº de llinia
                    AdoAux(0).Recordset.Find (AdoAux(0).Recordset.Fields(1).Name & " =" & v)
                End If
                LLamaLineas 0, ModoLineas 'ocultar txtAux
            End Select
            
            PosicionarData

            
            
    End Select
    ' *************************************************************************
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean
Dim I As Integer

    On Error GoTo EDatosOK

    'Considereaciones antes de insertar o modificar
    If vEmpresa.laboral Then
        'Los campos txtson necesarios

    
    End If

    DatosOk = False
    B = CompForm(Me)
    If Not B Then Exit Function
    
    ' ******************** canviar els arguments de la funcio i el mensage ****************
    If (Modo = 3) Then 'Insertar
        'comprobar si ya existe el campo de clave primaria
        If ExisteCP(Text1(0)) Then B = False
        
'         Datos = DevuelveDesdeBD("idTrabajador", "empleado", "idTrabajador", Text1(0).Text, "N")
'         If Datos <> "" Then
'            MsgBox "Ya existe el Código de Empleado: " & Text1(0).Text, vbExclamation
'            DatosOk = False
'            PonerFoco Text1(0)
'            Exit Function
'         End If
    End If
    ' *************************************************************************************
         
    DatosOk = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    ' *** canviar-ho per la clua primaria ***
    Cad = "( idcal = " & Text1(0).Text & ")"
    ' ***************************************
    
    If SituarData(Data1, Cad, Indicador) Then
       PonerModo 2
       lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

Private Function Eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

'    Conn.BeginTrans
    ' ***** canviar el nom de la clau primaria *******
    vWhere = " WHERE idcal =" & Data1.Recordset!idCal
    ' ************************************************
              
    conn.Execute "Delete from " & NombreTabla & vWhere
               
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
'        Conn.RollbackTrans
        Eliminar = False
    Else
'        Conn.CommitTrans
        Eliminar = True
    End If
End Function

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Not Text1(Index).MultiLine Then ConseguirFoco Text1(Index), Modo
End Sub



Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean








    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
'
'    ' ***************** configurar els camps de buscar codis *****************
    Select Case Index
        Case 0, 17, 18, 19
            PonerFormatoEntero Text1(Index)


        Case 8 'NIF
            Text1(Index).Text = UCase(Text1(Index).Text)
            'ValidarNIF Text1(Index).Text


        Case 11, 27, 28
            If Not EsFechaOK(Text1(Index)) Then
                Text1(Index).Text = ""
                PonerFoco Text1(Index)
            End If
        
        Case 13 To 16
            PonerFormatoEntero Text1(Index)
'        Case 9, 10 'telèfons, fax i mòbils
'            PosarFormatTelefon Text1(Index)

'        Case 12, 13 'loginweb i passwweb
'            Text1(Index).Text = LCase(Text1(Index).Text)
'            If Index = 12 Then 'login
'                If Not ComprobarLoginEmp(Text1(Index).Text) Then PonerFoco Text1(Index)
'            End If
'
'        Case 6, 21 'poblacion
'            Nuevo = False
'            If Index = 6 Then
''                PonerDatosPoblacion Text1(Index), Text2(Index), Text1(Index + 1), , , Nuevo, Text1(9)
'            ElseIf Index = 21 Then
''                PonerDatosPoblacion Text1(Index), Text2(Index), Text1(Index + 1), , , Nuevo
'            End If
'            If Nuevo Then
''                Indice = Index
''                Set frmPob = New frmPoblacio
''                frmPob.DatosADevolverBusqueda = "0|1|2|3|4|5|"
''                frmPob.NuevoCodigo = Text1(Index).Text
''                Text1(Index).Text = ""
''                TerminaBloquear
''                frmPob.Show vbModal
''                Set frmPob = Nothing
'                If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
'            End If
'
'        Case 3, 33 'empresa
'            If PonerFormatoEntero(Text1(Index)) Then
'                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "empresas", "nomempre", "codempre", "N")
'                If Text2(Index).Text = "" Then
'                    cadMen = "No existe la Empresa: " & Text1(Index).Text & vbCrLf
'                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
'                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
'                        Set frmEmp = New frmEmpresas
'                        frmEmp.DatosADevolverBusqueda = "0|1|"
'                        frmEmp.NuevoCodigo = Text1(Index).Text
'                        Text1(Index).Text = ""
'                        TerminaBloquear
'                        frmEmp.Show vbModal
'                        Set frmEmp = Nothing
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
'                    Else
'                        Text1(Index).Text = ""
'                    End If
'                    PonerFoco Text1(Index)
'                End If
'            Else
'                Text2(Index).Text = ""
'            End If
'
'            If Index = 3 And Text1(4).Text <> "" Then Text1_LostFocus (4)
'
'        Case 4 'agencia
'            If PonerFormatoEntero(Text1(Index)) Then
'                cadMen = Text1(3).Text 'empresa
'                If (cadMen = "" Or Not IsNumeric(cadMen)) Then
'                    Text1(4).Text = ""
'                    Text2(4).Text = ""
'                    Exit Sub
'                End If
'
'                Text2(Index).Text = DevuelveDesdeBDNew(cPTours, "agencias", "desagenc", "codempre", cadMen, "N", , "codagenc", Text1(Index).Text, "N")
'                FormateaCampo Text1(Index)
'                If Text2(Index).Text = "" And Text1(Index) <> "" Then
'                    cadMen = "No existe la Agencia: " & Text1(Index).Text & vbCrLf
'                    cadMen = cadMen & "para la Empresa: " & Text1(3).Text & "  " & Text2(3).Text & vbCrLf
'                    MsgBox cadMen, vbExclamation
''                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
''                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
''                        Set frmAge = New frmAgencias
''                        frmAge.DatosADevolverBusqueda = "0|1|"
''                        frmAge.NuevoCodigo = text1(Index).Text
''                        text1(Index).Text = ""
''                        TerminaBloquear
''                        frmAge.Show vbModal
''                        Set frmAge = Nothing
''                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
''                    Else
''                        text1(Index).Text = ""
''                    End If
'                    PonerFoco Text1(Index)
'                End If
'            Else
'                Text2(Index).Text = ""
'            End If
'
'        Case 15 'Cuenta Contable
'            If Text1(Index).Text = "" Then
'                Text2(Index).Text = ""
'                Exit Sub
'            End If
'            If Modo = 3 And ContieneCaracterBusqueda(Text1(Index).Text) Then Exit Sub     'Busquedas
'            Text2(Index).Text = PonerNombreCuenta(Text1(Index))
'
'        Case 16 'Cuenta Bancaria
'            If PonerFormatoEntero(Text1(Index)) Then
'                If Text1(Index).Text = "" Then Exit Sub
'                Text2(Index).Text = DevuelveDesdeBDNew(cPTours, "bancsofi", "nombanco", "codnacio", Combo1(2).ItemData(Combo1(2).ListIndex), "N", , "codbanco", Text1(Index).Text, "N")
'                If Text2(Index).Text = "" Then
'                    cadMen = "No existe el Banco: " & Text1(Index).Text & "  "
'                    cadMen = cadMen & "para el pais: " & Combo1(2).List(Combo1(2).ListIndex) & vbCrLf
'                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
'                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
'                        Set frmBan = New frmBancsofi
'                        frmBan.DatosADevolverBusqueda = "0|1|"
'                        frmBan.NuevoCodigo = Text1(Index).Text
'                        frmBan.NuevoPais = Combo1(2).ItemData(Combo1(2).ListIndex)
'                        Text1(Index).Text = ""
'                        TerminaBloquear
'                        frmBan.Show vbModal
'                        Set frmBan = Nothing
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
'                    Else
'                        Text1(Index).Text = ""
'                    End If
'                    PonerFoco Text1(Index)
'                End If
'            Else
'                Text2(Index).Text = ""
'            End If
'
'        Case 24 'Tipo de Nómina
'            If PonerFormatoEntero(Text1(Index)) Then
'                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "tiponomi", "desnomin")
'                If Text2(Index).Text = "" And Text1(Index) <> "" Then
'                    cadMen = "No existe el Tipo de Nómina: " & Text1(Index).Text & vbCrLf
'                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
'                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
'                        Set frmTiN = New frmTiponomi
'                        frmTiN.DatosADevolverBusqueda = "0|1|"
'                        frmTiN.NuevoCodigo = Text1(Index).Text
'                        Text1(Index).Text = ""
'                        TerminaBloquear
'                        frmTiN.Show vbModal
'                        Set frmTiN = Nothing
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
'                    Else
'                        Text1(Index).Text = ""
'                    End If
'                    PonerFoco Text1(Index)
'                End If
'            Else
'                Text2(Index).Text = ""
'            End If
'
'        Case 30 'Tipo de Empleado
'            If PonerFormatoEntero(Text1(Index)) Then
'                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "tiposemp", "desemple")
'                If Text2(Index).Text = "" Then
'                    cadMen = "No existe el Tipo de Empleado: " & Text1(Index).Text & vbCrLf
'                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
'                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
'                        Set frmTiE = New frmTiposemp
'                        frmTiE.DatosADevolverBusqueda = "0|1|"
'                        frmTiE.NuevoCodigo = Text1(Index).Text
'                        Text1(Index).Text = ""
'                        TerminaBloquear
'                        frmTiE.Show vbModal
'                        Set frmTiE = Nothing
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
'                    Else
'                        Text1(Index).Text = ""
'                    End If
'                    PonerFoco Text1(Index)
'                End If
'            Else
'                Text2(Index).Text = ""
'            End If
'
'        Case 27, 28, 29 'dates
'            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
    End Select
End Sub




Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not Text1(Index).MultiLine Then
        If KeyAscii = teclaBuscar Then
            If Modo = 1 Or Modo = 3 Or Modo = 4 Then
                Select Case Index
                    Case 3: KEYBusqueda KeyAscii, 0 'empresa
                    Case 4: KEYBusqueda KeyAscii, 1 'agencia
                    Case 6: KEYBusqueda KeyAscii, 2 'poblacion
                    Case 33: KEYBusqueda KeyAscii, 8 'empresa alta
                    Case 24: KEYBusqueda KeyAscii, 6 'tipo de nomina
                    Case 30: KEYBusqueda KeyAscii, 7 'tipo de empleado
                    Case 15: KEYBusqueda KeyAscii, 3 'cuenta contable
                    Case 16: KEYBusqueda KeyAscii, 4 'banco oficial
                    Case 21: KEYBusqueda KeyAscii, 5 'poblacion
                    
                    Case 27: KEYFecha KeyAscii, 27
                    Case 28: KEYFecha KeyAscii, 28
                    Case 29: KEYFecha KeyAscii, 29
                End Select
            End If
        Else
            KeyPress KeyAscii
        End If
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Not Text1(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
        If (Modo = 0 Or Modo = 2) Then Unload Me
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    'imgBuscar_Click (indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
   ' imgFec_Click (Indice)
End Sub


'Private Sub CargaCombo(Index As Integer)
'
'
'    Combo1(Index).Clear
'
'    ' ******* configurar els distints Combos **********
'    Select Case Index
'        Case 0  'Seccion
'            cad_meua = "SELECT * FROM secciones ORDER BY idseccion"
'            Set miRs = New ADODB.Recordset
'            miRs.Open cad_meua, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'            While Not miRs.EOF
'                Combo1(Index).AddItem miRs!Nombre
'                Combo1(Index).ItemData(Combo1(Index).NewIndex) = miRs!idseccion
'                miRs.MoveNext
'            Wend
'
'            miRs.Close
'
'        Case 1 'forma de cobro
'
'            cad_meua = "SELECT * FROM categorias ORDER BY idcategoria"
'            Set miRs = New ADODB.Recordset
'            miRs.Open cad_meua, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'            While Not miRs.EOF
'                Combo1(Index).AddItem miRs!nomcategoria
'                Combo1(Index).ItemData(Combo1(Index).NewIndex) = miRs!idCategoria
'                miRs.MoveNext
'            Wend
'
'            miRs.Close
'
'        Case 2
'
'            cad_meua = "SELECT * FROM stipocontrol ORDER BY tipocontrol"
'            Set miRs = New ADODB.Recordset
'            miRs.Open cad_meua, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'            While Not miRs.EOF
'                Combo1(Index).AddItem miRs!desccontrol
'                Combo1(Index).ItemData(Combo1(Index).NewIndex) = miRs!tipocontrol
'                miRs.MoveNext
'            Wend
'
'            miRs.Close
'
'        Case 3
'            If Not TieneLaboral Then Exit Sub
'
'            cad_meua = "SELECT * FROM tipocontrato ORDER BY idcontrato"
'            Set miRs = New ADODB.Recordset
'            miRs.Open cad_meua, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'            While Not miRs.EOF
'                Combo1(Index).AddItem miRs!desccontrato
'                Combo1(Index).ItemData(Combo1(Index).NewIndex) = miRs!idcontrato
'                miRs.MoveNext
'            Wend
'
'            miRs.Close
'    End Select
'
''    If Index <> 2 Then 'excepte per a ibanpais sempre seleccione la 1ª opció per defecte
''        Combo1(Index).ListIndex = 0
''    Else
''        'per defecte seleccione ES
''        PosicionarCombo Combo1(Index), 724
''        For j = 0 To Combo1(Index).ListCount - 1
'''            If Combo1(Index).ItemData(j) = 724 Then
'''                Combo1(Index).ListIndex = j
'''                Exit For
'''            End If
'''        Next j
''    End If
''    ' **************************************************************
'End Sub
'

Private Sub BotonImprimir()
'Dim cadParam As String
'Dim cadFormula As String
'
'    'Añadir el parametro de Empresa
'    cadParam = "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
'    'selecciona solo los empleados de esa empresa
'    cadFormula = "{empleado.codempre}=" & vEmpresa.codEmpre
'    With frmImprimir
'        .FormulaSeleccion = cadFormula
'        .OtrosParametros = cadParam
'        .NumeroParametros = 1 'Solo parametro de la empresa
'        .SoloImprimir = False
'        .Opcion = 13 'Opcionlistado
'        .Show vbModal
'    End With
End Sub


Private Sub printNou()
    
'    With frmImprimir2
'        .cadTabla2 = "empleado"
'        .Informe2 = "rEmpleados.rpt"
'        If CadB <> "" Then
'            .cadRegSelec = SQL2SF(CadB)
'        Else
'            .cadRegSelec = ""
'        End If
'        .cadRegActua = POS2SF(Data1, Me)
'        .cadTodosReg = ""
'        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomEmpre & "'|pOrden={empleado.apeemple}|"
'        .NumeroParametros2 = 2
'        .MostrarTree2 = False
'        .InfConta2 = False
'        .ConSubInforme2 = False
'
'        .Show vbModal
'    End With
End Sub


'Private Sub InsertaItem(H As Integer, FI As Date, FF As Date)
'Dim IT As ListItem
'    Set IT = ListView1.ListItems.Add()
'    IT.Text = Format(FI, "dd/mm/yyyy")
'    FF = DateAdd("d", -1, FF)
'    If FI <> FF Then
'        IT.SubItems(1) = Format(FF, "dd/mm/yyyy")
'    Else
'        IT.SubItems(1) = ""
'    End If
'    miSQL = DevuelveDesdeBD("nomhorario", "horarios", "idhorario", CStr(kCampo), "N")
'    IT.SubItems(2) = miSQL
'
'    If Now >= FI Then
'        If Format(Now, "dd/mm/yyyy") <= FF Then
'            IT.EnsureVisible
'            IT.Selected = True
'            IT.Bold = True
'            IT.ListSubItems(2).Bold = True
'            IT.ListSubItems(1).ForeColor = vbBlue
'            IT.ListSubItems(2).ForeColor = vbBlue
'            IT.ForeColor = vbBlue
'            Set ListView1.SelectedItem = IT
'        End If
'    End If
'End Sub
'
'Private Sub CargaCalendario()
'Dim F As Date
'Dim F2 As Date
'    ListView1.ListItems.Clear
'    Set miRs = New ADODB.Recordset
'    F = DateAdd("d", -14, Now)
'    F2 = DateAdd("m", 1, Now)
'    miSQL = "Select * from calendarioT where fecha >= '" & Format(F, FormatoFecha) & "'"
'    miSQL = miSQL & " and fecha <= '" & Format(F2, FormatoFecha) & "' and idtrabajador =" & Data1.Recordset!idTrabajador
'    miSQL = miSQL & " order by fecha"
'    miRs.Open miSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    kCampo = 0
'    F = FechaIni
'    While Not miRs.EOF
'        If kCampo <> miRs!idhorario Then
'            If kCampo > 0 Then InsertaItem kCampo, F, miRs!fecha
'            F = miRs!fecha
'            kCampo = miRs!idhorario
'        End If
'        F2 = miRs!fecha
'        miRs.MoveNext
'    Wend
'    If kCampo > 0 Then
'        If F <> F2 Then InsertaItem kCampo, F, F2
'    End If
'    miRs.Close
'
'    Set miRs = Nothing
'End Sub



Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim B As Boolean
Dim tots As String
On Error GoTo ECarga

    'b = DataGridAux(Index).Enabled
    'DataGridAux(Index).Enabled = False
    
    AdoAux(Index).ConnectionString = conn
    AdoAux(Index).RecordSource = MontaSQLCarga(Index, enlaza)
    AdoAux(Index).CursorType = adOpenDynamic
    AdoAux(Index).LockType = adLockPessimistic
    DataGridAux(Index).ScrollBars = dbgNone
    AdoAux(Index).Refresh
    Set DataGridAux(Index).DataSource = AdoAux(Index)
    
    DataGridAux(Index).AllowRowSizing = False
    DataGridAux(Index).RowHeight = 290
    If PrimeraVez Then
        DataGridAux(Index).ClearFields
        DataGridAux(Index).ReBind
        DataGridAux(Index).Refresh
    End If
    
    'DataGridAux(Index).Enabled = b
    PrimeraVez = False
    
    Select Case Index
        Case 0 'Viajeros
            'si es visible|control|tipo campo|nombre campo|ancho control|formato campo|
'            tots = "N||||0|;S|txtAux(1)|T|NºLinea|800|;" 'numexped,numlinea
            tots = "N||||0|;" 'idcal
            tots = tots & "S|txtAux(1)|T|Fecha|1100|;S|txtAux(2)|T|Descripcion|2800|;" 'nombre, apellido
            arregla tots, DataGridAux(Index), Me
'            DataGridAux(Index).Columns(4).Alignment = dbgCenter
    End Select
    DataGridAux(Index).ScrollBars = dbgAutomatic
    
'    If Not AdoAux(Index).Recordset.EOF Then
'        DataGridAux_RowColChange Index, 1, 1
'    Else
''        LimpiarCamposFrame Index
'    End If
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub



Private Function MontaSQLCarga(Index As Integer, enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
Dim Tabla As String
    
    Select Case Index
        Case 0 'Vi
            SQL = "SELECT * FROM calendariof WHERE idcal = "
            If enlaza Then
                SQL = SQL & Data1.Recordset.Fields(0)
            Else
                SQL = SQL & "-1"
            End If
            
            If Me.chkActual.Value = 1 Then
                SQL = SQL & " AND fecha >='" & Format(vEmpresa.FechaInicio, FormatoFecha)
                SQL = SQL & "' AND fecha <='" & Format(vEmpresa.FechaFin, FormatoFecha) & "'"
            End If
            SQL = SQL & " ORDER BY fecha"
    End Select
    MontaSQLCarga = SQL
End Function









Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vTabla As String
Dim anc As Single
Dim I As Integer
    
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    'If ModificaLineas = 2 Then Exit Sub
    ModoLineas = 1 'Ponemos Modo Añadir Linea
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Cabecera
        cmdAceptar_Click
        'No se ha insertado la cabecera
        If ModoLineas = 0 Then Exit Sub
'        'si la cabecera no esta insertada salir
'        If DevuelveDesdeBD("codprove", "proveedo", "codprove", text1(0).Text, "N") = "" Then
'            Exit Sub
'        End If
    End If
    

    PonerModo 5

    'Situamos el grid al final
    AnyadirLinea DataGridAux(Index), AdoAux(Index)

    anc = DataGridAux(Index).Top
    If DataGridAux(Index).Row < 0 Then
        anc = anc + 220
    Else
        anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 15
    End If
    
    LLamaLineas Index, ModoLineas, anc
    
    Select Case Index

            
        Case 0 'Viajeros
            txtAux(0).Text = Text1(0).Text 'numexped
            txtAux(1).Text = ""
            txtAux(2).Text = ""
            
           
           
            PonerFoco txtAux(1)
    End Select
End Sub


Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim I As Integer
    Dim J As Integer
    
    ModoLineas = 2 'Modificar llínia
    
    If Modo = 4 Then 'Modificar Cabecera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
    
  
    PonerModo 5
    
    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub

'    Me.lblIndicador.Caption = "MODIFICAR LINEA"
    
    If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
        I = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
        DataGridAux(Index).Scroll 0, I
        DataGridAux(Index).Refresh
    End If
      
    anc = DataGridAux(Index).Top
    If DataGridAux(Index).Row < 0 Then
        anc = anc + 210
    Else
        anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 15
    End If

    Select Case Index
'        Case 0 'cuentas bancarias
'            For j = 0 To 2
'                txtAux(j).Text = DataGridAux(Index).Columns(j).Text
'            Next j
'
'            SelComboBool AdoAux(Index).Recordset!codNacio, cmbAux(0)
'
'            For j = 5 To 7
'                txtAux(j).Text = DataGridAux(Index).Columns(j).Text
'            Next j
'
'            SelComboBool AdoAux(Index).Recordset!ctactiva, cmbAux(9)
'            SelComboBool AdoAux(Index).Recordset!ctaprpal, cmbAux(10)
'            txtAux(11).Text = DataGridAux(Index).Columns("direccio").Text
'            txtAux(12).Text = DataGridAux(Index).Columns("observac").Text
'            BloquearTxt txtAux(11), False
'            BloquearTxt txtAux(12), False
            
'        Case 1 'departamentos
'            For j = 13 To 17
'                txtAux(j).Text = DataGridAux(Index).Columns(j - 13).Text
'            Next j
'
'            SelComboBool AdoAux(Index).Recordset!facturac, cmbAux(18)
'            SelComboBool AdoAux(Index).Recordset!document, cmbAux(19)
'            SelComboBool AdoAux(Index).Recordset!princpal, cmbAux(20)
'
'            For i = 21 To 24
'                BloquearTxt txtAux(i), False
'            Next i
            
        Case 0 'viajeros
            For J = 0 To 2
                txtAux(J).Text = DataGridAux(Index).Columns(J).Text
            Next J
            'cmbAux(0).ListIndex = DataGridAux(Index).Columns(j).Text
            BloquearTxt txtAux(1), True 'bloquea numlinea
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    Select Case Index
'        Case 0 'cuentas bancarias
'            PonerFoco txtAux(3)
'        Case 1 'departamentos
'            PonerFoco txtAux(16)
        Case 0 'Viajeros
            PonerFoco txtAux(2)
    End Select
End Sub


'Private Sub BotonImprimirLinea(Index As Integer)
'Dim SQL As String, tabla As String
'Dim cadParam As String, cadFormula As String
'Dim cadSelect As String
'Dim numParam As Byte
'Dim OpcionListado As Integer
'On Error GoTo EImprimirLin
'
'    Select Case Index
'        Case 2 'COMISIONES/ PRODUCTOS
'            If Text1(0).Text = "" Then Exit Sub
'            'no se llama a frmListado, sino directamente a Imprimir
'            'inicializamos aqui parametros y formula seleccion
'            tabla = "provprod"
'            OpcionListado = 15
'
'            'Añadir el parametro de Empresa
'            cadParam = "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
'            numParam = 1
'            cadFormula = "{provprod.codprove}=" & Text1(0).Text
'            cadSelect = cadFormula
'
'             'Comprobar si hay registros a Mostrar antes de abrir el Informe
'            SQL = "Select count(*) FROM " & tabla
'            If cadSelect <> "" Then
'                cadSelect = QuitarCaracterACadena(cadSelect, "{")
'                cadSelect = QuitarCaracterACadena(cadSelect, "}")
'                SQL = SQL & " WHERE " & cadSelect
'            End If
'            If RegistrosAListar(SQL) = 0 Then
'                MsgBox "No hay datos para mostrar en el Informe.", vbInformation
'                Exit Sub
'            End If
'    End Select
'
''            LlamarImprimir
'    With frmImprimir
'        .FormulaSeleccion = cadFormula
'        .OtrosParametros = cadParam
'        .NumeroParametros = numParam
'        .SoloImprimir = False
'        .Opcion = OpcionListado
'        .Show vbModal
'    End With
'EImprimirLin:
'    If Err.Number <> 0 Then MuestraError Err.Number, "Imprimir Linea", Err.Description
'End Sub




Private Sub BotonEliminarLinea(Index As Integer)
Dim SQL As String, Cad As String
Dim Eliminar As Boolean
    On Error GoTo Error2

    ModoLineas = 3 'Posem Modo Eliminar Llínia
    
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
 
    PonerModo 5

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    'If Not SepuedeBorrar(Index) Then Exit Sub

    Eliminar = False
    
    Select Case Index
'        Case 0 'cltebanc
'            SQL = "Seguro que desea eliminar la Cuenta: "
'            SQL = SQL & Format(AdoAux(Index).Recordset!CodBanco, "0000") & "-" & Format(AdoAux(Index).Recordset!CodSucur, "0000") & "-" & Format(AdoAux(Index).Recordset!digcontr, "00") & "-" & Format(AdoAux(Index).Recordset!ctabanco, "0000000000")
'            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
'                Eliminar = True
'                SQL = "DELETE FROM provbanc"
'                SQL = SQL & ObtenerWhereCab(True) & " AND numlinea= " & AdoAux(Index).Recordset!numlinea
''                TerminaBloquear
''                Conn.Execute SQL
''                If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
''                CargaGrid Index, True
''                SituarTab (NumTabMto)
''                SSTab1.Tab = 1
''                SSTab2.Tab = NumTabMto
'            End If
'        Case 1 'departamentos
''            If Not SepuedeBorrar(Index) Then Exit Sub
'            SQL = "Seguro que desea eliminar el Departamento: "
'            SQL = SQL & AdoAux(Index).Recordset!nomdepto
'            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
'                Eliminar = True
'                SQL = "DELETE FROM provdpto"
'                SQL = SQL & ObtenerWhereCab(True) & " AND numlinea= " & AdoAux(Index).Recordset!numlinea
''                TerminaBloquear
''                Conn.Execute SQL
''                If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
''                CargaGrid Index, True
''                SituarTab (NumTabMto)
''                SSTab1.Tab = 1
''                SSTab2.Tab = NumTabMto
'            End If

        Case 0 'Viajeros
'            SQL = "Seguro que desea eliminar la linea: " & Format(DBLet(AdoAux(Index).Recordset!numlinea), FormatoCampo(txtAux(1)))
'            cad = DBLet(AdoAux(Index).Recordset!apepasaj)
'            If DBLet(AdoAux(Index).Recordset!nompasaj) <> "" Then
'                If cad <> "" Then cad = cad & ", " & AdoAux(Index).Recordset!nompasaj
'            End If
'            SQL = SQL & vbCrLf & cad
            SQL = "¿Seguro que desea eliminar la linea?"
            SQL = SQL & vbCrLf & "Fecha: " & AdoAux(0).Recordset!Fecha
            SQL = SQL & vbCrLf & "Descripcion: " & AdoAux(0).Recordset!descripcion
            SQL = SQL & vbCrLf & Cad
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
                SQL = "DELETE FROM calendariof where idcal =" & Me.Data1.Recordset.Fields(0)
                SQL = SQL & " and fecha = '" & Format(AdoAux(0).Recordset!Fecha, FormatoFecha) & "'"
                Eliminar = True
                'SQL = SQL & ObtenerWhereCab(True) & " AND numlinea= " & AdoAux(Index).Recordset!numlinea
            End If
    End Select
    
    If Eliminar Then
        TerminaBloquear
        If EjecutaSQL(SQL) Then
            CargaGrid Index, True
            
            'SituarDataTrasEliminar AdoAux(Index), NumRegElim
            'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
        End If
    End If
    
    ModoLineas = 0
    PosicionarData
    
    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub


Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim B As Boolean

    On Error GoTo ELLamaLin

    DeseleccionaGrid DataGridAux(Index)
    'PonerModo xModo + 1
    
    B = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Lineas
    Select Case Index
        Case 0 'Viajeros
            For jj = 1 To 2
                txtAux(jj).Top = alto
                txtAux(jj).Visible = B
            Next jj
 '           cmbAux(0).Top = alto - 15
 '           cmbAux(0).Visible = b
    End Select
ELLamaLin:
    Err.Clear
End Sub



Private Sub CargaCalendario()
Dim F As Date
Dim F2 As Date
    ListView1.ListItems.Clear
    Set miRs = New ADODB.Recordset
    F = DateAdd("m", -2, Now)
    F2 = DateAdd("m", 2, Now)
    miSQL = "Select * from calendariol where fecha >= '" & Format(F, FormatoFecha) & "'"
    miSQL = miSQL & " and fecha <= '" & Format(F2, FormatoFecha) & "' and idcal =" & Data1.Recordset!idCal
    
    miSQL = miSQL & " order by fecha"
    miRs.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    kCampo = 0
    F = vEmpresa.FechaInicio
    While Not miRs.EOF
        If kCampo <> miRs!IdHorario Then
            If kCampo > 0 Then InsertaItem kCampo, F, miRs!Fecha
            F = miRs!Fecha
            kCampo = miRs!IdHorario
        End If
        F2 = miRs!Fecha
        miRs.MoveNext
    Wend
    If kCampo > 0 Then
        If F <> F2 Then InsertaItem kCampo, F, F2
    End If
    miRs.Close
    
    Set miRs = Nothing
End Sub


Private Sub InsertaItem(H As Integer, FI As Date, FF As Date)
Dim IT As ListItem
    Set IT = ListView1.ListItems.Add()
    IT.Text = Format(FI, "dd/mm/yyyy")
    FF = DateAdd("d", -1, FF)
    If FI <> FF Then
        IT.SubItems(1) = Format(FF, "dd/mm/yyyy")
    Else
        IT.SubItems(1) = ""
    End If
    miSQL = DevuelveDesdeBD("nomhorario", "horarios", "idhorario", CStr(kCampo), "N")
    IT.SubItems(2) = miSQL
 
    If Now >= FI Then
        If Format(Now, "dd/mm/yyyy") <= FF Then
            IT.EnsureVisible
            IT.Selected = True
            IT.Bold = True
            IT.ListSubItems(1).Bold = True
            IT.ListSubItems(2).Bold = True
            IT.ListSubItems(1).ForeColor = vbBlue
            IT.ListSubItems(2).ForeColor = vbBlue
            IT.ForeColor = vbBlue
            Set ListView1.SelectedItem = IT
        End If
    End If
End Sub





Private Sub InsertarLinea()
'Inserta registro en las tablas de Lineas: provbanc, provdpto
Dim nomFrame As String
Dim B As Boolean
On Error Resume Next

'    Select Case NumTabMto
'        Case 0: nomFrame = "FrameAux0" 'viajeros
'        Case 1: nomFrame = "FrameAux1" 'Departamentos
'        Case 2: nomFrame = "FrameAux2" 'Productos
'    End Select
    
    nomFrame = "FrameAux"
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomFrame) Then
'            If NumTabMto = 0 Then
'                'si ha seleccionat "cuenta principal", actualise totes les atres a "no"
'                If (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) Then
'                    ActualisaCtaprpal (txtAux(2).Text)
'                End If
'            End If
            B = BLOQUEADesdeFormulario2(Me, Data1, 1)
            CargaGrid 0, True
            If B Then BotonAnyadirLinea 0
           ' SituarTab (0)
        End If
    End If
End Sub


Private Sub ModificarLinea()
'Modifica registro en las tablas de Lineas: provbanc, provdpto
Dim nomFrame As String
Dim v As String
On Error GoTo EModificarLin

'    Select Case NumTabMto
'        Case 0: nomFrame = "FrameAux0" 'cuentas Bancarias
'        Case 1: nomFrame = "FrameAux1" 'Departamentos
'        Case 2: nomFrame = "FrameAux2" 'Productos
'    End Select
    nomFrame = "FrameAux"
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomFrame) Then
            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
'            If NumTabMto = 0 Then
'                'si ha seleccionat "cuenta principal", actualise totes les atres a "no"
'                If (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) Then
'                    ActualisaCtaprpal (txtAux(2).Text)
'                End If
'            End If
            v = "'" & Format(AdoAux(0).Recordset.Fields(1), FormatoFecha) & "'" 'el 2 es el nº de llinia
            ModoLineas = 0
            CargaGrid 0, True
'            SituarTab (0)
'            SSTab1.Tab = 1
'            SSTab2.Tab = NumTabMto
'            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'per a que es quede en modificar
'            PonerModo 4
            DataGridAux(0).SetFocus
            AdoAux(0).Recordset.Find ("fecha =" & v)

            LLamaLineas 0, 0
        End If
    End If
EModificarLin:
    If Err.Number <> 0 Then MuestraError Err.Number, "Modificar Linea", Err.Description
End Sub



Private Function DatosOkLlin(nomFrame As String) As Boolean
Dim B As Boolean
On Error GoTo EDatosOKLlin

    DatosOkLlin = False
        
    B = CompForm2(Me, 2, nomFrame) 'Comprobar formato datos ok
    If Not B Then Exit Function
    
    If CDate(txtAux(1).Text) > vEmpresa.FechaFin Then
        If MsgBox("Fecha fuera de temporada.  ¿Continuar?", vbQuestion + vbYesNo) <> vbYes Then Exit Function
    End If
    
    
'    Select Case NumTabMto
'    End Select
         
    DatosOkLlin = B
EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub



Private Sub txtAux_LostFocus(Index As Integer)
    txtAux(Index).Text = Trim(txtAux(Index).Text)
    If txtAux(Index).Text = "" Then Exit Sub
    
    If Index = 1 Then
        If Not EsFechaOK(txtAux(Index)) Then
            MsgBox "Fecha incorrecta: " & txtAux(Index).Text, vbExclamation
            txtAux(Index).Text = ""
            PonerFoco txtAux(Index)
        End If
    End If
End Sub
