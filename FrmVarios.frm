VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form FrmVarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7920
   Icon            =   "FrmVarios.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameGenSemana 
      Height          =   3615
      Left            =   720
      TabIndex        =   36
      Top             =   1200
      Visible         =   0   'False
      Width           =   5055
      Begin VB.ComboBox cboSeccion 
         Height          =   315
         Left            =   360
         TabIndex        =   52
         Text            =   "Combo1"
         Top             =   1080
         Width           =   4215
      End
      Begin VB.CommandButton cmdCalcularHorasTrabajadasSemana 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2040
         TabIndex        =   40
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   3480
         TabIndex        =   41
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   6
         Left            =   3360
         TabIndex        =   39
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   38
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Sección"
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   64
         Top             =   2520
         Width           =   4335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Sección"
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   53
         Top             =   840
         Width           =   615
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   6
         Left            =   3120
         Picture         =   "FrmVarios.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Inicio"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   43
         Top             =   1920
         Width           =   615
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   5
         Left            =   960
         Picture         =   "FrmVarios.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Fin"
         Height          =   255
         Index           =   7
         Left            =   2520
         TabIndex        =   42
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Generación horas semanales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   240
         Index           =   1
         Left            =   1080
         TabIndex        =   37
         Top             =   360
         Width           =   3060
      End
   End
   Begin VB.Frame FrameAjusteparadas 
      Height          =   3735
      Left            =   120
      TabIndex        =   54
      Top             =   240
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3000
         TabIndex        =   63
         Top             =   3120
         Width           =   1215
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   495
         Left            =   4800
         TabIndex        =   62
         Top             =   2040
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   873
         _Version        =   327681
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtDecimal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   3360
         TabIndex        =   59
         Text            =   "Text2"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtDecimal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "Text2"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   4440
         TabIndex        =   56
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Paradas"
         Height          =   255
         Index           =   13
         Left            =   3360
         TabIndex        =   61
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Total"
         Height          =   195
         Index           =   12
         Left            =   480
         TabIndex        =   60
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label Label4 
         Caption         =   "Asignar vacaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   57
         Top             =   1080
         Width           =   5295
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ajustes paradas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Index           =   3
         Left            =   240
         TabIndex        =   55
         Top             =   360
         Width           =   2805
      End
   End
   Begin VB.Frame FramePedirMes 
      Height          =   2295
      Left            =   1440
      TabIndex        =   44
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   50
         Text            =   "Text2"
         Top             =   1080
         Width           =   975
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         Index           =   0
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdDevuelveMEs 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   1440
         TabIndex        =   47
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   2760
         TabIndex        =   46
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Año"
         Height          =   195
         Index           =   10
         Left            =   2040
         TabIndex        =   51
         Top             =   840
         Width           =   300
      End
      Begin VB.Label Label3 
         Caption         =   "Mes"
         Height          =   195
         Index           =   9
         Left            =   480
         TabIndex        =   49
         Top             =   840
         Width           =   300
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   45
         Top             =   240
         Width           =   3885
      End
   End
   Begin VB.Frame FrAsignaVaca 
      Height          =   5175
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   5655
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   2880
         Width           =   2295
      End
      Begin VB.CommandButton cmdVacas 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2760
         TabIndex        =   33
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   4
         Left            =   3960
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   3
         Left            =   2040
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optPeriodo 
         Caption         =   "Option1"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   25
         Top             =   2280
         Width           =   2295
      End
      Begin VB.OptionButton optPeriodo 
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   24
         Top             =   1920
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   4200
         TabIndex        =   21
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Mes"
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   35
         Top             =   2880
         Width           =   495
      End
      Begin VB.Image imgFechasSueltas 
         Height          =   255
         Left            =   1320
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Calendario"
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   32
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label3 
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   31
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Fin"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   30
         Top             =   1200
         Width           =   495
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   4
         Left            =   3720
         Picture         =   "FrmVarios.frx":0122
         ToolTipText     =   "Buscar fecha"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Inicio"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   28
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Intervalo"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   27
         Top             =   840
         Width           =   735
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1800
         Picture         =   "FrmVarios.frx":01AD
         ToolTipText     =   "Buscar fecha"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Asignar vacaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label Label3 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   5295
      End
   End
   Begin VB.Frame FrModiCal 
      Height          =   2295
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   4215
      Begin VB.CommandButton cmdCambioCalen 
         Caption         =   "Cambiar"
         Height          =   375
         Left            =   1560
         TabIndex        =   15
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   2760
         TabIndex        =   16
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   2
         Left            =   2520
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "a partir de la fecha:"
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
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Modificar el calendario del trabajador"
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
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   3495
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   2040
         Picture         =   "FrmVarios.frx":0238
         ToolTipText     =   "Buscar fecha"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Inicio"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   3855
      End
   End
   Begin VB.Frame FrameModiHora 
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton cmdModifiHorario 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   1
         Left            =   3600
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   600
         Width           =   3615
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   3840
         TabIndex        =   4
         Top             =   1680
         Width           =   975
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   3240
         Picture         =   "FrmVarios.frx":02C3
         ToolTipText     =   "Buscar fecha"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   720
         Picture         =   "FrmVarios.frx":034E
         ToolTipText     =   "Buscar fecha"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Inicio"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Modificar el horario:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Fin"
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   9
         Top             =   1080
         Width           =   375
      End
   End
   Begin VB.Frame FrameVerImg 
      Height          =   6855
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cerrar"
         Height          =   375
         Index           =   1
         Left            =   4560
         TabIndex        =   12
         Top             =   6360
         Width           =   975
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4575
      End
      Begin VB.Image imgt 
         Height          =   5500
         Left            =   120
         Top             =   600
         Width           =   4800
      End
   End
End
Attribute VB_Name = "FrmVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Integer
    ' 0.- Pide datos para Modificar el horario de un trabajador para un dia
    ' 1.- Ver imagen en grande
    ' 2.- Se ha modificado el calendario en el trabajador. Vamos a preguntar
    '        la fecha a partir de la cual el cambio es efectivo
    
    ' 3.- Asignar vacciones para un trabajador
    
    ' 4.- Pedir D/H fehas para proceso generacion Horas semanales
    
    ' 5.- pedir Mes
    
    ' 6.- Ajuste paradas en procesar fecha
    
    
    ' 7. Generacion de horas semanales para FRUXERESA (SOLO alzicoop, obviamente)
    
    
Public Parametros As String

Private WithEvents frmc As frmCal
Attribute frmc.VB_VarHelpID = -1


Dim FInicioSeccion As Date
Dim i As Integer
Dim cad As String

Private Sub Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub




Private Sub cboSeccion_Click()

    If Me.cboSeccion.ListIndex < 0 Then Exit Sub
    cboSeccion.Tag = 0  'Puede ir por ms de una semana
    
    
    If Opcion = 4 Then
        '-------------------  Proceso nominas
    
            cad = DevuelveDesdeBD("nominas", "secciones", "idseccion", Me.cboSeccion.ItemData(cboSeccion.ListIndex), "N")
            cboSeccion.Tag = Val(cad)
            FInicioSeccion = "01/01/2001"
            cad = DevuelveDesdeBD("max(fechafin)", "jornadassemanalesproceso", "seccion", Me.cboSeccion.ItemData(cboSeccion.ListIndex), "N")
            If cad <> "" Then
                Me.txtFecha(5).Text = Format(cad, "dd/mm/yyyy")
                Me.txtFecha(5).Text = DateAdd("d", 1, CDate(Me.txtFecha(5).Text))
                FInicioSeccion = CDate(txtFecha(5).Text)
                
                'seccion.nominas =1
                If cboSeccion.Tag = 1 Then
                    'Sera desde hasta el domingo de esa semana
                    i = Format(CDate(Me.txtFecha(5).Text), "w", vbMonday)
                    i = 7 - i
                    Me.txtFecha(6).Text = Format(DateAdd("d", i, CDate(Me.txtFecha(5).Text)), "dd/mm/yyyy")
                Else
                    Me.txtFecha(6).Text = ""
                End If
            End If
        
    Else
        'Solo ALzira
        'Es generar datos para FRUXERESA
        'ESTOY AQUIIIIIII
            cad = DevuelveDesdeBD("nominas", "secciones", "idseccion", Me.cboSeccion.ItemData(cboSeccion.ListIndex), "N")
            cboSeccion.Tag = Val(cad)
            FInicioSeccion = "01/01/2001"
            cad = DevuelveDesdeBD("max(fechafin)", "jornadassemanalesproceso", "seccion", Me.cboSeccion.ItemData(cboSeccion.ListIndex), "N")
            If cad <> "" Then
                Me.txtFecha(5).Text = Format(cad, "dd/mm/yyyy")
                Me.txtFecha(5).Text = DateAdd("d", 1, CDate(Me.txtFecha(5).Text))
                FInicioSeccion = CDate(txtFecha(5).Text)
                
                'seccion.nominas =1
                If cboSeccion.Tag = 1 Then
                    'Sera desde hasta el domingo de esa semana
                    i = Format(CDate(Me.txtFecha(5).Text), "w", vbMonday)
                    i = 7 - i
                    Me.txtFecha(6).Text = Format(DateAdd("d", i, CDate(Me.txtFecha(5).Text)), "dd/mm/yyyy")
                Else
                    Me.txtFecha(6).Text = ""
                End If
        
            End If
  End If
End Sub

Private Sub cmdCambioCalen_Click()
    If txtFecha(2).Text = "" Then Exit Sub
    
    CadenaDesdeOtroForm = txtFecha(2).Text
    Unload Me
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdDevuelveMEs_Click()
    If Me.txtNumero(0).Text = "" Then Exit Sub
        
    CadenaDesdeOtroForm = "01/" & Format(Me.cboMes(0).ListIndex + 1, "00") & "/" & Me.txtNumero(0).Text
    Unload Me
End Sub

Private Sub cmdModifiHorario_Click()
    frmAsignaHorario.Opcion = 1
    frmAsignaHorario.FeIni = CDate(Me.txtFecha(0).Text)
    frmAsignaHorario.FeFin = CDate(Me.txtFecha(1).Text)
    frmAsignaHorario.Opcion = 1
    frmAsignaHorario.OtrosDatos = RecuperaValor(Parametros, 1) & "|" & RecuperaValor(Parametros, 4) & "|"
    frmAsignaHorario.Show vbModal
    Unload Me
    
End Sub



Private Sub cmdVacas_Click()

    'ASignar vacaciones
    If Me.txtFecha(3).Text <> "" Or txtFecha(4).Text <> "" Then
        If Me.txtFecha(4).Text = "" Or txtFecha(3).Text = "" Then
            MsgBox "Ponga desde / hasta fecha inicio", vbExclamation
            Exit Sub
        End If
        
        
        If CDate(Me.txtFecha(3).Text) > CDate(txtFecha(4).Text) Then
            MsgBox "Fecha inicio mayor que fin", vbExclamation
            Exit Sub
        End If
        
        
        If GenerarVacaciones(False) Then Unload Me
        
    Else
        If cmbFecha.ListIndex > 0 Then
            'Un mes de vacaciones
            If GenerarVacaciones(True) Then Unload Me
        Else
            MsgBox "Indique las vacaciones", vbExclamation
        End If
    End If
    
    
    
End Sub

Private Function DevuelveFechaUltimoProcesado() As Date
    Set miRsAux = New ADODB.Recordset
    'COMPROBAMOS QUE NO SE metan vacaciones en dias YA procesados
    'Veamos cual es la fecha de ultimo proceso
    DevuelveFechaUltimoProcesado = "01/01/2001"
    cad = "Select max(fecha) from marcajes"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then DevuelveFechaUltimoProcesado = miRsAux.Fields(0)
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
End Function

Private Function GenerarVacaciones(MesEntero As Boolean) As Boolean
Dim FechaUltimoProcesado As Date
Dim F1 As Date
Dim F2 As Date
    
    GenerarVacaciones = False
    FechaUltimoProcesado = DevuelveFechaUltimoProcesado
    'Voy a utlizar el txtfecha(3) aunque venga del mes. Luego borro
    '---------------------------------------------------------------
    If MesEntero Then
        'AHora veremos si las fecha de vacaciones YA hay dias procesados
        If Year(vEmpresa.FechaInicio) = Year(vEmpresa.FechaFin) Then
                'AÑOS NATURALES. Vamos, como en la conta
                i = Year(vEmpresa.FechaInicio)
                If Me.optPeriodo(1).Value Then i = i + 1
                    
        Else
            'Años partidos
            i = Year(vEmpresa.FechaFin)
            If Not (vEmpresa.FechaFin <= FechaUltimoProcesado) Then i = i + 1
                
        End If
        txtFecha(3).Text = "01/" & cmbFecha.ItemData(cmbFecha.ListIndex) & "/" & i
    End If
    If CDate(txtFecha(3).Text) <= FechaUltimoProcesado Then
        'Vuelvo a poner la txt en blanco
        If MesEntero Then txtFecha(3).Text = ""
        MsgBox "Fecha de incio esta ya procesada: " & FechaUltimoProcesado, vbExclamation
        Exit Function
    End If
    
    
    If MesEntero Then
        F1 = CDate("01/" & cmbFecha.ItemData(cmbFecha.ListIndex) & "/" & i)
        i = DiasMes(cmbFecha.ItemData(cmbFecha.ListIndex), i)
        F2 = i & "/" & cmbFecha.ItemData(cmbFecha.ListIndex) & "/" & Year(F1)
        
    Else
        'Son dias de intervalo
        F1 = CDate(txtFecha(3).Text)
        F2 = CDate(txtFecha(4).Text)
    End If

    cad = "UPDATE calendariot Set tipodia = 2 WHERE  idTrabajador =  " & Val(RecuperaValor(Parametros, 2))
    cad = cad & " AND fecha = '"
    While F1 <= F2
        EjecutaSQL cad & Format(F1, FormatoFecha) & "'"
        F1 = DateAdd("d", 1, F1)
    Wend
    
    GenerarVacaciones = True
End Function


Private Sub cmdCalcularHorasTrabajadasSemana_Click()
    
    
     'Muuuchas cosas a comprobar
    cad = ""
    If txtFecha(5).Text = "" Or txtFecha(6).Text = "" Then cad = "Ponga las fechas"
        
    If cboSeccion.ListIndex < 0 Then
        If vEmpresa.QueEmpresa = 4 Then
            cboSeccion.Tag = 1
            
            'Procesamos todas las secciones a la vez
            
            'FInicioSeccion = txtFecha(5).Text
           
            
        Else
            cad = "Falta sección" & vbCrLf & cad
        End If
    End If
    If cad <> "" Then
        MsgBox cad, vbExclamation
        Exit Sub
    End If


    'Faltara ver si esta la semana completa o es final de semana
    cad = ""
    If Not IsDate(txtFecha(5).Text) Or Not IsDate(txtFecha(6).Text) Then
        cad = "Fechas incorrectas"
    Else
        'VA POR SEMANAS
        If Me.cboSeccion.Tag = 1 Then
    
            If DateDiff("d", txtFecha(5).Text, txtFecha(6).Text) > 6 Then
                'VA por semanas
                 cad = "Mas de una semana seleccionada"
                
            ElseIf Format(CDate(txtFecha(5).Text), "ww", vbMonday) <> Format(CDate(txtFecha(6).Text), "ww", vbMonday) Then
                'Procesamos de semana en semana. Deben pertenecer a la misma semana
                cad = "Semanas distintas"
            Else
                'OK. Faltara ver mas casos
                
                
                'Agosto 2014
                '-----------------------------------------------------
                'Las semanas
                'La fecha inicio NO puede ser inferior a la que le corresponde, que esta guardado en FInicioSeccion
                If CDate(txtFecha(5).Text) < FInicioSeccion Then
                   cad = "Intervalo ya procesado. Inicio debe ser: " & FInicioSeccion
                Else
                    'Si no procesa desde qyue le corresponde
                    'veremos si hay datos pendientes de procesar
                     If CDate(txtFecha(5).Text) > FInicioSeccion Then
                          'Si la fecha es mayor a la que ponen como inicio del intervalor, veremos si ya hay procesados
                          'para esas fechas y esa seccion
                          cad = "fecha >=" & DBSet(FInicioSeccion, "F") & " AND fecha <" & DBSet(txtFecha(5).Text, "F")
                          cad = cad & " AND idtrabajador IN (select idtrabajador from trabajadores  "
                          
                          If cboSeccion.ListIndex > 0 Then cad = cad & " WHERE Seccion = " & cboSeccion.ItemData(cboSeccion.ListIndex)
                          
                          cad = cad & ") AND 1"
                          cad = DevuelveDesdeBD("count(*)", "jornadassemanalesalz", cad, "1")
                          If Val(cad) > 0 Then
                               cad = "Existen datos procesados entre: " & FInicioSeccion & " y " & txtFecha(5).Text
                          Else
                               cad = ""
                          End If
                    
                           'Veremos sy hay marcajes pendientes de procesar en ese periodo que es
                           'desde que le corresponde hasta donde empieza
                          If cad = "" Then
                             cad = "fecha >=" & DBSet(FInicioSeccion, "F") & " AND fecha <" & DBSet(txtFecha(5).Text, "F")
                             cad = cad & " AND idtrabajador IN (select idtrabajador from trabajadores where 1=1"
                             If cboSeccion.ListIndex > 0 Then cad = cad & " AND seccion=" & cboSeccion.ItemData(cboSeccion.ListIndex)
                             cad = cad & ") AND 1"
                             cad = DevuelveDesdeBD("count(*)", "marcajes", cad, "1")
                             If Val(cad) > 0 Then
                                  cad = "Existen marcajes entre: " & FInicioSeccion & " y " & txtFecha(5).Text & " que no entran en el intervalo para procesar"
                             Else
                                  cad = ""
                             End If
                        
                          End If
                       
                       
                          'Febrero 2015
                          If cad = "" Then
                                'Comrpobacion una. que el dia es menor que jueves
                                i = Weekday(txtFecha(5).Text, vbMonday)
                                If i >= 5 Then
                                    'Es viernes. Por lo tanto, el dia hasta tiene que ser domingo
                                    If Weekday(CDate(txtFecha(6).Text), vbMonday) <> 7 Then cad = "Fecha final de intervalo debe ser domingo. Proceso semana completo"
                                        
                                End If
                          End If
                    
                    End If  'de mayor que fecha inicioseccion
                    
                    If cad = "" Then
                           cad = "fecha >=" & DBSet(txtFecha(5).Text, "F") & " AND fecha <=" & DBSet(txtFecha(6).Text, "F")
                             cad = cad & " AND idtrabajador IN (select idtrabajador from trabajadores where 1=1"
                             If cboSeccion.ListIndex > 0 Then cad = cad & " AND seccion=" & cboSeccion.ItemData(cboSeccion.ListIndex)
                             cad = cad & ") AND correcto "
                             cad = DevuelveDesdeBD("count(*)", "marcajes", cad, "0")
                             If Val(cad) > 0 Then
                                  cad = "Existen marcajes INCORRECTOS entre: " & txtFecha(5).Text & " y " & txtFecha(6).Text
                             Else
                                  cad = ""
                             End If
                    End If
                    
                End If
                            
                            
                
                
            End If
        End If  'de por semanas
    End If
    If cad <> "" Then
    
        MsgBox cad, vbExclamation
        Exit Sub
    End If
    
    
   
    
    
    
    
    
    
    
    
    
    
    
    

    
    
    
    

    Screen.MousePointer = vbHourglass
    HazCalcularHorasTrabajadasSemana
    Screen.MousePointer = vbDefault
End Sub

Private Sub HazCalcularHorasTrabajadasSemana()
Dim ColTraba  As Collection
Dim N As Byte
Dim Salir As Boolean
        
        
        
        
        
        
    Set miRsAux = New ADODB.Recordset
    
    
    
    'Abril 2014
    'Una comprobacion
    'Luego lee los festivos del calendario, y estamos procesando una seccion
    'es decir NO dejo pasar si para la seccion hay mas de una calendario
    If cboSeccion.ListIndex >= 0 Then
        cad = "select distinct(idcal) from trabajadores where seccion=" & cboSeccion.ItemData(cboSeccion.ListIndex)
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic
        cad = ""
        While Not miRsAux.EOF
            cad = cad & "1"
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    Else
        cad = "1"
    End If
    If Len(cad) <> 1 Then
        If vEmpresa.QueEmpresa <> 5 Then
            MsgBox "Solo puede haber un calendario para la seccion. ", vbExclamation
            Exit Sub
        End If
    End If
    
    
    
    
    Set ColTraba = New Collection
    
    cad = "Select marcajes.idtrabajador FROM marcajes,trabajadores WHERE marcajes.idtrabajador=trabajadores.idtrabajador AND"
    cad = cad & TipoAlziraEntreFechas
    If cboSeccion.ListIndex >= 0 Then cad = cad & " AND seccion=" & cboSeccion.ItemData(cboSeccion.ListIndex)
    cad = cad & " GROUP BY 1"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    cad = ""
    While Not miRsAux.EOF
        cad = cad & ", " & miRsAux!idTrabajador
        NumRegElim = NumRegElim + 1
        If NumRegElim > 50 Then
            ColTraba.Add Mid(cad, 2)
            cad = ""
            NumRegElim = 0
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If NumRegElim > 0 Then ColTraba.Add Mid(cad, 2)
    
    '
    cad = "DELETE FROM tmphorastipoalzira WHERE codusu =" & vUsu.Codigo
    conn.Execute cad
    
    
    Salir = True
    If ColTraba.Count = 0 Then
        MsgBox "Ningun trabajador a procesar para el intervalo", vbExclamation
        Salir = False
        
    Else
    
        For N = 1 To ColTraba.Count
            If Me.cboSeccion.Tag = 1 Then
                'Proceso de horas. Horas normales, extra, extrucutrales...
                ProcesoCalculaHorasTipoAlzira " AND idtrabajador in (" & ColTraba.Item(N) & ")"
                
            Else
                'Proceso de calculo de horas por conteo (sums)
                'Este calculo es sencillo ya que los trabajadores de estas secciones
                'NO, repito NO, tripito NO,
                'llevan control de horas , por lo tanto, todas las horas trabajadas son de tipo 0
                MsgBox "Proceso en desarrollo", vbExclamation
                Exit Sub
            End If
        Next
        
        CadenaDesdeOtroForm = txtFecha(5).Text & "|" & txtFecha(6).Text & "|" & IIf(Me.cboSeccion.ListIndex < 0, 1, 0) & "|"
    End If
    
    Set miRsAux = Nothing
    
    If Salir Then Unload Me
End Sub




'Llevara UN list de trabajadores 23,35,36....
Private Sub ProcesoCalculaHorasTipoAlzira(ListaTrabajadores As String)
Dim Insert As String
Dim ColSabados As Collection
Dim AuxTra As String
Dim RT As ADODB.Recordset
Dim Fin As Boolean
    
Dim HN As Currency
Dim HI As Currency
Dim T1 As Currency
Dim T2 As Currency
    
Dim HorasExceso As Currency
Dim HorasmaximoNormalesDia As Integer
Dim HoraSabadoExtras As Date

    
    
        
    HorasmaximoNormalesDia = 9
    If vEmpresa.QueEmpresa = 4 Then HorasmaximoNormalesDia = 8
    
    'Hay empresas CASTELDUC, que hacen una compensacion en nomina a final de mes. Y ya esta. No van dia a dia
    If vEmpresa.CompensaHorasNominaMES Then HorasmaximoNormalesDia = vEmpresa.CompensaMES_HorasDia   'HorasmaximoNormalesDia = 10
        
        
    
    
    
    Insert = "INSERT INTO tmphorastipoalzira(codusu,idtrabajador,diasem,fecha,TipoHoras,horastrabajadas) "
    
    'Varios pasos
    'Primero los domingos y festivos ENTRAN Con todas las horas extras
    '----------------------------------------------------------------
    cad = 2 'HORA EXTRA
    cad = "select " & vUsu.Codigo & ", idtrabajador,date_format(fecha,'%w') diasem,fecha," & cad & ",horastrabajadas"
    cad = cad & " from marcajes where " & TipoAlziraEntreFechas
    cad = cad & ListaTrabajadores & " AND "
    
    
    conn.Execute Insert & cad & " date_format(fecha,'%w')=0" 'domingos
    
    conn.Execute Insert & cad & " Festivo = 1 and date_format(fecha,'%w')<>0"  'festivos que no sean domingos
    
    
    
        
        cad = 0 'HORA normales
        cad = "select " & vUsu.Codigo & ", idtrabajador,date_format(fecha,'%w') diasem,fecha," & cad & ",if(horastrabajadas>" & HorasmaximoNormalesDia & "," & HorasmaximoNormalesDia & ",horastrabajadas)"
        cad = cad & " from marcajes where " & TipoAlziraEntreFechas
        cad = cad & ListaTrabajadores & " AND "
        conn.Execute Insert & cad & " date_format(fecha,'%w') in (1,2,3,4,5) AND    Festivo = 0"
        
        'Las que se pasen de 9 (HorasmaximoNormalesDia) van a horas estrcutrales
        cad = 1 'HORA estrucutrales
        cad = "select " & vUsu.Codigo & ", idtrabajador,date_format(fecha,'%w') diasem,fecha," & cad & ",horastrabajadas-" & HorasmaximoNormalesDia
        cad = cad & " from marcajes where " & TipoAlziraEntreFechas
        cad = cad & ListaTrabajadores & " AND "
        cad = cad & " date_format(fecha,'%w') in (1,2,3,4,5) AND Festivo = 0 AND horastrabajadas>" & HorasmaximoNormalesDia
        conn.Execute Insert & cad
        
        
    
    
    'ALZRIRA
    'Los sabados, a partir de las 14:30 son extras
    '------------------------------------------------
    'Los sabados, que no sean festivos
    cad = "select fecha from marcajes where " & TipoAlziraEntreFechas
    cad = cad & ListaTrabajadores & " and date_format(fecha,'%w')=6 and festivo=0 GROUP BY 1"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set ColSabados = New Collection
    While Not miRsAux.EOF
        ColSabados.Add CStr(miRsAux!Fecha)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    If vEmpresa.QueEmpresa = 5 Then
        HoraSabadoExtras = "13:30:00"
    Else
        If vEmpresa.QueEmpresa = 4 Then
            HoraSabadoExtras = "14:00:00"
        Else
            HoraSabadoExtras = "14:30:00"
        End If
    End If
    If vEmpresa.HoraSabadoExtras <> "" Then HoraSabadoExtras = Format(CDate(vEmpresa.HoraSabadoExtras), "hh:nn:ss")
    
    
    For i = 1 To ColSabados.Count
                
            'Veremos que trabajadores tienen un fichaje mas alla de las 14:30(HoraSabadoExtras )
            cad = "select idtrabajador from  entradamarcajes where fecha=" & DBSet(ColSabados.Item(i), "F") & " and hora>" & DBSet(HoraSabadoExtras, "H")
            cad = cad & ListaTrabajadores & " GROUP BY 1"
            miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            AuxTra = ""
            While Not miRsAux.EOF
                AuxTra = AuxTra & ", " & miRsAux!idTrabajador
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            'YA tengo los trabadores que el SABADO ha trabajado mas alla de las HoraSabadoExtras
            
            'Los que NO han ido mas alla HoraSabadoExtras
                cad = 1 'HORA estrucutrales
                cad = "select " & vUsu.Codigo & ", idtrabajador,date_format(fecha,'%w') diasem,fecha," & cad & ",if(horastrabajadas>" & HorasmaximoNormalesDia & "," & HorasmaximoNormalesDia & ",horastrabajadas)"
                cad = cad & " from marcajes where fecha=" & DBSet(ColSabados.Item(i), "F") & " AND horastrabajadas>" & HorasmaximoNormalesDia
                cad = cad & ListaTrabajadores
                If AuxTra <> "" Then
                    AuxTra = Mid(AuxTra, 2)
                    'En auxtra estan los que han trabajado mas alla de las 14:30
                    cad = cad & " and not idtrabajador in  (" & AuxTra & ")"
    
                End If
                conn.Execute Insert & cad
                
                cad = 0 'HORA normales
                cad = "select " & vUsu.Codigo & ", idtrabajador,date_format(fecha,'%w') diasem,fecha," & cad & ",horastrabajadas"
                cad = cad & " from marcajes where fecha=" & DBSet(ColSabados.Item(i), "F") & " AND horastrabajadas<=" & HorasmaximoNormalesDia
                cad = cad & ListaTrabajadores
                If AuxTra <> "" Then
                    'En auxtra estan los que han trabajado mas alla de las  HoraSabadoExtras
                    cad = cad & " and not idtrabajador in  (" & AuxTra & ")"
                End If
                conn.Execute Insert & cad
                
          
                
            
            'Los que han trabajado el sabado REESTRUCTURAMOS sus horas
            If AuxTra <> "" Then
                Set RT = New ADODB.Recordset
                
                'REESTABLECEMOS LAS HORAS PARA AQUELLOS  han trabajado mas alla de las HoraSabadoExtras
                cad = "Select marcajes.*,ExcesoDefecto from marcajes,incidencias where idinci=IncFinal AND fecha=" & DBSet(ColSabados.Item(i), "F") & " and idtrabajador in  (" & AuxTra & ")"
                miRsAux.Open cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                While Not miRsAux.EOF
                    Debug.Print miRsAux!idTrabajador
                    'FALTA####
                    'If miRsAux!idTrabajador = 9 Then St op
                    HN = miRsAux!HorasTrabajadas
                    
                    HorasExceso = 0
                    
                    If vEmpresa.QueEmpresa = 4 Then
                        'En catadau, son normales las 8 primeras
                        HI = 0
                    
                    Else
                        If miRsAux!ExcesoDefecto = 0 Then
                            HI = 0
                        Else
                            'Horas incidencia
                            HI = miRsAux!HorasIncid
                            HN = HN - miRsAux!HorasIncid
                        End If
                    End If
                    
                    'Vere las fichadas de ese dia que superen las HoraSabadoExtras
                    cad = "Select * from entradamarcajes  where fecha=" & DBSet(miRsAux!Fecha, "F") & " AND idtrabajador =" & miRsAux!idTrabajador
                    cad = cad & "  and hora>" & DBSet(HoraSabadoExtras, "H") & "  and hora <='23:59:59' ORDER by hora desc"
                    RT.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    Fin = False
                    Do
                        T1 = CCur(DevuelveValorHora(RT!Hora))
                        RT.MoveNext
                        If RT.EOF Then
                            T2 = CCur(DevuelveValorHora(HoraSabadoExtras))
                        Else
                            T2 = CCur(DevuelveValorHora(RT!Hora))
                            RT.MoveNext
                        End If
                        If RT.EOF Then Fin = True
                        
                        
                        'Calculamos la diferencia
                        T2 = T1 - T2
                        If T2 >= 0.1 Then
                           HorasExceso = HorasExceso + T2
                           
                           'Si hay inicidencia se las sumo
                           If HI > 0 Then
                                If T2 >= HI Then
                                    T2 = T2 - HI
                                    HI = 0
                                Else
                                    HI = HI - T2
                                    T2 = 0
                                End If
                            End If
                            
                            If T2 > HN Then
                                MsgBox "NO puedo quitar mas horas", vbExclamation
                            Else
                                HN = HN - T2
                            End If
                           
                           
                           
                        End If
                    Loop Until Fin
                    
                    
                    'Insertamos en la tmp
                    'Normales
                    'tmphorastipoalzira(codusu,idtrabajador,diasem,fecha,TipoHoras,horastrabajadas)
                    cad = " VALUES (" & vUsu.Codigo & "," & miRsAux!idTrabajador & ",5," & DBSet(miRsAux!Fecha, "F") & ","
                    
                    cad = Insert & cad
                    
                    'Las normales
                        
                        conn.Execute cad & "0," & DBSet(HN, "N") & ")"   'normales
                        
                        If HorasExceso > 0 Then conn.Execute cad & "2," & DBSet(HorasExceso, "N") & ")"  'LS QUE VAN MAS ALLA SON EXTRA
                        If HI > 0 Then conn.Execute cad & "1," & DBSet(HI, "N") & ")"   'LS QUE VAN MAS ALLA SON estruc

                    
                    RT.Close
                    miRsAux.MoveNext 'siguiente trabajador
                Wend
                miRsAux.Close
            End If
        
        

                    
    
                
    Next i
    
        'Si los sabados fueran proceso normal seria el trozo de aqui abajo
        'COOPIC. Sabados, proceso normal  correinte
        'cad = 0 'HORA normales
        'cad = "select " & vUsu.Codigo & ", idtrabajador,date_format(fecha,'%w') diasem,fecha," & cad & ",if(horastrabajadas>" & HorasmaximoNormalesDia & "," & HorasmaximoNormalesDia & ",horastrabajadas)"
        'cad = cad & " from marcajes where " & TipoAlziraEntreFechas
        'cad = cad & ListaTrabajadores & " AND "
        'conn.Execute Insert & cad & " date_format(fecha,'%w') in (6) AND    Festivo = 0"
        
        'Las que se pasen de 9 (HorasmaximoNormalesDia) van a horas estrcutrales
        'cad = 1 'HORA estrucutrales
        'cad = "select " & vUsu.Codigo & ", idtrabajador,date_format(fecha,'%w') diasem,fecha," & cad & ",horastrabajadas-" & HorasmaximoNormalesDia
        'cad = cad & " from marcajes where " & TipoAlziraEntreFechas
        'cad = cad & ListaTrabajadores & " AND "
        'cad = cad & " date_format(fecha,'%w') in (6) AND Festivo = 0 AND horastrabajadas>" & HorasmaximoNormalesDia
        'conn.Execute Insert & cad
            
    
    
    
    
    cad = "Delete FROM tmphorastipoalzira WHERE codusu =" & vUsu.Codigo & " AND horastrabajadas=0"
    conn.Execute cad
    
    
    
    
    
    Label3(14).Caption = "Dias-nomina"
    Label3(14).Refresh
    
    
    
    
    
End Sub








Private Function TipoAlziraEntreFechas()
    TipoAlziraEntreFechas = " fecha Between " & DBSet(txtFecha(5).Text, "F") & " and " & DBSet(txtFecha(6).Text, "F")
End Function




Private Sub Command1_Click()
    i = 0
    CadenaDesdeOtroForm = ""
    If Me.txtDecimal(1).Text <> "" Then
        If ImporteFormateado(txtDecimal(1).Text) > 0 Then
            'Si es visible
            If Me.txtDecimal(0).Visible Then
                If ImporteFormateado(txtDecimal(1).Text) > ImporteFormateado(txtDecimal(0).Text) Then
                    MsgBox "No puede quitar mas horas de las que tiene trabajadas", vbExclamation
                    Exit Sub
                End If
            End If
            
        End If
        CadenaDesdeOtroForm = txtDecimal(1).Tag
    End If
    If Not Me.txtDecimal(0).Visible Then
        cad = "Desea asignar como horas de parada: " & txtDecimal(1).Text & "  a :" & vbCrLf & Parametros
        If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    Select Case Opcion
    Case 0
        PonerFoco Me.txtFecha(0)
    Case 3
        PonerFoco Me.txtFecha(3)
    Case 4
            
    Case 6
        PonerFoco Me.txtFecha(3)
    End Select
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer
Dim W As Integer

    Me.Icon = frmMain.Icon

    Me.FrameModiHora.Visible = False
    Me.FrameVerImg.Visible = False
    Me.FrModiCal.Visible = False
    FrAsignaVaca.Visible = False
    FrameGenSemana.Visible = False
    FramePedirMes.Visible = False
    FrameAjusteparadas.Visible = False
    Select Case Opcion
    Case 0
        H = Me.FrameModiHora.Height
        W = Me.FrameModiHora.Width
        Caption = "Modificar horario"
        Text1(0).Text = RecuperaValor(Parametros, 1)
        Me.txtFecha(0).Text = RecuperaValor(Parametros, 2)
        Me.txtFecha(1).Text = RecuperaValor(Parametros, 3)
        FrameModiHora.Visible = True
    Case 1
        H = Me.FrameVerImg.Height
        W = Me.FrameVerImg.Width
        Caption = "Ver imagen"
        FrameVerImg.Visible = True
    Case 2
        H = Me.FrModiCal.Height
        W = Me.FrModiCal.Width
        Caption = "Modificar calendario"
        Label1(5).Caption = RecuperaValor(Parametros, 1)
        Me.txtFecha(2).Text = Format(DateAdd("d", 1, Now), "dd/mm/yyyy")
        FrModiCal.Visible = True
        
    Case 3
        H = Me.FrAsignaVaca.Height
        W = Me.FrAsignaVaca.Width
        Caption = "Asignar calendario vacaciones"
        Label4(0).Caption = RecuperaValor(Parametros, 1)
        txtFecha(3).Text = ""
        txtFecha(4).Text = ""
        FrAsignaVaca.Visible = True
        optPeriodo(0).Caption = Format(vEmpresa.FechaInicio, "dd/mm/yyyy") & " - " & Format(vEmpresa.FechaFin, "dd/mm/yyyy")
        optPeriodo(1).Caption = Format(DateAdd("yyyy", 1, vEmpresa.FechaInicio), "dd/mm/yyyy") & " - "
        optPeriodo(1).Caption = optPeriodo(1).Caption & Format(DateAdd("yyyy", 1, vEmpresa.FechaFin), "dd/mm/yyyy")
        CargaCombo 0
        Me.imgFechasSueltas.Picture = frmPpal.imgListImages16.ListImages(3).Picture
    
    
    
    Case 4, 7
        'GENERA SEMANA
    
        Label3(14).Caption = ""  'indicador proceso
        PonerFrameVisible FrameGenSemana, H, W
        If Opcion = 4 Then
            Caption = "Horas laboral"
            Label4(1).Caption = "Generación horas semanales"
            Label4(1).ForeColor = &H80&
        Else
            Caption = "Horas Fruxeresa"
            Label4(1).Caption = "Gen HORAS Fruxeresa"
            Label4(1).ForeColor = &HC000C0
        End If
        
        
        cad = "DELETE FROM tmphorastipoalzira WHERE codusu =" & vUsu.Codigo
        conn.Execute cad
        
        
        cboSeccion.Clear
        Set miRsAux = New ADODB.Recordset
        cad = "select * from secciones where idseccion in (select seccion from trabajadores) "
        If Opcion = 4 Then cad = cad & " and nominas=1"
        cad = cad & " ORDER BY nombre"
        
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        

        While Not miRsAux.EOF
            cboSeccion.AddItem miRsAux!Nombre & " (" & miRsAux!IdSeccion & ")"
            cboSeccion.ItemData(cboSeccion.NewIndex) = miRsAux!IdSeccion
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        
        
        If vEmpresa.QueEmpresa = 4 Then
                    
            cad = "select seccion,max(fechaini) fi,max(fechafin) ff from jornadassemanalesproceso  group by 1"
            miRsAux.Open cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
            cad = ""
            
            While Not miRsAux.EOF
                If cad = "" Then
                    cad = Format(IIf(IsNull(miRsAux!FI), "01/01/1900", miRsAux!FI), "dd/mm/yyyy") & Format(IIf(IsNull(miRsAux!FF), "01/01/1900", miRsAux!FF), "dd/mm/yyyy")
                Else
                    CadenaDesdeOtroForm = Format(IIf(IsNull(miRsAux!FI), "01/01/19000", miRsAux!FI), "dd/mm/yyyy") & Format(IIf(IsNull(miRsAux!FF), "01/01/19000", miRsAux!FF), "dd/mm/yyyy")
                    If cad <> CadenaDesdeOtroForm Then
                        cad = "MAL"
                        miRsAux.MoveLast
                    End If
                End If
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            If cad <> "" Then
                If cad = "MAL" Then
                    MsgBox "Distintas fechas en la generacion por secciones", vbExclamation
                Else
                    
                    FInicioSeccion = CDate(Mid(cad, 11, 10))
                    FInicioSeccion = DateAdd("d", 1, FInicioSeccion)
                    txtFecha(5).Text = Format(FInicioSeccion, "dd/mm/yyyy")
                    i = Weekday(FInicioSeccion, vbMonday)
                    i = 7 - i
                    If i > 0 Then txtFecha(6).Text = Format(DateAdd("d", i, FInicioSeccion), "dd/mm/yyyy")
                    
                            
                    
                End If
            End If
            CadenaDesdeOtroForm = ""
        End If
        Set miRsAux = Nothing
    Case 5
        'Pedir mes
        PonerFrameVisible FramePedirMes, H, W
        Label4(2).Caption = "Listado horas mensual"
        Caption = "Listado"
        CargaCombo 1
        
    Case 6
        'Ajustes paradas
        Caption = "Ajuste paradas"
        PonerFrameVisible FrameAjusteparadas, H, W
        
        i = RecuperaValor(Parametros, 1)  'Vemos si es uno o mas de uno
        txtDecimal(0).Visible = i = 1
        Label3(12).Visible = i = 1
        
        txtDecimal(0).Text = RecuperaValor(Parametros, 2)
        cad = Trim(RecuperaValor(Parametros, 3))
        If cad = "" Then cad = "0"
        txtDecimal(1).Tag = ImporteFormateado(cad)
        txtDecimal(1).Text = Format(txtDecimal(1).Tag, FormatoImporte)
            
        Parametros = RecuperaValor(Parametros, 4) 'Dejo el listado de trabajadores
        If i = 1 Then
            Label4(4).Caption = Parametros
        Else
            Label4(4).Caption = "Trabajadores seleccionados: " & i
        End If
        
        
    End Select
    
    Me.cmdCancelar(Opcion).Cancel = True
    Me.Height = H + 520
    Me.Width = W + 240
End Sub



Private Sub PonerFrameVisible(ByRef Fr As Frame, HE As Integer, Wi As Integer)
        HE = Fr.Height
        Wi = Fr.Width
        Fr.Top = 0
        Fr.Left = 120
        Fr.Visible = True
End Sub

Private Sub frmc_Selec(vFecha As Date)
    'Me.Caption = vFecha & "   ind: " & imgFec(0).Tag
    cad = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub imgFec_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim Obj As Object

    Set frmc = New frmCal
    esq = imgFec(Index).Left
    dalt = imgFec(Index).Top
    
    ' *** adrede ***
    'If Index <> 49 Then
    '    esq = imgFec(Index).Left
    '    dalt = imgFec(Index).Top
    'Else
    '    esq = btnFec(Index).Left
    '    dalt = btnFec(Index).Top
    'End If
    ' *************

'    Set obj = imgFec(Index).Container
'
'    While imgFec(Index).Parent.Name <> obj.Name
'        esq = esq + obj.Left
'        dalt = dalt + obj.Top
'        Set obj = obj.Container
'    Wend
    
    ' *** adrede ***
  '  If Index <> 49 Then
        Set Obj = imgFec(Index).Container

        While imgFec(Index).Parent.Name <> Obj.Name
            esq = esq + Obj.Left
            dalt = dalt + Obj.Top
            Set Obj = Obj.Container
        Wend
'    Else
'        Set obj = btnFec(Index).Container
'
'        While btnFec(Index).Parent.Name <> obj.Name
'            esq = esq + obj.Left
'            dalt = dalt + obj.Top
'            Set obj = obj.Container
'        Wend
'
'    End If
    ' *************
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    ' *** adredre ***
'    If Index <> 49 Then 'dreta i baix
        frmc.Left = esq + imgFec(Index).Parent.Left + 30
        frmc.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40
'    Else 'esquerra i dalt
'        frmC.Left = esq + btnFec(Index).Parent.Left - frmC.Width + btnFec(Index).Width + 40
'        frmC.Top = dalt + btnFec(Index).Parent.Top - frmC.Height + menu - 25
'    End If
    ' ***************

    imgFec(0).Tag = Index '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtFecha(Index).Text <> "" Then frmc.NovaData = txtFecha(Index).Text
    ' ********************************************
    cad = ""
    frmc.Show vbModal
    Set frmc = Nothing
    
    If cad <> "" Then
        'Me.Caption = Cad
        txtFecha(Index).Text = cad
    
        ' *** repasar si el camp es txtAux o Text1 ***
        PonerFoco txtFecha(Index) '<===
        ' ********************************************
    End If
End Sub

Private Sub imgFechasSueltas_Click()
    With frmAsignaHorario
        If optPeriodo(0).Value Then
            i = 0
        Else
            i = 1
        End If
        
        .OtrosDatos = Parametros & DevuelveFechaUltimoProcesado & "|"
        .FeIni = DateAdd("yyyy", i, vEmpresa.FechaInicio)
        .FeFin = DateAdd("yyyy", i, vEmpresa.FechaFin)
        .Opcion = 3
        .Show vbModal
        
        If CadenaDesdeOtroForm <> "" Then Unload Me
            
    End With
End Sub

Private Sub txtDecimal_GotFocus(Index As Integer)
    ConseguirFoco txtDecimal(Index), 3
End Sub

Private Sub txtDecimal_KeyPress(Index As Integer, KeyAscii As Integer)
     Keypress KeyAscii
End Sub

Private Sub txtDecimal_LostFocus(Index As Integer)
Dim C As Currency

    If Index = 0 Then Exit Sub
    txtDecimal(Index).Text = Trim(txtDecimal(Index).Text)
    C = 0
    If txtDecimal(Index).Text <> "" Then
       If EsNumerico(txtDecimal(Index).Text) Then
            If InStr(txtDecimal(Index).Text, ",") = 0 Then
                C = CCur(TransformaPuntosComas(txtDecimal(Index).Text))
            Else
                C = ImporteFormateado(txtDecimal(Index).Text)
            End If
        End If
    End If
    
    If C = 0 Then
        txtDecimal(Index).Text = 0
        txtDecimal(Index).Tag = 0
    Else
        txtDecimal(Index).Tag = C
        VerificarCampo
    End If
    
    
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub txtFecha_LostFocus(Index As Integer)
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    If txtFecha(Index).Text = "" Then Exit Sub
    
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    If Not EsFechaOK(txtFecha(Index)) Then
        txtFecha(Index).Text = ""
        PonerFoco txtFecha(Index)
    End If
    'Cuando el indice es 3 pongo en hasta(el 4),si esta vacio, lo mismo
    If Index = 3 Then
        If txtFecha(4).Text = "" Then txtFecha(4).Text = txtFecha(3).Text
    End If
End Sub



Private Sub CargaCombo(Cual As Integer)


    Select Case Cual
    Case 0
        cmbFecha.Clear
        cmbFecha.AddItem ("Seleccione un mes")
        
        If Year(vEmpresa.FechaInicio) = Year(vEmpresa.FechaFin) Then
            'AÑOS NATURALES
            For i = 1 To 12
                cad = Format(CDate("01/" & i & "/2000"), "mmmm")
                cmbFecha.AddItem cad
                cmbFecha.ItemData(cmbFecha.NewIndex) = i
            Next i
        Else
            'Años partidos
            For i = Month(vEmpresa.FechaInicio) To Month(vEmpresa.FechaFin) + 12
                If (i Mod 12) = 0 Then
                    cad = Format(CDate("01/12/2000"), "mmmm")
                Else
                    cad = Format(CDate("01/" & (i Mod 12) & "/2000"), "mmmm")
                End If
                cmbFecha.AddItem cad
                cmbFecha.ItemData(cmbFecha.NewIndex) = i
            Next i
        
        End If
        
    Case 1
        Me.cboMes(0).Clear
        For i = 1 To 12
            cad = Format(CDate("01/" & i & "/2000"), "mmmm")
            cboMes(0).AddItem cad
            cboMes(0).ItemData(cboMes(0).NewIndex) = i 'NO HACE FALTA
        Next i
        i = Month(Now)
        If i = 1 Then
            Me.cboMes(0).ListIndex = 11
            Me.txtNumero(0).Text = Year(Now) - 1
        Else
            Me.cboMes(0).ListIndex = i - 1 - 1
            Me.txtNumero(0).Text = Year(Now)
        End If
        
        
        
    End Select
End Sub

Private Sub txtNumero_GotFocus(Index As Integer)
    ConseguirFoco txtNumero(Index), 3
End Sub


Private Sub txtNumero_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub txtNumero_LostFocus(Index As Integer)
    txtNumero(Index).Text = Trim(txtNumero(Index).Text)
    If txtNumero(Index).Text = "" Then Exit Sub
    
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    If Not IsNumeric(txtNumero(Index).Text) Then
        txtNumero(Index).Text = ""
        PonerFoco txtNumero(Index)
    End If

End Sub



Private Sub HorasTrabajadasSeccionesSinNomina(ListaTrabajadores As String)

Dim C As String
'INSERT INTO jornadassemanalesalz(idtrabajador,fecha,TipoHoras,horastrabajadas,ParaEmpresa,Ajuste)

    C = "INSERT INTO jornadassemanalesalz(idtrabajador,fecha,TipoHoras,horastrabajadas,ParaEmpresa,Ajuste) VALUES"
    cad = "select idtrabajador,fecha," & cad & ",if(horastrabajadas>9,9,horastrabajadas)"
    cad = cad & " from marcajes where " & TipoAlziraEntreFechas
    cad = cad & ListaTrabajadores
    
End Sub

Private Sub txtParadas_Change(Index As Integer)

End Sub

Private Sub IncrementarDecrementarParadas(Subir As Boolean)
    If Subir Then
        txtDecimal(1).Tag = txtDecimal(1).Tag + 0.25
    Else
        txtDecimal(1).Tag = txtDecimal(1).Tag - 0.25
    End If
    
    VerificarCampo
End Sub


Private Sub VerificarCampo()
    If txtDecimal(1).Tag > 2 Then
        MsgBox "Hora superior a 2", vbCritical
        txtDecimal(1).Tag = 2
    Else
        If txtDecimal(1).Tag < 0 Then txtDecimal(1).Tag = 0
    End If
    
    txtDecimal(1).Text = Format(txtDecimal(1).Tag, FormatoImporte)
End Sub

Private Sub UpDown1_DownClick()
    IncrementarDecrementarParadas False
End Sub

Private Sub UpDown1_UpClick()
    IncrementarDecrementarParadas True
End Sub
