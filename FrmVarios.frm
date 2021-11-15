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
    
    
    ' 7. Generacion de horas semanales para "segunda empresa" fruixeras-motilla (SOLO alzicoop, obviamente)
    
    
Public Parametros As String

Private WithEvents frmc As frmCal
Attribute frmc.VB_VarHelpID = -1


Dim FInicioSeccion As Date
Dim I As Integer
Dim Cad As String


Private Sub KeyPress(KeyAscii As Integer)
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
    
            Cad = DevuelveDesdeBD("nominas", "secciones", "idseccion", Me.cboSeccion.ItemData(cboSeccion.ListIndex), "N")
            cboSeccion.Tag = Val(Cad)
            FInicioSeccion = "01/01/2001"
            Cad = DevuelveDesdeBD("max(fechafin)", "jornadassemanalesproceso", "seccion", Me.cboSeccion.ItemData(cboSeccion.ListIndex), "N")
            If Cad <> "" Then
                Me.txtFecha(5).Text = Format(Cad, "dd/mm/yyyy")
                Me.txtFecha(5).Text = DateAdd("d", 1, CDate(Me.txtFecha(5).Text))
                FInicioSeccion = CDate(txtFecha(5).Text)
                
                'seccion.nominas =1
                If cboSeccion.Tag = 1 Then
                    'Sera desde hasta el domingo de esa semana
                    I = Format(CDate(Me.txtFecha(5).Text), "w", vbMonday)
                    I = 7 - I
                    Me.txtFecha(6).Text = Format(DateAdd("d", I, CDate(Me.txtFecha(5).Text)), "dd/mm/yyyy")
                Else
                    Me.txtFecha(6).Text = ""
                End If
            End If
        
    Else
        'Solo ALzira
        'Es generar datos para FRUXERESA -motilla
        'E
            Cad = DevuelveDesdeBD("nominas", "secciones", "idseccion", Me.cboSeccion.ItemData(cboSeccion.ListIndex), "N")
            cboSeccion.Tag = Val(Cad)
            FInicioSeccion = "01/01/2001"
            Cad = DevuelveDesdeBD("max(fechafin)", "jornadassemanalesproceso", "seccion", Me.cboSeccion.ItemData(cboSeccion.ListIndex), "N")
            If Cad <> "" Then
                Me.txtFecha(5).Text = Format(Cad, "dd/mm/yyyy")
                Me.txtFecha(5).Text = DateAdd("d", 1, CDate(Me.txtFecha(5).Text))
                FInicioSeccion = CDate(txtFecha(5).Text)
                
                'seccion.nominas =1
                If cboSeccion.Tag = 1 Then
                    'Sera desde hasta el domingo de esa semana
                    I = Format(CDate(Me.txtFecha(5).Text), "w", vbMonday)
                    I = 7 - I
                    Me.txtFecha(6).Text = Format(DateAdd("d", I, CDate(Me.txtFecha(5).Text)), "dd/mm/yyyy")
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
    Cad = "Select max(fecha) from marcajes"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
                I = Year(vEmpresa.FechaInicio)
                If Me.optPeriodo(1).Value Then I = I + 1
                    
        Else
            'Años partidos
            I = Year(vEmpresa.FechaFin)
            If Not (vEmpresa.FechaFin <= FechaUltimoProcesado) Then I = I + 1
                
        End If
        txtFecha(3).Text = "01/" & cmbFecha.ItemData(cmbFecha.ListIndex) & "/" & I
    End If
    If CDate(txtFecha(3).Text) <= FechaUltimoProcesado Then
        'Vuelvo a poner la txt en blanco
        If MesEntero Then txtFecha(3).Text = ""
        MsgBox "Fecha de incio esta ya procesada: " & FechaUltimoProcesado, vbExclamation
        Exit Function
    End If
    
    
    If MesEntero Then
        F1 = CDate("01/" & cmbFecha.ItemData(cmbFecha.ListIndex) & "/" & I)
        I = DiasMes(cmbFecha.ItemData(cmbFecha.ListIndex), I)
        F2 = I & "/" & cmbFecha.ItemData(cmbFecha.ListIndex) & "/" & Year(F1)
        
    Else
        'Son dias de intervalo
        F1 = CDate(txtFecha(3).Text)
        F2 = CDate(txtFecha(4).Text)
    End If

    Cad = "UPDATE calendariot Set tipodia = 2 WHERE  idTrabajador =  " & Val(RecuperaValor(Parametros, 2))
    Cad = Cad & " AND fecha = '"
    While F1 <= F2
        EjecutaSQL Cad & Format(F1, FormatoFecha) & "'"
        F1 = DateAdd("d", 1, F1)
    Wend
    
    GenerarVacaciones = True
End Function


Private Sub cmdCalcularHorasTrabajadasSemana_Click()
    
    
     'Muuuchas cosas a comprobar
    Cad = ""
    If txtFecha(5).Text = "" Or txtFecha(6).Text = "" Then Cad = "Ponga las fechas"
        
    If cboSeccion.ListIndex < 0 Then
        If vEmpresa.QueEmpresa = 4 Then
            cboSeccion.Tag = 1
            
            'Procesamos todas las secciones a la vez
            
            'FInicioSeccion = txtFecha(5).Text
           
            
        Else
            Cad = "Falta sección" & vbCrLf & Cad
        End If
    End If
    If Cad <> "" Then
        MsgBox Cad, vbExclamation
        Exit Sub
    End If


    'Faltara ver si esta la semana completa o es final de semana
    Cad = ""
    If Not IsDate(txtFecha(5).Text) Or Not IsDate(txtFecha(6).Text) Then
        Cad = "Fechas incorrectas"
    Else
        'VA POR SEMANAS
        If Me.cboSeccion.Tag = 1 Then
    
            If DateDiff("d", txtFecha(5).Text, txtFecha(6).Text) > 6 Then
                'VA por semanas
                 Cad = "Mas de una semana seleccionada"
                
            ElseIf Format(CDate(txtFecha(5).Text), "ww", vbMonday) <> Format(CDate(txtFecha(6).Text), "ww", vbMonday) Then
                'Procesamos de semana en semana. Deben pertenecer a la misma semana
                Cad = "Semanas distintas"
            Else
                'OK. Faltara ver mas casos
                
                
                'Agosto 2014
                '-----------------------------------------------------
                'Las semanas
                'La fecha inicio NO puede ser inferior a la que le corresponde, que esta guardado en FInicioSeccion
                If CDate(txtFecha(5).Text) < FInicioSeccion Then
                   Cad = "Intervalo ya procesado. Inicio debe ser: " & FInicioSeccion
                Else
                    'Si no procesa desde qyue le corresponde
                    'veremos si hay datos pendientes de procesar
                     If CDate(txtFecha(5).Text) > FInicioSeccion Then
                          'Si la fecha es mayor a la que ponen como inicio del intervalor, veremos si ya hay procesados
                          'para esas fechas y esa seccion
                          Cad = "fecha >=" & DBSet(FInicioSeccion, "F") & " AND fecha <" & DBSet(txtFecha(5).Text, "F")
                          Cad = Cad & " AND idtrabajador IN (select idtrabajador from trabajadores  "
                          
                          If cboSeccion.ListIndex > 0 Then Cad = Cad & " WHERE Seccion = " & cboSeccion.ItemData(cboSeccion.ListIndex)
                          
                          Cad = Cad & ") AND 1"
                          Cad = DevuelveDesdeBD("count(*)", "jornadassemanalesalz", Cad, "1")
                          If Val(Cad) > 0 Then
                               Cad = "Existen datos procesados entre: " & FInicioSeccion & " y " & txtFecha(5).Text
                          Else
                               Cad = ""
                          End If
                    
                           'Veremos sy hay marcajes pendientes de procesar en ese periodo que es
                           'desde que le corresponde hasta donde empieza
                          If Cad = "" Then
                             Cad = "fecha >=" & DBSet(FInicioSeccion, "F") & " AND fecha <" & DBSet(txtFecha(5).Text, "F")
                             Cad = Cad & " AND idtrabajador IN (select idtrabajador from trabajadores where 1=1"
                             If cboSeccion.ListIndex > 0 Then Cad = Cad & " AND seccion=" & cboSeccion.ItemData(cboSeccion.ListIndex)
                             Cad = Cad & ") AND 1"
                             Cad = DevuelveDesdeBD("count(*)", "marcajes", Cad, "1")
                             If Val(Cad) > 0 Then
                                  Cad = "Existen marcajes entre: " & FInicioSeccion & " y " & txtFecha(5).Text & " que no entran en el intervalo para procesar"
                             Else
                                  Cad = ""
                             End If
                        
                          End If
                       
                       
                          'Febrero 2015
                          If Cad = "" Then
                                'Comrpobacion una. que el dia es menor que jueves
                                I = Weekday(txtFecha(5).Text, vbMonday)
                                If I >= 5 Then
                                    'Es viernes. Por lo tanto, el dia hasta tiene que ser domingo
                                    If Weekday(CDate(txtFecha(6).Text), vbMonday) <> 7 Then Cad = "Fecha final de intervalo debe ser domingo. Proceso semana completo"
                                        
                                End If
                          End If
                    
                    End If  'de mayor que fecha inicioseccion
                    
                    If Cad = "" Then
                           Cad = "fecha >=" & DBSet(txtFecha(5).Text, "F") & " AND fecha <=" & DBSet(txtFecha(6).Text, "F")
                             Cad = Cad & " AND idtrabajador IN (select idtrabajador from trabajadores where 1=1"
                             If cboSeccion.ListIndex > 0 Then Cad = Cad & " AND seccion=" & cboSeccion.ItemData(cboSeccion.ListIndex)
                             Cad = Cad & ") AND correcto "
                             Cad = DevuelveDesdeBD("count(*)", "marcajes", Cad, "0")
                             If Val(Cad) > 0 Then
                                  Cad = "Existen marcajes INCORRECTOS entre: " & txtFecha(5).Text & " y " & txtFecha(6).Text
                             Else
                                  Cad = ""
                             End If
                    End If
                    
                End If
                            
                            
                
                
            End If
        End If  'de por semanas
    End If
    If Cad <> "" Then
    
        MsgBox Cad, vbExclamation
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
        Cad = "select distinct(idcal) from trabajadores where seccion=" & cboSeccion.ItemData(cboSeccion.ListIndex)
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic
        Cad = ""
        While Not miRsAux.EOF
            Cad = Cad & "1"
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    Else
        Cad = "1"
    End If
    If Len(Cad) <> 1 Then
        If vEmpresa.QueEmpresa <> 5 Then
            MsgBox "Solo puede haber un calendario para la seccion. ", vbExclamation
            Exit Sub
        End If
    End If
    
    
    
    
    Set ColTraba = New Collection
    
    Cad = "Select marcajes.idtrabajador FROM marcajes,trabajadores WHERE marcajes.idtrabajador=trabajadores.idtrabajador AND"
    Cad = Cad & TipoAlziraEntreFechas
    If cboSeccion.ListIndex >= 0 Then Cad = Cad & " AND seccion=" & cboSeccion.ItemData(cboSeccion.ListIndex)
    Cad = Cad & " GROUP BY 1"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    Cad = ""
    While Not miRsAux.EOF
        Cad = Cad & ", " & miRsAux!idTrabajador
        NumRegElim = NumRegElim + 1
        If NumRegElim > 50 Then
            ColTraba.Add Mid(Cad, 2)
            Cad = ""
            NumRegElim = 0
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If NumRegElim > 0 Then ColTraba.Add Mid(Cad, 2)
    
    '
    Cad = "DELETE FROM tmphorastipoalzira WHERE codusu =" & vUsu.Codigo
    conn.Execute Cad
    
    
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


'Horas por Area, alzira
'Veremos cuantos trabajadores (y cuales), han trabajado en distintas areas
Dim Cadena1 As String
Dim Cadena2 As String
Dim F As Date
Dim ContErrores As Integer
    
Dim rAlmuerzo As ADODB.Recordset
Dim hAlmuerzo As Currency

Dim CurrencyAux As Currency
Dim VariableAux As String
        
    HorasmaximoNormalesDia = 9
    If vEmpresa.QueEmpresa = vbCatadau Then HorasmaximoNormalesDia = 8
    
    'Hay empresas CASTELDUC, que hacen una compensacion en nomina a final de mes. Y ya esta. No van dia a dia
    If vEmpresa.CompensaHorasNominaMES Then HorasmaximoNormalesDia = vEmpresa.CompensaMES_HorasDia   'HorasmaximoNormalesDia = 10
        
    
    Label3(14).Caption = "Calculando horas-tipo-area"
    Label3(14).Refresh
    
    
    Insert = "INSERT INTO tmphorastipoalzira(codusu,idtrabajador,diasem,fecha,TipoHoras,horastrabajadas) "
    
    'Varios pasos
    'Primero los domingos y festivos ENTRAN Con todas las horas extras
    '----------------------------------------------------------------
    Cad = 2 'HORA EXTRA
    Cad = "select " & vUsu.Codigo & ", idtrabajador,date_format(fecha,'%w') diasem,fecha," & Cad & ",horastrabajadas"
    Cad = Cad & " from marcajes where " & TipoAlziraEntreFechas
    Cad = Cad & ListaTrabajadores & " AND "
    
    
    conn.Execute Insert & Cad & " date_format(fecha,'%w')=0" 'domingos
    
    conn.Execute Insert & Cad & " Festivo = 1 and date_format(fecha,'%w')<>0"  'festivos que no sean domingos
    
    
    
        
        Cad = 0 'HORA normales
        Cad = "select " & vUsu.Codigo & ", idtrabajador,date_format(fecha,'%w') diasem,fecha," & Cad & ",if(horastrabajadas>" & HorasmaximoNormalesDia & "," & HorasmaximoNormalesDia & ",horastrabajadas)"
        Cad = Cad & " from marcajes where " & TipoAlziraEntreFechas
        Cad = Cad & ListaTrabajadores & " AND "
        conn.Execute Insert & Cad & " date_format(fecha,'%w') in (1,2,3,4,5) AND    Festivo = 0"
        
        'Las que se pasen de 9 (HorasmaximoNormalesDia) van a horas estrcutrales
        Cad = 1 'HORA estrucutrales
        Cad = "select " & vUsu.Codigo & ", idtrabajador,date_format(fecha,'%w') diasem,fecha," & Cad & ",horastrabajadas-" & HorasmaximoNormalesDia
        Cad = Cad & " from marcajes where " & TipoAlziraEntreFechas
        Cad = Cad & ListaTrabajadores & " AND "
        Cad = Cad & " date_format(fecha,'%w') in (1,2,3,4,5) AND Festivo = 0 AND horastrabajadas>" & HorasmaximoNormalesDia
        conn.Execute Insert & Cad
        
        
    
    
    'ALZRIRA
    'Los sabados, a partir de las 14:30 son extras
    '------------------------------------------------
    'Los sabados, que no sean festivos
    Label3(14).Caption = "Vers diasema: 5"
    Label3(14).Refresh
    Cad = "select fecha from marcajes where " & TipoAlziraEntreFechas
    Cad = Cad & ListaTrabajadores & " and date_format(fecha,'%w')=6 and festivo=0 GROUP BY 1"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
    
    
    For I = 1 To ColSabados.Count
            
            
            'Veremos que trabajadores tienen un fichaje mas alla de las 14:30(HoraSabadoExtras )
            Cad = "select idtrabajador from  entradamarcajes where fecha=" & DBSet(ColSabados.Item(I), "F") & " and hora>" & DBSet(HoraSabadoExtras, "H")
            Cad = Cad & ListaTrabajadores & " GROUP BY 1"
            miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            AuxTra = ""
            While Not miRsAux.EOF
                AuxTra = AuxTra & ", " & miRsAux!idTrabajador
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            'YA tengo los trabadores que el SABADO ha trabajado mas alla de las HoraSabadoExtras
            
            'Los que NO han ido mas alla HoraSabadoExtras
                Cad = 1 'HORA estrucutrales
                Cad = "select " & vUsu.Codigo & ", idtrabajador,date_format(fecha,'%w') diasem,fecha," & Cad & ",if(horastrabajadas>" & HorasmaximoNormalesDia & "," & HorasmaximoNormalesDia & ",horastrabajadas)"
                Cad = Cad & " from marcajes where fecha=" & DBSet(ColSabados.Item(I), "F") & " AND horastrabajadas>" & HorasmaximoNormalesDia
                Cad = Cad & ListaTrabajadores
                If AuxTra <> "" Then
                    AuxTra = Mid(AuxTra, 2)
                    'En auxtra estan los que han trabajado mas alla de las 14:30
                    Cad = Cad & " and not idtrabajador in  (" & AuxTra & ")"
    
                End If
                conn.Execute Insert & Cad
                
                Cad = 0 'HORA normales
                Cad = "select " & vUsu.Codigo & ", idtrabajador,date_format(fecha,'%w') diasem,fecha," & Cad & ",horastrabajadas"
                Cad = Cad & " from marcajes where fecha=" & DBSet(ColSabados.Item(I), "F") & " AND horastrabajadas<=" & HorasmaximoNormalesDia
                Cad = Cad & ListaTrabajadores
                If AuxTra <> "" Then
                    'En auxtra estan los que han trabajado mas alla de las  HoraSabadoExtras
                    Cad = Cad & " and not idtrabajador in  (" & AuxTra & ")"
                End If
                conn.Execute Insert & Cad
                
          
                
            
            'Los que han trabajado el sabado REESTRUCTURAMOS sus horas
            If AuxTra <> "" Then
                Set RT = New ADODB.Recordset
                
                'REESTABLECEMOS LAS HORAS PARA AQUELLOS  han trabajado mas alla de las HoraSabadoExtras
                Cad = "Select marcajes.*,ExcesoDefecto from marcajes,incidencias where idinci=IncFinal AND fecha=" & DBSet(ColSabados.Item(I), "F") & " and idtrabajador in  (" & AuxTra & ")"
                miRsAux.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                While Not miRsAux.EOF
                
                    Label3(14).Caption = "sabados " & miRsAux!idTrabajador
                    Label3(14).Refresh
                
                    Debug.Print miRsAux!idTrabajador
                    'FALTA####
                    'If miRsAux!idTrabajador = 142 Then St op
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
                    Cad = "Select entradamarcajes.*,hour(hora) lahora, ADDTIME(hora , '-24:00:00' ) horaajustada from entradamarcajes  where fecha=" & DBSet(miRsAux!Fecha, "F") & " AND idtrabajador =" & miRsAux!idTrabajador
                    'Cad = Cad & "  and hora>" & DBSet(HoraSabadoExtras, "H") & "  and hora <='23:59:59' ORDER by hora desc"
                    Cad = Cad & "  and hora>" & DBSet(HoraSabadoExtras, "H") '& "  and hora <='23:59:59' ORDER by hora desc"
                    Cad = Cad & "  ORDER by hora desc"
                    
                    
                    RT.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    Fin = False
                    Do
                        If Val(RT!LaHora) >= 24 Then
                            T1 = CCur(DevuelveValorHora(RT!horaajustada)) + 24
                        Else
                            'Lo que habia
                            T1 = CCur(DevuelveValorHora(RT!Hora))
                        End If
                        
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
                            
                            
                            CurrencyAux = T2 - HN
                            
                            'Noviembre 2020
                            'If T2 > HN Then
                            If CurrencyAux >= 1 Then
                                MsgBox "NO puedo quitar mas horas:" & miRsAux!Fecha & " - " & miRsAux!idTrabajador, vbExclamation
                                Debug.Print miRsAux!idTrabajador
                            Else
                                If CurrencyAux > 0 Then
                                    'Puede ser que haya parado a merendar, con lo cual , en horas exceo tengo que quitarle esta diferencia
                                    HorasExceso = HorasExceso - CurrencyAux
                                    HN = 0
                                Else
                                    'Lo que habia
                                    HN = HN - T2
                                End If
                            End If
                           
                           
                           
                        End If
                    Loop Until Fin
                    
                    
                    'Insertamos en la tmp
                    'Normales
                    'tmphorastipoalzira(codusu,idtrabajador,diasem,fecha,TipoHoras,horastrabajadas)
                    Cad = " VALUES (" & vUsu.Codigo & "," & miRsAux!idTrabajador & ",5," & DBSet(miRsAux!Fecha, "F") & ","
                    
                    Cad = Insert & Cad
                    
                    'Las normales
                        
                        conn.Execute Cad & "0," & DBSet(HN, "N") & ")"   'normales
                        
                        If HorasExceso > 0 Then conn.Execute Cad & "2," & DBSet(HorasExceso, "N") & ")"  'LS QUE VAN MAS ALLA SON EXTRA
                        If HI > 0 Then conn.Execute Cad & "1," & DBSet(HI, "N") & ")"   'LS QUE VAN MAS ALLA SON estruc

                    
                    RT.Close
                    miRsAux.MoveNext 'siguiente trabajador
                Wend
                miRsAux.Close
            End If
        
        

                    
    
                
    Next I
    
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
            
    
    
    
    
    Cad = "Delete FROM tmphorastipoalzira WHERE codusu =" & vUsu.Codigo & " AND horastrabajadas=0"
    conn.Execute Cad
    
    'Nuevo proceso de Areas
    'ALZRIRA
    'Noviembre 2020
    'Areas
    Cad = "(select distinct area from terminales) as mister "
    Cad = DevuelveDesdeBD("count(*)", Cad, "1", "1")
    If Val(Cad) > 1 Then
    
        Label3(14).Caption = "Areas-secciones 1"
        Label3(14).Refresh
    
        conn.Execute "DELETE from tmphorasArea WHERE codusu =" & vUsu.Codigo
        'para los posible errores
        Cad = "Delete FROM tmppresencia WHERE codusu =" & vUsu.Codigo
        conn.Execute Cad
    
        
        'En la tabla
        
        
        
        'Marcajes que solo tienen  1 zona, con lo cual la zona es la que es
        '--------------------------------------------------------------------
        Cad = "select idtrabajador, fecha, count(*) from "
        Cad = Cad & " ("
        Cad = Cad & " select distinct idtrabajador,fecha,area"
        Cad = Cad & " from entradamarcajes,terminales where entradamarcajes.reloj =terminales.id"
        Cad = Cad & " AND   " & TipoAlziraEntreFechas
        Cad = Cad & ")  aaa   "  'subselect
        Cad = Cad & " group by 1,2"
        Cad = Cad & " having count(*) = 1"
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cadena1 = ""
        I = 0
        While Not miRsAux.EOF
        
            
        
           Cadena1 = Cadena1 & ", (" & miRsAux!idTrabajador & "," & DBSet(miRsAux!Fecha, "F") & ")"
           I = I + 1
           miRsAux.MoveNext
           If miRsAux.EOF Then I = 100
           If I > 20 Then
           
           
                Label3(14).Caption = "Areas-secciones 1. " & Format(Now, "hh:mm:ss")
                Label3(14).Refresh
            

                'insert INTO tmpinformehorasmes(codusu,fecha,idTrabajador,DT,H1)
                Cadena1 = Mid(Cadena1, 2)
                Cad = "INSERT INTO tmphorasArea(codusu,idtra,Fecha,Area,masdenArea,Horas) "
                Cad = Cad & " SELECT " & vUsu.Codigo & ",idtrabajador,fecha,area,0,0 "   'Como solo tiene un area, NO hace falta calcular las horas
                Cad = Cad & " from entradamarcajes,terminales where entradamarcajes.reloj =terminales.id"
                Cad = Cad & " AND (idtrabajador,fecha) in ("
                Cad = Cad & Cadena1 & " ) group by fecha,idtrabajador"
                conn.Execute Cad
                Cadena1 = ""
                I = 0
           End If
        Wend
        miRsAux.Close
        
        
        
        'Los que tienen mas de una zona
        Label3(14).Caption = "Areas-secciones 2"
        Label3(14).Refresh
        Cad = "select idtrabajador, fecha, count(*) from "
        Cad = Cad & " ("
        Cad = Cad & " select distinct idtrabajador,fecha,area"
        Cad = Cad & " from entradamarcajes,terminales where entradamarcajes.reloj =terminales.id"
        Cad = Cad & " AND   " & TipoAlziraEntreFechas
        Cad = Cad & ")  aaa   "  'subselect
        Cad = Cad & " group by 1,2"
        Cad = Cad & " having count(*) >1"
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
        Set rAlmuerzo = New ADODB.Recordset
        Cad = "Select idTrabajador,Fecha, HorasDto FROM marcajes where horasdto>0 AND  "
        Cad = Cad & TipoAlziraEntreFechas
        Cad = Cad & " ORDER BY 1,2"
        rAlmuerzo.Open Cad, conn, adOpenKeyset, adCmdText
        
        
        Cadena1 = ""
        NumRegElim = 0
        ContErrores = 0
        Set RT = New ADODB.Recordset
        While Not miRsAux.EOF
            Cadena1 = Cadena1 & ", (" & miRsAux!idTrabajador & "," & DBSet(miRsAux!Fecha, "F") & ")"
            NumRegElim = NumRegElim + 1
            miRsAux.MoveNext
            If miRsAux.EOF Then NumRegElim = 100
            If NumRegElim > 20 Then
                'Tengo trabajador / fecha
                Label3(14).Caption = "Areas distintas"
                Label3(14).Refresh
                Cadena1 = Mid(Cadena1, 2)
                Cad = "Select * from entradamarcajes,terminales where entradamarcajes.reloj =terminales.id"
                Cad = Cad & " AND (idtrabajador,fecha) IN (" & Cadena1 & ") ORDER BY  idtrabajador,fecha,hora,reloj"
                RT.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                AuxTra = ""
                F = "01/01/1900"
                Cadena2 = ""
                While Not RT.EOF
                    Label3(14).Caption = "Areas distintas.  " & RT!idTrabajador & "  " & RT!Fecha
                    Label3(14).Refresh
                    
                    If RT!idTrabajador <> AuxTra Or F <> RT!Fecha Or CStr(RT!Area) <> Insert Then
                        If AuxTra <> "" Then
                            
                            If hAlmuerzo > 0 Then
                                If hAlmuerzo < HN Then
                                    If hAlmuerzo > 0.5 Then
                                        'stop
                                        MsgBox "Hora parada > 0.5 (Merienda?)", vbExclamation
                                    End If
                                    HN = HN - hAlmuerzo
                                    hAlmuerzo = 0
                                Else
                                    MsgBox "MENOS HORAS que el almuerzo del dia"
                                    'Stop
                                End If
                            End If
                            'INSERT INTO tmphorasArea(codusu,idtra,Fecha,Area,masdenArea,Horas) "
                            Cad = ", (" & vUsu.Codigo & "," & AuxTra & "," & DBSet(F, "F") & "," & Insert & "," & I & "," & DBSet(HN, "N") & ")"
                            Cadena2 = Cadena2 & Cad
                        
                        
                            If Len(Cadena2) > 3000 Then
                                Cadena2 = Mid(Cadena2, 2)
                                Cadena2 = "INSERT INTO tmphorasArea(codusu,idtra,Fecha,Area,masdenArea,Horas) VALUES " & Cadena2
                                conn.Execute Cadena2
                                Cadena2 = ""
                            End If
                        
                        
                        End If
                        'Si el trabajador o la fecha es distinto reseteo I  , que se grabara en "masdeunazona" y servira para ordenar las zonas por hora
                        If RT!idTrabajador <> AuxTra Or F <> RT!Fecha Then
                            I = 0
                            hAlmuerzo = 0
                            'Veremos si tiene almuerzo
                            If LocalizaRegistropradas(rAlmuerzo, RT!idTrabajador, RT!Fecha) Then
                                hAlmuerzo = rAlmuerzo!HorasDto
                           
                            End If
                            
                        End If
                        
                        I = I + 1
                        AuxTra = RT!idTrabajador
                        F = RT!Fecha
                        Insert = RT!Area
                        HN = 0
                                                    
                        'Ejemplo de como quedara tmphorasArea   (trabajador 6)
                        'tra   fec      ZON sec  horas
                        '6   2020-10-01  1   1   7.50   Empieza trabajadno 7.5 horas en area 1
                        '6   2020-10-01  2   2   3.50    y va 3.5 a area 2
                        '6   2020-10-02  2   1   7.00   dia 2   Emipeza 7.5 horas en zona 2
                        '6   2020-10-02  1   2   3.50       y luego va a zona 1 durant 3.5 horas
                        '6   2020-10-03  0   0   0.00    Trabaja todo el dia en zona 0
                                                
                    End If
                    
                    
                    
                    
                    T1 = CCur(DevuelveValorHora(RT!Hora))
                    RT.MoveNext
                    If RT.EOF Then
                        'ERROR
                        ContErrores = ContErrores + 1
                        Cad = "Marcajes impares"
                        Cad = "INSERT INTO tmppresencia(codusu,Id,NomTrabajador,idtra,fecha) VALUES (" & vUsu.Codigo & "," & ContErrores & "," & DBSet(Cad, "T") & "," & AuxTra & "," & DBSet(F, "F") & ")"
                        conn.Execute Cad
                        T2 = T1
                    Else
                        T2 = CCur(DevuelveValorHora(RT!Hora))
                        If Insert <> RT!Area Then
                            'Dos marcajes con distinta AREA
                            ContErrores = ContErrores + 1
                            Cad = "Marcajes distinta area."
                            Cad = "INSERT INTO tmppresencia(codusu,Id,NomTrabajador,idtra,fecha) VALUES (" & vUsu.Codigo & "," & ContErrores & "," & DBSet(Cad, "T") & "," & AuxTra & "," & DBSet(F, "F") & ")"
                            conn.Execute Cad
                        End If
                        
                    End If
                    
                    
                    
                    'Calculamos la diferencia
                    HN = HN + (T2 - T1)
                            
                    
                    RT.MoveNext
                Wend
                RT.Close
                
                 If AuxTra <> "" Then
                    'INSERT INTO tmphorasArea(codusu,idtra,Fecha,Area,masdenArea,Horas) "
                    Cad = ", (" & vUsu.Codigo & "," & AuxTra & "," & DBSet(F, "F") & "," & Insert & "," & I & "," & DBSet(HN, "N") & ")"
                    Cadena2 = Cadena2 & Cad
                End If
                
                Cadena1 = ""
                NumRegElim = 0
                If Cadena2 <> "" Then
                    Cadena2 = Mid(Cadena2, 2)
                    Cadena2 = "INSERT INTO tmphorasArea(codusu,idtra,Fecha,Area,masdenArea,Horas) VALUES " & Cadena2
                    conn.Execute Cadena2
                End If
            End If
        Wend
        miRsAux.Close
        rAlmuerzo.Close
        Set rAlmuerzo = Nothing
       
        
        
        If ContErrores > 0 Then
            
            With frmImprimir
                .FormulaSeleccion = "{tmppresencia.codusu} = " & vUsu.Codigo
                .NombreRPT100 = "rErrorProcHoras.rpt"
                .Titulo100 = "Errores proceso horas"
                Cad = LCase(TipoAlziraEntreFechas)
                Cad = Replace(Cad, "fecha", "")
                Cad = Replace(Cad, "between", "")
                Cad = Replace(Cad, "and", "-")
                Cad = Trim(Replace(Cad, """", ""))
                .OtrosParametros = "|pEmp=""" & vEmpresa.NomEmpresa & """|datos= """ & Cad & " ""|"
                .Opcion = 100
                .NumeroParametros = 2
                .Show vbModal
            End With
            
            
            
        End If
        
        
        
        
    End If
    
    
    
    
    
    
    
    Label3(14).Caption = "Dias-nomina"
    Label3(14).Refresh
    
    
    
    
    
End Sub








Private Function TipoAlziraEntreFechas()
    TipoAlziraEntreFechas = " fecha Between " & DBSet(txtFecha(5).Text, "F") & " and " & DBSet(txtFecha(6).Text, "F")
End Function




Private Sub Command1_Click()
    I = 0
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
        Cad = "Desea asignar como horas de parada: " & txtDecimal(1).Text & "  a :" & vbCrLf & Parametros
        If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
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
            'fruix-motilla
            Caption = "Horas " & "vEmpresa.SegundaEmpresa"
            Label4(1).Caption = "Gen HORAS " & "vEmpresa.SegundaEmpresa"
            Label4(1).ForeColor = &HC000C0
        End If
        
        
        Cad = "DELETE FROM tmphorastipoalzira WHERE codusu =" & vUsu.Codigo
        conn.Execute Cad
        
        
        cboSeccion.Clear
        Set miRsAux = New ADODB.Recordset
        Cad = "select * from secciones where idseccion in (select seccion from trabajadores) "
        If Opcion = 4 Then Cad = Cad & " and nominas=1"
        Cad = Cad & " ORDER BY nombre"
        
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        

        While Not miRsAux.EOF
            cboSeccion.AddItem miRsAux!Nombre & " (" & miRsAux!IdSeccion & ")"
            cboSeccion.ItemData(cboSeccion.NewIndex) = miRsAux!IdSeccion
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        
        
        If vEmpresa.QueEmpresa = 4 Then
                    
            Cad = "select seccion,max(fechaini) fi,max(fechafin) ff from jornadassemanalesproceso  WHERE true "
            Cad = Cad & " AND seccion in (select idseccion from secciones where nominas=1)"
            Cad = Cad & " group by 1"
            miRsAux.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
            Cad = ""
            
            While Not miRsAux.EOF
                If Cad = "" Then
                    Cad = Format(IIf(IsNull(miRsAux!FI), "01/01/1900", miRsAux!FI), "dd/mm/yyyy") & Format(IIf(IsNull(miRsAux!FF), "01/01/1900", miRsAux!FF), "dd/mm/yyyy")
                Else
                    CadenaDesdeOtroForm = Format(IIf(IsNull(miRsAux!FI), "01/01/1900", miRsAux!FI), "dd/mm/yyyy") & Format(IIf(IsNull(miRsAux!FF), "01/01/1900", miRsAux!FF), "dd/mm/yyyy")
                    If Cad <> CadenaDesdeOtroForm Then
                        Cad = "MAL"
                        miRsAux.MoveLast
                    End If
                End If
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            If Cad <> "" Then
                If Cad = "MAL" Then
                    MsgBox "Distintas fechas en la generacion por secciones", vbExclamation
                Else
                    
                    FInicioSeccion = CDate(Mid(Cad, 11, 10))
                    FInicioSeccion = DateAdd("d", 1, FInicioSeccion)
                    txtFecha(5).Text = Format(FInicioSeccion, "dd/mm/yyyy")
                    I = Weekday(FInicioSeccion, vbMonday)
                    I = 7 - I
                    If I > 0 Then txtFecha(6).Text = Format(DateAdd("d", I, FInicioSeccion), "dd/mm/yyyy")
                    
                            
                    
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
        
        I = RecuperaValor(Parametros, 1)  'Vemos si es uno o mas de uno
        txtDecimal(0).Visible = I = 1
        Label3(12).Visible = I = 1
        
        txtDecimal(0).Text = RecuperaValor(Parametros, 2)
        Cad = Trim(RecuperaValor(Parametros, 3))
        If Cad = "" Then Cad = "0"
        txtDecimal(1).Tag = ImporteFormateado(Cad)
        txtDecimal(1).Text = Format(txtDecimal(1).Tag, FormatoImporte)
            
        Parametros = RecuperaValor(Parametros, 4) 'Dejo el listado de trabajadores
        If I = 1 Then
            Label4(4).Caption = Parametros
        Else
            Label4(4).Caption = "Trabajadores seleccionados: " & I
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
    Cad = Format(vFecha, "dd/mm/yyyy")
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
    Cad = ""
    frmc.Show vbModal
    Set frmc = Nothing
    
    If Cad <> "" Then
        'Me.Caption = Cad
        txtFecha(Index).Text = Cad
    
        ' *** repasar si el camp es txtAux o Text1 ***
        PonerFoco txtFecha(Index) '<===
        ' ********************************************
    End If
End Sub

Private Sub imgFechasSueltas_Click()
    With frmAsignaHorario
        If optPeriodo(0).Value Then
            I = 0
        Else
            I = 1
        End If
        
        .OtrosDatos = Parametros & DevuelveFechaUltimoProcesado & "|"
        .FeIni = DateAdd("yyyy", I, vEmpresa.FechaInicio)
        .FeFin = DateAdd("yyyy", I, vEmpresa.FechaFin)
        .Opcion = 3
        .Show vbModal
        
        If CadenaDesdeOtroForm <> "" Then Unload Me
            
    End With
End Sub

Private Sub txtDecimal_GotFocus(Index As Integer)
    ConseguirFoco txtDecimal(Index), 3
End Sub

Private Sub txtDecimal_KeyPress(Index As Integer, KeyAscii As Integer)
     KeyPress KeyAscii
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
    KeyPress KeyAscii
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
            For I = 1 To 12
                Cad = Format(CDate("01/" & I & "/2000"), "mmmm")
                cmbFecha.AddItem Cad
                cmbFecha.ItemData(cmbFecha.NewIndex) = I
            Next I
        Else
            'Años partidos
            For I = Month(vEmpresa.FechaInicio) To Month(vEmpresa.FechaFin) + 12
                If (I Mod 12) = 0 Then
                    Cad = Format(CDate("01/12/2000"), "mmmm")
                Else
                    Cad = Format(CDate("01/" & (I Mod 12) & "/2000"), "mmmm")
                End If
                cmbFecha.AddItem Cad
                cmbFecha.ItemData(cmbFecha.NewIndex) = I
            Next I
        
        End If
        
    Case 1
        Me.cboMes(0).Clear
        For I = 1 To 12
            Cad = Format(CDate("01/" & I & "/2000"), "mmmm")
            cboMes(0).AddItem Cad
            cboMes(0).ItemData(cboMes(0).NewIndex) = I 'NO HACE FALTA
        Next I
        I = Month(Now)
        If I = 1 Then
            Me.cboMes(0).ListIndex = 11
            Me.txtNumero(0).Text = Year(Now) - 1
        Else
            Me.cboMes(0).ListIndex = I - 1 - 1
            Me.txtNumero(0).Text = Year(Now)
        End If
        
        
        
    End Select
End Sub

Private Sub txtNumero_GotFocus(Index As Integer)
    ConseguirFoco txtNumero(Index), 3
End Sub


Private Sub txtNumero_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
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
    Cad = "select idtrabajador,fecha," & Cad & ",if(horastrabajadas>9,9,horastrabajadas)"
    Cad = Cad & " from marcajes where " & TipoAlziraEntreFechas
    Cad = Cad & ListaTrabajadores
    
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



'Rs vien ordenado por idtra,fecha
Private Function LocalizaRegistropradas(ByRef RS1 As ADODB.Recordset, idTRa As Long, F As Date) As Boolean
Dim Fin As Boolean
Dim Encontrado As Boolean

    LocalizaRegistropradas = False
    
    'localiza trabajador
    RS1.MoveFirst
    Encontrado = False
    Fin = False
    While Not Fin
        If RS1!idTrabajador = idTRa Then
            Fin = True
            Encontrado = True
        Else
            RS1.MoveNext
            If RS1.EOF Then
                Fin = True
            Else
                If RS1!idTrabajador > idTRa Then Fin = True
            End If
        End If
    Wend
    
    If Not Encontrado Then Exit Function: MsgBox "No localzado registro parada " & idTRa & " " & F: Debug.Print "Stop"
    
    Encontrado = False
    Fin = False
    While Not Fin
        If RS1!Fecha = F Then
            Encontrado = True
            Fin = True
        Else
            RS1.MoveNext
            If RS1.EOF Then
                Fin = True
            Else
                If RS1!Fecha > F Then Fin = True
            End If
        End If
    Wend
    
    LocalizaRegistropradas = Encontrado
    'Localiza fecha
    
    
End Function
