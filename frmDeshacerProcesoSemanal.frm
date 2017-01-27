VERSION 5.00
Begin VB.Form frmDeshacerProcesoSemanal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Jorandas horas trabajadas"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2040
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdDeshacer 
      Caption         =   "Deshacer"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Caption         =   "Deshacer proceso generación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Index           =   12
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "frmDeshacerProcesoSemanal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fini As Date
Dim FFin As Date


Private Sub cmdDeshacer_Click()
    
    
    If UCase(InputBox("Contraseña para deshacer el proceso")) = "ARIADNA" Then
        If HacerProceso Then
            
            'Preguntamos si quiere otro , o no
            If MsgBox("Desea contrinuar con otro proceso?", vbQuestion + vbYesNo) = vbYes Then
                CargaDatosDeshacer
            Else
                Unload Me
            End If
        
        End If
    End If
    
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    cmdDeshacer.Enabled = False
    If vUsu.Nivel <= 1 Then
        cmdDeshacer.Enabled = True
        CargaDatosDeshacer
    End If
    
End Sub


Private Sub CargaDatosDeshacer()
Dim Cad As String
    Cad = "select * from jornadassemanalesproceso order by fecha desc"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        Text1(0).Text = miRsAux!Fecha
        Text1(1).Text = miRsAux!fechaini & "  al  " & miRsAux!FechaFin
        Text1(2).Text = miRsAux!Nombre & "   Resgistros: " & miRsAux!sumatorios
        Fini = miRsAux!fechaini
        FFin = miRsAux!FechaFin
    End If
    miRsAux.Close
    Set miRsAux = Nothing
End Sub


Private Function HacerProceso() As Boolean
    Screen.MousePointer = vbHourglass
    conn.BeginTrans
    Set miRsAux = New ADODB.Recordset
    HacerProceso = DesHacer()
    If HacerProceso Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
End Function


Private Function DesHacer() As Boolean
Dim Cad As String

    On Error GoTo eDesHacer
    DesHacer = False
    
    
    'Proceso
    Cad = "DELETE FROM jornadassemanalesalz WHERE fecha >=" & DBSet(Fini, "F")
    Cad = Cad & " AND fecha <= " & DBSet(FFin, "F")
    conn.Execute Cad
    
    
    'Reestablecemos la bolasa dehoras
    Cad = "DELETE from trabajadoresbolsahoras where (IdTrabajador,ParaEmpresa,TipoHora) in (select IdTrabajador,ParaEmpresa,TipoHora"
    Cad = Cad & " from jornadassemanaleshcobolsa where fecha=" & DBSet(Me.Text1(0).Text, "FH") & ")"
    conn.Execute Cad
    
    'Dejamos como estaba la bolsa
    Cad = "INSERT INTO trabajadoresbolsahoras (IdTrabajador ,ParaEmpresa ,TipoHora ,HorasBolsa) "
    Cad = Cad & " select IdTrabajador,ParaEmpresa,TipoHora,HorasBolsa from jornadassemanaleshcobolsa where fecha=" & DBSet(Me.Text1(0).Text, "FH")
    conn.Execute Cad
    
    
    'Borramos la entrada
    Cad = "DELETE from jornadassemanalesproceso where fecha=" & DBSet(Me.Text1(0).Text, "FH")
    conn.Execute Cad
    
    DesHacer = True
    Exit Function
eDesHacer:
    MuestraError Err.Number, Err.Description
End Function
