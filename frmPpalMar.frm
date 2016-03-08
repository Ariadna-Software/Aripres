VERSION 5.00
Begin VB.Form frmPpalMar 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   1995
   ClientTop       =   1995
   ClientWidth     =   8580
   Icon            =   "frmPpalMar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPpalMar.frx":030A
   ScaleHeight     =   3135
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   2280
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   6240
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   6240
      TabIndex        =   5
      Top             =   2100
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   4
      Left            =   4830
      TabIndex        =   4
      Top             =   2685
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajador"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   3
      Left            =   4635
      TabIndex        =   3
      Top             =   2145
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Generar marcaje"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Index           =   2
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ariadna Software"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Generar marcaje"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   480
      Index           =   0
      Left            =   4620
      TabIndex        =   0
      Top             =   120
      Width           =   3795
   End
End
Attribute VB_Name = "frmPpalMar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vUsu As Usuario

Private Sub Form_Load()
    PonLabel
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then
            Unload Me
        End If
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)
    
    'Comprobamos si los dos estan con datos
    If Text1(0).Text <> "" And Text1(1).Text <> "" Then
        'Probar conexion usuario
        Validar
    End If
        
    
End Sub



Private Sub Validar()

Dim Ok As Byte

    'Validaremos el usuario y despues el password
    Set vUsu = New Usuario
    
    If vUsu.Leer(Text1(0).Text) = 0 Then
        'Con exito
        If vUsu.PasswdPROPIO = Text1(1).Text Then
            Ok = 0
        Else
            Ok = 1
        End If

    Else
        Ok = 2
    End If
    
    If Ok <> 0 Then
        MsgBox "Usuario-Clave Incorrecto", vbExclamation
        Set vUsu = Nothing
        Text1(1).Text = ""
        Text1(0).SetFocus
    Else
        'OK
        Acciones
        
        Screen.MousePointer = vbHourglass
        Unload Me
    End If

End Sub



Private Sub Acciones()
Dim C As String

    On Error GoTo MAL
    C = "Select max(secuencia),curdate(),curtime() from entradafichajes"
    RS.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        C = "0"
    Else
        If IsNull(RS.Fields(0)) Then
            C = "0"
        Else
            C = RS.Fields(0)
        End If
    End If
    
    C = CStr(Val(C) + 1)
    C = "INSERT INTO entradafichajes (Secuencia, idTrabajador, Fecha, Hora, HoraReal, idInci) VALUES (" & C
    C = C & "," & vUsu.Codigo & ",'" & Format(RS.Fields(1), "yyyy-mm-dd") & "','" & Format(RS.Fields(2), "hh:mm:ss")
    C = C & "','" & Format(RS.Fields(2), "hh:mm:ss") & "',0)"
    Conn.Execute C
    Exit Sub
MAL:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Timer1_Timer()
    
    
        Fecha = DateAdd("s", 1, Fecha)
        PonLabel
End Sub


Private Sub PonLabel()
        Label1(2).Caption = "Fecha / hora en servidor: " & Format(Fecha, "dd/mm/yyyy") & "  " & Format(Fecha, "hh:mm:ss")
        Label1(2).Refresh
End Sub
