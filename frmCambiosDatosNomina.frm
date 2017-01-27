VERSION 5.00
Begin VB.Form frmCambiosDatosNomina 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar datos trabajador"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameNomina 
      Height          =   3015
      Left            =   120
      TabIndex        =   31
      Top             =   0
      Width           =   7455
      Begin VB.TextBox txtHN 
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtDias 
         Height          =   285
         Index           =   4
         Left            =   5880
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtHN 
         Height          =   285
         Index           =   4
         Left            =   5880
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtHN 
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtDias 
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   6000
         TabIndex        =   16
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton cmdNomina 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   4680
         TabIndex        =   15
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lblIdTra 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   38
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblTra 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   37
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label1 
         Caption         =   "Dias trabajados"
         Height          =   255
         Index           =   9
         Left            =   4080
         TabIndex        =   36
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Anticipos"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   35
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Horas compensables"
         Height          =   255
         Index           =   7
         Left            =   4080
         TabIndex        =   34
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Horas normales"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   33
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Dias nomina"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   32
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame FrameGenerando 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.TextBox txtDias 
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtHN 
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtDias 
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox txtHN 
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtHC 
         Height          =   285
         Index           =   1
         Left            =   3240
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtDias 
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox txtHN 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtHC 
         Height          =   285
         Index           =   2
         Left            =   3240
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2520
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   0
         Left            =   6240
         TabIndex        =   4
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   5040
         TabIndex        =   3
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtBolsa 
         Height          =   285
         Index           =   1
         Left            =   6240
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtBolsa 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   5160
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1320
         Width           =   735
      End
      Begin VB.Line Line2 
         X1              =   5160
         X2              =   7200
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line1 
         X1              =   1320
         X2              =   3960
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label lblTra 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   30
         Top             =   240
         Width           =   5775
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
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1320
         Width           =   1035
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
         Left            =   240
         TabIndex        =   28
         Top             =   1920
         Width           =   1080
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
         Left            =   240
         TabIndex        =   27
         Top             =   2520
         Width           =   645
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
         Left            =   4560
         TabIndex        =   26
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label lblIdTra 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Dias"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   24
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "HN"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   23
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "HC"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   22
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Antes"
         Height          =   255
         Index           =   3
         Left            =   5280
         TabIndex        =   21
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Despues"
         Height          =   255
         Index           =   4
         Left            =   6240
         TabIndex        =   20
         Top             =   960
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmCambiosDatosNomina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    '0.- desde generar
    '1- Ya en nominas
    
Dim B As Boolean
Dim Cad As String

Private Sub cmdAceptar_Click()
Dim T As TextBox
    Cad = ""
    'Todos los campos con valor
    For Each T In txtHC
        If T.Text = "" Then Cad = "C"
    Next
    For Each T In txtHN
        If T.Text = "" And T.Index < 3 Then Cad = "C"
    Next
    For Each T In txtBolsa
        If T.Text = "" Then Cad = "C"
    Next
    For Each T In txtDias
        If T.Text = "" And T.Index < 3 Then Cad = "C"
    Next
    If Cad <> "" Then
        MsgBox "Todos los campos son requeridos", vbExclamation
        Exit Sub
    End If
    
    
    
    'Actualizamos datos
    If MsgBox("Seguro que desea cambiar los datos del mes?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
         
        'Horas oficiles
        Cad = "UPDATE tmpDatosMes SET "
        Cad = Cad & " mesdias = " & txtDias(0).Text
        Cad = Cad & " ,meshoras = " & TransformaComasPuntos(ImporteFormateadoAmoneda(txtHN(0).Text))
        Cad = Cad & " ,DiasTrabajados = " & txtDias(1).Text
        Cad = Cad & " ,horasn = " & TransformaComasPuntos(ImporteFormateadoAmoneda(txtHN(1).Text))
        Cad = Cad & " ,horasc = " & TransformaComasPuntos(ImporteFormateadoAmoneda(txtHC(1).Text))
        
        
        'Saldo lo poenmos a 0
        Cad = Cad & " ,saldodias=0"
        Cad = Cad & " ,saldoh=0"
        
        'Compensadas en NOMINA
        Cad = Cad & " ,diasperiodo = " & txtDias(2).Text
        Cad = Cad & " ,extras = " & TransformaComasPuntos(ImporteFormateadoAmoneda(txtHC(2).Text))
        
        'Bolsa
        Cad = Cad & " ,bolsadespues = " & TransformaComasPuntos(ImporteFormateadoAmoneda(txtBolsa(1).Text))
        
        Cad = Cad & " WHERE Trabajador = " & Me.lblIdTra(0).Caption
        Cad = Cad & " AND mes = " & Mid(Caption, 1, 3)
    
        On Error Resume Next
        conn.Execute Cad
        If Err.Number <> 0 Then
            MuestraError Err.Number, Err.Description
            Exit Sub
        End If
    'Para que refresque el view
    VariableCompartida = "OK"
    Unload Me
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdNomina_Click()
'                'Dias
'            frmCambiosDatosNomina.txtDias(3).Text = !Dias
'            'TRABAJADAS
'            frmCambiosDatosNomina.txtDias(4).Text = DBLet(!diastra, "N")
'            'HN
'            frmCambiosDatosNomina.txtHN(3).Text = !HN
'            'HC
'            frmCambiosDatosNomina.txtHN(4).Text = !HC
'            'Anticipos
'            frmCambiosDatosNomina.txtHN(5).Text = !Anticipos
    Cad = "N"
    If txtHN(3).Text = "" Or txtHN(4).Text = "" Or txtHN(4).Text = "" Then Cad = ""
    If Cad <> "" Then
        If txtDias(3).Text = "" Or txtDias(4).Text = "" Then Cad = ""
    End If
    If Cad = "" Then
        MsgBox "Todos los campos son obligatorios", vbExclamation
        Exit Sub
        
    End If
    
    Cad = "UPDATE Nominas SET Dias = " & txtDias(3).Text
    Cad = Cad & " ,DiasTra = " & txtDias(4).Text
    Cad = Cad & " ,HN = " & TransformaComasPuntos(ImporteFormateadoAmoneda(txtHN(3).Text))
    Cad = Cad & " ,HC = " & TransformaComasPuntos(ImporteFormateadoAmoneda(txtHN(4).Text))
    Cad = Cad & " ,Anticipos = " & TransformaComasPuntos(ImporteFormateadoAmoneda(txtHN(5).Text))
    Cad = Cad & " WHERE idTrabajador = " & Me.lblIdTra(1).Caption
    Cad = Cad & " AND fecha = #" & Format(Caption, FormatoFecha) & "#"
        On Error Resume Next
        conn.Execute Cad
        If Err.Number <> 0 Then
            MuestraError Err.Number, Err.Description
            Exit Sub
        End If
    'Para que refresque el view
    VariableCompartida = "OK"
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    
    
    FrameGenerando.Visible = Opcion = 0
    FrameNomina.Visible = Opcion = 1
    Me.cmdCancel(Opcion).Cancel = True
    Limpiar Me
End Sub





Private Sub txtBolsa_GotFocus(Index As Integer)
    PonerFoco txtBolsa(Index)
End Sub

Private Sub txtBolsa_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub txtBolsa_LostFocus(Index As Integer)

    txtBolsa(Index).Text = Trim(txtBolsa(Index).Text)
    If txtBolsa(Index).Text = "" Then Exit Sub
    B = False
    If IsNumeric(txtBolsa(Index).Text) Then
        If InStr(1, txtBolsa(Index).Text, ",") Then
            'NO hago nada, ya esta formateado
        
        Else
            Cad = TextBoxAImporte(txtBolsa(Index))
        End If
        B = True
    End If

    If Not B Then
        MsgBox "Horas incorrecto. " & txtBolsa(Index).Text, vbExclamation
        txtBolsa(Index).Text = ""
    End If
    
End Sub

Private Sub txtDias_GotFocus(Index As Integer)
    PonerFoco txtDias(Index)
End Sub


Private Sub Txtdias_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub Txtdias_LostFocus(Index As Integer)

    txtDias(Index).Text = Trim(txtDias(Index).Text)
    If txtDias(Index).Text = "" Then Exit Sub
    B = False
    If IsNumeric(txtDias(Index).Text) Then
        txtDias(Index).Text = CInt(txtDias(Index).Text)
        If Val(txtDias(Index).Text) <= 31 Then B = True
    End If
    If Not B Then
        MsgBox "Dias incorrecto. " & txtDias(Index).Text, vbExclamation
        txtDias(Index).Text = ""
    End If
End Sub





Private Sub txtHC_GotFocus(Index As Integer)
    PonerFoco txtHN(Index)
End Sub

Private Sub txtHC_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub txtHC_LostFocus(Index As Integer)
    txtHC(Index).Text = Trim(txtHC(Index).Text)
    If txtHC(Index).Text = "" Then Exit Sub
    B = False
    If IsNumeric(txtHC(Index).Text) Then
        If InStr(1, txtHC(Index).Text, ",") Then
            'NO hago nada, ya esta formateado
        
        Else
            Cad = TextBoxAImporte(txtHC(Index))
        End If
        B = True
    End If

    If Not B Then
        MsgBox "Horas incorrecto. " & txtHC(Index).Text, vbExclamation
        txtHC(Index).Text = ""
    End If
End Sub

Private Sub txtHN_GotFocus(Index As Integer)
    PonerFoco txtHN(Index)
End Sub

Private Sub txtHN_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub


    
Private Sub txtHN_LostFocus(Index As Integer)
    txtHN(Index).Text = Trim(txtHN(Index).Text)
    If txtHN(Index).Text = "" Then Exit Sub
    B = False
    If IsNumeric(txtHN(Index).Text) Then
        If InStr(1, txtHN(Index).Text, ",") Then
            'NO hago nada, ya esta formateado
        
        Else
            Cad = TextBoxAImporte(txtHN(Index))
        End If
        B = True
    End If

    If Not B Then
        MsgBox "Horas incorrecto. " & txtHN(Index).Text, vbExclamation
        txtHN(Index).Text = ""
    End If
End Sub
