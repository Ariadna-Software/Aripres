VERSION 5.00
Begin VB.Form frmSoloHora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Introducción marcajes"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "frmSoloHora.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5235
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAcabalgado 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   315
      Index           =   1
      Left            =   3720
      TabIndex        =   6
      Top             =   2280
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   315
      Index           =   0
      Left            =   2400
      TabIndex        =   5
      Top             =   2280
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   2220
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1500
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   660
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   1155
   End
   Begin VB.Label Label3 
      Caption         =   "Puedes  poner los dos puntos de las horas con el punto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   1140
      Picture         =   "frmSoloHora.frx":6852
      Top             =   1740
      Width           =   240
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Nuevo marcaje"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   4755
   End
   Begin VB.Label Label1 
      Caption         =   "Incidencia"
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
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   1740
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Hora"
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
      Left            =   180
      TabIndex        =   0
      Top             =   900
      Width           =   735
   End
End
Attribute VB_Name = "frmSoloHora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1

Public Event Seleccionar(vHora As Date, vInci As Integer, Acabagado As Byte)

Public Hora As String
Public Inci As Integer
Public CadInci As String
Public TipoAcabalgada As Byte   ' 0: normal     1: Hora inferior a 00:00    2: hora superior a 23:59:59
Private PrimeraVez As Boolean


Private Sub Command1_Click(Index As Integer)
Dim Acabalg As Byte

    If Index = 0 Then
        If DatosOk Then
            If Me.chkAcabalgado.Visible Then
                If Hora = "" Then
                    'Nuevo
                    'esta marcado el check
                    If Me.chkAcabalgado.Value = 1 Then
                        If vEmpresa.AcabalgadoDiaInicio Then
                            Acabalg = 1
                        Else
                            Acabalg = 2
                        End If
                    End If
                Else
                    'Al modificar NO dejo cambiar el tipo de hora acabalgada
                    Acabalg = TipoAcabalgada
                End If
                    
            Else
                Acabalg = 0
            End If
            RaiseEvent Seleccionar(CDate(Text1(0).Text), CInt(Text1(1).Text), Acabalg)
            Else
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Me.Top = 1000
    End If
End Sub

Private Sub Form_Load()
Dim TipoAca As Byte
    Text1(0).Text = Hora
    Text1(1).Text = Inci
    Text1(2).Text = CadInci
    If Hora = "" Then
        Label2.Caption = "Nuevo marcaje"
        chkAcabalgado.Enabled = True
        Else
        Label2.Caption = "Modificar"
        chkAcabalgado.Enabled = False
    End If
    chkAcabalgado.Visible = False
    If vEmpresa.HorarioNocturno2 Then
        If vEmpresa.QueEmpresa = 2 Then
            
            
            
            If Hora = "" Then
                chkAcabalgado.Visible = True
                If vEmpresa.AcabalgadoDiaInicio Then
                    TipoAca = 2
                Else
                    TipoAca = 1
                End If
            Else
                chkAcabalgado.Value = 0
                
                If Me.TipoAcabalgada > 0 Then
                    TipoAca = Me.TipoAcabalgada
                    chkAcabalgado.Value = 1
                    chkAcabalgado.Visible = True
                Else
                    If vEmpresa.AcabalgadoDiaInicio Then
                        TipoAca = 2
                    Else
                        TipoAca = 1
                    End If
                End If
            End If
            
            'Label check
            If TipoAca = 2 Then
                chkAcabalgado.Caption = "Horas del dia anterior"
            Else
                chkAcabalgado.Caption = "Horas del dia siguiente"
            End If

            
        End If
    End If
    
    
    
    
    PrimeraVez = True
End Sub


Private Function DatosOk() As Boolean
    DatosOk = False
    If Text1(0).Text = "" Then
        MsgBox "Escriba una fecha", vbExclamation
        Exit Function
    End If
    
    If Not IsDate(Text1(0).Text) Then
        MsgBox "No es una fecha válida", vbExclamation
        Exit Function
    End If
    'Compruebo que en la cadena hay dos puntos
    If InStr(1, Text1(0).Text, ":") = 0 Then
        MsgBox "No es un hora válida", vbExclamation
        Exit Function
    End If
    
    'Ahora la incidencia
    If Text1(1).Text = "" Then
        MsgBox "Seleccione una incidencia.", vbExclamation
        Exit Function
    End If
    
    If Not IsNumeric(Text1(1).Text) Then
        MsgBox "Número de incidencia incorrecto.", vbExclamation
        Exit Function
    End If
    
    If CInt(Text1(1).Text) < 0 Then
        MsgBox "Número de incidencia incorrecta.", vbExclamation
        Exit Function
    End If
    DatosOk = True
End Function

Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
Text1(1).Text = vCodigo
Text1(2).Text = vCadena
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Text1(1).Text = RecuperaValor(CadenaDevuelta, 1)
    Text1(2).Text = RecuperaValor(CadenaDevuelta, 2)
End Sub

Private Sub Image1_Click()
            
            CadInci = "Código|idinci|N||25·"
            CadInci = CadInci & "Descrpción|nominci|T||60·"
            Set frmB = New frmBuscaGrid
            frmB.vTabla = "incidencias"
            frmB.vCampos = CadInci
            frmB.vDevuelve = "0|1|"
            frmB.vSelElem = 0
            frmB.vTitulo = "INCIDENCIAS"
            frmB.Show vbModal
            CadInci = ""
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim I As Integer
Dim C As String

Select Case Index
Case 0
    Do
        I = InStr(1, Text1(0).Text, ".")
        If I > 0 Then
            C = Mid(Text1(0).Text, I + 1)
            If Len(C) = 1 Then
                If Val(C) > 5 Then
                    C = "0" & C
                Else
                    C = C & "0"
                End If
            End If
            Text1(0).Text = Mid(Text1(0).Text, 1, I - 1) & ":" & C
        End If
    Loop While I <> 0
    
    If Text1(0).Text <> "" Then
        If Not IsDate(Text1(0).Text) Then
            MsgBox "Error en el campo hora: " & Text1(0).Text, vbExclamation
            Text1(0).Text = ""
            PonerFoco Text1(0)
    
        End If
    End If
Case 1
    If Text1(1).Text = "" Then Exit Sub
    If Not IsNumeric(Text1(1).Text) Then
        MsgBox "La incidencia tiene que ser un número.", vbExclamation
        Text1(1).Text = -1
        Text1(2).Text = ""
        PonerFoco Text1(1)
        Exit Sub
    End If
    'Incidencia
    C = DevuelveDesdeBD("nominci", "incidencias", "idinci", Text1(1).Text, "N")
    
    If C = "" Then
        
        Text1(2).Text = "NO EXISTE :" & Text1(1).Text
        Text1(1).Text = 0
        PonerFoco Text1(1)
        Else
            Text1(2).Text = C
    End If
    
End Select

End Sub

Private Sub KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
         Unload Me
    End If
End Sub
