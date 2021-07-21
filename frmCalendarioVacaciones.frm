VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#17.2#0"; "Codejock.Calendar.v17.2.0.ocx"
Begin VB.Form frmCalendarioVacaciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vacaciones"
   ClientHeight    =   11235
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11235
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   7
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   17
      Tag             =   "0"
      Text            =   "Text1"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   8
      Left            =   13560
      Locked          =   -1  'True
      TabIndex        =   16
      Tag             =   "0"
      Text            =   "Text1"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   5
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   14
      Tag             =   "0"
      Text            =   "Text1"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   6
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   12
      Tag             =   "0"
      Text            =   "Text1"
      Top             =   960
      Width           =   975
   End
   Begin VB.Frame frameButt 
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   3015
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   360
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   635
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Appearance      =   1
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver hoy"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Solicitar dias"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar solicitud"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Conceder dias"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Sailr"
            EndProperty
         EndProperty
         BorderStyle     =   1
         Begin VB.CheckBox chkVistaPrevia 
            Caption         =   "Vista previa"
            Height          =   195
            Index           =   0
            Left            =   8520
            TabIndex        =   11
            Top             =   120
            Width           =   1215
         End
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   4
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   5
      Tag             =   "0"
      Text            =   "Text1"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   12240
      Locked          =   -1  'True
      TabIndex        =   4
      Tag             =   "0"
      Text            =   "Text1"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   960
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   975
   End
   Begin XtremeCalendarControl.DatePicker DatePicker1 
      Height          =   9495
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   14415
      _Version        =   1114114
      _ExtentX        =   25426
      _ExtentY        =   16748
      _StockProps     =   64
      AutoSize        =   0   'False
      ShowTodayButton =   0   'False
      ShowNoneButton  =   0   'False
      ShowNonMonthDays=   0   'False
      Show3DBorder    =   0
      RowCount        =   3
      ColumnCount     =   4
      AskDayMetrics   =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Solicitados"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   3
      Left            =   12120
      TabIndex        =   22
      Top             =   690
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DISPONIBLES"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   8
      Left            =   13320
      TabIndex        =   21
      Top             =   690
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DÍAS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   6600
      TabIndex        =   20
      Top             =   690
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Disfrutados"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   7440
      TabIndex        =   19
      Top             =   690
      Width           =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "Pendientes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   10920
      TabIndex        =   18
      Top             =   690
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " Vacaciones "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Index           =   4
      Left            =   10200
      TabIndex        =   8
      Top             =   420
      Width           =   1365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      X1              =   7440
      X2              =   14520
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label Label1 
      Caption         =   "Aprobados"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   8640
      TabIndex        =   15
      Top             =   690
      Width           =   1110
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   9840
      TabIndex        =   13
      Top             =   690
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   1560
      TabIndex        =   7
      Top             =   690
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   690
      Width           =   660
   End
End
Attribute VB_Name = "frmCalendarioVacaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Trabajador As Long
Private idCalendario As Integer

Dim CadenaVacaciones As String    'LLevara dd/mm/yyyy-00   fecha-situacion  10 + 3
Dim CadObservaciones As String
Dim FechaInicio As Date
Dim DiaSeleccionado As Date
Dim FechaFin As Date
Dim I As Integer


Private Sub DatePicker1_MonthChanged()
 DatePicker1.EnsureVisible FechaInicio
End Sub

Private Sub DatePicker1_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
Dim k As Integer
Dim Situ As Byte
Dim Aux As String

    If Weekday(Day) = vbSunday Then
      '  Set Metrics.Font = Me.Font
       ' Metrics.ForeColor = vbRed
        Metrics.Font.Bold = True
        Metrics.ForeColor = vbRed
    End If
    
    
    k = InStr(1, CadenaVacaciones, Format(Day, "dd/mm/yyyy"))
    If k > 0 Then
        'Stop
        Metrics.Font.Bold = True
        Aux = Mid(CadenaVacaciones, k + 11, 2)
        If Aux = "00" Then
            'Solicitadoas
            Metrics.ForeColor = vbWhite
            Metrics.BackColor = &HC0C0&
            
        Else
            If Aux = "01" Then
                'Aprobadas o peendientes
                    Metrics.ForeColor = vbWhite
                    Metrics.BackColor = &HC00000
                
            Else
                If Aux = "02" Then
                    'Aprobadas pendientes
                    Metrics.ForeColor = vbWhite
                    Metrics.BackColor = &H8000&
                Else
                    'Vacaciones calendario laborables
                    Metrics.ForeColor = vbRed
                    
                End If
            End If
        End If
    End If
    
    
End Sub




Private Sub DatePicker1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim CurDate
Dim k As Long
Dim J As Long
    CurDate = DatePicker1.HitTest
    
    If CurDate <> 0 Then
        If DiaSeleccionado <> CurDate Then
            DiaSeleccionado = CurDate
            k = InStr(1, CadObservaciones, Format(CurDate, "dd/mm/yyyy"))
            If k > 0 Then
                J = InStr(k, CadObservaciones, "|")
                k = k + 10
                If J > 0 Then DatePicker1.ToolTipText = Mid(CadObservaciones, k, J - k)
            End If
        Else
            'DatePicker1.ToolTipText = ""
        End If
    Else
        DatePicker1.ToolTipText = ""
    End If
End Sub

Private Sub DatePicker1_SelectionChanged()
Dim MoverSel As Boolean

    MoverSel = False
    
    If DatePicker1.Selection.BlocksCount > 0 Then
        If DatePicker1.Selection.Blocks(0).DateBegin < FechaInicio Then
            MoverSel = True
        Else
            If DatePicker1.Selection.Blocks(0).DateEnd > FechaFin Then MoverSel = True
        End If
        
        If MoverSel Then
        
            MsgBox "Fecha fuera temporada", vbExclamation
            DatePicker1.EnsureVisible vEmpresa.FechaFin
            DatePicker1.EnsureVisible FechaInicio
            DatePicker1.ClearSelection
        End If
    End If
 
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    I = 0
    Me.Icon = frmPpal.Icon
    
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 15
        .Buttons(3).Image = 3
        .Buttons(4).Image = 5
        .Buttons(5).Image = 17
        .Buttons(7).Image = 11
        'el 13 i el 14 son separadors
    '    .Buttons(btnPrimero).Image = 6  'Primero
    '    .Buttons(btnPrimero + 1).Image = 7 'Anterior
    '    .Buttons(btnPrimero + 2).Image = 8 'Siguiente
    '    .Buttons(btnPrimero + 3).Image = 9 'Último
    End With

    Me.frameButt.BorderStyle = 0

    FechaInicio = vEmpresa.FechaInicio
    If Month(vEmpresa.FechaInicio) <> 1 Then
        If True Then FechaInicio = "01/01/" & Year(vEmpresa.FechaInicio)
    End If
            
    FechaFin = DateAdd("yyyy", 1, FechaInicio)
    FechaFin = DateAdd("d", -1, FechaFin)
    DatePicker1.HighlightToday = False
    
    'Trabajador = 1
    PonerDatosTrabajador

    
    
    
    
    DatePicker1.EnsureVisible FechaInicio
    
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Cad As String
Dim F As Date
Dim ColDias As Collection
Dim Aux As String
Dim k As Integer


    If Button.Index = 7 Then
        Unload Me
        Exit Sub
    End If
    
    
    
    
    'Previo para solicitar, anular
    If Button.Index = 3 Or Button.Index = 4 Then
        'Va a solicitar los dias
        If DatePicker1.Selection.BlocksCount <> 1 Then
            MsgBox "Seleccione uno (o más ) días", vbExclamation
            Exit Sub
        End If
    End If
    
    
    If Button.Index = 1 Then
        'Ver hoy
        If Me.DatePicker1.HighlightToday = True Then
            Me.DatePicker1.HighlightToday = False
            Button.ToolTipText = "Mostrar hoy"
        Else
            Me.DatePicker1.HighlightToday = True
            Button.ToolTipText = "Ocultar hoy"
        End If
        
        DatePicker1.RedrawControl
        
        
    ElseIf Button.Index = 3 Then
        
        
        'So0licitar
        Cad = ""
        Set ColDias = New Collection
        I = 0
        F = DatePicker1.Selection.Blocks(0).DateBegin
        Do
            'Dias laborables
            If Weekday(F, vbMonday) <= 5 Then
                If InStr(1, CadenaVacaciones, Format(F, "dd/mm/yyyy")) > 0 Then
                    
            
                Else
                    ColDias.Add F
                    Cad = Cad & "     " & Format(F, "dd/mm/yyyy")
                    I = I + 1
                    If (I Mod 5) = 0 Then Cad = Cad & vbCrLf
                End If
            End If
            F = DateAdd("d", 1, F)
        Loop Until F > DatePicker1.Selection.Blocks(0).DateEnd
        
        If I = 0 Then
            MsgBox "Ningun dia laborable en la selecccion", vbExclamation
        Else
        
            If I > Me.Text1(8).Tag Then MsgBox String(40, "*") & vbCrLf & "Solicita mas dias que disponibles " & vbCrLf & I & "    //   " & Text1(8).Text, vbCritical
            
        
            Cad = "Va a solicitar los siguientes dias (" & I & ") : " & vbCrLf & Cad & vbCrLf & vbCrLf & "¿Continuar?"
            If MsgBox(Cad, vbQuestion + vbYesNoCancel) = vbYes Then
                Screen.MousePointer = vbHourglass
                'Grabamos y resfrescamos
                Cad = ""
                For I = 1 To ColDias.Count
                    'trabajadoresvacaciones(idtrabajador,fecha,situacion,observa)
                    Cad = Cad & ", (" & Trabajador & "," & DBSet(ColDias.Item(I), "F") & ",0," & DBSet("Solicitada desde Aripres.  " & vUsu.Login, "T") & ")"
                Next
                Cad = Mid(Cad, 2)
                Cad = "INSERT INTO trabajadoresvacaciones(idtrabajador,fecha,situacion,observa) VALUES " & Cad
                EjecutaSQL Cad
                
                DatePicker1.ClearSelection
                
                CargarDias
                DatePicker1.RedrawControl
                
                Me.Refresh
                Screen.MousePointer = vbDefault
            End If
        End If
    ElseIf Button.Index = 4 Or Button.Index = 5 Then
        
        'Quitar // conceder solicitudes
                
        Cad = ""
        Set ColDias = New Collection
        I = 0
        F = DatePicker1.Selection.Blocks(0).DateBegin
        Do
            'Dias laborables
            If Weekday(F, vbMonday) <= 5 Then
                k = InStr(1, CadenaVacaciones, Format(F, "dd/mm/yyyy"))
                If k = 0 Then
                    
            
                Else
                
                    
                    Aux = Mid(CadenaVacaciones, k + 11, 2)
                    If Aux = "00" Then
                        'OK es una solicitud. SE PUEDE quitar
                        ColDias.Add F
                        Cad = Cad & "     " & Format(F, "dd/mm/yyyy")
                        I = I + 1
                        If (I Mod 5) = 0 Then Cad = Cad & vbCrLf
                    End If
                End If
            End If
            F = DateAdd("d", 1, F)
        Loop Until F > DatePicker1.Selection.Blocks(0).DateEnd
        
        If I = 0 Then
            MsgBox "Ningun dia solicitado en la selecccion", vbExclamation
        Else
            Cad = " los siguientes dias (" & I & ") : " & vbCrLf & Cad & vbCrLf & vbCrLf & "¿Continuar?"
            Cad = IIf(Button.Index = 4, "Va a anular la solicitud de ", "Va a CONCEDER ") & Cad
            If MsgBox(Cad, vbQuestion + vbYesNoCancel) = vbYes Then
                Screen.MousePointer = vbHourglass
                'Grabamos y resfrescamos
                Cad = ""
                For I = 1 To ColDias.Count
                    'trabajadoresvacaciones(idtrabajador,fecha,situacion,observa)
                    Cad = Cad & ", " & DBSet(ColDias.Item(I), "F")
                Next
                Cad = Mid(Cad, 2)
                Cad = " WHERE idtrabajador = " & Trabajador & " AND fecha in (" & Cad & ")"
                
                If Button.Index = 4 Then
                    Cad = "DELETE FROM  trabajadoresvacaciones " & Cad
                Else
                    Cad = "UPDATE trabajadoresvacaciones SET situacion=1 " & Cad
                End If
                EjecutaSQL Cad
                
                DatePicker1.ClearSelection
                CargarDias
                DatePicker1.RedrawControl
                Screen.MousePointer = vbDefault
                Me.Refresh
            End If
        End If

        
        
    
    ElseIf Button.Index = 7 Then
    
    
    
    
    Else
        Stop
        
    
    End If
    
    Set ColDias = Nothing
End Sub

Private Sub Limpia()
    For I = 0 To Text1.Count - 1
        Text1(I).Text = ""
        Text1(I).Tag = 0
        Text1(I).Locked = True
    Next

End Sub

Private Sub PonerDatosTrabajador()
Dim Cad As String
    
    On Error GoTo ePonerDatosTrabajador
    
    
    
    
    Limpia
    idCalendario = -1
    
    Cad = "Select IdTrabajador,NomTrabajador,DiasVacaciones,idcal FROM trabajadores where IdTrabajador =" & Trabajador
    'Cad = Cad & " AND fecha between  " & DBSet(F1, "F") & " AND " & DBSet(F2, "F")
    Set miRsAux = New ADODB.Recordset
    
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Text1(0).Text = Format(miRsAux!idTrabajador, "0000")
    Text1(1).Text = miRsAux!nomtrabajador
    Text1(2).Text = miRsAux!DiasVacaciones
    Text1(2).Tag = miRsAux!DiasVacaciones
    idCalendario = miRsAux!idCal
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    CargarDias
    
ePonerDatosTrabajador:
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
        Limpia
        Set miRsAux = Nothing
    End If
    
    
End Sub

Private Sub CargarDias()
Dim Cad As String
Dim Situacion As Byte

    
    For Situacion = 3 To 8
        Text1(Situacion).Tag = 0
    Next
    


    Cad = "Select  situacion , fecha ,coalesce(observa,'') observa,idtrabajador FROM trabajadoresvacaciones where IdTrabajador =" & Trabajador
    Cad = Cad & " AND fecha between  " & DBSet(FechaInicio, "F") & " AND " & DBSet(FechaFin, "F")
    Set miRsAux = New ADODB.Recordset
    CadenaVacaciones = ""
    CadObservaciones = ""
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        
        Situacion = 0
        If miRsAux!Situacion = 1 Then
            If miRsAux!Fecha <= Now Then
                Situacion = 1
                Text1(4).Tag = Text1(4).Tag + 1
            Else
                Situacion = 2
                Text1(5).Tag = Text1(5).Tag + 1
            End If
            
        Else
            'Soliocitado
            Text1(3).Tag = Text1(3).Tag + 1
        End If
        CadenaVacaciones = CadenaVacaciones & Format(miRsAux!Fecha, "dd/mm/yyyy") & "-" & Format(Situacion, "00")
        If Situacion = 0 And miRsAux!Observa <> "" Then     'me he garantizado que no sea null con coalesce
            'Es una solicutd
            CadObservaciones = CadObservaciones & Format(miRsAux!Fecha, "dd/mm/yyyy") & miRsAux!Observa & "|"
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Pendientes
    Text1(6).Tag = Text1(4).Tag + Val(Text1(5).Tag)
    Text1(7).Tag = Text1(2).Tag - Text1(6).Tag
    'Disponibles
    'pendientes - solicitados
    Text1(8).Tag = Text1(7).Tag - Text1(3).Tag
    
    For Situacion = 2 To 8
        
        If Text1(Situacion).Tag = 0 Then
            Text1(Situacion).Text = ""
        Else
            Text1(Situacion).Text = Format(Text1(Situacion).Tag, "00")
        End If
    Next
    
    
    Cad = "Select   fecha ,Descripcion FROM calendariof  where idcal=" & idCalendario
    Cad = Cad & " AND fecha between  " & DBSet(FechaInicio, "F") & " AND " & DBSet(FechaFin, "F")
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        CadenaVacaciones = CadenaVacaciones & Format(miRsAux!Fecha, "dd/mm/yyyy") & "-" & Format(3, "00")
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub
