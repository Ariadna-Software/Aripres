VERSION 5.00
Begin VB.Form frmRelojZKTeco 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ZKTeco"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Leer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   5895
   End
   Begin VB.Image Image1 
      Height          =   3450
      Left            =   0
      Picture         =   "frmZKTeco.frx":0000
      Top             =   120
      Width           =   6000
   End
End
Attribute VB_Name = "frmRelojZKTeco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim Cad As String
Dim NF As Integer
Private Sub Command1_Click()
    
    Unload Me
   
End Sub

Private Sub Command2_Click()

    

    Screen.MousePointer = vbHourglass
    If LeerDatos Then
        
        Label1.Caption = "Lectura realizada correctamente"
        Command2.Enabled = False
    Else
        Me.Label1.Caption = ""
    
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    Label1.Caption = ""
End Sub




Private Function LeerDatos() As Boolean
Dim cArc As Collection
Dim fil As Collection
Dim I As Integer
Dim Fic As String
    On Error GoTo eLeerDatos
    LeerDatos = False
    
    Label1.Caption = "Comprobando carpeta"
    Label1.Refresh
    If Dir(vEmpresa.DirMarcajes, vbDirectory) = "" Then Err.Raise 513, , "No existe la carpeta " & vEmpresa.DirMarcajes
    
    Set cArc = New Collection
    
    Cad = Dir(vEmpresa.DirMarcajes & "\" & vEmpresa.NomFich, vbArchive)    ' Recupera la primera entrada.
    Do While Cad <> ""   ' Inicia el bucle.
        Fic = Cad
        Label1.Caption = "Ver fic: " & Fic
        Label1.Refresh
        If preComprobacion Then
            cArc.Add CStr(Fic)
        Else
            Err.Raise 513, Cad, Cad
        End If
      
       Cad = Dir   ' Obtiene siguiente entrada.
    Loop

        
    Set fil = New Collection
    For I = 1 To cArc.Count
        Label1.Caption = "Bloquear fichero: " & cArc.Item(I)
        Label1.Refresh
        'Renombramos para bloquear el fichero
        Cad = cArc.Item(I)
        NumRegElim = InStrRev(Cad, ".")
        Cad = Mid(Cad, 1, NumRegElim) & "ari"
        Cad = vEmpresa.DirMarcajes & "\" & Cad
        Name vEmpresa.DirMarcajes & "\" & cArc.Item(I) As Cad
        fil.Add CStr(Cad)
       
        
    Next
   
    
    
    If fil.Count > 0 Then
        
        DoEvents
    
        
        For I = 1 To fil.Count
            Cad = DevuelveDesdeBD("max(secuencia)", "entradafichajes", "1", "1")
            NumRegElim = Val(Cad) + 1
        
        
            Cad = fil.Item(I)
            If InsertarDatos Then
                Cad = fil.Item(I)
                MatarFichero CStr(Cad)
            End If
            
        Next I
        
        
        
        Label1.Caption = "Entradas repetidas"
        Label1.Refresh
    
        
        EntradasRepetidasProceso Me.Label1
        
    End If
    
    
    If cArc.Count = 0 Then MsgBox "Ningun fichero a procesar", vbExclamation
    
eLeerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set fil = Nothing
    Set cArc = Nothing
End Function



Private Function preComprobacion() As Boolean
Dim Tr As String
Dim Aux As String
    On Error GoTo epreComprobacion
    
    Cad = vEmpresa.DirMarcajes & "\" & Cad
    
    NF = FreeFile
    Tr = "|"
    Open Cad For Input As #NF
    While Not EOF(NF)
        Line Input #NF, Cad
        'Primera linea: <PSD Copy File v.2.0>
        'Lineas de marca: 100720181228  Ejemplo len=12
        If UCase(Mid(Cad, 1, 5)) <> "<PSD " Then
            If Len(Cad) > 12 Then
                'Ejemplo. Va con tabulaciones
                ' tr  fe          hora
                '5   20180710    124851  1   0
                Cad = Mid(Cad, 1, InStr(1, Cad, Chr(9)) - 1)
                If InStr(1, Tr, "|" & Cad & "|") = 0 Then Tr = Tr & Cad & "|"
        
            End If
        End If
    Wend
    Close #NF
    
    Tr = Mid(Tr, 2) 'quitamos el primer |
    If Len(Tr) > 1 Then
        Do
            NF = InStr(1, Tr, "|")
            If NF = 0 Then
                Tr = ""
            Else
                Cad = Mid(Tr, 1, NF - 1)
                Tr = Mid(Tr, NF + 1)
            
                Aux = DevuelveDesdeBD("idtrabajador", "trabajadores", "numtarjeta", Cad, "T")
                If Aux = "" Then
                    Cad = "NO existe trabajador en BD: " & Cad & vbCrLf
                    Exit Function
                End If
            End If
        Loop Until Tr = ""
    End If
    preComprobacion = True
    Exit Function
epreComprobacion:
    Cad = Err.Description
    Err.Clear
    If NF > 0 Then Close #NF
End Function




Private Function InsertarDatos() As Boolean
Dim Aux As String
Dim Inser As String
Dim J As Integer

    On Error GoTo epreComprobacion
    
    Label1.Caption = Cad
    Label1.Refresh
    espera 0.5
    
    'Borramos de la temporal en aripres
    conn.Execute "DELETE FROM tmppresencia where codusu =" & vUsu.Codigo
    
    
    
    NF = FreeFile
    Open Cad For Input As #NF
    While Not EOF(NF)
        Line Input #NF, Cad
        'Primera linea: <PSD Copy File v.2.0>
        'Lineas de marca: 100720181228  Ejemplo len=12
        If UCase(Mid(Cad, 1, 5)) <> "<PSD " Then
            If Len(Cad) > 12 Then
                'Ejemplo. Va con tabulaciones
                ' tr  fe          hora
                '5   20180710    124851  1   0
                'tmppresencia(Id,idtra,Fecha,H1,codusu,Incidencias)
                    
                J = InStr(1, Cad, Chr(9))
                Aux = Mid(Cad, 1, J - 1)
                Cad = Mid(Cad, J + 1)
                NumRegElim = NumRegElim + 1
                '
                Inser = Inser & ", (" & NumRegElim & "," & Aux & ","
                Aux = Mid(Cad, 1, 8)
                Aux = Mid(Aux, 7, 2) & "/" & Mid(Aux, 5, 2) & "/" & Mid(Aux, 1, 4)
                Inser = Inser & DBSet(Aux, "F") & ",'"
                Aux = Mid(Cad, 10, 6)
                Aux = Mid(Aux, 1, 2) & ":" & Mid(Aux, 3, 2) & ":" & Mid(Aux, 5, 2)
                Inser = Inser & Aux & "'," & vUsu.Codigo & ",0)"
                
            End If
        End If
    Wend
    Close #NF
    NF = -1
    
    
    
    
    If Inser <> "" Then
        Label1.Caption = "Insertando"
        Label1.Refresh
        
        Inser = Mid(Inser, 2)
        Inser = "insert tmppresencia(Id,idtra,Fecha,H1,codusu,Incidencias) VALUES " & Inser
        conn.Execute Inser
        espera 0.5
    
    
        
        Set miRsAux = New ADODB.Recordset
        
        Cad = "    select distinct idtra from tmppresencia WHERE codusu = " & vUsu.Codigo
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            Cad = miRsAux!idTRa
            Aux = DevuelveDesdeBD("idtrabajador", "trabajadores", "numtarjeta", Cad, "T")
            If Aux = "" Then Err.Raise 513, , "NO existe trabajador. Tarjeta: " & Cad
            Aux = "UPDATE tmppresencia set seccion =" & Aux & " WHERE codusu =" & vUsu.Codigo & " AND idtra=" & miRsAux!idTRa
            conn.Execute Aux
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        Cad = "tmppresencia where codusu =" & vUsu.Codigo
        Cad = "SELECT id,seccion,fecha ,h1,incidencias,h1 FROM " & Cad & " ORDER BY fecha,h1"
        Cad = "INSERT INTO entradafichajes(Secuencia,idTrabajador,Fecha,Hora,idInci,HoraReal) " & Cad
        conn.Execute Cad
        
    
    
    
    End If
        
    InsertarDatos = True
    Exit Function
epreComprobacion:
    MuestraError Err.Number, Err.Description
    If NF > 0 Then Close #NF
    Set miRsAux = Nothing
End Function

Private Sub MatarFichero(Origen As String)
    On Error GoTo eM
    Cad = vEmpresa.DirProcesados & "\" & Format(Now, "yyyymmdd_hhnnss") & ".dat"
    
    FileCopy Origen, Cad
    Kill Origen
       
    Exit Sub
eM:
    MsgBox "Error eliminando fichero. El programa continuará. Avise soporte técnico", vbCritical
    Err.Clear

End Sub
