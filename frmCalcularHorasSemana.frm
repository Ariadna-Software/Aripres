VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCalcularHorasSemana 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Horas semanales"
   ClientHeight    =   9870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12735
   Icon            =   "frmCalcularHorasSemana.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9870
   ScaleWidth      =   12735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   9360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Ajustar"
      Height          =   375
      Left            =   8280
      TabIndex        =   2
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9720
      TabIndex        =   1
      Top             =   9360
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   8895
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   15690
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
         Text            =   "Codigo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   2011
      EndProperty
   End
End
Attribute VB_Name = "frmCalcularHorasSemana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TodasSecciones As Boolean

Dim ColumnaDondeEmpiezanHoras As Byte
Dim J As Integer


Dim CuantosTiposHoraTrabaja As Byte  'FALTA meter en parametros
Dim IdSeccion As Byte
Dim Sumas() As Currency



Dim InicioProceso As Date
Dim FinProceso As Date
Dim FechaInicioSemana As Date
Dim Previsualizacion As Boolean
Dim ProcesoDeNominasAlzira As Boolean  'Hay secciones que son conteos de horas
Dim PrimeraVez As Boolean


Private Sub cmdAceptar_Click()
    
    
    'msgbox y comrpobaciones
    If Not Previsualizacion Then
        If MsgBox("�Desea cerrar el computo del periodo ?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    
    
    'ACeptamos
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    
    If Previsualizacion Then
        AjustarHoras
        Previsualizacion = False
        If Me.ListView1.ListItems.Count > 0 Then
            cmdAceptar.Caption = "Guardar"
            cmdImprimir.Visible = True
        
        Else
            
            HacerGeneracionPeriodo
            Me.cmdAceptar.Enabled = False
        End If
    Else
        If Me.ListView1.ListItems.Count = 0 Then
            MsgBox "Ningun dato para generar", vbExclamation
            Exit Sub
        End If
        HacerGeneracionPeriodo

        Me.cmdAceptar.Enabled = False
    End If
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
    If Me.cmdAceptar.Enabled = False Then Unload Me
    
End Sub




Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdImprimir_Click()

    
    If CargarDatosImpresion Then
        Me.Tag = CadenaDesdeOtroForm   'en el frmimp se pone a ""
        With frmImprimir
            .FormulaSeleccion = "{tmpcombinada.codusu} = " & vUsu.Codigo
            .NombreRPT100 = "PreAjusteSemana.rpt"
            .Titulo100 = "Pre ajuste semanal"
            .OtrosParametros = "FechaFin= ""Desde " & RecuperaValor(CadenaDesdeOtroForm, 1) & " hasta " & RecuperaValor(CadenaDesdeOtroForm, 2) & """|"
            .Opcion = 100
            .NumeroParametros = 1
            .Show vbModal
        End With
        CadenaDesdeOtroForm = Me.Tag
        Me.Tag = ""
    End If
    
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        CargaDatos  'Horas trabajadas para la coperativa
        
        Previsualizacion = True
        
        If vEmpresa.CompensaHorasNominaMES Then cmdAceptar_Click
    End If
End Sub

Private Sub Form_Load()

    
    
    Me.Icon = frmMain.Icon
    
    
    PrimeraVez = True
End Sub


'Dim ParaLaCooperativa As Byte    ** YA NO LA PASAMOS.
Private Sub CargaDatos()
Dim Cad As String
Dim idTrabajador As Long
Dim Fecha As Date
Dim IT As ListItem
Dim diasTrabajados As Byte 'Laborables semana
Dim F2 As Date
Dim PintaColumnaDiasNominaAnterior As Boolean

    Set miRsAux = New ADODB.Recordset
    
    
    'Coje los festivos del CALENDARIO 1, pero para la seccion. Obtengo la seccion
    
    Cad = DevuelveDesdeBD("min(idtrabajador)", "tmphorastipoalzira", "codusu", CStr(vUsu.Codigo))
    If Cad <> "" Then
        idTrabajador = Val(DevuelveDesdeBD("idcal", "trabajadores", "idtrabajador", Cad))
        Cad = DevuelveDesdeBD("seccion", "trabajadores", "idtrabajador", Cad)
    Else
        Cad = "1"
    End If
    IdSeccion = Val(Cad)
    'Ver si la seccion tiene proceso de nominas compensables estructurlaes...
    Cad = DevuelveDesdeBD("Nominas", "secciones", "idseccion", Cad)
    ProcesoDeNominasAlzira = Cad = "1"
    If vEmpresa.CompensaHorasNominaMES Then ProcesoDeNominasAlzira = False
    
    
    
    
    Cad = RecuperaValor(CadenaDesdeOtroForm, 1)
    InicioProceso = CDate(Cad)
    Cad = RecuperaValor(CadenaDesdeOtroForm, 2)
    FinProceso = CDate(Cad)
    
    
    FechaInicioSemana = InicioProceso
    If ProcesoDeNominasAlzira Then
        'Es proceso semanal,
        'buscare el primer dia de la semana
        J = Weekday(FechaInicioSemana, vbMonday) - 1
        FechaInicioSemana = DateAdd("d", -J, FechaInicioSemana)
    End If


    
    If ProcesoDeNominasAlzira Then
        'Aunque pida un periodo corto, siempre es una semana trabajada, 5 dias
        diasTrabajados = 5
        Fecha = DateAdd("d", 6, FechaInicioSemana)
        Cad = " fecha between " & DBSet(FechaInicioSemana, "F") & " AND " & DBSet(Fecha, "F")
        Cad = Cad & " AND idcal=" & idTrabajador
        Cad = "Select * from calendariof WHERE " & Cad
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic
        
        Cad = ""
        While Not miRsAux.EOF
            
            If Weekday(miRsAux!Fecha, vbMonday) <= 5 Then diasTrabajados = diasTrabajados - 1
                        
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    
    
    Else
        diasTrabajados = DateDiff("d", InicioProceso, FinProceso)
        Cad = " fecha between " & DBSet(InicioProceso, "F") & " AND " & DBSet(FinProceso, "F")
        Cad = Cad & " AND idcal=" & idTrabajador
        Cad = "Select * from calendariof WHERE " & Cad
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic
        
        Cad = ""
        While Not miRsAux.EOF
            If Weekday(miRsAux!Fecha, vbMonday) <= 5 Then diasTrabajados = diasTrabajados - 1
                        
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    End If
    idTrabajador = 0
    
    
    
    

    Cad = "Select * from tiposhora ORDER BY TipoHora"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic
    CuantosTiposHoraTrabaja = 0
    Cad = ""
    While Not miRsAux.EOF
        Cad = Cad & Mid(miRsAux!Desctipohora, 1, 3) & "|"
        CuantosTiposHoraTrabaja = CuantosTiposHoraTrabaja + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    For J = 1 To CuantosTiposHoraTrabaja
        Me.ListView1.ColumnHeaders.Add , , RecuperaValor(Cad, J), 800, 1
    Next
    
    
    'A�adiremos una columna para los dias trabajados. WIdth=0
    Me.ListView1.ColumnHeaders.Add , , "Dias", 0, 1
    
    'La columna para dias nomina
    Me.ListView1.ColumnHeaders.Add , , "Lab.", 800, 2
    
    
    Me.ListView1.Width = 6000 + (CuantosTiposHoraTrabaja * 800) + 900  'DE laboral
    Me.Width = Me.ListView1.Width + 120 + 240
    Me.cmdCancelar.Left = Me.Width - Me.cmdCancelar.Width - 240
    Me.cmdAceptar.Left = Me.cmdCancelar.Left - Me.cmdAceptar.Width - 240
    
    
    ColumnaDondeEmpiezanHoras = 3
    
    
    PintaColumnaDiasNominaAnterior = InicioProceso <> FechaInicioSemana
    
    
    Set miRs = New ADODB.Recordset
    
    
    Cad = "select  trabajadores.idtrabajador,nomtrabajador"
    Cad = Cad & " ,tmphorastipoalzira.*,DescTipoHora"
    Cad = Cad & " from trabajadores,tmphorastipoalzira,tiposhora Where trabajadores.idTrabajador"
    Cad = Cad & " = tmphorastipoalzira.idTrabajador And tiposhora.TipoHora = tmphorastipoalzira.tipohoras"
    'No separamos por cooperativa o fruxeresa. El desdeoble lo hacen luego
    'Cad = Cad & " and tmphorastipoalzira.ParaEmpresa =" & ParaLaCooperativa
    Cad = Cad & " and tmphorastipoalzira.codusu =" & vUsu.Codigo
    Cad = Cad & " order by trabajadores.idtrabajador,fecha,tipohoras"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ReDim Sumas(CuantosTiposHoraTrabaja)
    
    idTrabajador = -1
    While Not miRsAux.EOF
        If miRsAux!idTrabajador <> idTrabajador Then
            
            
            
            'InsertoSumatorio
            If idTrabajador >= 0 Then SumaHorasTrabajador idTrabajador, diasTrabajados
                
            
            
            Set IT = ListView1.ListItems.Add()
            IT.Text = miRsAux!idTrabajador
            IT.Tag = 0 'Trabajador
            IT.SubItems(1) = miRsAux!nomtrabajador
            IT.SubItems(2) = " "
            
            'El hco de horas
            For J = 1 To CuantosTiposHoraTrabaja
                IT.SubItems(ColumnaDondeEmpiezanHoras - 1 + J) = 0  'cargamos un CERO
                Sumas(J - 1) = 0 'reestablecemos sumas
            Next J
            IT.SubItems(ColumnaDondeEmpiezanHoras) = " "
            
            If PintaColumnaDiasNominaAnterior Then
            
                'IT.SubItems(ColumnaDondeEmpiezanHoras + CuantosTiposHoraTrabaja + 1) = "0"
                'IT.ListSubItems(ColumnaDondeEmpiezanHoras + CuantosTiposHoraTrabaja + 1).ForeColor = &H404040
                'IT.ListSubItems(ColumnaDondeEmpiezanHoras + CuantosTiposHoraTrabaja + 1).Bold = True
                
            End If
                
                idTrabajador = miRsAux!idTrabajador
                Fecha = "01/01/1900"
            
           
            '
            ' Y pongo este
            For J = 1 To CuantosTiposHoraTrabaja
                IT.SubItems(ColumnaDondeEmpiezanHoras + J) = " "  'cargamos un CERO
                Sumas(J) = 0
            Next J
            
            
            
            
            If ProcesoDeNominasAlzira Then
                'Si la fecha inicio del periodo es distinto de la fecha, pondre los datos de los dias YA procesados
                If FechaInicioSemana <> InicioProceso Then
                    
                    'Cad = "Select * from jornadassemanalesalz where ParaEmpresa = " & ParaLaCooperativa
                    'Noviembre 2014
                    Cad = "Select idtrabajador,fecha,tipohoras,sum(horastrabajadas) horastrabajadas,sum(laborable) labor from jornadassemanalesalz WHERE "
                    
                    Cad = Cad & " IdTrabajador =" & idTrabajador
                    
      
                    Cad = Cad & " AND fecha >=" & DBSet(FechaInicioSemana, "F") & " AND fecha <" & DBSet(InicioProceso, "F")
                    
                    'nov2014
                    Cad = Cad & " GROUP BY 1,2,3"
                    
                    Cad = Cad & " ORDER BY fecha,tipohoras"
                    miRs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    'Vemos los datos semanas anteriors
                    F2 = "01/01/1900"
                    While Not miRs.EOF
                        If F2 <> miRs!Fecha Then
                            Set IT = ListView1.ListItems.Add()
                            IT.Tag = 4 'Horas que ya estan procesadas
                            IT.Text = " "
                            IT.SubItems(1) = " "
                            IT.SubItems(2) = miRs!Fecha
                            IT.ListSubItems(2).ForeColor = &H8080&
                            For J = 1 To CuantosTiposHoraTrabaja
                                IT.SubItems(2 + J) = " "
                            Next J
                            'En J tengo ya la columna
                            J = CuantosTiposHoraTrabaja + CuantosTiposHoraTrabaja
                            'IT.ListSubItems(7).ForeColor = &H8080&
                            IT.SubItems(J) = miRs!labor
                            IT.ListSubItems(J).ForeColor = &H8080&
                            F2 = miRs!Fecha
                        End If
                        J = miRs!TipoHoras
                        IT.SubItems(ColumnaDondeEmpiezanHoras + J) = Format(miRs!HorasTrabajadas, "0.00")
                        Sumas(J) = Sumas(J) + miRs!HorasTrabajadas
                
                         
                         
                         
                         miRs.MoveNext
                    Wend
                    miRs.Close
                
                End If
            End If
        End If
        If miRsAux!Fecha <> Fecha Then
        
            Set IT = ListView1.ListItems.Add()
            IT.Tag = 1 'horas
            IT.Text = " "
            IT.SubItems(1) = " "
            IT.SubItems(2) = miRsAux!Fecha
            
            For J = 1 To CuantosTiposHoraTrabaja
                IT.SubItems(2 + J) = " "
            Next J
            Fecha = miRsAux!Fecha
        End If
        ColumnaDondeEmpiezanHoras = 3
        'Que columna pinto
        
        IT.SubItems(ColumnaDondeEmpiezanHoras + miRsAux!TipoHoras) = Format(miRsAux!HorasTrabajadas, "0.00")
        Sumas(miRsAux!TipoHoras) = Sumas(miRsAux!TipoHoras) + miRsAux!HorasTrabajadas
    
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    If idTrabajador >= 0 Then SumaHorasTrabajador idTrabajador, diasTrabajados
       
    Exit Sub
       
    Dim Aux As String
    Dim FechaAux As Date
    Dim I As Integer
    Dim DiasLaborablesInicioSemana As Integer
    'Febrero 2015
    'Dias nomina trabajador
    'Primero. Dias
    If idTrabajador >= 0 Then
        Cad = DevuelveDesdeBD("idcal", "trabajadores", "idTrabajador", CStr(idTrabajador))
        idTrabajador = Val(Cad) 'Celandario
                
        Cad = "fecha>=" & DBSet(FechaInicioSemana, "F") & " AND fecha<="
        'Ultimo dia de proceso
        I = Weekday(FinProceso, vbMonday)
        If I > vbFriday Then
            FechaAux = DateAdd("d", -(I - 1), CDate(FinProceso))
        Else
            FechaAux = FinProceso
        End If
            
        Cad = Cad & DBSet(FechaAux, "F") & " AND idcal"
        Cad = DevuelveDesdeBD("count(*)", "calendariof", Cad, CStr(idTrabajador))
        'ya tengo los festivos que hay en esa periodo de facturacion
        idTrabajador = Val(Cad) 'numero festivos del periodo
        
        I = DateDiff("d", FechaInicioSemana, FechaAux) + 1 'Dias del proceso
        I = I - idTrabajador 'dias proceso
        If I < 0 Then
            MsgBox "Diferencia de dias <0. " & idTrabajador & " "
            Stop 'que de error
            
        End If
        
    End If
    
    'Desde inicio semana hasta el dia antes del dia a procesar
    'FechaInicioSemana
    FechaAux = DateAdd("d", -1, InicioProceso)
    Cad = "select idtrabajador,sum(laborable) from jornadassemanalesalz where "
    Cad = Cad & " fecha>=" & DBSet(FechaInicioSemana, "F") & " AND fecha <=" & DBSet(FechaAux, "F")
    Cad = Cad & " GROUP BY idtrabajador"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If DBLet(miRsAux.Fields(1), "N") > 0 Then
            For I = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(I).Tag = 0 Then
                    If Val(ListView1.ListItems(I).Text) = miRsAux!idTrabajador Then
                        'Este es
                        
                        ListView1.ListItems(J).SubItems(ColumnaDondeEmpiezanHoras + CuantosTiposHoraTrabaja + 1) = miRsAux.Fields(1)
                        Exit For
                    End If
                End If
            Next
            If I > ListView1.ListItems.Count Then
              '  MsgBox "No se ha encotrado al trabajador: " & miRsAux!idTrabajador
              '  Stop
            End If
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
    Set miRs = Nothing
    Set miRsAux = Nothing
    
    
End Sub



'Calculo estandart. Coopic y Alzira
Private Sub SumaHorasTrabajador(idTrab As Long, Dias As Byte)
    If vEmpresa.QueEmpresa = 4 Then
        SumaHorasTrabajadorCatadau idTrab, Dias
    Else
        SumaHorasTrabajadorStd idTrab, Dias
    End If
End Sub


Private Sub SumaHorasTrabajadorStd(idTrab As Long, Dias As Byte)
Dim J As Integer
Dim IT
Dim Aux As Currency
Dim Ajustado As Boolean
Dim HorasSem As Integer

     'InsertoSumatorio
        Set IT = ListView1.ListItems.Add()
        IT.Text = " "
        IT.Tag = 2 'suma
        IT.SubItems(1) = " "
        IT.SubItems(2) = "SUMA "
        IT.ListSubItems(2).Bold = True
        IT.ListSubItems(2).ForeColor = vbGreen
        For J = 0 To CuantosTiposHoraTrabaja - 1
            IT.SubItems(ColumnaDondeEmpiezanHoras + J) = Format(Sumas(J), "0.00")
        Next J
        IT.SubItems(ColumnaDondeEmpiezanHoras + J) = Dias  'ultima columna
        
        'Ajuste hora
        HorasSem = Dias * 8
        If vEmpresa.CompensaHorasNominaMES Then HorasSem = 10000   'No hay limite de horas semanales
        
        Ajustado = False
        If Sumas(0) > HorasSem Then
            'No
            Aux = Sumas(0) - HorasSem
            Sumas(0) = HorasSem
            
            Sumas(1) = Sumas(1) + Aux
            Ajustado = True
        End If
        'Las estruturales y extraordinarias no pueden pasar de 80
        If Sumas(1) + Sumas(2) > 80 Then
            'MsgBox "Falta ajustar limite anual", vbExclamation
            'Ajustado = True
        End If
            
            
            
        If Ajustado Then
            Set IT = ListView1.ListItems.Add()
            IT.Text = " "
            IT.SubItems(1) = " "
            IT.SubItems(2) = "AJUSTE (" & idTrab & ")"
            IT.ListSubItems(2).Bold = True
            IT.ListSubItems(2).ForeColor = vbRed
            For J = 0 To CuantosTiposHoraTrabaja - 1
                IT.SubItems(ColumnaDondeEmpiezanHoras + J) = Format(Sumas(J), "0.00")
            Next J
            IT.SubItems(ColumnaDondeEmpiezanHoras + J) = Dias  'ultima columna
            IT.Tag = 3 'ajuste
            
        End If
        
End Sub

'Arrastra las estrucutirales
Private Sub SumaHorasTrabajadorCatadau(idTrab As Long, Dias As Byte)
Dim J As Integer
Dim IT
Dim Aux As Currency
Dim Ajustado As Boolean
Dim HorasSem As Integer
Dim Hest As Currency
Dim Fin As Boolean
Dim k As Integer
Dim EstrcuCompensadas As Currency
Dim B1 As Boolean

        'If idTrab = 57 Then Stop
        Ajustado = False
        
        
        If Sumas(1) > 0 Then
            'Tiene estructurales
            J = ListView1.ListItems.Count
            Fin = False
            Do
                If ListView1.ListItems(J).Tag <> 1 Then
                    Fin = True
                Else
                    J = J - 1
                End If
            Loop Until Fin
            
            For k = J + 1 To ListView1.ListItems.Count
                If Sumas(1) > 0 Then
                    
                    B1 = False
                    If Trim(ListView1.ListItems(k).SubItems(ColumnaDondeEmpiezanHoras)) <> "" Then
                        If ListView1.ListItems(k).SubItems(ColumnaDondeEmpiezanHoras) < 8 Then B1 = True
                    End If
                    If B1 Then
                        
                        Hest = 8 - ListView1.ListItems(k).SubItems(ColumnaDondeEmpiezanHoras)
                        If Hest > Sumas(1) Then Hest = Sumas(1)
                        EstrcuCompensadas = EstrcuCompensadas + Hest
                        ListView1.ListItems(k).SubItems(ColumnaDondeEmpiezanHoras) = ListView1.ListItems(k).SubItems(ColumnaDondeEmpiezanHoras) + Hest
                        Sumas(1) = Sumas(1) - Hest
                        Sumas(0) = Sumas(0) + Hest
                        ListView1.ListItems(k).SubItems(ColumnaDondeEmpiezanHoras) = ListView1.ListItems(k).SubItems(ColumnaDondeEmpiezanHoras)
                        ListView1.ListItems(k).ListSubItems(ColumnaDondeEmpiezanHoras).ForeColor = vbBlue
                        Ajustado = True
                    End If
                End If
            Next k
            
            'Si ha ajustado , entonces es que HA quitado estructurales
            If Ajustado Then
                
                For k = J + 1 To ListView1.ListItems.Count
                    If EstrcuCompensadas > 0 Then
                        If ListView1.ListItems(k).SubItems(ColumnaDondeEmpiezanHoras + 1) <> " " Then
                            
                            Hest = ListView1.ListItems(k).SubItems(ColumnaDondeEmpiezanHoras + 1) 'esrcuturales
                            If Hest > EstrcuCompensadas Then Hest = EstrcuCompensadas
                            ListView1.ListItems(k).SubItems(ColumnaDondeEmpiezanHoras + 1) = ListView1.ListItems(k).SubItems(ColumnaDondeEmpiezanHoras + 1) - Hest
                            EstrcuCompensadas = EstrcuCompensadas - Hest
                            ListView1.ListItems(k).ListSubItems(ColumnaDondeEmpiezanHoras + 1).ForeColor = vbBlue
                            
                        End If
                    End If
                Next k
            End If
        End If
        
        
        'InsertoSumatorio
        Set IT = ListView1.ListItems.Add()
        IT.Text = " "
        IT.Tag = 2 'suma
        IT.SubItems(1) = " "
        IT.SubItems(2) = "SUMA "
        IT.ListSubItems(2).Bold = True
        IT.ListSubItems(2).ForeColor = vbGreen
        For J = 0 To CuantosTiposHoraTrabaja - 1
            IT.SubItems(ColumnaDondeEmpiezanHoras + J) = Format(Sumas(J), "0.00")
        Next J
        IT.SubItems(ColumnaDondeEmpiezanHoras + J) = Dias  'ultima columna
        
        
        
        
        
        
        
        'Catadau. Si un
        HorasSem = Dias * 8
        
        
        
        
        
        If Sumas(0) > HorasSem Then
            'No
            Aux = Sumas(0) - HorasSem
            Sumas(0) = HorasSem
            
            Sumas(1) = Sumas(1) + Aux
            Ajustado = True
        End If
            
                    
            
        If Ajustado Then
            Set IT = ListView1.ListItems.Add()
            IT.Text = " "
            IT.SubItems(1) = " "
            IT.SubItems(2) = "AJUSTE (" & idTrab & ")"
            IT.ListSubItems(2).Bold = True
            IT.ListSubItems(2).ForeColor = vbRed
            For J = 0 To CuantosTiposHoraTrabaja - 1
                IT.SubItems(ColumnaDondeEmpiezanHoras + J) = Format(Sumas(J), "0.00")
            Next J
            IT.SubItems(ColumnaDondeEmpiezanHoras + J) = Dias  'ultima columna
            IT.Tag = 3 'ajuste
            
        End If
        
End Sub




Private Sub HacerGeneracionPeriodo()
Dim Cad As String
Dim C As String
Dim idTrabajador As Long
Dim Columnas As Byte
Dim I As Byte
Dim Horas As Currency
Dim Laborable As Byte
Dim HaSidoAjustada As Boolean

Dim TieneZonas As Boolean
Dim CuadraHorasZonas As Boolean
Dim AuxH As String
Dim CadenaZonasAreas As String
Dim codArea As Integer
Dim HZ1 As Currency
Dim HZ2 As Currency
Dim HorasArea As Currency
Dim KK As Integer
Dim Kuantos As Integer
Dim FinB As Boolean


Dim vecHorasZona() As Currency
Dim vecZona() As Integer
Dim PunteroZona As Integer
    'Encadena desde otro form llevare las fechas del intervalor
    Cad = "select paraempresa,count(*) as cuantos from tmphorastipoalzira where codusu =" & vUsu.Codigo & " group by 1"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    If miRsAux.EOF Then
        Cad = Cad & "0�0|"
    Else
    
        While Not miRsAux.EOF
            Cad = Cad & miRsAux!paraempresa & "�" & miRsAux!Cuantos & "|"
            miRsAux.MoveNext
        Wend
    End If
    miRsAux.Close
    
    'Si tiene ZONAS
    TieneZonas = False
    Cad = "(select distinct area from terminales) as mister "
    Cad = DevuelveDesdeBD("count(*)", Cad, "1", "1")
    If Val(Cad) > 1 Then TieneZonas = True
        
    
    
    
    'jornadassemanalesproceso fecha,fechaIni,fechaFin,Sumatorios,codusu,Nombre
    If Me.TodasSecciones Then
        Cad = "SELECT now()," & DBSet(RecuperaValor(CadenaDesdeOtroForm, 1), "F") & "," & DBSet(RecuperaValor(CadenaDesdeOtroForm, 2), "F") & ",count(*)"
        Cad = Cad & "," & vUsu.Codigo & "," & DBSet(vUsu.Nombre, "T") & ", seccion"
        Cad = Cad & " from trabajadores,tmphorastipoalzira,tiposhora Where"
        Cad = Cad & " trabajadores.idTrabajador = tmphorastipoalzira.idTrabajador And"
        Cad = Cad & " tiposhora.TipoHora = tmphorastipoalzira.tipohoras And tmphorastipoalzira.codusu = " & vUsu.Codigo
        Cad = Cad & " group by seccion"
        
        
    
    Else
        Cad = "now()," & DBSet(RecuperaValor(CadenaDesdeOtroForm, 1), "F") & "," & DBSet(RecuperaValor(CadenaDesdeOtroForm, 2), "F") & "," & DBSet(Cad, "T")
        Cad = Cad & "," & vUsu.Codigo & "," & DBSet(vUsu.Nombre, "T") & "," & IdSeccion
        Cad = " VALUES (" & Cad & ")"
    End If
    
    
    Cad = "INSERT INTO jornadassemanalesproceso(fecha,fechaIni,fechaFin,Sumatorios,codusu,Nombre,Seccion) " & Cad
    conn.Execute Cad
    

    'Si va por todas las secciones, las que no haya procesado, porque no tienen datos entre las fechas, la metemos en el proceso
    '    con valor - 1
    
    If TodasSecciones Then
        espera 0.5
        Cad = "INSERT IGNORE INTO jornadassemanalesproceso(fecha,fechaIni,fechaFin,Sumatorios,codusu,Nombre,Seccion) "
        Cad = Cad & " SELECT now()," & DBSet(RecuperaValor(CadenaDesdeOtroForm, 1), "F") & "," & DBSet(RecuperaValor(CadenaDesdeOtroForm, 2), "F") & ",'-1'"
        Cad = Cad & "," & vUsu.Codigo & "," & DBSet(vUsu.Nombre, "T") & ", idseccion"
        Cad = Cad & " from secciones where Nominas =1 and not idseccion in (select seccion from jornadassemanalesproceso where fechaini = " & DBSet(RecuperaValor(CadenaDesdeOtroForm, 1), "F") & ")"
        conn.Execute Cad
    End If

    
    If vEmpresa.QueEmpresa = vbAlzira Then
        'COOPIC tiene proceso final bolsa horas
        Cad = "select now(), IdTrabajador,ParaEmpresa,TipoHora,HorasBolsa "
        Cad = Cad & " from trabajadoresbolsahoras"
        Cad = "insert into jornadassemanalesHcoBolsa(fecha,IdTrabajador,ParaEmpresa,TipoHora,HorasBolsa) " & Cad
        conn.Execute Cad
    End If
    
    'Insertamos en la que tendra que dia , que horas
    'Recorremos el LISTVIEW y haremos un insert con cada hora, bien sea nor estuct...
    
    C = "INSERT INTO jornadassemanalesalz(idtrabajador,fecha,TipoHoras,horastrabajadas,ParaEmpresa,Ajuste,laborable,codarea) VALUES"
    
    idTrabajador = 0
    Cad = ""
    If ListView1.ListItems.Count > 0 Then Columnas = Me.ListView1.ListItems(1).ListSubItems.Count
    Set miRsAux = New ADODB.Recordset
    For J = 1 To Me.ListView1.ListItems.Count
        If ListView1.ListItems(J).Tag = 0 Then
            idTrabajador = CLng(ListView1.ListItems(J).Text)
            
            'If idTrabajador = 1743 Then St op
            
            If TieneZonas Then
                If J > 1 Then miRsAux.Close  '1 es el primer trasbajador
                AuxH = "select * from tmphorasArea where codusu =" & vUsu.Codigo & " AND   idtra=" & idTrabajador & " order by fecha,masdenArea"
                miRsAux.Open AuxH, conn, adOpenKeyset, adCmdText
                If Not miRsAux.EOF Then CuadraHorasZonas = True
                    
            End If
            
            
        Else
            If ListView1.ListItems(J).Tag = 1 Then
                            
                
                'Id trabajador
                Laborable = Val((ListView1.ListItems(J).SubItems(ColumnaDondeEmpiezanHoras + CuantosTiposHoraTrabaja + 1)))
                
                Kuantos = IIf(vEmpresa.QueEmpresa = vbAlzira, 1, 0)
                CadenaZonasAreas = Format(Kuantos, "0000") & "-" & TransformaComasPuntos(CStr(HZ1)) & "|"
                Kuantos = 1
                If TieneZonas Then
                            
                    AuxH = "fecha =" & DBSet(Me.ListView1.ListItems(J).SubItems(ColumnaDondeEmpiezanHoras - 1), "F")
                    miRsAux.Find AuxH, , adSearchForward, 1
                    If miRsAux.EOF Then
                        'Ya esta arraiba
                        
                        ' CadenaZonasAreas      Kuantos
                        CuadraHorasZonas = False
                    Else
                        If DBLet(miRsAux!masdenarea, "N") < 1 Then
                            'una unica area de trabajo
                            codArea = DBLet(miRsAux!Area)
                            CuadraHorasZonas = False
                        Else
                            'Empieza la fiesta
                            
                            AuxH = ""
                            Kuantos = 0
                            FinB = False
                            While Not FinB
                                If miRsAux!Fecha = Me.ListView1.ListItems(J).SubItems(ColumnaDondeEmpiezanHoras - 1) Then
                                
                                    AuxH = AuxH & Format(miRsAux!Area, "0000") & "-" & miRsAux!Horas & "|"
                                    Kuantos = Kuantos + 1
                                    miRsAux.MoveNext
                                    If miRsAux.EOF Then FinB = True
                                Else
                                    FinB = True
                                End If
                                
                            Wend
                            If Kuantos = 1 Then
                               
                                MsgBox "No deberia haber llegado. Soporte."
                                 Stop
                                'NO DEBERIA HABER LLGADO
                            Else
                                ReDim vecZona(Kuantos - 1)
                                ReDim vecHorasZona(Kuantos - 1)
                                For I = 0 To Kuantos - 1
                                    CadenaZonasAreas = RecuperaValor(AuxH, I + 1)
                                    vecZona(I) = Mid(CadenaZonasAreas, 1, 4)
                                    vecHorasZona(I) = Mid(CadenaZonasAreas, 6)
                                Next
                                PunteroZona = 0
                            End If
                        End If
                    End If
                
                End If
                
                For I = ColumnaDondeEmpiezanHoras To Columnas
                        
                    
                    If Trim(ListView1.ListItems(J).SubItems(I)) <> "" Then
                            'Ha sido ajustada
                            HaSidoAjustada = False
                            If ListView1.ListItems(J).ListSubItems(ColumnaDondeEmpiezanHoras - 1).ForeColor = vbBlue Then
                                HaSidoAjustada = True
                            Else
                                If ListView1.ListItems(J).ListSubItems(ColumnaDondeEmpiezanHoras).ForeColor = vbBlue Then
                                    HaSidoAjustada = True
                                Else
                                    If ListView1.ListItems(J).ListSubItems(ColumnaDondeEmpiezanHoras + 1).ForeColor = vbBlue Then HaSidoAjustada = True
                                End If
                            End If
                            HZ1 = ImporteFormateado(Me.ListView1.ListItems(J).SubItems(I))  'Horas
                            
                            Do
                                
                                If CuadraHorasZonas Then
                                    If PunteroZona <= UBound(vecZona) Then
                                        
                                        If HZ1 >= vecHorasZona(PunteroZona) Then
                                            'Son mas horas de las que tiene en esa zona.
                                            ' Por lo tanto, gasta las que necesite en esa zona y repetira para la siguiente
                                            codArea = vecZona(PunteroZona)
                                            HZ2 = vecHorasZona(PunteroZona)
                                            HZ1 = HZ1 - HZ2
                                            PunteroZona = PunteroZona + 1
                                            
                                        Else
                                            codArea = vecZona(PunteroZona)
                                            HZ2 = HZ1
                                            HZ1 = 0
                                            vecHorasZona(PunteroZona) = vecHorasZona(PunteroZona) - HZ2
                                            If vecHorasZona(PunteroZona) < 0 Then MsgBox "Error puntero ZONA . Soporte "
                                        End If
                                    
                                    Else
                                        'Ya no hay mas zonas. Se queda en la que este
                                        'Stop
                                    End If
                                Else
                                    HZ2 = HZ1
                                    HZ1 = 0
                                End If
                                
                                
                                
                            If HZ2 = 0 Then
                                
                                Debug.Print idTrabajador & " - " & DBSet(Me.ListView1.ListItems(J).SubItems(ColumnaDondeEmpiezanHoras - 1), "F")
                                'No hacemos insert
                            Else
                                Cad = Cad & ", (" & idTrabajador & "," & DBSet(Me.ListView1.ListItems(J).SubItems(ColumnaDondeEmpiezanHoras - 1), "F") & "," & I - ColumnaDondeEmpiezanHoras & ","
                                Cad = Cad & DBSet(HZ2, "N") & ",0,"  'antes ponia: TransformaComasPuntos(Me.ListView1.ListItems(J).SubItems(I))
                            
                                Cad = Cad & Abs(HaSidoAjustada)
                                'Dias nomina
                                If Laborable > 0 Then
                                    Cad = Cad & ",1"
                                    Laborable = 0
                                Else
                                    Cad = Cad & ",0"
                                End If
                        
                                'Nov 2020   Area
                                Cad = Cad & "," & codArea
                                'Fin
                                Cad = Cad & ")"
                            End If
                            Loop Until HZ1 = 0
                    End If
                Next
                'Febr 2016. Dias nomina
                If Len(Cad) > 20000 Then
                    Cad = Mid(Cad, 2)
                    Cad = C & Cad
                    conn.Execute Cad
                    Cad = ""
                End If
                
            End If
        End If
    Next J
    Set miRsAux = Nothing
    If Cad <> "" Then
        Cad = Mid(Cad, 2)
        Cad = C & Cad
        conn.Execute Cad
    End If
    
    
    
    
    'Mete los ajustes semanales y recalcula bolsas
    If vEmpresa.QueEmpresa = 2 Then HacerAjustesSobreBD
    

End Sub





Private Sub AjustarHoras()
Dim Difer As Currency
Dim PrimerDiaTrabajador As Integer
Dim PrimerDiaParaAjustar As Integer
Dim I As Integer
Dim Llevo As Currency
Dim Horas As Currency
Dim HorasSemama As Integer
Dim HemosAjustado As Boolean
    
        
        
        'Como vamos a ajustar las horas.
        'Para todos aquellos que haya que ajustar
        HemosAjustado = False
        For J = Me.ListView1.ListItems.Count To 1 Step -1
            If ListView1.ListItems(J).Tag = 3 Then
            
                ListView1.ListItems(J).EnsureVisible
            
                'OK, este es el de ajustar
                
                'Difer = ImporteFormateado(ListView1.ListItems(J - 1).SubItems(ColumnaDondeEmpiezanHoras))
                'Difer = Difer - ImporteFormateado(ListView1.ListItems(J).SubItems(ColumnaDondeEmpiezanHoras))
                
                
                HorasSemama = ListView1.ListItems(J).SubItems(ColumnaDondeEmpiezanHoras + CuantosTiposHoraTrabaja) * 8
                If HorasSemama = 0 Then HorasSemama = 8
                'Puede haber procesado algun dia de esta semana. Por lo tanto, esos dias NO se pueden tocar
                'Buscare el oprimer dia a procesar.
                PrimerDiaTrabajador = 0
                PrimerDiaParaAjustar = 0
                I = J - 1
                While PrimerDiaTrabajador = 0
                    If Me.ListView1.ListItems(I).Tag = 0 Then   'columna nombre
                        'Primer dia a trabajar
                        PrimerDiaTrabajador = I + 1
                        If PrimerDiaParaAjustar = 0 Then PrimerDiaParaAjustar = PrimerDiaTrabajador
                    Else
                        'Si ya esta ajustado no podre tocarlo
                        If Me.ListView1.ListItems(I).Tag = 4 Then
                            If PrimerDiaParaAjustar = 0 Then PrimerDiaParaAjustar = I + 1
                        End If
                        I = I - 1
                        
                    End If
                Wend
                         
                If PrimerDiaTrabajador <> PrimerDiaParaAjustar Then
                    'Porceso ya parte de semana
                    Llevo = 0
                    For I = PrimerDiaTrabajador To PrimerDiaParaAjustar - 1 'horas YA Procesadas
                        Horas = ImporteFormateado(Trim(ListView1.ListItems(I).SubItems(ColumnaDondeEmpiezanHoras)))
                        Llevo = Llevo + Horas
                    Next
                    If Llevo > HorasSemama Then
                        MsgBox "Horas ya procesadas superan las horas semanales", vbExclamation
                        
                    End If
                    PrimerDiaTrabajador = PrimerDiaParaAjustar
                    HemosAjustado = True
                Else
                    'Iniciamos de CERO el proceso
                    Llevo = 0
                End If
                
               
                For I = PrimerDiaTrabajador To J - 2 'Los dias
                
                    Horas = ImporteFormateado(Trim(ListView1.ListItems(I).SubItems(ColumnaDondeEmpiezanHoras)))
                    If Horas + Llevo > HorasSemama Then
                        'Ya las pasa. Son todas ESTRUCTURALES excepto si son del sabado
                        Difer = (Horas + Llevo) - HorasSemama
                        
                        'HT tiene una DIFER menos
                        ListView1.ListItems(I).SubItems(ColumnaDondeEmpiezanHoras) = Format(Horas - Difer, "0.00")
                        
                        
                        'HEstructurales tiene una SI NO ES SABADO
                        Horas = ImporteFormateado(Trim(ListView1.ListItems(I).SubItems(ColumnaDondeEmpiezanHoras + 1)))
                        Horas = Horas + Difer
                        ListView1.ListItems(I).SubItems(ColumnaDondeEmpiezanHoras + 1) = Format(Horas, "0.00")
                        ListView1.ListItems(I).ListSubItems(ColumnaDondeEmpiezanHoras + 1).ForeColor = vbBlue
                        ListView1.ListItems(I).ListSubItems(ColumnaDondeEmpiezanHoras - 1).ForeColor = vbBlue
                        
                        Llevo = HorasSemama 'Ya tiene las semanales cumplidas
                    Else
                        Llevo = Llevo + Horas
                    End If
                Next I
                
                
                
            End If
        Next J
        
        'If J = 0 Then
        If HemosAjustado Then
            'NO hay ninguno para ajustar
            Exit Sub
        End If
        
        Dim Fin As Boolean
        Dim Cad As String
        Dim DiasSemana As Integer
        Dim DiasTr As Integer
        Dim FechaAux As Date
        
        Cad = DevuelveDesdeBD("min(idtrabajador)", "tmphorastipoalzira", "codusu", CStr(vUsu.Codigo))
        If Cad = "" Then
            Cad = "1"
        Else
            Cad = DevuelveDesdeBD("idcal", "trabajadores", "idtrabajador", Cad)
            If Cad = "" Then Cad = "1"
        End If
        J = Weekday(FinProceso, vbMonday)
        If J > 5 Then
            J = J - 5
            FechaAux = DateAdd("d", -J, FinProceso)
            DiasSemana = 5
        Else
            FechaAux = FinProceso
            DiasSemana = DateDiff("d", FechaInicioSemana, FinProceso) + 1
        End If
        
        Cad = " idcal =" & Cad
        Cad = Cad & " AND fecha between " & DBSet(FechaInicioSemana, "F") & " AND " & DBSet(FinProceso, "F") & " AND 1"
        Cad = DevuelveDesdeBD("count(*)", "calendariof", Cad, "1")
        
        DiasSemana = DiasSemana - Val(Cad)
        
        
        
        
        'Ajustes dias nomina
        Fin = False
        
        J = 1
        While J < ListView1.ListItems.Count
        
            If ListView1.ListItems(J).Tag = 0 Then
                'Estamos dentro del trabajador
                
                'Dias trabajados anterior al dias proceso
                DiasTr = 0
                Fin = False
                Do
                    J = J + 1
                   
                    If ListView1.ListItems(J).Tag = 2 Then
                        'Final de lineas del trabajador
                        ListView1.ListItems(J).SubItems(ColumnaDondeEmpiezanHoras + CuantosTiposHoraTrabaja + 1) = DiasTr
                        Fin = True
                        J = J + 1
                        'Significa que lleva la linea de ajuste. Hay que sumar uno mas a J
                        If J < ListView1.ListItems.Count Then
                            If ListView1.ListItems(J).Tag = 3 Then J = J + 1
                        End If
                        
                    Else
                        'Si ha trabajado horas Normales y o estruturales
                        If ListView1.ListItems(J).Tag = 1 Then
                            Horas = ImporteFormateado(Trim(ListView1.ListItems(J).SubItems(ColumnaDondeEmpiezanHoras)))
                            Horas = Horas + ImporteFormateado(Trim(ListView1.ListItems(J).SubItems(ColumnaDondeEmpiezanHoras + 1)))
                            
                            If Horas <> 0 Then
                                'Tiene horas este dia, que son laborales
                                'Veremos si ha trabado todos los dias que podia
                                If DiasSemana > DiasTr Then
                                    DiasTr = DiasTr + 1
                                    ListView1.ListItems(J).SubItems(ColumnaDondeEmpiezanHoras + CuantosTiposHoraTrabaja + 1) = 1 'Dia nomina
                                End If
                            Else
                                'Son extras. NO suman
                                
                            End If
                        Else
                            If ListView1.ListItems(J).Tag = 4 Then
                                
                                'Dias trabajados con anterioridad
                                DiasTr = DiasTr + Val(ListView1.ListItems(J).SubItems(ColumnaDondeEmpiezanHoras + CuantosTiposHoraTrabaja + 1))  'Dia nomina
                            Else
                               ' Stop
                            End If
                        End If
                    End If
                Loop Until Fin
            End If
            
        
        
        
        Wend
        
        
        
        
        
        ListView1.Refresh
End Sub

Private Sub HacerAjustesSobreBD()
Dim Aux As String
Dim C As String
Dim Byt As Byte
Dim Difer As Currency
Dim IdTr As Long
Dim RBolsa As ADODB.Recordset
Dim HAnterior As Currency

    Set RBolsa = New ADODB.Recordset
    
    For J = 1 To Me.ListView1.ListItems.Count
        
        If ListView1.ListItems(J).Tag = 3 Then
            'OK. Ha habido ajuste
            
            'Lo que esta entre paraentesis es el trabajador
            
            
            'Trabajador
            C = Mid(ListView1.ListItems(J).SubItems(2), InStr(1, ListView1.ListItems(J).SubItems(2), "(") + 1)
            C = Mid(C, 1, Len(C) - 1)
            IdTr = Val(C)
            
            
            'Leo bolsa horas
            C = "select idtrabajador,sum(if(tipohora=1,HorasBolsa,0)) estruct,"
            C = C & "sum(if(tipohora=2,HorasBolsa,0)) extra, sum(if(tipohora=3,HorasBolsa,0)) pactad from "
            C = C & " trabajadoresbolsahoras where idtrabajador=" & IdTr & " and paraempresa=0 group by 1"
            RBolsa.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            
           'Para la bolsa de horas
            conn.Execute "DELETE FROM trabajadoresbolsahoras WHERE idtrabajador=" & IdTr & " AND ParaEmpresa =0"
            Aux = ""
           
            'El 0 son horas trabajadas
            For Byt = 1 To CuantosTiposHoraTrabaja - 1
                
                HAnterior = 0
                If Not RBolsa.EOF Then HAnterior = DBLet(RBolsa.Fields(CInt(Byt)), "N")
                
                
                Difer = ImporteFormateado(ListView1.ListItems(J).SubItems(ColumnaDondeEmpiezanHoras + Byt))
               
                '        IdTrabajador,ParaEmpresa,TipoHora,HorasBolsa"
                Difer = Difer + HAnterior
                If Difer <> 0 Then Aux = Aux & ", (" & IdTr & ",0," & Byt & "," & DBSet(Difer, "N") & ")"
                
            Next Byt
            If Aux <> "" Then
                C = "INSERT INTO trabajadoresbolsahoras(IdTrabajador,ParaEmpresa,TipoHora,HorasBolsa) VALUES "
                Aux = Mid(Aux, 2)
                Aux = C & Aux
                conn.Execute Aux
            End If
            RBolsa.Close
        End If
    Next
    '
    Set RBolsa = Nothing
End Sub



    



Private Function CargarDatosImpresion() As Boolean
Dim Aux As String
Dim IdTr As Long
Dim Cad As String
Dim TieneAjustes As Boolean
Dim Byt As Byte
Dim Impor As Currency
Dim DiasLaborables As Integer

    'tmpcombinada(IdTrabajador,Fecha,idinci,HT,HE,HR)
    NumRegElim = 1
    conn.Execute "DELETE FROM  tmpcombinada WHERE codusu = " & vUsu.Codigo
    
    IdTr = 0
    Cad = ""
    
    For J = 1 To Me.ListView1.ListItems.Count
        
           If ListView1.ListItems(J).Tag = 0 Then
                IdTr = Val(ListView1.ListItems(J).Text)
                
                DiasLaborables = 0
                For Byt = 0 To CuantosTiposHoraTrabaja - 1
                    Impor = ImporteFormateado(Trim(ListView1.ListItems(J).SubItems(ColumnaDondeEmpiezanHoras + Byt)))
                    Sumas(Byt) = Impor
                Next Byt
                NumRegElim = 0
            Else
                If ListView1.ListItems(J).Tag = 1 Then
                    NumRegElim = NumRegElim + 1
                    
                      For Byt = 0 To CuantosTiposHoraTrabaja - 2 'NO vamos a ver PACTADAS todavia
                            Impor = ImporteFormateado(Trim(ListView1.ListItems(J).SubItems(ColumnaDondeEmpiezanHoras + Byt)))
                            Sumas(Byt) = Sumas(Byt) + Impor
                      Next Byt
                      If ListView1.ListItems(J).SubItems(CuantosTiposHoraTrabaja + 4) <> "" Then DiasLaborables = DiasLaborables + ListView1.ListItems(J).SubItems(CuantosTiposHoraTrabaja + 4)
                Else
                    If ListView1.ListItems(J).Tag = 2 Then
                                            
'                        TieneAjustes = False
'                        If J = ListView1.ListItems.Count Then
'                            'Es el ultimo. NO tiene ajuste
'                        Else
'                            If ListView1.ListItems(J + 1).Tag = 3 Then TieneAjustes = True
'                        End If
'
'                        'El 0 son horas trabajadas
'                        Aux = ""
'                        For Byt = 0 To CuantosTiposHoraTrabaja - 2 'NO vamos a ver PACTADAS todavia
'                            If TieneAjustes Then
'                                Impor = ImporteFormateado(Trim(ListView1.ListItems(J + 1).SubItems(ColumnaDondeEmpiezanHoras + Byt)))
'                            Else
'                                Impor = ImporteFormateado(Trim(ListView1.ListItems(J).SubItems(ColumnaDondeEmpiezanHoras + Byt)))
'                            End If
'                            Impor = Impor - Sumas(Byt)
'                            '
'                            ''tmpcombinada(IdTrabajador,Fecha,idinci,HT,HE,HR)
'                            Aux = Aux & "," & DBSet(Impor, "N", "N")
'
'                        Next Byt
'                        Cad = Cad & ", (" & vUsu.Codigo & "," & IdTr & ",'1972-04-12'," & NumRegElim & Aux & ")"
                        Aux = ""
                        For Byt = 0 To CuantosTiposHoraTrabaja - 2 'NO vamos a ver PACTADAS todavia
                            Aux = Aux & "," & DBSet(Sumas(Byt), "N")
                        Next
                        'Dias para la laborable 'Marzo 2016
                        
                        Aux = Aux & "," & DiasLaborables
                        Cad = Cad & ", (" & vUsu.Codigo & "," & IdTr & ",'1972-04-12'," & NumRegElim & Aux & ")"
                    End If
                End If
                
            
        End If
    Next
    CargarDatosImpresion = True
    Cad = Mid(Cad, 2)
    Cad = "INSERT INTO tmpcombinada(codusu,IdTrabajador,Fecha,idinci,HT,HE,HR,H1) VALUES " & Cad
    conn.Execute Cad
    

End Function

Private Sub ListView1_DblClick()
Dim QueTrabajador  As Integer
Dim Cad As String


    If ListView1.SelectedItem Is Nothing Then Exit Sub
    Cad = ""
    J = ListView1.SelectedItem.Index
    While J > 0
        If Me.ListView1.ListItems(J).Tag <> 0 Then   'columna nombre
            'uno patras
            J = J - 1
           
        Else
            'Este es el trabajador
             Cad = ListView1.ListItems(J).Text & "|" & ListView1.ListItems(J).SubItems(1) & "|" & InicioProceso & "|" & FinProceso & "|"
             J = 0
        End If
       
    Wend
    If Cad <> "" Then
        frmMarcajesPantalla.QuieroVerDatos = Cad
        frmMarcajesPantalla.Show vbModal
    End If
End Sub




