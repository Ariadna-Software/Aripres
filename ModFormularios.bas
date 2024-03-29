Attribute VB_Name = "ModFormularios"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
'   FUNCIONES GENERALES
'----------------------------------------------------------------------------
'----------------------------------------------------------------------------


'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla
Public Function SugerirCodigoSiguienteStr(NomTabla As String, NomCodigo As String, Optional CondLineas As String) As String
Dim SQL As String
Dim RS As ADODB.Recordset

    On Error GoTo ESugerirCodigo

    'SQL = "Select Max(codtipar) from stipar"
    SQL = "Select Max(" & NomCodigo & ") from " & NomTabla
    If CondLineas <> "" Then
        SQL = SQL & " WHERE " & CondLineas
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, , , adCmdText
    SQL = "1"
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            If IsNumeric(RS.Fields(0)) Then
                SQL = CStr(RS.Fields(0) + 1)
            Else
                If Asc(Left(RS.Fields(0), 1)) <> 122 Then 'Z
                SQL = Left(RS.Fields(0), 1) & CStr(Asc(Right(RS.Fields(0), 1)) + 1)
                End If
            End If
        End If
    End If
    RS.Close
    Set RS = Nothing
    SugerirCodigoSiguienteStr = SQL
ESugerirCodigo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Public Sub BloquearFrameAux(ByRef formulario As Form, nom_frame As String, Modo As Byte, Optional NumTabMto As Integer)
Dim I As Byte
Dim B As Boolean
Dim Control As Object

    On Error GoTo EBloquear

    'b = (Modo = 3 Or Modo = 4 Or Modo = 5)
    B = (Modo = 5) 'And (NumTabMto = 3)
    
    For Each Control In formulario.Controls
        'If (Control.Tag <> "") And (Control.Visible = True) And (Control.Container.Name = nom_frame) Then
        If (Control.Tag <> "") Then
           If (Control.Container.Name = nom_frame) Then
                If (TypeOf Control Is TextBox) And (Control.Name = "txtAux") Then
                    Control.Locked = Not B
                    If B Then
                        Control.BackColor = vbWhite
                    Else
                        Control.BackColor = &H80000018 'Amarillo Claro
                    End If
                    If Modo = 3 Then Control.Text = "" 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
                    
                ElseIf (TypeOf Control Is ComboBox) And (Control.Name = "cmbAux") Then
                    'Control.Locked = Not b
                    Control.Enabled = B
                    If B Then
                        Control.BackColor = vbWhite
                    Else
                        Control.BackColor = &H80000018 'Amarillo Claro
                    End If
                    If Modo = 3 Then Control.ListIndex = 0 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
                End If
            End If
        End If
    
    Next Control

EBloquear:
    If Err.Number <> 0 Then Err.Clear
End Sub





Public Sub BloquearFrameAux2(ByRef formulario As Form, nom_frame As String, bloquea As Boolean)
Dim B As Boolean
Dim Control As Object

    On Error GoTo EBloquear

    'b = (Modo = 3 Or Modo = 4 Or Modo = 5)
'    b = (Modo = 5) And (NumTabMto = 3)
    B = bloquea
    
    For Each Control In formulario.Controls
        If (TypeOf Control Is TextBox) Then 'TEXT
            If (Control.Name = "txtAux") And (Control.Container.Name = nom_frame) Then
                If (Control.Tag <> "") Then
                    Control.Locked = B
                    If Not B Then
                        Control.BackColor = vbWhite
                    Else
                        Control.BackColor = &H80000018 'Amarillo Claro
                    End If
'                    If Modo = 3 Then Control.Text = "" 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
                End If
            End If
            
        ElseIf (TypeOf Control Is ComboBox) Then 'COMBO
            If (Control.Name = "cmbAux") And (Control.Container.Name = nom_frame) Then
                Control.Enabled = Not B
                If Not B Then
                    Control.BackColor = vbWhite
                Else
                    Control.BackColor = &H80000018 'Amarillo Claro
                End If
'                If Modo = 3 Then Control.ListIndex = 0 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
            End If
        End If
    Next Control

EBloquear:
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub BloquearText1(ByRef formulario As Form, Modo As Byte)
'Bloquea controles q se llamen TEXT1 si no estamos en Modo: 3.-Insertar, 4.-Modificar
'IN ->  formulario: formulario en el que se van a poner los controles textbox en modo visualizaci�n
'       Modo: modo del mantenimiento (Insertar, Modificar,Buscar...)
Dim I As Byte
Dim B As Boolean
Dim vtag As CTag
On Error Resume Next

    With formulario
        'b = (Modo = 3 Or Modo = 4 Or Modo = 1 Or Modo = 5) 'And ModoLineas = 1))
        B = (Modo = 3 Or Modo = 4 Or Modo = 1) '06/09/2005, lleve el modo 5 per a que no es puga modificar la cap�alera mentre treballe en les ll�nies
        
        For I = 0 To .Text1.Count - 1 'En principio todos los TExt1 tiene TAG
            Set vtag = New CTag
            vtag.Cargar .Text1(I)
            If vtag.Cargado Then
                If vtag.EsClave And (Modo = 4 Or Modo = 5) Then
                    .Text1(I).Locked = True
                    .Text1(I).BackColor = &H80000018 'groc
                Else
                     .Text1(I).Locked = Not B  '((Not b) And (Modo <> 1))
                    If B Then
                        .Text1(I).BackColor = vbWhite
                    Else
                        .Text1(I).BackColor = &H80000018 'groc
                    End If
                    If Modo = 3 Then .Text1(I).Text = "" 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
                End If
'            Else
'                .text1(i).Locked = Not b  '((Not b) And (Modo <> 1))
'                If b Then
'                    .text1(i).BackColor = vbWhite
'                Else
'                    .text1(i).BackColor = &H80000018 'groc
'                End If
'                If Modo = 3 Then .text1(i).Text = "" 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
            End If
            Set vtag = Nothing
        Next I
    End With
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BloquearTxt(ByRef Text As TextBox, B As Boolean, Optional EsContador As Boolean)
'Bloquea un control de tipo TextBox
'Si lo bloquea lo poner de color amarillo claro sino lo pone en color blanco (sino es contador)
'pero si es contador lo pone color azul claro
On Error Resume Next

    Text.Locked = B
    If Not B And Text.Enabled = False Then Text.Enabled = True
    If B Then
        If EsContador Then
            'Si Es un campo que se obtiene de un contador poner color azul
            Text.BackColor = &H80000013 'Azul Claro
        Else
            Text.BackColor = &H80000018 'Amarillo Claro
        End If
    Else
        Text.BackColor = vbWhite
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BloquearCmb(ByRef Cmb As ComboBox, B As Boolean, Optional EsContador As Boolean)
'Bloqueja un control de tipo ComboBox
'Si el bloqueja el posa de color gris claro, sino el posa de color blanc (sino es contador)
'pero si es contador el posa color blau clar
On Error Resume Next

    'Cmb.Locked = b
    Cmb.Enabled = Not B
    'If Not b And Cmb.Enabled = False Then Cmb.Enabled = True
    If B Then
        If EsContador Then
            'Si Es un campo que se obtiene de un contador poner color azul
            Cmb.BackColor = &H80000013 'Azul Claro
        Else
            Cmb.BackColor = &H80000018 'Amarillo Claro
        End If
    Else
        Cmb.BackColor = vbWhite
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BloquearCheck1(ByRef formulario As Form, Modo As Byte)
'Bloquea controles q sean CheckBox si no estamos en Modo: 3.-Insertar, 4.-Modificar
'IN ->  formulario: formulario en el que se van a poner los controles textbox en modo visualizaci�n
'       Modo: modo del mantenimiento (Insertar, Modificar,Buscar...)
    Dim B As Boolean
'    Dim Control As Control

    On Error Resume Next

    B = (Modo = 3 Or Modo = 4 Or Modo = 1)
    With formulario
        For I = 0 To .Check1.Count - 1
            .Check1(I).Enabled = B
            If Modo = 3 Then .Check1(I).Value = 0 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
        Next I
    End With
    
    If Err.Number <> 0 Then Err.Clear
End Sub




Public Sub BloquearChk(ByRef chk As CheckBox, B As Boolean)
'Bloquea un control de tipo CheckBox
'(IN) b : sera true o false segun si bloquea o no
    On Error Resume Next

    chk.Enabled = Not B
   
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BloquearChecks(ByRef formulario As Form, Modo As Byte)
'Bloquea controles q sean CheckBox si no estamos en Modo: 3.-Insertar, 4.-Modificar
'IN ->  formulario: formulario en el que se van a poner los controles textbox en modo visualizaci�n
'       Modo: modo del mantenimiento (Insertar, Modificar,Buscar...)
Dim B As Boolean
Dim Control As Control
    
    On Error Resume Next

    B = (Modo = 3 Or Modo = 4 Or Modo = 1)
    
    With formulario
        For Each Control In formulario.Controls
            If TypeOf Control Is CheckBox Then
                If InStr(1, Control.Name, "Aux") Then
                
                Else
                    Control.Enabled = B
                    If Modo = 3 Then Control.Value = 0 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
                End If
            End If
        Next Control
    End With
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub BloquearCombo(ByRef formulario As Form, Modo As Byte)
'Bloquea controles q sean ComboBox si no estamos en Modo: 3.-Insertar, 4.-Modificar
'IN ->  formulario: formulario en el que se van a poner los controles textbox en modo visualizaci�n
'       Modo: modo del mantenimiento (Insertar, Modificar,Buscar,...)
Dim B As Boolean
    
    On Error Resume Next

    'b = (Modo = 3 Or Modo = 4 Or Modo = 1 Or Modo = 5)
    B = (Modo = 3 Or Modo = 4 Or Modo = 1) '06/09/2005, lleve el modo 5 per a que no es puga modificar la cap�alera mentre treballe en les ll�nies
    
    With formulario
        For I = 0 To .Combo1.Count - 1
            Set vtag = New CTag
            vtag.Cargar .Combo1(I)
            If vtag.Cargado Then
                If vtag.EsClave And (Modo = 4 Or Modo = 5) Then
                    .Combo1(I).Enabled = False
                    .Combo1(I).BackColor = &H80000018 'groc
                Else
                    .Combo1(I).Enabled = B
                    If B Then
                        .Combo1(I).BackColor = vbWhite
                    Else
                        .Combo1(I).BackColor = &H80000018 'Amarillo Claro
                    End If
                    If Modo = 3 Then .Combo1(I).ListIndex = -1 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
                End If
            End If
        Next I
    End With
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub BloquearComboANTIC(ByRef formulario As Form, Modo As Byte)
''Bloquea controles q sean ComboBox si no estamos en Modo: 3.-Insertar, 4.-Modificar
''IN ->  formulario: formulario en el que se van a poner los controles textbox en modo visualizaci�n
''       Modo: modo del mantenimiento (Insertar, Modificar,Buscar,...)
'Dim b As Boolean
'On Error Resume Next
'
'    b = (Modo = 3 Or Modo = 4 Or Modo = 1)
'    With formulario
'        For i = 0 To .Combo1.Count - 1
'            .Combo1(i).Enabled = b
'            If b Then
'                .Combo1(i).BackColor = vbWhite
'            Else
'                .Combo1(i).BackColor = &H80000018 'Amarillo Claro
'            End If
'            If Modo = 3 Then .Combo1(i).ListIndex = 0 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
'        Next i
'    End With
'    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub BloquearImgBuscar(ByRef formulario As Form, Modo As Byte, Optional ModoLineas As Byte)
'Bloquea controles q sean ComboBox si no estamos en Modo: 3.-Insertar, 4.-Modificar
'IN ->  formulario: formulario en el que se van a poner los controles textbox en modo visualizaci�n
'       Modo: modo del mantenimiento (Insertar, Modificar,Buscar, Insertar/Modificar Lineas...)
Dim B As Boolean
On Error Resume Next

'    b = (Modo = 3 Or Modo = 4 Or Modo = 1 Or (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2)))
    B = (Modo = 3 Or Modo = 4 Or Modo = 1)
    
    With formulario
        For I = 0 To .imgBuscar.Count - 1
            .imgBuscar(I).Enabled = B
            .imgBuscar(I).Visible = B
        Next I
    End With
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BloquearImgBuscar2(ByRef formulario As Form, Modo As Byte, Optional ModoLineas As Byte)
'Bloquea controles q sean ComboBox si no estamos en Modo: 3.-Insertar, 4.-Modificar
'IN ->  formulario: formulario en el que se van a poner los controles textbox en modo visualizaci�n
'       Modo: modo del mantenimiento (Insertar, Modificar,Buscar, Insertar/Modificar Lineas...)
'En el TAG del ImgBuscar pongo un 1 si la imagen pertenece a  al tabla principal
'y un 0 si pertenece a los frame txtAux
Dim B As Boolean
'Dim bAux As Boolean
    
    On Error Resume Next

    B = (Modo = 3 Or Modo = 4 Or Modo = 1) 'Or (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2)))
'    bAux = (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2))
    
    With formulario
        For I = 0 To .imgBuscar.Count - 1
            If .imgBuscar(I).Tag = 1 Then 'esta en la cabecera
                .imgBuscar(I).Enabled = B
                .imgBuscar(I).Visible = B
            Else 'esta en las lineas
                .imgBuscar(I).Enabled = False
                .imgBuscar(I).Visible = False
            End If
        Next I
    End With
    If Err.Number <> 0 Then Err.Clear
End Sub






Public Sub BloquearImgZoom(ByRef formulario As Form, Modo As Byte, Optional ModoLineas As Byte)
'Bloquea los controles q sean Image zoom si no estamos en Modo: 3.-Insertar, 4.-Modificar
'(IN) -> formulario: formulario en el que se van a poner los controles Image zoom en modo visualizaci�n
'(IN) -> Modo: modo del mantenimiento (Insertar, Modificar,Buscar, Insertar/Modificar Lineas...)

    Dim B As Boolean

    On Error Resume Next

    B = (Modo = 3 Or Modo = 4 Or Modo = 2 Or (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2)))
    With formulario
        For I = 0 To .imgZoom.Count - 1
            .imgZoom(I).Enabled = B
            .imgZoom(I).Visible = B
        Next I
    End With
    
    If Err.Number <> 0 Then Err.Clear
End Sub




Public Sub BloquearImgFec(ByRef formulario As Form, Index As Integer, Modo As Byte, Optional ModoLineas As Byte)
Dim B As Boolean
    On Error Resume Next

    B = (Modo = 3 Or Modo = 4 Or Modo = 1 Or (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2)))
    formulario.imgFec(Index).Enabled = B
    formulario.imgFec(Index).Visible = B
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BloquearImage(ByRef img As Image, B As Boolean)

    On Error Resume Next
    
    img.Enabled = Not B
    img.Visible = Not B
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BloquearList(ByRef List As ListBox, B As Boolean)
On Error Resume Next

    'List.Locked = b
    List.Enabled = Not B
    'If Not b And List.Enabled = False Then List.Enabled = True
    If B Then
        List.BackColor = &H80000018 'Amarillo Claro
    Else
        List.BackColor = vbWhite
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BloquearOption(ByRef Opt As OptionButton, B As Boolean)
On Error Resume Next

    'Opt.Locked = b
    Opt.Enabled = Not B
    'If Not b And Opt.Enabled = False Then Opt.Enabled = True
    If B Then
        Opt.BackColor = &H80000018 'Amarillo Claro
    Else
        Opt.BackColor = vbWhite
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub PonerLongCamposGnral(ByRef formulario As Form, Modo As Byte, Opcion As Byte)
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'ya que en busqueda se permite introducir criterios m�s largos del tama�o del campo
'en busqueda permitimos escribir: "0001:0004"
'en cambio al insertar o modificar la longitud solo debe permitir ser: "0001"
'(IN) formulario y Modo en que se encuentra el formulario
'(IN) Opcion : 1 para los TEXT1, 3 para los txtAux

    Dim I As Integer
    
    On Error Resume Next

    With formulario
        If Modo = 1 Then 'BUSQUEDA
            Select Case Opcion
                Case 1 'Para los TEXT1
                    For I = 0 To .Text1.Count - 1
                        With .Text1(I)
                            If .MaxLength <> 0 Then
                               .HelpContextID = .MaxLength 'guardamos es maxlenth para reestablecerlo despues
                                .MaxLength = 0 'tama�o infinito
                            End If
                        End With
                    Next I
                
                Case 3 'para los TXTAUX
                    For I = 0 To .txtAux.Count - 1
                        With .txtAux(I)
                            If .MaxLength <> 0 Then
                               .HelpContextID = .MaxLength 'guardamos es maxlenth para reestablecerlo despues
                                .MaxLength = 0 'tama�o infinito
                            End If
                        End With
                    Next I
            End Select
            
        Else 'resto de modos
            Select Case Opcion
                Case 1 'par los Text1
                    For I = 0 To .Text1.Count - 1
                        With .Text1(I)
                            If .HelpContextID <> 0 Then
                                .MaxLength = .HelpContextID 'volvemos a poner el valor real del maxlenth
                                .HelpContextID = 0
                            End If
                        End With
                    Next I
                Case 3 'para los txtAux
                    For I = 0 To .txtAux.Count - 1
                        With .txtAux(I)
                            If .HelpContextID <> 0 Then
                                .MaxLength = .HelpContextID 'volvemos a poner el valor real del maxlenth
                                .HelpContextID = 0
                            End If
                        End With
                    Next I
            End Select
        End If
    End With
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub DesplazamientoData(ByRef vData As Adodc, Index As Integer)
'Para desplazarse por los registros de control Data
    If vData.Recordset.EOF Then Exit Sub
    Select Case Index
        Case 0 'Primer Registro
            If Not vData.Recordset.BOF Then vData.Recordset.MoveFirst
        Case 1 'Anterior
            vData.Recordset.MovePrevious
            If vData.Recordset.BOF Then vData.Recordset.MoveFirst
        Case 2 'Siguiente
            vData.Recordset.MoveNext
            If vData.Recordset.EOF Then vData.Recordset.MoveLast
        Case 3 'Ultimo
            vData.Recordset.MoveLast
    End Select
End Sub


'===========================
Public Function SituarData(ByRef vData As Adodc, vWhere As String, ByRef Indicador As String, Optional NoRefresca As Boolean) As Boolean
'Situa un DataControl en el registo que cumple vwhere
    On Error GoTo ESituarData

        'Actualizamos el recordset
        If Not NoRefresca Then vData.Refresh
        
        'El sql para que se situe en el registro en especial es el siguiente
        vData.Recordset.Find vWhere
        If vData.Recordset.EOF Then
            If vData.Recordset.RecordCount > 0 Then vData.Recordset.MoveFirst
            GoTo ESituarData
        End If
        Indicador = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
        SituarData = True
        Exit Function

ESituarData:
        If Err.Number <> 0 Then Err.Clear
        SituarData = False
End Function

'===========================
Public Function SituarDataMULTI(ByRef vData As Adodc, vWhere As String, ByRef Indicador As String, Optional NoRefresca As Boolean) As Boolean
'Situa un DataControl en el registo que cumple vwhere
On Error GoTo ESituarData
        'Actualizamos el recordset
        If Not NoRefresca Then vData.Refresh
        'El sql para que se situe en el registro en especial es el siguiente
        Multi_Find vData.Recordset, vWhere
        'vData.Recordset.Find vWhere
        If vData.Recordset.EOF Then GoTo ESituarData
        Indicador = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
        SituarDataMULTI = True
        Exit Function
ESituarData:
        If Err.Number <> 0 Then Err.Clear
        SituarDataMULTI = False
End Function


Public Sub Multi_Find(ByRef oRs As ADODB.Recordset, sCriteria As String)

    Dim clone_rs As ADODB.Recordset
    Set clone_rs = oRs.Clone
    
    clone_rs.Filter = sCriteria
    
    If clone_rs.EOF Or clone_rs.BOF Then
     oRs.MoveLast
     oRs.MoveNext
    Else
     oRs.Bookmark = clone_rs.Bookmark
    End If
    
    clone_rs.Close
    Set clone_rs = Nothing

End Sub

'===========================
'## SUSTITUIR POR SituarDataMulti
Public Function SituarDataGen(ByRef vData As Adodc, ByRef T1 As TextBox, ByRef T2 As TextBox, Indicador As String) As Boolean
'Situa un DataControl en el registo que cumple vwhere
Dim mTag1 As CTag, mTag2 As CTag
Dim Valor1 As Variant, valor2 As Variant
Dim Dato1, Dato2
Dim Encontrado As Boolean
On Error GoTo ESituarData

    SituarDataGen = False
    
    If T1.Tag <> "" And T2.Tag <> "" Then
        'Cargamos el Tag del TEXT1
        Set mTag1 = New CTag
        mTag1.Cargar T1
        If mTag1.Cargado Then
            Select Case mTag1.TipoDato
                Case "T": Valor1 = T1.Text
                Case "N": Valor1 = Val(T1.Text)
            End Select
        Else
            Exit Function
        End If
        
        'Cargamos el Tag del TEXT2
        Set mTag2 = New CTag
        mTag2.Cargar T2
        If mTag2.Cargado Then
            Select Case mTag2.TipoDato
                Case "T": valor2 = T2.Text
                Case "N": valor2 = Val(T2.Text)
            End Select
        Else
            Exit Function
        End If
        
        'Actualizamos el recordset
        vData.Refresh
        If vData.Recordset.EOF Then GoTo ESituarData
        
        Encontrado = False
        While Not Encontrado And Not vData.Recordset.EOF
            'valor del dato de la columna asociada al Text1
            Select Case mTag1.TipoDato
                Case "T": Dato1 = vData.Recordset.Fields(mTag1.columna).Value
                Case "N": Dato1 = Val(vData.Recordset.Fields(mTag1.columna).Value)
            End Select

            'valor del dato de la columna asociada al Text2
            Select Case mTag2.TipoDato
                Case "T": Dato2 = vData.Recordset.Fields(mTag2.columna).Value
                Case "N": Dato2 = Val(vData.Recordset.Fields(mTag2.columna).Value)
            End Select

            If Dato1 = Valor1 And Dato2 = valor2 Then
'                If cod3 = "" Then
                    Encontrado = True
'                Else
'                    Select Case T3
'                        Case "T": Dato3 = vData.Recordset.Fields(2).Value
'                        Case "N": Dato3 = Val(vData.Recordset.Fields(2).Value)
'                    End Select
'                    If Dato3 = valor3 Then
'                        encontrado = True
'                    Else
'                        vData.Recordset.MoveNext
'                    End If
'                End If
            Else
                vData.Recordset.MoveNext
            End If
        Wend
        Set mTag1 = Nothing
        Set mTag2 = Nothing
        Indicador = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
        SituarDataGen = True
        Exit Function
    End If
ESituarData:
        If Err.Number <> 0 Then Err.Clear
        SituarDataGen = False
End Function




'===========================
Public Function SituarDataPosicion(ByRef vData As Adodc, NumPos As Long, ByRef Indicador As String) As Boolean
'Situa un DataControl en el registro que ocupa la posicion NumPos
Dim TotalReg As Long

    On Error GoTo ESituarDataPosicion
    
'        'Actualizamos el recordset
'        If Not NoRefresca Then vdata.Refresh

        TotalReg = vData.Recordset.RecordCount
        
        If vData.Recordset.EOF Then GoTo ESituarDataPosicion
        
        If NumPos <= TotalReg Then
            vData.Recordset.Move NumPos - 1
        Else
            vData.Recordset.Move NumPos
        End If
        Indicador = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
        SituarDataPosicion = True
        Exit Function
        
ESituarDataPosicion:
        If Err.Number <> 0 Then Err.Clear
        SituarDataPosicion = False
End Function






Public Function SituarDataTrasEliminar(ByRef vData As Adodc, NumReg, Optional no_refre As Boolean) As Boolean
    On Error GoTo ESituarDataElim

    If Not no_refre Then vData.Refresh 'quan siga False o no es passe a la funci�, es refrescar�. Hi ha que passar-lo com a True quan el manteniment siga Grid per a que no refresque
    
    If Not vData.Recordset.EOF Then    'Solo habia un registro
        If NumReg > vData.Recordset.RecordCount Then
            vData.Recordset.MoveLast
        Else
            vData.Recordset.MoveFirst
            vData.Recordset.Move NumReg - 1
        End If
        SituarDataTrasEliminar = True
    Else
        SituarDataTrasEliminar = False
    End If
        
ESituarDataElim:
    If Err.Number <> 0 Then
        Err.Clear
        SituarDataTrasEliminar = False
    End If
End Function


Public Sub PonerFoco(ByRef Text As TextBox)
On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PonerFocoBtn(ByRef Btn As CommandButton)
On Error Resume Next
    If Btn.Visible Then Btn.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PonerFocoCmb(ByRef combo As ComboBox)
On Error Resume Next
    combo.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub PonerFocoChk(ByRef chk As CheckBox)
On Error Resume Next
    chk.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub PonerFocoGrid(ByRef DGrid As DataGrid)
    On Error Resume Next
    DGrid.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub




Public Sub ConseguirFoco(ByRef Text As TextBox, Modo As Byte)
'Acciones que se realizan en el evento:GotFocus de los TextBox:Text1
'en los formularios de Mantenimiento
On Error Resume Next

    If (Modo <> 0 And Modo <> 2) Then
        If Modo = 1 Then 'Modo 1: Busqueda
            Text.BackColor = vbYellow
        End If
        Text.SelStart = 0
        Text.SelLength = Len(Text.Text)
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub ConseguirFocoLin(ByRef Text As TextBox)
'Acciones que se realizan en el evento:GotFocus de los TextBox:TxtAux para LINEAS
'en los formularios de Mantenimiento
On Error Resume Next

'    If (Modo <> 0 And Modo <> 2) Then
'        If Modo = 1 Then 'Modo 1: Busqueda
'            Text.BackColor = vbYellow
'        End If
        With Text
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
'    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function PerderFocoGnral(ByRef Text As TextBox, Modo As Byte) As Boolean
Dim Comprobar As Boolean
'Dim mTag As CTag

    On Error Resume Next

    If Screen.ActiveForm.ActiveControl.Name = "cmdCancelar" Then
        PerderFocoGnral = False
        Exit Function
    End If

    With Text
        'Quitamos blancos por los lados
        .Text = Trim(.Text)
        
        
         If .BackColor = vbYellow Then
            If .Locked Then
                .BackColor = &H80000018
            Else
                .BackColor = vbWhite
            End If
        End If
        
        
        'Si no estamos en modo: 3=Insertar o 4=Modificar o 1=Busqueda, no hacer ninguna comprobacion
        If (Modo <> 3 And Modo <> 4 And Modo <> 1 And Modo <> 5) Then
            PerderFocoGnral = False
            Exit Function
        End If
        
        If Modo = 1 Then
            'Si estamos en modo busqueda y contiene un caracter especial no realizar
            'las comprobaciones
            Comprobar = ContieneCaracterBusqueda(.Text)
            If Comprobar Then
                PerderFocoGnral = False
                Exit Function
            End If
        End If
        PerderFocoGnral = True
    End With
    
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function PerderFocoGnralLineas(ByRef Txt As TextBox, ModoLineas As Byte) As Boolean
'Para el LostFocus de los txtAux de Mto de lineas


    On Error Resume Next

    If Screen.ActiveForm.ActiveControl.Name = "cmdCancelar" Then
        PerderFocoGnralLineas = False
        Exit Function
    End If
    
    With Txt
        'Quitamos blancos por los lados
        .Text = Trim(.Text)

        If .BackColor = vbYellow Then
'    '        Text1(Index).BackColor = &H80000018
            .BackColor = vbWhite
        End If

        'Si no estamos en modo: 1=Insertar o 4=Modificar o 1=Busqueda, no hacer ninguna comprobacion
        If (ModoLineas <> 1 And ModoLineas <> 2) Then
            PerderFocoGnralLineas = False
            Exit Function
        End If
    End With

    PerderFocoGnralLineas = True
    If Err.Number <> 0 Then Err.Clear



'Dim Comprobar As Boolean
'On Error Resume Next
'    With Txt
'
'        'Quitamos blancos por los lados
'        .Text = Trim(.Text)
'
'        If .BackColor = vbYellow Then
'    '        Text1(Index).BackColor = &H80000018
'            .BackColor = vbWhite
'        End If
'
'        'Si no estamos en modo: 1=Insertar o 4=Modificar o 1=Busqueda, no hacer ninguna comprobacion
'        If (ModoLineas <> 1 And ModoLineas <> 2 And ModoLineas <> 1) Then
'            PerderFocoGnralLineas = False
'            Exit Function
'        End If
'
'        If ModoLineas = 1 Then
'            'Si estamos en modo busqueda y contiene un caracter especial no realizar
'            'las comprobaciones
'            Comprobar = ContieneCaracterBusqueda(.Text)
'            If Comprobar Then
'                PerderFocoGnralLineas = False
'                Exit Function
'            End If
'        End If
'        PerderFocoGnralLineas = True
'    End With
'    If Err.Number <> 0 Then Err.Clear
End Function


Public Sub Limpiar(ByRef formulario As Form)
    Dim Control As Object

    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Control.Text = ""
        End If
    Next Control
End Sub



Public Sub LimpiarText1(ByRef formulario As Form)
'Dim i As Integer
'
'    With formulario
'        For i = 0 To .Text1.Count - 1
'            .Text1(i).Text = ""
'        Next i
'    End With
End Sub


Public Sub LimpiarTxtAux(ByRef formulario As Form)
'Dim i As Integer
'
'    With formulario
'        For i = 0 To .txtAux.Count - 1
'            .txtAux(i).Text = ""
'        Next i
'    End With
End Sub


Public Sub LimpiarLin(ByRef formulario As Form, nomFrame As String)
'Limpiar los controles Text que esten dentro del frame nomFrame
    Dim Control As Object

    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            If Control.Container.Name = nomFrame Then
                Control.Text = ""
            End If
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Container.Name = nomFrame Then
                Control.ListIndex = -1
            End If
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Container.Name = nomFrame Then
                Control.Value = 0
            End If
        End If
    Next Control
End Sub



Public Function EsVacio(ByRef Campo As TextBox) As Boolean
'    If (campo.Text = "" Or campo.Text = "0") Then
'        EsVacio = True
'    Else
'        EsVacio = False
'    End If
End Function




Public Sub DesplazamientoVisible(ByRef toolb As Toolbar, iniBoton As Byte, bol As Boolean, nreg As Byte)
'Oculta o Muestra las botones de desplazamiento de la toolbar
Dim I As Byte

    Select Case nreg
        Case 0, 1 '0 o 1 registro no mostrar los botones despl.
            For I = iniBoton To iniBoton + 3
                toolb.Buttons(I).Visible = False
            Next I
        Case Else '>1 reg, mostrar si bol
            For I = iniBoton To iniBoton + 3
                toolb.Buttons(I).Visible = bol
            Next I
    End Select
End Sub



Public Function EsNumerico(Texto As String) As Boolean
Dim I As Integer
Dim C As Integer
Dim L As Integer
Dim Cad As String
Dim B As Boolean
    
    EsNumerico = False
    B = True
    Cad = ""
    If Not IsNumeric(Texto) Then
        Cad = "El campo debe ser num�rico"
        B = False
        '======= A�ade Laura
        'formato: (.25)
        I = InStr(1, Texto, ".")
        If I = 1 Then
            If IsNumeric(Mid(Texto, 2, Len(Texto))) Then B = True
        End If
        '======================
    Else
        'Vemos si ha puesto mas de un punto
        C = 0
        L = 1
        Do
            I = InStr(L, Texto, ",")
            If I > 0 Then
                L = I + 1
                C = C + 1
            End If
        Loop Until I = 0
        If C > 1 Then
            Cad = "Numero de comas incorrecto"
            B = False
        End If
        
        'Si no ha puesto ninguna coma y tiene m�s de un punto
        If C = 0 Then
            L = 1
            Do
                I = InStr(L, Texto, ".")
                If I > 0 Then
                    L = I + 1
                    C = C + 1
                End If
            Loop Until I = 0
            If C > 1 Then
                Cad = "Numero incorrecto"
                B = False
            End If
        End If
    End If
    If Not B Then
        MsgBox Cad, vbExclamation
    Else
        EsNumerico = B
    End If
End Function


Public Function PonerFormatoEntero(ByRef T As TextBox) As Boolean
'Comprueba que el valor del textbox es un entero y le pone el formato
Dim mTag As CTag
Dim Cad As String
Dim Formato As String
On Error GoTo EPonerFormato

    
    If T.Text = "" Then Exit Function
    PonerFormatoEntero = True
    
    Set mTag = New CTag
    mTag.Cargar T
    If mTag.Cargado Then
       Cad = mTag.Nombre 'descripcion del campo
       Formato = mTag.Formato
    End If
    Set mTag = Nothing

    If Not EsEntero(T.Text) Then
        PonerFormatoEntero = False
        MsgBox "El campo " & Cad & " tiene que ser num�rico.", vbExclamation
        PonerFoco T
    Else
         'T.Text = Format(T.Text, Formato)
         ' **** 21-11-2005 Canvi de C�sar. Per a que formatetge be si es posa un
         ' n�mero negatiu, li lleve un 0 a la m�scara per a que el n�mero
         ' c�piga dins del textbox en el maxlength asignat.
         ' Si es crida a esta funci� la m�scara es del tipo 0000
         If T.Text < 0 Then _
            Formato = Replace(Formato, "0", "", 1, 1)
        ' *************************************************************************
         
         T.Text = Format(T.Text, Formato)
    End If
    
EPonerFormato:
    If Err.Number <> 0 Then Err.Clear
End Function

Public Function PosarFormatTelefon(ByRef T As TextBox) As Boolean
'Comprova que el Tel�fon/Fax/M�bil no te espais en blanc i nom�s t� n�meros
Dim mTag As CTag
Dim Cad As String

On Error GoTo EPosarFormatTelefon

    If T.Text = "" Then Exit Function
    PosarFormatTelefon = True
    
    T.Text = Replace(T.Text, " ", "")
       
    Set mTag = New CTag
    mTag.Cargar T
    If mTag.Cargado Then
       Cad = mTag.Nombre 'descripci� del camp
    End If
    Set mTag = Nothing

    If (InStr(1, T.Text, ",") > 0) Or (InStr(1, T.Text, ".") > 0) Or (InStr(1, T.Text, "+") > 0) Or (InStr(1, T.Text, "-") > 0) Or (Not IsNumeric(T.Text)) Then
        PosarFormatTelefon = False
        MsgBox "El campo " & Cad & " tiene que ser num�rico.", vbExclamation
        PonerFoco T
    End If
    
EPosarFormatTelefon:
    If Err.Number <> 0 Then Err.Clear
End Function


'=================================
Public Function PonerFormatoDecimal(ByRef T As TextBox, tipoF As Single) As Boolean
'tipoF: tipo de Formato a aplicar
'  1 -> Decimal(12,2)
'  2 -> Decimal(10,4)
'  3 -> Decimal(10,2)
'  4 -> Decimal(5,2)
'  5 -> Decimal(8,4)
Dim valor As Currency
Dim PEntera As Currency
Dim NoOK As Boolean
Dim I As Byte
Dim cadEnt As String
'Dim mTas As CTag

    If T.Text = "" Then Exit Function
    PonerFormatoDecimal = False
    NoOK = False
    With T
'        If Not EsEntero(.Text) Then
        If Not EsNumerico(CStr(.Text)) Then
'             MsgBox "El campo debe ser num�rico.", vbExclamation
'            .Text = ""
            PonerFoco T
            Exit Function
        End If


        If InStr(1, .Text, ",") > 0 Then
            valor = ImporteFormateado(.Text)
        Else
            cadEnt = .Text
            I = InStr(1, cadEnt, ".")
            If I > 0 Then cadEnt = Mid(cadEnt, 1, I - 1)
            If tipoF = 1 And Len(cadEnt) > 10 Then
                MsgBox "El valor no puede ser mayor de 9999999999,99", vbExclamation
                NoOK = True
            End If
            If NoOK Then
'                    .Text = ""
                T.SetFocus
                Exit Function
            End If
            valor = CCur(TransformaPuntosComas(.Text))
        End If
            
        'Comprobar la longitud de la Parte Entera
        PEntera = Int(valor)
        Select Case tipoF 'Comprobar longitud
            Case 1 'Decimal(12,2)
                If Len(CStr(PEntera)) > 10 Then
                    MsgBox "El valor no puede ser mayor de 9999999999,99", vbExclamation
                    NoOK = True
                End If
            Case 2 'Decimal(10,4)
                If Len(CStr(PEntera)) > 6 Then
                    MsgBox "El valor no puede ser mayor de 999999,9999", vbExclamation
                    NoOK = True
                End If
            Case 3 'Decimal(10,2)
                If Len(CStr(PEntera)) > 8 Then
                    MsgBox "El valor no puede ser mayor de 999999,99", vbExclamation
                    NoOK = True
                End If
            Case 4 'Decimal(5,2)
                If Len(CStr(PEntera)) > 3 Or ((Len(CStr(PEntera)) = 3) And (valor > 100)) Then
                    MsgBox "El valor no puede ser mayor de 100,00", vbExclamation
                    NoOK = True
                End If
            Case 5 'Decimal(8,4)
                If Len(CStr(PEntera)) > 4 Then
                    MsgBox "El valor no puede ser mayor de 9999,9999", vbExclamation
                    NoOK = True
                End If
        End Select




'       valor = CCur(TransformaPuntosComas(.Text))
'        If Not EsNumerico(CStr(valor)) Then
'             MsgBox "El campo debe ser num�rico.", vbExclamation
''            .Text = ""
'            PonerFoco T
'        Else
'            Set mTag = New CTag
'            If mTag.Cargar(T) Then
'                NoOK = mTag.Comprobar(T)
'                If NoOK = False Then Exit Function
'            End If
'            Set mTag = Nothing
            
           
            
            If NoOK Then
                PonerFormatoDecimal = False
'                .Text = ""
                T.SetFocus
                Exit Function
            End If
            
            'Poner el Formato
            Select Case tipoF
                Case 1 'Formato Decimal(12,2)
                    .Text = Format(valor, FormatoImporte)
                Case 2 'Formato Decimal(10,4)
                    .Text = Format(valor, FormatoPrecio)
                Case 3 'Formato Decimal(10,2)
                    .Text = Format(valor, FormatoDec10d2)
                Case 4 'Formato Decimal(5,2)
                    .Text = Format(valor, FormatoPorcen)
                Case 5 'Formato Decimal(8,4)
                    .Text = Format(valor, FormatoKms)
            End Select
            PonerFormatoDecimal = True
'        End If
    End With
End Function


Public Function PonerNombreDeCod(ByRef Txt As TextBox, Tabla As String, Campo As String, Optional Codigo As String, Optional tipo As String, Optional cBD As Byte, Optional codigo2 As String, Optional valor2 As String, Optional tipo2 As String) As String
'Devuelve el nombre/Descripci�n asociado al C�digo correspondiente
'Adem�s pone formato al campo txt del c�digo a partir del Tag
Dim SQL As String
Dim devuelve As String
Dim vtag As CTag
Dim ValorCodigo As String

    On Error GoTo EPonerNombresDeCod

    ValorCodigo = Txt.Text
    If ValorCodigo <> "" Then
        Set vtag = New CTag
        If vtag.Cargar(Txt) Then
            If Codigo = "" Then Codigo = vtag.columna
            If tipo = "" Then tipo = vtag.TipoDato
            
            If cBD = 0 Then cBD = cPTours
            SQL = DevuelveDesdeBDNew(cBD, Tabla, Campo, Codigo, ValorCodigo, tipo, , codigo2, valor2, tipo2)
            If vtag.TipoDato = "N" Then ValorCodigo = Format(ValorCodigo, vtag.Formato)
            Txt.Text = ValorCodigo 'Valor codigo formateado
            If SQL = "" Then
'                If vtag.Nombre <> "" Then
'                    devuelve = "No existe el " & vtag.Nombre & ": " & ValorCodigo
'                Else
'                    devuelve = "No existe el " & Texto & ": " & ValorCodigo
'                End If
'                MsgBox devuelve, vbExclamation
'                Txt.Text = ""
'                PonerFoco Txt
            Else
                PonerNombreDeCod = SQL 'Descripcion del codigo
            End If
        End If
        Set vtag = Nothing
    Else
        PonerNombreDeCod = ""
    End If
'    Exit Function
EPonerNombresDeCod:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Nombre asociado a c�digo: " & Codigo, Err.Description
End Function





Public Sub PonerIndicador(ByRef lblIndicador As Label, Modo As Byte, Optional ModoLineas As Byte)
'Pone el titulo del label lblIndicador
    Select Case Modo
        Case 0    'Modo Inicial
            lblIndicador.Caption = ""
        Case 1 'Modo Buscar
            lblIndicador.Caption = "BUSQUEDA"
        Case 2    'Preparamos para que pueda Modificar
'            lblIndicador.Caption = ""

        Case 3 'Modo Insertar
            lblIndicador.Caption = "INSERTAR"
        Case 4 'MODIFICAR
            lblIndicador.Caption = "MODIFICAR"
            
        Case 5 'Modo Lineas
            If ModoLineas = 1 Then
                lblIndicador.Caption = "INSERTAR LINEA"
            ElseIf ModoLineas = 2 Then
                lblIndicador.Caption = "MODIFICAR LINEA"
            End If
        Case Else
            lblIndicador.Caption = ""
    End Select
End Sub

Public Function PonerContRegistros(ByRef vData As Adodc) As String
'indicador del registro donde nos encontramos: "1 de 20"
    On Error GoTo EPonerReg
    
    If Not vData.Recordset.EOF Then
        PonerContRegistros = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
    Else
        PonerContRegistros = ""
    End If
    
EPonerReg:
    If Err.Number <> 0 Then
        Err.Clear
        PonerContRegistros = ""
    End If
End Function


Public Sub KEYdown(KeyCode As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
On Error Resume Next
    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
            SendKeys "+{tab}"
        Case 40 'Desplazamiento Flecha Hacia Abajo
            SendKeys "{tab}"
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub AnyadirLinea(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
On Error Resume Next

    vDataGrid.AllowAddNew = True
    If vData.Recordset.RecordCount > 0 Then
        vDataGrid.HoldFields
        vData.Recordset.MoveLast
        vDataGrid.Row = vDataGrid.Row + 1
    End If
    vDataGrid.Enabled = False
    
    If Err.Number <> 0 Then Err.Clear
End Sub





Public Function LanzaHomeGnral(nomWeb As String) As Boolean
On Error GoTo ELanzaHome
Dim Ruta As String

    LanzaHomeGnral = False
    'Obtenemos la pagina web de los parametros
'    CadenaDesdeOtroForm = DevuelveDesdeBDNew(conAri, "spara1", Opcion, "codigo", "1", "N")
'    If CadenaDesdeOtroForm = "" Then
'        MsgBox "Falta configurar los datos en Par�metros de la Aplicaci�n.", vbExclamation
'        Exit Function
'    End If
    If nomWeb = "" Then
        MsgBox "No hay una direcci�n Web para mostrar.", vbInformation
        Exit Function
    End If

'    If Opcion = "webversion" Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "?version=" & App.Major & "." & App.Minor & "." & App.Revision

    'Lanzamos
    Ruta = "C:\Archivos de programa\Internet Explorer\IEXPLORE.EXE"
'    If vConfig.Explorador <> "" Then
'       Shell vConfig.Explorador & " " & nomWeb, vbMaximizedFocus
        Shell Ruta & " " & nomWeb, vbMaximizedFocus
        LanzaHomeGnral = True
'    End If
ELanzaHome:
    If Err.Number <> 0 Then MuestraError Err.Number, nomWeb & vbCrLf & Err.Description
'    CadenaDesdeOtroForm = ""
End Function



Public Function LanzaMailGnral(dirMail As String) As Boolean
'LLama al Programa de Correo (Outlook,...)
On Error GoTo ELanzaHome

    LanzaMailGnral = False
    If dirMail = "" Then
        MsgBox "No hay direcci�n e-mail a la que enviar.", vbExclamation
        Exit Function
    End If

    Call ShellExecute(Hwnd, "Open", "mailto: " & dirMail, "", "", vbNormalFocus)
    LanzaMailGnral = True
    
ELanzaHome:
    If Err.Number <> 0 Then MuestraError Err.Number, vbCrLf & Err.Description
'    CadenaDesdeOtroForm = ""
End Function


'Public Sub SubirItemList(ByRef LView As ListView)
''Subir el item seleccionado del listview una posicion
'Dim i As Byte, Item As Byte
'Dim Aux As String
'On Error Resume Next
'
'    For i = 2 To LView.ListItems.Count
'        If LView.ListItems(i).Selected Then
'            Item = i
'            Aux = LView.ListItems(i).Text
'            LView.ListItems(i).Text = LView.ListItems(i - 1).Text
'            LView.ListItems(i - 1).Text = Aux
'        End If
'    Next i
'    If Item <> 0 Then
'        LView.ListItems(Item).Selected = False
'        LView.ListItems(Item - 1).Selected = True
'    End If
'    LView.SetFocus
'    If Err.Number <> 0 Then Err.Clear
'End Sub
'
'
'Public Sub BajarItemList(ByRef LView As ListView)
''Bajar el item seleccionado del listview una posicion
'Dim i As Byte, Item As Byte
'Dim Aux As String
'On Error Resume Next
'
'    For i = 1 To LView.ListItems.Count - 1
'        If LView.ListItems(i).Selected Then
'            Item = i
'            Aux = LView.ListItems(i).Text
'            LView.ListItems(i).Text = LView.ListItems(i + 1).Text
'            LView.ListItems(i + 1).Text = Aux
'        End If
'    Next i
'    If Item <> 0 Then
'        LView.ListItems(Item).Selected = False
'        LView.ListItems(Item + 1).Selected = True
'    End If
'    LView.SetFocus
'    If Err.Number <> 0 Then Err.Clear
'End Sub


Public Function EsCodigoCero(Cod As String, Formato As String) As Boolean
    EsCodigoCero = False
    If Cod <> "" Then
        If IsNumeric(Cod) Then
            If Val(Cod) = Val(0) Then
                EsCodigoCero = True
                MsgBox "El c�digo " & Formato & " no se puede modificar ni eliminar.", vbExclamation
                Screen.MousePointer = vbDefault
            End If
        End If
    End If
End Function




Public Sub CargaGridGnral(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, SQL As String, PrimeraVez As Boolean)
    On Error GoTo ECargaGRid

    vDataGrid.Enabled = True
    '    vdata.Recordset.Cancel
    vData.ConnectionString = conn
    vData.RecordSource = SQL
    vData.CursorType = adOpenDynamic
    vData.LockType = adLockPessimistic
    vDataGrid.ScrollBars = dbgNone
    vData.Refresh
    
    Set vDataGrid.DataSource = vData
    vDataGrid.AllowRowSizing = False
  
    vDataGrid.RowHeight = 290
   
    If PrimeraVez Then
        vDataGrid.ClearFields
        vDataGrid.ReBind
        vDataGrid.Refresh
    End If
    
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "CargaGrid", Err.Description
End Sub


Public Sub DeseleccionaGrid(ByRef vDataGrid As DataGrid)
    On Error GoTo EDeseleccionaGrid

    While vDataGrid.SelBookmarks.Count > 0
        vDataGrid.SelBookmarks.Remove 0
    Wend
    vDataGrid.SelStartCol = -1
    vDataGrid.SelEndCol = -1
    
    Exit Sub
        
EDeseleccionaGrid:
    Err.Clear
End Sub



Public Sub PosicionarCombo(Combo1 As ComboBox, valor As Integer)
'Situa el combo en la posicion de un valor concreto
Dim J As Integer

    On Error GoTo EPosCombo
    
    For J = 0 To Combo1.ListCount - 1
        If Combo1.ItemData(J) = valor Then
            Combo1.ListIndex = J
            Exit For
        End If
    Next J

EPosCombo:
    If Err.Number <> 0 Then Err.Clear
End Sub






'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
'   FUNCIONES Para PLANNER TOURS
'----------------------------------------------------------------------------
'----------------------------------------------------------------------------

Public Sub DatosPoblacion(codPobla As String, desPobla As String, CPostal As String, Provi As String, PAIS As String, Optional Prefix As String)
'IN --> codPobla
'OUT -> desPobla (Descripcion de la poblacion)
'        CPostal, Provi, Pais
Dim SQL As String
Dim RS As ADODB.Recordset

    If codPobla <> "" Then
        If EsEntero(codPobla) Then
            SQL = "SELECT poblacio.despobla,poblacio.codposta, provinci.desprovi, naciones.desnacio, provinci.preprovi"
            SQL = SQL & " FROM poblacio, provinci, naciones WHERE codpobla= " & codPobla
            SQL = SQL & " AND provinci.codprovi = poblacio.codprovi AND naciones.codnacio = provinci.codnacio"

            Set RS = New ADODB.Recordset
            RS.Open SQL, conn, , , adCmdText
            If Not RS.EOF Then
                codPobla = Format(codPobla, "000000")
                desPobla = RS.Fields!desPobla
                CPostal = DBLet(RS.Fields!codposta, "T")
                Provi = RS.Fields!desProvi
                PAIS = RS.Fields!desnacio
                If Not IsNull(RS.Fields!preprovi) Then _
                    Prefix = CStr(RS.Fields!preprovi)
            Else
'                MsgBox "No existe el c�digo de Poblaci�n: " & codPobla, vbInformation
                codPobla = "NoExiste"
                desPobla = ""
                CPostal = ""
                Provi = ""
                PAIS = ""
                Prefix = ""
            End If
            RS.Close
            Set RS = Nothing
        Else
             MsgBox "El C�digo de Poblaci�n debe ser num�rico.", vbInformation
             codPobla = ""
        End If
    Else
        codPobla = ""
        desPobla = ""
        CPostal = ""
        Provi = ""
        PAIS = ""
    End If
End Sub


Public Sub PonerDatosPoblacion(ByRef Tcpob As TextBox, ByRef Tdpob As TextBox, Optional Tcp As TextBox, Optional Tdprov As TextBox, Optional Tdpai As TextBox, Optional Nuevo As Boolean, Optional Telefon As TextBox)
Dim codPobla As String, desPobla As String
Dim CPostal As String
Dim desProvi As String, desPais As String
Dim Prefix As String
Dim cadMen As String

    codPobla = Tcpob.Text
    DatosPoblacion codPobla, desPobla, CPostal, desProvi, desPais, Prefix
    Tdpob.Text = desPobla
    'Tcp.Text = CPostal
    If Not Tcp Is Nothing Then Tcp.Text = CPostal
    If Not Tdprov Is Nothing Then Tdprov.Text = desProvi
    If Not Tdpai Is Nothing Then Tdpai.Text = desPais
    If (Not Telefon Is Nothing) Then _
        If (Telefon.Text = "") Then Telefon.Text = Prefix
'    If Not Tdprov = Nothing Then
'        Tdprov.Text = desProvi
'    End If
'    Tdpai.Text = desPais
    If codPobla = "NoExiste" Then
        'cadMen = "No existe el c�digo de Poblaci�n: " & Format(Tcpob.Text, "000000")
        cadMen = "No existe la Poblaci�n: " & Format(Tcpob.Text, "000000")
        cadMen = cadMen & vbCrLf & "�Desea Crearla?"
        If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
            Nuevo = True
'                    Indice = Index
'            Set frmPob = New frmPoblacio
'            frmPob.DatosADevolverBusqueda = "0|1|2|3|4|"
'            frmPob.NuevoCodigo = Tcpob.Text
'            Tcpob.Text = ""
'            TerminaBloquear
'            frmPob.Show vbModal
'            Set frmPob = Nothing
'            If Modo = 4 Then Bloquea = True
        Else
            codPobla = ""
            Tcpob.Text = codPobla
        End If
            PonerFoco Tcpob
    Else
        Tcpob.Text = codPobla 'Devuelve el campo formateado
    End If
End Sub




'Public Function PonerNomCliente(ByRef T As TextBox) As String
''Obtiene la cadena "apellido, nombre" o "nom.comercial" del cliente del codigo en T
''segun sea una persona o empresa.
'Dim cad As String, cadNom As String
'Dim tipCli As String 'tipo de cliente (persona/empresa)
'On Error Resume Next
'
'    If T.Text = "" Then
'        PonerNomCliente = ""
'        Exit Function
'    End If
'
'    If PonerFormatoEntero(T) Then
''    If Not EsEntero(T.Text) Then
''        '***************+ canviar el mensage ***********************
''        MsgBox "El C�digo de Cliente tiene que ser num�rico", vbExclamation
''        '**********************************************************++
''        T.Text = ""
''        PonerNomCliente = ""
''        PonerFoco T
''        Exit Function
''    Else
'        cad = "nom_come" 'nombre persona/nom comercial empresa
'        tipCli = DevuelveDesdeBDNew(cPTours, "clientes", "tipclien", "codclien", T.Text, "N", cad)
'        If tipCli = "" Then
'            MsgBox "No existe el cliente: " & T.Text, vbExclamation
'            T.Text = ""
'            PonerFoco T
'        ElseIf tipCli = 1 Then 'persona
'            T.Text = Format(T.Text, "000000")
'            'obtenemos el Apellido
'            cadNom = DevuelveDesdeBDNew(cPTours, "clientes", "ape_raso", "codclien", T.Text, "N")
'            If cadNom <> "" Then
'                cadNom = cadNom & ", " & cad 'apellido, nombre
'                PonerNomCliente = cadNom
'            End If
'        ElseIf tipCli = 2 Then 'empresa
'            T.Text = Format(T.Text, "000000")
'            PonerNomCliente = cad
'        End If
'    End If
'    If Err.Number <> 0 Then Err.Clear
'End Function
'



'Public Function PonerNomClienteNew(ByRef T As TextBox, Optional cadNIF As String, Optional MuestraMen As Boolean) As String
''Obtiene la cadena "apellido, nombre" o "nom.comercial" del cliente del codigo en T
''segun sea una persona o empresa.
''(IN) T: campo Text del codigo de cliente del cual queremos obtener el nombre
''(OUT) cadNIF : NIF del cliente
'
'    Dim cadNom As String
''    Dim tipCli As String 'tipo de cliente (persona/empresa)
'    Dim cCli As CCliente
'
'    On Error Resume Next
'
'    If T.Text = "" Then
'        PonerNomClienteNew = ""
'        Exit Function
'    End If
'
'    If PonerFormatoEntero(T) Then
'        Set cCli = New CCliente
'        If cCli.LeerDatos(T.Text) Then
'            T.Text = Format(T.Text, "000000")
'
'            If cCli.TipoClien = 1 Then 'Persona
'                cadNom = cCli.Ape_RazSoc & ", " & cCli.Nom_Come
'                cadNIF = cCli.NIF_CIF
'            ElseIf cCli.TipoClien = 2 Then 'Empresa
'                cadNom = cCli.Nom_Come
'                cadNIF = cCli.NIF_CIF
'            End If
'            PonerNomClienteNew = cadNom
'        ElseIf MuestraMen Then
'            MsgBox "No existe el cliente: " & T.Text, vbExclamation
''            T.Text = ""
'            PonerFoco T
'        End If
'        Set cCli = Nothing
'    End If
'
'    If Err.Number <> 0 Then Err.Clear
'End Function
'
'
'Public Function PonerForPagoCliente(codCli As String, codEmp As String, ByRef txtFP As TextBox, Optional MuestraMen As Boolean) As String
'    Dim cCli As CCliente
'
'    On Error Resume Next
'
'    If Not (codCli <> "" And codEmp <> "") Then
'        PonerForPagoCliente = ""
'        Exit Function
'    End If
'
''    If PonerFormatoEntero(T) Then
'        Set cCli = New CCliente
'        If cCli.LeerDatosFactu(codCli, codEmp) Then
''            T.Text = Format(T.Text, "000000")
'            txtFP.Text = cCli.ForPago
'            FormateaCampo txtFP
'            PonerForPagoCliente = cCli.DescForPago
'
'        ElseIf MuestraMen Then
''            codEmp = "No existe la forma de pago : " & txtFP
''            MsgBox "No existe la forma de pago : " & T.Text, vbExclamation
'''            T.Text = ""
''            PonerFoco T
'        End If
'        Set cCli = Nothing
''    End If
'
'    If Err.Number <> 0 Then Err.Clear
'End Function
'
'
'
'
'Public Function PonerNomEmpleado(ByRef T As TextBox, Optional MuestraMen As Boolean) As String
''Obtiene la cadena "apellido, nombre" del empleado del codigo en T
''(IN) T: campo Text del codigo de empleado del cual queremos obtener el nombre
'    Dim cEmp As CEmpleado
'
'    On Error GoTo ENomEmple
'
'    If T.Text = "" Then
'        PonerNomEmpleado = ""
'        Exit Function
'    End If
'
'    If PonerFormatoEntero(T) Then
'        Set cEmp = New CEmpleado
'        If cEmp.LeerDatos(T.Text) Then
'            FormateaCampo T
'            PonerNomEmpleado = cEmp.NombreEmple & " " & cEmp.ApellidoEmple
'        ElseIf MuestraMen Then
'            MsgBox "No existe el empleado: " & T.Text, vbExclamation
'            T.Text = ""
'            PonerFoco T
'        End If
'        Set cEmp = Nothing
'    End If
'
'ENomEmple:
'    If Err.Number <> 0 Then Err.Clear
'End Function
'
'
'
'Public Function PonerNomGuia(ByRef T As TextBox, Optional MuestraMen As Boolean) As String
''Obtiene la cadena "nombre apellido1 apellido2" del guia del codigo en T
''(IN) T: campo Text del codigo del guia del cual queremos obtener el nombre
'    Dim cGui As CGuia
'
'    On Error GoTo ENomGuia
'
'    If T.Text = "" Then
'        PonerNomGuia = ""
'        Exit Function
'    End If
'
'    If PonerFormatoEntero(T) Then
'        Set cGui = New CGuia
'        If cGui.LeerDatos(T.Text) Then
'            FormateaCampo T
'            PonerNomGuia = cGui.NombreGuia & " " & cGui.Apellido1Guia & " " & cGui.Apellido2Guia
'        ElseIf MuestraMen Then
'            MsgBox "No existe el guia de viaje: " & T.Text, vbExclamation
'            T.Text = ""
'            PonerFoco T
'        End If
'        Set cGui = Nothing
'    End If
'
'ENomGuia:
'    If Err.Number <> 0 Then Err.Clear
'End Function
'
'
'
'
'Public Function PonerNomProveedor(ByRef T As TextBox) As String
''Obtiene la cadena "nombre" del proveedor del codigo en T
'Dim cadNom As String
'Dim cadMen As String
'
'    On Error Resume Next
'
''    Nuevo = False
'
'    If T.Text = "" Then
'        PonerNomProveedor = ""
'        Exit Function
'    End If
'
'    If Not EsEntero(T.Text) Then
'        '***************+ canviar el mensage ***********************
'        MsgBox "El C�digo de Proveedor tiene que ser num�rico", vbExclamation
'        '**********************************************************++
'        T.Text = ""
'        PonerNomProveedor = ""
'        PonerFoco T
'        Exit Function
'    Else
'        cadNom = PonerNombreDeCod(T, "proveedo", "nomcomer", "codprove", "N")
''        cadNom = DevuelveDesdeBDnew(cPTours, "proveedor", "nomcomer", "codprove", T.Text, "N")
'        If cadNom = "" Then
'            cadMen = "No existe el proveedor: " & T.Text & vbCrLf
'            MsgBox cadMen, vbExclamation
''            cadMen = cadMen & "�Desea crearlo?" & vbCrLf
''            If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
''                Nuevo = True
''            Else
'                 T.Text = ""
''            End If
'            PonerNomProveedor = ""
'            PonerFoco T
'        Else 'empresa
''            T.Text = Format(T.Text, "000000")
'            PonerNomProveedor = cadNom
'        End If
'    End If
'    If Err.Number <> 0 Then Err.Clear
'End Function
'
'
'Public Function PonerNomProveedorNew(ByRef T As TextBox, EsPres As Boolean, Optional cadNIF As String, Optional MuestraMen As Boolean) As String
''Obtiene la cadena "nom.comercial" del proveedor del codigo en T
''(IN) T: campo Text del codigo del proveedor del cual queremos obtener el nombre
''(IN) EsPres: si es o no prestatario
''(OUT) cadNIF : NIF del proveedor
'
'Dim cadNom As String
'Dim cPro As CProveedor
'
'    On Error Resume Next
'
'    If T.Text = "" Then
'        PonerNomProveedorNew = ""
'        Exit Function
'    End If
'
'    If PonerFormatoEntero(T) Then
'        Set cPro = New CProveedor
'        If cPro.LeerDatos(T.Text) Then
'            T.Text = Format(T.Text, "000000")
'            cadNom = cPro.NomComer
'            cadNIF = cPro.NIFProve
'
'            'comprobar en los proveedores que el CIF tiene valor
'            If EsPres = False Then
'                'los prestatario no tiene CIF si tiene es porque es proveedor
'                If cadNIF = "" Then
'                    cadNom = ""
'                    cadNIF = ""
'                    MsgBox "El c�digo:" & T.Text & " es prestatario pero no proveedor.", vbExclamation
'                    PonerFoco T
'                End If
'            End If
'            PonerNomProveedorNew = cadNom
'
'        ElseIf MuestraMen Then
'            If EsPres Then
'                MsgBox "No existe el prestatario: " & T.Text, vbExclamation
'            Else
'                MsgBox "No existe el proveedor: " & T.Text, vbExclamation
'            End If
''            T.Text = ""
'            PonerFoco T
'        End If
'        Set cPro = Nothing
'    End If
'
'    If Err.Number <> 0 Then Err.Clear
'End Function
'
'
'
'
'
'
'
'Public Function PonerBancoPropio(codEmpre As String, codBanpr As String, nomBanpr As String) As String
''devuelve la cuenta: ES-2077-0014-11-01010225252
''en nomBanco devuelve el nombre del banco
'Dim SQL As String
'Dim nomEmpre As String
'Dim RS As ADODB.Recordset
'
'     'Poner banco Propio
'    If codBanpr <> "" Then
'        'comprobamos que existe el banco propio en la BD
'        SQL = DevuelveDesdeBDNew(cPTours, "bancctas", "codbanpr", "codempre", codEmpre, "N", , "codbanpr", codBanpr, "N")
'        If SQL = "" Then 'No existe el cod. banpr
'            nomEmpre = DevuelveDesdeBDNew(cPTours, "empresas", "nomempre", "codempre", codEmpre, "N")
'            SQL = "No existe el c�digo de Banco Propio: " & codBanpr
'            SQL = SQL & vbCrLf & "para la empresa: " & Format(codEmpre, "000") & " - " & nomEmpre
'            MsgBox SQL, vbExclamation
'            PonerBancoPropio = ""
'            nomBanpr = "Error"
'        Else
'            SQL = "SELECT DISTINCT naciones.ibanpais, bancctas.codbanco, bancctas.codsucur, bancctas.digcontr, bancctas.ctabanco, bancsofi.nombanco "
'            SQL = SQL & " FROM bancctas, naciones, bancsofi WHERE codempre = " & codEmpre & " AND codbanpr= " & codBanpr
'            SQL = SQL & " AND bancctas.codnacio = naciones.codnacio "
'            SQL = SQL & " AND (bancctas.codnacio = bancsofi.codnacio AND bancctas.codbanco = bancsofi.codbanco) "
'            Set RS = New ADODB.Recordset
'            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'            PonerBancoPropio = RS.Fields(0).Value & "-" & Format(RS.Fields(1).Value, "0000") & "-" & Format(RS.Fields(2).Value, "0000") & "-" & Format(RS.Fields(3).Value, "00") & "-" & Format(RS.Fields(4).Value, "0000000000")
'            nomBanpr = RS.Fields!NomBanco
'            RS.Close
'            Set RS = Nothing
'        End If
'    Else
'        PonerBancoPropio = ""
'        nomBanpr = ""
'    End If
'End Function



Public Function ValidarCuentaBancaria(ByRef txtB As TextBox, ByRef txtS As TextBox, ByRef txtDC As TextBox, ByRef txtC As TextBox) As Boolean
''IN: Controles textbox a Validar
'
'    ValidarCuentaBancaria = False
'
'    'Banco
'    If txtB.Text <> "" And Len(txtB.Text) < 4 Then
'            MsgBox "El campo Banco" & " debe tener 4 d�gitos", vbExclamation
'            PonerFoco txtB
'            Exit Function
'    End If
'
'    'Sucursal
'    If txtS.Text <> "" And Len(txtS.Text) < 4 Then
'            MsgBox "El campo Sucursal" & " debe tener 4 d�gitos", vbExclamation
'            PonerFoco txtS
'            Exit Function
'    End If
'
'    'Digito de Control
'    If txtDC.Text <> "" And Len(txtDC.Text) < 2 Then
'        MsgBox "El campo digito de control debe tener 2 d�gitos", vbExclamation
'        PonerFoco txtDC
'        Exit Function
'    End If
'
'    'Cuenta Bancaria
'    If txtC.Text <> "" And Len(txtC.Text) < 10 Then
'        MsgBox "El campo Cuenta Bancaria debe tener 10 d�gitos", vbExclamation

'        PonerFoco txtC
'        Exit Function
'    End If
'    ValidarCuentaBancaria = True
End Function






'Abrir visor documentos MIME
Public Function LanzaVisorMimeDocumento(Formhwnd As Long, Archivo As String)
    Call ShellExecute(Formhwnd, "Open", Archivo, "", "", 1)
End Function

