Attribute VB_Name = "ModFunciones"
'////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////
'   En este modulo estan las funciones que recorren el form
'   usando el each for
'   Estas son
'
'   CamposSiguiente -> Nos devuelve el el text siguiente en
'           el orden del tabindex
'
'   CompForm -> Compara los valores con su tag
'
'   InsertarDesdeForm - > Crea el sql de insert e inserta
'
'   Limpiar -> Pone a "" todos los objetos text de un fo<rm
'
'   ObtenerBusqueda -> A partir de los text crea el sql a
'       partir del WHERE ( sin el).
'
'   ModifcarDesdeFormulario -> Opcion modificar. Genera el SQL
'
'   PonerDatosForma -> Pone los datos del RECORDSET en el form
'////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////
Option Explicit

Public Const ValorNulo = "Null"

Public Function CompForm(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Carga As Boolean
Dim Correcto As Boolean

    CompForm = False
    Set mTag = New CTag
    For Each Control In formulario.Controls
        
        'TEXT BOX
        If TypeOf Control Is TextBox And Control.Visible = True Then
            Carga = mTag.Cargar(Control)
            If Control.Name = "Text1" Then Debug.Print Control.Index
            If Carga = True Then
                Correcto = mTag.Comprobar(Control)
                If Not Correcto Then
                    PonerFoco Control
                    Exit Function
                End If
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If
        'COMBOBOX
        ElseIf TypeOf Control Is ComboBox And Control.Visible = True Then
            'Comprueba que los campos estan bien puestos
            If Control.Tag <> "" Then
                Carga = mTag.Cargar(Control)
                If Carga = False Then
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function

                Else
                    If mTag.Vacio = "N" And Control.ListIndex < 0 Then
                        
                           MsgBox "Seleccione una dato para: " & mTag.Nombre, vbExclamation
                           PonerFocoCmb Control
                           Exit Function
                    End If
                End If
            End If
        End If
    Next Control
    CompForm = True
End Function

'Añade: CESAR
'Para utilizar los campos con TAG dentro de un Frame
Public Function CompForm2(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Carga As Boolean
Dim Correcto As Boolean

    CompForm2 = False
    Set mTag = New CTag
    For Each Control In formulario.Controls
        'TEXT BOX
        If TypeOf Control Is TextBox Then
            If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                Carga = mTag.Cargar(Control)
                If Carga = True Then
                    Correcto = mTag.Comprobar(Control)
                    If Not Correcto Then Exit Function
                Else
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                End If
            End If
        'COMBOBOX
        ElseIf TypeOf Control Is ComboBox And Control.Visible = True Then
            'Comprueba que los campos estan bien puestos
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    Carga = mTag.Cargar(Control)
                    If Carga = False Then
                        MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                        Exit Function
    
                    Else
                        If mTag.Vacio = "N" And Control.ListIndex < 0 Then
                                MsgBox "Seleccione una dato para: " & mTag.Nombre, vbExclamation
                                Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next Control
    CompForm2 = True
End Function




'Public Function CampoSiguiente(ByRef formulario As Form, valor As Integer) As Control
'Dim Fin As Boolean
'Dim Control As Object
'
'On Error GoTo ECampoSiguiente
'
'    'Debug.Print "Llamada:  " & Valor
'    'Vemos cual es el siguiente
'    Do
'        valor = valor + 1
'        For Each Control In formulario.Controls
'            'Debug.Print "-> " & Control.Name & " - " & Control.TabIndex
'            'Si es texto monta esta parte de sql
'            If Control.TabIndex = valor Then
'                    Set CampoSiguiente = Control
'                    Fin = True
'                    Exit For
'            End If
'        Next Control
'        If Not Fin Then
'            valor = -1
'        End If
'    Loop Until Fin
'    Exit Function
'ECampoSiguiente:
'    Set CampoSiguiente = Nothing
'    Err.Clear
'End Function



'-----------------------------------
Public Function ValorParaSQL(valor, ByRef vtag As CTag) As String
Dim Dev As String
Dim d As Single
Dim I As Integer
Dim v
    Dev = ""
    If valor <> "" Then
        Select Case vtag.TipoDato
        Case "N"
            v = valor
            If InStr(1, valor, ",") Then
                If InStr(1, valor, ".") Then
                    'ABRIL 2004

                    'Ademas de la coma lleva puntos
                    v = ImporteFormateado(CStr(valor))
                    valor = v
                Else

                    v = CSng(valor)
                    valor = v
                End If
            Else

            End If
            Dev = TransformaComasPuntos(CStr(valor))

        Case "F"
            Dev = "'" & Format(valor, FormatoFecha) & "'"
            
        Case "FH"
            Dev = DBSet(valor, "FH")
        Case "T"
            Dev = CStr(valor)
            NombreSQL Dev
            Dev = "'" & Dev & "'"
        Case Else
            Dev = "'" & valor & "'"
        End Select

    Else
        'Si se permiten nulos, la "" ponemos un NULL
        If vtag.Vacio = "S" Then Dev = ValorNulo
    End If
    ValorParaSQL = Dev
End Function


Public Function InsertarDesdeForm(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim Cad As String
    
    On Error GoTo EInsertarF
    
    'Exit Function
    Set mTag = New CTag
    InsertarDesdeForm = False
    Der = ""
    Izda = ""
    For Each Control In formulario.Controls
        
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.Visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then

                    If mTag.columna <> "" Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Access
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & ""
                    
                        'Parte VALUES
                        Cad = ValorParaSQL(Control.Text, mTag)
                        If Der <> "" Then Der = Der & ","
                        Der = Der & Cad
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.Visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Izda <> "" Then Izda = Izda & ","
                'Access
                'Izda = Izda & "[" & mTag.Columna & "]"
                Izda = Izda & "" & mTag.columna & ""
                If Control.Value = 1 Then
                    Cad = "1"
                    Else
                    Cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then Cad = Abs(CBool(Cad))
                Der = Der & Cad
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox And Control.Visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Izda <> "" Then Izda = Izda & ","
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & ""
                    If Control.ListIndex = -1 Then
                        Cad = ValorNulo
                    Else
                        Cad = Control.ItemData(Control.ListIndex)
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & Cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    Cad = "INSERT INTO " & mTag.Tabla
    Cad = Cad & " (" & Izda & ") VALUES (" & Der & ");"
    conn.Execute Cad, , adCmdText
    
    InsertarDesdeForm = True
Exit Function

EInsertarF:
    MuestraError Err.Number, "Inserta. " & Err.Description
End Function


Public Function InsertarDesdeForm2(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim Cad As String
    
    On Error GoTo EInsertarF
    
    'Exit Function
    Set mTag = New CTag
    InsertarDesdeForm2 = False
    Der = ""
    Izda = ""
    
    For Each Control In formulario.Controls
    
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            If Izda <> "" Then Izda = Izda & ","
                            'Access
                            'Izda = Izda & "[" & mTag.Columna & "]"
                            Izda = Izda & "" & mTag.columna & ""
                        
                            'Parte VALUES
                            Cad = ValorParaSQL(Control.Text, mTag)
                            If Der <> "" Then Der = Der & ","
                            Der = Der & Cad
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If Izda <> "" Then Izda = Izda & ","
                    'Access
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & ""
                    If Control.Value = 1 Then
                        Cad = "1"
                        Else
                        Cad = "0"
                    End If
                    If Der <> "" Then Der = Der & ","
                    If mTag.TipoDato = "N" Then Cad = Abs(CBool(Cad))
                    Der = Der & Cad
                End If
            End If
            
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & ""
                        If Control.ListIndex = -1 Then
                            Cad = ValorNulo
                        ElseIf mTag.TipoDato = "N" Then
                            Cad = Control.ItemData(Control.ListIndex)
                        Else
                            Cad = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        If Der <> "" Then Der = Der & ","
                        Der = Der & Cad
                    End If
                End If
            End If
            
        'OPTION BUTTON
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.Value Then
                            If Izda <> "" Then Izda = Izda & ","
                            Izda = Izda & "" & mTag.columna & ""
                            Cad = Control.Index
                            If Der <> "" Then Der = Der & ","
                            Der = Der & Cad
                        End If
                    End If
                End If
            End If
            
'        ElseIf TypeOf Control Is DTPicker Then
'            If Control.Tag <> "" Then
'                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
'                    mTag.Cargar Control
'                    If mTag.Cargado Then
'                        If Control.Value Then
'                            If Izda <> "" Then Izda = Izda & ","
'                            Izda = Izda & "" & mTag.columna & ""
'                            cad = Control.index
'                            If Der <> "" Then Der = Der & ","
'                            Der = Der & cad
'                        End If
'                        If Izda <> "" Then Izda = Izda & ","
'                        Izda = Izda & "" & mTag.columna & ""
'
'                        'Parte VALUES
'                        If Control.Visible Then
'                            cad = ValorParaSQL(Control.Value, mTag)
'                        Else
'                            cad = ValorNulo
'                        End If
'                        If Der <> "" Then Der = Der & ","
'                        Der = Der & cad
'                    End If
'                End If
 '           End If
        End If
        
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    Cad = "INSERT INTO " & mTag.Tabla
    Cad = Cad & " (" & Izda & ") VALUES (" & Der & ");"
    conn.Execute Cad, , adCmdText
    
    InsertarDesdeForm2 = True
Exit Function

EInsertarF:
    MuestraError Err.Number, "Inserta. " & vbCrLf & Err.Description
End Function


Public Function CadenaInsertarDesdeForm(ByRef formulario As Form) As String
'Equivale a InsertarDesdeForm, excepto que devuelve la candena SQL y hace el execute fuera de la función.
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim Cad As String
    
    On Error GoTo EInsertarF
    
    'Exit Function
    Set mTag = New CTag
    CadenaInsertarDesdeForm = ""
    Der = ""
    Izda = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.Visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.columna <> "" Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Access
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & ""
                    
                        'Parte VALUES
                        Cad = ValorParaSQL(Control.Text, mTag)
                        If Der <> "" Then Der = Der & ","
                        Der = Der & Cad
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.Visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Izda <> "" Then Izda = Izda & ","
                'Access
                'Izda = Izda & "[" & mTag.Columna & "]"
                Izda = Izda & "" & mTag.columna & ""
                If Control.Value = 1 Then
                    Cad = "1"
                    Else
                    Cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then Cad = Abs(CBool(Cad))
                Der = Der & Cad
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox And Control.Visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Izda <> "" Then Izda = Izda & ","
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & ""
                    If Control.ListIndex = -1 Then
                        Cad = ValorNulo
                    Else
                        Cad = Control.ItemData(Control.ListIndex)
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & Cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    Cad = "INSERT INTO " & mTag.Tabla
    Cad = Cad & " (" & Izda & ") VALUES (" & Der & ");"
'    Conn.Execute cad, , adCmdText
    
    CadenaInsertarDesdeForm = Cad
Exit Function
EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function


Public Function PonerCamposForma(ByRef formulario As Form, ByRef vData As Adodc) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Cad As String
Dim valor As Variant
Dim Campo As String  'Campo en la base de datos
Dim I As Integer


    On Error GoTo EPonerCamposForma

    Set mTag = New CTag
    PonerCamposForma = False

    For Each Control In formulario.Controls
        'Debug.Print "Control: " & Control.Name
        'TEXTO
        If (TypeOf Control Is TextBox) And (Control.Visible = True) Then
'                If TypeOf control Is TextBox Then

            'Comprobamos que tenga tag
            mTag.Cargar Control
            If Control.Tag <> "" Then
                If mTag.Cargado Then
                    'Columna en la BD
                    If mTag.columna <> "" Then
                        
                        
                        Debug.Print mTag.columna
                        Campo = mTag.columna
                        
                        If mTag.Vacio = "S" Then
                            valor = DBLet(vData.Recordset.Fields(Campo))
                        Else
                            valor = vData.Recordset.Fields(Campo)
                        End If
                        
                        
                        
                        If mTag.Formato <> "" And CStr(valor) <> "" Then
                            If mTag.TipoDato = "N" Then
                                'Es numerico, entonces formatearemos y sustituiremos
                                ' La coma por el punto
                                Cad = Format(valor, mTag.Formato)
                                'Antiguo
                                'Control.Text = TransformaComasPuntos(cad)
                                'nuevo
                                Control.Text = Cad
                            Else
                                Control.Text = Format(valor, mTag.Formato)
                            End If
                        Else
                            Control.Text = valor
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf (TypeOf Control Is CheckBox) And (Control.Visible = True) Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Columna en la BD
                    Campo = mTag.columna
                    valor = vData.Recordset.Fields(Campo)
                    Else
                        valor = 0
                End If
                If IsNull(valor) Then valor = 0
                Control.Value = valor
            End If

         'COMBOBOX
         ElseIf (TypeOf Control Is ComboBox) And Control.Visible Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    Campo = mTag.columna
                    valor = DBLet(vData.Recordset.Fields(Campo))
                    I = 0
                    For I = 0 To Control.ListCount - 1
                        If Control.ItemData(I) = Val(valor) Then
                            Control.ListIndex = I
                            Exit For
                        End If
                    Next I
                    If I = Control.ListCount Then Control.ListIndex = -1
                End If 'de cargado
            End If 'de <>""
        End If
    Next Control

    'Veremos que tal
    PonerCamposForma = True
Exit Function
EPonerCamposForma:
    MuestraError Err.Number, "Poner campos formulario. "
End Function

'Añade: CESAR
'Para utilizar los campos con TAG dentro de un Frame
Public Function PonerCamposForma2(ByRef formulario As Form, ByRef vData As Adodc, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Cad As String
Dim valor As Variant
Dim Campo As String  'Campo en la base de datos
Dim I As Integer

    Set mTag = New CTag
    PonerCamposForma2 = False

    For Each Control In formulario.Controls
        'TEXTO
        If (TypeOf Control Is TextBox) Then
            'Comprobamos que tenga tag
            mTag.Cargar Control
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    If mTag.Cargado Then
                        'Columna en la BD
                        If mTag.columna <> "" Then
                            Campo = mTag.columna
                            If mTag.Vacio = "S" Then
                                valor = DBLet(vData.Recordset.Fields(Campo))
                            Else
                                valor = vData.Recordset.Fields(Campo)
                            End If
                            If mTag.Formato <> "" And CStr(valor) <> "" Then
                                If mTag.TipoDato = "N" Then
                                    'Es numerico, entonces formatearemos y sustituiremos
                                    ' La coma por el punto
                                    Cad = Format(valor, mTag.Formato)
                                    'Antiguo
                                    'Control.Text = TransformaComasPuntos(cad)
                                    'nuevo
                                    Control.Text = Cad
                                Else
                                    Control.Text = Format(valor, mTag.Formato)
                                End If
                            Else
                                Control.Text = valor
                            End If
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf (TypeOf Control Is CheckBox) Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Columna en la BD
                        Campo = mTag.columna
                        valor = vData.Recordset.Fields(Campo)
                    Else
                        valor = 0
                    End If
                    If IsNull(valor) Then valor = 0
                    Control.Value = valor
                End If
            End If

         'COMBOBOX
         ElseIf (TypeOf Control Is ComboBox) Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        Campo = mTag.columna
                        valor = DBLet(vData.Recordset.Fields(Campo))
                        I = 0
                        For I = 0 To Control.ListCount - 1
                            If Control.ItemData(I) = Val(valor) Then
                                Control.ListIndex = I
                                Exit For
                            End If
                        Next I
                        If I = Control.ListCount Then Control.ListIndex = -1
                    End If 'de cargado
                End If
            End If 'de <>""
            
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Columna en la BD
                        Campo = mTag.columna
                        valor = vData.Recordset.Fields(Campo)
                        If IsNull(valor) Then valor = 0
                        If Control.Index = valor Then
                            Control.Value = True
                        Else
                            Control.Value = False
                        End If
                    End If
                End If
            End If
            
'        ElseIf TypeOf Control Is DTPicker Then
'            If Control.Tag <> "" Then
'                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
'                    mTag.Cargar Control
'                    If mTag.Cargado Then
'                        'Columna en la BD
'                        campo = mTag.columna
'                        valor = vData.Recordset.Fields(campo)
'                        If IsNull(valor) Then valor = Now
'                        Control.Value = Format(valor, mTag.Formato)
'                    End If
'                End If
'            End If
        End If
    Next Control

    'Veremos que tal
    PonerCamposForma2 = True
Exit Function
EPonerCamposForma2:
    MuestraError Err.Number, "Poner campos formulario 2. "
End Function


Public Function ForaGrid(ByRef formulari As Form, ByRef vGrid As DataGrid, Control As Object) As Boolean
Dim mTag As CTag
Dim Cad As String
Dim valor As Variant
Dim camp As String  'Camp en la BDA
Dim I As Integer

    Set mTag = New CTag
    ForaGrid = False

    If (TypeOf Control Is TextBox) Then 'text
        mTag.Cargar Control
        If Control.Tag <> "" Then
            If mTag.Cargado Then
                If mTag.columna <> "" Then
                    camp = mTag.columna
                    If mTag.Vacio = "S" Then
                        valor = DBLet(vGrid.Columns(camp).Text)
                        'valor = DBLet(vGrid.Recordset.Fields(campo))
                    Else
                        'valor = vGrid.Columns!camp
                        valor = vGrid.Columns(camp).Text
                    End If
                    If mTag.Formato <> "" And CStr(valor) <> "" Then
                        If mTag.TipoDato = "N" Then
                            Cad = Format(valor, mTag.Formato)
                            Control.Text = Cad
                        Else
                            Control.Text = Format(valor, mTag.Formato)
                        End If
                    Else
                        Control.Text = valor
                    End If
                End If
            End If
        End If

'        'CheckBOX
'        ElseIf (TypeOf Control Is CheckBox) Then
'            If Control.Tag <> "" Then
'                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
'                    mTag.Cargar Control
'                    If mTag.Cargado Then
'                        'Columna en la BD
'                        campo = mTag.columna
'                        valor = vData.Recordset.Fields(campo)
'                        Else
'                            valor = 0
'                    End If
'                    If IsNull(valor) Then valor = 0
'                    Control.Value = valor
'                End If
'            End If
'
'         'COMBOBOX
'         ElseIf (TypeOf Control Is ComboBox) Then
'            If Control.Tag <> "" Then
'                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
'                    mTag.Cargar Control
'                    If mTag.Cargado Then
'                        campo = mTag.columna
'                        valor = DBLet(vData.Recordset.Fields(campo))
'                        i = 0
'                        For i = 0 To Control.ListCount - 1
'                            If Control.ItemData(i) = Val(valor) Then
'                                Control.ListIndex = i
'                                Exit For
'                            End If
'                        Next i
'                        If i = Control.ListCount Then Control.ListIndex = -1
'                    End If 'de cargado
'                End If
'            End If 'de <>""
    End If

    'Veremos que tal
    ForaGrid = True
Exit Function
EPosarCampsGrid:
    MuestraError Err.Number, "Poner campos grid. "
End Function


'Public Function PonerCamposFormaFrame(ByRef formulario As Form, NomTxtBox As String, ByRef vData As Adodc, Optional NomCheck As String, Optional NomCombo As String) As Boolean
'Dim Control As Object
'Dim mTag As CTag
'Dim cad As String
'Dim valor As Variant
'Dim campo As String  'Campo en la base de datos
'Dim i As Integer
'
'    Set mTag = New CTag
'    PonerCamposFormaFrame = False
'
'
'        For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox And Control.Visible = True And Control.Name = NomTxtBox Then
'            'Comprobamos que tenga tag
'            mTag.Cargar Control
''            Debug.Print Control.Parent
'            If Control.Tag <> "" Then
'                If mTag.Cargado Then
'                    'Columna en la BD
'                    If mTag.Columna <> "" Then
'                        campo = mTag.Columna
'                        If mTag.Vacio = "S" Then
'                            valor = DBLet(vData.Recordset.Fields(campo))
'                        Else
'                            valor = vData.Recordset.Fields(campo)
'                        End If
'                        If mTag.Formato <> "" And CStr(valor) <> "" Then
'                            If mTag.TipoDato = "N" Then
'                                'Es numerico, entonces formatearemos y sustituiremos
'                                ' La coma por el punto
'                                cad = Format(valor, mTag.Formato)
'                                'Antiguo
'                                'Control.Text = TransformaComasPuntos(cad)
'                                'nuevo
'                                Control.Text = cad
'                            Else
'                                Control.Text = Format(valor, mTag.Formato)
'                            End If
'                        Else
'                            Control.Text = valor
'                        End If
'                    End If
'                End If
'            End If
'        'CheckBOX
'        ElseIf TypeOf Control Is CheckBox And Control.Visible = True And Control.Name = NomCheck Then
'            If Control.Tag <> "" Then
'                mTag.Cargar Control
'                If mTag.Cargado Then
'                    'Columna en la BD
'                    campo = mTag.Columna
'                    valor = vData.Recordset.Fields(campo)
'                    Else
'                        valor = 0
'                End If
'                Control.Value = valor
'            End If
'
'         'COMBOBOX
'         ElseIf TypeOf Control Is ComboBox And Control.Visible = True And Control.Name = NomCombo Then
'            If Control.Tag <> "" Then
'                mTag.Cargar Control
'                If mTag.Cargado Then
'                    campo = mTag.Columna
'                    valor = vData.Recordset.Fields(campo)
'                    i = 0
'                    For i = 0 To Control.ListCount - 1
'                        If Control.ItemData(i) = Val(valor) Then
'                            Control.ListIndex = i
'                            Exit For
'                        End If
'                    Next i
'                    If i = Control.ListCount Then Control.ListIndex = -1
'                End If 'de cargado
'            End If 'de <>""
'        End If
'
'    Next Control
'
'    'Veremos que tal
'    PonerCamposFormaFrame = True
'Exit Function
'EPonerCamposForma:
'    MuestraError Err.Number, "Poner campos formulario. "
'End Function


Private Function ObtenerMaximoMinimo(vSQL As String, Optional vBD As Byte) As String
Dim RS As Recordset
    ObtenerMaximoMinimo = ""
    Set RS = New ADODB.Recordset
    'If vBD = cConta Then
    '    RS.Open vSQL, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    'Else
        RS.Open vSQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    'End If
    If Not RS.EOF Then
        If Not IsNull(RS.EOF) Then
            ObtenerMaximoMinimo = CStr(RS.Fields(0))
        End If
    End If
    RS.Close
    Set RS = Nothing
End Function


'====DAVID
'Public Function ObtenerBusqueda(ByRef formulario As Form) As String
'    Dim Control As Object
'    Dim Carga As Boolean
'    Dim mTag As CTag
'    Dim Aux As String
'    Dim cad As String
'    Dim SQL As String
'    Dim tabla As String
'    Dim RC As Byte
'
'    On Error GoTo EObtenerBusqueda
'
'    'Exit Function
'    Set mTag = New CTag
'    ObtenerBusqueda = ""
'    SQL = ""
'
'    'Recorremos los text en busca de ">>" o "<<"
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            Aux = Trim(Control.Text)
'            If Aux = ">>" Or Aux = "<<" Then
'                Carga = mTag.Cargar(Control)
'                If Carga Then
'                    If Aux = ">>" Then
'                        cad = " MAX(" & mTag.Columna & ")"
'                    Else
'                        cad = " MIN(" & mTag.Columna & ")"
'                    End If
'                    SQL = "Select " & cad & " from " & mTag.tabla
'                    SQL = ObtenerMaximoMinimo(SQL)
'                    Select Case mTag.TipoDato
'                    Case "N"
'                        SQL = mTag.tabla & "." & mTag.Columna & " = " & TransformaComasPuntos(SQL)
'                    Case "F"
'                        SQL = mTag.tabla & "." & mTag.Columna & " = '" & Format(SQL, "yyyy-mm-dd") & "'"
'                    Case Else
'                        SQL = mTag.tabla & "." & mTag.Columna & " = '" & SQL & "'"
'                    End Select
'                    SQL = "(" & SQL & ")"
'                End If
'            End If
'        End If
'    Next
'
'
'
'    'Recorremos los text en busca del NULL
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            Aux = Trim(Control.Text)
'            If UCase(Aux) = "NULL" Then
'                Carga = mTag.Cargar(Control)
'                If Carga Then
'
'                    SQL = mTag.tabla & "." & mTag.Columna & " is NULL"
'                    SQL = "(" & SQL & ")"
'                    Control.Text = ""
'                End If
'            End If
'        End If
'    Next
'
'
'
'    'Recorremos los textbox
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            'Cargamos el tag
'            Carga = mTag.Cargar(Control)
'            If Carga Then
'                If mTag.Cargado Then
'                    Aux = Trim(Control.Text)
'                    If Aux <> "" Then
'                        If mTag.tabla <> "" Then
'                            tabla = mTag.tabla & "."
'                        Else
'                            tabla = ""
'                        End If
'                    RC = SeparaCampoBusqueda(mTag.TipoDato, tabla & mTag.Columna, Aux, cad)
'                    If RC = 0 Then
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'            Else
'                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
'                Exit Function
'            End If
'
'        'COMBO BOX
'        ElseIf TypeOf Control Is ComboBox Then
'            mTag.Cargar Control
'            If mTag.Cargado Then
'                If Control.ListIndex > -1 Then
'                    If mTag.TipoDato <> "T" Then
'                        cad = Control.ItemData(Control.ListIndex)
'                        cad = mTag.tabla & "." & mTag.Columna & " = " & cad
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    Else
'                        cad = Control.List(Control.ListIndex)
'                        cad = mTag.tabla & "." & mTag.Columna & " = '" & cad & "'"
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'
'
'        'CHECK
'        ElseIf TypeOf Control Is CheckBox Then
'            If Control.Tag <> "" Then
'                mTag.Cargar Control
'                If mTag.Cargado Then
'                    If Control.Value = 1 Then
'                        cad = mTag.tabla & "." & mTag.Columna & " = 1"
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'        End If
'
'
'    Next Control
'    ObtenerBusqueda = SQL
'Exit Function
'EObtenerBusqueda:
'    ObtenerBusqueda = ""
'    MuestraError Err.Number, "Obtener búsqueda. "
'End Function

Public Function ObtenerBusqueda(ByRef formulario As Form, Optional vBD As Byte, Optional cadWhere As String) As String
    Dim Control As Object
    Dim Carga As Boolean
    Dim mTag As CTag
    Dim Aux As String
    Dim Cad As String
    Dim SQL As String
    Dim Tabla As String
    Dim RC As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda = ""
    SQL = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                If Control.Tag <> "" Then
                    Carga = mTag.Cargar(Control)
                    If Carga Then
                        If Aux = ">>" Then
                            Cad = " MAX(" & mTag.columna & ")"
                        Else
                            Cad = " MIN(" & mTag.columna & ")"
                        End If
                        SQL = "Select " & Cad & " from " & mTag.Tabla
                        If cadWhere <> "" Then SQL = SQL & " WHERE " & cadWhere
                        SQL = ObtenerMaximoMinimo(SQL, vBD)
                        Select Case mTag.TipoDato
                        Case "N"
                            SQL = mTag.Tabla & "." & mTag.columna & " = " & TransformaComasPuntos(SQL)
                        Case "F"
                            SQL = mTag.Tabla & "." & mTag.columna & " = '" & Format(SQL, "yyyy-mm-dd") & "'"
                        Case Else
                            SQL = mTag.Tabla & "." & mTag.columna & " = '" & SQL & "'"
                        End Select
                        SQL = "(" & SQL & ")"
                    End If
                End If
            End If
        End If
    Next

    
    
    'Recorremos los text en busca del NULL
    For Each Control In formulario.Controls
        If (TypeOf Control Is TextBox) And Control.Visible Then
            Aux = Trim(Control.Text)
            If UCase(Aux) = "NULL" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If True Then   'ponia:  Not paraRPT
                        SQL = mTag.Tabla & "." & mTag.columna & " is NULL"
                    Else
                        SQL = "{" & mTag.Tabla & "." & mTag.columna & "} is NULL"
                    End If
                    SQL = "(" & SQL & ")"
                    Control.Text = ""
                End If
            End If
        End If
    Next
    
    
    
    

    'Recorremos los textbox
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                'Cargamos el tag
                Carga = mTag.Cargar(Control)
                If Carga Then
                    Aux = Trim(Control.Text)
                    If Aux <> "" Then
                        If mTag.Tabla <> "" Then
                            Tabla = mTag.Tabla & "."
                            Else
                            Tabla = ""
                        End If
                        RC = SeparaCampoBusqueda(mTag.TipoDato, Tabla & mTag.columna, Aux, Cad)
                        If RC = 0 Then
                            If SQL <> "" Then SQL = SQL & " AND "
                            SQL = SQL & "(" & Cad & ")"
                        End If
                    End If
                Else
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                End If
            End If
        
        
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex > -1 Then
                        If mTag.TipoDato = "N" Then
                            Cad = Control.ItemData(Control.ListIndex)
                        Else
                            Cad = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        Cad = mTag.Tabla & "." & mTag.columna & " = " & Cad
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & Cad & ")"
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is CheckBox Then
            '=============== Añade: Laura, 15/04/05
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.Value = 1 Then
                        Cad = Control.Value
                        Cad = mTag.Tabla & "." & mTag.columna & " = " & Cad
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & Cad & ")"
                    End If
                End If
            End If
            '===================
        End If
    Next Control
    ObtenerBusqueda = SQL
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda = ""
    MuestraError Err.Number, "Obtener búsqueda. " & vbCrLf & Err.Description
End Function

'Añade: CESAR
'Para utilizar los campos con TAG dentro de un Frame
Public Function ObtenerBusqueda2(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As String
    Dim Control As Object
    Dim Carga As Boolean
    Dim mTag As CTag
    Dim Aux As String
    Dim Cad As String
    Dim SQL As String
    Dim Tabla As String
    Dim RC As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda2 = ""
    SQL = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                        If Aux = ">>" Then
                            Cad = " MAX(" & mTag.columna & ")"
                        Else
                            Cad = " MIN(" & mTag.columna & ")"
                        End If
                        SQL = "Select " & Cad & " from " & mTag.Tabla
                        SQL = ObtenerMaximoMinimo(SQL)
                        Select Case mTag.TipoDato
                        Case "N"
                            SQL = mTag.Tabla & "." & mTag.columna & " = " & TransformaComasPuntos(SQL)
                        Case "F"
                            SQL = mTag.Tabla & "." & mTag.columna & " = '" & Format(SQL, "yyyy-mm-dd") & "'"
                        Case Else
                            SQL = mTag.Tabla & "." & mTag.columna & " = '" & SQL & "'"
                        End Select
                        SQL = "(" & SQL & ")"
                    End If
                End If
            End If
        End If
    Next

    'Recorremos los textbox
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
          If Control.Tag <> "" Then
            'Cargamos el tag
            Carga = mTag.Cargar(Control)
            If Carga Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    Aux = Trim(Control.Text)
                    If Aux <> "" Then
                        If mTag.Tabla <> "" Then
                            Tabla = mTag.Tabla & "."
                            Else
                            Tabla = ""
                        End If
                        RC = SeparaCampoBusqueda(mTag.TipoDato, Tabla & mTag.columna, Aux, Cad)
                        If RC = 0 Then
                            If SQL <> "" Then SQL = SQL & " AND "
                            SQL = SQL & "(" & Cad & ")"
                        End If
                    End If
                End If
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If
        End If
        
        
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then ' +-+- 12/05/05: canvi de Cèsar, no te sentit passar-li un control que no té TAG +-+-
                mTag.Cargar Control
                If mTag.Cargado Then
                    If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                        If Control.ListIndex > -1 Then
                            Cad = Control.ItemData(Control.ListIndex)
                            Cad = mTag.Tabla & "." & mTag.columna & " = " & Cad
                            If SQL <> "" Then SQL = SQL & " AND "
                            SQL = SQL & "(" & Cad & ")"
                        End If
                    End If
                End If
            End If
            
         ElseIf TypeOf Control Is CheckBox Then
            '=============== Añade: Laura, 27/04/05
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.Value = 1 Then
                        Cad = Control.Value
                        Cad = mTag.Tabla & "." & mTag.columna & " = " & Cad
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & Cad & ")"
                    End If
                End If
            End If
            '===================
        End If
    Next Control
    ObtenerBusqueda2 = SQL
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda2 = ""
    MuestraError Err.Number, "Obtener búsqueda. " & vbCrLf & Err.Description
End Function


Public Function ModificaDesdeFormulario(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWhere As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadWhere = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.Visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.columna <> "" Then
                        'Sea para el where o para el update esto lo necesito
                        Aux = ValorParaSQL(Control.Text, mTag)
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If cadWhere <> "" Then cadWhere = cadWhere & " AND "
                             cadWhere = cadWhere & "(" & mTag.columna & " = " & Aux & ")"

                        Else
                            If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                            cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.Visible Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
            End If

        ElseIf TypeOf Control Is ComboBox And Control.Visible Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        'Lo pondremos para el WHERE
                         If cadWhere <> "" Then cadWhere = cadWhere & " AND "
                         cadWhere = cadWhere & "(" & mTag.columna & " = " & Aux & ")"
                    Else
                        If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                        cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                    End If
'
'
'                   If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
'                   'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
'                   cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE PaísDestinatario = 'México';
    If cadWhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.Tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWhere
    conn.Execute Aux, , adCmdText

    ModificaDesdeFormulario = True
    Exit Function
    
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function


Public Function ModificaDesdeFormulario2(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWhere As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormulario2 = False
    Set mTag = New CTag
    Aux = ""
    cadWhere = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            'Sea para el where o para el update esto lo necesito
                            Aux = ValorParaSQL(Control.Text, mTag)
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                            If mTag.EsClave Then
                                'Lo pondremos para el WHERE
                                 If cadWhere <> "" Then cadWhere = cadWhere & " AND "
                                 cadWhere = cadWhere & "(" & mTag.columna & " = " & Aux & ")"
    
                            Else
                                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                            End If
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If Control.Value = 1 Then
                        Aux = "TRUE"
                    Else
                        Aux = "FALSE"
                    End If
                    If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'Esta es para access
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If

        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.ListIndex = -1 Then
                            Aux = ValorNulo
                        ElseIf mTag.TipoDato = "N" Then
                            Aux = Control.ItemData(Control.ListIndex)
                        Else
                            Aux = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If cadWhere <> "" Then cadWhere = cadWhere & " AND "
                             cadWhere = cadWhere & "(" & mTag.columna & " = " & Aux & ")"
                        Else
                            If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                            cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                        End If
'
'
'                        If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
'                        'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
'                        cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.Value Then
                            Aux = Control.Index
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                              If mTag.EsClave Then
                                  'Lo pondremos para el WHERE
                                   If cadWhere <> "" Then cadWhere = cadWhere & " AND "
                                   cadWhere = cadWhere & "(" & mTag.columna & " = " & Aux & ")"
                              Else
                                  If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                  cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                              End If
                        End If
                    End If
                End If
            End If
            
'        ElseIf TypeOf Control Is DTPicker Then
'            If Control.Tag <> "" Then
'                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
'                    mTag.Cargar Control
'                    If mTag.Cargado Then
'                        If Control.Value Then
'                         If mTag.columna <> "" Then
'                            Aux = Control.index
'                            If Control.Visible Then
'                                Aux = ValorParaSQL(Control.Value, mTag)
'                            Else
'                                Aux = ValorNulo
'                            End If
'                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
'                            'dentro del WHERE
'                            If mTag.EsClave Then
'                                'Lo pondremos para el WHERE
'                                If cadWhere <> "" Then cadWhere = cadWhere & " AND "
'                                cadWhere = cadWhere & "(" & mTag.columna & " = " & Aux & ")"
'                            Else
'                                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
'                                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
'                            End If
'                        End If
'                    End If
'                End If
'            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE PaísDestinatario = 'México';
    If cadWhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.Tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWhere
    conn.Execute Aux, , adCmdText

    ModificaDesdeFormulario2 = True
    Exit Function
    
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar 2. " & Err.Description
End Function


Public Sub FormateaCampo(vTex As TextBox)
    Dim mTag As CTag
    Dim Cad As String
    On Error GoTo EFormateaCampo
    Set mTag = New CTag
    mTag.Cargar vTex
    If mTag.Cargado Then
        If vTex.Text <> "" Then
            If mTag.Formato <> "" Then
                Cad = TransformaPuntosComas(vTex.Text)
                Cad = Format(Cad, mTag.Formato)
                vTex.Text = Cad
            End If
        End If
    End If
EFormateaCampo:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Sub


Public Function FormatoCampo(ByRef vTex As TextBox) As String
'Devuelve el formato del campo en el TAg: "0000"
Dim mTag As CTag
Dim Cad As String
    
    On Error GoTo EFormatoCampo

    Set mTag = New CTag
    mTag.Cargar vTex
    If mTag.Cargado Then
        FormatoCampo = mTag.Formato
    End If
    
EFormatoCampo:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function


'Añade: CESAR
'Para utilizalo en el arreglaGrid
Public Function FormatoCampo2(ByRef objec As Object) As String
'Devuelve el formato del campo en el TAg: "0000"
Dim mTag As CTag
Dim Cad As String

    On Error GoTo EFormatoCampo2

    Set mTag = New CTag
    mTag.Cargar objec
    If mTag.Cargado Then
        FormatoCampo2 = mTag.Formato
    End If
    
EFormatoCampo2:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function


Public Function TipoCamp(ByRef objec As Object) As String
Dim mTag As CTag
Dim Cad As String

    On Error GoTo ETipoCamp

    Set mTag = New CTag
    mTag.Cargar objec
    If mTag.Cargado Then
        TipoCamp = mTag.TipoDato
    End If

ETipoCamp:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function


'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function RecuperaValor(ByRef CADENA As String, Orden As Integer) As String
Dim I As Integer
Dim J As Integer
Dim Cont As Integer
Dim Cad As String

    I = 0
    Cont = 1
    Cad = ""
    Do
        J = I + 1
        I = InStr(J, CADENA, "|")
        If I > 0 Then
            If Cont = Orden Then
                Cad = Mid(CADENA, J, I - J)
                I = Len(CADENA) 'Para salir del bucle
                Else
                    Cont = Cont + 1
            End If
        End If
    Loop Until I = 0
    RecuperaValor = Cad
End Function


'-----------------------------------------------------------------------
'Deshabilitar ciertas opciones del menu
'EN funcion del nivel de usuario
'Esto es a nivel general, cuando el Toolba es el mismo

'Para ello en el tag del button tendremos k poner un numero k nos diara hasta k nivel esta permitido

Public Sub PonerOpcionesMenuGeneral(ByRef formulario As Form)
Dim I As Integer
Dim J As Integer
'Dim bol As Boolean

On Error GoTo EPonerOpcionesMenuGeneral
'bol = vUsu.Nivel < 2

'Añadir, modificar y borrar deshabilitados si no nivel
'With Formulario
'    For i = 1 To .Toolbar1.Buttons.Count
'        If .Toolbar1.Buttons(i).Tag <> "" Then
'            j = Val(.Toolbar1.Buttons(i).Tag)
'            If j < vUsu.Nivel Then
'                .Toolbar1.Buttons(i).Enabled = False
'            End If
'        End If
'    Next i
'End With

Exit Sub
EPonerOpcionesMenuGeneral:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub


Public Sub PonerModoMenuGral(ByRef formulario As Form, activo As Boolean)
Dim I As Integer
'Dim j As Integer

On Error GoTo PonerModoMenuGral

'Añadir, modificar y borrar deshabilitados si no Modo
    With formulario
        For I = 1 To .Toolbar1.Buttons.Count
            Select Case .Toolbar1.Buttons(I).ToolTipText
                Case "Nuevo"
                    .Toolbar1.Buttons(I).Visible = Not .DeConsulta
                Case "Modificar", "Eliminar", "Imprimir"
                    .Toolbar1.Buttons(I).Visible = Not .DeConsulta
                    .Toolbar1.Buttons(I).Enabled = activo
'                Case "Modificar"
'                Case "Eliminar"
'                Case "Imprimir"
            End Select
        Next I
        
        
        'El menu Visible
        .mnModificar.Visible = Not .DeConsulta
        .mnEliminar.Visible = Not .DeConsulta
        'El menu activo
        .mnModificar.Enabled = activo
        .mnEliminar.Enabled = activo
    End With
    
    
Exit Sub
PonerModoMenuGral:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub




'Este modifica las claves prinipales y todo
'la sentenca del WHERE cod=1 and .. viene en claves
Public Function ModificaDesdeFormularioClaves(ByRef formulario As Form, Claves As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWhere As String
Dim cadUPDATE As String
Dim I As Integer

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormularioClaves = False
    Set mTag = New CTag
    Aux = ""
    cadWhere = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Sea para el where o para el update esto lo necesito
                    Aux = ValorParaSQL(Control.Text, mTag)
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
            End If

        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If
        End If
    Next Control
    cadWhere = Claves
    'Construimos el SQL
    If cadWhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.Tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWhere
    conn.Execute Aux, , adCmdText

ModificaDesdeFormularioClaves = True
Exit Function
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function

Public Function BLOQUEADesdeFormulario(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWhere As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQUEADesdeFormulario
    BLOQUEADesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadWhere = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.Visible = True Then
            If Control.Tag <> "" Then

                mTag.Cargar Control
                If mTag.Cargado Then
                    'Sea para el where o para el update esto lo necesito
                    Aux = ValorParaSQL(Control.Text, mTag)
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        'Lo pondremos para el WHERE
                         If cadWhere <> "" Then cadWhere = cadWhere & " AND "
                         cadWhere = cadWhere & "(" & mTag.columna & " = " & Aux & ")"
                    End If
                End If
            End If
        End If
    Next Control

    If cadWhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "select * FROM " & mTag.Tabla
        Aux = Aux & " WHERE " & cadWhere & " FOR UPDATE"

        'Intenteamos bloquear
        PreparaBloquear
        conn.Execute Aux, , adCmdText
        BLOQUEADesdeFormulario = True
    End If
    
EBLOQUEADesdeFormulario:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        TerminaBloquear
    End If
    Screen.MousePointer = AntiguoCursor
End Function


'Añade: CESAR
'Para utilizar los campos con TAG dentro de un Frame
Public Function BLOQUEADesdeFormulario2(ByRef formulario As Form, ByRef ado As Adodc, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWhere As String
Dim AntiguoCursor As Byte
Dim nomcamp As String

    On Error GoTo EBLOQUEADesdeFormulario2
    
    BLOQUEADesdeFormulario2 = False
    Set mTag = New CTag
    Aux = ""
    cadWhere = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If (TypeOf Control Is TextBox) Or (TypeOf Control Is ComboBox) Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Sea para el where o para el update esto lo necesito
                        'Aux = ValorParaSQL(Control.Text, mTag)
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                        
                            If TypeOf Control Is ComboBox Then
                                Aux = CStr(ado.Recordset.Fields(mTag.columna))
                                Aux = Control.ItemData(CInt(Aux))
                            Else
                        
                                Aux = ValorParaSQL(CStr(ado.Recordset.Fields(mTag.columna)), mTag)
                            End If
                            'Lo pondremos para el WHERE
                             If cadWhere <> "" Then cadWhere = cadWhere & " AND "
                             cadWhere = cadWhere & "(" & mTag.columna & " = " & Aux & ")"
                        End If
                    End If
                End If
            End If
        End If
    Next Control

    If cadWhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "select * FROM " & mTag.Tabla
        Aux = Aux & " WHERE " & cadWhere & " FOR UPDATE"

        'Intenteamos bloquear
        PreparaBloquear
        conn.Execute Aux, , adCmdText
        BLOQUEADesdeFormulario2 = True
    End If
    
EBLOQUEADesdeFormulario2:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla 2" & Err.Description
'        BLOQUEADesdeFormulario2 = False
        TerminaBloquear
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Public Function BloqueaRegistro(cadTABLA As String, cadWhere As String) As Boolean
Dim Aux As String

    On Error GoTo EBloqueaRegistro
        
    BloqueaRegistro = False
    Aux = "select * FROM " & cadTABLA
    Aux = Aux & " WHERE " & cadWhere & " FOR UPDATE"

    'Intenteamos bloquear
    PreparaBloquear
    conn.Execute Aux, , adCmdText
    BloqueaRegistro = True

EBloqueaRegistro:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        TerminaBloquear
    End If
End Function


Public Function BloqueaRegistroForm(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim AuxDef As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQ
    BloqueaRegistroForm = False
    Set mTag = New CTag
    Aux = ""
    AuxDef = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        Aux = ValorParaSQL(Control.Text, mTag)
                        AuxDef = AuxDef & Aux & "|"
                    End If
                End If
            End If
        End If
    Next Control

    If AuxDef = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
'        Aux = "Insert into zBloqueos(codusu,tabla,clave) VALUES(" & vUsu.Codigo & ",'" & mTag.tabla
        Aux = Aux & "',""" & AuxDef & """)"
        conn.Execute Aux
        BloqueaRegistroForm = True
    End If
EBLOQ:
    If Err.Number <> 0 Then
        Aux = ""
        If conn.Errors.Count > 0 Then
            If conn.Errors(0).NativeError = 1062 Then
                '¡Ya existe el registro, luego esta bloqueada
                Aux = "BLOQUEO"
            End If
        End If
        If Aux = "" Then
            MuestraError Err.Number, "Bloqueo tabla"
        Else
            MsgBox "Registro bloqueado por otro usuario", vbExclamation
        End If
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Public Function DesBloqueaRegistroForm(ByRef TextBoxConTag As TextBox) As Boolean
Dim mTag As CTag
Dim SQL As String

'Solo me interesa la tabla
On Error Resume Next
    Set mTag = New CTag
    mTag.Cargar TextBoxConTag
    If mTag.Cargado Then
'        SQL = "DELETE from zBloqueos where codusu=" & vUsu.Codigo & " and tabla='" & mTag.tabla & "'"
        conn.Execute SQL
        If Err.Number <> 0 Then
            Err.Clear
        End If
    End If
    Set mTag = Nothing
End Function


'====================== LAURA

Public Function ComprobarCero(valor As String) As String
    If valor = "" Then
        ComprobarCero = "0"
    Else
        ComprobarCero = valor
    End If
End Function




'Public Function PonerClaseAgencia(claAgen As Byte) As String
'Dim cad As String
'
'    If claAgen = 1 Then
'        cad = "MINORISTA"
'    ElseIf claAgen = 2 Then
'        cad = "MAYORISTA"
'    End If
'
'    PonerClaseAgencia = cad
'End Function



'Public Function PonerClaseVenta(claVen As Byte) As String
'Dim cad As String
'
'    If claVen = 1 Then
'        cad = "INDIVIDUALES"
'    ElseIf claVen = 2 Then
'        cad = "GRUPOS"
'    End If
'
'    PonerClaseVenta = cad
'End Function



Public Function CalcularImporte(Cantidad As String, Precio As String, Dto1 As String, Dto2 As String) As String
'Calcula el Importe de una linea de Oferta, Pedido, Albaran, ...
'Importe=cantidad * precio - (descuentos)
'Tipo Descuento: 0=aditivo, 1=sobre resto
'Dim vImp As Currency
'Dim vDto1 As Currency, vDto2 As Currency
'On Error Resume Next
'
'    'Como son de tipo string comprobar que si vale "" lo ponemos a 0
'    Cantidad = ComprobarCero(Cantidad)
'    Precio = ComprobarCero(Precio)
'    Dto1 = ComprobarCero(Dto1)
'    Dto2 = ComprobarCero(Dto2)
'
'
'    vImp = CCur(Cantidad) * CCur(Precio)
'    If vParamAplic.TipoDtos = 0 Then 'Dto Aditivo
'        vDto1 = (CCur(Dto1) * vImp) / 100
'        vDto2 = (CCur(Dto2) * vImp) / 100
'        vImp = vImp - vDto1 - vDto2
'    ElseIf vParamAplic.TipoDtos = 1 Then 'Sobre Resto
'        vDto1 = (CCur(Dto1) * vImp) / 100
'        vImp = vImp - vDto1
'        vDto2 = (CCur(Dto2) * vImp) / 100
'        vImp = vImp - vDto2
'    End If
'    vImp = Round(vImp, 2)
'    CalcularImporte = CStr(vImp)
End Function




'Public Function PonerNomEmple(codEmp As String) As String
'Dim nomEmp As String
'Dim cad As String
'
'    'apellidos i nombre del empleado
'    If (codEmp <> "") Then
'        nomEmp = "nomemple"
'        cad = DevuelveDesdeBDNew(cPTours, "empleado", "apeemple", "codemple", codEmp, "N", nomEmp, "codempre", CStr(vSesion.Empresa), "N", "codagenc", CStr(vSesion.Agencia), "N")
'        If cad <> "" Then cad = cad & ", " & nomEmp
'    End If
'    PonerNomEmple = cad
'End Function



Public Function ExisteCP(T As TextBox) As Boolean
'comprueba para un campo de texto que sea clave primaria, si ya existe un
'registro con ese valor
Dim vtag As CTag
Dim devuelve As String

    On Error GoTo ErrExiste

    ExisteCP = False
    If T.Text <> "" Then
        If T.Tag <> "" Then
            Set vtag = New CTag
            If vtag.Cargar(T) Then
'                If vtag.EsClave Then
                    'devuelve = DevuelveDesdeBDNew(cPTours, vtag.tabla, vtag.columna, vtag.columna, T.Text, vtag.TipoDato)
                    devuelve = DevuelveDesdeBDNew(0, vtag.Tabla, vtag.columna, vtag.columna, T.Text, vtag.TipoDato)
                    If devuelve <> "" Then
    '                    MsgBox "Ya existe un registro para " & vtag.Nombre & ": " & T.Text, vbExclamation
                        MsgBox "Ya existe el " & vtag.Nombre & ": " & T.Text, vbExclamation
                        ExisteCP = True
                        PonerFoco T
                    End If
'                End If
            End If
            Set vtag = Nothing
        End If
    End If
    Exit Function
    
ErrExiste:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar código.", Err.Description
End Function




Public Function TotalRegistros(vSQL As String) As Long
'Devuelve el valor de la SQL
'para obtener COUNT(*) de la tabla
Dim RS As ADODB.Recordset

    On Error Resume Next

    Set RS = New ADODB.Recordset
    RS.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    TotalRegistros = 0
    If Not RS.EOF Then
        If RS.Fields(0).Value > 0 Then TotalRegistros = RS.Fields(0).Value  'Solo es para saber que hay registros que mostrar
    End If
    RS.Close
    Set RS = Nothing

    If Err.Number <> 0 Then
        TotalRegistros = 0
        Err.Clear
    End If
End Function







'En el primer valor llevara el codigo (itemdata
'elñ field(1) sera descripcion
Public Sub CargaComboDesdeBD(ByRef CBO As ComboBox, SQL As String)
Dim RS As ADODB.Recordset

    CBO.Clear
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        CBO.AddItem RS.Fields(1)
        CBO.ItemData(CBO.NewIndex) = RS.Fields(0)
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
End Sub




'==== LAURA
Public Function PonerFormatoHora(ByRef T As TextBox) As Boolean
Dim Cad As String

        Cad = T.Text
        PonerFormatoHora = False
        If Cad <> "" Then
            If Not EsHoraOK(Cad) Then
                MsgBox "Hora incorrecta. (hh:mm:ss)", vbExclamation
                Cad = "mal"
            End If
            If Cad <> "" And Cad <> "mal" Then
                T.Text = Cad
                PonerFormatoHora = True
            Else
                T.Text = ""
                PonerFoco T
            End If
        End If
End Function

'=== DAVID (estaba en Modulo:bus)
Public Function EsHoraOK(T As String) As Boolean
Dim Cad As String
    
    Cad = T
    If InStr(1, Cad, ":") = 0 Then
        Select Case Len(T)
            Case 8
                Cad = Mid(Cad, 1, 2) & ":" & Mid(Cad, 3, 2) & ":" & Mid(Cad, 5)
            Case 6
                Cad = Mid(Cad, 1, 2) & ":" & Mid(Cad, 3, 2) & ":" & Mid(Cad, 5)
            Case 4
                Cad = Mid(Cad, 1, 2) & ":" & Mid(Cad, 3, 2) & ":00"
        End Select
    End If
    
    If IsDate(Cad) Then
        EsHoraOK = True
        T = Format(Cad, "hh:mm:ss")
    Else
        EsHoraOK = False
    End If
End Function




Public Function TextBoxAImporte(ByRef T As TextBox) As String
Dim Mon As Currency

    T = Trim(T.Text)
    If Not IsNumeric(T.Text) Then
        TextBoxAImporte = "No es campo numerico"
        Exit Function
    End If
    
    If InStr(1, T.Text, ",") Then
        'Ya esta formateado
        TextBoxAImporte = ""
        Exit Function
    End If
    
    'Llegados aqui solo hay un punto. Luego lo pasamos a moneda
    Mon = CCur(TransformaPuntosComas(T.Text))
    T.Text = Format(Mon, "##,###,##0.00")
        
    
End Function

