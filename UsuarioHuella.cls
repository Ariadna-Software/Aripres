VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UsuarioHuella"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvarCodUsuario As String 'copia local
Private mvarFIR As String 'copia local
Private mvarMensaje As String 'copia local
Private mvarGesLabID As Long 'copia local
Private SQL As String
Private Rs As ADODB.Recordset
'-- Declaraciones relacionadas con la huella
Dim objNBioBSP As NBioBSPCOMLib.NBioBSP
Dim objDevice As IDevice
Dim objExtraction As IExtraction
Dim objMatching As IMatching
Dim objFPData As IFPData
Dim objFPImage As IFPImage
Dim szTextEncodeFIR As String ' Lleva la huella en formato texto
Dim vntminutiae1 As Variant
Dim vntminutiae2 As Variant
Private Conectado As Boolean
Private Recibido As String
'-- Fin de declaraciones relacionadas con la huella

Public Property Let GesLabID(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.GesLabID = 5
    mvarGesLabID = vData
End Property


Public Property Get GesLabID() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.GesLabID
    GesLabID = mvarGesLabID
End Property



Public Property Let Mensaje(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Mensaje = 5
    mvarMensaje = vData
End Property


Public Property Get Mensaje() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Mensaje
    Mensaje = mvarMensaje
End Property



Public Property Let FIR(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.FIR = 5
    mvarFIR = vData
End Property


Public Property Get FIR() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.FIR
    FIR = mvarFIR
End Property



Public Property Let CodUsuario(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.CodUsuario = 5
    mvarCodUsuario = vData
    '-- Forzamos a que la longitud sea 10 con ceros por delante
    mvarCodUsuario = Right(String(10, "0") & Trim(mvarCodUsuario), 10)
End Property


Public Property Get CodUsuario() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.CodUsuario
    CodUsuario = mvarCodUsuario
End Property

Public Function Leer(vCodUsuario As String, Optional Comprobando As Boolean) As Boolean
    vCodUsuario = Right(String(10, "0") & vCodUsuario, 10)
    SQL = "select * from usuarios where CodUsuario = " & GesHuellaDB.Texto(vCodUsuario)
    Set Rs = GesHuellaDB.cursor(SQL)
    If Not Rs.EOF Then
        If Not Comprobando Then
            mvarCodUsuario = Rs!CodUsuario
            mvarMensaje = Rs!Mensaje
            mvarFIR = Rs!FIR
            mvarGesLabID = Rs!GesLabID
        End If
        Leer = True
    End If
End Function

Public Function Guardar() As Boolean
    If Leer(mvarCodUsuario, True) Then
        SQL = "update usuarios set "
        SQL = SQL & " Mensaje = " & GesHuellaDB.Texto(mvarMensaje) & ", "
        SQL = SQL & " FIR = " & GesHuellaDB.Texto(mvarFIR) & ", "
        SQL = SQL & " GeslabID = " & GesHuellaDB.Numero(mvarGesLabID)
        SQL = SQL & " where CodUsuario = " & GesHuellaDB.Texto(mvarCodUsuario)
    Else
        SQL = " insert into usuarios(CodUsuario, Mensaje, FIR, GeslabID) values ("
        SQL = SQL & GesHuellaDB.Texto(mvarCodUsuario) & ","
        SQL = SQL & GesHuellaDB.Texto(mvarMensaje) & ","
        SQL = SQL & GesHuellaDB.Texto(mvarFIR) & ","
        SQL = SQL & GesHuellaDB.Numero(mvarGesLabID) & ")"
    End If
    GesHuellaDB.ejecutar SQL
End Function

'CalidadImagen
' de 0 a 100
Public Function CapturaHuella(CalidadImagen As Byte) As Boolean
    If mvarCodUsuario = "" Then Exit Function
    Dim Mens As String
    Dim FingerID As String ' identificador de huella
    CapturaHuella = False
    '-- Abrir el periférico
    objDevice.Open NBioAPI_DEVICE_ID_AUTO_DETECT
    '-- Capturar la huella del usuario con código correspondiente (Número de tarjeta (Text1(2)))
    
    objExtraction.EnrollImageQuality = CalidadImagen 'Enero 2012
    
    Call objExtraction.Enroll(mvarCodUsuario, Null)

    '-- Una vez capturada la huella cerramos el periférico
    objDevice.Close NBioAPI_DEVICE_ID_AUTO_DETECT
    If objExtraction.ErrorCode = NBioAPIERROR_NONE Then
        '-- Si no ha habido errores
        mvarFIR = objExtraction.TextEncodeFIR ' Obtenemos la huella en formato texto
        'mens = "La calidad de la imagen capturada es: " & CStr(objExtraction.EnrollImageQuality)
        'MsgBox mens, vbInformation
        CapturaHuella = True
        '-- Guardamos la iimagen
        ImagenHuella vEmpresa.DirHuellas & "\" & mvarCodUsuario & ".jpg"
    Else
        MsgBox objExtraction.ErrorDescription, vbCritical
    End If
End Function

Public Function ImagenHuella(Fichero) As Boolean
    objFPImage.Export
    objFPImage.Save Fichero, NBioAPI_IMG_TYPE_JPG, objFPImage.FingerID(0)
End Function

Public Function CargarHuella() As Boolean
    Dim FingerID As String ' identificador de huella
    '-- La expotamos al formato que necesitamos para trabajar con el lector del reloj
    Call objFPData.Export(mvarFIR, MINCONV_TYPE_FIM01_HV)
    If objFPData.TotalFingerCount > 0 Then
        FingerID = objFPData.FingerID(0)
        vntminutiae1 = objFPData.FPData(FingerID, 0) 'Info Template 1
        vntminutiae2 = objFPData.FPData(FingerID, 1) 'Info Template 2
    End If
    CargarHuella = True
End Function

Public Function CargarEnTerminal(k2 As Kreta2) As Boolean
    Dim i As Integer
    Dim B As Byte
    Dim sData As String
    Dim sExtraData As String
    Dim sIP As String
    Dim Envio As String
    Dim Espero As String
    
    Dim sTrama As String
    Dim sLonED As String
    Dim sDatED As String
    Dim sCRCED As String

    
    Dim arrbyte1() As Byte
    ReDim arrbyte1(0 To 403)
    Dim arrbyte2() As Byte
    ReDim arrbyte2(0 To 403)
    CargarEnTerminal = False
    CargarHuella
    
    arrbyte1 = vntminutiae1
    arrbyte2 = vntminutiae2
    sExtraData = ""
    If k2.ComprobarConexion = False Then Exit Function
    'Información de la primera huella.
    For i = 0 To 403
      sExtraData = sExtraData + hex2(arrbyte1(i))
    Next i
    'Información de la segunda huella.
    For i = 0 To 403
      sExtraData = sExtraData + hex2(arrbyte2(i))
    Next i
    'Información de usuario.
    '-- "564" + NumTrabajador + Usar PIN + PIN + Semanal + Presente + "3" 0123456789
    sData = "564" + Right("0000000000" & Trim(mvarCodUsuario), 10) + "00" + "1234" + "01" + "00" + "3"
    '-- Montaje de la huella en la trama
    sDatED = sExtraData
    sLonED = Right("0000" + Hex(Len(sDatED) / 2), 4)
    sCRCED = CalcCRC(sLonED + sDatED)
    Envio = sData + Chr(&H17) + sLonED + sDatED + sCRCED
    Espero = "66411"
    
    'Antes Noviembre 2011
    'If k2.EnviarRecibir(Envio, Espero) Then CargarEnTerminal = True
    If k2.EnviarRecibir(Envio, Espero) = 0 Then
        CargarEnTerminal = True
    Else
        CargarEnTerminal = False
        mvarMensaje = "" 'para que no grabe la siguiente
    End If
    '-- Si tiene mensaje lo grabamos tambien
    If mvarMensaje <> "" Then
        Envio = "594" + Right("0000000000" & Trim(mvarCodUsuario), 10) + _
                    Left(mvarMensaje & String(20, " "), 20)
        Espero = "69411"
        k2.EnviarRecibir Envio, Espero
    End If
    
    
End Function




'--------------------------------------------
'Pruebas de alta de usuario SIN huella
Public Function CargarEnTerminalSINHUELLA(k2 As Kreta2) As Boolean

    Dim B As Byte
    Dim sData As String
    
    Dim sIP As String
    Dim Envio As String
    Dim Espero As String
    
    Dim sTrama As String
    
    
    
    CargarEnTerminalSINHUELLA = False

    

  
    If k2.ComprobarConexion = False Then Exit Function
    'Información de la primera huella.

    'Información de usuario.
    '-- "564" + NumTrabajador + Usar PIN + PIN + Semanal + Presente + "3" 0123456789
    'SIN HUELLA
    '"564" + "0000000222" + "B0" + "0000" + "00" + "00"
    
    
    sData = "564" + Right("0000000000" & mvarCodUsuario, 10) + "B0" + "1234" + "01" + "00"
    
    'Envio = sData + Chr(&H17) & DevuelveHuellaGenerica
    Envio = sData
    Espero = "66411"
    
    
    'Envio desde el programa:   56400000006ECB012340000
    ' dese kreta                56400000006ECB012340100

    
    'Antes Noviembre 2011
    'If k2.EnviarRecibir(Envio, Espero) Then CargarEnTerminal = True
    If k2.EnviarRecibir(Envio, Espero) = 0 Then
        CargarEnTerminalSINHUELLA = True
    Else
        CargarEnTerminalSINHUELLA = False
        mvarMensaje = "" 'para que no grabe la siguiente
    End If
    '-- Si tiene mensaje lo grabamos tambien
    If mvarMensaje <> "" Then
        Envio = "594" + Right("0000000000" & Trim(mvarCodUsuario), 10) + _
                    Left(mvarMensaje & String(20, " "), 20)
        Espero = "69411"
        k2.EnviarRecibir Envio, Espero
    End If
    
    
End Function



Public Sub Eliminar()
    If Leer(mvarCodUsuario, True) Then
        SQL = "DELETE FROM usuarios  where CodUsuario = " & GesHuellaDB.Texto(mvarCodUsuario)
        GesHuellaDB.ejecutar SQL
    End If
End Sub






Private Sub Class_Initialize()
    '-- Inicio carga de los objetos relacionados con la huella
    Set objNBioBSP = New NBioBSPCOMLib.NBioBSP
    Set objDevice = objNBioBSP.Device
    Set objExtraction = objNBioBSP.Extraction
    Set objMatching = objNBioBSP.Matching
    Set objFPData = objNBioBSP.FPData
    Set objFPImage = objNBioBSP.FPImage
    '-- Fin carga de los objectos relacionados con la huella
End Sub

