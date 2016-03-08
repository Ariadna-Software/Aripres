Attribute VB_Name = "modErrorer"
Option Explicit

Private Fich_ErrorLineas As Integer
Private B_FErroresLin As Boolean
Private Fecha_FicheErrorLinea As Date
Private nFichero As String
Private ErroresPorFichero As Integer
                        'Esto servira para la integracion de varios ficheros(robotics)
                                'Para poner el encabezado el nombre del fichero

'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'                 ERRORES PROCESANDO LAS LINEAS
'
'
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------


Public Sub InicializaErroresLinea(vF As Date)
    B_FErroresLin = False
    Fecha_FicheErrorLinea = vF
    ErroresPorFichero = 0
    nFichero = ""
End Sub

Public Sub AsignaNomFichero(Nombre As String)
    If B_FErroresLin Then
        'Fichero abierto
        If ErroresPorFichero <> 0 Then
            'Fichero abierto. Es , como minimo, el segundo archivo que da error
            ' COn lo cual cierro el parrafo
            Print #Fich_ErrorLineas, "": Print #Fich_ErrorLineas, "": Print #Fich_ErrorLineas, ""
            Print #Fich_ErrorLineas, "          Fin fichero: " & nFichero
            Print #Fich_ErrorLineas, "  ---------------------------------"
            Print #Fich_ErrorLineas, "": Print #Fich_ErrorLineas, "": Print #Fich_ErrorLineas, ""
        End If
    End If
    nFichero = Nombre
    ErroresPorFichero = 0
End Sub


Public Sub FinErroresLinea()
Dim Cad As String

If B_FErroresLin Then
    Close #Fich_ErrorLineas
    'Optativo
    Cad = Format(Fecha_FicheErrorLinea, "ddmmyy") & "_" & Format(Fecha_FicheErrorLinea, "hhmm")
    Cad = App.Path & "\Err" & Cad & ".log"
    MsgBox "Se han producido errores en la importacion de archivos." & vbCrLf & _
        "Vease el archivo: " & Cad & " para más información.", vbExclamation
End If
End Sub

Private Sub AbrirFicheroErrores()
Dim Cad As String

On Error GoTo Errores1
Cad = Format(Fecha_FicheErrorLinea, "ddmmyy") & "_" & Format(Fecha_FicheErrorLinea, "hhmm")
Cad = App.Path & "\Err" & Cad & ".log"
Fich_ErrorLineas = FreeFile
Open Cad For Output As #Fich_ErrorLineas
B_FErroresLin = True
Exit Sub
Errores1:
    Cad = "Error GRAVE: " & vbCrLf & Err.Number & " - " & Err.Description
    Cad = Cad & vbCrLf & vbCrLf & "La aplicacion finalizara"
    MsgBox Cad, vbCritical
    End
End Sub


Public Sub EscribeErrorLinea(Lin As String)
    If Not B_FErroresLin Then AbrirFicheroErrores
    If ErroresPorFichero = 0 Then ImprimeEncabezado
    Print #Fich_ErrorLineas, Lin
    ErroresPorFichero = ErroresPorFichero + 1
End Sub


Private Sub ImprimeEncabezado()
Dim I As Integer
    If nFichero <> "" Then
        For I = 1 To 3
            Print #Fich_ErrorLineas, ""
        Next I
        Print #Fich_ErrorLineas, "--------------------------------------------------------"
        Print #Fich_ErrorLineas, "-"
        Print #Fich_ErrorLineas, "-           " & nFichero
        Print #Fich_ErrorLineas, "-"
        Print #Fich_ErrorLineas, "-"
        
    End If
End Sub
