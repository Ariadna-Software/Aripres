Attribute VB_Name = "modBackup"
Option Explicit


Public Sub BACKUP_TablaIzquierda(ByRef RS As ADODB.Recordset, ByRef cadena As String)
Dim I As Integer
Dim nexo As String

    cadena = ""
    nexo = ""
    For I = 0 To RS.Fields.Count - 1
        cadena = cadena & nexo & RS.Fields(I).Name
        nexo = ","
    Next I
    cadena = "(" & cadena & ")"
End Sub





'---------------------------------------------------
'El fichero siempre sera NF
Public Sub BACKUP_Tabla(ByRef RS As ADODB.Recordset, ByRef Derecha As String)
Dim I As Integer
Dim nexo As String
Dim valor As String
Dim tipo As Integer


    On Error GoTo EBACKUP

    Derecha = ""
    nexo = ""
    For I = 0 To RS.Fields.Count - 1
        tipo = RS.Fields(I).Type
        
        If IsNull(RS.Fields(I)) Then
            valor = "NULL"
        Else
            
            'pruebas
            Select Case tipo
            'TEXTO
            Case 129, 200, 201
                valor = RS.Fields(I)
                NombreSQL valor    '.-----------> 23 Octubre 2003.
                valor = "'" & valor & "'"
            'Fecha
            Case 133
                valor = CStr(RS.Fields(I))
                valor = "'" & Format(valor, FormatoFecha) & "'"
                
            'Numero normal, sin decimales
            Case 2, 3, 16 To 19
                valor = RS.Fields(I)
            
            'Numero con decimales
            Case 131
                valor = CStr(RS.Fields(I))
                valor = TransformaComasPuntos(valor)
                
            Case 134
                'Campo HORA
                valor = "'" & Format(RS.Fields(I), "hh:mm:ss") & "'"
                
            Case 205
                
                valor = ""
                'FALTA###
                'LeerBinaryEnString RS.Fields(I), valor
            Case Else
                valor = "Error grave. Tipo de datos no tratado." & vbCrLf
                valor = valor & vbCrLf & "SQL: " & RS.Source
                valor = valor & vbCrLf & "Pos: " & I
                valor = valor & vbCrLf & "Campo: " & RS.Fields(I).Name
                valor = valor & vbCrLf & "Valor: " & RS.Fields(I)
                MsgBox valor, vbExclamation
                MsgBox "El programa finalizará. Avise al soporte técnico.", vbCritical
                End
            End Select
        End If
        Derecha = Derecha & nexo & valor
        nexo = ","
    Next I
    Derecha = "(" & Derecha & ")"
    
    
    Exit Sub
    
EBACKUP:
    MuestraError Err.Number, "Tipo dato: " & tipo & "     Valor: " & valor & vbCrLf & vbCrLf & RS.Source
End Sub
