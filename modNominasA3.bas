Attribute VB_Name = "modNominasA3"
Option Explicit

Public Function GeneraNominaA3(FechaPago As Date) As Boolean
Dim Regs As Integer
Dim Im As Currency
Dim cad As String
Dim Aux As String
Dim NFic As Integer
Dim EsPersonaJuridica2 As Boolean
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Dim VectorDias As String
Dim RegImpBruto As String
Dim RegDias As String

Dim Importe As String
Dim Dias As Integer
Dim FecPag As String

'Dim miRsAux As ADODB.Recordset

    On Error GoTo EGen3
    GeneraNominaA3 = False

    NFic = -1

    NFic = FreeFile
    Open App.Path & "\nominaA3.txt" For Output As NFic

    cad = "03" & "00017" & "00000" ' tipo de registro + codigo de empresa + centro o codigo de trabajador
    
    FecPag = Format(Year(FechaPago), "0000") & Format(Month(FechaPago), "00") & Format(Day(FechaPago), "00")

    '                           char char  char  cur    cur
    'tmppagosmes   idTrabajador,Nombre,IRPF,SS,importe1,importe2
    Sql = "select * from tmppagosmes "
    
    Set Rs2 = New ADODB.Recordset
    Rs2.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Regs = 0
    While Not Rs2.EOF
        ' para cada trabajador he de generar 2 registros (
    
        Regs = Regs + 1
        
        ' importe bruto
        Importe = Format(Int(DBLet(Rs2!Importe1, "N")), "00000") & Format((DBLet(Rs2!Importe1, "N") - Int(DBLet(Rs2!Importe1, "N"))) * 100, "00")
        If DBLet(Rs2!Importe1, "N") >= 0 Then
            Importe = Importe & "+"
        Else
            Importe = Importe & "-"
        End If
        
        Importe = Importe & "000000000+"
        
        RegImpBruto = cad & Format(Rs2!idTrabajador, "000000") & FecPag & "001" & "001" & Importe 'cad+codtraba+fecha+incidencia+001+importe bruto
        Print #NFic, RegImpBruto
        
        ' dias trabajados
        Dias = Format(Int(DBLet(Rs2!IRPF, "N")), "00")
        VectorDias = Mid(DBLet(Rs2!Nombre, "T") & "NNNNN", 1, 31)
        RegDias = cad & Format(Rs2!idTrabajador, "000000") & FecPag & "016" & Format(Dias, "00") & "00" & DBLet(VectorDias, "T") & "00000000000000" 'cad+codtraba+fecha+016+dias+00+SSNNS..+"
        Print #NFic, RegDias
        
        '22:   DATOS DEVENGO PAGAS EXTRAS
        Importe = Format(Int(DBLet(Rs2!Importe2, "N")), "00000") & Format((DBLet(Rs2!Importe2, "N") - Int(DBLet(Rs2!Importe2, "N"))) * 100, "00")
        If DBLet(Rs2!Importe2, "N") >= 0 Then
            Importe = Importe & "+"
        Else
            Importe = Importe & "-"
        End If
        
        Importe = Importe & "000000000+"
        
        RegImpBruto = cad & Format(Rs2!idTrabajador, "000000") & FecPag & "001" & "001" & Importe 'cad+codtraba+fecha+incidencia+001+importe bruto
        Print #NFic, RegImpBruto
        
        
        Rs2.MoveNext
    Wend
    Rs2.Close
    Set Rs2 = Nothing
    
    Close (NFic)
    NFic = -1
    
    If Regs > 0 Then GeneraNominaA3 = True
    Exit Function
EGen3:
    MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    If NFic > 0 Then Close (NFic)
End Function

