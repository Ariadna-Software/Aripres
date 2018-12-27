Attribute VB_Name = "LibNorma34"
Option Explicit


'----------------------------------------------------------------------
'  Copia fichero generado bajo
Public Function CopiarFicheroNorma43_(Destino As String) As Boolean

    
    CopiarFicheroNorma43_ = CopiarEnDisquette(Destino) 'A disco
    
    
End Function


Private Function CopiarEnDisquette(Destino As String) As Boolean


On Error Resume Next

        CopiarEnDisquette = False
    
        FileCopy App.Path & "\norma34.txt", Destino
        If Err.Number <> 0 Then
           MsgBox "Error creando copia fichero." & vbCrLf & Err.Description, vbCritical
           Err.Clear
        Else
           MsgBox "El fichero esta guardado como: " & vbCrLf & Destino, vbInformation
           CopiarEnDisquette = True
        End If
    

End Function





'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'Cuenta propia tendra empipados entidad|sucursal|cc|cuenta|
Public Function GeneraFicheroNorma34(CIF As String, Fecha As Date, CuentaPropia As String, ConceptoTransferencia As String) As Boolean
Dim NFich As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim Importe As Currency
Dim RS As ADODB.Recordset
Dim Aux As String
Dim cad As String

    On Error GoTo EGen
    GeneraFicheroNorma34 = False
    
    
    NFich = FreeFile
    Open App.Path & "\norma34.txt" For Output As #NFich
    
    'Codigo ordenante
    CodigoOrdenante = Right("    " & CIF, 9)   'CIF EMPRESA
    
    'CABECERA
    Cabecera1 NFich, CodigoOrdenante, Fecha, CuentaPropia, cad
    Cabecera2 NFich, CodigoOrdenante, cad
    Cabecera3 NFich, CodigoOrdenante, cad
    Cabecera4 NFich, CodigoOrdenante, cad
    
    
    
    'Imprimimos las lineas
    'Para ello abrimos la tabla tmpNorma34
    Set RS = New ADODB.Recordset
    Aux = "Select * from tmpnorma34"
    RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Importe = 0
    If RS.EOF Then
        'No hayningun registro
        
    Else
        Regs = 0
        While Not RS.EOF
            Aux = RellenaAceros(RS!codsoc, False, 12)
            'Cad = "06"
            'Cad = Cad & "56"
            'Cad = Cad & " "
            Aux = "06" & "56" & " " & CodigoOrdenante & Aux  'Ordenante y socio juntos
        
            Linea1 NFich, Aux, RS, cad, ConceptoTransferencia
            Linea2 NFich, Aux, RS, cad
            Linea3 NFich, Aux, RS, cad
            Linea4 NFich, Aux, RS, cad
            Linea5 NFich, Aux, RS, cad
            Linea6 NFich, Aux, RS, cad
           
        
        
        
        
            Importe = Importe + RS!Importe
            Regs = Regs + 1
            RS.MoveNext
        Wend
        'Imprimimos totales
        Totales NFich, CodigoOrdenante, Importe, Regs, cad
    End If
    RS.Close
    Set RS = Nothing
    Close (NFich)
    If Regs > 0 Then GeneraFicheroNorma34 = True
    Exit Function
EGen:
    MuestraError Err.Number, Err.Description

End Function


Private Function RellenaABlancos(CADENA As String, PorLaDerecha As Boolean, Longitud As Integer) As String
Dim cad As String
    
    cad = Space(Longitud)
    If PorLaDerecha Then
        cad = CADENA & cad
        RellenaABlancos = Left(cad, Longitud)
    Else
        cad = cad & CADENA
        RellenaABlancos = Right(cad, Longitud)
    End If
    
End Function



Private Function RellenaAceros(CADENA As String, PorLaDerecha As Boolean, Longitud As Integer) As String
Dim cad As String
    
    cad = Mid("00000000000000000000", 1, Longitud)
    If PorLaDerecha Then
        cad = CADENA & cad
        RellenaAceros = Left(cad, Longitud)
    Else
        cad = cad & CADENA
        RellenaAceros = Right(cad, Longitud)
    End If
    
End Function



'Private Sub Cabecera1(NF As Integer,ByRef CodOrde As String)
'Dim Cad As String
'
'End Sub

Private Sub Cabecera1(NF As Integer, ByRef CodOrde As String, Fecha As Date, Cta As String, ByRef cad As String)

    cad = "03"
    cad = cad & "56"
    cad = cad & " "
    cad = cad & CodOrde
    cad = cad & Space(12) & "001"
    cad = cad & Format(Now, "ddmmyy")
    cad = cad & Format(Fecha, "ddmmyy")
    'Cuenta bancaria
    cad = cad & RecuperaValor(Cta, 1)
    cad = cad & RecuperaValor(Cta, 2)
    cad = cad & RecuperaValor(Cta, 4)
    cad = cad & "0"  'Sin relacion
    cad = cad & "   " & RecuperaValor(Cta, 3)  'Digito de control bancario
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub



Private Sub Cabecera2(NF As Integer, ByRef CodOrde As String, ByRef cad As String)
    cad = "03"
    cad = cad & "56"
    cad = cad & " "
    cad = cad & CodOrde
    cad = cad & Space(12) & "002"
    
    cad = cad & RellenaABlancos(vEmpresa.NomEmpresa, True, 30)  'Nombre empresa
  
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Cabecera3(NF As Integer, ByRef CodOrde As String, ByRef cad As String)
    cad = "03"
    cad = cad & "56"
    cad = cad & " "
    cad = cad & CodOrde
    cad = cad & Space(12) & "003"
    
    cad = cad & RellenaABlancos(vEmpresa.DirEmpresa, True, 30)   'Nombre empresa
  
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub



Private Sub Cabecera4(NF As Integer, ByRef CodOrde As String, ByRef cad As String)

    cad = "03"
    cad = cad & "56"
    cad = cad & " "
    cad = cad & CodOrde
    cad = cad & Space(12) & "004"
    
    cad = cad & RellenaABlancos(vEmpresa.CodPosEmpresa, False, 5)
    cad = cad & " "
    cad = cad & RellenaABlancos(vEmpresa.PobEmpresa, True, 30)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub



'ConceptoTransferencia
'1.- Abono nomina
'9.- Transferencia ordinaria
Private Sub Linea1(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String, vConceptoTransferencia As String)

    cad = CodOrde   'llevara tb la ID del socio
    cad = cad & "010"
    cad = cad & RellenaAceros(CStr(Round(RS1!Importe, 2) * 100), False, 12)
    
    cad = cad & RellenaABlancos(RS1!banco1, False, 4)    'Entidad
    cad = cad & RellenaABlancos(RS1!banco2, False, 4)   'Sucur
    cad = cad & RellenaABlancos(RS1!banco4, False, 10)  'Cta
    cad = cad & "1" & vConceptoTransferencia
    cad = cad & "  "
    cad = cad & RellenaABlancos(RS1!banco3, False, 2)  'Cta
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Linea2(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "011"
    cad = cad & RellenaABlancos(RS1!Nombre, False, 36)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Linea3(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "012"
    cad = cad & RellenaABlancos(RS1!Domicilio, False, 36)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Linea4(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "013"
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Linea5(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "014"
    cad = cad & RellenaABlancos(RS1!codpos, False, 5) & " "
    cad = cad & RellenaABlancos(RS1!Poblacion, False, 30)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Linea6(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "016"
    cad = cad & RellenaABlancos(RS1!concepto, False, 35)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Totales(NF As Integer, ByRef CodOrde As String, Total As Currency, Registros As Integer, ByRef cad As String)
    cad = "08" & "56 "
    cad = cad & CodOrde    'llevara tb la ID del socio
    cad = cad & Space(15)
    cad = cad & RellenaAceros(CStr(Int(Round(Total * 100, 2))), False, 12)
    cad = cad & RellenaAceros(CStr(Registros), False, 8)
    cad = cad & RellenaAceros(CStr((Registros * 6) + 4 + 1), False, 10)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub






'*******************************************************************
'SEPA
'*******************************************************************
Public Function GeneraFicheroNorma34SEPA(CIF As String, Fecha As Date, CuentaPropia2 As String, ConceptoTr As String, SufijoOEM As String) As Boolean
Dim SepaXML As Boolean

    SepaXML = DevuelveDesdeBD("SepaXML", "empresas", "1", "1") = "1"
    If False Then
        GeneraFicheroNorma34SEPA = GenFichNorma34SEPA(CIF, Fecha, CuentaPropia2, ConceptoTr, SufijoOEM)
    Else
        GeneraFicheroNorma34SEPA = GeneraFicheroNorma34SEPA_XML(CIF, Fecha, CuentaPropia2, ConceptoTr, SufijoOEM)
    End If
End Function



Public Function GenFichNorma34SEPA(CIF As String, Fecha As Date, CuentaPropia2 As String, ConceptoTr As String, SufijoOEM As String) As Boolean
Dim Regs As Integer
Dim Importe As Currency
Dim Im As Currency
Dim cad As String
Dim Aux As String

Dim miRsAux As ADODB.Recordset
Dim NF As Integer



    On Error GoTo EGen2
    GenFichNorma34SEPA = False
    

    
    
    'Cargamos la cuenta
 
    Set miRsAux = New ADODB.Recordset

    cad = RecuperaValor(CuentaPropia2, 5) & RecuperaValor(CuentaPropia2, 1) & RecuperaValor(CuentaPropia2, 2) & RecuperaValor(CuentaPropia2, 3) & RecuperaValor(CuentaPropia2, 4)
    CuentaPropia2 = cad
  
    If Len(cad) <> 24 Then
        MsgBox "Error IBAN banco : " & CuentaPropia2, vbExclamation
        Exit Function
    End If
    
    NF = FreeFile
    Open App.Path & "\norma34.txt" For Output As #NF
    
    
    
    'SEPA
    '1.- Cabecera ordenante
    '------------------------------------------------------------------------
    cad = "01" & "ORD" & "34145" & "001" & CIF
        
    'sufijo (Tenemos el OEM, que se utiliza para las otras normas antiguas
    cad = cad & SufijoOEM
    cad = cad & Format(Now, "yyyymmdd")
    cad = cad & Format(Fecha, "yyyymmdd")
    cad = cad & "A" 'IBAN
     
    'EL IBAN propiamente
    cad = cad & FrmtStr(CuentaPropia2, 34)
    cad = cad & "1" 'Cargo por cada operacion
    'Nombre
   
    cad = cad & FrmtStr(vEmpresa.NomEmpresa, 70)
    
    cad = cad & FrmtStr(Trim(vEmpresa.DirEmpresa), 50)
    cad = cad & FrmtStr(Trim(vEmpresa.CodPosEmpresa & " " & vEmpresa.PobEmpresa), 50)
    cad = cad & FrmtStr(DBLet(vEmpresa.ProvEmpresa, "T"), 40)
    
    'Pais y libre
    cad = cad & "ES" & FrmtStr("", 311)
    Print #NF, cad
  
  
  
    '2.- Registro cabecera TRANSFERENCIA
    '------------------------------------------------------------------------
    cad = "02" & "SCT" & "34145" & CIF
        
    'sufijo (Tenemos el OEM, que se utiliza para las otras normas
    cad = cad & SufijoOEM
    cad = cad & FrmtStr("", 578)
    Print #NF, cad
    
    
    
    
    cad = "SELECT tmpNorma34.*, Trabajadores.*, sbic.bic"
    cad = cad & " FROM (tmpNorma34 INNER JOIN Trabajadores ON tmpNorma34.CodSoc = Trabajadores.IdTrabajador)"
    cad = cad & " LEFT JOIN sbic ON Trabajadores.entidad = sbic.entidad;"
    
    
    
    miRsAux.Open cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    If Not miRsAux.EOF Then
        cad = "#"
        While Not miRsAux.EOF
            If IsNull(miRsAux!BIC) Then
                If InStr(1, cad, "#" & miRsAux!banco1 & "#") = 0 Then cad = cad & miRsAux!banco1 & "#"
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.MoveFirst
        
        
        If Len(cad) > 1 Then
            cad = Mid(cad, 2)
            cad = Mid(cad, 1, Len(cad) - 1)
            cad = Replace(cad, "#", "   /   ")
            cad = "Bancos sin BIC asignado:" & vbCrLf & cad & vbCrLf & vbCrLf & "¿Continuar?"
            If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then
                miRsAux.Close
                Close (NF)
                Exit Function
            End If
        End If
        
    End If
    
    
    Regs = 0
    Importe = 0
    If miRsAux.EOF Then
        'No hayningun registro

    Else
        While Not miRsAux.EOF
            
       
                Im = miRsAux!Importe
     
            Importe = Importe + Im
            Regs = Regs + 1
            
            'Campo 1,2,3
            cad = "03" & "SCT" & "34145" & "002"
            
            'Campo 5 . Referencia del ordenante
            If IsNull(miRsAux!numdni) Then
                Aux = miRsAux!codsoc
            Else
                Aux = miRsAux!numdni
            End If
                
         
            Aux = FrmtStr(Aux, 10)
            Aux = DBLet(miRsAux!concepto, "T") & " Tra:" & Format(miRsAux!codsoc, "0000") & " F:" & Format(Fecha, "dd/mm")
            
            cad = cad & FrmtStr(Aux, 35)
            
            'Campo 6
            cad = cad & "A"
            
            'IBAN
            cad = cad & FrmtStr(IBAN_Destino(miRsAux), 34)
            
            
            
            'Campo8 Importe
            cad = cad & Format(Im * 100, String(11, "0")) ' Importe
            
            'Campo9
            cad = cad & "3" 'gastos compartidos
            'Campo 10
            cad = cad & FrmtStr(DBLet(miRsAux!BIC, "T"), 11) 'BIC

            'nommacta,dirdatos,codposta,dirdatos,despobla,impvenci,scobro.codmacta
            'Datos Basicos del beneficiario
            cad = cad & DatosBasicosDelDeudor(miRsAux)
            
            'Campo16 ID del pago. Concepto
            
            Aux = DBLet(miRsAux!concepto, "T") & " " & DBLet(Fecha, "T") & " Importe " & Format(Im, FormatoImporte)
          
            cad = cad & FrmtStr(Aux, 140)
            
            'Campo17
            cad = cad & FrmtStr("", 35)  'Reservado
            
            'Campo18  campo19
            
           
            
            If ConceptoTr = "1" Then
                cad = cad & "SALASALA"
            ElseIf ConceptoTr = "0" Then
                cad = cad & "PENSPENS"
            Else
                cad = cad & "TRADTRAD"
            End If
            
           
            
            cad = cad & FrmtStr("", 99)  'libre
            
            Print #NF, cad
            
            miRsAux.MoveNext
        Wend
        
    
        'TOTALES
        '----------------------------------
        'Total trasnferencia SEPA
        'Campo 1,2
        cad = "04" & "SCT"
        
        'Campo3 Importe total
        cad = cad & Format(Importe * 100, String(17, "0")) ' Importe
        cad = cad & Format(Regs, String(8, "0")) ' Importe
        'Total registros son
        'Reg(numreo de adeudos + 1 reg01 + un reg02 + reg04
        cad = cad & Format(Regs + 2, String(10, "0")) ' Importe   '2014-01-29  HABIA un reg + 3
        cad = cad & FrmtStr("", 560)  'libre
        Print #NF, cad
        
        'Total general
        cad = "99" & "ORD"
        
        'Campo3 Importe total
        cad = cad & Format(Importe * 100, String(17, "0")) ' Importe
        cad = cad & Format(Regs, String(8, "0")) ' Importe
        
        'Igual que arriba as uno
        'Reg(numreo de adeudos + 1 reg01 + un reg02 + reg04  +1
        cad = cad & Format(Regs + 4, String(10, "0")) ' Importe
        cad = cad & FrmtStr("", 560)  'libre
        Print #NF, cad
        
        
        
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    Close (NF)
    If Regs > 0 Then GenFichNorma34SEPA = True
    Exit Function
EGen2:
    MuestraError Err.Number, Err.Description

End Function


Private Function FrmtStr(Campo As String, Longitud As Integer) As String
    FrmtStr = Mid(Trim(Campo) & Space(Longitud), 1, Longitud)
End Function

Private Function IBAN_Destino(ByRef miRsAux) As String

        IBAN_Destino = FrmtStr(DBLet(miRsAux!IBAN, "T"), 4) ' ES00
        IBAN_Destino = IBAN_Destino & Format(miRsAux!banco1, "0000") ' Código de entidad receptora
        IBAN_Destino = IBAN_Destino & Format(miRsAux!banco2, "0000") ' Código de oficina receptora
        IBAN_Destino = IBAN_Destino & Format(miRsAux!banco3, "00") ' Dígitos de control
        IBAN_Destino = IBAN_Destino & Format(miRsAux!banco4, "0000000000") ' Código de cuenta

End Function
    
Private Function DatosBasicosDelDeudor(ByRef miRsAux) As String
        DatosBasicosDelDeudor = FrmtStr(miRsAux!nomtrabajador, 70)
        'dirdatos,codposta,despobla,pais desprovi
        DatosBasicosDelDeudor = DatosBasicosDelDeudor & FrmtStr(DBLet(miRsAux!domtrabajador, "T"), 50)
        DatosBasicosDelDeudor = DatosBasicosDelDeudor & FrmtStr(Trim(DBLet(miRsAux!codpostrabajador, "T") & " " & DBLet(miRsAux!pobtrabajador, "T")), 50)
        DatosBasicosDelDeudor = DatosBasicosDelDeudor & FrmtStr(DBLet(miRsAux!pobtrabajador, "T"), 40)
        
       
        DatosBasicosDelDeudor = DatosBasicosDelDeudor & "ES"
        
End Function











'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'
'
'
'
'
'               Norma 34 SEPA XML
'
'
'
'
'
'
'
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************


Public Function GeneraFicheroNorma34SEPA_XML(CIF As String, Fecha As Date, CuentaPropia2 As String, ConceptoTr As String, SufijoOEM As String) As Boolean
Dim Regs As Integer
Dim Importe As Currency
Dim Im As Currency
Dim cad As String
Dim Aux As String
Dim NFic As Integer
Dim EsPersonaJuridica2 As Boolean
Dim miRsAux As ADODB.Recordset

    On Error GoTo EGen3
    GeneraFicheroNorma34SEPA_XML = False
    
    NFic = -1
    
    
    Set miRsAux = New ADODB.Recordset

    cad = RecuperaValor(CuentaPropia2, 5) & RecuperaValor(CuentaPropia2, 1) & RecuperaValor(CuentaPropia2, 2) & RecuperaValor(CuentaPropia2, 3) & RecuperaValor(CuentaPropia2, 4)
    CuentaPropia2 = cad
  
    If Len(cad) <> 24 Then
        MsgBox "Error IBAN banco : " & CuentaPropia2, vbExclamation
        Exit Function
    End If
    
    'Esta comprobacion deberia hacerla antes
    
    cad = "SELECT tmpNorma34.CodSoc, tmpNorma34.Nombre, tmpNorma34.Banco1, tmpNorma34.Banco2, tmpNorma34.Banco3"
    cad = cad & ", tmpNorma34.Banco4, tmpNorma34.Domicilio, tmpNorma34.Codpos, tmpNorma34.Poblacion, tmpNorma34.Concepto,"
    cad = cad & "tmpNorma34.Importe, tmpNorma34.tipo"
    
    cad = cad & ",Trabajadores.*, sbic.bic"
    cad = cad & " FROM (tmpNorma34 INNER JOIN Trabajadores ON tmpNorma34.CodSoc = Trabajadores.IdTrabajador)"
    cad = cad & " LEFT JOIN sbic ON Trabajadores.entidad = sbic.entidad;"
    miRsAux.Open cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    If Not miRsAux.EOF Then
        cad = "#"
        While Not miRsAux.EOF
            If IsNull(miRsAux!BIC) Then
                If InStr(1, cad, "#" & miRsAux!banco1 & "#") = 0 Then cad = cad & miRsAux!banco1 & "#"
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.MoveFirst
        
        
        If Len(cad) > 1 Then
            cad = Mid(cad, 2)
            cad = Mid(cad, 1, Len(cad) - 1)
            cad = Replace(cad, "#", "   /   ")
            cad = "Bancos sin BIC asignado:" & vbCrLf & cad & vbCrLf & vbCrLf & "¿Continuar?"
            If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then
                miRsAux.Close
                Set miRsAux = Nothing
                Exit Function
            End If
        End If
        
    End If
    miRsAux.Close
    
    
    
    
    
    
    
    
    
    
    
    NFic = FreeFile
    Open App.Path & "\norma34.txt" For Output As NFic
    
    
    Print #NFic, "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
    Print #NFic, "<Document xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""urn:iso:std:iso:20022:tech:xsd:pain.001.001.03"">"
    Print #NFic, "<CstmrCdtTrfInitn>"
    Print #NFic, "   <GrpHdr>"
    
    '                   NumeroTransferencia
    cad = "TRANPAG" & Format(0, "000000") & "F" & Format(Now, "yyyymmddThhnnss")
    Print #NFic, "      <MsgId>" & cad & "</MsgId>"
    Print #NFic, "      <CreDtTm>" & Format(Now, "yyyy-mm-ddThh:nn:ss") & "</CreDtTm>"
    
    'Registrp cabecera con totales
    
    Aux = "importe"
    cad = "tmpNorma34"

    cad = "Select count(*),sum(" & Aux & ") FROM " & cad & " WHERE 1 =1"
    Aux = "0|0|"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(1)) Then Aux = miRsAux.Fields(0) & "|" & miRsAux.Fields(1) & "|"
    End If
    miRsAux.Close
    
    
    
    Print #NFic, "      <NbOfTxs>" & RecuperaValor(Aux, 1) & "</NbOfTxs>"
    Print #NFic, "      <CtrlSum>" & TransformaComasPuntos(RecuperaValor(Aux, 2)) & "</CtrlSum>"
    Print #NFic, "      <InitgPty>"
    Print #NFic, "         <Nm>" & XML(vEmpresa.NomEmpresa) & "</Nm>"
    Print #NFic, "         <Id>"
    cad = Mid(CIF, 1, 1)
    
    EsPersonaJuridica2 = Not IsNumeric(cad)

    
    
    
    cad = "PrvtId"
    If EsPersonaJuridica2 Then cad = "OrgId"
    
    Print #NFic, "           <" & cad & ">"
    Print #NFic, "               <Othr>"
    Print #NFic, "                  <Id>" & CIF & SufijoOEM & "</Id>"
    Print #NFic, "               </Othr>"
    Print #NFic, "           </" & cad & ">"
    
    Print #NFic, "         </Id>"
    Print #NFic, "      </InitgPty>"
    Print #NFic, "   </GrpHdr>"

    Print #NFic, "   <PmtInf>"
    
    Print #NFic, "      <PmtInfId>" & Format(Now, "yyyymmddhhnnss") & CIF & "</PmtInfId>"
    Print #NFic, "      <PmtMtd>TRF</PmtMtd>"
    Print #NFic, "      <ReqdExctnDt>" & Format(Fecha, "yyyy-mm-dd") & "</ReqdExctnDt>"
    Print #NFic, "      <Dbtr>"
    
     'Nombre
    Print #NFic, "         <Nm>" & XML(vEmpresa.NomEmpresa) & "</Nm>"
    Print #NFic, "         <PstlAdr>"
    Print #NFic, "            <Ctry>ES</Ctry>"

    cad = vEmpresa.DirEmpresa & " "
    cad = cad & Trim(vEmpresa.PobEmpresa) & " " & vEmpresa.ProvEmpresa & " "
   
    Print #NFic, "            <AdrLine>" & XML(Trim(cad)) & "</AdrLine>"
    
    Print #NFic, "         </PstlAdr>"
    Print #NFic, "         <Id>"
    
    Aux = "PrvtId"
    If EsPersonaJuridica2 Then Aux = "OrgId"
   
    
    Print #NFic, "            <" & Aux & ">"
    
    Print #NFic, "               <Othr>"
    Print #NFic, "                  <Id>" & CIF & SufijoOEM & "</Id>"
    Print #NFic, "               </Othr>"
    Print #NFic, "            </" & Aux & ">"
    Print #NFic, "         </Id>"
    Print #NFic, "    </Dbtr>"
    
    
    Print #NFic, "    <DbtrAcct>"
    Print #NFic, "       <Id>"
    Print #NFic, "          <IBAN>" & Trim(CuentaPropia2) & "</IBAN>"
    Print #NFic, "       </Id>"
    Print #NFic, "       <Ccy>EUR</Ccy>"
    Print #NFic, "    </DbtrAcct>"
    Print #NFic, "    <DbtrAgt>"
    Print #NFic, "       <FinInstnId>"
    
    cad = Mid(CuentaPropia2, 5, 4)
    cad = DevuelveDesdeBD("bic", "sbic", "entidad", cad, "T")
    Print #NFic, "          <BIC>" & Trim(cad) & "</BIC>"
    Print #NFic, "       </FinInstnId>"
    Print #NFic, "    </DbtrAgt>"
    
    
    
    
    cad = "SELECT tmpNorma34.CodSoc, tmpNorma34.Nombre, tmpNorma34.Banco1, tmpNorma34.Banco2, tmpNorma34.Banco3"
    cad = cad & ", tmpNorma34.Banco4, tmpNorma34.Domicilio, tmpNorma34.Codpos, tmpNorma34.Poblacion, tmpNorma34.Concepto,"
    cad = cad & "tmpNorma34.Importe, tmpNorma34.tipo,Trabajadores.*, sbic.bic"
    cad = cad & " FROM (tmpNorma34 INNER JOIN Trabajadores ON tmpNorma34.CodSoc = Trabajadores.IdTrabajador)"
    cad = cad & " LEFT JOIN sbic ON Trabajadores.entidad = sbic.entidad;"
        

    miRsAux.Open cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    Regs = 0
    While Not miRsAux.EOF
        Print #NFic, "   <CdtTrfTxInf>"
        Print #NFic, "      <PmtId>"
        
         
        'IDentificador
         If IsNull(miRsAux!numdni) Then
            Aux = DBLet(miRsAux!concepto, "T") & " Tra:" & Format(miRsAux!codsoc, "0000") & " F:" & Format(Fecha, "dd/mm")
        Else
            Aux = miRsAux!numdni
        End If
    
        
        Print #NFic, "         <EndToEndId>" & Aux & "</EndToEndId>"
        Print #NFic, "      </PmtId>"
        Print #NFic, "      <PmtTpInf>"
        
        'Importe
        Im = miRsAux!Importe
        
        
        
        'Persona fisica o juridica
        cad = DBLet(miRsAux!numdni, "T")
        cad = Mid(cad, 1, 1)
        EsPersonaJuridica2 = Not IsNumeric(cad)
        'Como da problemas Cajamar, siempre ponemos Perosna juridica. Veremos
        'EsPersonaJuridica2 = True
        
        
        Importe = Importe + Im
        Regs = Regs + 1
        
        Print #NFic, "          <SvcLvl><Cd>SEPA</Cd></SvcLvl>"
        If ConceptoTr = "1" Then
            Aux = "SALA"
        ElseIf ConceptoTr = "0" Then
            Aux = "PENS"
        Else
            Aux = "TRAD"
        End If
        Print #NFic, "          <CtgyPurp><Cd>" & Aux & "</Cd></CtgyPurp>"
        Print #NFic, "       </PmtTpInf>"
        Print #NFic, "       <Amt>"
        Print #NFic, "          <InstdAmt Ccy=""EUR"">" & TransformaComasPuntos(CStr(Im)) & "</InstdAmt>"
        Print #NFic, "       </Amt>"
        Print #NFic, "       <CdtrAgt>"
        Print #NFic, "          <FinInstnId>"
        cad = DBLet(miRsAux!BIC, "T")
        If cad = "" Then Err.Raise 513, , "No existe BIC" & miRsAux!Nombre & vbCrLf & "Entidad: " & miRsAux!Entidad
        Print #NFic, "             <BIC>" & DBLet(miRsAux!BIC, "T") & "</BIC>"
        Print #NFic, "          </FinInstnId>"
        Print #NFic, "       </CdtrAgt>"
        Print #NFic, "       <Cdtr>"
        Print #NFic, "          <Nm>" & XML(miRsAux!Nombre) & "</Nm>"
        
        
        'Como cajamar da problemas, lo quitamos para todos
        'Print #NFic, "          <PstlAdr>"
        '
        'Cad = "ES"
        'If Not IsNull(miRsAux!PAIS) Then Cad = Mid(miRsAux!PAIS, 1, 2)
        'Print #NFic, "              <Ctry>" & Cad & "</Ctry>"
        '
        'If Not IsNull(miRsAux!dirdatos) Then Print #NFic, "              <AdrLine>" & XML(miRsAux!dirdatos) & "</AdrLine>"
        'Cad = XML(Trim(DBLet(miRsAux!codposta, "T") & " " & DBLet(miRsAux!despobla, "T")))
        'If Cad <> "" Then Print #NFic, "              <AdrLine>" & Cad & "</AdrLine>"
        'If Not IsNull(miRsAux!desprovi) Then Print #NFic, "              <AdrLine>" & XML(miRsAux!desprovi) & "</AdrLine>"
        'Print #NFic, "           </PstlAdr>"
        
        
        
        Print #NFic, "           <Id>"
        Aux = "PrvtId"
        If EsPersonaJuridica2 Then Aux = "OrgId"
      
        Print #NFic, "               <" & Aux & ">"
        Print #NFic, "                  <Othr>"
        
        Print #NFic, "                     <Id>" & miRsAux!numdni & "</Id>"
        'Da problemas.... con Cajamar
        'Print #NFic, "                     <Issr>NIF</Issr>"
        Print #NFic, "                  </Othr>"
        Print #NFic, "               </" & Aux & ">"
        Print #NFic, "           </Id>"
        Print #NFic, "        </Cdtr>"
        Print #NFic, "        <CdtrAcct>"
        Print #NFic, "           <Id>"
        Print #NFic, "              <IBAN>" & IBAN_Destino(miRsAux) & "</IBAN>"
        Print #NFic, "           </Id>"
        Print #NFic, "        </CdtrAcct>"
        Print #NFic, "      <Purp>"
        
        
        If ConceptoTr = "1" Then
            Aux = "SALA"
        ElseIf ConceptoTr = "0" Then
            Aux = "PENS"
        Else
            Aux = "TRAD"
        End If
        
        Print #NFic, "         <Cd>" & Aux & "</Cd>"
        Print #NFic, "      </Purp>"
        Print #NFic, "      <RmtInf>"
        
        Aux = DBLet(miRsAux!concepto, "T") & " " & DBLet(Fecha, "T") & " Importe" & Format(Im, FormatoImporte)
        If Trim(Aux) = "" Then Aux = miRsAux!Nommacta
        Print #NFic, "         <Ustrd>" & XML(Trim(Aux)) & "</Ustrd>"
        Print #NFic, "      </RmtInf>"
        Print #NFic, "   </CdtTrfTxInf>"
 
       
    
            
        miRsAux.MoveNext
    Wend
    Print #NFic, "   </PmtInf>"
    Print #NFic, "</CstmrCdtTrfInitn></Document>"
    
    
    miRsAux.Close
    Set miRsAux = Nothing
    Close (NFic)
    NFic = -1
    If Regs > 0 Then GeneraFicheroNorma34SEPA_XML = True
    Exit Function
EGen3:
    MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    If NFic > 0 Then Close (NFic)
End Function



Private Function XML(CADENA As String) As String
Dim i As Integer
Dim Aux As String
Dim Le As String
Dim C As Integer
    'Carácter no permitido en XML  Representación ASCII
    '& (ampersand)          &amp;
    '< (menor que)          &lt;
    ' > (mayor que)         &gt;
    '“ (dobles comillas)    &quot;
    '' (apóstrofe)          &apos;
    
    'La ISO recomienda trabajar con los carcateres:
    'a b c d e f g h i j k l m n o p q r s t u v w x y z
    'A B C D E F G H I J K L M N O P Q R S T U V W X Y Z
    '0 1 2 3 4 5 6 7 8 9
    '/ - ? : ( ) . , ' +
    'Espacio
    Aux = ""
    For i = 1 To Len(CADENA)
        Le = Mid(CADENA, i, 1)
        C = Asc(Le)
        
        
        Select Case C
        Case 40 To 57
            'Caracteres permitidos y numeros
            
        Case 65 To 90
            'Letras mayusculas
            
        Case 97 To 122
            'Letras minusculas
            
        Case 32
            'espacio en balanco
            
        Case Else
            Le = " "
        End Select
        Aux = Aux & Le
    Next
    XML = Aux
End Function





