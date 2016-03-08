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
Dim Cad As String

    On Error GoTo EGen
    GeneraFicheroNorma34 = False
    
    
    NFich = FreeFile
    Open App.Path & "\norma34.txt" For Output As #NFich
    
    'Codigo ordenante
    CodigoOrdenante = Right("    " & CIF, 9)   'CIF EMPRESA
    
    'CABECERA
    Cabecera1 NFich, CodigoOrdenante, Fecha, CuentaPropia, Cad
    Cabecera2 NFich, CodigoOrdenante, Cad
    Cabecera3 NFich, CodigoOrdenante, Cad
    Cabecera4 NFich, CodigoOrdenante, Cad
    
    
    
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
        
            Linea1 NFich, Aux, RS, Cad, ConceptoTransferencia
            Linea2 NFich, Aux, RS, Cad
            Linea3 NFich, Aux, RS, Cad
            Linea4 NFich, Aux, RS, Cad
            Linea5 NFich, Aux, RS, Cad
            Linea6 NFich, Aux, RS, Cad
           
        
        
        
        
            Importe = Importe + RS!Importe
            Regs = Regs + 1
            RS.MoveNext
        Wend
        'Imprimimos totales
        Totales NFich, CodigoOrdenante, Importe, Regs, Cad
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
Dim Cad As String
    
    Cad = Space(Longitud)
    If PorLaDerecha Then
        Cad = CADENA & Cad
        RellenaABlancos = Left(Cad, Longitud)
    Else
        Cad = Cad & CADENA
        RellenaABlancos = Right(Cad, Longitud)
    End If
    
End Function



Private Function RellenaAceros(CADENA As String, PorLaDerecha As Boolean, Longitud As Integer) As String
Dim Cad As String
    
    Cad = Mid("00000000000000000000", 1, Longitud)
    If PorLaDerecha Then
        Cad = CADENA & Cad
        RellenaAceros = Left(Cad, Longitud)
    Else
        Cad = Cad & CADENA
        RellenaAceros = Right(Cad, Longitud)
    End If
    
End Function



'Private Sub Cabecera1(NF As Integer,ByRef CodOrde As String)
'Dim Cad As String
'
'End Sub

Private Sub Cabecera1(NF As Integer, ByRef CodOrde As String, Fecha As Date, Cta As String, ByRef Cad As String)

    Cad = "03"
    Cad = Cad & "56"
    Cad = Cad & " "
    Cad = Cad & CodOrde
    Cad = Cad & Space(12) & "001"
    Cad = Cad & Format(Now, "ddmmyy")
    Cad = Cad & Format(Fecha, "ddmmyy")
    'Cuenta bancaria
    Cad = Cad & RecuperaValor(Cta, 1)
    Cad = Cad & RecuperaValor(Cta, 2)
    Cad = Cad & RecuperaValor(Cta, 4)
    Cad = Cad & "0"  'Sin relacion
    Cad = Cad & "   " & RecuperaValor(Cta, 3)  'Digito de control bancario
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub



Private Sub Cabecera2(NF As Integer, ByRef CodOrde As String, ByRef Cad As String)
    Cad = "03"
    Cad = Cad & "56"
    Cad = Cad & " "
    Cad = Cad & CodOrde
    Cad = Cad & Space(12) & "002"
    
    Cad = Cad & RellenaABlancos(vEmpresa.NomEmpresa, True, 30)  'Nombre empresa
  
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Cabecera3(NF As Integer, ByRef CodOrde As String, ByRef Cad As String)
    Cad = "03"
    Cad = Cad & "56"
    Cad = Cad & " "
    Cad = Cad & CodOrde
    Cad = Cad & Space(12) & "003"
    
    Cad = Cad & RellenaABlancos(vEmpresa.DirEmpresa, True, 30)   'Nombre empresa
  
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub



Private Sub Cabecera4(NF As Integer, ByRef CodOrde As String, ByRef Cad As String)

    Cad = "03"
    Cad = Cad & "56"
    Cad = Cad & " "
    Cad = Cad & CodOrde
    Cad = Cad & Space(12) & "004"
    
    Cad = Cad & RellenaABlancos(vEmpresa.CodPosEmpresa, False, 5)
    Cad = Cad & " "
    Cad = Cad & RellenaABlancos(vEmpresa.PobEmpresa, True, 30)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub



'ConceptoTransferencia
'1.- Abono nomina
'9.- Transferencia ordinaria
Private Sub Linea1(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Cad As String, vConceptoTransferencia As String)

    Cad = CodOrde   'llevara tb la ID del socio
    Cad = Cad & "010"
    Cad = Cad & RellenaAceros(CStr(Round(RS1!Importe, 2) * 100), False, 12)
    
    Cad = Cad & RellenaABlancos(RS1!banco1, False, 4)    'Entidad
    Cad = Cad & RellenaABlancos(RS1!banco2, False, 4)   'Sucur
    Cad = Cad & RellenaABlancos(RS1!banco4, False, 10)  'Cta
    Cad = Cad & "1" & vConceptoTransferencia
    Cad = Cad & "  "
    Cad = Cad & RellenaABlancos(RS1!banco3, False, 2)  'Cta
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Linea2(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Cad As String)
    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "011"
    Cad = Cad & RellenaABlancos(RS1!Nombre, False, 36)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Linea3(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Cad As String)
    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "012"
    Cad = Cad & RellenaABlancos(RS1!Domicilio, False, 36)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Linea4(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Cad As String)
    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "013"
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Linea5(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Cad As String)
    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "014"
    Cad = Cad & RellenaABlancos(RS1!codpos, False, 5) & " "
    Cad = Cad & RellenaABlancos(RS1!Poblacion, False, 30)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Linea6(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Cad As String)
    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "016"
    Cad = Cad & RellenaABlancos(RS1!concepto, False, 35)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Totales(NF As Integer, ByRef CodOrde As String, Total As Currency, Registros As Integer, ByRef Cad As String)
    Cad = "08" & "56 "
    Cad = Cad & CodOrde    'llevara tb la ID del socio
    Cad = Cad & Space(15)
    Cad = Cad & RellenaAceros(CStr(Int(Round(Total * 100, 2))), False, 12)
    Cad = Cad & RellenaAceros(CStr(Registros), False, 8)
    Cad = Cad & RellenaAceros(CStr((Registros * 6) + 4 + 1), False, 10)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
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
Dim Cad As String
Dim Aux As String

Dim miRsAux As ADODB.Recordset
Dim NF As Integer



    On Error GoTo EGen2
    GenFichNorma34SEPA = False
    

    
    
    'Cargamos la cuenta
 
    Set miRsAux = New ADODB.Recordset

    Cad = RecuperaValor(CuentaPropia2, 5) & RecuperaValor(CuentaPropia2, 1) & RecuperaValor(CuentaPropia2, 2) & RecuperaValor(CuentaPropia2, 3) & RecuperaValor(CuentaPropia2, 4)
    CuentaPropia2 = Cad
  
    If Len(Cad) <> 24 Then
        MsgBox "Error IBAN banco : " & CuentaPropia2, vbExclamation
        Exit Function
    End If
    
    NF = FreeFile
    Open App.Path & "\norma34.txt" For Output As #NF
    
    
    
    'SEPA
    '1.- Cabecera ordenante
    '------------------------------------------------------------------------
    Cad = "01" & "ORD" & "34145" & "001" & CIF
        
    'sufijo (Tenemos el OEM, que se utiliza para las otras normas antiguas
    Cad = Cad & SufijoOEM
    Cad = Cad & Format(Now, "yyyymmdd")
    Cad = Cad & Format(Fecha, "yyyymmdd")
    Cad = Cad & "A" 'IBAN
     
    'EL IBAN propiamente
    Cad = Cad & FrmtStr(CuentaPropia2, 34)
    Cad = Cad & "1" 'Cargo por cada operacion
    'Nombre
   
    Cad = Cad & FrmtStr(vEmpresa.NomEmpresa, 70)
    
    Cad = Cad & FrmtStr(Trim(vEmpresa.DirEmpresa), 50)
    Cad = Cad & FrmtStr(Trim(vEmpresa.CodPosEmpresa & " " & vEmpresa.PobEmpresa), 50)
    Cad = Cad & FrmtStr(DBLet(vEmpresa.ProvEmpresa, "T"), 40)
    
    'Pais y libre
    Cad = Cad & "ES" & FrmtStr("", 311)
    Print #NF, Cad
  
  
  
    '2.- Registro cabecera TRANSFERENCIA
    '------------------------------------------------------------------------
    Cad = "02" & "SCT" & "34145" & CIF
        
    'sufijo (Tenemos el OEM, que se utiliza para las otras normas
    Cad = Cad & SufijoOEM
    Cad = Cad & FrmtStr("", 578)
    Print #NF, Cad
    
    
    
    
    Cad = "SELECT tmpNorma34.*, Trabajadores.*, sbic.bic"
    Cad = Cad & " FROM (tmpNorma34 INNER JOIN Trabajadores ON tmpNorma34.CodSoc = Trabajadores.IdTrabajador)"
    Cad = Cad & " LEFT JOIN sbic ON Trabajadores.entidad = sbic.entidad;"
    
    
    
    miRsAux.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    If Not miRsAux.EOF Then
        Cad = "#"
        While Not miRsAux.EOF
            If IsNull(miRsAux!BIC) Then
                If InStr(1, Cad, "#" & miRsAux!banco1 & "#") = 0 Then Cad = Cad & miRsAux!banco1 & "#"
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.MoveFirst
        
        
        If Len(Cad) > 1 Then
            Cad = Mid(Cad, 2)
            Cad = Mid(Cad, 1, Len(Cad) - 1)
            Cad = Replace(Cad, "#", "   /   ")
            Cad = "Bancos sin BIC asignado:" & vbCrLf & Cad & vbCrLf & vbCrLf & "¿Continuar?"
            If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then
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
            Cad = "03" & "SCT" & "34145" & "002"
            
            'Campo 5 . Referencia del ordenante
            If IsNull(miRsAux!numdni) Then
                Aux = miRsAux!codsoc
            Else
                Aux = miRsAux!numdni
            End If
                
         
            Aux = FrmtStr(Aux, 10)
            Aux = DBLet(miRsAux!concepto, "T") & " Tra:" & Format(miRsAux!codsoc, "0000") & " F:" & Format(Fecha, "dd/mm")
            
            Cad = Cad & FrmtStr(Aux, 35)
            
            'Campo 6
            Cad = Cad & "A"
            
            'IBAN
            Cad = Cad & FrmtStr(IBAN_Destino(miRsAux), 34)
            
            
            
            'Campo8 Importe
            Cad = Cad & Format(Im * 100, String(11, "0")) ' Importe
            
            'Campo9
            Cad = Cad & "3" 'gastos compartidos
            'Campo 10
            Cad = Cad & FrmtStr(DBLet(miRsAux!BIC, "T"), 11) 'BIC

            'nommacta,dirdatos,codposta,dirdatos,despobla,impvenci,scobro.codmacta
            'Datos Basicos del beneficiario
            Cad = Cad & DatosBasicosDelDeudor(miRsAux)
            
            'Campo16 ID del pago. Concepto
            
            Aux = DBLet(miRsAux!concepto, "T") & " " & DBLet(Fecha, "T") & " Importe " & Format(Im, FormatoImporte)
          
            Cad = Cad & FrmtStr(Aux, 140)
            
            'Campo17
            Cad = Cad & FrmtStr("", 35)  'Reservado
            
            'Campo18  campo19
            
           
            
            If ConceptoTr = "1" Then
                Cad = Cad & "SALASALA"
            ElseIf ConceptoTr = "0" Then
                Cad = Cad & "PENSPENS"
            Else
                Cad = Cad & "TRADTRAD"
            End If
            
           
            
            Cad = Cad & FrmtStr("", 99)  'libre
            
            Print #NF, Cad
            
            miRsAux.MoveNext
        Wend
        
    
        'TOTALES
        '----------------------------------
        'Total trasnferencia SEPA
        'Campo 1,2
        Cad = "04" & "SCT"
        
        'Campo3 Importe total
        Cad = Cad & Format(Importe * 100, String(17, "0")) ' Importe
        Cad = Cad & Format(Regs, String(8, "0")) ' Importe
        'Total registros son
        'Reg(numreo de adeudos + 1 reg01 + un reg02 + reg04
        Cad = Cad & Format(Regs + 2, String(10, "0")) ' Importe   '2014-01-29  HABIA un reg + 3
        Cad = Cad & FrmtStr("", 560)  'libre
        Print #NF, Cad
        
        'Total general
        Cad = "99" & "ORD"
        
        'Campo3 Importe total
        Cad = Cad & Format(Importe * 100, String(17, "0")) ' Importe
        Cad = Cad & Format(Regs, String(8, "0")) ' Importe
        
        'Igual que arriba as uno
        'Reg(numreo de adeudos + 1 reg01 + un reg02 + reg04  +1
        Cad = Cad & Format(Regs + 4, String(10, "0")) ' Importe
        Cad = Cad & FrmtStr("", 560)  'libre
        Print #NF, Cad
        
        
        
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
Dim Cad As String
Dim Aux As String
Dim NFic As Integer
Dim EsPersonaJuridica2 As Boolean
Dim miRsAux As ADODB.Recordset

    On Error GoTo EGen3
    GeneraFicheroNorma34SEPA_XML = False
    
    NFic = -1
    
    
    Set miRsAux = New ADODB.Recordset

    Cad = RecuperaValor(CuentaPropia2, 5) & RecuperaValor(CuentaPropia2, 1) & RecuperaValor(CuentaPropia2, 2) & RecuperaValor(CuentaPropia2, 3) & RecuperaValor(CuentaPropia2, 4)
    CuentaPropia2 = Cad
  
    If Len(Cad) <> 24 Then
        MsgBox "Error IBAN banco : " & CuentaPropia2, vbExclamation
        Exit Function
    End If
    
    'Esta comprobacion deberia hacerla antes
    
    Cad = "SELECT tmpNorma34.CodSoc, tmpNorma34.Nombre, tmpNorma34.Banco1, tmpNorma34.Banco2, tmpNorma34.Banco3"
    Cad = Cad & ", tmpNorma34.Banco4, tmpNorma34.Domicilio, tmpNorma34.Codpos, tmpNorma34.Poblacion, tmpNorma34.Concepto,"
    Cad = Cad & "tmpNorma34.Importe, tmpNorma34.tipo"
    
    Cad = Cad & ",Trabajadores.*, sbic.bic"
    Cad = Cad & " FROM (tmpNorma34 INNER JOIN Trabajadores ON tmpNorma34.CodSoc = Trabajadores.IdTrabajador)"
    Cad = Cad & " LEFT JOIN sbic ON Trabajadores.entidad = sbic.entidad;"
    miRsAux.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    If Not miRsAux.EOF Then
        Cad = "#"
        While Not miRsAux.EOF
            If IsNull(miRsAux!BIC) Then
                If InStr(1, Cad, "#" & miRsAux!banco1 & "#") = 0 Then Cad = Cad & miRsAux!banco1 & "#"
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.MoveFirst
        
        
        If Len(Cad) > 1 Then
            Cad = Mid(Cad, 2)
            Cad = Mid(Cad, 1, Len(Cad) - 1)
            Cad = Replace(Cad, "#", "   /   ")
            Cad = "Bancos sin BIC asignado:" & vbCrLf & Cad & vbCrLf & vbCrLf & "¿Continuar?"
            If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then
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
    Cad = "TRANPAG" & Format(0, "000000") & "F" & Format(Now, "yyyymmddThhnnss")
    Print #NFic, "      <MsgId>" & Cad & "</MsgId>"
    Print #NFic, "      <CreDtTm>" & Format(Now, "yyyy-mm-ddThh:nn:ss") & "</CreDtTm>"
    
    'Registrp cabecera con totales
    
    Aux = "importe"
    Cad = "tmpNorma34"

    Cad = "Select count(*),sum(" & Aux & ") FROM " & Cad & " WHERE 1 =1"
    Aux = "0|0|"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(1)) Then Aux = miRsAux.Fields(0) & "|" & miRsAux.Fields(1) & "|"
    End If
    miRsAux.Close
    
    
    
    Print #NFic, "      <NbOfTxs>" & RecuperaValor(Aux, 1) & "</NbOfTxs>"
    Print #NFic, "      <CtrlSum>" & TransformaComasPuntos(RecuperaValor(Aux, 2)) & "</CtrlSum>"
    Print #NFic, "      <InitgPty>"
    Print #NFic, "         <Nm>" & XML(vEmpresa.NomEmpresa) & "</Nm>"
    Print #NFic, "         <Id>"
    Cad = Mid(CIF, 1, 1)
    
    EsPersonaJuridica2 = Not IsNumeric(Cad)

    
    
    
    Cad = "PrvtId"
    If EsPersonaJuridica2 Then Cad = "OrgId"
    
    Print #NFic, "           <" & Cad & ">"
    Print #NFic, "               <Othr>"
    Print #NFic, "                  <Id>" & CIF & SufijoOEM & "</Id>"
    Print #NFic, "               </Othr>"
    Print #NFic, "           </" & Cad & ">"
    
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

    Cad = vEmpresa.DirEmpresa & " "
    Cad = Cad & Trim(vEmpresa.PobEmpresa) & " " & vEmpresa.ProvEmpresa & " "
   
    Print #NFic, "            <AdrLine>" & XML(Trim(Cad)) & "</AdrLine>"
    
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
    
    Cad = Mid(CuentaPropia2, 5, 4)
    Cad = DevuelveDesdeBD("bic", "sbic", "entidad", Cad, "T")
    Print #NFic, "          <BIC>" & Trim(Cad) & "</BIC>"
    Print #NFic, "       </FinInstnId>"
    Print #NFic, "    </DbtrAgt>"
    
    
    
    
    Cad = "SELECT tmpNorma34.CodSoc, tmpNorma34.Nombre, tmpNorma34.Banco1, tmpNorma34.Banco2, tmpNorma34.Banco3"
    Cad = Cad & ", tmpNorma34.Banco4, tmpNorma34.Domicilio, tmpNorma34.Codpos, tmpNorma34.Poblacion, tmpNorma34.Concepto,"
    Cad = Cad & "tmpNorma34.Importe, tmpNorma34.tipo,Trabajadores.*, sbic.bic"
    Cad = Cad & " FROM (tmpNorma34 INNER JOIN Trabajadores ON tmpNorma34.CodSoc = Trabajadores.IdTrabajador)"
    Cad = Cad & " LEFT JOIN sbic ON Trabajadores.entidad = sbic.entidad;"
        

    miRsAux.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
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
        Cad = DBLet(miRsAux!numdni, "T")
        Cad = Mid(Cad, 1, 1)
        EsPersonaJuridica2 = Not IsNumeric(Cad)
        'Como da problemas Cajamar, siempre ponemos Perosna juridica. Veremos
        EsPersonaJuridica2 = True
        
        
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
        Cad = DBLet(miRsAux!BIC, "T")
        If Cad = "" Then Err.Raise 513, , "No existe BIC" & miRsAux!Nombre & vbCrLf & "Entidad: " & miRsAux!Entidad
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





