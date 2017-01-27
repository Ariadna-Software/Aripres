VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmVisReport 
   Caption         =   "Visor de informes"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8430
   Icon            =   "frmVisReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   8430
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   5415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8055
      lastProp        =   600
      _cx             =   14208
      _cy             =   9551
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "frmVisReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Informe As String

'estas varriables las trae del formulario de impresion
Public FormulaSeleccion As String
Public SoloImprimir As Boolean
Public OtrosParametros As String   ' El grupo acaba en |                            ' param1=valor1|param2=valor2|
Public NumeroParametros As Integer   'Cuantos parametros hay.  EMPRESA(EMP) no es parametro. Es fijo en todos los informes
Public MostrarTree As Boolean

Public ExportarPDF As Boolean


Public ConSubinforme As Boolean


Dim mapp As CRAXDRT.Application
Dim mrpt As CRAXDRT.Report
Dim Argumentos() As String
Dim PrimeraVez As Boolean
Dim HayQueCerrar As Boolean


Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)

    
      UseDefault = False
     
      If mrpt.PrinterSetupEx(Me.Hwnd) = 0 Then
         
         mrpt.PrintOut False
         
     
     End If
    
    
End Sub

Private Sub Form_Activate()

    If PrimeraVez Then
        PrimeraVez = False
        If SoloImprimir Or Me.ExportarPDF Then
            Screen.MousePointer = vbHourglass
            Unload Me
            
        Else
            If HayQueCerrar Then Unload Me
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim I As Integer
Dim J As Integer
On Error GoTo Err_Carga

        Me.Icon = frmMain.Icon
    HayQueCerrar = False
    Screen.MousePointer = vbHourglass
    
    
    
    Set mapp = CreateObject("CrystalRuntime.Application")
    'Informe = "C:\Programas\Conta\Contabilidad\InformesD\sumas12.rpt"
    Set mrpt = mapp.OpenReport(Informe)



'Conectar a la BD de la Empresa

    For I = 1 To mrpt.Database.Tables.Count
       mrpt.Database.Tables(I).SetLogOnInfo "Aripres4", ValorBD          ', vConfig.User, vConfig.Password
'       If InStr(1, Right(mrpt.Database.Tables(i).Name, 2), "_") = 0 Then
       If InStr(1, mrpt.Database.Tables(I).Name, "_") = 0 Then
               mrpt.Database.Tables(I).Location = ValorBD & "." & mrpt.Database.Tables(I).Name
       ElseIf InStr(1, mrpt.Database.Tables(I).Name, "alias") <> 0 Then
            J = InStr(1, mrpt.Database.Tables(I).Name, "_")
            mrpt.Database.Tables(I).Location = ValorBD & "." & Mid(mrpt.Database.Tables(I).Name, 1, J - 1)
       End If

    Next I

    If ConSubinforme Then AbrirSubreport

    PrimeraVez = True
    CargaArgumentos
    CRViewer1.EnableGroupTree = MostrarTree
    CRViewer1.DisplayGroupTree = MostrarTree
    
    If FormulaSeleccion <> "" Then
        If mrpt.RecordSelectionFormula <> "" Then FormulaSeleccion = " AND " & FormulaSeleccion
        mrpt.RecordSelectionFormula = mrpt.RecordSelectionFormula & FormulaSeleccion
    End If
    'Si es a mail
    If Me.ExportarPDF Then
        Exportar
        Exit Sub
    End If
    
    
    'lOS MARGENES
    PonerMargen
    
    CRViewer1.ReportSource = mrpt
    If SoloImprimir Then
        mrpt.PrintOut False
    Else
        CRViewer1.ViewReport
    End If
    
    Exit Sub
Err_Carga:
    HayQueCerrar = True
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & Informe, vbCritical
    Set mapp = Nothing
    Set mrpt = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub CargaArgumentos()
Dim Parametro As String
Dim I As Integer
    'El primer parametro es el nombre de la empresa para todas las empresas
    ' Por lo tanto concaatenaremos con otros parametros
    ' Y sumaremos uno
    'Luego iremos recogiendo para cada formula su valor y viendo si esta en
    ' La cadena de parametros
    'Si esta asignaremos su valor
    
    OtrosParametros = "|Emp= """ & vEmpresa.NomEmpresa & """|" & OtrosParametros
    NumeroParametros = NumeroParametros + 1
    
    For I = 1 To mrpt.FormulaFields.Count
        Parametro = mrpt.FormulaFields(I).Name
        Parametro = Mid(Parametro, 3)  'Quitamos el {@
        Parametro = Mid(Parametro, 1, Len(Parametro) - 1) ' el } del final
        'Debug.Print Parametro
        If DevuelveValor(Parametro) Then mrpt.FormulaFields(I).Text = Parametro
        'Debug.Print " -- " & Parametro
    Next I
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrpt = Nothing
    Set mapp = Nothing
End Sub


Private Function DevuelveValor(ByRef valor As String) As Boolean
Dim I As Integer
Dim J As Integer
    valor = "|" & valor & "="
    DevuelveValor = False
    I = InStr(1, OtrosParametros, valor, vbTextCompare)
    If I > 0 Then
        I = I + Len(valor) + 1
        J = InStr(I, OtrosParametros, "|")
        If J > 0 Then
            valor = Mid(OtrosParametros, I, J - I)
            If valor = "" Then
                valor = " "
            Else
                'Si no tiene el salto
                If InStr(1, valor, "chr(13)") = 0 Then CompruebaComillas valor
            End If
            DevuelveValor = True
        End If
    End If
End Function


Private Sub CompruebaComillas(ByRef Valor1 As String)
Dim Aux As String
Dim J As Integer
Dim I As Integer

    If Mid(Valor1, 1, 1) = Chr(34) Then
        'Tiene comillas. Con lo cual tengo k poner las dobles
        Aux = Mid(Valor1, 2, Len(Valor1) - 2)
        I = -1
        Do
            J = I + 2
            I = InStr(J, Aux, """")
            If I > 0 Then
              Aux = Mid(Aux, 1, I - 1) & """" & Mid(Aux, I)
            End If
        Loop Until I = 0
        Aux = """" & Aux & """"
        Valor1 = Aux
    End If
End Sub

Private Sub Exportar()
    mrpt.ExportOptions.DiskFileName = App.Path & "\docum.pdf"
    mrpt.ExportOptions.DestinationType = crEDTDiskFile
    mrpt.ExportOptions.PDFExportAllPages = True
    mrpt.ExportOptions.FormatType = crEFTPortableDocFormat
    mrpt.Export False
    'Si ha generado bien entonces
    CadenaDesdeOtroForm = "OK"
End Sub

Private Sub PonerMargen()
Dim Cad As String
Dim I As Integer
    On Error GoTo EPon
    Cad = Dir(App.Path & "\*.mrg")
    If Cad <> "" Then
        I = InStr(1, Cad, ".")
        If I > 0 Then
            Cad = Mid(Cad, 1, I - 1)
            If IsNumeric(Cad) Then
                If Val(Cad) > 4000 Then Cad = "4000"
                If Val(Cad) > 0 Then
                    mrpt.BottomMargin = mrpt.BottomMargin + Val(Cad)
                End If
            End If
        End If
    End If
    
    Exit Sub
EPon:
    Err.Clear
End Sub




'======== LAURA ===============================================================

Private Sub AbrirSubreport()

'Para cada subReport que encuentre en el Informe pone las tablas del subReport

'apuntando a la BD correspondiente

Dim crxSection As CRAXDRT.Section

Dim crxObject As Object

Dim crxSubreportObject As CRAXDRT.SubreportObject
Dim smrpt
Dim I As Byte

 

    For Each crxSection In mrpt.Sections

        For Each crxObject In crxSection.ReportObjects

             If TypeOf crxObject Is SubreportObject Then

                Set crxSubreportObject = crxObject

                Set smrpt = mrpt.OpenSubreport(crxSubreportObject.SubreportName)

                For I = 1 To smrpt.Database.Tables.Count 'para cada tabla

                    '------ Añade Laura: 09/06/2005

 '                   If smrpt.Database.Tables(i).ConnectionProperties.Item("DSN") = "Aripres4" Then

                        smrpt.Database.Tables(I).SetLogOnInfo "Aripres4", ValorBD

                        If (InStr(1, smrpt.Database.Tables(I).Name, "_") = 0) Then

                           smrpt.Database.Tables(I).Location = ValorBD & "." & smrpt.Database.Tables(I).Name

                        End If

'                    ElseIf smrpt.Database.Tables(i).ConnectionProperties.Item("DSN") = "vConta" Then
'
'                        smrpt.Database.Tables(i).SetLogOnInfo "vConta", "conta" & vParamAplic.NumeroConta, vParamAplic.UsuarioConta, vParamAplic.PasswordConta
'
'                        If (InStr(1, smrpt.Database.Tables(i).Name, "_") = 0) Then
'
'                           smrpt.Database.Tables(i).Location = "conta" & vParamAplic.NumeroConta & "." & smrpt.Database.Tables(i).Name
'
'                        End If
'
'                    End If

                    '------

                Next I

             End If

        Next crxObject

    Next crxSection

    

    Set crxSubreportObject = Nothing

End Sub

