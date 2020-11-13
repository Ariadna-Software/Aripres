VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPpal 
   Caption         =   "Form1"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command11 
      Caption         =   "Asignar calendario"
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Cosas"
      Height          =   495
      Left            =   7560
      TabIndex        =   5
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Ver calendario"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      Cancel          =   -1  'True
      Caption         =   "SALIR"
      Height          =   735
      Left            =   6840
      TabIndex        =   3
      Top             =   5880
      Width           =   3135
   End
   Begin MSComctlLib.ImageList imgListPpal 
      Left            =   6600
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun 
      Left            =   7560
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_BN 
      Left            =   7560
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_OM 
      Left            =   7560
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun16 
      Left            =   8520
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_BN16 
      Left            =   8520
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_OM16 
      Left            =   8520
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListImages16 
      Left            =   8520
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":0811
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":0DAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":17BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":21CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":3261
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgListComun32 
      Left            =   6600
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4560
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageListReloj 
      Left            =   5880
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":3C73
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":3F8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":42A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":45C1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "32x32"
      Height          =   255
      Index           =   2
      Left            =   6720
      TabIndex        =   2
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "24x24"
      Height          =   255
      Index           =   1
      Left            =   7560
      TabIndex        =   1
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "16x16"
      Height          =   255
      Index           =   0
      Left            =   8520
      TabIndex        =   0
      Top             =   1560
      Width           =   615
   End
End
Attribute VB_Name = "frmPpal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub GetIconsFromLibrary(ByVal sLibraryFilePath As String, ByVal op As Integer, ByVal tam As Integer)
    Dim i As Integer
    Dim tRes As ResType, iCount As Integer
        
    opcio = op
    tamany = tam
    ghmodule = LoadLibraryEx(sLibraryFilePath, 0, DONT_RESOLVE_DLL_REFERENCES)

    If ghmodule = 0 Then
        MsgBox "Invalid library file.", vbCritical
        Exit Sub
    End If
        
    For tRes = RT_FIRST To RT_LAST
        DoEvents
        EnumResourceNames ghmodule, tRes, AddressOf EnumResNameProc, 0
    Next
    FreeLibrary ghmodule
             
End Sub





Private Sub Form_Load()


    GetIconsFromLibrary App.Path & "\iconos.dll", 1, 32
'    GetIconsFromLibrary App.Path & "\iconosPTours.dll", 2, 32
    
    GetIconsFromLibrary App.Path & "\iconos.dll", 1, 24
    GetIconsFromLibrary App.Path & "\iconos_BN.dll", 2, 24
    GetIconsFromLibrary App.Path & "\iconos_OM.dll", 3, 24
    
    GetIconsFromLibrary App.Path & "\iconosPTours.dll", 4, 24
    
    GetIconsFromLibrary App.Path & "\iconos.dll", 1, 16
    GetIconsFromLibrary App.Path & "\iconos_BN.dll", 2, 16
    GetIconsFromLibrary App.Path & "\iconos_OM.dll", 3, 16
    

End Sub


