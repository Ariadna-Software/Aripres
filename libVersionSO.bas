Attribute VB_Name = "libVersionSO"
Option Explicit

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Const VER_NT_WORKSTATION = 1
Private Const VER_NT_DOMAIN_CONTROLLER = 2
Private Const VER_NT_SERVER = 3

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
    wServicePackMajor As Integer 'win2000 only
    wServicePackMinor As Integer 'win2000 only
    wSuiteMask As Integer 'win2000 only
    wProductType As Byte 'win2000 only
    wReserved As Byte
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFOEX) As Long

Private Function GetVersionInfo() As String
    Dim myOS As OSVERSIONINFOEX
    Dim bExInfo As Boolean
    Dim sOS As String

    myOS.dwOSVersionInfoSize = Len(myOS) 'should be 148/156
    'try win2000 version
    If GetVersionEx(myOS) = 0 Then
        'if fails
        myOS.dwOSVersionInfoSize = 148 'ignore reserved data
        If GetVersionEx(myOS) = 0 Then
            GetVersionInfo = "Microsoft Windows (Unknown)"
            Exit Function
        End If
    Else
        bExInfo = True
    End If
   
    With myOS
        'is version 4
        If .dwPlatformId = VER_PLATFORM_WIN32_NT Then
            'nt platform
            Select Case .dwMajorVersion
            Case 3, 4
                sOS = "Microsoft Windows NT"
            Case 5
                sOS = "Microsoft Windows 2000"
            End Select
            If bExInfo Then
                'workstation/server?
                If .wProductType = VER_NT_SERVER Then
                    sOS = sOS & " Server"
                ElseIf .wProductType = VER_NT_DOMAIN_CONTROLLER Then
                    sOS = sOS & " Domain Controller"
                ElseIf .wProductType = VER_NT_WORKSTATION Then
                    sOS = sOS & " Workstation"
                End If
            End If
           
            'get version/build no
            sOS = sOS & " Version " & .dwMajorVersion & "." & .dwMinorVersion & " " & StripTerminator(.szCSDVersion) & " (Build " & .dwBuildNumber & ")"
           
        ElseIf .dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
            'get minor version info
            If .dwMinorVersion = 0 Then
                sOS = "Microsoft Windows 95"
            ElseIf .dwMinorVersion = 10 Then
                sOS = "Microsoft Windows 98"
            ElseIf .dwMinorVersion = 90 Then
                sOS = "Microsoft Windows Millenium"
            Else
                sOS = "Microsoft Windows 9?"
            End If
            'get version/build no
            sOS = sOS & "Version " & .dwMajorVersion & "." & .dwMinorVersion & " " & StripTerminator(.szCSDVersion) & " (Build " & .dwBuildNumber & ")"
        End If
    End With
    GetVersionInfo = sOS
End Function
Private Function StripTerminator(sString As String) As String
    StripTerminator = Left$(sString, InStr(sString, Chr$(0)) - 1)
End Function



Public Function VersionSO() As String
    VersionSO = GetVersionInfo
End Function
