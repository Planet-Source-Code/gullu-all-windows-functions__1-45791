Attribute VB_Name = "Module1"
Public Declare Function RasEnumConnections Lib "RasApi32.dll" Alias "RasEnumConnectionsA" (lpRasCon As Any, lpcb As Long, lpcConnections As Long) As Long
Public Declare Function RasGetConnectStatus Lib "RasApi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As Any) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwReserved As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName$, ByVal lpdwReserved As Long, lpdwType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long
Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetDoubleClickTime& Lib "user32" ()
Declare Function GetTickCount& Lib "kernel32" ()
Declare Sub SHAddToRecentDocs Lib "shell32.dll" (ByVal uFlags As Long, ByVal pv As String)
Public Const SPI_SETDESKWALLPAPER = 20
Public Const MAX_PATH = 260
Public Const SPI_GETSCREENSAVEACTIVE = 16
Public Const RAS95_MaxEntryName = 256
Public Const RAS95_MaxDeviceType = 16
Public Const RAS95_MaxDeviceName = 32
Public Const SM_CMOUSEBUTTONS = 43
Private Const PHYSICALOFFSETX = 112
Public Const EXIT_LOGOFF = 0
Public Const EXIT_SHUTDOWN = 1
Public Const PROCESSOR_INTEL_386 = 386
Public Const PROCESSOR_INTEL_486 = 486
Public Const PROCESSOR_INTEL_PENTIUM = 586
Public Const PROCESSOR_MIPS_R4000 = 4000
Public Const PROCESSOR_ALPHA_21064 = 21064
Public Const EXIT_REBOOT = 2
Private Const PHYSICALOFFSETY = 113
Private Const PLANES = 14
Private Const BITSPIXEL = 12
Public Type RASCONN95
       dwSize As Long
       hRasCon As Long
       szEntryName(RAS95_MaxEntryName) As Byte
       szDeviceType(RAS95_MaxDeviceType) As Byte
       szDeviceName(RAS95_MaxDeviceName) As Byte
End Type
Public Type RASCONNSTATUS95
       dwSize As Long
       RasConnState As Long
       dwError As Long
       szDeviceType(RAS95_MaxDeviceType) As Byte
       szDeviceName(RAS95_MaxDeviceName) As Byte
End Type
Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type
Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
Function RegGetString$(hInKey As Long, ByVal subkey$, ByVal valname$)
Dim RetVal$, hSubKey As Long, dwType As Long, SZ As Long
    Dim R As Long
    RetVal$ = ""
    Const KEY_ALL_ACCESS As Long = &HF0063
    Const ERROR_SUCCESS As Long = 0
    Const REG_SZ As Long = 1
    R = RegOpenKeyEx(hInKey, subkey$, 0, KEY_ALL_ACCESS, hSubKey)
    If R <> ERROR_SUCCESS Then GoTo Quit_Now
    SZ = 256: v$ = String$(SZ, 0)
    R = RegQueryValueEx(hSubKey, valname$, 0, dwType, ByVal v$, SZ)
    If R = ERROR_SUCCESS And dwType = REG_SZ Then
        RetVal$ = Left$(v$, SZ - 1)
    Else
        RetVal$ = "--Not String--"
    End If
    If hInKey = 0 Then
        R = RegCloseKey(hSubKey)
    End If
Quit_Now:
        RegGetString$ = RetVal$
End Function
Function hideStartButton()
OurParent& = FindWindow("Shell_TrayWnd", "")
   OurHandle& = FindWindowEx(OurParent&, 0, "Button", vbNullString)
   ShowWindow OurHandle&, 0
End Function
Function showStartButton()
OurParent& = FindWindow("Shell_TrayWnd", "")
   OurHandle& = FindWindowEx(OurParent&, 0, "Button", vbNullString)
   ShowWindow OurHandle&, 5
End Function
Public Function IsModemConnected() As Boolean
    Dim TRasCon(255) As RASCONN95
    Dim lg As Long
    Dim lpcon As Long
    Dim RetVal As Long
    Dim Tstatus As RASCONNSTATUS95
    TRasCon(0).dwSize = 412
    lg = 256 * TRasCon(0).dwSize
    RetVal = RasEnumConnections(TRasCon(0), lg, lpcon)
    If RetVal <> 0 Then
         MsgBox "ERROR"
         Exit Function
    End If
    Tstatus.dwSize = 160
    RetVal = RasGetConnectStatus(TRasCon(0).hRasCon, Tstatus)
    If Tstatus.RasConnState = &H2000 Then
      IsModemConnected = True
     Else
      IsModemConnected = False
    End If
End Function
Sub SystemInformation()
    Dim msg As String
    Dim NewLine As String
    Dim ret As Integer
    Dim ver_major As Integer
    Dim ver_minor As Integer
    Dim Build As Long
    NewLine = Chr(13) + Chr(10)
    Dim verinfo As OSVERSIONINFO
    verinfo.dwOSVersionInfoSize = Len(verinfo)
    ret = GetVersionEx(verinfo)
    If ret = 0 Then
        MsgBox "Error Getting Version Information"
        End
    End If
    Select Case verinfo.dwPlatformId
        Case 0
            msg = msg + "Windows 32s "
        Case 1
            msg = msg + "Windows 95 "
        Case 2
            msg = msg + "Windows NT "
    End Select
    ver_major = verinfo.dwMajorVersion
    ver_minor = verinfo.dwMinorVersion
    Build = verinfo.dwBuildNumber
    msg = msg & ver_major & "." & ver_minor
    msg = msg & " (Build " & Build & ")" & NewLine & NewLine
    Dim sysinfo As SYSTEM_INFO
    GetSystemInfo sysinfo
    msg = msg + "CPU: "
    Select Case sysinfo.dwProcessorType
        Case PROCESSOR_INTEL_386
            msg = msg + "Intel 386" + NewLine
        Case PROCESSOR_INTEL_486
            msg = msg + "Intel 486" + NewLine
        Case PROCESSOR_INTEL_PENTIUM
            msg = msg + "Intel Pentium" + NewLine
        Case PROCESSOR_MIPS_R4000
            msg = msg + "MIPS R4000" + NewLine
        Case PROCESSOR_ALPHA_21064
            msg = msg + "DEC Alpha 21064" + NewLine
        Case Else
            msg = msg + "(unknown)" + NewLine
    End Select
    msg = msg + NewLine
    Dim memsts As MEMORYSTATUS
    Dim memory As Long
    GlobalMemoryStatus memsts
    memory = memsts.dwTotalPhys
    msg = msg + "Total Physical Memory: "
    msg = msg + Format(memory \ 1024, "###,###,###") + "K" + NewLine
    memory = memsts.dwAvailPhys
    msg = msg + "Available Physical Memory: "
    msg = msg + Format(memory \ 1024, "###,###,###") + "K" + NewLine
    memory = memsts.dwTotalVirtual
    msg = msg + "Total Virtual Memory: "
    msg = msg + Format(memory \ 1024, "###,###,###") + "K" + NewLine
    memory = memsts.dwAvailVirtual
    msg = msg + "Available Virtual Memory: "
    msg = msg + Format(memory \ 1024, "###,###,###") + "K" + NewLine
    MsgBox msg, vbOKOnly, "System Info"
End Sub

