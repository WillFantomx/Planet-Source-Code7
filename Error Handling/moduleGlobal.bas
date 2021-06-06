Attribute VB_Name = "moduleGlobal"
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

Private Type SYSTEMTIME
    wYear                   As Integer
    wMonth                  As Integer
    wDayOfWeek              As Integer
    wDay                    As Integer
    wHour                   As Integer
    wMinute                 As Integer
    wSecond                 As Integer
    wMilliseconds           As Integer
End Type


Private Type TIME_ZONE_INFORMATION
    Bias                    As Long
    StandardName(63)        As Byte
    StandardDate            As SYSTEMTIME
    StandardBias            As Long
    DaylightName(63)        As Byte
    DaylightDate            As SYSTEMTIME
    DaylightBias            As Long
End Type


Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
   (LpVersionInformation As OSVERSIONINFO) As Long
Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As _
   MEMORYSTATUS)
Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As _
   SYSTEM_INFO)

Private Const PROCESSOR_INTEL_386 = 386
Private Const PROCESSOR_INTEL_486 = 486
Private Const PROCESSOR_INTEL_PENTIUM = 586
Private Const PROCESSOR_MIPS_R4000 = 4000
Private Const PROCESSOR_ALPHA_21064 = 21064

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Const TIME_ZONE_ID_UNKNOWN = 0
Private Const TIME_ZONE_ID_STANDARD = 1
Private Const TIME_ZONE_ID_DAYLIGHT = 2

Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Public Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)


'Change below email address to corresponding email.
Const EmailTo = "youremailaddress@serviceprovider.com"


Public Function GetLocalDateTime() As String
  Dim TimeZoneInfo As TIME_ZONE_INFORMATION
  Dim currentBias As Long
  Dim currentLocaltime As SYSTEMTIME


  'Windows returns the inverse of the bias we need (East Coast StdTime is returned as +0500, we want -0500)
  If GetTimeZoneInformation(TimeZoneInfo) = TIME_ZONE_ID_DAYLIGHT Then
    currentBias = -(TimeZoneInfo.Bias + TimeZoneInfo.DaylightBias)
  Else
    currentBias = -(TimeZoneInfo.Bias + TimeZoneInfo.StandardBias)
  End If

  GetLocalTime currentLocaltime

  With currentLocaltime
    GetLocalDateTime = Format(.wDay, "00") & "/" & Format(.wMonth, "00") & "/" & Format$(.wYear, "0000") & _
                              " " & Format$(.wHour, "00") & ":" & Format(.wMinute, "00") & ":" & Format(.wSecond, "00")
  End With
End Function 'GetDateInUniversalFormat



    'Public Function GetDirectXVersion()
'    Dim lVersion As Long
'    Dim lpData As String
'    lpData = Space$(40)
'    lVersion = 40
'    Call RegQueryValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\DirectX", "Version", lpData, lVersion)
'    GetDirectXVersion = lpData
'End Function

Public Function GetOSVersion()
    ' Get operating system and version.
    Dim verinfo As OSVERSIONINFO
    Dim build As String, ver_major As String, ver_minor As String
    Dim ret As Long
    verinfo.dwOSVersionInfoSize = Len(verinfo)
    ret = GetVersionEx(verinfo)
    If ret = 0 Then
        MsgBox "Error Getting Version Information"
        End
    End If
    Select Case verinfo.dwPlatformId
        Case 0
            GetOSVersion = GetOSVersion & "Windows 32s "
        Case 1
            GetOSVersion = GetOSVersion & "Windows 95/98 "
        Case 2
            GetOSVersion = GetOSVersion & "Windows NT "
    End Select

    ver_major = verinfo.dwMajorVersion
    ver_minor = verinfo.dwMinorVersion
    build = verinfo.dwBuildNumber
    GetOSVersion = GetOSVersion & ver_major & "." & ver_minor
    GetOSVersion = GetOSVersion & " (Build " & build & ")"
End Function

Public Function GetCPUType()
    ' Get CPU type and operating mode.
    Dim sysinfo As SYSTEM_INFO
    GetSystemInfo sysinfo
    GetCPUType = GetCPUType
    Select Case sysinfo.dwProcessorType
        Case PROCESSOR_INTEL_386
            GetCPUType = GetCPUType & "Intel 386"
        Case PROCESSOR_INTEL_486
            GetCPUType = GetCPUType & "Intel 486"
        Case PROCESSOR_INTEL_PENTIUM
            GetCPUType = GetCPUType & "Intel Pentium"
        Case PROCESSOR_MIPS_R4000
            GetCPUType = GetCPUType & "MIPS R4000"
        Case PROCESSOR_ALPHA_21064
            GetCPUType = GetCPUType & "DEC Alpha 21064"
        Case Else
            GetCPUType = GetCPUType & "(unknown)"
    End Select
End Function

'MSDN Code   ID: Q161088
Public Function SendMessage(DisplayMsg As Boolean, Optional AttachmentPath)
    Dim objOutlook As Outlook.Application
    Dim objOutlookMsg As Outlook.MailItem
    Dim objOutlookRecip As Outlook.Recipient
    Dim objOutlookAttach As Outlook.Attachment

    ' Create the Outlook session.
    Set objOutlook = CreateObject("Outlook.Application")

    ' Create the message.
    Set objOutlookMsg = objOutlook.CreateItem(olMailItem)

    With objOutlookMsg
        ' Add the To recipient(s) to the message.
        Set objOutlookRecip = .Recipients.Add(EmailTo)
        objOutlookRecip.Type = olTo

        ' Add the CC recipient(s) to the message.
        'Set objOutlookRecip = .Recipients.Add("youremailaddress@yahoo.com")
        'objOutlookRecip.Type = olCC

        ' Add the BCC recipient(s) to the message.
        'Set objOutlookRecip = .Recipients.Add("if you want any BCC")
        'objOutlookRecip.Type = olBCC

       ' Set the Subject, Body, and Importance of the message.
       .Subject = "Error Log"
       .Body = "Find attachment for error log. " & vbCrLf & vbCrLf
       .Importance = olImportanceHigh  'High importance

       ' Add attachments to the message.
       If Not IsMissing(AttachmentPath) Then
           Set objOutlookAttach = .Attachments.Add(AttachmentPath)
       End If

       ' Resolve each Recipient's name.
       For Each objOutlookRecip In .Recipients
           objOutlookRecip.Resolve
       Next

       ' Should we display the message before sending?
       If DisplayMsg Then
           .Display
       Else
           .Save
           .Send
       End If
    End With
    Set objOutlook = Nothing
End Function

