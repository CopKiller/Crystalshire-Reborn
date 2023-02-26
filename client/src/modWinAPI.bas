Attribute VB_Name = "modWinAPI"
Option Explicit

' API Declares
Public myHWnd As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
                                      (lpVersionInformation As OSVERSIONINFO) As Long

' ACCESS TOKEN RELATED JAZZ (Check for admin and stuff)
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Long, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Const TOKEN_READ As Long = &H20008    ' Used to read the token data
Private Const TOKEN_INFO_CLASS_TokenElevation As Long = 20    ' Used to check whether token is elevated or not

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'''''''''''''''''''''''''''''''''''''''''''''''
' HIGH RESOLUTION TIMER
Private Declare Sub GetSystemTime Lib "kernel32.dll" Alias "GetSystemTimeAsFileTime" (ByRef lpSystemTimeAsFileTime As Currency)
Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private GetSystemTimeOffset As Currency
'''''''''''''''''''''''''''''''''''''''''''''''

Private Type OSVERSIONINFO
    OSVSize As Long
    dwVerMajor As Long
    dwVerMinor As Long
    dwBuildNumber As Long
    PlatformID As Long
    szCSDVersion As String * 128
End Type

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Public Sub InitTime()
    
    ' Set the high-resolution timer
    timeBeginPeriod 1
    
    ' Get the initial time, time starting from this point will be calculated relative to this value
    GetSystemTime GetSystemTimeOffset

End Sub

Public Function getTime() As Currency

    ' The roll over still happens but the advantage is that you don't have to restart your pc, just restart the server
    ' This is getTimeCount starts counting from when your PC has started, but this method starts counting from when the server has started
    
    Dim CurrentTime As Currency

    ' Grab the current time (we have to pass a variable ByRef instead of a function return like the other timers)

    GetSystemTime CurrentTime

    ' Calculate the difference between the 64-bit times, return as a 32-bit time
    getTime = CurrentTime - GetSystemTimeOffset

End Function

Public Function GetWindowsVersion() As String
    Dim osv As OSVERSIONINFO
    osv.OSVSize = Len(osv)

    If GetVersionEx(osv) = 1 Then
        Select Case osv.PlatformID
        Case VER_PLATFORM_WIN32s
            GetWindowsVersion = "Win32s on Windows 3.1"
        Case VER_PLATFORM_WIN32_NT
            GetWindowsVersion = "Windows NT"

            Select Case osv.dwVerMajor
            Case 3
                GetWindowsVersion = "Windows NT 3.5"
            Case 4
                GetWindowsVersion = "Windows NT 4.0"
            Case 5
                Select Case osv.dwVerMinor
                Case 0
                    GetWindowsVersion = "Windows 2000"
                Case 1
                    GetWindowsVersion = "Windows XP"
                Case 2
                    GetWindowsVersion = "Windows Server 2003"
                End Select
            Case 6
                Select Case osv.dwVerMinor
                Case 0
                    GetWindowsVersion = "Windows Vista/Server 2008"
                Case 1
                    GetWindowsVersion = "Windows 7/Server 2008 R2"
                Case 2
                    GetWindowsVersion = "Windows 8 and 10"
                End Select
            End Select

        Case VER_PLATFORM_WIN32_WINDOWS:
            Select Case osv.dwVerMinor
            Case 0
                GetWindowsVersion = "Windows 95"
            Case 90
                GetWindowsVersion = "Windows Me"
            Case Else
                GetWindowsVersion = "Windows 98"
            End Select
        End Select
    Else
        GetWindowsVersion = "Não foi identificada a versão do windows!."
    End If
End Function

Public Function IsElevatedAccess() As Boolean
    Dim PID As Long, hToken As Long, Elevated As Long, ReturnLength As Long

    PID = GetCurrentProcess    ' = -1 always, it's a value that when passed to API maps to the current, correct PID

    If OpenProcessToken(PID, TOKEN_READ, hToken) Then
        Call GetTokenInformation(hToken, TOKEN_INFO_CLASS_TokenElevation, Elevated, LenB(Elevated), ReturnLength)
        IsElevatedAccess = Not (Elevated = 0)    ' if Elevated = 0 then not elevated
        Call CloseHandle(hToken)
    End If
End Function

Function SecondsToHMS(ByRef Segundos As Long) As String
    Dim HR As Long, ms As Long, SS As Long, MM As Long
    Dim Total As Long, Count As Long

    If Segundos = 0 Then Exit Function

    HR = (Segundos \ 3600)
    MM = (Segundos \ 60)
    SS = Segundos
    'ms = (Segundos * 10)

    ' Pega o total de segundos pra trabalharmos melhor na variavel!
    Total = Segundos

    ' Verifica se tem mais de 1 hora em segundos!
    If HR > 0 Then
        '// Horas
        Do While (Total >= 3600)
            Total = Total - 3600
            Count = Count + 1
        Loop
        If Count > 0 Then
            SecondsToHMS = Count & "h "
            Count = 0
        End If
        '// Minutos
        Do While (Total >= 60)
            Total = Total - 60
            Count = Count + 1
        Loop
        If Count > 0 Then
            SecondsToHMS = SecondsToHMS & Count & "m "
            Count = 0
        End If
        '// Segundos
        Do While (Total > 0)
            Total = Total - 1
            Count = Count + 1
        Loop
        If Count > 0 Then
            SecondsToHMS = SecondsToHMS & Count & "s "
            Count = 0
        End If
    ElseIf MM > 0 Then
        '// Minutos
        Do While (Total >= 60)
            Total = Total - 60
            Count = Count + 1
        Loop
        If Count > 0 Then
            SecondsToHMS = SecondsToHMS & Count & "m "
            Count = 0
        End If
        '// Segundos
        Do While (Total > 0)
            Total = Total - 1
            Count = Count + 1
        Loop
        If Count > 0 Then
            SecondsToHMS = SecondsToHMS & Count & "s "
            Count = 0
        End If
    ElseIf SS > 0 Then
        ' Joga na função esse segundo.
        SecondsToHMS = SS & "s "
        Total = Total - SS
    End If
End Function
