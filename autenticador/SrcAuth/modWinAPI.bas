Attribute VB_Name = "modWinAPI"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''
' HIGH RESOLUTION TIMER
Private Declare Sub GetSystemTime Lib "kernel32.dll" Alias "GetSystemTimeAsFileTime" (ByRef lpSystemTimeAsFileTime As Currency)
Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private GetSystemTimeOffset As Currency
Public Tick As Currency
'''''''''''''''''''''''''''''''''''''''''''''''

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
