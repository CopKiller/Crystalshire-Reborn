Attribute VB_Name = "modGlobals"
Option Explicit

' Used for closing key doors again
Public KeyTimer As Currency

' Used for gradually giving back npcs hp
Public GiveNPCHPTimer As Currency

' Used for logging
Public ServerLog As Boolean

' Text vars
Public vbQuote As String

' Maximum classes
Public Max_Classes As Byte

' Used for server loop
Public ServerOnline As Boolean

' Used for outputting text
Public NumLines As Long

' Used to handle shutting down server with countdown.
Public isShuttingDown As Boolean
Public Secs As Long
Public TotalPlayersOnline As Long

' GameCPS
Public GameCPS As Long
Public ElapsedTime As Currency

' high indexing
Public Player_HighIndex As Long

' lock the CPS?
Public CPSUnlock As Boolean

' Game Time
Public GameSeconds As Byte
Public GameMinutes As Byte
Public GameHours As Byte
Public DayTime As Boolean
Public GameSecondsPerSecond As Byte
Public GameMinutesPerMinute As Byte
