Attribute VB_Name = "modVars"
Option Explicit

Public NumLines As Long
Public Const MAX_LINES As Long = 100

Public Const ACCOUNT_LENGTH As Byte = 12
Public Const NAME_LENGTH As Long = 20

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const CLIENT_MAJOR As Byte = 1
Public Const CLIENT_MINOR As Byte = 8
Public Const CLIENT_REVISION As Byte = 0

Public Const MAX_BYTE As Byte = 255
Public Const MAX_INTEGER As Integer = 32767
Public Const MAX_LONG As Long = 2147483647

Public Const GAME_SERVER_IP As String = "127.0.0.1" ' "46.23.70.66"
Public Const GAME_SERVER_PORT As Long = 7001 ' the port used by the main game server
Public Const AUTH_SERVER_PORT As Long = 7002 ' the port used for people to connect to auth server
Public Const SERVER_AUTH_PORT As Long = 7003 ' the portal used for server to talk to auth server

Public Const GAME_NAME As String = "Crystalshire"
Public Const GAME_WEBSITE As String = "http://www.crystalshire.com"

Public Const MAX_PLAYERS As Byte = 50

' Tempo limite da conexao no sistema.
' Se uma conexao em que o token do login nao for confirmado.
' Tem uma duracao de 5 segundos para permancer no sistema.
Public Const MAX_CONNECTED_TIME As Long = 5000

' Codigo de autenticacao do cliente.
Public AuthCode As String

Public classMD5 As clsMD5
Public cryptography As clsCryptography

Public HandleDataSub(CMSG_COUNT) As Long

' dialogue alert strings
Public Const DIALOGUE_MSG_CONNECTION As Byte = 1
Public Const DIALOGUE_MSG_BANNED As Byte = 2
Public Const DIALOGUE_MSG_KICKED As Byte = 3
Public Const DIALOGUE_MSG_OUTDATED As Byte = 4
Public Const DIALOGUE_MSG_USERLENGTH As Byte = 5
Public Const DIALOGUE_MSG_ILLEGALNAME As Byte = 6
Public Const DIALOGUE_MSG_REBOOTING As Byte = 7
Public Const DIALOGUE_MSG_NAMETAKEN As Byte = 8
Public Const DIALOGUE_MSG_NAMELENGTH As Byte = 9
Public Const DIALOGUE_MSG_NAMEILLEGAL As Byte = 10
Public Const DIALOGUE_MSG_MYSQL As Byte = 11
Public Const DIALOGUE_MSG_WRONGPASS As Byte = 12
Public Const DIALOGUE_MSG_ACTIVATED As Byte = 13
