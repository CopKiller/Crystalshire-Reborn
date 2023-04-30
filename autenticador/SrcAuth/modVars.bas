Attribute VB_Name = "modVars"
Option Explicit

Public isShuttingDown As Boolean

' Valores globais recebidos do servidor ao se conectar!
' Player
Public Const MAX_INV As Byte = 35
Public Const MAX_PLAYER_SPELLS As Byte = 35
Public Const MAX_HOTBAR As Byte = 12
' Player Database
Public Const MAX_QUESTS As Integer = 70
Public Const MAX_SERIAL_NUMBER As Integer = 100
Public Const MAX_BANK As Integer = 100

' Constantes que precisam ser alteradas aqui e no servidor!
Public Const MAX_PLAYERS As Byte = 50
Public Const GAME_NAME As String = "Crystalshire"
Public Const GAME_WEBSITE As String = "http://www.crystalshire.com"

' Sex constants
Public Const SEX_MALE As Byte = 0
Public Const SEX_FEMALE As Byte = 1

' Classes
Public Max_Classes As Byte

' Boolean constants
Public Const NO As Byte = 0
Public Const YES As Byte = 1

Public NumLines As Long
Public Const MAX_LINES As Long = 100

Public Const NAME_LENGTH As Byte = 20
Public Const ACCOUNT_LENGTH As Byte = 12
Public Const EMAIL_LENGTH As Byte = 25
Public Const DATA_LENGTH As Byte = 10 ' Count Characters in Date, Ex: 28/08/2022 <- have 10 characters (contagem pra não dar problema ao enviar a estrutura de dados pro servidor de save)
Public Const TASK_DEFEAT_LENGTH As Byte = 100

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const CLIENT_MAJOR As Byte = 1
Public Const CLIENT_MINOR As Byte = 8
Public Const CLIENT_REVISION As Byte = 0

Public Const MAX_BYTE As Byte = 255
Public Const MAX_INTEGER As Integer = 32767
Public Const MAX_LONG As Long = 2147483647

' Connection details
Public Const GAME_SERVER_IP As String = "127.0.0.1"    ' "46.23.70.66"
Public Const AUTH_SERVER_IP As String = "127.0.0.1"    ' "46.23.70.66"
Public Const EVENT_SERVER_IP As String = "127.0.0.1"    ' "46.23.70.66"
Public Const GAME_SERVER_PORT As Long = 7001    ' the port used by the main game server
Public Const AUTH_SERVER_PORT As Long = 7002    ' the port used for people to connect to auth server
Public Const SERVER_AUTH_PORT As Long = 7003    ' the portal used for server to talk to auth server
Public Const EVENT_SERVER_PORT As Long = 7004    ' the portal used for server to talk to auth server

' Tempo limite da conexao no sistema.
' Se uma conexao em que o token do login nao for confirmado.
' Tem uma duracao de 5 segundos para permancer no sistema.
Public Const MAX_CONNECTED_TIME As Long = 2000

' Codigo de autenticacao do cliente.
Public AuthCode As String

Public HandleDataSub(CMSG_COUNT) As Long

' dialogue alert strings
Public Const DIALOGUE_MSG_CONNECTION As Byte = 1
Public Const DIALOGUE_MSG_BANNED As Byte = 2
Public Const DIALOGUE_MSG_KICKED As Byte = 3
Public Const DIALOGUE_MSG_OUTDATED As Byte = 4
Public Const DIALOGUE_MSG_USERLENGTH As Byte = 5
Public Const DIALOGUE_MSG_USERILLEGAL As Byte = 6
Public Const DIALOGUE_MSG_REBOOTING As Byte = 7
Public Const DIALOGUE_MSG_NAMETAKEN As Byte = 8
Public Const DIALOGUE_MSG_NAMELENGTH As Byte = 9
Public Const DIALOGUE_MSG_NAMEILLEGAL As Byte = 10
Public Const DIALOGUE_MSG_WRONGPASS As Byte = 11
Public Const DIALOGUE_ACCOUNT_CREATED As Byte = 12
Public Const DIALOGUE_ACCOUNT_EMAILINVALID As Byte = 13
Public Const DIALOGUE_ACCOUNT_PASSLENGTH As Byte = 14
Public Const DIALOGUE_ACCOUNT_PASSNULL As Byte = 15
Public Const DIALOGUE_ACCOUNT_USERNULL As Byte = 16
Public Const DIALOGUE_ACCOUNT_PASSCONFIRM As Byte = 17
Public Const DIALOGUE_ACCOUNT_CAPTCHAINCORRECT As Byte = 18
Public Const DIALOGUE_SERIAL_INCORRECT As Byte = 19
Public Const DIALOGUE_SERIAL_CLAIMED As Byte = 20
Public Const DIALOGUE_BIRTHDAY_INCORRECT As Byte = 21
