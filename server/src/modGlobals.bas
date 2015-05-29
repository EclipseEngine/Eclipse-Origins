Attribute VB_Name = "modGlobals"
Option Explicit
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const conSwNormal = 1

Public StartTime As Long

Public CharMode As Long

Public Credits As String
Public News As String

Public DebugMode As Boolean
Public ErrorCount As Long

' Used for closing key doors again
Public KeyTimer As Long

' Used for gradually giving back npcs hp
Public GiveNPCHPTimer As Long
Public GivePetHPTimer As Long

' Used for logging
Public ServerLog As Boolean

' Text vars
Public vbQuote As String

' Maximum classes
Public Max_Classes As Long

' Used for server loop
Public ServerOnline As Boolean

' Used for outputting text
Public NumLines As Long

' Used to handle shutting down server with countdown.
Public isShuttingDown As Boolean
Public shutDownType As Long
Public Secs As Long
Public TotalPlayersOnline As Long

' GameCPS
Public GameCPS As Long
Public ElapsedTime As Long

' high indexing
Public Player_HighIndex As Long

' lock the CPS?
Public CPSUnlock As Boolean

Public EditingPlayer As Long
Public EditInv(1 To MAX_INV) As PlayerInvRec
Public EditSpell(1 To MAX_PLAYER_SPELLS) As Long



