Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'For Clear functions
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal errLine As Long)
Dim filename As String
Static fileOpen As Boolean
    filename = App.path & "\data\logs\errors.txt"
    If fileOpen = False Then
        fileOpen = True
        Open filename For Append As #1
            Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
            Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
            Print #1, "Error occured on line " & errLine
            Print #1, ""
        Close #1
        fileOpen = False
    End If
    ErrorCount = ErrorCount + 1
    UpdateCaption
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)

   On Error GoTo errorhandler

    If LCase$(Dir(tDir & tName, vbDirectory)) <> LCase(tName) Then Call MkDir(tDir & tName)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ChkDir", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' Outputs string to text file
Sub AddLog(ByVal Text As String, ByVal FN As String)
    Dim filename As String
    Dim F As Long


   On Error GoTo errorhandler

    If ServerLog Then
        filename = App.path & "\data\logs\" & FN

        If Not FileExist(filename, True) Then
            F = FreeFile
            Open filename For Output As #F
            Close #F
        End If

        F = FreeFile
        Open filename For Append As #F
        Print #F, Time & ": " & Text
        Close #F
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "AddLog", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' gets a string from a text file
Public Function GetVar(File As String, Header As String, ByVal Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    Dim Value As String

   On Error GoTo errorhandler

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
    GetVar = Replace(GetVar, CStr(Chr$(237)), vbCrLf, , , vbTextCompare)

   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetVar", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, ByVal Value As String)
Dim i As Long

   On Error GoTo errorhandler
   
    Value = Replace(Value, vbCrLf, CStr(Chr$(237)), , , vbTextCompare)
    Call WritePrivateProfileString$(Header, Var, Value, File)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PutVar", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Function FileExist(ByVal filename As String, Optional RAW As Boolean = False) As Boolean


   On Error GoTo errorhandler

    If Not RAW Then
        If LenB(Dir(App.path & "\" & filename)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(Dir(filename)) > 0 Then
            FileExist = True
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "FileExist", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Public Sub SaveOptions()
    

   On Error GoTo errorhandler

    PutVar App.path & "\options.ini", "OPTIONS", "Game_Name", Options.Game_Name
    PutVar App.path & "\options.ini", "OPTIONS", "Port", STR(Options.Port)
    PutVar App.path & "\options.ini", "OPTIONS", "MOTD", Options.MOTD
    PutVar App.path & "\options.ini", "OPTIONS", "Website", Options.Website
    PutVar App.path & "\options.ini", "OPTIONS", "KEY", Options.Key
    PutVar App.path & "\options.ini", "OPTIONS", "DataFolder", Options.DataFolder
    PutVar App.path & "\options.ini", "OPTIONS", "UpdateURL", Options.UpdateURL
    PutVar App.path & "\options.ini", "OPTIONS", "StaffOnly", STR(Options.StaffOnly)
    PutVar App.path & "\options.ini", "OPTIONS", "DisableRemoteRestart", STR(Options.DisableRemoteRestart)
    
    'New Stuff
    PutVar App.path & "\options.ini", "GameOptions", "NewCombat", STR(NewOptions.CombatMode)
    PutVar App.path & "\options.ini", "GameOptions", "MaxLevel", STR(NewOptions.MaxLevel)
    PutVar App.path & "\options.ini", "GameOptions", "MainMenuMusic", NewOptions.MainMenuMusic
    PutVar App.path & "\options.ini", "GameOptions", "ItemLoss", STR(NewOptions.ItemLoss)
    PutVar App.path & "\options.ini", "GameOptions", "ExpLoss", STR(NewOptions.ExpLoss)
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveOptions", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
    
End Sub

Public Sub LoadOptions()
Dim filepath As String, killfile As Boolean

   On Error GoTo errorhandler
    
    If FileExist(App.path & "\options.ini", True) = False Then
        If FileExist(App.path & "\data\options.ini", True) = True Then
            filepath = App.path & "\data\options.ini"
            killfile = True
        Else
            
        End If
    Else
        filepath = App.path & "\options.ini"
    End If

    Options.Game_Name = GetVar(filepath, "OPTIONS", "Game_Name")
    Options.Port = Val(GetVar(filepath, "OPTIONS", "Port"))
    Options.MOTD = GetVar(filepath, "OPTIONS", "MOTD")
    Options.Website = GetVar(filepath, "OPTIONS", "Website")
    Options.SilentStartup = Val(GetVar(filepath, "OPTIONS", "SilentStartup"))
    Options.Key = Trim$(GetVar(filepath, "OPTIONS", "KEY"))
    Options.DataFolder = Trim$(GetVar(filepath, "OPTIONS", "DataFolder"))
    Options.UpdateURL = Trim$(GetVar(filepath, "OPTIONS", "UpdateURL"))
    Options.StaffOnly = Val(GetVar(filepath, "OPTIONS", "StaffOnly"))
    Options.DisableRemoteRestart = Val(GetVar(filepath, "OPTIONS", "DisableRemoteRestart"))
    
    If Options.StaffOnly = 1 Then
        frmServer.chkStaffOnly.Value = 1
    End If
    
    If Options.DisableRemoteRestart = 1 Then
        frmServer.chkDisableRestart.Value = 1
    End If
    
    
    
    If Options.Key = "" Then Options.Key = GenerateOptionsKey: SaveOptions
    
    Dim iFileNumber As Integer
    Dim handle As Integer
    Dim filetext As String
    'Loading News and Credits Now too... easier this way :D.
    If FileExist(App.path & "\data\news.txt", True) Then
        handle = FreeFile
        Open App.path & "\data\news.txt" For Input As #handle
        filetext = Input$(LOF(handle), handle)
        Close #handle
        News = filetext
        frmServer.txtNews.Text = News
    Else
        News = ""
        iFileNumber = FreeFile
        Open App.path & "\data\news.txt" For Output As #iFileNumber
        Print #iFileNumber, News
        Close #iFileNumber
    End If
    
    If FileExist(App.path & "\data\credits.txt", True) Then
        handle = FreeFile
        Open App.path & "\data\credits.txt" For Input As #handle
        filetext = Input$(LOF(handle), handle)
        Close #handle
        Credits = filetext
        frmServer.txtCredits.Text = News
    Else
        Credits = ""
        iFileNumber = FreeFile
        Open App.path & "\data\credits.txt" For Output As #iFileNumber
        Print #iFileNumber, Credits
        Close #iFileNumber
    End If
    
    If Trim$(GetVar(filepath, "GameOptions", "NewCombat")) = "" Then
        NewOptions.CombatMode = 1
        NewOptions.MaxLevel = 100
        SaveOptions
    Else
        NewOptions.CombatMode = Val(GetVar(filepath, "GameOptions", "NewCombat"))
        If NewOptions.CombatMode <> 1 And NewOptions.CombatMode <> 2 Then NewOptions.CombatMode = 1: SaveOptions
        NewOptions.MaxLevel = Val(GetVar(filepath, "GameOptions", "MaxLevel"))
        If NewOptions.MaxLevel <= 0 Then NewOptions.MaxLevel = 100: SaveOptions
        NewOptions.MainMenuMusic = GetVar(filepath, "GameOptions", "MainMenuMusic")
        NewOptions.ItemLoss = Val(GetVar(filepath, "GameOptions", "ItemLoss"))
        NewOptions.ExpLoss = Val(GetVar(filepath, "GameOptions", "ExpLoss"))
    End If
    
    MAX_LEVELS = NewOptions.MaxLevel
    
    If killfile Then
        SaveOptions
        Kill filepath
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadOptions", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function Ban(ByVal Index As Long, ByVal accname As String, ByVal online As Boolean, Optional reason As String = "") As Boolean
    Dim filename As String
    Dim ip As String
    Dim F As Long
    Dim i As Long

   On Error GoTo errorhandler

    filename = App.path & "\data\banlist.bin"

    ' Make sure the file exists
    If Not FileExist("data\banlist.bin") Then
        F = FreeFile
        Open filename For Binary As #F
        Close #F
    End If
    
    If BanCount > 0 Then
        For i = 1 To BanCount
            If online Then
                If Trim$(Bans(i).BanName) = Trim$(Player(Index).login) Then
                    Exit Function
                End If
            Else
                If Trim$(Bans(i).BanName) = Trim$(accname) Then
                    Exit Function
                End If
            End If
        Next
    End If
    
    BanCount = BanCount + 1
    ReDim Preserve Bans(BanCount)
    
    With Bans(BanCount)
        If online Then
            accname = Trim$(Player(Index).login)
            .BanChar = Trim$(Player(Index).characters(TempPlayer(Index).CurChar).Name)
            .BanName = Trim$(Player(Index).login)
            .IPAddress = Trim$(GetPlayerIP(Index))
        Else
            .BanChar = "N/A"
            .BanName = Trim$(accname)
            i = FindAccount(accname)
            If i > 0 Then
                .IPAddress = Trim$(account(i).ip)
                If Trim$(.IPAddress) = "" Then
                    .IPAddress = "N/A"
                End If
            Else
                .IPAddress = "N/A"
            End If
        End If
        .BanReason = reason
    End With

    F = FreeFile
    Open filename For Binary As #F
    Put #F, , BanCount
    For i = 1 To BanCount
        Put #F, , Bans(i).BanName
        Put #F, , Bans(i).BanChar
        Put #F, , Bans(i).IPAddress
        Put #F, , Bans(i).BanReason
    Next
    Close #F
    
    Call GlobalMsg("Account " & accname & " has been banned from " & Options.Game_Name & "!", White)
    Call AddLog("Account " & accname & " has been banned!", ADMIN_LOG)
    If online Then
        Call AlertMsg(Index, "You have been banned!")
    End If
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerAccess(i) >= ADMIN_CREATOR Then
                SendAccounts i
                SendBans i
            End If
        End If
    Next
    
    Ban = True


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "BanIndex", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Function

' **************
' ** Accounts **
' **************
Function AccountExist(ByVal Name As String) As Boolean
    Dim filename As String

   On Error GoTo errorhandler

    filename = "data\accounts\" & Trim(Name) & "\" & Trim(Name) & ".bin"

    If FileExist(filename) Then
        AccountExist = True
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "AccountExist", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
    Dim filename As String
    Dim RightPassword As String * NAME_LENGTH
    Dim nFileNum As Long


   On Error GoTo errorhandler

    If AccountExist(Name) Then
        filename = App.path & "\data\accounts\" & Trim$(Name) & "\" & Trim$(Name) & ".bin"
        nFileNum = FreeFile
        Open filename For Binary As #nFileNum
        Get #nFileNum, ACCOUNT_LENGTH, RightPassword
        Close #nFileNum

        If UCase$(Trim$(Password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "PasswordOK", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub AddAccount(ByVal Index As Long, ByVal Name As String, ByVal Password As String)
    Dim i As Long
    

   On Error GoTo errorhandler

    ClearPlayer Index
    
    Player(Index).login = Name
    Player(Index).Password = Password

    Call SavePlayer(Index)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "AddAccount", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub DeleteName(ByVal Name As String)
    Dim f1 As Long
    Dim f2 As Long
    Dim s As String

   On Error GoTo errorhandler

    Call FileCopy(App.path & "\data\accounts\charlist.txt", App.path & "\data\accounts\chartemp.txt")
    ' Destroy name from charlist
    f1 = FreeFile
    Open App.path & "\data\accounts\chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.path & "\data\accounts\charlist.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, s

        If Trim$(LCase$(s)) <> Trim$(LCase$(Name)) Then
            Print #f2, s
        End If

    Loop

    Close #f1
    Close #f2
    Call Kill(App.path & "\data\accounts\chartemp.txt")


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DeleteName", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ****************
' ** Characters **
' ****************
Function CharExist(ByVal Index As Long, ByVal CharNum As Long) As Boolean


   On Error GoTo errorhandler

    If LenB(Trim$(Player(Index).characters(CharNum).Name)) > 0 Then
        CharExist = True
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CharExist", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub AddChar(ByVal Index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Long, ByVal Sprite As Long, CharNum As Long)
    Dim F As Long
    Dim n As Long
    Dim spritecheck As Boolean


   On Error GoTo errorhandler

    If LenB(Trim$(Player(Index).characters(CharNum).Name)) = 0 Then
        
        spritecheck = False
        
        Player(Index).characters(CharNum).Name = Name
        Player(Index).characters(CharNum).Sex = Sex
        Player(Index).characters(CharNum).Class = ClassNum

        Player(Index).characters(CharNum).Level = 1

        For n = 1 To Stats.Stat_Count - 1
            Player(Index).characters(CharNum).stat(n) = Class(ClassNum).stat(n)
        Next n

        Player(Index).characters(CharNum).Dir = DIR_DOWN
        Player(Index).characters(CharNum).Map = START_MAP
        Player(Index).characters(CharNum).x = START_X
        Player(Index).characters(CharNum).y = START_Y
        Player(Index).characters(CharNum).Dir = DIR_DOWN
        Player(Index).characters(CharNum).Vital(Vitals.HP) = GetPlayerMaxVital(Index, Vitals.HP)
        Player(Index).characters(CharNum).Vital(Vitals.MP) = GetPlayerMaxVital(Index, Vitals.MP)
        
        ' set starter equipment
        If Class(ClassNum).startItemCount > 0 Then
            For n = 1 To Class(ClassNum).startItemCount
                If Class(ClassNum).StartItem(n) > 0 Then
                    ' item exist?
                    If Len(Trim$(Item(Class(ClassNum).StartItem(n)).Name)) > 0 Then
                        Player(Index).characters(CharNum).Inv(n).Num = Class(ClassNum).StartItem(n)
                        Player(Index).characters(CharNum).Inv(n).Value = Class(ClassNum).StartValue(n)
                    End If
                End If
            Next
        End If
        
        ' set start spells
        If Class(ClassNum).startSpellCount > 0 Then
            For n = 1 To Class(ClassNum).startSpellCount
                If Class(ClassNum).StartSpell(n) > 0 Then
                    ' spell exist?
                    If Len(Trim$(Spell(Class(ClassNum).StartSpell(n)).Name)) > 0 Then
                        Player(Index).characters(CharNum).Spell(n) = Class(ClassNum).StartSpell(n)
                    End If
                End If
            Next
        End If
        
        ' Append name to file
        F = FreeFile
        Open App.path & "\data\accounts\charlist.txt" For Append As #F
        Print #F, Name
        Close #F
        Call SavePlayer(Index)
        Exit Sub
        
        account(FindAccount(Player(Index).login)).characters(CharNum) = Trim$(Name)
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "AddChar", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function FindChar(ByVal Name As String, ByRef charcount As Long) As Boolean
    Dim F As Long
    Dim s As String

   On Error GoTo errorhandler

    F = FreeFile
    Open App.path & "\data\accounts\charlist.txt" For Input As #F
    charcount = 0
    Do While Not EOF(F)
        Input #F, s
        If Trim$(LCase(s)) <> "" Then
            charcount = charcount + 1
            If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
                FindChar = True
                Close #F
                Exit Function
            End If
        End If

    Loop

    Close #F


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "FindChar", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Function

' *************
' ** Players **
' *************
Sub SaveAllPlayersOnline()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            Call SavePlayer(i)
            Call SaveBank(i)
        End If

    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveAllPlayersOnline", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SavePlayer(ByVal Index As Long)
    Dim filename As String, i As Long
    Dim F As Long
    

   On Error GoTo errorhandler

    ChkDir App.path & "\data\accounts\", Trim$(Player(Index).login)

    filename = App.path & "\data\accounts\" & Trim$(Player(Index).login) & "\" & Trim$(Player(Index).login) & ".bin"
    
    F = FreeFile
    
    Open filename For Binary As #F
    Put #F, , Player(Index).login
    Put #F, , Player(Index).Password
    Put #F, , Player(Index).ip
    Close #F
    
    For i = 1 To MAX_PLAYER_CHARS
        filename = App.path & "\data\accounts\" & Trim$(Player(Index).login) & "\" & Trim$(Player(Index).login) & "_char" & CStr(i) & ".bin"
        F = FreeFile
        
        Open filename For Binary As #F
        Put #F, , Player(Index).characters(i)
        Close #F
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SavePlayer", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
    
End Sub

Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
    Dim filename As String, i As Long
    Dim F As Long

   On Error GoTo errorhandler

    Call ClearPlayer(Index)
    filename = App.path & "\data\accounts\" & Trim(Name) & "\" & Trim$(Name) & ".bin"
    F = FreeFile
    Open filename For Binary As #F
    Get #F, , Player(Index).login
    Get #F, , Player(Index).Password
    Player(Index).ip = Trim$(GetPlayerIP(Index))
    If AccountCount > 0 Then
        account(FindAccount(Trim$(Player(Index).login))).ip = Trim$(GetPlayerIP(Index))
    End If
    Close #F
    For i = 1 To MAX_PLAYER_CHARS
        filename = App.path & "\data\accounts\" & Trim(Name) & "\" & Trim$(Name) & "_char" & CStr(i) & ".bin"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Player(Index).characters(i)
        Close #F
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadPlayer", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearPlayer(ByVal Index As Long)
    Dim i As Long
    

   On Error GoTo errorhandler

    If frmServer.fraEditPlayer.Visible Then
        If EditingPlayer = Index Then
            EditingPlayer = 0
            frmServer.fraEditPlayer.Visible = False
            frmServer.lblNotifications.Caption = "Player Editing Canceled! - Reason: Player Disconnected."
            frmServer.lblNotifications.ForeColor = &HFF&
        End If
    End If
    
    Call ZeroMemory(ByVal VarPtr(TempPlayer(Index)), LenB(TempPlayer(Index)))
    Set TempPlayer(Index).Buffer = New clsBuffer
    
    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).login = vbNullString
    Player(Index).Password = vbNullString
    ReDim Player(Index).characters(MAX_PLAYER_CHARS)
    For i = 1 To MAX_PLAYER_CHARS
        Player(Index).characters(i).Name = vbNullString
        Player(Index).characters(i).Class = 1
    Next
    
    If Index > 0 Then
        frmServer.lvwInfo.ListItems(Index).SubItems(1) = vbNullString
        frmServer.lvwInfo.ListItems(Index).SubItems(2) = vbNullString
        frmServer.lvwInfo.ListItems(Index).SubItems(3) = vbNullString
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearPlayer", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' *************
' ** Classes **
' *************
Public Sub CreateClassesINI()
    Dim filename As String
    Dim File As String, x As Long

   On Error GoTo errorhandler

    filename = App.path & "\data\classes.ini"
    Max_Classes = 1
    
    ReDim Class(Max_Classes)
    
    For x = 1 To Max_Classes
        ReDim Class(x).StartItem(1)
        ReDim Class(x).StartValue(1)
        ReDim Class(x).StartSpell(1)
        Class(x).startItemCount = 0
        Class(x).startSpellCount = 0
    Next
    If Not FileExist(filename, True) Then
        File = FreeFile
        Open filename For Output As File
        Print #File, "[INIT]"
        Print #File, "MaxClasses=" & Max_Classes
        Print #File, "[Class1]"
        Print #File, "MaleSprite=1"
        Print #File, "FemaleSprite=1"
        Print #File, "Strength = 0"
        Print #File, "Endurance = 0"
        Print #File, "Intelligence = 0"
        Print #File, "Agility = 0"
        Print #File, "Willpower = 0"
        Print #File, "Name = Warrior"
        'Head Stuff
        Print #File, "MHair=1"
        Print #File, "MHeads=1"
        Print #File, "MClothes=1"
        Print #File, "MEars=1"
        Print #File, "MEtc=1"
        Print #File, "MEyebrows=1"
        Print #File, "MEyes=1"
        Print #File, "MMouth=1"
        Print #File, "MNose=1"
        Print #File, "FHair=1"
        Print #File, "FHeads=1"
        Print #File, "FClothes=1"
        Print #File, "FEars=1"
        Print #File, "FEtc=1"
        Print #File, "FEyebrows=1"
        Print #File, "FEyes=1"
        Print #File, "FMouth=1"
        Print #File, "FNose=1"
        'End Head Stuff
        Close File
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CreateClassesINI", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub LoadClasses()
    Dim filename As String
    Dim i As Long, n As Long
    Dim tmpSprite As String
    Dim tmpArray() As String
    Dim startItemCount As Long, startSpellCount As Long
    Dim x As Long


   On Error GoTo errorhandler

    If CheckClasses Then
        'ReDim Class(1 To Max_Classes)
        'Call SaveClasses
    Else
        filename = App.path & "\data\classes.ini"
        Max_Classes = Val(GetVar(filename, "INIT", "MaxClasses"))
        ReDim Class(1 To Max_Classes)
    End If

    Call ClearClasses
    
    frmServer.cmbClass.Clear
    
    filename = App.path & "\data\classes.ini"
    
    CharMode = Val(GetVar(filename, "GAME", "CharMode"))
    If CharMode = 0 Then CharMode = 1

    For i = 1 To Max_Classes
        SetLoadingProgress "Loading Classes.", 17, i / Max_Classes
        DoEvents
    
        Class(i).Name = GetVar(filename, "CLASS" & i, "Name")
        
        frmServer.cmbClass.AddItem (Trim$(Class(i).Name))
        
        'Read Head Configurations
        Class(i).MaleFaceParts.FHair = Trim$(GetVar(filename, "CLASS" & i, "MHair"))
        Class(i).MaleFaceParts.FHeads = Trim$(GetVar(filename, "CLASS" & i, "MHeads"))
        Class(i).MaleFaceParts.FCloth = Trim$(GetVar(filename, "CLASS" & i, "MClothes"))
        Class(i).MaleFaceParts.FEars = Trim$(GetVar(filename, "CLASS" & i, "MEars"))
        Class(i).MaleFaceParts.FEtc = Trim$(GetVar(filename, "CLASS" & i, "MEtc"))
        Class(i).MaleFaceParts.FEyebrows = Trim$(GetVar(filename, "CLASS" & i, "MEyebrows"))
        Class(i).MaleFaceParts.FEyes = Trim$(GetVar(filename, "CLASS" & i, "MEyes"))
        Class(i).MaleFaceParts.FMouth = Trim$(GetVar(filename, "CLASS" & i, "MMouth"))
        Class(i).MaleFaceParts.FNose = Trim$(GetVar(filename, "CLASS" & i, "MNose"))
        Class(i).MaleFaceParts.FFace = Trim$(GetVar(filename, "CLASS" & i, "MFaces"))
        
        Class(i).FemaleFaceParts.FHair = Trim$(GetVar(filename, "CLASS" & i, "FHair"))
        Class(i).FemaleFaceParts.FHeads = Trim$(GetVar(filename, "CLASS" & i, "FHeads"))
        Class(i).FemaleFaceParts.FCloth = Trim$(GetVar(filename, "CLASS" & i, "FClothes"))
        Class(i).FemaleFaceParts.FEars = Trim$(GetVar(filename, "CLASS" & i, "FEars"))
        Class(i).FemaleFaceParts.FEtc = Trim$(GetVar(filename, "CLASS" & i, "FEtc"))
        Class(i).FemaleFaceParts.FEyebrows = Trim$(GetVar(filename, "CLASS" & i, "FEyebrows"))
        Class(i).FemaleFaceParts.FEyes = Trim$(GetVar(filename, "CLASS" & i, "FEyes"))
        Class(i).FemaleFaceParts.FMouth = Trim$(GetVar(filename, "CLASS" & i, "FMouth"))
        Class(i).FemaleFaceParts.FNose = Trim$(GetVar(filename, "CLASS" & i, "FNose"))
        Class(i).FemaleFaceParts.FFace = Trim$(GetVar(filename, "CLASS" & i, "FFaces"))
        
        ' continue
        Class(i).stat(Stats.Strength) = Val(GetVar(filename, "CLASS" & i, "Strength"))
        Class(i).stat(Stats.Endurance) = Val(GetVar(filename, "CLASS" & i, "Endurance"))
        Class(i).stat(Stats.Intelligence) = Val(GetVar(filename, "CLASS" & i, "Intelligence"))
        Class(i).stat(Stats.Agility) = Val(GetVar(filename, "CLASS" & i, "Agility"))
        Class(i).stat(Stats.Willpower) = Val(GetVar(filename, "CLASS" & i, "Willpower"))
        
        ' how many starting items?
        startItemCount = Val(GetVar(filename, "CLASS" & i, "StartItemCount"))
        If startItemCount > 0 Then ReDim Class(i).StartItem(1 To startItemCount)
        If startItemCount > 0 Then ReDim Class(i).StartValue(1 To startItemCount)
        
        ' loop for items & values
        Class(i).startItemCount = startItemCount
        If startItemCount >= 1 And startItemCount <= MAX_INV Then
            For x = 1 To startItemCount
                Class(i).StartItem(x) = Val(GetVar(filename, "CLASS" & i, "StartItem" & x))
                Class(i).StartValue(x) = Val(GetVar(filename, "CLASS" & i, "StartValue" & x))
            Next
        End If
        
        ' how many starting spells?
        startSpellCount = Val(GetVar(filename, "CLASS" & i, "StartSpellCount"))
        If startSpellCount > 0 Then ReDim Class(i).StartSpell(1 To startSpellCount)
        
        ' loop for spells
        Class(i).startSpellCount = startSpellCount
        If startSpellCount >= 1 And startSpellCount <= MAX_INV Then
            For x = 1 To startSpellCount
                Class(i).StartSpell(x) = Val(GetVar(filename, "CLASS" & i, "StartSpell" & x))
            Next
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadClasses", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SaveClasses()
    Dim filename As String
    Dim i As Long
    Dim x As Long
    

   On Error GoTo errorhandler

    filename = App.path & "\data\classes.ini"

    For i = 1 To Max_Classes
        Call PutVar(filename, "CLASS" & i, "Name", Trim$(Class(i).Name))
        Call PutVar(filename, "CLASS" & i, "MaleSprite", "1")
        Call PutVar(filename, "CLASS" & i, "FemaleSprite", "1")
        Call PutVar(filename, "CLASS" & i, "Strength", STR(Class(i).stat(Stats.Strength)))
        Call PutVar(filename, "CLASS" & i, "Endurance", STR(Class(i).stat(Stats.Endurance)))
        Call PutVar(filename, "CLASS" & i, "Intelligence", STR(Class(i).stat(Stats.Intelligence)))
        Call PutVar(filename, "CLASS" & i, "Agility", STR(Class(i).stat(Stats.Agility)))
        Call PutVar(filename, "CLASS" & i, "Willpower", STR(Class(i).stat(Stats.Willpower)))
        ' loop for items & values
        For x = 1 To UBound(Class(i).StartItem)
            Call PutVar(filename, "CLASS" & i, "StartItem" & x, STR(Class(i).StartItem(x)))
            Call PutVar(filename, "CLASS" & i, "StartValue" & x, STR(Class(i).StartValue(x)))
        Next
        ' loop for spells
        For x = 1 To UBound(Class(i).StartSpell)
            Call PutVar(filename, "CLASS" & i, "StartSpell" & x, STR(Class(i).StartSpell(x)))
        Next
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveClasses", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function CheckClasses() As Boolean
    Dim filename As String

   On Error GoTo errorhandler

    filename = App.path & "\data\classes.ini"

    If Not FileExist(filename, True) Then
        Call CreateClassesINI
        CheckClasses = True
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CheckClasses", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub ClearClasses()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To Max_Classes
        Call ZeroMemory(ByVal VarPtr(Class(i)), LenB(Class(i)))
        Class(i).Name = vbNullString
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearClasses", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' ***********
' ** Items **
' ***********
Sub SaveItems()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveItems", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SaveItem(ByVal ItemNum As Long)
    Dim filename As String
    Dim F  As Long

   On Error GoTo errorhandler

    filename = App.path & "\data\items\item" & ItemNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Item(ItemNum)
    Close #F


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveItem", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub LoadItems()
    Dim filename As String
    Dim i As Long
    Dim F As Long

   On Error GoTo errorhandler

    Call CheckItems
    
    frmServer.cmbWeapon.Clear
    frmServer.cmbArmor.Clear
    frmServer.cmbHelmet.Clear
    frmServer.cmbShield.Clear
    frmServer.cmbItems.Clear
    
    frmServer.cmbWeapon.AddItem "None."
    frmServer.cmbArmor.AddItem "None."
    frmServer.cmbHelmet.AddItem "None."
    frmServer.cmbShield.AddItem "None."
    frmServer.cmbItems.AddItem "None."

    For i = 1 To MAX_ITEMS
        SetLoadingProgress "Loading Items.", 20, i / MAX_ITEMS
        DoEvents
    
        filename = App.path & "\data\Items\Item" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Item(i)
        Close #F
        
        frmServer.cmbWeapon.AddItem i & ". " & Trim$(Item(i).Name)
        frmServer.cmbArmor.AddItem i & ". " & Trim$(Item(i).Name)
        frmServer.cmbHelmet.AddItem i & ". " & Trim$(Item(i).Name)
        frmServer.cmbShield.AddItem i & ". " & Trim$(Item(i).Name)
        frmServer.cmbItems.AddItem i & ". " & Trim$(Item(i).Name)
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadItems", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub CheckItems()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_ITEMS

        If Not FileExist("\Data\Items\Item" & i & ".dat") Then
            Call SaveItem(i)
        End If

    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckItems", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearItem(ByVal Index As Long)

   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
    Item(Index).Desc = vbNullString
    Item(Index).Sound = "None."


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearItem", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearItems()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
        SetLoadingProgress "Clearing Items.", 12, i / MAX_ITEMS
        DoEvents
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearItems", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' ***********
' ** Shops **
' ***********
Sub SaveShops()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveShops", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SaveShop(ByVal shopNum As Long)
    Dim filename As String
    Dim F As Long

   On Error GoTo errorhandler

    filename = App.path & "\data\shops\shop" & shopNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Shop(shopNum)
    Close #F


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveShop", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub LoadShops()
    Dim filename As String
    Dim i As Long
    Dim F As Long

   On Error GoTo errorhandler

    Call CheckShops

    For i = 1 To MAX_SHOPS
        SetLoadingProgress "Loading Shops.", 23, i / MAX_SHOPS
        DoEvents
    
        filename = App.path & "\data\shops\shop" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Shop(i)
        Close #F
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadShops", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub CheckShops()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_SHOPS

        If Not FileExist("\Data\shops\shop" & i & ".dat") Then
            Call SaveShop(i)
        End If

    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckShops", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearShop(ByVal Index As Long)

   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).Name = vbNullString


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearShop", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearShops()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
        SetLoadingProgress "Clearing Shops.", 13, i / MAX_SHOPS
        DoEvents
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearShops", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' ************
' ** Spells **
' ************
Sub SaveSpell(ByVal Spellnum As Long)
    Dim filename As String
    Dim F As Long

   On Error GoTo errorhandler

    filename = App.path & "\data\spells\spells" & Spellnum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Spell(Spellnum)
    Close #F


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveSpell", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SaveSpells()
    Dim i As Long

   On Error GoTo errorhandler

    Call SetStatus("Saving spells... ")

    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveSpells", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub LoadSpells()
    Dim filename As String
    Dim i As Long
    Dim F As Long

   On Error GoTo errorhandler

    Call CheckSpells
    
    frmServer.cmbSpells.Clear
    frmServer.cmbSpells.AddItem "None."

    For i = 1 To MAX_SPELLS
        SetLoadingProgress "Loading Spells.", 24, i / MAX_SPELLS
        DoEvents
        filename = App.path & "\data\spells\spells" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Spell(i)
        Close #F
        frmServer.cmbSpells.AddItem i & ". " & Trim$(Spell(i).Name)
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadSpells", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub CheckSpells()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_SPELLS

        If Not FileExist("\Data\spells\spells" & i & ".dat") Then
            Call SaveSpell(i)
        End If

    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckSpells", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearSpell(ByVal Index As Long)

   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).Name = vbNullString
    Spell(Index).LevelReq = 1 'Needs to be 1 for the spell editor
    Spell(Index).Desc = vbNullString
    Spell(Index).Sound = "None."


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearSpell", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearSpells()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
        SetLoadingProgress "Clearing Spells.", 14, i / MAX_SPELLS
        DoEvents
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearSpells", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' **********
' ** NPCs **
' **********
Sub SaveNpcs()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveNpcs", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SaveNpc(ByVal npcnum As Long)
    Dim filename As String
    Dim F As Long

   On Error GoTo errorhandler

    filename = App.path & "\data\npcs\npc" & npcnum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Npc(npcnum)
    Close #F


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveNpc", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub LoadNpcs()
    Dim filename As String
    Dim i As Long
    Dim F As Long

   On Error GoTo errorhandler

    Call CheckNpcs

    For i = 1 To MAX_NPCS
        SetLoadingProgress "Loading NPCs.", 21, i / MAX_NPCS
        DoEvents
    
        filename = App.path & "\data\npcs\npc" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Npc(i)
        Close #F
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadNpcs", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub CheckNpcs()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_NPCS

        If Not FileExist("\Data\npcs\npc" & i & ".dat") Then
            Call SaveNpc(i)
        End If

    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckNpcs", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearNpc(ByVal Index As Long)

   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Npc(Index)), LenB(Npc(Index)))
    Npc(Index).Name = vbNullString
    Npc(Index).AttackSay = vbNullString
    Npc(Index).Sound = "None."


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearNpc", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearNpcs()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
        SetLoadingProgress "Clearing NPCs.", 10, i / MAX_NPCS
        DoEvents
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearNpcs", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' **********
' ** Resources **
' **********
Sub SaveResources()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_RESOURCES
        Call SaveResource(i)
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveResources", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SaveResource(ByVal ResourceNum As Long)
    Dim filename As String
    Dim F As Long

   On Error GoTo errorhandler

    filename = App.path & "\data\resources\resource" & ResourceNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Put #F, , Resource(ResourceNum)
    Close #F


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveResource", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub LoadResources()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim sLen As Long
    

   On Error GoTo errorhandler

    Call CheckResources

    For i = 1 To MAX_RESOURCES
        SetLoadingProgress "Loading Resources.", 22, i / MAX_RESOURCES
        DoEvents
    
        filename = App.path & "\data\resources\resource" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
            Get #F, , Resource(i)
        Close #F
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadResources", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub CheckResources()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_RESOURCES
        If Not FileExist("\Data\Resources\Resource" & i & ".dat") Then
            Call SaveResource(i)
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckResources", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearResource(ByVal Index As Long)

   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).Name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).Sound = "None."


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearResource", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearResources()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
        SetLoadingProgress "Clearing Resources.", 11, i / MAX_RESOURCES
        DoEvents
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearResources", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' **********
' ** animations **
' **********
Sub SaveAnimations()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_ANIMATIONS
        Call SaveAnimation(i)
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveAnimations", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SaveAnimation(ByVal AnimationNum As Long)
    Dim filename As String
    Dim F As Long

   On Error GoTo errorhandler

    filename = App.path & "\data\animations\animation" & AnimationNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Put #F, , Animation(AnimationNum)
    Close #F


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveAnimation", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub LoadAnimations()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim sLen As Long
    

   On Error GoTo errorhandler

    Call CheckAnimations

    For i = 1 To MAX_ANIMATIONS
        SetLoadingProgress "Loading Animations.", 25, i / MAX_ANIMATIONS
        DoEvents
        filename = App.path & "\data\animations\animation" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
            Get #F, , Animation(i)
        Close #F
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadAnimations", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub CheckAnimations()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_ANIMATIONS

        If Not FileExist("\Data\animations\animation" & i & ".dat") Then
            Call SaveAnimation(i)
        End If

    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckAnimations", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearAnimation(ByVal Index As Long)

   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).Name = vbNullString
    Animation(Index).Sound = "None."


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearAnimation", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearAnimations()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
        SetLoadingProgress "Clearing Animations.", 15, i / MAX_ANIMATIONS
        DoEvents
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearAnimations", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' **********
' ** Maps **
' **********
Sub SaveMap(ByVal MapNum As Long)
    Dim filename As String
    Dim F As Long
    Dim x As Long
    Dim y As Long, i As Long, z As Long, w As Long

   On Error GoTo errorhandler
   
    'Only save non-instanced maps
    If MapNum < 1 Then Exit Sub

    filename = App.path & "\data\maps\map" & MapNum & ".dat"
    F = FreeFile
    
    Open filename For Binary As #F
    Put #F, , Map(MapNum).Name
    Put #F, , Map(MapNum).Music
    Put #F, , Map(MapNum).BGS
    Put #F, , Map(MapNum).Revision
    Put #F, , Map(MapNum).Moral
    Put #F, , Map(MapNum).Up
    Put #F, , Map(MapNum).Down
    Put #F, , Map(MapNum).Left
    Put #F, , Map(MapNum).Right
    Put #F, , Map(MapNum).BootMap
    Put #F, , Map(MapNum).BootX
    Put #F, , Map(MapNum).BootY
    
    Put #F, , Map(MapNum).Weather
    Put #F, , Map(MapNum).WeatherIntensity
    
    Put #F, , Map(MapNum).Fog
    Put #F, , Map(MapNum).FogSpeed
    Put #F, , Map(MapNum).FogOpacity
    
    Put #F, , Map(MapNum).Red
    Put #F, , Map(MapNum).Green
    Put #F, , Map(MapNum).Blue
    Put #F, , Map(MapNum).Alpha
    
    Put #F, , Map(MapNum).MaxX
    Put #F, , Map(MapNum).MaxY

    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY
            Put #F, , Map(MapNum).Tile(x, y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Put #F, , Map(MapNum).Npc(x)
        Put #F, , Map(MapNum).NpcSpawnType(x)
    Next
    
    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY
            Put #F, , Map(MapNum).ExTile(x, y)
        Next
    Next
    
    Close #F
    
    'This is for event saving, it is in .ini files becuase there are non-limited values (strings) that cannot easily be loaded/saved in the normal manner.
    filename = App.path & "\data\maps\map" & MapNum & "_eventdata.dat"
    PutVar filename, "Events", "EventCount", Val(Map(MapNum).EventCount)
    
    If Map(MapNum).EventCount > 0 Then
        For i = 1 To Map(MapNum).EventCount
            With Map(MapNum).Events(i)
                PutVar filename, "Event" & i, "Name", .Name
                PutVar filename, "Event" & i, "Global", Val(.Global)
                PutVar filename, "Event" & i, "x", Val(.x)
                PutVar filename, "Event" & i, "y", Val(.y)
                PutVar filename, "Event" & i, "PageCount", Val(.PageCount)
            End With
            If Map(MapNum).Events(i).PageCount > 0 Then
                For x = 1 To Map(MapNum).Events(i).PageCount
                    With Map(MapNum).Events(i).Pages(x)
                        PutVar filename, "Event" & i & "Page" & x, "chkVariable", Val(.chkVariable)
                        PutVar filename, "Event" & i & "Page" & x, "VariableIndex", Val(.VariableIndex)
                        PutVar filename, "Event" & i & "Page" & x, "VariableCondition", Val(.VariableCondition)
                        PutVar filename, "Event" & i & "Page" & x, "VariableCompare", Val(.VariableCompare)
                        
                        PutVar filename, "Event" & i & "Page" & x, "chkSwitch", Val(.chkSwitch)
                        PutVar filename, "Event" & i & "Page" & x, "SwitchIndex", Val(.SwitchIndex)
                        PutVar filename, "Event" & i & "Page" & x, "SwitchCompare", Val(.SwitchCompare)
                        
                        PutVar filename, "Event" & i & "Page" & x, "chkHasItem", Val(.chkHasItem)
                        PutVar filename, "Event" & i & "Page" & x, "HasItemIndex", Val(.HasItemIndex)
                        PutVar filename, "Event" & i & "Page" & x, "HasItemAmount", Val(.HasItemAmount)
                        
                        PutVar filename, "Event" & i & "Page" & x, "chkSelfSwitch", Val(.chkSelfSwitch)
                        PutVar filename, "Event" & i & "Page" & x, "SelfSwitchIndex", Val(.SelfSwitchIndex)
                        PutVar filename, "Event" & i & "Page" & x, "SelfSwitchCompare", Val(.SelfSwitchCompare)
                        
                        PutVar filename, "Event" & i & "Page" & x, "GraphicType", Val(.GraphicType)
                        PutVar filename, "Event" & i & "Page" & x, "Graphic", Val(.Graphic)
                        PutVar filename, "Event" & i & "Page" & x, "GraphicX", Val(.GraphicX)
                        PutVar filename, "Event" & i & "Page" & x, "GraphicY", Val(.GraphicY)
                        PutVar filename, "Event" & i & "Page" & x, "GraphicX2", Val(.GraphicX2)
                        PutVar filename, "Event" & i & "Page" & x, "GraphicY2", Val(.GraphicY2)
                        
                        PutVar filename, "Event" & i & "Page" & x, "MoveType", Val(.MoveType)
                        PutVar filename, "Event" & i & "Page" & x, "MoveSpeed", Val(.MoveSpeed)
                        PutVar filename, "Event" & i & "Page" & x, "MoveFreq", Val(.MoveFreq)
                        
                        PutVar filename, "Event" & i & "Page" & x, "IgnoreMoveRoute", Val(.IgnoreMoveRoute)
                        PutVar filename, "Event" & i & "Page" & x, "RepeatMoveRoute", Val(.RepeatMoveRoute)
                        
                        PutVar filename, "Event" & i & "Page" & x, "MoveRouteCount", Val(.MoveRouteCount)
                        
                        If .MoveRouteCount > 0 Then
                            For y = 1 To .MoveRouteCount
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Index", Val(.MoveRoute(y).Index)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data1", Val(.MoveRoute(y).Data1)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data2", Val(.MoveRoute(y).Data2)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data3", Val(.MoveRoute(y).Data3)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data4", Val(.MoveRoute(y).Data4)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data5", Val(.MoveRoute(y).data5)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data6", Val(.MoveRoute(y).data6)
                            Next
                        End If
                        
                        PutVar filename, "Event" & i & "Page" & x, "WalkAnim", Val(.WalkAnim)
                        PutVar filename, "Event" & i & "Page" & x, "DirFix", Val(.DirFix)
                        PutVar filename, "Event" & i & "Page" & x, "WalkThrough", Val(.WalkThrough)
                        PutVar filename, "Event" & i & "Page" & x, "ShowName", Val(.ShowName)
                        PutVar filename, "Event" & i & "Page" & x, "Trigger", Val(.Trigger)
                        PutVar filename, "Event" & i & "Page" & x, "CommandListCount", Val(.CommandListCount)
                        
                        PutVar filename, "Event" & i & "Page" & x, "Position", Val(.Position)
                        PutVar filename, "Event" & i & "Page" & x, "QuestNum", Val(.questnum)
                    End With
                    
                    If Map(MapNum).Events(i).Pages(x).CommandListCount > 0 Then
                        For y = 1 To Map(MapNum).Events(i).Pages(x).CommandListCount
                            PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "CommandCount", Val(Map(MapNum).Events(i).Pages(x).CommandList(y).CommandCount)
                            PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "ParentList", Val(Map(MapNum).Events(i).Pages(x).CommandList(y).ParentList)
                            If Map(MapNum).Events(i).Pages(x).CommandList(y).CommandCount > 0 Then
                                For z = 1 To Map(MapNum).Events(i).Pages(x).CommandList(y).CommandCount
                                    With Map(MapNum).Events(i).Pages(x).CommandList(y).Commands(z)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Index", Val(.Index)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Text1", .Text1
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Text2", .Text2
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Text3", .Text3
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Text4", .Text4
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Text5", .Text5
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Data1", Val(.Data1)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Data2", Val(.Data2)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Data3", Val(.Data3)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Data4", Val(.Data4)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Data5", Val(.data5)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Data6", Val(.data6)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "ConditionalBranchCommandList", Val(.ConditionalBranch.CommandList)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "ConditionalBranchCondition", Val(.ConditionalBranch.Condition)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "ConditionalBranchData1", Val(.ConditionalBranch.Data1)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "ConditionalBranchData2", Val(.ConditionalBranch.Data2)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "ConditionalBranchData3", Val(.ConditionalBranch.Data3)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "ConditionalBranchElseCommandList", Val(.ConditionalBranch.ElseCommandList)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRouteCount", Val(.MoveRouteCount)
                                        If .MoveRouteCount > 0 Then
                                            For w = 1 To .MoveRouteCount
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRoute" & w & "Index", Val(.MoveRoute(w).Index)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRoute" & w & "Data1", Val(.MoveRoute(w).Data1)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRoute" & w & "Data2", Val(.MoveRoute(w).Data2)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRoute" & w & "Data3", Val(.MoveRoute(w).Data3)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRoute" & w & "Data4", Val(.MoveRoute(w).Data4)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRoute" & w & "Data5", Val(.MoveRoute(w).data5)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRoute" & w & "Data6", Val(.MoveRoute(w).data6)
                                            Next
                                        End If
                                    End With
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        Next
    End If
    
    
    DoEvents


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveMap", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SaveMaps()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_MAPS
        Call SaveMap(i)
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveMaps", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub LoadMaps()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim x As Long
    Dim y As Long, z As Long, p As Long, w As Long
    Dim newtileset As Long, newtiley As Long

  ' On Error GoTo errorhandler

    Call CheckMaps

    For i = 1 To MAX_MAPS
        SetLoadingProgress "Loading Maps.", 18, i / MAX_MAPS
        DoEvents
        filename = App.path & "\data\maps\map" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Map(i).Name
        Get #F, , Map(i).Music
        Get #F, , Map(i).BGS
        Get #F, , Map(i).Revision
        Get #F, , Map(i).Moral
        Get #F, , Map(i).Up
        Get #F, , Map(i).Down
        Get #F, , Map(i).Left
        Get #F, , Map(i).Right
        Get #F, , Map(i).BootMap
        Get #F, , Map(i).BootX
        Get #F, , Map(i).BootY
        
        Get #F, , Map(i).Weather
        Get #F, , Map(i).WeatherIntensity
        
        Get #F, , Map(i).Fog
        Get #F, , Map(i).FogSpeed
        Get #F, , Map(i).FogOpacity
        
        Get #F, , Map(i).Red
        Get #F, , Map(i).Green
        Get #F, , Map(i).Blue
        Get #F, , Map(i).Alpha
        
        Get #F, , Map(i).MaxX
        Get #F, , Map(i).MaxY
        ' have to set the tile()
        ReDim Map(i).Tile(0 To Map(i).MaxX, 0 To Map(i).MaxY)
        ReDim Map(i).ExTile(0 To Map(i).MaxX, 0 To Map(i).MaxY)
        
        For x = 0 To Map(i).MaxX
            For y = 0 To Map(i).MaxY
                Get #F, , Map(i).Tile(x, y)
            Next
        Next

        For x = 1 To MAX_MAP_NPCS
            Get #F, , Map(i).Npc(x)
            Get #F, , Map(i).NpcSpawnType(x)
            MapNpc(i).Npc(x).Num = Map(i).Npc(x)
        Next
        
        For x = 0 To Map(i).MaxX
            For y = 0 To Map(i).MaxY
                Get #F, , Map(i).ExTile(x, y)
            Next
        Next

        Close #F
        
        ClearTempTile i
        CacheResources i
        DoEvents
        CacheMapBlocks i
    Next
    
    For z = 1 To MAX_MAPS
        filename = App.path & "\data\maps\map" & z & "_eventdata.dat"
        Map(z).EventCount = Val(GetVar(filename, "Events", "EventCount"))
        
        If Map(z).EventCount > 0 Then
            ReDim Map(z).Events(0 To Map(z).EventCount)
            For i = 1 To Map(z).EventCount
                If Val(GetVar(filename, "Event" & i, "PageCount")) > 0 Then
                    With Map(z).Events(i)
                        .Name = GetVar(filename, "Event" & i, "Name")
                        .Global = Val(GetVar(filename, "Event" & i, "Global"))
                        .x = Val(GetVar(filename, "Event" & i, "x"))
                        .y = Val(GetVar(filename, "Event" & i, "y"))
                        .PageCount = Val(GetVar(filename, "Event" & i, "PageCount"))
                    End With
                    If Map(z).Events(i).PageCount > 0 Then
                        ReDim Map(z).Events(i).Pages(0 To Map(z).Events(i).PageCount)
                        For x = 1 To Map(z).Events(i).PageCount
                            With Map(z).Events(i).Pages(x)
                                .chkVariable = Val(GetVar(filename, "Event" & i & "Page" & x, "chkVariable"))
                                .VariableIndex = Val(GetVar(filename, "Event" & i & "Page" & x, "VariableIndex"))
                                .VariableCondition = Val(GetVar(filename, "Event" & i & "Page" & x, "VariableCondition"))
                                .VariableCompare = Val(GetVar(filename, "Event" & i & "Page" & x, "VariableCompare"))
                                
                                .chkSwitch = Val(GetVar(filename, "Event" & i & "Page" & x, "chkSwitch"))
                                .SwitchIndex = Val(GetVar(filename, "Event" & i & "Page" & x, "SwitchIndex"))
                                .SwitchCompare = Val(GetVar(filename, "Event" & i & "Page" & x, "SwitchCompare"))
                                
                                .chkHasItem = Val(GetVar(filename, "Event" & i & "Page" & x, "chkHasItem"))
                                .HasItemIndex = Val(GetVar(filename, "Event" & i & "Page" & x, "HasItemIndex"))
                                .HasItemAmount = Val(GetVar(filename, "Event" & i & "Page" & x, "HasItemAmount"))
                                
                                .chkSelfSwitch = Val(GetVar(filename, "Event" & i & "Page" & x, "chkSelfSwitch"))
                                .SelfSwitchIndex = Val(GetVar(filename, "Event" & i & "Page" & x, "SelfSwitchIndex"))
                                .SelfSwitchCompare = Val(GetVar(filename, "Event" & i & "Page" & x, "SelfSwitchCompare"))
                                
                                .GraphicType = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicType"))
                                .Graphic = Val(GetVar(filename, "Event" & i & "Page" & x, "Graphic"))
                                .GraphicX = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicX"))
                                .GraphicY = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicY"))
                                .GraphicX2 = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicX2"))
                                .GraphicY2 = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicY2"))
                                
                                .MoveType = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveType"))
                                .MoveSpeed = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveSpeed"))
                                .MoveFreq = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveFreq"))
                                
                                .IgnoreMoveRoute = Val(GetVar(filename, "Event" & i & "Page" & x, "IgnoreMoveRoute"))
                                .RepeatMoveRoute = Val(GetVar(filename, "Event" & i & "Page" & x, "RepeatMoveRoute"))
                                
                                .MoveRouteCount = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRouteCount"))
                                
                                If .MoveRouteCount > 0 Then
                                    ReDim Map(z).Events(i).Pages(x).MoveRoute(0 To .MoveRouteCount)
                                    For y = 1 To .MoveRouteCount
                                        .MoveRoute(y).Index = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Index"))
                                        .MoveRoute(y).Data1 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data1"))
                                        .MoveRoute(y).Data2 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data2"))
                                        .MoveRoute(y).Data3 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data3"))
                                        .MoveRoute(y).Data4 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data4"))
                                        .MoveRoute(y).data5 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data5"))
                                        .MoveRoute(y).data6 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data6"))
                                    Next
                                End If
                                
                                .WalkAnim = Val(GetVar(filename, "Event" & i & "Page" & x, "WalkAnim"))
                                .DirFix = Val(GetVar(filename, "Event" & i & "Page" & x, "DirFix"))
                                .WalkThrough = Val(GetVar(filename, "Event" & i & "Page" & x, "WalkThrough"))
                                .ShowName = Val(GetVar(filename, "Event" & i & "Page" & x, "ShowName"))
                                .Trigger = Val(GetVar(filename, "Event" & i & "Page" & x, "Trigger"))
                                .CommandListCount = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandListCount"))
                             
                                .Position = Val(GetVar(filename, "Event" & i & "Page" & x, "Position"))
                                .questnum = Val(GetVar(filename, "Event" & i & "Page" & x, "QuestNum"))
                            End With
                                
                            If Map(z).Events(i).Pages(x).CommandListCount > 0 Then
                                ReDim Map(z).Events(i).Pages(x).CommandList(0 To Map(z).Events(i).Pages(x).CommandListCount)
                                For y = 1 To Map(z).Events(i).Pages(x).CommandListCount
                                    Map(z).Events(i).Pages(x).CommandList(y).CommandCount = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "CommandCount"))
                                    Map(z).Events(i).Pages(x).CommandList(y).ParentList = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "ParentList"))
                                    If Map(z).Events(i).Pages(x).CommandList(y).CommandCount > 0 Then
                                        ReDim Map(z).Events(i).Pages(x).CommandList(y).Commands(Map(z).Events(i).Pages(x).CommandList(y).CommandCount)
                                        For p = 1 To Map(z).Events(i).Pages(x).CommandList(y).CommandCount
                                            With Map(z).Events(i).Pages(x).CommandList(y).Commands(p)
                                                .Index = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Index"))
                                                .Text1 = GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Text1")
                                                .Text2 = GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Text2")
                                                .Text3 = GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Text3")
                                                .Text4 = GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Text4")
                                                .Text5 = GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Text5")
                                                .Data1 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Data1"))
                                                .Data2 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Data2"))
                                                .Data3 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Data3"))
                                                .Data4 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Data4"))
                                                .data5 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Data5"))
                                                .data6 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Data6"))
                                                .ConditionalBranch.CommandList = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "ConditionalBranchCommandList"))
                                                .ConditionalBranch.Condition = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "ConditionalBranchCondition"))
                                                .ConditionalBranch.Data1 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "ConditionalBranchData1"))
                                                .ConditionalBranch.Data2 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "ConditionalBranchData2"))
                                                .ConditionalBranch.Data3 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "ConditionalBranchData3"))
                                                .ConditionalBranch.ElseCommandList = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "ConditionalBranchElseCommandList"))
                                                .MoveRouteCount = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRouteCount"))
                                                If .MoveRouteCount > 0 Then
                                                    ReDim .MoveRoute(1 To .MoveRouteCount)
                                                    For w = 1 To .MoveRouteCount
                                                        .MoveRoute(w).Index = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRoute" & w & "Index"))
                                                        .MoveRoute(w).Data1 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRoute" & w & "Data1"))
                                                        .MoveRoute(w).Data2 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRoute" & w & "Data2"))
                                                        .MoveRoute(w).Data3 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRoute" & w & "Data3"))
                                                        .MoveRoute(w).Data4 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRoute" & w & "Data4"))
                                                        .MoveRoute(w).data5 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRoute" & w & "Data5"))
                                                        .MoveRoute(w).data6 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRoute" & w & "Data6"))
                                                    Next
                                                End If
                                            End With
                                        Next
                                    End If
                                Next
                            End If
                        Next
                    End If
                End If
            Next
        End If
        DoEvents
    Next
    UpdateMapReport


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadMaps", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub CheckMaps()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_MAPS

        If Not FileExist("\Data\maps\map" & i & ".dat") Then
            Map(i).Name = "New Map"
            Call SaveMap(i)
        End If

    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckMaps", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearMapItem(ByVal Index As Long, ByVal MapNum As Long)

   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapItem(MapNum, Index)), LenB(MapItem(MapNum, Index)))
    MapItem(MapNum, Index).playerName = vbNullString


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearMapItem", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearMapItems()
    Dim x As Long
    Dim y As Long


   On Error GoTo errorhandler

    For y = MIN_MAPS To MAX_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, y)
        Next
        SetLoadingProgress "Clearing Map Items.", 7, y / MAX_MAPS
        DoEvents
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearMapItems", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearMapNpc(ByVal Index As Long, ByVal MapNum As Long)

   On Error GoTo errorhandler

    ReDim MapNpc(MapNum).Npc(1 To MAX_MAP_NPCS)
    Call ZeroMemory(ByVal VarPtr(MapNpc(MapNum).Npc(Index)), LenB(MapNpc(MapNum).Npc(Index)))


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearMapNpc", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearMapNpcs()
    Dim x As Long
    Dim y As Long


   On Error GoTo errorhandler

    For y = MIN_MAPS To MAX_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, y)
        Next
        SetLoadingProgress "Clearing Map Npcs.", 8, y / MAX_MAPS
        DoEvents
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearMapNpcs", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearZoneNpc(ZoneNum As Long, npcnum As Long)

   On Error GoTo errorhandler

    ReDim ZoneNpc(ZoneNum).Npc(1 To MAX_MAP_NPCS * 2)
    Call ZeroMemory(ByVal VarPtr(ZoneNpc(ZoneNum).Npc(npcnum)), LenB(ZoneNpc(ZoneNum).Npc(npcnum)))


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearZoneNpc", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearZoneNpcs()
    Dim x As Long
    Dim y As Long


   On Error GoTo errorhandler

    For y = 1 To MAX_ZONES
        For x = 1 To MAX_MAP_NPCS * 2
            Call ClearZoneNpc(y, x)
        Next
        SetLoadingProgress "Clearing Zone NPCs.", 9, y / MAX_ZONES
        DoEvents
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearZoneNpcs", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearMap(ByVal MapNum As Long)

   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Map(MapNum)), LenB(Map(MapNum)))
    Map(MapNum).Name = vbNullString
    Map(MapNum).MaxX = MAX_MAPX
    Map(MapNum).MaxY = MAX_MAPY
    ReDim Map(MapNum).Tile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
    ReDim Map(MapNum).ExTile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
    ' Reset the map cache array for this map.
    MapCache(MapNum).data = vbNullString


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearMap", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearMaps()
    Dim i As Long


   On Error GoTo errorhandler

    For i = MIN_MAPS To MAX_MAPS
        Call ClearMap(i)
        SetLoadingProgress "Clearing Maps.", 5, i / MAX_MAPS
        DoEvents
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearMaps", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearZone(ByVal ZoneNum As Long)

   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapZones(ZoneNum)), LenB(MapZones(ZoneNum)))
    MapZones(ZoneNum).Name = ""


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearZone", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearZones()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_ZONES
        Call ClearZone(i)
        SetLoadingProgress "Clearing Zones.", 6, i / MAX_ZONES
        DoEvents
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearZones", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function GetClassName(ByVal ClassNum As Long) As String

   On Error GoTo errorhandler

    GetClassName = Trim$(Class(ClassNum).Name)


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetClassName", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Function GetClassMaxVital(ByVal ClassNum As Long, ByVal Vital As Vitals) As Long

   On Error GoTo errorhandler

    Select Case Vital
        Case HP
            With Class(ClassNum)
                GetClassMaxVital = 100 + (.stat(Endurance) * 5) + 2
            End With
        Case MP
            With Class(ClassNum)
                GetClassMaxVital = 30 + (.stat(Intelligence) * 10) + 2
            End With
    End Select


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetClassMaxVital", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Function GetClassStat(ByVal ClassNum As Long, ByVal stat As Stats) As Long

   On Error GoTo errorhandler

    GetClassStat = Class(ClassNum).stat(stat)


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetClassStat", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SaveBank(ByVal Index As Long)
    Dim filename As String
    Dim F As Long
    

   On Error GoTo errorhandler

    filename = App.path & "\data\accounts\" & Trim$(Player(Index).login) & "\" & Trim$(Player(Index).login) & "_char" & CStr(TempPlayer(Index).CurChar) & "_bank.bin"
    
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Bank(Index)
    Close #F


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveBank", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub LoadBank(ByVal Index As Long, ByVal Name As String)
    Dim filename As String
    Dim F As Long


   On Error GoTo errorhandler

    Call ClearBank(Index)

    filename = App.path & "\data\accounts\" & Trim$(Player(Index).login) & "\" & Trim$(Player(Index).login) & "_char" & CStr(TempPlayer(Index).CurChar) & "_bank.bin"
    
    If Not FileExist(filename, True) Then
        Call SaveBank(Index)
        Exit Sub
    End If

    F = FreeFile
    Open filename For Binary As #F
        Get #F, , Bank(Index)
    Close #F


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadBank", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearBank(ByVal Index As Long)

   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Bank(Index)), LenB(Bank(Index)))


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearBank", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearParty(ByVal partyNum As Long)

   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Party(partyNum)), LenB(Party(partyNum)))


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearParty", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SaveSwitches()
Dim i As Long, filename As String

   On Error GoTo errorhandler

filename = App.path & "\data\switches.ini"

For i = 1 To MAX_SWITCHES
    Call PutVar(filename, "Switches", "Switch" & CStr(i) & "Name", Switches(i))
Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveSwitches", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SaveVariables()
Dim i As Long, filename As String

   On Error GoTo errorhandler

filename = App.path & "\data\variables.ini"

For i = 1 To MAX_VARIABLES
    Call PutVar(filename, "Variables", "Variable" & CStr(i) & "Name", Variables(i))
Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveVariables", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub LoadSwitches()
Dim i As Long, filename As String

   On Error GoTo errorhandler

filename = App.path & "\data\switches.ini"

For i = 1 To MAX_SWITCHES
    SetLoadingProgress "Loading Switches.", 26, i / MAX_SWITCHES
    DoEvents
    Switches(i) = GetVar(filename, "Switches", "Switch" & CStr(i) & "Name")
Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadSwitches", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub LoadVariables()
Dim i As Long, filename As String

   On Error GoTo errorhandler

filename = App.path & "\data\variables.ini"

For i = 1 To MAX_VARIABLES
    SetLoadingProgress "Loading Variables.", 27, i / MAX_VARIABLES
    DoEvents
    Variables(i) = GetVar(filename, "Variables", "Variable" & CStr(i) & "Name")
Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadVariables", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub LoadHouses()
Dim i As Long, filepath As String

   On Error GoTo errorhandler

filepath = App.path & "\data\HouseConfig.ini"
    frmServer.lstHouses.Clear
    For i = 1 To MAX_HOUSES
        SetLoadingProgress "Loading Houses.", 28, i / MAX_HOUSES
        DoEvents
        HouseConfig(i).BaseMap = Val(GetVar(filepath, "House" & CStr(i), "BaseMap"))
        HouseConfig(i).ConfigName = Trim$(GetVar(filepath, "House" & CStr(i), "Name"))
        HouseConfig(i).MaxFurniture = Val(GetVar(filepath, "House" & CStr(i), "MaxFurniture"))
        HouseConfig(i).price = Val(GetVar(filepath, "House" & CStr(i), "Price"))
        HouseConfig(i).x = Val(GetVar(filepath, "House" & CStr(i), "X"))
        HouseConfig(i).y = Val(GetVar(filepath, "House" & CStr(i), "Y"))
        frmServer.lstHouses.AddItem i & ". " & HouseConfig(i).ConfigName
    Next
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendHouseConfigs i
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadHouses", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SaveHouse(Index As Long)
Dim filepath As String

   On Error GoTo errorhandler

filepath = App.path & "\data\HouseConfig.ini"
    If Index > 0 And Index <= MAX_HOUSES Then
        Call PutVar(filepath, "House" & CStr(Index), "BaseMap", CStr(HouseConfig(Index).BaseMap))
        Call PutVar(filepath, "House" & CStr(Index), "Name", HouseConfig(Index).ConfigName)
        Call PutVar(filepath, "House" & CStr(Index), "MaxFurniture", CStr(HouseConfig(Index).MaxFurniture))
        Call PutVar(filepath, "House" & CStr(Index), "Price", CStr(HouseConfig(Index).price))
        Call PutVar(filepath, "House" & CStr(Index), "X", CStr(HouseConfig(Index).x))
        Call PutVar(filepath, "House" & CStr(Index), "Y", CStr(HouseConfig(Index).y))
    End If
    LoadHouses


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveHouse", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SaveHouses()
Dim i As Long

   On Error GoTo errorhandler

For i = 1 To MAX_HOUSES
    SaveHouse i
Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveHouses", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function GetAccountFromCharacterName(Name As String, ByRef slot As Long) As String
Dim File As String, F As Long, i As Long, MyName As String, MyPath As String

   On Error GoTo errorhandler

    Name = Trim$(Name)
    Name = LCase(Name)
    MyPath = App.path & "\data\accounts\"
    MyName = Dir(MyPath, vbDirectory)
    Do While MyName <> ""
       If MyName <> "." And MyName <> ".." Then
          If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
            For i = 1 To MAX_PLAYER_CHARS
                F = FreeFile
                Open App.path & "\data\accounts\" & MyName & "\" & MyName & "_char" & CStr(i) & ".bin" For Binary As #F
                Get #F, , MonkeyPlayer.characters(i)
                Close #F
                If Trim$(LCase(MonkeyPlayer.characters(i).Name)) = Name Then
                    GetAccountFromCharacterName = Trim$(MyName)
                    slot = i
                    Exit Function
                End If
            Next
          End If
       End If
       MyName = Dir
    Loop


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetAccountFromCharacterName", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SaveZones()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_ZONES
        Call SaveZone(i)
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveZones", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SaveZone(ByVal Index As Long)
    Dim filename As String
    Dim i As Long

   On Error GoTo errorhandler

    ChkDir App.path & "\Data\", "zones"
    'This is for event saving, it is in .ini files becuase there are non-limited values (strings) that cannot easily be loaded/saved in the normal manner.
    filename = App.path & "\data\zones\zone" & Index & ".ini"
    PutVar filename, "Zone", "Name", Trim$(MapZones(Index).Name)
    PutVar filename, "Zone", "MapCount", CStr(MapZones(Index).MapCount)
    
    If MapZones(Index).MapCount > 0 Then
        For i = 1 To MapZones(Index).MapCount
            PutVar filename, "Zone", "Map" & CStr(i), CStr(MapZones(Index).Maps(i))
        Next
    End If
    
    For i = 1 To MAX_MAP_NPCS * 2
        PutVar filename, "Zone", "NPC" & CStr(i), CStr(MapZones(Index).NPCs(i))
    Next
    For i = 1 To 5
        PutVar filename, "Zone", "Weather" & CStr(i), CStr(MapZones(Index).Weather(i))
    Next
    PutVar filename, "Zone", "WeatherIntensity", CStr(MapZones(Index).WeatherIntensity)
    DoEvents


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveZone", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub LoadZones()
    Dim i As Long

   On Error GoTo errorhandler

    CheckZones
    For i = 1 To MAX_ZONES
        Call LoadZone(i)
        SetLoadingProgress "Loading Zones.", 19, i / MAX_ZONES
        DoEvents
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadZones", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub LoadZone(ByVal Index As Long)
    Dim filename As String
    Dim i As Long
    
    'This is for event saving, it is in .ini files becuase there are non-limited values (strings) that cannot easily be loaded/saved in the normal manner.

   On Error GoTo errorhandler

    filename = App.path & "\data\zones\zone" & Index & ".ini"
    MapZones(Index).Name = Trim$(GetVar(filename, "Zone", "Name"))
    MapZones(Index).MapCount = Val(GetVar(filename, "Zone", "MapCount"))
    
    If MapZones(Index).MapCount > 0 Then
        ReDim MapZones(Index).Maps(MapZones(Index).MapCount)
        For i = 1 To MapZones(Index).MapCount
            MapZones(Index).Maps(i) = Val(GetVar(filename, "Zone", "Map" & CStr(i)))
        Next
    End If
    
    For i = 1 To MAX_MAP_NPCS * 2
        MapZones(Index).NPCs(i) = Val(GetVar(filename, "Zone", "NPC" & CStr(i)))
    Next
    For i = 1 To 5
        MapZones(Index).Weather(i) = Val(GetVar(filename, "Zone", "Weather" & CStr(i)))
    Next
    MapZones(Index).WeatherIntensity = Val(GetVar(filename, "Zone", "WeatherIntensity"))
    DoEvents


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadZone", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub CheckZones()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_ZONES

        If Not FileExist("\Data\zones\zone" & i & ".ini") Then
            Call SaveZone(i)
        End If

    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckZones", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Function GenerateSerial(Length As Long) As String
    Dim serial As String, i As Long

   On Error GoTo errorhandler

    Do Until Len(serial) >= Length
        i = rand(0, 10)
        If i <= 5 Then
            i = rand(0, 9)
            serial = serial & CStr(i)
        Else
            i = rand(0, 10)
            If i < 5 Then
                i = rand(65, 90)
                If i < 65 Then i = 65
                If i > 90 Then i = 90
                serial = serial & Chr(i)
            Else
                i = rand(97, 122)
                If i < 97 Then i = 97
                If i > 122 Then i = 122
                serial = serial & Chr(i)
            End If
        End If
    Loop
    
    If Len(serial) > Length Then
        serial = Right(serial, Length)
    End If
    GenerateSerial = serial


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GenerateSerial", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Sub UpdateMapReport()
Dim i As Long

   On Error GoTo errorhandler

    frmServer.lstMaps.Clear
    For i = 1 To MAX_MAPS
        If Trim$(Map(i).Name) <> "" And Trim$(Map(i).Name) <> "New Map" Then
            frmServer.lstMaps.AddItem (i & ". " & Trim$(Map(i).Name))
        Else
            frmServer.lstMaps.AddItem (i & ". Unnamed")
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "UpdateMapReport", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function GenerateOptionsKey() As String
Dim Key As String, i As Long
   On Error GoTo errorhandler
    Do
        i = Random(1000, 9999)
        If GetSetting("Eclipse Origins", "Server" & i, "InUse") = "" Then
            GenerateOptionsKey = CStr(i)
            SaveSetting "Eclipse Origins", "Server" & i, "InUse", "1"
            Exit Function
        End If
    Loop
   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GenerateOptionsKey", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub LoadBans()
Dim i As Long, filename As String, F As Long
   On Error GoTo errorhandler

    filename = App.path & "\data\banlist.bin"
    BanCount = 0
    ' Check if file exists
    If Not FileExist("data\banlist.bin") Then
        F = FreeFile
        Open filename For Binary As #F
        Close #F
        Exit Sub
    End If

    F = FreeFile
    Open filename For Binary As #F
    Get #F, , BanCount
    ReDim Bans(BanCount)
    If BanCount > 0 Then
        For i = 1 To BanCount
            Get #F, , Bans(i).BanName
            Get #F, , Bans(i).BanChar
            Get #F, , Bans(i).IPAddress
            Get #F, , Bans(i).BanReason
        Next
    End If

    Close #F


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadBans", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub LoadAccounts()
Dim dirLoc As String, filename As String, x As Long, i As Long
Dim dirLoc2 As String, x2 As String, i2 As String

   On Error GoTo errorhandler

    AccountCount = 0
    dirLoc = Dir$(App.path & "\data\accounts\*.*", vbDirectory)
    x = 1
    Do While dirLoc <> ""
        If dirLoc <> "." And dirLoc <> ".." Then
            On Error Resume Next
            If GetAttr(App.path & "\data\accounts\" & dirLoc) = vbDirectory Then
                filename = App.path & "\data\accounts\" & dirLoc & "\" & dirLoc & ".bin"
                If FileExist(filename, True) Then
                    AccountCount = AccountCount + 1
                    ReDim Preserve account(AccountCount)
                    ReDim Preserve account(AccountCount).characters(MAX_PLAYER_CHARS)
                    LoadAccount AccountCount, dirLoc
                End If
            End If
        End If
        dirLoc = Dir$(App.path & "\data\accounts\*.*", vbDirectory)
        For i = 1 To x
            dirLoc = Dir$
        Next
        dirLoc = Dir$
        x = x + 1
    Loop


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadAccounts", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub LoadAccount(Index As Long, Name As String)
    Dim i As Long, x As Long, F As Long, pass As String * NAME_LENGTH, filename As String
    

   On Error GoTo errorhandler
   
   filename = App.path & "\data\accounts\" & Trim$(Name) & "\" & Trim$(Name) & ".bin"

    With account(Index)
        F = FreeFile
        
        Open filename For Binary As #F
            Get #F, , .login
            Get #F, , pass
            .pass = Len(Trim$(pass))
            Get #F, , .ip
        Close #F
        For i = 1 To MAX_PLAYER_CHARS
            filename = App.path & "\data\accounts\" & Trim(Name) & "\" & Trim$(Name) & "_char" & CStr(i) & ".bin"
            F = FreeFile
            Open filename For Binary As #F
            Get #F, , MonkeyPlayer.characters(i)
            Close #F
            .characters(i) = Trim$(MonkeyPlayer.characters(i).Name)
            If MonkeyPlayer.characters(i).access > .access Then
                .access = MonkeyPlayer.characters(i).access
            End If
        Next
        
    End With


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadAccount", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Function FindAccount(accname As String) As Long
Dim i As Long

   On Error GoTo errorhandler
    accname = Trim$(accname)
    If AccountCount > 0 Then
        For i = 1 To AccountCount
            If Trim$(account(i).login) = accname Then
                FindAccount = i
                Exit Function
            End If
        Next
    End If
   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "FindAccount", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Function
' *****************
' ** Projectiles **
' *****************
Sub SaveProjectiles()
Dim i As Long

   On Error GoTo errorhandler

    For i = 1 To MAX_PROJECTILES
        Call SaveProjectile(i)
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveProjectiles", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SaveProjectile(ByVal ProjectileNum As Long)
    Dim filename As String
    Dim F As Long

   On Error GoTo errorhandler

    filename = App.path & "\data\projectiles\Projectile" & ProjectileNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Put #F, , Projectiles(ProjectileNum)
    Close #F


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveProjectile", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub LoadProjectiles()
    Dim filename As String
    Dim i As Integer
    Dim F As Long
    Dim sLen As Long
    

   On Error GoTo errorhandler

    Call CheckProjectile

    For i = 1 To MAX_PROJECTILES
        filename = App.path & "\data\projectiles\Projectile" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
            Get #F, , Projectiles(i)
        Close #F
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadProjectiles", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub CheckProjectile()
    Dim i As Long

   On Error GoTo errorhandler

    For i = 1 To MAX_PROJECTILES
        If Not FileExist("\Data\projectiles\Projectile" & i & ".dat") Then
            Call SaveProjectile(i)
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckProjectile", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearMapProjectiles()
Dim x As Long
Dim y As Long

   On Error GoTo errorhandler

    For x = MIN_MAPS To MAX_MAPS
        For y = 1 To MAX_PROJECTILES
            ClearMapProjectile x, y
        Next
    Next

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearMapProjectiles", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearMapProjectile(ByVal MapNum As Long, ByVal Index As Long)

   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapProjectiles(MapNum, Index)), LenB(MapProjectiles(MapNum, Index)))

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearMapProjectile", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearProjectile(ByVal Index As Long)

   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Projectiles(Index)), LenB(Projectiles(Index)))
    Projectiles(Index).Name = vbNullString


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearProjectile", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearProjectiles()
    Dim i As Long

   On Error GoTo errorhandler

    For i = 1 To MAX_PROJECTILES
        Call ClearProjectile(i)
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearProjectiles", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

