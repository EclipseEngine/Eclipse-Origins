Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long

Public Sub HandleError(ByVal procName As String, ByVal contName As String, erNumber As Long, erDesc As String, errLine As Long)
Dim filename As String

    ' filename = App.path & "\data files\logs\errors.txt"
   ' Open filename For Append As #1
   '     Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
   '     Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
   '     Print #1, "The error occured on line " & errLine
   '     Print #1, ""
   ' Close #1

   ' ErrorCount = ErrorCount + 1
   ' UpdateDebugCaption
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)


   On Error GoTo errorhandler

    If LCase$(dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ChkDir", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Function FileExist(ByVal filename As String, Optional RAW As Boolean = False) As Boolean


   On Error GoTo errorhandler

    If Not RAW Then
        If LenB(dir(App.path & "\" & filename)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(dir(filename)) > 0 Then
            FileExist = True
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "FileExist", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found


   On Error GoTo errorhandler

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetVar", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)

   On Error GoTo errorhandler

    Call WritePrivateProfileString$(Header, Var, Value, File)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PutVar", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SaveOptions()
Dim filename As String



   On Error GoTo errorhandler

    filename = App.path & "\Data Files\config.ini"
    Call PutVar(filename, "Options", "Music", str(Options.Music))
    Call PutVar(filename, "Options", "Sound", str(Options.sound))
    Call PutVar(filename, "Options", "Debug", str(Options.Debug))
    Call PutVar(filename, "Options", "Render", str(Options.Render))
    Call PutVar(filename, "Options", "ClickToWalk", str(Options.ClicktoWalk))
    Call PutVar(filename, "Options", "Fullscreen", str(Options.FullScreen))
    Call PutVar(filename, "Options", "DefaultServer", str(Options.DefaultServer))
    Call PutVar(filename, "Options", "GfxMode", str(Options.GfxMode))
    Call PutVar(filename, "Options", "HideServerList", str(Options.HideServerList))

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveOptions", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SaveServers()
Dim filename As String, i As Long

   On Error GoTo errorhandler

    filename = App.path & "\Data Files\config.ini"
    Call PutVar(filename, "Servers", "ServerCount", str(ServerCount))
    If ServerCount > 0 Then
        For i = 1 To ServerCount
            Call PutVar(filename, "Server" & CStr(i), "Game_Name", Trim$(Servers(i).Game_Name))
            Call PutVar(filename, "Server" & CStr(i), "IP", Trim$(Servers(i).ip))
            Call PutVar(filename, "Server" & CStr(i), "Port", str(Servers(i).port))
            Call PutVar(filename, "Server" & CStr(i), "Username", Trim$(Servers(i).Username))
            Call PutVar(filename, "Server" & CStr(i), "SavePass", str(1))
            If Servers(i).SavePass = 1 Then
                Call PutVar(filename, "Server" & CStr(i), "Password", Trim$(Servers(i).Password))
            Else
                Call PutVar(filename, "Server" & CStr(i), "Password", "")
            End If
        Next
    End If

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveOptions", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub LoadOptions()
Dim filename As String, i As Long



   On Error GoTo errorhandler

    filename = App.path & "\Data Files\config.ini"
    If Not FileExist(filename, True) Then
        Options.Music = 1
        Options.sound = 1
        Options.Debug = 0
        Options.Render = 5
        Options.ClicktoWalk = 1
        Options.FullScreen = 0
        Options.DefaultServer = 0
        Options.GfxMode = 0
        Options.HideServerList = 0
        ServerCount = 0
        ReDim Servers(0)
        SaveOptions
    Else
        Options.Music = GetVar(filename, "Options", "Music")
        Options.sound = GetVar(filename, "Options", "Sound")
        Options.Debug = GetVar(filename, "Options", "Debug")
        Options.Render = Val(GetVar(filename, "Options", "Render"))
        Options.ClicktoWalk = Val(GetVar(filename, "Options", "ClickToWalk"))
        Options.FullScreen = Val(GetVar(filename, "Options", "FullScreen"))
        Options.DefaultServer = Val(GetVar(filename, "Options", "DefaultServer"))
        Options.GfxMode = Val(GetVar(filename, "Options", "GfxMode"))
        Options.HideServerList = Val(GetVar(filename, "Options", "HideServerList"))
        ServerCount = Val(GetVar(filename, "Servers", "ServerCount"))
        ReDim Servers(ServerCount)
        If ServerCount > 0 Then
            For i = 1 To ServerCount
                Servers(i).Game_Name = Trim$(GetVar(filename, "Server" & CStr(i), "Game_Name"))
                Servers(i).ip = Trim$(GetVar(filename, "Server" & CStr(i), "IP"))
                Servers(i).port = Val(GetVar(filename, "Server" & CStr(i), "Port"))
                Servers(i).SavePass = Val(GetVar(filename, "Server" & CStr(i), "SavePass"))
                Servers(i).Username = Trim$(GetVar(filename, "Server" & CStr(i), "Username"))
                Servers(i).Password = Trim$(GetVar(filename, "Server" & CStr(i), "Password"))
            Next
        End If
    End If
    
    If Options.Debug = 1 Then DebugMode = True Else DebugMode = False


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadOptions", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SaveMap(ByVal MapNum As Long)
Dim filename As String
Dim f As Long
Dim X As Long
Dim Y As Long, i As Long, z As Long, w As Long



   On Error GoTo errorhandler
   
    'Check for instance
    If MapNum < 1 Then Exit Sub

    filename = App.path & MAP_PATH & "map" & MapNum & MAP_EXT

    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Map.Name
    Put #f, , Map.Music
    Put #f, , Map.BGS
    Put #f, , Map.Revision
    Put #f, , Map.Moral
    Put #f, , Map.Up
    Put #f, , Map.Down
    Put #f, , Map.Left
    Put #f, , Map.Right
    Put #f, , Map.BootMap
    Put #f, , Map.BootX
    Put #f, , Map.BootY
    Put #f, , Map.Weather
    Put #f, , Map.WeatherIntensity
    Put #f, , Map.Fog
    Put #f, , Map.FogSpeed
    Put #f, , Map.FogOpacity
    Put #f, , Map.Red
    Put #f, , Map.Green
    Put #f, , Map.Blue
    Put #f, , Map.Alpha
    Put #f, , Map.MaxX
    Put #f, , Map.MaxY

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            Put #f, , Map.Tile(X, Y)
        Next

        DoEvents
    Next

    For X = 1 To MAX_MAP_NPCS
        Put #f, , Map.Npc(X)
        Put #f, , Map.NpcSpawnType(X)
    Next
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            Put #f, , Map.exTile(X, Y)
        Next

        DoEvents
    Next

    Close #f


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveMap", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub LoadMap(ByVal MapNum As Long)
Dim filename As String
Dim f As Long
Dim X As Long
Dim Y As Long, i As Long, z As Long, w As Long, p As Long



   On Error GoTo errorhandler

    filename = App.path & MAP_PATH & "map" & MapNum & MAP_EXT
    ClearMap
    f = FreeFile
    Open filename For Binary As #f
    Get #f, , Map.Name
    Get #f, , Map.Music
    Get #f, , Map.BGS
    Get #f, , Map.Revision
    Get #f, , Map.Moral
    Get #f, , Map.Up
    Get #f, , Map.Down
    Get #f, , Map.Left
    Get #f, , Map.Right
    Get #f, , Map.BootMap
    Get #f, , Map.BootX
    Get #f, , Map.BootY
    Get #f, , Map.Weather
    Get #f, , Map.WeatherIntensity
        Get #f, , Map.Fog
    Get #f, , Map.FogSpeed
    Get #f, , Map.FogOpacity
        Get #f, , Map.Red
    Get #f, , Map.Green
    Get #f, , Map.Blue
    Get #f, , Map.Alpha
    Get #f, , Map.MaxX
    Get #f, , Map.MaxY
    ' have to set the tile()
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)
    ReDim Map.exTile(0 To Map.MaxX, 0 To Map.MaxY)
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            Get #f, , Map.Tile(X, Y)
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Get #f, , Map.Npc(X)
        Get #f, , Map.NpcSpawnType(X)
    Next
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            Get #f, , Map.exTile(X, Y)
        Next
    Next

    Close #f
    ClearTempTile


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadMap", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CheckTilesets()
Dim i As Long



   On Error GoTo errorhandler

    i = 1
    NumTileSets = 1
    ReDim Tex_Tileset(1)

    While FileExist(GFX_PATH & "tilesets\" & i & GFX_EXT)
        ReDim Preserve Tex_Tileset(NumTileSets)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Tileset(NumTileSets).filepath = App.path & GFX_PATH & "tilesets\" & i & GFX_EXT
        Tex_Tileset(NumTileSets).Texture = NumTextures
        NumTileSets = NumTileSets + 1
        i = i + 1
    Wend
    NumTileSets = NumTileSets - 1
    If NumTileSets = 0 Then Exit Sub


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckTilesets", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CheckCharacters()
Dim i As Long



   On Error GoTo errorhandler

    i = 1
    NumCharacters = 1
    ReDim Tex_Character(1)

    While FileExist(GFX_PATH & "characters\" & i & GFX_EXT)
        ReDim Preserve Tex_Character(NumCharacters)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Character(NumCharacters).filepath = App.path & GFX_PATH & "characters\" & i & GFX_EXT
        Tex_Character(NumCharacters).Texture = NumTextures
        NumCharacters = NumCharacters + 1
        i = i + 1
    Wend
    NumCharacters = NumCharacters - 1
    If NumCharacters = 0 Then Exit Sub


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckCharacters", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CheckPaperdolls()
Dim i As Long



   On Error GoTo errorhandler

    i = 1
    NumPaperdolls = 1
    ReDim Tex_Paperdoll(1)

    While FileExist(GFX_PATH & "paperdolls\" & i & GFX_EXT)
        ReDim Preserve Tex_Paperdoll(NumPaperdolls)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Paperdoll(NumPaperdolls).filepath = App.path & GFX_PATH & "paperdolls\" & i & GFX_EXT
        Tex_Paperdoll(NumPaperdolls).Texture = NumTextures
        NumPaperdolls = NumPaperdolls + 1
        i = i + 1
    Wend
    NumPaperdolls = NumPaperdolls - 1
    If NumPaperdolls = 0 Then Exit Sub


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckPaperdolls", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CheckAnimations()
Dim i As Long



   On Error GoTo errorhandler

    i = 1
    NumAnimations = 1
    ReDim Tex_Animation(1)
    ReDim AnimationTimer(1)

    While FileExist(GFX_PATH & "animations\" & i & GFX_EXT)
        ReDim Preserve Tex_Animation(NumAnimations)
        ReDim Preserve AnimationTimer(NumAnimations)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Animation(NumAnimations).Texture = NumTextures
        Tex_Animation(NumAnimations).filepath = App.path & GFX_PATH & "animations\" & i & GFX_EXT
        NumAnimations = NumAnimations + 1
        i = i + 1
    Wend
    NumAnimations = NumAnimations - 1
    If NumAnimations = 0 Then Exit Sub

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckAnimations", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CheckItems()
Dim i As Long



   On Error GoTo errorhandler

    i = 1
    numitems = 1
    ReDim Tex_Item(1)

    While FileExist(GFX_PATH & "items\" & i & GFX_EXT)
        ReDim Preserve Tex_Item(numitems)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Item(numitems).filepath = App.path & GFX_PATH & "items\" & i & GFX_EXT
        Tex_Item(numitems).Texture = NumTextures
        numitems = numitems + 1
        i = i + 1
    Wend
    numitems = numitems - 1
    If numitems = 0 Then Exit Sub


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckItems", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CheckPics()
Dim i As Long



   On Error GoTo errorhandler

    i = 1
    NumPics = 1
    ReDim Tex_Pic(1)

    While FileExist(GFX_PATH & "pictures\" & i & GFX_EXT)
        ReDim Preserve Tex_Pic(NumPics)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Pic(NumPics).filepath = App.path & GFX_PATH & "pictures\" & i & GFX_EXT
        Tex_Pic(NumPics).Texture = NumTextures
        NumPics = NumPics + 1
        i = i + 1
    Wend
    NumPics = NumPics - 1
    If NumPics = 0 Then Exit Sub


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckPics", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CheckFurniture()
Dim i As Long



   On Error GoTo errorhandler

    i = 1
    NumFurniture = 1
    ReDim Tex_Furniture(1)

    While FileExist(GFX_PATH & "furniture\" & i & GFX_EXT)
        ReDim Preserve Tex_Furniture(NumFurniture)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Furniture(NumFurniture).filepath = App.path & GFX_PATH & "furniture\" & i & GFX_EXT
        Tex_Furniture(NumFurniture).Texture = NumTextures
        NumFurniture = NumFurniture + 1
        i = i + 1
    Wend
    NumFurniture = NumFurniture - 1
    If NumFurniture = 0 Then Exit Sub


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckFurniture", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CheckProjectiles()
Dim i As Long

   On Error GoTo errorhandler

    i = 1
    NumProjectiles = 1
    ReDim Tex_Projectiles(1)

    While FileExist(GFX_PATH & "projectiles\" & i & GFX_EXT)
        ReDim Preserve Tex_Projectiles(NumProjectiles)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Projectiles(NumProjectiles).filepath = App.path & GFX_PATH & "projectiles\" & i & GFX_EXT
        Tex_Projectiles(NumProjectiles).Texture = NumTextures
        NumProjectiles = NumProjectiles + 1
        i = i + 1
    Wend
    NumProjectiles = NumProjectiles - 1
    If NumProjectiles = 0 Then Exit Sub

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckProjectiles", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CheckBodies()
Dim i As Long



   On Error GoTo errorhandler

    i = 1
    NumCharBodies = 1
    ReDim Tex_CharBodies(1)
    While FileExist(GFX_PATH & "character creation\Characters\bodies\" & i & GFX_EXT)
        ReDim Preserve Tex_CharBodies(NumCharBodies)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_CharBodies(NumCharBodies).filepath = App.path & GFX_PATH & "character creation\Characters\bodies\" & i & GFX_EXT
        Tex_CharBodies(NumCharBodies).Texture = NumTextures
        NumCharBodies = NumCharBodies + 1
        i = i + 1
    Wend
    NumCharBodies = NumCharBodies - 1
    i = 1
    NumCharHair = 1
    ReDim Tex_CharHair(1)
    While FileExist(GFX_PATH & "character creation\Characters\hair\" & i & GFX_EXT)
        ReDim Preserve Tex_CharHair(NumCharHair)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_CharHair(NumCharHair).filepath = App.path & GFX_PATH & "character creation\Characters\hair\" & i & GFX_EXT
        Tex_CharHair(NumCharHair).Texture = NumTextures
        NumCharHair = NumCharHair + 1
        i = i + 1
    Wend
    NumCharHair = NumCharHair - 1
    i = 1
    NumMaleLegs = 1
    ReDim Tex_MaleLegs(1)
    While FileExist(GFX_PATH & "character creation\Characters\pants\male\" & i & GFX_EXT)
        ReDim Preserve Tex_MaleLegs(NumMaleLegs)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_MaleLegs(NumMaleLegs).filepath = App.path & GFX_PATH & "character creation\Characters\pants\male\" & i & GFX_EXT
        Tex_MaleLegs(NumMaleLegs).Texture = NumTextures
        NumMaleLegs = NumMaleLegs + 1
        i = i + 1
    Wend
    NumMaleLegs = NumMaleLegs - 1
    i = 1
    NumFemaleLegs = 1
    ReDim Tex_FemaleLegs(1)
    While FileExist(GFX_PATH & "character creation\Characters\pants\female\" & i & GFX_EXT)
        ReDim Preserve Tex_FemaleLegs(NumFemaleLegs)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_FemaleLegs(NumFemaleLegs).filepath = App.path & GFX_PATH & "character creation\Characters\pants\female\" & i & GFX_EXT
        Tex_FemaleLegs(NumFemaleLegs).Texture = NumTextures
        NumFemaleLegs = NumFemaleLegs + 1
        i = i + 1
    Wend
    NumFemaleLegs = NumFemaleLegs - 1
    i = 1
    NumCharShirts = 1
    ReDim Tex_CharShirts(1)
    While FileExist(GFX_PATH & "character creation\Characters\shirts\" & i & GFX_EXT)
        ReDim Preserve Tex_CharShirts(NumCharShirts)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_CharShirts(NumCharShirts).filepath = App.path & GFX_PATH & "character creation\Characters\shirts\" & i & GFX_EXT
        Tex_CharShirts(NumCharShirts).Texture = NumTextures
        NumCharShirts = NumCharShirts + 1
        i = i + 1
    Wend
    NumCharShirts = NumCharShirts - 1
    i = 1
    NumMaleShoes = 1
    ReDim Tex_MaleShoes(1)
    While FileExist(GFX_PATH & "character creation\Characters\shoes\male\" & i & GFX_EXT)
        ReDim Preserve Tex_MaleShoes(NumMaleShoes)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_MaleShoes(NumMaleShoes).filepath = App.path & GFX_PATH & "character creation\Characters\shoes\male\" & i & GFX_EXT
        Tex_MaleShoes(NumMaleShoes).Texture = NumTextures
        NumMaleShoes = NumMaleShoes + 1
        i = i + 1
    Wend
    NumMaleShoes = NumMaleShoes - 1
    i = 1
    NumFemaleShoes = 1
    ReDim Tex_FemaleShoes(1)
    While FileExist(GFX_PATH & "character creation\Characters\shoes\female\" & i & GFX_EXT)
        ReDim Preserve Tex_FemaleShoes(NumFemaleShoes)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_FemaleShoes(NumFemaleShoes).filepath = App.path & GFX_PATH & "character creation\Characters\shoes\female\" & i & GFX_EXT
        Tex_FemaleShoes(NumFemaleShoes).Texture = NumTextures
        NumFemaleShoes = NumFemaleShoes + 1
        i = i + 1
    Wend
    NumFemaleShoes = NumFemaleShoes - 1


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckBodies", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
Public Sub CheckCharFaces()
Dim i As Long



   On Error GoTo errorhandler

    i = 1
    NumFaceHair = 1
    ReDim Tex_FHair(1)
    While FileExist(GFX_PATH & "character creation\Faces\Hair\" & i & GFX_EXT)
        ReDim Preserve Tex_FHair(NumFaceHair)
        ReDim Preserve Tex_FHairB(NumFaceHair)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_FHair(NumFaceHair).filepath = App.path & GFX_PATH & "character creation\Faces\Hair\" & i & GFX_EXT
        Tex_FHair(NumFaceHair).Texture = NumTextures
            If FileExist(GFX_PATH & "character creation\Faces\Hair\" & i & "_b" & GFX_EXT) = True Then
            NumTextures = NumTextures + 1
            ReDim Preserve gTexture(NumTextures)
            Tex_FHairB(NumFaceHair).filepath = App.path & GFX_PATH & "character creation\Faces\Hair\" & i & "_b" & GFX_EXT
            Tex_FHairB(NumFaceHair).Texture = NumTextures
        End If
        NumFaceHair = NumFaceHair + 1
        i = i + 1
    Wend
    NumFaceHair = NumFaceHair - 1
    i = 1
    NumFaceHeads = 1
    ReDim Tex_FHeads(1)
    While FileExist(GFX_PATH & "character creation\Faces\Heads\" & i & GFX_EXT)
        ReDim Preserve Tex_FHeads(NumFaceHeads)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_FHeads(NumFaceHeads).filepath = App.path & GFX_PATH & "character creation\Faces\Heads\" & i & GFX_EXT
        Tex_FHeads(NumFaceHeads).Texture = NumTextures
        NumFaceHeads = NumFaceHeads + 1
        i = i + 1
    Wend
    NumFaceHeads = NumFaceHeads - 1
    i = 1
    NumFaceEyes = 1
    ReDim Tex_FEyes(1)
    While FileExist(GFX_PATH & "character creation\Faces\Eyes\" & i & GFX_EXT)
        ReDim Preserve Tex_FEyes(NumFaceEyes)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_FEyes(NumFaceEyes).filepath = App.path & GFX_PATH & "character creation\Faces\Eyes\" & i & GFX_EXT
        Tex_FEyes(NumFaceEyes).Texture = NumTextures
        NumFaceEyes = NumFaceEyes + 1
        i = i + 1
    Wend
    NumFaceEyes = NumFaceEyes - 1
    i = 1
    NumFaceEyebrows = 1
    ReDim Tex_FEyebrows(1)
    While FileExist(GFX_PATH & "character creation\Faces\Eyebrows\" & i & GFX_EXT)
        ReDim Preserve Tex_FEyebrows(NumFaceEyebrows)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_FEyebrows(NumFaceEyebrows).filepath = App.path & GFX_PATH & "character creation\Faces\Eyebrows\" & i & GFX_EXT
        Tex_FEyebrows(NumFaceEyebrows).Texture = NumTextures
        NumFaceEyebrows = NumFaceEyebrows + 1
        i = i + 1
    Wend
    NumFaceEyebrows = NumFaceEyebrows - 1
    i = 1
    NumFaceEars = 1
    ReDim Tex_FEars(1)
    While FileExist(GFX_PATH & "character creation\Faces\Ears\" & i & GFX_EXT)
        ReDim Preserve Tex_FEars(NumFaceEars)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_FEars(NumFaceEars).filepath = App.path & GFX_PATH & "character creation\Faces\Ears\" & i & GFX_EXT
        Tex_FEars(NumFaceEars).Texture = NumTextures
        NumFaceEars = NumFaceEars + 1
        i = i + 1
    Wend
    NumFaceEars = NumFaceEars - 1
    i = 1
    NumFaceMouths = 1
    ReDim Tex_FMouths(1)
    While FileExist(GFX_PATH & "character creation\Faces\Mouths\" & i & GFX_EXT)
        ReDim Preserve Tex_FMouth(NumFaceMouths)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_FMouth(NumFaceMouths).filepath = App.path & GFX_PATH & "character creation\Faces\Mouths\" & i & GFX_EXT
        Tex_FMouth(NumFaceMouths).Texture = NumTextures
        NumFaceMouths = NumFaceMouths + 1
        i = i + 1
    Wend
    NumFaceMouths = NumFaceMouths - 1
    i = 1
    NumFaceNoses = 1
    ReDim Tex_FNoses(1)
    While FileExist(GFX_PATH & "character creation\Faces\Noses\" & i & GFX_EXT)
        ReDim Preserve Tex_FNose(NumFaceNoses)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_FNose(NumFaceNoses).filepath = App.path & GFX_PATH & "character creation\Faces\Noses\" & i & GFX_EXT
        Tex_FNose(NumFaceNoses).Texture = NumTextures
        NumFaceNoses = NumFaceNoses + 1
        i = i + 1
    Wend
    NumFaceNoses = NumFaceNoses - 1
    i = 1
    NumFaceEtc = 1
    ReDim Tex_FEtc(1)
    While FileExist(GFX_PATH & "character creation\Faces\Etc\" & i & GFX_EXT)
        ReDim Preserve Tex_FEtc(NumFaceEtc)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_FEtc(NumFaceEtc).filepath = App.path & GFX_PATH & "character creation\Faces\Etc\" & i & GFX_EXT
        Tex_FEtc(NumFaceEtc).Texture = NumTextures
        NumFaceEtc = NumFaceEtc + 1
        i = i + 1
    Wend
    NumFaceEtc = NumFaceEtc - 1
    i = 1
    NumFaceShirts = 1
    ReDim Tex_FShirts(1)
    While FileExist(GFX_PATH & "character creation\Faces\Clothes\" & i & GFX_EXT)
        ReDim Preserve Tex_FShirts(NumFaceShirts)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_FShirts(NumFaceShirts).filepath = App.path & GFX_PATH & "character creation\Faces\Clothes\" & i & GFX_EXT
        Tex_FShirts(NumFaceShirts).Texture = NumTextures
        NumFaceShirts = NumFaceShirts + 1
        i = i + 1
    Wend
    NumFaceShirts = NumFaceShirts - 1


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckCharFaces", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CheckResources()
Dim i As Long



   On Error GoTo errorhandler

    i = 1
    NumResources = 1
    ReDim Tex_Resource(1)

    While FileExist(GFX_PATH & "resources\" & i & GFX_EXT)
        ReDim Preserve Tex_Resource(NumResources)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Resource(NumResources).filepath = App.path & GFX_PATH & "resources\" & i & GFX_EXT
        Tex_Resource(NumResources).Texture = NumTextures
        NumResources = NumResources + 1
        i = i + 1
    Wend
    NumResources = NumResources - 1
    If NumResources = 0 Then Exit Sub


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckResources", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CheckSpellIcons()
Dim i As Long



   On Error GoTo errorhandler

    i = 1
    NumSpellIcons = 1
    ReDim Tex_SpellIcon(1)

    While FileExist(GFX_PATH & "spellicons\" & i & GFX_EXT)
        ReDim Preserve Tex_SpellIcon(NumSpellIcons)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_SpellIcon(NumSpellIcons).filepath = App.path & GFX_PATH & "spellicons\" & i & GFX_EXT
        Tex_SpellIcon(NumSpellIcons).Texture = NumTextures
        NumSpellIcons = NumSpellIcons + 1
        i = i + 1
    Wend
    NumSpellIcons = NumSpellIcons - 1
    If NumSpellIcons = 0 Then Exit Sub


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckSpellIcons", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CheckFaces()
Dim i As Long



   On Error GoTo errorhandler

    i = 1
    NumFaces = 1
    ReDim Tex_Face(1)

    While FileExist(GFX_PATH & "Faces\" & i & GFX_EXT)
        ReDim Preserve Tex_Face(NumFaces)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Face(NumFaces).filepath = App.path & GFX_PATH & "faces\" & i & GFX_EXT
        Tex_Face(NumFaces).Texture = NumTextures
        NumFaces = NumFaces + 1
        i = i + 1
    Wend
    If i - 1 < NumCharacters Then
        For i = i To NumCharacters
            ReDim Preserve Tex_Face(NumFaces)
            If FileExist(GFX_PATH & "Faces\" & i & GFX_EXT) Then
                NumTextures = NumTextures + 1
                ReDim Preserve gTexture(NumTextures)
                Tex_Face(NumFaces).filepath = App.path & GFX_PATH & "faces\" & i & GFX_EXT
                Tex_Face(NumFaces).Texture = NumTextures
            End If
            NumFaces = NumFaces + 1
        Next
    End If
    NumFaces = NumFaces - 1
    
    If NumFaces = 0 Then Exit Sub


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckFaces", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CheckFogs()
Dim i As Long



   On Error GoTo errorhandler

    i = 1
    NumFogs = 1
    ReDim Tex_Fog(1)
    While FileExist(GFX_PATH & "fogs\" & i & GFX_EXT)
        ReDim Preserve Tex_Fog(NumFogs)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Fog(NumFogs).filepath = App.path & GFX_PATH & "fogs\" & i & GFX_EXT
        Tex_Fog(NumFogs).Texture = NumTextures
        NumFogs = NumFogs + 1
        i = i + 1
    Wend
    NumFogs = NumFogs - 1
    If NumFogs = 0 Then Exit Sub


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckFogs", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CheckGUIs()
Dim i As Long



   On Error GoTo errorhandler

    i = 1
    NumGUI = 1
    ReDim Tex_GUI(1)
    While FileExist(GFX_PATH & "gui\" & i & GFX_EXT)
        ReDim Preserve Tex_GUI(NumGUI)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_GUI(NumGUI).filepath = App.path & GFX_PATH & "gui\" & i & GFX_EXT
        Tex_GUI(NumGUI).Texture = NumTextures
        NumGUI = NumGUI + 1
        i = i + 1
    Wend
    NumGUI = NumGUI - 1
    If NumGUI = 0 Then Exit Sub


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckGUIs", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearPlayer(ByVal Index As Long)
Dim oldhp As Long, oldmp As Long, oldmaxhp As Long, oldmaxmp As Long, oldname As String

   On Error GoTo errorhandler
    oldhp = GetPlayerVital(Index, HP)
    oldmp = GetPlayerVital(Index, MP)
    oldmaxhp = GetPlayerMaxVital(Index, HP)
    oldmaxmp = GetPlayerMaxVital(Index, MP)
    oldname = Player(Index).Name
    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).Name = oldname
    SetPlayerVital Index, HP, oldhp
    SetPlayerVital Index, MP, oldmp
    Player(Index).MaxVital(HP) = oldmaxhp
    Player(Index).MaxVital(MP) = oldmaxmp


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearPlayer", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearItem(ByVal Index As Long)


   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
    Item(Index).Desc = vbNullString
    Item(Index).sound = "None."


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
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearItems", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearAnimInstance(ByVal Index As Long)


   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(AnimInstance(Index)), LenB(AnimInstance(Index)))


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearAnimInstance", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearAnimation(ByVal Index As Long)


   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).Name = vbNullString
    Animation(Index).sound = "None."


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
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearAnimations", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearNPC(ByVal Index As Long)


   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Npc(Index)), LenB(Npc(Index)))
    Npc(Index).Name = vbNullString
    Npc(Index).sound = "None."


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearNPC", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearNpcs()
Dim i As Long



   On Error GoTo errorhandler

    For i = 1 To MAX_NPCS
        Call ClearNPC(i)
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearNpcs", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearSpell(ByVal Index As Long)


   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(spell(Index)), LenB(spell(Index)))
    spell(Index).Name = vbNullString
    spell(Index).Desc = vbNullString
    spell(Index).sound = "None."


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
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearSpells", "modDatabase", Err.Number, Err.Description, Erl
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
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearShops", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearResource(ByVal Index As Long)


   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).Name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).sound = "None."


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
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearResources", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearMapItem(ByVal Index As Long)


   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapItem(Index)), LenB(MapItem(Index)))


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearMapItem", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearMap()


   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Map), LenB(Map))
    Map.Name = vbNullString
    Map.MaxX = MAX_MAPX
    Map.MaxY = MAX_MAPY
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)
    ReDim Map.exTile(0 To Map.MaxX, 0 To Map.MaxY)
    initAutotiles


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearMap", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearMapItems()
Dim i As Long



   On Error GoTo errorhandler

    For i = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(i)
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearMapItems", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearMapNpc(ByVal Index As Long)


   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapNpc(Index)), LenB(MapNpc(Index)))


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearMapNpc", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearZoneNpc(zonenum As Long, npcNum As Long)


   On Error GoTo errorhandler
    If npcNum <= 0 Then Exit Sub
    Call ZeroMemory(ByVal VarPtr(ZoneNPC(zonenum).Npc(npcNum)), LenB(ZoneNPC(zonenum).Npc(npcNum)))


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearZoneNpc", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearMapNpcs()
Dim i As Long, X As Long



   On Error GoTo errorhandler

    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(i)
    Next
    For i = 1 To MAX_ZONES
        For X = 1 To MAX_MAP_NPCS
            Call ClearZoneNpc(i, X)
        Next
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearMapNpcs", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' **********************
' ** Player functions **
' **********************
Function GetPlayerName(ByVal Index As Long) As String


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).Name)


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerName", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Name = Name


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerName", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function GetPlayerClass(ByVal Index As Long) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerClass = Player(Index).Class


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerClass", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Class = ClassNum


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerClass", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(Index).Level


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerLevel", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Level = Level


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerLevel", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function GetPlayerExp(ByVal Index As Long) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerExp = Player(Index).Exp


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerExp", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function GetPlayerNextLevel(ByVal Index As Long) As Long

   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    If Player(Index).Level >= MAX_LEVELS Then GetPlayerNextLevel = 0: Exit Function
    GetPlayerNextLevel = (50 / 3) * ((GetPlayerLevel(Index) + 1) ^ 3 - (6 * (GetPlayerLevel(Index) + 1) ^ 2) + 17 * (GetPlayerLevel(Index) + 1) - 12)


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerNextLevel", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Exp = Exp
    If GetPlayerLevel(Index) = MAX_LEVELS And Player(Index).Exp > GetPlayerNextLevel(Index) Then
        Player(Index).Exp = GetPlayerNextLevel(Index)
        Exit Sub
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerExp", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(Index).Access


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerAccess", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Access = Access


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerAccess", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function GetPlayerPK(ByVal Index As Long) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(Index).PK


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerPK", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).PK = PK


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerPK", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(Index).Vital(Vital)


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerVital", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Vital(Vital) = Value

    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerVital", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerMaxVital = Player(Index).MaxVital(Vital)



   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerMaxVital", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function GetPlayerStat(ByVal Index As Long, stat As Stats) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerStat = Player(Index).stat(stat)


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerStat", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub SetPlayerStat(ByVal Index As Long, stat As Stats, ByVal Value As Long)


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    If Value <= 0 Then Value = 1
    If Value > MAX_BYTE Then Value = MAX_BYTE
    Player(Index).stat(stat) = Value


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerStat", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(Index).Points


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal Points As Long)


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Points = Points


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function GetPlayerMap(ByVal Index As Long) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Or Index <= 0 Then Exit Function
    GetPlayerMap = Player(Index).Map


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerMap", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Map = MapNum


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerMap", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function GetPlayerX(ByVal Index As Long) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(Index).X


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerX", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).X = X


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerX", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function GetPlayerY(ByVal Index As Long) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(Index).Y


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerY", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Y = Y


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerY", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function GetPlayerDir(ByVal Index As Long) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(Index).dir


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerDir", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal dir As Long)


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).dir = dir


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerDir", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal invslot As Long) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function
    GetPlayerInvItemNum = PlayerInv(invslot).Num


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal invslot As Long, ByVal ItemNum As Long)
   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invslot).Num = ItemNum


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal invslot As Long) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = PlayerInv(invslot).Value


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal invslot As Long, ByVal ItemValue As Long)


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invslot).Value = ItemValue


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function GetPlayerEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Or Index <= 0 Then Exit Function
    GetPlayerEquipment = Player(Index).Equipment(EquipmentSlot)


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerEquipment", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)


   On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Equipment(EquipmentSlot) = InvNum


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerEquipment", "modDatabase", Err.Number, Err.Description, Erl
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

Sub ClearMapProjectile(ByVal ProjectileNum As Long)
    
   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapProjectiles(ProjectileNum)), LenB(MapProjectiles(ProjectileNum)))

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearMapProjectile", "modDatabase", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

