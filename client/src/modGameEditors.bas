Attribute VB_Name = "modGameEditors"
Option Explicit
Public cpEvent As EventRec
Const LB_SETHORIZONTALEXTENT = &H194
Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public EventList() As EventListRec

' ////////////////
' // Map Editor //
' ////////////////
Public Sub MapEditorInit()
Dim i As Long
Dim smusic() As String

    ' set the width

   On Error GoTo errorhandler

    frmEditor_Map.Width = 7425
    ' we're in the map editor
    InMapEditor = True
    ' show the form
    frmEditor_Map.Visible = True
    ' set the scrolly bars
    frmEditor_Map.scrlTileSet.max = NumTileSets
    frmEditor_Map.fraTileSet.Caption = "Tileset: " & 1
    frmEditor_Map.scrlTileSet.Value = 1
    ' set the scrollbars
    frmEditor_Map.scrlPictureY.max = (Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height \ PIC_Y) - (frmEditor_Map.picBack.Height \ PIC_Y)
    frmEditor_Map.scrlPictureX.max = (Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width \ PIC_X) - (frmEditor_Map.picBack.Width \ PIC_X)
    MapEditorTileScroll
    ' set shops for the shop attribute
    frmEditor_Map.cmbShop.AddItem "None"
    For i = 1 To MAX_SHOPS
        frmEditor_Map.cmbShop.AddItem i & ": " & Shop(i).Name
    Next
    ' we're not in a shop
    frmEditor_Map.cmbShop.ListIndex = 0




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MapEditorInit", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub MapEditorProperties()
Dim X As Long
Dim Y As Long
Dim i As Long
    ' populate the cache if we need to

   On Error GoTo errorhandler

    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_MapProperties.lstMusic.Clear
    frmEditor_MapProperties.lstMusic.AddItem "None."
    For i = 1 To UBound(musicCache)
        frmEditor_MapProperties.lstMusic.AddItem musicCache(i)
    Next
    frmEditor_MapProperties.cmbSound.Clear
    frmEditor_MapProperties.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_MapProperties.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    With frmEditor_MapProperties
        .txtName.Text = Trim$(Map.Name)
            ' find the music we have set
        If .lstMusic.ListCount >= 0 Then
            .lstMusic.ListIndex = 0
            For i = 0 To .lstMusic.ListCount
                If .lstMusic.List(i) = Trim$(Map.Music) Then
                    .lstMusic.ListIndex = i
                End If
            Next
        End If
            If .cmbSound.ListCount >= 0 Then
            .cmbSound.ListIndex = 0
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(Map.BGS) Then
                    .cmbSound.ListIndex = i
                End If
            Next
        End If
            ' rest of it
        .txtUp.Text = CStr(Map.Up)
        .txtDown.Text = CStr(Map.Down)
        .txtLeft.Text = CStr(Map.Left)
        .txtRight.Text = CStr(Map.Right)
        .cmbMoral.ListIndex = Map.Moral
        .txtBootMap.Text = CStr(Map.BootMap)
        .txtBootX.Text = CStr(Map.BootX)
        .txtBootY.Text = CStr(Map.BootY)
            .CmbWeather.ListIndex = Map.Weather
        .scrlWeatherIntensity.Value = Map.WeatherIntensity
            .ScrlFog.Value = Map.Fog
        .ScrlFogSpeed.Value = Map.FogSpeed
        .scrlFogOpacity.Value = Map.FogOpacity
            .ScrlR.Value = Map.Red
        .ScrlG.Value = Map.Green
        .ScrlB.Value = Map.Blue
        .scrlA.Value = Map.Alpha

        ' show the map npcs
        .lstNpcs.Clear
        For X = 1 To MAX_MAP_NPCS
            If Map.Npc(X) > 0 Then
            .lstNpcs.AddItem X & ": " & Trim$(Npc(Map.Npc(X)).Name)
            Else
                .lstNpcs.AddItem X & ": No NPC"
            End If
        Next
        .lstNpcs.ListIndex = 0
            ' show the npc selection combo
        .cmbNpc.Clear
        .cmbNpc.AddItem "No NPC"
        For X = 1 To MAX_NPCS
            .cmbNpc.AddItem X & ": " & Trim$(Npc(X).Name)
        Next
            ' set the combo box properly
        Dim tmpString() As String
        Dim npcNum As Long
        tmpString = Split(.lstNpcs.List(.lstNpcs.ListIndex))
        npcNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
        .cmbNpc.ListIndex = Map.Npc(npcNum)
        ' show the current map
        .lblMap.Caption = "Current map: " & GetPlayerMap(MyIndex)
        .txtMaxX.Text = Map.MaxX
        .txtMaxY.Text = Map.MaxY
    End With




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MapEditorProperties", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub MapEditorSetTile(ByVal X As Long, ByVal Y As Long, ByVal CurLayer As Long, Optional ByVal multitile As Boolean = False, Optional ByVal theAutotile As Byte = 0)
Dim X2 As Long, Y2 As Long



   On Error GoTo errorhandler

    If theAutotile > 0 Then
        If CurLayer > MapLayer.Layer_Count - 1 Then
            With Map.exTile(X, Y)
                ' set layer
                .Layer(CurLayer - (MapLayer.Layer_Count - 1)).X = EditorTileX
                .Layer(CurLayer - (MapLayer.Layer_Count - 1)).Y = EditorTileY
                .Layer(CurLayer - (MapLayer.Layer_Count - 1)).Tileset = frmEditor_Map.scrlTileSet.Value
                .Autotile(CurLayer - (MapLayer.Layer_Count - 1)) = theAutotile
                CacheRenderState X, Y, CurLayer
            End With

        Else
            With Map.Tile(X, Y)
                ' set layer
                .Layer(CurLayer).X = EditorTileX
                .Layer(CurLayer).Y = EditorTileY
                .Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
                .Autotile(CurLayer) = theAutotile
                CacheRenderState X, Y, CurLayer
            End With
        End If
        ' do a re-init so we can see our changes
        initAutotiles
        Exit Sub
    End If

    If Not multitile Then ' single
        If CurLayer > (MapLayer.Layer_Count - 1) Then
            With Map.exTile(X, Y)
                ' set layer
                .Layer(CurLayer - (MapLayer.Layer_Count - 1)).X = EditorTileX
                .Layer(CurLayer - (MapLayer.Layer_Count - 1)).Y = EditorTileY
                .Layer(CurLayer - (MapLayer.Layer_Count - 1)).Tileset = frmEditor_Map.scrlTileSet.Value
                .Autotile(CurLayer - (MapLayer.Layer_Count - 1)) = 0
                CacheRenderState X, Y, CurLayer
            End With
        Else
            With Map.Tile(X, Y)
                ' set layer
                .Layer(CurLayer).X = EditorTileX
                .Layer(CurLayer).Y = EditorTileY
                .Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
                .Autotile(CurLayer) = 0
                CacheRenderState X, Y, CurLayer
            End With
        End If
    Else ' multitile
        Y2 = 0 ' starting tile for y axis
        For Y = CurY To CurY + EditorTileHeight - 1
            X2 = 0 ' re-set x count every y loop
            For X = CurX To CurX + EditorTileWidth - 1
                If X >= 0 And X <= Map.MaxX Then
                    If Y >= 0 And Y <= Map.MaxY Then
                        If CurLayer > (MapLayer.Layer_Count - 1) Then
                            With Map.exTile(X, Y)
                                .Layer(CurLayer - (MapLayer.Layer_Count - 1)).X = EditorTileX + X2
                                .Layer(CurLayer - (MapLayer.Layer_Count - 1)).Y = EditorTileY + Y2
                                .Layer(CurLayer - (MapLayer.Layer_Count - 1)).Tileset = frmEditor_Map.scrlTileSet.Value
                                .Autotile(CurLayer - (MapLayer.Layer_Count - 1)) = 0
                                CacheRenderState X, Y, CurLayer
                            End With

                        Else
                            With Map.Tile(X, Y)
                                .Layer(CurLayer).X = EditorTileX + X2
                                .Layer(CurLayer).Y = EditorTileY + Y2
                                .Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
                                .Autotile(CurLayer) = 0
                                CacheRenderState X, Y, CurLayer
                            End With
                        End If
                    End If
                End If
                X2 = X2 + 1
            Next
            Y2 = Y2 + 1
        Next
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MapEditorSetTile", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub MapEditorMouseDown(ByVal Button As Integer, ByVal X As Long, ByVal Y As Long, Optional ByVal movedMouse As Boolean = True)
Dim i As Long
Dim CurLayer As Long
Dim tmpDir As Byte


    ' find which layer we're on

   On Error GoTo errorhandler
   
    
    If frmEditor_Map.optLayer(1).Value = True Then
        CurLayer = 1
    ElseIf frmEditor_Map.optLayer(2).Value = True Then
        If frmEditor_Map.scrlLayerNum.Value <= 2 Then
            CurLayer = 1 + frmEditor_Map.scrlLayerNum.Value
        Else
            CurLayer = MapLayer.Layer_Count - 1 + frmEditor_Map.scrlLayerNum.Value - 2
        End If
    ElseIf frmEditor_Map.optLayer(4).Value = True Then
        If frmEditor_Map.scrlLayerNum.Value <= 2 Then
            CurLayer = 3 + frmEditor_Map.scrlLayerNum.Value
        Else
            CurLayer = MapLayer.Layer_Count - 1 + 3 + frmEditor_Map.scrlLayerNum.Value - 2
        End If
    Else
    End If
    
    If Not isInBounds Then Exit Sub
    If Button = vbLeftButton Then
        If frmEditor_Map.optLayers.Value Then
            ' no autotiling
            If EditorTileWidth = 1 And EditorTileHeight = 1 Then 'single tile
                MapEditorSetTile CurX, CurY, CurLayer, , frmEditor_Map.scrlAutotile.Value
            Else ' multi tile!
                If frmEditor_Map.scrlAutotile.Value = 0 Then
                    MapEditorSetTile CurX, CurY, CurLayer, True
                Else
                    MapEditorSetTile CurX, CurY, CurLayer, , frmEditor_Map.scrlAutotile.Value
                End If
            End If
        ElseIf frmEditor_Map.optEvent.Value Then
            If frmEditor_Events.Visible = False Then
                AddEvent CurX, CurY
            End If
        ElseIf frmEditor_Map.optAttribs.Value Then
            With Map.Tile(CurX, CurY)
                ' blocked tile
                If frmEditor_Map.optBlocked.Value Then .type = TILE_TYPE_BLOCKED
                ' warp tile
                If frmEditor_Map.optWarp.Value Then
                    .type = TILE_TYPE_WARP
                    .Data1 = EditorWarpMap
                    .data2 = EditorWarpX
                    .Data3 = EditorWarpY
                    .Data4 = ""
                End If
                ' item spawn
                If frmEditor_Map.optItem.Value Then
                    .type = TILE_TYPE_ITEM
                    .Data1 = ItemEditorNum
                    .data2 = ItemEditorValue
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' npc avoid
                If frmEditor_Map.optNpcAvoid.Value Then
                    .type = TILE_TYPE_NPCAVOID
                    .Data1 = 0
                    .data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' key
                If frmEditor_Map.optKey.Value Then
                    .type = TILE_TYPE_KEY
                    .Data1 = KeyEditorNum
                    .data2 = KeyEditorTake
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' key open
                If frmEditor_Map.optKeyOpen.Value Then
                    .type = TILE_TYPE_KEYOPEN
                    .Data1 = KeyOpenEditorX
                    .data2 = KeyOpenEditorY
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' resource
                If frmEditor_Map.optResource.Value Then
                    .type = TILE_TYPE_RESOURCE
                    .Data1 = ResourceEditorNum
                    .data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' door
                If frmEditor_Map.optDoor.Value Then
                    .type = TILE_TYPE_DOOR
                    .Data1 = EditorWarpMap
                    .data2 = EditorWarpX
                    .Data3 = EditorWarpY
                    .Data4 = ""
                End If
                ' npc spawn
                If frmEditor_Map.optNpcSpawn.Value Then
                    .type = TILE_TYPE_NPCSPAWN
                    .Data1 = SpawnNpcNum
                    .data2 = SpawnNpcDir
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' shop
                If frmEditor_Map.optShop.Value Then
                    .type = TILE_TYPE_SHOP
                    .Data1 = EditorShop
                    .data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' bank
                If frmEditor_Map.optBank.Value Then
                    .type = TILE_TYPE_BANK
                    .Data1 = 0
                    .data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' heal
                If frmEditor_Map.optHeal.Value Then
                    .type = TILE_TYPE_HEAL
                    .Data1 = MapEditorHealType
                    .data2 = MapEditorHealAmount
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' trap
                If frmEditor_Map.optTrap.Value Then
                    .type = TILE_TYPE_TRAP
                    .Data1 = MapEditorHealAmount
                    .data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' slide
                If frmEditor_Map.optSlide.Value Then
                    .type = TILE_TYPE_SLIDE
                    .Data1 = MapEditorSlideDir
                    .data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' sound
                If frmEditor_Map.optSound.Value Then
                    .type = TILE_TYPE_SOUND
                    .Data1 = 0
                    .data2 = 0
                    .Data3 = 0
                    .Data4 = MapEditorSound
                End If
                ' sound
                If frmEditor_Map.optHouse.Value Then
                    .type = TILE_TYPE_HOUSE
                    .Data1 = HouseTileIndex
                    .data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' instance
                If frmEditor_Map.optInstance.Value Then
                    .type = TILE_TYPE_INSTANCE
                    .Data1 = EditorWarpMap
                    .data2 = EditorWarpX
                    .Data3 = EditorWarpY
                    .Data4 = ""
                End If
                ' random dungeon
                If frmEditor_Map.optRandomDungeon.Value Then
                    .type = TILE_TYPE_RANDOMDUNGEON
                    .Data1 = MapEditorRandomDungeon
                    .data2 = MapEditorFloorNum
                    .Data3 = 0
                    .Data4 = ""
                End If
            End With
        ElseIf frmEditor_Map.optBlock.Value Then
             With Map.Tile(CurX, CurY)
                .type = TILE_TYPE_BLOCKED
                .Data1 = 0
                .data2 = 0
                .Data3 = 0
                .Data4 = ""
            End With
        End If
    End If

    If Button = vbRightButton Then
        If frmEditor_Map.optLayers.Value Then
            If CurLayer > MapLayer.Layer_Count - 1 Then
                With Map.exTile(CurX, CurY)
                    ' clear layer
                    .Layer(CurLayer - (MapLayer.Layer_Count - 1)).X = 0
                    .Layer(CurLayer - (MapLayer.Layer_Count - 1)).Y = 0
                    .Layer(CurLayer - (MapLayer.Layer_Count - 1)).Tileset = 0
                    If .Autotile(CurLayer - (MapLayer.Layer_Count - 1)) > 0 Then
                        .Autotile(CurLayer - (MapLayer.Layer_Count - 1)) = 0
                        ' do a re-init so we can see our changes
                        initAutotiles
                    End If
                    CacheRenderState X, Y, CurLayer
                End With
            Else
                With Map.Tile(CurX, CurY)
                    ' clear layer
                    .Layer(CurLayer).X = 0
                    .Layer(CurLayer).Y = 0
                    .Layer(CurLayer).Tileset = 0
                    If .Autotile(CurLayer) > 0 Then
                        .Autotile(CurLayer) = 0
                        ' do a re-init so we can see our changes
                        initAutotiles
                    End If
                    CacheRenderState X, Y, CurLayer
                End With
            End If
        ElseIf frmEditor_Map.optEvent.Value Then
            Call DeleteEvent(CurX, CurY)
        ElseIf frmEditor_Map.optAttribs.Value Then
            With Map.Tile(CurX, CurY)
                ' clear attribute
                .type = 0
                .Data1 = 0
                .data2 = 0
                .Data3 = 0
            End With
        ElseIf frmEditor_Map.optBlock.Value Then
            If Map.Tile(CurX, CurY).type = TILE_TYPE_BLOCKED Then Map.Tile(CurX, CurY).type = 0
        End If
    End If

    CacheResources




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MapEditorMouseDown", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub MapEditorChooseTile(Button As Integer, X As Single, Y As Single)


   On Error GoTo errorhandler

    If Button = vbLeftButton Then
        EditorTileWidth = 1
        EditorTileHeight = 1
            EditorTileX = X \ PIC_X
        EditorTileY = Y \ PIC_Y
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MapEditorChooseTile", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub MapEditorDrag(Button As Integer, X As Single, Y As Single)


   On Error GoTo errorhandler

    If Button = vbLeftButton Then
        ' convert the pixel number to tile number
        X = (X \ PIC_X) + 1
        Y = (Y \ PIC_Y) + 1
        ' check it's not out of bounds
        If X < 0 Then X = 0
        If X > Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width / PIC_X Then X = Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width / PIC_X
        If Y < 0 Then Y = 0
        If Y > Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height / PIC_Y Then Y = Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height / PIC_Y
        ' find out what to set the width + height of map editor to
        If X > EditorTileX Then ' drag right
            EditorTileWidth = X - EditorTileX
        Else ' drag left
            ' TO DO
        End If
        If Y > EditorTileY Then ' drag down
            EditorTileHeight = Y - EditorTileY
        Else ' drag up
            ' TO DO
        End If
    End If





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MapEditorDrag", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub MapEditorTileScroll()

    ' horizontal scrolling

   On Error GoTo errorhandler

    If Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width < frmEditor_Map.picBack.Width Then
        frmEditor_Map.scrlPictureX.Enabled = False
    Else
        frmEditor_Map.scrlPictureX.Enabled = True
    End If
    ' vertical scrolling
    If Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height < frmEditor_Map.picBack.Height Then
        frmEditor_Map.scrlPictureY.Enabled = False
    Else
        frmEditor_Map.scrlPictureY.Enabled = True
    End If





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MapEditorTileScroll", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub MapEditorSend()


   On Error GoTo errorhandler

    Call SendMap
    InMapEditor = False
    Unload frmEditor_Map
    GettingMap = True



   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MapEditorSend", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub MapEditorCancel()
    Dim buffer As clsBuffer
   On Error GoTo errorhandler
    InMapEditor = False
    Set buffer = New clsBuffer
    buffer.WriteLong CNeedMap
    buffer.WriteLong GetPlayerMap(MyIndex)
    SendData buffer.ToArray()
    Set buffer = Nothing
    GettingMap = True
    Unload frmEditor_Map
    

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MapEditorCancel", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub MapEditorClearLayer()
Dim i As Long
Dim X As Long
Dim Y As Long
Dim CurLayer As Long


    ' find which layer we're on

   On Error GoTo errorhandler

    If frmEditor_Map.optLayer(1).Value = True Then
        CurLayer = 1
    ElseIf frmEditor_Map.optLayer(2).Value = True Then
        If frmEditor_Map.scrlLayerNum.Value <= 2 Then
            CurLayer = 1 + frmEditor_Map.scrlLayerNum.Value
        Else
            CurLayer = MapLayer.Layer_Count - 1 + frmEditor_Map.scrlLayerNum.Value - 2
        End If
    ElseIf frmEditor_Map.optLayer(4).Value = True Then
        If frmEditor_Map.scrlLayerNum.Value <= 2 Then
            CurLayer = 3 + frmEditor_Map.scrlLayerNum.Value
        Else
            CurLayer = MapLayer.Layer_Count - 1 + 3 + frmEditor_Map.scrlLayerNum.Value - 2
        End If
    Else
    End If
    If CurLayer = 0 Then Exit Sub

    ' ask to clear layer
    If MsgBox("Are you sure you wish to clear this layer?", vbYesNo, Servers(ServerIndex).Game_Name) = vbYes Then
        If CurLayer > MapLayer.Layer_Count - 1 Then
            For X = 0 To Map.MaxX
                For Y = 0 To Map.MaxY
                    Map.exTile(X, Y).Layer(CurLayer - (MapLayer.Layer_Count - 1)).X = 0
                    Map.exTile(X, Y).Layer(CurLayer - (MapLayer.Layer_Count - 1)).Y = 0
                    Map.exTile(X, Y).Layer(CurLayer - (MapLayer.Layer_Count - 1)).Tileset = 0
                    CacheRenderState X, Y, CurLayer
                Next
            Next
        Else
            For X = 0 To Map.MaxX
                For Y = 0 To Map.MaxY
                    Map.Tile(X, Y).Layer(CurLayer).X = 0
                    Map.Tile(X, Y).Layer(CurLayer).Y = 0
                    Map.Tile(X, Y).Layer(CurLayer).Tileset = 0
                    CacheRenderState X, Y, CurLayer
                Next
            Next
        End If
        initAutotiles
    End If





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MapEditorClearLayer", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub MapEditorFillLayer()
Dim i As Long
Dim X As Long
Dim Y As Long
Dim CurLayer As Long


    ' find which layer we're on

   On Error GoTo errorhandler

    If frmEditor_Map.optLayer(1).Value = True Then
        CurLayer = 1
    ElseIf frmEditor_Map.optLayer(2).Value = True Then
        If frmEditor_Map.scrlLayerNum.Value <= 2 Then
            CurLayer = 1 + frmEditor_Map.scrlLayerNum.Value
        Else
            CurLayer = MapLayer.Layer_Count - 1 + frmEditor_Map.scrlLayerNum.Value - 2
        End If
    ElseIf frmEditor_Map.optLayer(4).Value = True Then
        If frmEditor_Map.scrlLayerNum.Value <= 2 Then
            CurLayer = 3 + frmEditor_Map.scrlLayerNum.Value
        Else
            CurLayer = MapLayer.Layer_Count - 1 + 3 + frmEditor_Map.scrlLayerNum.Value - 2
        End If
    Else
    End If
    If CurLayer = 0 Then Exit Sub

    If MsgBox("Are you sure you wish to fill this layer?", vbYesNo, Servers(ServerIndex).Game_Name) = vbYes Then
        If CurLayer > MapLayer.Layer_Count - 1 Then
            For X = 0 To Map.MaxX
                For Y = 0 To Map.MaxY
                    Map.exTile(X, Y).Layer(CurLayer - (MapLayer.Layer_Count - 1)).X = EditorTileX
                    Map.exTile(X, Y).Layer(CurLayer - (MapLayer.Layer_Count - 1)).Y = EditorTileY
                    Map.exTile(X, Y).Layer(CurLayer - (MapLayer.Layer_Count - 1)).Tileset = frmEditor_Map.scrlTileSet.Value
                    Map.exTile(X, Y).Autotile(CurLayer - (MapLayer.Layer_Count - 1)) = frmEditor_Map.scrlAutotile.Value
                    CacheRenderState X, Y, CurLayer
                Next
            Next
        Else
            For X = 0 To Map.MaxX
                For Y = 0 To Map.MaxY
                    Map.Tile(X, Y).Layer(CurLayer).X = EditorTileX
                    Map.Tile(X, Y).Layer(CurLayer).Y = EditorTileY
                    Map.Tile(X, Y).Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
                    Map.Tile(X, Y).Autotile(CurLayer) = frmEditor_Map.scrlAutotile.Value
                    CacheRenderState X, Y, CurLayer
                Next
            Next
        End If
            ' now cache the positions
        initAutotiles
    End If





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MapEditorFillLayer", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub MapEditorClearAttribs()
Dim X As Long
Dim Y As Long



   On Error GoTo errorhandler

    If MsgBox("Are you sure you wish to clear the attributes on this map?", vbYesNo, Servers(ServerIndex).Game_Name) = vbYes Then

        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Map.Tile(X, Y).type = 0
            Next
        Next

    End If





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MapEditorClearAttribs", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub MapEditorLeaveMap()


   On Error GoTo errorhandler

    If InMapEditor Then
        If MsgBox("Save changes to current map?", vbYesNo) = vbYes Then
            Call MapEditorSend
        Else
            Call MapEditorCancel
        End If
    End If





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MapEditorLeaveMap", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' /////////////////
' // Item Editor //
' /////////////////
Public Sub ItemEditorInit()
Dim i As Long
Dim SoundSet As Boolean


   On Error GoTo errorhandler

    If frmEditor_Item.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Item.cmbSound.Clear
    frmEditor_Item.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Item.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating

    With Item(EditorIndex)
        frmEditor_Item.txtName.Text = Trim$(.Name)
        If .pic > frmEditor_Item.scrlPic.max Then .pic = 0
        frmEditor_Item.scrlPic.Value = .pic
        frmEditor_Item.chkStackable.Value = .Stackable
        frmEditor_Item.cmbType.ListIndex = .type
        frmEditor_Item.scrlAnim.Value = .Animation
        frmEditor_Item.txtDesc.Text = Trim$(.Desc)
            ' find the sound we have set
        If frmEditor_Item.cmbSound.ListCount >= 0 Then
            For i = 0 To frmEditor_Item.cmbSound.ListCount
                If frmEditor_Item.cmbSound.List(i) = Trim$(.sound) Then
                    frmEditor_Item.cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or frmEditor_Item.cmbSound.ListIndex = -1 Then frmEditor_Item.cmbSound.ListIndex = 0
        End If

        ' Type specific settings
        If (frmEditor_Item.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmEditor_Item.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
            frmEditor_Item.fraEquipment.Visible = True
            frmEditor_Item.scrlProjectile.Value = .Data1
            frmEditor_Item.scrlDamage.Value = .data2
            frmEditor_Item.cmbTool.ListIndex = .Data3

            If .speed < 100 Then .speed = 100
            frmEditor_Item.scrlSpeed.Value = .speed
                    ' loop for stats
            For i = 1 To Stats.Stat_Count - 1
                frmEditor_Item.scrlStatBonus(i).Value = .Add_Stat(i)
            Next
            If .Paperdoll > 0 And .Paperdoll <= frmEditor_Item.scrlPaperdoll.max Then
                frmEditor_Item.scrlPaperdoll = .Paperdoll
            End If
        Else
            frmEditor_Item.fraEquipment.Visible = False
        End If

        If frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_CONSUME Then
            frmEditor_Item.fraVitals.Visible = True
            frmEditor_Item.scrlAddHp.Value = .AddHP
            frmEditor_Item.scrlAddMP.Value = .AddMP
            frmEditor_Item.scrlAddExp.Value = .AddEXP
            frmEditor_Item.scrlCastSpell.Value = .CastSpell
            frmEditor_Item.chkInstant.Value = .instaCast
        Else
            frmEditor_Item.fraVitals.Visible = False
        End If

        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
            frmEditor_Item.fraSpell.Visible = True
            frmEditor_Item.scrlSpell.Value = .Data1
        Else
            frmEditor_Item.fraSpell.Visible = False
        End If
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_PET) Then
            frmEditor_Item.fraPet.Visible = True
            frmEditor_Item.scrlPet.Value = .Data1
        Else
            frmEditor_Item.fraPet.Visible = False
        End If

        If frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_FURNITURE Then
            frmEditor_Item.fraFurniture.Visible = True
            If Item(EditorIndex).data2 > 0 And Item(EditorIndex).data2 <= NumFurniture Then
                frmEditor_Item.scrlFurniture.Value = Item(EditorIndex).data2
            Else
                frmEditor_Item.scrlFurniture.Value = 1
            End If
            frmEditor_Item.cmbFurnitureType.ListIndex = Item(EditorIndex).Data1
        Else
            frmEditor_Item.fraFurniture.Visible = False
        End If

        ' Basic requirements
        frmEditor_Item.scrlAccessReq.Value = .AccessReq
        frmEditor_Item.scrlLevelReq.Value = .LevelReq
            ' loop for stats
        For i = 1 To Stats.Stat_Count - 1
            frmEditor_Item.scrlStatReq(i).Value = .Stat_Req(i)
        Next
            ' Build cmbClassReq
        frmEditor_Item.cmbClassReq.Clear
        frmEditor_Item.cmbClassReq.AddItem "None"

        For i = 1 To Max_Classes
            frmEditor_Item.cmbClassReq.AddItem Class(i).Name
        Next

        frmEditor_Item.cmbClassReq.ListIndex = .classReq
        ' Info
        frmEditor_Item.scrlPrice.Value = .Price
        frmEditor_Item.cmbBind.ListIndex = .BindType
        frmEditor_Item.scrlRarity.Value = .Rarity
             EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1
    End With
    Item_Changed(EditorIndex) = True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ItemEditorInit", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub ItemEditorOk()
Dim i As Long



   On Error GoTo errorhandler

    For i = 1 To MAX_ITEMS
        If Item_Changed(i) Then
            Call SendSaveItem(i)
        End If
    Next
    Unload frmEditor_Item
    Editor = 0
    ClearChanged_Item




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ItemEditorOk", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub ItemEditorCancel()


   On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Item
    ClearChanged_Item
    ClearItems
    SendRequestItems




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ItemEditorCancel", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub ClearChanged_Item()


   On Error GoTo errorhandler

    ZeroMemory Item_Changed(1), MAX_ITEMS * 2 ' 2 = boolean length




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearChanged_Item", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub ClearChanged_Zone()


   On Error GoTo errorhandler

    ZeroMemory Zone_Changed(1), MAX_ZONES * 2 ' 2 = boolean length




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearChanged_Zone", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub ClearChanged_House()


   On Error GoTo errorhandler

    ZeroMemory House_Changed(1), MAX_HOUSES * 2 ' 2 = boolean length




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearChanged_House", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' /////////////////
' // Animation Editor //
' /////////////////
Public Sub AnimationEditorInit()
Dim i As Long
Dim SoundSet As Boolean


   On Error GoTo errorhandler

    If frmEditor_Animation.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Animation.cmbSound.Clear
    frmEditor_Animation.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Animation.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating

    With Animation(EditorIndex)
        frmEditor_Animation.txtName.Text = Trim$(.Name)
            ' find the sound we have set
        If frmEditor_Animation.cmbSound.ListCount >= 0 Then
            For i = 0 To frmEditor_Animation.cmbSound.ListCount
                If frmEditor_Animation.cmbSound.List(i) = Trim$(.sound) Then
                    frmEditor_Animation.cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or frmEditor_Animation.cmbSound.ListIndex = -1 Then frmEditor_Animation.cmbSound.ListIndex = 0
        End If
            For i = 0 To 1
            frmEditor_Animation.scrlSprite(i).Value = .Sprite(i)
            frmEditor_Animation.scrlFrameCount(i).Value = .Frames(i)
            frmEditor_Animation.scrlLoopCount(i).Value = .LoopCount(i)
                    If .looptime(i) > 0 Then
                frmEditor_Animation.scrlLoopTime(i).Value = .looptime(i)
            Else
                frmEditor_Animation.scrlLoopTime(i).Value = 45
            End If
                Next
             EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    End With
    Animation_Changed(EditorIndex) = True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "AnimationEditorInit", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub AnimationEditorOk()
Dim i As Long



   On Error GoTo errorhandler

    For i = 1 To MAX_ANIMATIONS
        If Animation_Changed(i) Then
            Call SendSaveAnimation(i)
        End If
    Next
    Unload frmEditor_Animation
    Editor = 0
    ClearChanged_Animation




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "AnimationEditorOk", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub AnimationEditorCancel()


   On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Animation
    ClearChanged_Animation
    ClearAnimations
    SendRequestAnimations




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "AnimationEditorCancel", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub ClearChanged_Animation()


   On Error GoTo errorhandler

    ZeroMemory Animation_Changed(1), MAX_ANIMATIONS * 2 ' 2 = boolean length




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearChanged_Animation", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' ////////////////
' // Npc Editor //
' ////////////////
Public Sub NpcEditorInit()
Dim i As Long
Dim SoundSet As Boolean


   On Error GoTo errorhandler

    If frmEditor_NPC.Visible = False Then Exit Sub
    EditorIndex = frmEditor_NPC.lstIndex.ListIndex + 1
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_NPC.cmbSound.Clear
    frmEditor_NPC.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_NPC.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    With frmEditor_NPC
        .txtName.Text = Trim$(Npc(EditorIndex).Name)
        .txtAttackSay.Text = Trim$(Npc(EditorIndex).AttackSay)
        If Npc(EditorIndex).Sprite < 0 Or Npc(EditorIndex).Sprite > .scrlSprite.max Then Npc(EditorIndex).Sprite = 0
        .scrlSprite.Value = Npc(EditorIndex).Sprite
        .txtSpawnSecs.Text = CStr(Npc(EditorIndex).SpawnSecs)
        .cmbBehaviour.ListIndex = Npc(EditorIndex).Behaviour
        .scrlRange.Value = Npc(EditorIndex).Range
        .txtChance.Text = CStr(Npc(EditorIndex).DropChances(1))
        .scrlNum.Value = Npc(EditorIndex).DropItems(1)
        .scrlValue.Value = Npc(EditorIndex).DropItemValues(1)
        .txtHP.Text = Npc(EditorIndex).HP
        .txtExp.Text = Npc(EditorIndex).Exp
        .txtLevel.Text = Npc(EditorIndex).Level
        .txtDamage.Text = Npc(EditorIndex).Damage
        .scrlAnimation.Value = Npc(EditorIndex).Animation
            ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(Npc(EditorIndex).sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
            For i = 1 To Stats.Stat_Count - 1
            .scrlStat(i).Value = Npc(EditorIndex).stat(i)
        Next
        
        If Npc(EditorIndex).ItemBehaviour = 1 Then
            .chkDropItems.Value = 1
        Else
            .chkDropItems.Value = 0
        End If
        .chkProjectile.Value = Npc(EditorIndex).Projectile
    End With
    
    NPC_Changed(EditorIndex) = True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "NpcEditorInit", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub NpcEditorOk()
Dim i As Long



   On Error GoTo errorhandler

    For i = 1 To MAX_NPCS
        If NPC_Changed(i) Then
            Call SendSaveNpc(i)
        End If
    Next
    Unload frmEditor_NPC
    Editor = 0
    ClearChanged_NPC




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "NpcEditorOk", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub NpcEditorCancel()


   On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_NPC
    ClearChanged_NPC
    ClearNpcs
    SendRequestNPCS




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "NpcEditorCancel", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub ClearChanged_NPC()


   On Error GoTo errorhandler

    ZeroMemory NPC_Changed(1), MAX_NPCS * 2 ' 2 = boolean length




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearChanged_NPC", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' ////////////////
' // Resource Editor //
' ////////////////
Public Sub ResourceEditorInit()
Dim i As Long
Dim SoundSet As Boolean



   On Error GoTo errorhandler

    If frmEditor_Resource.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Resource.lstIndex.ListIndex + 1
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Resource.cmbSound.Clear
    frmEditor_Resource.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Resource.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    With frmEditor_Resource
        .scrlExhaustedPic.max = NumResources
        .scrlNormalPic.max = NumResources
        .scrlAnimation.max = MAX_ANIMATIONS
            .txtName.Text = Trim$(Resource(EditorIndex).Name)
        .txtMessage.Text = Trim$(Resource(EditorIndex).SuccessMessage)
        .txtMessage2.Text = Trim$(Resource(EditorIndex).EmptyMessage)
        .cmbType.ListIndex = Resource(EditorIndex).ResourceType
        .scrlNormalPic.Value = Resource(EditorIndex).ResourceImage
        .scrlExhaustedPic.Value = Resource(EditorIndex).ExhaustedImage
        .scrlReward.Value = Resource(EditorIndex).ItemReward
        .scrlTool.Value = Resource(EditorIndex).ToolRequired
        .scrlHealth.Value = Resource(EditorIndex).Health
        .scrlRespawn.Value = Resource(EditorIndex).RespawnTime
        .scrlAnimation.Value = Resource(EditorIndex).Animation
            ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(Resource(EditorIndex).sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
    End With
    Resource_Changed(EditorIndex) = True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ResourceEditorInit", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub ResourceEditorOk()
Dim i As Long



   On Error GoTo errorhandler

    For i = 1 To MAX_RESOURCES
        If Resource_Changed(i) Then
            Call SendSaveResource(i)
        End If
    Next
    Unload frmEditor_Resource
    Editor = 0
    ClearChanged_Resource




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ResourceEditorOk", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub ResourceEditorCancel()


   On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Resource
    ClearChanged_Resource
    ClearResources
    SendRequestResources




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ResourceEditorCancel", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub ClearChanged_Resource()


   On Error GoTo errorhandler

    ZeroMemory Resource_Changed(1), MAX_RESOURCES * 2 ' 2 = boolean length




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearChanged_Resource", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' /////////////////
' // Shop Editor //
' /////////////////
Public Sub ShopEditorInit()
Dim i As Long


   On Error GoTo errorhandler

    If frmEditor_Shop.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Shop.lstIndex.ListIndex + 1
    frmEditor_Shop.txtName.Text = Trim$(Shop(EditorIndex).Name)
    If Shop(EditorIndex).BuyRate > 0 Then
        frmEditor_Shop.scrlBuy.Value = Shop(EditorIndex).BuyRate
    Else
        frmEditor_Shop.scrlBuy.Value = 100
    End If
    frmEditor_Shop.cmbItem.Clear
    frmEditor_Shop.cmbItem.AddItem "None"
    frmEditor_Shop.cmbCostItem.Clear
    frmEditor_Shop.cmbCostItem.AddItem "None"

    For i = 1 To MAX_ITEMS
        frmEditor_Shop.cmbItem.AddItem i & ": " & Trim$(Item(i).Name)
        frmEditor_Shop.cmbCostItem.AddItem i & ": " & Trim$(Item(i).Name)
    Next

    frmEditor_Shop.cmbItem.ListIndex = 0
    frmEditor_Shop.cmbCostItem.ListIndex = 0
    UpdateShopTrade
    Shop_Changed(EditorIndex) = True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ShopEditorInit", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub UpdateShopTrade(Optional ByVal tmpPos As Long = 0)
Dim i As Long



   On Error GoTo errorhandler

    frmEditor_Shop.lstTradeItem.Clear

    For i = 1 To MAX_TRADES
        With Shop(EditorIndex).TradeItem(i)
            ' if none, show as none
            If .Item = 0 And .CostItem = 0 Then
                frmEditor_Shop.lstTradeItem.AddItem "Empty Trade Slot"
            Else
                frmEditor_Shop.lstTradeItem.AddItem i & ": " & .ItemValue & "x " & Trim$(Item(.Item).Name) & " for " & .CostValue & "x " & Trim$(Item(.CostItem).Name)
            End If
        End With
    Next

    frmEditor_Shop.lstTradeItem.ListIndex = tmpPos




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "UpdateShopTrade", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub ShopEditorOk()
Dim i As Long



   On Error GoTo errorhandler

    For i = 1 To MAX_SHOPS
        If Shop_Changed(i) Then
            Call SendSaveShop(i)
        End If
    Next
    Unload frmEditor_Shop
    Editor = 0
    ClearChanged_Shop




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ShopEditorOk", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub ShopEditorCancel()


   On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Shop
    ClearChanged_Shop
    ClearShops
    SendRequestShops




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ShopEditorCancel", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub ClearChanged_Shop()


   On Error GoTo errorhandler

    ZeroMemory Shop_Changed(1), MAX_SHOPS * 2 ' 2 = boolean length




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearChanged_Shop", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' //////////////////
' // Spell Editor //
' //////////////////
Public Sub SpellEditorInit()
Dim i As Long
Dim SoundSet As Boolean


   On Error GoTo errorhandler

    If frmEditor_Spell.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Spell.lstIndex.ListIndex + 1
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Spell.cmbSound.Clear
    frmEditor_Spell.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Spell.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    With frmEditor_Spell
        ' set max values
        .scrlAnimCast.max = MAX_ANIMATIONS
        .scrlAnim.max = MAX_ANIMATIONS
        .scrlAOE.max = MAX_BYTE
        .scrlProjectile.max = MAX_PROJECTILES
        .scrlRange.max = MAX_BYTE
        .scrlMap.max = MAX_MAPS
            ' build class combo
        .cmbClass.Clear
        .cmbClass.AddItem "None"
        For i = 1 To Max_Classes
            .cmbClass.AddItem Trim$(Class(i).Name)
        Next
            If spell(EditorIndex).classReq > -1 And spell(EditorIndex).classReq <= Max_Classes Then
            .cmbClass.ListIndex = spell(EditorIndex).classReq
        End If
            ' set values
        .txtName.Text = Trim$(spell(EditorIndex).Name)
        .txtDesc.Text = Trim$(spell(EditorIndex).Desc)
        .cmbType.ListIndex = spell(EditorIndex).type
        .scrlMP.Value = spell(EditorIndex).MPCost
        .scrlLevel.Value = spell(EditorIndex).LevelReq
        .scrlAccess.Value = spell(EditorIndex).AccessReq
        .cmbClass.ListIndex = spell(EditorIndex).classReq
        .scrlCast.Value = spell(EditorIndex).CastTime
        .scrlCool.Value = spell(EditorIndex).CDTime
        .scrlIcon.Value = spell(EditorIndex).Icon
        .scrlMap.Value = spell(EditorIndex).Map
        .scrlX.Value = spell(EditorIndex).X
        .scrlY.Value = spell(EditorIndex).Y
        .scrlDir.Value = spell(EditorIndex).dir
        If spell(EditorIndex).type = SPELL_TYPE_PET Then
            .scrlVital.Value = spell(EditorIndex).Pet
            .lblVital.Caption = "Pet: " & spell(EditorIndex).Pet & ". " & Trim$(Pet(spell(EditorIndex).Pet).Name)
        Else
            .scrlVital.Value = spell(EditorIndex).Vital
            .lblVital.Caption = "Vital: " & spell(EditorIndex).Vital
        End If
        .scrlVital.Value = spell(EditorIndex).Vital
        .scrlDuration.Value = spell(EditorIndex).Duration
        .scrlInterval.Value = spell(EditorIndex).Interval
        .scrlRange.Value = spell(EditorIndex).Range
        If spell(EditorIndex).IsAoE Then
            .chkAOE.Value = 1
        Else
            .chkAOE.Value = 0
        End If
        .scrlAOE.Value = spell(EditorIndex).AoE
        .scrlAnimCast.Value = spell(EditorIndex).CastAnim
        .scrlAnim.Value = spell(EditorIndex).SpellAnim
        .scrlStun.Value = spell(EditorIndex).StunDuration
        If spell(EditorIndex).IsProjectile Then
            .chkProjectile.Value = 1
        Else
            .chkProjectile.Value = 0
        End If
        .scrlProjectile.Value = spell(EditorIndex).Projectile
            ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(spell(EditorIndex).sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
    End With
    Spell_Changed(EditorIndex) = True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SpellEditorInit", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SpellEditorOk()
Dim i As Long



   On Error GoTo errorhandler

    For i = 1 To MAX_SPELLS
        If Spell_Changed(i) Then
            Call SendSaveSpell(i)
        End If
    Next
    Unload frmEditor_Spell
    Editor = 0
    ClearChanged_Spell




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SpellEditorOk", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SpellEditorCancel()


   On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Spell
    ClearChanged_Spell
    ClearSpells
    SendRequestSpells




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SpellEditorCancel", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub ClearChanged_Spell()


   On Error GoTo errorhandler

    ZeroMemory Spell_Changed(1), MAX_SPELLS * 2 ' 2 = boolean length




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearChanged_Spell", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub ClearAttributeDialogue()


   On Error GoTo errorhandler

    frmEditor_Map.fraNpcSpawn.Visible = False
    frmEditor_Map.fraResource.Visible = False
    frmEditor_Map.fraMapItem.Visible = False
    frmEditor_Map.fraMapKey.Visible = False
    frmEditor_Map.fraKeyOpen.Visible = False
    frmEditor_Map.fraMapWarp.Visible = False
    frmEditor_Map.fraShop.Visible = False
    frmEditor_Map.fraSoundEffect.Visible = False
    frmEditor_Map.fraRandomDungeon.Visible = False




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearAttributeDialogue", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub


'Event Editor Stuffz Also includes event functions from the map editor (copy/paste/delete)

Sub CopyEvent_Map(X As Long, Y As Long)
Dim count As Long, i As Long

   On Error GoTo errorhandler

    count = Map.EventCount
    If count = 0 Then Exit Sub
    For i = 1 To count
        If Map.Events(i).X = X And Map.Events(i).Y = Y Then
            ' copy it
            'CopyMemory ByVal VarPtr(cpEvent), ByVal VarPtr(Map.Events(i)), LenB(Map.Events(i))
             CopyEvent = Map.Events(i)
            ' exit
            Exit Sub
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CopyEvent_Map", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub PasteEvent_Map(X As Long, Y As Long)
Dim count As Long, i As Long, EventNum As Long

   On Error GoTo errorhandler

    count = Map.EventCount
    If count > 0 Then
        For i = 1 To count
            If Map.Events(i).X = X And Map.Events(i).Y = Y Then
                ' already an event - paste over it
                EventNum = i
            End If
        Next
    End If
    ' couldn't find one - create one
    If EventNum = 0 Then
        ' increment count
        AddEvent X, Y, True
        EventNum = count + 1
    End If
    ' copy it
    'CopyMemory ByVal VarPtr(Map.Events(eventNum)), ByVal VarPtr(cpEvent), LenB(cpEvent)
    Map.Events(EventNum) = CopyEvent
    ' set position
    Map.Events(EventNum).X = X
    Map.Events(EventNum).Y = Y


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PasteEvent_Map", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub DeleteEvent(X As Long, Y As Long)
Dim count As Long, i As Long, lowIndex As Long

   On Error GoTo errorhandler

    If Not InMapEditor Then Exit Sub
    If frmEditor_Events.Visible = True Then Exit Sub
    count = Map.EventCount
    For i = 1 To count
        If Map.Events(i).X = X And Map.Events(i).Y = Y Then
            ' delete it
            ClearEvent i
            lowIndex = i
            Exit For
        End If
    Next
    ' not found anything
    If lowIndex = 0 Then Exit Sub
    ' move everything down an index
    For i = lowIndex To count - 1
        Map.Events(i) = Map.Events(i + 1)
    Next
    ' delete the last index
    ClearEvent count
    ' set the new count
    Map.EventCount = count - 1


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DeleteEvent", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub AddEvent(X As Long, Y As Long, Optional ByVal cancelLoad As Boolean = False)
Dim count As Long, pageCount As Long, i As Long

   On Error GoTo errorhandler

    count = Map.EventCount + 1
    ' make sure there's not already an event
    If count - 1 > 0 Then
        For i = 1 To count - 1
            If Map.Events(i).X = X And Map.Events(i).Y = Y Then
                ' already an event - edit it
                If Not cancelLoad Then EventEditorInit i
                Exit Sub
            End If
        Next
    End If
    ' increment count
    Map.EventCount = count
    ReDim Preserve Map.Events(0 To count)
    ' set the new event
    Map.Events(count).X = X
    Map.Events(count).Y = Y
    ' give it a new page
    pageCount = Map.Events(count).pageCount + 1
    Map.Events(count).pageCount = pageCount
    ReDim Preserve Map.Events(count).Pages(pageCount)
    ' load the editor
    If Not cancelLoad Then EventEditorInit count


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "AddEvent", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearEvent(EventNum As Long)

   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Map.Events(EventNum)), LenB(Map.Events(EventNum)))


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearEvent", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub EventEditorInit(EventNum As Long)
Dim i As Long

   On Error GoTo errorhandler

    EditorEvent = EventNum
    ' copy the event data to the temp event
    'CopyMemory ByVal VarPtr(tmpEvent), ByVal VarPtr(Map.Events(eventNum)), LenB(Map.Events(eventNum))
    tmpEvent = Map.Events(EventNum)
    frmEditor_Events.InitEventEditorForm
    ' populate form
    With frmEditor_Events
        ' set the tabs
        .tabPages.Tabs.Clear
        For i = 1 To tmpEvent.pageCount
            .tabPages.Tabs.Add , , str(i)
        Next
        ' items
        .cmbHasItem.Clear
        .cmbHasItem.AddItem "None"
        For i = 1 To MAX_ITEMS
            .cmbHasItem.AddItem i & ": " & Trim$(Item(i).Name)
        Next
            ' variables
        .cmbPlayerVar.Clear
        .cmbPlayerVar.AddItem "None"
        For i = 1 To MAX_VARIABLES
            .cmbPlayerVar.AddItem i & ". " & Variables(i)
        Next
            ' variables
        .cmbPlayerSwitch.Clear
        .cmbPlayerSwitch.AddItem "None"
        For i = 1 To MAX_SWITCHES
            .cmbPlayerSwitch.AddItem i & ". " & Switches(i)
        Next
                ' name
        .txtName.Text = tmpEvent.Name
        ' enable delete button
        If tmpEvent.pageCount > 1 Then
            .cmdDeletePage.Enabled = True
        Else
            .cmdDeletePage.Enabled = False
        End If
        .cmdPastePage.Enabled = False
        ' Load page 1 to start off with
        curPageNum = 1
        EventEditorLoadPage curPageNum
    End With
    ' show the editor
    frmEditor_Events.Show


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EventEditorInit", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub EventEditorLoadPage(pageNum As Long)
    ' populate form

   On Error GoTo errorhandler

    With tmpEvent.Pages(pageNum)
        GraphicSelX = .GraphicX
        GraphicSelY = .GraphicY
        GraphicSelX2 = .GraphicX2
        GraphicSelY2 = .GraphicY2
        frmEditor_Events.cmbGraphic.ListIndex = .GraphicType
        frmEditor_Events.cmbHasItem.ListIndex = .HasItemIndex
        If .HasItemAmount = 0 Then
            frmEditor_Events.scrlCondition_HasItem.Value = 1
        Else
            frmEditor_Events.scrlCondition_HasItem.Value = .HasItemAmount
        End If
        frmEditor_Events.cmbMoveFreq.ListIndex = .MoveFreq
        frmEditor_Events.cmbMoveSpeed.ListIndex = .MoveSpeed
        frmEditor_Events.cmbMoveType.ListIndex = .MoveType
        frmEditor_Events.cmbPlayerVar.ListIndex = .VariableIndex
        frmEditor_Events.cmbPlayerSwitch.ListIndex = .SwitchIndex
        frmEditor_Events.cmbSelfSwitch.ListIndex = .SelfSwitchIndex
        frmEditor_Events.cmbSelfSwitchCompare.ListIndex = .SelfSwitchCompare
        frmEditor_Events.cmbPlayerSwitchCompare.ListIndex = .SwitchCompare
        frmEditor_Events.cmbPlayerVarCompare.ListIndex = .VariableCompare
        frmEditor_Events.chkGlobal.Value = tmpEvent.Global
        frmEditor_Events.cmbTrigger.ListIndex = .Trigger
        frmEditor_Events.chkDirFix.Value = .DirFix
        frmEditor_Events.chkHasItem.Value = .chkHasItem
        frmEditor_Events.chkPlayerVar.Value = .chkVariable
        frmEditor_Events.chkPlayerSwitch.Value = .chkSwitch
        frmEditor_Events.chkSelfSwitch.Value = .chkSelfSwitch
        frmEditor_Events.chkWalkAnim.Value = .WalkAnim
        frmEditor_Events.chkWalkThrough.Value = .WalkThrough
        frmEditor_Events.chkShowName.Value = .ShowName
        frmEditor_Events.txtPlayerVariable = .VariableCondition
        frmEditor_Events.scrlGraphic.Value = .Graphic
        If frmEditor_Events.cmbEventQuest.ListCount > 0 Then
            If .questnum >= 0 And .questnum <= frmEditor_Events.cmbEventQuest.ListCount Then
                frmEditor_Events.cmbEventQuest.ListIndex = .questnum
            End If
        End If
        If frmEditor_Events.cmbEventQuest.ListIndex = -1 Then frmEditor_Events.cmbEventQuest.ListIndex = 0
        If .chkHasItem = 0 Then
            frmEditor_Events.cmbHasItem.Enabled = False
        Else
            frmEditor_Events.cmbHasItem.Enabled = True
        End If
        If .chkSelfSwitch = 0 Then
            frmEditor_Events.cmbSelfSwitch.Enabled = False
            frmEditor_Events.cmbSelfSwitchCompare.Enabled = False
        Else
            frmEditor_Events.cmbSelfSwitch.Enabled = True
            frmEditor_Events.cmbSelfSwitchCompare.Enabled = True
        End If
        If .chkSwitch = 0 Then
            frmEditor_Events.cmbPlayerSwitch.Enabled = False
            frmEditor_Events.cmbPlayerSwitchCompare.Enabled = False
        Else
            frmEditor_Events.cmbPlayerSwitch.Enabled = True
            frmEditor_Events.cmbPlayerSwitchCompare.Enabled = True
        End If
                If .chkVariable = 0 Then
            frmEditor_Events.cmbPlayerVar.Enabled = False
            frmEditor_Events.txtPlayerVariable.Enabled = False
            frmEditor_Events.cmbPlayerVarCompare.Enabled = False
        Else
            frmEditor_Events.cmbPlayerVar.Enabled = True
            frmEditor_Events.txtPlayerVariable.Enabled = True
            frmEditor_Events.cmbPlayerVarCompare.Enabled = True
        End If
            If frmEditor_Events.cmbMoveType.ListIndex = 2 Then
            frmEditor_Events.cmdMoveRoute.Enabled = True
        Else
            frmEditor_Events.cmdMoveRoute.Enabled = False
        End If
            frmEditor_Events.cmbPositioning.ListIndex = .Position
            ' show the commands
        EventListCommands
    End With


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EventEditorLoadPage", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub EventEditorOK()
    ' copy the event data from the temp event
    'CopyMemory ByVal VarPtr(Map.Events(EditorEvent)), ByVal VarPtr(tmpEvent), LenB(tmpEvent)

   On Error GoTo errorhandler

    Map.Events(EditorEvent) = tmpEvent
    ' unload the form
    Unload frmEditor_Events


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EventEditorOK", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub EventListCommands()
Dim i As Long, curlist As Long, oldI As Long, X As Long, indent As String, listleftoff() As Long, conditionalstage() As Long

   On Error GoTo errorhandler

    frmEditor_Events.lstCommands.Clear
    If tmpEvent.Pages(curPageNum).CommandListCount > 0 Then
    ReDim listleftoff(1 To tmpEvent.Pages(curPageNum).CommandListCount)
    ReDim conditionalstage(1 To tmpEvent.Pages(curPageNum).CommandListCount)
        'Start Up at 1
        curlist = 1
        X = -1
newlist:
        For i = 1 To tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
            If listleftoff(curlist) > 0 Then
                If (tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(listleftoff(curlist)).Index = EventType.evCondition Or tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(listleftoff(curlist)).Index = EventType.evShowChoices) And conditionalstage(curlist) <> 0 Then
                    i = listleftoff(curlist)
                ElseIf listleftoff(curlist) >= i Then
                    i = listleftoff(curlist) + 1
                End If
            End If
            If i <= tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then
                If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Index = EventType.evCondition Then
                    X = X + 1
                    Select Case conditionalstage(curlist)
                        Case 0
                            ReDim Preserve EventList(X)
                            EventList(X).CommandList = curlist
                            EventList(X).CommandNum = i
                            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Condition
                                Case 0
                                    Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.data2
                                        Case 0
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] == " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data3
                                        Case 1
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] >= " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data3
                                        Case 2
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] <= " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data3
                                        Case 3
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] > " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data3
                                        Case 4
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] < " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data3
                                        Case 5
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] != " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data3
                                    End Select
                                Case 1
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Switch [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Switches(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] == " & "True"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Switch [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Switches(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] == " & "False"
                                    End If
                                Case 2
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Has Item [" & Trim$(Item(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1).Name) & "] x" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.data2
                                Case 3
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Class Is [" & Trim$(Class(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1).Name) & "]"
                                Case 4
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Knows Skill [" & Trim$(spell(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1).Name) & "]"
                                Case 5
                                    Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.data2
                                        Case 0
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is == " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                        Case 1
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is >= " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                        Case 2
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is <= " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                        Case 3
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is > " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                        Case 4
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is < " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                        Case 5
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is NOT " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                    End Select
                                Case 6
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.data2 = 0 Then
                                        Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                            Case 0
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [A] == " & "True"
                                            Case 1
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [B] == " & "True"
                                            Case 2
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [C] == " & "True"
                                            Case 3
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [D] == " & "True"
                                        End Select
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.data2 = 1 Then
                                        Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                            Case 0
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [A] == " & "False"
                                            Case 1
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [B] == " & "False"
                                            Case 2
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [C] == " & "False"
                                            Case 3
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [D] == " & "False"
                                        End Select
                                    End If
                                Case 7
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.data2 = 0 Then
                                        Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data3
                                            Case 0
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Quest [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & "] not started."
                                            Case 1
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Quest [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & "] is started."
                                            Case 2
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Quest [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & "] is completed."
                                            Case 3
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Quest [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & "] can be started."
                                            Case 4
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Quest [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & "] can be ended. (All tasks complete)"
                                        End Select
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.data2 = 1 Then
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Quest [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & "] in progress and on task #" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data3
                                    End If
                            End Select
                                                    indent = indent & "       "
                            listleftoff(curlist) = i
                            conditionalstage(curlist) = 1
                            curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.CommandList
                            GoTo newlist
                        Case 1
                            ReDim Preserve EventList(X)
                            EventList(X).CommandList = curlist
                            EventList(X).CommandNum = 0
                            frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "Else"
                            listleftoff(curlist) = i
                            conditionalstage(curlist) = 2
                            curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.ElseCommandList
                            GoTo newlist
                        Case 2
                            ReDim Preserve EventList(X)
                            EventList(X).CommandList = curlist
                            EventList(X).CommandNum = 0
                            frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "End Branch"
                            indent = Mid(indent, 1, Len(indent) - 7)
                            listleftoff(curlist) = i
                            conditionalstage(curlist) = 0
                    End Select
                ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Index = EventType.evShowChoices Then
                    X = X + 1
                    Select Case conditionalstage(curlist)
                        Case 0
                            ReDim Preserve EventList(X)
                            EventList(X).CommandList = curlist
                            EventList(X).CommandNum = i
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data5 > 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Show Choices - Prompt: " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "... - Face: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data5
                            Else
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Show Choices - Prompt: " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "... - No Face"
                            End If
                            indent = indent & "       "
                            listleftoff(curlist) = i
                            conditionalstage(curlist) = 1
                            GoTo newlist
                        Case 1
                            If Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text2) <> "" Then
                                ReDim Preserve EventList(X)
                                EventList(X).CommandList = curlist
                                EventList(X).CommandNum = 0
                                frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "When [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text2) & "]"
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 2
                                curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1
                                GoTo newlist
                            Else
                                X = X - 1
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 2
                                curlist = curlist
                                GoTo newlist
                            End If
                        Case 2
                            If Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text3) <> "" Then
                                ReDim Preserve EventList(X)
                                EventList(X).CommandList = curlist
                                EventList(X).CommandNum = 0
                                frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "When [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text3) & "]"
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 3
                                curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2
                                GoTo newlist
                            Else
                                X = X - 1
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 3
                                curlist = curlist
                                GoTo newlist
                            End If
                        Case 3
                            If Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text4) <> "" Then
                                ReDim Preserve EventList(X)
                                EventList(X).CommandList = curlist
                                EventList(X).CommandNum = 0
                                frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "When [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text4) & "]"
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 4
                                curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3
                                GoTo newlist
                            Else
                                X = X - 1
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 4
                                curlist = curlist
                                GoTo newlist
                            End If
                        Case 4
                            If Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text5) <> "" Then
                                ReDim Preserve EventList(X)
                                EventList(X).CommandList = curlist
                                EventList(X).CommandNum = 0
                                frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "When [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text5) & "]"
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 5
                                curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data4
                                GoTo newlist
                            Else
                                X = X - 1
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 5
                                curlist = curlist
                                GoTo newlist
                            End If
                        Case 5
                            ReDim Preserve EventList(X)
                            EventList(X).CommandList = curlist
                            EventList(X).CommandNum = 0
                            frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "Branch End"
                            indent = Mid(indent, 1, Len(indent) - 7)
                            listleftoff(curlist) = i
                            conditionalstage(curlist) = 0
                    End Select
                Else
                    X = X + 1
                    ReDim Preserve EventList(X)
                    EventList(X).CommandList = curlist
                    EventList(X).CommandNum = i
                    Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Index
                        Case EventType.evAddText
                            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2
                                Case 0
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Add Text - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "... - Color: " & GetColorString(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & " - Chat Type: Player"
                                Case 1
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Add Text - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "... - Color: " & GetColorString(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & " - Chat Type: Map"
                                Case 2
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Add Text - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "... - Color: " & GetColorString(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & " - Chat Type: Global"
                            End Select
                        Case EventType.evShowText
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Show Text - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "... - No Face"
                            Else
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Show Text - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "... - Face: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1
                            End If
                        Case EventType.evPlayerVar
                            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2
                                Case 0
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & "] == " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3
                                Case 1
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & "] + " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3
                                Case 2
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & "] - " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3
                                Case 3
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & "] Random Between " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & " and " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data4
                            End Select
                        Case EventType.evPlayerSwitch
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Switch [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & ". " & Switches(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & "] == True"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Switch [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & ". " & Switches(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & "] == False"
                            End If
                        Case EventType.evSelfSwitch
                            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1
                                Case 0
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [A] to ON"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [A] to OFF"
                                    End If
                                Case 1
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [B] to ON"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [B] to OFF"
                                    End If
                                Case 2
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [C] to ON"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [C] to OFF"
                                    End If
                                Case 3
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [D] to ON"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [D] to OFF"
                                    End If
                            End Select
                        Case EventType.evExitProcess
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Exit Event Processing"
                                            Case EventType.evChangeItems
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Item Amount of [" & Trim$(Item(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name) & "] to " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Give Player " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & " " & Trim$(Item(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name) & "(s)"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2 = 2 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Take " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & " " & Trim$(Item(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name) & "(s) from Player."
                            End If
                                                Case EventType.evRestoreHP
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Restore Player HP"
                        Case EventType.evRestoreMP
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Restore Player MP"
                        Case EventType.evLevelUp
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Level Up Player"
                        Case EventType.evChangeLevel
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Level to " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1
                        Case EventType.evChangeSkills
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Teach Player Skill [" & Trim$(spell(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name) & "]"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Remove Player Skill [" & Trim$(spell(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name) & "]"
                            End If
                        Case EventType.evChangeClass
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Class to " & Trim$(Class(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name)
                        Case EventType.evChangeSprite
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Sprite to " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1
                        Case EventType.evChangeSex
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Sex to Male."
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Sex to Female."
                            End If
                        Case EventType.evChangePK
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player PK to No."
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player PK to Yes."
                            End If
                        Case EventType.evWarpPlayer
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data4 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Warp Player To Map: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & ") while retaining direction."
                            Else
                                Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data4 - 1
                                    Case DIR_UP
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Warp Player To Map: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & ") facing upward."
                                    Case DIR_DOWN
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Warp Player To Map: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & ") facing downward."
                                    Case DIR_LEFT
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Warp Player To Map: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & ") facing left."
                                    Case DIR_RIGHT
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Warp Player To Map: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & ") facing right."
                                End Select
                            End If
                        Case EventType.evSetMoveRoute
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 <= Map.EventCount Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Move Route for Event #" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " [" & Trim$(Map.Events(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name) & "]"
                            Else
                               frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Move Route for COULD NOT FIND EVENT!"
                            End If
                        Case EventType.evPlayAnimation
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Play Animation " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " [" & Trim$(Animation(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name) & "]" & " on Player"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Play Animation " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " [" & Trim$(Animation(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name) & "]" & " on Event #" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & " [" & Trim$(Map.Events(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3).Name) & "]"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2 = 2 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Play Animation " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " [" & Trim$(Animation(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name) & "]" & " on Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data4 & ")"
                            End If
                        Case EventType.evCustomScript
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Execute Custom Script Case: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1
                        Case EventType.evPlayBGM
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Play BGM [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1) & "]"
                        Case EventType.evFadeoutBGM
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Fadeout BGM"
                        Case EventType.evPlaySound
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Play Sound [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1) & "]"
                        Case EventType.evStopSound
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Stop Sound"
                        Case EventType.evOpenBank
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Open Bank"
                        Case EventType.evOpenMail
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Open Mail Box"
                        Case EventType.evOpenShop
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Open Shop [" & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & ". " & Trim$(Shop(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name) & "]"
                        Case EventType.evSetAccess
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Access [" & frmEditor_Events.cmbSetAccess.List(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & "]"
                        Case EventType.evGiveExp
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Give Player " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " Experience."
                        Case EventType.evShowChatBubble
                            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1
                                Case TARGET_TYPE_PLAYER
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Show Chat Bubble - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "... - On Player"
                                Case TARGET_TYPE_NPC
                                    If Map.Npc(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2) <= 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Show Chat Bubble - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "... - On NPC [" & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2) & ". ]"
                                    Else
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Show Chat Bubble - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "... - On NPC [" & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2) & ". " & Trim$(Npc(Map.Npc(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2)).Name) & "]"
                                    End If
                                Case TARGET_TYPE_EVENT
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Show Chat Bubble - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "... - On Event [" & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2) & ". " & Trim$(Map.Events(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2).Name) & "]"
                            End Select
                        Case EventType.evLabel
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Label: [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1) & "]"
                        Case EventType.evGotoLabel
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Jump to Label: [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1) & "]"
                        Case EventType.evSpawnNpc
                            If Map.Npc(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) <= 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Spawn NPC: [" & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & ". " & "]"
                            Else
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Spawn NPC: [" & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & ". " & Trim$(Npc(Map.Npc(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1)).Name) & "]"
                            End If
                        Case EventType.evFadeIn
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Fade In"
                        Case EventType.evFadeOut
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Fade Out"
                        Case EventType.evFlashWhite
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Flash White"
                        Case EventType.evSetFog
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Fog [Fog: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & " Speed: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2) & " Opacity: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3) & "]"
                        Case EventType.evSetWeather
                            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1
                                Case WEATHER_TYPE_NONE
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Weather [None]"
                                Case WEATHER_TYPE_RAIN
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Weather [Rain - Intensity: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2) & "]"
                                Case WEATHER_TYPE_SNOW
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Weather [Snow - Intensity: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2) & "]"
                                Case WEATHER_TYPE_SANDSTORM
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Weather [Sand Storm - Intensity: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2) & "]"
                                Case WEATHER_TYPE_STORM
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Weather [Storm - Intensity: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2) & "]"
                            End Select
                        Case EventType.evSetTint
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Map Tint RGBA [" & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & "," & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2) & "," & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3) & "," & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data4) & "]"
                        Case EventType.evWait
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Wait " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & " Ms"
                        Case EventType.evBeginQuest
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Begin Quest: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & ". " & Trim$(quest(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name)
                        Case EventType.evEndQuest
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "End Quest: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & ". " & Trim$(quest(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name)
                        Case EventType.evQuestTask
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Complete Quest Task: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & ". " & Trim$(quest(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name) & " - Task# " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2
                        Case EventType.evShowPicture
                            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3
                                Case 1
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Show Picture " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 + 1) & ": Pic=" & str(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2) & " Top Left, X: " & str(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data4) & " Y: " & str(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data5)
                                Case 2
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Show Picture " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 + 1) & ": Pic=" & str(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2) & " Center Screen, X: " & str(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data4) & " Y: " & str(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data5)
                                Case 3
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Show Picture " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 + 1) & ": Pic=" & str(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).data2) & " On Player, X: " & str(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data4) & " Y: " & str(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data5)
                            End Select
                        Case EventType.evHidePicture
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Hide Picture " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 + 1)
                        Case EventType.evWaitMovement
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 <= Map.EventCount Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Wait for Event #" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " [" & Trim$(Map.Events(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name) & "] to complete move route."
                            Else
                               frmEditor_Events.lstCommands.AddItem indent & "@>" & "Wait for COULD NOT FIND EVENT to complete move route."
                            End If
                        Case EventType.evHoldPlayer
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Hold Player [Do not allow player to move.]"
                        Case EventType.evReleasePlayer
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Release Player [Allow player to turn and move again.]"
                        Case Else
                            'Ghost
                            X = X - 1
                            If X = -1 Then
                                ReDim EventList(0)
                            Else
                                ReDim Preserve EventList(X)
                            End If
                    End Select
                End If
            End If
        Next
        If curlist > 1 Then
            X = X + 1
            ReDim Preserve EventList(X)
            EventList(X).CommandList = curlist
            EventList(X).CommandNum = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount + 1
            frmEditor_Events.lstCommands.AddItem indent & "@> "
            curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).ParentList
            GoTo newlist
        End If
    End If
    frmEditor_Events.lstCommands.AddItem indent & "@> "
    
    Dim z As Long
    X = 0
    For i = 0 To frmEditor_Events.lstCommands.ListCount - 1
        X = frmEditor_Events.TextWidth(frmEditor_Events.lstCommands.List(i))
        If X > z Then z = X
    Next
    
    ScrollCommands (z)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EventListCommands", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub ScrollCommands(size As Integer)

   On Error GoTo errorhandler

  Call SendMessage(frmEditor_Events.lstCommands.hwnd, LB_SETHORIZONTALEXTENT, (size) + 6, 0&)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ScrollCommands", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ListCommandAdd(s As String)
Static X As Long

   On Error GoTo errorhandler

    frmEditor_Events.lstCommands.AddItem s
    ' scrollbar
    If X < frmEditor_Events.TextWidth(s & "  ") Then
       X = frmEditor_Events.TextWidth(s & "  ")
      If frmEditor_Events.ScaleMode = vbTwips Then X = X / Screen.TwipsPerPixelX ' if twips change to pixels
      SendMessageByNum frmEditor_Events.lstCommands.hwnd, LB_SETHORIZONTALEXTENT, X, 0
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ListCommandAdd", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub AddCommand(Index As Long)
    Dim curlist As Long, i As Long, X As Long, curslot As Long, p As Long, oldCommandList As CommandListRec

   On Error GoTo errorhandler

    If tmpEvent.Pages(curPageNum).CommandListCount = 0 Then
        tmpEvent.Pages(curPageNum).CommandListCount = 1
        ReDim tmpEvent.Pages(curPageNum).CommandList(1)
    End If
    If frmEditor_Events.lstCommands.ListIndex = frmEditor_Events.lstCommands.ListCount - 1 Then
        curlist = 1
    Else
        curlist = EventList(frmEditor_Events.lstCommands.ListIndex).CommandList
    End If
        If tmpEvent.Pages(curPageNum).CommandListCount = 0 Then
        tmpEvent.Pages(curPageNum).CommandListCount = 1
        ReDim tmpEvent.Pages(curPageNum).CommandList(curlist)
    End If
    oldCommandList = tmpEvent.Pages(curPageNum).CommandList(curlist)
    tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount + 1
    p = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
    If p <= 0 Then
        ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(0)
    Else
        ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(1 To p)
        tmpEvent.Pages(curPageNum).CommandList(curlist).ParentList = oldCommandList.ParentList
        tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = p
        For i = 1 To p - 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i) = oldCommandList.Commands(i)
        Next
    End If
    If frmEditor_Events.lstCommands.ListIndex = frmEditor_Events.lstCommands.ListCount - 1 Then
        curslot = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
    Else
        i = EventList(frmEditor_Events.lstCommands.ListIndex).CommandNum
        If i < tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then
            For X = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount - 1 To i Step -1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(X + 1) = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(X)
            Next
            curslot = EventList(frmEditor_Events.lstCommands.ListIndex).CommandNum
        Else
            curslot = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
        End If
    End If
    Select Case Index
        Case EventType.evAddText
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtAddText_Text.Text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlAddText_Colour.Value
            If frmEditor_Events.optAddText_Player.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 0
            ElseIf frmEditor_Events.optAddText_Map.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 1
            ElseIf frmEditor_Events.optAddText_Global.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 2
            End If
        Case EventType.evCondition
            'This is the part where the whole entire source goes to hell :D
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandListCount = tmpEvent.Pages(curPageNum).CommandListCount + 2
            ReDim Preserve tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount)
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.CommandList = tmpEvent.Pages(curPageNum).CommandListCount - 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.ElseCommandList = tmpEvent.Pages(curPageNum).CommandListCount
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.CommandList).ParentList = curlist
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.ElseCommandList).ParentList = curlist
            For i = 0 To 7
                If frmEditor_Events.optCondition_Index(i).Value = True Then X = i
            Next
                    Select Case X
                Case 0 'Player Var
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 0
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_PlayerVarIndex.ListIndex + 1
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.data2 = frmEditor_Events.cmbCondition_PlayerVarCompare.ListIndex
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data3 = Val(frmEditor_Events.txtCondition_PlayerVarCondition.Text)
                Case 1 'Player Switch
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 1
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_PlayerSwitch.ListIndex + 1
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.data2 = frmEditor_Events.cmbCondtion_PlayerSwitchCondition.ListIndex
                Case 2 'Has Item
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 2
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_HasItem.ListIndex + 1
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.data2 = Val(frmEditor_Events.scrlCondition_HasItem.Value)
                Case 3 'Class Is
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 3
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_ClassIs.ListIndex + 1
                Case 4 'Learnt Skill
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 4
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_LearntSkill.ListIndex + 1
                Case 5 'Level Is
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 5
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = Val(frmEditor_Events.txtCondition_LevelAmount.Text)
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.data2 = frmEditor_Events.cmbCondition_LevelCompare.ListIndex
                Case 6 'Self Switch
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 6
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_SelfSwitch.ListIndex
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.data2 = frmEditor_Events.cmbCondition_SelfSwitchCondition.ListIndex
                Case 7 'Quest Shiz
                   tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 7
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.scrlCondition_Quest.Value
                    If frmEditor_Events.optCondition_Quest(0).Value Then
                        tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.data2 = 0
                        tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data3 = frmEditor_Events.cmbCondition_General.ListIndex
                    Else
                        tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.data2 = 1
                        tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data3 = frmEditor_Events.scrlCondition_QuestTask.Value
                    End If
            End Select
        Case EventType.evShowText
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtShowText.Text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlShowTextFace.Value
        Case EventType.evShowChoices
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtChoicePrompt.Text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text2 = frmEditor_Events.txtChoices(1).Text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text3 = frmEditor_Events.txtChoices(2).Text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text4 = frmEditor_Events.txtChoices(3).Text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text5 = frmEditor_Events.txtChoices(4).Text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data5 = frmEditor_Events.scrlShowChoicesFace.Value
            tmpEvent.Pages(curPageNum).CommandListCount = tmpEvent.Pages(curPageNum).CommandListCount + 4
            ReDim Preserve tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount)
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = tmpEvent.Pages(curPageNum).CommandListCount - 3
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = tmpEvent.Pages(curPageNum).CommandListCount - 2
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = tmpEvent.Pages(curPageNum).CommandListCount - 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = tmpEvent.Pages(curPageNum).CommandListCount
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount - 3).ParentList = curlist
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount - 2).ParentList = curlist
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount - 1).ParentList = curlist
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount).ParentList = curlist
        Case EventType.evPlayerVar
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbVariable.ListIndex + 1
            For i = 0 To 3
                If frmEditor_Events.optVariableAction(i).Value = True Then
                    Exit For
                End If
            Next
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = i
            If i = 3 Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = Val(frmEditor_Events.txtVariableData(i).Text)
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = Val(frmEditor_Events.txtVariableData(i + 1).Text)
            Else
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = Val(frmEditor_Events.txtVariableData(i).Text)
            End If
        Case EventType.evPlayerSwitch
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSwitch.ListIndex + 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = frmEditor_Events.cmbPlayerSwitchSet.ListIndex
        Case EventType.evSelfSwitch
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSetSelfSwitch.ListIndex
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = frmEditor_Events.cmbSetSelfSwitchTo.ListIndex
        Case EventType.evExitProcess
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evChangeItems
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeItemIndex.ListIndex + 1
            If frmEditor_Events.optChangeItemSet.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 0
            ElseIf frmEditor_Events.optChangeItemAdd.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 1
            ElseIf frmEditor_Events.optChangeItemRemove.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 2
            End If
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = Val(frmEditor_Events.txtChangeItemsAmount.Text)
        Case EventType.evRestoreHP
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evRestoreMP
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evLevelUp
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evChangeLevel
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlChangeLevel.Value
        Case EventType.evChangeSkills
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeSkills.ListIndex + 1
            If frmEditor_Events.optChangeSkillsAdd.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 0
            ElseIf frmEditor_Events.optChangeSkillsRemove.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 1
            End If
        Case EventType.evChangeClass
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeClass.ListIndex + 1
        Case EventType.evChangeSprite
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlChangeSprite.Value
        Case EventType.evChangeSex
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            If frmEditor_Events.optChangeSexMale.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0
            ElseIf frmEditor_Events.optChangeSexFemale.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1
            End If
        Case EventType.evChangePK
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            If frmEditor_Events.optChangePKYes.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1
            ElseIf frmEditor_Events.optChangePKNo.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0
            End If
        Case EventType.evWarpPlayer
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlWPMap.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = frmEditor_Events.scrlWPX.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.scrlWPY.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = frmEditor_Events.cmbWarpPlayerDir.ListIndex
        Case EventType.evSetMoveRoute
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = ListOfEvents(frmEditor_Events.cmbEvent.ListIndex)
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = frmEditor_Events.chkIgnoreMove.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.chkRepeatRoute.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRouteCount = TempMoveRouteCount
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRoute = TempMoveRoute
        Case EventType.evPlayAnimation
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbPlayAnim.ListIndex + 1
            If frmEditor_Events.optPlayAnimPlayer.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 0
            ElseIf frmEditor_Events.optPlayAnimEvent.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.cmbPlayAnimEvent.ListIndex + 1
            ElseIf frmEditor_Events.optPlayAnimTile.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 2
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.scrlPlayAnimTileX.Value
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = frmEditor_Events.scrlPlayAnimTileY.Value
            End If
        Case EventType.evCustomScript
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlCustomScript.Value
        Case EventType.evPlayBGM
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = musicCache(frmEditor_Events.cmbPlayBGM.ListIndex + 1)
        Case EventType.evFadeoutBGM
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evPlaySound
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = soundCache(frmEditor_Events.cmbPlaySound.ListIndex + 1)
        Case EventType.evStopSound
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evOpenBank
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evOpenMail
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evOpenShop
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbOpenShop.ListIndex + 1
        Case EventType.evSetAccess
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSetAccess.ListIndex
        Case EventType.evGiveExp
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlGiveExp.Value
        Case EventType.evShowChatBubble
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtChatbubbleText.Text
            If frmEditor_Events.optChatBubbleTarget(0).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = TARGET_TYPE_PLAYER
            ElseIf frmEditor_Events.optChatBubbleTarget(1).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = TARGET_TYPE_NPC
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = frmEditor_Events.cmbChatBubbleTarget.ListIndex + 1
            ElseIf frmEditor_Events.optChatBubbleTarget(2).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = TARGET_TYPE_EVENT
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = frmEditor_Events.cmbChatBubbleTarget.ListIndex + 1
            End If
        Case EventType.evLabel
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtLabelName.Text
        Case EventType.evGotoLabel
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtGotoLabel.Text
        Case EventType.evSpawnNpc
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSpawnNPC.ListIndex + 1
        Case EventType.evFadeIn
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evFadeOut
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evFlashWhite
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evSetFog
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.ScrlFogData(0).Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = frmEditor_Events.ScrlFogData(1).Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.ScrlFogData(2).Value
        Case EventType.evSetWeather
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.CmbWeather.ListIndex
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = frmEditor_Events.scrlWeatherIntensity.Value
        Case EventType.evSetTint
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlMapTintData(0).Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = frmEditor_Events.scrlMapTintData(1).Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.scrlMapTintData(2).Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = frmEditor_Events.scrlMapTintData(3).Value
        Case EventType.evWait
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlWaitAmount.Value
        Case EventType.evBeginQuest
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbBeginQuest.ListIndex + 1
        Case EventType.evEndQuest
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbEndQuest.ListIndex + 1
        Case EventType.evQuestTask
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlCompleteQuestTaskQuest.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = frmEditor_Events.scrlCompleteQuestTask.Value
        Case EventType.evShowPicture
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbPicIndex.ListIndex
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = frmEditor_Events.scrlShowPicture.Value
            If frmEditor_Events.optPic(1).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = 1
            ElseIf frmEditor_Events.optPic(2).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = 2
            Else
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = 3
            End If
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = Val(frmEditor_Events.txtPicOffset(1).Text)
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data5 = Val(frmEditor_Events.txtPicOffset(2).Text)
        Case EventType.evHidePicture
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbHidePic.ListIndex
        Case EventType.evWaitMovement
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = ListOfEvents(frmEditor_Events.cmbMoveWait.ListIndex)
        Case EventType.evHoldPlayer
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evReleasePlayer
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
    End Select
    EventListCommands


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "AddCommand", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub EditEventCommand()
    Dim i As Long, X As Long, z As Long, curlist As Long, curslot As Long

   On Error GoTo errorhandler

    i = frmEditor_Events.lstCommands.ListIndex
    If i = -1 Then Exit Sub
    If i > UBound(EventList) Then Exit Sub

    curlist = EventList(i).CommandList
    curslot = EventList(i).CommandNum
    If curlist = 0 Then Exit Sub
    If curslot = 0 Then Exit Sub
    If curlist > tmpEvent.Pages(curPageNum).CommandListCount Then Exit Sub
    If curslot > tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then Exit Sub
    Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index
        Case EventType.evAddText
            isEdit = True
            frmEditor_Events.txtAddText_Text.Text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1
            frmEditor_Events.scrlAddText_Colour.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2
                Case 0
                    frmEditor_Events.optAddText_Player.Value = True
                Case 1
                    frmEditor_Events.optAddText_Map.Value = True
                Case 2
                    frmEditor_Events.optAddText_Global.Value = True
            End Select
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(2).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evCondition
            isEdit = True
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(7).Visible = True
            frmEditor_Events.fraCommands.Visible = False
            frmEditor_Events.ClearConditionFrame
            frmEditor_Events.optCondition_Index(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition).Value = True
            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition
                Case 0
                    frmEditor_Events.cmbCondition_PlayerVarIndex.Enabled = True
                    frmEditor_Events.cmbCondition_PlayerVarCompare.Enabled = True
                    frmEditor_Events.txtCondition_PlayerVarCondition.Enabled = True
                    frmEditor_Events.cmbCondition_PlayerVarIndex.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 - 1
                    frmEditor_Events.cmbCondition_PlayerVarCompare.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.data2
                    frmEditor_Events.txtCondition_PlayerVarCondition.Text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data3
                Case 1
                    frmEditor_Events.cmbCondition_PlayerSwitch.Enabled = True
                    frmEditor_Events.cmbCondtion_PlayerSwitchCondition.Enabled = True
                    frmEditor_Events.cmbCondition_PlayerSwitch.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 - 1
                    frmEditor_Events.cmbCondtion_PlayerSwitchCondition.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.data2
                Case 2
                    frmEditor_Events.cmbCondition_HasItem.Enabled = True
                    frmEditor_Events.scrlCondition_HasItem.Enabled = True
                    frmEditor_Events.cmbCondition_HasItem.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 - 1
                    frmEditor_Events.scrlCondition_HasItem.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.data2
                Case 3
                    frmEditor_Events.cmbCondition_ClassIs.Enabled = True
                    frmEditor_Events.cmbCondition_ClassIs.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 - 1
                Case 4
                    frmEditor_Events.cmbCondition_LearntSkill.Enabled = True
                    frmEditor_Events.cmbCondition_LearntSkill.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 - 1
                Case 5
                    frmEditor_Events.cmbCondition_LevelCompare.Enabled = True
                    frmEditor_Events.txtCondition_LevelAmount.Enabled = True
                    frmEditor_Events.txtCondition_LevelAmount.Text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1
                    frmEditor_Events.cmbCondition_LevelCompare.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.data2
                Case 6
                    frmEditor_Events.cmbCondition_SelfSwitch.Enabled = True
                    frmEditor_Events.cmbCondition_SelfSwitchCondition.Enabled = True
                    frmEditor_Events.cmbCondition_SelfSwitch.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1
                    frmEditor_Events.cmbCondition_SelfSwitchCondition.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.data2
                Case 7
                    frmEditor_Events.scrlCondition_Quest.Enabled = True
                    frmEditor_Events.scrlCondition_Quest.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1
                    frmEditor_Events.fraConditions_Quest.Visible = True
                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.data2 = 0 Then
                        frmEditor_Events.optCondition_Quest(0).Value = True
                        frmEditor_Events.cmbCondition_General.Enabled = True
                        frmEditor_Events.cmbCondition_General.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data3
                        frmEditor_Events.lblConditionQuest.Caption = "Quest: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data3
                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.data2 = 1 Then
                        frmEditor_Events.optCondition_Quest(1).Value = True
                        frmEditor_Events.scrlCondition_QuestTask.Enabled = True
                        frmEditor_Events.scrlCondition_QuestTask.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data3
                        frmEditor_Events.lblCondition_QuestTask.Caption = "#" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data3
                    End If
            End Select
        Case EventType.evShowText
            isEdit = True
            frmEditor_Events.txtShowText.Text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1
            frmEditor_Events.scrlShowTextFace.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(0).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evShowChoices
            isEdit = True
            frmEditor_Events.txtChoicePrompt.Text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1
            frmEditor_Events.txtChoices(1).Text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text2
            frmEditor_Events.txtChoices(2).Text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text3
            frmEditor_Events.txtChoices(3).Text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text4
            frmEditor_Events.txtChoices(4).Text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text5
            frmEditor_Events.scrlShowChoicesFace.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data5
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(1).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evPlayerVar
            isEdit = True
            frmEditor_Events.cmbVariable.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2
                Case 0
                    frmEditor_Events.optVariableAction(0).Value = True
                    frmEditor_Events.txtVariableData(0).Text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
                Case 1
                    frmEditor_Events.optVariableAction(1).Value = True
                    frmEditor_Events.txtVariableData(1).Text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
                Case 2
                    frmEditor_Events.optVariableAction(2).Value = True
                    frmEditor_Events.txtVariableData(2).Text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
                Case 3
                    frmEditor_Events.optVariableAction(3).Value = True
                    frmEditor_Events.txtVariableData(3).Text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
                    frmEditor_Events.txtVariableData(4).Text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4
            End Select
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(4).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evPlayerSwitch
            isEdit = True
            frmEditor_Events.cmbSwitch.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            frmEditor_Events.cmbPlayerSwitchSet.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(5).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSelfSwitch
            isEdit = True
            frmEditor_Events.cmbSetSelfSwitch.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.cmbSetSelfSwitchTo.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(6).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeItems
            isEdit = True
            frmEditor_Events.cmbChangeItemIndex.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 0 Then
                frmEditor_Events.optChangeItemSet.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 1 Then
                frmEditor_Events.optChangeItemAdd.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 2 Then
                frmEditor_Events.optChangeItemRemove.Value = True
            End If
            frmEditor_Events.txtChangeItemsAmount.Text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(10).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeLevel
            isEdit = True
            frmEditor_Events.scrlChangeLevel.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(11).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeSkills
            isEdit = True
            frmEditor_Events.cmbChangeSkills.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 0 Then
                frmEditor_Events.optChangeSkillsAdd.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 1 Then
                frmEditor_Events.optChangeSkillsRemove.Value = True
            End If
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(12).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeClass
            isEdit = True
            frmEditor_Events.cmbChangeClass.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(13).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeSprite
            isEdit = True
            frmEditor_Events.scrlChangeSprite.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(14).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeSex
            isEdit = True
            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0 Then
                frmEditor_Events.optChangeSexMale.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1 Then
                frmEditor_Events.optChangeSexFemale.Value = True
            End If
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(15).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangePK
            isEdit = True
            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1 Then
                frmEditor_Events.optChangePKYes.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0 Then
                frmEditor_Events.optChangePKNo.Value = True
            End If
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(16).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evWarpPlayer
            isEdit = True
            frmEditor_Events.scrlWPMap.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.scrlWPX.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2
            frmEditor_Events.scrlWPY.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
            frmEditor_Events.cmbWarpPlayerDir.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(18).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSetMoveRoute
            isEdit = True
            frmEditor_Events.fraMoveRoute.Visible = True
            frmEditor_Events.lstMoveRoute.Clear
            frmEditor_Events.cmbEvent.Clear
            ReDim ListOfEvents(0 To Map.EventCount)
            ListOfEvents(0) = EditorEvent
            frmEditor_Events.cmbEvent.AddItem "This Event"
            frmEditor_Events.cmbEvent.ListIndex = 0
            frmEditor_Events.cmbEvent.Enabled = True
            For i = 1 To Map.EventCount
                If i <> EditorEvent Then
                    frmEditor_Events.cmbEvent.AddItem Trim$(Map.Events(i).Name)
                    X = X + 1
                    ListOfEvents(X) = i
                    If i = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 Then frmEditor_Events.cmbEvent.ListIndex = X
                End If
            Next
            
            IsMoveRouteCommand = True
            frmEditor_Events.chkIgnoreMove.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2
            frmEditor_Events.chkRepeatRoute.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
            TempMoveRouteCount = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRouteCount
            TempMoveRoute = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRoute
            For i = 1 To TempMoveRouteCount
                Select Case TempMoveRoute(i).Index
                    Case 1
                        frmEditor_Events.lstMoveRoute.AddItem "Move Up"
                    Case 2
                        frmEditor_Events.lstMoveRoute.AddItem "Move Down"
                    Case 3
                        frmEditor_Events.lstMoveRoute.AddItem "Move Left"
                    Case 4
                        frmEditor_Events.lstMoveRoute.AddItem "Move Right"
                    Case 5
                        frmEditor_Events.lstMoveRoute.AddItem "Move Randomly"
                    Case 6
                        frmEditor_Events.lstMoveRoute.AddItem "Move Towards Player"
                    Case 7
                        frmEditor_Events.lstMoveRoute.AddItem "Move Away From Player"
                    Case 8
                        frmEditor_Events.lstMoveRoute.AddItem "Step Forward"
                    Case 9
                        frmEditor_Events.lstMoveRoute.AddItem "Step Back"
                    Case 10
                        frmEditor_Events.lstMoveRoute.AddItem "Wait 100ms"
                    Case 11
                        frmEditor_Events.lstMoveRoute.AddItem "Wait 500ms"
                    Case 12
                        frmEditor_Events.lstMoveRoute.AddItem "Wait 1000ms"
                    Case 13
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Up"
                    Case 14
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Down"
                    Case 15
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Left"
                    Case 16
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Right"
                    Case 17
                        frmEditor_Events.lstMoveRoute.AddItem "Turn 90 Degrees To the Right"
                    Case 18
                        frmEditor_Events.lstMoveRoute.AddItem "Turn 90 Degrees To the Left"
                    Case 19
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Around 180 Degrees"
                    Case 20
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Randomly"
                    Case 21
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Towards Player"
                    Case 22
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Away from Player"
                    Case 23
                        frmEditor_Events.lstMoveRoute.AddItem "Set Speed 8x Slower"
                    Case 24
                        frmEditor_Events.lstMoveRoute.AddItem "Set Speed 4x Slower"
                    Case 25
                        frmEditor_Events.lstMoveRoute.AddItem "Set Speed 2x Slower"
                    Case 26
                        frmEditor_Events.lstMoveRoute.AddItem "Set Speed to Normal"
                    Case 27
                        frmEditor_Events.lstMoveRoute.AddItem "Set Speed 2x Faster"
                    Case 28
                        frmEditor_Events.lstMoveRoute.AddItem "Set Speed 4x Faster"
                    Case 29
                        frmEditor_Events.lstMoveRoute.AddItem "Set Frequency Lowest"
                    Case 30
                        frmEditor_Events.lstMoveRoute.AddItem "Set Frequency Lower"
                    Case 31
                        frmEditor_Events.lstMoveRoute.AddItem "Set Frequency Normal"
                    Case 32
                        frmEditor_Events.lstMoveRoute.AddItem "Set Frequency Higher"
                    Case 33
                        frmEditor_Events.lstMoveRoute.AddItem "Set Frequency Highest"
                    Case 34
                        frmEditor_Events.lstMoveRoute.AddItem "Turn On Walking Animation"
                    Case 35
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Off Walking Animation"
                    Case 36
                        frmEditor_Events.lstMoveRoute.AddItem "Turn On Fixed Direction"
                    Case 37
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Off Fixed Direction"
                    Case 38
                        frmEditor_Events.lstMoveRoute.AddItem "Turn On Walk Through"
                    Case 39
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Off Walk Through"
                    Case 40
                        frmEditor_Events.lstMoveRoute.AddItem "Set Position Below Player"
                    Case 41
                        frmEditor_Events.lstMoveRoute.AddItem "Set Position at Player Level"
                    Case 42
                        frmEditor_Events.lstMoveRoute.AddItem "Set Position Above Player"
                    Case 43
                        frmEditor_Events.lstMoveRoute.AddItem "Set Graphic"
                End Select
            Next
                        frmEditor_Events.fraMoveRoute.Width = 841
            frmEditor_Events.fraMoveRoute.Height = 609
            frmEditor_Events.fraMoveRoute.Visible = True
                    frmEditor_Events.fraDialogue.Visible = False
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evPlayAnimation
            isEdit = True
            frmEditor_Events.lblPlayAnimX.Visible = False
            frmEditor_Events.lblPlayAnimY.Visible = False
            frmEditor_Events.scrlPlayAnimTileX.Visible = False
            frmEditor_Events.scrlPlayAnimTileY.Visible = False
            frmEditor_Events.cmbPlayAnimEvent.Visible = False
            frmEditor_Events.cmbPlayAnim.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            frmEditor_Events.cmbPlayAnimEvent.Clear
            For i = 1 To Map.EventCount
                frmEditor_Events.cmbPlayAnimEvent.AddItem i & ". " & Trim$(Map.Events(i).Name)
            Next
            frmEditor_Events.cmbPlayAnimEvent.ListIndex = 0
            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 0 Then
                frmEditor_Events.optPlayAnimPlayer.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 1 Then
                frmEditor_Events.optPlayAnimEvent.Value = True
                frmEditor_Events.cmbPlayAnimEvent.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 - 1
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 2 Then
                frmEditor_Events.optPlayAnimTile.Value = True
                frmEditor_Events.scrlPlayAnimTileX.max = Map.MaxX
                frmEditor_Events.scrlPlayAnimTileY.max = Map.MaxY
                frmEditor_Events.scrlPlayAnimTileX.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
                frmEditor_Events.scrlPlayAnimTileY.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4
            End If
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(20).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evCustomScript
            isEdit = True
            frmEditor_Events.scrlCustomScript.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(29).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evPlayBGM
            isEdit = True
            For i = 1 To UBound(musicCache())
                If musicCache(i) = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 Then
                    frmEditor_Events.cmbPlayBGM.ListIndex = i - 1
                End If
            Next
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(25).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evPlaySound
            isEdit = True
            For i = 1 To UBound(soundCache())
                If soundCache(i) = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 Then
                    frmEditor_Events.cmbPlaySound.ListIndex = i - 1
                End If
            Next
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(26).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evOpenShop
            isEdit = True
            frmEditor_Events.cmbOpenShop.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(21).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSetAccess
            isEdit = True
            frmEditor_Events.cmbSetAccess.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(28).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evGiveExp
            isEdit = True
            frmEditor_Events.scrlGiveExp.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.lblGiveExp.Caption = "Give Exp: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(17).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evShowChatBubble
            isEdit = True
            frmEditor_Events.txtChatbubbleText.Text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1
            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
                Case TARGET_TYPE_PLAYER
                    frmEditor_Events.optChatBubbleTarget(0).Value = True
                Case TARGET_TYPE_NPC
                    frmEditor_Events.optChatBubbleTarget(1).Value = True
                Case TARGET_TYPE_EVENT
                    frmEditor_Events.optChatBubbleTarget(1).Value = True
            End Select
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(3).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evLabel
            isEdit = True
            frmEditor_Events.txtLabelName.Text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(8).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evGotoLabel
            isEdit = True
            frmEditor_Events.txtGotoLabel.Text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(9).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSpawnNpc
            isEdit = True
            frmEditor_Events.cmbSpawnNPC.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(19).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSetFog
            isEdit = True
            frmEditor_Events.ScrlFogData(0).Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.ScrlFogData(1).Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2
            frmEditor_Events.ScrlFogData(2).Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(22).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSetWeather
            isEdit = True
            frmEditor_Events.CmbWeather.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.scrlWeatherIntensity.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(23).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSetTint
            isEdit = True
            frmEditor_Events.scrlMapTintData(0).Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.scrlMapTintData(1).Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2
            frmEditor_Events.scrlMapTintData(2).Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
            frmEditor_Events.scrlMapTintData(3).Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(24).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evWait
            isEdit = True
            frmEditor_Events.scrlWaitAmount.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(27).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evBeginQuest
            isEdit = True
            frmEditor_Events.cmbBeginQuest.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(30).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evEndQuest
            isEdit = True
            frmEditor_Events.cmbEndQuest.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(31).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evQuestTask
            isEdit = True
            frmEditor_Events.scrlCompleteQuestTaskQuest.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.scrlCompleteQuestTask.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(32).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evShowPicture
            isEdit = True
            frmEditor_Events.cmbPicIndex.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.scrlShowPicture.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2
            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = 1 Then
                frmEditor_Events.optPic(1).Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = 2 Then
                frmEditor_Events.optPic(2).Value = True
            Else
                frmEditor_Events.optPic(3).Value = True
            End If
            frmEditor_Events.txtPicOffset(1).Text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4
            frmEditor_Events.txtPicOffset(2).Text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data5
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(33).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evHidePicture
            isEdit = True
            frmEditor_Events.cmbHidePic.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(34).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evWaitMovement
            isEdit = True
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(35).Visible = True
            frmEditor_Events.fraCommands.Visible = False
            frmEditor_Events.cmbMoveWait.Clear
            ReDim ListOfEvents(0 To Map.EventCount)
            ListOfEvents(0) = EditorEvent
            frmEditor_Events.cmbMoveWait.AddItem "This Event"
            frmEditor_Events.cmbMoveWait.ListIndex = 0
            For i = 1 To Map.EventCount
                If i <> EditorEvent Then
                    frmEditor_Events.cmbMoveWait.AddItem Trim$(Map.Events(i).Name)
                    X = X + 1
                    ListOfEvents(X) = i
                    If i = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 Then frmEditor_Events.cmbMoveWait.ListIndex = X
                End If
            Next
    End Select


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EditEventCommand", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub DeleteEventCommand()
    Dim i As Long, X As Long, z As Long, curlist As Long, curslot As Long, p As Long, oldCommandList As CommandListRec

   On Error GoTo errorhandler

    i = frmEditor_Events.lstCommands.ListIndex
    If i = -1 Then Exit Sub
    If i > UBound(EventList) Then Exit Sub
    curlist = EventList(i).CommandList
    curslot = EventList(i).CommandNum
    If curlist = 0 Then Exit Sub
    If curslot = 0 Then Exit Sub
    If curlist > tmpEvent.Pages(curPageNum).CommandListCount Then Exit Sub
    If curslot > tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then Exit Sub
    If curslot = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then
        tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount - 1
        p = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
        If p <= 0 Then
            ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(0)
        Else
            oldCommandList = tmpEvent.Pages(curPageNum).CommandList(curlist)
            ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(p)
            X = 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).ParentList = oldCommandList.ParentList
            tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = p
            For i = 1 To p + 1
                If i <> curslot Then
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(X) = oldCommandList.Commands(i)
                    X = X + 1
                End If
            Next
        End If
    Else
        tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount - 1
        p = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
        oldCommandList = tmpEvent.Pages(curPageNum).CommandList(curlist)
        X = 1
        If p <= 0 Then
            ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(0)
        Else
            ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(p)
            tmpEvent.Pages(curPageNum).CommandList(curlist).ParentList = oldCommandList.ParentList
            tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = p
            For i = 1 To p + 1
                If i <> curslot Then
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(X) = oldCommandList.Commands(i)
                    X = X + 1
                End If
            Next
        End If
    End If
    EventListCommands


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DeleteEventCommand", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub ClearEventCommands()

   On Error GoTo errorhandler

    ReDim tmpEvent.Pages(curPageNum).CommandList(1)
    tmpEvent.Pages(curPageNum).CommandListCount = 1
    EventListCommands


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearEventCommands", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub EditCommand()
    Dim i As Long, X As Long, z As Long, curlist As Long, curslot As Long

   On Error GoTo errorhandler

    i = frmEditor_Events.lstCommands.ListIndex
    If i = -1 Then Exit Sub
    If i > UBound(EventList) Then Exit Sub

    curlist = EventList(i).CommandList
    curslot = EventList(i).CommandNum
    If curlist = 0 Then Exit Sub
    If curslot = 0 Then Exit Sub
    If curlist > tmpEvent.Pages(curPageNum).CommandListCount Then Exit Sub
    If curslot > tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then Exit Sub
    Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index
        Case EventType.evAddText
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtAddText_Text.Text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlAddText_Colour.Value
            If frmEditor_Events.optAddText_Player.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 0
            ElseIf frmEditor_Events.optAddText_Map.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 1
            ElseIf frmEditor_Events.optAddText_Global.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 2
            End If
        Case EventType.evCondition
            If frmEditor_Events.optCondition_Index(0).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 0
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_PlayerVarIndex.ListIndex + 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.data2 = frmEditor_Events.cmbCondition_PlayerVarCompare.ListIndex
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data3 = Val(frmEditor_Events.txtCondition_PlayerVarCondition.Text)
            ElseIf frmEditor_Events.optCondition_Index(1).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_PlayerSwitch.ListIndex + 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.data2 = frmEditor_Events.cmbCondtion_PlayerSwitchCondition.ListIndex
            ElseIf frmEditor_Events.optCondition_Index(2).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 2
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_HasItem.ListIndex + 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.data2 = Val(frmEditor_Events.scrlCondition_HasItem.Value)
            ElseIf frmEditor_Events.optCondition_Index(3).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 3
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_ClassIs.ListIndex + 1
            ElseIf frmEditor_Events.optCondition_Index(4).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 4
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_LearntSkill.ListIndex + 1
            ElseIf frmEditor_Events.optCondition_Index(5).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 5
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = Val(frmEditor_Events.txtCondition_LevelAmount.Text)
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.data2 = frmEditor_Events.cmbCondition_LevelCompare.ListIndex
            ElseIf frmEditor_Events.optCondition_Index(6).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 6
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_SelfSwitch.ListIndex
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.data2 = frmEditor_Events.cmbCondition_SelfSwitchCondition.ListIndex
            ElseIf frmEditor_Events.optCondition_Index(7).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 7
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.scrlCondition_Quest.Value
                If frmEditor_Events.optCondition_Quest(0).Value Then
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.data2 = 0
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data3 = frmEditor_Events.cmbCondition_General.ListIndex
                Else
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.data2 = 1
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data3 = frmEditor_Events.scrlCondition_QuestTask.Value
                End If
            End If
        Case EventType.evShowText
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtShowText.Text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlShowTextFace.Value
        Case EventType.evShowChoices
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtChoicePrompt.Text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text2 = frmEditor_Events.txtChoices(1).Text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text3 = frmEditor_Events.txtChoices(2).Text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text4 = frmEditor_Events.txtChoices(3).Text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text5 = frmEditor_Events.txtChoices(4).Text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data5 = frmEditor_Events.scrlShowChoicesFace.Value
        Case EventType.evPlayerVar
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbVariable.ListIndex + 1
            For i = 0 To 3
                If frmEditor_Events.optVariableAction(i).Value = True Then
                    Exit For
                End If
            Next
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = i
            If i = 3 Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = Val(frmEditor_Events.txtVariableData(i).Text)
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = Val(frmEditor_Events.txtVariableData(i + 1).Text)
            Else
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = Val(frmEditor_Events.txtVariableData(i).Text)
            End If
        Case EventType.evPlayerSwitch
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSwitch.ListIndex + 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = frmEditor_Events.cmbPlayerSwitchSet.ListIndex
        Case EventType.evSelfSwitch
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSetSelfSwitch.ListIndex
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = frmEditor_Events.cmbSetSelfSwitchTo.ListIndex
        Case EventType.evChangeItems
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeItemIndex.ListIndex + 1
            If frmEditor_Events.optChangeItemSet.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 0
            ElseIf frmEditor_Events.optChangeItemAdd.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 1
            ElseIf frmEditor_Events.optChangeItemRemove.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 2
            End If
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = Val(frmEditor_Events.txtChangeItemsAmount.Text)
        Case EventType.evChangeLevel
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlChangeLevel.Value
        Case EventType.evChangeSkills
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeSkills.ListIndex + 1
            If frmEditor_Events.optChangeSkillsAdd.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 0
            ElseIf frmEditor_Events.optChangeSkillsRemove.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 1
            End If
        Case EventType.evChangeClass
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeClass.ListIndex + 1
        Case EventType.evChangeSprite
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlChangeSprite.Value
        Case EventType.evChangeSex
            If frmEditor_Events.optChangeSexMale.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0
            ElseIf frmEditor_Events.optChangeSexFemale.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1
            End If
        Case EventType.evChangePK
            If frmEditor_Events.optChangePKYes.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1
            ElseIf frmEditor_Events.optChangePKNo.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0
            End If
        Case EventType.evWarpPlayer
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlWPMap.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = frmEditor_Events.scrlWPX.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.scrlWPY.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = frmEditor_Events.cmbWarpPlayerDir.ListIndex
        Case EventType.evSetMoveRoute
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = ListOfEvents(frmEditor_Events.cmbEvent.ListIndex)
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = frmEditor_Events.chkIgnoreMove.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.chkRepeatRoute.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRouteCount = TempMoveRouteCount
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRoute = TempMoveRoute
        Case EventType.evPlayAnimation
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbPlayAnim.ListIndex + 1
            If frmEditor_Events.optPlayAnimPlayer.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 0
            ElseIf frmEditor_Events.optPlayAnimEvent.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.cmbPlayAnimEvent.ListIndex + 1
            ElseIf frmEditor_Events.optPlayAnimTile.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = 2
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.scrlPlayAnimTileX.Value
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = frmEditor_Events.scrlPlayAnimTileY.Value
            End If
        Case EventType.evCustomScript
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlCustomScript.Value
        Case EventType.evPlayBGM
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = musicCache(frmEditor_Events.cmbPlayBGM.ListIndex + 1)
        Case EventType.evPlaySound
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = soundCache(frmEditor_Events.cmbPlaySound.ListIndex + 1)
        Case EventType.evOpenShop
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbOpenShop.ListIndex + 1
        Case EventType.evSetAccess
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSetAccess.ListIndex
        Case EventType.evGiveExp
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlGiveExp.Value
        Case EventType.evShowChatBubble
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtChatbubbleText.Text
            If frmEditor_Events.optChatBubbleTarget(0).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = TARGET_TYPE_PLAYER
            ElseIf frmEditor_Events.optChatBubbleTarget(1).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = TARGET_TYPE_NPC
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = frmEditor_Events.cmbChatBubbleTarget.ListIndex + 1
            ElseIf frmEditor_Events.optChatBubbleTarget(2).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = TARGET_TYPE_EVENT
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = frmEditor_Events.cmbChatBubbleTarget.ListIndex + 1
            End If
        Case EventType.evLabel
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtLabelName.Text
        Case EventType.evGotoLabel
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtGotoLabel.Text
        Case EventType.evSpawnNpc
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSpawnNPC.ListIndex + 1
        Case EventType.evSetFog
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.ScrlFogData(0).Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = frmEditor_Events.ScrlFogData(1).Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.ScrlFogData(2).Value
        Case EventType.evSetWeather
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.CmbWeather.ListIndex
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = frmEditor_Events.scrlWeatherIntensity.Value
        Case EventType.evSetTint
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlMapTintData(0).Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = frmEditor_Events.scrlMapTintData(1).Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.scrlMapTintData(2).Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = frmEditor_Events.scrlMapTintData(3).Value
        Case EventType.evWait
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlWaitAmount.Value
        Case EventType.evBeginQuest
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbBeginQuest.ListIndex + 1
        Case EventType.evEndQuest
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbEndQuest.ListIndex + 1
        Case EventType.evQuestTask
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlCompleteQuestTaskQuest.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = frmEditor_Events.scrlCompleteQuestTask.Value
        Case EventType.evShowPicture
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbPicIndex.ListIndex
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).data2 = frmEditor_Events.scrlShowPicture.Value
            If frmEditor_Events.optPic(1).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = 1
            ElseIf frmEditor_Events.optPic(2).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = 2
            Else
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = 3
            End If
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = Val(frmEditor_Events.txtPicOffset(1).Text)
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data5 = Val(frmEditor_Events.txtPicOffset(2).Text)
        Case EventType.evHidePicture
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbHidePic.ListIndex
        Case EventType.evWaitMovement
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = ListOfEvents(frmEditor_Events.cmbMoveWait.ListIndex)
    End Select
    EventListCommands


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EditCommand", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub RequestSwitchesAndVariables()
Dim i As Long, buffer As clsBuffer

   On Error GoTo errorhandler

Set buffer = New clsBuffer
buffer.WriteLong CRequestSwitchesAndVariables
SendData buffer.ToArray
Set buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "RequestSwitchesAndVariables", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendSwitchesAndVariables()
Dim i As Long, buffer As clsBuffer

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CSwitchesAndVariables
    For i = 1 To MAX_SWITCHES
        buffer.WriteString Switches(i)
    Next
    For i = 1 To MAX_VARIABLES
        buffer.WriteString Variables(i)
    Next
    SendData buffer.ToArray
Set buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendSwitchesAndVariables", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Public Sub ZoneEditorInit()
Dim i As Long, X As Long


   On Error GoTo errorhandler

    If frmEditor_Zone.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Zone.lstIndex.ListIndex + 1
    i = EditorIndex
    With MapZones(EditorIndex)
        frmEditor_Zone.cmbNpc.ListIndex = 0
        frmEditor_Zone.txtZoneName.Text = Trim$(.Name)
        frmEditor_Zone.txtAddMap.Text = "0"
        For X = 1 To 5
            frmEditor_Zone.scrlWeather(X).Value = .Weather(X)
        Next
        frmEditor_Zone.cmbNpc.ListIndex = 0
        frmEditor_Zone.scrlWeatherIntensity.Value = .WeatherIntensity

        frmEditor_Zone.lstNpcs.Clear
        For X = 1 To MAX_MAP_NPCS * 2
            If MapZones(i).NPCs(X) > 0 Then
                frmEditor_Zone.lstNpcs.AddItem CStr(MapZones(i).NPCs(X)) & ". " & Trim$(Npc(MapZones(i).NPCs(X)).Name)
            Else
                frmEditor_Zone.lstNpcs.AddItem "No NPC"
            End If
        Next
            frmEditor_Zone.lstMaps.Clear
        If MapZones(i).MapCount > 0 Then
            For X = 1 To MapZones(i).MapCount
                frmEditor_Zone.lstMaps.AddItem "Map #" & MapZones(i).Maps(X)
            Next
        End If
            EditorIndex = frmEditor_Zone.lstIndex.ListIndex + 1
    End With
    Zone_Changed(EditorIndex) = True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ZoneEditorInit", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub ZoneEditorCancel()


   On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Zone
    ClearChanged_Zone




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ZoneEditorCancel", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub ZoneEditorOk()
Dim i As Long, buffer As clsBuffer, count As Long, X As Long


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CSaveZones
    For i = 1 To MAX_ZONES
        If Zone_Changed(i) Then count = count + 1
    Next
    buffer.WriteLong count
    If count > 0 Then
        For i = 1 To MAX_ZONES
            If Zone_Changed(i) Then
                buffer.WriteLong i
                buffer.WriteString Trim$(MapZones(i).Name)
                buffer.WriteLong MapZones(i).MapCount
                If MapZones(i).MapCount > 0 Then
                    For X = 1 To MapZones(i).MapCount
                        buffer.WriteLong MapZones(i).Maps(X)
                    Next
                End If
                For X = 1 To MAX_MAP_NPCS * 2
                    buffer.WriteLong MapZones(i).NPCs(X)
                Next
                For X = 1 To 5
                    buffer.WriteByte MapZones(i).Weather(X)
                Next
                buffer.WriteByte MapZones(i).WeatherIntensity
            End If
        Next
    End If
    SendData buffer.ToArray
    Set buffer = Nothing
    Unload frmEditor_Zone
    Editor = 0
    ClearChanged_Zone




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ZoneEditorOk", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub HouseEditorInit()
Dim i As Long, X As Long


   On Error GoTo errorhandler

    If frmEditor_House.Visible = False Then Exit Sub
    EditorIndex = frmEditor_House.lstIndex.ListIndex + 1
    i = EditorIndex
    With House(EditorIndex)
        frmEditor_House.txtName.Text = Trim$(.ConfigName)
        frmEditor_House.txtBaseMap.Text = Trim$(CStr(.BaseMap))
        frmEditor_House.txtXEntrance.Text = Trim$(CStr(.X))
        frmEditor_House.txtYEntrance.Text = Trim$(CStr(.Y))
        frmEditor_House.txtHousePrice.Text = Trim$(CStr(.Price))
        frmEditor_House.txtHouseFurniture.Text = Trim$(CStr(.MaxFurniture))
    End With
    House_Changed(EditorIndex) = True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ZoneEditorInit", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
Public Sub HouseEditorCancel()


   On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_House
    ClearChanged_House




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HouseEditorCancel", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub HouseEditorOk()
Dim i As Long, buffer As clsBuffer, count As Long, X As Long


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CSaveHouses
    For i = 1 To MAX_HOUSES
        If House_Changed(i) Then count = count + 1
    Next
    buffer.WriteLong count
    If count > 0 Then
        For i = 1 To MAX_HOUSES
            If House_Changed(i) Then
                buffer.WriteLong i
                buffer.WriteString Trim$(House(i).ConfigName)
                buffer.WriteLong House(i).BaseMap
                buffer.WriteLong House(i).X
                buffer.WriteLong House(i).Y
                buffer.WriteLong House(i).Price
                buffer.WriteLong House(i).MaxFurniture
            End If
        Next
    End If
    SendData buffer.ToArray
    Set buffer = Nothing
    Unload frmEditor_House
    Editor = 0
    ClearChanged_House
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HouseEditorOk", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub InitPlayerEditor()
Dim i As Long

   On Error GoTo errorhandler

    With frmAdmin
        .picEditPlayer.Visible = True
        .fraCharList.Visible = False
        .fraEditPlayer.Visible = True
        .txtLogin.Text = Trim$(frmAdmin.EditingAccount)
        .txtCharName.Text = Trim$(Player(0).Name)
        .cmbSex.ListIndex = Player(0).Sex
        .cmbClass.Clear
        If Max_Classes > 0 Then
            For i = 1 To Max_Classes
                .cmbClass.AddItem Trim$(Class(i).Name)
            Next
            .cmbClass.ListIndex = Player(0).Class - 1
        End If
        .txtLevel.Text = CStr(Player(0).Level)
        .txtExp.Text = CStr(Player(0).Exp)
        .cmbAccess.ListIndex = Player(0).Access
        .cmbPK.ListIndex = Player(0).PK
        .txtHP.Text = Player(0).Vital(Vitals.HP)
        .txtMP.Text = Player(0).Vital(Vitals.MP)
        .txtStrength.Text = Player(0).stat(Stats.Strength)
        .txtEndurance.Text = Player(0).stat(Stats.Endurance)
        .txtIntelligence.Text = Player(0).stat(Stats.Intelligence)
        .txtAgility.Text = Player(0).stat(Stats.Agility)
        .txtWillPower.Text = Player(0).stat(Stats.Willpower)
        .txtPoints.Text = Player(0).Points
        .cmbWeapon.Clear
        .cmbArmor.Clear
        .cmbHelmet.Clear
        .cmbShield.Clear
        .cmbWeapon.AddItem "None."
        .cmbArmor.AddItem "None."
        .cmbHelmet.AddItem "None."
        .cmbShield.AddItem "None."
        .cmbItems.AddItem "None."
        For i = 1 To MAX_ITEMS
            .cmbWeapon.AddItem i & ". " & Trim$(Item(i).Name)
            .cmbArmor.AddItem i & ". " & Trim$(Item(i).Name)
            .cmbHelmet.AddItem i & ". " & Trim$(Item(i).Name)
            .cmbShield.AddItem i & ". " & Trim$(Item(i).Name)
            .cmbItems.AddItem i & ". " & Trim$(Item(i).Name)
        Next
        .cmbWeapon.ListIndex = Player(0).Equipment(Equipment.Weapon)
        .cmbShield.ListIndex = Player(0).Equipment(Equipment.Shield)
        .cmbArmor.ListIndex = Player(0).Equipment(Equipment.Armor)
        .cmbHelmet.ListIndex = Player(0).Equipment(Equipment.Helmet)
        .txtMap.Text = Player(0).Map
        .txtX.Text = Player(0).X
        .txtY.Text = Player(0).Y
        .cmbDir.ListIndex = Player(0).dir
        .cmbSpells.Clear
        .cmbSpells.AddItem "None"
        For i = 1 To MAX_SPELLS
            .cmbSpells.AddItem i & ". " & Trim$(spell(i).Name)
        Next
        
        InitPlayerSpellEditor
        .lstSpells.ListIndex = 0
        InitPlayerItemEditor
        .cmbInvSlot.ListIndex = 0
        
    End With
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "InitPlayerEditor", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub InitPlayerSpellEditor()
Dim i As Long, X As Long

   On Error GoTo errorhandler

    With frmAdmin
        X = .lstSpells.ListIndex
        .lstSpells.Clear
        For i = 1 To MAX_PLAYER_SPELLS
            If TempPlayerSpells(i) = 0 Then
                .lstSpells.AddItem i & ". No Spell"
            Else
                .lstSpells.AddItem i & ". " & Trim$(spell(TempPlayerSpells(i)).Name)
            End If
        Next
        .lstSpells.ListIndex = X
    End With
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "InitPlayerSpellEditor", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub InitPlayerItemEditor()
Dim i As Long, X As Long

   On Error GoTo errorhandler

    With frmAdmin
        X = .cmbInvSlot.ListIndex
        .cmbInvSlot.Clear
        For i = 1 To MAX_INV
            If TempPlayerInv(i).Num = 0 Then
                .cmbInvSlot.AddItem i & ". No Item"
            Else
                If TempPlayerInv(i).Value = 0 Or TempPlayerInv(i).Value = 1 Then
                    .cmbInvSlot.AddItem i & ". " & Trim$(Item(TempPlayerInv(i).Num).Name) & " x1"
                Else
                    .cmbInvSlot.AddItem i & ". " & Trim$(Item(TempPlayerInv(i).Num).Name) & " x" & TempPlayerInv(i).Value
                End If
            End If
        Next
        .cmbInvSlot.ListIndex = X
    End With
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "InitPlayerItemEditor", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub



' ///////////////////////
' // Projectile Editor //
' ///////////////////////
Public Sub ProjectileEditorInit()
Dim i As Long, X As Long
Dim MusicFound As Boolean
Dim tmpString() As String
Dim npcNum As Long


   On Error GoTo errorhandler

    If frmEditor_Projectile.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Projectile.lstIndex.ListIndex + 1
    
    'Rec Bullshit
    With Projectiles(EditorIndex)
        frmEditor_Projectile.txtName.Text = Trim$(.Name)
        frmEditor_Projectile.scrlPic.Value = .Sprite
        frmEditor_Projectile.scrlRange.Value = .Range
        frmEditor_Projectile.scrlSpeed.Value = .speed
        frmEditor_Projectile.scrlDamage.Value = .Damage
    End With
    
    Projectile_Changed(EditorIndex) = True

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ProjectileEditorInit", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub ProjectileEditorOk()
Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_PROJECTILES
        If Projectile_Changed(i) Then
            Call SendSaveProjectile(i)
        End If
    Next
    
    Unload frmEditor_Projectile
    Editor = 0
    ClearChanged_Projectile


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ProjectileEditorOk", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub ProjectileEditorCancel()

   On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Projectile
    ClearChanged_Projectile
    ClearProjectiles
    SendRequestProjectiles


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ProjectileEditorCancel", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub ClearChanged_Projectile()

   On Error GoTo errorhandler

   ZeroMemory Projectile_Changed(1), MAX_PROJECTILES * 2 ' 2 = boolean length


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearChanged_Projctile", "modGameEditors", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
