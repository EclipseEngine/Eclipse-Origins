Attribute VB_Name = "modText"
Option Explicit
' Stuffs
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type CharVA
    Vertex(0 To 3) As TLVERTEX
End Type

Public Type VFH
    BitmapWidth As Long
    BitmapHeight As Long
    CellWidth As Long
    CellHeight As Long
    BaseCharOffset As Byte
    CharWidth(0 To 255) As Byte
    CharVA(0 To 255) As CharVA
End Type

Public Type CustomFont
    HeaderInfo As VFH
    Texture As DX8TextureRec
    RowPitch As Integer
    RowFactor As Single
    ColFactor As Single
    CharHeight As Byte
End Type

' Chat Buffer
Public ChatVA() As TLVERTEX
Public ChatVAS() As TLVERTEX

Public Const ChatTextBufferSize As Integer = 200
Public ChatBufferChunk As Single

'Text buffer
Public Type ChatTextBuffer
    Text As String
    color As Long
End Type

'Chat vertex buffer information
Public ChatArrayUbound As Long
Public ChatVB As Direct3DVertexBuffer8
Public ChatVBS As Direct3DVertexBuffer8
Public ChatTextBuffer(1 To ChatTextBufferSize) As ChatTextBuffer

Public Font_Default As CustomFont
Public Font_Georgia As CustomFont

Public Const FVF_SIZE As Long = 28

Public Sub RenderText(ByRef UseFont As CustomFont, ByVal Text As String, ByVal X As Long, ByVal Y As Long, ByVal color As Long, Optional ByVal Alpha As Long = 0, Optional Shadow As Boolean = True, Optional GameScreen As Boolean = False)
Dim TempVA(0 To 3)  As TLVERTEX
Dim TempVAS(0 To 3) As TLVERTEX
Dim TempStr() As String
Dim count As Integer
Dim Ascii() As Byte
Dim Row As Integer
Dim u As Single
Dim v As Single
Dim i As Long
Dim j As Long
Dim KeyPhrase As Byte
Dim TempColor As Long
Dim ResetColor As Byte
Dim srcRect As rect
Dim v2 As D3DVECTOR2
Dim v3 As D3DVECTOR2
Dim YOffset As Single

    ' set the color

   On Error GoTo ErrorHandler

    Alpha = 255 - Alpha
    color = dx8Colour(color, Alpha)
    'Check for valid text to render
    If LenB(Text) = 0 Then Exit Sub
    'Get the text into arrays (split by vbCrLf)
    TempStr = Split(Text, vbCrLf)
    'Set the temp color (or else the first character has no color)
    TempColor = color
    'Set the texture
    Direct3D_Device.SetTexture 0, gTexture(UseFont.Texture.Texture).Texture
    'CurrentTexture = -1
    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(TempStr)
        If Len(TempStr(i)) > 0 Then
            YOffset = i * UseFont.CharHeight
            count = 0
            'Convert the characters to the ascii value
            Ascii() = StrConv(TempStr(i), vbFromUnicode)
                    'Loop through the characters
            For j = 1 To Len(TempStr(i))
                'Copy from the cached vertex array to the temp vertex array
                Call CopyMemory(TempVA(0), UseFont.HeaderInfo.CharVA(Ascii(j - 1)).Vertex(0), FVF_SIZE * 4)
                            'Set up the verticies
                TempVA(0).X = X + count
                TempVA(0).Y = Y + YOffset
                TempVA(1).X = TempVA(1).X + X + count
                TempVA(1).Y = TempVA(0).Y
                TempVA(2).X = TempVA(0).X
                TempVA(2).Y = TempVA(2).Y + TempVA(0).Y
                TempVA(3).X = TempVA(1).X
                TempVA(3).Y = TempVA(2).Y
                            If GameScreen Then
                    Dim X1 As Long, Y1 As Long, srcX As Double, srcY As Double, dw As Long, dH As Long, sW As Double, sH As Double, dX As Long, dY As Long
                    Dim tOffsetX As Long, tOffsetY As Long
                    'Lets do some trimming
                    dX = TempVA(0).X
                    dY = TempVA(0).Y
                    srcX = TempVA(0).TU
                    srcY = TempVA(0).TV
                    dw = TempVA(1).X - TempVA(0).X
                    dH = TempVA(2).Y - TempVA(1).Y
                    sW = TempVA(1).TU - TempVA(0).TU
                    sH = TempVA(2).TV - TempVA(0).TV
                                                                    If dX + dw < GameScreenBounds.Left Then
                        dw = 0
                    Else
                        'Trimming
                        If dX < GameScreenBounds.Left Then
                            dw = 0
                        End If
                    End If
                                            If dY + dH < GameScreenBounds.Top Then
                        dH = 0
                    Else
                        'Trimming
                        If dY < GameScreenBounds.Top Then
                            dw = 0
                        End If
                    End If
                                    If dX > GameScreenBounds.Right Then
                        sW = 0
                        dw = 0
                    Else
                        If X1 + dw > GameScreenBounds.Right Then
                            dw = 0
                        End If
                    End If
                                    If Y1 > GameScreenBounds.Bottom Then
                        sH = 0
                        dH = 0
                    Else
                        If dY + dH > GameScreenBounds.Bottom Then
                            dw = 0
                        End If
                    End If
                                                    Y1 = dY + tOffsetY
                    X1 = dX + tOffsetX

                                    TempVA(0).X = X1
                    TempVA(0).Y = Y1
                    TempVA(0).TU = srcX
                    TempVA(0).TV = srcY
                                    TempVA(1).X = TempVA(0).X + dw
                    TempVA(1).Y = TempVA(0).Y
                    TempVA(1).TU = TempVA(0).TU + sW
                    TempVA(1).TV = TempVA(0).TV
                                    TempVA(2).X = TempVA(0).X
                    TempVA(2).Y = TempVA(0).Y + dH
                    TempVA(2).TU = TempVA(0).TU
                    TempVA(2).TV = TempVA(1).TV + sH
                                    TempVA(3).X = TempVA(1).X
                    TempVA(3).Y = TempVA(2).Y
                    TempVA(3).TU = TempVA(1).TU
                    TempVA(3).TV = TempVA(2).TV
                            End If
                            'Set the colors
                TempVA(0).color = TempColor
                TempVA(1).color = TempColor
                TempVA(2).color = TempColor
                TempVA(3).color = TempColor
                            'Draw the verticies
                Call Direct3D_Device.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, TempVA(0), Len(TempVA(0)))
                            'Shift over the the position to render the next character
                count = count + UseFont.HeaderInfo.CharWidth(Ascii(j - 1))
                            'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = color
                End If
            Next j
        End If
    Next i


   On Error GoTo 0
   Exit Sub
ErrorHandler:
    HandleError "RenderText", "modText", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub EngineInitFontTextures()
    ' FONT DEFAULT

   On Error GoTo ErrorHandler

    NumTextures = NumTextures + 1
    ReDim Preserve gTexture(NumTextures)
    Font_Default.Texture.Texture = NumTextures
    Font_Default.Texture.filepath = App.path & FONT_PATH & "texdefault.png"
    LoadTexture1 Font_Default.Texture
    ' Georgia
    NumTextures = NumTextures + 1
    ReDim Preserve gTexture(NumTextures)
    Font_Georgia.Texture.Texture = NumTextures
    Font_Georgia.Texture.filepath = App.path & FONT_PATH & "georgia.png"
    LoadTexture1 Font_Georgia.Texture


   On Error GoTo 0
   Exit Sub
ErrorHandler:
    HandleError "EngineInitFontTextures", "modText", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub UnloadFontTextures()

   On Error GoTo ErrorHandler

    UnloadFont Font_Default
    UnloadFont Font_Georgia


   On Error GoTo 0
   Exit Sub
ErrorHandler:
    HandleError "UnloadFontTextures", "modText", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Sub UnloadFont(font As CustomFont)

   On Error GoTo ErrorHandler

    font.Texture.Texture = 0


   On Error GoTo 0
   Exit Sub
ErrorHandler:
    HandleError "UnloadFont", "modText", Err.Number, Err.Description, Erl
    Err.Clear
End Sub


Sub LoadFontHeader(ByRef theFont As CustomFont, ByVal filename As String)
Dim fileNum As Byte
Dim LoopChar As Long
Dim Row As Single
Dim u As Single
Dim v As Single


    'Load the header information

   On Error GoTo ErrorHandler

    fileNum = FreeFile
    Open App.path & FONT_PATH & filename For Binary As #fileNum
        Get #fileNum, , theFont.HeaderInfo
    Close #fileNum
    'Calculate some common values
    theFont.CharHeight = theFont.HeaderInfo.CellHeight - 4
    theFont.RowPitch = theFont.HeaderInfo.BitmapWidth \ theFont.HeaderInfo.CellWidth
    theFont.ColFactor = theFont.HeaderInfo.CellWidth / theFont.HeaderInfo.BitmapWidth
    theFont.RowFactor = theFont.HeaderInfo.CellHeight / theFont.HeaderInfo.BitmapHeight
    'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - theFont.HeaderInfo.BaseCharOffset) \ theFont.RowPitch
        u = ((LoopChar - theFont.HeaderInfo.BaseCharOffset) - (Row * theFont.RowPitch)) * theFont.ColFactor
        v = Row * theFont.RowFactor
            'Set the verticies
        With theFont.HeaderInfo.CharVA(LoopChar)
            .Vertex(0).color = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
            .Vertex(0).RHW = 1
            .Vertex(0).TU = u
            .Vertex(0).TV = v
            .Vertex(0).X = 0
            .Vertex(0).Y = 0
            .Vertex(0).z = 0
            .Vertex(1).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).RHW = 1
            .Vertex(1).TU = u + theFont.ColFactor
            .Vertex(1).TV = v
            .Vertex(1).X = theFont.HeaderInfo.CellWidth
            .Vertex(1).Y = 0
            .Vertex(1).z = 0
            .Vertex(2).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).RHW = 1
            .Vertex(2).TU = u
            .Vertex(2).TV = v + theFont.RowFactor
            .Vertex(2).X = 0
            .Vertex(2).Y = theFont.HeaderInfo.CellHeight
            .Vertex(2).z = 0
            .Vertex(3).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).RHW = 1
            .Vertex(3).TU = u + theFont.ColFactor
            .Vertex(3).TV = v + theFont.RowFactor
            .Vertex(3).X = theFont.HeaderInfo.CellWidth
            .Vertex(3).Y = theFont.HeaderInfo.CellHeight
            .Vertex(3).z = 0
        End With
    Next LoopChar


   On Error GoTo 0
   Exit Sub
ErrorHandler:
    HandleError "LoadFontHeader", "modText", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub EngineInitFontSettings()

   On Error GoTo ErrorHandler

    LoadFontHeader Font_Default, "texdefault.dat"
    LoadFontHeader Font_Georgia, "georgia.dat"


   On Error GoTo 0
   Exit Sub
ErrorHandler:
    HandleError "EngineInitFontSettings", "modText", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Public Function dx8Colour(ByVal colourNum As Long, ByVal Alpha As Long) As Long

   On Error GoTo ErrorHandler

    Select Case colourNum
        Case 0 ' Black
            dx8Colour = D3DColorARGB(Alpha, 0, 0, 0)
        Case 1 ' Blue
            dx8Colour = D3DColorARGB(Alpha, 16, 104, 237)
        Case 2 ' Green
            dx8Colour = D3DColorARGB(Alpha, 119, 188, 84)
        Case 3 ' Cyan
            dx8Colour = D3DColorARGB(Alpha, 16, 224, 237)
        Case 4 ' Red
            dx8Colour = D3DColorARGB(Alpha, 201, 0, 0)
        Case 5 ' Magenta
            dx8Colour = D3DColorARGB(Alpha, 255, 0, 255)
        Case 6 ' Brown
            dx8Colour = D3DColorARGB(Alpha, 175, 149, 92)
        Case 7 ' Grey
            dx8Colour = D3DColorARGB(Alpha, 192, 192, 192)
        Case 8 ' DarkGrey
            dx8Colour = D3DColorARGB(Alpha, 128, 128, 128)
        Case 9 ' BrightBlue
            dx8Colour = D3DColorARGB(Alpha, 126, 182, 240)
        Case 10 ' BrightGreen
            dx8Colour = D3DColorARGB(Alpha, 126, 240, 137)
        Case 11 ' BrightCyan
            dx8Colour = D3DColorARGB(Alpha, 157, 242, 242)
        Case 12 ' BrightRed
            dx8Colour = D3DColorARGB(Alpha, 255, 0, 0)
        Case 13 ' Pink
            dx8Colour = D3DColorARGB(Alpha, 255, 118, 221)
        Case 14 ' Yellow
            dx8Colour = D3DColorARGB(Alpha, 255, 255, 0)
        Case 15 ' White
            dx8Colour = D3DColorARGB(Alpha, 255, 255, 255)
        Case 16 ' dark brown
            dx8Colour = D3DColorARGB(Alpha, 98, 84, 52)
        Case 17 'Orange
            dx8Colour = D3DColorARGB(Alpha, 255, 96, 0)
        Case 18 'Darkcolor
            dx8Colour = D3DColorARGB(Alpha, 25, 25, 25)
        Case 19 'Purple
            dx8Colour = D3DColorARGB(Alpha, 160, 32, 240)
    End Select


   On Error GoTo 0
   Exit Function
ErrorHandler:
    HandleError "dx8Colour", "modText", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function EngineGetTextWidth(ByRef UseFont As CustomFont, ByVal Text As String) As Integer
Dim LoopI As Integer

    'Make sure we have text

   On Error GoTo ErrorHandler

    If LenB(Text) = 0 Then Exit Function
    'Loop through the text
    For LoopI = 1 To Len(Text)
        EngineGetTextWidth = EngineGetTextWidth + UseFont.HeaderInfo.CharWidth(Asc(Mid$(Text, LoopI, 1)))
    Next LoopI


   On Error GoTo 0
   Exit Function
ErrorHandler:
    HandleError "EngineGetTextWidth", "modText", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Public Sub DrawPlayerName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim color As Long
Dim Name As String


   On Error GoTo ErrorHandler

    If Player(MyIndex).InHouse > 0 Then
        If Player(Index).InHouse <> Player(MyIndex).InHouse Then Exit Sub
    End If

    ' Check access level
    If GetPlayerPK(Index) = NO Then

        Select Case GetPlayerAccess(Index)
            Case 0
                color = Orange
            Case 1
                color = White
            Case 2
                color = Cyan
            Case 3
                color = BrightGreen
            Case 4
                color = Yellow
        End Select

    Else
        color = BrightRed
    End If

    Name = Trim$(Player(Index).Name)
    If Player(Index).Access > 0 Then
        Name = "[GM] " & Trim$(Player(Index).Name) & " Lv. " & Player(Index).Level
    Else
        Name = Trim$(Player(Index).Name) & " Lv. " & Player(Index).Level
    End If
    ' calc pos
    TextX = ConvertMapX(GetPlayerX(Index) * PIC_X) + Player(Index).XOffset + (PIC_X \ 2) - (getWidth(Font_Default, (Trim$(Name))) / 2)
    If CharMode = 1 Then
        If Player(Index).Sprite(SpriteEnum.Body) < 1 Or Player(Index).Sprite(SpriteEnum.Body) > NumCharacters Then
            TextY = ConvertMapY((GetPlayerY(Index) * PIC_Y) + Player(Index).YOffset - 16)
        Else
            ' Determine location for text
            TextY = ConvertMapY((GetPlayerY(Index) * PIC_Y) + Player(Index).YOffset - Tex_CharBodies(Player(Index).Sprite(SpriteEnum.Body)).Height / 4 + 16)
        End If
    Else
        If Player(Index).Sprite(FaceEnum.Head) < 1 Or Player(Index).Sprite(FaceEnum.Head) > NumCharacters Then
            TextY = ConvertMapY((GetPlayerY(Index) * PIC_Y) + Player(Index).YOffset - 16)
        Else
            ' Determine location for text
            TextY = ConvertMapY((GetPlayerY(Index) * PIC_Y) + Player(Index).YOffset - Tex_Character(Player(Index).Face(FaceEnum.Head)).Height / 4)
        End If
    End If
    
    

    ' Draw name
    'Call DrawText(TexthDC, TextX, TextY, Name, Color)
    RenderText Font_Default, Name, TextX, TextY, color, 0, True, True




   On Error GoTo 0
   Exit Sub
ErrorHandler:
    HandleError "DrawPlayerName", "modText", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawZoneNpcName(ByVal zonenum As Long, ByVal zoneNpcNum As Long)
Dim TextX As Long
Dim TextY As Long
Dim color As Long
Dim Name As String
Dim npcNum As Long



   On Error GoTo ErrorHandler

    npcNum = ZoneNPC(zonenum).Npc(zoneNpcNum).Num

    Select Case Npc(npcNum).Behaviour
        Case NPC_BEHAVIOUR_ATTACKONSIGHT
            color = BrightRed
        Case NPC_BEHAVIOUR_ATTACKWHENATTACKED
            color = Yellow
        Case NPC_BEHAVIOUR_GUARD
            color = Grey
        Case Else
            color = BrightGreen
    End Select

    Name = Trim$(Npc(npcNum).Name)
    TextX = ConvertMapX(ZoneNPC(zonenum).Npc(zoneNpcNum).X * PIC_X) + ZoneNPC(zonenum).Npc(zoneNpcNum).XOffset + (PIC_X \ 2) - (getWidth(Font_Default, (Trim$(Name))) / 2)
    If Npc(npcNum).Sprite < 1 Or Npc(npcNum).Sprite > NumCharacters Then
        TextY = ConvertMapY((ZoneNPC(zonenum).Npc(zoneNpcNum).Y * PIC_Y) + ZoneNPC(zonenum).Npc(zoneNpcNum).YOffset - 16)
    Else
        ' Determine location for text
        TextY = ConvertMapY((ZoneNPC(zonenum).Npc(zoneNpcNum).Y * PIC_Y) + ZoneNPC(zonenum).Npc(zoneNpcNum).YOffset - (Tex_Character(Npc(npcNum).Sprite).Height / 4))
    End If

    ' Draw name
    'Call DrawText(TexthDC, TextX, TextY, Name, Color)
    RenderText Font_Default, Name, TextX, TextY, color, 0, True, True

   On Error GoTo 0
   Exit Sub
ErrorHandler:
    HandleError "DrawZoneNpcName", "modText", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawNpcName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim color As Long
Dim Name As String
Dim npcNum As Long
Dim blstr As String * NAME_LENGTH



   On Error GoTo ErrorHandler

    npcNum = MapNpc(Index).Num

    Select Case Npc(npcNum).Behaviour
        Case NPC_BEHAVIOUR_ATTACKONSIGHT
            color = BrightRed
        Case NPC_BEHAVIOUR_ATTACKWHENATTACKED
            color = Yellow
        Case NPC_BEHAVIOUR_GUARD
            color = Grey
        Case Else
            color = BrightGreen
    End Select

    Name = Trim$(Npc(npcNum).Name)
    If Name = blstr Then Exit Sub
    TextX = ConvertMapX(MapNpc(Index).X * PIC_X) + MapNpc(Index).XOffset + (PIC_X \ 2) - (getWidth(Font_Default, (Trim$(Name))) / 2)
    If Npc(npcNum).Sprite < 1 Or Npc(npcNum).Sprite > NumCharacters Then
        TextY = ConvertMapY((MapNpc(Index).Y * PIC_Y) + MapNpc(Index).YOffset - 16)
    Else
        ' Determine location for text
        TextY = ConvertMapY((MapNpc(Index).Y * PIC_Y) + MapNpc(Index).YOffset - (Tex_Character(Npc(npcNum).Sprite).Height / 4))
    End If

    ' Draw name
    'Call DrawText(TexthDC, TextX, TextY, Name, Color)
    RenderText Font_Default, Name, TextX, TextY, color, 0, True, True
    Dim i As Long





   On Error GoTo 0
   Exit Sub
ErrorHandler:
    HandleError "DrawNpcName", "modText", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Function DrawMapAttributes()
    Dim X As Long
    Dim Y As Long
    Dim tx As Long
    Dim ty As Long


   On Error GoTo ErrorHandler

    If frmEditor_Map.optAttribs.Value Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(X, Y) Then
                    With Map.Tile(X, Y)
                        tx = ((ConvertMapX(X * PIC_X)) - 4) + (PIC_X * 0.5)
                        ty = ((ConvertMapY(Y * PIC_Y)) - 7) + (PIC_Y * 0.5)
                        Select Case .type
                            Case TILE_TYPE_BLOCKED
                                RenderText Font_Default, "B", tx, ty, BrightRed, 0, True, True
                            Case TILE_TYPE_WARP
                                RenderText Font_Default, "W", tx, ty, BrightBlue, 0, True, True
                            Case TILE_TYPE_ITEM
                                RenderText Font_Default, "I", tx, ty, White, 0, True, True
                            Case TILE_TYPE_NPCAVOID
                                RenderText Font_Default, "N", tx, ty, White, 0, True, True
                            Case TILE_TYPE_KEY
                                RenderText Font_Default, "K", tx, ty, White, 0, True, True
                            Case TILE_TYPE_KEYOPEN
                                RenderText Font_Default, "O", tx, ty, White, 0, True, True
                            Case TILE_TYPE_RESOURCE
                                RenderText Font_Default, "B", tx, ty, Green, 0, True, True
                            Case TILE_TYPE_DOOR
                                RenderText Font_Default, "D", tx, ty, Brown, 0, True, True
                            Case TILE_TYPE_NPCSPAWN
                                RenderText Font_Default, "S", tx, ty, Yellow, 0, True, True
                            Case TILE_TYPE_SHOP
                                RenderText Font_Default, "S", tx, ty, BrightBlue, 0, True, True
                            Case TILE_TYPE_BANK
                                RenderText Font_Default, "B", tx, ty, Blue, 0, True, True
                            Case TILE_TYPE_HEAL
                                RenderText Font_Default, "H", tx, ty, BrightGreen, 0, True, True
                            Case TILE_TYPE_TRAP
                                RenderText Font_Default, "T", tx, ty, BrightRed, 0, True, True
                            Case TILE_TYPE_SLIDE
                                RenderText Font_Default, "S", tx, ty, BrightCyan, 0, True, True
                            Case TILE_TYPE_SOUND
                                RenderText Font_Default, "S", tx, ty, Orange, 0, True, True
                            Case TILE_TYPE_HOUSE
                                RenderText Font_Default, "H", tx, ty, Orange, 0, True, True
                            Case TILE_TYPE_INSTANCE
                                RenderText Font_Default, "I", tx, ty, Black, 0
                            Case TILE_TYPE_RANDOMDUNGEON
                                RenderText Font_Default, "RD", tx, ty, BrightRed, 0
                        End Select
                    End With
                End If
            Next
        Next
    End If




   On Error GoTo 0
   Exit Function
ErrorHandler:
    HandleError "DrawMapAttributes", "modText", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub DrawActionMsg(ByVal Index As Long, ByVal updatePos As Boolean)
    Dim X As Long, Y As Long, i As Long, Time As Long
    ' does it exist

   On Error GoTo ErrorHandler
    
    If ActionMsg(Index).Created = 0 Then Exit Sub
    ' how long we want each message to appear
    Select Case ActionMsg(Index).type
        Case ACTIONMSG_STATIC
            Time = 1500
    
            If ActionMsg(Index).Y > 0 Then
                X = ActionMsg(Index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) - 2
            Else
                X = ActionMsg(Index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) + 18
            End If
    
        Case ACTIONMSG_SCROLL
            Time = 1500
                If ActionMsg(Index).Y > 0 Then
                X = ActionMsg(Index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) - 2 - (ActionMsg(Index).Scroll * 0.6)
                If updatePos Then ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            Else
                X = ActionMsg(Index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) + 18 + (ActionMsg(Index).Scroll * 0.6)
                If updatePos Then ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            End If
    
        Case ACTIONMSG_SCREEN
            Time = 3000
    
            ' This will kill any action screen messages that there in the system
            For i = MAX_BYTE To 1 Step -1
                If ActionMsg(i).type = ACTIONMSG_SCREEN Then
                    If i <> Index Then
                        ClearActionMsg Index
                        Index = i
                    End If
                End If
            Next
            X = (frmMain.Width \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
            Y = 425
    End Select
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)

    If GetTickCount < ActionMsg(Index).Created + Time Then
        RenderText Font_Default, ActionMsg(Index).Message, X, Y, ActionMsg(Index).color, 0, True, True
    Else
        ClearActionMsg Index
    End If

   On Error GoTo 0
   Exit Sub
ErrorHandler:
    HandleError "DrawActionMsg", "modText", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Function getWidth(font As CustomFont, ByVal Text As String) As Long

   On Error GoTo ErrorHandler

    getWidth = EngineGetTextWidth(font, Text)



   On Error GoTo 0
   Exit Function
ErrorHandler:
    HandleError "getWidth", "modText", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Public Sub DrawEventName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim color As Long
Dim Name As String, i As Long


   On Error GoTo ErrorHandler

    If InMapEditor Then Exit Sub

    color = White

    Name = Trim$(Map.MapEvents(Index).Name)
    ' calc pos
    TextX = ConvertMapX(Map.MapEvents(Index).X * PIC_X) + Map.MapEvents(Index).XOffset + (PIC_X \ 2) - (getWidth(Font_Default, (Trim$(Name))) / 2)
    If Map.MapEvents(Index).GraphicType = 0 Then
        TextY = ConvertMapY(Map.MapEvents(Index).Y * PIC_Y) + Map.MapEvents(Index).YOffset - 16
    ElseIf Map.MapEvents(Index).GraphicType = 1 Then
        If Map.MapEvents(Index).GraphicNum < 1 Or Map.MapEvents(Index).GraphicNum > NumCharacters Then
            TextY = ConvertMapY(Map.MapEvents(Index).Y * PIC_Y) + Map.MapEvents(Index).YOffset - 16
        Else
            ' Determine location for text
            TextY = ConvertMapY(Map.MapEvents(Index).Y * PIC_Y) + Map.MapEvents(Index).YOffset - (Tex_Character(Map.MapEvents(Index).GraphicNum).Height / 4) + 16
        End If
    ElseIf Map.MapEvents(Index).GraphicType = 2 Then
        If Map.MapEvents(Index).GraphicY2 > 0 Then
            TextY = ConvertMapY(Map.MapEvents(Index).Y * PIC_Y) + Map.MapEvents(Index).YOffset - ((Map.MapEvents(Index).GraphicY2 - Map.MapEvents(Index).GraphicY) * 32) + 16
        Else
            TextY = ConvertMapY(Map.MapEvents(Index).Y * PIC_Y) + Map.MapEvents(Index).YOffset - 32 + 16
        End If
    End If

    ' Draw name
    RenderText Font_Default, Name, TextX, TextY, color, 0, True, True
    For i = 1 To MAX_QUESTS
        'check if the npc is the starter to any quest: [!] symbol
        'can accept the quest as a new one?
        If Player(MyIndex).PlayerQuest(i).state = QUEST_NOT_STARTED Or Player(MyIndex).PlayerQuest(i).state = QUEST_COMPLETED_BUT Or (Player(MyIndex).PlayerQuest(i).state = QUEST_COMPLETED And quest(i).Repeatable = 1) Then
                'the npc gives this quest?
            If Map.MapEvents(Index).questnum = i Then
                Name = "[!]"
                TextX = ConvertMapX(Map.MapEvents(Index).X * PIC_X) + Map.MapEvents(Index).XOffset + (PIC_X \ 2) - getWidth(Font_Default, (Trim$("[!]"))) + 8
                TextY = TextY - 16
                If quest(i).Repeatable = 1 Then
                    RenderText Font_Default, Name, TextX, TextY, White, 0, True, True
                Else
                    RenderText Font_Default, Name, TextX, TextY, Yellow, 0, True, True
                End If
                Exit For
            End If
        ElseIf Player(MyIndex).PlayerQuest(i).state = QUEST_STARTED Then
            If Map.MapEvents(Index).questnum = i Then
                Name = "[*]"
                TextX = ConvertMapX(Map.MapEvents(Index).X * PIC_X) + Map.MapEvents(Index).XOffset + (PIC_X \ 2) - getWidth(Font_Default, (Trim$("[*]"))) + 8
                TextY = TextY - 16
                RenderText Font_Default, Name, TextX, TextY, Yellow, 0, True, True
                Exit For
            End If
        End If
    Next




   On Error GoTo 0
   Exit Sub
ErrorHandler:
    HandleError "DrawEventName", "modText", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawChatBubble(ByVal Index As Long)
Dim theArray() As String, X As Long, Y As Long, i As Long, MaxWidth As Long, X2 As Long, Y2 As Long, colour As Long

   On Error GoTo ErrorHandler

    With chatBubble(Index)
        If .targetType = TARGET_TYPE_PLAYER Then
            ' it's a player
            If GetPlayerMap(.target) = GetPlayerMap(MyIndex) Then
                ' it's on our map - get co-ords
                X = ConvertMapX((Player(.target).X * 32) + Player(.target).XOffset) + 16
                Y = ConvertMapY((Player(.target).Y * 32) + Player(.target).YOffset) - 40
            End If
        ElseIf .targetType = TARGET_TYPE_NPC Then
            ' it's on our map - get co-ords
            X = ConvertMapX((MapNpc(.target).X * 32) + MapNpc(.target).XOffset) + 16
            Y = ConvertMapY((MapNpc(.target).Y * 32) + MapNpc(.target).YOffset) - 40
        ElseIf .targetType = TARGET_TYPE_EVENT Then
            X = ConvertMapX((Map.MapEvents(.target).X * 32) + Map.MapEvents(.target).XOffset) + 16
            Y = ConvertMapY((Map.MapEvents(.target).Y * 32) + Map.MapEvents(.target).YOffset) - 40
        End If
            ' word wrap the text
        WordWrap_Array .Msg, ChatBubbleWidth, theArray
                    ' find max width
        For i = 1 To UBound(theArray)
            If EngineGetTextWidth(Font_Default, theArray(i)) > MaxWidth Then MaxWidth = EngineGetTextWidth(Font_Default, theArray(i))
        Next
                    ' calculate the new position
        X2 = X - (MaxWidth \ 2)
        Y2 = Y - (UBound(theArray) * 12)
                    ' render bubble - top left
        RenderTexture Tex_ChatBubble, X2 - 9, Y2 - 5, 0, 0, 9, 5, 9, 5, -1, True
        ' top right
        RenderTexture Tex_ChatBubble, X2 + MaxWidth, Y2 - 5, 119, 0, 9, 5, 9, 5, -1, True
        ' top
        RenderTexture Tex_ChatBubble, X2, Y2 - 5, 10, 0, MaxWidth, 5, 5, 5, -1, True
        ' bottom left
        RenderTexture Tex_ChatBubble, X2 - 9, Y, 0, 19, 9, 6, 9, 6, -1, True
        ' bottom right
        RenderTexture Tex_ChatBubble, X2 + MaxWidth, Y, 119, 19, 9, 6, 9, 6, -1, True
        ' bottom - left half
        RenderTexture Tex_ChatBubble, X2, Y, 10, 19, (MaxWidth \ 2) - 5, 6, 9, 6, -1, True
        ' bottom - right half
        RenderTexture Tex_ChatBubble, X2 + (MaxWidth \ 2) + 6, Y, 10, 19, (MaxWidth \ 2) - 5, 6, 9, 6, -1, True
        ' left
        RenderTexture Tex_ChatBubble, X2 - 9, Y2, 0, 6, 9, (UBound(theArray) * 12), 9, 1, -1, True
        ' right
        RenderTexture Tex_ChatBubble, X2 + MaxWidth, Y2, 119, 6, 9, (UBound(theArray) * 12), 9, 1, -1, True
        ' center
        RenderTexture Tex_ChatBubble, X2, Y2, 9, 5, MaxWidth, (UBound(theArray) * 12), 1, 1, -1, True
        ' little pointy bit
        RenderTexture Tex_ChatBubble, X - 5, Y, 58, 19, 11, 11, 11, 11, -1, True
                    ' render each line centralised
        For i = 1 To UBound(theArray)
            RenderText Font_Georgia, theArray(i), X - (EngineGetTextWidth(Font_Default, theArray(i)) / 2), Y2, DarkBrown, 0, True, True
            Y2 = Y2 + 12
        Next
        ' check if it's timed out - close it if so
        If .Timer + 5000 < GetTickCount Then
            .active = False
        End If
    End With


   On Error GoTo 0
   Exit Sub
ErrorHandler:
    HandleError "DrawChatBubble", "modText", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub WordWrap_Array(ByVal Text As String, ByVal MaxLineLen As Long, ByRef theArray() As String)
Dim lineCount As Long, i As Long, size As Long, lastSpace As Long, b As Long
    'Too small of text

   On Error GoTo ErrorHandler

    If Len(Text) < 2 Then
        ReDim theArray(1 To 1) As String
        theArray(1) = Text
        Exit Sub
    End If
    ' default values
    b = 1
    lastSpace = 1
    size = 0
    For i = 1 To Len(Text)
        ' if it's a space, store it
        Select Case Mid$(Text, i, 1)
            Case " ": lastSpace = i
            Case "_": lastSpace = i
            Case "-": lastSpace = i
        End Select
            'Add up the size
        size = size + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(Text, i, 1)))
            'Check for too large of a size
        If size > MaxLineLen Then
            'Check if the last space was too far back
            If i - lastSpace > 12 Then
                'Too far away to the last space, so break at the last character
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(Text, b, (i - 1) - b))
                b = i - 1
                size = 0
            Else
                'Break at the last space to preserve the word
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(Text, b, lastSpace - b))
                b = lastSpace + 1
                            'Count all the words we ignored (the ones that weren't printed, but are before "i")
                size = EngineGetTextWidth(Font_Default, Mid$(Text, lastSpace, i - lastSpace))
            End If
        End If
            ' Remainder
        If i = Len(Text) Then
            If b <> i Then
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = theArray(lineCount) & Mid$(Text, b, i)
            End If
        End If
    Next


   On Error GoTo 0
   Exit Sub
ErrorHandler:
    HandleError "WordWrap_Array", "modText", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Function WordWrap(ByVal Text As String, ByVal MaxLineLen As Integer) As String
Dim TempSplit() As String
Dim TSLoop As Long
Dim lastSpace As Long
Dim size As Long
Dim i As Long
Dim b As Long

    'Too small of text

   On Error GoTo ErrorHandler

    If Len(Text) < 2 Then
        WordWrap = Text
        Exit Function
    End If

    'Check if there are any line breaks - if so, we will support them
    TempSplit = Split(Text, vbNewLine)
    For TSLoop = 0 To UBound(TempSplit)
        'Clear the values for the new line
        size = 0
        b = 1
        lastSpace = 1
            'Add back in the vbNewLines
        If TSLoop < UBound(TempSplit()) Then TempSplit(TSLoop) = TempSplit(TSLoop) & vbNewLine
            'Only check lines with a space
        If InStr(1, TempSplit(TSLoop), " ") Or InStr(1, TempSplit(TSLoop), "-") Or InStr(1, TempSplit(TSLoop), "_") Then
                    'Loop through all the characters
            For i = 1 To Len(TempSplit(TSLoop))
                        'If it is a space, store it so we can easily break at it
                Select Case Mid$(TempSplit(TSLoop), i, 1)
                    Case " ": lastSpace = i
                    Case "_": lastSpace = i
                    Case "-": lastSpace = i
                End Select
                'Add up the size
                size = size + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(TempSplit(TSLoop), i, 1)))
 
                'Check for too large of a size
                If size > MaxLineLen Then
                    'Check if the last space was too far back
                    If i - lastSpace > 12 Then
                        'Too far away to the last space, so break at the last character
                        WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), b, (i - 1) - b)) & vbNewLine
                        b = i - 1
                        size = 0
                    Else
                        'Break at the last space to preserve the word
                        WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), b, lastSpace - b)) & vbNewLine
                        b = lastSpace + 1
                                            'Count all the words we ignored (the ones that weren't printed, but are before "i")
                        size = EngineGetTextWidth(Font_Default, Mid$(TempSplit(TSLoop), lastSpace, i - lastSpace))
                    End If
                End If
                            'This handles the remainder
                If i = Len(TempSplit(TSLoop)) Then
                    If b <> i Then
                        WordWrap = WordWrap & Mid$(TempSplit(TSLoop), b, i)
                    End If
                End If
            Next i
        Else
            WordWrap = WordWrap & TempSplit(TSLoop)
        End If
    Next TSLoop


   On Error GoTo 0
   Exit Function
ErrorHandler:
    HandleError "WordWrap", "modText", Err.Number, Err.Description, Erl
    Err.Clear
End Function



' Chat Box
Public Sub RenderChatTextBuffer()
Dim srcRect As rect
Dim v2 As D3DVECTOR2
Dim v3 As D3DVECTOR2
Dim i As Long

    'Clear the LastTexture, letting the rest of the engine know that the texture needs to be changed for next rect render

   On Error GoTo ErrorHandler

    Direct3D_Device.SetTexture 0, gTexture(Font_Default.Texture.Texture).Texture

    If ChatArrayUbound > 0 Then
        Direct3D_Device.SetStreamSource 0, ChatVBS, FVF_SIZE
        Direct3D_Device.DrawPrimitive D3DPT_TRIANGLELIST, 0, (ChatArrayUbound + 1) \ 3
        Direct3D_Device.SetStreamSource 0, ChatVB, FVF_SIZE
        Direct3D_Device.DrawPrimitive D3DPT_TRIANGLELIST, 0, (ChatArrayUbound + 1) \ 3
    End If


   On Error GoTo 0
   Exit Sub
ErrorHandler:
    HandleError "RenderChatTextBuffer", "modText", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub UpdateChatArray()
Dim Chunk As Integer
Dim count As Integer
Dim LoopC As Byte
Dim Ascii As Byte
Dim Row As Long
Dim Pos As Long
Dim u As Single
Dim v As Single
Dim X As Single
Dim Y As Single
Dim Y2 As Single
Dim i As Long
Dim j As Long
Dim size As Integer
Dim KeyPhrase As Byte
Dim ResetColor As Byte
Dim TempColor As Long
Dim YOffset As Long

    ' set the offset of each line

   On Error GoTo ErrorHandler

    YOffset = 14

    'Set the position
    If ChatBufferChunk <= 1 Then ChatBufferChunk = 1
    Chunk = ChatScroll
    'Get the number of characters in all the visible buffer
    size = 0
    For LoopC = (Chunk * ChatBufferChunk) - (ChatLines - 1) To Chunk * ChatBufferChunk
        If LoopC > ChatTextBufferSize Then Exit For
        size = size + Len(ChatTextBuffer(LoopC).Text)
    Next
    size = size - j
    ChatArrayUbound = size * 6 - 1
    If ChatArrayUbound < 0 Then Exit Sub
    ReDim ChatVA(0 To ChatArrayUbound) 'Size our array to fix the 6 verticies of each character
    ReDim ChatVAS(0 To ChatArrayUbound)
    'Set the base position
    X = ChatOffsetX
    Y = ChatOffsetY + 30

    'Loop through each buffer string
    For LoopC = (Chunk * ChatBufferChunk) - (ChatLines - 1) To Chunk * ChatBufferChunk
        If LoopC > ChatTextBufferSize Then Exit For
        If ChatBufferChunk * Chunk > ChatTextBufferSize Then ChatBufferChunk = ChatBufferChunk - 1
            'Set the temp color
        TempColor = ChatTextBuffer(LoopC).color
            'Set the Y position to be used
        Y2 = Y - (LoopC * YOffset) + (Chunk * ChatBufferChunk * YOffset) - 32
            'Loop through each line if there are line breaks (vbCrLf)
        count = 0   'Counts the offset value we are on
        If LenB(ChatTextBuffer(LoopC).Text) <> 0 Then  'Dont bother with empty strings
                    'Loop through the characters
            For j = 1 To Len(ChatTextBuffer(LoopC).Text)
                        'Convert the character to the ascii value
                Ascii = Asc(Mid$(ChatTextBuffer(LoopC).Text, j, 1))
                            'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
                Row = (Ascii - Font_Default.HeaderInfo.BaseCharOffset) \ Font_Default.RowPitch
                u = ((Ascii - Font_Default.HeaderInfo.BaseCharOffset) - (Row * Font_Default.RowPitch)) * Font_Default.ColFactor
                v = Row * Font_Default.RowFactor

                ' ****** Rectangle | Top Left ******
                With ChatVA(0 + (6 * Pos))
                    .color = TempColor
                    .X = (X) + count
                    .Y = (Y2)
                    .TU = u
                    .TV = v
                    .RHW = 1
                End With
                            ' ****** Rectangle | Bottom Left ******
                With ChatVA(1 + (6 * Pos))
                    .color = TempColor
                    .X = (X) + count
                    .Y = (Y2) + Font_Default.HeaderInfo.CellHeight
                    .TU = u
                    .TV = v + Font_Default.RowFactor
                    .RHW = 1
                End With
                            ' ****** Rectangle | Bottom Right ******
                With ChatVA(2 + (6 * Pos))
                    .color = TempColor
                    .X = (X) + count + Font_Default.HeaderInfo.CellWidth
                    .Y = (Y2) + Font_Default.HeaderInfo.CellHeight
                    .TU = u + Font_Default.ColFactor
                    .TV = v + Font_Default.RowFactor
                    .RHW = 1
                End With
                                        'Triangle 2 (only one new vertice is needed)
                ChatVA(3 + (6 * Pos)) = ChatVA(0 + (6 * Pos)) 'Top-left corner
                            ' ****** Rectangle | Top Right ******
                With ChatVA(4 + (6 * Pos))
                    .color = TempColor
                    .X = (X) + count + Font_Default.HeaderInfo.CellWidth
                    .Y = (Y2)
                    .TU = u + Font_Default.ColFactor
                    .TV = v
                    .RHW = 1
                End With

                ChatVA(5 + (6 * Pos)) = ChatVA(2 + (6 * Pos))

                'Update the character we are on
                Pos = Pos + 1

                'Shift over the the position to render the next character
                count = count + Font_Default.HeaderInfo.CharWidth(Ascii)
                            'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = ChatTextBuffer(LoopC).color
                End If
            Next
        End If
    Next LoopC
        If Not Direct3D_Device Is Nothing Then   'Make sure the D3DDevice exists - this will only return false if we received messages before it had time to load
        Set ChatVBS = Direct3D_Device.CreateVertexBuffer(FVF_SIZE * Pos * 6, 0, FVF_TLVERTEX, D3DPOOL_MANAGED)
        D3DVertexBuffer8SetData ChatVBS, 0, FVF_SIZE * Pos * 6, 0, ChatVAS(0)
        Set ChatVB = Direct3D_Device.CreateVertexBuffer(FVF_SIZE * Pos * 6, 0, FVF_TLVERTEX, D3DPOOL_MANAGED)
        D3DVertexBuffer8SetData ChatVB, 0, FVF_SIZE * Pos * 6, 0, ChatVA(0)
    End If
    Erase ChatVAS()
    Erase ChatVA()


   On Error GoTo 0
   Exit Sub
ErrorHandler:
    HandleError "UpdateChatArray", "modText", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub AddText(ByVal Text As String, ByVal tColor As Long, Optional ByVal Alpha As Long = 255)
Dim TempSplit() As String
Dim TSLoop As Long
Dim lastSpace As Long
Dim size As Long
Dim i As Long
Dim b As Long
Dim color As Long
Dim u As String * 1

   On Error GoTo ErrorHandler
   
    For i = 1 To Len(Text)
        If StrComp(Mid(Text, i, 1), i, vbTextCompare) <> 0 Then
            
        Else
            If i = Len(Text) Then
                Exit Sub
            End If
        End If
    Next
    
    If u = Text Then Exit Sub

    color = dx8Colour(tColor, Alpha)

    'Check if there are any line breaks - if so, we will support them
    TempSplit = Split(Text, vbCrLf)
    For TSLoop = 0 To UBound(TempSplit)
        'Clear the values for the new line
        size = 0
        b = 1
        lastSpace = 1
            'Loop through all the characters
        For i = 1 To Len(TempSplit(TSLoop))
                'If it is a space, store it so we can easily break at it
            Select Case Mid$(TempSplit(TSLoop), i, 1)
                Case " ": lastSpace = i
                Case "_": lastSpace = i
                Case "-": lastSpace = i
            End Select
                    'Add up the size
            size = size + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(TempSplit(TSLoop), i, 1)))
                    'Check for too large of a size
            If size > ChatWidth Then
                            'Check if the last space was too far back
                If i - lastSpace > 10 Then
                                'Too far away to the last space, so break at the last character
                    AddToChatTextBuffer_Overflow Trim$(Mid$(TempSplit(TSLoop), b, (i - 1) - b)), color
                    b = i - 1
                    size = 0
                Else
                    'Break at the last space to preserve the word
                    AddToChatTextBuffer_Overflow Trim$(Mid$(TempSplit(TSLoop), b, lastSpace - b)), color
                    b = lastSpace + 1
                    'Count all the words we ignored (the ones that weren't printed, but are before "i")
                    size = EngineGetTextWidth(Font_Default, Mid$(TempSplit(TSLoop), lastSpace, i - lastSpace))
                End If
            End If
                    'This handles the remainder
            If i = Len(TempSplit(TSLoop)) Then
                If b <> i Then AddToChatTextBuffer_Overflow Mid$(TempSplit(TSLoop), b, i), color
            End If
        Next i
    Next TSLoop
    'Only update if we have set up the text (that way we can add to the buffer before it is even made)
    If Font_Default.RowPitch = 0 Then Exit Sub
    If ChatScroll > 8 Then ChatScroll = ChatScroll + 1

    'Update the array
    UpdateChatArray


   On Error GoTo 0
   Exit Sub
ErrorHandler:
    HandleError "AddText", "modText", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub AddToChatTextBuffer_Overflow(ByVal Text As String, ByVal color As Long)
Dim LoopC As Long

    'Move all other text up

   On Error GoTo ErrorHandler

    For LoopC = (ChatTextBufferSize - 1) To 1 Step -1
        ChatTextBuffer(LoopC + 1) = ChatTextBuffer(LoopC)
    Next LoopC
    'Set the values
    ChatTextBuffer(1).Text = Text
    ChatTextBuffer(1).color = color
    ' set the total chat lines
    totalChatLines = totalChatLines + 1
    If totalChatLines > ChatTextBufferSize - 1 Then totalChatLines = ChatTextBufferSize - 1


   On Error GoTo 0
   Exit Sub
ErrorHandler:
    HandleError "AddToChatTextBuffer_Overflow", "modText", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub UpdateShowChatText()
Dim CHATOFFSET As Long, i As Long, X As Long


   On Error GoTo ErrorHandler

    CHATOFFSET = 52
    If EngineGetTextWidth(Font_Default, MyText) > 760 - CHATOFFSET Then
        For i = Len(MyText) To 1 Step -1
            X = X + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(MyText, i, 1)))
            If X > 760 - CHATOFFSET Then
                RenderChatText = Right$(MyText, Len(MyText) - i + 1)
                Exit For
            End If
        Next
    Else
        RenderChatText = MyText
    End If


   On Error GoTo 0
   Exit Sub
ErrorHandler:
    HandleError "UpdateShowChatText", "modText", Err.Number, Err.Description, Erl
    Err.Clear
End Sub


Public Function LimitText(font As CustomFont, str As String, lmt As Long) As String

   On Error GoTo ErrorHandler

    Do Until EngineGetTextWidth(font, str) <= lmt - 1
        str = Right(str, Len(str) - 1)
    Loop
    LimitText = str


   On Error GoTo 0
   Exit Function
ErrorHandler:
    HandleError "LimitText", "modText", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function TrimText(font As CustomFont, ByVal str As String, lmt As Long) As String

   On Error GoTo ErrorHandler

    Do Until EngineGetTextWidth(font, str) <= lmt - 1
        str = Left(str, Len(str) - 1)
    Loop
    TrimText = str


   On Error GoTo 0
   Exit Function
ErrorHandler:
    HandleError "TrimText", "modText", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function Starz(l As Long) As String
    Dim s As String, i As Long

   On Error GoTo ErrorHandler

    If l > 0 Then
        For i = 1 To l
            s = s & "*"
        Next
    End If
    Starz = s


   On Error GoTo 0
   Exit Function
ErrorHandler:
    HandleError "Starz", "modText", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function nTrim(s As String) As String
Dim a As String * 1

   On Error GoTo ErrorHandler
    
    s = Replace(s, a, "")
    s = Trim(s)
    nTrim = s

   On Error GoTo 0
   Exit Function
ErrorHandler:
    HandleError "nTrim", "modText", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Sub DrawPlayerPetName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim color As Long
Dim Name As String


   On Error GoTo ErrorHandler

    color = BrightRed

    Name = Trim$(GetPlayerName(Index)) & "'s " & Trim$(Pet(Player(Index).Pet.Num).Name)
    ' calc pos
    TextX = ConvertMapX(Player(Index).Pet.X * PIC_X) + Player(Index).Pet.XOffset + (PIC_X \ 2) - (EngineGetTextWidth(Font_Default, Name) / 2)
    If Pet(Player(Index).Pet.Num).Sprite < 1 Or Pet(Player(Index).Pet.Num).Sprite > NumCharacters Then
        TextY = ConvertMapY(Player(Index).Pet.Y * PIC_Y) + Player(Index).Pet.YOffset - 16
    Else
        ' Determine location for text
        TextY = ConvertMapY(Player(Index).Pet.Y * PIC_Y) + Player(Index).Pet.YOffset - (Tex_Character(Pet(Player(Index).Pet.Num).Sprite).Height / 4) + 16
    End If

    ' Draw name
    RenderText Font_Default, Name, TextX, TextY, color, 0, True, True


   On Error GoTo 0
   Exit Sub
ErrorHandler:
    HandleError "DrawPlayerPetName", "modText", Err.Number, Err.Description, Erl
    Err.Clear
    
End Sub

