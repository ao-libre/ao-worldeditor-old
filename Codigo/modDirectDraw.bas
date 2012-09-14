Attribute VB_Name = "modDirectDraw"
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

''
' modDirectDraw
'
' @remarks Funciones de DirectDraw y Visualizacion
' @author unkwown
' @version 0.0.20
' @date 20061015

Option Explicit

Public bTecho       As Boolean 'hay techo?

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Function LoadWavetoDSBuffer(ByRef DS As DirectSound, ByRef DSB As DirectSoundBuffer, ByRef sfile As String) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
    
    Dim bufferDesc As DSBUFFERDESC
    Dim waveFormat As WAVEFORMATEX
    
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    
    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2
    waveFormat.lSamplesPerSec = 22050
    waveFormat.nBitsPerSample = 16
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
    Set DSB = DS.CreateSoundBufferFromFile(sfile, bufferDesc, waveFormat)
    
    If Err.Number <> 0 Then
        Exit Function
    End If
    
    LoadWavetoDSBuffer = True
End Function

Function DeInitTileEngine() As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 26/05/06
'*************************************************
Dim loopc As Integer

'****** Clear DirectX objects ******
Set PrimarySurface = Nothing
Set PrimaryClipper = Nothing
Set BackBufferSurface = Nothing

Set SurfaceDB = Nothing

Set DirectDraw = Nothing

'Reset any channels that are done
For loopc = 1 To NumSoundBuffers
    Set DSBuffers(loopc) = Nothing
Next loopc

Set DirectSound = Nothing

Set DirectX = Nothing

DeInitTileEngine = True

End Function

Sub ShowNextFrame(ByVal DisplayFormTop As Integer, ByVal DisplayFormLeft As Integer, ByVal MouseViewX As Integer, ByVal MouseViewY As Integer)
'***************************************************
'Author: Arron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Updates the game's model and renders everything.
'***************************************************
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
    
    '****** Set main view rectangle ******
    MainViewRect.Left = (DisplayFormLeft / Screen.TwipsPerPixelX) + MainViewLeft
    MainViewRect.Top = (DisplayFormTop / Screen.TwipsPerPixelY) + MainViewTop
    MainViewRect.Right = MainViewRect.Left + MainViewWidth
    MainViewRect.Bottom = MainViewRect.Top + MainViewHeight
    
    If UserMoving Then
        '****** Move screen Left and Right if needed ******
        If AddtoUserPos.X <> 0 Then
            OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame
            If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                OffsetCounterX = 0
                AddtoUserPos.X = 0
                UserMoving = False
            End If
        End If
        
        '****** Move screen Up and Down if needed ******
        If AddtoUserPos.y <> 0 Then
            OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.y * timerTicksPerFrame
            If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.y) Then
                OffsetCounterY = 0
                AddtoUserPos.y = 0
                UserMoving = False
            End If
        End If
    End If
    
    '****** Update screen ******
    Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY)
    
    'Display front-buffer!
    Call PrimarySurface.Blt(MainViewRect, BackBufferSurface, MainDestRect, DDBLT_WAIT)
    
    'Limit FPS to 100 (an easy number higher than monitor's vertical refresh rates)
    While (DirectX.TickCount - fpsLastCheck) \ 10 < FramesPerSecCounter
        Sleep 5
    Wend
    
    'FPS update
    If fpsLastCheck + 1000 < DirectX.TickCount Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 1
        fpsLastCheck = DirectX.TickCount
    Else
        FramesPerSecCounter = FramesPerSecCounter + 1
    End If
    
    'Get timing info
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim X As Integer
    Dim y As Integer
    Dim tx As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            y = -1
        
        Case E_Heading.EAST
            X = 1
        
        Case E_Heading.SOUTH
            y = 1
        
        Case E_Heading.WEST
            X = -1
    End Select
    
    'Fill temp pos
    tx = UserPos.X + X
    tY = UserPos.y + y
    
    'Check to see if its out of bounds
    If tx < MinXBorder Or tx > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.X = X
        UserPos.X = tx
        AddtoUserPos.y = y
        UserPos.y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.X, UserPos.y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.y).Trigger = 4, True, False)
    End If
End Sub

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tx As Integer, ByRef tY As Integer)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    tx = UserPos.X + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
    tY = UserPos.y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2
End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal y As Integer)
On Error Resume Next
    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    With CharList(CharIndex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .Active = 0 Then _
            NumChars = NumChars + 1
        
        .iHead = Head
        .iBody = Body
        .Head = HeadData(Head)
        .Body = BodyData(Body)
        
        .Heading = Heading
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.X = X
        .Pos.y = y
        
        'Make active
        .Active = 1
    End With
    
    'Plot on map
    MapData(X, y).CharIndex = CharIndex
    
    bRefreshRadar = True ' GS
End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)
    With CharList(CharIndex)
        .Active = 0
        
        .Moving = 0
        .Pos.X = 0
        .Pos.y = 0
    End With
End Sub

Sub EraseChar(ByVal CharIndex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
On Error Resume Next
    If CharIndex = 0 Then Exit Sub
    
    CharList(CharIndex).Active = 0
    
    'Update lastchar
    If CharIndex = LastChar Then
        Do Until CharList(LastChar).Active = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    
    MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.y).CharIndex = 0

    Call ResetCharInfo(CharIndex)
    
    'Update NumChars
    NumChars = NumChars - 1

    bRefreshRadar = True ' GS
End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal grhIndex As Integer, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    Grh.grhIndex = grhIndex
    
    If grhIndex = 0 Then Exit Sub
    
    If Started = 2 Then
        If GrhData(Grh.grhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.grhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.grhIndex).Speed
End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim addX As Integer
    Dim addY As Integer
    Dim X As Integer
    Dim y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    With CharList(CharIndex)
        X = .Pos.X
        y = .Pos.y
        
        'Figure out which way to move
        Select Case nHeading
            Case E_Heading.NORTH
                addY = -1
        
            Case E_Heading.EAST
                addX = 1
        
            Case E_Heading.SOUTH
                addY = 1
            
            Case E_Heading.WEST
                addX = -1
        End Select
        
        nX = X + addX
        nY = y + addY
        
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.y = nY
        MapData(X, y).CharIndex = 0
        
        .MoveOffsetX = -1 * (TilePixelWidth * addX)
        .MoveOffsetY = -1 * (TilePixelHeight * addY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = addX
        .scrollDirectionY = addY
    End With
    
    'areas viejos
    'If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
    '    If CharIndex <> UserCharIndex Then
    '        Call EraseChar(CharIndex)
    '    End If
    'End If
End Sub

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
On Error Resume Next
    Dim X As Integer
    Dim y As Integer
    Dim addX As Integer
    Dim addY As Integer
    Dim nHeading As E_Heading
    
    With CharList(CharIndex)
        X = .Pos.X
        y = .Pos.y
        
        MapData(X, y).CharIndex = 0
        
        addX = nX - X
        addY = nY - y
        
        If Sgn(addX) = 1 Then
            nHeading = E_Heading.EAST
        ElseIf Sgn(addX) = -1 Then
            nHeading = E_Heading.WEST
        ElseIf Sgn(addY) = -1 Then
            nHeading = E_Heading.NORTH
        ElseIf Sgn(addY) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        
        .Pos.X = nX
        .Pos.y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * addX)
        .MoveOffsetY = -1 * (TilePixelHeight * addY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addX)
        .scrollDirectionY = Sgn(addY)
    End With

    bRefreshRadar = True ' GS
End Sub

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
    Dim loopc As Long
    
    loopc = 1
    Do While (CharList(loopc).Active = 1) And (loopc <= UBound(CharList))
        loopc = loopc + 1
    Loop
    
    NextOpenChar = loopc
End Function

Function LegalPos(ByVal X As Integer, ByVal y As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 - GS
'*************************************************

LegalPos = True

'Check to see if its out of bounds
If Not InMapLegalBounds(X, y) Then
    LegalPos = False
    Exit Function
End If

'Check to see if its blocked
If MapData(X, y).Blocked = 1 Then
    LegalPos = False
    Exit Function
End If

'Check for character
If MapData(X, y).CharIndex > 0 Then
    LegalPos = False
    Exit Function
End If

End Function

Function InMapLegalBounds(ByVal X As Integer, ByVal y As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

If (X < MinXBorder) Or (X > MaxXBorder) Or (y < MinYBorder) Or (y > MaxYBorder) Then
    InMapLegalBounds = False
    Exit Function
End If

InMapLegalBounds = True

End Function

Public Sub DDrawGrhtoSurface(ByRef Surface As DirectDrawSurface7, ByRef Grh As Grh, ByVal X As Integer, ByVal y As Integer, ByVal Center As Byte, ByVal Animate As Byte)
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
On Error GoTo error
        
    If Grh.grhIndex = 0 Then Exit Sub
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.grhIndex).NumFrames / Grh.Speed)
            If Grh.FrameCounter > GrhData(Grh.grhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.grhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.grhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call Surface.BltFast(X, y, SurfaceDB.Surface(.FileNum), SourceRect, DDBLTFAST_WAIT)
    End With
Exit Sub

error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.Number & " ] Error"
        End
    End If
End Sub

Public Sub DDrawTransGrhIndextoSurface(ByRef Surface As DirectDrawSurface7, ByVal grhIndex As Integer, ByVal X As Integer, ByVal y As Integer, ByVal Center As Byte)
    Dim SourceRect As RECT
    
    If grhIndex = 0 Then Exit Sub
    
    With GrhData(grhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call Surface.BltFast(X, y, SurfaceDB.Surface(.FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
    End With
End Sub

Public Sub DDrawTransGrhtoSurface(ByRef Surface As DirectDrawSurface7, ByRef Grh As Grh, ByVal X As Integer, ByVal y As Integer, ByVal Center As Byte, ByVal Animate As Byte)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
    Dim ddsdDest As DDSURFACEDESC2
    
On Error GoTo error

    If Grh.grhIndex = 0 Then Exit Sub
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.grhIndex).NumFrames / Grh.Speed)
            
            If Grh.FrameCounter > GrhData(Grh.grhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.grhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.grhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
                
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call Surface.BltFast(X, y, SurfaceDB.Surface(.FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
    End With
Exit Sub

error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.Number & " ] Error"
        End
    End If
End Sub

Sub DrawBackBufferSurface()
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
PrimarySurface.Blt MainViewRect, BackBufferSurface, MainDestRect, DDBLT_WAIT
End Sub

Sub DrawGrhtoHdc(ByVal hdc As Long, ByVal grhIndex As Integer, ByRef SourceRect As RECT, ByRef destRect As RECT)
'*****************************************************************
'Draws a Grh's portion to the given area of any Device Context
'*****************************************************************
    Call SurfaceDB.Surface(GrhData(grhIndex).FileNum).BltToDC(hdc, SourceRect, destRect)
End Sub

Sub PlayWaveDS(ByRef file As String)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

    'Cylce through avaiable sound buffers
    LastSoundBufferUsed = LastSoundBufferUsed + 1
    If LastSoundBufferUsed > NumSoundBuffers Then
        LastSoundBufferUsed = 1
    End If
    
    If LoadWavetoDSBuffer(DirectSound, DSBuffers(LastSoundBufferUsed), file) Then
        DSBuffers(LastSoundBufferUsed).Play DSBPLAY_DEFAULT
    End If

End Sub

Sub RenderScreen(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 31/05/06 by GS
'*************************************************
Dim y           As Long     'Keeps track of where on map we are
Dim X           As Long     'Keeps track of where on map we are
Dim screenminY  As Integer  'Start Y pos on current screen
Dim screenmaxY  As Integer  'End Y pos on current screen
Dim screenminX  As Integer  'Start X pos on current screen
Dim screenmaxX  As Integer  'End X pos on current screen
Dim MinY        As Integer  'Start Y pos on current map
Dim MaxY        As Integer  'End Y pos on current map
Dim MinX        As Integer  'Start X pos on current map
Dim MaxX        As Integer  'End X pos on current map
Dim ScreenX     As Integer  'Keeps track of where to place tile on screen
Dim ScreenY     As Integer  'Keeps track of where to place tile on screen
Dim ScreenXOffset   As Integer
Dim ScreenYOffset   As Integer
Dim minXOffset  As Integer
Dim minYOffset  As Integer
Dim PixelOffsetXTemp As Integer 'For centering grhs
Dim PixelOffsetYTemp As Integer 'For centering grhs
Dim Grh         As Grh      'Temp Grh for show tile and blocked
                    
    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    MinY = screenminY - TileBufferSize
    MaxY = screenmaxY + TileBufferSize
    MinX = screenminX - TileBufferSize
    MaxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If MinY < YMinMapSize Then
        minYOffset = YMinMapSize - MinY
        MinY = YMinMapSize
    End If
    
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
    
    If MinX < XMinMapSize Then
        minXOffset = XMinMapSize - MinX
        MinX = XMinMapSize
    End If
    
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        ScreenYOffset = (YMinMapSize - screenminY) + 1
        screenminY = YMinMapSize
    End If
    
    If screenmaxY < YMaxMapSize Then
        screenmaxY = screenmaxY + 1
    ElseIf screenmaxY > YMaxMapSize Then
        screenmaxY = YMaxMapSize
    End If
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        ScreenXOffset = (XMinMapSize - screenminX) + 1
        screenminX = XMinMapSize
    End If
    
    If screenmaxX < XMaxMapSize Then
        screenmaxX = screenmaxX + 1
    ElseIf screenmaxX > XMaxMapSize Then
        screenmaxX = XMaxMapSize
    End If
    
    Call CleanViewPort
    
    'Draw floor layer
    ScreenY = ScreenYOffset
    For y = screenminY To screenmaxY
        ScreenX = ScreenXOffset
        For X = screenminX To screenmaxX
            
            'Layer 1 **********************************
            If MapData(X, y).Graphic(1).grhIndex <> 0 Then
                Call DDrawGrhtoSurface(BackBufferSurface, MapData(X, y).Graphic(1), _
                    (ScreenX - 1) * TilePixelWidth + PixelOffsetX + TileBufferPixelOffsetX, _
                    (ScreenY - 1) * TilePixelHeight + PixelOffsetY + TileBufferPixelOffsetY, _
                    0, 1)
            End If
                
            If bSelectSup Then
                If CurLayer = 1 Then
                    If X = SobreX And y = SobreY Then
                        If MosaicoChecked Then
                            Call DDrawGrhtoSurface(BackBufferSurface, CurrentGrh(((X + DespX) Mod mAncho) + 1, ((y + DespY) Mod MAlto) + 1), _
                                (ScreenX - 1) * TilePixelWidth + PixelOffsetX + TileBufferPixelOffsetX, _
                                (ScreenY - 1) * TilePixelHeight + PixelOffsetY + TileBufferPixelOffsetY, _
                                0, 1)
                        Else
                            Call DDrawGrhtoSurface(BackBufferSurface, CurrentGrh(0), _
                                (ScreenX - 1) * TilePixelWidth + PixelOffsetX + TileBufferPixelOffsetX, _
                                (ScreenY - 1) * TilePixelHeight + PixelOffsetY + TileBufferPixelOffsetY, _
                                0, 1)
                        End If
                    End If
                End If
            End If
            '******************************************
            
            ScreenX = ScreenX + 1
        Next X
        
        'Increment ScreenY
        ScreenY = ScreenY + 1
    Next y
        
    If bVerCapa(2) Then
        'Draw floor layer 2
        ScreenY = minYOffset
        For y = MinY To MaxY
            ScreenX = minXOffset
            For X = MinX To MaxX
                
                'Layer 2 **********************************
                If MapData(X, y).Graphic(2).grhIndex <> 0 Then
                    Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(X, y).Graphic(2), _
                            (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                            (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                            1, 1)
                End If
                
                If bSelectSup Then
                    If CurLayer = 2 Then
                        If (X = SobreX) And (y = SobreY) Then
                            If MosaicoChecked Then
                                Call DDrawTransGrhtoSurface(BackBufferSurface, CurrentGrh(((X + DespX) Mod mAncho) + 1, ((y + DespY) Mod MAlto) + 1), _
                                    (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                                    (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                                    1, 1)
                            Else
                                Call DDrawTransGrhtoSurface(BackBufferSurface, CurrentGrh(0), _
                                    (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                                    (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                                    1, 1)
                            End If
                        End If
                    End If
                End If
                '******************************************
                
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next y
    Else
        If bSelectSup Then
            If CurLayer = 2 Then
                X = SobreX
                y = SobreY
                ScreenX = (X - MinX) + minXOffset
                ScreenY = (y - MinY) + minYOffset
                
                If MosaicoChecked Then
                    Call DDrawTransGrhtoSurface(BackBufferSurface, CurrentGrh(((X + DespX) Mod mAncho) + 1, ((y + DespY) Mod MAlto) + 1), _
                        (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                        (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                        1, 1)
                Else
                    Call DDrawTransGrhtoSurface(BackBufferSurface, CurrentGrh(0), _
                        (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                        (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                        1, 1)
                End If
            End If
        End If
    End If
    
    'Draw Transparent Layers
    ScreenY = minYOffset
    For y = MinY To MaxY
        ScreenX = minXOffset
        For X = MinX To MaxX
            PixelOffsetXTemp = (ScreenX - 1) * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = (ScreenY - 1) * TilePixelHeight + PixelOffsetY
            
            With MapData(X, y)
                'Object Layer **********************************
                If (.ObjGrh.grhIndex <> 0) And bVerObjetos Then
                    Call DDrawTransGrhtoSurface(BackBufferSurface, .ObjGrh, _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                End If
                '***********************************************
                
                
                'Char layer ************************************
                If (.CharIndex <> 0) And bVerNpcs Then
                    Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp)
                End If
                '*************************************************
                
                
                'Layer 3 *****************************************
                If (.Graphic(3).grhIndex <> 0) And bVerCapa(3) Then
                    'Draw
                    Call DDrawTransGrhtoSurface(BackBufferSurface, .Graphic(3), _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                End If
                '************************************************
            End With
            
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next y
    
    Grh.FrameCounter = 1
    Grh.Started = 0
    
    If bVerCapa(4) Then
        'Draw layer 4
        ScreenY = minYOffset
        For y = MinY To MaxY
            ScreenX = minXOffset
            For X = MinX To MaxX
                With MapData(X, y)
                    'Layer 4 **********************************
                    If .Graphic(4).grhIndex <> 0 Then
                        'Draw
                        Call DDrawTransGrhtoSurface(BackBufferSurface, .Graphic(4), _
                            (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                            (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                            1, 1)
                    End If
                    '**********************************
                End With
                
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next y
    End If
    
    'Draw trans, bloqs, triggers and select tiles
    ScreenY = ScreenYOffset
    For y = screenminY To screenmaxY
        ScreenX = ScreenXOffset
        For X = screenminX To screenmaxX
            With MapData(X, y)
                PixelOffsetXTemp = (ScreenX - 1) * TilePixelWidth + PixelOffsetX + TileBufferPixelOffsetX
                PixelOffsetYTemp = (ScreenY - 1) * TilePixelHeight + PixelOffsetY + TileBufferPixelOffsetY
                
                '**********************************
                If (.TileExit.Map <> 0) And bTranslados Then
                    Grh.grhIndex = 3
                    
                    Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, _
                        PixelOffsetXTemp, _
                        PixelOffsetYTemp, _
                        1, 0)
                End If
                
                'Show blocked tiles
                If (.Blocked = 1) And bBloqs Then
                    Grh.grhIndex = 4
                    
                    Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, _
                        PixelOffsetXTemp, _
                        PixelOffsetYTemp, _
                        1, 0)
                    
                    'BackBufferSurface.SetFillColor vbRed
                    
                    'Call BackBufferSurface.DrawBox( _
                        PixelOffsetXTemp + TilePixelWidth \ 2, _
                        PixelOffsetYTemp + TilePixelHeight \ 2, _
                        (PixelOffsetXTemp + 5) + TilePixelWidth \ 2, _
                        (PixelOffsetYTemp + 5) + TilePixelHeight \ 2)
                End If
                
                If bTriggers Then
                    Call TextDrawer.AddText(PixelOffsetXTemp + TilePixelWidth \ 2, PixelOffsetYTemp + TilePixelHeight \ 2, vbRed, str(.Trigger), True)
                End If
                
                If .Select Then
                    BackBufferSurface.SetForeColor vbGreen
                    BackBufferSurface.SetFillStyle 1
                    BackBufferSurface.DrawBox PixelOffsetXTemp, PixelOffsetYTemp, PixelOffsetXTemp + TilePixelWidth, PixelOffsetYTemp + TilePixelHeight
                End If
                '******************************************
            
                ScreenX = ScreenX + 1
            End With
        Next X
        
        'Increment ScreenY
        ScreenY = ScreenY + 1
    Next y
    
    Dim DC As Long
    
    DC = BackBufferSurface.GetDC
    
    Call TextDrawer.DrawTextToDC(DC)
    Call BackBufferSurface.ReleaseDC(DC)
End Sub

Private Sub CharRender(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 16/09/2010 (Zama)
'Draw char's to screen without offcentering them
'16/09/2010: ZaMa - Ya no se dibujan los bodies cuando estan invisibles.
'***************************************************
    Dim moved As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim color As Long
    
    With CharList(CharIndex)
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                        (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
        
        'If done moving stop animation
        If Not moved Then
            'Stop animations
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
            
            .Moving = False
        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        'Draw Body
        If .Body.Walk(.Heading).grhIndex Then _
            Call DDrawTransGrhtoSurface(BackBufferSurface, .Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
    
        'Draw Head
        If .Head.Head(.Heading).grhIndex Then _
            Call DDrawTransGrhtoSurface(BackBufferSurface, .Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.y, 1, 0)
    End With
End Sub

Public Sub RenderText(ByVal lngXPos As Integer, ByVal lngYPos As Integer, ByRef strText As String, ByVal lngColor As Long)
    If LenB(strText) > 0 Then
        'TextDrawer.DrawText lngXPos - 2, lngYPos - 1, strText, vbBlack, BackBufferSurface
        TextDrawer.DrawText lngXPos, lngYPos, strText, lngColor, BackBufferSurface
    End If
End Sub

Public Function InitTileEngine(ByVal setDisplayFormhWnd As Long, ByVal setMainViewTop As Integer, ByVal setMainViewLeft As Integer, ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal setWindowTileHeight As Integer, ByVal setWindowTileWidth As Integer, ByVal setTileBufferSize As Integer, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer, ByVal engineSpeed As Single) As Boolean
'***************************************************
'Author: Aaron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Creates all DX objects and configures the engine to start running.
'***************************************************
    Dim surfaceDesc As DDSURFACEDESC2
    Dim ddck As DDCOLORKEY
    
    'Fill startup variables
    MainViewTop = setMainViewTop
    MainViewLeft = setMainViewLeft
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    TileBufferSize = setTileBufferSize
    
    HalfWindowTileHeight = setWindowTileHeight \ 2
    HalfWindowTileWidth = setWindowTileWidth \ 2
    
    'Compute offset in pixels when rendering tile buffer.
    'We diminish by one to get the top-left corner of the tile for rendering.
    TileBufferPixelOffsetX = ((TileBufferSize - 1) * TilePixelWidth)
    TileBufferPixelOffsetY = ((TileBufferSize - 1) * TilePixelHeight)
    
    engineBaseSpeed = engineSpeed
    
    'Set FPS value to 60 for startup
    FPS = 60
    FramesPerSecCounter = 60
    
    MinXBorder = XMinMapSize + (ClienteWidth \ 2)
    MaxXBorder = XMaxMapSize - (ClienteWidth \ 2)
    MinYBorder = YMinMapSize + (ClienteHeight \ 2)
    MaxYBorder = YMaxMapSize - (ClienteHeight \ 2)
    
    MainViewWidth = TilePixelWidth * WindowTileWidth
    MainViewHeight = TilePixelHeight * WindowTileHeight
    
    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set intial user position
    UserPos.X = MinXBorder
    UserPos.y = MinYBorder
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY
    
    'Set the view rect
    With MainViewRect
        .Left = MainViewLeft
        .Top = MainViewTop
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
    'Set the dest rect
    With MainDestRect
        .Left = TilePixelWidth * TileBufferSize - TilePixelWidth
        .Top = TilePixelHeight * TileBufferSize - TilePixelHeight
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
On Error Resume Next
    Set DirectX = New DirectX7
    
    If Err Then
        MsgBox "No se puede iniciar DirectX. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Function
    End If

    
    '****** INIT DirectDraw ******
    ' Create the root DirectDraw object
    Set DirectDraw = DirectX.DirectDrawCreate("")
    
    If Err Then
        MsgBox "No se puede iniciar DirectDraw. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Function
    End If
    
On Error GoTo 0
    Call DirectDraw.SetCooperativeLevel(setDisplayFormhWnd, DDSCL_NORMAL)
    
    'Primary Surface
    ' Fill the surface description structure
    With surfaceDesc
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    End With
    ' Create the surface
    Set PrimarySurface = DirectDraw.CreateSurface(surfaceDesc)
    
    'Create Primary Clipper
    Set PrimaryClipper = DirectDraw.CreateClipper(0)
    Call PrimaryClipper.SetHWnd(frmMain.hwnd)
    Call PrimarySurface.SetClipper(PrimaryClipper)
    
    With BackBufferRect
        .Left = 0
        .Top = 0
        .Right = TilePixelWidth * (WindowTileWidth + 2 * TileBufferSize)
        .Bottom = TilePixelHeight * (WindowTileHeight + 2 * TileBufferSize)
    End With
    
    With surfaceDesc
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        If ClientSetup.bUseVideo Then
            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
        Else
            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        End If
        .lHeight = BackBufferRect.Bottom
        .lWidth = BackBufferRect.Right
    End With
    
    ' Create surface
    Set BackBufferSurface = DirectDraw.CreateSurface(surfaceDesc)
    
    'Set color key
    ddck.low = 0
    ddck.high = 0
    Call BackBufferSurface.SetColorKey(DDCKEY_SRCBLT, ddck)
    
    'Set font transparency
    Call BackBufferSurface.SetFontTransparency(D_TRUE)
    
    'Load graphic data into memory
    modIndices.CargarIndicesDeGraficos

    frmCargando.X.Caption = "Iniciando Control de Superficies..."
    Call SurfaceDB.Initialize(DirectDraw, ClientSetup.bUseVideo, DirGraficos, ClientSetup.byMemory)
    
    'Wave Sound
    Set DirectSound = DirectX.DirectSoundCreate("")
    DirectSound.SetCooperativeLevel setDisplayFormhWnd, DSSCL_PRIORITY
    LastSoundBufferUsed = 1
    
    InitTileEngine = True
End Function

Private Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Private Sub CleanViewPort()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Fills the viewport with black.
'***************************************************
    Dim r As RECT
    Call BackBufferSurface.BltColorFill(r, vbBlack)
End Sub
