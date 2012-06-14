Attribute VB_Name = "modGeneral"
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
' modGeneral
'
' @remarks Funciones Generales
' @author unkwown
' @version 0.4.11
' @date 20061015

Option Explicit

Private lFrameTimer As Long

''
' Realiza acciones de desplasamiento segun las teclas que hallamos precionado
'

Public Sub CheckKeys()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
Static LastMovement As Long
    
    If Not Application.IsAppActive Then Exit Sub
    If Not HotKeysAllow Then Exit Sub
        
    If WalkMode Then
        'Control movement interval (this enforces the 1 step loss when meditating / resting client-side)
        If GetTickCount - LastMovement > 56 Then
            LastMovement = GetTickCount
        Else
            Exit Sub
        End If
    End If

    If GetAsyncKeyState(vbKeyUp) < 0 Then
        If UserPos.y > YMinMapSize Then
            If WalkMode And (UserMoving = 0) Then
                MoveTo E_Heading.NORTH
            ElseIf WalkMode = False Then
                UserPos.y = UserPos.y - 1
            End If
            
            bRefreshRadar = True ' Radar
        End If
    End If

    If GetAsyncKeyState(vbKeyRight) < 0 Then
        If UserPos.X < XMaxMapSize Then
            If WalkMode And (UserMoving = 0) Then
                MoveTo E_Heading.EAST
            ElseIf WalkMode = False Then
                UserPos.X = UserPos.X + 1
            End If
            
            bRefreshRadar = True ' Radar
        End If
    End If

    If GetAsyncKeyState(vbKeyDown) < 0 Then
        If UserPos.y < YMaxMapSize Then
            If WalkMode And (UserMoving = 0) Then
                MoveTo E_Heading.SOUTH
            ElseIf WalkMode = False Then
                UserPos.y = UserPos.y + 1
            End If
            
            bRefreshRadar = True ' Radar
        End If
    End If

    If GetAsyncKeyState(vbKeyLeft) < 0 Then
        If UserPos.X > XMinMapSize Then
            If WalkMode And (UserMoving = 0) Then
                MoveTo E_Heading.WEST
            ElseIf WalkMode = False Then
                UserPos.X = UserPos.X - 1
            End If
            
            bRefreshRadar = True ' Radar
        End If
    End If
End Sub

Public Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Integer) As String
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim i As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String

Seperator = Chr$(SepASCII)
LastPos = 0
FieldNum = 0

For i = 1 To Len(Text)
    CurChar = mid$(Text, i, 1)
    If CurChar = Seperator Then
        FieldNum = FieldNum + 1
        
        If FieldNum = Pos Then
            ReadField = mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = i
    End If
Next i
FieldNum = FieldNum + 1

If FieldNum = Pos Then
    ReadField = mid$(Text, LastPos + 1)
End If

End Function

''
' Completa y corrije un path
'
' @param Path Especifica el path con el que se trabajara
' @return   Nos devuelve el path completado

Private Function autoCompletaPath(ByVal path As String) As String
'*************************************************
'Author: ^[GS]^
'Last modified: 22/05/06
'*************************************************
path = Replace(path, "/", "\")

If Left$(path, 1) = "\" Then
    ' agrego app.path & path
    path = App.path & path
End If

If Right$(path, 1) <> "\" Then
    ' me aseguro que el final sea con "\"
    path = path & "\"
End If

autoCompletaPath = path
End Function

''
' Carga la configuracion del WorldEditor de WorldEditor.ini
'

Private Sub CargarMapIni()
'*************************************************
'Author: ^[GS]^
'Last modified: 02/10/06
'*************************************************
On Error GoTo Fallo
Dim tStr As String
Dim Leer As New clsIniReader
Dim i As Long

IniPath = App.path & "\"

If Not FileExist(IniPath & "WorldEditor.ini", vbArchive) Then
    frmMain.mnuGuardarUltimaConfig.Checked = True
    DirGraficos = IniPath & "Graficos\"
    DirIndex = IniPath & "INIT\"
    DirMidi = IniPath & "MIDI\"
    DirMp3 = IniPath & "MP3\"
    DirDats = IniPath & "DATS\"
    UserPos.X = 50
    UserPos.y = 50
    PantallaX = 21
    PantallaY = 19
    
    MsgBox "Falta el archivo 'WorldEditor.ini' de configuración.", vbInformation
    Exit Sub
End If

Call Leer.Initialize(IniPath & "WorldEditor.ini")

' Obj de Translado
Cfg_TrOBJ = Val(Leer.GetValue("CONFIGURACION", "ObjTranslado"))
frmMain.mnuAutoCapturarTranslados.Checked = Val(Leer.GetValue("CONFIGURACION", "AutoCapturarTrans"))
frmMain.mnuAutoCapturarSuperficie.Checked = Val(Leer.GetValue("CONFIGURACION", "AutoCapturarSup"))
frmMain.mnuUtilizarDeshacer.Checked = Val(Leer.GetValue("CONFIGURACION", "UtilizarDeshacer"))

' Guardar Ultima Configuracion
frmMain.mnuGuardarUltimaConfig.Checked = Val(Leer.GetValue("CONFIGURACION", "GuardarConfig"))

'Reciente
frmMain.Dialog.InitDir = autoCompletaPath(Leer.GetValue("PATH", "UltimoMapa"))
DirGraficos = autoCompletaPath(Leer.GetValue("PATH", "DirGraficos"))

If DirGraficos = "\" Then
    DirGraficos = IniPath & "Graficos\"
End If

If FileExist(DirGraficos, vbDirectory) = False Then
    MsgBox "El directorio de Graficos es incorrecto", vbCritical + vbOKOnly
    End
End If

DirMidi = autoCompletaPath(Leer.GetValue("PATH", "DirMidi"))

If DirMidi = "\" Then
    DirMidi = IniPath & "MIDI\"
End If

DirMp3 = autoCompletaPath(Leer.GetValue("PATH", "DirMp3"))

If DirMp3 = "\" Then
    DirMp3 = IniPath & "MP3\"
End If

If FileExist(DirMidi, vbDirectory) = False Then
    MsgBox "El directorio de MIDI es incorrecto", vbCritical + vbOKOnly
    End
End If

DirIndex = autoCompletaPath(Leer.GetValue("PATH", "DirIndex"))

If DirIndex = "\" Then
    DirIndex = IniPath & "INIT\"
End If

If FileExist(DirIndex, vbDirectory) = False Then
    MsgBox "El directorio de Index es incorrecto", vbCritical + vbOKOnly
    End
End If

DirDats = autoCompletaPath(Leer.GetValue("PATH", "DirDats"))

If DirDats = "\" Then
    DirDats = IniPath & "DATS\"
End If

If FileExist(DirDats, vbDirectory) = False Then
    MsgBox "El directorio de Dats es incorrecto", vbCritical + vbOKOnly
    End
End If

tStr = Leer.GetValue("MOSTRAR", "LastPos") ' x-y
UserPos.X = Val(ReadField(1, tStr, Asc("-")))
UserPos.y = Val(ReadField(2, tStr, Asc("-")))

If UserPos.X < XMinMapSize Or UserPos.X > XMaxMapSize Then
    UserPos.X = 50
End If

If UserPos.y < YMinMapSize Or UserPos.y > YMaxMapSize Then
    UserPos.y = 50
End If

' Menu Mostrar
frmMain.mnuVerAutomatico.Checked = Val(Leer.GetValue("MOSTRAR", "ControlAutomatico"))

For i = 2 To 4
    bVerCapa(i) = Val(Leer.GetValue("MOSTRAR", "Capa" & i))
    frmMain.mnuVerCapa(i).Checked = bVerCapa(i)
Next i

bTranslados = Val(Leer.GetValue("MOSTRAR", "Translados"))
bTriggers = Val(Leer.GetValue("MOSTRAR", "Triggers"))
bBloqs = Val(Leer.GetValue("MOSTRAR", "Bloqueos"))
bVerNpcs = Val(Leer.GetValue("MOSTRAR", "NPCs"))
bVerObjetos = Val(Leer.GetValue("MOSTRAR", "Objetos"))

frmMain.mnuVerTranslados.Checked = bTranslados
frmMain.mnuVerObjetos.Checked = bVerObjetos
frmMain.mnuVerNPCs.Checked = bVerNpcs
frmMain.mnuVerTriggers.Checked = bTriggers
frmMain.mnuVerBloqueos.Checked = bBloqs

frmMain.cVerTriggers.value = bTriggers
frmMain.cVerBloqueos.value = bBloqs

' Tamaño de visualizacion
PantallaX = Val(Leer.GetValue("MOSTRAR", "PantallaX"))
PantallaY = Val(Leer.GetValue("MOSTRAR", "PantallaY"))

If PantallaX > 27 Or PantallaX <= 3 Then PantallaX = 21
If PantallaY > 25 Or PantallaY <= 3 Then PantallaY = 19

ClienteHeight = Val(Leer.GetValue("MOSTRAR", "ClienteHeight"))
ClienteWidth = Val(Leer.GetValue("MOSTRAR", "ClienteWidth"))

NumMaps = Val(GetVar(DirDats & "Map.dat", "INIT", "NumMaps"))
Exit Sub
Fallo:
    MsgBox "ERROR " & Err.Number & " en WorldEditor.ini" & vbCrLf & Err.Description, vbCritical
    Resume Next
End Sub

Function MoveToLegalPos(ByVal X As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 01/08/2009
'Checks to see if a tile position is legal, including if there is a casper in the tile
'10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
'01/08/2009: ZaMa - Now invisible admins can't change position with caspers.
'*****************************************************************
    Dim CharIndex As Integer
    
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(X, y).Blocked = 1 Then
        Exit Function
    End If
    
    CharIndex = MapData(X, y).CharIndex
    '¿Hay un personaje?
    If CharIndex > 0 Then
        Exit Function
    End If
    
    MoveToLegalPos = True
End Function

Sub MoveTo(ByVal Direccion As E_Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/28/2008
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
' 06/28/2008: NicoNZ - Saqué lo que impedía que si el usuario estaba paralizado se ejecute el sub.
'***************************************************
    Dim LegalOk As Boolean
    
    Select Case Direccion
        Case E_Heading.NORTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.y - 1)
        Case E_Heading.EAST
            LegalOk = MoveToLegalPos(UserPos.X + 1, UserPos.y)
        Case E_Heading.SOUTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.y + 1)
        Case E_Heading.WEST
            LegalOk = MoveToLegalPos(UserPos.X - 1, UserPos.y)
    End Select
    
    If LegalOk Then
        MoveCharbyHead UserCharIndex, Direccion
        MoveScreen Direccion
    Else
        CharList(UserCharIndex).Heading = Direccion
    End If
End Sub

Public Sub Main()
'*************************************************
'Author: Unkwown
'Last modified: 15/10/06 - GS
'*************************************************
On Error Resume Next

Dim OffsetCounterX As Integer
Dim OffsetCounterY As Integer
Dim Chkflag As Integer

    If App.PrevInstance Then End
    
    'Load ao.dat config file
    Call LoadClientSetup
    
    If ClientSetup.bDinamic Then
        Set SurfaceDB = New clsSurfaceManDyn
    Else
        Set SurfaceDB = New clsSurfaceManStatic
    End If
    
    Call CargarMapIni
    Call IniciarCabecera(MiCabecera)
    DoEvents
    
    If FileExist(IniPath & "WorldEditor.jpg", vbArchive) Then frmCargando.Picture1.Picture = LoadPicture(IniPath & "WorldEditor.jpg")
    
    frmCargando.verX = "v" & App.Major & "." & App.Minor & "." & App.Revision
    frmCargando.Show
    frmCargando.SetFocus
    frmCargando.X.Caption = "Iniciando DirectSound..."
    DoEvents
    
    IniciarDirectSound
    frmCargando.X.Caption = "Cargando Indice de Superficies..."
    modIndices.CargarIndicesSuperficie
    frmCargando.X.Caption = "Indexando Cargado de Imagenes..."
    DoEvents
    
    If InitTileEngine(frmMain.hwnd, frmMain.MainViewShp.Top + 50, frmMain.MainViewShp.Left + 4, 32, 32, PantallaY, PantallaX, 9, 8, 8, 0.018) Then ' 30/05/2006
        'Display form handle, View window offset from 0,0 of display form, Tile Size, Display size in tiles, Screen buffer
        frmCargando.P1.Visible = True
        frmCargando.L(0).Visible = True
        frmCargando.X.Caption = "Cargando Cuerpos..."
        modIndices.CargarIndicesDeCuerpos
        DoEvents
        
        frmCargando.P2.Visible = True
        frmCargando.L(1).Visible = True
        frmCargando.X.Caption = "Cargando Cabezas..."
        modIndices.CargarIndicesDeCabezas
        DoEvents
        
        frmCargando.P3.Visible = True
        frmCargando.L(2).Visible = True
        frmCargando.X.Caption = "Cargando NPC's..."
        modIndices.CargarIndicesNPC
        DoEvents
        
        frmCargando.P4.Visible = True
        frmCargando.L(3).Visible = True
        frmCargando.X.Caption = "Cargando Objetos..."
        modIndices.CargarIndicesOBJ
        DoEvents
        
        frmCargando.P5.Visible = True
        frmCargando.L(4).Visible = True
        frmCargando.X.Caption = "Cargando Triggers..."
        modIndices.CargarIndicesTriggers
        DoEvents
        
        frmCargando.P6.Visible = True
        frmCargando.L(5).Visible = True
    End If
    
    Set TextDrawer = New clsTextDrawer
    Call TextDrawer.InitText(DirectDraw, ClientSetup.bUseVideo)
    
    frmCargando.SetFocus
    frmCargando.X.Caption = "Iniciando Ventana de Edición..."
    
    frmCargando.Hide
    frmMain.Show
    modMapIO.NuevoMapa
    
    Call ActualizarMosaico
    
    prgRun = True
    FPS = 0
    Chkflag = 0
    dTiempoGT = GetTickCount
    CurLayer = 1
    
    Do While prgRun
        'Sólo dibujamos si la ventana no está minimizada
        If frmMain.WindowState <> 1 And frmMain.Visible Then
            Call ShowNextFrame(frmMain.Top, frmMain.Left, frmMain.MouseX, frmMain.MouseY)
            
            'Play ambient sounds
            'Call RenderSounds
            
            Call CheckKeys
        End If
        
        'FPS Counter - mostramos las FPS
        If GetTickCount - lFrameTimer >= 1000 Then
            CaptionWorldEditor frmMain.Dialog.FileName, (MapInfo.Changed = 1)
            frmMain.FPS.Caption = "FPS: " & FPS
            
            lFrameTimer = GetTickCount
        End If
        
        If bRefreshRadar Then Call RefreshAllChars

        'If frmMain.PreviewGrh.Visible Then Call modPaneles.VistaPreviaDeSup
        DoEvents
    Loop
        
    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            modMapIO.GuardarMapa frmMain.Dialog.FileName
        End If
    End If
    
    DeInitTileEngine
    LiberarDirectSound
    
    Call StopMusic
    
    Dim f
    
    For Each f In Forms
        Unload f
    Next
    End

End Sub

Public Function GetVar(ByRef file As String, ByRef Main As String, ByRef Var As String) As String
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found

szReturn = vbNullString
sSpaces = Space$(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish
GetPrivateProfileString Main, Var, szReturn, sSpaces, Len(sSpaces), file

GetVar = RTrim$(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Public Sub WriteVar(ByRef file As String, ByRef Main As String, ByRef Var As String, ByRef value As String)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
writeprivateprofilestring Main, Var, value, file
End Sub

Public Sub ToggleWalkMode()
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 - GS
'*************************************************
On Error GoTo fin:

WalkMode = Not WalkMode

If Not WalkMode Then
    frmMain.mnuModoCaminata.Checked = False
    
    'Erase character
    Call EraseChar(UserCharIndex)
    MapData(UserPos.X, UserPos.y).CharIndex = 0
Else
    'MakeCharacter
    If LegalPos(UserPos.X, UserPos.y) Then
        Call MakeChar(NextOpenChar(), 1, 1, SOUTH, UserPos.X, UserPos.y)
        UserCharIndex = MapData(UserPos.X, UserPos.y).CharIndex
        frmMain.mnuModoCaminata.Checked = True
    Else
        MsgBox "ERROR: Ubicacion ilegal."
        WalkMode = False
    End If
End If
fin:
End Sub

Public Sub FixCoasts(ByVal grhIndex As Integer, ByVal X As Integer, ByVal y As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

If grhIndex = 7284 Or grhIndex = 7290 Or grhIndex = 7291 Or grhIndex = 7297 Or _
   grhIndex = 7300 Or grhIndex = 7301 Or grhIndex = 7302 Or grhIndex = 7303 Or _
   grhIndex = 7304 Or grhIndex = 7306 Or grhIndex = 7308 Or grhIndex = 7310 Or _
   grhIndex = 7311 Or grhIndex = 7313 Or grhIndex = 7314 Or grhIndex = 7315 Or _
   grhIndex = 7316 Or grhIndex = 7317 Or grhIndex = 7319 Or grhIndex = 7321 Or _
   grhIndex = 7325 Or grhIndex = 7326 Or grhIndex = 7327 Or grhIndex = 7328 Or grhIndex = 7332 Or _
   grhIndex = 7338 Or grhIndex = 7339 Or grhIndex = 7345 Or grhIndex = 7348 Or _
   grhIndex = 7349 Or grhIndex = 7350 Or grhIndex = 7351 Or grhIndex = 7352 Or _
   grhIndex = 7349 Or grhIndex = 7350 Or grhIndex = 7351 Or _
   grhIndex = 7354 Or grhIndex = 7357 Or grhIndex = 7358 Or grhIndex = 7360 Or _
   grhIndex = 7362 Or grhIndex = 7363 Or grhIndex = 7365 Or grhIndex = 7366 Or _
   grhIndex = 7367 Or grhIndex = 7368 Or grhIndex = 7369 Or grhIndex = 7371 Or _
   grhIndex = 7373 Or grhIndex = 7375 Or grhIndex = 7376 Then MapData(X, y).Graphic(2).grhIndex = 0

End Sub

Public Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Randomize Timer
RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
End Function


''
' Actualiza todos los Chars en el mapa
'

Public Sub RefreshAllChars()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
On Error Resume Next
Dim loopc As Integer

frmMain.ApuntadorRadar.Move UserPos.X - 12, UserPos.y - 10
frmMain.picRadar.Cls

For loopc = 1 To LastChar
    If CharList(loopc).Active = 1 Then
        MapData(CharList(loopc).Pos.X, CharList(loopc).Pos.y).CharIndex = loopc
        If CharList(loopc).Heading <> 0 Then
            frmMain.picRadar.ForeColor = vbGreen
            frmMain.picRadar.Line (0 + CharList(loopc).Pos.X, 0 + CharList(loopc).Pos.y)-(2 + CharList(loopc).Pos.X, 0 + CharList(loopc).Pos.y)
            frmMain.picRadar.Line (0 + CharList(loopc).Pos.X, 1 + CharList(loopc).Pos.y)-(2 + CharList(loopc).Pos.X, 1 + CharList(loopc).Pos.y)
        End If
    End If
Next loopc

bRefreshRadar = False
End Sub


''
' Actualiza el Caption del menu principal
'
' @param Trabajando Indica el path del mapa con el que se esta trabajando
' @param Editado Indica si el mapa esta editado

Public Sub CaptionWorldEditor(ByVal Trabajando As String, ByVal Editado As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If Trabajando = vbNullString Then
    Trabajando = "Nuevo Mapa"
End If
frmMain.Caption = "WorldEditor v" & App.Major & "." & App.Minor & " Build " & App.Revision & " - [" & Trabajando & "]"
If Editado = True Then
    frmMain.Caption = frmMain.Caption & " (modificado)"
End If
End Sub

Private Sub LoadClientSetup()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/19/09
'11/19/09: Pato - Is optional show the frmGuildNews form
'**************************************************************
    Dim fHandle As Integer
    
    If FileExist(App.path & "\init\ao.dat", vbArchive) Then
        fHandle = FreeFile
        
        Open App.path & "\init\ao.dat" For Binary Access Read Lock Write As fHandle
            Get fHandle, , ClientSetup
        Close fHandle
    Else
        'Use dynamic by default
        ClientSetup.bDinamic = True
    End If
    
    If InStr(1, ClientSetup.sGraficos, "Graficos") Then
        GraphicsFile = ClientSetup.sGraficos
    Else
        GraphicsFile = "Graficos3.ind"
    End If
End Sub

Public Function fullyBlack(ByVal grhIndex As Long) As Boolean
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 10/27/2011
'Return true if the grh is fully black
'*************************************************************
    Dim color As Long
    Dim X As Long
    Dim y As Long
    Dim srchdc As Long
    Dim Surface As DirectDrawSurface7
    
    With GrhData(GrhData(grhIndex).Frames(1))
        Set Surface = SurfaceDB.Surface(.FileNum)
        
        srchdc = Surface.GetDC
        
        For y = .sY To .sY + .pixelHeight - 1
            For X = .sX To .sX + .pixelWidth - 1
                color = GetPixel(srchdc, X, y)
                
                If color <> 0 Then
                    Call Surface.ReleaseDC(srchdc)
                    
                    fullyBlack = False
                    Exit Function
                End If
            Next X
        Next y
    End With
    
    Call Surface.ReleaseDC(srchdc)
    fullyBlack = True
End Function
