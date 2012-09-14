Attribute VB_Name = "modDeclaraciones"
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
' modDeclaraciones
'
' @remarks Declaraciones
' @author ^[GS]^
' @version 0.1.10
' @date 20061002

Option Explicit

Public Const MSGMod As String = "Este mapa há sido modificado." & vbCrLf & "Si no lo guardas perderas todos los cambios ¿Deseas guardarlo?"
Public Const MSGDang As String = "CUIDADO! Este comando puede arruinar el mapa." & vbCrLf & "¿Estas seguro que desea continuar?"

Public Const ENDL As String * 2 = vbCrLf

' Objeto de Translado
Public Cfg_TrOBJ As Integer

'Path
Public IniPath As String
Public DirGraficos As String
Public DirMidi As String
Public DirMp3 As String
Public DirIndex As String
Public DirDats As String

Public bAutoGuardarMapa As Byte
Public bAutoGuardarMapaCount As Byte
Public HotKeysAllow As Boolean  ' Control Automatico de HotKeys
Public vMostrando As Byte
Public WORK As Boolean
Public PATH_Save As String
Public NumMap_Save As Integer
Public NameMap_Save As String

Public ClienteHeight As Integer
Public ClienteWidth As Integer

' DX Config
Public PantallaX As Integer
Public PantallaY As Integer

Public Type tSetupMods
    bDinamic    As Boolean
    byMemory    As Byte
    bUseVideo   As Boolean
    bNoMusic    As Boolean
    bNoSound    As Boolean
    bNoRes      As Boolean ' 24/06/2006 - ^[GS]^
    bNoSoundEffects As Boolean
    sGraficos   As String * 13
    bGuildNews  As Boolean ' 11/19/09
    bDie        As Boolean ' 11/23/09 - FragShooter
    bKill       As Boolean ' 11/23/09 - FragShooter
    byMurderedLevel As Byte ' 11/23/09 - FragShooter
    bActive     As Boolean
    bGldMsgConsole As Boolean
    bCantMsgs   As Byte
End Type

Public ClientSetup As tSetupMods

Public SobreX As Byte   ' Posicion X bajo el Cursor
Public SobreY As Byte   ' Posicion Y bajo el Cursor

' Radar
Public MiRadarX As Integer
Public MiRadarY As Integer
Public bRefreshRadar As Boolean

Public Type SupData
    name As String
    Grh As Integer
    Width As Byte
    Height As Byte
    block As Boolean
    Capa As Byte
End Type

Public MaxSup As Integer
Public SupData() As SupData

Public Type NpcData
    name As String
    Body As Integer
    Head As Integer
    Heading As Byte
    Hostile As Boolean
End Type

Public NumNPCs As Integer
Public NpcData() As NpcData

Public Type ObjData
    name As String 'Nombre del obj
    ObjType As Integer 'Tipo enum que determina cuales son las caract del obj
    grhIndex As Integer ' Indice del grafico que representa el obj
    GrhSecundario As Integer
    Info As String
    Ropaje As Integer 'Indice del grafico del ropaje
    WeaponAnim As Integer ' Apunta a una anim de armas
    ShieldAnim As Integer ' Apunta a una anim de escudo
    Texto As String
End Type

Public NumOBJs As Integer
Public ObjData() As ObjData

Public Conexion As New Connection
Public prgRun As Boolean
Public CurrentGrh() As Grh
Public MapaCargado As Boolean
Public dTiempoGT As Double
Public dLastWalk As Double
Public DespX As Integer
Public DespY As Integer
Public mAncho As Byte
Public MAlto As Byte

Public MinSelectX As Byte
Public MaxSelectX As Byte
Public MinSelectY As Byte
Public MaxSelectY As Byte
Public FirstSelectX As Byte
Public FirstSelectY As Byte
Public MouseDownX As Integer
Public MouseDownY As Integer

'Hold info about each map
Public Type MapInfo
    Music As String
    name As String
    MapVersion As Integer
    PK As Boolean
    MagiaSinEfecto As Byte
    NoEncriptarMP As Byte
    Terreno As String
    Zona As String
    Restringir As String
    BackUp As Byte
    Changed As Byte ' flag for WorldEditor
End Type

'********** CONSTANTS ***********

'Map sizes in tiles
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

''
'Sets a Grh animation to loop indefinitely.
Public Const INFINITE_LOOPS As Integer = -1

'********** TYPES ***********
'Holds a local position
Public Type Position
    X As Integer
    y As Integer
End Type

'Holds a world position
Public Type WorldPos
    Map As Integer
    X As Integer
    y As Integer
End Type

'Points to a grhData and keeps animation info
Public Type Grh
    grhIndex As Integer
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer
End Type

'Holds data about where a bmp can be found,
'How big it is and animation info
Public Type GrhData
    sX As Integer
    sY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    Speed As Single
End Type

' Cuerpos body.dat
Public Type tIndiceCuerpo
    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

' Lista de Cuerpos body.dat
Public Type tBodyData
    Walk(1 To 4) As Grh
    HeadOffset As Position
End Type

' body.dat
Public BodyData() As tBodyData
Public NumBodies As Integer

'Lista de cabezas
Public Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type

'Heads list
Public Type tHeadData
    Head(1 To 4) As Grh
End Type

Public HeadData() As tHeadData

'Apariencia del personaje
Public Type Char
    Active As Byte
    Heading As E_Heading
    Pos As Position
    
    iHead As Integer
    iBody As Integer
    Body As tBodyData
    Head As tHeadData
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
End Type

'Holds info about a object
Public Type Obj
    objindex As Integer
    Amount As Integer
End Type

'Holds info about each tile position
Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    Trigger As Integer
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    Select As Byte
End Type

'********** Public VARS ***********
'Where the map borders are.. Set during load
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'Object Constants
Public Const MAX_INVENORY_OBJS  As Integer = 10000

' Deshacer
Public Const maxDeshacer As Integer = 10
Public MapData_Deshacer(1 To maxDeshacer, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

Type tDeshacerInfo
    Libre As Boolean
    Desc As String
End Type

Public MapData_Deshacer_Info(1 To maxDeshacer) As tDeshacerInfo

'********** Public ARRAYS ***********
Public GrhData() As GrhData 'Holds all the grh data
Public MapData() As MapBlock 'Holds map data for current map
Public MapInfo As MapInfo 'Holds map info for current map
Public CharList(1 To 10000) As Char 'Holds info about all characters on map

'Encabezado bmp
Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

' DirectSound
Public Const NumSoundBuffers As Byte = 20
Public DirectSound As DirectSound
Public DSBuffers(1 To NumSoundBuffers) As DirectSoundBuffer
Public LastSoundBufferUsed As Integer

Public gDespX As Integer
Public gDespY As Integer

'User status vars
Public CurMap As Integer 'Current map loaded
Public UserIndex As Integer
Global UserBody As Integer
Global UserHead As Integer
Public UserPos As Position 'Holds current user pos
Public AddtoUserPos As Position 'For moving user
Public UserCharIndex As Integer

Public FramesPerSec As Integer
Public FramesPerSecCounter As Long

'Main view size size in tiles
Public WindowTileWidth As Integer
Public WindowTileHeight As Integer

Public HalfWindowTileWidth As Integer
Public HalfWindowTileHeight As Integer

'Pixel offset of main view screen from 0,0
Public MainViewTop As Integer
Public MainViewLeft As Integer

'How many tiles the engine "looks ahead" when
'drawing the screen
Public TileBufferSize As Integer

Public TileBufferPixelOffsetX As Integer
Public TileBufferPixelOffsetY As Integer

Public timerElapsedTime As Single
Public timerTicksPerFrame As Single
Public engineBaseSpeed As Single

Public FPS As Long
Public fpsLastCheck As Long

Public UserMoving As Byte

'Handle to where all the drawing is going to take place
Public DisplayFormhWnd As Long

'Tile size in pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

'Map editor variables
Public WalkMode As Boolean

'Totals
Public NumMaps As Integer 'Number of maps
Public Numheads As Integer
Public NumGrhFiles As Integer 'Number of bmps
Global NumChars As Integer
Global LastChar As Integer

'Direcciones
Public Enum E_Heading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

Public CurLayer As Byte
Public GraphicsFile As String 'Que graficos.ind usamos
Public bTriggers As Boolean
Public bBloqs As Boolean
Public bTranslados As Boolean
Public bVerCapa(2 To 4) As Boolean
Public bAutoCompletarSuperficies As Boolean
Public bVerNpcs As Boolean
Public bVerObjetos As Boolean
Public bSelectSup As Boolean
Public MosaicoChecked As Boolean

'********** Direct X ***********
Public MainViewRect As RECT
Public MainDestRect As RECT
Public MainViewWidth As Integer
Public MainViewHeight As Integer
Public BackBufferRect As RECT

Public DirectX As New DirectX7
Public DirectDraw As DirectDraw7

Public PrimarySurface As DirectDrawSurface7
Public PrimaryClipper As DirectDrawClipper
Public SecundaryClipper As DirectDrawClipper
Public BackBufferSurface As DirectDrawSurface7
Public SurfaceDB As clsSurfaceManager

Public TextDrawer As clsTextDrawer
'********** OUTSIDE FUNCTIONS ***********

'For Get and Write Var
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long

'For KeyInput
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long) As Long
