Attribute VB_Name = "modMapIO"
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
' modMapIO
'
' @remarks Funciones Especificas al trabajo con Archivos de Mapas
' @author gshaxor@gmail.com
' @version 0.1.15
' @date 20060602

Option Explicit

Private MapTitulo As String     ' GS > Almacena el titulo del mapa para el .dat

''
' Obtener el tamaño de un archivo
'
' @param FileName Especifica el path del archivo
' @return   Nos devuelve el tamaño

Public Function FileSize(ByRef FileName As String) As Long
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

    On Error GoTo FalloFile
    Dim nFileNum As Integer
    Dim lFileSize As Long
    
    nFileNum = FreeFile
    Open FileName For Input As nFileNum
    lFileSize = LOF(nFileNum)
    Close nFileNum
    FileSize = lFileSize
    
    Exit Function
FalloFile:
    FileSize = -1
End Function

''
' Nos dice si existe el archivo/directorio
'
' @param file Especifica el path
' @param FileType Especifica el tipo de archivo/directorio
' @return   Nos devuelve verdadero o falso

Public Function FileExist(ByRef file As String, ByVal FileType As VbFileAttribute) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 26/05/06
'*************************************************

FileExist = (LenB(Dir$(file, FileType)) > 0)
End Function

''
' Abre un Mapa
'
' @param Path Especifica el path del mapa

Public Sub AbrirMapa(ByRef path As String, ByRef Buffer() As MapBlock, Optional ByVal SoloMap As Boolean = False)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If FileSize(path) = 130273 Then
    Call MapaV1_Cargar(path, Buffer, SoloMap)
    frmMain.mnuUtirialNuevoFormato.Checked = False
Else
    Call MapaV2_Cargar(path, Buffer, SoloMap)
    frmMain.mnuUtirialNuevoFormato.Checked = True
End If
End Sub

''
' Guarda el Mapa
'
' @param Path Especifica el path del mapa

Public Sub GuardarMapa(Optional ByRef path As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************

frmMain.Dialog.CancelError = True
On Error GoTo ErrHandler

If LenB(path) = 0 Then
    frmMain.ObtenerNombreArchivo True
    path = frmMain.Dialog.FileName
    If LenB(path) = 0 Then Exit Sub
End If

If frmMain.mnuUtirialNuevoFormato.Checked = True Then
    Call MapaV2_Guardar(path)
Else
    Call MapaV1_Guardar(path)
End If

ErrHandler:
End Sub

''
' Nos pregunta donde guardar el mapa en caso de modificarlo
'
' @param Path Especifica si existiera un path donde guardar el mapa

Public Sub DeseaGuardarMapa(Optional ByRef path As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If MapInfo.Changed = 1 Then
    If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
        GuardarMapa path
    End If
End If
End Sub

''
' Limpia todo el mapa a uno nuevo
'

Public Sub NuevoMapa()
'*************************************************
'Author: ^[GS]^
'Last modified: 21/05/06
'*************************************************

On Error Resume Next

Dim loopc As Integer
Dim y As Integer
Dim X As Integer

bAutoGuardarMapaCount = 0

frmMain.mnuUtirialNuevoFormato.Checked = True
frmMain.mnuReAbrirMapa.Enabled = False
frmMain.TimAutoGuardarMapa.Enabled = False
frmMain.lblMapVersion.Caption = 0

MapaCargado = False

For loopc = 0 To frmMain.MapPest.Count - 1
    frmMain.MapPest(loopc).Enabled = False
Next

frmMain.MousePointer = 11

ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

For loopc = 1 To LastChar
    If CharList(loopc).Active = 1 Then Call EraseChar(loopc)
Next loopc

MapInfo.MapVersion = 0
MapInfo.name = "Nuevo Mapa"
MapInfo.Music = 0
MapInfo.PK = True
MapInfo.MagiaSinEfecto = 0
MapInfo.Terreno = "BOSQUE"
MapInfo.Zona = "CAMPO"
MapInfo.Restringir = "NO"
MapInfo.NoEncriptarMP = 0

Call MapInfo_Actualizar

bRefreshRadar = True ' Radar

'Set changed flag
MapInfo.Changed = 0
frmMain.MousePointer = 0

' Vacio deshacer
modEdicion.Deshacer_Clear

MapaCargado = True

frmMain.SetFocus

End Sub

''
' Guardar Mapa con el formato V2
'
' @param SaveAs Especifica donde guardar el mapa

Public Sub MapaV2_Guardar(ByVal SaveAs As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error GoTo ErrorSave
Dim FreeFileMap As Long
Dim FreeFileInf As Long
Dim loopc As Long
Dim TempInt As Integer
Dim y As Long
Dim X As Long
Dim ByFlags As Byte

If FileExist(SaveAs, vbNormal) = True Then
    If MsgBox("¿Desea sobrescribir " & SaveAs & "?", vbCritical + vbYesNo) = vbNo Then
        Exit Sub
    Else
        Kill SaveAs
    End If
End If

frmMain.MousePointer = 11

' y borramos el .inf tambien
If FileExist(Left$(SaveAs, Len(SaveAs) - 4) & ".inf", vbNormal) = True Then
    Kill Left$(SaveAs, Len(SaveAs) - 4) & ".inf"
End If

'Open .map file
FreeFileMap = FreeFile
Open SaveAs For Binary As FreeFileMap
Seek FreeFileMap, 1

SaveAs = Left$(SaveAs, Len(SaveAs) - 4)
SaveAs = SaveAs & ".inf"

'Open .inf file
FreeFileInf = FreeFile
Open SaveAs For Binary As FreeFileInf
Seek FreeFileInf, 1

    'map Header
    
    ' Version del Mapa
    If frmMain.lblMapVersion.Caption < 32767 Then
        frmMain.lblMapVersion.Caption = frmMain.lblMapVersion + 1
        frmMapInfo.txtMapVersion = frmMain.lblMapVersion.Caption
    End If
    Put FreeFileMap, , CInt(frmMain.lblMapVersion.Caption)
    Put FreeFileMap, , MiCabecera
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    
    'inf Header
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    
    'Write .map file
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
                ByFlags = 0
                
                If MapData(X, y).Blocked = 1 Then ByFlags = ByFlags Or 1
                If MapData(X, y).Graphic(2).grhIndex Then ByFlags = ByFlags Or 2
                If MapData(X, y).Graphic(3).grhIndex Then ByFlags = ByFlags Or 4
                If MapData(X, y).Graphic(4).grhIndex Then ByFlags = ByFlags Or 8
                If MapData(X, y).Trigger Then ByFlags = ByFlags Or 16
                
                Put FreeFileMap, , ByFlags
                
                Put FreeFileMap, , MapData(X, y).Graphic(1).grhIndex
                
                For loopc = 2 To 4
                    If MapData(X, y).Graphic(loopc).grhIndex Then _
                        Put FreeFileMap, , MapData(X, y).Graphic(loopc).grhIndex
                Next loopc
                
                If MapData(X, y).Trigger Then _
                    Put FreeFileMap, , MapData(X, y).Trigger
                
                '.inf file
                
                ByFlags = 0
                
                If MapData(X, y).TileExit.Map Then ByFlags = ByFlags Or 1
                If MapData(X, y).NPCIndex Then ByFlags = ByFlags Or 2
                If MapData(X, y).OBJInfo.objindex Then ByFlags = ByFlags Or 4
                
                Put FreeFileInf, , ByFlags
                
                If MapData(X, y).TileExit.Map Then
                    Put FreeFileInf, , MapData(X, y).TileExit.Map
                    Put FreeFileInf, , MapData(X, y).TileExit.X
                    Put FreeFileInf, , MapData(X, y).TileExit.y
                End If
                
                If MapData(X, y).NPCIndex Then
                
                    Put FreeFileInf, , CInt(MapData(X, y).NPCIndex)
                End If
                
                If MapData(X, y).OBJInfo.objindex Then
                    Put FreeFileInf, , MapData(X, y).OBJInfo.objindex
                    Put FreeFileInf, , MapData(X, y).OBJInfo.Amount
                End If
            
        Next X
    Next y
    
    'Close .map file
    Close FreeFileMap
    
    'Close .inf file
    Close FreeFileInf


Call Pestañas(SaveAs)

'write .dat file
SaveAs = Left$(SaveAs, Len(SaveAs) - 4) & ".dat"
MapInfo_Guardar SaveAs

'Change mouse icon
frmMain.MousePointer = 0
MapInfo.Changed = 0

Exit Sub

ErrorSave:
    MsgBox "Error en GuardarV2, nro. " & Err.Number & " - " & Err.Description
End Sub

''
' Guardar Mapa con el formato V1
'
' @param SaveAs Especifica donde guardar el mapa

Public Sub MapaV1_Guardar(SaveAs As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error GoTo ErrorSave
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim loopc As Long
    Dim TempInt As Integer
    Dim y As Long
    Dim X As Long
    
    If FileExist(SaveAs, vbNormal) = True Then
        If MsgBox("¿Desea sobrescribir " & SaveAs & "?", vbCritical + vbYesNo) = vbNo Then
            Exit Sub
        Else
            Kill SaveAs
        End If
    End If
    
    'Change mouse icon
    frmMain.MousePointer = 11
    
    If FileExist(Left$(SaveAs, Len(SaveAs) - 4) & ".inf", vbNormal) = True Then
        Kill Left$(SaveAs, Len(SaveAs) - 4) & ".inf"
    End If
    
    'Open .map file
    FreeFileMap = FreeFile
    Open SaveAs For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    
    SaveAs = Left$(SaveAs, Len(SaveAs) - 4)
    SaveAs = SaveAs & ".inf"
    'Open .inf file
    FreeFileInf = FreeFile
    Open SaveAs For Binary As FreeFileInf
    Seek FreeFileInf, 1
    'map Header
    If frmMain.lblMapVersion.Caption < 32767 Then
        frmMain.lblMapVersion.Caption = frmMain.lblMapVersion + 1
        frmMapInfo.txtMapVersion = frmMain.lblMapVersion.Caption
    End If
    Put FreeFileMap, , CInt(frmMain.lblMapVersion.Caption)
    Put FreeFileMap, , MiCabecera
    
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    
    'inf Header
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    
    'Write .map file
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            '.map file
            
            ' Bloqueos
            Put FreeFileMap, , MapData(X, y).Blocked
            
            ' Capas
            For loopc = 1 To 4
                If loopc = 2 Then Call FixCoasts(MapData(X, y).Graphic(loopc).grhIndex, X, y)
                Put FreeFileMap, , MapData(X, y).Graphic(loopc).grhIndex
            Next loopc
            
            ' Triggers
            Put FreeFileMap, , MapData(X, y).Trigger
            Put FreeFileMap, , TempInt
            
            '.inf file
            'Tile exit
            Put FreeFileInf, , MapData(X, y).TileExit.Map
            Put FreeFileInf, , MapData(X, y).TileExit.X
            Put FreeFileInf, , MapData(X, y).TileExit.y
            
            'NPC
            Put FreeFileInf, , MapData(X, y).NPCIndex
            
            'Object
            Put FreeFileInf, , MapData(X, y).OBJInfo.objindex
            Put FreeFileInf, , MapData(X, y).OBJInfo.Amount
            
            'Empty place holders for future expansion
            Put FreeFileInf, , TempInt
            Put FreeFileInf, , TempInt
            
        Next X
    Next y
    
    'Close .map file
    Close FreeFileMap
    'Close .inf file
    Close FreeFileInf
    
    Call Pestañas(SaveAs)
    
    'write .dat file
    SaveAs = Left$(SaveAs, Len(SaveAs) - 4) & ".dat"
    MapInfo_Guardar SaveAs
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0
    
Exit Sub
ErrorSave:
    MsgBox "Error " & Err.Number & " - " & Err.Description
End Sub

''
' Abrir Mapa con el formato V2
'
' @param Map Especifica el Path del mapa

Public Sub MapaV2_Cargar(ByVal Map As String, ByRef Buffer() As MapBlock, ByVal SoloMap As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error Resume Next
    Dim loopc As Integer
    Dim TempInt As Integer
    Dim Body As Integer
    Dim Head As Integer
    Dim Heading As Byte
    Dim y As Integer
    Dim X As Integer
    Dim ByFlags As Byte
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    
    'Change mouse icon
    frmMain.MousePointer = 11
       
    'Open files
    FreeFileMap = FreeFile
    Open Map For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    If Not SoloMap Then
        Map = Left$(Map, Len(Map) - 4)
        Map = Map & ".inf"
        
        FreeFileInf = FreeFile
        Open Map For Binary As FreeFileInf
        Seek FreeFileInf, 1
    End If
    
    'Cabecera map
    Get FreeFileMap, , MapInfo.MapVersion
    Get FreeFileMap, , MiCabecera
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    
    If Not SoloMap Then
        'Cabecera inf
        Get FreeFileInf, , TempInt
        Get FreeFileInf, , TempInt
        Get FreeFileInf, , TempInt
        Get FreeFileInf, , TempInt
        Get FreeFileInf, , TempInt
    End If
    
    'Load arrays
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            Get FreeFileMap, , ByFlags
            
            Buffer(X, y).Blocked = (ByFlags And 1)
            
            Get FreeFileMap, , Buffer(X, y).Graphic(1).grhIndex
            InitGrh Buffer(X, y).Graphic(1), Buffer(X, y).Graphic(1).grhIndex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get FreeFileMap, , Buffer(X, y).Graphic(2).grhIndex
                InitGrh Buffer(X, y).Graphic(2), Buffer(X, y).Graphic(2).grhIndex
            Else
                Buffer(X, y).Graphic(2).grhIndex = 0
            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get FreeFileMap, , Buffer(X, y).Graphic(3).grhIndex
                InitGrh Buffer(X, y).Graphic(3), Buffer(X, y).Graphic(3).grhIndex
            Else
                Buffer(X, y).Graphic(3).grhIndex = 0
            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get FreeFileMap, , Buffer(X, y).Graphic(4).grhIndex
                InitGrh Buffer(X, y).Graphic(4), Buffer(X, y).Graphic(4).grhIndex
            Else
                Buffer(X, y).Graphic(4).grhIndex = 0
            End If
            
             
            'Trigger used?
            If ByFlags And 16 Then
                Get FreeFileMap, , Buffer(X, y).Trigger
            Else
                Buffer(X, y).Trigger = 0
            End If
            
            If Not SoloMap Then
                '.inf file
                Get FreeFileInf, , ByFlags
                
                If ByFlags And 1 Then
                    Get FreeFileInf, , Buffer(X, y).TileExit.Map
                    Get FreeFileInf, , Buffer(X, y).TileExit.X
                    Get FreeFileInf, , Buffer(X, y).TileExit.y
                End If
        
                If ByFlags And 2 Then
                    'Get and make NPC
                    Get FreeFileInf, , Buffer(X, y).NPCIndex
        
                    If Buffer(X, y).NPCIndex < 0 Then
                        Buffer(X, y).NPCIndex = 0
                    Else
                        Body = NpcData(Buffer(X, y).NPCIndex).Body
                        Head = NpcData(Buffer(X, y).NPCIndex).Head
                        Heading = NpcData(Buffer(X, y).NPCIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, X, y)
                    End If
                End If
        
                If ByFlags And 4 Then
                    'Get and make Object
                    Get FreeFileInf, , Buffer(X, y).OBJInfo.objindex
                    Get FreeFileInf, , Buffer(X, y).OBJInfo.Amount
                    If Buffer(X, y).OBJInfo.objindex > 0 Then
                        InitGrh Buffer(X, y).ObjGrh, ObjData(Buffer(X, y).OBJInfo.objindex).grhIndex
                    End If
                End If
            End If
        Next X
    Next y
    
    'Close files
    Close FreeFileMap
    
    If Not SoloMap Then
        Close FreeFileInf
        
        Call Pestañas(Map)
        
        bRefreshRadar = True ' Radar
        
        Map = Left$(Map, Len(Map) - 4) & ".dat"
        
        MapInfo_Cargar Map
        frmMain.lblMapVersion.Caption = MapInfo.MapVersion
        
        'Set changed flag
        MapInfo.Changed = 0
        
        ' Vacia el Deshacer
        modEdicion.Deshacer_Clear
    End If
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapaCargado = True
End Sub

''
' Abrir Mapa con el formato V1
'
' @param Map Especifica el Path del mapa

Public Sub MapaV1_Cargar(ByVal Map As String, ByRef Buffer() As MapBlock, ByVal SoloMap As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    On Error Resume Next
    Dim TBlock As Byte
    Dim loopc As Integer
    Dim TempInt As Integer
    Dim Body As Integer
    Dim Head As Integer
    Dim Heading As Byte
    Dim y As Integer
    Dim X As Integer
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    
    'Change mouse icon
    frmMain.MousePointer = 11
    
    'Open files
    FreeFileMap = FreeFile
    Open Map For Binary As FreeFileMap
    
    
    Seek FreeFileMap, 1
    
    If Not SoloMap Then
        Map = Left$(Map, Len(Map) - 4)
        Map = Map & ".inf"
        FreeFileInf = FreeFile
        Open Map For Binary As #2
        Seek FreeFileInf, 1
    End If
    
    'Cabecera map
    Get FreeFileMap, , MapInfo.MapVersion
    Get FreeFileMap, , MiCabecera
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    
    If Not SoloMap Then
        'Cabecera inf
        Get FreeFileInf, , TempInt
        Get FreeFileInf, , TempInt
        Get FreeFileInf, , TempInt
        Get FreeFileInf, , TempInt
        Get FreeFileInf, , TempInt
    End If
    
    'Load arrays
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            '.map file
            Get FreeFileMap, , Buffer(X, y).Blocked
            
            For loopc = 1 To 4
                Get FreeFileMap, , Buffer(X, y).Graphic(loopc).grhIndex
                'Set up GRH
                If Buffer(X, y).Graphic(loopc).grhIndex > 0 Then
                    InitGrh Buffer(X, y).Graphic(loopc), Buffer(X, y).Graphic(loopc).grhIndex
                End If
            Next loopc
            'Trigger
            Get FreeFileMap, , Buffer(X, y).Trigger
            
            Get FreeFileMap, , TempInt
            
            If Not SoloMap Then
                '.inf file
                
                'Tile exit
                Get FreeFileInf, , Buffer(X, y).TileExit.Map
                Get FreeFileInf, , Buffer(X, y).TileExit.X
                Get FreeFileInf, , Buffer(X, y).TileExit.y
                              
                'make NPC
                Get FreeFileInf, , Buffer(X, y).NPCIndex
                If Buffer(X, y).NPCIndex > 0 Then
                    Body = NpcData(Buffer(X, y).NPCIndex).Body
                    Head = NpcData(Buffer(X, y).NPCIndex).Head
                    Heading = NpcData(Buffer(X, y).NPCIndex).Heading
                    Call MakeChar(NextOpenChar(), Body, Head, Heading, X, y)
                End If
                
                'Make obj
                Get FreeFileInf, , Buffer(X, y).OBJInfo.objindex
                Get FreeFileInf, , Buffer(X, y).OBJInfo.Amount
                If Buffer(X, y).OBJInfo.objindex > 0 Then
                    InitGrh Buffer(X, y).ObjGrh, ObjData(Buffer(X, y).OBJInfo.objindex).grhIndex
                End If
                
                'Empty place holders for future expansion
                Get FreeFileInf, , TempInt
                Get FreeFileInf, , TempInt
            End If
        Next X
    Next y
    
    'Close files
    Close FreeFileMap
    
    If Not SoloMap Then
        Close FreeFileInf
        
        Call Pestañas(Map)
        
        bRefreshRadar = True ' Radar
        
        Map = Left$(Map, Len(Map) - 4) & ".dat"
            
        MapInfo_Cargar Map
        frmMain.lblMapVersion.Caption = MapInfo.MapVersion
        
        'Set changed flag
        MapInfo.Changed = 0
        
        ' Vacia el Deshacer
        modEdicion.Deshacer_Clear
    End If
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapaCargado = True
End Sub



' *****************************************************************************
' MAPINFO *********************************************************************
' *****************************************************************************

''
' Guardar Informacion del Mapa (.dat)
'
' @param Archivo Especifica el Path del archivo .DAT

Public Sub MapInfo_Guardar(ByVal Archivo As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************

    If LenB(MapTitulo) = 0 Then
        MapTitulo = NameMap_Save
    End If

    Call WriteVar(Archivo, MapTitulo, "Name", MapInfo.name)
    Call WriteVar(Archivo, MapTitulo, "MusicNum", MapInfo.Music)
    Call WriteVar(Archivo, MapTitulo, "MagiaSinefecto", Val(MapInfo.MagiaSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "NoEncriptarMP", Val(MapInfo.NoEncriptarMP))

    Call WriteVar(Archivo, MapTitulo, "Terreno", MapInfo.Terreno)
    Call WriteVar(Archivo, MapTitulo, "Zona", MapInfo.Zona)
    Call WriteVar(Archivo, MapTitulo, "Restringir", MapInfo.Restringir)
    Call WriteVar(Archivo, MapTitulo, "BackUp", str(MapInfo.BackUp))

    If MapInfo.PK Then
        Call WriteVar(Archivo, MapTitulo, "Pk", "0")
    Else
        Call WriteVar(Archivo, MapTitulo, "Pk", "1")
    End If
End Sub

''
' Abrir Informacion del Mapa (.dat)
'
' @param Archivo Especifica el Path del archivo .DAT

Public Sub MapInfo_Cargar(ByVal Archivo As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 02/06/06
'*************************************************

On Error Resume Next
    Dim Leer As New clsIniReader
    Dim loopc As Integer
    Dim path As String
    
    MapTitulo = Empty
    Leer.Initialize Archivo

    For loopc = Len(Archivo) To 1 Step -1
        If mid$(Archivo, loopc, 1) = "\" Then
            path = Left$(Archivo, loopc)
            Exit For
        End If
    Next loopc
    
    Archivo = Right$(Archivo, Len(Archivo) - (Len(path)))
    MapTitulo = UCase$(Left$(Archivo, Len(Archivo) - 4))

    MapInfo.name = Leer.GetValue(MapTitulo, "Name")
    MapInfo.Music = Leer.GetValue(MapTitulo, "MusicNum")
    MapInfo.MagiaSinEfecto = Val(Leer.GetValue(MapTitulo, "MagiaSinEfecto"))
    MapInfo.NoEncriptarMP = Val(Leer.GetValue(MapTitulo, "NoEncriptarMP"))
    
    If Val(Leer.GetValue(MapTitulo, "Pk")) = 0 Then
        MapInfo.PK = True
    Else
        MapInfo.PK = False
    End If
    
    MapInfo.Terreno = Leer.GetValue(MapTitulo, "Terreno")
    MapInfo.Zona = Leer.GetValue(MapTitulo, "Zona")
    MapInfo.Restringir = Leer.GetValue(MapTitulo, "Restringir")
    MapInfo.BackUp = Val(Leer.GetValue(MapTitulo, "BACKUP"))
    
    Call MapInfo_Actualizar
    
End Sub

''
' Actualiza el formulario de MapInfo
'

Public Sub MapInfo_Actualizar()
'*************************************************
'Author: ^[GS]^
'Last modified: 02/06/06
'*************************************************

On Error Resume Next
    ' Mostrar en Formularios
    frmMapInfo.txtMapNombre.Text = MapInfo.name
    frmMapInfo.txtMapMusica.Text = MapInfo.Music
    frmMapInfo.txtMapTerreno.Text = MapInfo.Terreno
    frmMapInfo.txtMapZona.Text = MapInfo.Zona
    frmMapInfo.txtMapRestringir.Text = MapInfo.Restringir
    frmMapInfo.chkMapBackup.Value = MapInfo.BackUp
    frmMapInfo.chkMapMagiaSinEfecto.Value = MapInfo.MagiaSinEfecto
    frmMapInfo.chkMapNoEncriptarMP.Value = MapInfo.NoEncriptarMP
    frmMapInfo.chkMapPK.Value = IIf(MapInfo.PK = True, 1, 0)
    frmMapInfo.txtMapVersion = MapInfo.MapVersion
    frmMain.lblMapNombre = MapInfo.name
    frmMain.lblMapMusica = MapInfo.Music

End Sub

''
' Calcula la orden de Pestañas
'
' @param Map Especifica path del mapa

Public Sub Pestañas(ByVal Map As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
On Error Resume Next
Dim loopc As Integer

PATH_Save = Left$(Map, InStrRev(Map, "\"))
Map = Right$(Map, Len(Map) - Len(PATH_Save))
Map = Left$(Map, Len(Map) - 4) 'Sacamos la extension

For loopc = Len(Map) To 1 Step -1
    If Not IsNumeric(mid$(Map, loopc)) Then
        NumMap_Save = Val(mid$(Map, loopc + 1))
        NameMap_Save = Left$(Map, loopc)
        Exit For
    End If
Next loopc

For loopc = (NumMap_Save - 4) To (NumMap_Save + 8)
    If FileExist(PATH_Save & NameMap_Save & loopc & ".map", vbArchive) Then
        frmMain.MapPest(loopc - NumMap_Save + 4).Visible = True
        frmMain.MapPest(loopc - NumMap_Save + 4).Enabled = True
        frmMain.MapPest(loopc - NumMap_Save + 4).Caption = NameMap_Save & loopc
    Else
        frmMain.MapPest(loopc - NumMap_Save + 4).Visible = False
    End If
Next loopc
End Sub
