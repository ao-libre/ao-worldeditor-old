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

Public Function FileSize(ByVal FileName As String) As Long
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************

    On Error GoTo FalloFile

    Dim nFileNum  As Integer

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

Public Function General_File_Exist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean

    '*************************************************
    'Author: Unkwown
    'Last modified: 26/05/06
    '*************************************************
    If LenB(Dir(file, FileType)) = 0 Then
        General_File_Exist = False
    Else
        General_File_Exist = True

    End If

End Function

''
' Abre un Mapa
'
' @param Path Especifica el path del mapa

Public Sub AbrirMapa(ByVal Path As String, ByRef Buffer() As MapBlock, Optional ByVal SoloMap As Boolean = False)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************

    Call MapaV2_Cargar(Path, Buffer, SoloMap)

End Sub

''
' Guarda el Mapa
'
' @param Path Especifica el path del mapa

Public Sub GuardarMapa(Optional Path As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************

    frmMain.Dialog.CancelError = True

    On Error GoTo ErrHandler

    If LenB(Path) = 0 Then
        frmMain.ObtenerNombreArchivo True
        Path = frmMain.Dialog.FileName

        If LenB(Path) = 0 Then Exit Sub

    End If

    Call MapaV2_Guardar(Path)

ErrHandler:

End Sub

''
' Nos pregunta donde guardar el mapa en caso de modificarlo
'
' @param Path Especifica si existiera un path donde guardar el mapa

Public Sub DeseaGuardarMapa(Optional Path As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            GuardarMapa Path

        End If

    End If

End Sub

Public Sub NuevoMapa()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 21/05/06
    '*************************************************

    On Error Resume Next

    Dim loopc As Integer

    bAutoGuardarMapaCount = 0
    
    With frmMain
    
        '.mnuUtirialNuevoFormato.Checked = True
        .mnuReAbrirMapa.Enabled = False
        .TimAutoGuardarMapa.Enabled = False
        .lblMapVersion.Caption = 0

        MapaCargado = False

        For loopc = 0 To frmMain.MapPest.Count - 1
            .MapPest(loopc).Enabled = False
        Next

        .MousePointer = 11
    
    End With
    
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

    For loopc = 1 To LastChar
        If CharList(loopc).Active = 1 Then Call EraseChar(loopc)
    Next loopc
    
    With MapInfo
    
        .MapVersion = 0
        .name = "Nuevo Mapa"
        .Music = 0
        .PK = True
        .MagiaSinEfecto = 0
        .Terreno = "BOSQUE"
        .Zona = "CAMPO"
        .Restringir = "NO"
        .NoEncriptarMP = 0
    
    End With

    Call MapInfo_Actualizar

    bRefreshRadar = True ' Radar

    'Set changed flag
    MapInfo.Changed = 0
    frmMain.MousePointer = 0

    ' Vacio deshacer
    Call modEdicion.Deshacer_Clear

    MapaCargado = True

    frmMain.SetFocus

End Sub

Public Sub MapaV2_Guardar(ByVal SaveAs As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error GoTo ErrorSave

    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim loopc       As Long
    Dim TempInt     As Integer
    Dim Y           As Long
    Dim X           As Long
    Dim ByFlags     As Byte
    Dim R           As Byte
    Dim G           As Byte
    Dim B           As Byte

    If General_File_Exist(SaveAs, vbNormal) = True Then
        If MsgBox("¿Desea sobrescribir " & SaveAs & "?", vbCritical + vbYesNo) = vbNo Then
            Exit Sub
        Else
            Kill SaveAs

        End If

    End If

    frmMain.MousePointer = 11

    ' y borramos el .inf tambien
    If General_File_Exist(left$(SaveAs, Len(SaveAs) - 4) & ".inf", vbNormal) = True Then
        Kill left$(SaveAs, Len(SaveAs) - 4) & ".inf"

    End If

    'Open .map file
    FreeFileMap = FreeFile
    Open SaveAs For Binary As FreeFileMap
    Seek FreeFileMap, 1

    SaveAs = left$(SaveAs, Len(SaveAs) - 4)
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
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            ByFlags = 0
                
            If MapData(X, Y).Blocked = 1 Then ByFlags = ByFlags Or 1
            If MapData(X, Y).Graphic(2).grh_index Then ByFlags = ByFlags Or 2
            If MapData(X, Y).Graphic(3).grh_index Then ByFlags = ByFlags Or 4
            If MapData(X, Y).Graphic(4).grh_index Then ByFlags = ByFlags Or 8
            If MapData(X, Y).Trigger Then ByFlags = ByFlags Or 16
            If MapData(X, Y).particle_group_index Then ByFlags = ByFlags Or 32
            If MapData(X, Y).light_index Then ByFlags = ByFlags Or 64
            If MapData(X, Y).AlturaPoligonos(0) Or MapData(X, Y).AlturaPoligonos(1) Or MapData(X, Y).AlturaPoligonos(2) Or MapData(X, Y).AlturaPoligonos(3) Then ByFlags = ByFlags Or 128
            Put FreeFileMap, , ByFlags
                
            Put FreeFileMap, , MapData(X, Y).Graphic(1).grh_index
                
            For loopc = 2 To 4

                If MapData(X, Y).Graphic(loopc).grh_index Then Put FreeFileMap, , MapData(X, Y).Graphic(loopc).grh_index
            Next loopc
                
            If MapData(X, Y).Trigger Then Put FreeFileMap, , MapData(X, Y).Trigger
                
            If MapData(X, Y).particle_group_index Then Put FreeFileMap, , MapData(X, Y).parti_index

            If MapData(X, Y).light_index Then
                Put FreeFileMap, , Lights(MapData(X, Y).light_index).Range
                R = Lights(MapData(X, Y).light_index).RGBCOLOR.R
                G = Lights(MapData(X, Y).light_index).RGBCOLOR.G
                B = Lights(MapData(X, Y).light_index).RGBCOLOR.B
                Put FreeFileMap, , R
                Put FreeFileMap, , G
                Put FreeFileMap, , B

            End If
                
            If MapData(X, Y).AlturaPoligonos(0) Or MapData(X, Y).AlturaPoligonos(1) Or MapData(X, Y).AlturaPoligonos(2) Or MapData(X, Y).AlturaPoligonos(3) Then
                Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(0)
                Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(1)
                Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(2)
                Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(3)
                    
                If MapData(X, Y).AlturaPoligonos(0) Then Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(0)

                If MapData(X, Y).AlturaPoligonos(1) Then Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(1)

                If MapData(X, Y).AlturaPoligonos(2) Then Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(2)

                If MapData(X, Y).AlturaPoligonos(3) Then Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(3)

            End If
                
            '.inf file
                
            ByFlags = 0
                
            If MapData(X, Y).TileExit.Map Then ByFlags = ByFlags Or 1
            If MapData(X, Y).NPCIndex Then ByFlags = ByFlags Or 2
            If MapData(X, Y).OBJInfo.objindex Then ByFlags = ByFlags Or 4
                
            Put FreeFileInf, , ByFlags
                
            If MapData(X, Y).TileExit.Map Then
                Put FreeFileInf, , MapData(X, Y).TileExit.Map
                Put FreeFileInf, , MapData(X, Y).TileExit.X
                Put FreeFileInf, , MapData(X, Y).TileExit.Y

            End If
                
            If MapData(X, Y).NPCIndex Then
                
                Put FreeFileInf, , CInt(MapData(X, Y).NPCIndex)

            End If
                
            If MapData(X, Y).OBJInfo.objindex Then
                Put FreeFileInf, , MapData(X, Y).OBJInfo.objindex
                Put FreeFileInf, , MapData(X, Y).OBJInfo.Amount

            End If
            
        Next X
    Next Y
    
    'Close .map file
    Close FreeFileMap
    
    'Close .inf file
    Close FreeFileInf

    Call Pestañas(SaveAs)

    'write .dat file
    SaveAs = left$(SaveAs, Len(SaveAs) - 4) & ".dat"
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

    Dim loopc       As Long

    Dim TempInt     As Integer

    Dim T           As String

    Dim Y           As Long

    Dim X           As Long
    
    If General_File_Exist(SaveAs, vbNormal) = True Then
        If MsgBox("¿Desea sobrescribir " & SaveAs & "?", vbCritical + vbYesNo) = vbNo Then
            Exit Sub
        Else
            Kill SaveAs

        End If

    End If
    
    'Change mouse icon
    frmMain.MousePointer = 11
    T = SaveAs

    If General_File_Exist(left(SaveAs, Len(SaveAs) - 4) & ".inf", vbNormal) = True Then
        Kill left(SaveAs, Len(SaveAs) - 4) & ".inf"

    End If
    
    'Open .map file
    FreeFileMap = FreeFile
    Open SaveAs For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    SaveAs = left(SaveAs, Len(SaveAs) - 4)
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
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            '.map file
            
            ' Bloqueos
            Put FreeFileMap, , MapData(X, Y).Blocked
            
            ' Capas
            For loopc = 1 To 4

                If loopc = 2 Then Call FixCoasts(MapData(X, Y).Graphic(loopc).grh_index, X, Y)
                Put FreeFileMap, , MapData(X, Y).Graphic(loopc).grh_index
            Next loopc
            
            ' Triggers
            Put FreeFileMap, , MapData(X, Y).Trigger
            Put FreeFileMap, , TempInt
            
            '.inf file
            'Tile exit
            Put FreeFileInf, , MapData(X, Y).TileExit.Map
            Put FreeFileInf, , MapData(X, Y).TileExit.X
            Put FreeFileInf, , MapData(X, Y).TileExit.Y
            
            'NPC
            Put FreeFileInf, , MapData(X, Y).NPCIndex
            
            'Object
            Put FreeFileInf, , MapData(X, Y).OBJInfo.objindex
            Put FreeFileInf, , MapData(X, Y).OBJInfo.Amount
            
            'Empty place holders for future expansion
            Put FreeFileInf, , TempInt
            Put FreeFileInf, , TempInt
            
        Next X
    Next Y
    
    'Close .map file
    Close FreeFileMap
    'Close .inf file
    Close FreeFileInf
    FreeFileMap = FreeFile
    Open T & "2" For Binary Access Write As FreeFileMap
    Put FreeFileMap, , MapData
    Close FreeFileMap
    Call Pestañas(SaveAs)
    
    'write .dat file
    SaveAs = left(SaveAs, Len(SaveAs) - 4) & ".dat"
    MapInfo_Guardar SaveAs
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0
    
    Exit Sub
ErrorSave:
    MsgBox "Error " & Err.Number & " - " & Err.Description

End Sub

Public Sub MapaV2_Cargar(ByVal Map As String, ByRef Buffer() As MapBlock, ByVal SoloMap As Boolean)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error Resume Next

    Dim TempInt     As Integer
    Dim Body        As Integer
    Dim Head        As Integer
    Dim Heading     As Byte
    Dim Y           As Integer
    Dim X           As Integer
    Dim ByFlags     As Byte
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim TempLng     As Long
    Dim TempByte1   As Byte
    Dim TempByte2   As Byte
    Dim TempByte3   As Byte
           
    Call LightDestroyAll
    Call Particle_Group_Remove_All
    Call Map_ResetMontañita
    
    'Change mouse icon
    frmMain.MousePointer = 11
       
    'Open files
    FreeFileMap = FreeFile
    Open Map For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    If Not SoloMap Then
        Map = left$(Map, Len(Map) - 4)
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
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            Get FreeFileMap, , ByFlags
            
            Buffer(X, Y).Blocked = (ByFlags And 1)
            
            Get FreeFileMap, , Buffer(X, Y).Graphic(1).grh_index
            Grh_Initialize Buffer(X, Y).Graphic(1), Buffer(X, Y).Graphic(1).grh_index
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get FreeFileMap, , Buffer(X, Y).Graphic(2).grh_index
                Grh_Initialize Buffer(X, Y).Graphic(2), Buffer(X, Y).Graphic(2).grh_index
            Else
                Buffer(X, Y).Graphic(2).grh_index = 0

            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get FreeFileMap, , Buffer(X, Y).Graphic(3).grh_index
                Grh_Initialize Buffer(X, Y).Graphic(3), Buffer(X, Y).Graphic(3).grh_index
            Else
                Buffer(X, Y).Graphic(3).grh_index = 0

            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get FreeFileMap, , Buffer(X, Y).Graphic(4).grh_index
                Grh_Initialize Buffer(X, Y).Graphic(4), Buffer(X, Y).Graphic(4).grh_index
            Else
                Buffer(X, Y).Graphic(4).grh_index = 0

            End If
             
            'Trigger used?
            If ByFlags And 16 Then
                Get FreeFileMap, , Buffer(X, Y).Trigger
            Else
                Buffer(X, Y).Trigger = 0

            End If
            
            If ByFlags And 32 Then
                Get FreeFileMap, , TempInt
                MapData(X, Y).particle_group_index = General_Particle_Create(TempInt, X, Y, -1)

            End If
            
            If ByFlags And 64 Then
                Get FreeFileMap, , TempLng
                Get FreeFileMap, , TempByte1
                Get FreeFileMap, , TempByte2
                Get FreeFileMap, , TempByte3
                Call LightSet(X, Y, True, TempLng, TempByte1, TempByte2, TempByte3)

            End If
            
            If ByFlags And 128 Then
                Get FreeFileMap, , MapData(X, Y).AlturaPoligonos(0)
                Get FreeFileMap, , MapData(X, Y).AlturaPoligonos(1)
                Get FreeFileMap, , MapData(X, Y).AlturaPoligonos(2)
                Get FreeFileMap, , MapData(X, Y).AlturaPoligonos(3)
                
                If MapData(X, Y).AlturaPoligonos(0) Then Get FreeFileMap, , MapData(X, Y).AlturaPoligonos(0)
                
                If MapData(X, Y).AlturaPoligonos(1) Then Get FreeFileMap, , MapData(X, Y).AlturaPoligonos(1)
                
                If MapData(X, Y).AlturaPoligonos(2) Then Get FreeFileMap, , MapData(X, Y).AlturaPoligonos(2)
                
                If MapData(X, Y).AlturaPoligonos(3) Then Get FreeFileMap, , MapData(X, Y).AlturaPoligonos(3)

            End If

            If Not SoloMap Then
                '.inf file
                Get FreeFileInf, , ByFlags
                
                If ByFlags And 1 Then
                    Get FreeFileInf, , Buffer(X, Y).TileExit.Map
                    Get FreeFileInf, , Buffer(X, Y).TileExit.X
                    Get FreeFileInf, , Buffer(X, Y).TileExit.Y

                End If
        
                If ByFlags And 2 Then
                    'Get and make NPC
                    Get FreeFileInf, , Buffer(X, Y).NPCIndex
        
                    If Buffer(X, Y).NPCIndex < 0 Then
                        Buffer(X, Y).NPCIndex = 0
                    Else
                        Body = NpcData(Buffer(X, Y).NPCIndex).Body
                        Head = NpcData(Buffer(X, Y).NPCIndex).Head
                        Heading = NpcData(Buffer(X, Y).NPCIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, X, Y)

                    End If

                End If
        
                If ByFlags And 4 Then
                    'Get and make Object
                    Get FreeFileInf, , Buffer(X, Y).OBJInfo.objindex
                    Get FreeFileInf, , Buffer(X, Y).OBJInfo.Amount

                    If Buffer(X, Y).OBJInfo.objindex > 0 Then
                        Grh_Initialize Buffer(X, Y).ObjGrh, ObjData(Buffer(X, Y).OBJInfo.objindex).grh_index

                    End If

                End If

            End If

        Next X
    Next Y
    
    'Close files
    Close FreeFileMap
    
    If Not SoloMap Then
        Close FreeFileInf
        
        Call Pestañas(Map)
        
        bRefreshRadar = True ' Radar
        
        Map = left$(Map, Len(Map) - 4) & ".dat"
        
        Call MapInfo_Cargar(Map)
        frmMain.lblMapVersion.Caption = MapInfo.MapVersion
        
        'Set changed flag
        MapInfo.Changed = 0
        
        ' Vacia el Deshacer
        Call modEdicion.Deshacer_Clear

    End If
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapaCargado = True

End Sub

''
' Abrir Mapa con el formato V1
'
' @param Map Especifica el Path del mapa

Public Sub MapaV1_Cargar(ByVal Map As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error Resume Next

    Dim loopc       As Integer
    Dim TempInt     As Integer
    Dim Body        As Integer
    Dim Head        As Integer
    Dim Heading     As Byte
    Dim Y           As Integer
    Dim X           As Integer
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long

    DoEvents
    
    'Change mouse icon
    frmMain.MousePointer = 11
    
    'Open files
    FreeFileMap = FreeFile
    Open Map For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    Map = left(Map, Len(Map) - 4)
    Map = Map & ".inf"
    FreeFileInf = FreeFile
    Open Map For Binary As #2
    Seek FreeFileInf, 1
    
    'Cabecera map
    Get FreeFileMap, , MapInfo.MapVersion
    Get FreeFileMap, , MiCabecera
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    
    'Cabecera inf
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    
    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            '.map file
            Get FreeFileMap, , MapData(X, Y).Blocked
            
            For loopc = 1 To 4
                Get FreeFileMap, , MapData(X, Y).Graphic(loopc).grh_index

                'Set up GRH
                If MapData(X, Y).Graphic(loopc).grh_index > 0 Then
                    Grh_Initialize MapData(X, Y).Graphic(loopc), MapData(X, Y).Graphic(loopc).grh_index

                End If

            Next loopc

            'Trigger
            Get FreeFileMap, , MapData(X, Y).Trigger
            
            Get FreeFileMap, , TempInt
            '.inf file
            
            'Tile exit
            Get FreeFileInf, , MapData(X, Y).TileExit.Map
            Get FreeFileInf, , MapData(X, Y).TileExit.X
            Get FreeFileInf, , MapData(X, Y).TileExit.Y
                          
            'make NPC
            Get FreeFileInf, , MapData(X, Y).NPCIndex

            If MapData(X, Y).NPCIndex > 0 Then
                Body = NpcData(MapData(X, Y).NPCIndex).Body
                Head = NpcData(MapData(X, Y).NPCIndex).Head
                Heading = NpcData(MapData(X, Y).NPCIndex).Heading
                Call MakeChar(NextOpenChar(), Body, Head, Heading, X, Y)

            End If
            
            'Make obj
            Get FreeFileInf, , MapData(X, Y).OBJInfo.objindex
            Get FreeFileInf, , MapData(X, Y).OBJInfo.Amount

            If MapData(X, Y).OBJInfo.objindex > 0 Then
                Grh_Initialize MapData(X, Y).ObjGrh, ObjData(MapData(X, Y).OBJInfo.objindex).grh_index

            End If
            
            'Empty place holders for future expansion
            Get FreeFileInf, , TempInt
            Get FreeFileInf, , TempInt
                 
        Next X
    Next Y
    
    'Close files
    Close FreeFileMap
    Close FreeFileInf
     
    Call Pestañas(Map)
    
    Map = left(Map, Len(Map) - 4) & ".dat"
        
    MapInfo_Cargar Map
    frmMain.lblMapVersion.Caption = MapInfo.MapVersion
    
    'Set changed flag
    MapInfo.Changed = 0
    
    ' Vacia el Deshacer
    modEdicion.Deshacer_Clear
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapaCargado = True

End Sub

Public Sub MapaV3_Cargar(ByVal Map As String)
    '*************************************************
    'Author: Loopzer
    'Last modified: 22/11/07
    '*************************************************

    On Error Resume Next

    Dim FreeFileMap As Long

    DoEvents
    'Change mouse icon
    frmMain.MousePointer = 11
    
    FreeFileMap = FreeFile
    Open Map For Binary Access Read As FreeFileMap
    Get FreeFileMap, , MapData
    Close FreeFileMap
    
    Call Pestañas(Map)
    
    Map = left(Map, Len(Map) - 4) & ".dat"
        
    MapInfo_Cargar Map
    frmMain.lblMapVersion.Caption = MapInfo.MapVersion
    
    'Set changed flag
    MapInfo.Changed = 0
    
    ' Vacia el Deshacer
    modEdicion.Deshacer_Clear
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapaCargado = True

End Sub

Public Sub MapaV3_Guardar(Mapa As String)

    '*************************************************
    'Author: Loopzer
    'Last modified: 22/11/07
    '*************************************************
    'copy&paste RLZ
    On Error GoTo ErrorSave

    Dim FreeFileMap As Long
    
    If General_File_Exist(Mapa, vbNormal) = True Then
        If MsgBox("¿Desea sobrescribir " & Mapa & "?", vbCritical + vbYesNo) = vbNo Then
            Exit Sub
        Else
            Kill Mapa

        End If

    End If
    
    frmMain.MousePointer = 11
    
    FreeFileMap = FreeFile
    Open Mapa For Binary Access Write As FreeFileMap
    Put FreeFileMap, , MapData
    Close FreeFileMap
    Call Pestañas(Mapa)
    
    Mapa = left(Mapa, Len(Mapa) - 4) & ".dat"
    MapInfo_Guardar Mapa
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0
    
    Exit Sub
ErrorSave:
    MsgBox "Error " & Err.Number & " - " & Err.Description

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
    Call WriteVar(Archivo, MapTitulo, "InviSinEfecto", Val(MapInfo.InviSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "ResuSinEfecto", Val(MapInfo.ResuSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "NoEncriptarMP", Val(MapInfo.NoEncriptarMP))

    Call WriteVar(Archivo, MapTitulo, "Terreno", MapInfo.Terreno)
    Call WriteVar(Archivo, MapTitulo, "Zona", MapInfo.Zona)
    Call WriteVar(Archivo, MapTitulo, "Restringir", MapInfo.Restringir)
    Call WriteVar(Archivo, MapTitulo, "BackUp", Str(MapInfo.BackUp))

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

    Dim Leer  As New clsIniReader
    Dim loopc As Integer
    Dim Path  As String

    MapTitulo = Empty
    
    Call Leer.Initialize(Archivo)

    For loopc = Len(Archivo) To 1 Step -1

        If mid(Archivo, loopc, 1) = "\" Then
            Path = left(Archivo, loopc)
            Exit For

        End If

    Next
    Archivo = Right(Archivo, Len(Archivo) - (Len(Path)))
    MapTitulo = UCase(left(Archivo, Len(Archivo) - 4))
    
    With MapInfo
    
        .name = Leer.GetValue(MapTitulo, "Name")
        .Music = Leer.GetValue(MapTitulo, "MusicNum")
        .MagiaSinEfecto = Val(Leer.GetValue(MapTitulo, "MagiaSinEfecto"))
        .InviSinEfecto = Val(Leer.GetValue(MapTitulo, "InviSinEfecto"))
        .ResuSinEfecto = Val(Leer.GetValue(MapTitulo, "ResuSinEfecto"))
        .NoEncriptarMP = Val(Leer.GetValue(MapTitulo, "NoEncriptarMP"))
    
        If Val(Leer.GetValue(MapTitulo, "Pk")) = 0 Then
            MapInfo.PK = True
        Else
            MapInfo.PK = False

        End If
    
        .Terreno = Leer.GetValue(MapTitulo, "Terreno")
        .Zona = Leer.GetValue(MapTitulo, "Zona")
        .Restringir = Leer.GetValue(MapTitulo, "Restringir")
        .BackUp = Val(Leer.GetValue(MapTitulo, "BACKUP"))
    
    End With
    
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
    With frmMapInfo
        .txtMapNombre.Text = MapInfo.name
        .txtMapMusica.Text = MapInfo.Music
        .txtMapTerreno.Text = MapInfo.Terreno
        .txtMapZona.Text = MapInfo.Zona
        .txtMapRestringir.Text = MapInfo.Restringir
        .chkMapBackup.value = MapInfo.BackUp
        .chkMapMagiaSinEfecto.value = MapInfo.MagiaSinEfecto
        .chkMapInviSinEfecto.value = MapInfo.InviSinEfecto
        .chkMapResuSinEfecto.value = MapInfo.ResuSinEfecto
        .chkMapNoEncriptarMP.value = MapInfo.NoEncriptarMP
        .chkMapPK.value = IIf(MapInfo.PK = True, 1, 0)
        .txtMapVersion = MapInfo.MapVersion
    End With
    
    With frmMain
        .lblMapNombre = MapInfo.name
        .lblMapMusica = MapInfo.Music
    End With
    

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

    For loopc = Len(Map) To 1 Step -1

        If mid(Map, loopc, 1) = "\" Then
            PATH_Save = left(Map, loopc)
            Exit For

        End If

    Next
    Map = Right(Map, Len(Map) - (Len(PATH_Save)))

    For loopc = Len(left(Map, Len(Map) - 4)) To 1 Step -1

        If IsNumeric(mid(left(Map, Len(Map) - 4), loopc, 1)) = False Then
            NumMap_Save = Right(left(Map, Len(Map) - 4), Len(left(Map, Len(Map) - 4)) - loopc)
            NameMap_Save = left(Map, loopc)
            Exit For

        End If

    Next

    For loopc = (NumMap_Save - 4) To (NumMap_Save + 8)

        If General_File_Exist(PATH_Save & NameMap_Save & loopc & ".map", vbArchive) = True Then
            frmMain.MapPest(loopc - NumMap_Save + 4).Visible = True
            frmMain.MapPest(loopc - NumMap_Save + 4).Enabled = True
            frmMain.MapPest(loopc - NumMap_Save + 4).Caption = NameMap_Save & loopc
        Else
            frmMain.MapPest(loopc - NumMap_Save + 4).Visible = False

        End If

    Next

End Sub

