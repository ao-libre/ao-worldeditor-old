Attribute VB_Name = "modIndices"
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
' modIndices
'
' @remarks Funciones Especificas al Trabajo con Indices
' @author gshaxor@gmail.com
' @version 0.1.05
' @date 20060530

Option Explicit


''
' Carga los indices de Graficos
'

Public Sub CargarIndicesDeGraficos()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************

On Error GoTo ErrorHandler

    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim handle As Integer
    Dim fileVersion As Long
    
    If Not FileExist(DirIndex & GraphicsFile, vbArchive) Then
        MsgBox "Falta el archivo " & GraphicsFile & " en " & DirIndex, vbCritical
        End
    End If
    
    'Open files
    handle = FreeFile()
    
    Open DirIndex & GraphicsFile For Binary Access Read As handle
    Seek handle, 1
    
    'Get file version
    Get handle, , fileVersion
    
    'Get number of grhs
    Get handle, , grhCount
    
    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData
    
    While Not EOF(handle)
        Get handle, , Grh
        
        If Grh Then
            With GrhData(Grh)
                'Get number of frames
                Get handle, , .NumFrames
                If .NumFrames <= 0 Then GoTo ErrorHandler
                
                ReDim .Frames(1 To GrhData(Grh).NumFrames)
                
                If .NumFrames > 1 Then
                    'Read a animation GRH set
                    For Frame = 1 To .NumFrames
                        Get handle, , .Frames(Frame)
                        If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                            GoTo ErrorHandler
                        End If
                    Next Frame
                    
                    Get handle, , .Speed
                    
                    If .Speed <= 0 Then GoTo ErrorHandler
                    
                    'Compute width and height
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    If .TileWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileHeight = GrhData(.Frames(1)).TileHeight
                    If .TileHeight <= 0 Then GoTo ErrorHandler
                Else
                    'Read in normal GRH data
                    Get handle, , .FileNum
                    If .FileNum <= 0 Then GoTo ErrorHandler
                    
                    Get handle, , GrhData(Grh).sX
                    If .sX < 0 Then GoTo ErrorHandler
                    
                    Get handle, , .sY
                    If .sY < 0 Then GoTo ErrorHandler
                    
                    Get handle, , .pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    Get handle, , .pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    'Compute width and height
                    .TileWidth = .pixelWidth / TilePixelHeight
                    .TileHeight = .pixelHeight / TilePixelWidth
                    
                    .Frames(1) = Grh
                End If
            End With
        End If
    Wend
    
    Close handle
    
    Exit Sub

ErrorHandler:
Close handle
    MsgBox "Error al intentar cargar el Grh " & Grh & " de graficos.ind en " & DirIndex & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Superficie
'

Public Sub CargarIndicesSuperficie()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************

On Error GoTo Fallo
    If Not FileExist(IniPath & "GrhIndex\indices.ini", vbArchive) Then
        MsgBox "Falta el archivo 'GrhIndex\indices.ini'", vbCritical
        End
    End If
    
    Dim Leer As New clsIniReader
    Dim i As Integer
    
    Leer.Initialize IniPath & "GrhIndex\indices.ini"
    
    MaxSup = Leer.GetValue("INIT", "Referencias")
    
    ReDim SupData(MaxSup) As SupData
    frmMain.lListado(0).Clear
    
    For i = 0 To MaxSup
        SupData(i).name = Leer.GetValue("REFERENCIA" & i, "Nombre")
        SupData(i).Grh = Val(Leer.GetValue("REFERENCIA" & i, "GrhIndice"))
        SupData(i).Width = Val(Leer.GetValue("REFERENCIA" & i, "Ancho"))
        SupData(i).Height = Val(Leer.GetValue("REFERENCIA" & i, "Alto"))
        SupData(i).Block = IIf(Val(Leer.GetValue("REFERENCIA" & i, "Bloquear")) = 1, True, False)
        SupData(i).Capa = Val(Leer.GetValue("REFERENCIA" & i, "Capa"))
        frmMain.lListado(0).AddItem SupData(i).name & " - #" & i
    Next i
    
    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el indice " & i & " de GrhIndex\indices.ini" & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly
End Sub

''
' Carga los indices de Objetos
'

Public Sub CargarIndicesOBJ()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error GoTo Fallo
    If Not FileExist(DirDats & "\OBJ.dat", vbArchive) Then
        MsgBox "Falta el archivo 'OBJ.dat' en " & DirDats, vbCritical
        End
    End If
    
    Dim Obj As Integer
    Dim Leer As New clsIniReader
    
    Call Leer.Initialize(DirDats & "\OBJ.dat")
    
    frmMain.lListado(3).Clear
    NumOBJs = Val(Leer.GetValue("INIT", "NumOBJs"))
    
    ReDim ObjData(1 To NumOBJs) As ObjData
    
    For Obj = 1 To NumOBJs
        frmCargando.X.Caption = "Cargando Datos de Objetos..." & Obj & "/" & NumOBJs
        DoEvents
        ObjData(Obj).name = Leer.GetValue("OBJ" & Obj, "Name")
        ObjData(Obj).GrhIndex = Val(Leer.GetValue("OBJ" & Obj, "GrhIndex"))
        ObjData(Obj).ObjType = Val(Leer.GetValue("OBJ" & Obj, "ObjType"))
        ObjData(Obj).Ropaje = Val(Leer.GetValue("OBJ" & Obj, "NumRopaje"))
        ObjData(Obj).Info = Leer.GetValue("OBJ" & Obj, "Info")
        ObjData(Obj).WeaponAnim = Val(Leer.GetValue("OBJ" & Obj, "Anim"))
        ObjData(Obj).Texto = Leer.GetValue("OBJ" & Obj, "Texto")
        ObjData(Obj).GrhSecundario = Val(Leer.GetValue("OBJ" & Obj, "GrhSec"))
        frmMain.lListado(3).AddItem ObjData(Obj).name & " - #" & Obj
    Next Obj
    Exit Sub
Fallo:
MsgBox "Error al intentar cargar el Objteto " & Obj & " de OBJ.dat en " & DirDats & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Triggers
'

Public Sub CargarIndicesTriggers()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************

On Error GoTo Fallo
    If Not FileExist(DirIndex & "Triggers.ini", vbArchive) Then
        MsgBox "Falta el archivo 'Triggers.ini' en " & DirIndex, vbCritical
        End
    End If
    
    Dim NumT As Integer
    Dim T As Integer
    Dim Leer As New clsIniReader
    
    Call Leer.Initialize(DirIndex & "Triggers.ini")
    
    frmMain.lListado(4).Clear
    NumT = Val(Leer.GetValue("INIT", "NumTriggers"))
    
    For T = 0 To NumT
         frmMain.lListado(4).AddItem Leer.GetValue("Trig" & T, "Name") & " - #" & T
    Next T

Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el Trigger " & T & " de Triggers.ini en " & DirIndex & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Cuerpos
'

Public Sub CargarIndicesDeCuerpos()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************

On Error GoTo Fallo
    If Not FileExist(DirIndex & "Personajes.ind", vbArchive) Then
        MsgBox "Falta el archivo 'Personajes.ind' en " & DirIndex, vbCritical
        End
    End If
    
    Dim n As Integer
    Dim i As Integer
    
    n = FreeFile
    Open DirIndex & "Personajes.ind" For Binary Access Read As #n
        'cabecera
        Get #n, , MiCabecera
        'num de cabezas
        Get #n, , NumBodies
        
        'Resize array
        ReDim BodyData(1 To NumBodies) As tBodyData
        ReDim MisCuerpos(1 To NumBodies) As tIndiceCuerpo
        
        For i = 1 To NumBodies
            Get #n, , MisCuerpos(i)
            
            InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
            InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
            InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
            InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
            
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        Next i
    Close #n
Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el Cuerpo " & i & " de Personajes.ind en " & DirIndex & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Cabezas
'

Public Sub CargarIndicesDeCabezas()
On Error GoTo Fallo
    If Not FileExist(DirIndex & "Cabezas.ind", vbArchive) Then
        MsgBox "Falta el archivo 'Cabezas.ind' en " & DirIndex, vbCritical
        End
    End If
    
    Dim n As Integer
    Dim i As Integer
    Dim MisCabezas() As tIndiceCabeza
    
    n = FreeFile
    
    Open DirIndex & "Cabezas.ind" For Binary Access Read As #n
        'cabecera
        Get #n, , MiCabecera
        'num de cabezas
        Get #n, , Numheads
        'Resize array
        ReDim HeadData(0 To Numheads) As tHeadData
        ReDim MisCabezas(0 To Numheads) As tIndiceCabeza
        
        For i = 1 To Numheads
            Get #n, , MisCabezas(i)
            If MisCabezas(i).Head(1) Then
                Call InitGrh(HeadData(i).Head(1), MisCabezas(i).Head(1), 0)
                Call InitGrh(HeadData(i).Head(2), MisCabezas(i).Head(2), 0)
                Call InitGrh(HeadData(i).Head(3), MisCabezas(i).Head(3), 0)
                Call InitGrh(HeadData(i).Head(4), MisCabezas(i).Head(4), 0)
            End If
        Next i
    Close #n
Exit Sub
Fallo:
    MsgBox "Error al intentar cargar la Cabeza " & i & " de Cabezas.ind en " & DirIndex & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de NPCs
'

Public Sub CargarIndicesNPC()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************

On Error GoTo Fallo
    If Not FileExist(DirDats & "\NPCs.dat", vbArchive) Then
        MsgBox "Falta el archivo 'NPCs.dat' en " & DirDats, vbCritical
        End
    End If
    
    Dim NPC As Integer
    Dim Leer As New clsIniReader
    
    frmMain.lListado(1).Clear
    frmMain.lListado(2).Clear
    
    Call Leer.Initialize(DirDats & "\NPCs.dat")
    NumNPCs = Val(Leer.GetValue("INIT", "NumNPCs"))
    
    ReDim NpcData(1 To NumNPCs) As NpcData
    
    For NPC = 1 To NumNPCs
        With NpcData(NPC)
            .name = Leer.GetValue("NPC" & NPC, "Name")
            .Body = Val(Leer.GetValue("NPC" & NPC, "Body"))
            .Head = Val(Leer.GetValue("NPC" & NPC, "Head"))
            .Heading = Val(Leer.GetValue("NPC" & NPC, "Heading"))
            .Hostile = CBool(Val(Leer.GetValue("NPC" & NPC, "Hostile")))
            
            If .Hostile Then
                frmMain.lListado(2).AddItem .name & " - #" & NPC
            Else
                frmMain.lListado(1).AddItem .name & " - #" & NPC
            End If
        End With
    Next NPC
    
    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el NPC " & NPC & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub
