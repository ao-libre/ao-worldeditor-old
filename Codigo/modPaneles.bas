Attribute VB_Name = "modPaneles"
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
' modPaneles
'
' @remarks Funciones referentes a los Paneles de Funcion
' @author gshaxor@gmail.com
' @version 0.3.28
' @date 20060530

Option Explicit

''
' Activa/Desactiva el Estado de la Funcion en el Panel Superior
'
' @param Numero Especifica en numero de funcion
' @param Activado Especifica si esta o no activado

Public Sub EstSelectPanel(ByVal Numero As Byte, ByVal Activado As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/05/06
'*************************************************

    If Activado Then
        frmMain.SelectPanel(Numero).GradientMode = lv_Bottom2Top
        frmMain.SelectPanel(Numero).HoverBackColor = frmMain.SelectPanel(Numero).GradientColor
        If frmMain.mnuVerAutomatico.Checked = True Then
            Select Case Numero
                Case 0
                    If CurLayer <> 1 Then
                        frmMain.mnuVerCapa(CurLayer).Tag = CInt(frmMain.mnuVerCapa(CurLayer).Checked)
                        frmMain.mnuVerCapa(CurLayer).Checked = True
                            
                        bVerCapa(CurLayer) = True
                    End If
                Case 2
                    frmMain.cVerBloqueos.Tag = CInt(frmMain.cVerBloqueos.value)
                    frmMain.cVerBloqueos.value = True
                    frmMain.mnuVerBloqueos.Checked = frmMain.cVerBloqueos.value
                Case 6
                    frmMain.cVerTriggers.Tag = CInt(frmMain.cVerTriggers.value)
                    frmMain.cVerTriggers.value = True
                    frmMain.mnuVerTriggers.Checked = frmMain.cVerTriggers.value
            End Select
        End If
    Else
        frmMain.SelectPanel(Numero).HoverBackColor = frmMain.SelectPanel(Numero).BackColor
        frmMain.SelectPanel(Numero).GradientMode = lv_NoGradient
        
        If frmMain.mnuVerAutomatico.Checked Then
            Select Case Numero
                Case 0
                    If CurLayer <> 1 Then
                        If LenB(frmMain.mnuVerCapa(CurLayer).Tag) <> 0 Then
                            frmMain.mnuVerCapa(CurLayer).Checked = CBool(frmMain.mnuVerCapa(CurLayer).Tag)
                            bVerCapa(CurLayer) = frmMain.mnuVerCapa(CurLayer).Checked
                        End If
                    End If
                Case 2
                    If LenB(frmMain.cVerBloqueos.Tag) = 0 Then frmMain.cVerBloqueos.Tag = 0
                    frmMain.cVerBloqueos.value = CBool(frmMain.cVerBloqueos.Tag)
                    frmMain.mnuVerBloqueos.Checked = frmMain.cVerBloqueos.value
                Case 6
                    If LenB(frmMain.cVerTriggers.Tag) = 0 Then frmMain.cVerTriggers.Tag = 0
                    frmMain.cVerTriggers.value = CBool(frmMain.cVerTriggers.Tag)
                    frmMain.mnuVerTriggers.Checked = frmMain.cVerTriggers.value
            End Select
        End If
    End If
End Sub

''
' Muestra los controles que componen a la funcion seleccionada del Panel
'
' @param Numero Especifica el numero de Funcion
' @param Ver Especifica si se va a ver o no
' @param Normal Inidica que ahi que volver todo No visible

Public Sub VerFuncion(ByVal Numero As Byte, ByVal Ver As Boolean, Optional Normal As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    If Normal Then Call VerFuncion(vMostrando, False, False)
    
    Select Case Numero
        Case 0 ' Superficies
            frmMain.lListado(0).Visible = Ver
            frmMain.cFiltro(0).Visible = Ver
            frmMain.cCapas.Visible = Ver
            frmMain.cGrh.Visible = Ver
            frmMain.cQuitarEnEstaCapa.Visible = Ver
            frmMain.cQuitarEnTodasLasCapas.Visible = Ver
            frmMain.cSeleccionarSuperficie.Visible = Ver
            frmMain.lbFiltrar(0).Visible = Ver
            frmMain.lbCapas.Visible = Ver
            frmMain.lbGrh.Visible = Ver
            frmMain.PreviewGrh.Visible = Ver
            If Ver = True Then
                frmMain.StatTxt.Top = 672
                frmMain.StatTxt.Height = 37
            Else
                frmMain.StatTxt.Top = 416
                frmMain.StatTxt.Height = 293
            End If
        Case 1 ' Translados
            frmMain.lMapN.Visible = Ver
            frmMain.lXhor.Visible = Ver
            frmMain.lYver.Visible = Ver
            frmMain.tTMapa.Visible = Ver
            frmMain.tTX.Visible = Ver
            frmMain.tTY.Visible = Ver
            frmMain.cInsertarTrans.Visible = Ver
            frmMain.cInsertarTransOBJ.Visible = Ver
            frmMain.cUnionManual.Visible = Ver
            frmMain.cUnionAuto.Visible = Ver
            frmMain.cQuitarTrans.Visible = Ver
        Case 2 ' Bloqueos
            frmMain.cQuitarBloqueo.Visible = Ver
            frmMain.cInsertarBloqueo.Visible = Ver
            frmMain.cVerBloqueos.Visible = Ver
        Case 3  ' NPCs
            frmMain.lListado(1).Visible = Ver
            frmMain.cFiltro(1).Visible = Ver
            frmMain.lbFiltrar(1).Visible = Ver
            frmMain.lNumFunc(Numero - 3).Visible = Ver
            frmMain.cNumFunc(Numero - 3).Visible = Ver
            frmMain.cInsertarFunc(Numero - 3).Visible = Ver
            frmMain.cQuitarFunc(Numero - 3).Visible = Ver
            frmMain.cAgregarFuncalAzar(Numero - 3).Visible = Ver
            frmMain.lCantFunc(Numero - 3).Visible = Ver
            frmMain.cCantFunc(Numero - 3).Visible = Ver
        Case 4 ' NPCs Hostiles
            frmMain.lListado(2).Visible = Ver
            frmMain.cFiltro(2).Visible = Ver
            frmMain.lbFiltrar(2).Visible = Ver
            frmMain.lNumFunc(Numero - 3).Visible = Ver
            frmMain.cNumFunc(Numero - 3).Visible = Ver
            frmMain.cInsertarFunc(Numero - 3).Visible = Ver
            frmMain.cQuitarFunc(Numero - 3).Visible = Ver
            frmMain.cAgregarFuncalAzar(Numero - 3).Visible = Ver
            frmMain.lCantFunc(Numero - 3).Visible = Ver
            frmMain.cCantFunc(Numero - 3).Visible = Ver
        Case 5 ' OBJs
            frmMain.lListado(3).Visible = Ver
            frmMain.cFiltro(3).Visible = Ver
            frmMain.lbFiltrar(3).Visible = Ver
            frmMain.lNumFunc(Numero - 3).Visible = Ver
            frmMain.cNumFunc(Numero - 3).Visible = Ver
            frmMain.cInsertarFunc(Numero - 3).Visible = Ver
            frmMain.cQuitarFunc(Numero - 3).Visible = Ver
            frmMain.cAgregarFuncalAzar(Numero - 3).Visible = Ver
            frmMain.lCantFunc(Numero - 3).Visible = Ver
            frmMain.cCantFunc(Numero - 3).Visible = Ver
        Case 6 ' Triggers
            frmMain.cQuitarTrigger.Visible = Ver
            frmMain.cInsertarTrigger.Visible = Ver
            frmMain.cVerTriggers.Visible = Ver
            frmMain.lListado(4).Visible = Ver
    End Select
    
    If Ver Then
        vMostrando = Numero
        If Numero < 0 Or Numero > 6 Then Exit Sub
        If frmMain.SelectPanel(Numero).value = False Then
            frmMain.SelectPanel(Numero).value = True
        End If
    Else
        If Numero < 0 Or Numero > 6 Then Exit Sub
        If frmMain.SelectPanel(Numero).value = True Then
            frmMain.SelectPanel(Numero).value = False
        End If
    End If
End Sub

''
' Filtra del Listado de Elementos de una Funcion
'
' @param Numero Indica la funcion a Filtrar

Public Sub Filtrar(ByVal Numero As Byte)
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************

    Dim vDatos As String
    Dim i As Long
    Dim Filtro As String
    
    If frmMain.cFiltro(Numero).ListCount > 5 Then
        frmMain.cFiltro(Numero).RemoveItem 0
    End If
    
    frmMain.cFiltro(Numero).AddItem frmMain.cFiltro(Numero).Text
    frmMain.lListado(Numero).Clear
        
    Filtro = frmMain.cFiltro(Numero).Text
    
    Select Case Numero
        Case 0 ' superficie
            For i = 0 To MaxSup
                vDatos = SupData(i).name
                
                If (LenB(Filtro) = 0) Or (InStr(1, UCase$(vDatos), UCase$(Filtro))) Then
                    frmMain.lListado(Numero).AddItem vDatos & " - #" & i
                End If
            Next i
            
        Case 1 ' NPCs
            For i = 1 To NumNPCs
                If Not NpcData(i).Hostile Then
                    vDatos = NpcData(i).name
                    
                    If (LenB(Filtro) = 0) Or (InStr(1, UCase$(vDatos), UCase$(Filtro))) Then
                        frmMain.lListado(Numero).AddItem vDatos & " - #" & i
                    End If
                End If
            Next i
        Case 2 ' NPCs Hostiles
            For i = 1 To NumNPCs
                If NpcData(i).Hostile Then
                    vDatos = NpcData(i).name
                    
                    If (LenB(Filtro) = 0) Or (InStr(1, UCase$(vDatos), UCase$(Filtro))) Then
                        frmMain.lListado(Numero).AddItem vDatos & " - #" & i
                    End If
                End If
            Next i
            
        Case 3 ' Objetos
            For i = 1 To NumOBJs
                vDatos = ObjData(i).name
                
                If (LenB(Filtro) = 0) Or (InStr(1, UCase$(vDatos), UCase$(Filtro))) Then
                    frmMain.lListado(Numero).AddItem vDatos & " - #" & i
                End If
            Next i
    End Select
End Sub

Public Function DameGrhIndex(ByVal GrhIn As Integer) As Integer
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

DameGrhIndex = SupData(GrhIn).Grh

If SupData(GrhIn).Width > 0 Then
    frmConfigSup.MOSAICO.value = vbChecked
    frmConfigSup.MAncho.Text = SupData(GrhIn).Width
    frmConfigSup.mLargo.Text = SupData(GrhIn).Height
Else
    frmConfigSup.MOSAICO.value = vbUnchecked
    frmConfigSup.MAncho.Text = "0"
    frmConfigSup.mLargo.Text = "0"
End If

End Function

Public Sub ActualizarMosaico()
If frmConfigSup.MOSAICO.value = vbChecked Then
    MAncho = Val(frmConfigSup.MAncho)
    MAlto = Val(frmConfigSup.mLargo)
    
    ReDim CurrentGrh(1 To MAncho, 1 To MAlto) As Grh
Else
    ReDim CurrentGrh(0) As Grh
End If

Call fPreviewGrh(frmMain.cGrh.Text)
End Sub

Public Sub fPreviewGrh(ByVal GrhIn As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 22/05/06
'*************************************************
Dim X As Byte
Dim Y As Byte

If Val(GrhIn) < 1 Then
    frmMain.cGrh.Text = UBound(GrhData)
    Exit Sub
End If

If Val(GrhIn) > UBound(GrhData) Then
    frmMain.cGrh.Text = 1
    Exit Sub
End If

If frmConfigSup.MOSAICO.value = vbChecked Then
    For Y = 1 To MAlto
        For X = 1 To MAncho
            'Change CurrentGrh
            InitGrh CurrentGrh(X, Y), GrhIn
            
            GrhIn = GrhIn + 1
        Next X
    Next Y
Else
    InitGrh CurrentGrh(0), GrhIn
End If
End Sub

''
' Indica la accion de mostrar Vista Previa de la Superficie seleccionada
'

Public Sub VistaPreviaDeSup()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
Dim SR As RECT, DR As RECT
    
    frmGrafico.ShowPic = frmGrafico.Picture1
    
    If frmConfigSup.MOSAICO = vbUnchecked Then
        If CurrentGrh(0).GrhIndex Then
            With GrhData(CurrentGrh(0).GrhIndex)
                DR.Left = 0
                DR.Top = 0
                DR.Bottom = .pixelHeight
                DR.Right = .pixelWidth
                
                SR.Left = .sX
                SR.Top = .sY
                SR.Right = SR.Left + .pixelWidth
                SR.Bottom = SR.Top + .pixelHeight
                
                Call DrawGrhtoHdc(frmGrafico.ShowPic.hdc, .Frames(1), SR, DR)
            End With
        End If
    Else
        Dim X As Integer, Y As Integer
        
        For X = 1 To MAncho
            For Y = 1 To MAlto
                If CurrentGrh(X, Y).GrhIndex Then
                    With GrhData(CurrentGrh(X, Y).GrhIndex)
                        DR.Left = (X - 1) * .pixelWidth
                        DR.Top = (Y - 1) * .pixelHeight
                        DR.Right = X * .pixelWidth
                        DR.Bottom = Y * .pixelHeight
                        
                        SR.Left = .sX
                        SR.Top = .sY
                        SR.Right = SR.Left + .pixelWidth
                        SR.Bottom = SR.Top + .pixelHeight
                        
                        Call DrawGrhtoHdc(frmGrafico.ShowPic.hdc, .Frames(1), SR, DR)
                    End With
                End If
            Next Y
        Next X
    End If
    
    frmGrafico.ShowPic.Picture = frmGrafico.ShowPic.Image
    frmMain.PreviewGrh = frmGrafico.ShowPic
End Sub
