Attribute VB_Name = "modDXEngine"
Option Explicit

'Private Const DegreeToRadian As Single = 0.0174532925

'***************************
'Estructures
'***************************
'This structure describes a transformed and lit vertex.
Private Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    Color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type

Private Type tGraphicChar
    Src_X As Integer
    Src_Y As Integer
End Type

Private Type tGraphicFont
    texture_index As Long
    Caracteres(0 To 255) As tGraphicChar 'Ascii Chars
    Char_Size As Byte 'In pixels
End Type

Private Type DXFont
    dFont As D3DXFont
    Size As Integer
End Type


Public Enum FontAlignment
    fa_center = DT_CENTER
    fa_top = DT_TOP
    fa_left = DT_LEFT
    fa_topleft = DT_TOP Or DT_LEFT
    fa_bottomleft = DT_BOTTOM Or DT_LEFT
    fa_bottom = DT_BOTTOM
    fa_right = DT_RIGHT
    fa_bottomright = DT_BOTTOM Or DT_RIGHT
    fa_topright = DT_TOP Or DT_RIGHT
End Enum

'***************************
'Variables
'***************************
'Major DX Objects
Public dX As DirectX8
Public d3d As Direct3D8
Public ddevice As Direct3DDevice8
Public d3dx As D3DX8

Dim d3dpp As D3DPRESENT_PARAMETERS

'Texture Manager for Dinamic Textures
Public DXPool As clsTextureManager

'Main form handle
Dim form_hwnd As Long

'Display variables
Dim screen_hwnd As Long

'FPS Counters
Dim fps_last_time As Long 'When did we last check the frame rate?
Dim fps_frame_counter As Long 'How many frames have been drawn
Dim FPS As Long 'What the current frame rate is.....

Dim engine_render_started As Boolean

'Graphic Font List
Dim gfont_list() As tGraphicFont
Dim gfont_count As Long

'Font List
Private font_list() As DXFont
Private font_count As Integer


'***************************
'Constants
'***************************
'Engine
Private Const COLOR_KEY As Long = &HFF000000
Private Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
'PI
Private Const PI As Single = 3.14159265358979

'Old fashion BitBlt functions
Private Const SRCCOPY = &HCC0020

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcsrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

'Initialization
Public Function DXEngine_Initialize(ByVal f_hwnd As Long, ByVal s_hwnd As Long, ByVal windowed As Boolean)
'On Error GoTo errhandler
    Dim d3dcaps As D3DCAPS8
    Dim d3ddm As D3DDISPLAYMODE
    
    DXEngine_Initialize = True
    
    'Main display
    screen_hwnd = s_hwnd
    form_hwnd = f_hwnd
    
    '*******************************
    'Initialize root DirectX8 objects
    '*******************************
    Set dX = New DirectX8
    'Create the Direct3D object
    Set d3d = dX.Direct3DCreate
    'Create helper class
    Set d3dx = New D3DX8
    
    '*******************************
    'Initialize video device
    '*******************************
    Dim DevType As CONST_D3DDEVTYPE
    DevType = D3DDEVTYPE_HAL
    'Get the capabilities of the Direct3D device that we specify. In this case,
    'we'll be using the adapter default (the primiary card on the system).
    Call d3d.GetDeviceCaps(D3DADAPTER_DEFAULT, DevType, d3dcaps)
    'Grab some information about the current display mode.
    Call d3d.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, d3ddm)
    
    'Now we'll go ahead and fill the D3DPRESENT_PARAMETERS type.
    With d3dpp
        .windowed = 1
        .SwapEffect = D3DSWAPEFFECT_COPY
        .BackBufferFormat = d3ddm.Format 'current display depth
    End With
    'create device
    Set ddevice = d3d.CreateDevice(D3DADAPTER_DEFAULT, DevType, screen_hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)

    DeviceRenderStates
    
    '****************************************************
    'Inicializamos el manager de texturas
    '****************************************************
    Call DXPool.Texture_Initialize(500)
    
    '****************************************************
    'Clears the buffer to start rendering
    '****************************************************
    Device_Clear
    '****************************************************
    'Load Misc
    '****************************************************
    LoadGraphicFonts
    LoadFonts
    CargarParticulas
    Exit Function
ErrHandler:
    DXEngine_Initialize = False
End Function

Public Function DXEngine_BeginRender() As Boolean
On Error GoTo ErrorHandler:
    DXEngine_BeginRender = True
    
    'Check if we have the device
    If ddevice.TestCooperativeLevel <> D3D_OK Then
        Do
            DoEvents
        Loop While ddevice.TestCooperativeLevel = D3DERR_DEVICELOST
        
        DXPool.Texture_Remove_All
        Fonts_Destroy
        Device_Reset
        
        DeviceRenderStates
        LoadFonts
        LoadGraphicFonts
    End If
    
    '****************************************************
    'Render
    '****************************************************
    '*******************************
    'Erase the backbuffer so that it can be drawn on again
    Device_Clear
    '*******************************
    '*******************************
    'Start the scene
    ddevice.BeginScene
    '*******************************
    
    engine_render_started = True
Exit Function
ErrorHandler:
    DXEngine_BeginRender = False
    MsgBox "Error in Engine_Render_Start: " & Err.Number & ": " & Err.Description
End Function

Public Function DXEngine_EndRender() As Boolean
On Error GoTo ErrorHandler:
    DXEngine_EndRender = True

    If engine_render_started = False Then
        Exit Function
    End If
    
    '*******************************
    'End scene
    ddevice.EndScene
    '*******************************
    
    '*******************************
    'Flip the backbuffer to the screen
    Device_Flip
    '*******************************
    
    '*******************************
    'Calculate current frames per second
    If GetTickCount >= (fps_last_time + 1000) Then
        FPS = fps_frame_counter
        fps_frame_counter = 0
        fps_last_time = GetTickCount
    Else
        fps_frame_counter = fps_frame_counter + 1
    End If
    '*******************************
    

    
    
    engine_render_started = False
Exit Function
ErrorHandler:
    DXEngine_EndRender = False
    MsgBox "Error in Engine_Render_End: " & Err.Number & ": " & Err.Description
End Function

Private Sub Device_Clear()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    'Clear the back buffer
    ddevice.Clear 0, ByVal 0&, D3DCLEAR_TARGET, 0, 1#, 0
End Sub

Private Function Device_Reset() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Resets the device
'**************************************************************
On Error GoTo ErrHandler:
'On Error Resume Next

    'Be sure the scene is finished
    ddevice.EndScene
    'Reset device
    ddevice.Reset d3dpp
    
    DeviceRenderStates
       
Exit Function
ErrHandler:
    Device_Reset = Err.Number
End Function
Public Sub DXEngine_TextureRenderAdvance(ByVal texture_index As Long, ByVal dest_x As Long, ByVal dest_y As Long, ByVal Src_X As Long, ByVal Src_Y As Long, _
                                             ByVal dest_width As Long, ByVal dest_height As Long, ByVal src_width As Long, ByVal src_height As Long, ByRef rgb_list() As Long, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'This sub allow texture resizing
'
'**************************************************************

    
    Dim src_rect As RECT
    Dim dest_rect As RECT
    Dim temp_verts(3) As TLVERTEX
    Dim Texture As Direct3DTexture8
    Dim texture_width As Integer
    Dim texture_height As Integer

    'rgb_list(0) = RGB(255, 255, 255)
    'rgb_list(1) = RGB(255, 255, 255)
    'rgb_list(2) = RGB(255, 255, 255)
    'rgb_list(3) = RGB(255, 255, 255)
    
    'Set up the destination rectangle
    With dest_rect
        .Bottom = dest_y + dest_height
        .left = dest_x
        .Right = dest_x + dest_width
        .top = dest_y
    End With
    
    With src_rect
        .Bottom = Src_Y + src_height
        .Right = Src_X + src_width
        .top = Src_Y
        .left = Src_X
    End With
    
    Set Texture = DXPool.GetTexture(texture_index)
    Call DXPool.Texture_Dimension_Get(texture_index, texture_width, texture_height)
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, rgb_list(), texture_width, texture_height, angle
    
    'Set Texture
    ddevice.SetTexture 0, Texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square texture
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
End Sub
Public Sub DXEngine_TextureRender(ByVal texture_index As Long, ByVal dest_x As Long, ByVal dest_y As Long, ByVal src_width As Long, _
                                            ByVal src_height As Long, ByRef rgb_list() As Long, ByVal Src_X As Long, _
                                            ByVal Src_Y As Long, ByVal dest_width As Long, ByVal dest_height As Long, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'This sub doesnt allow texture resizing
'
'**************************************************************
    Dim src_rect As RECT
    Dim dest_rect As RECT
    Dim temp_verts(3) As TLVERTEX
    Dim texture_height As Integer
    Dim texture_width As Integer
    Dim Texture As Direct3DTexture8
    
    'Set up the source rectangle
    With src_rect
        .Bottom = Src_Y + src_height - 1
        .left = Src_X
        .Right = Src_X + src_width - 1
        .top = Src_Y
    End With
        
    'Set up the destination rectangle
    With dest_rect
        .Bottom = dest_y + dest_height
        .left = dest_x
        .Right = dest_x + dest_width
        .top = dest_y
    End With
    
    'ESTO NO ME GUSTA
    Set Texture = DXPool.GetTexture(texture_index)
    Call DXPool.Texture_Dimension_Get(texture_index, texture_width, texture_height)
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, rgb_list(), texture_height, texture_width, angle
    'Set Texture
    ddevice.SetTexture 0, Texture
    
    'Enable alpha-blending
    'ddevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    
    If alpha_blend Then
       'Set Rendering for alphablending
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square texture
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
    
    'Turn off alphablending after we're done
    'ddevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
End Sub
Private Function Geometry_Create_TLVertex(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
                                            ByVal rhw As Single, ByVal Color As Long, ByVal Specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.X = X
    Geometry_Create_TLVertex.Y = Y
    Geometry_Create_TLVertex.Z = Z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.Color = Color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function

Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef texture_width As Integer, Optional ByRef texture_height As Integer, Optional ByVal angle As Single)
'**************************************************************
'Authors: Aaron Perkins;
'Last Modify Date: 5/07/2002
'
' * v1 *    v3
' |     \   |
' |     \   |
' |     \   |
' |     \   |
' |     \   |
' * v0 *    v2
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single
    
    If angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.left + (dest.Right - dest.left - 1) / 2
        y_center = dest.top + (dest.Bottom - dest.top - 1) / 2
        
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.Bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point
    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.left
        y_Cor = dest.Bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
    End If
    
    
    '0 - Bottom left vertex
    If texture_width And texture_height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.left / texture_width, (src.Bottom) / texture_height)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.left
        y_Cor = dest.top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
    End If
    
    
    '1 - Top left vertex
    If texture_width And texture_height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.left / texture_width, src.top / texture_height)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
    End If
    
    
    '2 - Bottom right vertex
    If texture_width And texture_height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right) / texture_width, (src.Bottom) / texture_height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
    End If
    
    
    '3 - Top right vertex
    If texture_width And texture_height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right) / texture_width, src.top / texture_height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 0)
    End If
End Sub

Public Sub DXEngine_GraphicTextRender(Font_Index As Integer, ByVal Text As String, ByVal top As Long, ByVal left As Long, _
                                  ByVal Color As Long)

    If Len(Text) > 255 Then Exit Sub
    
    Dim i As Byte
    Dim X As Integer
    Dim rgb_list(3) As Long
    
    For i = 0 To 3
        rgb_list(i) = Color
    Next i
    
    X = -1
    Dim Char As Integer
    For i = 1 To Len(Text)
        Char = AscB(mid$(Text, i, 1)) - 32
        
        If Char = 0 Then
            X = X + 1
        Else
            X = X + 1
            Call DXEngine_TextureRenderAdvance(gfont_list(Font_Index).texture_index, left + X * gfont_list(Font_Index).Char_Size, _
                                                        top, gfont_list(Font_Index).Caracteres(Char).Src_X, gfont_list(Font_Index).Caracteres(Char).Src_Y, _
                                                            gfont_list(Font_Index).Char_Size, gfont_list(Font_Index).Char_Size, gfont_list(Font_Index).Char_Size, gfont_list(Font_Index).Char_Size, _
                                                                rgb_list(), False)
        End If
    Next i
    
    
    
End Sub

Public Sub DXEngine_Deinitialize()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
On Error Resume Next

    'El manager de texturas es ahora independiente del engine.
    Call DXPool.Texture_Remove_All
    
    Set d3dx = Nothing
    Set ddevice = Nothing
    Set d3d = Nothing
    Set dX = Nothing
    Set DXPool = Nothing
End Sub

Private Sub LoadChars(ByVal Font_Index As Integer)
    Dim i As Integer
    Dim X As Integer
    Dim Y As Integer
    
    For i = 0 To 255
        With gfont_list(Font_Index).Caracteres(i)
            X = (i Mod 16) * gfont_list(Font_Index).Char_Size
            If X = 0 Then '16 chars per line
                Y = Y + 1
            End If
            .Src_X = X
            .Src_Y = (Y * gfont_list(Font_Index).Char_Size) - gfont_list(Font_Index).Char_Size
        End With
    Next i
End Sub
Public Sub LoadGraphicFonts()
    Dim i As Byte
    Dim file_path As String

    file_path = DirIndex & "GUIFonts.ini"

    If General_File_Exist(file_path, vbArchive) Then
        gfont_count = general_var_get(file_path, "INIT", "FontCount")
        If gfont_count > 0 Then
            ReDim gfont_list(1 To gfont_count) As tGraphicFont
            For i = 1 To gfont_count
                With gfont_list(i)
                    .Char_Size = general_var_get(file_path, "FONT" & i, "Size")
                    .texture_index = general_var_get(file_path, "FONT" & i, "Graphic")
                    If .texture_index > 0 Then Call DXPool.Texture_Load(.texture_index, 0)
                    LoadChars (i)
                End With
            Next i
        End If
    End If
End Sub

Public Sub DXEngine_StatsRender()
    'fps
    Call DXEngine_TextRender(1, FPS & " FPS", 0, 0, D3DColorXRGB(255, 255, 255))
End Sub

Private Sub Device_Flip()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    'Draw the graphics to the front buffer.
    ddevice.Present ByVal 0&, ByVal 0&, screen_hwnd, ByVal 0&
End Sub

Private Sub DeviceRenderStates()
    With ddevice
        'Set the vertex shader to an FVF that contains texture coords,
        'and transformed and lit vertex coords.
        .SetVertexShader FVF
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        
        'No se para q mierda sera esto.
        '.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        '.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        '.SetRenderState D3DRS_ZENABLE, True
        '.SetRenderState D3DRS_ZWRITEENABLE, False
        
        'Particle engine settings
        '.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        '.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        '.SetRenderState D3DRS_POINTSCALE_ENABLE, 0


    End With
End Sub

Private Sub Font_Make(ByVal Style As String, ByVal Size As Long, ByVal italic As Boolean, ByVal bold As Boolean)
    font_count = font_count + 1
    ReDim Preserve font_list(1 To font_count)
    
    Dim font_desc As IFont
    Dim fnt As New StdFont
    fnt.name = Style
    fnt.Size = Size
    fnt.bold = bold
    fnt.italic = italic
    Set font_desc = fnt
    font_list(font_count).Size = Size
    Set font_list(font_count).dFont = d3dx.CreateFont(ddevice, font_desc.hFont)
End Sub

Private Sub LoadFonts()
    Dim num_fonts As Integer
    Dim i As Integer
    Dim file_path As String
    
    file_path = DirIndex & "fonts.ini"
    
    If Not General_File_Exist(file_path, vbArchive) Then Exit Sub
    
    num_fonts = general_var_get(file_path, "INIT", "FontCount")
    
    For i = 1 To num_fonts
        Call Font_Make(general_var_get(file_path, "FONT" & i, "Name"), general_var_get(file_path, "FONT" & i, "Size"), general_var_get(file_path, "FONT" & i, "Cursiva"), general_var_get(file_path, "FONT" & i, "Negrita"))
    Next i
End Sub
Public Sub DXEngine_TextRender(ByVal Font_Index As Integer, ByVal Text As String, ByVal left As Integer, ByVal top As Integer, ByVal Color As Long, Optional ByVal Alingment As Byte = DT_LEFT, Optional ByVal Width As Integer = 0, Optional ByVal Height As Integer = 0)
    If Not Font_Check(Font_Index) Then Exit Sub
    
    Dim TextRect As RECT 'This defines where it will be
    'Dim BorderColor As Long
    
    'Set width and height if no specified
    If Width = 0 Then Width = Len(Text) * (font_list(Font_Index).Size + 1)
    If Height = 0 Then Height = font_list(Font_Index).Size * 2
    
    'DrawBorder
    
    'BorderColor = D3DColorXRGB(0, 0, 0)
    
    'TextRect.top = top - 1
    'TextRect.left = left
    'TextRect.bottom = top + height
    'TextRect.Right = left + width
    'd3dx.DrawText font_list(Font_Index).dFont, BorderColor, Text, TextRect, Alingment
    'TextRect.top = top
    'TextRect.left = left - 1
    'TextRect.bottom = top + height
    'TextRect.Right = left + width
    'd3dx.DrawText font_list(Font_Index).dFont, BorderColor, Text, TextRect, Alingment
    'TextRect.top = top + 1
    'TextRect.left = left
    'TextRect.bottom = top + height
    'TextRect.Right = left + width
    'd3dx.DrawText font_list(Font_Index).dFont, BorderColor, Text, TextRect, Alingment
    'TextRect.top = top
    'TextRect.left = left + 1
    'TextRect.bottom = top + height
    'TextRect.Right = left + width
    'd3dx.DrawText font_list(Font_Index).dFont, BorderColor, Text, TextRect, Alingment
    
    TextRect.top = top
    TextRect.left = left
    TextRect.Bottom = top + Height
    TextRect.Right = left + Width
    d3dx.DrawText font_list(Font_Index).dFont, Color, Text, TextRect, Alingment

End Sub
Private Function Font_Check(ByVal Font_Index As Long) As Boolean
    If Font_Index > 0 And Font_Index <= font_count Then
        Font_Check = True
    End If
End Function

Private Sub Fonts_Destroy()
    Dim i As Integer
    
    For i = 1 To font_count
        Set font_list(i).dFont = Nothing
        font_list(i).Size = 0
    Next i
    font_count = 0
End Sub

Public Function D3DColorValueGet(ByVal A As Byte, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte) As D3DCOLORVALUE
    D3DColorValueGet.A = A
    D3DColorValueGet.R = R
    D3DColorValueGet.G = G
    D3DColorValueGet.B = B
End Function

Public Sub DXEngine_TextureToHdcRender(ByVal texture_index As Long, desthdc As Long, ByVal screen_x As Long, ByVal screen_Y As Long, ByVal SX As Integer, ByVal SY As Integer, ByVal sW As Integer, ByVal sH As Integer, Optional transparent As Boolean = False)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/02/2003
'This method is SLOW... Don't use in a loop if you care about
'speed!
'*************************************************************

    Dim file_path As String
    Dim Src_X As Long
    Dim Src_Y As Long
    Dim src_width As Long
    Dim src_height As Long
    Dim hdcsrc As Long

    file_path = DirGraficos & texture_index & ".bmp"
    
    Src_X = SX
    Src_Y = SY
    src_width = sW
    src_height = sH

    hdcsrc = CreateCompatibleDC(desthdc)
    
    SelectObject hdcsrc, LoadPicture(file_path)
    
    If transparent = False Then
        BitBlt desthdc, screen_x, screen_Y, src_width, src_height, hdcsrc, Src_X, Src_Y, SRCCOPY
    Else
        TransparentBlt desthdc, screen_x, screen_Y, src_width, src_height, hdcsrc, Src_X, Src_Y, src_width, src_height, COLOR_KEY
    End If
        
    DeleteDC hdcsrc
End Sub

Public Sub DXEngine_BeginSecondaryRender()
    Device_Clear
    ddevice.BeginScene
End Sub
Public Sub DXEngine_EndSecondaryRender(ByVal hWnd As Long, ByVal Width As Integer, ByVal Height As Integer)
    Dim DR As RECT
    DR.left = 0
    DR.top = 0
    DR.Bottom = Height
    DR.Right = Width
    
    ddevice.EndScene
    ddevice.Present DR, ByVal 0&, hWnd, ByVal 0&
End Sub

Public Sub DXEngine_DrawBox(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Color As Long, Optional ByVal border_width = 1)
    Dim VertexB(3) As TLVERTEX
    Dim box_rect As RECT
    
    With box_rect
        .Bottom = Y + Height
        .left = X
        .Right = X + Width
        .top = Y
    End With
    
    ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
    ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        
    ddevice.SetTexture 0, Nothing
    
    'Upper Line
    VertexB(0) = Geometry_Create_TLVertex(box_rect.left, box_rect.top, 0, 1, Color, 0, 0, 0)
    VertexB(1) = Geometry_Create_TLVertex(box_rect.Right, box_rect.top, 0, 1, Color, 0, 0, 0)
    VertexB(2) = Geometry_Create_TLVertex(box_rect.left, box_rect.top + border_width, 0, 1, Color, 0, 0, 0)
    VertexB(3) = Geometry_Create_TLVertex(box_rect.Right, box_rect.top + border_width, 0, 1, Color, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexB(0), Len(VertexB(0))
    'Left Line
    VertexB(0) = Geometry_Create_TLVertex(box_rect.left + border_width, box_rect.top, 0, 1, Color, 0, 0, 0)
    VertexB(1) = Geometry_Create_TLVertex(box_rect.left + border_width, box_rect.Bottom, 0, 1, Color, 0, 0, 0)
    VertexB(2) = Geometry_Create_TLVertex(box_rect.left, box_rect.top, 0, 2, Color, 0, 0, 0)
    VertexB(3) = Geometry_Create_TLVertex(box_rect.left, box_rect.Bottom, 0, 2, Color, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexB(0), Len(VertexB(0))
    'Right Border
    VertexB(0) = Geometry_Create_TLVertex(box_rect.Right, box_rect.top, 0, 1, Color, 0, 0, 0)
    VertexB(1) = Geometry_Create_TLVertex(box_rect.Right, box_rect.Bottom, 0, 1, Color, 0, 0, 0)
    VertexB(2) = Geometry_Create_TLVertex(box_rect.Right - border_width, box_rect.top, 0, 3, Color, 0, 0, 0)
    VertexB(3) = Geometry_Create_TLVertex(box_rect.Right - border_width, box_rect.Bottom, 0, 3, Color, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexB(0), Len(VertexB(0))
    'Bottom Border
    VertexB(0) = Geometry_Create_TLVertex(box_rect.left, box_rect.Bottom - border_width, 0, 1, Color, 0, 0, 0)
    VertexB(1) = Geometry_Create_TLVertex(box_rect.Right, box_rect.Bottom - border_width, 0, 1, Color, 0, 0, 0)
    VertexB(2) = Geometry_Create_TLVertex(box_rect.left, box_rect.Bottom, 0, 1, Color, 0, 0, 0)
    VertexB(3) = Geometry_Create_TLVertex(box_rect.Right, box_rect.Bottom, 0, 1, Color, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexB(0), Len(VertexB(0))
    
    ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
End Sub
Public Sub D3DColorToRgbList(rgb_list() As Long, Color As D3DCOLORVALUE)
    rgb_list(0) = D3DColorARGB(Color.A, Color.R, Color.G, Color.B)
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub

Private Function Particle_Group_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
   
    loopc = 1
    Do Until particle_group_list(loopc).Active = False
        If loopc = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
   
    Particle_Group_Next_Open = loopc
Exit Function
ErrorHandler:
    Particle_Group_Next_Open = 1
End Function
Public Function Particle_Group_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean) As Long
   
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 12/15/2002
'Returns the particle_group_index if successful, else 0
'**************************************************************
    If (map_x <> -1) And (map_y <> -1) Then
    If Map_Particle_Group_Get(map_x, map_y) = 0 Then
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, X1, Y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, X2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin
    Else
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, X1, Y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, X2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin
    End If
    End If
End Function
 
Public Function Particle_Group_Remove(ByVal particle_group_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        Particle_Group_Destroy particle_group_index
        Particle_Group_Remove = True
    End If
End Function
 
Public Function Particle_Group_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim index As Long
   
    For index = 1 To particle_group_last
        'Make sure it's a legal index
        If Particle_Group_Check(index) Then
            Particle_Group_Destroy index
        End If
    Next index
   
    Particle_Group_Remove_All = True
End Function
 
Public Function Particle_Group_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
   
    loopc = 1
    Do Until particle_group_list(loopc).id = id
        If loopc = particle_group_last Then
            Particle_Group_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
   
    Particle_Group_Find = loopc
Exit Function
ErrorHandler:
    Particle_Group_Find = 0
End Function
 
Private Sub Particle_Group_Make(ByVal particle_group_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Makes a new particle effect
'*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
   
    'Make active
    particle_group_list(particle_group_index).Active = True
   
    'Map pos
    If (map_x <> -1) And (map_y <> -1) Then
        particle_group_list(particle_group_index).map_x = map_x
        particle_group_list(particle_group_index).map_y = map_y
    End If
   
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
   
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
   
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
   
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
   
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
   
    particle_group_list(particle_group_index).X1 = X1
    particle_group_list(particle_group_index).Y1 = Y1
    particle_group_list(particle_group_index).X2 = X2
    particle_group_list(particle_group_index).Y2 = Y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
   
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
   
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
   
    'plot particle group on map
    MapData(map_x, map_y).particle_group_index = particle_group_index
End Sub
 
Private Sub Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Integer, ByVal screen_Y As Integer, _
                            ByVal grh_index As Long, ByRef rgb_list() As Long, _
                            Optional ByVal alpha_blend As Boolean, Optional ByVal no_move As Boolean, _
                            Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                            Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                            Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                            Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                            Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                            Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                            Optional ByVal bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
                            Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                            Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                            Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                            Optional grh_resizex As Integer, Optional grh_resizey As Integer, _
                            Optional ByVal Radio As Integer, Optional ByVal Count As Integer, Optional ByVal index As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/24/2003
'
'**************************************************************

    If no_move = False Then
                If temp_particle.alive_counter = 0 Then
                    Grh_Initialize temp_particle.Grh, grh_index, alpha_blend
                    If Radio = 0 Then
                        temp_particle.X = General_Random_Number(X1, X2)
                        temp_particle.Y = General_Random_Number(Y1, Y2)
                    Else
                        temp_particle.X = (General_Random_Number(X1, X2) + Radio) + Radio * Cos(PI * 2 * index / Count)
                        temp_particle.Y = (General_Random_Number(Y1, Y2) + Radio) + Radio * Sin(PI * 2 * index / Count)
                    End If
                    temp_particle.vector_x = General_Random_Number(vecx1, vecx2)
                    temp_particle.vector_y = General_Random_Number(vecy1, vecy2)
                    temp_particle.angle = angle
                    temp_particle.alive_counter = General_Random_Number(life1, life2)
                    temp_particle.friction = fric
                Else
                    'Continue old particle
                    'Do gravity
                    If gravity = True Then
                        temp_particle.vector_y = temp_particle.vector_y + grav_strength
                        If temp_particle.Y > 0 Then
                            'bounce
                            temp_particle.vector_y = bounce_strength
                        End If
                    End If
                    'Do rotation
                   If spin = True Then temp_particle.Grh.angle = temp_particle.Grh.angle + (General_Random_Number(spin_speedL, spin_speedH) / 100)
                    If temp_particle.angle >= 360 Then
                        temp_particle.angle = 0
                    End If
                                
                    If XMove = True Then temp_particle.vector_x = General_Random_Number(move_x1, move_x2)
                    If YMove = True Then temp_particle.vector_y = General_Random_Number(move_y1, move_y2)
                End If

        'Add in vector
        temp_particle.X = temp_particle.X + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.Y = temp_particle.Y + (temp_particle.vector_y \ temp_particle.friction)
    
        'decrement counter
         temp_particle.alive_counter = temp_particle.alive_counter - 1
    End If
    
 
    'Draw it
    If temp_particle.Grh.grh_index Then
        Grh_Render temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_Y, rgb_list()
    End If
End Sub
Public Sub Particle_Group_Render(ByVal particle_group_index As Long, ByVal screen_x As Integer, ByVal screen_Y As Integer)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 12/15/2002
'Renders a particle stream at a paticular screen point
'*****************************************************************
On Error GoTo Err
    Dim loopc As Long
    Dim temp_rgb(3) As Long
    Dim no_move As Boolean

    'Set colors
    temp_rgb(0) = particle_group_list(particle_group_index).rgb_list(0)
    temp_rgb(1) = particle_group_list(particle_group_index).rgb_list(1)
    temp_rgb(2) = particle_group_list(particle_group_index).rgb_list(2)
    temp_rgb(3) = particle_group_list(particle_group_index).rgb_list(3)
       
    If particle_group_list(particle_group_index).alive_counter Then
    
        'See if it is time to move a particle
        particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timer_ticks_per_frame
        If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
            particle_group_list(particle_group_index).frame_counter = 0
            no_move = False
        Else
            no_move = True
        End If
    
    
        'If it's still alive render all the particles inside
        For loopc = 1 To particle_group_list(particle_group_index).particle_count
        
            'Render particle
            Particle_Render particle_group_list(particle_group_index).particle_stream(loopc), _
                            screen_x, screen_Y, _
                            particle_group_list(particle_group_index).grh_index_list(Round(General_Random_Number(1, particle_group_list(particle_group_index).grh_index_count), 0)), _
                            temp_rgb(), _
                            particle_group_list(particle_group_index).alpha_blend, no_move, _
                            particle_group_list(particle_group_index).X1, particle_group_list(particle_group_index).Y1, particle_group_list(particle_group_index).angle, _
                            particle_group_list(particle_group_index).vecx1, particle_group_list(particle_group_index).vecx2, _
                            particle_group_list(particle_group_index).vecy1, particle_group_list(particle_group_index).vecy2, _
                            particle_group_list(particle_group_index).life1, particle_group_list(particle_group_index).life2, _
                            particle_group_list(particle_group_index).fric, particle_group_list(particle_group_index).spin_speedL, _
                            particle_group_list(particle_group_index).gravity, particle_group_list(particle_group_index).grav_strength, _
                            particle_group_list(particle_group_index).bounce_strength, particle_group_list(particle_group_index).X2, _
                            particle_group_list(particle_group_index).Y2, particle_group_list(particle_group_index).XMove, _
                            particle_group_list(particle_group_index).move_x1, particle_group_list(particle_group_index).move_x2, _
                            particle_group_list(particle_group_index).move_y1, particle_group_list(particle_group_index).move_y2, _
                            particle_group_list(particle_group_index).YMove, particle_group_list(particle_group_index).spin_speedH, _
                            particle_group_list(particle_group_index).spin, , , , _
                            , particle_group_list(particle_group_index).particle_count, loopc
                            
        Next loopc
        
        If no_move = False Then
            'Update the group alive counter
            If particle_group_list(particle_group_index).never_die = False Then
                particle_group_list(particle_group_index).alive_counter = particle_group_list(particle_group_index).alive_counter - 1
            End If
        End If
    
    Else
        'If it's dead destroy it
        particle_group_list(particle_group_index).particle_count = particle_group_list(particle_group_index).particle_count - 1
        If particle_group_list(particle_group_index).particle_count <= 0 Then Particle_Group_Destroy particle_group_index
    End If
    
Err:
    temp_rgb(0) = 0
    temp_rgb(1) = 0
    temp_rgb(2) = 0
    temp_rgb(3) = 0
End Sub
 
Public Function Particle_Type_Get(ByVal particle_Index As Long) As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 8/27/2003
'Returns the stream type of a particle stream
'*****************************************************************
    If Particle_Group_Check(particle_Index) Then
        Particle_Type_Get = particle_group_list(particle_Index).stream_type
    End If
End Function
 
Private Function Particle_Group_Check(ByVal particle_group_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check index
    If particle_group_index > 0 And particle_group_index <= particle_group_last Then
        If particle_group_list(particle_group_index).Active Then
            Particle_Group_Check = True
        End If
    End If
End Function
 
Public Function Particle_Group_Map_Pos_Set(ByVal particle_group_index As Long, ByVal map_x As Long, ByVal map_y As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/27/2003
'Returns true if successful, else false
'**************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        'Make sure it's a legal move
        If InMapBounds(map_x, map_y) Then
            'Move it
            particle_group_list(particle_group_index).map_x = map_x
            particle_group_list(particle_group_index).map_y = map_y
   
            Particle_Group_Map_Pos_Set = True
        End If
    End If
End Function
 
Private Sub Particle_Group_Destroy(ByVal particle_group_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As particle_group
    
    If particle_group_list(particle_group_index).map_x > 0 And particle_group_list(particle_group_index).map_y > 0 Then
        MapData(particle_group_list(particle_group_index).map_x, particle_group_list(particle_group_index).map_y).particle_group_index = 0
    End If
    
    particle_group_list(particle_group_index) = temp
            
    'Update array size
    If particle_group_index = particle_group_last Then
        Do Until particle_group_list(particle_group_last).Active
            particle_group_last = particle_group_last - 1
            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count - 1
End Sub


Public Function Map_Particle_Group_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long
 
    If InMapBounds(map_x, map_y) Then
        Map_Particle_Group_Get = MapData(map_x, map_y).particle_group_index
    Else
        Map_Particle_Group_Get = 0
    End If
End Function
