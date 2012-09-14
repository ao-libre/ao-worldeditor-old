VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextDrawer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECTO ' xD
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, _
    ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, _
    ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, _
    ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function InvalidateRect Lib "user32" ( _
    ByVal hwnd As Long, lpRect As RECTO, ByVal bErase As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_PAINT As Long = &HF

Private Const TEXT_TRANSPARENT As Long = 1
Private Const TEXT_OPAQUE As Long = 2

Private Const FW_DONTCARE As Long = 0
Private Const FW_THIN  As Long = 100
Private Const FW_EXTRALIGHT  As Long = 200
Private Const FW_ULTRALIGHT  As Long = 200
Private Const FW_LIGHT  As Long = 300
Private Const FW_NORMAL As Long = 400
Private Const FW_REGULAR As Long = 400
Private Const FW_MEDIUM  As Long = 500
Private Const FW_SEMIBOLD As Long = 600
Private Const FW_DEMIBOLD As Long = 600
Private Const FW_BOLD  As Long = 700
Private Const FW_EXTRABOLD  As Long = 800
Private Const FW_ULTRABOLD As Long = 800
Private Const FW_BLACK  As Long = 900
Private Const FW_HEAVY  As Long = 900

Private Const ANSI_CHARSET As Long = 0
Private Const DEFAULT_CHARSET As Long = 1
Private Const SYMBOL_CHARSET As Long = 2
Private Const SHIFTJIS_CHARSET As Long = 128
Private Const OEM_CHARSET  As Long = 255

Private Const OUT_TT_ONLY_PRECIS  As Long = 7
Private Const CLIP_DEFAULT_PRECIS  As Long = 0
Private Const CLIP_LH_ANGLES  As Long = &H10
Private Const PROOF_QUALITY  As Long = 2
Private Const TRUETYPE_FONTTYPE As Long = &H4

Private Const TEXTBUFFER_SIZE As Long = 1024 * 2

Private Type TextInfo
    X As Long
    y As Long
    color As Long
    centered As Boolean
    Text As String
End Type

Private Type Size
    cx As Long
    cy As Long
End Type

Private TextBuffer(TEXTBUFFER_SIZE - 1) As TextInfo
Private textCount As Long

Private lngFont As Long
Private oldFont As Long

Private Const MAX_COLORS As Byte = 255

Private DirectDraw As DirectDraw7
Private useVideoMemory As Boolean
Private TextSurfaces(MAX_COLORS - 1) As Long
Private Surface(MAX_COLORS - 1) As DirectDrawSurface7
Private CantColors As Long
Private surfacesize As Size
Private letters As String
Private lettersrect(255) As RECT

Public Function InitText(ByRef DD As DirectDraw7, ByVal videoMemory As Boolean)
'**************************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modify Date: 07/31/2010
'
'**************************************************************

Dim ret As Size
Dim tsize As Integer
Dim i As Long
Dim hdc As Long

Set DirectDraw = DD
useVideoMemory = videoMemory

hdc = GetDC(frmMain.hwnd)
lngFont = CreateFont(13, 0, 0, 0, FW_BOLD, 0, 0, 0, DEFAULT_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_LH_ANGLES Or CLIP_DEFAULT_PRECIS, PROOF_QUALITY, TRUETYPE_FONTTYPE, "Tahoma") ' Hardocodeo la que usan siempre, la de frmMain.
oldFont = SelectObject(hdc, lngFont)
tsize = 0
letters = ""

For i = 32 To 255
    letters = letters & Chr$(i)
    
    Call GetTextExtentPoint32(hdc, Chr$(i), 1, ret)
    
    lettersrect(i).Left = tsize
    lettersrect(i).Right = tsize + ret.cx
    lettersrect(i).Top = 0
    lettersrect(i).Bottom = 13
    
    tsize = tsize + ret.cx
    If i = 126 Then i = 160
Next i

surfacesize.cy = ret.cy
surfacesize.cx = tsize

If oldFont <> 0 Then
    Call SelectObject(hdc, oldFont)
End If
    
DeleteObject lngFont
Call ReleaseDC(frmMain.hwnd, hdc)

Set Surface(0) = CreateSurface(0) ' Creo el color Negro
CantColors = 1
End Function

Private Function GetTextSurface(ByVal color As Long) As DirectDrawSurface7
'**************************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modify Date: 07/31/2010
'
'**************************************************************
Dim i As Long

For i = 0 To CantColors
    If TextSurfaces(i) = color Then
        Set GetTextSurface = Surface(i)
        Exit Function
    End If
Next i

If CantColors < MAX_COLORS Then 'creo la nueva surface
    Set Surface(CantColors) = CreateSurface(color)
    TextSurfaces(CantColors) = color
    Set GetTextSurface = Surface(CantColors)
    CantColors = CantColors + 1
End If

End Function

Private Function CreateSurface(ByVal color As Long) As DirectDrawSurface7
'**************************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modify Date: 07/31/2010
'
'**************************************************************
On Error GoTo ErrHandler

    Dim newSurface As DirectDrawSurface7
    Dim ddsd As DDSURFACEDESC2
    Dim ddck As DDCOLORKEY
    Dim bkcolor As Long
    Dim hdc As Long
    
     'Set up the surface desc
    ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    If useVideoMemory Then
        ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
    Else
        ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End If
    ddsd.lWidth = surfacesize.cx
    ddsd.lHeight = surfacesize.cy
        
    Set newSurface = DirectDraw.CreateSurface(ddsd)
    bkcolor = &H0
    If color = vbBlack Then bkcolor = &H10101
    
    'Dibujo el texto
    hdc = newSurface.GetDC
    lngFont = CreateFont(13, 0, 0, 0, FW_BOLD, 0, 0, 0, DEFAULT_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_LH_ANGLES Or CLIP_DEFAULT_PRECIS, PROOF_QUALITY, TRUETYPE_FONTTYPE, "Tahoma") ' Hardocodeo la que usan siempre, la de frmMain.
    oldFont = SelectObject(hdc, lngFont)
    SetBkMode hdc, TEXT_OPAQUE
    SetBkColor hdc, bkcolor
    SetTextColor hdc, color
    TextOut hdc, 0, 0, letters, Len(letters)
    If oldFont <> 0 Then Call SelectObject(hdc, oldFont)
    DeleteObject lngFont
    newSurface.ReleaseDC hdc
    'Podria haber usado el DrawText del dx aca, pero soy jodido
    
    ddck.high = bkcolor
    ddck.low = bkcolor
    Call newSurface.SetColorKey(DDCKEY_SRCBLT, ddck)
    Set CreateSurface = newSurface
    
Exit Function

ErrHandler:
    If Err.Number = DDERR_OUTOFVIDEOMEMORY Then
        ' No video memory? Try system memory! We want all surfaces loaded!
        ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Resume
    End If
    'If Err.Number = DDERR_OUTOFMEMORY Or Err.Number = DDERR_OUTOFVIDEOMEMORY Then
    '        MsgBox "No hay memoria disponible! El programa abortará. Cierra algunos programas e intenta de nuevo"
    '        End
    'Else
        MsgBox "Un error inesperado ocurrió al intentar crear el texto" & ". " & vbCrLf & _
                "El código de error es " & CStr(Err.Number) & " - " & Err.Description & vbCrLf & vbCrLf & "Copia este mensaje y notifica a los administradores.", _
                vbOKOnly Or vbCritical Or vbExclamation, "Error"
        End
    'End If
End Function

Public Function DrawText(ByVal X As Long, ByVal y As Long, ByRef Text As String, ByVal color As Long, ByRef Surface As DirectDrawSurface7)
'**************************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modify Date: 07/31/2010
'
'**************************************************************
    Dim i As Integer
    Dim tx As Long
    Dim textSurface As DirectDrawSurface7
    Dim bytestring() As Byte
    
    bytestring = StrConv(Text, vbFromUnicode)
    tx = X
    
    Set textSurface = GetTextSurface(color)
    
    For i = 0 To UBound(bytestring)
        Call Surface.BltFast(tx, y, textSurface, lettersrect(bytestring(i)), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
        tx = tx + lettersrect(bytestring(i)).Right - lettersrect(bytestring(i)).Left
    Next i
End Function

Public Function DrawTextToDC(ByVal hdc As Long, Optional ByVal Shadow As Boolean, Optional ByVal FontSize As Long = 13) As Long
'**************************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modify Date: 07/31/2010
'
'**************************************************************

'Forma alternativa de dibujar texto con GDI
'La forma de usarla es llamando a la funcion AddText de ahi abajo
Dim i As Long
Dim lastcolor As Long
Dim ret As Size
Dim tRect As RECTO

'Aplico mi font
lngFont = CreateFont(FontSize, 0, 0, 0, FW_BOLD, 0, 0, 0, DEFAULT_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_LH_ANGLES Or CLIP_DEFAULT_PRECIS, PROOF_QUALITY, TRUETYPE_FONTTYPE, "Tahoma") ' Hardocodeo la que usan siempre, la de frmMain.
oldFont = SelectObject(hdc, lngFont)

'Le doy fondo transparente
SetBkMode hdc, TEXT_TRANSPARENT

'Primero dibujo toda la capa en negro
SetTextColor hdc, vbBlack

If Shadow Then
    For i = 0 To textCount - 1
        With (TextBuffer(i))
            If .centered Then ' ya lo dejo acomodado
                Call GetTextExtentPoint32(hdc, .Text, Len(.Text), ret)
                .X = .X - ret.cx \ 2
                .y = .y - ret.cy \ 2
            End If
            
            TextOut hdc, .X - 2, .y - 1, .Text, Len(.Text)
        End With
    Next i
End If

'Y ahora toda la capa en color
For i = 0 To textCount - 1
    With (TextBuffer(i))
        If .color <> lastcolor Then ' si es igual me ahorro una llamada a un api, costo beneficio me va
            SetTextColor hdc, .color
            lastcolor = .color
        End If
        
        If (Not Shadow) And .centered Then
            Call GetTextExtentPoint32(hdc, .Text, Len(.Text), ret)
            .X = .X - ret.cx \ 2
            .y = .y - ret.cy \ 2
        End If
        
        TextOut hdc, .X, .y, .Text, Len(.Text)
    End With
Next i


textCount = 0 ' Se resetea
If oldFont <> 0 Then Call SelectObject(hdc, oldFont)
DeleteObject lngFont

DrawTextToDC = 0
End Function

Public Function AddText(ByVal xPos As Integer, ByVal yPos As Integer, ByVal fontcolor As Long, ByRef strText As String, Optional ByVal centered As Boolean = False) As Long
'**************************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modify Date: 07/31/2010
'
'**************************************************************

If LenB(strText) > 0 Then
    If textCount < TEXTBUFFER_SIZE Then
        With TextBuffer(textCount)
            .X = xPos
            .y = yPos
            .color = fontcolor
            .Text = strText
            .centered = centered
        End With
        textCount = textCount + 1
        AddText = textCount
    Else
        AddText = -1
    End If
End If

End Function

Private Sub Class_Terminate()
'**************************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modify Date: 07/31/2010
'
'**************************************************************

Dim i As Long

For i = 0 To CantColors
    Set Surface(i) = Nothing
Next i

End Sub
