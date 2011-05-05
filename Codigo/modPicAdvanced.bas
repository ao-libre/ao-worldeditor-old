Attribute VB_Name = "modPicAdvanced"
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
' modPicAdvanced
'
' @author Unknown
' @version Unknown
' @date Unknown

Option Explicit

' ----==== GDIPlus Const ====----
Public Const GdiPlusVersion As Long = 1
Private Const mimeJPG As String = "image/jpeg"
Private Const mimePNG As String = "image/png"
Private Const mimeTIFF As String = "image/tiff"

Private Const EncoderParameterValueTypeLong As Long = 4
Private Const EncoderQuality As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
Private Const EncoderCompression As String = "{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"
' ----==== Sonstige Types ====----
Private Type PICTDESC
    cbSizeOfStruct As Long
    picType As Long
    hgdiObj As Long
    hPalOrXYExt As Long
End Type

Private Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

' ----==== GDIPlus Types ====----
Private Type GDIPlusStartupInput
    GdiPlusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type GdiplusStartupOutput
    NotificationHook As Long
    NotificationUnhook As Long
End Type

Private Type EncoderParameter
    GUID As GUID
    NumberOfValues As Long
    type As Long
    Value As Long
End Type

Private Type EncoderParameters
    Count As Long
    Parameter(15) As EncoderParameter
End Type

Private Type ImageCodecInfo
    Clsid As GUID
    FormatID As GUID
    CodecNamePtr As Long
    DllNamePtr As Long
    FormatDescriptionPtr As Long
    FilenameExtensionPtr As Long
    MimeTypePtr As Long
    Flags As Long
    Version As Long
    SigCount As Long
    SigSize As Long
    SigPatternPtr As Long
    SigMaskPtr As Long
End Type

' ----==== GDI+ 5.xx und 6.xx Enumerationen ====----
Private Type ARGB
    Blue As Byte
    Green As Byte
    Red As Byte
    Alpha As Byte
End Type

Private Type ColorPalette
    Flags As PaletteFlags
    Count As Long
    Entries As ARGB
End Type

Public Enum EncoderValueConstants
    EncoderValueColorTypeCMYK = 0
    EncoderValueColorTypeYCCK = 1
    EncoderValueCompressionLZW = 2
    EncoderValueCompressionCCITT3 = 3
    EncoderValueCompressionCCITT4 = 4
    EncoderValueCompressionRle = 5
    EncoderValueCompressionNone = 6
    EncoderValueScanMethodInterlaced = 7
    EncoderValueScanMethodNonInterlaced = 8
    EncoderValueVersionGif87 = 9
    EncoderValueVersionGif89 = 10
    EncoderValueRenderProgressive = 11
    EncoderValueRenderNonProgressive = 12
    EncoderValueTransformRotate90 = 13
    EncoderValueTransformRotate180 = 14
    EncoderValueTransformRotate270 = 15
    EncoderValueTransformFlipHorizontal = 16
    EncoderValueTransformFlipVertical = 17
    EncoderValueMultiFrame = 18
    EncoderValueLastFrame = 19
    EncoderValueFlush = 20
    EncoderValueFrameDimensionTime = 21
    EncoderValueFrameDimensionResolution = 22
    EncoderValueFrameDimensionPage = 23
End Enum

Private Enum PaletteFlags
    PaletteFlagsHasAlpha = &H1
    PaletteFlagsGrayScale = &H2
    PaletteFlagsHalftone = &H4
End Enum

Private Enum PixelFormats
    PixelFormatUndefined = &H0&
    PixelFormatDontCare = PixelFormatUndefined
    PixelFormatMax = &HF&
    PixelFormat1_8 = &H100&
    PixelFormat4_8 = &H400&
    PixelFormat8_8 = &H800&
    PixelFormat16_8 = &H1000&
    PixelFormat24_8 = &H1800&
    PixelFormat32_8 = &H2000&
    PixelFormat48_8 = &H3000&
    PixelFormat64_8 = &H4000&
    PixelFormat16bppRGB555 = &H21005
    PixelFormat16bppRGB565 = &H21006
    PixelFormat16bppGrayScale = &H101004
    PixelFormat16bppARGB1555 = &H61007
    PixelFormat24bppRGB = &H21808
    PixelFormat32bppRGB = &H22009
    PixelFormat32bppARGB = &H26200A
    PixelFormat32bppPARGB = &HD200B
    PixelFormat48bppRGB = &H10300C
    PixelFormat64bppARGB = &H34400D
    PixelFormat64bppPARGB = &H1C400E
    PixelFormatGDI = &H20000
    PixelFormat1bppIndexed = &H30101
    PixelFormat4bppIndexed = &H30402
    PixelFormat8bppIndexed = &H30803
    PixelFormatAlpha = &H40000
    PixelFormatIndexed = &H10000
    PixelFormatPAlpha = &H80000
    PixelFormatExtended = &H100000
    PixelFormatCanonical = &H200000
End Enum
' ----==== Sonstige Enumerationen ====----
Public Enum TifCompressionType
    ' EncoderValueConstants.EncoderValueCompressionLZW
    TiffCompressionLZW = 2
    'EncoderValueConstants.EncoderValueCompressionCCITT3
    TiffCompressionCCITT3 = 3
    'EncoderValueConstants.EncoderValueCompressionCCITT4
    TiffCompressionCCITT4 = 4
    'EncoderValueConstants.EncoderValueCompressionRle
    TiffCompressionRle = 5
    'EncoderValueConstants.EncoderValueCompressionNone
    TiffCompressionNone = 6
End Enum
' ----==== GDIPlus Enums ====----
Public Enum Status 'GDI+ Status
    OK = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
    ProfileNotFound = 21
End Enum
' ----==== GDI+ 6.xx Enumerationen ====----
Private Enum DitherType
    DitherTypeNone = 0
    DitherTypeSolid = 1
    DitherTypeOrdered4x4 = 2
    DitherTypeOrdered8x8 = 3
    DitherTypeOrdered16x16 = 4
    DitherTypeOrdered91x91 = 5
    DitherTypeSpiral4x4 = 6
    DitherTypeSpiral8x8 = 7
    DitherTypeDualSpiral4x4 = 8
    DitherTypeDualSpiral8x8 = 9
    DitherTypeErrorDiffusion = 10
End Enum

Private Enum PaletteType
    PaletteTypeCustom = 0
    PaletteTypeOptimal = 1
    PaletteTypeFixedBW = 2
    PaletteTypeFixedHalftone8 = 3
    PaletteTypeFixedHalftone27 = 4
    PaletteTypeFixedHalftone64 = 5
    PaletteTypeFixedHalftone125 = 6
    PaletteTypeFixedHalftone216 = 7
    PaletteTypeFixedHalftone252 = 8
    PaletteTypeFixedHalftone256 = 9
End Enum
' ----==== GDI+ 5.xx und 6.xx API Deklarationen ====----
Private Declare Function GdipCloneBitmapArea Lib "gdiplus" _
    (ByVal X As Single, ByVal y As Single, ByVal Width As Single, _
    ByVal Height As Single, ByVal format As PixelFormats, _
    ByVal srcBitmap As Long, ByRef dstBitmap As Long) As Status

Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" _
    (ByVal FileName As Long, ByRef BITMAP As Long) As Status

Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" _
    (ByVal hbm As Long, ByVal hpal As Long, _
    ByRef BITMAP As Long) As Status

Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" _
    (ByVal BITMAP As Long, ByRef hbmReturn As Long, _
    ByVal background As Long) As Status

Private Declare Function GdipDisposeImage Lib "gdiplus" _
    (ByVal image As Long) As Status

Private Declare Function GdipGetImageEncoders Lib "gdiplus" _
    (ByVal numEncoders As Long, ByVal Size As Long, _
    ByRef Encoders As Any) As Status

Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" _
    (ByRef numEncoders As Long, ByRef Size As Long) As Status

Private Declare Function GdipGetImagePixelFormat Lib "gdiplus" _
    (ByVal image As Long, ByRef PixelFormat As PixelFormats) As Status

Private Declare Function GdipGetImageDimension Lib "gdiplus" _
    (ByVal image As Long, ByRef sngWidth As Single, _
    ByRef sngHeight As Single) As Status

Private Declare Function GdiplusShutdown Lib "gdiplus" _
    (ByVal token As Long) As Status

Private Declare Function GdiplusStartup Lib "gdiplus" _
    (ByRef token As Long, ByRef lpInput As GDIPlusStartupInput, _
    Optional ByRef lpOutput As Any) As Status

Private Declare Function GdipSaveImageToFile Lib "gdiplus" _
    (ByVal image As Long, ByVal FileName As Long, _
    ByRef clsidEncoder As GUID, _
    ByRef encoderParams As Any) As Status

' ----==== GDI+ 6.xx API Deklarationen ====----
Private Declare Function GdipBitmapConvertFormat Lib "gdiplus" _
    (ByVal pInputBitmap As Long, _
    ByVal PixelFormat As PixelFormats, _
    ByVal DitherType As DitherType, _
    ByVal PaletteType As PaletteType, _
    ByVal palette As Any, _
    ByVal alphaThresholdPercent As Single) As Status

Private Declare Function GdipInitializePalette Lib "gdiplus" _
    (ByVal palette As Any, _
    ByVal PaletteType As PaletteType, _
    ByVal optimalColors As Long, _
    ByVal useTransparentColor As Long, _
    ByVal BITMAP As Long) As Status
 
' ----==== OLE API Declarations ====----
Private Declare Function CLSIDFromString Lib "ole32" _
    (ByVal str As Long, id As GUID) As Long

Private Declare Sub OleCreatePictureIndirect Lib "oleaut32.dll" _
    (lpPictDesc As PICTDESC, riid As IID, ByVal fOwn As Boolean, _
    lplpvObj As Object)

' ----==== Kernel API Declarations ====----
Private Declare Function lstrlenW Lib "kernel32" _
    (lpString As Any) As Long

Private Declare Function lstrcpyW Lib "kernel32" _
    (lpString1 As Any, lpString2 As Any) As Long
    
Private Declare Function GetModuleHandle Lib "kernel32" _
    Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Declare Function LoadLibrary Lib "kernel32" _
    Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" _
    (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" _
    (ByVal hLibModule As Long) As Long
' ----==== Variablen ====----

Private GdipToken As Long
Private GdipInitialized As Boolean
Public UseGDI6 As Boolean

Public Function StartUpGDIPlus(ByVal GdipVersion As Long) As Status
 ' Initialisieren der GDI+ Instanz
 Dim GdipStartupInput As GDIPlusStartupInput
 GdipStartupInput.GdiPlusVersion = GdipVersion
 StartUpGDIPlus = GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
End Function

Public Function ShutdownGDIPlus() As Status
 ' Beendet GDI+ Instanz
 ShutdownGDIPlus = GdiplusShutdown(GdipToken)
End Function

Public Function Execute(ByVal lReturn As Status) As Status
 Dim lCurErr As Status
 If lReturn = Status.OK Then
 lCurErr = Status.OK
 Else
 lCurErr = lReturn
 MsgBox GdiErrorString(lReturn) & " GDI+ Error:" & lReturn, _
   vbOKOnly, "GDI Error"
 End If
 Execute = lCurErr
End Function

Private Function GdiErrorString(ByVal lError As Status) As String
Dim s As String
 
Select Case lError
    Case GenericError:
        s = "Generic Error."
        
    Case InvalidParameter
        s = "Invalid Parameter."
        
    Case OutOfMemory
        s = "Out Of Memory."
        
    Case ObjectBusy
        s = "Object Busy."
        
    Case InsufficientBuffer
        s = "Insufficient Buffer."
        
    Case NotImplemented
        s = "Not Implemented."
        
    Case Win32Error
        s = "Win32 Error."
        
    Case WrongState
        s = "Wrong State."
        
    Case Aborted
        s = "Aborted."
        
    Case FileNotFound
        s = "File Not Found."
        
    Case ValueOverflow
        s = "Value Overflow."
        
    Case AccessDenied
        s = "Access Denied."
        
    Case UnknownImageFormat
        s = "Unknown Image Format."
        
    Case FontFamilyNotFound
        s = "FontFamily Not Found."
        
    Case FontStyleNotFound
        s = "FontStyle Not Found."
        
    Case NotTrueTypeFont
        s = "Not TrueType Font."
        
    Case UnsupportedGdiplusVersion
        s = "Unsupported Gdiplus Version."
        
    Case GdiplusNotInitialized
        s = "Gdiplus Not Initialized."
        
    Case PropertyNotFound
        s = "Property Not Found."
        
    Case PropertyNotSupported
        s = "Property Not Supported."
        
    Case Else
        s = "Unknown GDI+ Error."
End Select
 
GdiErrorString = s
End Function

Public Function LoadPicturePlus(ByVal FileName As String) As StdPicture
Dim retStatus As Status
Dim lBitmap As Long
Dim hBitmap As Long
 
 ' Öffnet die Bilddatei in lBitmap
retStatus = Execute(GdipCreateBitmapFromFile(StrPtr(FileName), lBitmap))
 
If retStatus = OK Then
 
    ' Erzeugen einer GDI Bitmap lBitmap -> hBitmap
    retStatus = Execute(GdipCreateHBITMAPFromBitmap(lBitmap, hBitmap, 0))
 
    If retStatus = OK Then
        ' Erzeugen des StdPicture Objekts von hBitmap
        Set LoadPicturePlus = HandleToPicture(hBitmap, vbPicTypeBitmap)
    End If
 
    ' Lösche lBitmap
    Call Execute(GdipDisposeImage(lBitmap))
End If

End Function

Private Function HandleToPicture(ByVal hGDIHandle As Long, _
    ByVal ObjectType As PictureTypeConstants, _
    Optional ByVal hpal As Long = 0) As StdPicture
 
Dim tPictDesc As PICTDESC
Dim IID_IPicture As IID
Dim oPicture As IPicture
 
' Initialisiert die PICTDESC Structur
With tPictDesc
    .cbSizeOfStruct = Len(tPictDesc)
    .picType = ObjectType
    .hgdiObj = hGDIHandle
    .hPalOrXYExt = hpal
End With

' Initialisiert das IPicture Interface ID
With IID_IPicture
    .Data1 = &H7BF80981
    .Data2 = &HBF32
    .Data3 = &H101A
    .Data4(0) = &H8B
    .Data4(1) = &HBB
    .Data4(3) = &HAA
    .Data4(5) = &H30
    .Data4(6) = &HC
    .Data4(7) = &HAB
End With

' Erzeugen des Objekts
OleCreatePictureIndirect tPictDesc, IID_IPicture, True, oPicture

' Rückgabe des Pictureobjekts
Set HandleToPicture = oPicture
 
End Function

Private Function GetEncoderClsid(mimeType As String, pClsid As GUID) As Boolean
Dim num As Long
Dim Size As Long
Dim pImageCodecInfo() As ImageCodecInfo
Dim j As Long
Dim buffer As String
 
Call GdipGetImageEncodersSize(num, Size)

If (Size = 0) Then
    GetEncoderClsid = False '// fehlgeschlagen
    Exit Function
End If

ReDim pImageCodecInfo(0 To Size \ Len(pImageCodecInfo(0)) - 1)
Call GdipGetImageEncoders(num, Size, pImageCodecInfo(0))

For j = 0 To num - 1
    buffer = Space$(lstrlenW(ByVal pImageCodecInfo(j).MimeTypePtr))
    
    Call lstrcpyW(ByVal StrPtr(buffer), ByVal _
        pImageCodecInfo(j).MimeTypePtr)
      
    If (StrComp(buffer, mimeType, vbTextCompare) = 0) Then
        pClsid = pImageCodecInfo(j).Clsid
        Erase pImageCodecInfo
        GetEncoderClsid = True '// erfolgreich
        Exit Function
    End If
Next j

Erase pImageCodecInfo
GetEncoderClsid = False '// fehlgeschlagen
End Function

Public Function UseGDI_v_6xx() As Boolean
Dim hMod As Long
Dim Loaded As Boolean
Dim sFunction As String
Dim sModule As String

' GDIPLUS.DLL
sModule = "GDIPLUS"

' eine Funktion die erst ab der
' GDI+ 6.xx vorhanden ist
sFunction = "GdipDrawImageFX"

'Handle der DLL erhalten
hMod = GetModuleHandle(sModule)

' Falls DLL nicht registriert ...
If hMod = 0 Then
    ' DLL in den Speicher laden.
    hMod = LoadLibrary(sModule)
    If hMod Then Loaded = True
End If

If hMod Then
    If GetProcAddress(hMod, sFunction) Then UseGDI_v_6xx = True
End If

If Loaded Then Call FreeLibrary(hMod)
 
End Function

Private Function ConvertTo1bppIndexedAndSaveAsTiffGDI5( _
    ByVal sFileName As String, _
    ByVal lInBitmap As Long, _
    Optional ByVal eTifCompression As EncoderValueConstants _
    = EncoderValueCompressionNone) As Status
 
Dim lNewBitmap As Long
Dim sWidth As Single
Dim sHeight As Single
Dim tPicEncoder As GUID
Dim tEncoderParameters As EncoderParameters

' Ermitteln der CLSID vom mimeType Encoder
If GetEncoderClsid(mimeTIFF, tPicEncoder) = True Then

    ' Initialisieren der Encoderparameter
    tEncoderParameters.Count = 1
    
    With tEncoderParameters.Parameter(0)
        ' Setzen der Kompressions GUID
        CLSIDFromString StrPtr(EncoderCompression), .GUID
        .NumberOfValues = 1
        .type = EncoderParameterValueTypeLong
        ' Kompressionstyp
        .Value = VarPtr(eTifCompression)
    End With
    
    ' Dimensionen von lInBitmap ermitteln
    If Execute(GdipGetImageDimension(lInBitmap, sWidth, sHeight)) = OK Then
        ' 1bppIndexed kopie von lInBitmap
        ' erstellen -> lNewBitmap
        If Execute(GdipCloneBitmapArea( _
            0, 0, sWidth, sHeight, _
            PixelFormat1bppIndexed, _
            lInBitmap, lNewBitmap)) = OK Then
            
            ' Speichert lNewBitmap als
            ' 1bppIndexed Tiff
            ConvertTo1bppIndexedAndSaveAsTiffGDI5 = _
                Execute(GdipSaveImageToFile( _
                lNewBitmap, StrPtr(sFileName), _
                tPicEncoder, tEncoderParameters))
            
            ' Lösche lNewBitmap
            Call Execute(GdipDisposeImage(lNewBitmap))
        End If
    End If
Else
    ' speichern nicht erfolgreich
    ConvertTo1bppIndexedAndSaveAsTiffGDI5 = Aborted
    MsgBox "Konnte keinen passenden Encoder ermitteln.", vbOKOnly, "Encoder Error"
End If
End Function

Private Function ConvertTo1bppIndexedAndSaveAsTiffGDI6( _
    ByVal sFileName As String, _
    ByVal lInBitmap As Long, _
    Optional ByVal eTifCompression As EncoderValueConstants _
    = EncoderValueCompressionNone) As Status
 
Dim lNewBitmap As Long
Dim sWidth As Single
Dim sHeight As Single
Dim tPicEncoder As GUID
Dim ePixelFormat As PixelFormats
Dim tEncoderParameters As EncoderParameters
Dim tCPal() As ColorPalette

' Palette für 1bppIndexed dimensionieren
ReDim tCPal(0 To 1)

' Anzahl der Farben setzen
tCPal(0).Count = UBound(tCPal) + 1

' Ermitteln der CLSID vom mimeType Encoder
If GetEncoderClsid(mimeTIFF, tPicEncoder) = True Then
    
    ' Initialisieren der Encoderparameter
    tEncoderParameters.Count = 1
    
    With tEncoderParameters.Parameter(0)
        ' Setzen der Kompressions GUID
        CLSIDFromString StrPtr(EncoderCompression), .GUID
        .NumberOfValues = 1
        .type = EncoderParameterValueTypeLong
        ' Kompressionstyp
        .Value = VarPtr(eTifCompression)
    End With
    
    ' Dimensionen von lInBitmap ermitteln
    If Execute(GdipGetImageDimension( _
        lInBitmap, sWidth, sHeight)) = OK Then
      
        ' PixelFormat von lInBitmap ermitteln
        If Execute(GdipGetImagePixelFormat(lInBitmap, ePixelFormat)) = OK Then
          
            ' kopie von lInBitmap erstellen
            ' -> lNewBitmap
            If Execute(GdipCloneBitmapArea(0, 0, sWidth, sHeight, ePixelFormat, _
                lInBitmap, lNewBitmap)) = OK Then
              
                ' optimierte 1bppIndexed Palette
                ' für lNewBitmap erzeugen
                If Execute(GdipInitializePalette( _
                    VarPtr(tCPal(0)), PaletteTypeOptimal, _
                    tCPal(0).Count, CLng(Abs(False)), _
                    lNewBitmap)) = OK Then
                      
                    ' lNewBitmap zu 1bppIndexed Bitmap mit
                    ' erzeugter Palette konvertieren
                    If Execute(GdipBitmapConvertFormat( _
                        lNewBitmap, PixelFormat1bppIndexed, _
                        DitherTypeDualSpiral8x8, _
                        PaletteTypeOptimal, _
                        VarPtr(tCPal(0)), 0)) = OK Then
                        
                        ' Speichert lNewBitmap als
                        ' 1bppIndexed Tiff mit
                        ' optimierter Palette
                        ConvertTo1bppIndexedAndSaveAsTiffGDI6 = _
                            Execute(GdipSaveImageToFile( _
                            lNewBitmap, StrPtr(sFileName), _
                            tPicEncoder, tEncoderParameters))
                    End If
                End If
                
                ' Lösche lNewBitmap
                Call Execute(GdipDisposeImage(lNewBitmap))
            End If
        End If
    End If
Else
    ' speichern nicht erfolgreich
    ConvertTo1bppIndexedAndSaveAsTiffGDI6 = Aborted
    MsgBox "Konnte keinen passenden Encoder ermitteln.", _
        vbOKOnly, "Encoder Error"
End If
End Function

Private Function SaveAsTiff(ByVal sFileName As String, _
    ByVal lInBitmap As Long, _
    Optional ByVal eTifCompression As EncoderValueConstants _
    = EncoderValueCompressionNone) As Status
 
Dim tPicEncoder As GUID
Dim tEncoderParameters As EncoderParameters

' Ermitteln der CLSID vom mimeType Encoder
If GetEncoderClsid(mimeTIFF, tPicEncoder) = True Then
    
    ' Initialisieren der Encoderparameter
    tEncoderParameters.Count = 1
    
    With tEncoderParameters.Parameter(0)
        ' Setzen der Kompressions GUID
        CLSIDFromString StrPtr(EncoderCompression), .GUID
        .NumberOfValues = 1
        .type = EncoderParameterValueTypeLong
        ' Kompressionstyp
        .Value = VarPtr(eTifCompression)
    End With
    
    ' Speichert lInBitmap als Tiff
    SaveAsTiff = Execute(GdipSaveImageToFile(lInBitmap, _
        StrPtr(sFileName), tPicEncoder, _
        tEncoderParameters))

Else
    ' speichern nicht erfolgreich
    SaveAsTiff = Aborted
    MsgBox "Konnte keinen passenden Encoder ermitteln.", _
    vbOKOnly, "Encoder Error"
End If
End Function

Public Function SavePictureAsTiff(ByVal Pic As StdPicture, _
    ByVal sFileName As String, _
    Optional ByVal eTifCompression As EncoderValueConstants _
    = EncoderValueCompressionNone) As Boolean
 
Dim lRet As Status
Dim lBitmap As Long

' Erzeugt eine GDI+ Bitmap vom
' StdPicture Handle -> lBitmap
If Execute(GdipCreateBitmapFromHBITMAP( _
    Pic.handle, 0, lBitmap)) = OK Then
    
    ' Kompressionstyp
    Select Case eTifCompression
      
        Case EncoderValueCompressionNone, _
            EncoderValueCompressionLZW
          
            lRet = SaveAsTiff( _
                sFileName, lBitmap, eTifCompression)
          
        Case Else 'RLE, CCITT3, CCITT4
          
            ' für die Komprimierungsmodi RLE, CCITT3, CCITT4
            ' muss die Bitmap in ein 1bppIndexed Bitmap
            ' konvertiert werden
            
            ' wird GDI+ v6.xx verwendet
            If UseGDI6 Then
            
                ' !!! ab GDI+ Version 6.xx und höher !!!
                lRet = ConvertTo1bppIndexedAndSaveAsTiffGDI6( _
                sFileName, lBitmap, eTifCompression)
            Else
                ' oder GDI+ v5.xx
                
                ' !!! ab GDI+ Version 5.xx und höher !!!
                lRet = ConvertTo1bppIndexedAndSaveAsTiffGDI5( _
                sFileName, lBitmap, eTifCompression)
            End If
    
    End Select
    
    If lRet = OK Then
        ' speichern erfolgreich
        SavePictureAsTiff = True
    Else
        ' speichern nicht erfolgreich
        SavePictureAsTiff = False
    End If
    
    ' Lösche lBitmap
    Call Execute(GdipDisposeImage(lBitmap))
End If
End Function

Public Function SavePictureAsJPG(ByVal Pic As StdPicture, _
    ByVal FileName As String, Optional ByVal Quality As Long = 85) _
    As Boolean
 
Dim retStatus As Status
Dim retVal As Boolean
Dim lBitmap As Long
 
' Erzeugt eine GDI+ Bitmap vom StdPicture Handle -> lBitmap
retStatus = Execute(GdipCreateBitmapFromHBITMAP(Pic.handle, 0, lBitmap))

If retStatus = OK Then
 
    Dim PicEncoder As GUID
    Dim tParams As EncoderParameters
 
    '// Ermitteln der CLSID vom mimeType Encoder
    retVal = GetEncoderClsid(mimeJPG, PicEncoder)
    
    If retVal = True Then
      
        If Quality > 100 Then Quality = 100
        If Quality < 0 Then Quality = 0
      
        ' Initialisieren der Encoderparameter
        tParams.Count = 1
      
        With tParams.Parameter(0) ' Quality
            ' Setzen der Quality GUID
            CLSIDFromString StrPtr(EncoderQuality), .GUID
            .NumberOfValues = 1
            .type = EncoderParameterValueTypeLong
            .Value = VarPtr(Quality)
        End With
      
        ' Speichert lBitmap als JPG
        retStatus = Execute(GdipSaveImageToFile(lBitmap, _
        StrPtr(FileName), PicEncoder, tParams))
        
        If retStatus = OK Then
            SavePictureAsJPG = True
        Else
            SavePictureAsJPG = False
        End If
    Else
        SavePictureAsJPG = False
        MsgBox "Konnte keinen passenden Encoder ermitteln.", _
            vbOKOnly, "Encoder Error"
    End If
 
    ' Lösche lBitmap
    Call Execute(GdipDisposeImage(lBitmap))
End If
End Function

Public Function SavePictureAsPNG(ByVal Pic As StdPicture, _
    ByVal sFileName As String) As Boolean
 
Dim lBitmap As Long
Dim tPicEncoder As GUID

' Erzeugt eine GDI+ Bitmap vom
' StdPicture Handle -> lBitmap
If Execute(GdipCreateBitmapFromHBITMAP( _
    Pic.handle, 0, lBitmap)) = OK Then
    
    ' Ermitteln der CLSID vom mimeType Encoder
    If GetEncoderClsid(mimePNG, tPicEncoder) = True Then
      
        ' Speichert lBitmap als PNG
        If Execute(GdipSaveImageToFile(lBitmap, _
            StrPtr(sFileName), tPicEncoder, ByVal 0)) = OK Then
            
            ' speichern erfolgreich
            SavePictureAsPNG = True
        Else
            ' speichern nicht erfolgreich
            SavePictureAsPNG = False
        End If
    Else
        ' speichern nicht erfolgreich
        SavePictureAsPNG = False
        MsgBox "Konnte keinen passenden Encoder ermitteln.", _
            vbOKOnly, "Encoder Error"
    End If
    
    ' Lösche lBitmap
    Call Execute(GdipDisposeImage(lBitmap))

End If
End Function
