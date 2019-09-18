VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMeteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This Object provides info about the Ambient.
'Handles the rain, day and night.

'Time info
Private mHour      As Byte

Private mMinutes   As Byte

Private mSeconds   As Byte

Private Rainning   As Byte

'Light Info
Private MedioDia   As D3DCOLORVALUE

Private MediaNoche As D3DCOLORVALUE

Private LuzActual  As D3DCOLORVALUE

Public Sub Initialize()
    MedioDia.R = 255
    MedioDia.G = 255
    MedioDia.B = 255
    
    MediaNoche.R = 100
    MediaNoche.G = 100
    MediaNoche.B = 120

End Sub

Public Sub Set_Time(ByVal Hour As Byte, ByVal Minutes As Byte)
    mHour = Hour
    mMinutes = Minutes
    mSeconds = 0
    
    AmbientCalculate

End Sub

Private Sub AmbientCalculate()

    If mHour >= 12 Then
        Call D3DXColorLerp(LuzActual, MedioDia, MediaNoche, mHour / 23)
    ElseIf mHour < 12 Then
        Call D3DXColorLerp(LuzActual, MediaNoche, MedioDia, mHour / 12)

    End If
    
End Sub

Public Sub Get_AmbientLight(Color As D3DCOLORVALUE)
    Color = LuzActual

End Sub

Public Sub Rain_Toggle()

    If Rainning Then
        Rain_Stop
    Else
        Rain_Start

    End If

End Sub

Private Sub Rain_Stop()
    Rainning = 0

    'DXEngine_ParticleGroup_Destroy RainParticleIndex
    'RainParticleIndex = 0
End Sub

Private Sub Rain_Start()

    If Rainning Then Exit Sub
    
    Rainning = 1
    
    'RainParticleIndex = DXEngine_ParticleGroupCreate(0, 0, e_ParticleType.Rain, -1)
End Sub

Public Sub Get_AmbientColorRGB(Rgb_List() As Long)
    Rgb_List(0) = D3DColorARGB(LuzActual.A, LuzActual.R, LuzActual.G, LuzActual.B)
    Rgb_List(1) = Rgb_List(0)
    Rgb_List(2) = Rgb_List(0)
    Rgb_List(3) = Rgb_List(0)

End Sub
