VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WorldEditor"
   ClientHeight    =   10740
   ClientLeft      =   390
   ClientTop       =   840
   ClientWidth     =   15705
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   716
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1047
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Montaña"
      Height          =   495
      Left            =   120
      TabIndex        =   118
      Top             =   10200
      Width           =   1455
   End
   Begin VB.PictureBox picRadar 
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      Height          =   1590
      Left            =   120
      ScaleHeight     =   106
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   107
      TabIndex        =   115
      Top             =   120
      Width           =   1605
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   1365
         Left            =   120
         Top             =   105
         Width           =   1365
      End
      Begin VB.Shape ApuntadorRadar 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   6  'Inside Solid
         DrawMode        =   6  'Mask Pen Not
         FillColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   600
         Top             =   600
         Width           =   375
      End
      Begin VB.Label POSX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X: ?? - Y:??"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   150
         Left            =   120
         TabIndex        =   117
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label FPS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FPS: ??"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   150
         Left            =   1065
         TabIndex        =   116
         Top             =   1440
         Width           =   450
      End
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   8
      Left            =   13560
      TabIndex        =   99
      Top             =   30
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1826
      Caption         =   "Particula"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":628A
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.PictureBox MainViewPic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   9165
      Left            =   4560
      ScaleHeight     =   611
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   700
      TabIndex        =   97
      Top             =   1425
      Width           =   10500
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   6
      Left            =   11160
      TabIndex        =   47
      Top             =   30
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1826
      Caption         =   "Tri&gger's (F12)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":6850
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   5
      Left            =   9840
      TabIndex        =   46
      Top             =   30
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   1826
      Caption         =   "&Objetos (F11)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":6E16
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   3
      Left            =   8640
      TabIndex        =   45
      Top             =   30
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1826
      Caption         =   "&NPC's (F8)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":7317
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   2
      Left            =   7200
      TabIndex        =   44
      Top             =   30
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1826
      Caption         =   "&Bloqueos (F7)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":76CB
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   1
      Left            =   5760
      TabIndex        =   43
      Top             =   30
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1826
      Caption         =   "&Translados (F6)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "frmMain.frx":7A4C
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   0
      Left            =   5160
      TabIndex        =   42
      Top             =   30
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1826
      Caption         =   "&Superficie (F5)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "frmMain.frx":B0AC
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H cmdQuitarFunciones 
      Height          =   435
      Left            =   1800
      TabIndex        =   41
      ToolTipText     =   "Quitar Todas las Funciones Activadas"
      Top             =   1320
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   767
      Caption         =   "&Quitar Funciones (F4)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632319
   End
   Begin VB.Timer TimAutoGuardarMapa 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3960
      Top             =   1920
   End
   Begin VB.TextBox StatTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3675
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "frmMain.frx":E5F2
      Top             =   6360
      Width           =   4320
   End
   Begin VB.PictureBox pPaneles 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   120
      Picture         =   "frmMain.frx":E632
      ScaleHeight     =   4365
      ScaleWidth      =   4365
      TabIndex        =   8
      Top             =   1800
      Width           =   4395
      Begin VB.Frame cParticulas 
         BackColor       =   &H80000007&
         Caption         =   "Particles"
         ForeColor       =   &H80000009&
         Height          =   1785
         Left            =   360
         TabIndex        =   110
         Top             =   360
         Visible         =   0   'False
         Width           =   1860
         Begin VB.TextBox txtParticula 
            Alignment       =   2  'Center
            BackColor       =   &H80000006&
            ForeColor       =   &H80000005&
            Height          =   375
            Left            =   900
            TabIndex        =   111
            Top             =   255
            Width           =   555
         End
         Begin WorldEditor.lvButtons_H cInsertarParticula 
            Height          =   405
            Left            =   75
            TabIndex        =   112
            Top             =   765
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   714
            Caption         =   "Insertar Particula"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin WorldEditor.lvButtons_H cQuitarParticula 
            Height          =   390
            Left            =   75
            TabIndex        =   113
            Top             =   1215
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   688
            Caption         =   "Quitar Particula"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000012&
            Caption         =   "Particula:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   150
            TabIndex        =   114
            Top             =   300
            Width           =   660
         End
      End
      Begin VB.Frame cLuces 
         BackColor       =   &H80000012&
         Caption         =   "Luces"
         ForeColor       =   &H80000009&
         Height          =   3405
         Left            =   165
         TabIndex        =   100
         Top             =   135
         Visible         =   0   'False
         Width           =   3975
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmMain.frx":20C96
            Left            =   2400
            List            =   "frmMain.frx":20CA6
            TabIndex        =   119
            Text            =   "Mañana"
            Top             =   600
            Width           =   1335
         End
         Begin VB.CheckBox LuzRedonda 
            BackColor       =   &H80000012&
            Caption         =   "Luces redondas"
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   2055
            TabIndex        =   109
            Top             =   2565
            Width           =   1515
         End
         Begin VB.Frame RGBCOLOR 
            BackColor       =   &H80000012&
            Caption         =   "RGB"
            ForeColor       =   &H00FFFFFF&
            Height          =   690
            Left            =   135
            TabIndex        =   103
            Top             =   315
            Width           =   1680
            Begin VB.TextBox G 
               BackColor       =   &H80000012&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000014&
               Height          =   315
               Left            =   600
               TabIndex        =   106
               Text            =   "1"
               Top             =   270
               Width           =   450
            End
            Begin VB.TextBox B 
               BackColor       =   &H80000012&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000014&
               Height          =   315
               Left            =   1095
               TabIndex        =   105
               Text            =   "1"
               Top             =   270
               Width           =   450
            End
            Begin VB.TextBox R 
               BackColor       =   &H80000012&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000014&
               Height          =   315
               Left            =   105
               TabIndex        =   104
               Text            =   "1"
               Top             =   270
               Width           =   450
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H80000012&
            Caption         =   "Rango"
            ForeColor       =   &H8000000E&
            Height          =   660
            Left            =   135
            TabIndex        =   101
            Top             =   1080
            Width           =   1695
            Begin VB.TextBox cRango 
               BackColor       =   &H80000012&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000014&
               Height          =   315
               Left            =   105
               TabIndex        =   102
               Text            =   "1"
               Top             =   240
               Width           =   1035
            End
         End
         Begin WorldEditor.lvButtons_H cInsertarLuz 
            Height          =   600
            Left            =   180
            TabIndex        =   107
            Top             =   2565
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   1058
            Caption         =   "Insertar Luz"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin WorldEditor.lvButtons_H cQuitarLuz 
            Height          =   600
            Left            =   180
            TabIndex        =   108
            Top             =   1860
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   1058
            Caption         =   "Quitar Luz"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   -2147483633
         End
      End
      Begin VB.TextBox tTY 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   86
         Text            =   "1"
         Top             =   960
         Visible         =   0   'False
         Width           =   2900
      End
      Begin VB.TextBox tTX 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   85
         Text            =   "1"
         Top             =   600
         Visible         =   0   'False
         Width           =   2900
      End
      Begin VB.TextBox tTMapa 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   84
         Text            =   "1"
         Top             =   240
         Visible         =   0   'False
         Width           =   2900
      End
      Begin WorldEditor.lvButtons_H cInsertarTrans 
         Height          =   375
         Left            =   240
         TabIndex        =   87
         Top             =   1320
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Insertar Translado"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarTransOBJ 
         Height          =   375
         Left            =   240
         TabIndex        =   88
         Top             =   1680
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "Colocar automaticamente &Objeto"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cUnionManual 
         Height          =   375
         Left            =   240
         TabIndex        =   89
         Top             =   2160
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Union con Mapa Adyacente (manual)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cUnionAuto 
         Height          =   375
         Left            =   240
         TabIndex        =   90
         Top             =   2520
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "Union con Mapas &Adyacentes (auto)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarTrans 
         Height          =   375
         Left            =   240
         TabIndex        =   91
         Top             =   3000
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Quitar Translados"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cCapas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         ItemData        =   "frmMain.frx":20CC5
         Left            =   1080
         List            =   "frmMain.frx":20CD5
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cGrh 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Left            =   2880
         TabIndex        =   73
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         Left            =   600
         TabIndex        =   72
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   0
         ItemData        =   "frmMain.frx":20CE5
         Left            =   120
         List            =   "frmMain.frx":20CE7
         Sorted          =   -1  'True
         TabIndex        =   71
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin WorldEditor.lvButtons_H cQuitarEnTodasLasCapas 
         Height          =   375
         Left            =   120
         TabIndex        =   74
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Quitar en &Capas 2 y 3"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarEnEstaCapa 
         Height          =   375
         Left            =   120
         TabIndex        =   75
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar en esta Capa"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cSeleccionarSuperficie 
         Height          =   735
         Left            =   2400
         TabIndex        =   76
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar Superficie"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         ItemData        =   "frmMain.frx":20CE9
         Left            =   3360
         List            =   "frmMain.frx":20CEB
         TabIndex        =   67
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         ItemData        =   "frmMain.frx":20CED
         Left            =   840
         List            =   "frmMain.frx":20CEF
         TabIndex        =   0
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   3
         ItemData        =   "frmMain.frx":20CF1
         Left            =   120
         List            =   "frmMain.frx":20CF3
         TabIndex        =   66
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   3
         Left            =   600
         TabIndex        =   65
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         ItemData        =   "frmMain.frx":20CF5
         Left            =   840
         List            =   "frmMain.frx":20CF7
         TabIndex        =   58
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         ItemData        =   "frmMain.frx":20CF9
         Left            =   3360
         List            =   "frmMain.frx":20CFB
         TabIndex        =   57
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         Left            =   600
         TabIndex        =   56
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   1
         ItemData        =   "frmMain.frx":20CFD
         Left            =   120
         List            =   "frmMain.frx":20CFF
         TabIndex        =   55
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   3210
         Index           =   4
         ItemData        =   "frmMain.frx":20D01
         Left            =   120
         List            =   "frmMain.frx":20D03
         TabIndex        =   54
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.PictureBox Picture5 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   10
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture6 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   11
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture7 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   12
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture8 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   13
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture9 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   14
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture11 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   50
         Top             =   0
         Width           =   0
      End
      Begin WorldEditor.lvButtons_H cQuitarTrigger 
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar Trigger's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cVerTriggers 
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Mostrar Trigger's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarTrigger 
         Height          =   735
         Left            =   2400
         TabIndex        =   53
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar Trigger"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   59
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar NPC's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   60
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   0
         Left            =   2400
         TabIndex        =   61
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cVerBloqueos 
         Height          =   495
         Left            =   120
         TabIndex        =   62
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   873
         Caption         =   "&Mostrar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarBloqueo 
         Height          =   735
         Left            =   120
         TabIndex        =   63
         Top             =   720
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1296
         Caption         =   "&Insertar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarBloqueo 
         Height          =   735
         Left            =   120
         TabIndex        =   64
         Top             =   1560
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1296
         Caption         =   "&Quitar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   68
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar OBJ's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   69
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar OBJ's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   2
         Left            =   2400
         TabIndex        =   70
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar Objetos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   1
         Left            =   2400
         TabIndex        =   83
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   82
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   81
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar NPC's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         ItemData        =   "frmMain.frx":20D05
         Left            =   840
         List            =   "frmMain.frx":20D07
         TabIndex        =   77
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         Left            =   600
         TabIndex        =   78
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   2
         ItemData        =   "frmMain.frx":20D09
         Left            =   120
         List            =   "frmMain.frx":20D0B
         TabIndex        =   79
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         ItemData        =   "frmMain.frx":20D0D
         Left            =   3360
         List            =   "frmMain.frx":20D0F
         TabIndex        =   80
         Text            =   "500"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lYver 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Y vertical:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   94
         Top             =   1005
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lXhor 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "X horizontal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   93
         Top             =   645
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lMapN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Mapa:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   92
         Top             =   285
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lbCapas 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Capa Actual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   120
         TabIndex        =   25
         Top             =   3195
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lbGrh 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Sup Actual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   2040
         TabIndex        =   24
         Top             =   3195
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de NPC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   2160
         TabIndex        =   23
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de OBJ:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   2160
         TabIndex        =   20
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de NPC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   2160
         TabIndex        =   16
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   5460
      Left            =   60
      ScaleHeight     =   364
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   297
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6270
      Width           =   4455
      Begin VB.CommandButton cmdConversor 
         Caption         =   "Convertir a .CSM"
         Height          =   420
         Left            =   1680
         TabIndex        =   120
         Top             =   3960
         Width           =   1620
      End
      Begin VB.PictureBox PreviewGrh 
         BackColor       =   &H00004000&
         FillColor       =   &H00C0C0C0&
         Height          =   3780
         Left            =   45
         ScaleHeight     =   3720
         ScaleWidth      =   4305
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   4365
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   2565
      Top             =   2025
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1290
      Left            =   1680
      TabIndex        =   2
      Top             =   30
      Width           =   3225
      Begin WorldEditor.lvButtons_H cmdInformacionDelMapa 
         Height          =   375
         Left            =   100
         TabIndex        =   95
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         Caption         =   "&Información del Mapa"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label lblMapVersion 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1440
         TabIndex        =   29
         Top             =   1010
         Width           =   105
      End
      Begin VB.Label lblMapMusica 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1440
         TabIndex        =   28
         Top             =   352
         Width           =   90
      End
      Begin VB.Label lblMapNombre 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Nuevo Mapa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1440
         TabIndex        =   27
         Top             =   90
         Width           =   900
      End
      Begin VB.Label lblFMusica 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Musica:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   270
         Left            =   105
         TabIndex        =   7
         Top             =   320
         Width           =   3015
      End
      Begin VB.Label lblFVersion 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Versión:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Left            =   105
         TabIndex        =   6
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label lblFNombreMapa 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre del Mapa:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   270
         Left            =   105
         TabIndex        =   3
         Top             =   60
         Width           =   3015
      End
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   675
      Index           =   4
      Left            =   9840
      TabIndex        =   96
      Top             =   240
      Visible         =   0   'False
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1191
      Caption         =   "none"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":20D11
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   7
      Left            =   12240
      TabIndex        =   98
      Top             =   30
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1826
      Caption         =   "Luces "
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":210C5
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.Line Separacion1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   329
      X2              =   329
      Y1              =   8
      Y2              =   88
   End
   Begin VB.Line Separacion2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   337
      X2              =   337
      Y1              =   8
      Y2              =   88
   End
   Begin VB.Line Separacion2 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   336
      X2              =   336
      Y1              =   8
      Y2              =   88
   End
   Begin VB.Line Separacion1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   328
      X2              =   328
      Y1              =   8
      Y2              =   88
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   14340
      TabIndex        =   49
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   13575
      TabIndex        =   48
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5925
      TabIndex        =   40
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6690
      TabIndex        =   39
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   7455
      TabIndex        =   38
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   8220
      TabIndex        =   37
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   8985
      TabIndex        =   36
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   9750
      TabIndex        =   35
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   10515
      TabIndex        =   34
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   11280
      TabIndex        =   33
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   12045
      TabIndex        =   32
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   31
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   12810
      TabIndex        =   30
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Menu FileMnu 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuArchivoLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNuevoMapa 
         Caption         =   "&Nuevo Mapa"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuAbrirMapa 
         Caption         =   "&Abrir Mapa"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuArchivoLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReAbrirMapa 
         Caption         =   "&Re-Abrir Mapa"
      End
      Begin VB.Menu mnuArchivoLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardarMapa 
         Caption         =   "&Guardar Mapa"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuGuardarMapaComo 
         Caption         =   "Guardar Mapa &como..."
      End
      Begin VB.Menu mnuArchivoLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardarcomoBMP 
         Caption         =   "Guardar Render en &BMP"
      End
      Begin VB.Menu mnuGuardarcomoJPG 
         Caption         =   "Guardar Render en &JPG"
      End
      Begin VB.Menu mnuArchivoLine7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
      Begin VB.Menu mnuArchivoLine6 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "&Edición"
      Begin VB.Menu mnuComo 
         Caption         =   "¿ Como seleccionar ? ---- Mantener SHIFT y arrastrar el cursor."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCortar 
         Caption         =   "C&ortar Selección"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopiar 
         Caption         =   "&Copiar Selección"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPegar 
         Caption         =   "&Pegar Selección"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuBloquearS 
         Caption         =   "&Bloquear Selección"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuRealizarOperacion 
         Caption         =   "&Realizar Operación en Selección"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuDeshacerPegado 
         Caption         =   "Deshacer P&egado de Selección"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuLineEdicion0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeshacer 
         Caption         =   "&Deshacer"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuUtilizarDeshacer 
         Caption         =   "&Utilizar Deshacer"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuInfoMap 
         Caption         =   "&Información del Mapa"
      End
      Begin VB.Menu mnuLineEdicion1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertar 
         Caption         =   "&Insertar"
         Begin VB.Menu mnuInsertarTransladosAdyasentes 
            Caption         =   "&Translados a Mapas Adyasentes"
         End
         Begin VB.Menu mnuInsertarSuperficieAlAzar 
            Caption         =   "Superficie al &Azar"
         End
         Begin VB.Menu mnuInsertarSuperficieEnBordes 
            Caption         =   "Superficie en los &Bordes del Mapa"
         End
         Begin VB.Menu mnuInsertarSuperficieEnTodo 
            Caption         =   "Superficie en Todo el Mapa"
         End
         Begin VB.Menu mnuBloquearBordes 
            Caption         =   "Bloqueo en &Bordes del Mapa"
         End
         Begin VB.Menu mnuBloquearMapa 
            Caption         =   "Bloqueo en &Todo el Mapa"
         End
      End
      Begin VB.Menu mnuQuitar 
         Caption         =   "&Quitar"
         Begin VB.Menu mnuQuitarTranslados 
            Caption         =   "Todos los &Translados"
         End
         Begin VB.Menu mnuQuitarBloqueos 
            Caption         =   "Todos los &Bloqueos"
         End
         Begin VB.Menu mnuQuitarNPCs 
            Caption         =   "Todos los &NPC's"
         End
         Begin VB.Menu mnuQuitarNPCsHostiles 
            Caption         =   "Todos los NPC's &Hostiles"
         End
         Begin VB.Menu mnuQuitarObjetos 
            Caption         =   "Todos los &Objetos"
         End
         Begin VB.Menu mnuQuitarTriggers 
            Caption         =   "Todos los Tri&gger's"
         End
         Begin VB.Menu mnuQuitarSuperficieBordes 
            Caption         =   "Superficie de los B&ordes"
         End
         Begin VB.Menu mnuQuitarSuperficieDeCapa 
            Caption         =   "Superficie de la &Capa Seleccionada"
         End
         Begin VB.Menu mnuLineEdicion2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuQuitarTODO 
            Caption         =   "TODO"
         End
      End
      Begin VB.Menu mnuLineEdicion3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFunciones 
         Caption         =   "&Funciones"
         Begin VB.Menu mnuQuitarFunciones 
            Caption         =   "&Quitar Funciones"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuAutoQuitarFunciones 
            Caption         =   "Auto-&Quitar Funciones"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuConfigAvanzada 
         Caption         =   "Configuracion A&vanzada de Superficie"
      End
      Begin VB.Menu mnuLineEdicion4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoCompletarSuperficies 
         Caption         =   "Auto-Completar &Superficies"
      End
      Begin VB.Menu mnuAutoCapturarSuperficie 
         Caption         =   "Auto-C&apturar información de la Superficie"
      End
      Begin VB.Menu mnuAutoCapturarTranslados 
         Caption         =   "Auto-&Capturar información de los Translados"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAutoGuardarMapas 
         Caption         =   "Configuración de Auto-&Guardar Mapas"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver"
      Begin VB.Menu mnuCapas 
         Caption         =   "...&Capas"
         Begin VB.Menu mnuVerCapa1 
            Caption         =   "Capa &1 (Piso)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa2 
            Caption         =   "Capa &2 (costas, etc)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa3 
            Caption         =   "Capa &3 (arboles, etc)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa4 
            Caption         =   "Capa &4 (techos, etc)"
         End
      End
      Begin VB.Menu mnuVerTranslados 
         Caption         =   "...&Translados"
      End
      Begin VB.Menu mnuVerBloqueos 
         Caption         =   "...&Bloqueos"
      End
      Begin VB.Menu mnuVerNPCs 
         Caption         =   "...&NPC's"
      End
      Begin VB.Menu mnuVerObjetos 
         Caption         =   "...&Objetos"
      End
      Begin VB.Menu mnuVerTriggers 
         Caption         =   "...Tri&gger's"
      End
      Begin VB.Menu mnuVerGrilla 
         Caption         =   "...Gri&lla"
      End
      Begin VB.Menu mnuLinMostrar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVerAutomatico 
         Caption         =   "Control &Automaticamente"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuPaneles 
      Caption         =   "&Paneles"
      Begin VB.Menu mnuSuperficie 
         Caption         =   "&Superficie"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuTranslados 
         Caption         =   "&Translados"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuBloquear 
         Caption         =   "&Bloquear"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuNPCs 
         Caption         =   "&NPC's"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuNPCsHostiles 
         Caption         =   "NPC's &Hostiles"
         Shortcut        =   {F9}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuObjetos 
         Caption         =   "&Objetos"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuTriggers 
         Caption         =   "Tri&gger's"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuPanelesLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQSuperficie 
         Caption         =   "Ocultar Superficie"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuQTranslados 
         Caption         =   "Ocultar Translados"
         Shortcut        =   +{F6}
      End
      Begin VB.Menu mnuQBloquear 
         Caption         =   "Ocultar Bloquear"
         Shortcut        =   +{F7}
      End
      Begin VB.Menu mnuQNPCs 
         Caption         =   "Ocultar NPC's"
         Shortcut        =   +{F8}
      End
      Begin VB.Menu mnuQNPCsHostiles 
         Caption         =   "Ocultar NPC's Hostiles"
         Shortcut        =   +{F9}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQObjetos 
         Caption         =   "Ocultar Objetos"
         Shortcut        =   +{F11}
      End
      Begin VB.Menu mnuQTriggers 
         Caption         =   "Ocultar Trigger's"
         Shortcut        =   +{F12}
      End
      Begin VB.Menu mnuFuncionesLine1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuInformes 
         Caption         =   "&Informes"
      End
      Begin VB.Menu mnuActualizarIndices 
         Caption         =   "&Actualizar Indices de..."
         Begin VB.Menu mnuActualizarSuperficies 
            Caption         =   "&Superficies"
         End
         Begin VB.Menu mnuActualizarNPCs 
            Caption         =   "&NPC's"
         End
         Begin VB.Menu mnuActualizarObjs 
            Caption         =   "&Objetos"
         End
         Begin VB.Menu mnuActualizarTriggers 
            Caption         =   "&Trigger's"
         End
         Begin VB.Menu mnuActualizarCabezas 
            Caption         =   "C&abezas"
         End
         Begin VB.Menu mnuActualizarCuerpos 
            Caption         =   "C&uerpos"
         End
         Begin VB.Menu mnuActualizarGraficos 
            Caption         =   "&Graficos"
         End
      End
      Begin VB.Menu mnuModoCaminata 
         Caption         =   "Modalidad &Caminata"
      End
      Begin VB.Menu mnuGRHaBMP 
         Caption         =   "&GRH => BMP"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptimizar 
         Caption         =   "Optimi&zar Mapa"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardarUltimaConfig 
         Caption         =   "&Guardar Ultima Configuración"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ay&uda"
      Begin VB.Menu mnuManual 
         Caption         =   "&Manual"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuLineAyuda1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAcercaDe 
         Caption         =   "&Acerca de..."
      End
   End
   Begin VB.Menu mnuObjSc 
      Caption         =   "mnuObjSc"
      Visible         =   0   'False
      Begin VB.Menu mnuConfigObjTrans 
         Caption         =   "&Utilizar como Objeto de Translados"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Option Explicit

Private Sub PonerAlAzar(ByVal n As Integer, T As Byte)

    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06 by GS
    '*************************************************
    Dim objindex As Long

    Dim NPCIndex As Long

    Dim X, Y, i

    Dim Head    As Integer

    Dim Body    As Integer

    Dim Heading As Byte

    i = n

    modEdicion.Deshacer_Add "Aplicar " & IIf(T = 0, "Objetos", "NPCs") & " al Azar" ' Hago deshacer

    Do While i > 0
        X = CInt(RandomNumber(XMinMapSize, XMaxMapSize - 1))
        Y = CInt(RandomNumber(YMinMapSize, YMaxMapSize - 1))
    
        Select Case T

            Case 0

                If MapData(X, Y).OBJInfo.objindex = 0 Then
                    i = i - 1

                    If cInsertarBloqueo.Value = True Then
                        MapData(X, Y).Blocked = 1
                    Else
                        MapData(X, Y).Blocked = 0

                    End If

                    If cNumFunc(2).Text > 0 Then
                        objindex = cNumFunc(2).Text
                        Grh_Initialize MapData(X, Y).ObjGrh, ObjData(objindex).GrhIndex
                        MapData(X, Y).OBJInfo.objindex = objindex
                        MapData(X, Y).OBJInfo.Amount = Val(cCantFunc(2).Text)

                        Select Case ObjData(objindex).ObjType ' GS

                            Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos
                                MapData(X, Y).Graphic(3) = MapData(X, Y).ObjGrh

                        End Select

                    End If

                End If

            Case 1

                If MapData(X, Y).Blocked = 0 Then
                    i = i - 1

                    If cNumFunc(T - 1).Text > 0 Then
                        NPCIndex = cNumFunc(T - 1).Text
                        Body = NpcData(NPCIndex).Body
                        Head = NpcData(NPCIndex).Head
                        Heading = NpcData(NPCIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, CInt(X), CInt(Y))
                        MapData(X, Y).NPCIndex = NPCIndex

                    End If

                End If

            Case 2

                If MapData(X, Y).Blocked = 0 Then
                    i = i - 1

                    If cNumFunc(T - 1).Text >= 0 Then
                        NPCIndex = cNumFunc(T - 1).Text
                        Body = NpcData(NPCIndex).Body
                        Head = NpcData(NPCIndex).Head
                        Heading = NpcData(NPCIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, CInt(X), CInt(Y))
                        MapData(X, Y).NPCIndex = NPCIndex

                    End If

                End If

        End Select

        DoEvents
    Loop

End Sub

Private Sub cAgregarFuncalAzar_Click(index As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    On Error Resume Next

    If IsNumeric(cCantFunc(index).Text) = False Or cCantFunc(index).Text > 200 Then
        MsgBox "El Valor de Cantidad introducido no es soportado!" & vbCrLf & "El valor maximo es 200.", vbCritical
        Exit Sub

    End If

    cAgregarFuncalAzar(index).Enabled = False
    Call PonerAlAzar(CInt(cCantFunc(index).Text), 1 + (IIf(index = 2, -1, index)))
    cAgregarFuncalAzar(index).Enabled = True

End Sub

Private Sub cCantFunc_Change(index As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    On Error Resume Next

    If Val(cCantFunc(index)) < 1 Then
        cCantFunc(index).Text = 1

    End If

    If Val(cCantFunc(index)) > 10000 Then
        cCantFunc(index).Text = 10000

    End If

End Sub

Private Sub cCapas_Change()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 31/05/06
    '*************************************************
    If Val(cCapas.Text) < 1 Then
        cCapas.Text = 1

    End If

    If Val(cCapas.Text) > 4 Then
        cCapas.Text = 4

    End If

    cCapas.Tag = vbNullString

End Sub

Private Sub cCapas_KeyPress(KeyAscii As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0

End Sub

Private Sub cFiltro_GotFocus(index As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    HotKeysAllow = False

End Sub

Private Sub cFiltro_KeyPress(index As Integer, KeyAscii As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If KeyAscii = 13 Then
        Call Filtrar(index)

    End If

End Sub

Private Sub cFiltro_LostFocus(index As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    HotKeysAllow = True

End Sub

Private Sub cGrh_KeyPress(KeyAscii As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error GoTo Fallo

    If KeyAscii = 13 Then
        Call fPreviewGrh(cGrh.Text)

        If frmMain.cGrh.ListCount > 5 Then
            frmMain.cGrh.RemoveItem 0

        End If

        frmMain.cGrh.AddItem frmMain.cGrh.Text

    End If

    Exit Sub
Fallo:
    cGrh.Text = 1

End Sub

Private Sub cInsertarFunc_Click(index As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cInsertarFunc(index).Value = True Then
        cQuitarFunc(index).Enabled = False
        cAgregarFuncalAzar(index).Enabled = False

        If index <> 2 Then cCantFunc(index).Enabled = False
        Call modPaneles.EstSelectPanel((index) + 3, True)
    Else
        cQuitarFunc(index).Enabled = True
        cAgregarFuncalAzar(index).Enabled = True

        If index <> 2 Then cCantFunc(index).Enabled = True
        Call modPaneles.EstSelectPanel((index) + 3, False)

    End If

End Sub

Private Sub cInsertarLuz_Click()

    If cInsertarLuz.Value Then
        cQuitarLuz.Enabled = False
    Else
        cQuitarLuz.Enabled = True

    End If

End Sub

Private Sub cInsertarParticula_Click()

    If cInsertarParticula.Value Then
        cQuitarParticula.Enabled = False
    Else
        cQuitarParticula.Enabled = True

    End If

End Sub

Private Sub cInsertarTrans_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 22/05/06
    '*************************************************
    If cInsertarTrans.Value = True Then
        cQuitarTrans.Enabled = False
        Call modPaneles.EstSelectPanel(1, True)
    Else
        cQuitarTrans.Enabled = True
        Call modPaneles.EstSelectPanel(1, False)

    End If

End Sub

Private Sub cInsertarTrigger_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cInsertarTrigger.Value = True Then
        cQuitarTrigger.Enabled = False
        Call modPaneles.EstSelectPanel(6, True)
    Else
        cQuitarTrigger.Enabled = True
        Call modPaneles.EstSelectPanel(6, False)

    End If

End Sub

Private Sub cmdConversor_Click()
    
    Dim PathMAP As String
        PathMAP = InputBox("Donde estan los archivos .MAP (Path absoluto, por favor) ?")
    
    If Len(PathMAP) = 0 Then Exit Sub
    
    Dim PathCSM As String
        PathCSM = InputBox("Donde queres que se guarden los archivos .CSM (Path absoluto, por favor) ?")
    
    If Len(PathCSM) = 0 Then Exit Sub
    
    Dim Desde As String
        Desde = InputBox("Desde que mapa queres convertir a .CSM ?")
    
    If Len(Desde) = 0 Then Exit Sub
    
    Dim Hasta As String
        Hasta = InputBox("Hasta que mapa queres convertir a .CSM ?")
    
    If Len(Hasta) = 0 Then Exit Sub
    
    Call Convertir(PathMAP, PathCSM, CInt(Desde), CInt(Hasta))
    
End Sub

Private Sub cmdInformacionDelMapa_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    frmMapInfo.Show
    frmMapInfo.Visible = True

End Sub

Private Sub cmdQuitarFunciones_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call mnuQuitarFunciones_Click

End Sub

Private Sub Combo1_Click()

    Select Case Combo1.List(Combo1.ListIndex)

        Case "Mañana"
            Call Meteo.Set_Time(7, 0)

        Case "Dia"
            Call Meteo.Set_Time(12, 0)

        Case "Tarde"
            Call Meteo.Set_Time(18, 0)

        Case "Noche"
            Call Meteo.Set_Time(21, 0)

    End Select

End Sub

Private Sub Command1_Click()
    frmEngineTest.Show

End Sub

Private Sub cQuitarLuz_Click()

    If cQuitarLuz.Value Then
        cInsertarLuz.Enabled = False
    Else
        cInsertarLuz.Enabled = True

    End If

End Sub

Private Sub cQuitarParticula_Click()

    If cQuitarParticula.Value Then
        cInsertarParticula.Enabled = False
    Else
        cInsertarParticula.Enabled = True

    End If

End Sub

Private Sub cUnionManual_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    cInsertarTrans.Value = (cUnionManual.Value = True)
    Call cInsertarTrans_Click

End Sub

Private Sub cverBloqueos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    mnuVerBloqueos.Checked = cVerBloqueos.Value

End Sub

Private Sub cverTriggers_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    mnuVerTriggers.Checked = cVerTriggers.Value

End Sub

Private Sub cNumFunc_KeyPress(index As Integer, KeyAscii As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    On Error Resume Next

    If KeyAscii = 13 Then

        Dim Cont As String

        Cont = frmMain.cNumFunc(index).Text
        Call cNumFunc_LostFocus(index)

        If Cont <> frmMain.cNumFunc(index).Text Then Exit Sub
        If frmMain.cNumFunc(index).ListCount > 5 Then
            frmMain.cNumFunc(index).RemoveItem 0

        End If

        frmMain.cNumFunc(index).AddItem frmMain.cNumFunc(index).Text
        Exit Sub
    ElseIf KeyAscii = 8 Then
    
    ElseIf IsNumeric(Chr(KeyAscii)) = False Then
        KeyAscii = 0
        Exit Sub

    End If

End Sub

Private Sub cNumFunc_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    On Error Resume Next

    If cNumFunc(index).Text = vbNullString Then
        frmMain.cNumFunc(index).Text = IIf(index = 1, 500, 1)

    End If

End Sub

Private Sub cNumFunc_LostFocus(index As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    On Error Resume Next

    If index = 0 Then
        If frmMain.cNumFunc(index).Text > 499 Or frmMain.cNumFunc(index).Text < 1 Then
            frmMain.cNumFunc(index).Text = 1

        End If

    ElseIf index = 1 Then

        If frmMain.cNumFunc(index).Text < 500 Or frmMain.cNumFunc(index).Text > 32000 Then
            frmMain.cNumFunc(index).Text = 500

        End If

    ElseIf index = 2 Then

        If frmMain.cNumFunc(index).Text < 1 Or frmMain.cNumFunc(index).Text > 32000 Then
            frmMain.cNumFunc(index).Text = 1

        End If

    End If

End Sub

Private Sub cInsertarBloqueo_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    cInsertarBloqueo.Tag = vbNullString

    If cInsertarBloqueo.Value = True Then
        cQuitarBloqueo.Enabled = False
        Call modPaneles.EstSelectPanel(2, True)
    Else
        cQuitarBloqueo.Enabled = True
        Call modPaneles.EstSelectPanel(2, False)

    End If

End Sub

Private Sub cQuitarBloqueo_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    cInsertarBloqueo.Tag = vbNullString

    If cQuitarBloqueo.Value = True Then
        cInsertarBloqueo.Enabled = False
        Call modPaneles.EstSelectPanel(2, True)
    Else
        cInsertarBloqueo.Enabled = True
        Call modPaneles.EstSelectPanel(2, False)

    End If

End Sub

Private Sub cQuitarEnEstaCapa_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cQuitarEnEstaCapa.Value = True Then
        lListado(0).Enabled = False
        cFiltro(0).Enabled = False
        cGrh.Enabled = False
        cSeleccionarSuperficie.Enabled = False
        cQuitarEnTodasLasCapas.Enabled = False
        Call modPaneles.EstSelectPanel(0, True)
    Else
        lListado(0).Enabled = True
        cFiltro(0).Enabled = True
        cGrh.Enabled = True
        cSeleccionarSuperficie.Enabled = True
        cQuitarEnTodasLasCapas.Enabled = True
        Call modPaneles.EstSelectPanel(0, False)

    End If

End Sub

Private Sub cQuitarEnTodasLasCapas_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cQuitarEnTodasLasCapas.Value = True Then
        cCapas.Enabled = False
        lListado(0).Enabled = False
        cFiltro(0).Enabled = False
        cGrh.Enabled = False
        cSeleccionarSuperficie.Enabled = False
        cQuitarEnEstaCapa.Enabled = False
        Call modPaneles.EstSelectPanel(0, True)
    Else
        cCapas.Enabled = True
        lListado(0).Enabled = True
        cFiltro(0).Enabled = True
        cGrh.Enabled = True
        cSeleccionarSuperficie.Enabled = True
        cQuitarEnEstaCapa.Enabled = True
        Call modPaneles.EstSelectPanel(0, False)

    End If

End Sub

Private Sub cQuitarFunc_Click(index As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cQuitarFunc(index).Value = True Then
        cInsertarFunc(index).Enabled = False
        cAgregarFuncalAzar(index).Enabled = False
        cCantFunc(index).Enabled = False
        cNumFunc(index).Enabled = False
        cFiltro((index) + 1).Enabled = False
        lListado((index) + 1).Enabled = False
        Call modPaneles.EstSelectPanel((index) + 3, True)
    Else
        cInsertarFunc(index).Enabled = True
        cAgregarFuncalAzar(index).Enabled = True
        cCantFunc(index).Enabled = True
        cNumFunc(index).Enabled = True
        cFiltro((index) + 1).Enabled = True
        lListado((index) + 1).Enabled = True
        Call modPaneles.EstSelectPanel((index) + 3, False)

    End If

End Sub

Private Sub cQuitarTrans_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cQuitarTrans.Value = True Then
        cInsertarTransOBJ.Enabled = False
        cInsertarTrans.Enabled = False
        cUnionManual.Enabled = False
        cUnionAuto.Enabled = False
        tTMapa.Enabled = False
        tTX.Enabled = False
        tTY.Enabled = False
        mnuInsertarTransladosAdyasentes.Enabled = False
        Call modPaneles.EstSelectPanel(1, True)
    Else
        tTMapa.Enabled = True
        tTX.Enabled = True
        tTY.Enabled = True
        cUnionAuto.Enabled = True
        cUnionManual.Enabled = True
        cInsertarTrans.Enabled = True
        cInsertarTransOBJ.Enabled = True
        mnuInsertarTransladosAdyasentes.Enabled = True
        Call modPaneles.EstSelectPanel(1, False)

    End If

End Sub

Private Sub cQuitarTrigger_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cQuitarTrigger.Value = True Then
        lListado(4).Enabled = False
        cInsertarTrigger.Enabled = False
        Call modPaneles.EstSelectPanel(6, True)
    Else
        lListado(4).Enabled = True
        cInsertarTrigger.Enabled = True
        Call modPaneles.EstSelectPanel(6, False)

    End If

End Sub

Private Sub cSeleccionarSuperficie_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cSeleccionarSuperficie.Value = True Then
        cQuitarEnTodasLasCapas.Enabled = False
        cQuitarEnEstaCapa.Enabled = False
        Call modPaneles.EstSelectPanel(0, True)
    Else
        cQuitarEnTodasLasCapas.Enabled = True
        cQuitarEnEstaCapa.Enabled = True
        Call modPaneles.EstSelectPanel(0, False)

    End If

End Sub

Private Sub cUnionAuto_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    frmUnionAdyacente.Show

End Sub

Private Sub Form_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Me.SetFocus

End Sub

Private Sub MainViewPic_DblClick()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************

    If Not MapaCargado Then Exit Sub

    If SobreX > 0 And SobreY > 0 Then
        DobleClick Val(SobreX), Val(SobreY)

    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 24/11/08
    '*************************************************
    ' HotKeys
    If HotKeysAllow = False Then Exit Sub

    Select Case UCase(Chr(KeyAscii))

        Case "S" ' Activa/Desactiva Insertar Superficie
            cSeleccionarSuperficie.Value = (cSeleccionarSuperficie.Value = False)
            Call cSeleccionarSuperficie_Click

        Case "T" ' Activa/Desactiva Insertar Translados
            cInsertarTrans.Value = (cInsertarTrans.Value = False)
            Call cInsertarTrans_Click

        Case "B" ' Activa/Desactiva Insertar Bloqueos
            cInsertarBloqueo.Value = (cInsertarBloqueo.Value = False)
            Call cInsertarBloqueo_Click

        Case "N" ' Activa/Desactiva Insertar NPCs
            cInsertarFunc(0).Value = (cInsertarFunc(0).Value = False)
            Call cInsertarFunc_Click(0)

            ' Case "H" ' Activa/Desactiva Insertar NPCs Hostiles
            '     cInsertarFunc(1).value = (cInsertarFunc(1).value = False)
            '     Call cInsertarFunc_Click(1)
        Case "O" ' Activa/Desactiva Insertar Objetos
            cInsertarFunc(2).Value = (cInsertarFunc(2).Value = False)
            Call cInsertarFunc_Click(2)

        Case "G" ' Activa/Desactiva Insertar Triggers
            cInsertarTrigger.Value = (cInsertarTrigger.Value = False)
            Call cInsertarTrigger_Click

        Case "Q" ' Quitar Funciones
            Call mnuQuitarFunciones_Click

    End Select

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'If Seleccionando Then CopiarSeleccion
End Sub

Private Sub lListado_Click(index As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    On Error Resume Next

    If HotKeysAllow = False Then
        lListado(index).Tag = lListado(index).ListIndex

        Select Case index

            Case 0
                cGrh.Text = DameGrhIndex(general_field_read(2, lListado(index).Text, Asc("#")))

                If SupData(general_field_read(2, lListado(index).Text, Asc("#"))).Capa <> 0 Then
                    If LenB(general_field_read(2, lListado(index).Text, Asc("#"))) = 0 Then cCapas.Tag = cCapas.Text
                    cCapas.Text = SupData(general_field_read(2, lListado(index).Text, Asc("#"))).Capa
                Else

                    If LenB(cCapas.Tag) <> 0 Then
                        cCapas.Text = cCapas.Tag
                        cCapas.Tag = vbNullString

                    End If

                End If

                If SupData(general_field_read(2, lListado(index).Text, Asc("#"))).Block = True Then
                    If LenB(cInsertarBloqueo.Tag) = 0 Then cInsertarBloqueo.Tag = IIf(cInsertarBloqueo.Value = True, 1, 0)
                    cInsertarBloqueo.Value = True
                    Call cInsertarBloqueo_Click
                Else

                    If LenB(cInsertarBloqueo.Tag) <> 0 Then
                        cInsertarBloqueo.Value = IIf(Val(cInsertarBloqueo.Tag) = 1, True, False)
                        cInsertarBloqueo.Tag = vbNullString
                        Call cInsertarBloqueo_Click

                    End If

                End If

                Call fPreviewGrh(cGrh.Text)
            
                If frmMain.PreviewGrh.Visible = True Then
                    Call modPaneles.VistaPreviaDeSup

                End If

            Case 1
                cNumFunc(0).Text = general_field_read(2, lListado(index).Text, Asc("#"))

            Case 2
                cNumFunc(1).Text = general_field_read(2, lListado(index).Text, Asc("#"))

            Case 3
                cNumFunc(2).Text = general_field_read(2, lListado(index).Text, Asc("#"))

        End Select

    Else
        lListado(index).ListIndex = lListado(index).Tag

    End If

End Sub

Private Sub lListado_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    If index = 3 And Button = 2 Then
        If lListado(3).ListIndex > -1 Then Me.PopupMenu mnuObjSc

    End If

End Sub

Private Sub lListado_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 22/05/06
    '*************************************************
    On Error Resume Next

    HotKeysAllow = False

End Sub

Private Sub MapPest_Click(index As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If (index + NumMap_Save - 4) <> NumMap_Save Then
        Dialog.CancelError = True

        On Error GoTo ErrHandler

        Dialog.FileName = PATH_Save & NameMap_Save & (index + NumMap_Save - 4) & ".map"

        If MapInfo.Changed = 1 Then
            If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
                modMapIO.GuardarMapa Dialog.FileName

            End If

        End If

        Call modMapIO.NuevoMapa
        DoEvents
        modMapIO.AbrirMapa Dialog.FileName, MapData
        EngineRun = True
        Exit Sub
    
ErrHandler:
        MsgBox Err.Description

    End If

End Sub

Private Sub mnuAbrirMapa_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 04/11/2015
    '*************************************************
    
    Dialog.CancelError = True

    On Error GoTo ErrHandler

    Call DeseaGuardarMapa(Dialog.FileName)

    Call ObtenerNombreArchivo(False)

    If Len(Dialog.FileName) < 3 Then Exit Sub

    If WalkMode = True Then
        Call modGeneral.ToggleWalkMode

    End If
    
    Call modMapIO.NuevoMapa
    
    Select Case frmMain.Dialog.FilterIndex
    
        Case 1
            Call modMapIO.AbrirMapa(Dialog.FileName, MapData)
            
        Case 2
            Call modMapIO.Cargar_CSM(Dialog.FileName)
            
    End Select

    DoEvents
    
    mnuReAbrirMapa.Enabled = True
    
    EngineRun = True
    
    Exit Sub
    
ErrHandler:

End Sub

Private Sub mnuActualizarCabezas_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    Call modIndices.CargarIndicesDeCabezas

End Sub

Private Sub mnuActualizarCuerpos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    Call modIndices.CargarIndicesDeCuerpos

End Sub

Private Sub mnuActualizarGraficos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    Call Animations_Initialize(0.03, 32)

End Sub

Private Sub mnuActualizarSuperficies_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modIndices.CargarIndicesSuperficie

End Sub

Private Sub mnuacercade_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    frmAbout.Show

End Sub

Private Sub mnuActualizarNPCs_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modIndices.CargarIndicesNPC

End Sub

Private Sub mnuActualizarObjs_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modIndices.CargarIndicesOBJ

End Sub

Private Sub mnuActualizarTriggers_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modIndices.CargarIndicesTriggers

End Sub

Private Sub mnuAutoCapturarTranslados_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    mnuAutoCapturarTranslados.Checked = (mnuAutoCapturarTranslados.Checked = False)

End Sub

Private Sub mnuAutoCapturarSuperficie_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    mnuAutoCapturarSuperficie.Checked = (mnuAutoCapturarSuperficie.Checked = False)

End Sub

Private Sub mnuAutoCompletarSuperficies_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    mnuAutoCompletarSuperficies.Checked = (mnuAutoCompletarSuperficies.Checked = False)

End Sub

Private Sub mnuAutoGuardarMapas_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    frmAutoGuardarMapa.Show

End Sub

Private Sub mnuAutoQuitarFunciones_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    mnuAutoQuitarFunciones.Checked = (mnuAutoQuitarFunciones.Checked = False)

End Sub

Private Sub mnuBloquear_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Dim i As Byte

    For i = 0 To 6

        If i <> 2 Then
            frmMain.SelectPanel(i).Value = False
            Call VerFuncion(i, False)

        End If

    Next

    modPaneles.VerFuncion 2, True

End Sub

Private Sub mnuBloquearBordes_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Bloquear_Bordes

End Sub

Private Sub mnuBloquearMapa_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Bloqueo_Todo(1)

End Sub

Private Sub mnuBloquearS_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    Call modEdicion.Deshacer_Add("Bloquear Selección")
    Call BlockearSeleccion

End Sub

Private Sub mnuConfigAvanzada_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    frmConfigSup.Show

End Sub

Private Sub mnuConfigObjTrans_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    Cfg_TrOBJ = cNumFunc(2).Text

End Sub

Private Sub mnuCopiar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    Call CopiarSeleccion

End Sub

Private Sub mnuCortar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    Call modEdicion.Deshacer_Add("Cortar Selección")
    Call CortarSeleccion

End Sub

Private Sub mnuDeshacer_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 15/10/06
    '*************************************************
    Call modEdicion.Deshacer_Recover

End Sub

Private Sub mnuDeshacerPegado_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    Call modEdicion.Deshacer_Add("Deshacer Pegado de Selección")
    Call DePegar

End Sub

Private Sub mnuGRHaBMP_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    frmGRHaBMP.Show

End Sub

Private Sub mnuGuardarcomoBMP_Click()

    '*************************************************
    'Author: Salvito
    'Last modified: 01/05/2008 - ^[GS]^
    '*************************************************
    Dim Ratio As Integer

    Ratio = CInt(Val(InputBox("En que escala queres Renderizar? Entre 1 y 20.", "Elegi Escala", "1")))

    If Ratio < 1 Then Ratio = 1
    If Ratio >= 1 And Ratio <= 20 Then

        'RenderToPicture Ratio, True
    End If

End Sub

Private Sub mnuGuardarcomoJPG_Click()

    '*************************************************
    'Author: Salvito
    'Last modified: 01/05/2008 - ^[GS]^
    '*************************************************
    Dim Ratio As Integer

    Ratio = CInt(Val(InputBox("En que escala queres Renderizar? Entre 1 y 20.", "Elegi Escala", "1")))

    If Ratio < 1 Then Ratio = 1
    If Ratio >= 1 And Ratio <= 20 Then

        'RenderToPicture Ratio, False
    End If

End Sub

Private Sub mnuGuardarMapa_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    modMapIO.GuardarMapa Dialog.FileName

End Sub

Private Sub mnuGuardarMapaComo_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    modMapIO.GuardarMapa

End Sub

Private Sub mnuGuardarUltimaConfig_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 23/05/06
    '*************************************************
    mnuGuardarUltimaConfig.Checked = (mnuGuardarUltimaConfig.Checked = False)

End Sub

Private Sub mnuInfoMap_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    frmMapInfo.Show
    frmMapInfo.Visible = True

End Sub

Private Sub mnuInformes_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    frmInformes.Show

End Sub

Private Sub mnuInsertarSuperficieAlAzar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Superficie_Azar

End Sub

Private Sub mnuInsertarSuperficieEnBordes_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Superficie_Bordes

End Sub

Private Sub mnuInsertarSuperficieEnTodo_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Superficie_Todo

End Sub

Private Sub mnuInsertarTransladosAdyasentes_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    frmUnionAdyacente.Show

End Sub

Private Sub mnuManual_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 24/11/08
    '*************************************************
    If LenB(Dir(App.Path & "\manual\index.html", vbArchive)) <> 0 Then
        Call Shell("explorer " & App.Path & "\manual\index.html")
        DoEvents

    End If

End Sub

Private Sub mnuModoCaminata_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    ToggleWalkMode

End Sub

Private Sub mnuNPCs_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Dim i As Byte

    For i = 0 To 6

        If i <> 3 Then
            frmMain.SelectPanel(i).Value = False
            Call VerFuncion(i, False)

        End If

    Next
    modPaneles.VerFuncion 3, True

End Sub

'Private Sub mnuNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'Dim i As Byte
'For i = 0 To 6
'    If i <> 4 Then
'        frmMain.SelectPanel(i).value = False
'        Call VerFuncion(i, False)
'    End If
'Next
'modPaneles.VerFuncion 4, True
'End Sub

Private Sub mnuNuevoMapa_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    On Error Resume Next

    Dim loopc As Integer

    DeseaGuardarMapa Dialog.FileName

    For loopc = 0 To frmMain.MapPest.Count
        frmMain.MapPest(loopc).Visible = False
    Next

    frmMain.Dialog.FileName = Empty

    If WalkMode = True Then
        Call modGeneral.ToggleWalkMode

    End If

    Call modMapIO.NuevoMapa

    Call cmdInformacionDelMapa_Click

End Sub

Private Sub mnuObjetos_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Dim i As Byte

    For i = 0 To 6

        If i <> 5 Then
            frmMain.SelectPanel(i).Value = False
            Call VerFuncion(i, False)

        End If

    Next
    modPaneles.VerFuncion 5, True

End Sub

Private Sub mnuOptimizar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 22/09/06
    '*************************************************
    frmOptimizar.Show

End Sub

Private Sub mnuPegar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    Call modEdicion.Deshacer_Add("Pegar Selección")
    Call PegarSeleccion

End Sub

Private Sub mnuQBloquear_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    modPaneles.VerFuncion 2, False

End Sub

Private Sub mnuQNPCs_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    modPaneles.VerFuncion 3, False

End Sub

'Private Sub mnuQNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'modPaneles.VerFuncion 4, False
'End Sub

Private Sub mnuQObjetos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    modPaneles.VerFuncion 5, False

End Sub

Private Sub mnuQSuperficie_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    modPaneles.VerFuncion 0, False

End Sub

Private Sub mnuQTranslados_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    modPaneles.VerFuncion 1, False

End Sub

Private Sub mnuQTriggers_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    modPaneles.VerFuncion 6, False

End Sub

Private Sub mnuQuitarBloqueos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Bloqueo_Todo(0)

End Sub

Private Sub mnuQuitarFunciones_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    ' Superficies
    cSeleccionarSuperficie.Value = False
    Call cSeleccionarSuperficie_Click
    cQuitarEnEstaCapa.Value = False
    Call cQuitarEnEstaCapa_Click
    cQuitarEnTodasLasCapas.Value = False
    Call cQuitarEnTodasLasCapas_Click
    ' Translados
    cQuitarTrans.Value = False
    Call cQuitarTrans_Click
    cInsertarTrans.Value = False
    Call cInsertarTrans_Click
    ' Bloqueos
    cQuitarBloqueo.Value = False
    Call cQuitarBloqueo_Click
    cInsertarBloqueo.Value = False
    Call cInsertarBloqueo_Click
    ' Otras funciones
    cInsertarFunc(0).Value = False
    Call cInsertarFunc_Click(0)
    cInsertarFunc(1).Value = False
    Call cInsertarFunc_Click(1)
    cInsertarFunc(2).Value = False
    Call cInsertarFunc_Click(2)
    cQuitarFunc(0).Value = False
    Call cQuitarFunc_Click(0)
    cQuitarFunc(1).Value = False
    Call cQuitarFunc_Click(1)
    cQuitarFunc(2).Value = False
    Call cQuitarFunc_Click(2)
    ' Triggers
    cInsertarTrigger.Value = False
    Call cInsertarTrigger_Click
    cQuitarTrigger.Value = False
    Call cQuitarTrigger_Click

    'Luces
    cInsertarLuz.Value = False
    Call cInsertarLuz_Click

    cQuitarLuz.Value = False
    Call cQuitarLuz_Click

    'Particulas
    cQuitarParticula.Value = False
    Call cQuitarParticula_Click

    cInsertarParticula.Value = False
    Call cInsertarParticula_Click

End Sub

Private Sub mnuQuitarNPCs_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Quitar_NPCs(False)

End Sub

'Private Sub mnuQuitarNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'Call modEdicion.Quitar_NPCs(True)
'End Sub

Private Sub mnuQuitarObjetos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Quitar_Objetos

End Sub

Private Sub mnuQuitarSuperficieBordes_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Quitar_Bordes

End Sub

Private Sub mnuQuitarSuperficieDeCapa_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Quitar_Capa(cCapas.Text)

End Sub

Private Sub mnuQuitarTODO_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Borrar_Mapa

End Sub

Private Sub mnuQuitarTranslados_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 16/10/06
    '*************************************************
    Call modEdicion.Quitar_Translados

End Sub

Private Sub mnuQuitarTriggers_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Quitar_Triggers

End Sub

Private Sub mnuReAbrirMapa_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    On Error GoTo ErrHandler

    If FileExist(Dialog.FileName, vbArchive) = False Then Exit Sub
    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            modMapIO.GuardarMapa Dialog.FileName

        End If

    End If

    Call modMapIO.NuevoMapa
    modMapIO.AbrirMapa Dialog.FileName, MapData
    DoEvents
    mnuReAbrirMapa.Enabled = True
    EngineRun = True
    Exit Sub
ErrHandler:

End Sub

Private Sub mnuRealizarOperacion_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    Call modEdicion.Deshacer_Add("Realizar Operación en Selección")
    Call AccionSeleccion

End Sub

Private Sub mnuSalir_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Unload Me

End Sub

Private Sub mnuSuperficie_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Dim i As Byte

    For i = 0 To 6

        If i <> 0 Then
            frmMain.SelectPanel(i).Value = False
            Call VerFuncion(i, False)

        End If

    Next
    modPaneles.VerFuncion 0, True

End Sub

Private Sub mnuTranslados_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Dim i As Byte

    For i = 0 To 6

        If i <> 1 Then
            frmMain.SelectPanel(i).Value = False
            Call VerFuncion(i, False)

        End If

    Next
    modPaneles.VerFuncion 1, True

End Sub

Private Sub mnuTriggers_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Dim i As Byte

    For i = 0 To 6

        If i <> 6 Then
            frmMain.SelectPanel(i).Value = False
            Call VerFuncion(i, False)

        End If

    Next
    modPaneles.VerFuncion 6, True

End Sub

Private Sub mnuUtilizarDeshacer_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 16/10/06
    '*************************************************
    mnuUtilizarDeshacer.Checked = (mnuUtilizarDeshacer.Checked = False)

End Sub

Private Sub mnuVerAutomatico_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    mnuVerAutomatico.Checked = (mnuVerAutomatico.Checked = False)

End Sub

Private Sub mnuVerBloqueos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    cVerBloqueos.Value = (cVerBloqueos.Value = False)
    mnuVerBloqueos.Checked = cVerBloqueos.Value

End Sub

Private Sub mnuVerCapa1_Click()
    mnuVerCapa1.Checked = (mnuVerCapa1.Checked = False)

End Sub

Private Sub mnuVerCapa2_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    mnuVerCapa2.Checked = (mnuVerCapa2.Checked = False)

End Sub

Private Sub mnuVerCapa3_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    mnuVerCapa3.Checked = (mnuVerCapa3.Checked = False)

End Sub

Private Sub mnuVerCapa4_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    mnuVerCapa4.Checked = (mnuVerCapa4.Checked = False)

End Sub

Private Sub mnuVerGrilla_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 24/11/08
    '*************************************************
    VerGrilla = (VerGrilla = False)
    mnuVerGrilla.Checked = VerGrilla

End Sub

Private Sub mnuVerNPCs_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 26/05/06
    '*************************************************
    mnuVerNPCs.Checked = (mnuVerNPCs.Checked = False)

End Sub

Private Sub mnuVerObjetos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 26/05/06
    '*************************************************
    mnuVerObjetos.Checked = (mnuVerObjetos.Checked = False)

End Sub

Private Sub mnuVerTranslados_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 26/05/06
    '*************************************************
    mnuVerTranslados.Checked = (mnuVerTranslados.Checked = False)

End Sub

Private Sub mnuVerTriggers_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    cVerTriggers.Value = (cVerTriggers.Value = False)
    mnuVerTriggers.Checked = cVerTriggers.Value

End Sub

Private Sub picRadar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    If X < MinXBorder Then X = 11
    If X > MaxXBorder Then X = 89
    If Y < MinYBorder Then Y = 10
    If Y > MaxYBorder Then Y = 92

    UserPos.X = X
    UserPos.Y = Y
    bRefreshRadar = True

End Sub

Private Sub picRadar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    MiRadarX = X
    MiRadarY = Y

End Sub

Private Sub MainViewPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06 - GS
    'Last modified: 20/11/07 - Loopzer
    '*************************************************

    Dim tX As Integer

    Dim tY As Integer

    If Not MapaCargado Then Exit Sub

    ConvertCPtoTP 0, 0, X, Y, tX, tY

    'If Shift = 1 And Button = 2 Then PegarSeleccion tX, tY: Exit Sub
    If Shift = 1 And Button = 1 Then
        Seleccionando = True
        SeleccionIX = tX '+ UserPos.X
        SeleccionIY = tY '+ UserPos.Y
    Else
        ClickEdit Button, tX, tY

    End If

End Sub

Private Sub MainViewpic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06 - GS
    '*************************************************

    Dim tX As Integer

    Dim tY As Integer

    'Make sure map is loaded
    If Not MapaCargado Then Exit Sub
    HotKeysAllow = True

    ConvertCPtoTP 0, 0, X, Y, tX, tY

    POSX.Caption = "X: " & tX & " - Y: " & tY

    If tX < 10 Or tY < 10 Or tX > 90 Or tY > 90 Then
        POSX.ForeColor = vbRed
    Else
        POSX.ForeColor = vbWhite

    End If

    If Shift = 1 And Button = 1 Then
        Seleccionando = True
        SeleccionFX = tX '+ TileX
        SeleccionFY = tY '+ TileY
    Else
        ClickEdit Button, tX, tY

    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 24/11/08
    '*************************************************

    ' Guardar configuración
    WriteVar inipath & "WorldEditor.ini", "CONFIGURACION", "GuardarConfig", IIf(frmMain.mnuGuardarUltimaConfig.Checked = True, "1", "0")

    If frmMain.mnuGuardarUltimaConfig.Checked = True Then
        WriteVar inipath & "WorldEditor.ini", "PATH", "UltimoMapa", Dialog.FileName
        WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "ControlAutomatico", IIf(frmMain.mnuVerAutomatico.Checked = True, "1", "0")
        WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "Capa2", IIf(frmMain.mnuVerCapa2.Checked = True, "1", "0")
        WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "Capa3", IIf(frmMain.mnuVerCapa3.Checked = True, "1", "0")
        WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "Capa4", IIf(frmMain.mnuVerCapa4.Checked = True, "1", "0")
        WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "Translados", IIf(frmMain.mnuVerTranslados.Checked = True, "1", "0")
        WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "Objetos", IIf(frmMain.mnuVerObjetos.Checked = True, "1", "0")
        WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "NPCs", IIf(frmMain.mnuVerNPCs.Checked = True, "1", "0")
        WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "Triggers", IIf(frmMain.mnuVerTriggers.Checked = True, "1", "0")
        WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "Grilla", IIf(frmMain.mnuVerGrilla.Checked = True, "1", "0")
        WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "Bloqueos", IIf(frmMain.mnuVerBloqueos.Checked = True, "1", "0")
        WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "LastPos", UserPos.X & "-" & UserPos.Y
        WriteVar inipath & "WorldEditor.ini", "CONFIGURACION", "UtilizarDeshacer", IIf(frmMain.mnuUtilizarDeshacer.Checked = True, "1", "0")
        WriteVar inipath & "WorldEditor.ini", "CONFIGURACION", "AutoCapturarTrans", IIf(frmMain.mnuAutoCapturarTranslados.Checked = True, "1", "0")
        WriteVar inipath & "WorldEditor.ini", "CONFIGURACION", "AutoCapturarSup", IIf(frmMain.mnuAutoCapturarSuperficie.Checked = True, "1", "0")
        WriteVar inipath & "WorldEditor.ini", "CONFIGURACION", "ObjTranslado", Val(Cfg_TrOBJ)

    End If

    'Allow MainLoop to close program
    If prgRun = True Then
        prgRun = False
        Cancel = 1

    End If

End Sub

Private Sub SelectPanel_Click(index As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Dim i As Byte

    For i = 0 To 8

        If i <> index Then
            SelectPanel(i).Value = False
            Call VerFuncion(i, False)

        End If

    Next

    If mnuAutoQuitarFunciones.Checked = True Then Call mnuQuitarFunciones_Click
    Call VerFuncion(index, SelectPanel(index).Value)

End Sub

Private Sub TimAutoGuardarMapa_Timer()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If mnuAutoGuardarMapas.Checked = True Then
        bAutoGuardarMapaCount = bAutoGuardarMapaCount + 1

        If bAutoGuardarMapaCount >= bAutoGuardarMapa Then
            If MapInfo.Changed = 1 Then ' Solo guardo si el mapa esta modificado
                modMapIO.GuardarMapa Dialog.FileName

            End If

            bAutoGuardarMapaCount = 0

        End If

    End If

End Sub

Public Sub ObtenerNombreArchivo(ByVal Guardar As Boolean)
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************

    With Dialog
        .Filter = "Mapas de Argentum Online (*.map)|*.map|Mapas de IAOClon (*.csm)|*.csm"

        If Guardar Then
            .DialogTitle = "Guardar"
            .DefaultExt = ".txt"
            .FileName = vbNullString
            .flags = cdlOFNPathMustExist
            .ShowSave
        Else
            .DialogTitle = "Cargar"
            .FileName = vbNullString
            .flags = cdlOFNFileMustExist
            .ShowOpen

        End If

    End With

End Sub
