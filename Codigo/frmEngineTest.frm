VERSION 5.00
Begin VB.Form frmEngineTest 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Engine Test"
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame 
      Caption         =   "Montanas"
      Height          =   2295
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.Frame Frame 
         Caption         =   "Map Options"
         Height          =   1455
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1815
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   840
            TabIndex        =   10
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   840
            TabIndex        =   9
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "X"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Y"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   840
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Mountain Options"
         Height          =   1455
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   1815
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   840
            TabIndex        =   5
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   840
            TabIndex        =   4
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Radio"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Altura"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.CommandButton Command 
         Caption         =   "Crear"
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Reset"
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   1800
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmEngineTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command_Click()
    Map_CreateMontanita Text1.Text, Text2.Text, Text3.Text, Text4.Text

End Sub

Private Sub Command1_Click()
    Map_ResetMontanita

End Sub
