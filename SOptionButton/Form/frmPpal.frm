VERSION 5.00
Begin VB.Form frmPpal 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo SOptionButton"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6165
   Icon            =   "frmPpal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin OptionBt.SOptionButton SOptionButton10 
      Height          =   465
      Left            =   1785
      TabIndex        =   14
      Top             =   4200
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   820
      BackColor       =   16777215
      FocusColor      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      RoundColor      =   8812135
   End
   Begin VB.Frame Fram 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Align Right"
      Height          =   1050
      Index           =   1
      Left            =   3660
      TabIndex        =   10
      Top             =   2100
      Width           =   2385
      Begin OptionBt.SOptionButton SOptionButton9 
         Height          =   225
         Left            =   165
         TabIndex        =   11
         Top             =   345
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   397
         Alignment       =   1
         BackColor       =   16777215
         BorderColor     =   16744576
         Enabled         =   0   'False
         FocusColor      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         RoundColor      =   16744576
      End
      Begin OptionBt.SOptionButton SOptionButton11 
         Height          =   255
         Left            =   165
         TabIndex        =   12
         Top             =   630
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   450
         Alignment       =   1
         BackColor       =   16777215
         BorderColor     =   4110841
         Caption         =   "Colombia - ????"
         FocusColor      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         RoundColor      =   255
         Value           =   -1  'True
      End
   End
   Begin OptionBt.SOptionButton SOptionButton1 
      Height          =   225
      Left            =   195
      TabIndex        =   0
      Top             =   120
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   397
      BackColor       =   16777215
      BorderColor     =   8421376
      FocusColor      =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      RoundColor      =   8421376
   End
   Begin VB.Frame Fram 
      BackColor       =   &H00FFFFFF&
      Height          =   885
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   3405
      Begin OptionBt.SOptionButton SOptionButton5 
         Height          =   225
         Left            =   75
         TabIndex        =   5
         Top             =   225
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   397
         BackColor       =   16777215
         BorderColor     =   16744703
         FocusColor      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         RoundColor      =   16711935
      End
      Begin OptionBt.SOptionButton SOptionButton6 
         Height          =   225
         Left            =   1725
         TabIndex        =   6
         Top             =   225
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   397
         BackColor       =   16777215
         BorderColor     =   16711680
         FocusColor      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
         RoundColor      =   8388608
      End
      Begin OptionBt.SOptionButton SOptionButton7 
         Height          =   225
         Left            =   75
         TabIndex        =   7
         Top             =   525
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   397
         BackColor       =   16777215
         BorderColor     =   4210752
         FocusColor      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         RoundColor      =   0
      End
      Begin OptionBt.SOptionButton SOptionButton8 
         Height          =   225
         Left            =   1725
         TabIndex        =   8
         Top             =   525
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   397
         BackColor       =   16777215
         BorderColor     =   12686741
         Enabled         =   0   'False
         FocusColor      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         RoundColor      =   8934998
         Value           =   -1  'True
      End
   End
   Begin OptionBt.SOptionButton SOptionButton2 
      Height          =   225
      Left            =   2130
      TabIndex        =   1
      Top             =   120
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   397
      BackColor       =   16777215
      BorderColor     =   12583104
      FocusColor      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      RoundColor      =   8388736
   End
   Begin OptionBt.SOptionButton SOptionButton3 
      Height          =   225
      Left            =   195
      TabIndex        =   2
      Top             =   420
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   397
      BackColor       =   16777215
      BorderColor     =   255
      FocusColor      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      RoundColor      =   16576
   End
   Begin OptionBt.SOptionButton SOptionButton4 
      Height          =   225
      Left            =   2130
      TabIndex        =   3
      Top             =   420
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   397
      BackColor       =   16777215
      BorderColor     =   8421631
      FocusColor      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      RoundColor      =   255
   End
   Begin VB.Label lblComment 
      BackStyle       =   0  'Transparent
      Caption         =   "This is a very simple control of using, it is not necessary to explain to it."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C56A31&
      Height          =   660
      Index           =   0
      Left            =   3645
      TabIndex        =   13
      Top             =   3435
      Width           =   2415
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblComment 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmPpal.frx":058A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C56A31&
      Height          =   1635
      Index           =   4
      Left            =   3735
      TabIndex        =   9
      Top             =   165
      Width           =   2415
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   2445
      Left            =   0
      Picture         =   "frmPpal.frx":0652
      Top             =   1665
      Width           =   3570
   End
End
Attribute VB_Name = "frmPpal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
 SOptionButton11.Caption = "Colombia - "
 SOptionButton11.Caption = SOptionButton11.Caption & ChrW$(31252) & ChrW$(31175) & ChrW$(31215) & ChrW$(31188)
End Sub
