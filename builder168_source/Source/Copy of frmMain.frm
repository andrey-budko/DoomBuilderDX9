VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Doom Builder"
   ClientHeight    =   7710
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11175
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   514
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   745
   Begin VB.Timer tmr3DRedraw 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2280
      Top             =   2790
   End
   Begin VB.Timer tmrTerminate 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1770
      Top             =   2790
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1125
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   157
      TabStop         =   0   'False
      Top             =   2115
      Visible         =   0   'False
      Width           =   750
   End
   Begin MSComctlLib.ImageList imglstToolbar 
      Left            =   8430
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2294
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":282E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3362
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3E96
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4430
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":49CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F64
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":54FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5E32
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":61CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6766
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6D00
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":729A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7834
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7DCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8368
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8902
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8E9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9436
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":99D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9F6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A504
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AA9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B038
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B5D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BB6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C106
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C6A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CC3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D1D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D76E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMap 
      BackColor       =   &H8000000C&
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      HasDC           =   0   'False
      Height          =   2145
      Left            =   5025
      ScaleHeight     =   139
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   263
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3555
      Width           =   4005
   End
   Begin VB.PictureBox picSBar 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      HasDC           =   0   'False
      Height          =   5460
      Left            =   9150
      ScaleHeight     =   364
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   135
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   2025
      Begin VB.Frame fraSThingPreview 
         Caption         =   " Preview "
         Height          =   1590
         Left            =   0
         TabIndex        =   154
         Top             =   1950
         Visible         =   0   'False
         Width           =   1980
         Begin VB.PictureBox picSThing 
            BackColor       =   &H8000000C&
            CausesValidation=   0   'False
            ClipControls    =   0   'False
            HasDC           =   0   'False
            Height          =   1020
            Left            =   480
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   155
            TabStop         =   0   'False
            ToolTipText     =   "Sector Floor Texture"
            Top             =   240
            Width           =   1020
            Begin VB.Image imgSThing 
               Height          =   960
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   960
            End
         End
         Begin VB.Label lblSThing 
            Alignment       =   2  'Center
            Caption         =   "STARTAN3"
            Height          =   210
            Left            =   480
            TabIndex        =   156
            Top             =   1260
            UseMnemonic     =   0   'False
            Width           =   1020
         End
      End
      Begin VB.CommandButton cmdToggleSBar 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   9.75
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   147
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   300
         Width           =   270
      End
      Begin VB.Frame fraSThing 
         Caption         =   " Thing 0 "
         Height          =   1905
         Left            =   0
         TabIndex        =   110
         Top             =   0
         Visible         =   0   'False
         Width           =   1980
         Begin VB.Label lblSThingTag 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   900
            TabIndex        =   144
            Top             =   1260
            Width           =   90
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Tag:"
            Height          =   210
            Index           =   43
            Left            =   495
            TabIndex        =   143
            Top             =   1260
            Width           =   315
         End
         Begin VB.Label lblSThingAction 
            Caption         =   "0 - Normal"
            Height          =   210
            Left            =   900
            TabIndex        =   142
            Top             =   540
            UseMnemonic     =   0   'False
            Width           =   3105
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Action:"
            Height          =   210
            Index           =   42
            Left            =   300
            TabIndex        =   141
            Top             =   540
            Width           =   510
         End
         Begin VB.Label lblSThingType 
            Caption         =   "0 - Normal"
            Height          =   210
            Left            =   900
            TabIndex        =   118
            Top             =   300
            UseMnemonic     =   0   'False
            Width           =   1005
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Angle:"
            Height          =   210
            Index           =   25
            Left            =   345
            TabIndex        =   117
            Top             =   780
            UseMnemonic     =   0   'False
            Width           =   465
         End
         Begin VB.Label lblSThingAngle 
            Caption         =   "0"
            Height          =   210
            Left            =   900
            TabIndex        =   116
            Top             =   780
            UseMnemonic     =   0   'False
            Width           =   1005
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Flags:"
            Height          =   210
            Index           =   31
            Left            =   375
            TabIndex        =   115
            Top             =   1020
            Width           =   435
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Type:"
            Height          =   210
            Index           =   21
            Left            =   405
            TabIndex        =   114
            Top             =   300
            UseMnemonic     =   0   'False
            Width           =   405
         End
         Begin VB.Label lblSThingFlags 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   900
            TabIndex        =   113
            Top             =   1020
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "X,Y:"
            Height          =   210
            Index           =   17
            Left            =   495
            TabIndex        =   112
            Top             =   1500
            UseMnemonic     =   0   'False
            Width           =   315
         End
         Begin VB.Label lblSThingXY 
            AutoSize        =   -1  'True
            Caption         =   "0, 0"
            Height          =   210
            Left            =   915
            TabIndex        =   111
            Top             =   1500
            UseMnemonic     =   0   'False
            Width           =   270
         End
      End
      Begin VB.Frame fraSBackSidedef 
         Caption         =   " Back Side "
         Height          =   4170
         Left            =   0
         TabIndex        =   84
         Top             =   6345
         Width           =   1980
         Begin VB.PictureBox picSS2Lower 
            BackColor       =   &H8000000C&
            CausesValidation=   0   'False
            ClipControls    =   0   'False
            HasDC           =   0   'False
            Height          =   1020
            Left            =   480
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   87
            TabStop         =   0   'False
            ToolTipText     =   "Back Side Lower Texture"
            Top             =   2805
            Width           =   1020
            Begin VB.Image imgSS2Lower 
               Height          =   960
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   960
            End
         End
         Begin VB.PictureBox picSS2Middle 
            BackColor       =   &H8000000C&
            CausesValidation=   0   'False
            ClipControls    =   0   'False
            HasDC           =   0   'False
            Height          =   1020
            Left            =   480
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   86
            TabStop         =   0   'False
            ToolTipText     =   "Back Side Middle Texture"
            Top             =   1530
            Width           =   1020
            Begin VB.Image imgSS2Middle 
               Height          =   960
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   960
            End
         End
         Begin VB.PictureBox picSS2Upper 
            BackColor       =   &H8000000C&
            CausesValidation=   0   'False
            ClipControls    =   0   'False
            HasDC           =   0   'False
            Height          =   1020
            Left            =   480
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   85
            TabStop         =   0   'False
            ToolTipText     =   "Back Side Upper Texture"
            Top             =   240
            Width           =   1020
            Begin VB.Image imgSS2Upper 
               Height          =   960
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   960
            End
         End
         Begin VB.Label lblSS2Lower 
            Alignment       =   2  'Center
            Caption         =   "STARTAN3"
            Height          =   210
            Left            =   480
            TabIndex        =   90
            Top             =   3825
            UseMnemonic     =   0   'False
            Width           =   1020
         End
         Begin VB.Label lblSS2Middle 
            Alignment       =   2  'Center
            Caption         =   "STARTAN3"
            Height          =   210
            Left            =   480
            TabIndex        =   89
            Top             =   2550
            UseMnemonic     =   0   'False
            Width           =   1020
         End
         Begin VB.Label lblSS2Upper 
            Alignment       =   2  'Center
            Caption         =   "STARTAN3"
            Height          =   210
            Left            =   480
            TabIndex        =   88
            Top             =   1260
            UseMnemonic     =   0   'False
            Width           =   1020
         End
      End
      Begin VB.Frame fraSLinedef 
         Caption         =   " Linedef 0 "
         Height          =   2100
         Left            =   0
         TabIndex        =   66
         Top             =   0
         Visible         =   0   'False
         Width           =   1980
         Begin VB.Label lblSS2Sector 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   900
            TabIndex        =   127
            Top             =   1770
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "B Sector:"
            Height          =   210
            Index           =   12
            Left            =   135
            TabIndex        =   126
            Top             =   1770
            Width           =   675
         End
         Begin VB.Label lblSS1Sector 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   900
            TabIndex        =   125
            Top             =   1530
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "F Sector:"
            Height          =   210
            Index           =   32
            Left            =   150
            TabIndex        =   124
            Top             =   1530
            Width           =   660
         End
         Begin VB.Label lblSLinedefTag 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   900
            TabIndex        =   76
            Top             =   795
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tag:"
            Height          =   210
            Index           =   27
            Left            =   480
            TabIndex        =   75
            Top             =   795
            UseMnemonic     =   0   'False
            Width           =   315
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Action:"
            Height          =   210
            Index           =   11
            Left            =   300
            TabIndex        =   74
            Top             =   300
            UseMnemonic     =   0   'False
            Width           =   510
         End
         Begin VB.Label lblSLinedefLength 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   900
            TabIndex        =   73
            Top             =   540
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Length:"
            Height          =   210
            Index           =   18
            Left            =   270
            TabIndex        =   72
            Top             =   540
            UseMnemonic     =   0   'False
            Width           =   540
         End
         Begin VB.Label lblSLinedefType 
            Caption         =   "0 - Normal"
            Height          =   210
            Left            =   900
            TabIndex        =   71
            Top             =   300
            UseMnemonic     =   0   'False
            Width           =   1000
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "F Height:"
            Height          =   210
            Index           =   15
            Left            =   180
            TabIndex        =   70
            Top             =   1050
            Width           =   630
         End
         Begin VB.Label lblSS1Height 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   900
            TabIndex        =   69
            Top             =   1050
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "B Height:"
            Height          =   210
            Index           =   13
            Left            =   165
            TabIndex        =   68
            Top             =   1290
            Width           =   645
         End
         Begin VB.Label lblSS2height 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   900
            TabIndex        =   67
            Top             =   1290
            UseMnemonic     =   0   'False
            Width           =   90
         End
      End
      Begin VB.Frame fraSVertex 
         Caption         =   " Vertex 0 "
         Height          =   765
         Left            =   0
         TabIndex        =   63
         Top             =   0
         Width           =   1980
         Begin VB.Label lblSVertexXY 
            AutoSize        =   -1  'True
            Caption         =   "0, 0"
            Height          =   210
            Left            =   930
            TabIndex        =   65
            Top             =   330
            UseMnemonic     =   0   'False
            Width           =   270
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "X,Y:"
            Height          =   210
            Index           =   14
            Left            =   510
            TabIndex        =   64
            Top             =   330
            UseMnemonic     =   0   'False
            Width           =   315
         End
      End
      Begin VB.Frame fraSSectorFloor 
         Caption         =   " Floor "
         Height          =   1590
         Left            =   0
         TabIndex        =   107
         Top             =   3525
         Visible         =   0   'False
         Width           =   1980
         Begin VB.PictureBox picSFloor 
            BackColor       =   &H8000000C&
            CausesValidation=   0   'False
            ClipControls    =   0   'False
            HasDC           =   0   'False
            Height          =   1020
            Left            =   480
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   108
            TabStop         =   0   'False
            ToolTipText     =   "Sector Ceiling Texture"
            Top             =   240
            Width           =   1020
            Begin VB.Image imgSFloor 
               Height          =   960
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   960
            End
         End
         Begin VB.Label lblSFloor 
            Alignment       =   2  'Center
            Caption         =   "STARTAN3"
            Height          =   210
            Left            =   480
            TabIndex        =   109
            Top             =   1260
            UseMnemonic     =   0   'False
            Width           =   1020
         End
      End
      Begin VB.Frame fraSSectorCeiling 
         Caption         =   " Ceiling "
         Height          =   1590
         Left            =   0
         TabIndex        =   104
         Top             =   1905
         Visible         =   0   'False
         Width           =   1980
         Begin VB.PictureBox picSCeiling 
            BackColor       =   &H8000000C&
            CausesValidation=   0   'False
            ClipControls    =   0   'False
            HasDC           =   0   'False
            Height          =   1020
            Left            =   480
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   105
            TabStop         =   0   'False
            ToolTipText     =   "Sector Floor Texture"
            Top             =   240
            Width           =   1020
            Begin VB.Image imgSCeiling 
               Height          =   960
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   960
            End
         End
         Begin VB.Label lblSCeiling 
            Alignment       =   2  'Center
            Caption         =   "STARTAN3"
            Height          =   210
            Left            =   480
            TabIndex        =   106
            Top             =   1260
            UseMnemonic     =   0   'False
            Width           =   1020
         End
      End
      Begin VB.Frame fraSSector 
         Caption         =   " Sector 0 "
         Height          =   1875
         Left            =   0
         TabIndex        =   91
         Top             =   0
         Visible         =   0   'False
         Width           =   1980
         Begin VB.Label lblSSectorType 
            Caption         =   "0 - Normal"
            Height          =   210
            Left            =   900
            TabIndex        =   103
            Top             =   300
            UseMnemonic     =   0   'False
            Width           =   1005
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ceiling:"
            Height          =   210
            Index           =   22
            Left            =   300
            TabIndex        =   102
            Top             =   540
            UseMnemonic     =   0   'False
            Width           =   510
         End
         Begin VB.Label lblSSectorCeiling 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   900
            TabIndex        =   101
            Top             =   540
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Floor:"
            Height          =   210
            Index           =   23
            Left            =   405
            TabIndex        =   100
            Top             =   780
            Width           =   405
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Effect:"
            Height          =   210
            Index           =   30
            Left            =   330
            TabIndex        =   99
            Top             =   300
            UseMnemonic     =   0   'False
            Width           =   480
         End
         Begin VB.Label lblSSectorFloor 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   900
            TabIndex        =   98
            Top             =   780
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tag:"
            Height          =   210
            Index           =   26
            Left            =   480
            TabIndex        =   97
            Top             =   1020
            UseMnemonic     =   0   'False
            Width           =   315
         End
         Begin VB.Label lblSSectorTag 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   900
            TabIndex        =   96
            Top             =   1020
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Height:"
            Height          =   210
            Index           =   19
            Left            =   315
            TabIndex        =   95
            Top             =   1275
            UseMnemonic     =   0   'False
            Width           =   495
         End
         Begin VB.Label lblSSectorHeight 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   900
            TabIndex        =   94
            Top             =   1275
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Light:"
            Height          =   210
            Index           =   16
            Left            =   420
            TabIndex        =   93
            Top             =   1515
            UseMnemonic     =   0   'False
            Width           =   390
         End
         Begin VB.Label lblSSectorLight 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   900
            TabIndex        =   92
            Top             =   1515
            UseMnemonic     =   0   'False
            Width           =   90
         End
      End
      Begin VB.Frame fraSFrontSidedef 
         Caption         =   " Front Side "
         Height          =   4170
         Left            =   0
         TabIndex        =   77
         Top             =   2145
         Width           =   1980
         Begin VB.PictureBox picSS1Upper 
            BackColor       =   &H8000000C&
            CausesValidation=   0   'False
            ClipControls    =   0   'False
            HasDC           =   0   'False
            Height          =   1020
            Left            =   480
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   80
            TabStop         =   0   'False
            ToolTipText     =   "Front Side Upper Texture"
            Top             =   240
            Width           =   1020
            Begin VB.Image imgSS1Upper 
               Height          =   960
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   960
            End
         End
         Begin VB.PictureBox picSS1Middle 
            BackColor       =   &H8000000C&
            CausesValidation=   0   'False
            ClipControls    =   0   'False
            HasDC           =   0   'False
            Height          =   1020
            Left            =   480
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   79
            TabStop         =   0   'False
            ToolTipText     =   "Front Side Middle Texture"
            Top             =   1530
            Width           =   1020
            Begin VB.Image imgSS1Middle 
               Height          =   960
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   960
            End
         End
         Begin VB.PictureBox picSS1Lower 
            BackColor       =   &H8000000C&
            CausesValidation=   0   'False
            ClipControls    =   0   'False
            HasDC           =   0   'False
            Height          =   1020
            Left            =   480
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   78
            TabStop         =   0   'False
            ToolTipText     =   "Front Side Lower Texture"
            Top             =   2820
            Width           =   1020
            Begin VB.Image imgSS1Lower 
               Height          =   960
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   960
            End
         End
         Begin VB.Label lblSS1Upper 
            Alignment       =   2  'Center
            Caption         =   "STARTAN3"
            Height          =   210
            Left            =   480
            TabIndex        =   83
            Top             =   1260
            UseMnemonic     =   0   'False
            Width           =   1020
         End
         Begin VB.Label lblSS1Middle 
            Alignment       =   2  'Center
            Caption         =   "STARTAN3"
            Height          =   210
            Left            =   480
            TabIndex        =   82
            Top             =   2550
            UseMnemonic     =   0   'False
            Width           =   1020
         End
         Begin VB.Label lblSS1Lower 
            Alignment       =   2  'Center
            Caption         =   "STARTAN3"
            Height          =   210
            Left            =   480
            TabIndex        =   81
            Top             =   3840
            UseMnemonic     =   0   'False
            Width           =   1020
         End
      End
   End
   Begin VB.Timer tmrAutoScroll 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1260
      Top             =   2790
   End
   Begin VB.PictureBox picTexture 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      Height          =   375
      Left            =   240
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2115
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox picThings 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FontTransparent =   0   'False
      HasDC           =   0   'False
      Height          =   285
      Index           =   3
      Left            =   225
      Picture         =   "frmMain.frx":DD08
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.PictureBox picThings 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FontTransparent =   0   'False
      HasDC           =   0   'False
      Height          =   195
      Index           =   2
      Left            =   225
      Picture         =   "frmMain.frx":EF8A
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   132
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   1230
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.PictureBox picThings 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FontTransparent =   0   'False
      HasDC           =   0   'False
      Height          =   105
      Index           =   1
      Left            =   225
      Picture         =   "frmMain.frx":FA80
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   72
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   990
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox picThings 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FontTransparent =   0   'False
      HasDC           =   0   'False
      Height          =   45
      Index           =   0
      Left            =   225
      Picture         =   "frmMain.frx":100BA
      ScaleHeight     =   3
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Timer tmrMouseTimeout 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Top             =   2790
   End
   Begin VB.Timer tmrMouseOutside 
      Interval        =   107
      Left            =   750
      Top             =   2790
   End
   Begin VB.PictureBox picNumbers 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FontTransparent =   0   'False
      HasDC           =   0   'False
      Height          =   135
      Left            =   225
      Picture         =   "frmMain.frx":1055C
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   585
      Visible         =   0   'False
      Width           =   900
   End
   Begin MSComctlLib.Toolbar tlbToolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imglstToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   38
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FileNewMap"
            Object.ToolTipText     =   "New Map"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FileOpenMap"
            Object.ToolTipText     =   "Open Map"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FileSaveMap"
            Object.ToolTipText     =   "Save Map"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   14
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ModeMove"
            Object.ToolTipText     =   "Move Mode"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ModeVertices"
            Object.ToolTipText     =   "Vertices Mode"
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ModeLines"
            Object.ToolTipText     =   "Lines Mode"
            ImageIndex      =   6
            Style           =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ModeSectors"
            Object.ToolTipText     =   "Sectors Mode"
            ImageIndex      =   7
            Style           =   2
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ModeThings"
            Object.ToolTipText     =   "Things Mode"
            ImageIndex      =   8
            Style           =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Mode3D"
            Object.ToolTipText     =   "3D Mode"
            ImageIndex      =   9
            Style           =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   14
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FileBuild"
            Object.ToolTipText     =   "Build nodes"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FileTest"
            Object.ToolTipText     =   "Test Map"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   14
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EditUndo"
            Object.ToolTipText     =   "Undo"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EditRedo"
            Object.ToolTipText     =   "Redo"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   14
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EditFlipH"
            Object.ToolTipText     =   "Flip Selection Horizontal"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EditFlipV"
            Object.ToolTipText     =   "Flip Selection Vertical"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EditRotate"
            Object.ToolTipText     =   "Rotate Selection"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EditResize"
            Object.ToolTipText     =   "Resize Selection"
            ImageIndex      =   29
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   14
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EditGrid"
            Object.ToolTipText     =   "Grid Settings"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EditSnap"
            Object.ToolTipText     =   "Snap To Grid"
            ImageIndex      =   18
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EditStitch"
            Object.ToolTipText     =   "Stitch Vertices"
            ImageIndex      =   19
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EditCenterView"
            Object.ToolTipText     =   "Center View"
            ImageIndex      =   33
            Object.Width           =   14
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   14
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PrefabsInsert"
            Object.ToolTipText     =   "Insert Prefab from File"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PrefabsInsertPrevious"
            Object.ToolTipText     =   "Insert Previous Prefab"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   14
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "LinesFlip"
            Object.ToolTipText     =   "Flip Linedefs"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "LinesCurve"
            Object.ToolTipText     =   "Curve Linedefs"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SectorsJoin"
            Object.ToolTipText     =   "Join Sectors"
            ImageIndex      =   25
         EndProperty
         BeginProperty Button34 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SectorsMerge"
            Object.ToolTipText     =   "Merge Sectors"
            ImageIndex      =   26
         EndProperty
         BeginProperty Button35 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SectorsGradientBrightness"
            Object.ToolTipText     =   "Gradient Brightness"
            ImageIndex      =   30
         EndProperty
         BeginProperty Button36 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SectorsGradientFloors"
            Object.ToolTipText     =   "Gradient Floors"
            ImageIndex      =   31
         EndProperty
         BeginProperty Button37 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SectorsGradientCeilings"
            Object.ToolTipText     =   "Gradient Ceilings"
            ImageIndex      =   32
         EndProperty
         BeginProperty Button38 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ThingsFilter"
            Object.ToolTipText     =   "Things Filter"
            ImageIndex      =   27
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   7410
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   11
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2381
            MinWidth        =   2381
            Text            =   "0 vertices"
            TextSave        =   "0 vertices"
            Key             =   "numvertexes"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2381
            MinWidth        =   2381
            Text            =   "0 linedefs"
            TextSave        =   "0 linedefs"
            Key             =   "numlinedefs"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2381
            MinWidth        =   2381
            Text            =   "0 sidedefs"
            TextSave        =   "0 sidedefs"
            Key             =   "numsidedefs"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2381
            MinWidth        =   2381
            Text            =   "0 sectors"
            TextSave        =   "0 sectors"
            Key             =   "numsectors"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2381
            MinWidth        =   2381
            Text            =   "0 things"
            TextSave        =   "0 things"
            Key             =   "numthings"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2381
            MinWidth        =   2381
            Text            =   "Grid: 64"
            TextSave        =   "Grid: 64"
            Key             =   "gridsize"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2381
            MinWidth        =   2381
            Text            =   "Snap: OFF"
            TextSave        =   "Snap: OFF"
            Key             =   "snapmode"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2381
            MinWidth        =   2381
            Text            =   "Stitch: OFF"
            TextSave        =   "Stitch: OFF"
            Key             =   "stitchmode"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2381
            MinWidth        =   2381
            Text            =   "Zoom: 100%"
            TextSave        =   "Zoom: 100%"
            Key             =   "viewzoom"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2381
            MinWidth        =   2381
            Key             =   "mousex"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2381
            MinWidth        =   2381
            Key             =   "mousey"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBar 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FontTransparent =   0   'False
      HasDC           =   0   'False
      Height          =   1590
      Left            =   0
      ScaleHeight     =   106
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   745
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5820
      Visible         =   0   'False
      Width           =   11175
      Begin VB.Frame fraThingPreview 
         Caption         =   " Sprite "
         Height          =   1590
         Left            =   4170
         TabIndex        =   151
         Top             =   0
         Visible         =   0   'False
         Width           =   1305
         Begin VB.PictureBox picThing 
            BackColor       =   &H8000000C&
            CausesValidation=   0   'False
            ClipControls    =   0   'False
            HasDC           =   0   'False
            Height          =   1020
            Left            =   150
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   152
            TabStop         =   0   'False
            ToolTipText     =   "Sector Floor Texture"
            Top             =   240
            Width           =   1020
            Begin VB.Image imgThing 
               Height          =   960
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   960
            End
         End
         Begin VB.Label lblThing 
            Alignment       =   2  'Center
            Caption         =   "STARTAN3"
            Height          =   210
            Left            =   150
            TabIndex        =   153
            Top             =   1260
            UseMnemonic     =   0   'False
            Width           =   1020
         End
      End
      Begin VB.Frame fraSectorCeiling 
         Caption         =   " Ceiling "
         Height          =   1590
         Left            =   4170
         TabIndex        =   148
         Top             =   0
         Visible         =   0   'False
         Width           =   1305
         Begin VB.PictureBox picCeiling 
            BackColor       =   &H8000000C&
            CausesValidation=   0   'False
            ClipControls    =   0   'False
            HasDC           =   0   'False
            Height          =   1020
            Left            =   150
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   149
            TabStop         =   0   'False
            ToolTipText     =   "Sector Floor Texture"
            Top             =   240
            Width           =   1020
            Begin VB.Image imgCeiling 
               Height          =   960
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   960
            End
         End
         Begin VB.Label lblCeiling 
            Alignment       =   2  'Center
            Caption         =   "STARTAN3"
            Height          =   210
            Left            =   150
            TabIndex        =   150
            Top             =   1260
            UseMnemonic     =   0   'False
            Width           =   1020
         End
      End
      Begin VB.CommandButton cmdToggleBar 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   9.75
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10605
         TabIndex        =   145
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   0
         Width           =   300
      End
      Begin VB.Frame fraThing 
         Caption         =   " Thing 0 "
         Height          =   1590
         Left            =   45
         TabIndex        =   53
         Top             =   0
         Visible         =   0   'False
         Width           =   4065
         Begin VB.Label lblThingAction 
            Caption         =   "0 - Normal"
            Height          =   210
            Left            =   900
            TabIndex        =   140
            Top             =   510
            UseMnemonic     =   0   'False
            Width           =   3105
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Action:"
            Height          =   210
            Index           =   41
            Left            =   300
            TabIndex        =   139
            Top             =   510
            Width           =   510
         End
         Begin VB.Label lblThingTag 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   2415
            TabIndex        =   138
            Top             =   750
            Width           =   90
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Tag:"
            Height          =   210
            Index           =   40
            Left            =   2010
            TabIndex        =   137
            Top             =   750
            Width           =   315
         End
         Begin VB.Label lblThingXY 
            AutoSize        =   -1  'True
            Caption         =   "0, 0"
            Height          =   210
            Left            =   2445
            TabIndex        =   61
            Top             =   990
            UseMnemonic     =   0   'False
            Width           =   270
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "X,Y:"
            Height          =   210
            Index           =   37
            Left            =   2025
            TabIndex        =   60
            Top             =   990
            UseMnemonic     =   0   'False
            Width           =   315
         End
         Begin VB.Label lblThingFlags 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   900
            TabIndex        =   59
            Top             =   990
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Type:"
            Height          =   210
            Index           =   34
            Left            =   405
            TabIndex        =   58
            Top             =   270
            UseMnemonic     =   0   'False
            Width           =   405
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Flags:"
            Height          =   210
            Index           =   36
            Left            =   375
            TabIndex        =   57
            Top             =   990
            Width           =   435
         End
         Begin VB.Label lblThingAngle 
            Caption         =   "0"
            Height          =   210
            Left            =   900
            TabIndex        =   56
            Top             =   750
            UseMnemonic     =   0   'False
            Width           =   615
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Angle:"
            Height          =   210
            Index           =   35
            Left            =   345
            TabIndex        =   55
            Top             =   750
            UseMnemonic     =   0   'False
            Width           =   465
         End
         Begin VB.Label lblThingType 
            Caption         =   "0 - Normal"
            Height          =   210
            Left            =   900
            TabIndex        =   54
            Top             =   270
            UseMnemonic     =   0   'False
            Width           =   3105
         End
      End
      Begin VB.Frame fraLinedef 
         Caption         =   " Linedef 0 "
         Height          =   1590
         Left            =   45
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   4065
         Begin VB.Label lblS2Y 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   3600
            TabIndex        =   136
            Top             =   750
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblS2X 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   3600
            TabIndex        =   135
            Top             =   510
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblS1Y 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   2520
            TabIndex        =   134
            Top             =   750
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblS1X 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   2520
            TabIndex        =   133
            Top             =   510
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Back X:"
            Height          =   210
            Index           =   24
            Left            =   2955
            TabIndex        =   132
            Top             =   510
            Width           =   555
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Front X:"
            Height          =   210
            Index           =   33
            Left            =   1860
            TabIndex        =   131
            Top             =   510
            Width           =   570
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Back Y:"
            Height          =   210
            Index           =   20
            Left            =   2955
            TabIndex        =   129
            Top             =   750
            Width           =   570
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Front Y:"
            Height          =   210
            Index           =   29
            Left            =   1860
            TabIndex        =   128
            Top             =   750
            Width           =   585
         End
         Begin VB.Label lblS2Sector 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   2880
            TabIndex        =   123
            Top             =   1005
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Back Sector:"
            Height          =   210
            Index           =   28
            Left            =   1860
            TabIndex        =   121
            Top             =   1005
            Width           =   930
         End
         Begin VB.Label lblS2height 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   2880
            TabIndex        =   52
            Top             =   1245
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblLinedefType 
            Caption         =   "0 - Normal"
            Height          =   210
            Left            =   1215
            TabIndex        =   29
            Top             =   270
            UseMnemonic     =   0   'False
            Width           =   2790
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Action:"
            Height          =   210
            Index           =   0
            Left            =   600
            TabIndex        =   26
            Top             =   270
            UseMnemonic     =   0   'False
            Width           =   510
         End
         Begin VB.Label lblS1Height 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   1215
            TabIndex        =   50
            Top             =   1245
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Length:"
            Height          =   210
            Index           =   1
            Left            =   570
            TabIndex        =   28
            Top             =   510
            UseMnemonic     =   0   'False
            Width           =   540
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tag:"
            Height          =   210
            Index           =   38
            Left            =   795
            TabIndex        =   25
            Top             =   750
            UseMnemonic     =   0   'False
            Width           =   315
         End
         Begin VB.Label lblLinedefLength 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   1215
            TabIndex        =   27
            Top             =   510
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblLinedefTag 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   1215
            TabIndex        =   24
            Top             =   750
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblS1Sector 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   1215
            TabIndex        =   122
            Top             =   1005
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Front Sector:"
            Height          =   210
            Index           =   2
            Left            =   165
            TabIndex        =   120
            Top             =   1005
            Width           =   945
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Front Height:"
            Height          =   210
            Index           =   9
            Left            =   195
            TabIndex        =   130
            Top             =   1245
            Width           =   915
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Back Height:"
            Height          =   210
            Index           =   10
            Left            =   1860
            TabIndex        =   51
            Top             =   1245
            Width           =   900
         End
      End
      Begin VB.Frame fraVertex 
         Caption         =   " Vertex 0 "
         Height          =   1590
         Left            =   45
         TabIndex        =   6
         Top             =   0
         Width           =   3705
         Begin VB.Label lblVertexXY 
            AutoSize        =   -1  'True
            Caption         =   "0, 0"
            Height          =   210
            Left            =   1455
            TabIndex        =   8
            Top             =   510
            UseMnemonic     =   0   'False
            Width           =   270
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Coordinates:"
            Height          =   210
            Index           =   3
            Left            =   360
            TabIndex        =   7
            Top             =   510
            UseMnemonic     =   0   'False
            Width           =   915
         End
      End
      Begin VB.Frame fraSector 
         Caption         =   " Sector 0 "
         Height          =   1590
         Left            =   45
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   4065
         Begin VB.Label lblSectorLight 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   2670
            TabIndex        =   49
            Top             =   1230
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Light:"
            Height          =   210
            Index           =   8
            Left            =   2175
            TabIndex        =   48
            Top             =   1230
            UseMnemonic     =   0   'False
            Width           =   390
         End
         Begin VB.Label lblSectorHeight 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   900
            TabIndex        =   47
            Top             =   1230
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Height:"
            Height          =   210
            Index           =   7
            Left            =   300
            TabIndex        =   46
            Top             =   1230
            UseMnemonic     =   0   'False
            Width           =   495
         End
         Begin VB.Label lblSectorTag 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   900
            TabIndex        =   38
            Top             =   990
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tag:"
            Height          =   210
            Index           =   6
            Left            =   480
            TabIndex        =   37
            Top             =   990
            UseMnemonic     =   0   'False
            Width           =   315
         End
         Begin VB.Label lblSectorFloor 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   900
            TabIndex        =   36
            Top             =   750
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Effect:"
            Height          =   210
            Index           =   39
            Left            =   330
            TabIndex        =   35
            Top             =   270
            UseMnemonic     =   0   'False
            Width           =   480
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Floor:"
            Height          =   210
            Index           =   5
            Left            =   405
            TabIndex        =   34
            Top             =   750
            Width           =   405
         End
         Begin VB.Label lblSectorCeiling 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   210
            Left            =   900
            TabIndex        =   33
            Top             =   510
            UseMnemonic     =   0   'False
            Width           =   90
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ceiling:"
            Height          =   210
            Index           =   4
            Left            =   300
            TabIndex        =   32
            Top             =   510
            UseMnemonic     =   0   'False
            Width           =   510
         End
         Begin VB.Label lblSectorType 
            Caption         =   "0 - Normal"
            Height          =   210
            Left            =   900
            TabIndex        =   31
            Top             =   270
            UseMnemonic     =   0   'False
            Width           =   3105
         End
      End
      Begin VB.Frame fraSectorFloor 
         Caption         =   " Floor "
         Height          =   1590
         Left            =   5535
         TabIndex        =   39
         Top             =   0
         Visible         =   0   'False
         Width           =   1305
         Begin VB.PictureBox picFloor 
            BackColor       =   &H8000000C&
            CausesValidation=   0   'False
            ClipControls    =   0   'False
            HasDC           =   0   'False
            Height          =   1020
            Left            =   150
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   40
            TabStop         =   0   'False
            ToolTipText     =   "Sector Ceiling Texture"
            Top             =   240
            Width           =   1020
            Begin VB.Image imgFloor 
               Height          =   960
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   960
            End
         End
         Begin VB.Label lblFloor 
            Alignment       =   2  'Center
            Caption         =   "STARTAN3"
            Height          =   210
            Left            =   150
            TabIndex        =   41
            Top             =   1260
            UseMnemonic     =   0   'False
            Width           =   1020
         End
      End
      Begin VB.Frame fraBackSidedef 
         Caption         =   " Back Side "
         Height          =   1590
         Left            =   7695
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   3465
         Begin VB.PictureBox picS2Upper 
            BackColor       =   &H8000000C&
            CausesValidation=   0   'False
            ClipControls    =   0   'False
            HasDC           =   0   'False
            Height          =   1020
            Left            =   150
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   "Back Side Upper Texture"
            Top             =   240
            Width           =   1020
            Begin VB.Image imgS2Upper 
               Height          =   960
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   960
            End
         End
         Begin VB.PictureBox picS2Middle 
            BackColor       =   &H8000000C&
            CausesValidation=   0   'False
            ClipControls    =   0   'False
            HasDC           =   0   'False
            Height          =   1020
            Left            =   1230
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Back Side Middle Texture"
            Top             =   240
            Width           =   1020
            Begin VB.Image imgS2Middle 
               Height          =   960
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   960
            End
         End
         Begin VB.PictureBox picS2Lower 
            BackColor       =   &H8000000C&
            CausesValidation=   0   'False
            ClipControls    =   0   'False
            HasDC           =   0   'False
            Height          =   1020
            Left            =   2310
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Back Side Lower Texture"
            Top             =   240
            Width           =   1020
            Begin VB.Image imgS2Lower 
               Height          =   960
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   960
            End
         End
         Begin VB.Label lblS2Upper 
            Alignment       =   2  'Center
            Caption         =   "STARTAN3"
            Height          =   210
            Left            =   150
            TabIndex        =   22
            Top             =   1260
            UseMnemonic     =   0   'False
            Width           =   1020
         End
         Begin VB.Label lblS2Middle 
            Alignment       =   2  'Center
            Caption         =   "STARTAN3"
            Height          =   210
            Left            =   1230
            TabIndex        =   21
            Top             =   1260
            UseMnemonic     =   0   'False
            Width           =   1020
         End
         Begin VB.Label lblS2Lower 
            Alignment       =   2  'Center
            Caption         =   "STARTAN3"
            Height          =   210
            Left            =   2310
            TabIndex        =   20
            Top             =   1260
            UseMnemonic     =   0   'False
            Width           =   1020
         End
      End
      Begin VB.Frame fraFrontSidedef 
         Caption         =   " Front Side "
         Height          =   1590
         Left            =   4170
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   3465
         Begin VB.PictureBox picS1Lower 
            BackColor       =   &H8000000C&
            CausesValidation=   0   'False
            ClipControls    =   0   'False
            HasDC           =   0   'False
            Height          =   1020
            Left            =   2310
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Front Side Lower Texture"
            Top             =   240
            Width           =   1020
            Begin VB.Image imgS1Lower 
               Height          =   960
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   960
            End
         End
         Begin VB.PictureBox picS1Middle 
            BackColor       =   &H8000000C&
            CausesValidation=   0   'False
            ClipControls    =   0   'False
            HasDC           =   0   'False
            Height          =   1020
            Left            =   1230
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "Front Side Middle Texture"
            Top             =   240
            Width           =   1020
            Begin VB.Image imgS1Middle 
               Height          =   960
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   960
            End
         End
         Begin VB.PictureBox picS1Upper 
            BackColor       =   &H8000000C&
            CausesValidation=   0   'False
            ClipControls    =   0   'False
            HasDC           =   0   'False
            Height          =   1020
            Left            =   150
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Front Side Upper Texture"
            Top             =   240
            Width           =   1020
            Begin VB.Image imgS1Upper 
               Height          =   960
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   960
            End
         End
         Begin VB.Label lblS1Lower 
            Alignment       =   2  'Center
            Caption         =   "STARTAN3"
            Height          =   210
            Left            =   2310
            TabIndex        =   15
            Top             =   1260
            UseMnemonic     =   0   'False
            Width           =   1020
         End
         Begin VB.Label lblS1Middle 
            Alignment       =   2  'Center
            Caption         =   "STARTAN3"
            Height          =   210
            Left            =   1230
            TabIndex        =   14
            Top             =   1260
            UseMnemonic     =   0   'False
            Width           =   1020
         End
         Begin VB.Label lblS1Upper 
            Alignment       =   2  'Center
            Caption         =   "STARTAN3"
            Height          =   210
            Left            =   150
            TabIndex        =   13
            Top             =   1260
            UseMnemonic     =   0   'False
            Width           =   1020
         End
      End
      Begin VB.Label lblBarText 
         AutoSize        =   -1  'True
         Height          =   210
         Left            =   60
         TabIndex        =   146
         Top             =   15
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblMode 
         AutoSize        =   -1  'True
         Caption         =   "Vertices"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   555
         Left            =   720
         TabIndex        =   119
         Top             =   540
         Width           =   1905
      End
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Index           =   2
      Left            =   750
      Picture         =   "frmMain.frx":10BBA
      Top             =   3375
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Index           =   1
      Left            =   240
      Picture         =   "frmMain.frx":10EC4
      Top             =   3375
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMissingTexture 
      Height          =   960
      Left            =   105
      Picture         =   "frmMain.frx":111CE
      Top             =   3750
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image imgUnknownFlat 
      Height          =   960
      Left            =   2025
      Picture         =   "frmMain.frx":112E5
      Top             =   3750
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image imgUnknownTexture 
      Height          =   960
      Left            =   1065
      Picture         =   "frmMain.frx":113F2
      Top             =   3750
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Menu mnuFile 
      Caption         =   " &File "
      Begin VB.Menu itmFile 
         Caption         =   "&New Map"
         Index           =   0
      End
      Begin VB.Menu itmFile 
         Caption         =   "&Open Map..."
         Index           =   1
      End
      Begin VB.Menu itmFile 
         Caption         =   "&Close Map"
         Index           =   2
      End
      Begin VB.Menu itmFile 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu itmFile 
         Caption         =   "&Save Map"
         Index           =   4
      End
      Begin VB.Menu itmFile 
         Caption         =   "Save Map &As..."
         Index           =   5
      End
      Begin VB.Menu itmFile 
         Caption         =   "Save Map &Into...   "
         Index           =   6
      End
      Begin VB.Menu itmFile 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu itmFile 
         Caption         =   "&Export Map..."
         Index           =   8
      End
      Begin VB.Menu itmFile 
         Caption         =   "Export &Picture..."
         Index           =   9
      End
      Begin VB.Menu itmFile 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu itmFile 
         Caption         =   "&Build Nodes"
         Index           =   11
      End
      Begin VB.Menu itmFile 
         Caption         =   "&Test Map"
         Index           =   12
      End
      Begin VB.Menu itmFile 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu itmFileRecent 
         Caption         =   "itmFileRecent"
         Index           =   0
      End
      Begin VB.Menu itmFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu itmFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   " &Edit "
      Begin VB.Menu itmEditUndo 
         Caption         =   "&Undo ... "
      End
      Begin VB.Menu itmEditRedo 
         Caption         =   "&Redo ..."
      End
      Begin VB.Menu itmEditLine2 
         Caption         =   "-"
      End
      Begin VB.Menu itmEditMode 
         Caption         =   "&Move Move"
         Index           =   0
      End
      Begin VB.Menu itmEditMode 
         Caption         =   "&Vertices Mode"
         Index           =   1
      End
      Begin VB.Menu itmEditMode 
         Caption         =   "&Lines Mode"
         Index           =   2
      End
      Begin VB.Menu itmEditMode 
         Caption         =   "&Sectors Mode"
         Index           =   3
      End
      Begin VB.Menu itmEditMode 
         Caption         =   "&Things Mode"
         Index           =   4
      End
      Begin VB.Menu itmEditMode 
         Caption         =   "&3D Mode"
         Index           =   5
      End
      Begin VB.Menu itmEditLine1 
         Caption         =   "-"
      End
      Begin VB.Menu itmEditCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu itmEditCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu itmEditPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu itmEditDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu itmEditLine6 
         Caption         =   "-"
      End
      Begin VB.Menu itmEditFind 
         Caption         =   "Find..."
      End
      Begin VB.Menu itmEditReplace 
         Caption         =   "Find and Replace..."
      End
      Begin VB.Menu itmEditLine5 
         Caption         =   "-"
      End
      Begin VB.Menu itmEditFlipH 
         Caption         =   "Flip Horizontally"
      End
      Begin VB.Menu itmEditFlipV 
         Caption         =   "Flip Vertically"
      End
      Begin VB.Menu itmEditRotate 
         Caption         =   "Rotate"
      End
      Begin VB.Menu itmEditResize 
         Caption         =   "Resize"
      End
      Begin VB.Menu itmEditLine4 
         Caption         =   "-"
      End
      Begin VB.Menu itmEditSnapToGrid 
         Caption         =   "Snap To Grid"
         Checked         =   -1  'True
      End
      Begin VB.Menu itmEditStitch 
         Caption         =   "Stitch Vertices"
         Checked         =   -1  'True
      End
      Begin VB.Menu itmEditCenterView 
         Caption         =   "Center View"
      End
      Begin VB.Menu itmEditLine3 
         Caption         =   "-"
      End
      Begin VB.Menu itmEditMapOptions 
         Caption         =   "Map &Options...   "
      End
   End
   Begin VB.Menu mnuVertices 
      Caption         =   " &Vertices "
      Begin VB.Menu itmVerticesSnapToGrid 
         Caption         =   "&Snap to Grid"
      End
      Begin VB.Menu itmVerticesLine2 
         Caption         =   "-"
      End
      Begin VB.Menu itmVerticesStitch 
         Caption         =   "Stitch Vertices"
      End
      Begin VB.Menu itmVerticesLine1 
         Caption         =   "-"
      End
      Begin VB.Menu itmVerticesClearUnused 
         Caption         =   "&Clear Unused Vertices"
      End
   End
   Begin VB.Menu mnuLines 
      Caption         =   " &Lines "
      Begin VB.Menu itmLinesSnapToGrid 
         Caption         =   "&Snap to Grid"
      End
      Begin VB.Menu itmLinesLine4 
         Caption         =   "-"
      End
      Begin VB.Menu itmLinesAlign 
         Caption         =   "Autoalign Textures..."
      End
      Begin VB.Menu itmLinesLine3 
         Caption         =   "-"
      End
      Begin VB.Menu itmLinesSelect 
         Caption         =   "Select only 1-sided lines"
         Index           =   0
      End
      Begin VB.Menu itmLinesSelect 
         Caption         =   "Select only 2-sided lines"
         Index           =   1
      End
      Begin VB.Menu itmLinesLine1 
         Caption         =   "-"
      End
      Begin VB.Menu itmLinesFlipLinedefs 
         Caption         =   "Flip &Linedefs"
      End
      Begin VB.Menu itmLinesFlipSidedefs 
         Caption         =   "Flip &Sidedefs"
      End
      Begin VB.Menu itmLinesCurve 
         Caption         =   "Curve Linedefs"
      End
      Begin VB.Menu itmLinesLine2 
         Caption         =   "-"
      End
      Begin VB.Menu itmLinesCopy 
         Caption         =   "&Copy Linedef Properties"
      End
      Begin VB.Menu itmLinesPaste 
         Caption         =   "&Paste Linedef Properties"
      End
   End
   Begin VB.Menu mnuSectors 
      Caption         =   " &Sectors "
      Begin VB.Menu itmSectorsSnapToGrid 
         Caption         =   "&Snap to Grid"
      End
      Begin VB.Menu itmSectorsLine1 
         Caption         =   "-"
      End
      Begin VB.Menu itmSectorsJoin 
         Caption         =   "Join Sectors"
      End
      Begin VB.Menu itmSectorsMerge 
         Caption         =   "Merge Sectors"
      End
      Begin VB.Menu itmSectorsLine5 
         Caption         =   "-"
      End
      Begin VB.Menu itmSectorsGradientBrightness 
         Caption         =   "Gradient Brightness"
      End
      Begin VB.Menu itmSectorsGradientFloors 
         Caption         =   "Gradient Floors"
      End
      Begin VB.Menu itmSectorsGradientCeilings 
         Caption         =   "Gradient Ceilings"
      End
      Begin VB.Menu itmSectorsLine2 
         Caption         =   "-"
      End
      Begin VB.Menu itmSectorsRaiseFloor 
         Caption         =   "Raise Floor by 8"
      End
      Begin VB.Menu itmSectorsLowerFloor 
         Caption         =   "Lower Floor by 8"
      End
      Begin VB.Menu itmSectorsRaiseCeiling 
         Caption         =   "Raise Ceiling by 8"
      End
      Begin VB.Menu itmSectorsLowerCeiling 
         Caption         =   "Lower Ceiling by 8"
      End
      Begin VB.Menu itmSectorsLine3 
         Caption         =   "-"
      End
      Begin VB.Menu itmSectorsIncBrightness 
         Caption         =   "Increase Brightness"
      End
      Begin VB.Menu itmSectorsDecBrightness 
         Caption         =   "Decrease Brightness"
      End
      Begin VB.Menu itmSectorsLine4 
         Caption         =   "-"
      End
      Begin VB.Menu itmSectorsCopy 
         Caption         =   "&Copy Sector Properties"
      End
      Begin VB.Menu itmSectorsPaste 
         Caption         =   "&Paste Sector Properties"
      End
   End
   Begin VB.Menu mnuThings 
      Caption         =   " &Things "
      Begin VB.Menu itmThingsSnapToGrid 
         Caption         =   "&Snap to Grid"
      End
      Begin VB.Menu itmThingsLine1 
         Caption         =   "-"
      End
      Begin VB.Menu itmThingsCopy 
         Caption         =   "&Copy Thing Properties"
      End
      Begin VB.Menu itmThingsPaste 
         Caption         =   "&Paste Thing Properties"
      End
      Begin VB.Menu itmThingsLine2 
         Caption         =   "-"
      End
      Begin VB.Menu itmThingsFilter 
         Caption         =   "&Filter Things..."
      End
   End
   Begin VB.Menu mnuPrefabs 
      Caption         =   " &Prefabs "
      Begin VB.Menu itmPrefabQuick 
         Caption         =   "itmPrefabQuick"
         Index           =   0
      End
      Begin VB.Menu itmPrefabQuick 
         Caption         =   "itmPrefabQuick"
         Index           =   1
      End
      Begin VB.Menu itmPrefabQuick 
         Caption         =   "itmPrefabQuick"
         Index           =   2
      End
      Begin VB.Menu itmPrefabQuick 
         Caption         =   "itmPrefabQuick"
         Index           =   3
      End
      Begin VB.Menu itmPrefabQuick 
         Caption         =   "itmPrefabQuick"
         Index           =   4
      End
      Begin VB.Menu itmPrefabLine1 
         Caption         =   "-"
      End
      Begin VB.Menu itmPrefabPrevious 
         Caption         =   "Insert &Previous Prefab"
      End
      Begin VB.Menu itmPrefabInsert 
         Caption         =   "&Insert Prefab from File..."
      End
      Begin VB.Menu itmPrefabLine2 
         Caption         =   "-"
      End
      Begin VB.Menu itmPrefabSaveSel 
         Caption         =   "Save &selection as Prefab..."
      End
      Begin VB.Menu itmPrefabSaveMap 
         Caption         =   "Save &map as Prefab..."
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuScripts 
      Caption         =   " &Scripts "
      Begin VB.Menu itmScriptEdit 
         Caption         =   "itmScriptEdit"
         Index           =   0
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   " &Tools "
      Begin VB.Menu itmToolsFindErrors 
         Caption         =   "Find map errors..."
      End
      Begin VB.Menu itmToolsLine2 
         Caption         =   "-"
      End
      Begin VB.Menu itmToolsClearTextures 
         Caption         =   "&Remove Unused Textures"
      End
      Begin VB.Menu itmToolsFixTextures 
         Caption         =   "Fix &Missing Textures"
      End
      Begin VB.Menu itmToolsFixZeroLinedefs 
         Caption         =   "Fix &Zero-Length Linedefs"
      End
      Begin VB.Menu itmToolsLine1 
         Caption         =   "-"
      End
      Begin VB.Menu itmToolsConfiguration 
         Caption         =   "&Configuration..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   " &Help "
      Begin VB.Menu itmHelpWebsite 
         Caption         =   "Doom Builder &Website..."
      End
      Begin VB.Menu itmHelpFAQ 
         Caption         =   "Frequently Asked Questions..."
      End
      Begin VB.Menu itmHelpLine1 
         Caption         =   "-"
      End
      Begin VB.Menu itmHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'    Doom Builder
'    Copyright (c) 2003 Pascal vd Heiden, www.codeimp.com
'    This program is released under GNU General Public License
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'


'Do not allow any undeclared variables
Option Explicit

'Case sensitive comparisions
Option Compare Binary


Private DeselectAfterEdit As Boolean
Public OriginalMessageHandler As Long

Private GrabX As Single
Private GrabY As Single

Private StartSelection As Boolean
Private LastSelX As Single
Private LastSelY As Single

Private AutoScrollX As Long
Private AutoScrollY As Long

Private NoEditing As Boolean

Private F7Count As Long

Private LastMouseButton As Integer
Private LastMouseShift As Integer
Public LastMouseX As Single, LastMouseY As Single

Private DrawingCoords() As POINT
Private NumDrawingCoords As Long

Public Sub ApplyInterfaceConfiguration()
     
     'Show/hide toolbar
     tlbToolbar.Visible = (Val(Config("showtoolbar")) = vbChecked)
     
     'Show hide and position the details bar
     Select Case Val(Config("detailsbar"))
          
          Case 0:   'No details bar
               picBar.Visible = False
               picSBar.Visible = False
               
          Case 1:   'Bottom
               picBar.Visible = True
               picSBar.Visible = False
               picBar.Align = vbAlignBottom
               
          Case 2:   'Top
               picBar.Visible = True
               picSBar.Visible = False
               picBar.Align = vbAlignTop
               
          Case 3:   'Left
               picBar.Visible = False
               picSBar.Visible = True
               picSBar.Align = vbAlignLeft
               
               
          Case 4:   'Right
               picBar.Visible = False
               picSBar.Visible = True
               picSBar.Align = vbAlignRight
               
     End Select
     
     'Toggle bar
     If (Val(Config("togglebar")) = 0) Then InfoBarOpen Else InfoBarClose
End Sub

Public Sub CancelCurrentOperation()
     
     'Can only be done if a map is open
     If (mapfile <> "") Then
          
          'Cancel if in drawing operation
          Select Case submode
               Case ESM_SELECTING: CancelSelectOperation: RedrawMap False
               Case ESM_DRAGGING: CancelDragOperation: RedrawMap False
               Case ESM_DRAWING: CancelDrawOperation: RedrawMap False
               Case ESM_PASTING: CancelDragOperation: RedrawMap False
          End Select
          
          'Show highlight
          ShowHighlight LastX, LastY
          
          'Update status
          UpdateStatusBar
     End If
     
     'Not scrolling anymore
     Scrolling = False
     Screen.MousePointer = vbNormal
End Sub

Private Sub CancelDragOperation()
     
     'End of drag operation, set mode back to normal
     submode = ESM_NONE
     
     'Map was changed
     mapnodeschanged = True
     mapchanged = True
     
     'We dont need these anymore
     ReDim changedlines(0)
     numchangedlines = 0
     
     'Perform undo
     PerformUndo
     
     'Withdraw the redo (like nothing happend)
     WithdrawRedo
End Sub

Public Sub CancelDrawOperation()
     
     'Deselect all
     RemoveSelection False
     
     'End of drawing operation, set mode back to normal
     submode = ESM_NONE
     
     'Map was changed
     mapnodeschanged = True
     mapchanged = True
     
     'We dont need these anymore
     ReDim changedlines(0)
     numchangedlines = 0
     ReDim DrawingCoords(0)
     NumDrawingCoords = 0
     
     'Perform undo
     PerformUndo
     
     'Withdraw the redo (like nothing happend)
     WithdrawRedo
End Sub

Public Sub CancelSelectOperation()
     
     'End of select operation, set mode back to normal
     submode = ESM_NONE
End Sub

Public Sub ChangeAutoscroll(ByVal ForceDisable As Boolean)
     Dim px As Long, py As Long
     Dim sx As Long, sy As Long
     
     Const ScrollBounds As Long = 100
     Const ScrollMultiplier As Single = 0.6
     
     'Check if we should stop autoscroling
     If (Config("autoscroll") = vbUnchecked) Or (ForceDisable = True) Or _
        (mode = EM_MOVE) Or (submode = ESM_NONE) Then
          
          'Stop autoscrolling
          tmrAutoScroll.Enabled = False
          
          'Leave now
          Exit Sub
     End If
     
     'Get pixel coordinates
     px = -(ViewLeft - LastX) * ViewZoom
     py = -(ViewTop - LastY) * ViewZoom
     
     'Determine scrolling in X
     If (px < ScrollBounds) Then
          
          'Scroll to the left
          sx = -(ScrollBounds - px) * ScrollMultiplier
          If (sx < -ScrollBounds) Then sx = 0
          
     ElseIf (px > (picMap.width - 4) - ScrollBounds) Then
          
          'Scroll to the right
          sx = (px - ((picMap.width - 4) - ScrollBounds)) * ScrollMultiplier
          If (sx > ScrollBounds) Then sx = 0
     End If
     
     'Determine scrolling in Y
     If (py < ScrollBounds) Then
          
          'Scroll to the top
          sy = -(ScrollBounds - py) * ScrollMultiplier
          If (sy < -ScrollBounds) Then sy = 0
          
     ElseIf (py > (picMap.height - 4) - ScrollBounds) Then
          
          'Scroll to the bottom
          sy = (py - ((picMap.height - 4) - ScrollBounds)) * ScrollMultiplier
          If (sy > ScrollBounds) Then sy = 0
     End If
     
     'Check for scrolling
     If (sx <> 0) Or (sy <> 0) Then
          
          'Set scrolling
          AutoScrollX = sx / ViewZoom
          AutoScrollY = sy / ViewZoom
          tmrAutoScroll.Enabled = True
     Else
          
          'No scrolling
          tmrAutoScroll.Enabled = False
     End If
End Sub

Private Sub ChangeLinesHighlight(ByVal X As Long, ByVal Y As Long, Optional ByVal Forceupdate As Boolean)
     Dim distance As Long
     Dim nearest As Long
     Dim OldSelected As Long
     Dim xl As Long, yl As Long
     
     Dim Action As String
     Dim Length As Long
     Dim sTag As String
     Dim FHeight As String
     Dim bheight As String
     Dim FSector As String
     Dim BSector As String
     Dim fx As String
     Dim fy As String
     Dim bx As String
     Dim by As String
     Dim S1U As String
     Dim S1M As String
     Dim S1L As String
     Dim S2U As String
     Dim S2M As String
     Dim S2L As String
     
     'Keep old selection ifnot forced to update
     If (Not Forceupdate) Then OldSelected = currentselected Else OldSelected = -1
     
     'Check if a previous linedef was selected
     If (currentselected > -1) Then
          
          'Render the last selected linedef to normal (also vertices, those have been overdrawn)
          Render_AllLinedefs vertexes(0), linedefs(0), currentselected, currentselected, submode, indicatorsize
          If (Config("mode1vertices")) Then
               Render_AllVertices vertexes(0), linedefs(currentselected).v1, linedefs(currentselected).v1, vertexsize
               Render_AllVertices vertexes(0), linedefs(currentselected).v2, linedefs(currentselected).v2, vertexsize
          End If
          
          'Check if linedef has an action
          If (linedefs(currentselected).effect > 0) Then
               
               'Render the last tagged sectors
               If (linedefs(currentselected).tag <> 0) Then Render_TaggedSectors vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), numsectors, numlinedefs, linedefs(currentselected).tag, 0, indicatorsize, Config("mode1vertices"), vertexsize
               
               'Check if line has a known effect
               If (mapconfig("linedeftypes").Exists(CStr(linedefs(currentselected).effect)) = True) Then
                    
                    'Render by hexen format
                    If (mapconfig("linedeftypes")(CStr(linedefs(currentselected).effect))("mark1") = 1) And (linedefs(currentselected).arg0 > 0) Then Render_TaggedSectors vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), numsectors, numlinedefs, linedefs(currentselected).arg0, 0, indicatorsize, Config("mode1vertices"), vertexsize
                    If (mapconfig("linedeftypes")(CStr(linedefs(currentselected).effect))("mark2") = 1) And (linedefs(currentselected).arg1 > 0) Then Render_TaggedSectors vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), numsectors, numlinedefs, linedefs(currentselected).arg1, 0, indicatorsize, Config("mode1vertices"), vertexsize
                    If (mapconfig("linedeftypes")(CStr(linedefs(currentselected).effect))("mark3") = 1) And (linedefs(currentselected).arg2 > 0) Then Render_TaggedSectors vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), numsectors, numlinedefs, linedefs(currentselected).arg2, 0, indicatorsize, Config("mode1vertices"), vertexsize
                    If (mapconfig("linedeftypes")(CStr(linedefs(currentselected).effect))("mark4") = 1) And (linedefs(currentselected).arg3 > 0) Then Render_TaggedSectors vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), numsectors, numlinedefs, linedefs(currentselected).arg3, 0, indicatorsize, Config("mode1vertices"), vertexsize
                    If (mapconfig("linedeftypes")(CStr(linedefs(currentselected).effect))("mark5") = 1) And (linedefs(currentselected).arg4 > 0) Then Render_TaggedSectors vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), numsectors, numlinedefs, linedefs(currentselected).arg4, 0, indicatorsize, Config("mode1vertices"), vertexsize
                    
                    'Check if things are shown as well
                    If (Val(Config("modethings"))) Then
                         If ((mapconfig("linedeftypes")(CStr(linedefs(currentselected).effect))("mark1") = 2) And (linedefs(currentselected).arg0 > 0)) Or _
                            ((mapconfig("linedeftypes")(CStr(linedefs(currentselected).effect))("mark2") = 2) And (linedefs(currentselected).arg1 > 0)) Or _
                            ((mapconfig("linedeftypes")(CStr(linedefs(currentselected).effect))("mark3") = 2) And (linedefs(currentselected).arg2 > 0)) Or _
                            ((mapconfig("linedeftypes")(CStr(linedefs(currentselected).effect))("mark4") = 2) And (linedefs(currentselected).arg3 > 0)) Or _
                            ((mapconfig("linedeftypes")(CStr(linedefs(currentselected).effect))("mark5") = 2) And (linedefs(currentselected).arg4 > 0)) Then
                              
                              'Redraw map
                              Render_AllThingsDarkened things(0), 0, numthings - 1, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, filterthings, filtersettings
                              Render_AllLinedefs vertexes(0), linedefs(0), 0, numlinedefs - 1, submode, indicatorsize
                              If (Config("mode1vertices")) Then Render_AllVertices vertexes(0), 0, numvertexes - 1, vertexsize
                         End If
                    End If
               End If
          End If
     End If
     
     'Get the nearest linedef and select it if within allowed distance
     nearest = NearestLinedef(X, Y, vertexes(0), linedefs(0), numlinedefs, distance)
     If (distance <= Config("lineselectdistance") / ViewZoom) Then currentselected = nearest Else currentselected = -1
     
     'Check if a new selection is made
     If (currentselected > -1) Then
          
          'Check if linedef has an action
          If (linedefs(currentselected).effect > 0) Then
               
               'Render the tagged sectors if line has a tag
               If (linedefs(currentselected).tag <> 0) Then Render_TaggedSectors vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), numsectors, numlinedefs, linedefs(currentselected).tag, CLR_SECTORTAG, indicatorsize, Config("mode1vertices"), vertexsize
               
               'Check if we can render by hexen format
               If (mapconfig("linedeftypes").Exists(CStr(linedefs(currentselected).effect)) = True) Then
                    
                    'Render by hexen format
                    If (mapconfig("linedeftypes")(CStr(linedefs(currentselected).effect))("mark1") = 1) And (linedefs(currentselected).arg0 > 0) Then Render_TaggedSectors vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), numsectors, numlinedefs, linedefs(currentselected).arg0, CLR_SECTORTAG, indicatorsize, Config("mode1vertices"), vertexsize
                    If (mapconfig("linedeftypes")(CStr(linedefs(currentselected).effect))("mark2") = 1) And (linedefs(currentselected).arg1 > 0) Then Render_TaggedSectors vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), numsectors, numlinedefs, linedefs(currentselected).arg1, CLR_SECTORTAG, indicatorsize, Config("mode1vertices"), vertexsize
                    If (mapconfig("linedeftypes")(CStr(linedefs(currentselected).effect))("mark3") = 1) And (linedefs(currentselected).arg2 > 0) Then Render_TaggedSectors vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), numsectors, numlinedefs, linedefs(currentselected).arg2, CLR_SECTORTAG, indicatorsize, Config("mode1vertices"), vertexsize
                    If (mapconfig("linedeftypes")(CStr(linedefs(currentselected).effect))("mark4") = 1) And (linedefs(currentselected).arg3 > 0) Then Render_TaggedSectors vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), numsectors, numlinedefs, linedefs(currentselected).arg3, CLR_SECTORTAG, indicatorsize, Config("mode1vertices"), vertexsize
                    If (mapconfig("linedeftypes")(CStr(linedefs(currentselected).effect))("mark5") = 1) And (linedefs(currentselected).arg4 > 0) Then Render_TaggedSectors vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), numsectors, numlinedefs, linedefs(currentselected).arg4, CLR_SECTORTAG, indicatorsize, Config("mode1vertices"), vertexsize
                    
                    'Check if things are shown as well
                    If (Val(Config("modethings"))) Then
                         
                         'Render things by hexen format
                         If (mapconfig("linedeftypes")(CStr(linedefs(currentselected).effect))("mark1") = 2) And (linedefs(currentselected).arg0 > 0) Then Render_TaggedThings things(0), numthings - 1, linedefs(currentselected).arg0, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, Val(Config("allthingsrects")), ViewZoom, filterthings, filtersettings
                         If (mapconfig("linedeftypes")(CStr(linedefs(currentselected).effect))("mark2") = 2) And (linedefs(currentselected).arg1 > 0) Then Render_TaggedThings things(0), numthings - 1, linedefs(currentselected).arg1, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, Val(Config("allthingsrects")), ViewZoom, filterthings, filtersettings
                         If (mapconfig("linedeftypes")(CStr(linedefs(currentselected).effect))("mark3") = 2) And (linedefs(currentselected).arg2 > 0) Then Render_TaggedThings things(0), numthings - 1, linedefs(currentselected).arg2, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, Val(Config("allthingsrects")), ViewZoom, filterthings, filtersettings
                         If (mapconfig("linedeftypes")(CStr(linedefs(currentselected).effect))("mark4") = 2) And (linedefs(currentselected).arg3 > 0) Then Render_TaggedThings things(0), numthings - 1, linedefs(currentselected).arg3, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, Val(Config("allthingsrects")), ViewZoom, filterthings, filtersettings
                         If (mapconfig("linedeftypes")(CStr(linedefs(currentselected).effect))("mark5") = 2) And (linedefs(currentselected).arg4 > 0) Then Render_TaggedThings things(0), numthings - 1, linedefs(currentselected).arg4, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, Val(Config("allthingsrects")), ViewZoom, filterthings, filtersettings
                    End If
               End If
          End If
          
          'Render the selected linedef to selected (also vertices, those have been overdrawn)
          Render_LinedefLine vertexes(linedefs(currentselected).v1).X, vertexes(linedefs(currentselected).v1).Y, vertexes(linedefs(currentselected).v2).X, vertexes(linedefs(currentselected).v2).Y, CLR_LINEHIGHLIGHT, indicatorsize
          If (Config("mode1vertices")) Then
               Render_AllVertices vertexes(0), linedefs(currentselected).v1, linedefs(currentselected).v1, vertexsize
               Render_AllVertices vertexes(0), linedefs(currentselected).v2, linedefs(currentselected).v2, vertexsize
          End If
     End If
     
     'Show the rendered changes
     If (OldSelected <> currentselected) Then picMap.Refresh
     
     'Check if we should show the info
     If (currentselected > -1) Then
          
          'Only update when changed
          If (OldSelected <> currentselected) Then ShowLinedefInfo currentselected
     Else
          
          'Hide the info
          HideLinedefInfo
     End If
End Sub

Private Sub ChangeSectorsHighlight(ByVal X As Long, ByVal Y As Long, Optional ByVal Forceupdate As Boolean)
     Dim nearest As Long
     Dim OldSelected As Long
     Dim ld As Long, ldfound As Long
     
     Dim effect As String
     Dim Ceiling As Long
     Dim Floor As Long
     Dim sTag As Long
     Dim sHeight As Long
     Dim Brightness As Long
     Dim tceiling As String
     Dim tfloor As String
     
     'Keep old selection ifnot forced to update
     If (Not Forceupdate) Then OldSelected = currentselected Else OldSelected = -1
     
     'Check if a previous sector was selected
     If (currentselected > -1) Then
          
          'Go for all linedefs
          For ld = 0 To (numlinedefs - 1)
               
               'Check if one of the sidedefs belong to this sector
               ldfound = 0
               If (linedefs(ld).s1 > -1) Then If (sidedefs(linedefs(ld).s1).sector = currentselected) Then ldfound = 1
               If (linedefs(ld).s2 > -1) Then If (sidedefs(linedefs(ld).s2).sector = currentselected) Then ldfound = 1
               
               If (ldfound) Then
                    
                    'Render this linedef to normal (also vertices, those have been overdrawn)
                    Render_AllLinedefs vertexes(0), linedefs(0), ld, ld, submode, indicatorsize
                    If (Config("mode2vertices")) Then
                         Render_AllVertices vertexes(0), linedefs(ld).v1, linedefs(ld).v1, vertexsize
                         Render_AllVertices vertexes(0), linedefs(ld).v2, linedefs(ld).v2, vertexsize
                    End If
               End If
          Next ld
          
          'Check if sector has a tag
          If (sectors(currentselected).tag <> 0) Then
               
               'Render the last tagged linedefs
               Render_TaggedLinedefs vertexes(0), linedefs(0), numlinedefs, sectors(currentselected).tag, 1, 0, indicatorsize, 0, vertexsize
               
               'Check if things are shown as well
               If (Val(Config("modethings"))) Then
                    
                    'Redraw map
                    Render_AllThingsDarkened things(0), 0, numthings - 1, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, filterthings, filtersettings
                    Render_AllLinedefs vertexes(0), linedefs(0), 0, numlinedefs - 1, submode, indicatorsize
                    If (Config("mode2vertices")) Then Render_AllVertices vertexes(0), 0, numvertexes - 1, vertexsize
               End If
          End If
     End If
     
     'Get the intersecting sector and select it
     nearest = IntersectSector(X, Y, vertexes(0), linedefs(0), VarPtr(sidedefs(0)), numlinedefs, 0)
     currentselected = nearest
     
     'Check if a new sector is selected
     If (currentselected > -1) Then
          
          'Check if sector has a tag
          If (sectors(currentselected).tag <> 0) Then
               
               'Render the tagged linedefs
               Render_TaggedLinedefs vertexes(0), linedefs(0), numlinedefs, sectors(currentselected).tag, 1, CLR_SECTORTAG, indicatorsize, 0, vertexsize
               
               'Check if things are shown as well
               If (Val(Config("modethings"))) Then Render_TaggedArgThings things(0), numthings, sectors(currentselected).tag, 1, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, Val(Config("allthingsrects")), ViewZoom, filterthings, filtersettings
          End If
          
          'Go for all linedefs
          For ld = 0 To (numlinedefs - 1)
               
               'Check if one of the sidedefs belong to this sector
               ldfound = 0
               If (linedefs(ld).s1 > -1) Then If (sidedefs(linedefs(ld).s1).sector = currentselected) Then ldfound = 1
               If (linedefs(ld).s2 > -1) Then If (sidedefs(linedefs(ld).s2).sector = currentselected) Then ldfound = 1
               
               If (ldfound) Then
                    
                    'Render this linedef to selected (also vertices, those have been overdrawn)
                    Render_LinedefLine vertexes(linedefs(ld).v1).X, vertexes(linedefs(ld).v1).Y, vertexes(linedefs(ld).v2).X, vertexes(linedefs(ld).v2).Y, CLR_LINEHIGHLIGHT, indicatorsize
                    If (Config("mode2vertices")) Then
                         Render_AllVertices vertexes(0), linedefs(ld).v1, linedefs(ld).v1, vertexsize
                         Render_AllVertices vertexes(0), linedefs(ld).v2, linedefs(ld).v2, vertexsize
                    End If
               End If
          Next ld
     End If
     
     'Show the rendered changes
     If (OldSelected <> currentselected) Then picMap.Refresh
     
     'Check if we should show the info
     If (currentselected > -1) Then
          
          'Only update when changed
          If (OldSelected <> currentselected) Then
               
               'Show the information
               ShowSectorInfo currentselected
          End If
     Else
          
          'Hide the info
          HideSectorInfo
     End If
End Sub

Private Sub ChangeThingsHighlight(ByVal X As Long, ByVal Y As Long, Optional ByVal Forceupdate As Boolean)
     Dim distance As Long
     Dim nearest As Long
     Dim OldSelected As Long
     Dim tType As String
     Dim angle As String
     Dim Flags As String
     Dim nType As Long
     Dim spritename As String
     Dim tx As Long
     Dim ty As Long
     Dim SectorUndrawn As Boolean
     Dim sTag As String
     Dim sAction As String
     
     'Keep old selection if not forced to update
     If (Not Forceupdate) Then OldSelected = currentselected Else OldSelected = -1
     
     'Anything selected before?
     If (currentselected > -1) Then
          
          'Render the last current selected to normal
          Render_AllThings things(0), currentselected, currentselected, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, Val(Config("allthingsrects")), ViewZoom, filterthings, filtersettings
          If Config("thingrects") Then Render_BoxSwitched things(currentselected).X, things(currentselected).Y, things(currentselected).size * ViewZoom, PAL_NORMAL, (Config("thingrects") - 1), PAL_NORMAL
          
          'Check if thing has a known effect
          If (mapconfig("linedeftypes").Exists(CStr(things(currentselected).effect)) = True) Then
               
               'Render by hexen format
               If (mapconfig("linedeftypes")(CStr(things(currentselected).effect))("mark1") = 1) And (things(currentselected).arg0 > 0) Then Render_TaggedSectors vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), numsectors, numlinedefs, things(currentselected).arg0, 0, indicatorsize, 0, vertexsize: SectorUndrawn = True
               If (mapconfig("linedeftypes")(CStr(things(currentselected).effect))("mark2") = 1) And (things(currentselected).arg1 > 0) Then Render_TaggedSectors vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), numsectors, numlinedefs, things(currentselected).arg1, 0, indicatorsize, 0, vertexsize: SectorUndrawn = True
               If (mapconfig("linedeftypes")(CStr(things(currentselected).effect))("mark3") = 1) And (things(currentselected).arg2 > 0) Then Render_TaggedSectors vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), numsectors, numlinedefs, things(currentselected).arg2, 0, indicatorsize, 0, vertexsize: SectorUndrawn = True
               If (mapconfig("linedeftypes")(CStr(things(currentselected).effect))("mark4") = 1) And (things(currentselected).arg3 > 0) Then Render_TaggedSectors vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), numsectors, numlinedefs, things(currentselected).arg3, 0, indicatorsize, 0, vertexsize: SectorUndrawn = True
               If (mapconfig("linedeftypes")(CStr(things(currentselected).effect))("mark5") = 1) And (things(currentselected).arg4 > 0) Then Render_TaggedSectors vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), numsectors, numlinedefs, things(currentselected).arg4, 0, indicatorsize, 0, vertexsize: SectorUndrawn = True
               
               'When sector was undawn, redraw all things
               If (SectorUndrawn) Then
                    
                    'Redraw all things
                    Render_AllThings things(0), 0, numthings - 1, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, Val(Config("allthingsrects")), ViewZoom, filterthings, filtersettings
               Else
                    
                    'Check if things are shown as well
                    If (mapconfig("linedeftypes")(CStr(things(currentselected).effect))("mark1") = 2) And (things(currentselected).arg0 > 0) Then Render_TaggedThingsNormal things(0), numthings - 1, things(currentselected).arg0, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, Val(Config("allthingsrects")), ViewZoom, filterthings, filtersettings
                    If (mapconfig("linedeftypes")(CStr(things(currentselected).effect))("mark2") = 2) And (things(currentselected).arg1 > 0) Then Render_TaggedThingsNormal things(0), numthings - 1, things(currentselected).arg1, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, Val(Config("allthingsrects")), ViewZoom, filterthings, filtersettings
                    If (mapconfig("linedeftypes")(CStr(things(currentselected).effect))("mark3") = 2) And (things(currentselected).arg2 > 0) Then Render_TaggedThingsNormal things(0), numthings - 1, things(currentselected).arg2, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, Val(Config("allthingsrects")), ViewZoom, filterthings, filtersettings
                    If (mapconfig("linedeftypes")(CStr(things(currentselected).effect))("mark4") = 2) And (things(currentselected).arg3 > 0) Then Render_TaggedThingsNormal things(0), numthings - 1, things(currentselected).arg3, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, Val(Config("allthingsrects")), ViewZoom, filterthings, filtersettings
                    If (mapconfig("linedeftypes")(CStr(things(currentselected).effect))("mark5") = 2) And (things(currentselected).arg4 > 0) Then Render_TaggedThingsNormal things(0), numthings - 1, things(currentselected).arg4, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, Val(Config("allthingsrects")), ViewZoom, filterthings, filtersettings
                    
                    'Check if thing has a tag
                    If (things(currentselected).tag <> 0) Then
                                        
                         'Render things that refer to me
                         Render_TaggedArgThingsNormal things(0), numthings, things(currentselected).tag, 2, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, Val(Config("allthingsrects")), ViewZoom, filterthings, filtersettings
                    End If
               End If
          End If
     End If
     
     'Get the nearest thing and select it if within allowed distance
     nearest = NearestThing(X, Y, things(0), numthings, distance, filterthings, filtersettings)
     If (distance <= Config("thingselectdistance") / ViewZoom) Then currentselected = nearest Else currentselected = -1
     
     'Check if selection is made
     If (currentselected > -1) Then
          
          'Check if thing has an action
          If (things(currentselected).effect > 0) Then
               
               'Check if thing has a known effect
               If (mapconfig("linedeftypes").Exists(CStr(things(currentselected).effect)) = True) Then
                    
                    'Render sectors by hexen format
                    If (mapconfig("linedeftypes")(CStr(things(currentselected).effect))("mark1") = 1) And (things(currentselected).arg0 > 0) Then Render_TaggedSectors vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), numsectors, numlinedefs, things(currentselected).arg0, CLR_SECTORTAG, indicatorsize, 0, vertexsize
                    If (mapconfig("linedeftypes")(CStr(things(currentselected).effect))("mark2") = 1) And (things(currentselected).arg1 > 0) Then Render_TaggedSectors vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), numsectors, numlinedefs, things(currentselected).arg1, CLR_SECTORTAG, indicatorsize, 0, vertexsize
                    If (mapconfig("linedeftypes")(CStr(things(currentselected).effect))("mark3") = 1) And (things(currentselected).arg2 > 0) Then Render_TaggedSectors vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), numsectors, numlinedefs, things(currentselected).arg2, CLR_SECTORTAG, indicatorsize, 0, vertexsize
                    If (mapconfig("linedeftypes")(CStr(things(currentselected).effect))("mark4") = 1) And (things(currentselected).arg3 > 0) Then Render_TaggedSectors vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), numsectors, numlinedefs, things(currentselected).arg3, CLR_SECTORTAG, indicatorsize, 0, vertexsize
                    If (mapconfig("linedeftypes")(CStr(things(currentselected).effect))("mark5") = 1) And (things(currentselected).arg4 > 0) Then Render_TaggedSectors vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), numsectors, numlinedefs, things(currentselected).arg4, CLR_SECTORTAG, indicatorsize, 0, vertexsize
                    
                    'Render things by hexen format
                    If (mapconfig("linedeftypes")(CStr(things(currentselected).effect))("mark1") = 2) And (things(currentselected).arg0 > 0) Then Render_TaggedThings things(0), numthings - 1, things(currentselected).arg0, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, Val(Config("allthingsrects")), ViewZoom, filterthings, filtersettings
                    If (mapconfig("linedeftypes")(CStr(things(currentselected).effect))("mark2") = 2) And (things(currentselected).arg1 > 0) Then Render_TaggedThings things(0), numthings - 1, things(currentselected).arg1, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, Val(Config("allthingsrects")), ViewZoom, filterthings, filtersettings
                    If (mapconfig("linedeftypes")(CStr(things(currentselected).effect))("mark3") = 2) And (things(currentselected).arg2 > 0) Then Render_TaggedThings things(0), numthings - 1, things(currentselected).arg2, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, Val(Config("allthingsrects")), ViewZoom, filterthings, filtersettings
                    If (mapconfig("linedeftypes")(CStr(things(currentselected).effect))("mark4") = 2) And (things(currentselected).arg3 > 0) Then Render_TaggedThings things(0), numthings - 1, things(currentselected).arg3, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, Val(Config("allthingsrects")), ViewZoom, filterthings, filtersettings
                    If (mapconfig("linedeftypes")(CStr(things(currentselected).effect))("mark5") = 2) And (things(currentselected).arg4 > 0) Then Render_TaggedThings things(0), numthings - 1, things(currentselected).arg4, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, Val(Config("allthingsrects")), ViewZoom, filterthings, filtersettings
               End If
          End If
          
          'Check if thing has a tag
          If (things(currentselected).tag <> 0) Then
                              
               'Render things that refer to me
               Render_TaggedArgThings things(0), numthings, things(currentselected).tag, 2, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, Val(Config("allthingsrects")), ViewZoom, filterthings, filtersettings
          End If
          
          'Render the new current selected to selected
          If Config("thingrects") Then Render_BoxSwitched things(currentselected).X, things(currentselected).Y, things(currentselected).size * ViewZoom, PAL_THINGSELECTION, (Config("thingrects") - 1), PAL_THINGSELECTION
          Render_Bitmap ThingBitmapData(0), frmMain.picThings(thingsize).width, _
                        frmMain.picThings(thingsize).height, _
                        things(currentselected).image * frmMain.picThings(thingsize).height, 0, _
                        frmMain.picThings(thingsize).height, frmMain.picThings(thingsize).height, _
                        things(currentselected).X, things(currentselected).Y, _
                        CLR_THINGHIGHLIGHT, CLR_BACKGROUND
     End If
     
     'Show the rendered changes
     If (OldSelected <> currentselected) Then picMap.Refresh
     
     'Check if we should show the info
     If (currentselected > -1) Then
          
          'Only update when changed
          If (OldSelected <> currentselected) Then
               
               'Get information
               nType = things(currentselected).thing
               spritename = GetThingTypeSpriteName(things(currentselected).thing)
               tType = GetThingTypeDesc(things(currentselected).thing) & " (" & things(currentselected).thing & ")"
               angle = GetThingAngleDesc(things(currentselected).angle) & " (" & things(currentselected).angle & ")"
               Flags = Hexadecimal(things(currentselected).Flags, 4) & " (" & things(currentselected).Flags & ")"
               tx = things(currentselected).X
               ty = things(currentselected).Y
               sTag = things(currentselected).tag
               
               'Check if Thing has an action
               If (things(currentselected).effect > 0) Then
                    
                    'Check if the linedef type can be found
                    If (mapconfig("linedeftypes").Exists(CStr(things(currentselected).effect))) Then
                         sAction = things(currentselected).effect & " - " & Trim$(mapconfig("linedeftypes")(CStr(things(currentselected).effect))("title"))
                    Else
                         sAction = things(currentselected).effect & " - Unknown"
                    End If
               Else
                    sAction = "0 - None"
               End If
               
               'Check what panel to show the info on
               If picBar.Visible Then
                    
                    'Check if the bar is fully shown
                    If (cmdToggleBar.tag = "0") Then
                         
                         'Set the info
                         fraThing.Caption = " Thing " & currentselected & " "
                         lblThingType = ShortedText(tType, lblThingType.width \ Screen.TwipsPerPixelX)
                         lblThingAction.Caption = ShortedText(sAction, lblThingAction.width \ Screen.TwipsPerPixelX)
                         lblThingAngle = ShortedText(angle, lblThingAngle.width \ Screen.TwipsPerPixelX)
                         lblThingFlags = Flags
                         lblThingTag.Caption = sTag
                         lblThingXY = tx & ", " & ty
                         lblThing.Caption = spritename
                         GetScaledSpritePicture nType, imgThing, picThing.ScaleWidth, picThing.ScaleHeight
                         
                         'Show the info
                         fraThing.Visible = True
                         fraThingPreview.Visible = True
                    Else
                         
                         'Show single line only
                         lblBarText.Caption = " Thing " & currentselected & ":  " & tType
                    End If
                    
               ElseIf picSBar.Visible Then
                    
                    'Check if the bar is fully shown
                    If (cmdToggleBar.tag = "0") Then
                         
                         'Set the info
                         fraSThing.Caption = " Thing " & currentselected & " "
                         lblSThingType = ShortedText(tType, lblSThingType.width \ Screen.TwipsPerPixelX)
                         lblSThingAction.Caption = ShortedText(sAction, lblSThingAction.width \ Screen.TwipsPerPixelX)
                         lblSThingAngle = ShortedText(angle, lblSThingAngle.width \ Screen.TwipsPerPixelX)
                         lblSThingFlags = Flags
                         lblSThingTag.Caption = sTag
                         lblSThingXY = tx & ", " & ty
                         lblSThing.Caption = spritename
                         GetScaledSpritePicture nType, imgSThing, picSThing.ScaleWidth, picSThing.ScaleHeight
                         
                         'Show the info
                         fraSThing.Visible = True
                         fraSThingPreview.Visible = True
                    Else
                         
                         'Show single line only
                         lblBarText.Caption = " Thing " & currentselected & ":  " & tType
                    End If
               End If
               
               'Type on tooltip
               picMap.ToolTipText = ""
               If (Val(Config("showtooltips")) <> 0) Then picMap.ToolTipText = tType
          End If
     Else
          
          'Hide the info
          HideThingInfo
     End If
End Sub

Private Sub ChangeVertexHighlight(ByVal X As Long, ByVal Y As Long, Optional ByVal Forceupdate As Boolean)
     Dim distance As Long
     Dim nearest As Long
     Dim OldSelected As Long
     
     Dim vx As Long, vy As Long
     
     'Keep old selection ifnot forced to update
     If (Not Forceupdate) Then OldSelected = currentselected Else OldSelected = -1
     
     'Render the last current selected to normal
     If (currentselected > -1) Then Render_AllVertices vertexes(0), currentselected, currentselected, vertexsize
     
     'Get the nearest vertex and select it if within allowed distance
     nearest = NearestVertex(X, Y, vertexes(0), numvertexes, distance)
     If (distance <= Config("vertexselectdistance") / ViewZoom) Then currentselected = nearest Else currentselected = -1
     
     'Render the new current selected to selected
     If (currentselected > -1) Then Render_Box vertexes(currentselected).X, vertexes(currentselected).Y, vertexsize, CLR_VERTEXHIGHLIGHT, 1, CLR_VERTEXHIGHLIGHT
     
     'Show the rendered changes
     If (OldSelected <> currentselected) Then picMap.Refresh
     
     
     'Check if anything is selected
     If (currentselected > -1) Then
          
          'Only update when changed
          If (OldSelected <> currentselected) Then
               
               'Get information
               vx = vertexes(currentselected).X
               vy = vertexes(currentselected).Y
               
               'Coordiantes on Tooltip
               picMap.ToolTipText = ""
               If (Val(Config("showtooltips")) <> 0) Then picMap.ToolTipText = vx & ", " & vy
               
               'Check what panel to show info on
               If picBar.Visible Then
                    
                    'Check if the bar is fully shown
                    If (cmdToggleBar.tag = "0") Then
                         
                         'Set the info
                         fraVertex.Caption = " Vertex " & currentselected & " "
                         lblVertexXY.Caption = vx & ", " & vy
                         
                         'Show the info
                         fraVertex.Visible = True
                    Else
                         
                         'Single line only
                         lblBarText.Caption = " Vertex " & currentselected & ":  " & vx & ", " & vy
                    End If
               ElseIf picSBar.Visible Then
                    
                    'Check if the bar is fully shown
                    If (cmdToggleBar.tag = "0") Then
                         
                         'Set the info
                         fraSVertex.Caption = " Vertex " & currentselected & " "
                         lblSVertexXY.Caption = vx & ", " & vy
                         
                         'Show the info
                         fraSVertex.Visible = True
                    Else
                         
                         'Single line only
                         lblBarText.Caption = " Vertex " & currentselected & ":  " & vx & ", " & vy
                    End If
               End If
          End If
     Else
          
          'Hide the info
          HideVertexInfo
     End If
End Sub

Private Sub CreateSectorHere(ByVal X As Long, ByVal Y As Long)
     Dim distance As Single
     Dim v As Long
     Dim direction As Single
     Dim vx As Single, vy As Single
     Dim v0 As Long, v1 As Long, v2 As Long, v3 As Long
     Dim ss As Long, ns As Long
     Dim ld As Long
     
     'Remove higlight
     RemoveHighlight True
     
     'No more selection
     ResetSelections things(0), numthings, linedefs(0), numlinedefs, vertexes(0), numvertexes, VarPtr(sectors(0)), numsectors
     Set selected = New Dictionary
     numselected = 0
     selectedtype = EM_LINES
     
     'Check if we should snap the center to grid
     If snapmode Then
          
          'Snap X and Y to grid
          X = SnappedToGridX(X)
          Y = SnappedToGridY(Y)
     End If
     
     'Show indicator
     Render_Line X - 10 / ViewZoom, -Y, X + 10 / ViewZoom, -Y, CLR_MULTISELECT
     Render_Line X, -Y - 10 / ViewZoom, X, -Y + 10 / ViewZoom, CLR_MULTISELECT
     picMap.Refresh
     
     'Load the create sector dialog
     Load frmMakeSector
     
     'Set snap option default
     If snapmode Then frmMakeSector.chkSnap.Value = vbChecked
     
     'Show the dialog
     frmMakeSector.Show 1, Me
     
     'Check if not cancelled
     If (frmMakeSector.tag = "1") Then
          
          'Change mousepointer
          Screen.MousePointer = vbHourglass
          
          'This is drawing mode
          submode = ESM_DRAWING
          
          'Get diameter
          distance = frmMakeSector.scrDiameter.Value - 0.1
          
          'Make undo
          CreateUndo "sector insert"
          
          'Determine the sector in which this new sector will be created
          'This will be -1 for no surrounding sector
          ss = IntersectSector(X, Y, vertexes(0), linedefs(0), VarPtr(sidedefs(0)), numlinedefs, 0)
          
          'Create new sector
          ns = CreateSector
          
          'Check if we should copy properties from outer sector
          If (ss > -1) Then
               
               'Copy sector properties
               sectors(ns) = sectors(ss)
          Else
               
               'Set new properties
               With sectors(ns)
                    .Brightness = Config("defaultsector")("brightness")
                    .hceiling = Config("defaultsector")("hceiling")
                    .hfloor = Config("defaultsector")("hfloor")
                    .tceiling = UCase$(Config("defaultsector")("tceiling"))
                    .tfloor = UCase$(Config("defaultsector")("tfloor"))
                    .selected = 0
                    .special = 0
                    .tag = 0
               End With
          End If
          
          'New selection
          Set selected = New Dictionary
          numselected = frmMakeSector.scrVertices.Value
          
          'Go for the number of vertices
          For v = 1 To frmMakeSector.scrVertices.Value
               
               'Check if creating a square
               If (frmMakeSector.scrVertices.Value = 4) Then
                    
                    'Check what corner to build
                    Select Case v
                         
                         Case 1:   'Left top
                              vx = X - distance
                              vy = Y - distance
                              
                         Case 2:   'Right top
                              vx = X + distance
                              vy = Y - distance
                              
                         Case 3:   'Right bottom
                              vx = X + distance
                              vy = Y + distance
                              
                         Case 4:   'Left bottom
                              vx = X - distance
                              vy = Y + distance
                              
                    End Select
                    
               'Otherwise creating a circle
               Else
                    
                    'Calculate direction from center
                    direction = pi * (2 - ((2 / frmMakeSector.scrVertices.Value) * v))
                    
                    'Calculate exact X and Y
                    vx = X + sIn(direction) * distance
                    vy = Y + Cos(direction) * distance
               End If
               
               'Check if we should snap to grid
               If (frmMakeSector.chkSnap.Value = vbChecked) Then
                    
                    'Snap X and Y to grid
                    vx = SnappedToGridX(vx)
                    vy = SnappedToGridY(vy)
               End If
               
               'Keep previous vertex
               v1 = v2
               
               'Create new vertex here
               v2 = InsertVertex(vx, -vy)
               
               'Select the vertex
               vertexes(v2).selected = 1
               selected.Add CStr(v2), v2
               
               'Keep first and last vertex
               If (v = 1) Then v0 = v2
               If (v = frmMakeSector.scrVertices.Value) Then v3 = v2
               
               'Check if this is not the first vertex
               If (v > 1) Then
                    
                    'Create a linedef
                    ld = CreateLinedef
                    With linedefs(ld)
                         .v1 = v1
                         .v2 = v2
                         .tag = 0
                         .selected = 1
                         .effect = 0
                         
                         .s1 = CreateSidedef
                         With sidedefs(.s1)
                              .tx = 0
                              .ty = 0
                              If (ss > -1) Then .Upper = UCase$(Config("defaulttexture")("upper")) Else .Upper = "-"
                              If (ss = -1) Then .Middle = UCase$(Config("defaulttexture")("middle")) Else .Middle = "-"
                              If (ss > -1) Then .Lower = UCase$(Config("defaulttexture")("lower")) Else .Lower = "-"
                              .linedef = ld
                              .sector = ns
                         End With
                         
                         'Create a second sidedef if within a sector
                         If (ss > -1) Then
                              
                              'Make double sided
                              .Flags = LDF_TWOSIDED
                              
                              .s2 = CreateSidedef
                              With sidedefs(.s2)
                                   .tx = 0
                                   .ty = 0
                                   .Upper = UCase$(Config("defaulttexture")("upper"))
                                   .Middle = "-"
                                   .Lower = UCase$(Config("defaulttexture")("lower"))
                                   .linedef = ld
                                   .sector = ss
                              End With
                              
                         'Otherwise,
                         Else
                              
                              'Make it impassable
                              .Flags = LDF_IMPASSIBLE
                              .s2 = -1
                         End If
                    End With
                    
                    'Remove any unneeded textures
                    If (linedefs(ld).s1 > -1) Then RemoveUnusedSidedefTextures linedefs(ld).s1
                    If (linedefs(ld).s2 > -1) Then RemoveUnusedSidedefTextures linedefs(ld).s2
               End If
          Next v
          
          'Create a closing linedef from last to first vertex
          ld = CreateLinedef
          With linedefs(ld)
               .v1 = v3
               .v2 = v0
               .tag = 0
               .selected = 1
               .effect = 0
               
               .s1 = CreateSidedef
               With sidedefs(.s1)
                    .tx = 0
                    .ty = 0
                    If (ss > -1) Then .Upper = UCase$(Config("defaulttexture")("upper")) Else .Upper = "-"
                    If (ss = -1) Then .Middle = UCase$(Config("defaulttexture")("middle")) Else .Middle = "-"
                    If (ss > -1) Then .Lower = UCase$(Config("defaulttexture")("lower")) Else .Lower = "-"
                    .linedef = ld
                    .sector = ns
               End With
               
               'Create a second sidedef if within a sector
               If (ss > -1) Then
                    
                    'Make double sided
                    .Flags = LDF_TWOSIDED
                    
                    .s2 = CreateSidedef
                    With sidedefs(.s2)
                         .tx = 0
                         .ty = 0
                         .Upper = UCase$(Config("defaulttexture")("upper"))
                         .Middle = "-"
                         .Lower = UCase$(Config("defaulttexture")("lower"))
                         .linedef = ld
                         .sector = ss
                    End With
                    
               'Otherwise,
               Else
                    
                    'Make it impassable
                    .Flags = LDF_IMPASSIBLE
                    .s2 = -1
               End If
          End With
          
          'Select lines from vertices
          SelectLinedefsFromVertices
          
          'Setup the new sector
          NewSectorSetup True
          
          'DEBUG
          'DEBUG_FindUnusedSectors
          
          'Deselect all
          RemoveSelection False
          
          'No more changes lines
          ReDim changedlines(0)
          numchangedlines = 0
          
          'No more drag selection
          Set dragselected = New Dictionary
          dragnumselected = 0
          
          'Map has changed
          mapnodeschanged = True
          mapchanged = True
          
          'Normal mode
          submode = ESM_NONE
          
          'Reset mousepointer
          Screen.MousePointer = vbNormal
     End If
     
     'Show changes
     RedrawMap False
     
     'Unload dialog
     Unload frmMakeSector
     Set frmMakeSector = Nothing
End Sub

Private Sub CreateSubclassing()
     
     'Check if we are allowed to do subclassing
     If (CommandSwitch("-nosubclass") = False) Then
          
          'Keep original messages handler
          OriginalMessageHandler = GetWindowLong(Me.hWnd, GWL_WNDPROC)
          
          'Set our own messages handler
          SetWindowLong Me.hWnd, GWL_WNDPROC, AddressOf MainMessageHandler
     End If
End Sub

Private Sub DestroySubclassing()
     
     'Check if we are allowed to do subclassing
     If (CommandSwitch("-nosubclass") = False) Then
          
          'Restore original messages handler
          SetWindowLong Me.hWnd, GWL_WNDPROC, OriginalMessageHandler
     End If
End Sub

Private Sub DragSelection(ByVal Shift As ShiftConstants, ByVal X As Single, ByVal Y As Single)
     Dim i As Long, s As Long
     Dim DragSel As Variant
     Dim ox As Single, oy As Single
     
     'Check if dragging vertices
     If ((mode = EM_LINES) Or (mode = EM_SECTORS) Or (mode = EM_VERTICES)) Then
          
          'Go for dragging selected vertices
          DragSel = dragselected.Items
          For i = LBound(DragSel) To UBound(DragSel)
               
               'Get vertex
               s = DragSel(i)
               
               'Do not move the grabbed vertex yet
               If (s <> grabobject) Then
                    
                    'Calculate offset from this vertex to grabbed vertex
                    ox = vertexes(s).X - vertexes(grabobject).X
                    oy = vertexes(s).Y - vertexes(grabobject).Y
                    
                    'Move vertex
                    If (snapmode) And ((Shift And vbShiftMask) = 0) Then
                         vertexes(s).X = SnappedToGridX(X - GrabX) + ox
                         vertexes(s).Y = SnappedToGridY(-Y - GrabY) + oy
                    Else
                         vertexes(s).X = (X - GrabX) + ox
                         vertexes(s).Y = (-Y - GrabY) + oy
                    End If
               End If
          Next i
          
          'Move the grabbed vertex
          If (snapmode) And ((Shift And vbShiftMask) = 0) Then
               vertexes(grabobject).X = SnappedToGridX(X - GrabX)
               vertexes(grabobject).Y = SnappedToGridY(-Y - GrabY)
          Else
               vertexes(grabobject).X = X - GrabX
               vertexes(grabobject).Y = -Y - GrabY
          End If
     Else
          
          'Go for dragging selected things
          DragSel = dragselected.Items
          For i = LBound(DragSel) To UBound(DragSel)
               
               'Get thing
               s = DragSel(i)
               
               'Do not move the grabbed thing yet
               If (s <> grabobject) Then
                    
                    'Calculate offset from this thing to grabbed thing
                    ox = things(s).X - things(grabobject).X
                    oy = things(s).Y - things(grabobject).Y
                    
                    'Move thing
                    If (snapmode) And ((Shift And vbShiftMask) = 0) Then
                         things(s).X = SnappedToGridX(X - GrabX) + ox
                         things(s).Y = SnappedToGridY(-Y - GrabY) + oy
                    Else
                         things(s).X = (X - GrabX) + ox
                         things(s).Y = (-Y - GrabY) + oy
                    End If
                    
                    'Check if this is the 3D start position
                    If (things(s).thing = mapconfig("start3dmode")) Then ApplyPositionFromThing s
               End If
          Next i
          
          'Move the grabbed thing
          If (snapmode) And ((Shift And vbShiftMask) = 0) Then
               things(grabobject).X = SnappedToGridX(X - GrabX)
               things(grabobject).Y = SnappedToGridY(-Y - GrabY)
          Else
               things(grabobject).X = X - GrabX
               things(grabobject).Y = -Y - GrabY
          End If
          
          'Check if this is the 3D start position
          If (things(grabobject).thing = mapconfig("start3dmode")) Then ApplyPositionFromThing grabobject
     End If
     
     'Redraw map
     RedrawMap
End Sub

Private Sub DrawVertexHere(ByVal X As Long, ByVal Y As Long, Optional ByVal RefreshMap As Boolean = True)
     Dim vdistance As Long
     Dim ldistance As Long
     Dim ld As Long
     Dim nv As Long
     Dim nrv As Long
     
     'Snap X and Y if snap mode is on
     If snapmode Then
          X = SnappedToGridX(X)
          Y = SnappedToGridY(Y)
     End If
     
     'Record drawing coordinates
     ReDim Preserve DrawingCoords(0 To NumDrawingCoords)
     DrawingCoords(NumDrawingCoords).X = X
     DrawingCoords(NumDrawingCoords).Y = Y
     NumDrawingCoords = NumDrawingCoords + 1
     
     'No vertex inserted yet
     nv = -1
     
     'Get the nearest vertex
     nrv = NearestVertex(X, Y, vertexes(0), numvertexes, vdistance)
     
     'Check if vertex found
     If (nrv > -1) Then
          
          'Check if we should auto-stitch vertices
          If (stitchmode) Or (vertexes(nrv).selected <> 0) Then
               
               'Use this vertex instead of a new one if close enough for stitching
               If (vdistance <= Config("autostitchdistance")) Then nv = nrv
          End If
     End If
     
     'Check if we should insert a new vertex
     If (nv = -1) Then
          
          'Insert a vertex now
          nv = InsertVertex(X, -Y)
          
          'Get the nearest linedef
          ld = NearestLinedef(X, Y, vertexes(0), linedefs(0), numlinedefs, ldistance)
          
          'Check if distance is close enough for linedef split
          If (ldistance <= Config("linesplitdistance")) Then
               
               'Split the linedef with this vertex
               SplitLinedef ld, nv
               
               'Redraw the map
               If (RefreshMap) Then RedrawMap False
          End If
     End If
     
     'Select the vertex
     vertexes(nv).selected = 1
     
     'Draw the vertex
     If (RefreshMap) Then Render_AllVertices vertexes(0), nv, nv, vertexsize
     
     'Check if a vertex was previously drawn
     If (numselected > 0) Then
          
          'Insert a linedef
          ld = CreateLinedef
          
          'Set the linedef properties
          With linedefs(ld)
               
               'From previous vertex, to this vertex
               .v1 = selected.Items(selected.Count - 1)
               .v2 = nv
               
               'Set other line properties
               .Flags = LDF_IMPASSIBLE
               .effect = 0
               .tag = 0
               .selected = 1
               .s1 = -1
               .s2 = -1
          End With
          
          'Add to changing lines
          ReDim Preserve changedlines(0 To numchangedlines)
          changedlines(numchangedlines) = ld
          numchangedlines = numchangedlines + 1
          
          'Update status
          UpdateStatusBar
          
          'Draw the linedef
          If (RefreshMap) Then
               Render_AllLinedefs vertexes(0), linedefs(0), ld, ld, submode, indicatorsize
               Render_AllVertices vertexes(0), linedefs(ld).v1, linedefs(ld).v1, vertexsize
               Render_AllVertices vertexes(0), linedefs(ld).v2, linedefs(ld).v2, vertexsize
               
               'Render changing linedef lengths
               Render_ChangingLengths vertexes(0), linedefs(0), changedlines(0), numchangedlines, NumbersBitmapData(0), frmMain.picNumbers.width, frmMain.picNumbers.height, frmMain.picNumbers.width / 10, frmMain.picNumbers.height
          End If
          
          'Check if the new vertex is same as any drawn vertex
          If (selected.Exists(CStr(nv)) = True) Then
               
               'Refresh
               If (RefreshMap) Then picMap.Refresh
               
               'End drawing mode
               EndDrawOperation True
               
               'Leave now
               Exit Sub
          End If
     Else
          
          'Update status only (for the new vertex)
          UpdateStatusBar
     End If
     
     'Add vertex to selection list if not already in there
     If (selected.Exists(CStr(nv)) = False) Then
          selected.Add CStr(nv), nv
          numselected = numselected + 1
     End If
End Sub

Private Sub EndDragOperation()
     Dim Result As Boolean
     Dim DraggedSelection As Dictionary
     Dim DraggedSelectionType As ENUM_EDITMODE
     
     'Change mousepointer
     Screen.MousePointer = vbHourglass
     
     'Check if we were dragging vertices
     If ((mode = EM_LINES) Or (mode = EM_SECTORS) Or (mode = EM_VERTICES)) Then
          
          'Round all vertices
          RoundVertices vertexes(0), numvertexes
          
          'When user was pasting, replace -1 sectors with their parent
          If (submode = ESM_PASTING) Then ApplyParentSectors
          
          'Check if sector heights must be adjusted
          If (submode = ESM_PASTING) And (PrefabAdjustHeights = True) And _
             ((mode = EM_LINES) Or (mode = EM_SECTORS)) Then
               
               'Check if user wants sector heights to be adjusted
               If (Val(Config("pasteadjustsheights"))) Then
                    
                    'Check if selection must be converted
                    If (mode = EM_LINES) Then
                         
                         'Keep original selection
                         Set DraggedSelection = selected
                         DraggedSelectionType = selectedtype
                         
                         'Select dragged sectors only
                         SelectSectorsFromLinedefs True
                    End If
                    
                    'Apply heights adjustments
                    ApplySectorHeightAdjustments
                    
                    'Check if selection must be restored
                    If (mode = EM_LINES) Then
                         
                         'Restore original selection
                         Set selected = DraggedSelection
                         numselected = selected.Count
                         selectedtype = DraggedSelectionType
                    End If
               End If
          End If
          
          'Check if we should auto-stitch vertices
          If (stitchmode) Then
               
               'Make undo
               If (Config("subundos") = vbChecked) Then CreateUndo "stitch vertices"
               
               'Find all changing lines
               FindChangingLines True, True
               
               'Split linedefs with overlapping vertices
               Result = SelectedVerticesSplitLinedefs
               
               'Auto-stich vertices
               Result = Result Or AutoStitchDraggedSelection
               
               'Check result
               If (Result) Then
                    
                    'Remove looped linedefs
                    RemoveLoopedLinedefs
                    
                    'Find all changing lines
                    FindChangingLines True, True
                    
                    'Due to auto-stitch, linedefs could be overlapping
                    'Combine these into one now
                    MergeDoubleLinedefs
               Else
                    
                    'No stitching was done, no undo needed
                    If (Config("subundos") = vbChecked) Then WithdrawUndo
               End If
          End If
          
          'Fix glitches that other routines may cause
          SolveGlitches
          
          'When dragging lines or sector, deselect all vertices
          'that were temporarely seleced for dragging
          ResetSelections things(0), 0, linedefs(0), 0, vertexes(0), numvertexes, VarPtr(sectors(0)), 0
          If (mode = EM_VERTICES) Then ReapplyVerticesSelection
          
          'DEBUG
          'DEBUG_FindUnusedSectors
     End If
     
     'Check if we should deselect
     If (DeselectAfterEdit) Then RemoveSelection False
     
     'End of drag operation, set mode back to normal
     submode = ESM_NONE
     
     'Map was changed
     mapnodeschanged = True
     mapchanged = True
     
     'We dont need these anymore
     ReDim changedlines(0)
     numchangedlines = 0
     
     'Redraw the map
     RedrawMap
     
     'Show highlight
     ShowHighlight LastX, LastY
     
     'Update status
     UpdateStatusBar
     
     'Reset mousepointer
     Screen.MousePointer = vbNormal
End Sub

Private Sub EndDrawOperation(ByVal NewSector As Boolean)
     Dim sc1 As Long
     Dim sc2 As Long
     Dim ns As Long, os As Long
     Dim selectsectorindex As Long
     Dim isrelatedline As Long
     Dim ld As Long
     Dim sx As Single, sy As Single
     Dim Indices As Variant
     Dim i As Long
     Dim CopyFromSidedef As Long
     
     'Change mousepointer
     Screen.MousePointer = vbHourglass
     
     'Go for all lines to make a new selection list
     'so the selection list reflects the drawn lines
     Set selected = New Dictionary
     For ld = 0 To (numlinedefs - 1)
          
          'Check if selected
          If (linedefs(ld).selected <> 0) Then
               
               'Normal selection
               linedefs(ld).selected = 1
               
               'Add to list
               selected.Add CStr(ld), ld
          End If
     Next ld
     
     'Count selected linedefs
     numselected = selected.Count
     selectedtype = EM_LINES
     
     'Do not copy properties from
     'any sidedef/sector yet
     CopyFromSidedef = -1
     
     'Check if any line were drawn
     If (numselected > 0) Then
          
          'Get spot point on front side of the line
          GetLineSideSpot selected.Items(0), 1, True, sx, sy
          
          'Get the sector in which was drawn
          os = IntersectSector(sx, sy, vertexes(0), linedefs(0), VarPtr(sidedefs(0)), numlinedefs, 1)
          sc2 = os
          
          'Check if we should create a sector
          If (NewSector) Then
               
               'Check if drawn counterclockwise
               If Not point_in_polygon(selected.Items, selected.Count, sx, -sy) Then
                    
                    'Get a spot on the other side to ensure having a spot inside
                    GetLineSideSpot selected.Items(0), 1, False, sx, sy
                    
                    'Check if a parent sector exists
                    If (sc2 > -1) Then
                         
                         'sc1 will be the parent sector
                         sc1 = sc2
                         
                         'Make new sector at sc2
                         ns = CreateSector
                         sc2 = ns
                         
                         'Copy sector properties
                         sectors(sc2) = sectors(sc1)
                    Else
                         
                         'Go for all selected linedefs to flip
                         Indices = selected.Items
                         For ld = 0 To (numselected - 1)
                              
                              'Flip this linedef
                              FlipLinedefVertices Indices(ld)
                         Next ld
                         
                         'Make new sector at sc1
                         ns = CreateSector
                         sc1 = ns
                    End If
               Else
                    
                    'Make sector at sc1
                    ns = CreateSector
                    sc1 = ns
               End If
               
               'Select the new sector for editing
               selectsectorindex = ns
               
               'Check if we should copy properties from outer sector
               If (sc2 > -1) Then
                    
                    'Copy sector properties
                    sectors(sc1) = sectors(sc2)
               Else
                    
                    'Get index of first adjoining sidedef if any
                    CopyFromSidedef = FindAdjoiningSidedef(sx, sy)
                    
                    'Check if there is an adjacent line to copy properties from
                    If (CopyFromSidedef > -1) Then
                         
                         'Copy sector properties from adjoining sector
                         sectors(sc1) = sectors(sidedefs(CopyFromSidedef).sector)
                         
                         'Check if we should erase tag and actions
                         If (Config("copytagdraw") = vbUnchecked) Then
                              With sectors(sc1)
                                   .special = 0
                                   .tag = 0
                              End With
                         End If
                    Else
                         
                         'Set sector properties to defaults
                         With sectors(sc1)
                              .Brightness = Config("defaultsector")("brightness")
                              .hceiling = Config("defaultsector")("hceiling")
                              .hfloor = Config("defaultsector")("hfloor")
                              .tceiling = UCase$(Config("defaultsector")("tceiling"))
                              .tfloor = UCase$(Config("defaultsector")("tfloor"))
                              .selected = 0
                              .special = 0
                              .tag = 0
                         End With
                    End If
               End If
               
               'Setup the new lines
               SetupNewLinedefs sc1, sc2, CopyFromSidedef
               
               'All other lines (unselected) referring to the old sector must
               'be tested if inside new polygon and changed as needed
               
               'Check if we should auto-stitch vertices
               If (stitchmode) Then
                    
                    'Make dragged selection same as current selection
                    Set dragselected = selected
                    dragnumselected = numselected
                    
                    'Find and keep lines that have changed (added)
                    FindChangingLines True, True
                    
                    'Due to auto-stitch, linedefs could be overlapping
                    'Combine these into one now
                    MergeDoubleLinedefs
               End If
               
               'Go for all path linedefs
               For ld = 0 To (numlinedefs - 1)
                    
                    'Check if unselected
                    If (linedefs(ld).selected = 0) Then
                         
                         'Check sectors of the line
                         isrelatedline = False
                         If (linedefs(ld).s1 > -1) Then isrelatedline = (sidedefs(linedefs(ld).s1).sector = os) Else isrelatedline = (os = -1)
                         If (linedefs(ld).s2 > -1) Then isrelatedline = isrelatedline Or (sidedefs(linedefs(ld).s2).sector = os) Else isrelatedline = isrelatedline Or (os = -1)
                         
                         'Check if line is related to old sector
                         If isrelatedline Then
                              
                              'Get spot point on middle of the line
                              GetLineSideSpot ld, 0, True, sx, sy
                              
                              'Check if inside new sector
                              If point_in_polygon(selected.Items, numselected, sx, -sy) Then
                                   
                                   'Check if line has a front side
                                   If (linedefs(ld).s1 > -1) Then
                                        
                                        'Change sector if this refers to the old sector
                                        If (sidedefs(linedefs(ld).s1).sector = os) Then sidedefs(linedefs(ld).s1).sector = ns
                                        
                                   'Check if old sector is void
                                   ElseIf (os = -1) Then
                                        
                                        'Add sidedef
                                        linedefs(ld).s1 = CreateSidedef
                                        
                                        'Apply building defaults on sidedef
                                        With sidedefs(linedefs(ld).s1)
                                             .linedef = ld
                                             .Lower = UCase$(Config("defaulttexture")("lower"))
                                             If (linedefs(ld).s2 = -1) Then .Middle = UCase$(Config("defaulttexture")("middle")) Else .Middle = "-"
                                             .Upper = UCase$(Config("defaulttexture")("upper"))
                                             .sector = ns
                                             .tx = 0
                                             .ty = 0
                                        End With
                                        
                                        'Line becomes doublesided?
                                        If (linedefs(ld).s2 > -1) Then
                                             sidedefs(linedefs(ld).s2).Middle = "-"
                                             linedefs(ld).Flags = linedefs(ld).Flags Or LDF_TWOSIDED
                                             linedefs(ld).Flags = linedefs(ld).Flags And Not LDF_IMPASSIBLE
                                        End If
                                   End If
                                   
                                   'Check if line has a back side
                                   If (linedefs(ld).s2 > -1) Then
                                        
                                        'Change sector if this refers to the old sector
                                        If (sidedefs(linedefs(ld).s2).sector = os) Then sidedefs(linedefs(ld).s2).sector = ns
                                        
                                   'Check if old sector is void
                                   ElseIf (os = -1) Then
                                        
                                        'Add sidedef
                                        linedefs(ld).s2 = CreateSidedef
                                        
                                        'Apply building defaults on sidedef
                                        With sidedefs(linedefs(ld).s2)
                                             .linedef = ld
                                             .Lower = UCase$(Config("defaulttexture")("lower"))
                                             If (linedefs(ld).s1 = -1) Then .Middle = UCase$(Config("defaulttexture")("middle")) Else .Middle = "-"
                                             .Upper = UCase$(Config("defaulttexture")("upper"))
                                             .sector = ns
                                             .tx = 0
                                             .ty = 0
                                        End With
                                        
                                        'Line becomes doublesided?
                                        If (linedefs(ld).s1 > -1) Then
                                             sidedefs(linedefs(ld).s1).Middle = "-"
                                             linedefs(ld).Flags = linedefs(ld).Flags Or LDF_TWOSIDED
                                             linedefs(ld).Flags = linedefs(ld).Flags And Not LDF_IMPASSIBLE
                                        End If
                                   End If
                              End If
                         End If
                    End If
               Next ld
          Else
               
               'Check if inside a sector
               If (sc2 > -1) Then
                    
                    'Do a trace from one vertex to the other
                    ReDim SectorSplitLinesList(0)
                    SectorSplitNumLines = 0
                    TerminateRecursion = False
                    TraceSectorSplitVertex linedefs(selected.Items(0)).v1, linedefs(selected.Items(numselected - 1)).v2, sc2, SectorSplitLinesList(), 0
                    
                    'Remove line selection properties
                    ResetSelections things(0), 0, linedefs(0), numlinedefs, vertexes(0), 0, VarPtr(sectors(0)), 0
                    
                    'Check if a path was found
                    If (SectorSplitNumLines > 0) Then
                         
                         'We create a new sector
                         NewSector = True
                         
                         'Allocate memory for drawn lines into this array
                         ReDim Preserve SectorSplitLinesList(0 To SectorSplitNumLines + numselected - 1)
                         
                         'Go for all drawn linedefs
                         Indices = selected.Items
                         For i = 0 To (numselected - 1)
                              
                              'Add to the array to create a closed polygon
                              SectorSplitLinesList(SectorSplitNumLines + i) = Indices(i)
                              
                              'Select linedef
                              linedefs(Indices(i)).selected = 1
                         Next i
                         
                         'Check what side of line will be the new sector
                         If point_in_polygon(SectorSplitLinesList(), SectorSplitNumLines + numselected, sx, -sy) Then
                              
                              'Create new sector on front side
                              ns = CreateSector
                              sc1 = ns
                              
                              'Select the new sector for editing
                              selectsectorindex = ns
                              
                              'Copy sector properties
                              sectors(sc1) = sectors(sc2)
                         Else
                              
                              'sc1 will be the parent sector
                              sc1 = sc2
                              
                              'Create new sector on back side
                              ns = CreateSector
                              sc2 = ns
                              
                              'Select the old sector for editing
                              selectsectorindex = os
                              
                              'Copy sector properties
                              sectors(sc2) = sectors(sc1)
                         End If
                         
                         'All path lines referring to the old sector must
                         'now refer to the new created sector
                         
                         'Go for all path linedefs
                         For i = 0 To (SectorSplitNumLines - 1)
                              
                              'Get linedef
                              ld = SectorSplitLinesList(i)
                              
                              'Select linedef
                              linedefs(ld).selected = 1
                              
                              'Get spot point on front side of the line
                              GetLineSideSpot ld, 1, True, sx, sy
                              
                              'Check what side of the line is in the new polygon
                              If point_in_polygon(SectorSplitLinesList(), SectorSplitNumLines + numselected, sx, -sy) Then
                                   
                                   'Check if line has a front side
                                   If (linedefs(ld).s1 > -1) Then
                                        
                                        'Change sector if this refers to the old sector
                                        If (sidedefs(linedefs(ld).s1).sector = os) Then sidedefs(linedefs(ld).s1).sector = ns
                                   End If
                              Else
                                   
                                   'Check if line has a back side
                                   If (linedefs(ld).s2 > -1) Then
                                        
                                        'Change sector if this refers to the old sector
                                        If (sidedefs(linedefs(ld).s2).sector = os) Then sidedefs(linedefs(ld).s2).sector = ns
                                   End If
                              End If
                         Next i
                         
                         'Setup the new lines
                         SetupNewLinedefs sc1, sc2, CopyFromSidedef
                         
                         'All other lines (unselected) referring to the old sector must
                         'be tested if inside new polygon and changed as needed
                         
                         'Check if we should auto-stitch vertices
                         If (stitchmode) Then
                              
                              'Make dragged selection same as current selection
                              Set dragselected = selected
                              dragnumselected = numselected
                              
                              'Find and keep lines that have changed (added)
                              FindChangingLines True, True
                              
                              'Due to auto-stitch, linedefs could be overlapping
                              'Combine these into one now
                              MergeDoubleLinedefs
                         End If
                         
                         'Go for all path linedefs
                         For ld = 0 To (numlinedefs - 1)
                              
                              'Check if unselected
                              If (linedefs(ld).selected = 0) Then
                                   
                                   'Check sectors of the line
                                   isrelatedline = False
                                   If (linedefs(ld).s1 > -1) Then isrelatedline = (sidedefs(linedefs(ld).s1).sector = os)
                                   If (linedefs(ld).s2 > -1) Then isrelatedline = isrelatedline Or (sidedefs(linedefs(ld).s2).sector = os)
                                   
                                   'Check if line is related to old sector
                                   If isrelatedline Then
                                        
                                        'Get spot point on middle of the line
                                        GetLineSideSpot ld, 0, True, sx, sy
                                        
                                        'Check if inside new sector
                                        If point_in_polygon(SectorSplitLinesList(), SectorSplitNumLines + numselected, sx, -sy) Then
                                             
                                             'Check if line has a front side
                                             If (linedefs(ld).s1 > -1) Then
                                                  
                                                  'Change sector if this refers to the old sector
                                                  If (sidedefs(linedefs(ld).s1).sector = os) Then sidedefs(linedefs(ld).s1).sector = ns
                                             End If
                                             
                                             'Check if line has a back side
                                             If (linedefs(ld).s2 > -1) Then
                                                  
                                                  'Change sector if this refers to the old sector
                                                  If (sidedefs(linedefs(ld).s2).sector = os) Then sidedefs(linedefs(ld).s2).sector = ns
                                             End If
                                        End If
                                   End If
                              End If
                         Next ld
                    Else
                         
                         'Other side of the line is at the same sector
                         sc1 = sc2
                         
                         'Do not select a specific sector
                         selectsectorindex = -1
                         
                         'Setup the new lines
                         SetupNewLinedefs sc1, sc2, CopyFromSidedef
                    End If
               Else
                    
                    'Other side of the line is at the same sector
                    sc1 = sc2
                    
                    'Do not select a specific sector
                    selectsectorindex = -1
                    
                    'Setup the new lines
                    SetupNewLinedefs sc1, sc2, CopyFromSidedef
               End If
          End If
          
          'Remove line selection properties
          ResetSelections things(0), 0, linedefs(0), numlinedefs, vertexes(0), 0, VarPtr(sectors(0)), 0
          
          'Select a specific sector if given
          If (selectsectorindex > -1) Then
               
               'Select this sector
               SelectSector selectsectorindex
               Set selected = New Dictionary
               selected.Add CStr(selectsectorindex), selectsectorindex
               numselected = selected.Count
               
               'Convert selection to lines
               SelectLinedefsFromSectors
               selectedtype = EM_LINES
          End If
          
          'Setup the new sector
          NewSectorSetup (os = -1)
          
          'Remove looped linedefs
          RemoveLoopedLinedefs
          
          'Fix glitches that other routines may cause
          SolveGlitches
          
          'DEBUG
          'DEBUG_FindUnusedSectors
     End If
     
     'Deselect all
     RemoveSelection False
     
     'Rename the undo description
     If (NewSector) Then
          RenameUndo "sector draw"
     Else
          RenameUndo "line draw"
     End If
     
     'End of drawing operation, set mode back to normal
     submode = ESM_NONE
     
     'Map was changed
     mapnodeschanged = True
     mapchanged = True
     
     'We dont need these anymore
     ReDim changedlines(0)
     numchangedlines = 0
     ReDim DrawingCoords(0)
     NumDrawingCoords = 0
     
     'Redraw the map
     RedrawMap
     
     'Show highlight
     ShowHighlight LastX, LastY
     
     'Update status
     UpdateStatusBar
     
     'Reset mousepointer
     Screen.MousePointer = vbNormal
End Sub


Public Sub EndSelectOperation(ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
     Dim SelectRect As RECT
     Dim c As Long
     
     'Change mousepointer
     Screen.MousePointer = vbHourglass
     
     'Make rectangle from selection
     With SelectRect
          If (GrabX > X) Then .left = X: .right = GrabX Else .left = GrabX: .right = X
          If (-GrabY > -Y) Then .top = -Y: .bottom = -GrabY Else .top = -GrabY: .bottom = -Y
     End With
     
     'Check if we select vertices or things
     If (mode = EM_THINGS) Then
          
          'Select things in rectangle
          c = SelectThingsFromRect(SelectRect, (Shift And vbShiftMask) Or (Shift And vbCtrlMask))
     Else
          
          'Check if we should "downgrade" selection from lines/sectors
          If (mode = EM_LINES) Or (mode = EM_SECTORS) Then SelectVerticesFromLinedefs
          
          'Select vertices in rectangle
          c = SelectVerticesFromRect(SelectRect, (Shift And vbShiftMask) Or (Shift And vbCtrlMask))
          
          'Check if we should "upgrade" selection to lines
          If (mode = EM_LINES) Then
               
               'Upgrade to lines
               SelectLinedefsFromVertices
               
          'Check if we should "upgrade" selection to sectors
          ElseIf (mode = EM_SECTORS) Then
               
               'Upgrade to lines, then to sectors
               SelectLinedefsFromVertices
               SelectSectorsFromLinedefs
          End If
     End If
     
     'Deselect if preferred
     If (c = 0) And (Config("nothingdeselects") = vbChecked) Then RemoveSelection True
     
     'End of select operation, set mode back to normal
     submode = ESM_NONE
     
     'Redraw the map
     RedrawMap
     
     'Show highlight
     ShowHighlight X, Y
     
     'Update status
     UpdateStatusBar
     
     'Reset mousepointer
     Screen.MousePointer = vbNormal
End Sub

Public Sub HideLinedefInfo()
     
     'Hide the info
     fraLinedef.Visible = False
     fraFrontSidedef.Visible = False
     fraBackSidedef.Visible = False
     fraSLinedef.Visible = False
     fraSFrontSidedef.Visible = False
     fraSBackSidedef.Visible = False
     picMap.ToolTipText = ""
     lblBarText.Caption = ""
     Set imgS1Lower.Picture = Nothing
     Set imgS1Middle.Picture = Nothing
     Set imgS1Upper.Picture = Nothing
     Set imgS2Lower.Picture = Nothing
     Set imgS2Middle.Picture = Nothing
     Set imgS2Upper.Picture = Nothing
     Set imgSS1Lower.Picture = Nothing
     Set imgSS1Middle.Picture = Nothing
     Set imgSS1Upper.Picture = Nothing
     Set imgSS2Lower.Picture = Nothing
     Set imgSS2Middle.Picture = Nothing
     Set imgSS2Upper.Picture = Nothing
End Sub

Public Sub HideSectorInfo()
     
     'Hide the info
     fraSector.Visible = False
     fraSectorCeiling.Visible = False
     fraSectorFloor.Visible = False
     fraSSector.Visible = False
     fraSSectorCeiling.Visible = False
     fraSSectorFloor.Visible = False
     picMap.ToolTipText = ""
     lblBarText.Caption = ""
     Set imgCeiling.Picture = Nothing
     Set imgFloor.Picture = Nothing
     Set imgSCeiling.Picture = Nothing
     Set imgSFloor.Picture = Nothing
End Sub


Public Sub HideThingInfo()
     
     'Hide the info
     fraThing.Visible = False
     fraSThing.Visible = False
     fraThingPreview.Visible = False
     fraSThingPreview.Visible = False
     picMap.ToolTipText = ""
     lblBarText.Caption = ""
     Set imgThing.Picture = Nothing
     Set imgSThing.Picture = Nothing
End Sub

Public Sub HideVertexInfo()
     
     'Hide the info
     fraVertex.Visible = False
     fraSVertex.Visible = False
     picMap.ToolTipText = ""
     lblBarText.Caption = ""
End Sub

Private Sub InfoBarClose()
     
     'Close the bar
     Select Case Val(Config("detailsbar"))
          
          Case 1:   'Bottom
               picBar.Align = vbAlignNone
               picBar.height = cmdToggleBar.height
               picBar.Align = vbAlignBottom
               cmdToggleBar.Caption = "5"
               lblBarText.Visible = True
               
          Case 2:   'Top
               picBar.Align = vbAlignNone
               picBar.height = cmdToggleBar.height
               picBar.Align = vbAlignTop
               cmdToggleBar.Caption = "6"
               lblBarText.Visible = True
               
               'Change view so that the map doesnt appear to move
               If (mapfile <> "") Then ChangeView ViewLeft, ViewTop - (106 - cmdToggleBar.height) / ViewZoom, ViewZoom
               
          Case 3:   'Left
               picSBar.Align = vbAlignNone
               picSBar.width = cmdToggleSBar.width
               picSBar.Align = vbAlignLeft
               cmdToggleSBar.Caption = "4"
               
               'Change view so that the map doesnt appear to move
               If (mapfile <> "") Then ChangeView ViewLeft - (135 - cmdToggleSBar.width) / ViewZoom, ViewTop, ViewZoom
               
          Case 4:   'Right
               picSBar.Align = vbAlignNone
               picSBar.width = cmdToggleSBar.width
               picSBar.Align = vbAlignRight
               cmdToggleSBar.Caption = "3"
               
     End Select
     
     'Closed
     cmdToggleBar.tag = "1"
End Sub

Private Sub InfoBarOpen()
     
     'Open the bar
     Select Case Val(Config("detailsbar"))
          
          Case 1:   'Bottom
               picBar.Align = vbAlignNone
               picBar.height = 106
               picBar.Align = vbAlignBottom
               cmdToggleBar.Caption = "6"
               lblBarText.Visible = False
               
          Case 2:   'Top
               picBar.Align = vbAlignNone
               picBar.height = 106
               picBar.Align = vbAlignTop
               cmdToggleBar.Caption = "5"
               lblBarText.Visible = False
               
               'Change view so that the map doesnt appear to move
               If (mapfile <> "") Then ChangeView ViewLeft, ViewTop + (106 - cmdToggleBar.height) / ViewZoom, ViewZoom
               
          Case 3:   'Left
               picSBar.Align = vbAlignNone
               picSBar.width = 135
               picSBar.Align = vbAlignLeft
               cmdToggleSBar.Caption = "3"
               
               'Change view so that the map doesnt appear to move
               If (mapfile <> "") Then ChangeView ViewLeft + (135 - cmdToggleSBar.width) / ViewZoom, ViewTop, ViewZoom
               
          Case 4:   'Right
               picSBar.Align = vbAlignNone
               picSBar.width = 135
               picSBar.Align = vbAlignRight
               cmdToggleSBar.Caption = "4"
               
     End Select
     
     'Opened
     cmdToggleBar.tag = "0"
End Sub

Public Sub InfoBarToggle()
     
     'Tag of cmdToggleBar tells in what state the bar is
     '0 = Open
     '1 = Closed
     
     'Toggle bar
     If (Val(cmdToggleBar.tag) = 0) Then InfoBarClose Else InfoBarOpen
     
     'Save status
     Config("togglebar") = Val(cmdToggleBar.tag)
     
     'Remove highlight
     RemoveHighlight True
     
     'Call form resize to adjust sizes
     Form_Resize
     
     'Get rid of focus
     On Error Resume Next
     picSBar.SetFocus
     picBar.SetFocus
End Sub

Private Sub itmFileExportPicture_Click()
     Dim Result As String
     Dim FilterIndex As Long
     Dim GDIBitmap As clsGDIBitmap
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Ensure the splash dialog is gone
     Unload frmSplash: Set frmSplash = Nothing
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Show save dialog
     Result = SaveFile(Me.hWnd, "Export Picture As", "Windows Bitmap   *.bmp|*.bmp|Portable Network Graphic   *.png|*.png|JPEG Image   *.jpg|*.jpg", mapfilename, cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt, FilterIndex)
     frmMain.Refresh
     
     'Check if not cancelled
     If Result <> "" Then
          
          'Add extension
          If (LCase$(right$(Result, 4)) <> ".bmp") And (FilterIndex = 1) Then Result = Result & ".bmp"
          If (LCase$(right$(Result, 4)) <> ".png") And (FilterIndex = 2) Then Result = Result & ".png"
          If (LCase$(right$(Result, 4)) <> ".jpg") And (FilterIndex = 3) Then Result = Result & ".jpg"
          
          'Give export dialog
          Load frmExportPicture
          
          'Show dialog
          frmExportPicture.Show 1, Me
          
          'Check result
          If (frmExportPicture.tag = "OK") Then
               
               'Hourglass
               Screen.MousePointer = vbHourglass
               
               'Make the picture
               If RenderExportPicture() Then
                    
                    'Make GDI+ bitmap from picture
                    Set GDIBitmap = New clsGDIBitmap
                    GDIBitmap.CreateFromHBitmap picTexture.Picture.Handle
                    
                    'Kill file if already exists
                    If (Dir(Result) <> "") Then Kill Result
                    
                    'Convert and save the bitmap
                    Select Case FilterIndex
                         Case 1: GDIBitmap.SaveToFile Result, GDIBitmap.EncoderGuid(GDIBitmap.ExtensionExists("*.bmp")), 0
                         Case 2: GDIBitmap.SaveToFile Result, GDIBitmap.EncoderGuid(GDIBitmap.ExtensionExists("*.png")), 0
                         Case 3: GDIBitmap.SaveToFile Result, GDIBitmap.EncoderGuid(GDIBitmap.ExtensionExists("*.jpg")), 0
                    End Select
                    
                    'Clean up
                    Set GDIBitmap = Nothing
               End If
               
               'Normal mouse
               Screen.MousePointer = vbNormal
          End If
          
          'Unload dialog
          Unload frmExportPicture
          Set frmExportPicture = Nothing
     End If
End Sub

Private Sub RevertDrawingOperation()
     Dim RecordedCoords() As POINT
     Dim NumRecordedCoords As Long
     Dim i As Long
     
     'Copy the recorded coordinates
     RecordedCoords = DrawingCoords
     NumRecordedCoords = NumDrawingCoords
     
     'Undo complete drawing operation
     CancelDrawOperation
     
     'Restart drawing
     StartDrawOperation False
     
     'Anything to draw?
     If (NumRecordedCoords > 0) Then
          
          'Insert vertices at all recorded coords, except the last
          For i = 0 To NumRecordedCoords - 2
               
               'Repeat this vertex draw
               DrawVertexHere RecordedCoords(i).X, RecordedCoords(i).Y, False
          Next i
     End If
     
     'Redraw map
     RedrawMap
     
     'Render the line being drawn
     RenderDrawingLine PAL_MULTISELECTION, PAL_BACKGROUND, LastX, LastY
     
     'Show changes
     picMap.Refresh
End Sub

Public Sub ShowLinedefInfo(ByVal ld As Long)
     Dim distance As Long
     Dim nearest As Long
     Dim OldSelected As Long
     Dim xl As Long, yl As Long
     
     Dim Action As String
     Dim Length As Long
     Dim sTag As String
     Dim FHeight As String
     Dim bheight As String
     Dim FSector As String
     Dim BSector As String
     Dim fx As String
     Dim fy As String
     Dim bx As String
     Dim by As String
     Dim S1U As String
     Dim S1M As String
     Dim S1L As String
     Dim S2U As String
     Dim S2M As String
     Dim S2L As String
     
     'Get information
     If (mapconfig("mapformat") = 1) Then
          sTag = linedefs(ld).tag
     ElseIf (mapconfig("mapformat") = 2) Then
          If (linedefs(ld).argref0 > 0) Then If (sTag = "") Then sTag = sTag & linedefs(ld).arg0 Else sTag = sTag & ", " & linedefs(ld).arg0
          If (linedefs(ld).argref1 > 0) Then If (sTag = "") Then sTag = sTag & linedefs(ld).arg1 Else sTag = sTag & ", " & linedefs(ld).arg1
          If (linedefs(ld).argref2 > 0) Then If (sTag = "") Then sTag = sTag & linedefs(ld).arg2 Else sTag = sTag & ", " & linedefs(ld).arg2
          If (linedefs(ld).argref3 > 0) Then If (sTag = "") Then sTag = sTag & linedefs(ld).arg3 Else sTag = sTag & ", " & linedefs(ld).arg3
          If (linedefs(ld).argref4 > 0) Then If (sTag = "") Then sTag = sTag & linedefs(ld).arg4 Else sTag = sTag & ", " & linedefs(ld).arg4
          If (sTag = "") Then sTag = "0"
     End If
     
     'Calculate linedef length
     xl = vertexes(linedefs(ld).v2).X - vertexes(linedefs(ld).v1).X
     yl = vertexes(linedefs(ld).v2).Y - vertexes(linedefs(ld).v1).Y
     Length = CLng(Sqr(xl * xl + yl * yl))
     
     'Check if the linedef type can be found
     If (mapconfig("linedeftypes").Exists(CStr(linedefs(ld).effect))) Then
          Action = linedefs(ld).effect & " - " & Trim$(mapconfig("linedeftypes")(CStr(linedefs(ld).effect))("title"))
     Else
          'Check if generalized type
          If IsGenLinedefEffect(linedefs(ld).effect) Then
               Action = linedefs(ld).effect & " - " & GetGenLinedefCategory(linedefs(ld).effect)("title") & "*"
          Else
               Action = linedefs(ld).effect & " - Unknown"
          End If
     End If
     
     'Action on Tooltip
     picMap.ToolTipText = ""
     If (linedefs(ld).effect > 0) And (Val(Config("showtooltips")) <> 0) Then
          
          'Only when not in 3D Mode
          If (mode <> EM_3D) Then picMap.ToolTipText = Action
     End If
     
     'Check if there is a front sidedef
     If (linedefs(ld).s1 > -1) Then
          
          'Check for sector
          If (sidedefs(linedefs(ld).s1).sector > -1) Then
               
               'Get the front height
               FHeight = sectors(sidedefs(linedefs(ld).s1).sector).hceiling - sectors(sidedefs(linedefs(ld).s1).sector).hfloor
               FSector = sidedefs(linedefs(ld).s1).sector
          End If
          
          'Get offsets
          fx = sidedefs(linedefs(ld).s1).tx
          fy = sidedefs(linedefs(ld).s1).ty
          
          'Get textures
          S1U = sidedefs(linedefs(ld).s1).Upper
          S1M = sidedefs(linedefs(ld).s1).Middle
          S1L = sidedefs(linedefs(ld).s1).Lower
     Else
          FHeight = "-"
          FSector = "-"
          fx = "-"
          fy = "-"
     End If
     
     'Check if there is a back sidedef
     If (linedefs(ld).s2 > -1) Then
          
          'Check for sector
          If (sidedefs(linedefs(ld).s2).sector > -1) Then
               
               'Get the front height
               bheight = sectors(sidedefs(linedefs(ld).s2).sector).hceiling - sectors(sidedefs(linedefs(ld).s2).sector).hfloor
               BSector = sidedefs(linedefs(ld).s2).sector
          End If
          
          'Get offsets
          bx = sidedefs(linedefs(ld).s2).tx
          by = sidedefs(linedefs(ld).s2).ty
          
          'Get textures
          S2U = sidedefs(linedefs(ld).s2).Upper
          S2M = sidedefs(linedefs(ld).s2).Middle
          S2L = sidedefs(linedefs(ld).s2).Lower
     Else
          bheight = "-"
          BSector = "-"
          bx = "-"
          by = "-"
     End If
     
     
     'Check on what panel to show the info
     If picBar.Visible Then
          
          'Check if the bar is fully shown
          If (cmdToggleBar.tag = "0") Then
               
               'Set the info
               fraLinedef.Caption = " Linedef " & ld & " "
               lblLinedefType.Caption = ShortedText(Action, lblLinedefType.width \ Screen.TwipsPerPixelX)
               lblLinedefLength.Caption = Length
               lblLinedefTag.Caption = sTag
               lblS1Height.Caption = FHeight
               lblS2height.Caption = bheight
               lblS1Sector.Caption = FSector
               lblS2Sector.Caption = BSector
               lblS1X.Caption = fx
               lblS1Y.Caption = fy
               lblS2X.Caption = bx
               lblS2Y.Caption = by
               lblS1Upper.Caption = S1U
               lblS1Middle.Caption = S1M
               lblS1Lower.Caption = S1L
               lblS2Upper.Caption = S2U
               lblS2Middle.Caption = S2M
               lblS2Lower.Caption = S2L
               GetScaledTexturePicture S1U, imgS1Upper, , RequiresS1Upper(ld)
               GetScaledTexturePicture S1M, imgS1Middle, , RequiresS1Middle(ld)
               GetScaledTexturePicture S1L, imgS1Lower, , RequiresS1Lower(ld)
               GetScaledTexturePicture S2U, imgS2Upper, , RequiresS2Upper(ld)
               GetScaledTexturePicture S2M, imgS2Middle, , RequiresS2Middle(ld)
               GetScaledTexturePicture S2L, imgS2Lower, , RequiresS2Lower(ld)
               
               'Show panels
               fraLinedef.Visible = True
               fraFrontSidedef.Visible = (linedefs(ld).s1 > -1)
               fraBackSidedef.Visible = (linedefs(ld).s2 > -1)
          Else
               
               'Only show a sinle line
               lblBarText.Caption = " Linedef " & ld & ":  " & Action
          End If
          
     ElseIf picSBar.Visible Then
          
          'Check if the bar is fully shown
          If (cmdToggleBar.tag = "0") Then
               
               'Set the info
               fraSLinedef.Caption = " Linedef " & ld & " "
               lblSLinedefType.Caption = ShortedText(Action, lblSLinedefType.width \ Screen.TwipsPerPixelX)
               lblSLinedefLength.Caption = Length
               lblSLinedefTag.Caption = sTag
               lblSS1Height.Caption = FHeight
               lblSS2height.Caption = bheight
               lblSS1Sector.Caption = FSector
               lblSS2Sector.Caption = BSector
               lblSS1Upper.Caption = S1U
               lblSS1Middle.Caption = S1M
               lblSS1Lower.Caption = S1L
               lblSS2Upper.Caption = S2U
               lblSS2Middle.Caption = S2M
               lblSS2Lower.Caption = S2L
               GetScaledTexturePicture S1U, imgSS1Upper, , RequiresS1Upper(ld)
               GetScaledTexturePicture S1M, imgSS1Middle, , RequiresS1Middle(ld)
               GetScaledTexturePicture S1L, imgSS1Lower, , RequiresS1Lower(ld)
               GetScaledTexturePicture S2U, imgSS2Upper, , RequiresS2Upper(ld)
               GetScaledTexturePicture S2M, imgSS2Middle, , RequiresS2Middle(ld)
               GetScaledTexturePicture S2L, imgSS2Lower, , RequiresS2Lower(ld)
               
               'Show panels
               fraSLinedef.Visible = True
               fraSFrontSidedef.Visible = (linedefs(ld).s1 > -1)
               fraSBackSidedef.Visible = (linedefs(ld).s2 > -1)
          Else
               
               'Only show a sinle line
               lblBarText.Caption = " Linedef " & ld & ":  " & Action
          End If
     End If
End Sub


Public Sub ShowSectorInfo(ByVal sc As Long)
     Dim nearest As Long
     Dim OldSelected As Long
     Dim ld As Long, ldfound As Long
     
     Dim effect As String
     Dim Ceiling As Long
     Dim Floor As Long
     Dim sTag As Long
     Dim sHeight As Long
     Dim Brightness As Long
     Dim tceiling As String
     Dim tfloor As String
     
     'Get the information
     Ceiling = sectors(sc).hceiling
     Floor = sectors(sc).hfloor
     sTag = sectors(sc).tag
     sHeight = sectors(sc).hceiling - sectors(sc).hfloor
     Brightness = sectors(sc).Brightness
     tceiling = sectors(sc).tceiling
     tfloor = sectors(sc).tfloor
     
     'Check if the sector effect can be found
     If (Trim$(mapconfig("sectortypes").Exists(CStr(sectors(sc).special)))) Then
          effect = sectors(sc).special & " - " & Trim$(mapconfig("sectortypes")(CStr(sectors(sc).special)))
     Else
          effect = sectors(sc).special & " - Unknown"
     End If
     
     'Effect on Tooltip
     picMap.ToolTipText = ""
     If (sectors(sc).special > 0) And (Val(Config("showtooltips")) <> 0) Then
          
          'Only when not in 3D Mode
          If (mode <> EM_3D) Then picMap.ToolTipText = effect
     End If
     
     'Check what panel to show the info on
     If picBar.Visible Then
          
          'Check if the bar is fully shown
          If (cmdToggleBar.tag = "0") Then
               
               'Set the info
               fraSector.Caption = " Sector " & sc & " "
               lblSectorType = ShortedText(effect, lblSectorType.width \ Screen.TwipsPerPixelX)
               lblSectorTag.Caption = sTag
               lblSectorCeiling.Caption = Ceiling
               lblSectorFloor.Caption = Floor
               lblSectorHeight.Caption = sHeight
               lblSectorLight.Caption = Brightness
               lblCeiling.Caption = tceiling
               lblFloor.Caption = tfloor
               GetScaledFlatPicture tceiling, imgCeiling
               GetScaledFlatPicture tfloor, imgFloor
               
               'Show the info
               fraSector.Visible = True
               fraSectorCeiling.Visible = True
               fraSectorFloor.Visible = True
          Else
               
               'Show single line only
               lblBarText.Caption = " Sector " & sc & ":  " & effect
          End If
          
     ElseIf picSBar.Visible Then
          
          'Check if the bar is fully shown
          If (cmdToggleBar.tag = "0") Then
               
               'Set the info
               fraSSector.Caption = " Sector " & sc & " "
               lblSSectorType = ShortedText(effect, lblSSectorType.width \ Screen.TwipsPerPixelX)
               lblSSectorTag.Caption = sTag
               lblSSectorCeiling.Caption = Ceiling
               lblSSectorFloor.Caption = Floor
               lblSSectorHeight.Caption = sHeight
               lblSSectorLight.Caption = Brightness
               lblSCeiling.Caption = tceiling
               lblSFloor.Caption = tfloor
               GetScaledFlatPicture tceiling, imgSCeiling
               GetScaledFlatPicture tfloor, imgSFloor
               
               'Show the info
               fraSSector.Visible = True
               fraSSectorCeiling.Visible = True
               fraSSectorFloor.Visible = True
          Else
               
               'Show single line only
               lblBarText.Caption = " Sector " & sc & ":  " & effect
          End If
     End If
End Sub

Private Sub cmdToggleBar_Click()
     InfoBarToggle
End Sub

Private Sub cmdToggleSBar_Click()
     InfoBarToggle
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     Dim ShortcutCode As Long
     Dim DoUpdateStatusBar As Boolean
     Dim DoRedrawMap As Boolean
     
     'Adjust shift mask
     CurrentShiftMask = Shift
     
     'Ignore shift keys alone
     If (KeyCode = 16) Or (KeyCode = 17) Or (KeyCode = 18) Then Exit Sub
     
     'Make the shortcut code from keycode and shift
     ShortcutCode = KeyCode Or (Shift * (2 ^ 16))
     
     
     'Check if in 3D Mode
     If (mode = EM_3D) Then
          
          'Leave immediately if no map is loaded
          If (mapfile = "") Then Exit Sub
          
          'Leave when map edit is disabled
          If (picMap.Enabled = False) Then Exit Sub
          
          'Check how we should process data
          If TextureSelecting Then
               
               'Perform the action associated with the key
               KeydownTextureSelect ShortcutCode
          Else
               
               'Perform the action associated with the key
               Keydown3D ShortcutCode
          End If
          
          'Leave when form is unloaded
          If (IsLoaded(frmMain) = False) Then Exit Sub
     Else
          
          'Do menu keys
          KeypressMenus ShortcutCode
          
          'Leave immediately if no map is loaded
          If (mapfile = "") Then Exit Sub
          
          'Leave when map edit is disabled
          If (picMap.Enabled = False) Then Exit Sub
          
          'Do general keys
          KeypressGeneral ShortcutCode, DoUpdateStatusBar, DoRedrawMap
          
          'Leave when form is unloaded
          If (IsLoaded(frmMain) = False) Then Exit Sub

          'Do mode-specific keys
          Select Case mode
               Case EM_VERTICES: KeypressVertexes ShortcutCode
               Case EM_LINES: KeypressLines ShortcutCode
               Case EM_SECTORS: KeypressSectors ShortcutCode
               Case EM_THINGS: KeypressThings ShortcutCode
          End Select
          
          'Leave when form is unloaded
          If (IsLoaded(frmMain) = False) Then Exit Sub
     End If
     
     
     'Update status
     If DoUpdateStatusBar Then UpdateStatusBar
     
     'Check if we should redraw
     If DoRedrawMap Then
          
          'Redraw map
          RedrawMap
          
          'Show highlight
          If (submode = ESM_NONE) Then ShowHighlight LastX, LastY
     End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
     
     'Check if in 3D Mode
     If (mode = EM_3D) Then
          
          'Check how we should process data
          If TextureSelecting Then
               
               'Perform the action associated with the key
               KeypressTextureSelect KeyAscii
               
               'Leave when form is unloaded
               If (IsLoaded(frmMain) = False) Then Exit Sub
          End If
     End If
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
     Dim ShortcutCode As Long
     Dim DoUpdateStatusBar As Boolean
     Dim DoRedrawMap As Boolean
     
     'Adjust shift mask
     CurrentShiftMask = Shift
     
     'Ignore shift keys alone
     If (KeyCode = 16) Or (KeyCode = 17) Or (KeyCode = 18) Then Exit Sub
     
     'Make the shortcut code from keycode and shift
     ShortcutCode = KeyCode Or (Shift * (2 ^ 16))
     
     
     'Leave immediately if no map is loaded
     If (mapfile = "") Then Exit Sub
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     
     'Check if in 3D Mode
     If (mode = EM_3D) Then
          
          'Check how we should process data
          If Not TextureSelecting Then
               
               'Perform the action associated with the key
               Keyrelease3D ShortcutCode
          End If
     Else
          
          'Do general keys
          KeyreleaseGeneral ShortcutCode, DoUpdateStatusBar, DoRedrawMap
     End If
     
     'Leave when form is unloaded
     If (IsLoaded(frmMain) = False) Then Exit Sub
     
     'Update status
     If DoUpdateStatusBar Then UpdateStatusBar
     
     'Check if we should redraw
     If DoRedrawMap Then
          
          'Redraw map
          RedrawMap
          
          'Show highlight
          If (submode = ESM_NONE) Then ShowHighlight LastX, LastY
     End If
     
     'Reset the F7 count
     F7Count = 0
End Sub


Private Sub Form_Load()
     
     'Make subclassing
     CreateSubclassing
     
     'Apply configuration on interface
     ApplyInterfaceConfiguration
     
     'Disable controls for map editing (no map loaded yet)
     DisableMapEditing
     
     'Make menu item names with shortcuts
     UpdateMenuShortcuts
     
     'Update status bar
     UpdateStatusBar
     
     'Show recent files
     UpdateRecentFilesMenu
     
     'Apply window sizes
     With frmMain
          .left = Config("mainwindow")("left")
          .top = Config("mainwindow")("top")
          If (Config("mainwindow")("width") > 1500) Then .width = Config("mainwindow")("width")
          If (Config("mainwindow")("height") > 1500) Then .height = Config("mainwindow")("height")
          .WindowState = Config("mainwindow")("windowstate")
     End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     
     'Ensure the splash dialog is gone
     Unload frmSplash: Set frmSplash = Nothing
     
     'Unload map and terminate when unloaded
     If (MapUnload) Then
          
          'Remove subclassing
          DestroySubclassing
          
          'Do a resize to keep window sizes
          Form_Resize
          
          'Terminate program
          Terminate
     Else
          
          'Cancel the unload
          Cancel = True
     End If
End Sub

Public Sub Form_Resize()
     On Local Error Resume Next
     Dim ToolbarHeight As Long
     Dim BottomHeight As Long
     Dim TopHeight As Long
     Dim LeftHeight As Long
     Dim RightHeight As Long
     
     'Lock the viewport
     If (Running3D = False) Then LockWindowUpdate Me.hWnd 'picMap.hWnd
     
     'Move the bar toggle button
     cmdToggleBar.left = ScaleWidth - cmdToggleBar.width - 4
     If (Val(Config("detailsbar")) = 2) Then cmdToggleBar.top = picBar.ScaleHeight - cmdToggleBar.height Else cmdToggleBar.top = 0
     cmdToggleSBar.top = ScaleHeight - stbStatus.height - tlbToolbar.height - cmdToggleSBar.height - 4
     If (Val(Config("detailsbar")) = 3) Then cmdToggleSBar.left = picSBar.ScaleWidth - cmdToggleSBar.width Else cmdToggleSBar.left = 0
     
     'Make sure the renderer is terminated
     TerminateMapRenderer
     
     'Determine space to reserve
     If tlbToolbar.Visible Then ToolbarHeight = tlbToolbar.height
     If picBar.Visible And (picBar.Align = vbAlignBottom) Then BottomHeight = picBar.height
     If picBar.Visible And (picBar.Align = vbAlignTop) Then TopHeight = picBar.height
     If picSBar.Visible And (picSBar.Align = vbAlignLeft) Then LeftHeight = picSBar.width
     If picSBar.Visible And (picSBar.Align = vbAlignRight) Then RightHeight = picSBar.width
     
     'Map screen
     With picMap
          .top = ToolbarHeight + TopHeight + 2
          .left = LeftHeight + 2
          .width = ScaleWidth - RightHeight - picMap.left - 2
          .height = ScaleHeight - BottomHeight - picMap.top - stbStatus.height - 2
     End With
     
     'Check if in 3D Mode
     If (Running3D) Then
          
          'Free the mouse
          FreeMouse
          
          'Determine rendering area
          DetermineRenderScreenSize frmMain.picMap
          
          'Reclaim mouse
          CaptureMouse
          
          'Render now to update
          RunSingleFrame False, True
     Else
          
          'Only initialize renderer when a map is loaded
          If (mapfile <> "") Then
               
               'Initialize the map screen
               InitializeMapRenderer frmMain.picMap
               
               'Set the viewport
               ChangeView ViewLeft, ViewTop, ViewZoom
               
               'Redraw entire map
               RedrawMap
          End If
     End If
     
     'Unlock the viewport
     If (Running3D = False) Then LockWindowUpdate 0
     
     'Check if visible
     If (frmMain.Visible = True) Then
          
          'Save windowstate
          Config("mainwindow")("windowstate") = frmMain.WindowState
          
          'Check if it has a valid size now
          If (frmMain.WindowState = vbNormal) Then
               
               'Save window size
               Config("mainwindow")("left") = frmMain.left
               Config("mainwindow")("top") = frmMain.top
               Config("mainwindow")("width") = frmMain.width
               Config("mainwindow")("height") = frmMain.height
          End If
     End If
End Sub

Private Function InsertThingHere(ByVal X As Long, ByVal Y As Long) As Long
     Dim t As Long
     
     'Remove higlight
     RemoveHighlight True
     
     'No more selection
     ResetSelections things(0), numthings, linedefs(0), numlinedefs, vertexes(0), numvertexes, VarPtr(sectors(0)), numsectors
     Set selected = New Dictionary
     numselected = 0
     selectedtype = EM_THINGS
     
     'Make undo
     CreateUndo "thing insert"
     
     'Snap X and Y if snap mode is on
     If snapmode Then
          X = SnappedToGridX(X)
          Y = SnappedToGridY(Y)
     End If
     
     'Create a thing now
     t = CreateThing
     InsertThingHere = t
     
     'Set thing defaults
     things(t) = LastThing
     With things(t)
          .selected = 0
          .X = X
          .Y = -Y
     End With
     
     'Check if we should erase tag and actions
     If (Config("copytagdraw") = vbUnchecked) Then
          With things(t)
               .arg0 = 0
               .arg1 = 0
               .arg2 = 0
               .arg3 = 0
               .arg4 = 0
               .effect = 0
               .tag = 0
          End With
     End If
     
     'Update thing image and color
     UpdateThingImageColor t
     UpdateThingSize t
     UpdateThingCategory t
     
     'Check if we should edit the thing
     If (Config("newthingdialog") = vbChecked) Then
          
          'Select thing
          things(t).selected = 1
          selected.Add CStr(t), t
          numselected = 1
          
          'Show this
          RedrawMap
          
          'Load dialog
          Load frmThing
          
          'Dont make undo for this edit
          frmThing.lblMakeUndo.Caption = "No"
          
          'Show dialog
          picMap.Enabled = False
          frmThing.Show 1, Me
          picMap.Enabled = True
          
          'No more selection
          ResetSelections things(0), numthings, linedefs(0), numlinedefs, vertexes(0), numvertexes, VarPtr(sectors(0)), numsectors
          Set selected = New Dictionary
          numselected = 0
     End If
     
     'Map has changed
     mapchanged = True
End Function

Private Sub InsertVertexHere(ByVal X As Long, ByVal Y As Long)
     Dim distance As Long
     Dim nl As Long
     Dim nv As Long
     Dim d As Long

     'Snap X and Y if snap mode is on
     If snapmode Then
          X = SnappedToGridX(X)
          Y = SnappedToGridY(Y)
     End If
     
     'Get the nearest vertex
     nv = NearestVertex(X, Y, vertexes(0), numvertexes, d)
     
     'Check if no vertex already exists at these coordinates
     If (nv = -1) Or (d > 0) Then
          
          'Make undo
          CreateUndo "vertex insert"
          
          'Insert a vertex now
          nv = InsertVertex(X, -Y)
          
          'Draw the vertex
          Render_AllVertices vertexes(0), nv, nv, vertexsize
          
          'Get the nearest linedef
          nl = NearestLinedef(X, Y, vertexes(0), linedefs(0), numlinedefs, distance)
          
          'Check if distance is close enough for linedef split
          If (distance <= Config("linesplitdistance")) Then
               
               'Split the linedef with this vertex
               SplitLinedef nl, nv
               
               'Redraw the map
               RedrawMap False
                
               'DEBUG
               'DEBUG_FindUnusedSectors
              
               'Highlight whatever is under teh mouz0r
               If MouseInside Then picMap_MouseMove 0, 0, LastX, LastY
          End If
          
          'Map has changed
          mapchanged = True
          mapnodeschanged = True
     End If
End Sub

Private Sub itmEditCenterView_Click()
     Dim MapRect As RECT
     
     'Calculating map rect
     MapRect = CalculateMapRect
     
     'Center map in view
     CenterViewAt MapRect, True
     
     'Redraw map
     RedrawMap False
     
     'Show highlight
     ShowHighlight LastX, LastY
End Sub

Private Sub itmEditCopy_Click()
     Dim DeselectAfterCopy As Boolean
     
     'Leave if no selection is made
     If ((numselected = 0) And (currentselected = -1)) Then Exit Sub
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Check if not in normal mode
     If (submode <> ESM_NONE) Then
          
          'Show message, cant right now
          MsgBox "Cannot perform Copy while in an operation." & vbLf & "Please finish your operation first!", vbExclamation
          
          'Leave here
          Exit Sub
     End If
     
     'Change mousepointer
     Screen.MousePointer = vbHourglass
     
     'Check if no selection is made
     If (numselected = 0) Then
          
          'Temporarely select highlighted object for the mode we are in
          Select Case mode
               Case EM_VERTICES: SelectCurrentVertex
               Case EM_LINES: SelectCurrentLine
               Case EM_SECTORS: SelectCurrentSector
               Case EM_THINGS: SelectCurrentThing
          End Select
          
          'After copy, deselect this
          DeselectAfterCopy = True
     End If
     
     'Clear clipboard
     Clipboard.Clear
     
     'Clear file
     ClipboardCleanup
     
     'Save the selection to file
     SavePrefabSelection ClipboardFile
     
     'Set descriptor on clipboard
     ClipboardSetDescriptor
     
     'Check if we should deselect temporary selection
     If DeselectAfterCopy Then RemoveSelection True
     
     'Reset mousepointer
     Screen.MousePointer = vbNormal
End Sub

Private Sub itmEditCut_Click()
     
     'Leave if no selection is made
     If (numselected = 0) Then Exit Sub
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Check if not in normal mode
     If (submode <> ESM_NONE) Then
          
          'Show message, cant right now
          MsgBox "Cannot perform Cut while in an operation." & vbLf & "Please finish your operation first!", vbExclamation
          
          'Leave here
          Exit Sub
     End If
     
     
     'First copy
     itmEditCopy_Click
     
     'Delete selection/highlight
     DeleteSelection "cut"
     
     'Redraw map
     RedrawMap
End Sub

Private Sub itmEditDelete_Click()
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Check if not in normal mode
     If (submode <> ESM_NONE) And (submode <> ESM_PASTING) Then
          
          'Show message, cant right now
          MsgBox "Cannot perform Delete while in an operation." & vbLf & "Please finish your operation first!", vbExclamation
          
          'Leave here
          Exit Sub
     End If
     
     'Check if pasting
     If (submode = ESM_PASTING) Then
          
          'Just cancel
          CancelCurrentOperation
     Else
          
          'Delete selection/highlight
          DeleteSelection "delete"
          
          'Update status bar
          UpdateStatusBar
          
          'Redraw map
          RedrawMap
     End If
End Sub

Private Sub itmEditFind_Click()
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Load dialog
     Load frmFind
     
     'Set the height and hide replace controls
     frmFind.height = 1860
     frmFind.txtReplace.Visible = False
     frmFind.cmdBrowseReplace.Visible = False
     frmFind.chkReplaceOnly.Visible = False
     
     'Show dialog
     frmFind.Show 1, Me
End Sub

Private Sub itmEditFlipH_Click()
     Dim Deselect As Boolean
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     If (submode <> ESM_PASTING) Then CancelCurrentOperation
     
     'Check if no selection is made
     If (numselected = 0) And (submode <> ESM_PASTING) Then
          
          'If no highlight exists, leave
          If (currentselected = -1) Then Exit Sub
          
          'Make selection from highlight
          Select Case mode
               Case EM_VERTICES: SelectCurrentVertex
               Case EM_LINES: SelectCurrentLine
               Case EM_SECTORS: SelectCurrentSector
               Case EM_THINGS: SelectCurrentThing
          End Select
          
          'Remove selection after edit
          Deselect = True
     End If
     
     'Make undo
     If (submode <> ESM_PASTING) Then CreateUndo "flip horizontally"
     
     'Check if flipping things
     If (mode = EM_THINGS) Then
          
          'Perform the flip
          FlipThingsHorizontal
          
          'Map changed
          mapchanged = True
     Else
          
          'Check if the selection should be modified
          If ((mode = EM_LINES) Or (mode = EM_SECTORS)) And (submode <> ESM_PASTING) Then SelectVerticesFromLinedefs
          
          'Perform the flip
          FlipVerticesHorizontal
          
          'Check if the selection should be reversed
          If (submode <> ESM_PASTING) Then
               If (mode = EM_LINES) Or (mode = EM_SECTORS) Then SelectLinedefsFromVertices
               If (mode = EM_SECTORS) Then SelectSectorsFromLinedefs
          End If
          
          'Map changed
          mapchanged = True
          mapnodeschanged = True
     End If
     
     'Check if we should deselect
     If Deselect Then
          
          'Remove selection
          RemoveSelection False
     End If
     
     'Redraw map
     RedrawMap
End Sub

Private Sub itmEditFlipV_Click()
     Dim Deselect As Boolean
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     If (submode <> ESM_PASTING) Then CancelCurrentOperation
     
     'Check if no selection is made
     If (numselected = 0) And (submode <> ESM_PASTING) Then
          
          'If no highlight exists, leave
          If (currentselected = -1) Then Exit Sub
          
          'Make selection from highlight
          Select Case mode
               Case EM_VERTICES: SelectCurrentVertex
               Case EM_LINES: SelectCurrentLine
               Case EM_SECTORS: SelectCurrentSector
               Case EM_THINGS: SelectCurrentThing
          End Select
          
          'Remove selection after edit
          Deselect = True
     End If
     
     'Make undo
     If (submode <> ESM_PASTING) Then CreateUndo "flip horizontally"
     
     'Check if flipping things
     If (mode = EM_THINGS) Then
          
          'Perform the flip
          FlipThingsVertical
          
          'Map changed
          mapchanged = True
     Else
          
          'Check if the selection should be modified
          If ((mode = EM_LINES) Or (mode = EM_SECTORS)) And (submode <> ESM_PASTING) Then SelectVerticesFromLinedefs
          
          'Perform the flip
          FlipVerticesVertical
          
          'Check if the selection should be reversed
          If (submode <> ESM_PASTING) Then
               If (mode = EM_LINES) Or (mode = EM_SECTORS) Then SelectLinedefsFromVertices
               If (mode = EM_SECTORS) Then SelectSectorsFromLinedefs
          End If
          
          'Map changed
          mapchanged = True
          mapnodeschanged = True
     End If
     
     'Check if we should deselect
     If Deselect Then
          
          'Remove selection
          RemoveSelection False
     End If
     
     'Redraw map
     RedrawMap
End Sub

Public Sub itmEditMapOptions_Click()
     On Error GoTo MapOptionsError
     Dim OldGame As String
     Dim OldLumpName As String
     Dim MapLumpIndex As Long
     Dim OldMapFormat As Long
     Dim CurIWADFile As String
     Dim OldAddWad As String
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Remove highlight
     If (mapfile <> "") Then RemoveHighlight True
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Keep old settings
     OldGame = mapgame
     OldLumpName = maplumpname
     OldMapFormat = mapconfig("mapformat")
     OldAddWad = addwadfile
     
     'Change the map options
     If (ChangeMapOptions) And (mapfile <> "") Then
          
          'Change mousepointer
          Screen.MousePointer = vbHourglass
          
          'Show status dialog
          frmStatus.Show 0, frmMain
          frmMain.SetFocus
          frmMain.Refresh
          
          'Load the error log
          ErrorLog_Load
          
          'Load new map configuration
          DisplayStatus "Loading configuration..."
          LoadMapConfiguration mapgame
          
          'Check if the header lump was renamed
          If (OldLumpName <> maplumpname) Then
               
               'Remember the lump name as it was saved so we can remove that later
               If (mapoldlumpname = "") Then mapoldlumpname = OldLumpName
          End If
          
          'Check if map format changed
          If (OldMapFormat <> mapconfig("mapformat")) Then
               
               'Show warning
               MsgBox "WARNING: You have changed the map file format!" & vbLf & "Because your map is not designed for this format, it may not work correctly in the game!", vbCritical
          End If
          
          'Only reload textures and flat when game (IWAD) changed or additional file changed
          If (OldGame <> mapgame) Or (OldAddWad <> addwadfile) Then
               
               'Close additional wads
               IWAD.CloseFile
               AddWAD.CloseFile
               
               'Open additional wads
               OpenIWADFile
               OpenADDWADFile
               
               'Precache resources
               MapLoadResources
          End If
          
          'Create data structure optimizations
          DisplayStatus "Optimizing data structures..."
          CreateOptimizations
          
          'Unload status dialog
          Unload frmStatus: Set frmSplash = Nothing
          
          'Reset mousepointer
          Screen.MousePointer = vbDefault
          
          'Show the errors and warnings dialog
          ErrorLog_DisplayAndFlush
          
          'Update scripts menu
          UpdateScriptLumpsMenu
          
          'Re-select current editing mode
          'This will clear selection and redraw map
          itmEditMode_Click CInt(mode)
     End If
     
     'Map changed
     mapchanged = True
     
     'We're done here
     Exit Sub
     
     
MapOptionsError:
     
     'Show error message
     MsgBox "Error " & Err.number & " while change map options: " & Err.Description, vbCritical
     
     'Unload dialog
     Unload frmStatus: Set frmSplash = Nothing
     
     'Reset mousepointer
     Screen.MousePointer = vbDefault
End Sub

Public Sub itmEditMode_Click(Index As Integer)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Toggle to previous mode?
     If (mode = EM_MOVE) And (Index = EM_MOVE) Then
          
          'Go to previous mode!
          Index = PreviousMode
          
     'From move mode?
     ElseIf (mode = EM_MOVE) And (PreviousMode <> EM_3D) Then
          
          'Go to previous mode first!
          itmEditMode_Click EM_MOVE
     End If
     
     'Keep previous mode
     PreviousMode = mode
     
     'Check what to convert From
     Select Case mode
          
          Case EM_VERTICES ', EM_MOVE
               
               'Check what to convert To
               Select Case Index
                    Case EM_LINES: SelectLinedefsFromVertices
                    Case EM_SECTORS: SelectLinedefsFromVertices: SelectSectorsFromLinedefs
                    Case EM_THINGS: RemoveSelection False   'Dont convert, just deselect
               End Select
               
          Case EM_LINES
               
               'Check what to convert To
               Select Case Index
                    Case EM_VERTICES: SelectVerticesFromLinedefs      ', EM_MOVE
                    Case EM_SECTORS: SelectSectorsFromLinedefs
                    Case EM_THINGS: RemoveSelection False   'Dont convert, just deselect
               End Select
               
               'Hide toolbar buttons
               tlbToolbar.Buttons("LinesFlip").Visible = False
               tlbToolbar.Buttons("LinesCurve").Visible = False
               
          Case EM_SECTORS
               
               'Check what to convert To
               Select Case Index
                    Case EM_VERTICES: SelectVerticesFromLinedefs      ', EM_MOVE
                    Case EM_LINES: SelectLinedefsFromSectors
                    Case EM_THINGS: RemoveSelection False   'SelectThingsFromSectors is too slow, must find a faster solution
               End Select
               
               'Hide toolbar buttons
               tlbToolbar.Buttons("SectorsJoin").Visible = False
               tlbToolbar.Buttons("SectorsMerge").Visible = False
               tlbToolbar.Buttons("SectorsGradientBrightness").Visible = False
               tlbToolbar.Buttons("SectorsGradientFloors").Visible = False
               tlbToolbar.Buttons("SectorsGradientCeilings").Visible = False
               
          Case EM_THINGS
               
               'Dont convert, just deselect
               RemoveSelection False
               
               'Hide toolbar buttons
               tlbToolbar.Buttons("ThingsFilter").Visible = False
               
          Case EM_3D
               
               'Stop 3D Mode
               Stop3DMode
               
     End Select
     
     'Remove highlight
     RemoveHighlight True
     
     'Deselect current mode
     itmEditMode(mode).Checked = False
     
     'Hide menus
     mnuVertices.Visible = False
     mnuLines.Visible = False
     mnuSectors.Visible = False
     mnuThings.Visible = False
     
     'Change mode
     mode = Index
     submode = ESM_NONE
     
     'Select current mode
     itmEditMode(mode).Checked = True
     
     'Do stuff for this new mode
     Select Case Index
          
          Case EM_MOVE
               Set picMap.MouseIcon = imgCursor(1).Picture
               picMap.MousePointer = vbCustom
               tlbToolbar.Buttons("ModeMove").Value = tbrPressed
               selectedtype = EM_MOVE
               lblMode.Caption = "Move"
               
          Case EM_VERTICES
               picMap.MousePointer = vbNormal
               tlbToolbar.Buttons("ModeVertices").Value = tbrPressed
               mnuVertices.Visible = True
               selectedtype = EM_VERTICES
               lblMode.Caption = "Vertices"
               
          Case EM_LINES
               picMap.MousePointer = vbNormal
               tlbToolbar.Buttons("ModeLines").Value = tbrPressed
               mnuLines.Visible = True
               selectedtype = EM_LINES
               lblMode.Caption = "Lines"
               
               'Show toolbar buttons
               tlbToolbar.Buttons("LinesFlip").Visible = True
               tlbToolbar.Buttons("LinesCurve").Visible = True
               
          Case EM_SECTORS
               picMap.MousePointer = vbNormal
               tlbToolbar.Buttons("ModeSectors").Value = tbrPressed
               mnuSectors.Visible = True
               selectedtype = EM_SECTORS
               lblMode.Caption = "Sectors"
               
               'Show toolbar buttons
               tlbToolbar.Buttons("SectorsJoin").Visible = True
               tlbToolbar.Buttons("SectorsMerge").Visible = True
               tlbToolbar.Buttons("SectorsGradientBrightness").Visible = True
               tlbToolbar.Buttons("SectorsGradientFloors").Visible = True
               tlbToolbar.Buttons("SectorsGradientCeilings").Visible = True
               
          Case EM_THINGS
               picMap.MousePointer = vbNormal
               tlbToolbar.Buttons("ModeThings").Value = tbrPressed
               mnuThings.Visible = True
               selectedtype = EM_THINGS
               lblMode.Caption = "Things"
               
               'Show toolbar buttons
               tlbToolbar.Buttons("ThingsFilter").Visible = True
               
          Case EM_3D
               picMap.MousePointer = vbNormal
               lblMode.Caption = "3D Mode"
               tlbToolbar.Buttons("Mode3D").Value = tbrPressed
               
               'Check if anything exists
               If (numsectors < 2) Then
                    
                    'Cant build nodes
                    MsgBox "You need at least 2 sectors before going into 3D mode.", vbExclamation
                    
                    'Switch back to previous mode
                    itmEditMode_Click CInt(PreviousMode)
                    Exit Sub
               Else
                    
                    'Check if 3D mode is configured
                    While (Trim$(Config("videoadapterdesc")) = "") And (Val(Config("windowedvideo")) = 0)
                         
                         'Ask the user now
                         If (MsgBox("You have not yet configured your 3D Mode settings." & vbLf & _
                                   "Please click OK to configure the settings now.", vbInformation Or vbOKCancel) = VbMsgBoxResult.vbOK) Then
                              
                              'Show configuration
                              ShowConfiguration 8
                         Else
                              
                              'Switch back to previous mode
                              itmEditMode_Click CInt(PreviousMode)
                              Exit Sub
                         End If
                    Wend
                    
                    'Show status dialog
                    frmStatus.Show 0, frmMain
                    frmStatus.Refresh
                    frmMain.SetFocus
                    frmMain.Refresh
                    
                    'Must build nodes
                    If (mapnodeschanged) Or (TestStructures(TempWAD) = False) Then
                         
                         'Build nodes and check for errors
                         If MapBuild(True, False) = False Then
                              
                              'Unload status dialog
                              Screen.MousePointer = vbNormal
                              Unload frmStatus
                              Set frmStatus = Nothing
                              
                              'Nodebuilder failed!
                              MsgBox "The nodebuilder did not build the required structures." & vbLf & "Please check your map for errors or select a different nodebuilder!", vbCritical
                              
                              'Switch back to previous mode
                              itmEditMode_Click CInt(PreviousMode)
                              Exit Sub
                         End If
                    End If
                    
                    'Disable editing
                    picMap.Enabled = False
                    
                    'Set status
                    DisplayStatus "Building structures..."
                    
                    'Make triangles from SSECTORS
                    If PrepareStructures(TempWAD) Then
                         
                         'Set status
                         DisplayStatus "Switching to 3D mode..."
                         
                         'Hide
                         PreviousWindowstate = frmMain.WindowState
                         If (Val(Config("windowedvideo")) = 0) Then frmMain.WindowState = vbMinimized
                         
                         'Set the default settings
                         Init3DModeDefaults
                         
                         'Start the 3D Mode
                         If (Start3DMode) Then
                              
                              'Enable editing
                              picMap.Enabled = True
                              
                              'Run it now
                              Run3DMode
                         Else
                              
                              'Switch back to previous mode
                              itmEditMode_Click CInt(PreviousMode)
                         End If
                    Else
                         
                         'Switch back to previous mode
                         itmEditMode_Click CInt(PreviousMode)
                    End If
               End If
               
               'Leave here
               Exit Sub
     End Select
     
     'Redraw entire map
     RedrawMap
End Sub

Private Sub itmEditPaste_Click()
     Dim PasteMode As ENUM_PREFABINCLUDEMODE
     Dim PasteX As Long, PasteY As Long
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Remove highlight
     If (mapfile <> "") Then RemoveHighlight True
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Check if not in normal mode
     If (submode <> ESM_NONE) Then
          
          'Show message, cant right now
          MsgBox "Cannot perform Paste while in an operation." & vbLf & "Please finish your operation first!", vbExclamation
          
          'Leave here
          Exit Sub
     End If
     
     'Make undo
     CreateUndo "paste"
     
     'Check if we can paste at mouse location
     If MouseInside Then
          
          'Mouse position
          PasteX = LastX
          PasteY = LastY
     Else
          
          'Center of screen
          PasteX = picMap.ScaleLeft + picMap.ScaleWidth / 2
          PasteY = picMap.ScaleTop + picMap.ScaleHeight / 2
     End If
     
     'Determine paste mode by mode
     Select Case mode
          Case EM_VERTICES: PasteMode = PIM_VERTICES   'Paste vertices only
          Case EM_LINES: PasteMode = PIM_STRUCTURE
          Case EM_SECTORS: PasteMode = PIM_STRUCTURE
          Case EM_THINGS: PasteMode = PIM_THINGS
     End Select
     
     'Paste and check if anything was pasted
     If (InsertPrefab(ClipboardFile, PasteX, PasteY, PasteMode) > 0) Then
          
          'If in lines mode
          If (mode = EM_LINES) Then
               
               'Select lines from vertices
               SelectLinedefsFromVertices
               
          'If in sectors mode
          ElseIf (mode = EM_SECTORS) Then
               
               'Select sectors from vertices
               SelectLinedefsFromVertices
               SelectSectorsFromLinedefs
          End If
          
          'Grab it
          GrabX = PasteX
          GrabY = PasteY
          
          'Start pasting
          StartPasteOperation PasteX, PasteY
     Else
          
          'Nothing to paste
          MsgBox "Nothing to paste in the current mode.", vbInformation
     End If
     
     'Redraw map
     RedrawMap
End Sub

Private Sub itmEditRedo_Click()
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Remove highlight
     If (mapfile <> "") Then RemoveHighlight True
     
     'Check if not in normal mode
     If (submode <> ESM_NONE) Then
          
          'Show message, cant right now
          MsgBox "Cannot perform Redo while in an operation." & vbLf & "Please finish your operation first!", vbExclamation
          
          'Leave here
          Exit Sub
     End If
     
     'Change mousepointer
     Screen.MousePointer = vbHourglass
     
     'Remove selection and highlight
     RemoveHighlight
     RemoveSelection False
     
     'Do the redo
     PerformRedo
     
     'Update statusbar
     UpdateStatusBar
     
     'Redraw map
     RedrawMap
     
     'Reset mousepointer
     Screen.MousePointer = vbNormal
End Sub

Private Sub itmEditReplace_Click()
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Load dialog
     Load frmFind
     frmFind.Caption = "Find and Replace"
     
     'Show dialog
     frmFind.Show 1, Me
End Sub

Private Sub itmEditResize_Click()
     Dim Deselect As Boolean
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     If (submode <> ESM_PASTING) Then CancelCurrentOperation
     
     'Check if no selection is made
     If (numselected = 0) And (submode <> ESM_PASTING) Then
          
          'If no highlight exists, leave
          If (currentselected = -1) Then Exit Sub
          
          'Make selection from highlight
          Select Case mode
               Case EM_VERTICES: SelectCurrentVertex
               Case EM_LINES: SelectCurrentLine
               Case EM_SECTORS: SelectCurrentSector
               Case EM_THINGS: SelectCurrentThing
          End Select
          
          'Remove selection after edit
          Deselect = True
     End If
     
     'Make undo
     CreateUndo "resize"
     
     'Rotate
     Load frmResize
     frmResize.Show 1, Me
     
     'Check if we should deselect
     If Deselect Then
          
          'Remove selection
          RemoveSelection False
     End If
     
     'Redraw
     RedrawMap
End Sub

Private Sub itmEditRotate_Click()
     Dim Deselect As Boolean
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     If (submode <> ESM_PASTING) Then CancelCurrentOperation
     
     'Check if no selection is made
     If (numselected = 0) And (submode <> ESM_PASTING) Then
          
          'If no highlight exists, leave
          If (currentselected = -1) Then Exit Sub
          
          'Make selection from highlight
          Select Case mode
               Case EM_VERTICES: SelectCurrentVertex
               Case EM_LINES: SelectCurrentLine
               Case EM_SECTORS: SelectCurrentSector
               Case EM_THINGS: SelectCurrentThing
          End Select
          
          'Remove selection after edit
          Deselect = True
     End If
     
     'Make undo
     CreateUndo "rotate"
     
     'Rotate
     Load frmRotate
     frmRotate.Show 1, Me
     
     'Check if we should deselect
     If Deselect Then
          
          'Remove selection
          RemoveSelection False
     End If
     
     'Redraw
     RedrawMap
End Sub

Private Sub itmEditSnapToGrid_Click()
     
     'Toggle snap mode
     snapmode = Not snapmode
     itmEditSnapToGrid.Checked = snapmode
     tlbToolbar.Buttons("EditSnap").Value = Abs(snapmode)
     UpdateStatusBar
End Sub

Private Sub itmEditStitch_Click()
     
     'Toggle stitch mode
     stitchmode = Not stitchmode
     itmEditStitch.Checked = stitchmode
     tlbToolbar.Buttons("EditStitch").Value = Abs(stitchmode)
     UpdateStatusBar
End Sub

Private Sub itmEditUndo_Click()
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Remove highlight
     If (mapfile <> "") Then RemoveHighlight True
     
     'Check if not in normal mode
     If (submode <> ESM_NONE) Then
          
          'Just cancel the operation
          CancelCurrentOperation
     Else
          
          'Change mousepointer
          Screen.MousePointer = vbHourglass
          
          'Remove selection and highlight
          RemoveHighlight
          RemoveSelection False
          
          'Do the undo
          PerformUndo
          
          'Update statusbar
          UpdateStatusBar
          
          'Redraw map
          RedrawMap
          
          'Reset mousepointer
          Screen.MousePointer = vbNormal
     End If
End Sub

Private Sub itmFileBuild_Click()
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Change mousepointer
     Screen.MousePointer = vbHourglass
     
     'Check if anything exists
     If (numsectors < 2) Then
          
          'Cant build nodes
          MsgBox "You need at least 2 sectors before nodes can be build.", vbExclamation
     Else
          
          'Disable editing
          picMap.Enabled = False
          
          'Build nodes
          If MapBuild(False, False) = False Then
               
               'Nodebuilder failed!
               MsgBox "The nodebuilder did not build the required structures." & vbLf & "Please ensure you do not have any errors in your map!", vbCritical
          Else
               
               'Map changed
               mapchanged = True
          End If
     End If
     
     'Reset mousepointer
     Screen.MousePointer = vbDefault
     
     'Enable editing
     picMap.Enabled = True
End Sub

Private Sub itmFileCloseMap_Click()
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Unload map
     MapUnload
End Sub

Private Sub itmFile_Click(Index As Integer)
     Select Case Index
          
          Case 0: itmFileNew_Click
          Case 1: itmFileOpenMap_Click
          Case 2: itmFileCloseMap_Click
          
          Case 4: itmFileSaveMap_Click
          Case 5: itmFileSaveMapAs_Click
          Case 6: itmFileSaveMapInto_Click
          
          Case 8: itmFileExportMap_Click
          Case 9: itmFileExportPicture_Click
          
          Case 11: itmFileBuild_Click
          Case 12: itmFileTest_Click False
          
     End Select
End Sub

Private Sub itmFileExit_Click()
     Unload Me
End Sub

Private Sub itmFileExportMap_Click()
     Dim Result As String
     Dim FilterIndex As Long
     Dim HasChanged As Boolean
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Ensure the splash dialog is gone
     'It would be odd if it was still displayed though
     Unload frmSplash: Set frmSplash = Nothing
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Show save dialog
     Result = SaveFile(Me.hWnd, "Export Map As", "Doom/Heretic/Hexen WAD Files   *.wad|*.wad|Wavefront OBJ Files   *.obj|*.obj|All Files|*.*", mapfilename, cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt, FilterIndex)
     frmMain.Refresh

     'Check if not cancelled
     If Result <> "" Then
          
          'Add extension if needed
          If (LCase$(right$(Result, 4)) <> ".wad") And (FilterIndex = 1) Then Result = Result & ".wad"
          If (LCase$(right$(Result, 4)) <> ".obj") And (FilterIndex = 2) Then Result = Result & ".obj"
          
          'Check if exporting as OBJ
          If (FilterIndex = 2) Then
               
               'Delete file if it exists
               If (Dir(Result) <> "") Then Kill Result
               
               'Change mousepointer
               Screen.MousePointer = vbHourglass
               
               'Call the DLL API for writing the file
               ExportWavefrontObj Result, vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), things(0), numvertexes, numlinedefs, numsidedefs, numsectors, numthings
               
               'Reset mousepointer
               Screen.MousePointer = vbDefault
               
          Else
               
               'Make undo
               CreateUndo "Sidedefs compression", UGRP_NONE, 0, False
               
     '          'Give export dialog
     '          Load frmExport
     '          frmExport.txtTargetFile = Result
     '
     '          'Show dialog
     '          frmExport.Show 1, Me
     '
     '          'Check result
     '          If (frmExport.tag = "OK") Then
                    
                    'Keep changed status
                    HasChanged = mapchanged
                    
                    'Export the map
                    MapSave Result, SM_EXPORT, (Config("buildexportcompression") = vbChecked)
                    
                    'Restore changed status
                    '(The map is not saved to its original file)
                    mapchanged = HasChanged
     '          End If
     '
     '          'Unload dialog
     '          Unload frmExport
               
               'Restore undo
               PerformUndo False
               
               'Remove the redo
               WithdrawRedo
          End If
     End If
End Sub

Private Sub itmFileNew_Click()
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Ensure the splash dialog is gone
     Unload frmSplash
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Unload old map
     If (MapUnload) Then
          
          'Make new map
          MapNew "MAP01", True, True
     End If
End Sub

Private Sub itmFileOpenMap_Click()
     Dim Result As String
     Dim BeginFile As String
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Ensure the splash dialog is gone
     Unload frmSplash: Set frmSplash = Nothing
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Open dialog
     If (Config("recent").Exists("1") = True) Then BeginFile = Config("recent")("1")
     Result = OpenFile(Me.hWnd, "Open Map", "Doom/Heretic/Hexen WAD Files   *.wad|*.wad|All Files|*.*", BeginFile, cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames)
     frmMain.Refresh
     
     'Check if not cancelled
     If Result <> "" Then
          
          'Load the select map dialog
          Load frmMapSelect
          
          'Set the tag and caption
          frmMapSelect.tag = Result
          frmMapSelect.Caption = "Select Map from " & Dir(Result)
          
          'Show the select dialog
          frmMapSelect.Show 1, Me
     End If
End Sub

Private Sub itmFileRecent_Click(Index As Integer)
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Ensure the splash dialog is gone
     Unload frmSplash: Set frmSplash = Nothing
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Load the select map dialog
     Load frmMapSelect
     
     'Set the tag and caption
     frmMapSelect.tag = itmFileRecent(Index).tag
     frmMapSelect.Caption = "Select Map from " & Dir(itmFileRecent(Index).tag)
     
     'Show the select dialog
     frmMapSelect.Show 1, Me
End Sub

Public Sub itmFileSaveMap_Click()
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Check if changes were made
     If mapchanged Then
          
          'Check if the filename is different from filetitle
          If (mapsaved = True) And (Dir(mapfile) <> "") Then
               
               'Check if file is read-only
               If ((GetAttr(mapfile) And vbReadOnly) = vbReadOnly) Then
                    
                    'Save As...
                    itmFileSaveMapAs_Click
               Else
                    
                    'Save right away
                    If MapSave(mapfile, SM_SAVE) Then
                         
                         'Change the map filename
                         mapfilename = Dir(mapfile)
                         frmMain.Caption = App.Title & " - " & mapfilename & " (" & maplumpname & ")"
                         mapchanged = False
                         mapsaved = True
                    End If
               End If
          Else
               
               'Do a Save As...
               itmFileSaveMapAs_Click
          End If
     End If
End Sub

Private Sub itmFileSaveMapAs_Click()
     Dim Result As String
     Dim FilterIndex As Long
     Dim ResultOK As Boolean
     Dim SaveMethod As ENUM_SAVEMODES
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Ensure the splash dialog is gone
     'It would be odd if it was still displayed though
     Unload frmSplash: Set frmSplash = Nothing
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Continue
     Do
          'Show save dialog
          Result = SaveFile(Me.hWnd, "Save Map As", "Doom/Heretic/Hexen WAD Files   *.wad|*.wad|All Files|*.*", mapfilename, cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt, FilterIndex)
          frmMain.Refresh
          
          'Check if not cancelled
          If Result <> "" Then
               
               'Add extension
               If (LCase$(right$(Result, 4)) <> ".wad") And (FilterIndex = 1) Then Result = Result & ".wad"
               
               'Check if file exists
               If (Dir(Result) <> "") Then
                    
                    'Check if file is read-only
                    If ((GetAttr(Result) And vbReadOnly) = vbReadOnly) Then
                         
                         'Cannot save to read-only file
                         If (MsgBox("The file you selected is marked Read Only and cannot be changed." & vbLf & "Please enter a new file or select another file to overwrite.", vbExclamation Or vbOKCancel) = vbCancel) Then Result = ""
                    Else
                         
                         'OK
                         ResultOK = True
                    End If
               Else
                    
                    'OK
                    ResultOK = True
               End If
          End If
          
     'Continue until cancelled or OK
     Loop Until (Result = "") Or (ResultOK = True)
     
     'Check if not cancelled
     If (Result <> "") Then
          
          'Check if the same file as current file
          If (StrComp(Trim$(Result), Trim$(mapfile), vbTextCompare) = 0) Then
               
               'Save normally
               SaveMethod = SM_SAVE
          Else
               
               'Save as new file
               SaveMethod = SM_SAVEAS
          End If
          
          'Save the map here
          If MapSave(Result, SaveMethod) Then
               
               'Change the map filename
               mapfile = Result
               mapfilename = Dir(mapfile)
               frmMain.Caption = App.Title & " - " & mapfilename & " (" & maplumpname & ")"
               mapchanged = False
               mapsaved = True
          End If
     End If
End Sub

Private Sub itmFileSaveMapInto_Click()
     Dim Result As String
     Dim FilterIndex As Long
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Ensure the splash dialog is gone
     'It would be odd if it was still displayed though
     Unload frmSplash: Set frmSplash = Nothing
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Show save dialog
     Result = SaveFile(Me.hWnd, "Save Map Into", "Doom/Heretic/Hexen WAD Files   *.wad|*.wad|All Files|*.*", mapfilename, cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNFileMustExist, FilterIndex)
     frmMain.Refresh
     
     'Check if not cancelled
     If Result <> "" Then
          
          'Add extension
          If (LCase$(right$(Result, 4)) <> ".wad") And (FilterIndex = 1) Then Result = Result & ".wad"
          
          'Save the map here
          If MapSave(Result, SM_SAVEINTO) Then
               
               'Change the map filename
               mapfile = Result
               mapfilename = Dir(mapfile)
               frmMain.Caption = App.Title & " - " & mapfilename & " (" & maplumpname & ")"
               mapchanged = False
               mapsaved = True
          End If
     End If
End Sub

Public Sub itmFileTest_Click(ByVal ForceShowOptions As Boolean)
     Dim Parameters As String
     Dim TempMapFile As String
     Dim ExePath As String
     Dim OldWindowState As Long
     Dim OldScriptWindowState As Long
     Dim OldSelectColor As Long
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Check if we should show testing options first
     If (Config("testdialog") = vbChecked) Or (ForceShowOptions = True) Then ShowConfiguration 6
     
     'Check if not cancelled
     If (Not OptionsCancelled) Or (Config("testdialog") = vbUnchecked) Then
          
          'Check if the executable can be found
          If ((Config("testexec") <> "") And (Dir(Config("testexec")) <> "")) Then
               
               'Get executable path
               ExePath = PathOf(Config("testexec"))
               
               'Make temp filename
               TempMapFile = ExePath & "tempmap.wad"
               
               'Remove if file already exists
               If (Dir(TempMapFile) <> "") Then Kill TempMapFile
               
               'Save the map to temp file
               MapSave TempMapFile, SM_TEST
               
               'Get parameters
               Parameters = Config("testparams")
               
               'Replace placeholders in parameters
               Parameters = Replace$(Parameters, "%F", GetShortFileName(TempMapFile), , , vbTextCompare)
               Parameters = Replace$(Parameters, "%W", GetShortFileName(GetCurrentIWADFile), , , vbTextCompare)
               Parameters = Replace$(Parameters, "%D", Dir(GetCurrentIWADFile), , , vbTextCompare)
               Parameters = Replace$(Parameters, "%L", maplumpname, , , vbTextCompare)
               Parameters = Replace$(Parameters, "%A", GetShortFileName(addwadfile), , , vbTextCompare)
               Parameters = Replace$(Parameters, "%E", GetEpisodeNum(), , , vbTextCompare)
               Parameters = Replace$(Parameters, "%M", GetMapNum(), , , vbTextCompare)
               
               'Save selection color
               'This solves an bug with many software renderers on some systems
               OldSelectColor = GetSysColor(WCOLOR_HIGHLIGHT)
               
               'Minimize and hide
               OldWindowState = WindowState
               WindowState = vbMinimized
               
               'Same with Script editor if shown
               If (ScriptEditor) Then
                    OldScriptWindowState = frmScript.WindowState
                    frmScript.WindowState = vbMinimized
               End If
               
               'Disable editing
               picMap.Enabled = False
               
               'Focus to main window. This is a workaround from some
               'driver issues with microsoft mouse scrollwheel
               AppActivate frmMain.Caption
               frmMain.SetFocus
               
               'Launch the bitch
               If (Execute(Config("testexec"), Parameters, SW_SHOW, True) = False) Then MsgBox "Warning: Could not run the engine executable! Please try again later.", vbExclamation
               
               'Restore selection color
               SetSysColors 1, WCOLOR_HIGHLIGHT, OldSelectColor
               
               'Enable editing
               picMap.Enabled = True
               
               'Restore Script editor if shown
               If (ScriptEditor) Then frmScript.WindowState = OldScriptWindowState
               
               'Restore
               WindowState = OldWindowState
               Unload frmStatus
               
               'Remove temp map
               If (Dir(TempMapFile) <> "") Then Kill TempMapFile
          Else
               
               'Cant find engine
               MsgBox "Warning: Could not find the engine executable." & vbLf & "Please check your configuration!", vbExclamation
          End If
     End If
End Sub

Private Sub itmHelpAbout_Click()
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Ensure the splash dialog is gone
     Unload frmSplash: Set frmSplash = Nothing
     
     'Load splash
     Load frmSplash
     
     'Set version
     frmSplash.lblVersion = "Doom Builder version " & App.Major & "." & Format$(App.Minor, "00") & " build " & App.Revision
     
     'Change labels
     With frmSplash
          .lblStatus.Visible = False
          .lblVersion.Visible = True
          .lblWebsite.Visible = True
          .lblAbout1.Visible = True
     End With
     
     'Show about dialog
     frmSplash.Show 0, Me
End Sub

Private Sub itmHelpFAQ_Click()
     
     'Change mousepointer
     Screen.MousePointer = vbHourglass
     
     'Go to website
     Execute "http://www.doombuilder.com/builder_faq.php", "", SW_SHOW, False
     
     'Change mousepointer
     Screen.MousePointer = vbNormal
End Sub

Private Sub itmHelpWebsite_Click()
     
     'Change mousepointer
     Screen.MousePointer = vbHourglass
     
     'Go to website
     Execute "http://www.doombuilder.com", "", SW_SHOW, False
     
     'Change mousepointer
     Screen.MousePointer = vbNormal
End Sub

Private Sub itmLinesAlign_Click()
     Dim Deselect As Boolean
     Dim Indices As Variant
     Dim i As Long
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Check if no selection is made
     If (numselected = 0) Then
          
          'If no highlight exists, leave
          If (currentselected = -1) Then Exit Sub
          
          'Make selection from highlight
          SelectCurrentLine
          
          'Remove selection after edit
          Deselect = True
     End If
     
     'Show autoalign form
     frmAutoalign.Show 1, Me
     
     'Check result
     If (frmAutoalign.tag = "OK") Then
          
          'Make undo
          CreateUndo "autoalign textures", UGRP_TEXTUREALIGNMENT, -1, True
          
          'Get listing of selection
          Indices = selected.Items
          
          'Go for all linedefs
          For i = LBound(Indices) To UBound(Indices)
               
               'Align in X offsets?
               If (frmAutoalign.chkX.Value = vbChecked) Then
                    
                    'Remove linedef selections
                    ResetSelections things(0), 0, linedefs(0), numlinedefs, vertexes(0), 0, VarPtr(sectors(0)), 0
                    
                    'Check if it has a front side
                    If (linedefs(Indices(i)).s1 > -1) And (frmAutoalign.chkFront = vbChecked) Then
                         
                         'Autoalign this side
                         AlignTexturesX linedefs(Indices(i)).v1, sidedefs(linedefs(Indices(i)).s1).tx, frmAutoalign.lstTextures.List(frmAutoalign.lstTextures.ListIndex), False, Indices(i)
                    End If
                    
                    'Remove linedef selections
                    ResetSelections things(0), 0, linedefs(0), numlinedefs, vertexes(0), 0, VarPtr(sectors(0)), 0
                    
                    'Check if it has a back side
                    If (linedefs(Indices(i)).s2 > -1) And (frmAutoalign.chkBack = vbChecked) Then
                         
                         'Autoalign this side
                         AlignTexturesX linedefs(Indices(i)).v1, sidedefs(linedefs(Indices(i)).s2).tx, frmAutoalign.lstTextures.List(frmAutoalign.lstTextures.ListIndex), True, Indices(i)
                    End If
               End If
               
               'Align in Y offsets?
               If (frmAutoalign.chkY.Value = vbChecked) Then
                    
                    'Remove linedef selections
                    ResetSelections things(0), 0, linedefs(0), numlinedefs, vertexes(0), 0, VarPtr(sectors(0)), 0
                    
                    'Check if it has a front side
                    If (linedefs(Indices(i)).s1 > -1) And (frmAutoalign.chkFront = vbChecked) Then
                         
                         'Autoalign this side
                         AlignTexturesY linedefs(Indices(i)).v1, sidedefs(linedefs(Indices(i)).s1).ty, frmAutoalign.lstTextures.List(frmAutoalign.lstTextures.ListIndex), False, Indices(i)
                    End If
                    
                    'Remove linedef selections
                    ResetSelections things(0), 0, linedefs(0), numlinedefs, vertexes(0), 0, VarPtr(sectors(0)), 0
                    
                    'Check if it has a back side
                    If (linedefs(Indices(i)).s2 > -1) And (frmAutoalign.chkBack = vbChecked) Then
                         
                         'Autoalign this side
                         AlignTexturesY linedefs(Indices(i)).v1, sidedefs(linedefs(Indices(i)).s2).ty, frmAutoalign.lstTextures.List(frmAutoalign.lstTextures.ListIndex), True, Indices(i)
                    End If
               End If
          Next i
          
          'Remove linedef selections
          ResetSelections things(0), 0, linedefs(0), numlinedefs, vertexes(0), 0, VarPtr(sectors(0)), 0
          
          'Reselect lines
          ReselectLinedefs Indices(0)
     End If
     
     'Unload dialog
     Unload frmAutoalign
     
     'Check if we should deselect
     If Deselect Then
          
          'Remove selection
          RemoveSelection False
     End If
End Sub

Private Sub itmLinesCopy_Click()
     Dim CopyIndex As Long
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Check if not in normal mode
     If (submode <> ESM_NONE) Then
          
          'Show message, cant right now
          MsgBox "Cannot perform Copy Properties while in an operation." & vbLf & "Please finish your operation first!", vbExclamation
          
          'Leave here
          Exit Sub
     End If
     
     'Check if a selection is made
     If (numselected > 0) Then
          
          'Copy first selected
          CopyIndex = selected.Items(0)
          
     ElseIf (currentselected > -1) Then
          
          'Copy highlighted
          CopyIndex = currentselected
          
     Else
          
          'Nothing selected or highlighted
          Exit Sub
     End If
     
     'Copy the line properties
     CopiedLinedef = linedefs(CopyIndex)
     
     'Copy sidedefs if any
     If (linedefs(CopyIndex).s1 > -1) Then CopiedSidedef1 = sidedefs(linedefs(CopyIndex).s1)
     If (linedefs(CopyIndex).s2 > -1) Then CopiedSidedef2 = sidedefs(linedefs(CopyIndex).s2)
End Sub

Private Sub itmLinesCurve_Click()
     Dim Deselect As Boolean
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Check if no selection is made
     If (numselected = 0) Then
          
          'If no highlight exists, leave
          If (currentselected = -1) Then Exit Sub
          
          'Make selection from highlight
          SelectCurrentLine
          
          'Remove selection after edit
          Deselect = True
     End If
     
     'Make undo
     CreateUndo "curve"
     
     'Rotate
     Load frmCurve
     frmCurve.Show 1, Me
     
     'Check if we should deselect
     If Deselect Then
          
          'Remove selection
          RemoveSelection False
     End If
     
     'Redraw
     RedrawMap
End Sub

Private Sub itmLinesFlipLinedefs_Click()
     Dim Indices As Variant
     Dim i As Long
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Check if a selection is made
     If (numselected > 0) Then
          
          'Make undo
          CreateUndo "flip linedefs"
          
          'Go for all selected linedefs
          Indices = selected.Items
          For i = LBound(Indices) To UBound(Indices)
               
               'Flip vertices
               FlipLinedefVertices Indices(i)
               
               'Flip sidedefs (because they flipped with the linedef's vertices)
               FlipLinedefSidedefs Indices(i)
          Next i
          
     'Otherwise, check if a highlight is made
     ElseIf (currentselected > -1) Then
          
          'Make undo
          CreateUndo "flip linedefs"
          
          'Flip vertices
          FlipLinedefVertices currentselected
          
          'Flip sidedefs (because they flipped with the linedef's vertices)
          FlipLinedefSidedefs currentselected
     End If
     
     'Remove highlight
     RemoveHighlight True
     
     'Redraw map
     RedrawMap False
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
End Sub

Private Sub itmLinesFlipSidedefs_Click()
     Dim Indices As Variant
     Dim i As Long
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Check if a selection is made
     If (numselected > 0) Then
          
          'Make undo
          CreateUndo "flip sidedefs"
          
          'Go for all selected linedefs
          Indices = selected.Items
          For i = LBound(Indices) To UBound(Indices)
               
               'Flip sidedefs
               FlipLinedefSidedefs Indices(i)
          Next i
          
     'Otherwise, check if a highlight is made
     ElseIf (currentselected > -1) Then
          
          'Make undo
          CreateUndo "flip sidedefs"
          
          'Flip sidedefs
          FlipLinedefSidedefs currentselected
     End If
     
     'Remove highlight
     RemoveHighlight True
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
End Sub

Private Sub itmLinesPaste_Click()
     Dim Indices As Variant
     Dim i As Long
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Check if not in normal mode
     If (submode <> ESM_NONE) Then
          
          'Show message, cant right now
          MsgBox "Cannot perform Paste Properties while in an operation." & vbLf & "Please finish your operation first!", vbExclamation
          
          'Leave here
          Exit Sub
     End If
     
     'Make undo
     CreateUndo "paste linedef properties", , , True
     
     'Check if a selection is made
     If (numselected > 0) Then
          
          'Go for all selected linedefs
          Indices = selected.Items
          For i = LBound(Indices) To UBound(Indices)
               
               'Paste properties
               With linedefs(Indices(i))
                    .arg0 = CopiedLinedef.arg0
                    .arg1 = CopiedLinedef.arg1
                    .arg2 = CopiedLinedef.arg2
                    .arg3 = CopiedLinedef.arg3
                    .arg4 = CopiedLinedef.arg4
                    .effect = CopiedLinedef.effect
                    .Flags = (((CopiedLinedef.Flags And Not LDF_IMPASSIBLE) Or (.Flags And LDF_IMPASSIBLE)) And Not LDF_TWOSIDED) Or (.Flags And LDF_TWOSIDED)
                    .tag = CopiedLinedef.tag
               End With
               
               'Paste sidedef1 properties
               If (linedefs(Indices(i)).s1 > -1) And (CopiedLinedef.s1 > -1) Then
                    
                    With sidedefs(linedefs(Indices(i)).s1)
                         .Lower = CopiedSidedef1.Lower
                         .Middle = CopiedSidedef1.Middle
                         .tx = CopiedSidedef1.tx
                         .ty = CopiedSidedef1.ty
                         .Upper = CopiedSidedef1.Upper
                    End With
               End If
               
               'Paste sidedef2 properties
               If (linedefs(Indices(i)).s2 > -1) And (CopiedLinedef.s2 > -1) Then
                    
                    With sidedefs(linedefs(Indices(i)).s2)
                         .Lower = CopiedSidedef2.Lower
                         .Middle = CopiedSidedef2.Middle
                         .tx = CopiedSidedef2.tx
                         .ty = CopiedSidedef2.ty
                         .Upper = CopiedSidedef2.Upper
                    End With
               End If
          Next i
          
     'Otherwise, check if a highlight is made
     ElseIf (currentselected > -1) Then
          
          'Paste properties
          With linedefs(currentselected)
               .arg0 = CopiedLinedef.arg0
               .arg1 = CopiedLinedef.arg1
               .arg2 = CopiedLinedef.arg2
               .arg3 = CopiedLinedef.arg3
               .arg4 = CopiedLinedef.arg4
               .effect = CopiedLinedef.effect
               .Flags = CopiedLinedef.Flags
               .tag = CopiedLinedef.tag
          End With
          
          'Paste sidedef1 properties
          If (linedefs(currentselected).s1 > -1) And (CopiedLinedef.s1 > -1) Then
               
               With sidedefs(linedefs(currentselected).s1)
                    .Lower = CopiedSidedef1.Lower
                    .Middle = CopiedSidedef1.Middle
                    .tx = CopiedSidedef1.tx
                    .ty = CopiedSidedef1.ty
                    .Upper = CopiedSidedef1.Upper
               End With
          End If
          
          'Paste sidedef2 properties
          If (linedefs(currentselected).s2 > -1) And (CopiedLinedef.s2 > -1) Then
               
               With sidedefs(linedefs(currentselected).s2)
                    .Lower = CopiedSidedef2.Lower
                    .Middle = CopiedSidedef2.Middle
                    .tx = CopiedSidedef2.tx
                    .ty = CopiedSidedef2.ty
                    .Upper = CopiedSidedef2.Upper
               End With
          End If
     End If
     
     'Remove highlight
     RemoveHighlight True
     
     'Redraw map
     RedrawMap False
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
End Sub

Private Sub itmLinesSelect_Click(Index As Integer)
     Dim Indices As Variant
     Dim i As Long
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Check if a selection is made
     If (numselected > 0) Then
          
          'Go for all selected linedefs
          Indices = selected.Items
          For i = LBound(Indices) To UBound(Indices)
               
               'Check what to do
               Select Case Index
                    
                    'Select only 1 sided linedefs
                    Case 0:
                         
                         'Check if not 1-sided
                         If (linedefs(Indices(i)).s2 <> -1) And _
                            (linedefs(Indices(i)).s1 <> -1) Then
                              
                              'Remove selection
                              linedefs(Indices(i)).selected = 0
                              
                              'Remove index from selected objects
                              selected.Remove CStr(Indices(i))
                         End If
                         
                    'Select only 2 sided linedefs
                    Case 1:
                         
                         'Check if not 2-sided
                         If (linedefs(Indices(i)).s2 = -1) Or _
                            (linedefs(Indices(i)).s1 = -1) Then
                              
                              'Remove selection
                              linedefs(Indices(i)).selected = 0
                              
                              'Remove index from selected objects
                              selected.Remove CStr(Indices(i))
                         End If
               End Select
          Next i
          
          'Update number of selected items
          numselected = selected.Count
          
          'Redraw map
          RedrawMap
     End If
End Sub

Private Sub itmLinesSnapToGrid_Click()
     Dim Indices As Variant
     Dim i As Long
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Check if no selection is made, but a higlight
     If (numselected = 0) And (currentselected > -1) Then
          
          'Undo thise selecting after edit
          DeselectAfterEdit = True
          
          'Select current line
          SelectCurrentLine
     Else
          
          'Dont deselect after edit
          DeselectAfterEdit = False
     End If
     
     'Check if we have a selection
     If (numselected > 0) Then
          
          'Make Undo
          CreateUndo "snap to grid"
          
          'Go for all selected vertices
          Indices = SelectVerticesFromSelection.Items
          For i = LBound(Indices) To UBound(Indices)
               
               'Snap this vertex to grid now
               vertexes(Indices(i)).X = SnappedToGridX(vertexes(Indices(i)).X)
               vertexes(Indices(i)).Y = SnappedToGridY(vertexes(Indices(i)).Y)
          Next i
          
          'Map changed
          mapchanged = True
          mapnodeschanged = True
          
          'Deselect if we should
          If DeselectAfterEdit Then RemoveSelection False
          
          'Redraw map
          RedrawMap
     End If
End Sub

Private Sub itmPrefabInsert_Click()
     Dim PasteMode As ENUM_PREFABINCLUDEMODE
     Dim PasteX As Long, PasteY As Long
     Dim PasteFile As String
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Remove highlight
     If (mapfile <> "") Then RemoveHighlight True
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Check if not in normal mode
     If (submode <> ESM_NONE) Then
          
          'Show message, cant right now
          MsgBox "Cannot perform Prefab Insert while in an operation." & vbLf & "Please finish your operation first!", vbExclamation
          
          'Leave here
          Exit Sub
     End If
     
     'Browse for file
     picMap.Enabled = False
     If (Trim$(Config("prefabfolder")) <> "") Then
          PasteFile = OpenFile(Me.hWnd, "Insert Prefab file", "Doom Builder Prefab Files   *.dbp|*.dbp", Config("prefabfolder") & Dir(Config("prefabfolder") & "*"), cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames)
     Else
          PasteFile = OpenFile(Me.hWnd, "Insert Prefab file", "Doom Builder Prefab Files   *.dbp|*.dbp", "", cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames)
     End If
     DoEvents
     picMap.Enabled = True
     
     'Check if not cancelled
     If (Trim$(PasteFile) <> "") Then
          
          'Make undo
          CreateUndo "prefab insert"
          
          'Check if we can paste at mouse location
          If MouseInside Then
               
               'Mouse position
               PasteX = LastX
               PasteY = LastY
          Else
               
               'Center of screen
               PasteX = picMap.ScaleLeft + picMap.ScaleWidth / 2
               PasteY = picMap.ScaleTop + picMap.ScaleHeight / 2
          End If
          
          'Determine paste mode by mode
          Select Case mode
               Case EM_VERTICES: PasteMode = PIM_VERTICES   'Paste vertices only
               Case EM_LINES: PasteMode = PIM_STRUCTURE
               Case EM_SECTORS: PasteMode = PIM_STRUCTURE
               Case EM_THINGS: PasteMode = PIM_THINGS
          End Select
          
          'Paste and check if anything was pasted
          If (InsertPrefab(PasteFile, PasteX, PasteY, PasteMode) > 0) Then
               
               'If in lines mode
               If (mode = EM_LINES) Then
                    
                    'Select lines from vertices
                    SelectLinedefsFromVertices
                    
               'If in sectors mode
               ElseIf (mode = EM_SECTORS) Then
                    
                    'Select sectors from vertices
                    SelectLinedefsFromVertices
                    SelectSectorsFromLinedefs
               End If
               
               'Grab it
               GrabX = PasteX
               GrabY = PasteY
               
               'Start pasting
               StartPasteOperation PasteX, PasteY
               
               'Save last pasted
               LastPrefab = PasteFile
          Else
               
               'Nothing to paste
               MsgBox "This prefab does not have anything that can be inserted in the current mode.", vbInformation
          End If
          
          'Redraw map
          RedrawMap
     End If
End Sub

Private Sub itmPrefabPrevious_Click()
     Dim PasteMode As ENUM_PREFABINCLUDEMODE
     Dim PasteX As Long, PasteY As Long
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Remove highlight
     If (mapfile <> "") Then RemoveHighlight True
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Check if not in normal mode
     If (submode <> ESM_NONE) Then
          
          'Show message, cant right now
          MsgBox "Cannot perform Prefab Insert while in an operation." & vbLf & "Please finish your operation first!", vbExclamation
          
          'Leave here
          Exit Sub
     End If
     
     'Make undo
     CreateUndo "prefab insert"
     
     'Check if we can paste at mouse location
     If MouseInside Then
          
          'Mouse position
          PasteX = LastX
          PasteY = LastY
     Else
          
          'Center of screen
          PasteX = picMap.ScaleLeft + picMap.ScaleWidth / 2
          PasteY = picMap.ScaleTop + picMap.ScaleHeight / 2
     End If
     
     'Determine paste mode by mode
     Select Case mode
          Case EM_VERTICES: PasteMode = PIM_VERTICES   'Paste vertices only
          Case EM_LINES: PasteMode = PIM_STRUCTURE
          Case EM_SECTORS: PasteMode = PIM_STRUCTURE
          Case EM_THINGS: PasteMode = PIM_THINGS
     End Select
     
     'Paste and check if anything was pasted
     If (InsertPrefab(LastPrefab, PasteX, PasteY, PasteMode) > 0) Then
          
          'If in lines mode
          If (mode = EM_LINES) Then
               
               'Select lines from vertices
               SelectLinedefsFromVertices
               
          'If in sectors mode
          ElseIf (mode = EM_SECTORS) Then
               
               'Select sectors from vertices
               SelectLinedefsFromVertices
               SelectSectorsFromLinedefs
          End If
          
          'Grab it
          GrabX = PasteX
          GrabY = PasteY
          
          'Start pasting
          StartPasteOperation PasteX, PasteY
     Else
          
          'Nothing to paste
          MsgBox "This prefab does not have anything that can be inserted in the current mode.", vbInformation
     End If
     
     'Redraw map
     RedrawMap
End Sub

Private Sub itmPrefabQuick_Click(Index As Integer)
     Dim PasteMode As ENUM_PREFABINCLUDEMODE
     Dim PasteX As Long, PasteY As Long
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Remove highlight
     If (mapfile <> "") Then RemoveHighlight True
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Check if not in normal mode
     If (submode <> ESM_NONE) Then
          
          'Show message, cant right now
          MsgBox "Cannot perform Prefab Insert while in an operation." & vbLf & "Please finish your operation first!", vbExclamation
          
          'Leave here
          Exit Sub
     End If
     
     'Make undo
     CreateUndo "prefab insert"
     
     'Check if we can paste at mouse location
     If MouseInside Then
          
          'Mouse position
          PasteX = LastX
          PasteY = LastY
     Else
          
          'Center of screen
          PasteX = picMap.ScaleLeft + picMap.ScaleWidth / 2
          PasteY = picMap.ScaleTop + picMap.ScaleHeight / 2
     End If
     
     'Determine paste mode by mode
     Select Case mode
          Case EM_VERTICES: PasteMode = PIM_VERTICES   'Paste vertices only
          Case EM_LINES: PasteMode = PIM_STRUCTURE
          Case EM_SECTORS: PasteMode = PIM_STRUCTURE
          Case EM_THINGS: PasteMode = PIM_THINGS
     End Select
     
     'Paste and check if anything was pasted
     If (InsertPrefab(Config("quickprefab" & Index + 1), PasteX, PasteY, PasteMode) > 0) Then
          
          'If in lines mode
          If (mode = EM_LINES) Then
               
               'Select lines from vertices
               SelectLinedefsFromVertices
               
          'If in sectors mode
          ElseIf (mode = EM_SECTORS) Then
               
               'Select sectors from vertices
               SelectLinedefsFromVertices
               SelectSectorsFromLinedefs
          End If
          
          'Grab it
          GrabX = PasteX
          GrabY = PasteY
          
          'Start pasting
          StartPasteOperation PasteX, PasteY
          
          'Save last pasted
          LastPrefab = Config("quickprefab" & Index + 1)
     Else
          
          'Nothing to paste
          MsgBox "This prefab does not have anything that can be inserted in the current mode.", vbInformation
     End If
     
     'Redraw map
     RedrawMap
End Sub

Private Sub itmPrefabSaveSel_Click()
     Dim Result As String
     Dim FilterIndex As Long
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Browse for new file
     If (Trim$(Config("prefabfolder")) <> "") Then
          Result = SaveFile(Me.hWnd, "Save Prefab file", "Doom Builder Prefab Files   *.dbp|*.dbp|All Files|*.*", Config("prefabfolder") & Dir(Config("prefabfolder") & "*"), cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt, FilterIndex)
     Else
          Result = SaveFile(Me.hWnd, "Save Prefab file", "Doom Builder Prefab Files   *.dbp|*.dbp|All Files|*.*", "", cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt, FilterIndex)
     End If
     
     'Check if not cancelled
     If Result <> "" Then
          
          'Add extension
          If (LCase$(right$(Result, 4)) <> ".dbp") And (FilterIndex = 1) Then Result = Result & ".dbp"
          
          'Remove the file if already exists
          If (Dir(Result) <> "") Then Kill Result
          
          'Save the selection as prefab
          SavePrefabSelection Result
     End If
End Sub

Private Sub itmScriptEdit_Click(Index As Integer)
     Dim LumpName As String
     
     'Unload the script editor when its already loaded
     If (ScriptEditor) Then Unload frmScript
     
     'Get lump display name
     If (Trim$(itmScriptEdit(Index).tag) = "~") Then LumpName = maplumpname Else LumpName = itmScriptEdit(Index).tag
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Ensure the splash dialog is gone
     Unload frmSplash: Set frmSplash = Nothing
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Load the Script dialog
     Load frmScript
     
     'Set the lump name
     frmScript.lblLumpname.Caption = itmScriptEdit(Index).tag
     frmScript.Caption = "Doom Builder Map Script - " & LumpName
     
     'Load script
     frmScript.LoadLumpScript
     
     'Show the dialog
     frmScript.Show 0, Me
End Sub

Private Sub itmSectorsCopy_Click()
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Check if not in normal mode
     If (submode <> ESM_NONE) Then
          
          'Show message, cant right now
          MsgBox "Cannot perform Copy Properties while in an operation." & vbLf & "Please finish your operation first!", vbExclamation
          
          'Leave here
          Exit Sub
     End If
     
     'Check if a selection is made
     If (numselected > 0) Then
          
          'Copy the first selected sector
          CopiedSector = sectors(selected.Items(0))
          
     ElseIf (currentselected > -1) Then
          
          'Copy the highlighted sector
          CopiedSector = sectors(currentselected)
          
     Else
          
          'Nothing selected or highlighted
          Exit Sub
     End If
End Sub

Private Sub itmSectorsDecBrightness_Click()
     Dim Indices As Variant
     Dim i As Long
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Check if a selection is made
     If (numselected > 0) Then
          
          'Make undo
          CreateUndo "decrease brightness", , , True
          
          'Go for all selected sectors
          Indices = selected.Items
          For i = LBound(Indices) To UBound(Indices)
               
               'Decrease brightness
               sectors(Indices(i)).Brightness = sectors(Indices(i)).Brightness - 16
               If sectors(Indices(i)).Brightness < 0 Then sectors(Indices(i)).Brightness = 0
          Next i
          
     'Otherwise, check if a highlight is made
     ElseIf (currentselected > -1) Then
          
          'Make undo
          CreateUndo "decrease brightness", , , True
          
          'Decrease height
          sectors(currentselected).Brightness = sectors(currentselected).Brightness - 16
          If sectors(currentselected).Brightness < 0 Then sectors(currentselected).Brightness = 0
     End If
     
     'Reselect
     ChangeSectorsHighlight LastX, LastY, True
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
End Sub

Private Sub itmSectorsGradientBrightness_Click()
     Dim i As Long, p As Single
     Dim iSectors As Variant
     Dim v As Single, v1 As Single, v2 As Single
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Need at least 3 selected sectors
     If (numselected >= 3) Then
          
          'Make undo
          CreateUndo "gradient brightness"
          
          'Get sector indices
          iSectors = selected.Items
          
          'Get first and last values
          v1 = sectors(iSectors(LBound(iSectors))).Brightness
          v2 = sectors(iSectors(UBound(iSectors))).Brightness
          
          'Go for all selected sectors
          For i = LBound(iSectors) To UBound(iSectors)
               
               'Calculate interpolation
               p = i / UBound(iSectors)
               
               'Calculate new value
               v = v1 * (1 - p) + v2 * p
               
               'Apply the new value
               sectors(iSectors(i)).Brightness = v
          Next i
     End If
     
     'Redraw map
     RedrawMap False
     
     'Show highlight
     ShowHighlight LastX, LastY
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
End Sub

Private Sub itmSectorsGradientCeilings_Click()
     Dim i As Long, p As Single
     Dim iSectors As Variant
     Dim v As Single, v1 As Single, v2 As Single
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Need at least 3 selected sectors
     If (numselected >= 3) Then
          
          'Make undo
          CreateUndo "gradient ceilings"
          
          'Get sector indices
          iSectors = selected.Items
          
          'Get first and last values
          v1 = sectors(iSectors(LBound(iSectors))).hceiling
          v2 = sectors(iSectors(UBound(iSectors))).hceiling
          
          'Go for all selected sectors
          For i = LBound(iSectors) To UBound(iSectors)
               
               'Calculate interpolation
               p = i / UBound(iSectors)
               
               'Calculate new value
               v = v1 * (1 - p) + v2 * p
               
               'Apply the new value
               sectors(iSectors(i)).hceiling = v
          Next i
     End If
     
     'Redraw map
     RedrawMap False
     
     'Show highlight
     ShowHighlight LastX, LastY
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
End Sub


Private Sub itmSectorsGradientFloors_Click()
     Dim i As Long, p As Single
     Dim iSectors As Variant
     Dim v As Single, v1 As Single, v2 As Single
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Need at least 3 selected sectors
     If (numselected >= 3) Then
          
          'Make undo
          CreateUndo "gradient floors"
          
          'Get sector indices
          iSectors = selected.Items
          
          'Get first and last values
          v1 = sectors(iSectors(LBound(iSectors))).hfloor
          v2 = sectors(iSectors(UBound(iSectors))).hfloor
          
          'Go for all selected sectors
          For i = LBound(iSectors) To UBound(iSectors)
               
               'Calculate interpolation
               p = i / UBound(iSectors)
               
               'Calculate new value
               v = v1 * (1 - p) + v2 * p
               
               'Apply the new value
               sectors(iSectors(i)).hfloor = v
          Next i
     End If
     
     'Redraw map
     RedrawMap False
     
     'Show highlight
     ShowHighlight LastX, LastY
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
End Sub


Private Sub itmSectorsIncBrightness_Click()
     Dim Indices As Variant
     Dim i As Long
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Check if a selection is made
     If (numselected > 0) Then
          
          'Make undo
          CreateUndo "increase brightness", , , True
          
          'Go for all selected sectors
          Indices = selected.Items
          For i = LBound(Indices) To UBound(Indices)
               
               'Increase brightness
               sectors(Indices(i)).Brightness = sectors(Indices(i)).Brightness + 16
               If sectors(Indices(i)).Brightness > 255 Then sectors(Indices(i)).Brightness = 255
          Next i
          
     'Otherwise, check if a highlight is made
     ElseIf (currentselected > -1) Then
          
          'Make undo
          CreateUndo "increase brightness", , , True
          
          'Increase height
          sectors(currentselected).Brightness = sectors(currentselected).Brightness + 16
          If sectors(currentselected).Brightness > 255 Then sectors(currentselected).Brightness = 255
     End If
     
     'Reselect
     ChangeSectorsHighlight LastX, LastY, True
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
End Sub

Private Sub itmSectorsJoin_Click()
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Need at least 2 selected sectors
     If (numselected >= 2) Then
          
          'Make undo
          CreateUndo "join sectors"
          
          'Join selected sector together
          JoinSelectedSectors
          
          'DEBUG
          'DEBUG_FindUnusedSectors
     End If
     
     'Remove highlight
     RemoveHighlight True
     
     'Reselect
     ReselectSectors
     
     'Redraw map
     RedrawMap False
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
End Sub

Private Sub itmSectorsLowerCeiling_Click()
     Dim Indices As Variant
     Dim i As Long
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Check if a selection is made
     If (numselected > 0) Then
          
          'Make undo
          CreateUndo "lower ceiling", , , True
          
          'Go for all selected sectors
          Indices = selected.Items
          For i = LBound(Indices) To UBound(Indices)
               
               'Decrease height
               sectors(Indices(i)).hceiling = sectors(Indices(i)).hceiling - 8
          Next i
          
     'Otherwise, check if a highlight is made
     ElseIf (currentselected > -1) Then
          
          'Make undo
          CreateUndo "lower ceiling", , , True
          
          'Decrease height
          sectors(currentselected).hceiling = sectors(currentselected).hceiling - 8
     End If
     
     'Reselect
     ChangeSectorsHighlight LastX, LastY, True
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
End Sub

Private Sub itmSectorsLowerFloor_Click()
     Dim Indices As Variant
     Dim i As Long
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Check if a selection is made
     If (numselected > 0) Then
          
          'Make undo
          CreateUndo "lower floor", , , True
          
          'Go for all selected sectors
          Indices = selected.Items
          For i = LBound(Indices) To UBound(Indices)
               
               'Decrease height
               sectors(Indices(i)).hfloor = sectors(Indices(i)).hfloor - 8
          Next i
          
     'Otherwise, check if a highlight is made
     ElseIf (currentselected > -1) Then
          
          'Make undo
          CreateUndo "lower floor", , , True
          
          'Decrease height
          sectors(currentselected).hfloor = sectors(currentselected).hfloor - 8
     End If
     
     'Reselect
     ChangeSectorsHighlight LastX, LastY, True
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
End Sub

Private Sub itmSectorsMerge_Click()
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Need at least 2 selected sectors
     If (numselected >= 2) Then
          
          'Make undo
          CreateUndo "merge sectors"
          
          'Remove shared lindefs
          RemoveSelectedSharedLinedefs
          
          'Join selected sector together
          JoinSelectedSectors
          
          'DEBUG
          'DEBUG_FindUnusedSectors
     End If
     
     'Remove highlight
     RemoveHighlight True
     
     'Reselect
     ReselectSectors
     
     'Redraw map
     RedrawMap False
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
End Sub

Private Sub itmSectorsPaste_Click()
     Dim Indices As Variant
     Dim i As Long
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Check if not in normal mode
     If (submode <> ESM_NONE) Then
          
          'Show message, cant right now
          MsgBox "Cannot perform Paste Properties while in an operation." & vbLf & "Please finish your operation first!", vbExclamation
          
          'Leave here
          Exit Sub
     End If
     
     'Make undo
     CreateUndo "paste sector properties", , , True
     
     'Check if a selection is made
     If (numselected > 0) Then
          
          'Go for all selected linedefs
          Indices = selected.Items
          For i = LBound(Indices) To UBound(Indices)
               
               'Paste properties
               With sectors(Indices(i))
                    .Brightness = CopiedSector.Brightness
                    .hceiling = CopiedSector.hceiling
                    .hfloor = CopiedSector.hfloor
                    .special = CopiedSector.special
                    .tag = CopiedSector.tag
                    .tceiling = CopiedSector.tceiling
                    .tfloor = CopiedSector.tfloor
               End With
          Next i
          
     'Otherwise, check if a highlight is made
     ElseIf (currentselected > -1) Then
          
          'Paste properties
          With sectors(currentselected)
               .Brightness = CopiedSector.Brightness
               .hceiling = CopiedSector.hceiling
               .hfloor = CopiedSector.hfloor
               .special = CopiedSector.special
               .tag = CopiedSector.tag
               .tceiling = CopiedSector.tceiling
               .tfloor = CopiedSector.tfloor
          End With
     End If
     
     'Remove highlight
     RemoveHighlight True
     
     'Redraw map
     RedrawMap False
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
End Sub

Private Sub itmSectorsRaiseCeiling_Click()
     Dim Indices As Variant
     Dim i As Long
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Check if a selection is made
     If (numselected > 0) Then
          
          'Make undo
          CreateUndo "raise ceiling", , , True
          
          'Go for all selected sectors
          Indices = selected.Items
          For i = LBound(Indices) To UBound(Indices)
               
               'Increase height
               sectors(Indices(i)).hceiling = sectors(Indices(i)).hceiling + 8
          Next i
          
     'Otherwise, check if a highlight is made
     ElseIf (currentselected > -1) Then
          
          'Make undo
          CreateUndo "raise ceiling", , , True
          
          'Increase height
          sectors(currentselected).hceiling = sectors(currentselected).hceiling + 8
     End If
     
     'Reselect
     ChangeSectorsHighlight LastX, LastY, True
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
End Sub

Private Sub itmSectorsRaiseFloor_Click()
     Dim Indices As Variant
     Dim i As Long
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Check if a selection is made
     If (numselected > 0) Then
          
          'Make undo
          CreateUndo "raise floor", , , True
          
          'Go for all selected sectors
          Indices = selected.Items
          For i = LBound(Indices) To UBound(Indices)
               
               'Increase height
               sectors(Indices(i)).hfloor = sectors(Indices(i)).hfloor + 8
          Next i
          
     'Otherwise, check if a highlight is made
     ElseIf (currentselected > -1) Then
          
          'Make undo
          CreateUndo "raise floor", , , True
          
          'Increase height
          sectors(currentselected).hfloor = sectors(currentselected).hfloor + 8
     End If
     
     'Reselect
     ChangeSectorsHighlight LastX, LastY, True
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
End Sub

Private Sub itmSectorsSnapToGrid_Click()
     Dim Indices As Variant
     Dim i As Long
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Check if no selection is made, but a higlight
     If (numselected = 0) And (currentselected > -1) Then
          
          'Undo thise selecting after edit
          DeselectAfterEdit = True
          
          'Select current sector
          SelectCurrentSector
     Else
          
          'Dont deselect after edit
          DeselectAfterEdit = False
     End If
     
     'Check if we have a selection
     If (numselected > 0) Then
          
          'Make Undo
          CreateUndo "snap to grid"
          
          'Go for all selected vertices
          Indices = SelectVerticesFromSelection.Items
          For i = LBound(Indices) To UBound(Indices)
               
               'Snap this vertex to grid now
               vertexes(Indices(i)).X = SnappedToGridX(vertexes(Indices(i)).X)
               vertexes(Indices(i)).Y = SnappedToGridY(vertexes(Indices(i)).Y)
          Next i
          
          'Map changed
          mapchanged = True
          mapnodeschanged = True
          
          'Deselect if we should
          If DeselectAfterEdit Then RemoveSelection False
          
          'Redraw map
          RedrawMap
     End If
End Sub

Private Sub itmThingsCopy_Click()
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Check if not in normal mode
     If (submode <> ESM_NONE) Then
          
          'Show message, cant right now
          MsgBox "Cannot perform Copy Properties while in an operation." & vbLf & "Please finish your operation first!", vbExclamation
          
          'Leave here
          Exit Sub
     End If
     
     'Check if a selection is made
     If (numselected > 0) Then
          
          'Copy the first selected thing
          CopiedThing = things(selected.Items(0))
          
     ElseIf (currentselected > -1) Then
          
          'Copy the highlighted thing
          CopiedThing = things(currentselected)
          
     Else
          
          'Nothing selected or highlighted
          Exit Sub
     End If
End Sub

Private Sub itmThingsFilter_Click()
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Show filter dialog
     Load frmThingFilter
     frmThingFilter.Show 1, Me
End Sub

Private Sub itmThingsPaste_Click()
     Dim Indices As Variant
     Dim i As Long
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Check if not in normal mode
     If (submode <> ESM_NONE) Then
          
          'Show message, cant right now
          MsgBox "Cannot perform Paste Properties while in an operation." & vbLf & "Please finish your operation first!", vbExclamation
          
          'Leave here
          Exit Sub
     End If
     
     'Make undo
     CreateUndo "paste thing properties", , , True
     
     'Check if a selection is made
     If (numselected > 0) Then
          
          'Go for all selected linedefs
          Indices = selected.Items
          For i = LBound(Indices) To UBound(Indices)
               
               'Paste properties
               With things(Indices(i))
                    .angle = CopiedThing.angle
                    .arg0 = CopiedThing.arg0
                    .arg1 = CopiedThing.arg1
                    .arg2 = CopiedThing.arg2
                    .arg3 = CopiedThing.arg3
                    .arg4 = CopiedThing.arg4
                    .Color = CopiedThing.Color
                    .effect = CopiedThing.effect
                    .Flags = CopiedThing.Flags
                    .image = CopiedThing.image
                    .tag = CopiedThing.tag
                    .thing = CopiedThing.thing
                    .Z = CopiedThing.Z
               End With
          Next i
          
     'Otherwise, check if a highlight is made
     ElseIf (currentselected > -1) Then
          
          'Paste properties
          With things(currentselected)
               .angle = CopiedThing.angle
               .arg0 = CopiedThing.arg0
               .arg1 = CopiedThing.arg1
               .arg2 = CopiedThing.arg2
               .arg3 = CopiedThing.arg3
               .arg4 = CopiedThing.arg4
               .Color = CopiedThing.Color
               .effect = CopiedThing.effect
               .Flags = CopiedThing.Flags
               .image = CopiedThing.image
               .tag = CopiedThing.tag
               .thing = CopiedThing.thing
               .Z = CopiedThing.Z
          End With
     End If
     
     'Remove highlight
     RemoveHighlight True
     
     'Redraw map
     RedrawMap False
     
     'Map changed
     mapchanged = True
End Sub

Private Sub itmThingsSnapToGrid_Click()
     Dim Indices As Variant
     Dim i As Long
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Check if no selection is made, but a higlight
     If (numselected = 0) And (currentselected > -1) Then
          
          'Undo thise selecting after edit
          DeselectAfterEdit = True
          
          'Select current thing
          SelectCurrentThing
     Else
          
          'Youll only understand this if you are 1337 H4X0R!!! MWhahahah!!!
          DeselectAfterEdit = False
     End If
     
     'Check if we have a selection
     If (numselected > 0) Then
          
          'Make Undo
          CreateUndo "snap to grid"
          
          'Go for all selected vertices
          Indices = selected.Items
          For i = LBound(Indices) To UBound(Indices)
               
               'Snap this vertex to grid now
               things(Indices(i)).X = SnappedToGridX(things(Indices(i)).X)
               things(Indices(i)).Y = SnappedToGridY(things(Indices(i)).Y)
          Next i
          
          'Map changed
          mapchanged = True
          mapnodeschanged = True
          
          'Deselect if we should
          If DeselectAfterEdit Then RemoveSelection False
          
          'Redraw map
          RedrawMap
     End If
End Sub

Private Sub itmToolsClearTextures_Click()
     Dim ld As Long
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Make undo
     CreateUndo "clear unused textures"
     
     'Check if a selection is made
     If (numselected > 0) Then
          
          'Go for all linedefs
          For ld = 0 To (numlinedefs - 1)
               
               'Check if selected or none selected at all
               If (linedefs(ld).selected <> 0) Or (mode <> EM_LINES) Then
                    
                    'Check if linedef has a front sidedef
                    If (linedefs(ld).s1 > -1) Then
                         
                         'Remove upper texture if not required
                         If Not RequiresS1Upper(ld) Then sidedefs(linedefs(ld).s1).Upper = "-"
                         
                         'Remove middle texture if not required
                         'If Not RequiresS1Middle(ld) Then sidedefs(linedefs(ld).s1).Middle = "-")
                         
                         'Remove lower texture if not required
                         If Not RequiresS1Lower(ld) Then sidedefs(linedefs(ld).s1).Lower = "-"
                    End If
                    
                    'Check if linedef has a back sidedef
                    If (linedefs(ld).s2 > -1) Then
                         
                         'Remove upper texture if not required
                         If Not RequiresS2Upper(ld) Then sidedefs(linedefs(ld).s2).Upper = "-"
                         
                         'Remove middle texture if not required
                         'If Not RequiresS2Middle(ld) Then sidedefs(linedefs(ld).s2).Middle = "-")
                         
                         'Remove lower texture if not required
                         If Not RequiresS2Lower(ld) Then sidedefs(linedefs(ld).s2).Lower = "-"
                    End If
               End If
          Next ld
     Else
          
          'Please make a selection!
          'This may also remove textures from your changing sectors.
          MsgBox "Please make a selection so that you dont accedentially remove textures from sectors that may need them after being triggered.", vbExclamation
     End If
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
End Sub

Public Sub itmToolsConfiguration_Click()
     
     'Configure
     ShowConfiguration 1
End Sub

Private Sub itmToolsFindErrors_Click()
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Clear selection
     RemoveSelection True
     
     'Load dialog
     Load frmErrorCheck
     
     'Show dialog
     frmErrorCheck.Show 1, Me
End Sub

Private Sub itmToolsFixTextures_Click()
     Dim ld As Long
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Make undo
     CreateUndo "fix missing textures"
     
     'Ensure valid textures are used to build with
     CorrectDefaultTextures
     
     'Go for all linedefs
     For ld = 0 To (numlinedefs - 1)
          
          'Check if selected or none selected at all
          If (linedefs(ld).selected <> 0) Or (numselected = 0) Or (mode <> EM_LINES) Then
               
               'Check if linedef has a front sidedef
               If (linedefs(ld).s1 > -1) Then
                    
                    'Ensure upper texture if required
                    If RequiresS1Upper(ld) And Not IsTextureName(sidedefs(linedefs(ld).s1).Upper) Then sidedefs(linedefs(ld).s1).Upper = Config("defaulttexture")("upper")
                    
                    'Ensure middle texture if required
                    If RequiresS1Middle(ld) And Not IsTextureName(sidedefs(linedefs(ld).s1).Middle) Then sidedefs(linedefs(ld).s1).Middle = Config("defaulttexture")("middle")
                    
                    'Ensure lower texture if required
                    If RequiresS1Lower(ld) And Not IsTextureName(sidedefs(linedefs(ld).s1).Lower) Then sidedefs(linedefs(ld).s1).Lower = Config("defaulttexture")("lower")
               End If
               
               'Check if linedef has a back sidedef
               If (linedefs(ld).s2 > -1) Then
                    
                    'Ensure upper texture if required
                    If RequiresS2Upper(ld) And Not IsTextureName(sidedefs(linedefs(ld).s2).Upper) Then sidedefs(linedefs(ld).s2).Upper = Config("defaulttexture")("upper")
                    
                    'Ensure middle texture if required
                    If RequiresS2Middle(ld) And Not IsTextureName(sidedefs(linedefs(ld).s2).Middle) Then sidedefs(linedefs(ld).s2).Middle = Config("defaulttexture")("middle")
                    
                    'Ensure lower texture if required
                    If RequiresS2Lower(ld) And Not IsTextureName(sidedefs(linedefs(ld).s2).Lower) Then sidedefs(linedefs(ld).s2).Lower = Config("defaulttexture")("lower")
               End If
          End If
     Next ld
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
End Sub

Private Sub itmToolsFixZeroLinedefs_Click()
     Dim ld As Long
     Dim Count As Long
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Make undo
     CreateUndo "fix zero-length linedefs"
     
     'Go for all linedefs
     ld = numlinedefs - 1
     Do While (ld >= 0)
          
          'Check if linedef refers to same vertices
          If (linedefs(ld).v1 = linedefs(ld).v2) Then
               
               'Simply remove the linedef
               RemoveLinedef ld, True, True, True
               Count = Count + 1
          Else
               
               'Check if both vertices are at same location
               If (CLng(vertexes(linedefs(ld).v1).X) = CLng(vertexes(linedefs(ld).v2).X)) And _
                  (CLng(vertexes(linedefs(ld).v1).Y) = CLng(vertexes(linedefs(ld).v2).Y)) Then
                    
                    'Stitch the vertices
                    StitchVertices linedefs(ld).v1, linedefs(ld).v2
                    Count = Count + 1
               End If
          End If
          
          'Next linedef
          ld = ld - 1
     Loop
     
     'Redraw map
     RedrawMap False
     
     'Report
     If (Count > 0) Then
          
          'Show result
          MsgBox Count & " Zero-Length Linedefs have been solved.", vbInformation
          
          'Map changed
          mapchanged = True
          mapnodeschanged = True
     Else
          
          'Nothing found
          MsgBox "No Zero-Length Linedefs found.", vbInformation
     End If
End Sub

Private Sub itmVerticesClearUnused_Click()
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Make undo
     CreateUndo "remove unused vertices"
     
     'Rmeove unused vertices
     RemoveUnusedVertices
     
     'Remove selection
     RemoveSelection False
     
     'Redraw map
     RedrawMap
End Sub

Private Sub itmVerticesSnapToGrid_Click()
     Dim Indices As Variant
     Dim i As Long
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Check if no selection is made, but a higlight
     If (numselected = 0) And (currentselected > -1) Then
          
          'Undo thise selecting after edit
          DeselectAfterEdit = True
          
          'Select current vertex
          SelectCurrentVertex
     Else
          
          'Youll only understand this if you are 1337 H4X0R!!! MWhahahah!!!
          DeselectAfterEdit = False
     End If
     
     'Check if we have a selection
     If (numselected > 0) Then
          
          'Make Undo
          CreateUndo "snap to grid"
          
          'Go for all selected vertices
          Indices = selected.Items
          For i = LBound(Indices) To UBound(Indices)
               
               'Snap this vertex to grid now
               vertexes(Indices(i)).X = SnappedToGridX(vertexes(Indices(i)).X)
               vertexes(Indices(i)).Y = SnappedToGridY(vertexes(Indices(i)).Y)
          Next i
          
          'Map changed
          mapchanged = True
          mapnodeschanged = True
          
          'Deselect if we should
          If DeselectAfterEdit Then RemoveSelection False
          
          'Redraw map
          RedrawMap
     End If
End Sub

Private Sub KeypressGeneral(ByVal ShortcutCode As Long, ByRef DoUpdateStatusBar As Boolean, ByRef DoRedrawMap As Boolean)
     Dim Xdiff As Long, Ydiff As Long
     Dim NewZ As Single
     Dim Keybinds As Variant
     Dim AllowF7Message As Boolean
     Dim i As Long
     
     'Check what shortcut is pressed in ANY mode
     Select Case ShortcutCode
          
          Case Config("shortcuts")("zoomin")
               
               'Increase zoom
               NewZ = ViewZoom * (1 + Config("zoomspeed") / 1000)
               If NewZ > 100 Then NewZ = 100
               If (Val(Config("zoommouse")) <> 0) And (MouseInside = True) Then
                    Xdiff = ((ScreenWidth / NewZ) - (ScreenWidth / ViewZoom)) * ((LastX - picMap.ScaleLeft) / picMap.ScaleWidth)
                    Ydiff = ((ScreenHeight / NewZ) - (ScreenHeight / ViewZoom)) * ((LastY - picMap.ScaleTop) / picMap.ScaleHeight)
               Else
                    Xdiff = ((ScreenWidth / NewZ) - (ScreenWidth / ViewZoom)) / 2
                    Ydiff = ((ScreenHeight / NewZ) - (ScreenHeight / ViewZoom)) / 2
               End If
               ChangeView ViewLeft - Xdiff, ViewTop - Ydiff, NewZ
               DoRedrawMap = True
               DoUpdateStatusBar = True
               
          Case Config("shortcuts")("zoomout")
               
               'Decrease zoom
               NewZ = ViewZoom * (1 - Config("zoomspeed") / 1000)
               If NewZ < 0.05 Then NewZ = 0.05
               If (Val(Config("zoommouse")) <> 0) And (MouseInside = True) Then
                    Xdiff = ((ScreenWidth / NewZ) - (ScreenWidth / ViewZoom)) * ((LastX - picMap.ScaleLeft) / picMap.ScaleWidth)
                    Ydiff = ((ScreenHeight / NewZ) - (ScreenHeight / ViewZoom)) * ((LastY - picMap.ScaleTop) / picMap.ScaleHeight)
               Else
                    Xdiff = ((ScreenWidth / NewZ) - (ScreenWidth / ViewZoom)) / 2
                    Ydiff = ((ScreenHeight / NewZ) - (ScreenHeight / ViewZoom)) / 2
               End If
               ChangeView ViewLeft - Xdiff, ViewTop - Ydiff, NewZ
               DoRedrawMap = True
               DoUpdateStatusBar = True
               
          Case Config("shortcuts")("scrollleft")
               
               'Move left
               LastX = LastX - Config("scrollpixels") / ViewZoom
               ChangeView ViewLeft - Config("scrollpixels") / ViewZoom, ViewTop, ViewZoom
               DoRedrawMap = True
               DoUpdateStatusBar = True
               
               'Update cursor position in statusbar
               stbStatus.Panels("mousex").Text = "X " & CLng(LastX)
               stbStatus.Panels("mousey").Text = "Y " & -CLng(LastY)
               
          Case Config("shortcuts")("scrollright")
               
               'Move right
               LastX = LastX + Config("scrollpixels") / ViewZoom
               ChangeView ViewLeft + Config("scrollpixels") / ViewZoom, ViewTop, ViewZoom
               DoRedrawMap = True
               DoUpdateStatusBar = True
               
               'Update cursor position in statusbar
               stbStatus.Panels("mousex").Text = "X " & CLng(LastX)
               stbStatus.Panels("mousey").Text = "Y " & -CLng(LastY)
               
          Case Config("shortcuts")("scrollup")
               
               'Move up
               LastY = LastY - Config("scrollpixels") / ViewZoom
               ChangeView ViewLeft, ViewTop - Config("scrollpixels") / ViewZoom, ViewZoom
               DoRedrawMap = True
               DoUpdateStatusBar = True
               
               'Update cursor position in statusbar
               stbStatus.Panels("mousex").Text = "X " & CLng(LastX)
               stbStatus.Panels("mousey").Text = "Y " & -CLng(LastY)
               
          Case Config("shortcuts")("scrolldown")
               
               'Move down
               LastY = LastY + Config("scrollpixels") / ViewZoom
               ChangeView ViewLeft, ViewTop + Config("scrollpixels") / ViewZoom, ViewZoom
               DoRedrawMap = True
               DoUpdateStatusBar = True
               
               'Update cursor position in statusbar
               stbStatus.Panels("mousex").Text = "X " & CLng(LastX)
               stbStatus.Panels("mousey").Text = "Y " & -CLng(LastY)
               
          Case Config("shortcuts")("gridinc")
               
               'Decrease grid
               gridsizex = gridsizex / 2
               gridsizey = gridsizey / 2
               If (gridsizex < 1) Then gridsizex = 1
               If (gridsizey < 1) Then gridsizey = 1
               DoRedrawMap = True
               DoUpdateStatusBar = True
               
          Case Config("shortcuts")("griddec")
               
               'Increase grid
               gridsizex = gridsizex * 2
               gridsizey = gridsizey * 2
               If (gridsizex > 1024) Then gridsizex = 1024
               If (gridsizey > 1024) Then gridsizey = 1024
               DoRedrawMap = True
               DoUpdateStatusBar = True
               
          Case Config("shortcuts")("deselectall")
               
               'Only in normal mode
               If (submode = ESM_NONE) Then
                    
                    'Deselect and remove highlight
                    RemoveHighlight True
                    RemoveSelection True
               End If
               
          Case Config("shortcuts")("cancel")
               
               'Cancel if in drawing operation
               CancelCurrentOperation
               
          Case Config("shortcuts")("drawsector2")
               
               'Cancel if in drawing operation
               CancelCurrentOperation
               
               'Switch to lines mode if needed
               If (mode <> EM_LINES) And (mode <> EM_SECTORS) Then itmEditMode_Click EM_LINES
               
               'Start drawing
               StartDrawOperation
               
          Case Config("shortcuts")("reversedrawing")
               
               'Check if in drawing operation
               If (submode = ESM_DRAWING) Then RevertDrawingOperation
               
          Case Config("shortcuts")("place3dstart")
               
               'Check if mouse cursor is on the map
               If MouseInside Then
                    
                    'Place 3D Start mode Thing
                    Place3DModeStart LastX, LastY
                    DoUpdateStatusBar = True
                    DoRedrawMap = True
               End If
               
          Case Config("shortcuts")("editquickmove")
               
               'Toggle move mode
               If (mode <> EM_MOVE) Then
                    
                    'Switch to move mode now
                    'itmEditMode_Click EM_MOVE
                    
                    'Change mousecursor
                    Set picMap.MouseIcon = imgCursor(2).Picture
                    picMap.MousePointer = vbCustom
                    
                    'And drag now
                    submode = ESM_MOVING
               End If
               
          Case Config("shortcuts")("togglebar")
               
               'Toggle the Info Bar
               InfoBarToggle
               
          'Lession 1: Dont post flames against Doom Builder :P
          Case vbKeyF7
               
               'Have we hold F7 a while?
               If (F7Count = 100) Then
                    
                    'Presume we can display the message
                    AllowF7Message = True
                    
                    'Is there really no key assigned to F7?
                    Keybinds = Config("shortcuts").Items
                    For i = LBound(Keybinds) To UBound(Keybinds)
                         
                         'We dont want to interrupt any real features
                         If (Keybinds(i) = vbKeyF7) Then AllowF7Message = False
                    Next i
                    
                    'Bla
                    If (AllowF7Message) Then MsgBox "I am the great doom editing genius. Pay me. Pay me now!", vbExclamation, "Not Doom Builder"
                    
                    'Reset the count
                    F7Count = 0
               Else
                    
                    'Count the keypress
                    F7Count = F7Count + 1
               End If
     End Select
End Sub

Private Sub KeyreleaseGeneral(ByVal ShortcutCode As Long, ByRef DoUpdateStatusBar As Boolean, ByRef DoRedrawMap As Boolean)
     
     'Check what shortcut is released in ANY mode
     Select Case ShortcutCode
          
          Case Config("shortcuts")("editquickmove")
               
               'Return from move mode
               If (mode <> EM_MOVE) Then
                    
                    'No longer moving
                    Set picMap.MouseIcon = imgCursor(1).Picture
                    picMap.MousePointer = vbNormal
                    submode = ESM_NONE
               End If
     End Select
End Sub


Private Sub KeypressLines(ByVal ShortcutCode As Long)
     
     'Check what key is pressed
     Select Case ShortcutCode
          
          Case Config("shortcuts")("drawsector")
               
               'Only in normal mode
               If (submode = ESM_NONE) Then
                    
                    'Start drawing
                    StartDrawOperation
               End If
          
          Case Config("shortcuts")("createsector")
               
               'Only in normal mode
               If (submode = ESM_NONE) Then
                    
                    'Insert vertex
                    If MouseInside Then CreateSectorHere LastX, LastY
               End If
     End Select
End Sub

Private Sub KeypressMenus(ByVal ShortcutCode As Long)
     
     'Check what shortcut is pressed for menus
     Select Case ShortcutCode
          Case Config("shortcuts")("filenew"): mnuFile_Click: If itmFile(0).Enabled Then itmFileNew_Click
          Case Config("shortcuts")("fileopen"): mnuFile_Click: If itmFile(1).Enabled Then itmFileOpenMap_Click
          Case Config("shortcuts")("fileclose"): mnuFile_Click: If itmFile(2).Enabled Then itmFileCloseMap_Click
          Case Config("shortcuts")("filesave"): mnuFile_Click: If itmFile(4).Enabled Then itmFileSaveMap_Click
          Case Config("shortcuts")("filesaveas"): mnuFile_Click: If itmFile(5).Enabled Then itmFileSaveMapAs_Click
          Case Config("shortcuts")("filesaveinto"): mnuFile_Click: If itmFile(6).Enabled Then itmFileSaveMapInto_Click
          Case Config("shortcuts")("filebuildnodes"): mnuFile_Click: If itmFile(11).Enabled Then itmFileBuild_Click
          Case Config("shortcuts")("filetest"): mnuFile_Click: If itmFile(12).Enabled Then itmFileTest_Click False
          Case Config("shortcuts")("filetest2"): mnuFile_Click: If itmFile(12).Enabled Then itmFileTest_Click True
          Case Config("shortcuts")("fileexport"): mnuFile_Click: If itmFile(8).Enabled Then itmFileExportMap_Click
          Case Config("shortcuts")("fileexportpicture"): mnuFile_Click: If itmFile(9).Enabled Then itmFileExportPicture_Click
          
          Case Config("shortcuts")("editundo"): mnuEdit_Click: If itmEditUndo.Enabled Then itmEditUndo_Click
          Case Config("shortcuts")("editredo"): mnuEdit_Click: If itmEditRedo.Enabled Then itmEditRedo_Click
          Case Config("shortcuts")("editmove"): mnuEdit_Click: If itmEditMode(0).Enabled Then itmEditMode_Click 0
          Case Config("shortcuts")("editvertices"): mnuEdit_Click: If itmEditMode(1).Enabled Then itmEditMode_Click 1
          Case Config("shortcuts")("editlines"): mnuEdit_Click: If itmEditMode(2).Enabled Then itmEditMode_Click 2
          Case Config("shortcuts")("editsectors"): mnuEdit_Click: If itmEditMode(3).Enabled Then itmEditMode_Click 3
          Case Config("shortcuts")("editthings"): mnuEdit_Click: If itmEditMode(4).Enabled Then itmEditMode_Click 4
          Case Config("shortcuts")("edit3d"): mnuEdit_Click: If itmEditMode(5).Enabled Then itmEditMode_Click 5
          Case Config("shortcuts")("editcut"): mnuEdit_Click: If itmEditCut.Enabled Then itmEditCut_Click
          Case Config("shortcuts")("editcopy"): mnuEdit_Click: If itmEditCopy.Enabled Then itmEditCopy_Click
          Case Config("shortcuts")("editpaste"): mnuEdit_Click: If itmEditPaste.Enabled Then itmEditPaste_Click
          Case Config("shortcuts")("editdelete"): mnuEdit_Click: If itmEditDelete.Enabled Then itmEditDelete_Click
          Case Config("shortcuts")("editfind"): mnuEdit_Click: If itmEditFind.Enabled Then itmEditFind_Click
          Case Config("shortcuts")("editreplace"): mnuEdit_Click: If itmEditReplace.Enabled Then itmEditReplace_Click
          Case Config("shortcuts")("editoptions"): mnuEdit_Click: If itmEditMapOptions.Enabled Then itmEditMapOptions_Click
          Case Config("shortcuts")("editfliph"): mnuEdit_Click: If itmEditFlipH.Enabled Then itmEditFlipH_Click
          Case Config("shortcuts")("editflipv"): mnuEdit_Click: If itmEditFlipV.Enabled Then itmEditFlipV_Click
          Case Config("shortcuts")("editrotate"): mnuEdit_Click: If itmEditRotate.Enabled Then itmEditRotate_Click
          Case Config("shortcuts")("editresize"): mnuEdit_Click: If itmEditResize.Enabled Then itmEditResize_Click
          Case Config("shortcuts")("editcenterview"): mnuEdit_Click: If itmEditCenterView.Enabled Then itmEditCenterView_Click
          Case Config("shortcuts")("togglesnap"): mnuEdit_Click: If itmEditSnapToGrid.Enabled Then itmEditSnapToGrid_Click
          Case Config("shortcuts")("togglestitch"): mnuEdit_Click: If itmEditStitch.Enabled Then itmEditStitch_Click
          
          Case Config("shortcuts")("clearvertices"): If mnuVertices.Visible Then mnuVertices_Click: If itmVerticesClearUnused.Enabled Then itmVerticesClearUnused_Click
          Case Config("shortcuts")("stitchvertices"): If mnuVertices.Visible Then mnuVertices_Click: If itmVerticesStitch.Enabled Then itmVerticesStitch_Click
          
          Case Config("shortcuts")("linesautoalign"): If mnuLines.Visible Then mnuLines_Click: If itmLinesAlign.Enabled Then itmLinesAlign_Click
          Case Config("shortcuts")("select1sided"): If mnuLines.Visible Then mnuLines_Click: If itmLinesSelect(0).Enabled Then itmLinesSelect_Click 0
          Case Config("shortcuts")("select2sided"): If mnuLines.Visible Then mnuLines_Click: If itmLinesSelect(1).Enabled Then itmLinesSelect_Click 1
          Case Config("shortcuts")("fliplinedefs"): If mnuLines.Visible Then mnuLines_Click: If itmLinesFlipLinedefs.Enabled Then itmLinesFlipLinedefs_Click
          Case Config("shortcuts")("flipsidedefs"): If mnuLines.Visible Then mnuLines_Click: If itmLinesFlipSidedefs.Enabled Then itmLinesFlipSidedefs_Click
          Case Config("shortcuts")("curvelines"): If mnuLines.Visible Then mnuLines_Click: If itmLinesCurve.Enabled Then itmLinesCurve_Click
          
          Case Config("shortcuts")("joinsector"): If mnuSectors.Visible Then mnuSectors_Click: If itmSectorsJoin.Enabled Then itmSectorsJoin_Click
          Case Config("shortcuts")("mergesector"): If mnuSectors.Visible Then mnuSectors_Click: If itmSectorsMerge.Enabled Then itmSectorsMerge_Click
          Case Config("shortcuts")("raisefloor"): If mnuSectors.Visible Then mnuSectors_Click: If itmSectorsRaiseFloor.Enabled Then itmSectorsRaiseFloor_Click
          Case Config("shortcuts")("lowerfloor"): If mnuSectors.Visible Then mnuSectors_Click: If itmSectorsLowerFloor.Enabled Then itmSectorsLowerFloor_Click
          Case Config("shortcuts")("raiseceil"): If mnuSectors.Visible Then mnuSectors_Click: If itmSectorsRaiseCeiling.Enabled Then itmSectorsRaiseCeiling_Click
          Case Config("shortcuts")("lowerceil"): If mnuSectors.Visible Then mnuSectors_Click: If itmSectorsLowerCeiling.Enabled Then itmSectorsLowerCeiling_Click
          Case Config("shortcuts")("brightinc"): If mnuSectors.Visible Then mnuSectors_Click: If itmSectorsIncBrightness.Enabled Then itmSectorsIncBrightness_Click
          Case Config("shortcuts")("brightdec"): If mnuSectors.Visible Then mnuSectors_Click: If itmSectorsDecBrightness.Enabled Then itmSectorsDecBrightness_Click
          Case Config("shortcuts")("gradientbrightness"): If mnuSectors.Visible Then mnuSectors_Click: If itmSectorsGradientBrightness.Enabled Then itmSectorsGradientBrightness_Click
          Case Config("shortcuts")("gradientfloors"): If mnuSectors.Visible Then mnuSectors_Click: If itmSectorsGradientFloors.Enabled Then itmSectorsGradientFloors_Click
          Case Config("shortcuts")("gradientceilings"): If mnuSectors.Visible Then mnuSectors_Click: If itmSectorsGradientCeilings.Enabled Then itmSectorsGradientCeilings_Click
          
          Case Config("shortcuts")("thingsfilter"): If mnuThings.Visible Then mnuThings_Click: If itmThingsFilter.Enabled Then itmThingsFilter_Click
          
               
          Case Config("shortcuts")("errorcheck"): If mnuTools.Visible Then mnuTools_Click: If itmToolsFindErrors.Enabled Then itmToolsFindErrors_Click
          Case Config("shortcuts")("removetextures"): If mnuTools.Visible Then mnuTools_Click: If itmToolsClearTextures.Enabled Then itmToolsClearTextures_Click
          Case Config("shortcuts")("fixtextures"): If mnuTools.Visible Then mnuTools_Click: If itmToolsFixTextures.Enabled Then itmToolsFixTextures_Click
          Case Config("shortcuts")("fileconfig"): mnuTools_Click: If itmToolsConfiguration.Enabled Then itmToolsConfiguration_Click
          Case Config("shortcuts")("fixzerolengthlines"): mnuTools_Click: If itmToolsFixZeroLinedefs.Enabled Then itmToolsFixZeroLinedefs_Click
          
          Case Config("shortcuts")("prefabinsert"): If mnuPrefabs.Visible Then mnuPrefabs_Click: If itmPrefabInsert.Enabled Then itmPrefabInsert_Click
          Case Config("shortcuts")("prefabinsertlast"): If mnuPrefabs.Visible Then mnuPrefabs_Click: If itmPrefabPrevious.Enabled Then itmPrefabPrevious_Click
          Case Config("shortcuts")("prefabinsert1"): If mnuPrefabs.Visible Then mnuPrefabs_Click: If itmPrefabQuick(0).Enabled Then itmPrefabQuick_Click 0
          Case Config("shortcuts")("prefabinsert2"): If mnuPrefabs.Visible Then mnuPrefabs_Click: If itmPrefabQuick(1).Enabled Then itmPrefabQuick_Click 1
          Case Config("shortcuts")("prefabinsert3"): If mnuPrefabs.Visible Then mnuPrefabs_Click: If itmPrefabQuick(2).Enabled Then itmPrefabQuick_Click 2
          Case Config("shortcuts")("prefabinsert4"): If mnuPrefabs.Visible Then mnuPrefabs_Click: If itmPrefabQuick(3).Enabled Then itmPrefabQuick_Click 3
          Case Config("shortcuts")("prefabinsert5"): If mnuPrefabs.Visible Then mnuPrefabs_Click: If itmPrefabQuick(4).Enabled Then itmPrefabQuick_Click 4
          
          Case Config("shortcuts")("helpwebsite"): itmHelpWebsite_Click
          Case Config("shortcuts")("helpfaq"): itmHelpFAQ_Click
          Case Config("shortcuts")("helpabout"): itmHelpAbout_Click
          
          
          Case Config("shortcuts")("snaptogrid")
               If mnuVertices.Visible Then
                    mnuVertices_Click
                    If itmVerticesSnapToGrid.Enabled Then itmVerticesSnapToGrid_Click
               ElseIf mnuLines.Visible Then
                    mnuLines_Click
                    If itmLinesSnapToGrid.Enabled Then itmLinesSnapToGrid_Click
               ElseIf mnuSectors.Visible Then
                    mnuSectors_Click
                    If itmSectorsSnapToGrid.Enabled Then itmSectorsSnapToGrid_Click
               ElseIf mnuThings.Visible Then
                    mnuThings_Click
                    If itmThingsSnapToGrid.Enabled Then itmThingsSnapToGrid_Click
               End If
               
               
          Case Config("shortcuts")("switchmode")
               
               'Check what mode to switch to
               Select Case mode
                    
                    Case EM_MOVE
                         mnuEdit_Click
                         If itmEditMode(EM_VERTICES).Enabled Then itmEditMode_Click EM_VERTICES
                    
                    Case EM_VERTICES
                         mnuEdit_Click
                         If itmEditMode(EM_LINES).Enabled Then itmEditMode_Click EM_LINES
                    
                    Case EM_LINES
                         mnuEdit_Click
                         If itmEditMode(EM_SECTORS).Enabled Then itmEditMode_Click EM_SECTORS
                    
                    Case EM_SECTORS
                         mnuEdit_Click
                         If itmEditMode(EM_THINGS).Enabled Then itmEditMode_Click EM_THINGS
                    
                    Case EM_THINGS
                         mnuEdit_Click
                         If itmEditMode(EM_VERTICES).Enabled Then itmEditMode_Click EM_VERTICES
               End Select
          
          
          Case Config("shortcuts")("copyprops")
               If mnuLines.Visible Then
                    mnuLines_Click
                    If itmLinesCopy.Enabled Then itmLinesCopy_Click
               ElseIf mnuSectors.Visible Then
                    mnuSectors_Click
                    If itmSectorsCopy.Enabled Then itmSectorsCopy_Click
               ElseIf mnuThings.Visible Then
                    mnuThings_Click
                    If itmThingsCopy.Enabled Then itmThingsCopy_Click
               End If
          
          
          Case Config("shortcuts")("pasteprops")
               If mnuLines.Visible Then
                    mnuLines_Click
                    If itmLinesPaste.Enabled Then itmLinesPaste_Click
               ElseIf mnuSectors.Visible Then
                    mnuSectors_Click
                    If itmSectorsPaste.Enabled Then itmSectorsPaste_Click
               ElseIf mnuThings.Visible Then
                    mnuThings_Click
                    If itmThingsPaste.Enabled Then itmThingsPaste_Click
               End If
               
     End Select
     
End Sub

Private Sub KeypressSectors(ByVal ShortcutCode As Long)
     
     'Check what key is pressed
     Select Case ShortcutCode
          
          Case Config("shortcuts")("drawsector")
               
               'Only in normal mode
               If (submode = ESM_NONE) Then
                    
                    'Start drawing
                    StartDrawOperation
               End If
          
          Case Config("shortcuts")("createsector")
               
               'Only in normal mode
               If (submode = ESM_NONE) Then
                    
                    'Insert vertex
                    If MouseInside Then CreateSectorHere LastX, LastY
               End If
               
               
     End Select
End Sub

Private Sub KeypressThings(ByVal ShortcutCode As Long)
     
     'Check what key is pressed
     Select Case ShortcutCode
          
          Case Config("shortcuts")("insertthing")
               
               'Only in normal mode
               If (submode = ESM_NONE) Then
                    
                    'Insert vertex
                    If MouseInside Then InsertThingHere LastX, LastY
               End If
               
     End Select
End Sub

Private Sub KeypressVertexes(ByVal ShortcutCode As Long)
     
     'Check what key is pressed
     Select Case ShortcutCode
          
          Case Config("shortcuts")("insertvertex")
               
               'Only in normal mode
               If (submode = ESM_NONE) Then
                    
                    'Insert vertex
                    If MouseInside Then
                         InsertVertexHere LastX, LastY
                         picMap.Refresh
                    End If
               End If
               
     End Select
End Sub

Private Sub itmVerticesStitch_Click()
     Dim Indices As Variant
     Dim i As Long
     Dim firstv As Long
     
     'Switch back if in move mode
     If (mode = EM_MOVE) Then itmEditMode_Click CInt(PreviousMode)
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Check if a selection is made
     If (numselected > 1) Then
          
          'Make undo
          CreateUndo "stitch vertices"
          
          'Stitch em all
          StitchSelectedVertices
          
          'Remove highlight
          RemoveHighlight True
          
          'Redraw map
          RedrawMap False
          
          'Map changed
          mapchanged = True
          mapnodeschanged = True
     End If
End Sub

Public Sub mnuEdit_Click()
     
     'Enable/Disable paste
     itmEditPaste.Enabled = (PasteAvailable And (mapfile <> ""))
     
     'Enable/Disable copy, cut and delete
     itmEditCopy.Enabled = (((numselected > 0) Or (currentselected > -1)) And (mapfile <> ""))
     itmEditCut.Enabled = (((numselected > 0) Or (currentselected > -1)) And (mapfile <> ""))
     itmEditDelete.Enabled = (((numselected > 0) Or (currentselected > -1)) And (mapfile <> ""))
End Sub

Public Sub mnuFile_Click()
     
     'Remove highlight
     If (mapfile <> "") Then RemoveHighlight True
End Sub

Private Sub mnuHelp_Click()
     'Food
End Sub

Private Sub mnuLines_Click()
     'Restaurant
End Sub

Private Sub mnuPrefabs_Click()
     Dim i As Long
     
     'Check if a previous prefab can be inserted
     If (Trim$(LastPrefab) <> "") Then
          
          'Check if the file can be found
          If (Dir(LastPrefab) <> "") Then
               
               'Enable last prefab
               itmPrefabPrevious.Enabled = True
          Else
               
               'Cant find the file
               itmPrefabPrevious.Enabled = False
          End If
     Else
          
          'No last prefab
          itmPrefabPrevious.Enabled = False
     End If
     
     'Check if a selection is made
     If (numselected > 0) Then
          
          'Enable save selection
          itmPrefabSaveSel.Enabled = True
     Else
          
          'Nothing selected
          itmPrefabSaveSel.Enabled = False
     End If
     
     'Go for all items
     For i = 0 To 4
          
          'Check if a prefab is configured
          If (Trim$(Config("quickprefab" & i + 1)) <> "") And _
             (Dir(Trim$(Config("quickprefab" & i + 1))) <> "") Then
               
               'Enable this item
               itmPrefabQuick(i).Enabled = True
               itmPrefabQuick(i).Caption = MenuNameForShortcut("&" & i + 1 & " - " & Dir(Trim$(Config("quickprefab" & i + 1))), "prefabinsert" & i + 1)
          Else
               
               'Disable item
               itmPrefabQuick(i).Enabled = False
               itmPrefabQuick(i).Caption = MenuNameForShortcut("&" & i + 1 & " - Unused", "prefabinsert" & i + 1)
          End If
     Next i
End Sub

Private Sub mnuScripts_Click()
     'Blah
End Sub

Private Sub mnuSectors_Click()
     'pr0n
End Sub

Private Sub mnuThings_Click()
     'Cock
End Sub

Public Sub mnuTools_Click()
     'Pussy
End Sub

Private Sub mnuVertices_Click()
     'Cunt
End Sub

Public Sub MoveSelectOperation(ByVal X As Single, ByVal Y As Single)
     
     'Undraw selection rect
     Render_RectSwitched GrabX, -GrabY, LastSelX, -LastSelY, PAL_NORMAL, 2
     
     'Draw selection rect
     Render_RectSwitched GrabX, -GrabY, X, -Y, PAL_MULTISELECTION, 2
     
     'Show changes
     picMap.Refresh
     
     'Keep X and Y
     LastSelX = X
     LastSelY = Y
End Sub

Private Sub NewSectorSetup(ByVal CleamMiddleTextures As Boolean)
     Dim Indices As Variant
     Dim i As Long
     Dim ld As Long
     
     'Disable map editing
     picMap.Enabled = False
     
     'Select sector from lines
     SelectSectorsFromLinedefs
     
     'Check if we should show dialog
     If (Config("newsectordialog") = vbChecked) Then
          
          'Check if any sector(s) selected
          If (numselected > 0) Then
               
               'Load sector dialog
               Load frmSector
               
               'Dont make undo for this edit
               frmSector.lblMakeUndo.Caption = "No"
               
               'Reset mousepointer
               Screen.MousePointer = vbNormal
               
               'Show dialog
               frmSector.Show 1, Me
               
               'Change mousepointer
               Screen.MousePointer = vbHourglass
          End If
     End If
     
     'Select vertices from lines
     SelectVerticesFromLinedefs
     
     'Check if we should auto-stitch vertices
     If (stitchmode) Then
          
          'Make dragged selection same as current selection
          Set dragselected = selected
          dragnumselected = numselected
          
          'Find and keep lines that have changed (added)
          If (mode <> EM_THINGS) Then FindChangingLines True, True
          
          'Due to auto-stitch, linedefs could be overlapping
          'Combine these into one now
          MergeDoubleLinedefs
     End If
     
     'Select lines from vertices
     SelectLinedefsFromVertices
     
     'Go for all selected linedefs
     Indices = selected.Items
     For i = LBound(Indices) To UBound(Indices)
          
          'Get linedef
          ld = Indices(i)
          
          'Remove any unneeded textures
          'and add missing textures
          If (linedefs(ld).s1 > -1) Then
               RemoveUnusedSidedefTextures linedefs(ld).s1, CleamMiddleTextures
               FixMissingSidedefTextures linedefs(ld).s1
          End If
          If (linedefs(ld).s2 > -1) Then
               RemoveUnusedSidedefTextures linedefs(ld).s2, CleamMiddleTextures
               FixMissingSidedefTextures linedefs(ld).s2
          End If
     Next i
     
     'Check if we should show dialog
     If (Config("newlinesdialog") = vbChecked) Then
          
          'Check if any line(s) selected
          If (numselected > 0) Then
               
               'Load lines dialog
               Load frmLinedef
               
               'Dont make undo for this edit
               frmLinedef.lblMakeUndo.Caption = "No"
               
               'Reset mousepointer
               Screen.MousePointer = vbNormal
               
               'Show dialog
               frmLinedef.Show 1, Me
               
               'Change mousepointer
               Screen.MousePointer = vbHourglass
          End If
     End If
     
     'Enable map editing
     picMap.Enabled = True
End Sub

Private Sub picMap_DblClick()
     Dim StillDeselectAfterEdit As Boolean
     
     'Check if in 3D Mode
     If (mode = EM_3D) Then
          
          'Redo last mousebutton
          picMap_MouseDown LastMouseButton, LastMouseShift, LastMouseX, LastMouseY
          
     'When not in 3D Mode
     Else
          'Check for selection
          If (selected.Count <= 1) Then StillDeselectAfterEdit = True
          
          'Doubleclick does the same as normal right-click, EDIT
          picMap_MouseDown vbRightButton, 0, LastX, LastY
          
          'Remove selection after editing
          DeselectAfterEdit = StillDeselectAfterEdit
          
          'Do the rightclick
          picMap_MouseUp vbRightButton, 0, LastX, LastY
     End If
End Sub

Private Sub picMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Dim ShortcutCode As Long
     
     'Keep last position
     tmrMouseOutside.Enabled = True
     MouseInside = True
     
     'Disable autoscroll
     ChangeAutoscroll True
     
     'Leave immediately if no map is loaded
     If (mapfile = "") Then Exit Sub
     
     'Also leave when a timeout for this event is still active
     'If (tmrMouseTimeout.Enabled = True) Then Exit Sub
     
     'Check if in 3D Mode
     If (mode = EM_3D) Then
          
          'Keep last button and coords
          LastMouseButton = Button
          LastMouseShift = Shift
          LastMouseX = (CSng(X) / CSng(ScreenWidth)) * CSng(VideoParams.BackBufferWidth)
          LastMouseY = (CSng(Y) / CSng(ScreenHeight)) * CSng(VideoParams.BackBufferHeight)
          
          'Make the shortcut code from keycode and shift
          Select Case Button
               Case vbLeftButton: ShortcutCode = MOUSE_BUTTON_0 Or (Shift * (2 ^ 16))
               Case vbMiddleButton: ShortcutCode = MOUSE_BUTTON_2 Or (Shift * (2 ^ 16))
               Case vbRightButton: ShortcutCode = MOUSE_BUTTON_1 Or (Shift * (2 ^ 16))
          End Select
          
          'Check how we should process data
          If TextureSelecting Then
               
               'Perform the action associated with the key
               KeydownTextureSelect ShortcutCode
          Else
               
               'Perform the action associated with the key
               Keydown3D ShortcutCode
          End If
          
     'When not in 3D Mode
     Else
          
          'Check in what operation we are
          Select Case submode
               
               'No operation
               Case ESM_NONE
                    
                    'Dont start selecting when moving
                    StartSelection = False
                    
                    'Check if an object is highlighted
                    If (currentselected > -1) Then
                         
                         'Check if left button is used (SELECT)
                         If (Button = vbLeftButton) Then
                              
                              'Check if we should remove current selection first
                              If (((Shift And vbCtrlMask) = 0) And ((Shift And vbShiftMask) = 0) And (Config("additiveselect") = vbUnchecked)) Then RemoveSelection True
                              
                              'Select/Deselect highlighted object for the mode we are in
                              Select Case mode
                                   Case EM_VERTICES: SelectCurrentVertex
                                   Case EM_LINES: SelectCurrentLine
                                   Case EM_SECTORS: SelectCurrentSector
                                   Case EM_THINGS: SelectCurrentThing
                              End Select
                              
                              'Ready for editing
                              NoEditing = False
                              
                         'Check if right button is used (EDIT)
                         ElseIf (Button = vbRightButton) Then
                              
                              'Check if clicking outside selection
                              If (selected.Exists(CStr(currentselected)) = False) Then
                                   
                                   'Remove old selection
                                   RemoveSelection True
                                   
                                   'Select/Deselect highlighted object for the mode we are in
                                   Select Case mode
                                        Case EM_VERTICES: SelectCurrentVertex
                                        Case EM_LINES: SelectCurrentLine
                                        Case EM_SECTORS: SelectCurrentSector
                                        Case EM_THINGS: SelectCurrentThing
                                   End Select
                                   
                                   'We'll remove this selection after editing
                                   DeselectAfterEdit = True
                              Else
                                   
                                   'The selection was made manually, keep it
                                   DeselectAfterEdit = False
                              End If
                              
                              'Remove highlight
                              RemoveHighlight
                              
                              'Keep the grabbing postion
                              'When the mouse moves the drag mode will start
                              GrabX = X
                              GrabY = Y
                              
                              'If not dragging, the editing dialog will be shown on mouse release
                              NoEditing = False
                         End If
                    Else
                         
                         'Keep the grabbing postion
                         'When the mouse moves the drag mode will start
                         GrabX = X
                         GrabY = Y
                         
                         'Check if left button is used
                         If (Button = vbLeftButton) Then
                              
                              'Remove highlight
                              RemoveHighlight
                              
                              'Check mode
                              If (mode = EM_MOVE) Then
                                   
                                   'Change mousecursor
                                   Set picMap.MouseIcon = imgCursor(2).Picture
                                   submode = ESM_MOVING
                              Else
                                   
                                   'Start selecting when moving
                                   StartSelection = True
                              End If
                         ElseIf (Button = vbRightButton) Then
                              
                              'Check if no selection made
                              If (numselected = 0) Then
                                   
                                   'Check what to insert
                                   Select Case mode
                                        Case EM_VERTICES
                                             
                                             'Insert vertex
                                             InsertVertexHere X, Y
                                             
                                             'Highlight the vertex
                                             ShowHighlight X, Y
                                             
                                             'Select the vertex
                                             SelectCurrentVertex
                                             
                                             'We'll remove this selection after editing
                                             DeselectAfterEdit = True
                                             
                                             'Do not edit the vertex on mouse release
                                             NoEditing = True
                                             
                                        Case EM_THINGS
                                             
                                             'Insert thing
                                             InsertThingHere X, Y
                                             
                                             'Highlight the thing
                                             ShowHighlight X, Y
                                             
                                             'Check if we should select the thing for dragging
                                             If (Config("newthingdialog") = vbUnchecked) Then
                                                  
                                                  'Select the thing
                                                  SelectCurrentThing
                                                  
                                                  'We'll remove this selection after editing
                                                  DeselectAfterEdit = True
                                             End If
                                             
                                             'Do not edit the thing on mouse release
                                             NoEditing = True
                                   End Select
                              End If
                         End If
                    End If
                    
               'Paste operation
               Case ESM_PASTING
                    
                    'End of dragging operation
                    EndDragOperation
                    
                    'Remove selection
                    RemoveSelection True
                    RemoveHighlight True
                    
                    'Do a mousemove to update the highlight
                    picMap_MouseMove vbNormal, Shift, X, Y
                    
               'Drawing operation
               Case ESM_DRAWING
                    
                    'Remove previous drawn line
                    RenderDrawingLine PAL_NORMAL, PAL_NORMAL, LastX, LastY
                    
                    'Only draw with left mousebutton
                    If (Button = vbLeftButton) Then
                         
                         'Draw here
                         DrawVertexHere X, Y
                    End If
                    
                    'Show changes
                    picMap.Refresh
                    
          End Select
          
          'Keep last position
          LastX = X
          LastY = Y
          
          'Set a timeout for the next mousemove event
          tmrMouseTimeout.Enabled = True
     End If
End Sub

Private Sub picMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     
     'Keep last position
     tmrMouseOutside.Enabled = True
     MouseInside = True
     
     'Leave immediately if no map is loaded
     If (mapfile = "") Then Exit Sub
     
     'Check if in 3D Mode
     If (mode = EM_3D) Then
          
          'Keep last coords
          LastMouseX = (CSng(X) / CSng(ScreenWidth)) * CSng(VideoParams.BackBufferWidth)
          LastMouseY = (CSng(Y) / CSng(ScreenHeight)) * CSng(VideoParams.BackBufferHeight)
          
     'When not in 3D Mode
     Else
          
          'Also leave when a timeout for this event is still active
          If (tmrMouseTimeout.Enabled = True) Then Exit Sub
          
          'Update cursor position in statusbar
          stbStatus.Panels("mousex").Text = "X " & CLng(X)
          stbStatus.Panels("mousey").Text = "Y " & -CLng(Y)
          
          'Check in what operation we are
          Select Case submode
               
               'No operation
               Case ESM_NONE
                    
                    'Check if no mousebutton is hold
                    If (Button = vbNormal) Then
                         
                         'Show highlight
                         ShowHighlight X, Y
                         
                    'Check if right mousebutton is hold and a selection is made
                    ElseIf ((numselected > 0) And (Button = vbRightButton)) Then
                         
                         'Check if moved enough pixels to start drag
                         If ((Abs(X * ViewZoom - GrabX * ViewZoom) >= Config("dragpixels")) Or _
                             (Abs(Y * ViewZoom - GrabY * ViewZoom) >= Config("dragpixels"))) Then
                              
                              'Start drag operation
                              StartDragOperation X, Y
                         End If
                         
                    'Check if left mouse button is hold and no highlight
                    ElseIf ((Button = vbLeftButton) And (currentselected = -1) And StartSelection) Then
                         
                         'Remove highlight
                         RemoveHighlight True
                         
                         'Start multiselect mode
                         StartSelectOperation X, Y
                    End If
               
               'Drag operation
               Case ESM_DRAGGING
                    
                    'Dragging from right mousebutton
                    If (Button = vbRightButton) Then
                         
                         'Draw selection
                         DragSelection Shift, X, Y
                    End If
                    
                    'Check if we should autoscroll
                    If Config("autoscroll") Then tmrAutoScroll.Enabled = True
                    
               'Paste operation
               Case ESM_PASTING
                    
                    'Draw selection
                    DragSelection Shift, X, Y
                    
                    'Check if we should autoscroll
                    If Config("autoscroll") Then tmrAutoScroll.Enabled = True
                    
               'Select operation
               Case ESM_SELECTING
                    
                    'Selection from left mousebutton
                    If (Button = vbLeftButton) Then
                         
                         'Draw selection
                         MoveSelectOperation X, Y
                    End If
                    
                    'Check if we should autoscroll
                    If Config("autoscroll") Then tmrAutoScroll.Enabled = True
                    
               'Drawing operation
               Case ESM_DRAWING
                    
                    'Check if no button is hold
                    If (Button = vbNormal) Then
                         
                         'Remove previous drawn line
                         RenderDrawingLine PAL_NORMAL, PAL_NORMAL, LastX, LastY
                         
                         'Render line being drawn
                         RenderDrawingLine PAL_MULTISELECTION, PAL_BACKGROUND, X, Y
                         
                         'Show changes
                         picMap.Refresh
                    End If
                    
                    'Check if we should autoscroll
                    If Config("autoscroll") Then tmrAutoScroll.Enabled = True
                    
               'Moving map
               Case ESM_MOVING
                    
                    'Check if not the right button hold
                    If (Button <> vbRightButton) Then
                         
                         'Drag the map
                         ViewLeft = ViewLeft - (X - LastX)
                         ViewTop = ViewTop - (Y - LastY)
                         X = X - (X - LastX)
                         Y = Y - (Y - LastY)
                         
                         'Change viewport
                         ChangeView ViewLeft, ViewTop, ViewZoom
                         
                         'Redraw map
                         RedrawMap
                    End If
                    
          End Select
          
          'Keep last position
          LastX = X
          LastY = Y
          
          'Set autoscroll
          ChangeAutoscroll False
          
          'Set a timeout for the next mousemove event
          tmrMouseTimeout.Enabled = True
     End If
End Sub

Private Sub picMap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Dim ShortcutCode As Long
     
     'Keep last position
     tmrMouseOutside.Enabled = True
     MouseInside = True
     
     'Leave immediately if no map is loaded
     If (mapfile = "") Then Exit Sub
     
     'Also leave when a timeout for this event is still active
     'If (tmrMouseTimeout.Enabled = True) Then Exit Sub
     
     'Check if in 3D Mode
     If (mode = EM_3D) Then
          
          'Keep last coords
          LastMouseX = (CSng(X) / CSng(ScreenWidth)) * CSng(VideoParams.BackBufferWidth)
          LastMouseY = (CSng(Y) / CSng(ScreenHeight)) * CSng(VideoParams.BackBufferHeight)
          
          'Make the shortcut code from keycode and shift
          Select Case Button
               Case vbLeftButton: ShortcutCode = MOUSE_BUTTON_0 Or (Shift * (2 ^ 16))
               Case vbMiddleButton: ShortcutCode = MOUSE_BUTTON_2 Or (Shift * (2 ^ 16))
               Case vbRightButton: ShortcutCode = MOUSE_BUTTON_1 Or (Shift * (2 ^ 16))
          End Select
          
          'Check how we should process data
          If Not TextureSelecting Then
               
               'Perform the action associated with the key
               Keyrelease3D ShortcutCode
          End If
          
     'When not in 3D Mode
     Else
          
          'Check in what operation we are
          Select Case submode
               
               'No operation
               Case ESM_NONE
                    
                    'Show highlight
                    ShowHighlight X, Y
                    
                    'Check if no object was highlighted
                    If (currentselected = -1) Then
                         
                         'Deselect if preferred
                         If (Config("nothingdeselects") = vbChecked) Then RemoveSelection True
                         
                         'Check if right mousebutton used
                         If (Button = vbRightButton) Then
                              
                              'Check if no selection made
                              If (numselected = 0) Then
                                   
                                   'Check what to insert
                                   Select Case mode
                                        Case EM_LINES
                                             If ((Shift And vbShiftMask) = 0) And _
                                                ((Shift And vbCtrlMask) = 0) Then
                                                  
                                                  'Start drawing
                                                  StartDrawOperation
                                                  DrawVertexHere X, Y
                                             Else
                                                  
                                                  'Insert sector
                                                  CreateSectorHere X, Y
                                             End If
                                             
                                        Case EM_SECTORS
                                             If ((Shift And vbShiftMask) = 0) And _
                                                ((Shift And vbCtrlMask) = 0) Then
                                                  
                                                  'Start drawing
                                                  StartDrawOperation
                                                  DrawVertexHere X, Y
                                             Else
                                                  
                                                  'Insert sector
                                                  CreateSectorHere X, Y
                                             End If
                                   End Select
                              End If
                         End If
                    
                    'Check if a selection is made and we will edit it
                    ElseIf ((numselected > 0) And (Button = vbRightButton)) Then
                         
                         'Check if right button is used and allowed to edit
                         If (Button = vbRightButton) Then
                              
                              'Check if allowed to edit
                              If (NoEditing = False) Then
                                   
                                   'Remove any highlight
                                   RemoveHighlight True
                                   
                                   'Check what is gonna be edited
                                   Select Case mode
                                        Case EM_VERTICES:   'Just drag it bitch!
                                        Case EM_LINES: frmLinedef.Show 1, Me
                                        Case EM_SECTORS: frmSector.Show 1, Me
                                        Case EM_THINGS: frmThing.Show 1, Me
                                   End Select
                                   
                                   'Check if we should deselect
                                   If (DeselectAfterEdit) Then
                                        
                                        'Select/Deselect highlighted object for the mode we are in
                                        Select Case mode
                                             Case EM_VERTICES: RemoveVertexSelection
                                             Case EM_LINES: RemoveLinesSelection
                                             Case EM_SECTORS: RemoveSectorsSelection
                                             Case EM_THINGS: RemoveThingsSelection
                                        End Select
                                   End If
                                   
                                   'Redraw the map
                                   RedrawMap True
                              Else
                                   
                                   'Remove selection and highlight
                                   RemoveSelection False
                                   RemoveHighlight True
                                   
                                   'Show highlight
                                   ShowHighlight X, Y
                              End If
                         End If
                    Else
                         
                         'Show highlight
                         ShowHighlight X, Y
                    End If
                    
               'Drag operation
               Case ESM_DRAGGING
                    
                    'Dragging from right mousebutton
                    If (Button = vbRightButton) Then
                         
                         'End of dragging operation
                         EndDragOperation
                         
                         'Do a mousemove to update the highlight
                         picMap_MouseMove vbNormal, Shift, X, Y
                    End If
                    
               'Select operation
               Case ESM_SELECTING
                    
                    'Selection from left mousebutton
                    If (Button = vbLeftButton) Then
                         
                         'End of selecting operation
                         EndSelectOperation Shift, X, Y
                         
                         'No more select
                         StartSelection = False
                         
                         'Do a mousemove to update the highlight
                         picMap_MouseMove vbNormal, Shift, X, Y
                    End If
               
               'Drawing operation
               Case ESM_DRAWING
                    
                    'Check if right mousebutton used (END DRAWING)
                    If (Button = vbRightButton) Then
                         
                         'End drawing now
                         EndDrawOperation False
                         
                    'Otherwise check if left mousebutton used
                    ElseIf (Button = vbLeftButton) Then
                         
                         'Check if allowed to change
                         If picMap.Enabled Then
                              
                              'Render line being drawn
                              RenderDrawingLine PAL_MULTISELECTION, PAL_BACKGROUND, X, Y
                              
                              'Show changes
                              picMap.Refresh
                         End If
                    End If
                    
               'Moving map
               Case ESM_MOVING
                    
                    'No longer moving
                    Set picMap.MouseIcon = imgCursor(1).Picture
                    submode = ESM_NONE
                    
          End Select
          
          'Keep last position
          LastX = X
          LastY = Y
          
          'Set a timeout for the next mousemove event
          tmrMouseTimeout.Enabled = True
     End If
End Sub

Private Sub Place3DModeStart(ByVal X As Long, ByVal Y As Long)
     Dim t_found As Boolean
     Dim t As Long
     
     'Check if the position thing is within bounds
     If (PositionThing >= 0) And (PositionThing < numthings) Then
          
          'Check if the position thing is correct
          If (things(PositionThing).thing = mapconfig("start3dmode")) Then t_found = True
     End If
     
     'If no thing could be found, find a new one
     If (t_found = False) Then
          
          'Go for all things to find another positioning thing
          For t = 0 To (numthings - 1)
               
               'Check if this is a 3D start position
               If (things(t).thing = mapconfig("start3dmode")) Then
                    
                    'Use this
                    PositionThing = t
                    
                    'Found one
                    t_found = True
                    Exit For
               End If
          Next t
     End If
     
     'Check if a position thing could be found
     If t_found Then
          
          'Move the thing
          things(PositionThing).X = X
          things(PositionThing).Y = -Y
     Else
          
          'Create 3D Start thing
          PositionThing = CreateThing
          
          'Set tis properties
          With things(PositionThing)
               .angle = 0
               .arg0 = 0
               .arg1 = 0
               .arg2 = 0
               .arg3 = 0
               .arg4 = 0
               .effect = 0
               .Flags = 0
               .selected = 0
               .tag = 0
               .thing = mapconfig("start3dmode")
               .X = X
               .Y = -Y
               .Z = 0
          End With
          
          'Update image
          UpdateThingImageColor PositionThing
          UpdateThingSize PositionThing
          UpdateThingCategory PositionThing
     End If
     
     'Apply position
     ApplyPositionFromThing PositionThing
End Sub

Public Sub RemoveHighlight(Optional ByVal RemovePanels As Boolean = True)
     Dim ld As Long, ldfound As Long
     
     'Check the viewing mode
     Select Case mode
          
          Case EM_VERTICES
               
               'Render the last current selected to normal
               If (currentselected > -1) Then Render_AllVertices vertexes(0), currentselected, currentselected, vertexsize
               
               'Check if we should remove the info
               If (RemovePanels) Then HideVertexInfo
               
               
          Case EM_LINES
               
               'Check if a previous linedef was selected
               If (currentselected > -1) Then
                    
                    'Render the last selected linedef to normal (also vertices, those have been overdrawn)
                    Render_AllLinedefs vertexes(0), linedefs(0), currentselected, currentselected, submode, indicatorsize
                    If (Config("mode1vertices")) Then
                         Render_AllVertices vertexes(0), linedefs(currentselected).v1, linedefs(currentselected).v1, vertexsize
                         Render_AllVertices vertexes(0), linedefs(currentselected).v2, linedefs(currentselected).v2, vertexsize
                    End If
                    
                    'Render the last tagged sectors if line had a tag
                    If (linedefs(currentselected).tag > 0) Then Render_TaggedSectors vertexes(0), linedefs(0), VarPtr(sidedefs(0)), VarPtr(sectors(0)), numsectors, numlinedefs, linedefs(currentselected).tag, 0, indicatorsize, Config("mode1vertices"), vertexsize
               End If
               
               'Check if we should remove the info
               If (RemovePanels) Then HideLinedefInfo
               
               
          Case EM_SECTORS
               
               'Check if a previous sector was selected
               If (currentselected > -1) Then
                    
                    'Go for all linedefs
                    For ld = 0 To (numlinedefs - 1)
                         
                         'Check if one of the sidedefs belong to this sector
                         ldfound = 0
                         If (linedefs(ld).s1 > -1) Then If (sidedefs(linedefs(ld).s1).sector = currentselected) Then ldfound = 1
                         If (linedefs(ld).s2 > -1) Then If (sidedefs(linedefs(ld).s2).sector = currentselected) Then ldfound = 1
                         
                         If (ldfound) Then
                              
                              'Render this linedef to normal (also vertices, those have been overdrawn)
                              Render_AllLinedefs vertexes(0), linedefs(0), ld, ld, submode, indicatorsize
                              If (Config("mode2vertices")) Then
                                   Render_AllVertices vertexes(0), linedefs(ld).v1, linedefs(ld).v1, vertexsize
                                   Render_AllVertices vertexes(0), linedefs(ld).v2, linedefs(ld).v2, vertexsize
                              End If
                         End If
                    Next ld
                    
                    'Render the last tagged linedefs if sector had a tag
                    If (sectors(currentselected).tag > 0) Then Render_TaggedLinedefs vertexes(0), linedefs(0), numlinedefs, sectors(currentselected).tag, 1, 0, indicatorsize, 0, vertexsize
               End If
               
               'Check if we should remove the info
               If (RemovePanels) Then HideSectorInfo
               
               
          Case EM_THINGS
               
               'Render the last current selected to normal
               If (currentselected > -1) Then
                    Render_AllThings things(0), currentselected, currentselected, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, Val(Config("allthingsrects")), ViewZoom, filterthings, filtersettings
                    If Config("thingrects") Then Render_BoxSwitched things(currentselected).X, things(currentselected).Y, GetThingWidth(things(currentselected).thing) * ViewZoom, PAL_NORMAL, (Config("thingrects") - 1), PAL_NORMAL
               End If
               
               'Check if we should remove the info
               If (RemovePanels) Then HideThingInfo
               
               
          Case EM_3D
               
               'Check if we should remove the info
               If (RemovePanels) Then
                    HideLinedefInfo
                    HideSectorInfo
               End If
               
     End Select
     
     'Show map changes
     picMap.Refresh
     
     'No more highlight
     picMap.ToolTipText = ""
     currentselected = -1
End Sub

Private Sub RenderDrawingLine(ByVal Pal1 As ENUM_PALETTES, ByVal Pal2 As ENUM_PALETTES, ByVal X As Long, ByVal Y As Long)
     Dim x1 As Long, y1 As Long
     Dim x2 As Long, y2 As Long
     Dim lx As Long, ly As Long
     Dim Length As Long
     
     'Snap X and Y if snap mode is on
     If snapmode Then
          X = SnappedToGridX(X)
          Y = SnappedToGridY(Y)
     End If
     
     'Check if a vertex is already drawn
     If (numselected > 0) Then
          
          'Get line coordinates
          x1 = vertexes(selected.Items(selected.Count - 1)).X
          y1 = vertexes(selected.Items(selected.Count - 1)).Y
          x2 = X
          y2 = -Y
          
          'Line distances
          lx = x2 - x1
          ly = y2 - y1
          
          'Length of line
          Length = Int(Sqr(lx * lx + ly * ly))
          
          'Draw line switched
          Render_LinedefLineSwitched x1, y1, x2, y2, Pal1, indicatorsize
          
          'Draw previous vertex normal
          Render_BoxSwitched x1, y1, vertexsize, PAL_NORMAL, 1, PAL_NORMAL
          
          'Draw line length number
          Render_NumberSwitched Length, x1 + lx * 0.5, y1 + ly * 0.5, NumbersBitmapData(0), frmMain.picNumbers.width, frmMain.picNumbers.height, frmMain.picNumbers.width / 10, frmMain.picNumbers.height, Pal1, Pal2
     End If
     
     'Draw target vertex switched
     Render_BoxSwitched X, -Y, vertexsize, Pal1, 1, Pal1
End Sub

Private Sub SetupNewLinedefs(ByVal sc1 As Long, ByVal sc2 As Long, ByVal CopyFromSidedef As Long)
     Dim ld As Long
     Dim i As Long
     Dim Indices As Variant
     
     'Go for all selected linedefs
     Indices = selected.Items
     For i = 0 To (numselected - 1)
          
          'Get linedef
          ld = Indices(i)
          
          'Check if we should add Sidedef 1
          If (sc1 > -1) Then
               
               'Create sidedef
               linedefs(ld).s1 = CreateSidedef
               
               'Set sidedef properties
               With sidedefs(linedefs(ld).s1)
                    
                    'Reference to linedef
                    .linedef = ld
                    
                    'Reference to sector
                    .sector = sc1
                    
                    'Standard offsets
                    .tx = 0
                    .ty = 0
                    
                    'Check if lower can be copied
                    If (CopyFromSidedef > -1) Then
                         
                         'Check if source has a lower
                         If (IsTextureName(sidedefs(CopyFromSidedef).Lower)) Then
                              
                              'Copy from source
                              .Lower = sidedefs(CopyFromSidedef).Lower
                         Else
                              
                              'Check if source has a middle
                              If (IsTextureName(sidedefs(CopyFromSidedef).Middle)) Then
                                   
                                   'Copy from source middle
                                   .Lower = sidedefs(CopyFromSidedef).Middle
                              Else
                                   
                                   'Make default
                                   .Lower = Config("defaulttexture")("lower")
                              End If
                         End If
                    Else
                         
                         'Make default
                         .Lower = Config("defaulttexture")("lower")
                    End If
                    
                    'Check if a middle texture is needed
                    If (sc2 = -1) Then
                         
                         'Check if middle can be copied
                         If (CopyFromSidedef > -1) Then
                              
                              'Check if source has a middle
                              If (IsTextureName(sidedefs(CopyFromSidedef).Middle)) Then
                                   
                                   'Copy from source
                                   .Middle = sidedefs(CopyFromSidedef).Middle
                              Else
                                   
                                   'Make default
                                   .Middle = Config("defaulttexture")("middle")
                              End If
                         Else
                              
                              'Make default
                              .Middle = Config("defaulttexture")("middle")
                         End If
                    Else
                         
                         'No middle texture
                         .Middle = "-"
                    End If
                    
                    'Check if upper can be copied
                    If (CopyFromSidedef > -1) Then
                         
                         'Check if source has a upper
                         If (IsTextureName(sidedefs(CopyFromSidedef).Upper)) Then
                              
                              'Copy from source
                              .Upper = sidedefs(CopyFromSidedef).Upper
                         Else
                              
                              'Check if source has a middle
                              If (IsTextureName(sidedefs(CopyFromSidedef).Middle)) Then
                                   
                                   'Copy from source middle
                                   .Upper = sidedefs(CopyFromSidedef).Middle
                              Else
                                   
                                   'Make default
                                   .Upper = Config("defaulttexture")("upper")
                              End If
                         End If
                    Else
                         
                         'Make default
                         .Upper = Config("defaulttexture")("upper")
                    End If
               End With
          End If
          
          'Check if we should add Sidedef 2
          If (sc2 > -1) Then
               
               'Create sidedef
               linedefs(ld).s2 = CreateSidedef
               
               'Set sidedef properties
               With sidedefs(linedefs(ld).s2)
                    
                    'Reference to linedef
                    .linedef = ld
                    
                    'Reference to sector
                    .sector = sc2
                    
                    'Standard offsets
                    .tx = 0
                    .ty = 0
                    
                    'Check if lower can be copied
                    If (CopyFromSidedef > -1) Then
                         
                         'Check if source has a lower
                         If (IsTextureName(sidedefs(CopyFromSidedef).Lower)) Then
                              
                              'Copy from source
                              .Lower = sidedefs(CopyFromSidedef).Lower
                         Else
                              
                              'Check if source has a middle
                              If (IsTextureName(sidedefs(CopyFromSidedef).Middle)) Then
                                   
                                   'Copy from source middle
                                   .Lower = sidedefs(CopyFromSidedef).Middle
                              Else
                                   
                                   'Make default
                                   .Lower = Config("defaulttexture")("lower")
                              End If
                         End If
                    Else
                         
                         'Make default
                         .Lower = Config("defaulttexture")("lower")
                    End If
                    
                    'Check if a middle texture is needed
                    If (sc1 = -1) Then
                         
                         'Check if middle can be copied
                         If (CopyFromSidedef > -1) Then
                              
                              'Check if source has a middle
                              If (IsTextureName(sidedefs(CopyFromSidedef).Middle)) Then
                                   
                                   'Copy from source
                                   .Middle = sidedefs(CopyFromSidedef).Middle
                              Else
                                   
                                   'Make default
                                   .Middle = Config("defaulttexture")("middle")
                              End If
                         Else
                              
                              'Make default
                              .Middle = Config("defaulttexture")("middle")
                         End If
                    Else
                         
                         'No middle texture
                         .Middle = "-"
                    End If
                    
                    'Check if upper can be copied
                    If (CopyFromSidedef > -1) Then
                         
                         'Check if source has a upper
                         If (IsTextureName(sidedefs(CopyFromSidedef).Upper)) Then
                              
                              'Copy from source
                              .Upper = sidedefs(CopyFromSidedef).Upper
                         Else
                              
                              'Check if source has a middle
                              If (IsTextureName(sidedefs(CopyFromSidedef).Middle)) Then
                                   
                                   'Copy from source middle
                                   .Upper = sidedefs(CopyFromSidedef).Middle
                              Else
                                   
                                   'Make default
                                   .Upper = Config("defaulttexture")("upper")
                              End If
                         End If
                    Else
                         
                         'Make default
                         .Upper = Config("defaulttexture")("upper")
                    End If
               End With
               
               'Line is double sided
               linedefs(ld).Flags = linedefs(ld).Flags Or LDF_TWOSIDED
               linedefs(ld).Flags = linedefs(ld).Flags And Not LDF_IMPASSIBLE
          Else
               
               'Line is single sided
               linedefs(ld).Flags = linedefs(ld).Flags And Not LDF_TWOSIDED
               linedefs(ld).Flags = linedefs(ld).Flags Or LDF_IMPASSIBLE
          End If
     Next i
End Sub

Public Sub ShowConfiguration(ByVal showtab As Long)
     Dim OldIWAD As String
     Dim OldMixResources As Long
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Ensure the splash dialog is gone
     Unload frmSplash: Set frmSplash = Nothing
     
     'Cancel if in drawing operation
     CancelCurrentOperation
     
     'Keep the old settings
     OldIWAD = GetCurrentIWADFile
     OldMixResources = Val(Config("mixresources"))
     
     'Load dialog
     Load frmOptions
     
     'Select first tab
     frmOptions.tbsOptions.Tabs(showtab).selected = True
     
     'Show dialog
     frmOptions.Show 1, Me
     
     'Create rendering palette
     CreateRendererPalette
     
     'Only re-initialize renderer when a map is loaded
     If (mapfile <> "") Then
          
          'Change mousepointer
          Screen.MousePointer = vbHourglass
          
          'Show status dialog
          frmStatus.Show 0, frmMain
          frmStatus.Refresh
          frmMain.SetFocus
          frmMain.Refresh
          
          'Load the error log
          ErrorLog_Load
          
          'Reload textures and flat when needed
          If (OldIWAD <> GetCurrentIWADFile) Or (OldMixResources <> Val(Config("mixresources"))) Then
               
               'Close previous IWAD
               IWAD.CloseFile
               
               'Open associated IWAD
               OpenIWADFile
               
               'Precache resources
               MapLoadResources
          End If
          
          'Initialize the map screen
          InitializeMapRenderer frmMain.picMap
          
          'Unload status dialog
          Unload frmStatus: Set frmSplash = Nothing
          
          'Reset mousepointer
          Screen.MousePointer = vbDefault
          
          'Show the errors and warnings dialog
          ErrorLog_DisplayAndFlush
          
          'Set the viewport
          ChangeView ViewLeft, ViewTop, ViewZoom
     End If
     
     'Apply configuration on Interface
     ApplyInterfaceConfiguration
     
     'Update the controls
     Form_Resize
     
     'Update shortcuts
     UpdateMenuShortcuts
End Sub

Public Sub ShowHighlight(ByVal X As Long, ByVal Y As Long)
     
     'Cant show the highlight when no mouse inside
     If (MouseInside = False) Then Exit Sub
     
     'Highlight object for the mode we are in
     Select Case mode
          Case EM_VERTICES: ChangeVertexHighlight X, Y
          Case EM_LINES: ChangeLinesHighlight X, Y
          Case EM_SECTORS: ChangeSectorsHighlight X, Y
          Case EM_THINGS: ChangeThingsHighlight X, Y
     End Select
End Sub

Private Sub StartDragOperation(ByVal X As Single, ByVal Y As Single)
     Dim distance As Long
     
     'Make undo
     Select Case mode
          Case EM_THINGS: CreateUndo "thing drag"
          Case EM_VERTICES: CreateUndo "vertex drag"
          Case EM_LINES: CreateUndo "linedef drag"
          Case EM_SECTORS: CreateUndo "sector drag"
     End Select
     
     'never edit during or after drag
     NoEditing = True
     
     'Drag operation only drags vertices and things
     'so select vertices when dragging lines or sectors
     If ((mode = EM_LINES) Or (mode = EM_SECTORS)) Then
          
          'Make dragging selection
          Set dragselected = SelectVerticesFromSelection
          dragnumselected = dragselected.Count
     Else
          
          'Dragging same as selection
          Set dragselected = selected
          dragnumselected = numselected
     End If
     
     'Grab the nearest object
     If (mode = EM_THINGS) Then
          
          'Grab the nearest thing
          grabobject = NearestSelectedThing(X, Y, things(0), numthings, distance)
          
          'Calculate offset between mouse and grab object
          GrabX = GrabX - things(grabobject).X
          GrabY = -GrabY - things(grabobject).Y
     Else
          
          'Grab the nearest vertex
          grabobject = NearestSelectedVertex(X, Y, vertexes(0), numvertexes, distance)
          
          'Calculate offset between mouse and grab object
          GrabX = GrabX - vertexes(grabobject).X
          GrabY = -GrabY - vertexes(grabobject).Y
     End If
     
     'Find and keep lines that will be changed
     If (mode <> EM_THINGS) Then FindChangingLines False, True
     
     'Start drag operation
     currentselected = -1
     submode = ESM_DRAGGING
     
     'Redraw map
     RedrawMap
End Sub

Private Sub StartDrawOperation(Optional ByVal RefreshMap As Boolean = True)
     
     'Make undo
     CreateUndo "draw"
     
     'Deselect all
     RemoveSelection False
     RemoveHighlight True
     
     'Start draw operation
     currentselected = -1
     submode = ESM_DRAWING
     
     'No changed lines yet
     ReDim changedlines(0)
     numchangedlines = 0
     ReDim DrawingCoords(0)
     NumDrawingCoords = 0
     
     'Map has changed
     mapchanged = True
     mapnodeschanged = True
     
     'Redraw map
     If (RefreshMap) Then RedrawMap
End Sub

Private Sub StartPasteOperation(ByVal X As Single, ByVal Y As Single)
     Dim distance As Long
     
     'Set hourglass mousepointer
     Screen.MousePointer = vbArrowHourglass
     
     'Drag operation only drags vertices and things
     'so select vertices when dragging lines or sectors
     If ((mode = EM_LINES) Or (mode = EM_SECTORS)) Then
          
          'Make dragging selection
          Set dragselected = SelectVerticesFromSelection
          dragnumselected = dragselected.Count
     Else
          
          'Dragging same as selection
          Set dragselected = selected
          dragnumselected = numselected
     End If
     
     'Grab the nearest object
     If (mode = EM_THINGS) Then
          
          'Grab the nearest thing
          grabobject = NearestSelectedThing(X, Y, things(0), numthings, distance)
          
          'Check if anything
          If (grabobject > -1) Then
               
               'Calculate offset between mouse and grab object
               GrabX = GrabX - things(grabobject).X
               GrabY = -GrabY - things(grabobject).Y
          End If
     Else
          
          'Grab the nearest vertex
          grabobject = NearestSelectedVertex(X, Y, vertexes(0), numvertexes, distance)
          
          'Check if anything
          If (grabobject > -1) Then
               
               'Calculate offset between mouse and grab object
               GrabX = GrabX - vertexes(grabobject).X
               GrabY = -GrabY - vertexes(grabobject).Y
          End If
     End If
     
     'Set normal mousepointer
     Screen.MousePointer = vbDefault
     
     'Start paste operation
     currentselected = -1
     submode = ESM_PASTING
End Sub

Public Sub StartSelectOperation(ByVal X As Single, ByVal Y As Single)
     
     'Start drag operation
     currentselected = -1
     submode = ESM_SELECTING
     
     'Redraw map
     RedrawMap
     
     'Draw initial rect
     Render_RectSwitched GrabX, -GrabY, X, -Y, PAL_MULTISELECTION, 2
     
     'Show changes
     picMap.Refresh
     
     'Keep X and Y
     LastSelX = X
     LastSelY = Y
End Sub

Private Sub picMap_Paint()
     On Error Resume Next
     
     'Check if running 3D Mode
     If (Running3D) And (Me.Visible = True) Then
          
          'Check if windowed
          If (Val(Config("windowedvideo")) <> 0) And (Running3D = True) Then
               
               'Run a single frame to refresh
               RunSingleFrame False, True
          End If
     End If
End Sub

Private Sub stbStatus_PanelClick(ByVal Panel As MSComctlLib.Panel)
     
     'Leave immediately if no map is loaded
     If (mapfile = "") Then Exit Sub
     
     'Leave when map edit is disabled
     If (picMap.Enabled = False) Then Exit Sub
     
     'Unable to click these while in 3D mode
     If (mode = EM_3D) Then Exit Sub
     
     'Check what panel is clicked
     Select Case Panel.Key
          
          'Zoom
          Case "viewzoom": frmZoom.Show 1, Me
          
          'Grid
          Case "gridsize": frmGrid.Show 1, Me
          
          'Snap
          Case "snapmode": itmEditSnapToGrid_Click
          
          'Stitch
          Case "stitchmode": itmEditStitch_Click
          
     End Select
End Sub

Private Sub tlbToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
     
     'Unable to click these while in 3D mode
     If (mode = EM_3D) Then Exit Sub
     
     'Check what button is clicked
     Select Case Button.Key
          Case "FileNewMap": itmFileNew_Click
          Case "FileOpenMap": itmFileOpenMap_Click
          Case "FileSaveMap": itmFileSaveMap_Click
          Case "ModeMove": itmEditMode_Click 0
          Case "ModeVertices": itmEditMode_Click 1
          Case "ModeLines": itmEditMode_Click 2
          Case "ModeSectors": itmEditMode_Click 3
          Case "ModeThings": itmEditMode_Click 4
          Case "Mode3D": itmEditMode_Click 5
          Case "FileTest": itmFileTest_Click False
          Case "FileBuild": itmFileBuild_Click
          Case "EditUndo": itmEditUndo_Click
          Case "EditRedo": itmEditRedo_Click
          Case "EditGrid": frmGrid.Show 1, Me
          Case "EditSnap": itmEditSnapToGrid_Click
          Case "EditStitch": itmEditStitch_Click
          Case "EditFlipH": itmEditFlipH_Click
          Case "EditFlipV": itmEditFlipV_Click
          Case "EditRotate": itmEditRotate_Click
          Case "EditResize": itmEditResize_Click
          Case "EditCenterView": itmEditCenterView_Click
          Case "PrefabsInsert": If mnuPrefabs.Visible Then mnuPrefabs_Click: If itmPrefabInsert.Enabled Then itmPrefabInsert_Click
          Case "PrefabsInsertPrevious": If mnuPrefabs.Visible Then mnuPrefabs_Click: If itmPrefabPrevious.Enabled Then itmPrefabPrevious_Click
          Case "LinesFlip": itmLinesFlipLinedefs_Click
          Case "LinesCurve": itmLinesCurve_Click
          Case "SectorsJoin": itmSectorsJoin_Click
          Case "SectorsMerge": itmSectorsMerge_Click
          Case "SectorsGradientBrightness": itmSectorsGradientBrightness_Click
          Case "SectorsGradientFloors": itmSectorsGradientFloors_Click
          Case "SectorsGradientCeilings": itmSectorsGradientCeilings_Click
          Case "ThingsFilter": itmThingsFilter_Click
     End Select
End Sub

Private Sub tmr3DRedraw_Timer()
     On Error GoTo Leave3DMode
     Dim ErrNumber As Long
     Dim ErrDesc As String
     
     'After this fame input will be processed again
     IgnoreInput = False
     
     'Check if still in 3D Mode
     If (Running3D) Then
          
          'Check if editing allowed
          If (picMap.Enabled) Then
               
               'Check if not in texture selection
               If (TextureSelecting = False) Then
                    
                    'Do single frame in 3D Mode
                    RunSingleLoop
               End If
          End If
     End If
     
     'Done
     Exit Sub
     
     
Leave3DMode:
     
     'Keep error
     ErrNumber = Err.number
     ErrDesc = Err.Description
     
     'Stop 3D Mode now
     Stop3DMode
     
     'Display error if not device lost error
     If (ErrNumber <> -2005530520) Then MsgBox "Error " & ErrNumber & " in 3D Mode: " & ErrDesc, vbCritical
End Sub

Private Sub tmrAutoScroll_Timer()
     
     'Disable timer
     tmrAutoScroll.Enabled = False
     
     'Check if we should stop autoscroling
     If (Config("autoscroll") = vbUnchecked) Then Exit Sub
     If (mode = EM_MOVE) Or (submode = ESM_NONE) Or (submode = ESM_MOVING) Then Exit Sub
     
     'Restart timer
     tmrAutoScroll.Enabled = True
     
     'Apply scrolling
     ViewLeft = ViewLeft + AutoScrollX
     ViewTop = ViewTop + AutoScrollY
     
     'Change viewport
     ChangeView ViewLeft, ViewTop, ViewZoom
     
     'Redraw map
     RedrawMap
End Sub

Private Sub tmrMouseOutside_Timer()
     Dim ViewRect As RECT
     Dim pnt As POINT
     
     'Leave immediately if no map is loaded
     If (mapfile = "") Then Exit Sub
     
     'Get the screen coordinates of the map box
     ClientToScreen frmMain.picMap.hWnd, pnt
     
     'Make the rect of the map on screen
     With ViewRect
          .left = pnt.X
          .top = pnt.Y
          .right = pnt.X + (frmMain.picMap.width - 4)
          .bottom = pnt.Y + (frmMain.picMap.height - 4)
     End With
     
     'Get the mouse coordinates
     GetCursorPos pnt
     
     'Check if cursor is outside viewport
     If ((pnt.X < ViewRect.left) Or _
        (pnt.X > ViewRect.right) Or _
        (pnt.Y < ViewRect.top) Or _
        (pnt.Y > ViewRect.bottom)) Then
          
          'Update cursor position in statusbar
          stbStatus.Panels("mousex").Text = ""
          stbStatus.Panels("mousey").Text = ""
          
          'Remove drawline when in drawing mode
          If (submode = ESM_DRAWING) Then
               
               'Remove drawing line
               RenderDrawingLine PAL_NORMAL, PAL_NORMAL, LastX, LastY
               
               'Show changes
               picMap.Refresh
          End If
          
          'Remove highlighting
          If (currentselected > -1) Then RemoveHighlight
          
          'No scrolling
          ChangeAutoscroll True
          
          'Erase last position
          MouseInside = False
          
          'Disable timer
          tmrMouseOutside.Enabled = False
     End If
End Sub

Private Sub tmrMouseTimeout_Timer()
     
     'Disable timer
     tmrMouseTimeout.Enabled = False
End Sub

Private Sub UpdateMenuShortcuts()
     
     'Set all shortcut tips on the menu items
     itmFile(0).Caption = MenuNameForShortcut(itmFile(0).Caption, "filenew")
     itmFile(1).Caption = MenuNameForShortcut(itmFile(1).Caption, "fileopen")
     itmFile(2).Caption = MenuNameForShortcut(itmFile(2).Caption, "fileclose")
     itmFile(4).Caption = MenuNameForShortcut(itmFile(4).Caption, "filesave")
     itmFile(5).Caption = MenuNameForShortcut(itmFile(5).Caption, "filesaveas")
     itmFile(6).Caption = MenuNameForShortcut(itmFile(6).Caption, "filesaveinto")
     itmFile(11).Caption = MenuNameForShortcut(itmFile(11).Caption, "filebuildnodes")
     itmFile(12).Caption = MenuNameForShortcut(itmFile(12).Caption, "filetest")
     itmFile(8).Caption = MenuNameForShortcut(itmFile(8).Caption, "fileexport")
     itmFile(9).Caption = MenuNameForShortcut(itmFile(9).Caption, "fileexportpicture")
     
     itmEditUndo.Caption = MenuNameForShortcut(itmEditUndo.Caption, "editundo")
     itmEditRedo.Caption = MenuNameForShortcut(itmEditRedo.Caption, "editredo")
     itmEditCut.Caption = MenuNameForShortcut(itmEditCut.Caption, "editcut")
     itmEditCopy.Caption = MenuNameForShortcut(itmEditCopy.Caption, "editcopy")
     itmEditPaste.Caption = MenuNameForShortcut(itmEditPaste.Caption, "editpaste")
     itmEditDelete.Caption = MenuNameForShortcut(itmEditDelete.Caption, "editdelete")
     itmEditFind.Caption = MenuNameForShortcut(itmEditFind.Caption, "editfind")
     itmEditReplace.Caption = MenuNameForShortcut(itmEditReplace.Caption, "editreplace")
     itmEditMode(0).Caption = MenuNameForShortcut(itmEditMode(0).Caption, "editmove")
     itmEditMode(1).Caption = MenuNameForShortcut(itmEditMode(1).Caption, "editvertices")
     itmEditMode(2).Caption = MenuNameForShortcut(itmEditMode(2).Caption, "editlines")
     itmEditMode(3).Caption = MenuNameForShortcut(itmEditMode(3).Caption, "editsectors")
     itmEditMode(4).Caption = MenuNameForShortcut(itmEditMode(4).Caption, "editthings")
     itmEditMode(5).Caption = MenuNameForShortcut(itmEditMode(5).Caption, "edit3d")
     itmEditMapOptions.Caption = MenuNameForShortcut(itmEditMapOptions.Caption, "editoptions")
     itmEditFlipH.Caption = MenuNameForShortcut(itmEditFlipH.Caption, "editfliph")
     itmEditFlipV.Caption = MenuNameForShortcut(itmEditFlipV.Caption, "editflipv")
     itmEditRotate.Caption = MenuNameForShortcut(itmEditRotate.Caption, "editrotate")
     itmEditResize.Caption = MenuNameForShortcut(itmEditResize.Caption, "editresize")
     itmEditCenterView.Caption = MenuNameForShortcut(itmEditCenterView.Caption, "editcenterview")
     
     itmVerticesSnapToGrid.Caption = MenuNameForShortcut(itmVerticesSnapToGrid.Caption, "snaptogrid")
     itmVerticesClearUnused.Caption = MenuNameForShortcut(itmVerticesClearUnused.Caption, "clearvertices")
     itmVerticesStitch.Caption = MenuNameForShortcut(itmVerticesStitch.Caption, "stitchvertices")
     
     itmLinesSnapToGrid.Caption = MenuNameForShortcut(itmLinesSnapToGrid.Caption, "snaptogrid")
     itmLinesAlign.Caption = MenuNameForShortcut(itmLinesAlign.Caption, "linesautoalign")
     itmLinesSelect(0).Caption = MenuNameForShortcut(itmLinesSelect(0).Caption, "select1sided")
     itmLinesSelect(1).Caption = MenuNameForShortcut(itmLinesSelect(1).Caption, "select2sided")
     itmLinesFlipLinedefs.Caption = MenuNameForShortcut(itmLinesFlipLinedefs.Caption, "fliplinedefs")
     itmLinesFlipSidedefs.Caption = MenuNameForShortcut(itmLinesFlipSidedefs.Caption, "flipsidedefs")
     itmLinesCurve.Caption = MenuNameForShortcut(itmLinesCurve.Caption, "curvelines")
     itmLinesCopy.Caption = MenuNameForShortcut(itmLinesCopy.Caption, "copyprops")
     itmLinesPaste.Caption = MenuNameForShortcut(itmLinesPaste.Caption, "pasteprops")
     
     itmSectorsSnapToGrid.Caption = MenuNameForShortcut(itmSectorsSnapToGrid.Caption, "snaptogrid")
     itmSectorsJoin.Caption = MenuNameForShortcut(itmSectorsJoin.Caption, "joinsector")
     itmSectorsMerge.Caption = MenuNameForShortcut(itmSectorsMerge.Caption, "mergesector")
     itmSectorsRaiseFloor.Caption = MenuNameForShortcut(itmSectorsRaiseFloor.Caption, "raisefloor")
     itmSectorsLowerFloor.Caption = MenuNameForShortcut(itmSectorsLowerFloor.Caption, "lowerfloor")
     itmSectorsRaiseCeiling.Caption = MenuNameForShortcut(itmSectorsRaiseCeiling.Caption, "raiseceil")
     itmSectorsLowerCeiling.Caption = MenuNameForShortcut(itmSectorsLowerCeiling.Caption, "lowerceil")
     itmSectorsIncBrightness.Caption = MenuNameForShortcut(itmSectorsIncBrightness.Caption, "brightinc")
     itmSectorsDecBrightness.Caption = MenuNameForShortcut(itmSectorsDecBrightness.Caption, "brightdec")
     itmSectorsCopy.Caption = MenuNameForShortcut(itmSectorsCopy.Caption, "copyprops")
     itmSectorsPaste.Caption = MenuNameForShortcut(itmSectorsPaste.Caption, "pasteprops")
     itmSectorsGradientBrightness.Caption = MenuNameForShortcut(itmSectorsGradientBrightness.Caption, "gradientbrightness")
     itmSectorsGradientFloors.Caption = MenuNameForShortcut(itmSectorsGradientFloors.Caption, "gradientceilings")
     itmSectorsGradientCeilings.Caption = MenuNameForShortcut(itmSectorsGradientCeilings.Caption, "gradientfloors")
     
     itmThingsSnapToGrid.Caption = MenuNameForShortcut(itmThingsSnapToGrid.Caption, "snaptogrid")
     itmThingsCopy.Caption = MenuNameForShortcut(itmThingsCopy.Caption, "copyprops")
     itmThingsPaste.Caption = MenuNameForShortcut(itmThingsPaste.Caption, "pasteprops")
     itmThingsFilter.Caption = MenuNameForShortcut(itmThingsFilter.Caption, "thingsfilter")
     
     itmToolsFindErrors.Caption = MenuNameForShortcut(itmToolsFindErrors.Caption, "errorcheck")
     itmToolsClearTextures.Caption = MenuNameForShortcut(itmToolsClearTextures.Caption, "removetextures")
     itmToolsFixTextures.Caption = MenuNameForShortcut(itmToolsFixTextures.Caption, "fixtextures")
     itmToolsConfiguration.Caption = MenuNameForShortcut(itmToolsConfiguration.Caption, "fileconfig")
     itmToolsFixZeroLinedefs.Caption = MenuNameForShortcut(itmToolsFixZeroLinedefs.Caption, "fixzerolengthlines")
     
     itmPrefabInsert.Caption = MenuNameForShortcut(itmPrefabInsert.Caption, "prefabinsert")
     itmPrefabPrevious.Caption = MenuNameForShortcut(itmPrefabPrevious.Caption, "prefabinsertlast")
     itmPrefabQuick(0).Caption = MenuNameForShortcut(itmPrefabQuick(0).Caption, "prefabinsert1")
     itmPrefabQuick(1).Caption = MenuNameForShortcut(itmPrefabQuick(1).Caption, "prefabinsert2")
     itmPrefabQuick(2).Caption = MenuNameForShortcut(itmPrefabQuick(2).Caption, "prefabinsert3")
     itmPrefabQuick(3).Caption = MenuNameForShortcut(itmPrefabQuick(3).Caption, "prefabinsert4")
     itmPrefabQuick(4).Caption = MenuNameForShortcut(itmPrefabQuick(4).Caption, "prefabinsert5")
     
     itmHelpWebsite.Caption = MenuNameForShortcut(itmHelpWebsite.Caption, "helpwebsite")
     itmHelpFAQ.Caption = MenuNameForShortcut(itmHelpFAQ.Caption, "helpfaq")
     itmHelpAbout.Caption = MenuNameForShortcut(itmHelpAbout.Caption, "helpabout")
End Sub

Private Sub tmrTerminate_Timer()
     
     'Terminate program
     Terminate
End Sub


