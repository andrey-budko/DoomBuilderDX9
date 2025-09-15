VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuration"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8100
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
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   415
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      Height          =   4275
      Index           =   2
      Left            =   165
      TabIndex        =   75
      Tag             =   "Files"
      Top             =   960
      Visible         =   0   'False
      Width           =   7725
      Begin VB.CommandButton cmdGameIWADClear 
         Caption         =   "Clear"
         Height          =   345
         Left            =   6240
         TabIndex        =   79
         Top             =   3870
         Width           =   1065
      End
      Begin VB.CommandButton cmdGameIWAD 
         Caption         =   "Browse..."
         Height          =   345
         Left            =   5175
         TabIndex        =   78
         Top             =   3870
         Width           =   1065
      End
      Begin VB.TextBox txtGameIWAD 
         BackColor       =   &H80000000&
         Height          =   315
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   77
         Top             =   3885
         Width           =   4800
      End
      Begin MSComctlLib.ListView lstGames 
         Height          =   2655
         Left            =   300
         TabIndex        =   76
         Top             =   780
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Game"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label Label30 
         Caption         =   $"frmOptions.frx":000C
         Height          =   675
         Left            =   315
         TabIndex        =   80
         Top             =   90
         Width           =   6405
      End
      Begin VB.Label lblGameIWAD 
         AutoSize        =   -1  'True
         Caption         =   "Doom Shareware IWAD:"
         Height          =   210
         Left            =   315
         TabIndex        =   81
         Top             =   3645
         UseMnemonic     =   0   'False
         Width           =   1770
      End
   End
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      Height          =   4560
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Tag             =   "Editing"
      Top             =   1005
      Visible         =   0   'False
      Width           =   7635
      Begin VB.CheckBox chkPasteAdjustsHeights 
         Caption         =   "Adjust pasted sectors to correct relative heights"
         Height          =   225
         Left            =   555
         TabIndex        =   241
         Top             =   2970
         Width           =   3795
      End
      Begin VB.CheckBox chkStoreEditingInfo 
         Caption         =   "Save specific editing settings for each map"
         Height          =   225
         Left            =   2895
         TabIndex        =   239
         Top             =   4275
         Visible         =   0   'False
         Width           =   3420
      End
      Begin VB.CheckBox chkCopyTagsPaste 
         Caption         =   "Copy effects and tags when pasting"
         Height          =   225
         Left            =   555
         TabIndex        =   10
         Top             =   2505
         Width           =   3285
      End
      Begin VB.CheckBox chkCopyTagsDraw 
         Caption         =   "Copy adjacent effect and tag on drawing"
         Height          =   225
         Left            =   555
         TabIndex        =   9
         Top             =   2040
         Width           =   3285
      End
      Begin VB.CheckBox chkNewThingsDialog 
         Caption         =   "Show properties on creating Thing"
         Height          =   225
         Left            =   555
         TabIndex        =   6
         Top             =   1125
         Width           =   3225
      End
      Begin VB.CheckBox chkNewSectorDialog 
         Caption         =   "Show properties on creating Sector"
         Height          =   225
         Left            =   555
         TabIndex        =   7
         Top             =   1575
         Width           =   3225
      End
      Begin VB.CheckBox chkNewLinesDailog 
         Caption         =   "Show properties on creating Lines"
         Height          =   225
         Left            =   150
         TabIndex        =   8
         Top             =   4290
         Visible         =   0   'False
         Width           =   3225
      End
      Begin VB.CheckBox chkMixResources 
         Caption         =   "Mix Textures and Flats resources"
         Height          =   225
         Left            =   150
         TabIndex        =   3
         Top             =   4395
         Visible         =   0   'False
         Width           =   3105
      End
      Begin VB.CheckBox chkSubUndos 
         Caption         =   "Enable Undos for automatic features"
         Height          =   225
         Left            =   4695
         TabIndex        =   5
         Top             =   4350
         Visible         =   0   'False
         Width           =   3105
      End
      Begin VB.CheckBox chkSaveBackup 
         Caption         =   "Make backup when saving map"
         Height          =   225
         Left            =   3165
         TabIndex        =   4
         Top             =   4380
         Visible         =   0   'False
         Width           =   3105
      End
      Begin VB.CheckBox chkTexturePrecache 
         Caption         =   "Precache resources when loading map"
         Height          =   225
         Left            =   555
         TabIndex        =   2
         Top             =   660
         Width           =   3225
      End
      Begin DoomBuilder.ctlValueBox valMaxUndos 
         Height          =   360
         Left            =   6405
         TabIndex        =   16
         Top             =   4110
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         Max             =   100
         MaxLength       =   4
         Min             =   10
         SmallChange     =   10
         Value           =   "10"
      End
      Begin DoomBuilder.ctlValueBox valVertexSelectDistance 
         Height          =   360
         Left            =   6015
         TabIndex        =   11
         ToolTipText     =   "Screen Pixels"
         Top             =   135
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         Max             =   9990
         MaxLength       =   4
         Min             =   1
         Value           =   "2"
      End
      Begin DoomBuilder.ctlValueBox valLinedefSelectDistance 
         Height          =   360
         Left            =   6015
         TabIndex        =   12
         ToolTipText     =   "Screen Pixels"
         Top             =   645
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         Max             =   9990
         MaxLength       =   4
         Min             =   1
         Value           =   "2"
      End
      Begin DoomBuilder.ctlValueBox valThingSelectDistance 
         Height          =   360
         Left            =   6015
         TabIndex        =   13
         ToolTipText     =   "Screen Pixels"
         Top             =   1155
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         Max             =   9990
         MaxLength       =   4
         Min             =   1
         Value           =   "2"
      End
      Begin DoomBuilder.ctlValueBox valAutostitchDistance 
         Height          =   360
         Left            =   6015
         TabIndex        =   14
         ToolTipText     =   "Map Pixels"
         Top             =   1665
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         Max             =   9990
         MaxLength       =   4
         Value           =   "1"
      End
      Begin DoomBuilder.ctlValueBox valLinesplitDistance 
         Height          =   360
         Left            =   6015
         TabIndex        =   15
         ToolTipText     =   "Map Pixels"
         Top             =   2175
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         Max             =   9990
         MaxLength       =   4
         Value           =   "1"
      End
      Begin VB.CheckBox chkTextureCaching 
         Caption         =   "Cache resources in memory"
         Height          =   225
         Left            =   555
         TabIndex        =   1
         Top             =   195
         Width           =   3105
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Available Undo levels:"
         Height          =   210
         Left            =   4725
         TabIndex        =   21
         Top             =   4170
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Linedef split range:"
         Height          =   210
         Left            =   4530
         TabIndex        =   20
         Top             =   2235
         UseMnemonic     =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Stitch vertices range:"
         Height          =   210
         Left            =   4365
         TabIndex        =   18
         Top             =   1725
         UseMnemonic     =   0   'False
         Width           =   1545
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Thing selection range:"
         Height          =   210
         Left            =   4320
         TabIndex        =   19
         Top             =   1215
         UseMnemonic     =   0   'False
         Width           =   1590
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Linedef selection range:"
         Height          =   210
         Left            =   4170
         TabIndex        =   17
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   1740
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vertex selection range:"
         Height          =   210
         Left            =   4215
         TabIndex        =   22
         Top             =   195
         UseMnemonic     =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      Height          =   4425
      Index           =   3
      Left            =   165
      TabIndex        =   82
      Tag             =   "Defaults"
      Top             =   840
      Visible         =   0   'False
      Width           =   7635
      Begin VB.Frame fraMapDefaults 
         Caption         =   " Startup Mapping Defaults "
         Height          =   3180
         Left            =   135
         TabIndex        =   88
         Top             =   1110
         Width           =   7425
         Begin VB.TextBox txtDefaultTFloor 
            Height          =   315
            Left            =   5310
            MaxLength       =   8
            TabIndex        =   94
            Text            =   "STARTAN3"
            Top             =   750
            Width           =   1215
         End
         Begin VB.TextBox txtDefaultTCeiling 
            Height          =   315
            Left            =   5310
            MaxLength       =   8
            TabIndex        =   93
            Text            =   "STARTAN3"
            Top             =   345
            Width           =   1215
         End
         Begin VB.TextBox txtDefaultLower 
            Height          =   315
            Left            =   1905
            MaxLength       =   8
            TabIndex        =   91
            Text            =   "STARTAN3"
            Top             =   1155
            Width           =   1215
         End
         Begin VB.TextBox txtDefaultMiddle 
            Height          =   315
            Left            =   1905
            MaxLength       =   8
            TabIndex        =   90
            Text            =   "STARTAN3"
            Top             =   750
            Width           =   1215
         End
         Begin VB.TextBox txtDefaultUpper 
            Height          =   315
            Left            =   1905
            MaxLength       =   8
            TabIndex        =   89
            Text            =   "STARTAN3"
            Top             =   345
            Width           =   1215
         End
         Begin DoomBuilder.ctlValueBox valDefaultHCeiling 
            Height          =   360
            Left            =   5310
            TabIndex        =   95
            Top             =   1140
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   635
            MaxLength       =   4
            Min             =   -32767
            SmallChange     =   8
            Value           =   "2"
         End
         Begin DoomBuilder.ctlValueBox valDefaultHFloor 
            Height          =   360
            Left            =   5310
            TabIndex        =   96
            Top             =   1560
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   635
            MaxLength       =   4
            Min             =   -32767
            SmallChange     =   8
            Value           =   "2"
         End
         Begin DoomBuilder.ctlValueBox valDefaultBrightness 
            Height          =   360
            Left            =   5310
            TabIndex        =   97
            Top             =   1980
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   635
            Max             =   255
            MaxLength       =   4
            Value           =   "2"
         End
         Begin DoomBuilder.ctlValueBox valDefaultThing 
            Height          =   360
            Left            =   1905
            TabIndex        =   92
            Top             =   1560
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   635
            MaxLength       =   4
         End
         Begin VB.Label Label56 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Insert Thing:"
            Height          =   210
            Left            =   870
            TabIndex        =   104
            Top             =   1620
            UseMnemonic     =   0   'False
            Width           =   885
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Sector Brightness:"
            Height          =   210
            Left            =   3810
            TabIndex        =   106
            Top             =   2040
            UseMnemonic     =   0   'False
            Width           =   1350
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Floor Height:"
            Height          =   210
            Left            =   4260
            TabIndex        =   105
            Top             =   1620
            UseMnemonic     =   0   'False
            Width           =   900
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ceiling Height:"
            Height          =   210
            Left            =   4155
            TabIndex        =   103
            Top             =   1200
            UseMnemonic     =   0   'False
            Width           =   1005
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Floor Texture:"
            Height          =   210
            Left            =   4155
            TabIndex        =   101
            Top             =   795
            UseMnemonic     =   0   'False
            Width           =   1005
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ceiling Texture:"
            Height          =   210
            Left            =   4050
            TabIndex        =   99
            Top             =   390
            UseMnemonic     =   0   'False
            Width           =   1110
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Lower Texture:"
            Height          =   210
            Left            =   630
            TabIndex        =   102
            Top             =   1200
            UseMnemonic     =   0   'False
            Width           =   1125
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Middle Texture:"
            Height          =   210
            Left            =   660
            TabIndex        =   100
            Top             =   795
            UseMnemonic     =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Upper Texture:"
            Height          =   210
            Left            =   675
            TabIndex        =   98
            Top             =   390
            UseMnemonic     =   0   'False
            Width           =   1080
         End
      End
      Begin VB.Frame fraStartDefaults 
         Caption         =   " Map Open Defaults "
         Height          =   930
         Left            =   135
         TabIndex        =   83
         Top             =   90
         Width           =   7425
         Begin VB.CheckBox chkDefaultStitch 
            Caption         =   "Stitch vertices"
            Height          =   240
            Left            =   5115
            TabIndex        =   86
            Top             =   435
            Width           =   1545
         End
         Begin VB.CheckBox chkDefaultSnap 
            Caption         =   "Snap to grid"
            Height          =   240
            Left            =   3345
            TabIndex        =   85
            Top             =   435
            Width           =   1305
         End
         Begin DoomBuilder.ctlValueBox valDefaultGrid 
            Height          =   360
            Left            =   1620
            TabIndex        =   84
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   635
            Max             =   1024
            MaxLength       =   4
            Min             =   2
            SmallChange     =   8
            Value           =   "2"
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Grid Size:"
            Height          =   210
            Left            =   780
            TabIndex        =   87
            Top             =   435
            UseMnemonic     =   0   'False
            Width           =   705
         End
      End
   End
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      Height          =   4395
      Index           =   5
      Left            =   165
      TabIndex        =   128
      Tag             =   "Testing"
      Top             =   870
      Visible         =   0   'False
      Width           =   7635
      Begin VB.Frame Frame1 
         Caption         =   " Testing "
         Height          =   3795
         Left            =   135
         TabIndex        =   131
         Top             =   450
         Width           =   7425
         Begin VB.TextBox txtTestParams 
            Height          =   315
            Left            =   1980
            TabIndex        =   136
            Text            =   "-executethis - executethat"
            Top             =   1020
            Width           =   4845
         End
         Begin VB.CommandButton cmdBrowseEngine 
            Caption         =   "Browse..."
            Height          =   345
            Left            =   5580
            TabIndex        =   134
            Top             =   600
            Width           =   1245
         End
         Begin VB.TextBox txtTestExe 
            BackColor       =   &H80000000&
            Height          =   315
            Left            =   1980
            Locked          =   -1  'True
            TabIndex        =   133
            Text            =   "Nodebuilder.exe"
            Top             =   615
            Width           =   3525
         End
         Begin VB.CheckBox chkTestingDialog 
            Caption         =   "Always show me these options before testing"
            Height          =   255
            Left            =   1980
            TabIndex        =   142
            Top             =   3105
            Width           =   4125
         End
         Begin VB.Label Label83 
            Caption         =   "Use %M to indicate the Map number from E#M# or MAP## map name"
            Height          =   225
            Left            =   1980
            TabIndex        =   235
            Top             =   2685
            Width           =   5055
         End
         Begin VB.Label Label82 
            Caption         =   "Use %E to indicate the Episode number from E#M# map name"
            Height          =   225
            Left            =   1980
            TabIndex        =   234
            Top             =   2460
            Width           =   5055
         End
         Begin VB.Label Label50 
            Caption         =   "Use %F to indicate the edited PWAD file to be tested"
            Height          =   210
            Left            =   1980
            TabIndex        =   137
            Top             =   1380
            Width           =   5055
         End
         Begin VB.Label Label51 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Parameters:"
            Height          =   210
            Left            =   930
            TabIndex        =   135
            Top             =   1065
            UseMnemonic     =   0   'False
            Width           =   870
         End
         Begin VB.Label Label52 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Engine:"
            Height          =   210
            Left            =   1275
            TabIndex        =   132
            Top             =   660
            UseMnemonic     =   0   'False
            Width           =   525
         End
         Begin VB.Label Label53 
            Caption         =   "Use %W to indicate the current IWAD file to be used (path included)"
            Height          =   210
            Left            =   1980
            TabIndex        =   138
            Top             =   1590
            Width           =   5055
         End
         Begin VB.Label Label54 
            Caption         =   "Use %D to indicate the current IWAD file to be used (without path)"
            Height          =   210
            Left            =   1980
            TabIndex        =   139
            Top             =   1800
            Width           =   5055
         End
         Begin VB.Label Label55 
            Caption         =   "Use %L to indicate the lumpname as is set in the map options"
            Height          =   225
            Left            =   1980
            TabIndex        =   140
            Top             =   2010
            Width           =   5055
         End
         Begin VB.Label Label68 
            Caption         =   "Use %A to indicate the PWAD file with additional textures (if any)"
            Height          =   225
            Left            =   1980
            TabIndex        =   141
            Top             =   2235
            Width           =   5055
         End
      End
      Begin VB.ComboBox cmbTestQuickload 
         Height          =   330
         Left            =   5145
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   130
         Top             =   90
         Width           =   2415
      End
      Begin VB.Label Label81 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Select here a profile to quickly set the Testing parameters:"
         Height          =   210
         Left            =   765
         TabIndex        =   129
         Top             =   150
         Width           =   4200
      End
   End
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      Height          =   4680
      Index           =   7
      Left            =   165
      TabIndex        =   154
      Tag             =   "3D Mode"
      Top             =   870
      Visible         =   0   'False
      Width           =   7635
      Begin VB.Frame fraNo3D 
         BorderStyle     =   0  'None
         Height          =   4380
         Left            =   7185
         TabIndex        =   185
         Top             =   4125
         Visible         =   0   'False
         Width           =   7590
         Begin VB.Label lblNo3D 
            Alignment       =   2  'Center
            Caption         =   $"frmOptions.frx":0109
            Height          =   435
            Left            =   480
            TabIndex        =   186
            Top             =   840
            Width           =   6705
         End
      End
      Begin VB.CheckBox chkLinesSectorsInfo 
         Caption         =   "Show lines/sectors information in panel"
         Height          =   300
         Left            =   3090
         TabIndex        =   174
         Top             =   4530
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.CheckBox chkStandardTextureBrowse 
         Caption         =   "Use standard texture browser"
         Height          =   300
         Left            =   585
         TabIndex        =   169
         Top             =   3615
         Width           =   2850
      End
      Begin VB.CheckBox chkWindowed 
         Caption         =   "Windowed 3D Mode"
         Height          =   360
         Left            =   1515
         TabIndex        =   155
         Top             =   345
         Width           =   2850
      End
      Begin VB.CheckBox chkExclusivemouse 
         Caption         =   "Exclusive mouse access for Doom Builder"
         Height          =   300
         Left            =   4005
         TabIndex        =   173
         Top             =   3615
         Width           =   3375
      End
      Begin VB.CheckBox chkRaiseLowerCeiling 
         Caption         =   "Move ceiling when aimed at sidedef"
         Height          =   300
         Left            =   4005
         TabIndex        =   172
         Top             =   3255
         Width           =   3165
      End
      Begin VB.CheckBox chkBelowCeiling 
         Caption         =   "Stay below ceiling with gravity on"
         Height          =   300
         Left            =   75
         TabIndex        =   168
         Top             =   4335
         Visible         =   0   'False
         Width           =   3105
      End
      Begin VB.ComboBox cmbVideoDriver 
         Height          =   330
         ItemData        =   "frmOptions.frx":01C5
         Left            =   1515
         List            =   "frmOptions.frx":01C7
         Style           =   2  'Dropdown List
         TabIndex        =   156
         Top             =   795
         Width           =   2850
      End
      Begin VB.CheckBox chkFog 
         Caption         =   "Enable Fog"
         Height          =   300
         Left            =   585
         TabIndex        =   165
         Top             =   2895
         Width           =   1995
      End
      Begin VB.ComboBox cmbResolution 
         Height          =   330
         ItemData        =   "frmOptions.frx":01C9
         Left            =   1515
         List            =   "frmOptions.frx":01CB
         Style           =   2  'Dropdown List
         TabIndex        =   157
         Top             =   1245
         Width           =   2850
      End
      Begin VB.CheckBox chkInvertY 
         Caption         =   "Invert mouse Y axis"
         Height          =   300
         Left            =   585
         TabIndex        =   166
         Top             =   3255
         Width           =   2715
      End
      Begin VB.CheckBox chkAspect 
         Caption         =   "Fixed resolution aspect"
         Height          =   300
         Left            =   4005
         TabIndex        =   170
         Top             =   2895
         Width           =   2070
      End
      Begin VB.ComboBox cmbTextureFilter 
         Height          =   330
         ItemData        =   "frmOptions.frx":01CD
         Left            =   1515
         List            =   "frmOptions.frx":01DA
         Style           =   2  'Dropdown List
         TabIndex        =   158
         Top             =   1695
         Width           =   2850
      End
      Begin VB.CheckBox chkDirectXPrecache 
         Caption         =   "Precache resources when starting"
         Height          =   300
         Left            =   75
         TabIndex        =   167
         Top             =   4500
         Visible         =   0   'False
         Width           =   2850
      End
      Begin VB.CheckBox chkVertexbufferCache 
         Caption         =   "Cache structure data buffers"
         Height          =   300
         Left            =   3195
         TabIndex        =   171
         Top             =   4305
         Visible         =   0   'False
         Width           =   3315
      End
      Begin DoomBuilder.ctlValueBox txtFOV 
         Height          =   360
         Left            =   5925
         TabIndex        =   160
         ToolTipText     =   "Field Of View in 3D mode"
         Top             =   345
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         Max             =   160
         MaxLength       =   3
         Min             =   20
         SmallChange     =   10
         Value           =   "90"
      End
      Begin DoomBuilder.ctlValueBox txtMoveSpeed 
         Height          =   360
         Left            =   5925
         TabIndex        =   161
         ToolTipText     =   "Movement speed in map pixels per second"
         Top             =   795
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         Max             =   2048
         MaxLength       =   3
         Min             =   128
         SmallChange     =   32
         Value           =   "512"
      End
      Begin DoomBuilder.ctlValueBox txtMouseSpeed 
         Height          =   360
         Left            =   5925
         TabIndex        =   162
         ToolTipText     =   "Mouse sensitivity"
         Top             =   1245
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         Max             =   1000
         MaxLength       =   4
         Min             =   10
         SmallChange     =   10
         Value           =   "20"
      End
      Begin DoomBuilder.ctlValueBox txtGamma 
         Height          =   360
         Left            =   5925
         TabIndex        =   163
         ToolTipText     =   "Gamma correction (support dependant on videocard)"
         Top             =   1695
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         Max             =   100
         MaxLength       =   4
         Min             =   -100
         SmallChange     =   10
      End
      Begin DoomBuilder.ctlValueBox txtBrightness 
         Height          =   360
         Left            =   5925
         TabIndex        =   164
         ToolTipText     =   "Brightness (support dependant on videocard)"
         Top             =   2145
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         Max             =   100
         MaxLength       =   4
         Min             =   -100
         SmallChange     =   10
      End
      Begin DoomBuilder.ctlValueBox txtVideoDistance 
         Height          =   360
         Left            =   1515
         TabIndex        =   159
         ToolTipText     =   "Viewable distance in mappixels"
         Top             =   2100
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         Max             =   10000
         MaxLength       =   5
         Min             =   500
         SmallChange     =   500
         Value           =   ""
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         Caption         =   "mappixels"
         Height          =   210
         Left            =   2865
         TabIndex        =   182
         Top             =   2175
         Width           =   720
      End
      Begin VB.Label Label72 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "View distance:"
         Height          =   210
         Left            =   300
         TabIndex        =   181
         Top             =   2175
         Width           =   1095
      End
      Begin VB.Label Label70 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Brightness:"
         Height          =   210
         Left            =   4980
         TabIndex        =   184
         Top             =   2205
         Width           =   825
      End
      Begin VB.Label Label69 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Gamma:"
         Height          =   210
         Left            =   5220
         TabIndex        =   183
         Top             =   1755
         Width           =   585
      End
      Begin VB.Label lblVideoDriver 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Video Driver:"
         Height          =   210
         Left            =   450
         TabIndex        =   175
         Top             =   840
         Width           =   945
      End
      Begin VB.Label lblResolution 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Resolution:"
         Height          =   210
         Left            =   600
         TabIndex        =   177
         Top             =   1290
         Width           =   795
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "FOV:"
         Height          =   210
         Left            =   5400
         TabIndex        =   176
         Top             =   420
         Width           =   375
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Move Speed:"
         Height          =   210
         Left            =   4860
         TabIndex        =   178
         Top             =   870
         Width           =   945
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Mouse Speed:"
         Height          =   210
         Left            =   4770
         TabIndex        =   180
         Top             =   1320
         Width           =   1035
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Texture Filtering:"
         Height          =   210
         Left            =   195
         TabIndex        =   179
         Top             =   1740
         Width           =   1200
      End
   End
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      Height          =   4395
      Index           =   1
      Left            =   195
      TabIndex        =   23
      Tag             =   "Colors"
      Top             =   1110
      Visible         =   0   'False
      Width           =   7635
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   27
         Left            =   4290
         Style           =   1  'Graphical
         TabIndex        =   34
         Tag             =   "CLR_MAPBOUNDARY"
         Top             =   630
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   26
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   243
         Tag             =   "CLR_LINEBLOCKSOUND"
         Top             =   2880
         Width           =   375
      End
      Begin VB.CheckBox chkHighlighSyntax 
         Caption         =   "Syntax highlighting"
         Height          =   210
         Left            =   5385
         TabIndex        =   43
         Top             =   1035
         Width           =   1665
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   25
         Left            =   6615
         Style           =   1  'Graphical
         TabIndex        =   237
         Tag             =   "CLR_SCRIPTCONSTANT"
         Top             =   3630
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   24
         Left            =   6615
         Style           =   1  'Graphical
         TabIndex        =   45
         Tag             =   "CLR_SCRIPTLINENUMBERS"
         Top             =   1755
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   14
         Left            =   4290
         Style           =   1  'Graphical
         TabIndex        =   38
         Tag             =   "CLR_THINGSELECTED"
         Top             =   2130
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   23
         Left            =   6615
         Style           =   1  'Graphical
         TabIndex        =   49
         Tag             =   "CLR_SCRIPTSTRING"
         Top             =   3255
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   22
         Left            =   6615
         Style           =   1  'Graphical
         TabIndex        =   44
         Tag             =   "CLR_SCRIPTBACKGROUND"
         Top             =   1380
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   21
         Left            =   6615
         Style           =   1  'Graphical
         TabIndex        =   48
         Tag             =   "CLR_SCRIPTKEYWORD"
         Top             =   2880
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   20
         Left            =   6615
         Style           =   1  'Graphical
         TabIndex        =   47
         Tag             =   "CLR_SCRIPTCOMMENT"
         Top             =   2505
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   19
         Left            =   6615
         Style           =   1  'Graphical
         TabIndex        =   46
         Tag             =   "CLR_SCRIPTTEXT"
         Top             =   2130
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   0
         Left            =   4290
         Style           =   1  'Graphical
         TabIndex        =   33
         Tag             =   "CLR_BACKGROUND"
         Top             =   255
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   24
         Tag             =   "CLR_VERTEX"
         Top             =   255
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   4
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   27
         Tag             =   "CLR_LINE"
         Top             =   1380
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   6
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   28
         Tag             =   "CLR_LINESPECIAL"
         Top             =   1755
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   18
         Left            =   4290
         Style           =   1  'Graphical
         TabIndex        =   42
         Tag             =   "CLR_GRID64"
         Top             =   3630
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   17
         Left            =   4290
         Style           =   1  'Graphical
         TabIndex        =   41
         Tag             =   "CLR_GRID"
         Top             =   3255
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   16
         Left            =   4290
         Style           =   1  'Graphical
         TabIndex        =   40
         Tag             =   "CLR_MULTISELECT"
         Top             =   2880
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   15
         Left            =   4290
         Style           =   1  'Graphical
         TabIndex        =   39
         Tag             =   "CLR_THINGHIGHLIGHT"
         Top             =   2505
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   13
         Left            =   4290
         Style           =   1  'Graphical
         TabIndex        =   37
         Tag             =   "CLR_THINGUNKNOWN"
         Top             =   1755
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   12
         Left            =   4290
         Style           =   1  'Graphical
         TabIndex        =   36
         Tag             =   "CLR_SECTORTAG"
         Top             =   1380
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   11
         Left            =   4290
         Style           =   1  'Graphical
         TabIndex        =   35
         Tag             =   "CLR_THINGTAG"
         Top             =   1005
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   10
         Left            =   4305
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         Tag             =   "CLR_LINEDRAG"
         Top             =   -150
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   9
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   32
         Tag             =   "CLR_LINEHIGHLIGHT"
         Top             =   3630
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   8
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   31
         Tag             =   "CLR_LINESELECTED"
         Top             =   3255
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   7
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   30
         Tag             =   "CLR_LINESPECIALDOUBLE"
         Top             =   2505
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   5
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   29
         Tag             =   "CLR_LINEDOUBLE"
         Top             =   2130
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   3
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   26
         Tag             =   "CLR_VERTEXHIGHLIGHT"
         Top             =   1005
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   2
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   25
         Tag             =   "CLR_VERTEXSELECTED"
         Top             =   630
         Width           =   375
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Editing Boundaries:"
         Height          =   210
         Left            =   2850
         TabIndex        =   245
         Top             =   690
         UseMnemonic     =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sound Blocking Line:"
         Height          =   210
         Left            =   465
         TabIndex        =   244
         Top             =   2940
         UseMnemonic     =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label85 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Script Constant:"
         Height          =   210
         Left            =   5385
         TabIndex        =   238
         Top             =   3690
         UseMnemonic     =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label84 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Script Linenumbers:"
         Height          =   210
         Left            =   5100
         TabIndex        =   236
         Top             =   1815
         UseMnemonic     =   0   'False
         Width           =   1440
      End
      Begin VB.Label Label80 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Script String/Number:"
         Height          =   210
         Left            =   5010
         TabIndex        =   74
         Top             =   3315
         UseMnemonic     =   0   'False
         Width           =   1530
      End
      Begin VB.Label Label78 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Script Background:"
         Height          =   210
         Left            =   5160
         TabIndex        =   62
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label76 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Script Keyword:"
         Height          =   210
         Left            =   5355
         TabIndex        =   71
         Top             =   2940
         UseMnemonic     =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label75 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Script Comment:"
         Height          =   210
         Left            =   5370
         TabIndex        =   68
         Top             =   2565
         UseMnemonic     =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label74 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Script Text:"
         Height          =   210
         Left            =   5715
         TabIndex        =   65
         Top             =   2190
         UseMnemonic     =   0   'False
         Width           =   825
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "64 Pixels Grid:"
         Height          =   210
         Left            =   3195
         TabIndex        =   73
         Top             =   3690
         UseMnemonic     =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Custom Grid:"
         Height          =   210
         Left            =   3300
         TabIndex        =   70
         Top             =   3315
         UseMnemonic     =   0   'False
         Width           =   930
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Selection:"
         Height          =   210
         Left            =   3540
         TabIndex        =   67
         Top             =   2940
         UseMnemonic     =   0   'False
         Width           =   690
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Highlighted Thing:"
         Height          =   210
         Left            =   2970
         TabIndex        =   64
         Top             =   2565
         UseMnemonic     =   0   'False
         Width           =   1260
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Selected Thing:"
         Height          =   210
         Left            =   3120
         TabIndex        =   61
         Top             =   2190
         UseMnemonic     =   0   'False
         Width           =   1110
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Unknown Thing:"
         Height          =   210
         Left            =   3060
         TabIndex        =   59
         Top             =   1815
         UseMnemonic     =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tagged Sector:"
         Height          =   210
         Left            =   3120
         TabIndex        =   57
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   1110
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tagged Thing:"
         Height          =   210
         Left            =   3210
         TabIndex        =   54
         Top             =   1065
         UseMnemonic     =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Changing Linedef:"
         Height          =   210
         Left            =   2955
         TabIndex        =   50
         Top             =   -105
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Highlighted Line:"
         Height          =   210
         Left            =   810
         TabIndex        =   72
         Top             =   3690
         UseMnemonic     =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Selected Line:"
         Height          =   210
         Left            =   960
         TabIndex        =   69
         Top             =   3315
         UseMnemonic     =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Special Twosided Line:"
         Height          =   210
         Left            =   285
         TabIndex        =   66
         Top             =   2565
         UseMnemonic     =   0   'False
         Width           =   1680
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Special Line:"
         Height          =   210
         Left            =   1050
         TabIndex        =   60
         Top             =   1815
         UseMnemonic     =   0   'False
         Width           =   915
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Twosided Line:"
         Height          =   210
         Left            =   855
         TabIndex        =   63
         Top             =   2190
         UseMnemonic     =   0   'False
         Width           =   1110
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Common Line:"
         Height          =   210
         Left            =   960
         TabIndex        =   58
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Highlighted Vertex:"
         Height          =   210
         Left            =   600
         TabIndex        =   55
         Top             =   1065
         UseMnemonic     =   0   'False
         Width           =   1365
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Selected Vertex:"
         Height          =   210
         Left            =   750
         TabIndex        =   53
         Top             =   690
         UseMnemonic     =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Common Vertex:"
         Height          =   210
         Left            =   765
         TabIndex        =   51
         Top             =   315
         UseMnemonic     =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Background:"
         Height          =   210
         Left            =   3315
         TabIndex        =   52
         Top             =   315
         UseMnemonic     =   0   'False
         Width           =   915
      End
   End
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      Height          =   4695
      Index           =   6
      Left            =   165
      TabIndex        =   143
      Tag             =   "Shortcuts"
      Top             =   840
      Visible         =   0   'False
      Width           =   7725
      Begin VB.CheckBox chkModeKeys3D 
         Caption         =   "Use Mode switch keys also in 3D Mode"
         Height          =   285
         Left            =   150
         TabIndex        =   242
         Top             =   4350
         Width           =   3675
      End
      Begin VB.TextBox txtShortcut 
         Height          =   330
         Left            =   4050
         TabIndex        =   148
         TabStop         =   0   'False
         Top             =   2085
         Width           =   2265
      End
      Begin VB.ComboBox cmbShortcut 
         Height          =   330
         IntegralHeight  =   0   'False
         ItemData        =   "frmOptions.frx":022A
         Left            =   4050
         List            =   "frmOptions.frx":022C
         Style           =   2  'Dropdown List
         TabIndex        =   151
         TabStop         =   0   'False
         Top             =   2865
         Width           =   3465
      End
      Begin VB.PictureBox picWarning 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         HasDC           =   0   'False
         Height          =   720
         Left            =   3930
         ScaleHeight     =   46
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   248
         TabIndex        =   152
         TabStop         =   0   'False
         Top             =   3555
         Width           =   3750
         Begin VB.Label lblWarning 
            BackStyle       =   0  'Transparent
            Caption         =   "Warning: Some functions may not work when having keys assigned that are assigned to other functions as well."
            ForeColor       =   &H80000017&
            Height          =   645
            Left            =   345
            TabIndex        =   153
            Top             =   30
            UseMnemonic     =   0   'False
            Width           =   3330
         End
         Begin VB.Image imgWarning 
            Height          =   240
            Left            =   45
            Picture         =   "frmOptions.frx":022E
            Top             =   210
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdUnbind 
         Caption         =   "Clear"
         Enabled         =   0   'False
         Height          =   330
         Left            =   6390
         TabIndex        =   149
         TabStop         =   0   'False
         Top             =   2085
         Width           =   1110
      End
      Begin MSComctlLib.ListView lstFunctions 
         Height          =   4080
         Left            =   150
         TabIndex        =   144
         TabStop         =   0   'False
         Top             =   195
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   7197
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Function"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Keys"
            Object.Width           =   1765
         EndProperty
      End
      Begin VB.Label lblSpecialShortcut 
         AutoSize        =   -1  'True
         Caption         =   "Or select a special input type here:"
         Height          =   210
         Left            =   4050
         TabIndex        =   150
         Top             =   2625
         Width           =   2520
      End
      Begin VB.Label Label57 
         Caption         =   "Enter here the new key combination to assign to the selected function:"
         Height          =   465
         Left            =   4050
         TabIndex        =   147
         Top             =   1635
         Width           =   3180
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblFunctionDesc 
         Height          =   1155
         Left            =   4050
         TabIndex        =   146
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   3390
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblFunction 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4050
         TabIndex        =   145
         Top             =   195
         UseMnemonic     =   0   'False
         Width           =   3390
      End
   End
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      Height          =   4425
      Index           =   8
      Left            =   165
      TabIndex        =   187
      Tag             =   "Prefabs"
      Top             =   840
      Width           =   7635
      Begin VB.CommandButton cmdPrefabFolder 
         Caption         =   "Browse..."
         Height          =   345
         Left            =   5940
         TabIndex        =   205
         Top             =   3240
         Width           =   1275
      End
      Begin VB.TextBox txtPrefabFolder 
         Height          =   315
         Left            =   750
         TabIndex        =   204
         Top             =   3255
         Width           =   5115
      End
      Begin VB.CommandButton cmdBrowsePrefab 
         Caption         =   "Browse..."
         Height          =   345
         Index           =   4
         Left            =   5940
         TabIndex        =   202
         Top             =   2100
         Width           =   1275
      End
      Begin VB.TextBox txtQuickPrefab 
         Height          =   315
         Index           =   4
         Left            =   2040
         TabIndex        =   201
         Top             =   2115
         Width           =   3825
      End
      Begin VB.CommandButton cmdBrowsePrefab 
         Caption         =   "Browse..."
         Height          =   345
         Index           =   3
         Left            =   5940
         TabIndex        =   199
         Top             =   1680
         Width           =   1275
      End
      Begin VB.TextBox txtQuickPrefab 
         Height          =   315
         Index           =   3
         Left            =   2040
         TabIndex        =   198
         Top             =   1695
         Width           =   3825
      End
      Begin VB.CommandButton cmdBrowsePrefab 
         Caption         =   "Browse..."
         Height          =   345
         Index           =   2
         Left            =   5940
         TabIndex        =   196
         Top             =   1260
         Width           =   1275
      End
      Begin VB.TextBox txtQuickPrefab 
         Height          =   315
         Index           =   2
         Left            =   2040
         TabIndex        =   195
         Top             =   1275
         Width           =   3825
      End
      Begin VB.CommandButton cmdBrowsePrefab 
         Caption         =   "Browse..."
         Height          =   345
         Index           =   1
         Left            =   5940
         TabIndex        =   193
         Top             =   840
         Width           =   1275
      End
      Begin VB.TextBox txtQuickPrefab 
         Height          =   315
         Index           =   1
         Left            =   2040
         TabIndex        =   192
         Top             =   855
         Width           =   3825
      End
      Begin VB.CommandButton cmdBrowsePrefab 
         Caption         =   "Browse..."
         Height          =   345
         Index           =   0
         Left            =   5940
         TabIndex        =   190
         Top             =   420
         Width           =   1275
      End
      Begin VB.TextBox txtQuickPrefab 
         Height          =   315
         Index           =   0
         Left            =   2040
         TabIndex        =   189
         Top             =   435
         Width           =   3825
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         Caption         =   "Default folder to look for Prefabs when inserting (empty for last used folder):"
         Height          =   210
         Left            =   750
         TabIndex        =   203
         Top             =   2955
         Width           =   5550
      End
      Begin VB.Label Label65 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Quick Prefab 5:"
         Height          =   210
         Left            =   750
         TabIndex        =   200
         Top             =   2160
         UseMnemonic     =   0   'False
         Width           =   1110
      End
      Begin VB.Label Label64 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Quick Prefab 4:"
         Height          =   210
         Left            =   750
         TabIndex        =   197
         Top             =   1740
         UseMnemonic     =   0   'False
         Width           =   1110
      End
      Begin VB.Label Label63 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Quick Prefab 3:"
         Height          =   210
         Left            =   750
         TabIndex        =   194
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   1110
      End
      Begin VB.Label Label62 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Quick Prefab 2:"
         Height          =   210
         Left            =   750
         TabIndex        =   191
         Top             =   900
         UseMnemonic     =   0   'False
         Width           =   1110
      End
      Begin VB.Label Label58 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Quick Prefab 1:"
         Height          =   210
         Left            =   750
         TabIndex        =   188
         Top             =   480
         UseMnemonic     =   0   'False
         Width           =   1110
      End
   End
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      Height          =   4695
      Index           =   4
      Left            =   165
      TabIndex        =   107
      Tag             =   "Nodebuilder"
      Top             =   840
      Visible         =   0   'False
      Width           =   7635
      Begin VB.ComboBox cmbNodeQuickload 
         Height          =   330
         Left            =   5145
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   109
         Top             =   120
         Width           =   2415
      End
      Begin VB.Frame fraExportBuild 
         Caption         =   " Export Nodebuild "
         Height          =   1785
         Left            =   135
         TabIndex        =   115
         Top             =   2790
         Width           =   7425
         Begin VB.CheckBox chkCompressSidedefs 
            Caption         =   "Compress sidedefs when exporting"
            Height          =   255
            Left            =   1950
            TabIndex        =   119
            Top             =   1350
            Width           =   3375
         End
         Begin VB.TextBox txtExportNodebuilderParams 
            Height          =   315
            Left            =   1950
            TabIndex        =   118
            Text            =   "-executethis - executethat"
            Top             =   750
            Width           =   4845
         End
         Begin VB.CommandButton cmdBrowseExportExecutable 
            Caption         =   "Browse..."
            Height          =   345
            Left            =   5550
            TabIndex        =   117
            Top             =   330
            Width           =   1245
         End
         Begin VB.TextBox txtExportNodebuilderExe 
            BackColor       =   &H80000000&
            Height          =   315
            Left            =   1950
            Locked          =   -1  'True
            TabIndex        =   116
            Text            =   "Nodebuilder.exe"
            Top             =   345
            Width           =   3525
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            Caption         =   "Use same placeholders in these parameters as explained above."
            Height          =   210
            Left            =   1965
            TabIndex        =   127
            Top             =   1080
            Width           =   4680
         End
         Begin VB.Label Label60 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Parameters:"
            Height          =   210
            Left            =   900
            TabIndex        =   126
            Top             =   795
            UseMnemonic     =   0   'False
            Width           =   870
         End
         Begin VB.Label Label59 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Executable:"
            Height          =   210
            Left            =   930
            TabIndex        =   125
            Top             =   390
            UseMnemonic     =   0   'False
            Width           =   840
         End
      End
      Begin VB.Frame fraQuickBuild 
         Caption         =   " Quick Nodebuild "
         Height          =   2205
         Left            =   135
         TabIndex        =   110
         Top             =   480
         Width           =   7425
         Begin VB.TextBox txtNodebuilderExe 
            BackColor       =   &H80000000&
            Height          =   315
            Left            =   1980
            Locked          =   -1  'True
            TabIndex        =   111
            Text            =   "Nodebuilder.exe"
            Top             =   345
            Width           =   3525
         End
         Begin VB.CommandButton cmdBrowseExecutable 
            Caption         =   "Browse..."
            Height          =   345
            Left            =   5580
            TabIndex        =   112
            Top             =   330
            Width           =   1245
         End
         Begin VB.TextBox txtNodebuilderParams 
            Height          =   315
            Left            =   1980
            TabIndex        =   113
            Text            =   "-executethis - executethat"
            Top             =   750
            Width           =   4845
         End
         Begin VB.ComboBox cmbBuildNodes 
            Height          =   330
            ItemData        =   "frmOptions.frx":07B8
            Left            =   1980
            List            =   "frmOptions.frx":07C5
            Style           =   2  'Dropdown List
            TabIndex        =   114
            Top             =   1635
            Width           =   4845
         End
         Begin VB.Label Label77 
            AutoSize        =   -1  'True
            Caption         =   "Use %T to indicate the file to which nodes must be written"
            Height          =   210
            Left            =   1995
            TabIndex        =   123
            Top             =   1290
            Width           =   4215
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Executable:"
            Height          =   210
            Left            =   960
            TabIndex        =   120
            Top             =   390
            UseMnemonic     =   0   'False
            Width           =   840
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Parameters:"
            Height          =   210
            Left            =   930
            TabIndex        =   121
            Top             =   795
            UseMnemonic     =   0   'False
            Width           =   870
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "Use %F to indicate the file from which to read the map from"
            Height          =   210
            Left            =   1995
            TabIndex        =   122
            Top             =   1080
            Width           =   4290
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "When to rebuild:"
            Height          =   210
            Left            =   630
            TabIndex        =   124
            Top             =   1680
            Width           =   1170
         End
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         Caption         =   "Select here a profile to quickly set the Quick Nodebuild parameters:"
         Height          =   210
         Left            =   135
         TabIndex        =   108
         Top             =   180
         Width           =   4830
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   6390
      TabIndex        =   232
      Top             =   5775
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4725
      TabIndex        =   231
      Top             =   5775
      Width           =   1575
   End
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      Height          =   4785
      Index           =   9
      Left            =   165
      TabIndex        =   206
      Tag             =   "Interface"
      Top             =   780
      Visible         =   0   'False
      Width           =   7635
      Begin VB.CheckBox chkAutoScroll 
         Caption         =   "Automatic scrolling during operations"
         Height          =   225
         Left            =   570
         TabIndex        =   247
         Top             =   420
         Width           =   3375
      End
      Begin VB.CheckBox chkAlwaysAllTextures 
         Caption         =   "Always browse all textures and flats"
         Height          =   225
         Left            =   570
         TabIndex        =   246
         Top             =   3180
         Width           =   3105
      End
      Begin VB.CheckBox chkAutoCompleteTypedTex 
         Caption         =   "Autocomplete texture names while typing"
         Height          =   225
         Left            =   570
         TabIndex        =   240
         Top             =   3525
         Width           =   3330
      End
      Begin VB.CheckBox chkAutoCompleteTex 
         Caption         =   "Autocomplete texture names after typing"
         Height          =   225
         Left            =   570
         TabIndex        =   219
         Top             =   3870
         Width           =   3330
      End
      Begin VB.CheckBox chkAllThingsRects 
         Caption         =   "Outline all Things in Things mode"
         Height          =   225
         Left            =   570
         TabIndex        =   218
         Top             =   2835
         Width           =   3105
      End
      Begin VB.CheckBox chkIndicatorScaled 
         Caption         =   "Scale linedef indicator with Zoom"
         Height          =   225
         Left            =   570
         TabIndex        =   215
         Top             =   1800
         Width           =   3105
      End
      Begin VB.CheckBox chkNothingDeselects 
         Caption         =   "Selecting nothing will deselect all"
         Height          =   225
         Left            =   570
         TabIndex        =   216
         Top             =   2145
         Width           =   3105
      End
      Begin VB.CheckBox chkAdditiveSelect 
         Caption         =   "Additive select by default"
         Height          =   225
         Left            =   570
         TabIndex        =   217
         Top             =   2490
         Width           =   3105
      End
      Begin VB.ComboBox cmbThingRects 
         Height          =   330
         ItemData        =   "frmOptions.frx":080D
         Left            =   5985
         List            =   "frmOptions.frx":081A
         Style           =   2  'Dropdown List
         TabIndex        =   225
         Top             =   2580
         Width           =   1230
      End
      Begin VB.ComboBox cmbVertexSize 
         Height          =   330
         ItemData        =   "frmOptions.frx":0834
         Left            =   6285
         List            =   "frmOptions.frx":0841
         Style           =   2  'Dropdown List
         TabIndex        =   224
         Top             =   4455
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CheckBox chkLinesTree 
         Caption         =   "Show Linedefs listing in categorized tree"
         Height          =   225
         Left            =   570
         TabIndex        =   214
         Top             =   1455
         Width           =   3375
      End
      Begin VB.CheckBox chkThingsTree 
         Caption         =   "Show Things listing in categorized tree"
         Height          =   225
         Left            =   570
         TabIndex        =   213
         Top             =   1110
         Width           =   3375
      End
      Begin VB.CheckBox chkTooltips 
         Caption         =   "Show Tooltip for highlighted object"
         Height          =   225
         Left            =   120
         TabIndex        =   212
         Top             =   4620
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.CheckBox chkZoomMouse 
         Caption         =   "Zoom to or from mouse location"
         Height          =   225
         Left            =   150
         TabIndex        =   210
         Top             =   4500
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.ComboBox cmbDetailsBar 
         Height          =   330
         ItemData        =   "frmOptions.frx":085B
         Left            =   5970
         List            =   "frmOptions.frx":086E
         Style           =   2  'Dropdown List
         TabIndex        =   220
         Top             =   360
         Width           =   1230
      End
      Begin VB.CheckBox chkMode1Vertices 
         Caption         =   "Show Vertices in Lines mode"
         Height          =   225
         Left            =   3225
         TabIndex        =   207
         Top             =   4455
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.CheckBox chkMode2Vertices 
         Caption         =   "Show Vertices in Sectors mode"
         Height          =   225
         Left            =   1650
         TabIndex        =   208
         Top             =   4440
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.CheckBox chkModeAllThings 
         Caption         =   "Show Things in all modes"
         Height          =   225
         Left            =   3045
         TabIndex        =   209
         Top             =   4635
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.CheckBox chkHideToolbar 
         Caption         =   "Show Toolbar"
         Height          =   225
         Left            =   570
         TabIndex        =   211
         Top             =   765
         Width           =   3375
      End
      Begin DoomBuilder.ctlValueBox valScrollPixels 
         Height          =   360
         Left            =   5970
         TabIndex        =   221
         ToolTipText     =   "Screen Pixels"
         Top             =   885
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         Max             =   9990
         MaxLength       =   4
         Min             =   10
         SmallChange     =   10
         Value           =   "10"
      End
      Begin DoomBuilder.ctlValueBox valZoomSpeed 
         Height          =   360
         Left            =   5970
         TabIndex        =   222
         Top             =   1455
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         Max             =   9990
         MaxLength       =   4
         Min             =   10
         SmallChange     =   10
         Value           =   "10"
      End
      Begin DoomBuilder.ctlValueBox valIndicatorSize 
         Height          =   360
         Left            =   5985
         TabIndex        =   223
         ToolTipText     =   "Map or Screen Pixels depending on option ""Scale indicator size with Zoom"""
         Top             =   2010
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         Max             =   9990
         MaxLength       =   4
      End
      Begin VB.Label Label67 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Details Bar location:"
         Height          =   210
         Left            =   4440
         TabIndex        =   248
         Top             =   420
         Width           =   1425
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Linedef Indicator Size:"
         Height          =   210
         Left            =   4260
         TabIndex        =   228
         Top             =   2070
         UseMnemonic     =   0   'False
         Width           =   1605
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Thing size indicator:"
         Height          =   210
         Left            =   4425
         TabIndex        =   230
         Top             =   2640
         UseMnemonic     =   0   'False
         Width           =   1440
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         Caption         =   "Overall Vertex Size:"
         Height          =   210
         Left            =   4740
         TabIndex        =   229
         Top             =   4500
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Scroll speed:"
         Height          =   210
         Left            =   4920
         TabIndex        =   226
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   945
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Zoom speed:"
         Height          =   210
         Left            =   4920
         TabIndex        =   227
         Top             =   1515
         UseMnemonic     =   0   'False
         Width           =   945
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   5445
      Left            =   120
      TabIndex        =   233
      Top             =   165
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   9604
      MultiRow        =   -1  'True
      ShowTips        =   0   'False
      TabFixedWidth   =   2170
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   10
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  Editing  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  Colors  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  Files  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  Defaults  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  Nodebuilder  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  Testing  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  Shortcut Keys  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  3D Mode  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  Prefabs  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  Interface  "
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
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



Private EnumerateModes As Boolean
Private ChangeMode As Boolean
Private ModeWidth As Long
Private ModeHeight As Long
Private ModeFormat As Long
Private ModeRate As Long

Private NodebuilderProfiles As Dictionary
Private TestingProfiles As Dictionary

Private Sub chkWindowed_Click()
     
     'Enable/disable controls
     cmbVideoDriver.Enabled = (chkWindowed.Value = vbUnchecked)
     cmbResolution.Enabled = (chkWindowed.Value = vbUnchecked)
     lblVideoDriver.Enabled = (chkWindowed.Value = vbUnchecked)
     lblResolution.Enabled = (chkWindowed.Value = vbUnchecked)
     chkStandardTextureBrowse.Enabled = (chkWindowed.Value = vbChecked)
     chkLinesSectorsInfo.Enabled = (chkWindowed.Value = vbChecked)
End Sub

Private Sub cmbNodeQuickload_Change()
     Dim Settings As Dictionary
     Dim Index As Long
     
     'Leave when nothing selected
     If (cmbNodeQuickload.ListIndex < 0) Then Exit Sub
     
     'Get Index
     Index = cmbNodeQuickload.ItemData(cmbNodeQuickload.ListIndex)
     
     'Get settings
     Set Settings = NodebuilderProfiles.Items(Index)
     
     'Apply settings
     txtNodebuilderExe.Text = Settings("executable")
     txtNodebuilderParams.Text = Settings("parameters")
End Sub

Private Sub cmbNodeQuickload_Click()
     
     'Same as changing
     cmbNodeQuickload_Change
End Sub


Private Sub cmbNodeQuickload_KeyDown(KeyCode As Integer, Shift As Integer)
     
     'Same as changing
     cmbNodeQuickload_Change
End Sub


Private Sub cmbNodeQuickload_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Same as changing
     cmbNodeQuickload_Change
End Sub


Private Sub cmbResolution_Change()
     Dim ModeInfo As D3DDISPLAYMODE
     
     'Check if we may change mode
     If ChangeMode Then
          
          'Set the current mode
          D3D.EnumAdapterModes cmbVideoDriver.ListIndex, D3DFORMAT.D3DFMT_X8R8G8B8, cmbResolution.ItemData(cmbResolution.ListIndex), ModeInfo
          ModeWidth = ModeInfo.width
          ModeHeight = ModeInfo.height
          ModeFormat = ModeInfo.Format
          ModeRate = ModeInfo.RefreshRate
     End If
End Sub

Private Sub cmbResolution_Click()
     cmbResolution_Change
End Sub

Private Sub cmbResolution_KeyUp(KeyCode As Integer, Shift As Integer)
     cmbResolution_Change
End Sub

Private Sub cmbShortcut_Click()
     
     'Check if anything selected
     If (cmbShortcut.ListIndex > -1) Then
          
          'Set the combination on the tag of selected item
          lstFunctions.SelectedItem.tag = cmbShortcut.ItemData(cmbShortcut.ListIndex)
          
          'Show the combination name
          txtShortcut.Text = NameForKeycode((cmbShortcut.ItemData(cmbShortcut.ListIndex) And &HFFF), (cmbShortcut.ItemData(cmbShortcut.ListIndex) And &HFF0000) \ 2 ^ 16)
          
          lstFunctions.SelectedItem.ListSubItems(1) = txtShortcut.Text
     End If
     
     On Local Error Resume Next
     txtShortcut.SetFocus
End Sub

Private Sub cmbTestQuickload_Change()
     Dim Settings As Dictionary
     Dim Index As Long
     
     'Leave when nothing selected
     If (cmbTestQuickload.ListIndex < 0) Then Exit Sub
     
     'Get Index
     Index = cmbTestQuickload.ItemData(cmbTestQuickload.ListIndex)
     
     'Get settings
     Set Settings = TestingProfiles.Items(Index)
     
     'Apply settings
     txtTestExe.Text = Settings("executable")
     txtTestParams.Text = Settings("parameters")
End Sub

Private Sub cmbTestQuickload_Click()
     
     'Same as changing
     cmbTestQuickload_Change
End Sub


Private Sub cmbTestQuickload_KeyDown(KeyCode As Integer, Shift As Integer)
     
     'Same as changing
     cmbTestQuickload_Change
End Sub

Private Sub cmbTestQuickload_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Same as changing
     cmbTestQuickload_Change
End Sub


Private Sub cmbVideoDriver_Change()
     Dim CurrentAdapter As Long
     Dim ModesCount As Long
     Dim ModeInfo As D3DDISPLAYMODE
     Dim i As Long
     
     'Check if we're allowed to enumerate modes
     If EnumerateModes Then
          
          'Dont change mode
          ChangeMode = False
          
          'Get the current adapter
          CurrentAdapter = cmbVideoDriver.ListIndex
          
          'Clear modes combo
          cmbResolution.Clear
          
          'Fill with all modes that are allowed
          ModesCount = D3D.GetAdapterModeCount(CurrentAdapter, D3DFORMAT.D3DFMT_X8R8G8B8)
          For i = 0 To ModesCount - 1
               
               'Get the mode info
               D3D.EnumAdapterModes CurrentAdapter, D3DFORMAT.D3DFMT_X8R8G8B8, i, ModeInfo
               
               'Check if we allow this resolution
               If (ModeInfo.width >= 320) And _
                  (ModeInfo.height >= 200) Then
                    
                    'Add to combo
                    If ModeInfo.RefreshRate Then
                         
                         'With refreshrate information
                         cmbResolution.AddItem ModeInfo.width & " x " & ModeInfo.height & " x " & BitsFromFormat(ModeInfo.Format) & " @ " & ModeInfo.RefreshRate & " Hz"
                         cmbResolution.ItemData(cmbResolution.ListCount - 1) = i
                    Else
                         
                         'Without refresh rate information
                         cmbResolution.AddItem ModeInfo.width & " x " & ModeInfo.height & " x " & BitsFromFormat(ModeInfo.Format)
                         cmbResolution.ItemData(cmbResolution.ListCount - 1) = i
                    End If
                    
                    'Check if we should select it
                    If (cmbResolution.ListIndex = -1) And _
                       (ModeWidth = ModeInfo.width) And _
                       (ModeHeight = ModeInfo.height) And _
                       (ModeFormat = ModeInfo.Format) And _
                       ((ModeRate = ModeInfo.RefreshRate) Or (ModeRate = 0)) Then _
                         cmbResolution.ListIndex = cmbResolution.ListCount - 1
               End If
          Next i
          
          'Allow mode changes
          ChangeMode = True
     End If
End Sub

Private Sub cmbVideoDriver_Click()
     cmbVideoDriver_Change
End Sub

Private Sub cmbVideoDriver_KeyUp(KeyCode As Integer, Shift As Integer)
     cmbVideoDriver_Change
End Sub

Private Sub cmdBrowseEngine_Click()
     Dim NewFile As String
     
     'Browse for new file
     NewFile = OpenFile(fraOptions(5).hWnd, "Select Engine Executable", "Executable Files   *.exe|*.exe", txtTestExe.Text, cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames)
     
     'Check if not cancelled
     If (Trim$(NewFile) <> "") Then
          
          'No settings from profile
          cmbTestQuickload.ListIndex = -1
          
          'Set the new file in textbox
          txtTestExe.Text = NewFile
          txtTestExe.SelStart = Len(txtTestExe.Text)
          txtTestExe.SetFocus
     End If
End Sub

Private Sub cmdBrowseExecutable_Click()
     Dim NewFile As String
     
     'Change to local path
     ChDrive left$(App.Path, 1)
     ChDir App.Path
     
     'Browse for new file
     NewFile = OpenFile(fraOptions(4).hWnd, "Select Nodebuilder Executable", "Executable Files   *.exe|*.exe", txtNodebuilderExe.Text, cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames)
     
     'Check if not cancelled
     If (Trim$(NewFile) <> "") Then
          
          'No settings from profile
          cmbNodeQuickload.ListIndex = -1
          
          'Set the new file in textbox
          txtNodebuilderExe.Text = NewFile
          txtNodebuilderExe.SelStart = Len(txtNodebuilderExe.Text)
          txtNodebuilderExe.SetFocus
     End If
End Sub

Private Sub cmdBrowseExportExecutable_Click()
     Dim NewFile As String
     
     'Change to local path
     ChDrive left$(App.Path, 1)
     ChDir App.Path
     
     'Browse for new file
     NewFile = OpenFile(fraOptions(4).hWnd, "Select Nodebuilder Executable", "Executable Files   *.exe|*.exe", txtExportNodebuilderExe.Text, cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames)
     
     'Check if not cancelled
     If (Trim$(NewFile) <> "") Then
          
          'Set the new file in textbox
          txtExportNodebuilderExe.Text = NewFile
          txtExportNodebuilderExe.SelStart = Len(txtExportNodebuilderExe.Text)
          txtExportNodebuilderExe.SetFocus
     End If
End Sub

Private Sub cmdBrowsePrefab_Click(Index As Integer)
     Dim NewFile As String
     
     'Browse for new file
     NewFile = OpenFile(fraOptions(8).hWnd, "Select Prefab file", "Doom Builder Prefab Files   *.dbp|*.dbp", txtQuickPrefab(Index).Text, cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames)
     
     'Check if not cancelled
     If (Trim$(NewFile) <> "") Then
          
          'Set the new file in textbox
          txtQuickPrefab(Index).Text = NewFile
          txtQuickPrefab(Index).SelStart = Len(txtQuickPrefab(Index).Text)
          txtQuickPrefab(Index).SetFocus
     End If
End Sub

Private Sub cmdCancel_Click()
     OptionsCancelled = True
     
     'Terminate DirectX
     TerminateDirectX
     
     'Leave
     Unload Me
     Set frmOptions = Nothing
End Sub

Private Sub cmdColor_Click(Index As Integer)
     Dim NewColor As Long
     
     'Select new color
     NewColor = SelectColor(fraOptions(1).hWnd, cmdColor(Index).BackColor, cdlCCFullOpen Or cdlCCRGBInit, CustomColors())
     
     'Check if not cancelled
     If (NewColor <> -1) Then
          
          'Set the new color on the button
          cmdColor(Index).BackColor = NewColor
     End If
End Sub

Private Sub cmdGameIWAD_Click()
     Dim NewFile As String
     
     'Browse for new file
     NewFile = OpenFile(fraOptions(2).hWnd, "Select an IWAD", "IWAD Files   *.wad|*.wad", "", cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames)
     
     'Check if not cancelled
     If (Trim$(NewFile) <> "") Then
          
          'Set the new file in textbox
          txtGameIWAD.Text = NewFile
          lstGames.SelectedItem.ListSubItems(1).Text = NewFile
     End If
End Sub

Private Sub cmdGameIWADClear_Click()
     
     'Clear the file
     txtGameIWAD.Text = ""
     lstGames.SelectedItem.ListSubItems(1).Text = ""
End Sub


Private Sub cmdOK_Click()
     Dim i As Long
     Dim ColorRGB As BITMAPRGB
     OptionsCancelled = False
     
     'This will validate the last focused control
     DoEvents
     
     '=== Editing & Interface
     Config("indicatorsize") = Val(valIndicatorSize.Value)
     Config("scrollpixels") = Val(valScrollPixels.Value)
     Config("zoomspeed") = Val(valZoomSpeed.Value)
     Config("maxundos") = Val(valMaxUndos.Value)
     Config("thingrects") = cmbThingRects.ListIndex
     Config("vertexselectdistance") = Val(valVertexSelectDistance.Value)
     Config("lineselectdistance") = Val(valLinedefSelectDistance.Value)
     Config("thingselectdistance") = Val(valThingSelectDistance.Value)
     Config("autostitchdistance") = Val(valAutostitchDistance.Value)
     Config("linesplitdistance") = Val(valLinesplitDistance.Value)
     Config("indicatorscaled") = chkIndicatorScaled.Value
     Config("texturecaching") = chkTextureCaching.Value
     Config("textureprecache") = chkTexturePrecache.Value
     Config("mode1vertices") = chkMode1Vertices.Value
     Config("mode2vertices") = chkMode2Vertices.Value
     Config("savebackup") = chkSaveBackup.Value
     Config("nothingdeselects") = chkNothingDeselects.Value
     Config("additiveselect") = chkAdditiveSelect.Value
     Config("subundos") = chkSubUndos.Value
     Config("autoscroll") = chkAutoScroll.Value
     Config("showtoolbar") = chkHideToolbar.Value
     Config("modethings") = chkModeAllThings.Value
     Config("detailsbar") = cmbDetailsBar.ListIndex
     Config("zoommouse") = chkZoomMouse.Value
     Config("showtooltips") = chkTooltips.Value
     Config("vertexsize") = cmbVertexSize.ListIndex
     Config("thingstree") = chkThingsTree.Value
     Config("linestree") = chkLinesTree.Value
     Config("mixresources") = chkMixResources.Value
     Config("allthingsrects") = chkAllThingsRects.Value
     Config("newlinesdialog") = chkNewLinesDailog.Value
     Config("newsectordialog") = chkNewSectorDialog.Value
     Config("newthingdialog") = chkNewThingsDialog.Value
     Config("copytagdraw") = chkCopyTagsDraw.Value
     Config("copytagpaste") = chkCopyTagsPaste.Value
     Config("storeeditinginfo") = chkStoreEditingInfo.Value
     Config("autocompletetex") = chkAutoCompleteTex.Value
     Config("autocompletetypetex") = chkAutoCompleteTypedTex.Value
     Config("pasteadjustsheights") = chkPasteAdjustsHeights.Value
     Config("alwaysalltextures") = chkAlwaysAllTextures.Value
     
     
     '=== Colors
     For i = cmdColor.LBound To cmdColor.UBound
          
          'Get the RGB from button
          ColorRGB = WinLongToBITMAPRGB(cmdColor(i).BackColor)
          
          'Set the color in configuration
          Config("palette")(cmdColor(i).tag) = BITMAPRGBToLong(ColorRGB)
     Next i
     Config("syntaxhighlighting") = chkHighlighSyntax.Value
     
     
     '=== Files
     For i = 1 To lstGames.ListItems.Count
          
          'Set the iwad for the configation file
          Config("iwads")(lstGames.ListItems(i).ListSubItems("FILENAME").Text) = lstGames.ListItems(i).ListSubItems("IWAD").Text
     Next i
     
     
     '=== Defaults
     Config("defaultgrid") = Val(valDefaultGrid.Value)
     Config("defaultsnap") = chkDefaultSnap.Value
     Config("defaultstitch") = chkDefaultStitch.Value
     Config("defaulttexture")("upper") = txtDefaultUpper.Text
     Config("defaulttexture")("middle") = txtDefaultMiddle.Text
     Config("defaulttexture")("lower") = txtDefaultLower.Text
     Config("defaultsector")("tceiling") = txtDefaultTCeiling.Text
     Config("defaultsector")("tfloor") = txtDefaultTFloor.Text
     Config("defaultsector")("hceiling") = Val(valDefaultHCeiling.Value)
     Config("defaultsector")("hfloor") = Val(valDefaultHFloor.Value)
     Config("defaultsector")("brightness") = Val(valDefaultBrightness.Value)
     Config("defaultthing") = Val(valDefaultThing.Value)
     
     'Ensure valid textures are used to build with
     CorrectDefaultTextures
     
     
     '=== Nodebuilder
     Config("buildexec") = txtNodebuilderExe.Text
     Config("buildparams") = txtNodebuilderParams.Text
     Config("buildnodes") = cmbBuildNodes.ItemData(cmbBuildNodes.ListIndex)
     Config("buildexportexec") = txtExportNodebuilderExe.Text
     Config("buildexportparams") = txtExportNodebuilderParams.Text
     Config("buildexportcompression") = chkCompressSidedefs.Value
     
     
     '=== Testing
     Config("testexec") = txtTestExe.Text
     Config("testparams") = txtTestParams.Text
     Config("testdialog") = chkTestingDialog.Value
     
     
     '=== Shortcuts
     For i = 1 To lstFunctions.ListItems.Count
          
          'Set the item
          Config("shortcuts")(lstFunctions.ListItems(i).Key) = Val(lstFunctions.ListItems(i).tag)
     Next i
     
     'Shortcut keys options
     Config("modekeys3d") = chkModeKeys3D.Value
     
     
     '=== 3D Mode
     If (cmbVideoDriver.ListIndex > -1) And (cmbResolution.ListIndex > -1) Then
          Config("videoadapter") = cmbVideoDriver.ListIndex
          Config("videoadapterdesc") = cmbVideoDriver.Text
          Config("videowidth") = ModeWidth
          Config("videoheight") = ModeHeight
          Config("videoformat") = ModeFormat
          Config("videorate") = ModeRate
     Else
          Config("videoadapter") = 0
          Config("videoadapterdesc") = ""
     End If
     Config("videofov") = Val(txtFOV.Value)
     Config("texturefilter") = cmbTextureFilter.ListIndex
     Config("movespeed") = Val(txtMoveSpeed.Value)
     Config("mousespeed") = Val(txtMouseSpeed.Value)
     Config("showfog") = chkFog.Value
     Config("invertmousey") = chkInvertY.Value
     Config("videoaspect") = chkAspect.Value
     Config("directxprecache") = chkDirectXPrecache.Value
     Config("vertexbuffercache") = chkVertexbufferCache.Value
     Config("belowceiling") = chkBelowCeiling.Value
     Config("raiselowerceiling") = chkRaiseLowerCeiling.Value
     Config("videogamma") = Val(txtGamma.Value)
     Config("videobrightness") = Val(txtBrightness.Value)
     Config("videoviewdistance") = Val(txtVideoDistance.Value)
     Config("exclusivemouse") = chkExclusivemouse.Value
     Config("windowedvideo") = chkWindowed.Value
     Config("standardtexturebrowse") = chkStandardTextureBrowse.Value
     Config("linessectorsinfo") = chkLinesSectorsInfo.Value
     
     
     '=== Prefabs
     Config("quickprefab1") = txtQuickPrefab(0).Text
     Config("quickprefab2") = txtQuickPrefab(1).Text
     Config("quickprefab3") = txtQuickPrefab(2).Text
     Config("quickprefab4") = txtQuickPrefab(3).Text
     Config("quickprefab5") = txtQuickPrefab(4).Text
     If (right$(txtPrefabFolder.Text, 1) = "\") Then
          Config("prefabfolder") = txtPrefabFolder.Text
     Else
          Config("prefabfolder") = txtPrefabFolder.Text & "\"
     End If
     
     
     'Terminate DirectX
     TerminateDirectX
     
     'Leave
     Unload Me
     Set frmOptions = Nothing
End Sub

Private Sub cmdPrefabFolder_Click()
     Dim NewFolder As String
     
     'Browse for new folder
     NewFolder = SelectFolder(fraOptions(8).hWnd, "Select default Prefabs folder")
     
     'Check if not cancelled
     If (Trim$(NewFolder) <> "") Then
          
          'Set the new folder in textbox
          txtPrefabFolder.Text = NewFolder
          txtPrefabFolder.SelStart = Len(txtPrefabFolder.Text)
          txtPrefabFolder.SetFocus
     End If
End Sub

Private Sub cmdUnbind_Click()
     
     'Show the combination name
     txtShortcut.Text = ""
     
     'Set the combination on the tag of selected item
     lstFunctions.SelectedItem.tag = 0
     lstFunctions.SelectedItem.ListSubItems(1) = NameForKeycode(0, 0)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     
     'Adjust shift mask
     CurrentShiftMask = Shift
     
     'Check what key is pressed
     If (KeyCode = vbKeyTab) And (Shift = vbCtrlMask) Then
          
          'Switch to next panel
          If (tbsOptions.SelectedItem.Index = tbsOptions.Tabs.Count) Then
               tbsOptions.Tabs(1).selected = True
          Else
               tbsOptions.Tabs(tbsOptions.SelectedItem.Index + 1).selected = True
          End If
          
          'Focus to panel
          tbsOptions.SetFocus
     End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Adjust shift mask
     CurrentShiftMask = Shift
End Sub


Private Sub Form_Load()
     Dim i As Long
     Dim k As Long, s As Long
     Dim ColorRGB As BITMAPRGB
     Dim NewItem As ListItem
     Dim Shortcuts As New clsConfiguration
     Dim AllShortcuts As Dictionary
     Dim Parameters As New clsConfiguration
     Dim Key As String
     
     'This might give problems with new settings that dont exist in
     'the config yet, so ignore these now
     On Local Error Resume Next
     
     'Show hourglass mousepointer
     Screen.MousePointer = vbHourglass
     
     
     '=== Editing & Interface
     valIndicatorSize.Value = Config("indicatorsize")
     valScrollPixels.Value = Config("scrollpixels")
     valZoomSpeed.Value = Config("zoomspeed")
     valMaxUndos.Value = Config("maxundos")
     cmbThingRects.ListIndex = Config("thingrects")
     valVertexSelectDistance.Value = Config("vertexselectdistance")
     valLinedefSelectDistance.Value = Config("lineselectdistance")
     valThingSelectDistance.Value = Config("thingselectdistance")
     valAutostitchDistance.Value = Config("autostitchdistance")
     valLinesplitDistance.Value = Config("linesplitdistance")
     chkIndicatorScaled.Value = Config("indicatorscaled")
     chkTextureCaching.Value = Config("texturecaching")
     chkTexturePrecache.Value = Config("textureprecache")
     chkMode1Vertices.Value = Config("mode1vertices")
     chkMode2Vertices.Value = Config("mode2vertices")
     chkSaveBackup.Value = Config("savebackup")
     chkNothingDeselects.Value = Config("nothingdeselects")
     chkAdditiveSelect.Value = Config("additiveselect")
     chkSubUndos.Value = Config("subundos")
     chkAutoScroll.Value = Config("autoscroll")
     chkHideToolbar.Value = Config("showtoolbar")
     chkModeAllThings.Value = Config("modethings")
     cmbDetailsBar.ListIndex = Config("detailsbar")
     chkZoomMouse.Value = Config("zoommouse")
     chkTooltips.Value = Config("showtooltips")
     cmbVertexSize.ListIndex = Config("vertexsize")
     txtVideoDistance.Value = Config("videoviewdistance")
     chkThingsTree.Value = Config("thingstree")
     chkLinesTree.Value = Config("linestree")
     chkMixResources.Value = Config("mixresources")
     chkAllThingsRects.Value = Config("allthingsrects")
     chkNewLinesDailog.Value = Config("newlinesdialog")
     chkNewSectorDialog.Value = Config("newsectordialog")
     chkNewThingsDialog.Value = Config("newthingdialog")
     chkCopyTagsDraw.Value = Config("copytagdraw")
     chkCopyTagsPaste.Value = Config("copytagpaste")
     chkStoreEditingInfo.Value = Config("storeeditinginfo")
     chkAutoCompleteTypedTex.Value = Config("autocompletetypetex")
     chkPasteAdjustsHeights.Value = Config("pasteadjustsheights")
     chkAlwaysAllTextures.Value = Config("alwaysalltextures")
     
     
     '=== Colors
     For i = cmdColor.LBound To cmdColor.UBound
          
          'Get the RGB from configuration color
          ColorRGB = LongToBITMAPRGB(Config("palette")(cmdColor(i).tag))
          
          'Set the color on the button
          cmdColor(i).BackColor = RGB(ColorRGB.rgbRed, ColorRGB.rgbGreen, ColorRGB.rgbBlue)
     Next i
     chkHighlighSyntax.Value = Config("syntaxhighlighting")
     
     
     '=== Files
     
     'Go for al configs
     For i = 0 To (AllGameConfigs.Count - 1)
          
          'Add the game configuration to list
          Set NewItem = lstGames.ListItems.Add(, , AllGameConfigs.Keys(i))
          
          'Add the IWAD file to it
          NewItem.ListSubItems.Add , "IWAD", Config("iwads")(LCase$(Dir(AllGameConfigs(AllGameConfigs.Keys(i)))))
          
          'Add filename to it
          NewItem.ListSubItems.Add , "FILENAME", LCase$(Dir(AllGameConfigs(AllGameConfigs.Keys(i))))
          
          'Check if we should select this one
          If StrComp(AllGameConfigs.Keys(i), mapgame, vbTextCompare) = 0 Then NewItem.selected = True
     Next i
     
     'Select a default one
     If (lstGames.SelectedItem Is Nothing) Then lstGames.ListItems(1).selected = True
     lstGames_ItemClick lstGames.SelectedItem
     
     
     '=== Defaults
     valDefaultGrid.Value = Config("defaultgrid")
     chkDefaultSnap.Value = Config("defaultsnap")
     chkDefaultStitch.Value = Config("defaultstitch")
     txtDefaultUpper.Text = Config("defaulttexture")("upper")
     txtDefaultMiddle.Text = Config("defaulttexture")("middle")
     txtDefaultLower.Text = Config("defaulttexture")("lower")
     txtDefaultTCeiling.Text = Config("defaultsector")("tceiling")
     txtDefaultTFloor.Text = Config("defaultsector")("tfloor")
     valDefaultHCeiling.Value = Config("defaultsector")("hceiling")
     valDefaultHFloor.Value = Config("defaultsector")("hfloor")
     valDefaultBrightness.Value = Config("defaultsector")("brightness")
     valDefaultThing.Value = Config("defaultthing")
     
     '=== Nodebuilder
     txtNodebuilderExe.Text = Config("buildexec")
     txtNodebuilderParams.Text = Config("buildparams")
     txtExportNodebuilderExe.Text = Config("buildexportexec")
     txtExportNodebuilderParams.Text = Config("buildexportparams")
     chkCompressSidedefs.Value = Config("buildexportcompression")
     
     'Go for all nodebuild options
     For i = 0 To (cmbBuildNodes.ListCount - 1)
          If (cmbBuildNodes.ItemData(i) = Config("buildnodes")) Then cmbBuildNodes.ListIndex = i
     Next i
     
     '=== Testing
     txtTestExe.Text = Config("testexec")
     txtTestParams.Text = Config("testparams")
     chkTestingDialog.Value = Config("testdialog")
     
     
     '=== Shortcuts
     Shortcuts.LoadConfiguration App.Path & "\Shortcuts.cfg"
     Set AllShortcuts = Shortcuts.ReadSetting("shortcuts", New Dictionary, True)
     
     'Go for all items
     If (Config.Exists("shortcuts") = False) Then Config.Add "shortcuts", New Dictionary
     For i = 0 To (AllShortcuts.Count - 1)
          
          'Shortcut action key
          Key = AllShortcuts.Keys(i)
          
          'Add the list item
          Set NewItem = lstFunctions.ListItems.Add(, Key, AllShortcuts.Items(i)("title"))
          
          'Check if item exists in config
          If (Config("shortcuts").Exists(Key) = False) Then
               
               'Make the item
               Config("shortcuts").Add Key, 0
          End If
          
          'Set the value on the tag
          NewItem.tag = Val(Config("shortcuts")(Key))
          
          'Split keycode and shift
          k = (Val(Config("shortcuts")(Key)) And &HFFF)
          s = (Val(Config("shortcuts")(Key)) And &HFF0000) \ 2 ^ 16
          
          'Set the key string on the item
          NewItem.ListSubItems.Add , , NameForKeycode(k, s)
          
          'Add other properties
          With NewItem.ListSubItems
               .Add , "DESC", AllShortcuts.Items(i)("description")
               .Add , "UNBIND", AllShortcuts.Items(i)("unbind")
               .Add , "MOUSEBUTTONS", AllShortcuts.Items(i)("mousebuttons")
               .Add , "MOUSESCROLL", AllShortcuts.Items(i)("mousescroll")
          End With
     Next i
     
     'Select the first
     lstFunctions.ListItems(1).selected = True
     lstFunctions_ItemClick lstFunctions.SelectedItem
     
     'Shortcut keys options
     chkModeKeys3D.Value = Val(Config("modekeys3d"))
     
     
     '=== Prefabs
     txtQuickPrefab(0).Text = Config("quickprefab1")
     txtQuickPrefab(1).Text = Config("quickprefab2")
     txtQuickPrefab(2).Text = Config("quickprefab3")
     txtQuickPrefab(3).Text = Config("quickprefab4")
     txtQuickPrefab(4).Text = Config("quickprefab5")
     txtPrefabFolder.Text = Config("prefabfolder")
     
     
     'Initialize DirectX
     If (InitDirectX = True) Then
          
          'Setup 3D Mode panel
          Setup3DPanel
     Else
          
          'No 3D available
          tbsOptions.Tabs(7).tag = "NO"
          fraNo3D.visible = True
          fraNo3D.ZOrder 0
          fraNo3D.Move 0, 0
     End If
     
     'Other 3D Mode stuff
     txtFOV.Value = Val(Config("videofov"))
     cmbTextureFilter.ListIndex = Val(Config("texturefilter"))
     txtMoveSpeed.Value = Config("movespeed")
     txtMouseSpeed.Value = Config("mousespeed")
     chkFog.Value = Val(Config("showfog"))
     chkInvertY.Value = Val(Config("invertmousey"))
     chkAspect.Value = Config("videoaspect")
     chkAutoCompleteTex.Value = Config("autocompletetex")
     chkDirectXPrecache.Value = Config("directxprecache")
     chkVertexbufferCache.Value = Config("vertexbuffercache")
     chkBelowCeiling.Value = Config("belowceiling")
     chkRaiseLowerCeiling.Value = Config("raiselowerceiling")
     txtGamma.Value = Config("videogamma")
     txtBrightness.Value = Config("videobrightness")
     chkExclusivemouse.Value = Val(Config("exclusivemouse"))
     chkWindowed.Value = Val(Config("windowedvideo"))
     chkStandardTextureBrowse.Value = Val(Config("standardtexturebrowse"))
     chkLinesSectorsInfo.Value = Val(Config("linessectorsinfo"))
     
     
     'Load parameters
     Parameters.LoadConfiguration App.Path & "\Parameters.cfg"
     
     'Get Nodebuilder Profiles
     Set NodebuilderProfiles = Parameters.ReadSetting("nodebuilders", New Dictionary, True)
     
     'Fill Nodebuilder Profile boxes
     For i = 0 To NodebuilderProfiles.Count - 1
          
          'Get key
          Key = NodebuilderProfiles.Keys(i)
          
          'Add to combo
          cmbNodeQuickload.AddItem NodebuilderProfiles(Key)("title")
          cmbNodeQuickload.ItemData(cmbNodeQuickload.NewIndex) = i
     Next i
     
     'Get Testing Profiles
     Set TestingProfiles = Parameters.ReadSetting("engines", New Dictionary, True)
     
     'Fill Testing Profile boxes
     For i = 0 To TestingProfiles.Count - 1
          
          'Get key
          Key = TestingProfiles.Keys(i)
          
          'Add to combo
          cmbTestQuickload.AddItem TestingProfiles(Key)("title")
          cmbTestQuickload.ItemData(cmbTestQuickload.NewIndex) = i
     Next i
     
     
     'Done
     Screen.MousePointer = vbNormal
End Sub

Private Sub lstFunctions_ItemClick(ByVal Item As MSComctlLib.ListItem)
     On Local Error Resume Next
     
     'Show function title
     lblFunction.Caption = ShortedText(Item.Text, lblFunction.width / Screen.TwipsPerPixelX, True)
     lblFunctionDesc.Caption = ""
     lblFunctionDesc.Caption = Item.ListSubItems("DESC").Text
     
     'Clear the combobox options
     cmbShortcut.Clear
     
     'Mousebutton options
     If (Val(Item.ListSubItems("MOUSEBUTTONS").Text) <> 0) Then
          cmbShortcut.AddItem "Mouse1": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_0
          cmbShortcut.AddItem "Mouse2": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_1
          cmbShortcut.AddItem "Mouse3": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_2
          cmbShortcut.AddItem "Mouse4": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_3
          cmbShortcut.AddItem "Mouse5": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_4
          cmbShortcut.AddItem "Mouse6": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_5
          cmbShortcut.AddItem "Mouse7": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_6
          cmbShortcut.AddItem "Mouse8": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_7
          cmbShortcut.AddItem "Shift+Mouse1": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_0 Or (vbShiftMask * (2 ^ 16))
          cmbShortcut.AddItem "Shift+Mouse2": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_1 Or (vbShiftMask * (2 ^ 16))
          cmbShortcut.AddItem "Shift+Mouse3": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_2 Or (vbShiftMask * (2 ^ 16))
          cmbShortcut.AddItem "Shift+Mouse4": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_3 Or (vbShiftMask * (2 ^ 16))
          cmbShortcut.AddItem "Shift+Mouse5": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_4 Or (vbShiftMask * (2 ^ 16))
          cmbShortcut.AddItem "Shift+Mouse6": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_5 Or (vbShiftMask * (2 ^ 16))
          cmbShortcut.AddItem "Shift+Mouse7": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_6 Or (vbShiftMask * (2 ^ 16))
          cmbShortcut.AddItem "Shift+Mouse8": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_7 Or (vbShiftMask * (2 ^ 16))
          cmbShortcut.AddItem "Ctrl+Mouse1": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_0 Or (vbCtrlMask * (2 ^ 16))
          cmbShortcut.AddItem "Ctrl+Mouse2": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_1 Or (vbCtrlMask * (2 ^ 16))
          cmbShortcut.AddItem "Ctrl+Mouse3": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_2 Or (vbCtrlMask * (2 ^ 16))
          cmbShortcut.AddItem "Ctrl+Mouse4": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_3 Or (vbCtrlMask * (2 ^ 16))
          cmbShortcut.AddItem "Ctrl+Mouse5": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_4 Or (vbCtrlMask * (2 ^ 16))
          cmbShortcut.AddItem "Ctrl+Mouse6": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_5 Or (vbCtrlMask * (2 ^ 16))
          cmbShortcut.AddItem "Ctrl+Mouse7": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_6 Or (vbCtrlMask * (2 ^ 16))
          cmbShortcut.AddItem "Ctrl+Mouse8": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_7 Or (vbCtrlMask * (2 ^ 16))
          cmbShortcut.AddItem "Ctrl+Shift+Mouse1": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_0 Or (vbCtrlMask * (2 ^ 16)) Or (vbShiftMask * (2 ^ 16))
          cmbShortcut.AddItem "Ctrl+Shift+Mouse2": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_1 Or (vbCtrlMask * (2 ^ 16)) Or (vbShiftMask * (2 ^ 16))
          cmbShortcut.AddItem "Ctrl+Shift+Mouse3": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_2 Or (vbCtrlMask * (2 ^ 16)) Or (vbShiftMask * (2 ^ 16))
          cmbShortcut.AddItem "Ctrl+Shift+Mouse4": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_3 Or (vbCtrlMask * (2 ^ 16)) Or (vbShiftMask * (2 ^ 16))
          cmbShortcut.AddItem "Ctrl+Shift+Mouse5": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_4 Or (vbCtrlMask * (2 ^ 16)) Or (vbShiftMask * (2 ^ 16))
          cmbShortcut.AddItem "Ctrl+Shift+Mouse6": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_5 Or (vbCtrlMask * (2 ^ 16)) Or (vbShiftMask * (2 ^ 16))
          cmbShortcut.AddItem "Ctrl+Shift+Mouse7": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_6 Or (vbCtrlMask * (2 ^ 16)) Or (vbShiftMask * (2 ^ 16))
          cmbShortcut.AddItem "Ctrl+Shift+Mouse8": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_BUTTON_7 Or (vbCtrlMask * (2 ^ 16)) Or (vbShiftMask * (2 ^ 16))
     End If
     
     'Mousescroll options
     If (Val(Item.ListSubItems("MOUSESCROLL").Text) <> 0) Then
          cmbShortcut.AddItem "ScrollUp": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_SCROLL_UP
          cmbShortcut.AddItem "ScrollDown": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_SCROLL_DOWN
          cmbShortcut.AddItem "Shift+ScrollUp": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_SCROLL_UP Or (vbShiftMask * (2 ^ 16))
          cmbShortcut.AddItem "Shift+ScrollDown": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_SCROLL_DOWN Or (vbShiftMask * (2 ^ 16))
          cmbShortcut.AddItem "Ctrl+ScrollUp": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_SCROLL_UP Or (vbCtrlMask * (2 ^ 16))
          cmbShortcut.AddItem "Ctrl+ScrollDown": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_SCROLL_DOWN Or (vbCtrlMask * (2 ^ 16))
          cmbShortcut.AddItem "Ctrl+Shift+ScrollUp": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_SCROLL_UP Or (vbShiftMask * (2 ^ 16)) Or (vbCtrlMask * (2 ^ 16))
          cmbShortcut.AddItem "Ctrl+Shift+ScrollDown": cmbShortcut.ItemData(cmbShortcut.NewIndex) = MOUSE_SCROLL_DOWN Or (vbShiftMask * (2 ^ 16)) Or (vbCtrlMask * (2 ^ 16))
     End If
     
     'Enable controls
     cmdUnbind.Enabled = (Val(Item.ListSubItems("UNBIND").Text) <> 0)
     cmbShortcut.Enabled = (cmbShortcut.ListCount > 0)
     txtShortcut.Enabled = True
     If cmbShortcut.Enabled Then
          lblSpecialShortcut.ForeColor = vbWindowText
     Else
          lblSpecialShortcut.ForeColor = vbGrayText
     End If
     
     'Check if bound
     If (Item.ListSubItems(1).Text <> NameForKeycode(0, 0)) Then
          
          'Show shortcut
          txtShortcut.Text = Item.ListSubItems(1).Text
          txtShortcut.SelStart = Len(txtShortcut.Text)
     Else
          
          'Show nothing
          txtShortcut.Text = ""
     End If
End Sub

Private Sub lstFunctions_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     On Local Error Resume Next
     txtShortcut.SetFocus
End Sub

Private Sub lstGames_ItemClick(ByVal Item As MSComctlLib.ListItem)
     
     'Update label and textbox
     lblGameIWAD = Item.Text & " IWAD:"
     txtGameIWAD = Item.ListSubItems(1).Text
End Sub

Private Sub Setup3DPanel()
     On Error GoTo No3D
     Dim AdapterCount As Long
     Dim AdapterInfo As D3DADAPTER_IDENTIFIER9
     Dim AdapterDescription As String
     Dim CurrentAdapter As String
     Dim i As Long
     
     'Make sure the combo does not start enumerating modes already
     EnumerateModes = False
     
     'Get the current adapter mode from config
     ModeWidth = Val(Config("videowidth"))
     ModeHeight = Val(Config("videoheight"))
     ModeFormat = Val(Config("videoformat"))
     ModeRate = Val(Config("videorate"))
     
     'Get the current adapter from config
     CurrentAdapter = Config("videoadapterdesc")
'     D3D.GetAdapterIdentifier i, D3DENUM_NO_WHQL_LEVEL, AdapterInfo
'     CurrentAdapter = StringFromBytes(AdapterInfo.Description)
'     If (VarType(Config("videoadapterdesc")) = vbString) And _
'        (Trim$(Config("videoadapterdesc")) <> "") Then
'
'          'Use current set adapter
'          CurrentAdapter = Config("videoadapterdesc")
'     End If
     
     'Get the number of adapters
     AdapterCount = D3D.GetAdapterCount
     
     'Check if any adapters could be found
     If AdapterCount > 0 Then
          
          'Fill the Video Driver combo with adapters
          For i = 0 To (AdapterCount - 1)
               
               'Get the adapter info
               D3D.GetAdapterIdentifier i, 0, AdapterInfo
               AdapterDescription = StringFromBytes(AdapterInfo.Description)
               
               'Add to the combo
               cmbVideoDriver.AddItem AdapterDescription
               
               'Check if we should select this adapter
               If AdapterDescription = CurrentAdapter Then cmbVideoDriver.ListIndex = i
          Next i
          
          'Now modes may be enumerated
          EnumerateModes = True
          cmbVideoDriver_Change
     End If
     
     'Leave now
     Exit Sub
     
No3D:
     
     'No 3D available
     tbsOptions.Tabs(7).tag = "NO"
     fraNo3D.visible = True
     fraNo3D.ZOrder 0
     fraNo3D.Move 0, 0
End Sub

Private Sub tbsOptions_Click()
     Dim i As Long
     
     'Show the frame
     For i = fraOptions.LBound To fraOptions.UBound
          
          'Check if this tab is selected
          If (i = tbsOptions.SelectedItem.Index - 1) Then
               
               'Show the frame
               fraOptions(i).visible = True
               
               'Leave here
               Exit For
          End If
     Next i
     
     'Hide all other frames
     For i = fraOptions.LBound To fraOptions.UBound
          
          'Hide frame if not selected
          If (i <> tbsOptions.SelectedItem.Index - 1) Then fraOptions(i).visible = False
     Next i
     
     'Check if a warning should be displayed
     If (Val(tbsOptions.SelectedItem.tag) = 1) Then MsgBox "Warning: Could not detect 3D acceleration adapters or screen resolutions." & vbLf & "Please ensure that you have the latest DirectX installed and that you have a DirectX compitable videocard.", vbCritical
End Sub

Private Sub txtBrightness_GotFocus()
     SelectAllText txtBrightness
End Sub


Private Sub txtDefaultLower_GotFocus()
     SelectAllText txtDefaultLower
End Sub


Private Sub txtDefaultMiddle_GotFocus()
     SelectAllText txtDefaultMiddle
End Sub


Private Sub txtDefaultTCeiling_GotFocus()
     SelectAllText txtDefaultTCeiling
End Sub


Private Sub txtDefaultTFloor_GotFocus()
     SelectAllText txtDefaultTFloor
End Sub


Private Sub txtDefaultUpper_GotFocus()
     SelectAllText txtDefaultUpper
End Sub


Private Sub txtExportNodebuilderExe_GotFocus()
     SelectAllText txtExportNodebuilderExe
End Sub


Private Sub txtExportNodebuilderParams_GotFocus()
     SelectAllText txtExportNodebuilderParams
End Sub


Private Sub txtFOV_GotFocus()
     SelectAllText txtFOV
End Sub


Private Sub txtGameIWAD_GotFocus()
     SelectAllText txtGameIWAD
End Sub


Private Sub txtGamma_GotFocus()
     SelectAllText txtGamma
End Sub


Private Sub txtMouseSpeed_GotFocus()
     SelectAllText txtMouseSpeed
End Sub


Private Sub txtMoveSpeed_GotFocus()
     SelectAllText txtMoveSpeed
End Sub


Private Sub txtNodebuilderExe_GotFocus()
     SelectAllText txtNodebuilderExe
End Sub


Private Sub txtNodebuilderParams_GotFocus()
     SelectAllText txtNodebuilderParams
End Sub

Private Sub txtNodebuilderParams_KeyDown(KeyCode As Integer, Shift As Integer)
     
     'No settings from profile
     cmbNodeQuickload.ListIndex = -1
End Sub


Private Sub txtNodebuilderParams_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'No settings from profile
     cmbNodeQuickload.ListIndex = -1
End Sub


Private Sub txtPrefabFolder_GotFocus()
     SelectAllText txtPrefabFolder
End Sub


Private Sub txtQuickPrefab_GotFocus(Index As Integer)
     SelectAllText txtQuickPrefab(Index)
End Sub


Private Sub txtShortcut_GotFocus()
     cmdCancel.Cancel = False
     cmdCancel.TabStop = False
     cmdOK.Default = False
     cmdOK.TabStop = False
     tbsOptions.TabStop = False
End Sub

Private Sub txtShortcut_KeyDown(KeyCode As Integer, Shift As Integer)
     
     'Ignore shift keys alone
     If (KeyCode = 16) Or (KeyCode = 17) Or (KeyCode = 18) Then Exit Sub
     
     'Ignore CTRL+TAB
     If (KeyCode = vbKeyTab) And (Shift = vbCtrlMask) Then Exit Sub
     
     'Show the combination name
     txtShortcut.Text = NameForKeycode(KeyCode, Shift)
     
     'Set the combination on the tag of selected item
     lstFunctions.SelectedItem.tag = KeyCode Or (Shift * (2 ^ 16))
     lstFunctions.SelectedItem.ListSubItems(1) = txtShortcut.Text
     
     'Remove anything in the combo
     cmbShortcut.ListIndex = -1
End Sub

Private Sub txtShortcut_KeyPress(KeyAscii As Integer)
     KeyAscii = 0
     txtShortcut.SelStart = Len(txtShortcut.Text)
End Sub

Private Sub txtShortcut_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Ignore CTRL+TAB
     If (KeyCode = vbKeyTab) And (Shift = vbCtrlMask) Then Exit Sub
     
     'Update
     lstFunctions_ItemClick lstFunctions.SelectedItem
     txtShortcut.SelStart = Len(txtShortcut.Text)
End Sub

Private Sub txtShortcut_LostFocus()
     cmdCancel.TabStop = True
     cmdCancel.Cancel = True
     cmdOK.TabStop = True
     cmdOK.Default = True
     tbsOptions.TabStop = True
End Sub

Private Sub txtTestExe_GotFocus()
     SelectAllText txtTestExe
End Sub


Private Sub txtTestParams_GotFocus()
     SelectAllText txtTestParams
End Sub

Private Sub txtTestParams_KeyDown(KeyCode As Integer, Shift As Integer)
     
     'No settings from profile
     cmbTestQuickload.ListIndex = -1
End Sub


Private Sub txtTestParams_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'No settings from profile
     cmbTestQuickload.ListIndex = -1
End Sub


Private Sub txtVideoDistance_GotFocus()
     SelectAllText txtVideoDistance
End Sub


Private Sub valAutostitchDistance_GotFocus()
     SelectAllText valAutostitchDistance
End Sub


Private Sub valDefaultBrightness_GotFocus()
     SelectAllText valDefaultBrightness
End Sub


Private Sub valDefaultGrid_GotFocus()
     SelectAllText valDefaultGrid
End Sub


Private Sub valDefaultHCeiling_GotFocus()
     SelectAllText valDefaultHCeiling
End Sub


Private Sub valDefaultHFloor_GotFocus()
     SelectAllText valDefaultHFloor
End Sub


Private Sub valDefaultThing_GotFocus()
     SelectAllText valDefaultThing
End Sub


Private Sub valIndicatorSize_GotFocus()
     SelectAllText valIndicatorSize
End Sub


Private Sub valLinedefSelectDistance_GotFocus()
     SelectAllText valLinedefSelectDistance
End Sub


Private Sub valLinesplitDistance_GotFocus()
     SelectAllText valLinesplitDistance
End Sub


Private Sub valMaxUndos_GotFocus()
     SelectAllText valMaxUndos
End Sub


Private Sub valScrollPixels_GotFocus()
     SelectAllText valScrollPixels
End Sub


Private Sub valThingSelectDistance_GotFocus()
     SelectAllText valThingSelectDistance
End Sub


Private Sub valVertexSelectDistance_GotFocus()
     SelectAllText valVertexSelectDistance
End Sub


Private Sub valZoomSpeed_GotFocus()
     SelectAllText valZoomSpeed
End Sub


