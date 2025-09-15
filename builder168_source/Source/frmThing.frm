VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmThing 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Thing Selection"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
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
   Icon            =   "frmThing.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   401
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   658
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picThing 
      BackColor       =   &H8000000C&
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      HasDC           =   0   'False
      Height          =   1380
      Left            =   4455
      ScaleHeight     =   88
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   88
      TabIndex        =   59
      TabStop         =   0   'False
      ToolTipText     =   "Thing Sprite Preview"
      Top             =   3900
      Width           =   1380
      Begin VB.Image imgThing 
         Height          =   1260
         Left            =   60
         Stretch         =   -1  'True
         ToolTipText     =   "Thing Sprite Preview"
         Top             =   60
         Width           =   1200
      End
   End
   Begin MSComctlLib.ImageList imglstThings 
      Left            =   90
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThing.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThing.frx":05A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThing.frx":0B40
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThing.frx":10DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThing.frx":1674
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThing.frx":1C0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThing.frx":21A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThing.frx":2742
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThing.frx":2CDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThing.frx":3276
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThing.frx":3810
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThing.frx":3DAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThing.frx":4344
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThing.frx":48DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThing.frx":4E78
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThing.frx":5412
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   7305
      TabIndex        =   32
      Top             =   5550
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   5595
      TabIndex        =   31
      Top             =   5550
      Width           =   1575
   End
   Begin VB.Frame fraFlags 
      Caption         =   " Flags "
      Height          =   3135
      Left            =   4455
      TabIndex        =   44
      Top             =   555
      Width           =   5175
      Begin VB.TextBox txtRawFlags 
         Height          =   315
         Left            =   390
         MaxLength       =   6
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2670
         Width           =   945
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   0
         Left            =   390
         TabIndex        =   3
         Tag             =   "0"
         Top             =   300
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   1
         Left            =   390
         TabIndex        =   4
         Tag             =   "0"
         Top             =   585
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   2
         Left            =   390
         TabIndex        =   5
         Tag             =   "0"
         Top             =   870
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   3
         Left            =   390
         TabIndex        =   6
         Tag             =   "0"
         Top             =   1155
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   4
         Left            =   390
         TabIndex        =   7
         Tag             =   "0"
         Top             =   1440
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   5
         Left            =   390
         TabIndex        =   8
         Tag             =   "0"
         Top             =   1725
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   6
         Left            =   390
         TabIndex        =   9
         Tag             =   "0"
         Top             =   2010
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   7
         Left            =   390
         TabIndex        =   10
         Tag             =   "0"
         Top             =   2295
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   8
         Left            =   2550
         TabIndex        =   11
         Tag             =   "0"
         Top             =   300
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   9
         Left            =   2550
         TabIndex        =   12
         Tag             =   "0"
         Top             =   585
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   10
         Left            =   2550
         TabIndex        =   13
         Tag             =   "0"
         Top             =   870
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   11
         Left            =   2550
         TabIndex        =   14
         Tag             =   "0"
         Top             =   1155
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   12
         Left            =   2550
         TabIndex        =   15
         Tag             =   "0"
         Top             =   1440
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   13
         Left            =   2550
         TabIndex        =   16
         Tag             =   "0"
         Top             =   1725
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   14
         Left            =   2550
         TabIndex        =   17
         Tag             =   "0"
         Top             =   2010
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   15
         Left            =   2550
         TabIndex        =   18
         Tag             =   "0"
         Top             =   2295
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.Label Label2 
         Caption         =   "flags value"
         Height          =   195
         Left            =   1395
         TabIndex        =   58
         Top             =   2730
         Width           =   1515
      End
   End
   Begin VB.Frame fraThing 
      Caption         =   " Thing "
      Height          =   4725
      Left            =   255
      TabIndex        =   42
      Top             =   555
      Width           =   4065
      Begin MSComctlLib.TreeView trvThings 
         Height          =   3390
         Left            =   180
         TabIndex        =   1
         Top             =   330
         Visible         =   0   'False
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   5980
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   529
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "imglstThings"
         Appearance      =   1
      End
      Begin DoomBuilder.ctlValueBox txtThing 
         Height          =   360
         Left            =   690
         TabIndex        =   2
         Top             =   3795
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   635
         MaxLength       =   5
         Min             =   -32768
         EmptyAllowed    =   -1  'True
         Unsigned        =   -1  'True
      End
      Begin MSComctlLib.ListView lstThings 
         Height          =   3390
         Left            =   180
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   330
         Visible         =   0   'False
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   5980
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imglstThings"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Category"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Num"
            Object.Width           =   1111
         EndProperty
      End
      Begin VB.Label lblThingBlocks 
         AutoSize        =   -1  'True
         Caption         =   "?"
         Height          =   210
         Left            =   2985
         TabIndex        =   67
         Top             =   4290
         Width           =   90
      End
      Begin VB.Label lblThingHangs 
         AutoSize        =   -1  'True
         Caption         =   "?"
         Height          =   210
         Left            =   1620
         TabIndex        =   66
         Top             =   4290
         Width           =   90
      End
      Begin VB.Label lblThingHeight 
         AutoSize        =   -1  'True
         Caption         =   "?"
         Height          =   210
         Left            =   3555
         TabIndex        =   65
         Top             =   3855
         Width           =   90
      End
      Begin VB.Label lblThingWidth 
         AutoSize        =   -1  'True
         Caption         =   "?"
         Height          =   210
         Left            =   2505
         TabIndex        =   64
         Top             =   3855
         Width           =   90
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Hangs from ceiling:"
         Height          =   210
         Left            =   195
         TabIndex        =   63
         Top             =   4290
         Width           =   1380
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Blocking:"
         Height          =   210
         Left            =   2280
         TabIndex        =   62
         Top             =   4290
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Height:"
         Height          =   210
         Left            =   3015
         TabIndex        =   61
         Top             =   3855
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Width:"
         Height          =   210
         Left            =   1995
         TabIndex        =   60
         Top             =   3855
         Width           =   450
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   210
         Left            =   195
         TabIndex        =   43
         Top             =   3855
         UseMnemonic     =   0   'False
         Width           =   405
      End
   End
   Begin VB.Frame fraAngle 
      Caption         =   " Coordination "
      Height          =   1485
      Left            =   5940
      TabIndex        =   45
      Top             =   3795
      Width           =   3690
      Begin VB.OptionButton optAngle 
         Height          =   255
         Index           =   7
         Left            =   2430
         TabIndex        =   41
         TabStop         =   0   'False
         Tag             =   "135"
         Top             =   345
         Width           =   210
      End
      Begin VB.OptionButton optAngle 
         Height          =   255
         Index           =   6
         Left            =   2265
         TabIndex        =   40
         TabStop         =   0   'False
         Tag             =   "180"
         Top             =   660
         Width           =   210
      End
      Begin VB.OptionButton optAngle 
         Height          =   255
         Index           =   5
         Left            =   2430
         TabIndex        =   39
         TabStop         =   0   'False
         Tag             =   "225"
         Top             =   975
         Width           =   210
      End
      Begin VB.OptionButton optAngle 
         Height          =   255
         Index           =   3
         Left            =   2745
         TabIndex        =   38
         TabStop         =   0   'False
         Tag             =   "270"
         Top             =   1110
         Width           =   210
      End
      Begin VB.OptionButton optAngle 
         Height          =   255
         Index           =   0
         Left            =   2745
         TabIndex        =   34
         TabStop         =   0   'False
         Tag             =   "90"
         Top             =   210
         Width           =   210
      End
      Begin VB.OptionButton optAngle 
         Height          =   255
         Index           =   4
         Left            =   3075
         TabIndex        =   37
         TabStop         =   0   'False
         Tag             =   "315"
         Top             =   975
         Width           =   210
      End
      Begin VB.OptionButton optAngle 
         Height          =   255
         Index           =   2
         Left            =   3225
         TabIndex        =   36
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   660
         Width           =   210
      End
      Begin VB.OptionButton optAngle 
         Height          =   255
         Index           =   1
         Left            =   3075
         TabIndex        =   35
         TabStop         =   0   'False
         Tag             =   "45"
         Top             =   345
         Width           =   210
      End
      Begin DoomBuilder.ctlValueBox txtAngle 
         Height          =   360
         Left            =   975
         TabIndex        =   20
         Top             =   390
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   635
         MaxLength       =   6
         Min             =   -32767
         SmallChange     =   45
         Value           =   ""
         EmptyAllowed    =   -1  'True
         RelativeAllowed =   -1  'True
      End
      Begin DoomBuilder.ctlValueBox txtHeight 
         Height          =   360
         Left            =   975
         TabIndex        =   21
         Top             =   840
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   635
         MaxLength       =   6
         Min             =   -32767
         SmallChange     =   8
         Value           =   ""
         EmptyAllowed    =   -1  'True
         RelativeAllowed =   -1  'True
      End
      Begin VB.Label lblHeight 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Z Height:"
         Height          =   210
         Left            =   195
         TabIndex        =   55
         Top             =   915
         Width           =   645
      End
      Begin VB.Label lblAngle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Angle:"
         Height          =   210
         Left            =   375
         TabIndex        =   46
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   465
      End
      Begin VB.Line lineAngle 
         BorderWidth     =   3
         X1              =   2835
         X2              =   2850
         Y1              =   765
         Y2              =   780
      End
   End
   Begin VB.Frame fraAction 
      Caption         =   " Action "
      Height          =   3765
      Left            =   255
      TabIndex        =   48
      Top             =   1515
      Visible         =   0   'False
      Width           =   9375
      Begin VB.CommandButton cmdSelectType 
         Caption         =   "Select Action..."
         Height          =   345
         Left            =   2865
         TabIndex        =   25
         Top             =   450
         Width           =   1545
      End
      Begin DoomBuilder.ctlValueBox txtType 
         Height          =   360
         Left            =   1590
         TabIndex        =   24
         Top             =   435
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   635
         MaxLength       =   5
         Value           =   ""
         EmptyAllowed    =   -1  'True
      End
      Begin DoomBuilder.ctlValueBox txtArgument 
         Height          =   360
         Index           =   0
         Left            =   6930
         TabIndex        =   26
         Top             =   420
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   635
         Max             =   255
         MaxLength       =   3
         Value           =   ""
         EmptyAllowed    =   -1  'True
         RelativeAllowed =   -1  'True
      End
      Begin DoomBuilder.ctlValueBox txtArgument 
         Height          =   360
         Index           =   1
         Left            =   6930
         TabIndex        =   27
         Top             =   810
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   635
         Max             =   255
         MaxLength       =   3
         Value           =   ""
         EmptyAllowed    =   -1  'True
         RelativeAllowed =   -1  'True
      End
      Begin DoomBuilder.ctlValueBox txtArgument 
         Height          =   360
         Index           =   4
         Left            =   6930
         TabIndex        =   30
         Top             =   1980
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   635
         Max             =   255
         MaxLength       =   3
         Value           =   ""
         EmptyAllowed    =   -1  'True
         RelativeAllowed =   -1  'True
      End
      Begin DoomBuilder.ctlValueBox txtArgument 
         Height          =   360
         Index           =   2
         Left            =   6930
         TabIndex        =   28
         Top             =   1200
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   635
         Max             =   255
         MaxLength       =   3
         Value           =   ""
         EmptyAllowed    =   -1  'True
         RelativeAllowed =   -1  'True
      End
      Begin DoomBuilder.ctlValueBox txtArgument 
         Height          =   360
         Index           =   3
         Left            =   6930
         TabIndex        =   29
         Top             =   1590
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   635
         Max             =   255
         MaxLength       =   3
         Value           =   ""
         EmptyAllowed    =   -1  'True
         RelativeAllowed =   -1  'True
      End
      Begin VB.Label lblArgument 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Argument 4:"
         Height          =   210
         Index           =   3
         Left            =   5925
         TabIndex        =   54
         Top             =   1665
         Width           =   885
      End
      Begin VB.Label lblArgument 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Argument 3:"
         Height          =   210
         Index           =   2
         Left            =   5925
         TabIndex        =   53
         Top             =   1275
         Width           =   885
      End
      Begin VB.Label lblEffect 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Thing Action:"
         Height          =   210
         Left            =   525
         TabIndex        =   52
         Top             =   510
         UseMnemonic     =   0   'False
         Width           =   945
      End
      Begin VB.Label lblArgument 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Argument 1:"
         Height          =   210
         Index           =   0
         Left            =   5925
         TabIndex        =   51
         Top             =   480
         Width           =   885
      End
      Begin VB.Label lblArgument 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Argument 2:"
         Height          =   210
         Index           =   1
         Left            =   5925
         TabIndex        =   50
         Top             =   885
         Width           =   885
      End
      Begin VB.Label lblArgument 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Argument 5:"
         Height          =   210
         Index           =   4
         Left            =   5925
         TabIndex        =   49
         Top             =   2055
         Width           =   885
      End
   End
   Begin VB.Frame fraTag 
      Caption         =   " Tag "
      Height          =   885
      Left            =   255
      TabIndex        =   56
      Top             =   555
      Visible         =   0   'False
      Width           =   9375
      Begin VB.CommandButton cmdNextTag 
         Caption         =   "Next Unused"
         Height          =   345
         Left            =   2865
         TabIndex        =   23
         Top             =   315
         Width           =   1545
      End
      Begin DoomBuilder.ctlValueBox txtTag 
         Height          =   360
         Left            =   1590
         TabIndex        =   22
         Top             =   300
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   635
         MaxLength       =   5
         Min             =   -32768
         Value           =   ""
         EmptyAllowed    =   -1  'True
         RelativeAllowed =   -1  'True
         Unsigned        =   -1  'True
      End
      Begin VB.Label lblTag 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Thing Tag:"
         Height          =   210
         Left            =   690
         TabIndex        =   57
         Top             =   375
         UseMnemonic     =   0   'False
         Width           =   750
      End
   End
   Begin MSComctlLib.TabStrip tbsPanel 
      Height          =   5325
      Left            =   90
      TabIndex        =   33
      Top             =   105
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   9393
      TabWidthStyle   =   2
      ShowTips        =   0   'False
      TabFixedWidth   =   2646
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Properties"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Effects"
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
   Begin VB.Label lblMakeUndo 
      Height          =   210
      Left            =   540
      TabIndex        =   47
      Top             =   4740
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   900
   End
End
Attribute VB_Name = "frmThing"
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

Private Function CheckThingAngle() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first thing's angle
     CheckThingAngle = things(Indices(LBound(Indices))).angle
     
     'Go for all selected things
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if the angle is different
          If (things(Indices(i)).angle <> CheckThingAngle) Then
               CheckThingAngle = ""
               Exit Function
          End If
     Next i
End Function

Private Function CheckThingArg0() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set to first line arg
     CheckThingArg0 = things(Indices(LBound(Indices))).arg0
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if different
          If (things(Indices(i)).arg0 <> CheckThingArg0) Then
               CheckThingArg0 = ""
               Exit For
          End If
     Next i
End Function

Private Function CheckThingArg1() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set to first line arg
     CheckThingArg1 = things(Indices(LBound(Indices))).arg1
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if different
          If (things(Indices(i)).arg1 <> CheckThingArg1) Then
               CheckThingArg1 = ""
               Exit For
          End If
     Next i
End Function

Private Function CheckThingArg2() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set to first line arg
     CheckThingArg2 = things(Indices(LBound(Indices))).arg2
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if different
          If (things(Indices(i)).arg2 <> CheckThingArg2) Then
               CheckThingArg2 = ""
               Exit For
          End If
     Next i
End Function

Private Function CheckThingArg3() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set to first line arg
     CheckThingArg3 = things(Indices(LBound(Indices))).arg3
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if different
          If (things(Indices(i)).arg3 <> CheckThingArg3) Then
               CheckThingArg3 = ""
               Exit For
          End If
     Next i
End Function

Private Function CheckThingArg4() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set to first line arg
     CheckThingArg4 = things(Indices(LBound(Indices))).arg4
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if different
          If (things(Indices(i)).arg4 <> CheckThingArg4) Then
               CheckThingArg4 = ""
               Exit For
          End If
     Next i
End Function

Private Function CheckThingEffect() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first things's effect
     CheckThingEffect = things(Indices(LBound(Indices))).effect
     
     'Go for all selected things
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if the effect is different
          If (things(Indices(i)).effect <> CheckThingEffect) Then
               CheckThingEffect = ""
               Exit Function
          End If
     Next i
End Function

Private Function CheckThingFlag(ByRef Flag As Long) As Long
     Dim i As Long
     Dim Indices As Variant
     Dim Numchecked As Long
     
     'Go for all selected things
     Indices = selected.Items
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if the flag is set
          If ((things(Indices(i)).Flags And Flag) = Flag) Then Numchecked = Numchecked + 1
     Next i
     
     'Check what result to return
     If (Numchecked = 0) Then
          CheckThingFlag = vbUnchecked
     ElseIf (Numchecked = numselected) Then
          CheckThingFlag = vbChecked
     Else
          CheckThingFlag = vbGrayed
     End If
End Function

Private Function CheckThingHeight() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first thing's Z
     CheckThingHeight = things(Indices(LBound(Indices))).Z
     
     'Go for all selected things
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if the Z is different
          If (things(Indices(i)).Z <> CheckThingHeight) Then
               CheckThingHeight = ""
               Exit Function
          End If
     Next i
End Function

Private Function CheckThingRawFlags() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first thing's flags
     CheckThingRawFlags = things(Indices(LBound(Indices))).Flags
     
     'Go for all selected things
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if the type is different
          If (things(Indices(i)).Flags <> CheckThingRawFlags) Then
               CheckThingRawFlags = ""
               Exit Function
          End If
     Next i
End Function

Private Function CheckThingTag() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first thing's tag
     CheckThingTag = things(Indices(LBound(Indices))).tag
     
     'Go for all selected things
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if the tag is different
          If (things(Indices(i)).tag <> CheckThingTag) Then
               CheckThingTag = ""
               Exit Function
          End If
     Next i
End Function

Private Function CheckThingType() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first thing's type
     CheckThingType = things(Indices(LBound(Indices))).thing
     
     'Go for all selected things
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if the type is different
          If (things(Indices(i)).thing <> CheckThingType) Then
               CheckThingType = ""
               Exit Function
          End If
     Next i
End Function

Private Sub chkFlag_Click(Index As Integer)
     txtRawFlags.Text = ""
     txtRawFlags.Text = ""
End Sub

Private Sub cmdCancel_Click()
     Unload Me
     Set frmThing = Nothing
End Sub

Private Sub cmdNextTag_Click()
     txtTag.Text = NextThingTag
End Sub

Private Sub cmdOK_Click()
     Dim Indices As Variant
     Dim i As Long
     Dim f As Long
     Dim t As Long
     
     'Change mousepointer
     Screen.MousePointer = vbHourglass
     
     'Make undo
     If (lblMakeUndo.Caption = "") Then CreateUndo "thing edit"
     
     'Get the selection indices
     Indices = selected.Items
     
     'Go for all selected things
     For i = LBound(Indices) To UBound(Indices)
          
          'Get the thing index
          t = Indices(i)
          
          'Check if raw code set
          If (Trim$(txtRawFlags.Text) = "") Then
               
               'Go for all individual flags
               For f = 0 To 15
                    
                    'Check if this flag can be set
                    If (chkFlag(f).tag <> "0") Then
                         
                         'Check if the flag is marked to be set
                         If (chkFlag(f).Value = vbChecked) Then
                              
                              'Add the flag on the thing
                              things(t).Flags = things(t).Flags Or CLng(chkFlag(f).tag)
                         
                         'Check if the flag is marked to be removed
                         ElseIf (chkFlag(f).Value = vbUnchecked) Then
                              
                              'Remove the flag from the thing
                              things(t).Flags = things(t).Flags And Not CLng(chkFlag(f).tag)
                         End If
                    End If
               Next f
          Else
               
               'Set flags from raw
               On Error Resume Next
               things(t).Flags = Val(txtRawFlags.Text)
               On Error GoTo 0
          End If
          
          'Apply tag if a thing tag is specified
          If (txtTag.Text <> "") Then things(t).tag = txtTag.RelativeValue(things(t).tag)
          
          'Apply effect if a thing effect is specified
          If (txtType.Text <> "") Then things(t).effect = CLng(txtType.Text)
          
          'Check if this is an action which we know
          If (mapconfig("linedeftypes").Exists(CStr(things(t).effect)) = True) Then
               
               'Set the marking references on the linedef
               With things(t)
                    .argref0 = mapconfig("linedeftypes")(CStr(things(t).effect))("mark1")
                    .argref1 = mapconfig("linedeftypes")(CStr(things(t).effect))("mark2")
                    .argref2 = mapconfig("linedeftypes")(CStr(things(t).effect))("mark3")
                    .argref3 = mapconfig("linedeftypes")(CStr(things(t).effect))("mark4")
                    .argref4 = mapconfig("linedeftypes")(CStr(things(t).effect))("mark5")
               End With
          End If
          
          'Apply arguments if specified
          If (txtArgument(0).Text <> "") Then things(t).arg0 = txtArgument(0).RelativeValue(things(t).arg0)
          If (txtArgument(1).Text <> "") Then things(t).arg1 = txtArgument(1).RelativeValue(things(t).arg1)
          If (txtArgument(2).Text <> "") Then things(t).arg2 = txtArgument(2).RelativeValue(things(t).arg2)
          If (txtArgument(3).Text <> "") Then things(t).arg3 = txtArgument(3).RelativeValue(things(t).arg3)
          If (txtArgument(4).Text <> "") Then things(t).arg4 = txtArgument(4).RelativeValue(things(t).arg4)
          
          'Apply type if a thing type is specified
          If (txtThing.Text <> "") Then things(t).thing = Val(txtThing.Text)
          
          'Apply height if a thing height is specified
          If (txtHeight.Text <> "") Then things(t).Z = txtHeight.RelativeValue(things(t).Z)
          
          'Apply angle if an angle is specified
          If (txtAngle.Text <> "") Then things(t).angle = txtAngle.RelativeValue(things(t).angle)
          
          'Update thing image, color and size
          UpdateThingImageColor t
          UpdateThingSize t
          UpdateThingCategory t
          
          'Save last edited thing
          LastThing = things(t)
          
          'Check if this is the 3D start position
          If (things(t).thing = mapconfig("start3dmode")) Then ApplyPositionFromThing t
     Next i
     
     'Map has changed
     mapchanged = True
     
     'Reset mousepointer
     Screen.MousePointer = vbNormal
     
     'Leave now
     Unload Me
     Set frmThing = Nothing
End Sub

Private Sub cmdSelectType_Click()
     txtType.Text = SelectAction(txtType.Text, Me)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     
     'Adjust shift mask
     CurrentShiftMask = Shift
     
     'Check what key is pressed
     If (KeyCode = vbKeyTab) And (Shift = vbCtrlMask) Then
          
          'Switch to next panel
          If (tbsPanel.SelectedItem.Index = tbsPanel.Tabs.Count) Then
               tbsPanel.Tabs(1).selected = True
          Else
               tbsPanel.Tabs(tbsPanel.SelectedItem.Index + 1).selected = True
          End If
          
          'Focus to panel
          tbsPanel.SetFocus
     End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Adjust shift mask
     CurrentShiftMask = Shift
End Sub


Private Sub Form_Load()
     Dim Keys As Variant
     Dim nflag As Long
     Dim i As Long
     
     'Check if only one selected
     If (numselected = 1) Then Caption = Caption & " - Thing " & selected.Items(0)
     
     'Determine what thing stuff to show
     Select Case mapconfig("mapformat")
          
          'Doom map format
          Case MFMT_DOOM
               
               'No effects panel
               tbsPanel.Tabs.Remove 2
               
               'No Z Height
               lblAngle.top = lblAngle.top + 225
               txtAngle.top = txtAngle.top + 225
               lblHeight.visible = False
               txtHeight.visible = False
               
          'Hexen map format
          Case MFMT_HEXEN
               
               'Height
               txtHeight.Text = CheckThingHeight
               txtHeight.RelativeScroll = (numselected > 1)
               
               'Tag
               txtTag.Text = CheckThingTag
               
               'Effect
               txtType.Text = CheckThingEffect
               txtType_Change
               If (Trim$(txtType.Text) <> "") Then
                    
                    'Arguments
                    txtArgument(0).Value = CheckThingArg0
                    txtArgument(1).Value = CheckThingArg1
                    txtArgument(2).Value = CheckThingArg2
                    txtArgument(3).Value = CheckThingArg3
                    txtArgument(4).Value = CheckThingArg4
               Else
                    
                    'Args are not the same
                    txtArgument(0).Text = ""
                    txtArgument(1).Text = ""
                    txtArgument(2).Text = ""
                    txtArgument(3).Text = ""
                    txtArgument(4).Text = ""
               End If
               
     End Select
     
     'Check if showing tree or list
     If (Val(Config("thingstree")) = vbChecked) Then
          
          'Fill things tree
          FillThingsTree trvThings
          trvThings.visible = True
     Else
          
          'Fill things list
          FillThingsList lstThings
          lstThings.visible = True
     End If
     
     'Thing
     txtThing.Text = CheckThingType
     
     'Go for all flags
     Keys = mapconfig("thingflags").Keys
     For i = 0 To 15
          
          'Check if this flag is known
          If (mapconfig("thingflags").Exists(CStr(2 ^ i)) = True) Then
               
               'Check if not unset
               If CStr(mapconfig("thingflags")(CStr(2 ^ i))) <> "0" Then
                    
                    'Set the checkbox properties
                    chkFlag(nflag).tag = CStr(2 ^ i)
                    chkFlag(nflag).visible = True
                    chkFlag(nflag).Caption = mapconfig("thingflags")(CStr(2 ^ i))
                    
                    'Check this flag
                    chkFlag(nflag).Value = CheckThingFlag(2 ^ i)
                    
                    'Next flag
                    nflag = nflag + 1
               Else
                    
                    'Zero tag
                    chkFlag(nflag).tag = "0"
               End If
          Else
               
               'Zero tag
               chkFlag(nflag).tag = "0"
          End If
     Next i
     
     'Show raw flags
     txtRawFlags.Text = CheckThingRawFlags
     
     'Angle
     txtAngle.Text = CheckThingAngle
     'The option buttons dont work properly with RelativeScroll on :(
     'txtAngle.RelativeScroll = (numselected > 1)
End Sub

Private Sub lstThings_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
     
     'Check if already sorted by this column
     If lstThings.SortKey = (ColumnHeader.Index - 1) Then
          
          'Reverse sort
          If lstThings.SortOrder = lvwAscending Then
               lstThings.SortOrder = lvwDescending
          Else
               lstThings.SortOrder = lvwAscending
          End If
     Else
          
          'Change sort key
          lstThings.SortKey = ColumnHeader.Index - 1
          lstThings.SortOrder = lvwAscending
          lstThings.Sorted = True
     End If
     
     'Save sort
     If (lstThings.SortOrder = lvwAscending) Then
          Config("thingssort") = (lstThings.SortKey + 1)
     Else
          Config("thingssort") = -(lstThings.SortKey + 1)
     End If
End Sub

Private Sub lstThings_ItemClick(ByVal Item As MSComctlLib.ListItem)
     
     'Apply selection
     txtThing.Text = Trim$(Item.tag)
End Sub

Private Sub optAngle_Click(Index As Integer)
     txtAngle.Text = optAngle(Index).tag
End Sub

Private Sub tbsPanel_Click()
     If (tbsPanel.SelectedItem.Index = 1) Then
          fraThing.visible = True
          fraFlags.visible = True
          fraAngle.visible = True
          fraAction.visible = False
          fraTag.visible = False
          picThing.visible = True
     Else
          fraThing.visible = False
          fraFlags.visible = False
          fraAngle.visible = False
          fraAction.visible = True
          fraTag.visible = True
          picThing.visible = False
     End If
End Sub

Private Sub trvThings_NodeClick(ByVal Node As MSComctlLib.Node)
     
     'Check if node is a leaf
     If (Node.Children = 0) Then
          
          'Apply selection
          txtThing.Text = Trim$(Node.tag)
     End If
End Sub

Private Sub txtAngle_Change()
     On Error Resume Next
     Dim i As Long
     
     'Go for all options
     For i = 0 To 7
          optAngle(i).Value = (Val(txtAngle.Text) = optAngle(i).tag)
     Next i
     
     'Turn the line
     If (txtAngle.Text <> "") And (left$(txtAngle.Text, 2) <> "++") And (left$(txtAngle.Text, 2) <> "--") Then
          lineAngle.x2 = lineAngle.x1 + sIn((CSng(txtAngle.Text) + 90) / PiDiv) * 260
          lineAngle.y2 = lineAngle.y1 + Cos((CSng(txtAngle.Text) + 90) / PiDiv) * 260
     Else
          lineAngle.x2 = lineAngle.x1
          lineAngle.y2 = lineAngle.y1
     End If
End Sub

Private Sub txtAngle_GotFocus()
     SelectAllText txtAngle
End Sub


Private Sub txtArgument_GotFocus(Index As Integer)
     SelectAllText txtArgument(Index)
End Sub


Private Sub txtHeight_GotFocus()
     SelectAllText txtHeight
End Sub


Private Sub txtRawFlags_GotFocus()
     SelectAllText txtRawFlags
End Sub


Private Sub txtTag_GotFocus()
     SelectAllText txtTag
End Sub


Private Sub txtThing_Change()
     Dim Cat As String
     Dim a As Long
     Dim num As Long
     
     'Erase thing preview
     Set imgThing.Picture = Nothing
     
     'Select current thing type
     If (txtThing.Text <> "") Then
          
          'Do not give an error when the item cant be found
          On Local Error Resume Next
          trvThings.SelectedItem.selected = False
          trvThings.nodes("T" & txtThing.Text).selected = True
          trvThings.nodes("T" & txtThing.Text).EnsureVisible
          lstThings.SelectedItem.selected = False
          lstThings.ListItems("T" & txtThing.Text).selected = True
          lstThings.ListItems("T" & txtThing.Text).EnsureVisible
          On Local Error GoTo 0
          
          'Check if a thing type is given
          If (Trim$(txtThing.Text) <> "") Then
               
               'Get the number
               num = Val(txtThing.Text)
               
               'Show thing preview if possible
               GetScaledSpritePicture num, imgThing, picThing.ScaleWidth, picThing.ScaleHeight, False
               
               'Find its category
               Cat = GetThingTypeCategory(num)
               
               'Check if in any category
               If (Trim$(Cat) <> "") Then
                    
                    'Check if the thing has any arguments
                    If mapconfig("thingtypes")(Cat)(CStr(num)).Exists("arg1") Or _
                       mapconfig("thingtypes")(Cat)(CStr(num)).Exists("arg2") Or _
                       mapconfig("thingtypes")(Cat)(CStr(num)).Exists("arg3") Or _
                       mapconfig("thingtypes")(Cat)(CStr(num)).Exists("arg4") Or _
                       mapconfig("thingtypes")(Cat)(CStr(num)).Exists("arg5") Then
                         
                         'Disable thing effect
                         txtType.Enabled = False
                         lblEffect.ForeColor = vbGrayText
                         cmdSelectType.Enabled = False
                         
                         'Set all arguments
                         For a = 0 To 4
                              
                              'Check if argument is defined
                              If (mapconfig("thingtypes")(Cat)(txtThing.Text).Exists("arg" & a + 1)) Then
                                   
                                   'Set argument
                                   lblArgument(a).Caption = mapconfig("thingtypes")(Cat)(txtThing.Text)("arg" & a + 1) & ":"
                                   lblArgument(a).ForeColor = vbButtonText
                                   'txtArgument(a).Enabled = True
                              Else
                                   
                                   'Disable argument
                                   lblArgument(a).Caption = "Argument " & a + 1 & ":"
                                   lblArgument(a).ForeColor = vbGrayText
                                   'txtArgument(a).Enabled = False
                              End If
                         Next a
                    Else
                         
                         'Enable thing effect
                         txtType.Enabled = True
                         lblEffect.ForeColor = vbButtonText
                         cmdSelectType.Enabled = True
                         
                         'Let the arguments be set by effect
                         txtType_Change
                    End If
                    
                    'Show thing properties
                    lblThingWidth.Caption = GetThingWidth(Val(txtThing.Text))
                    lblThingHeight.Caption = GetThingHeight(Val(txtThing.Text))
                    lblThingHangs.Caption = YesNo(GetThingHangs(Val(txtThing.Text)))
                    lblThingBlocks.Caption = GetThingBlockingDesc(GetThingBlocking(Val(txtThing.Text)))
               Else
                    
                    'Enable thing effect
                    txtType.Enabled = True
                    lblEffect.ForeColor = vbButtonText
                    cmdSelectType.Enabled = True
                    
                    'No clue about its properties
                    lblThingWidth.Caption = "?"
                    lblThingHeight.Caption = "?"
                    lblThingHangs.Caption = "?"
                    lblThingBlocks.Caption = "?"
                    
                    'Let the arguments be set by effect
                    txtType_Change
               End If
          End If
     End If
End Sub

Private Sub txtThing_GotFocus()
     SelectAllText txtThing
End Sub


Private Sub txtType_Change()
     Dim a As Long
     
     'Check if a value is given
     If (Trim$(txtType.Value) <> "") Then
          
          'Check if the type is known
          If (mapconfig("linedeftypes").Exists(txtType.Value)) Then
               
               'Set all arguments
               For a = 0 To 4
                    
                    'Check if argument is defined
                    If (mapconfig("linedeftypes")(txtType.Value).Exists("arg" & a + 1)) Then
                         
                         'Set argument
                         lblArgument(a).Caption = mapconfig("linedeftypes")(txtType.Value)("arg" & a + 1) & ":"
                         lblArgument(a).ForeColor = vbButtonText
                         'txtArgument(a).Enabled = True
                    Else
                         
                         'Disable argument
                         lblArgument(a).Caption = "Argument " & a + 1 & ":"
                         lblArgument(a).ForeColor = vbGrayText
                         'txtArgument(a).Enabled = False
                    End If
               Next a
          Else
               
               'Disable all arguments
               For a = 0 To 4
                    
                    'Disable argument
                    lblArgument(a).Caption = "Argument " & a + 1 & ":"
                    lblArgument(a).ForeColor = vbGrayText
                    'txtArgument(a).Enabled = False
               Next a
          End If
     Else
          
          'Disable all arguments
          For a = 0 To 4
               
               'Disable argument
               lblArgument(a).Caption = "Argument " & a + 1 & ":"
               lblArgument(a).ForeColor = vbGrayText
               'txtArgument(a).Enabled = False
          Next a
     End If
End Sub

Private Sub txtType_GotFocus()
     SelectAllText txtType
End Sub


