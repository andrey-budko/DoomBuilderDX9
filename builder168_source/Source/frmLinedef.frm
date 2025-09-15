VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmLinedef 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Linedef Selection"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
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
   Icon            =   "frmLinedef.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   401
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   536
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picSectorWarning 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   300
      Left            =   240
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   503
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   4935
      Visible         =   0   'False
      Width           =   7575
      Begin VB.Label lblSectorWarning 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Warning:  Any new sidedefs that may be created will all be part of the specified sector!"
         ForeColor       =   &H80000017&
         Height          =   210
         Left            =   375
         TabIndex        =   49
         Top             =   30
         UseMnemonic     =   0   'False
         Width           =   6315
      End
      Begin VB.Image imgSectorWarning 
         Height          =   240
         Left            =   45
         Picture         =   "frmLinedef.frx":000C
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4665
      TabIndex        =   44
      Top             =   5550
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   6375
      TabIndex        =   46
      Top             =   5550
      Width           =   1575
   End
   Begin VB.Frame fraAction 
      Caption         =   " Action "
      Height          =   2295
      Left            =   240
      TabIndex        =   62
      Top             =   555
      Width           =   7575
      Begin VB.CommandButton cmdNextTag 
         Caption         =   "Next Unused"
         Height          =   345
         Left            =   2715
         TabIndex        =   22
         Top             =   780
         Width           =   1545
      End
      Begin VB.CommandButton cmdSelectType 
         Caption         =   "Select Action..."
         Height          =   345
         Left            =   2715
         TabIndex        =   19
         Top             =   270
         Width           =   1545
      End
      Begin VB.ComboBox cmbActivation 
         Height          =   330
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   780
         Width           =   2835
      End
      Begin DoomBuilder.ctlValueBox txtType 
         Height          =   360
         Left            =   1440
         TabIndex        =   18
         Top             =   255
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   635
         MaxLength       =   5
         Min             =   -32768
         Value           =   ""
         EmptyAllowed    =   -1  'True
         Unsigned        =   -1  'True
      End
      Begin DoomBuilder.ctlValueBox txtTag 
         Height          =   360
         Left            =   1440
         TabIndex        =   21
         Top             =   765
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
      Begin DoomBuilder.ctlValueBox txtArgument 
         Height          =   360
         Index           =   0
         Left            =   6570
         TabIndex        =   23
         Top             =   240
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
         Left            =   6570
         TabIndex        =   24
         Top             =   630
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
         Left            =   6570
         TabIndex        =   27
         Top             =   1800
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
         Left            =   6570
         TabIndex        =   25
         Top             =   1020
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
         Left            =   6570
         TabIndex        =   26
         Top             =   1410
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
         Caption         =   "Argument 5:"
         Height          =   210
         Index           =   4
         Left            =   5565
         TabIndex        =   70
         Top             =   1875
         Width           =   885
      End
      Begin VB.Label lblArgument 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Argument 2:"
         Height          =   210
         Index           =   1
         Left            =   5565
         TabIndex        =   69
         Top             =   705
         Width           =   885
      End
      Begin VB.Label lblArgument 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Argument 1:"
         Height          =   210
         Index           =   0
         Left            =   5565
         TabIndex        =   68
         Top             =   300
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Linedef Action:"
         Height          =   210
         Left            =   225
         TabIndex        =   67
         Top             =   330
         UseMnemonic     =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblTag 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sector Tag:"
         Height          =   210
         Left            =   450
         TabIndex        =   66
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   840
      End
      Begin VB.Label lblArgument 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Argument 3:"
         Height          =   210
         Index           =   2
         Left            =   5565
         TabIndex        =   65
         Top             =   1095
         Width           =   885
      End
      Begin VB.Label lblArgument 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Argument 4:"
         Height          =   210
         Index           =   3
         Left            =   5565
         TabIndex        =   64
         Top             =   1485
         Width           =   885
      End
      Begin VB.Label lblActivation 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Activation:"
         Height          =   210
         Left            =   555
         TabIndex        =   63
         Top             =   840
         Width           =   765
      End
   End
   Begin VB.Frame fraFlags 
      Caption         =   " Flags "
      Height          =   1905
      Left            =   240
      TabIndex        =   71
      Top             =   2940
      Width           =   7575
      Begin VB.CheckBox chkFlag 
         Caption         =   "Repeatable Effect"
         Height          =   255
         Index           =   0
         Left            =   165
         TabIndex        =   28
         Tag             =   "0"
         Top             =   270
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Linedef Flag"
         Height          =   255
         Index           =   1
         Left            =   165
         TabIndex        =   29
         Tag             =   "0"
         Top             =   525
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Linedef Flag"
         Height          =   255
         Index           =   2
         Left            =   165
         TabIndex        =   30
         Tag             =   "0"
         Top             =   780
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Linedef Flag"
         Height          =   255
         Index           =   3
         Left            =   165
         TabIndex        =   31
         Tag             =   "0"
         Top             =   1035
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Linedef Flag"
         Height          =   255
         Index           =   4
         Left            =   165
         TabIndex        =   32
         Tag             =   "0"
         Top             =   1290
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Linedef Flag"
         Height          =   255
         Index           =   5
         Left            =   165
         TabIndex        =   33
         Tag             =   "0"
         Top             =   1545
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Linedef Flag"
         Height          =   255
         Index           =   6
         Left            =   2670
         TabIndex        =   34
         Tag             =   "0"
         Top             =   270
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Linedef Flag"
         Height          =   255
         Index           =   7
         Left            =   2670
         TabIndex        =   35
         Tag             =   "0"
         Top             =   525
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Linedef Flag"
         Height          =   255
         Index           =   8
         Left            =   2670
         TabIndex        =   36
         Tag             =   "0"
         Top             =   780
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Linedef Flag"
         Height          =   255
         Index           =   9
         Left            =   2670
         TabIndex        =   37
         Tag             =   "0"
         Top             =   1035
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Linedef Flag"
         Height          =   255
         Index           =   10
         Left            =   2670
         TabIndex        =   38
         Tag             =   "0"
         Top             =   1290
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Linedef Flag"
         Height          =   255
         Index           =   11
         Left            =   2670
         TabIndex        =   39
         Tag             =   "0"
         Top             =   1545
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Linedef Flag"
         Height          =   255
         Index           =   12
         Left            =   5115
         TabIndex        =   40
         Tag             =   "0"
         Top             =   270
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Linedef Flag"
         Height          =   255
         Index           =   13
         Left            =   5115
         TabIndex        =   41
         Tag             =   "0"
         Top             =   525
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Linedef Flag"
         Height          =   255
         Index           =   14
         Left            =   5115
         TabIndex        =   42
         Tag             =   "0"
         Top             =   780
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Linedef Flag"
         Height          =   255
         Index           =   15
         Left            =   5115
         TabIndex        =   43
         Tag             =   "0"
         Top             =   1035
         Visible         =   0   'False
         Width           =   2325
      End
   End
   Begin VB.Frame fraSide2 
      Caption         =   "            "
      Height          =   2085
      Left            =   240
      TabIndex        =   50
      Top             =   2760
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CheckBox chkSetDefaultS2 
         Caption         =   "Set as build defaults"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4140
         TabIndex        =   17
         Top             =   1740
         Width           =   2175
      End
      Begin VB.TextBox txtS2Lower 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6420
         MaxLength       =   8
         TabIndex        =   16
         Text            =   " "
         Top             =   1350
         Width           =   1020
      End
      Begin VB.TextBox txtS2Middle 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5280
         MaxLength       =   8
         TabIndex        =   15
         Text            =   " "
         Top             =   1350
         Width           =   1020
      End
      Begin VB.TextBox txtS2Upper 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4140
         MaxLength       =   8
         TabIndex        =   14
         Text            =   " "
         Top             =   1350
         Width           =   1020
      End
      Begin VB.CheckBox chkBackSide 
         Caption         =   "Back Side"
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   0
         Width           =   1065
      End
      Begin VB.PictureBox picS2Lower 
         BackColor       =   &H8000000C&
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         Enabled         =   0   'False
         HasDC           =   0   'False
         Height          =   1020
         Left            =   6420
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   53
         TabStop         =   0   'False
         ToolTipText     =   "Back Side Lower Texture"
         Top             =   270
         Width           =   1020
         Begin VB.Image imgS2Lower 
            Height          =   960
            Left            =   0
            Stretch         =   -1  'True
            ToolTipText     =   "Back Side Lower Texture"
            Top             =   0
            Width           =   960
         End
      End
      Begin VB.PictureBox picS2Middle 
         BackColor       =   &H8000000C&
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         Enabled         =   0   'False
         HasDC           =   0   'False
         Height          =   1020
         Left            =   5280
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   52
         TabStop         =   0   'False
         ToolTipText     =   "Back Side Middle Texture"
         Top             =   270
         Width           =   1020
         Begin VB.Image imgS2Middle 
            Height          =   960
            Left            =   0
            Stretch         =   -1  'True
            ToolTipText     =   "Back Side Middle Texture"
            Top             =   0
            Width           =   960
         End
      End
      Begin VB.PictureBox picS2Upper 
         BackColor       =   &H8000000C&
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         Enabled         =   0   'False
         HasDC           =   0   'False
         Height          =   1020
         Left            =   4140
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   51
         TabStop         =   0   'False
         ToolTipText     =   "Back Side Upper Texture"
         Top             =   270
         Width           =   1020
         Begin VB.Image imgS2Upper 
            Height          =   960
            Left            =   0
            Stretch         =   -1  'True
            ToolTipText     =   "Back Side Upper Texture"
            Top             =   0
            Width           =   960
         End
      End
      Begin VB.CommandButton cmdS2Offset 
         Caption         =   "Visual Offset..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   1455
         TabIndex        =   13
         Top             =   1245
         Width           =   1740
      End
      Begin DoomBuilder.ctlValueBox txtS2Sector 
         Height          =   360
         Left            =   1455
         TabIndex        =   10
         Top             =   330
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   635
         MaxLength       =   5
         EmptyAllowed    =   -1  'True
      End
      Begin DoomBuilder.ctlValueBox txtS2OffsetX 
         Height          =   360
         Left            =   1455
         TabIndex        =   11
         Top             =   795
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   635
         MaxLength       =   5
         Min             =   -32767
         EmptyAllowed    =   -1  'True
         RelativeAllowed =   -1  'True
      End
      Begin DoomBuilder.ctlValueBox txtS2OffsetY 
         Height          =   360
         Left            =   2370
         TabIndex        =   12
         Top             =   795
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   635
         MaxLength       =   5
         Min             =   -32767
         EmptyAllowed    =   -1  'True
         RelativeAllowed =   -1  'True
      End
      Begin VB.Label lblLength2 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   210
         Left            =   1500
         TabIndex        =   73
         Top             =   1740
         Width           =   60
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Linedef length:"
         Height          =   210
         Left            =   210
         TabIndex        =   72
         Top             =   1740
         Width           =   1065
      End
      Begin VB.Label lblS2Sector 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sector:"
         Enabled         =   0   'False
         Height          =   210
         Left            =   750
         TabIndex        =   55
         Top             =   405
         UseMnemonic     =   0   'False
         Width           =   525
      End
      Begin VB.Label lblS2Offset 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Texture Offset:"
         Enabled         =   0   'False
         Height          =   210
         Left            =   165
         TabIndex        =   54
         Top             =   870
         UseMnemonic     =   0   'False
         Width           =   1110
      End
   End
   Begin VB.Frame fraSide1 
      Caption         =   "            "
      Height          =   2085
      Left            =   240
      TabIndex        =   56
      Top             =   555
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CheckBox chkSetDefaultS1 
         Caption         =   "Set as build defaults"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4140
         TabIndex        =   8
         Top             =   1740
         Width           =   2175
      End
      Begin VB.TextBox txtS1Lower 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6420
         MaxLength       =   8
         TabIndex        =   7
         Text            =   " "
         Top             =   1350
         Width           =   1020
      End
      Begin VB.TextBox txtS1Middle 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5280
         MaxLength       =   8
         TabIndex        =   6
         Text            =   " "
         Top             =   1350
         Width           =   1020
      End
      Begin VB.TextBox txtS1Upper 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4140
         MaxLength       =   8
         TabIndex        =   5
         Text            =   " "
         Top             =   1350
         Width           =   1020
      End
      Begin VB.CommandButton cmdS1Offset 
         Caption         =   "Visual Offset..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   1455
         TabIndex        =   4
         Top             =   1245
         Width           =   1740
      End
      Begin VB.PictureBox picS1Upper 
         BackColor       =   &H8000000C&
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         Enabled         =   0   'False
         HasDC           =   0   'False
         Height          =   1020
         Left            =   4140
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   59
         TabStop         =   0   'False
         ToolTipText     =   "Front Side Upper Texture"
         Top             =   270
         Width           =   1020
         Begin VB.Image imgS1Upper 
            Height          =   960
            Left            =   0
            Stretch         =   -1  'True
            ToolTipText     =   "Front Side Upper Texture"
            Top             =   0
            Width           =   960
         End
      End
      Begin VB.PictureBox picS1Middle 
         BackColor       =   &H8000000C&
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         Enabled         =   0   'False
         HasDC           =   0   'False
         Height          =   1020
         Left            =   5280
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   58
         TabStop         =   0   'False
         ToolTipText     =   "Front Side Middle Texture"
         Top             =   270
         Width           =   1020
         Begin VB.Image imgS1Middle 
            Height          =   960
            Left            =   0
            Stretch         =   -1  'True
            ToolTipText     =   "Front Side Middle Texture"
            Top             =   0
            Width           =   960
         End
      End
      Begin VB.PictureBox picS1Lower 
         BackColor       =   &H8000000C&
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         Enabled         =   0   'False
         HasDC           =   0   'False
         Height          =   1020
         Left            =   6420
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   57
         TabStop         =   0   'False
         ToolTipText     =   "Front Side Lower Texture"
         Top             =   270
         Width           =   1020
         Begin VB.Image imgS1Lower 
            Height          =   960
            Left            =   0
            Stretch         =   -1  'True
            ToolTipText     =   "Front Side Lower Texture"
            Top             =   0
            Width           =   960
         End
      End
      Begin VB.CheckBox chkFrontSide 
         Caption         =   "Front Side"
         Height          =   255
         Left            =   180
         TabIndex        =   0
         Top             =   0
         Width           =   1065
      End
      Begin DoomBuilder.ctlValueBox txtS1Sector 
         Height          =   360
         Left            =   1455
         TabIndex        =   1
         Top             =   330
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   635
         MaxLength       =   5
         EmptyAllowed    =   -1  'True
      End
      Begin DoomBuilder.ctlValueBox txtS1OffsetX 
         Height          =   360
         Left            =   1455
         TabIndex        =   2
         Top             =   795
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   635
         MaxLength       =   5
         Min             =   -32767
         EmptyAllowed    =   -1  'True
         RelativeAllowed =   -1  'True
      End
      Begin DoomBuilder.ctlValueBox txtS1OffsetY 
         Height          =   360
         Left            =   2370
         TabIndex        =   3
         Top             =   795
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   635
         MaxLength       =   5
         Min             =   -32767
         EmptyAllowed    =   -1  'True
         RelativeAllowed =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Linedef length:"
         Height          =   210
         Left            =   210
         TabIndex        =   75
         Top             =   1740
         Width           =   1065
      End
      Begin VB.Label lblLength1 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   210
         Left            =   1500
         TabIndex        =   74
         Top             =   1740
         Width           =   60
      End
      Begin VB.Label lblS1Offset 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Texture Offset:"
         Enabled         =   0   'False
         Height          =   210
         Left            =   165
         TabIndex        =   61
         Top             =   870
         UseMnemonic     =   0   'False
         Width           =   1110
      End
      Begin VB.Label lblS1Sector 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sector:"
         Enabled         =   0   'False
         Height          =   210
         Left            =   750
         TabIndex        =   60
         Top             =   405
         UseMnemonic     =   0   'False
         Width           =   525
      End
   End
   Begin MSComctlLib.TabStrip tbsPanel 
      Height          =   5265
      Left            =   90
      TabIndex        =   47
      Top             =   105
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   9287
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
            Caption         =   "Sidedefs"
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
      Left            =   600
      TabIndex        =   45
      Top             =   7140
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   900
   End
End
Attribute VB_Name = "frmLinedef"
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

Private Function CheckLinedefActivation() As Long
     Dim i As Long
     Dim a As Long
     Dim Indices As Variant
     Dim ATypes As Variant
     Dim la As Long
     
     'Get selection indices
     Indices = selected.Items
     
     'Get activation types
     ATypes = mapconfig("linedefactivations").Keys
     
     'Set result to first linedef's activation
     For a = 0 To (mapconfig("linedefactivations").Count - 1)
          If (linedefs(Indices(LBound(Indices))).Flags And Val(ATypes(a))) = Val(ATypes(a)) Then
               CheckLinedefActivation = a
          End If
     Next a
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check for other activation type
          For a = 0 To (mapconfig("linedefactivations").Count - 1)
               If (linedefs(Indices(i)).Flags And Val(ATypes(a))) = Val(ATypes(a)) Then
                    la = a
               End If
          Next a
          
          'Check if different
          If (la <> CheckLinedefActivation) Then
               CheckLinedefActivation = -1
               Exit For
          End If
     Next i
End Function

Private Function CheckLinedefArg0() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set to first line arg
     CheckLinedefArg0 = linedefs(Indices(LBound(Indices))).arg0
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if different
          If (linedefs(Indices(i)).arg0 <> CheckLinedefArg0) Then
               CheckLinedefArg0 = ""
               Exit For
          End If
     Next i
End Function

Private Function CheckLinedefArg1() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set to first line arg
     CheckLinedefArg1 = linedefs(Indices(LBound(Indices))).arg1
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if different
          If (linedefs(Indices(i)).arg1 <> CheckLinedefArg1) Then
               CheckLinedefArg1 = ""
               Exit For
          End If
     Next i
End Function

Private Function CheckLinedefArg2() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set to first line arg
     CheckLinedefArg2 = linedefs(Indices(LBound(Indices))).arg2
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if different
          If (linedefs(Indices(i)).arg2 <> CheckLinedefArg2) Then
               CheckLinedefArg2 = ""
               Exit For
          End If
     Next i
End Function

Private Function CheckLinedefArg3() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set to first line arg
     CheckLinedefArg3 = linedefs(Indices(LBound(Indices))).arg3
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if different
          If (linedefs(Indices(i)).arg3 <> CheckLinedefArg3) Then
               CheckLinedefArg3 = ""
               Exit For
          End If
     Next i
End Function

Private Function CheckLinedefArg4() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set to first line arg
     CheckLinedefArg4 = linedefs(Indices(LBound(Indices))).arg4
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if different
          If (linedefs(Indices(i)).arg4 <> CheckLinedefArg4) Then
               CheckLinedefArg4 = ""
               Exit For
          End If
     Next i
End Function

'Private FirstLinedef As Long

Private Function CheckLinedefFlag(ByRef Flag As Long) As Long
     Dim i As Long
     Dim Indices As Variant
     Dim Numchecked As Long
     
     'Go for all selected linedefs
     Indices = selected.Items
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if the flag is set
          If ((linedefs(Indices(i)).Flags And Flag) = Flag) Then Numchecked = Numchecked + 1
     Next i
     
     'Check what result to return
     If (Numchecked = 0) Then
          CheckLinedefFlag = vbUnchecked
     ElseIf (Numchecked = numselected) Then
          CheckLinedefFlag = vbChecked
     Else
          CheckLinedefFlag = vbGrayed
     End If
End Function

Private Function CheckLinedefSide1() As Long
     Dim i As Long
     Dim Indices As Variant
     Dim fs As Boolean
     
     'Get selected items
     Indices = selected.Items
     fs = (linedefs(Indices(LBound(Indices))).s1 > -1)
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if the sidedef is set
          If ((fs = (linedefs(Indices(i)).s1 > -1)) = False) Then
               
               'Sidedef is different
               CheckLinedefSide1 = vbGrayed
               Exit Function
          End If
     Next i
     
     'All are the same
     CheckLinedefSide1 = Abs(fs)
End Function

Private Function CheckLinedefSide1Lower() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first one
     For i = LBound(Indices) To UBound(Indices)
          If (linedefs(Indices(i)).s1 > -1) Then
               CheckLinedefSide1Lower = sidedefs(linedefs(Indices(i)).s1).Lower
               Exit For
          End If
     Next i
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if we should verify this sidedef
          If (linedefs(Indices(i)).s1 > -1) Then
               
               'Check if the type is different
               If (sidedefs(linedefs(Indices(i)).s1).Lower <> CheckLinedefSide1Lower) Then
                    
                    'Return nothing
                    CheckLinedefSide1Lower = ""
                    Exit Function
               End If
          End If
     Next i
End Function

Private Function CheckLinedefSide1Middle() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first one
     For i = LBound(Indices) To UBound(Indices)
          If (linedefs(Indices(i)).s1 > -1) Then
               CheckLinedefSide1Middle = sidedefs(linedefs(Indices(i)).s1).Middle
               Exit For
          End If
     Next i
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if we should verify this sidedef
          If (linedefs(Indices(i)).s1 > -1) Then
               
               'Check if the type is different
               If (sidedefs(linedefs(Indices(i)).s1).Middle <> CheckLinedefSide1Middle) Then
                    
                    'Return nothing
                    CheckLinedefSide1Middle = ""
                    Exit Function
               End If
          End If
     Next i
End Function

Private Function CheckLinedefSide1OffsetX() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first one
     For i = LBound(Indices) To UBound(Indices)
          If (linedefs(Indices(i)).s1 > -1) Then
               CheckLinedefSide1OffsetX = sidedefs(linedefs(Indices(i)).s1).tx
               Exit For
          End If
     Next i
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if we should verify this sidedef
          If (linedefs(Indices(i)).s1 > -1) Then
               
               'Check if the type is different
               If (sidedefs(linedefs(Indices(i)).s1).tx <> CheckLinedefSide1OffsetX) Then
                    
                    'Return nothing
                    CheckLinedefSide1OffsetX = ""
                    Exit Function
               End If
          End If
     Next i
End Function

Private Function CheckLinedefSide1OffsetY() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first one
     For i = LBound(Indices) To UBound(Indices)
          If (linedefs(Indices(i)).s1 > -1) Then
               CheckLinedefSide1OffsetY = sidedefs(linedefs(Indices(i)).s1).ty
               Exit For
          End If
     Next i
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if we should verify this sidedef
          If (linedefs(Indices(i)).s1 > -1) Then
               
               'Check if the type is different
               If (sidedefs(linedefs(Indices(i)).s1).ty <> CheckLinedefSide1OffsetY) Then
                    
                    'Return nothing
                    CheckLinedefSide1OffsetY = ""
                    Exit Function
               End If
          End If
     Next i
End Function

Private Function CheckLinedefSide1Sector() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first one
     For i = LBound(Indices) To UBound(Indices)
          If (linedefs(Indices(i)).s1 > -1) Then
               CheckLinedefSide1Sector = sidedefs(linedefs(Indices(i)).s1).sector
               Exit For
          End If
     Next i
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if we should verify this sidedef
          If (linedefs(Indices(i)).s1 > -1) Then
               
               'Check if the type is different
               If (sidedefs(linedefs(Indices(i)).s1).sector <> CheckLinedefSide1Sector) Then
                    
                    'Return nothing
                    CheckLinedefSide1Sector = ""
                    Exit Function
               End If
          End If
     Next i
End Function

Private Function CheckLinedefSide1Upper() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first one
     For i = LBound(Indices) To UBound(Indices)
          If (linedefs(Indices(i)).s1 > -1) Then
               CheckLinedefSide1Upper = sidedefs(linedefs(Indices(i)).s1).Upper
               Exit For
          End If
     Next i
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if we should verify this sidedef
          If (linedefs(Indices(i)).s1 > -1) Then
               
               'Check if the type is different
               If (sidedefs(linedefs(Indices(i)).s1).Upper <> CheckLinedefSide1Upper) Then
                    
                    'Return nothing
                    CheckLinedefSide1Upper = ""
                    Exit Function
               End If
          End If
     Next i
End Function

Private Function CheckLinedefSide2() As Long
     Dim i As Long
     Dim Indices As Variant
     Dim fs As Boolean
     
     'Get selected items
     Indices = selected.Items
     fs = (linedefs(Indices(LBound(Indices))).s2 > -1)
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if the sidedef is set
          If ((fs = (linedefs(Indices(i)).s2 > -1)) = False) Then
               
               'Sidedef is different
               CheckLinedefSide2 = vbGrayed
               Exit Function
          End If
     Next i
     
     'All are the same
     CheckLinedefSide2 = Abs(fs)
End Function

Private Function CheckLinedefSide2Lower() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first one
     For i = LBound(Indices) To UBound(Indices)
          If (linedefs(Indices(i)).s2 > -1) Then
               CheckLinedefSide2Lower = sidedefs(linedefs(Indices(i)).s2).Lower
               Exit For
          End If
     Next i
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if we should verify this sidedef
          If (linedefs(Indices(i)).s2 > -1) Then
               
               'Check if the type is different
               If (sidedefs(linedefs(Indices(i)).s2).Lower <> CheckLinedefSide2Lower) Then
                    
                    'Return nothing
                    CheckLinedefSide2Lower = ""
                    Exit Function
               End If
          End If
     Next i
End Function

Private Function CheckLinedefSide2Middle() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first one
     For i = LBound(Indices) To UBound(Indices)
          If (linedefs(Indices(i)).s2 > -1) Then
               CheckLinedefSide2Middle = sidedefs(linedefs(Indices(i)).s2).Middle
               Exit For
          End If
     Next i
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if we should verify this sidedef
          If (linedefs(Indices(i)).s2 > -1) Then
               
               'Check if the type is different
               If (sidedefs(linedefs(Indices(i)).s2).Middle <> CheckLinedefSide2Middle) Then
                    
                    'Return nothing
                    CheckLinedefSide2Middle = ""
                    Exit Function
               End If
          End If
     Next i
End Function

Private Function CheckLinedefSide2OffsetX() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first one
     For i = LBound(Indices) To UBound(Indices)
          If (linedefs(Indices(i)).s2 > -1) Then
               CheckLinedefSide2OffsetX = sidedefs(linedefs(Indices(i)).s2).tx
               Exit For
          End If
     Next i
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if we should verify this sidedef
          If (linedefs(Indices(i)).s2 > -1) Then
               
               'Check if the type is different
               If (sidedefs(linedefs(Indices(i)).s2).tx <> CheckLinedefSide2OffsetX) Then
                    
                    'Return nothing
                    CheckLinedefSide2OffsetX = ""
                    Exit Function
               End If
          End If
     Next i
End Function

Private Function CheckLinedefSide2OffsetY() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first one
     For i = LBound(Indices) To UBound(Indices)
          If (linedefs(Indices(i)).s2 > -1) Then
               CheckLinedefSide2OffsetY = sidedefs(linedefs(Indices(i)).s2).ty
               Exit For
          End If
     Next i
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if we should verify this sidedef
          If (linedefs(Indices(i)).s2 > -1) Then
               
               'Check if the type is different
               If (sidedefs(linedefs(Indices(i)).s2).ty <> CheckLinedefSide2OffsetY) Then
                    
                    'Return nothing
                    CheckLinedefSide2OffsetY = ""
                    Exit Function
               End If
          End If
     Next i
End Function

Private Function CheckLinedefSide2Sector() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first one
     For i = LBound(Indices) To UBound(Indices)
          If (linedefs(Indices(i)).s2 > -1) Then
               CheckLinedefSide2Sector = sidedefs(linedefs(Indices(i)).s2).sector
               Exit For
          End If
     Next i
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if we should verify this sidedef
          If (linedefs(Indices(i)).s2 > -1) Then
               
               'Check if the type is different
               If (sidedefs(linedefs(Indices(i)).s2).sector <> CheckLinedefSide2Sector) Then
                    
                    'Return nothing
                    CheckLinedefSide2Sector = ""
                    Exit Function
               End If
          End If
     Next i
End Function

Private Function CheckLinedefSide2Upper() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first one
     For i = LBound(Indices) To UBound(Indices)
          If (linedefs(Indices(i)).s2 > -1) Then
               CheckLinedefSide2Upper = sidedefs(linedefs(Indices(i)).s2).Upper
               Exit For
          End If
     Next i
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if we should verify this sidedef
          If (linedefs(Indices(i)).s2 > -1) Then
               
               'Check if the type is different
               If (sidedefs(linedefs(Indices(i)).s2).Upper <> CheckLinedefSide2Upper) Then
                    
                    'Return nothing
                    CheckLinedefSide2Upper = ""
                    Exit Function
               End If
          End If
     Next i
End Function

Private Function CheckLinedefTag() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first linedef's tag
     CheckLinedefTag = linedefs(Indices(LBound(Indices))).tag
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if the type is different
          If (linedefs(Indices(i)).tag <> CheckLinedefTag) Then
               CheckLinedefTag = ""
               Exit Function
          End If
     Next i
End Function

Private Function CheckLinedefType() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first linedef's type
     CheckLinedefType = linedefs(Indices(LBound(Indices))).effect
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if the type is different
          If (linedefs(Indices(i)).effect <> CheckLinedefType) Then
               CheckLinedefType = ""
               Exit Function
          End If
     Next i
End Function

Private Sub chkBackSide_Click()
     
     'Check if checkbox is checked
     If (chkBackSide.Value = vbChecked) Then
          
          'Enabled Controls
          lblS2Sector.Enabled = True
          lblS2Offset.Enabled = True
          txtS2Sector.Enabled = True
          txtS2OffsetX.Enabled = True
          txtS2OffsetY.Enabled = True
          cmdS2Offset.Enabled = True
          txtS2Upper.Enabled = True
          txtS2Middle.Enabled = True
          txtS2Lower.Enabled = True
          picS2Upper.Enabled = True
          picS2Middle.Enabled = True
          picS2Lower.Enabled = True
          chkSetDefaultS2.Enabled = True
          
          'Update texture previews
          txtS2Lower_Change
          txtS2Middle_Change
          txtS2Upper_Change
     Else
          
          'Disabled Controls
          lblS2Sector.Enabled = False
          lblS2Offset.Enabled = False
          txtS2Sector.Enabled = False
          txtS2OffsetX.Enabled = False
          txtS2OffsetY.Enabled = False
          cmdS2Offset.Enabled = False
          txtS2Upper.Enabled = False
          txtS2Middle.Enabled = False
          txtS2Lower.Enabled = False
          picS2Upper.Enabled = False
          picS2Middle.Enabled = False
          picS2Lower.Enabled = False
          chkSetDefaultS2.Enabled = False
          
          'Remove texture previews
          Set picS2Lower.Picture = Nothing
          Set picS2Middle.Picture = Nothing
          Set picS2Upper.Picture = Nothing
     End If
     
     'Check if we should show Sector Warning
     picSectorWarning.visible = (((chkFrontSide.tag <> vbChecked) And (chkFrontSide.Value = vbChecked)) Or ((chkBackSide.tag <> vbChecked) And (chkBackSide.Value = vbChecked)) And (numselected > 1))
     
     'Empty sector number allowed?
     txtS1Sector.EmptyAllowed = Not ((chkFrontSide.tag <> vbChecked) And (chkFrontSide.Value = vbChecked))
     txtS2Sector.EmptyAllowed = Not ((chkBackSide.tag <> vbChecked) And (chkBackSide.Value = vbChecked))
     If (txtS1Sector.EmptyAllowed = False) And (Trim$(txtS1Sector.Text) = "") Then txtS1Sector.Value = txtS1Sector.Max: txtS1Sector.Value = txtS1Sector.Min
     If (txtS2Sector.EmptyAllowed = False) And (Trim$(txtS2Sector.Text) = "") Then txtS2Sector.Value = txtS2Sector.Max: txtS2Sector.Value = txtS2Sector.Min
End Sub

Private Sub chkFrontSide_Click()
     
     'Check if checkbox is checked
     If (chkFrontSide.Value = vbChecked) Then
          
          'Enabled Controls
          lblS1Sector.Enabled = True
          lblS1Offset.Enabled = True
          txtS1Sector.Enabled = True
          txtS1OffsetX.Enabled = True
          txtS1OffsetY.Enabled = True
          cmdS1Offset.Enabled = True
          txtS1Upper.Enabled = True
          txtS1Middle.Enabled = True
          txtS1Lower.Enabled = True
          picS1Upper.Enabled = True
          picS1Middle.Enabled = True
          picS1Lower.Enabled = True
          chkSetDefaultS1.Enabled = True
          
          'Update texture previews
          txtS1Lower_Change
          txtS1Middle_Change
          txtS1Upper_Change
     Else
          
          'Disabled Controls
          lblS1Sector.Enabled = False
          lblS1Offset.Enabled = False
          txtS1Sector.Enabled = False
          txtS1OffsetX.Enabled = False
          txtS1OffsetY.Enabled = False
          cmdS1Offset.Enabled = False
          txtS1Upper.Enabled = False
          txtS1Middle.Enabled = False
          txtS1Lower.Enabled = False
          picS1Upper.Enabled = False
          picS1Middle.Enabled = False
          picS1Lower.Enabled = False
          chkSetDefaultS1.Enabled = False
          
          'Remove texture previews
          Set picS1Lower.Picture = Nothing
          Set picS1Middle.Picture = Nothing
          Set picS1Upper.Picture = Nothing
     End If
     
     'Check if we should show Sector Warning
     picSectorWarning.visible = (((chkFrontSide.tag <> vbChecked) And (chkFrontSide.Value = vbChecked)) Or ((chkBackSide.tag <> vbChecked) And (chkBackSide.Value = vbChecked)) And (numselected > 1))
     
     'Empty sector number allowed?
     txtS1Sector.EmptyAllowed = Not ((chkFrontSide.tag <> vbChecked) And (chkFrontSide.Value = vbChecked))
     txtS2Sector.EmptyAllowed = Not ((chkBackSide.tag <> vbChecked) And (chkBackSide.Value = vbChecked))
     If (txtS1Sector.EmptyAllowed = False) And (Trim$(txtS1Sector.Text) = "") Then txtS1Sector.Value = txtS1Sector.Max: txtS1Sector.Value = txtS1Sector.Min
     If (txtS2Sector.EmptyAllowed = False) And (Trim$(txtS2Sector.Text) = "") Then txtS2Sector.Value = txtS2Sector.Max: txtS2Sector.Value = txtS2Sector.Min
End Sub

Private Sub chkSetDefaultS1_Click()
     
     'Only one of two can be checked
     If (chkSetDefaultS1.Value = vbChecked) Then chkSetDefaultS2.Value = vbUnchecked
End Sub

Private Sub chkSetDefaultS2_Click()
     
     'Only one of two can be checked
     If (chkSetDefaultS2.Value = vbChecked) Then chkSetDefaultS1.Value = vbUnchecked
End Sub

Private Sub cmdCancel_Click()
     Unload Me
     Set frmLinedef = Nothing
End Sub

Private Sub cmdNextTag_Click()
     txtTag.Text = NextUnusedTag
End Sub

Private Sub cmdOK_Click()
     Dim Indices As Variant
     Dim i As Long
     Dim f As Long
     Dim ld As Long
     Dim a As Long
     Dim ATypes As Variant
     
     'Change mousepointer
     Screen.MousePointer = vbHourglass
     
     'This will validate the last focused control
     DoEvents
     
     'Get activation types
     ATypes = mapconfig("linedefactivations").Keys
     
     'Make undo
     If (lblMakeUndo.Caption = "") Then CreateUndo "linedef edit"
     
     'Get the selection indices
     Indices = selected.Items
     
     'Go for all linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Get the linedef index
          ld = Indices(i)
          
          'Go for all individual flags
          For f = 0 To 15
               
               'Check if this flag can be set
               If (chkFlag(f).tag <> "0") Then
                    
                    'Check if the flag is marked to be set
                    If (chkFlag(f).Value = vbChecked) Then
                         
                         'Add the flag on the linedef
                         linedefs(ld).Flags = linedefs(ld).Flags Or CLng(chkFlag(f).tag)
                    
                    'Check if the flag is marked to be removed
                    ElseIf (chkFlag(f).Value = vbUnchecked) Then
                         
                         'Remove the flag from the linedef
                         linedefs(ld).Flags = linedefs(ld).Flags And Not CLng(chkFlag(f).tag)
                    End If
               End If
          Next f
          
          'Apply type if a linedef type is specified
          If (txtType.Text <> "") Then linedefs(ld).effect = Val(txtType.Text)
          
          'Check if this is an action which we know
          If (mapconfig("linedeftypes").Exists(CStr(linedefs(ld).effect)) = True) Then
               
               'Set the marking references on the linedef
               With linedefs(ld)
                    .argref0 = mapconfig("linedeftypes")(CStr(linedefs(ld).effect))("mark1")
                    .argref1 = mapconfig("linedeftypes")(CStr(linedefs(ld).effect))("mark2")
                    .argref2 = mapconfig("linedeftypes")(CStr(linedefs(ld).effect))("mark3")
                    .argref3 = mapconfig("linedeftypes")(CStr(linedefs(ld).effect))("mark4")
                    .argref4 = mapconfig("linedeftypes")(CStr(linedefs(ld).effect))("mark5")
               End With
          End If
          
          'Apply tag if a tag is specified
          If (txtTag.Text <> "") Then linedefs(ld).tag = txtTag.RelativeValue(linedefs(ld).tag)
          
          'Apply activation if an activation is specified
          If (cmbActivation.ListIndex > -1) Then
               
               'Go for all activations to remove
               For a = 0 To (mapconfig("linedefactivations").Count - 1)
                    
                    'Ensure this activation is not set
                    linedefs(ld).Flags = linedefs(ld).Flags And Not Val(ATypes(a))
               Next a
               
               'Set the selected activation
               linedefs(ld).Flags = linedefs(ld).Flags Or Val(ATypes(cmbActivation.ListIndex))
          End If
          
          'Apply arguments if arguments are specified
          If (txtArgument(0).Text <> "") Then linedefs(ld).arg0 = txtArgument(0).RelativeValue(linedefs(ld).arg0)
          If (txtArgument(1).Text <> "") Then linedefs(ld).arg1 = txtArgument(1).RelativeValue(linedefs(ld).arg1)
          If (txtArgument(2).Text <> "") Then linedefs(ld).arg2 = txtArgument(2).RelativeValue(linedefs(ld).arg2)
          If (txtArgument(3).Text <> "") Then linedefs(ld).arg3 = txtArgument(3).RelativeValue(linedefs(ld).arg3)
          If (txtArgument(4).Text <> "") Then linedefs(ld).arg4 = txtArgument(4).RelativeValue(linedefs(ld).arg4)
          
          'Check if we should remove sidedef 1
          If ((linedefs(ld).s1 > -1) And (chkFrontSide.Value = vbUnchecked)) Then
               
               'Remove the sidedef
               RemoveSidedef linedefs(ld).s1, False, , False
               
          'Check if we should create or modify sidedef 1
          ElseIf (chkFrontSide.Value = vbChecked) Then
               
               'Create a sidedef if needed
               If (linedefs(ld).s1 = -1) Then linedefs(ld).s1 = CreateSidedef
               
               'Modify sidedef
               With sidedefs(linedefs(ld).s1)
                    If (txtS1Sector.Text <> "") Then .sector = Val(txtS1Sector.Text): mapnodeschanged = True
                    If (txtS1OffsetX.Text <> "") Then .tx = txtS1OffsetX.RelativeValue(.tx)
                    If (txtS1OffsetY.Text <> "") Then .ty = txtS1OffsetY.RelativeValue(.ty)
                    If (txtS1Upper.Text <> "") Then .Upper = txtS1Upper.Text
                    If (txtS1Middle.Text <> "") Then .Middle = txtS1Middle.Text
                    If (txtS1Lower.Text <> "") Then .Lower = txtS1Lower.Text
                    .linedef = ld
               End With
          End If
          
          
          'Check if we should remove sidedef 2
          If ((linedefs(ld).s2 > -1) And (chkBackSide.Value = vbUnchecked)) Then
               
               'Remove the sidedef
               RemoveSidedef linedefs(ld).s2, False, , False
               
          'Check if we should create or modify sidedef 2
          ElseIf (chkBackSide.Value = vbChecked) Then
               
               'Create a sidedef if needed
               If (linedefs(ld).s2 = -1) Then linedefs(ld).s2 = CreateSidedef
               
               'Modify sidedef
               With sidedefs(linedefs(ld).s2)
                    If (txtS2Sector.Text <> "") Then .sector = Val(txtS2Sector.Text): mapnodeschanged = True
                    If (txtS2OffsetX.Text <> "") Then .tx = txtS2OffsetX.RelativeValue(.tx)
                    If (txtS2OffsetY.Text <> "") Then .ty = txtS2OffsetY.RelativeValue(.ty)
                    If (txtS2Upper.Text <> "") Then .Upper = txtS2Upper.Text
                    If (txtS2Middle.Text <> "") Then .Middle = txtS2Middle.Text
                    If (txtS2Lower.Text <> "") Then .Lower = txtS2Lower.Text
                    .linedef = ld
               End With
          End If
     Next i
     
     'Map is modified
     mapnodeschanged = True
     mapchanged = True
     UpdateStatusBar
     
     'Make build defaults from S1 if requested
     If (chkSetDefaultS1.Value = vbChecked) Then
          
          'Set the build defaults
          If IsTextureName(txtS1Lower.Text) Then Config("defaulttexture")("lower") = txtS1Lower.Text
          If IsTextureName(txtS1Middle.Text) Then Config("defaulttexture")("middle") = txtS1Middle.Text
          If IsTextureName(txtS1Upper.Text) Then Config("defaulttexture")("upper") = txtS1Upper.Text
          
     'Make build defaults from S2 if requested
     ElseIf (chkSetDefaultS2.Value = vbChecked) Then
          
          'Set the build defaults
          If IsTextureName(txtS2Lower.Text) Then Config("defaulttexture")("lower") = txtS2Lower.Text
          If IsTextureName(txtS2Middle.Text) Then Config("defaulttexture")("middle") = txtS2Middle.Text
          If IsTextureName(txtS2Upper.Text) Then Config("defaulttexture")("upper") = txtS2Upper.Text
     End If
     
     'Reset mousepointer
     Screen.MousePointer = vbNormal
     
     'Leave here
     Unload Me
     Set frmLinedef = Nothing
End Sub

Private Sub cmdS1Offset_Click()
     Dim x As Long, y As Long
     
     'Determine start offset
     If (txtS1OffsetX.Text <> "") Then x = -CLng(txtS1OffsetX.Text)
     If (txtS1OffsetY.Text <> "") Then y = -CLng(txtS1OffsetY.Text)
     
     'Load the dialog
     Load frmTextureOffset
     
     'Initialize
     frmTextureOffset.Init txtS1Upper, txtS1Middle, txtS1Lower, x, y
     
     'Show dialog
     frmTextureOffset.Show 1, Me
     
     'Check the result
     If (frmTextureOffset.tag = "1") Then
          
          'Take the offsets
          txtS1OffsetX.Text = -frmTextureOffset.OffsetX
          txtS1OffsetY.Text = -frmTextureOffset.OffsetY
     End If
     
     'Unload the dialog
     Unload frmTextureOffset: Set frmTextureOffset = Nothing
End Sub

Private Sub cmdS2Offset_Click()
     Dim x As Long, y As Long
     
     'Determine start offset
     If (txtS2OffsetX.Text <> "") Then x = -CLng(txtS2OffsetX.Text)
     If (txtS2OffsetY.Text <> "") Then y = -CLng(txtS2OffsetY.Text)
     
     'Load the dialog
     Load frmTextureOffset
     
     'Initialize
     frmTextureOffset.Init txtS2Upper, txtS2Middle, txtS2Lower, x, y
     
     'Show dialog
     frmTextureOffset.Show 1, Me
     
     'Check the result
     If (frmTextureOffset.tag = "1") Then
          
          'Take the offsets
          txtS2OffsetX.Text = -frmTextureOffset.OffsetX
          txtS2OffsetY.Text = -frmTextureOffset.OffsetY
     End If
     
     'Unload the dialog
     Unload frmTextureOffset: Set frmTextureOffset = Nothing
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
     Dim i As Long
     Dim nflag As Long
     Dim xl As Long, yl As Long
     
     'Change mousepointer
     Screen.MousePointer = vbHourglass
     
     'Keep first selected linedef
     'FirstLinedef = selected.Items(0)
     
     'Check if only one selected
     If (numselected = 1) Then
          
          'Show caption
          Caption = Caption & " - Linedef " & selected.Items(0)
          
          'Calculate linedef length
          xl = vertexes(linedefs(selected.Items(0)).v2).x - vertexes(linedefs(selected.Items(0)).v1).x
          yl = vertexes(linedefs(selected.Items(0)).v2).y - vertexes(linedefs(selected.Items(0)).v1).y
          lblLength1.Caption = CLng(Sqr(xl * xl + yl * yl))
          lblLength2.Caption = lblLength1.Caption
     End If
     
     'Determine what linedef arguments to show
     Select Case mapconfig("mapformat")
          
          'Doom map format
          Case MFMT_DOOM
               For i = 0 To 4
                    lblArgument(i).visible = False
                    txtArgument(i).visible = False
               Next i
               lblTag.visible = True
               txtTag.visible = True
               cmdNextTag.visible = True
               lblActivation.visible = False
               cmbActivation.visible = False
               
          'Hexen map format
          Case MFMT_HEXEN
               For i = 0 To 4
                    lblArgument(i).visible = True
                    txtArgument(i).visible = True
               Next i
               lblTag.visible = False
               txtTag.visible = False
               cmdNextTag.visible = False
               lblActivation.visible = True
               cmbActivation.visible = True
               
               'Go for all activations
               Keys = mapconfig("linedefactivations").Keys
               For i = 0 To (mapconfig("linedefactivations").Count - 1)
                    
                    'Add to list
                    cmbActivation.AddItem mapconfig("linedefactivations")(Keys(i))
                    
                    'Set itemdata to flag value
                    cmbActivation.ItemData(cmbActivation.NewIndex) = Val(Keys(i))
               Next i
               
               'Activation
               i = CheckLinedefActivation
               If (i < cmbActivation.ListCount) Then cmbActivation.ListIndex = i
               
     End Select
     
     'Go for all flags
     Keys = mapconfig("linedefflags").Keys
     For i = 0 To 15
          
          'Check if this flag is known
          If (mapconfig("linedefflags").Exists(CStr(2 ^ i)) = True) Then
               
               'Check if not unset
               If CStr(mapconfig("linedefflags")(CStr(2 ^ i))) <> "0" Then
                    
                    'Set the checkbox properties
                    chkFlag(nflag).tag = CStr(2 ^ i)
                    chkFlag(nflag).visible = True
                    chkFlag(nflag).Caption = mapconfig("linedefflags")(CStr(2 ^ i))
                    
                    'Check this flag
                    chkFlag(nflag).Value = CheckLinedefFlag(2 ^ i)
                    
                    'Next flag
                    nflag = nflag + 1
               End If
          End If
     Next i
     
     'Set the type
     txtType.Text = CheckLinedefType
     txtType_Change
     If (Trim$(txtType.Text) <> "") Then
          
          'Arguments
          txtArgument(0).Value = CheckLinedefArg0
          txtArgument(1).Value = CheckLinedefArg1
          txtArgument(2).Value = CheckLinedefArg2
          txtArgument(3).Value = CheckLinedefArg3
          txtArgument(4).Value = CheckLinedefArg4
     Else
          
          'Args are not the same
          txtArgument(0).Text = ""
          txtArgument(1).Text = ""
          txtArgument(2).Text = ""
          txtArgument(3).Text = ""
          txtArgument(4).Text = ""
     End If
     
     'Linedef tag
     txtTag.Text = CheckLinedefTag
     
     'Sidedef checkboxes
     chkFrontSide.tag = CheckLinedefSide1
     chkBackSide.tag = CheckLinedefSide2
     chkFrontSide.Value = CheckLinedefSide1
     chkBackSide.Value = CheckLinedefSide2
     chkFrontSide_Click
     chkBackSide_Click
     
     'Front sidedef properties
     txtS1Sector.Text = CheckLinedefSide1Sector
     txtS1OffsetX.Text = CheckLinedefSide1OffsetX
     txtS1OffsetY.Text = CheckLinedefSide1OffsetY
     txtS1Upper.tag = CLng(LinedefsSide1UpperRequired)
     txtS1Middle.tag = CLng(LinedefsSide1MiddleRequired)
     txtS1Lower.tag = CLng(LinedefsSide1LowerRequired)
     txtS1Upper.Text = CheckLinedefSide1Upper
     txtS1Middle.Text = CheckLinedefSide1Middle
     txtS1Lower.Text = CheckLinedefSide1Lower
     
     'Back sidedef properties
     txtS2Sector.Text = CheckLinedefSide2Sector
     txtS2OffsetX.Text = CheckLinedefSide2OffsetX
     txtS2OffsetY.Text = CheckLinedefSide2OffsetY
     txtS2Upper.tag = CLng(LinedefsSide2UpperRequired)
     txtS2Middle.tag = CLng(LinedefsSide2MiddleRequired)
     txtS2Lower.tag = CLng(LinedefsSide2LowerRequired)
     txtS2Upper.Text = CheckLinedefSide2Upper
     txtS2Middle.Text = CheckLinedefSide2Middle
     txtS2Lower.Text = CheckLinedefSide2Lower
     
     'Limit sector fields to max sectors
     txtS1Sector.Max = numsectors - 1
     txtS2Sector.Max = numsectors - 1
     
     'Set relativescroll property
     'NOTE: This doesnt work nicely
     'txtS1OffsetX.RelativeScroll = (numselected > 1)
     'txtS1OffsetY.RelativeScroll = (numselected > 1)
     'txtS2OffsetX.RelativeScroll = (numselected > 1)
     'txtS2OffsetY.RelativeScroll = (numselected > 1)
     
     'Change mousepointer
     Screen.MousePointer = vbDefault
End Sub

Private Sub imgS1Lower_Click()
     picS1Lower_Click
End Sub

Private Sub imgS1Middle_Click()
     picS1Middle_Click
End Sub

Private Sub imgS1Upper_Click()
     picS1Upper_Click
End Sub

Private Sub imgS2Lower_Click()
     picS2Lower_Click
End Sub

Private Sub imgS2Middle_Click()
     picS2Middle_Click
End Sub

Private Sub imgS2Upper_Click()
     picS2Upper_Click
End Sub

Private Function LinedefsSide1LowerRequired() As Boolean
     Dim i As Long
     Dim Indices As Variant
     
     'Assume true
     LinedefsSide1LowerRequired = True
     
     'Get selection indices
     Indices = selected.Items
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if this linedef requires an upper texture
          If RequiresS1Lower(Indices(i)) = False Then
               
               'Not all lines require this
               LinedefsSide1LowerRequired = False
               Exit For
          End If
     Next i
End Function

Private Function LinedefsSide1MiddleRequired() As Boolean
     Dim i As Long
     Dim Indices As Variant
     
     'Assume true
     LinedefsSide1MiddleRequired = True
     
     'Get selection indices
     Indices = selected.Items
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if this linedef requires an upper texture
          If RequiresS1Middle(Indices(i)) = False Then
               
               'Not all lines require this
               LinedefsSide1MiddleRequired = False
               Exit For
          End If
     Next i
End Function

Private Function LinedefsSide1UpperRequired() As Boolean
     Dim i As Long
     Dim Indices As Variant
     
     'Assume true
     LinedefsSide1UpperRequired = True
     
     'Get selection indices
     Indices = selected.Items
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if this linedef requires an upper texture
          If RequiresS1Upper(Indices(i)) = False Then
               
               'Not all lines require this
               LinedefsSide1UpperRequired = False
               Exit For
          End If
     Next i
End Function

Private Function LinedefsSide2LowerRequired() As Boolean
     Dim i As Long
     Dim Indices As Variant
     
     'Assume true
     LinedefsSide2LowerRequired = True
     
     'Get selection indices
     Indices = selected.Items
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if this linedef requires an upper texture
          If RequiresS2Lower(Indices(i)) = False Then
               
               'Not all lines require this
               LinedefsSide2LowerRequired = False
               Exit For
          End If
     Next i
End Function

Private Function LinedefsSide2MiddleRequired() As Boolean
     Dim i As Long
     Dim Indices As Variant
     
     'Assume true
     LinedefsSide2MiddleRequired = True
     
     'Get selection indices
     Indices = selected.Items
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if this linedef requires an upper texture
          If RequiresS2Middle(Indices(i)) = False Then
               
               'Not all lines require this
               LinedefsSide2MiddleRequired = False
               Exit For
          End If
     Next i
End Function

Private Function LinedefsSide2UpperRequired() As Boolean
     Dim i As Long
     Dim Indices As Variant
     
     'Assume true
     LinedefsSide2UpperRequired = True
     
     'Get selection indices
     Indices = selected.Items
     
     'Go for all selected linedefs
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if this linedef requires an upper texture
          If RequiresS2Upper(Indices(i)) = False Then
               
               'Not all lines require this
               LinedefsSide2UpperRequired = False
               Exit For
          End If
     Next i
End Function

Private Sub picS1Lower_Click()
     txtS1Lower.Text = SelectTexture(txtS1Lower.Text, Me)
End Sub

Private Sub picS1Middle_Click()
     txtS1Middle.Text = SelectTexture(txtS1Middle.Text, Me)
End Sub

Private Sub picS1Upper_Click()
     txtS1Upper.Text = SelectTexture(txtS1Upper.Text, Me)
End Sub

Private Sub picS2Lower_Click()
     txtS2Lower.Text = SelectTexture(txtS2Lower.Text, Me)
End Sub

Private Sub picS2Middle_Click()
     txtS2Middle.Text = SelectTexture(txtS2Middle.Text, Me)
End Sub

Private Sub picS2Upper_Click()
     txtS2Upper.Text = SelectTexture(txtS2Upper.Text, Me)
End Sub

Private Sub tbsPanel_Click()
     If (tbsPanel.SelectedItem.Index = 1) Then
          fraFlags.visible = True
          fraAction.visible = True
          fraSide1.visible = False
          fraSide2.visible = False
     Else
          fraFlags.visible = False
          fraAction.visible = False
          fraSide1.visible = True
          fraSide2.visible = True
     End If
End Sub

Private Sub txtArgument_GotFocus(Index As Integer)
     SelectAllText txtArgument(Index)
End Sub


Private Sub txtS1Lower_Change()
     
     'Set the texture in the preview box
     GetScaledTexturePicture txtS1Lower.Text, imgS1Lower, , Val(txtS1Lower.tag)
End Sub

Private Sub txtS1Lower_GotFocus()
     SelectAllText txtS1Lower
End Sub


Private Sub txtS1Lower_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtS1Lower_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Complete texture name
     If Val(Config("autocompletetypetex")) Then CompleteTextureName KeyCode, Shift, txtS1Lower
End Sub

Private Sub txtS1Lower_Validate(Cancel As Boolean)
     
     'Find closest match if preferred
     If Val(Config("autocompletetex")) Then txtS1Lower.Text = GetNearestTextureName(txtS1Lower.Text)
End Sub

Private Sub txtS1Middle_Change()
     
     'Set the texture in the preview box
     GetScaledTexturePicture txtS1Middle.Text, imgS1Middle, , Val(txtS1Middle.tag)
End Sub

Private Sub txtS1Middle_GotFocus()
     SelectAllText txtS1Middle
End Sub


Private Sub txtS1Middle_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtS1Middle_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Complete texture name
     If Val(Config("autocompletetypetex")) Then CompleteTextureName KeyCode, Shift, txtS1Middle
End Sub


Private Sub txtS1Middle_Validate(Cancel As Boolean)
     
     'Find closest match if preferred
     If Val(Config("autocompletetex")) Then txtS1Middle.Text = GetNearestTextureName(txtS1Middle.Text)
End Sub

Private Sub txtS1OffsetX_GotFocus()
     SelectAllText txtS1OffsetX
End Sub


Private Sub txtS1OffsetY_GotFocus()
     SelectAllText txtS1OffsetY
End Sub


Private Sub txtS1Sector_GotFocus()
     SelectAllText txtS1Sector
End Sub


Private Sub txtS1Upper_Change()
     
     'Set the texture in the preview box
     GetScaledTexturePicture txtS1Upper.Text, imgS1Upper, , Val(txtS1Upper.tag)
End Sub

Private Sub txtS1Upper_GotFocus()
     SelectAllText txtS1Upper
End Sub


Private Sub txtS1Upper_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtS1Upper_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Complete texture name
     If Val(Config("autocompletetypetex")) Then CompleteTextureName KeyCode, Shift, txtS1Upper
End Sub


Private Sub txtS1Upper_Validate(Cancel As Boolean)
     
     'Find closest match if preferred
     If Val(Config("autocompletetex")) Then txtS1Upper.Text = GetNearestTextureName(txtS1Upper.Text)
End Sub

Private Sub txtS2Lower_Change()
     
     'Set the texture in the preview box
     GetScaledTexturePicture txtS2Lower.Text, imgS2Lower, , Val(txtS2Lower.tag)
End Sub

Private Sub txtS2Lower_GotFocus()
     SelectAllText txtS2Lower
End Sub


Private Sub txtS2Lower_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtS2Lower_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Complete texture name
     If Val(Config("autocompletetypetex")) Then CompleteTextureName KeyCode, Shift, txtS2Lower
End Sub

Private Sub txtS2Lower_Validate(Cancel As Boolean)
     
     'Find closest match if preferred
     If Val(Config("autocompletetex")) Then txtS2Lower.Text = GetNearestTextureName(txtS2Lower.Text)
End Sub

Private Sub txtS2Middle_Change()
     
     'Set the texture in the preview box
     GetScaledTexturePicture txtS2Middle.Text, imgS2Middle, , Val(txtS2Middle.tag)
End Sub

Private Sub txtS2Middle_GotFocus()
     SelectAllText txtS2Middle
End Sub


Private Sub txtS2Middle_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtS2Middle_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Complete texture name
     If Val(Config("autocompletetypetex")) Then CompleteTextureName KeyCode, Shift, txtS2Middle
End Sub

Private Sub txtS2Middle_Validate(Cancel As Boolean)
     
     'Find closest match if preferred
     If Val(Config("autocompletetex")) Then txtS2Middle.Text = GetNearestTextureName(txtS2Middle.Text)
End Sub

Private Sub txtS2OffsetX_GotFocus()
     SelectAllText txtS2OffsetX
End Sub


Private Sub txtS2OffsetY_GotFocus()
     SelectAllText txtS2OffsetY
End Sub


Private Sub txtS2Sector_GotFocus()
     SelectAllText txtS2Sector
End Sub


Private Sub txtS2Upper_Change()
     
     'Set the texture in the preview box
     GetScaledTexturePicture txtS2Upper.Text, imgS2Upper, , Val(txtS2Upper.tag)
End Sub

Private Sub txtS2Upper_GotFocus()
     SelectAllText txtS2Upper
End Sub


Private Sub txtS2Upper_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtS2Upper_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Complete texture name
     If Val(Config("autocompletetypetex")) Then CompleteTextureName KeyCode, Shift, txtS2Upper
End Sub

Private Sub txtS2Upper_Validate(Cancel As Boolean)
     
     'Find closest match if preferred
     If Val(Config("autocompletetex")) Then txtS2Upper.Text = GetNearestTextureName(txtS2Upper.Text)
End Sub

Private Sub txtTag_GotFocus()
     SelectAllText txtTag
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


