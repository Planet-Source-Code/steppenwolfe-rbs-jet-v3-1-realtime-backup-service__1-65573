VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIndex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RBS Jet V3 - Real Time Backup Service"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9210
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIndex.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pbInflate 
      Height          =   255
      Left            =   6660
      TabIndex        =   29
      Top             =   5130
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.PictureBox picBg 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4455
      Index           =   4
      Left            =   -30
      ScaleHeight     =   4425
      ScaleWidth      =   9135
      TabIndex        =   7
      Top             =   600
      Width           =   9165
      Begin VB.TextBox txtAbout 
         BackColor       =   &H80000018&
         Height          =   3735
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   28
         Text            =   "frmIndex.frx":57E2
         Top             =   300
         Width           =   8625
      End
   End
   Begin VB.PictureBox picBg 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4455
      Index           =   0
      Left            =   0
      ScaleHeight     =   4425
      ScaleWidth      =   9135
      TabIndex        =   6
      Top             =   600
      Width           =   9165
      Begin MSComctlLib.ListView lstArchiveCts 
         Height          =   3885
         Left            =   4230
         TabIndex        =   56
         Top             =   330
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   6853
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.TextBox txtExtractPath 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   3270
         Width           =   3705
      End
      Begin VB.CommandButton cmdFileExtract 
         Caption         =   "Delete File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   2190
         TabIndex        =   27
         Top             =   3720
         Width           =   1785
      End
      Begin VB.CommandButton cmdFileExtract 
         Caption         =   "Extract File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   3720
         Width           =   1785
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Deletion From Archive:"
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
         Index           =   11
         Left            =   240
         TabIndex        =   36
         Top             =   1500
         Width           =   1875
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Select the file from the list and click 'Delete File'."
         Height          =   210
         Index           =   10
         Left            =   240
         TabIndex        =   35
         Top             =   1740
         Width           =   3420
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Extraction Steps:"
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
         Index           =   9
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "3) Extract the file to that location."
         Height          =   210
         Index           =   8
         Left            =   240
         TabIndex        =   33
         Top             =   1050
         Width           =   2370
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "2) Choose the Extraction Path."
         Height          =   210
         Index           =   7
         Left            =   240
         TabIndex        =   32
         Top             =   780
         Width           =   2190
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Extract File To:"
         Height          =   210
         Index           =   6
         Left            =   240
         TabIndex        =   31
         Top             =   3060
         Width           =   1065
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "1) Select a File from the Archive list. "
         Height          =   210
         Index           =   5
         Left            =   240
         TabIndex        =   25
         Top             =   510
         Width           =   2655
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Archive Contents"
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
         Index           =   4
         Left            =   4230
         TabIndex        =   24
         Top             =   120
         Width           =   1440
      End
   End
   Begin VB.PictureBox picBg 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4455
      Index           =   3
      Left            =   0
      ScaleHeight     =   4425
      ScaleWidth      =   9135
      TabIndex        =   5
      Top             =   600
      Width           =   9165
      Begin VB.CommandButton cmdIdxChange 
         Caption         =   "Reinstall Index Service"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   6930
         TabIndex        =   55
         Top             =   3720
         Width           =   2145
      End
      Begin VB.CommandButton cmdIdxChange 
         Caption         =   "Uninstall Index Service"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   4680
         TabIndex        =   54
         Top             =   3720
         Width           =   2145
      End
      Begin VB.Frame Frame2 
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   180
         TabIndex        =   10
         Top             =   300
         Width           =   2265
         Begin VB.OptionButton optSchedule 
            Caption         =   "Every 24 Hours"
            Height          =   255
            Index           =   2
            Left            =   150
            TabIndex        =   52
            Top             =   1680
            Width           =   1665
         End
         Begin VB.OptionButton optSchedule 
            Caption         =   "Every 12 Hours"
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   51
            Top             =   1380
            Width           =   1665
         End
         Begin VB.OptionButton optSchedule 
            Caption         =   "Realtime"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   50
            Top             =   1080
            Value           =   -1  'True
            Width           =   1305
         End
         Begin VB.CheckBox chkOptions 
            Caption         =   "Log Change Events"
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   11
            ToolTipText     =   "Log Status Change Events"
            Top             =   390
            Value           =   1  'Checked
            Width           =   1785
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "Backup Schedule"
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
            Index           =   22
            Left            =   180
            TabIndex        =   53
            Top             =   810
            Width           =   1410
         End
      End
      Begin VB.CommandButton cmdIdxChange 
         Caption         =   "Start the Index Service"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   3720
         Width           =   2145
      End
      Begin VB.CommandButton cmdIdxChange 
         Caption         =   "Stop the Index Service"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2400
         TabIndex        =   8
         Top             =   3720
         Width           =   2145
      End
      Begin MSComctlLib.ListView lstDrives 
         Height          =   3075
         Left            =   2940
         TabIndex        =   12
         Top             =   390
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   5424
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Monitor Status"
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
         Index           =   21
         Left            =   2940
         TabIndex        =   49
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.PictureBox picBg 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4455
      Index           =   2
      Left            =   0
      ScaleHeight     =   4425
      ScaleWidth      =   9135
      TabIndex        =   2
      Top             =   600
      Width           =   9165
      Begin VB.CommandButton cmdAddFile 
         Caption         =   "Add to Watch List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   2160
         TabIndex        =   22
         Top             =   3720
         Width           =   1785
      End
      Begin VB.CommandButton cmdAddFile 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   3570
         TabIndex        =   21
         Top             =   3180
         Width           =   375
      End
      Begin VB.TextBox txtFilePath 
         Height          =   315
         Left            =   270
         TabIndex        =   20
         Top             =   3210
         Width           =   3165
      End
      Begin VB.ListBox lstFiles 
         Height          =   3420
         Left            =   4320
         TabIndex        =   19
         Top             =   450
         Width           =   4305
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Select a File:"
         Height          =   210
         Index           =   23
         Left            =   270
         TabIndex        =   57
         Top             =   3000
         Width           =   915
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "1) Select a File to be monitored"
         Height          =   210
         Index           =   20
         Left            =   240
         TabIndex        =   48
         Top             =   510
         Width           =   2220
      End
      Begin VB.Label lblInfo 
         Caption         =   "2) Choose 'Add to Watch List' to start monitoring that file"
         Height          =   450
         Index           =   19
         Left            =   240
         TabIndex        =   47
         Top             =   780
         Width           =   2790
      End
      Begin VB.Label lblInfo 
         Caption         =   "The Selected File will be marked for Backup each time its contents have Changed"
         Height          =   660
         Index           =   18
         Left            =   240
         TabIndex        =   46
         Top             =   1260
         Width           =   2835
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "File Monitoring Steps:"
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
         Index           =   2
         Left            =   240
         TabIndex        =   45
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Currently Monitored Files"
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
         Index           =   3
         Left            =   4320
         TabIndex        =   23
         Top             =   240
         Width           =   2115
      End
   End
   Begin VB.PictureBox picBg 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4455
      Index           =   1
      Left            =   0
      ScaleHeight     =   4425
      ScaleWidth      =   9165
      TabIndex        =   1
      Top             =   600
      Width           =   9195
      Begin VB.CommandButton cmdAddPath 
         Caption         =   "Add to Watch List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   2100
         TabIndex        =   16
         Top             =   3720
         Width           =   1785
      End
      Begin VB.CommandButton cmdAddPath 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   3510
         TabIndex        =   15
         Top             =   3270
         Width           =   375
      End
      Begin VB.TextBox txtPath 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   3300
         Width           =   3165
      End
      Begin VB.ListBox lstMonitored 
         Height          =   3630
         Left            =   4380
         TabIndex        =   13
         Top             =   360
         Width           =   4605
      End
      Begin MSComctlLib.ImageCombo icbFileType 
         Height          =   345
         Left            =   240
         TabIndex        =   18
         Top             =   2640
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   609
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblInfo 
         Caption         =   "4) All Files of the Selected Type will be Monitored for Changes. When a Change occurs, the File will be Archived Automatically."
         Height          =   660
         Index           =   17
         Left            =   240
         TabIndex        =   43
         Top             =   1590
         Width           =   3405
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Select Path to Monitor:"
         Height          =   210
         Index           =   16
         Left            =   240
         TabIndex        =   42
         Top             =   3060
         Width           =   1605
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Select a File Type:"
         Height          =   210
         Index           =   15
         Left            =   240
         TabIndex        =   41
         Top             =   2430
         Width           =   1320
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Monitor a Path For Files of Type :"
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
         Index           =   14
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   2685
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "3) Add the Selection to the Watch List."
         Height          =   210
         Index           =   13
         Left            =   240
         TabIndex        =   39
         Top             =   1290
         Width           =   2775
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "2) Choose the Path you want to Monitor."
         Height          =   210
         Index           =   12
         Left            =   240
         TabIndex        =   38
         Top             =   990
         Width           =   2910
      End
      Begin VB.Label lblInfo 
         Caption         =   "1) Select the type of file extension to Monitor from the dropdown list. "
         Height          =   450
         Index           =   0
         Left            =   240
         TabIndex        =   37
         Top             =   510
         Width           =   2685
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Currently Monitored Paths"
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
         Index           =   1
         Left            =   4380
         TabIndex        =   17
         Top             =   180
         Width           =   2190
      End
   End
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5055
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic16 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   270
      Picture         =   "frmIndex.frx":57E8
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   60
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":7F8A
            Key             =   "dft"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolbar 
      Left            =   30
      Top             =   1950
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":A73C
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":B590
            Key             =   "Folder Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":C3E4
            Key             =   "Item New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":D238
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":D7D4
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":DD70
            Key             =   "About"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":EBC4
            Key             =   "Book Red"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":F160
            Key             =   "Book Blue"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":F6FC
            Key             =   "Book Cyan"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":FC98
            Key             =   "Book Brown"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":10234
            Key             =   "Book Purple"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":107D0
            Key             =   "Import"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":10D6C
            Key             =   "Permanent"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":11BC0
            Key             =   "Folder New"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":12A14
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":12FB0
            Key             =   "Extract"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":1354C
            Key             =   "Ariel1"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":13E28
            Key             =   "Ariel"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":14C7C
            Key             =   "FileAdd"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":15218
            Key             =   "Folder Add"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":1606C
            Key             =   "Refresh"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolbar 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   1005
      ButtonWidth     =   1905
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open Archive"
            Key             =   "Open Archive"
            Object.ToolTipText     =   "Open the Backup Archive"
            ImageKey        =   "Folder Open"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add Path"
            Key             =   "Add Path"
            Object.ToolTipText     =   "Add a Path to Monitor"
            ImageKey        =   "Folder Add"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add Files"
            Key             =   "Add Files"
            Object.ToolTipText     =   "Add files to the Monitor"
            ImageKey        =   "FileAdd"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Options"
            Key             =   "Options"
            Object.ToolTipText     =   "Monitor Advanced Options"
            ImageKey        =   "Refresh"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Key             =   "About"
            Object.ToolTipText     =   "About RBS Jet"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsICB 
      Left            =   30
      Top             =   2550
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":16608
            Key             =   "dft"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picIcb 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   150
      Picture         =   "frmIndex.frx":1C8A2
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   44
      Top             =   2670
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComDlg.CommonDialog cdFile 
      Left            =   30
      Top             =   4050
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'/~ RBS Jet Test Harness
'/~ By John Underhill (Steppenwolfe)
'/~ June 05, 2006
'/~ Updated Archive and Service Classes June 30, 2006

'/* archive constants
Private Const COMP_NAME                 As String = "rbsarchive.rbc"
Private Const DECOMP_NAME               As String = "rbsarchive.rba"
Private Const SVC_NAME                  As String = "rbsidx.exe"

'/* icon flags
Private Const BASIC_SHGFI_FLAGS         As Double = _
    &H4 Or &H200 Or &H400 Or &H2000 Or &H4000

Private Enum eArchiveState
    NoArchive = 0
    Compressed = 1
    DeCompressed = 2
End Enum

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, _
                                                                                 ByVal dwFileAttributes As Long, _
                                                                                 psfi As SHFILEINFO, _
                                                                                 ByVal cbSizeFileInfo As Long, _
                                                                                 ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal hIml As Long, _
                                                            ByVal i As Long, _
                                                            ByVal hDCDest As Long, _
                                                            ByVal X As Long, _
                                                            ByVal Y As Long, _
                                                            ByVal Flags As Long) As Long

Private m_bStopped                      As Boolean
Private m_lImlHandle                    As Long
Private m_sRegPath                      As String
Private m_sIdxPath                      As String
Private m_cSIcon                        As Collection
Private tShInfo                         As SHFILEINFO
Private WithEvents cArchive             As clsArchive
Attribute cArchive.VB_VarHelpID = -1
Private cLightning                      As clsLightning
Private cService                        As clsService


Private Sub Form_Load()

    Set cArchive = New clsArchive
    Set cLightning = New clsLightning
    Set cService = New clsService
    m_sIdxPath = App.Path + "\Index\"
    m_sRegPath = "Software\" + App.ProductName + "\Index"
    '/* test service status
    ServiceCheck
    '/* check for started
    ServiceAcivate
    '/* archive config
    ArchiveDefaults
    '/* fetch contents of archive
    ArchiveInitList
    '/* get file extensions
    ExtensionsList
    '/* load file list
    InitFileList
    '/* load path list
    InitPathList
    '/* load drive list
    InitDriveList
    '/* list contents of archive
    ArchiveList
    '/* load about box
    LoadAbout
    tbToolbar_ButtonClick tbToolbar.Buttons.Item(1)

End Sub

'> Events
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub cArchive_eECompPMax(lMax As Long)
'/* progress max

    pbInflate.Max = lMax

End Sub

Private Sub cArchive_eECompPTick(lCnt As Long)
'/* progress tick

On Error Resume Next

    With pbInflate
        .Value = lCnt
        If lCnt = .Max Then
            .Value = 0
            .Visible = False
        End If
    End With
            
On Error GoTo 0

End Sub


'> Controls
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub cmdAddFile_Click(Index As Integer)
'/* add a file to the monitor

Dim sPath   As String
Dim cTemp   As Collection
Dim vItem   As Variant
Dim sDrive  As String

On Error GoTo Handler

    Select Case Index
    '/* select a file
    Case 0
        With cdFile
            .CancelError = True
            .InitDir = App.Path
            .ShowOpen
            txtFilePath.Text = .FileName
        End With
    '/* add to monitor
    Case 1
        If Len(txtFilePath.Text) = 0 Then
            MsgBox "Please Select a File to be Monitored before Proceeding.", _
                vbExclamation, "No File Selected!"
            Exit Sub
        End If
        If cArchive.File_Exists(txtFilePath.Text) Then
            Set cTemp = New Collection
            With cLightning
                '/* add file filter
                sPath = LCase$(txtFilePath.Text)
                If Not .Read_MultiCN(HKEY_LOCAL_MACHINE, m_sRegPath, "fltfle") Is Nothing Then
                    Set cTemp = .Read_MultiCN(HKEY_LOCAL_MACHINE, m_sRegPath, "fltfle")
                    '/* test for existing entry
                    For Each vItem In cTemp
                        If CStr(vItem) = sPath Then
                            GoTo Handler
                        End If
                    Next vItem
                    cTemp.Add sPath
                    .Delete_Value HKEY_LOCAL_MACHINE, m_sRegPath, "fltfle"
                    .Write_MultiCN HKEY_LOCAL_MACHINE, m_sRegPath, "fltfle", cTemp
                Else
                    Set cTemp = New Collection
                    cTemp.Add sPath
                    .Write_MultiCN HKEY_LOCAL_MACHINE, m_sRegPath, "fltfle", cTemp
                End If
                
                Set cTemp = New Collection
                '/* add the drive to global monitor
                sDrive = LCase$(Left$(txtFilePath.Text, 3))
                If Not .Read_MultiCN(HKEY_LOCAL_MACHINE, m_sRegPath, "drvpth") Is Nothing Then
                    Set cTemp = .Read_MultiCN(HKEY_LOCAL_MACHINE, m_sRegPath, "drvpth")
                    '/* test for existing entry
                    For Each vItem In cTemp
                        If CStr(vItem) = sDrive Then
                            GoTo skipdrv
                        End If
                    Next vItem
                    cTemp.Add sDrive
                    .Delete_Value HKEY_LOCAL_MACHINE, m_sRegPath, "drvpth"
                    .Write_MultiCN HKEY_LOCAL_MACHINE, m_sRegPath, "drvpth", cTemp
                Else
                    cTemp.Add sDrive
                    .Write_MultiCN HKEY_LOCAL_MACHINE, m_sRegPath, "drvpth", cTemp
                End If
skipdrv:
                '/* enable file monitor
                .Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "flemon", 1
                '/* set the change master flag
                .Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "mstflg", 1
            End With
            '/* start service
            ServiceAcivate
            InitFileList
        End If
    End Select
    
Handler:
    On Error GoTo 0

End Sub

Private Sub cmdAddPath_Click(Index As Integer)
'/* add a path a filter to the monitor

Dim sPath       As String
Dim sFile       As String
Dim sTitle      As String
Dim cTemp       As Collection
Dim sDrive      As String
Dim vItem       As Variant

On Error GoTo Handler

    Select Case Index
    '/* select a path
    Case 0
        sTitle = "Select a Path to Monitor"
        sPath = FolderBrowse(sTitle, Me.hWnd)
        If Len(sPath) = 0 Then Exit Sub
        txtPath.Text = sPath
    '/* add to list
    Case 1
        '/* validity checks
        sFile = Mid$(icbFileType.Text, InStrRev(icbFileType.Text, Chr$(46)))
        If Len(txtPath.Text) = 0 Then
            MsgBox "Please specify a path to be monitored before proceeding.", _
                vbExclamation, "No Path Specified!"
            Exit Sub
        ElseIf Len(sFile) = 0 Then
            MsgBox "Please specify a path to be monitored before proceeding.", _
                vbExclamation, "No Path Specified!"
            Exit Sub
        End If
        Set cTemp = New Collection
        sDrive = LCase$(Left$(txtPath.Text, 3))
        If Not MediaCheck(sDrive) = HardDrive Then
            MsgBox "File can not be archived to Portable media. Please select a Hard Drive to continue.", _
                vbExclamation, "Drive Type Not Supported!"
            Exit Sub
        End If
        '/* write monitoring params
        '/ and set master flag
        With cLightning
            '/* add qualified filter
            sPath = txtPath.Text + Chr$(30) + sFile
            Set cTemp = New Collection
            If Not .Read_MultiCN(HKEY_LOCAL_MACHINE, m_sRegPath, "fltext") Is Nothing Then
                Set cTemp = .Read_MultiCN(HKEY_LOCAL_MACHINE, m_sRegPath, "fltext")
                '/* test for existing entry
                For Each vItem In cTemp
                    If CStr(vItem) = sPath Then
                        GoTo Handler
                    End If
                Next vItem
                cTemp.Add sPath
                .Delete_Value HKEY_LOCAL_MACHINE, m_sRegPath, "fltext"
                .Write_MultiCN HKEY_LOCAL_MACHINE, m_sRegPath, "fltext", cTemp
            Else
                cTemp.Add sPath
                .Write_MultiCN HKEY_LOCAL_MACHINE, m_sRegPath, "fltext", cTemp
            End If
            
            Set cTemp = New Collection
            '/* add the drive to global monitor
            If Not .Read_MultiCN(HKEY_LOCAL_MACHINE, m_sRegPath, "drvpth") Is Nothing Then
                Set cTemp = .Read_MultiCN(HKEY_LOCAL_MACHINE, m_sRegPath, "drvpth")
                '/* test for existing entry
                For Each vItem In cTemp
                    If CStr(vItem) = sDrive Then
                        GoTo skipdrv
                    End If
                Next vItem
                cTemp.Add sDrive
                .Delete_Value HKEY_LOCAL_MACHINE, m_sRegPath, "drvpth"
                .Write_MultiCN HKEY_LOCAL_MACHINE, m_sRegPath, "drvpth", cTemp
            Else
                cTemp.Add sDrive
                .Write_MultiCN HKEY_LOCAL_MACHINE, m_sRegPath, "drvpth", cTemp
            End If
skipdrv:
            '/* enable path monitor
            .Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "pthmon", 1
            '/* set the change master flag
            .Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "mstflg", 1
            InitPathList
        End With
        '/* start service
        ServiceAcivate
    End Select

Handler:
    On Error GoTo 0

End Sub

Private Sub cmdFileExtract_Click(Index As Integer)
'/* archive manipulation

Dim sPath       As String
Dim sFile       As String
Dim sTitle      As String
Dim sName       As String

On Error GoTo Handler

    sFile = lstArchiveCts.SelectedItem.Text
    '/* input tests
    If Len(sFile) = 0 Then
        MsgBox "Please select an item from the list.", vbExclamation, "No File Selected!"
        Exit Sub
    End If
    
    '/* file paths
    txtExtractPath.Text = sPath
    
    With cArchive
        InExtraction True
        Select Case Index
        '/* extract
        Case 0
            sTitle = "Select the Extraction Location"
            sPath = FolderBrowse(sTitle, Me.hWnd)
            If Len(sPath) = 0 Then Exit Sub
            '/* delete from archive
            .p_Rebuild = True
            sName = Mid$(sFile, InStrRev(sFile, Chr$(92)) + 1)
            .Archive_Extract sFile, sPath + sName
            lstArchiveCts.ListItems.Remove (lstArchiveCts.SelectedItem.Index)
            stBar.SimpleText = "The Selected File has been Extracted.."
        '/* delete
        Case 1
            .Archive_Remove sFile
            lstArchiveCts.ListItems.Remove (lstArchiveCts.SelectedItem.Index)
            stBar.SimpleText = "The Selected File has been Deleted.."
        End Select
    End With

On Error GoTo 0
Handler:
    InExtraction False

End Sub

Private Sub chkOptions_Click(Index As Integer)
'/* option settings

    With cLightning
        Select Case Index
        Case 0
            If chkOptions(Index).Value = 1 Then
                stBar.SimpleText = "State Changes will be Logged.."
                '/* toggle logged property
                .Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "mdlogg", 1
            Else
                stBar.SimpleText = "State Changes will not be Logged.."
                .Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "mdlogg", 0
            End If
        End Select
        .Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "mstflg", 0
    End With

End Sub

Private Sub cmdIdxChange_Click(Index As Integer)
'/* index service functions

    With cService
        Select Case Index
        '/* start
        Case 0
            .Service_Start
        '/* stop
        Case 1
            .Service_Stop
        '/* uninstall
        Case 2
            .Service_Uninstall
            cLightning.Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "svcins", 0
        '/* reinstall
        Case 3
            .Service_Install
            cLightning.Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "svcins", 1
        End Select
    End With

End Sub

Private Sub optSchedule_Click(Index As Integer)
'/* backup interval

    With cLightning
        Select Case Index
        Case 0
            .Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "bkptmr", 1
        Case 1
            .Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "bkptmr", 2
            .Write_String HKEY_LOCAL_MACHINE, m_sRegPath, "bkpscd", CStr(Now)
        Case 2
            .Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "bkptmr", 3
            .Write_String HKEY_LOCAL_MACHINE, m_sRegPath, "bkpscd", CStr(Now)
        End Select
        .Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "mstflg", 1
    End With
    
End Sub

Private Sub tbToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
'/* toolbar control

Dim oPic    As PictureBox

    For Each oPic In picBg
        With oPic
            .Visible = False
            .Left = 0
            .Top = tbToolbar.Height
            .BorderStyle = 0
        End With
    Next oPic
    
    Select Case Button
    Case "Open Archive"
        picBg(0).Visible = True
        ArchiveList
        stBar.SimpleText = "Open an existing Archive and Extract Files.."
    Case "Add Path"
        picBg(1).Visible = True
        stBar.SimpleText = "Add a Directory to be Monitored .."
    Case "Add Files"
        picBg(2).Visible = True
        stBar.SimpleText = "Add a File to the Watch List.."
    Case "Options"
        picBg(3).Visible = True
        stBar.SimpleText = "Indexer Service Options.."
    Case "About"
        picBg(4).Visible = True
        stBar.SimpleText = "About RBS Jet.."
    End Select
    
End Sub


'> User Queues
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub ArchiveList()
'/* get list of archived files

Dim cTemp       As Collection
Dim vItem       As Variant
Dim aRes()      As String
Dim cItem       As ListItem

On Error GoTo Handler

    '/* service busy flag
    InExtraction True
    
    lstArchiveCts.ListItems.Clear
    With cArchive
        .Archive_List
        If Not .p_CReturn Is Nothing Then
            Set cTemp = .p_CReturn
            For Each vItem In cTemp
                aRes = Split(vItem, Chr$(30))
                With lstArchiveCts
                    Set cItem = .ListItems.Add(Text:=aRes(0))
                    cItem.SubItems(1) = aRes(1)
                    cItem.SubItems(2) = aRes(2) + " bytes"
                End With
            Next vItem
        End If
        .p_Remove = True
        .Archive_Compress
        Set .p_CReturn = New Collection
    End With

On Error GoTo 0
Handler:
    InExtraction False

End Sub

Private Sub InitFileList()
'/* list monitored files

Dim vItem   As Variant
Dim cTemp   As Collection

    lstFiles.Clear
    With cLightning
        If Not .Read_MultiCN(HKEY_LOCAL_MACHINE, m_sRegPath, "fltfle") Is Nothing Then
            Set cTemp = .Read_MultiCN(HKEY_LOCAL_MACHINE, m_sRegPath, "fltfle")
            '/* test for existing entry
            For Each vItem In cTemp
                lstFiles.AddItem vItem
            Next vItem
        End If
    End With
    
Handler:
    On Error GoTo 0

End Sub

Private Sub InitPathList()
'/* list monitored files

Dim vItem   As Variant
Dim cTemp   As Collection
Dim aRes()  As String

    lstMonitored.Clear
    With cLightning
        If Not .Read_MultiCN(HKEY_LOCAL_MACHINE, m_sRegPath, "fltext") Is Nothing Then
            Set cTemp = .Read_MultiCN(HKEY_LOCAL_MACHINE, m_sRegPath, "fltext")
            '/* test for existing entry
            For Each vItem In cTemp
                aRes = Split(vItem, Chr$(30))
                lstMonitored.AddItem "File: " + aRes(0) + "  Extension: " + aRes(1)
            Next vItem
        End If
    End With
    
Handler:
    On Error GoTo 0

End Sub

Private Sub ExtensionsList()
'/* populate icb list
'/* you could enumerate registered classes and get
'/* all file types here, seems like too many though

    With icbFileType
        Set .ImageList = ilsICB
        .ComboItems.Add Text:="Text Files .txt", Image:=IcbIcons("dummy.txt")
        .ComboItems.Add Text:="Documents .doc", Image:=IcbIcons("dummy.doc")
        .ComboItems.Add Text:="Excel Files .xls", Image:=IcbIcons("dummy.xls")
        .ComboItems.Add Text:="Access Files .mdb", Image:=IcbIcons("dummy.mdb")
        .ComboItems.Add Text:="VB Projects .vbp", Image:=IcbIcons("dummy.vbp")
        .ComboItems.Add Text:="VB Modules .bas", Image:=IcbIcons("dummy.bas")
        .ComboItems.Add Text:="VB Classes .cls", Image:=IcbIcons("dummy.cls")
        .ComboItems.Add Text:="VB Controls .ctl", Image:=IcbIcons("dummy.ctl")
        .ComboItems.Add Text:="VB Forms .frm", Image:=IcbIcons("dummy.frm")
    End With

End Sub

Private Function IcbIcons(ByVal sFile As String) As String
'/* use system image list to extract
'/* app icons for icb image list

Dim lIcon           As Long
Dim imgObj          As ListImage
Dim sKey            As String
Dim tFileInfo       As SHFILEINFO

On Error Resume Next

    If m_lImlHandle = 0 Then
        m_lImlHandle = mSundry.p_ListHandle
    End If
    lIcon = mSundry.IconIndex(sFile, tFileInfo)

    '/* load icon to picturebox
    If Not lIcon = 0 Then
        With picIcb
            Set .Picture = LoadPicture("")
            .AutoRedraw = True
            ImageList_Draw lIcon, tFileInfo.iIcon, .hDC, 0&, 0&, &H1
            .Refresh
        End With
        '/* use file extension as image key
        Set imgObj = ilsICB.ListImages.Add(Key:=sFile, Picture:=picIcb.Image)
        IcbIcons = sFile
    Else
        '/* no icon, use default image
        IcbIcons = "dft"
    End If
    
On Error GoTo 0

End Function

Private Sub InitDriveList()
'/* init the drive list

On Error Resume Next

    With lstDrives
        .SmallIcons = iml16
        .View = lvwReport
        .LabelEdit = lvwManual
        .ListItems.Clear
        .ColumnHeaders.Clear
        .FullRowSelect = True
        .GridLines = True
        .AllowColumnReorder = True
        .ColumnHeaders.Add 1, , "Volume", (.Width / 5) - 20
        .ColumnHeaders.Add 2, , "Status", .Width / 5
        .ColumnHeaders.Add 3, , "File System", .Width / 5
        .ColumnHeaders.Add 4, , "Capacity GB", .Width / 5
        .ColumnHeaders.Add 5, , "Free Space", .Width / 5
    End With
    
    '/* get drive data
    DriveListPopulate

On Error GoTo 0

End Sub

Private Sub DriveListPopulate()
'/* populate drive list

Dim vItem           As Variant
Dim cTemp           As Collection
Dim cItem           As ListItem
Dim sIcon           As String
Dim lCount          As Long

On Error GoTo Handler

    '/* check for mods first
    Set cTemp = DriveList
    If cTemp.Count = 0 Then Exit Sub

    '/* clear list
    lstDrives.ListItems.Clear
    
    '/* put to list
    For Each vItem In cTemp
        sIcon = DriveListIcon(CStr(vItem))
        Set cItem = lstDrives.ListItems.Add(Text:=vItem, SmallIcon:=sIcon)
        cItem.SubItems(2) = DriveType(CStr(vItem))
        cItem.SubItems(3) = DriveSize(CStr(vItem)).Item(1)
        cItem.SubItems(4) = DriveSize(CStr(vItem)).Item(2)
    Next vItem
    
    '/* active status
    If MonitorList Is Nothing Then
        For lCount = 1 To lstDrives.ListItems.Count
            lstDrives.ListItems.Item(lCount).SubItems(1) = "InActive"
        Next lCount
    Else
        Set cTemp = MonitorList
        For lCount = 1 To lstDrives.ListItems.Count
            For Each vItem In cTemp
                If LCase$(lstDrives.ListItems.Item(lCount)) = LCase$(Left$(CStr(vItem), 3)) Then
                    lstDrives.ListItems.Item(lCount).SubItems(1) = "Active"
                    Exit For
                Else
                    lstDrives.ListItems.Item(lCount).SubItems(1) = "InActive"
                End If
            Next vItem
        Next lCount
    End If
    
On Error GoTo 0

Handler:

End Sub

Private Function MonitorList() As Collection
'/* get active drives

On Error Resume Next

    With New clsLightning
        If Not .Read_MultiCN(HKEY_LOCAL_MACHINE, m_sRegPath, "drvpth").Count = 0 Then
            Set MonitorList = .Read_MultiCN(HKEY_LOCAL_MACHINE, m_sRegPath, "drvpth")
        End If
    End With
    
On Error GoTo 0

End Function

Private Function DriveListIcon(ByVal sPath As String) As String
'/* get listview icons

Dim lIcon           As Long
Dim imgObj          As ListImage
Dim sKey            As String

On Error Resume Next

    '/* get handle to icon
    lIcon = SHGetFileInfo(sPath, 0&, tShInfo, Len(tShInfo), BASIC_SHGFI_FLAGS Or &H1)
    '/* load icon to picturebox
    If Not lIcon = 0 Then
        With pic16
            Set .Picture = LoadPicture("")
            .AutoRedraw = True
            ImageList_Draw lIcon, tShInfo.iIcon, .hDC, 0&, 0&, &H1
            .Refresh
        End With
        
        '/* test for icon presence in collection
        sKey = m_cSIcon.Item(sPath)
        '/* if not present, add to collection
        If LenB(sKey) = 0 Then
            m_cSIcon.Add 1, sPath
            '/* add icon to image list
            '/* use file extension as image key
            Set imgObj = iml16.ListImages.Add(Key:=sPath, Picture:=pic16.Image)
        End If
        DriveListIcon = sPath
    Else
        '/* no icon, use default image
        DriveListIcon = "dft"
    End If
    
On Error GoTo 0

End Function

Private Sub LoadAbout()

Dim sAbout  As String

    sAbout = "RBS Jet v3.0 - by Steppenwolfe" & vbNewLine & _
    "Instructions: Open the Service project first, add the reference to ntsvchp.tlb type lib " & vbNewLine & _
    "found in the TLB folder. Compile the service component. Return to this project and run, " & vbNewLine & _
    "the service should have started automatically, if not, go to Options, and click, 'Install Service', " & vbNewLine & _
    "then 'Start Service' Check the service state by clicking on 'Service Status' to launch " & vbNewLine & _
    "the MMC Service console." & vbNewLine & vbNewLine & _
    "Operation: Add a file to be monitored to the watch list by clicking 'Add Files', and , " & vbNewLine & _
    "browsing to the file, then click 'Add to Watch List'. This file will then be monitored for " & vbNewLine & _
    "changes, and archived each time the file contents change. To add a path and all files " & vbNewLine & _
    "of a desired type, select 'Add Path', choose the file type to be monitored, select a directory, and click " & vbNewLine & _
    "'Add to Watch List', all files of that type in the path, will be monitored for changes, and archived " & vbNewLine & _
    "automatically. To extract a file, choose 'Open Archive', select the file, and choose Extract File."
    txtAbout.Text = sAbout

End Sub


'> Archive Engine
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub ArchiveDefaults()
'/* set archiver default paths

    With cArchive
        .p_CompName = m_sIdxPath + COMP_NAME
        .p_DecompName = m_sIdxPath + DECOMP_NAME
        .p_CompRatio = cLow
        .Startup_Check m_sIdxPath
    End With
    
End Sub

Private Sub ArchiveInitList()
'/* list setup

    With lstArchiveCts
        .View = lvwReport
        .AllowColumnReorder = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, Text:="File Name", Width:=(.Width / 4) * 2
        .ColumnHeaders.Add 2, Text:="File Added", Width:=.Width / 4
        .ColumnHeaders.Add 3, Text:="File Size", Width:=(.Width / 4) - 30
    End With

End Sub

Private Function ArchiveCompressed() As eArchiveState
'/* return current state of archive

    With cArchive
        If .File_Exists(m_sIdxPath + COMP_NAME) Then
            ArchiveCompressed = Compressed
        ElseIf .File_Exists(m_sIdxPath + DECOMP_NAME) Then
            ArchiveCompressed = DeCompressed
        Else
            ArchiveCompressed = NoArchive
        End If
    End With

End Function

'> Service Flags
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub InExtraction(ByVal bState As Boolean)
'/* signal a busy state to the service

    With cLightning
        If bState Then
            .Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "mdabsy", 1
        Else
            .Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "mdabsy", 0
        End If
    End With

End Sub

Private Sub ServiceCheck()
'/* get service state and install/start

Dim lState  As Long

    With cLightning
        '/* installed flag
        If .Read_DWord(HKEY_LOCAL_MACHINE, m_sRegPath, "svcins") = 0 Then
            If Not cService.Service_Install = True Then
                MsgBox "The Idexing Service is Not Installed, or Not Compiled!" + vbNewLine + _
                "Please refer to the Instructions Readme for more information.", vbExclamation, "No Service!"
                Exit Sub
            Else
                .Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "svcins", 1
            End If
        End If
    End With

End Sub

Private Sub ServiceStart()
    
    With cService
        '/* not started
        If Not .Service_State = 4 Then
            If Not .Service_Start = True Then
                MsgBox "The Idexing Service is Not Installed, or Not Compiled!" + vbNewLine + _
                "Please refer to the Instructions Readme for more information.", vbExclamation, "No Service!"
                Exit Sub
            Else
                '/* add description and set to autostart
                .Service_StartUp START_AUTO
                Dim sDesc As String
                sDesc = "RBS Jet Real Time Backup Service"
                .Service_Desc sDesc
            End If
        End If
    End With

End Sub

Private Function ServiceAcivate() As Boolean
'/* active drive list test

    If Not cLightning.Read_MultiCN(HKEY_LOCAL_MACHINE, m_sRegPath, "drvpth") Is Nothing Then
        ServiceAcivate = True
        cLightning.Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "svcins", 1
        If Not cService.Service_State = 4 Then
            ServiceStart
        End If
    End If

End Function

Private Function ServiceState() As Boolean
'/* service running status

    If cService.Service_State = 4 Then
        ServiceState = True
    End If

End Function

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next

    Set cArchive = Nothing
    Set cArchive = Nothing
    Set cLightning = Nothing
    Set cService = Nothing

End Sub
