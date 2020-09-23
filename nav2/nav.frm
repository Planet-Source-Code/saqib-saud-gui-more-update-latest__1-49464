VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form form1 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GUI"
   ClientHeight    =   7215
   ClientLeft      =   2415
   ClientTop       =   495
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   7560
   Begin Project1.chameleonButton chameleonButton4 
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   6720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "&Back"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "nav.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton chameleonButton3 
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Top             =   6720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "Gene&rate"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "nav.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton chameleonButton2 
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   6720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "Ne&xt"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "nav.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton chameleonButton1 
      Height          =   375
      Left            =   6240
      TabIndex        =   18
      Top             =   6720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "Fini&sh"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "nav.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   7575
      Begin VB.PictureBox Picture4 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   6555
         Left            =   0
         Picture         =   "nav.frx":0070
         ScaleHeight     =   6555
         ScaleWidth      =   7515
         TabIndex        =   24
         Top             =   0
         Width           =   7515
         Begin VB.TextBox cclr 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   5640
            TabIndex        =   5
            Text            =   "#FFFFFF"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox mclr 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   5640
            TabIndex        =   6
            Text            =   "#9999FF"
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox bgclr 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   5640
            TabIndex        =   9
            Text            =   "#000000"
            Top             =   2400
            Width           =   1695
         End
         Begin VB.ComboBox cpd 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   315
            ItemData        =   "nav.frx":A0812
            Left            =   5640
            List            =   "nav.frx":A0825
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1560
            Width           =   615
         End
         Begin VB.ComboBox tbw 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   315
            ItemData        =   "nav.frx":A0838
            Left            =   5640
            List            =   "nav.frx":A084B
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1920
            Width           =   615
         End
         Begin VB.ComboBox Combo3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   315
            ItemData        =   "nav.frx":A085E
            Left            =   5640
            List            =   "nav.frx":A0880
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   3000
            Width           =   855
         End
         Begin VB.TextBox tclr 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   5640
            TabIndex        =   11
            Text            =   "#808080"
            Top             =   3360
            Width           =   1695
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   5640
            TabIndex        =   12
            Text            =   "#FFFFFF"
            Top             =   3840
            Width           =   1695
         End
         Begin VB.ComboBox Combo4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   315
            ItemData        =   "nav.frx":A08A8
            Left            =   5640
            List            =   "nav.frx":A08C7
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   4320
            Width           =   1455
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   810
            ItemData        =   "nav.frx":A091F
            Left            =   5640
            List            =   "nav.frx":A092F
            TabIndex        =   14
            Top             =   4800
            Width           =   1455
         End
         Begin VB.TextBox tblw 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   5640
            TabIndex        =   15
            Text            =   "120"
            Top             =   6120
            Width           =   1695
         End
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   7575
      Begin VB.OptionButton Option2 
         Caption         =   "Preview"
         Height          =   255
         Left            =   1200
         TabIndex        =   33
         Top             =   6120
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Code"
         Height          =   255
         Left            =   2280
         TabIndex        =   32
         Top             =   6120
         Width           =   975
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   7575
         TabIndex        =   21
         Top             =   0
         Width           =   7575
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Code Generation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   240
            TabIndex        =   22
            Top             =   120
            Width           =   2175
         End
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   5415
         Left            =   240
         TabIndex        =   31
         Top             =   600
         Width           =   7215
         ExtentX         =   12726
         ExtentY         =   9551
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin VB.TextBox gentxt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   5415
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   25
         Top             =   600
         Width           =   7215
      End
      Begin VB.TextBox urlgen 
         Height          =   285
         Left            =   1200
         TabIndex        =   27
         Top             =   5520
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   6720
         TabIndex        =   34
         Top             =   6120
         Width           =   615
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   7560
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   7560
         X2              =   0
         Y1              =   6480
         Y2              =   6480
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   7575
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4665
         Left            =   0
         Picture         =   "nav.frx":A0956
         ScaleHeight     =   4665
         ScaleWidth      =   7545
         TabIndex        =   19
         Top             =   0
         Width           =   7545
         Begin Project1.chameleonButton delink 
            Height          =   375
            Left            =   4920
            TabIndex        =   30
            Top             =   2640
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            BTYPE           =   5
            TX              =   "De&lete Link"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   12632256
            BCOLO           =   12632256
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "nav.frx":113670
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Project1.chameleonButton Command3 
            Height          =   375
            Left            =   4920
            TabIndex        =   29
            Top             =   1800
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            BTYPE           =   5
            TX              =   "Move Link &Down"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   12632256
            BCOLO           =   12632256
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "nav.frx":11368C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Project1.chameleonButton Command2 
            Height          =   375
            Left            =   4920
            TabIndex        =   28
            Top             =   1320
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            BTYPE           =   5
            TX              =   "Move Link &Up"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   12632256
            BCOLO           =   12632256
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "nav.frx":1136A8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Project1.chameleonButton adlnk 
            Height          =   375
            Left            =   5280
            TabIndex        =   3
            Top             =   4290
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            BTYPE           =   5
            TX              =   "Add Link"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   12632256
            BCOLO           =   12632256
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "nav.frx":1136C4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComctlLib.ListView lnklst 
            Height          =   2535
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   4471
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   8421504
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Name"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "URL"
               Object.Width           =   3528
            EndProperty
         End
         Begin VB.TextBox lnkurl 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   3840
            TabIndex        =   2
            Text            =   "about:blank"
            Top             =   3840
            Width           =   3255
         End
         Begin VB.TextBox lnknm 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   3840
            TabIndex        =   1
            Text            =   "WiseSabre"
            Top             =   3360
            Width           =   3255
         End
      End
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub adlnk_Click()
Dim sFileName As String:    sFileName = lnknm.Text ' To get the filename
Dim sDirectory As String:   sDirectory = lnkurl.Text ' To get the path
Dim objLvi As MSComctlLib.ListItem: Set objLvi = form1.lnklst.ListItems.Add()
objLvi.Text = sFileName
objLvi.SubItems(1) = sDirectory ' Inserts the directory into SubItem(1)
End Sub

Private Sub chameleonButton1_Click()
Unload Me
End Sub

Private Sub chameleonButton2_Click()

If Frame3.Visible = True Then
visibler
chameleonButton4.Enabled = True
Frame1.Visible = True

Else
If Frame1.Visible = True Then
visibler
Frame2.Visible = True
chameleonButton3.Enabled = True
chameleonButton4.Enabled = True
chameleonButton2.Enabled = False

End If
End If
End Sub

Private Sub chameleonButton3_Click()
generate

WebBrowser1.Navigate2 (App.Path & "\wise.html")
End Sub


Private Sub chameleonButton4_Click()

If Frame2.Visible = True Then
visibler
chameleonButton2.Enabled = True
chameleonButton3.Enabled = False
Frame1.Visible = True

Else
If Frame1.Visible = True Then
visibler
Frame3.Visible = True
chameleonButton3.Enabled = True
chameleonButton2.Enabled = True
chameleonButton4.Enabled = False
chameleonButton3.Enabled = False
End If
End If
End Sub

Private Sub Command1_Click()

End Sub
Private Sub lnktxt()
lnknm.Text = ""
lnkurl.Text = ""
End Sub

Private Sub Command2_Click()
ItemUp
End Sub
Private Sub ItemUp()

If lnklst.SelectedItem.Index = 1 Then
Set lnklst.DropHighlight = lnklst.SelectedItem


Else
If lnklst.SelectedItem.Index = lnklst.ListItems.Count Then
    Set itmx = lnklst.ListItems.Add(lnklst.SelectedItem.Index - 1, , lnklst.SelectedItem.Text)
        itmx.SubItems(1) = lnklst.SelectedItem.SubItems(1)
   
        
 
        lnklst.ListItems.Remove (lnklst.SelectedItem.Index)
    Set lnklst.SelectedItem = lnklst.ListItems(lnklst.SelectedItem.Index - 1)
    Set lnklst.DropHighlight = lnklst.SelectedItem


Else
   Set itmx = lnklst.ListItems.Add(lnklst.SelectedItem.Index - 1, , lnklst.SelectedItem.Text)
        itmx.SubItems(1) = lnklst.SelectedItem.SubItems(1)
      
 
        lnklst.ListItems.Remove (lnklst.SelectedItem.Index)
    Set lnklst.SelectedItem = lnklst.ListItems(lnklst.SelectedItem.Index - 2)
    Set lnklst.DropHighlight = lnklst.SelectedItem

End If
End If
End Sub

Private Sub Command3_Click()
ItemDown

End Sub
Private Sub ItemDown()
If lnklst.SelectedItem.Index = lnklst.ListItems.Count Then
    Set lnklst.SelectedItem = lnklst.ListItems(lnklst.ListItems.Count)
    Set lnklst.DropHighlight = lnklst.SelectedItem


Else
    Set itmx = lnklst.ListItems.Add(lnklst.SelectedItem.Index + 2, , lnklst.SelectedItem.Text)
        itmx.SubItems(1) = lnklst.SelectedItem.SubItems(1)
        
 
        lnklst.ListItems.Remove (lnklst.SelectedItem.Index)
    Set lnklst.SelectedItem = lnklst.ListItems(lnklst.SelectedItem.Index + 1)
    Set lnklst.DropHighlight = lnklst.SelectedItem

End If



End Sub

Private Sub Command4_Click()
 linklst.Selected = True
 linklst.RemoveItem
End Sub

Private Sub Command5_Click()
 

End Sub

Private Sub Command6_Click()

End Sub
Private Sub generate()
css = "<style type=""text/css"">" & vbCrLf & "<!--" & vbCrLf & vbCrLf & ".Navlink {COLOR: " & tclr.Text & "; TEXT-DECORATION: none; font-family: " & Combo4.Text & "; font-size: " & Combo3.Text & "; font-weight:" & List1.Text & ";}" & vbCrLf & "a:link.Navlink  {color : " & tclr.Text & ";}" & vbCrLf & "a:visited.Navlink  {color : " & tclr.Text & ";}" & vbCrLf & "a:active.Navlink  {text-decoration: none;}" & vbCrLf & "a:hover.Navlink  {text-decoration: none;}" & vbCrLf & vbCrLf & "-->" & vbCrLf & "</style>"
script = "<script language = ""javascript"">" & vbCrLf & "<!--" & vbCrLf & vbCrLf & "function LmOver(elem, clr)" & vbCrLf & "{elem.style.backgroundColor = clr;" & vbCrLf & "elem.children.tags('A')[0].style.color = """ & Text5.Text & """;" & vbCrLf & "elem.style.cursor = 'hand'}" & vbCrLf & vbCrLf & "function LmOut(elem, clr)" & vbCrLf & "{elem.style.backgroundColor = clr;" & vbCrLf & "elem.children.tags('A')[0].style.color = """ & tclr.Text & """;}" & vbCrLf & vbCrLf & "function LmDown(elem, clr)" & vbCrLf & "{elem.style.backgroundColor = clr;" & vbCrLf & "elem.children.tags('A')[0].style.color = """ & cclr.Text & """;}" & vbCrLf & vbCrLf & "function LmUp(path)" & vbCrLf & "{location.href = path;}" & vbCrLf & vbCrLf & "//-->" & vbCrLf & "</script>"
myemailaddress = "<a href =" & " ""mailto:WiseSabre@hotmail.com"" Class=""navlink"" >" & "wisesabre@Hotmail.com" & "</a>"
htmlbody = "<font face = VERDANA size = 2 Class=""navlink"" ><center><br>" & vbCrLf & "THIS HTML IS GENERATED BY GUI v 2.0 <br>  WISESABRE <br>" & vbCrLf & vbCrLf & myhomepageurl & vbCrLf & vbCrLf & "<br>" & myemailaddress & vbCrLf & "<br></center>"
tabel = vbCrLf & vbCrLf & "<table border=""0"" width=""" & tblw.Text & """ bgcolor=""" & bgclr.Text & """ cellspacing=""0"" cellpadding=""0"">" & vbCrLf & "<tr><td width=""100%"">" & vbCrLf & vbCrLf & "<table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""" & cpd.Text & """>" & vbCrLf

etab = vbCrLf & "</table>" & vbCrLf & vbCrLf & "</td></tr>" & vbCrLf & "</table>"
Dim i
i = 1

 n = lnklst.ListItems.Count
Do Until i = n + 1
urlgen.SelText = "<tr><td width=""100%"" onMouseover=""LmOver(this, '" & mclr.Text & "')"" onMouseout=""LmOut(this, '" & cclr.Text & "')"" onMouseDown=""LmDown(this, '" & cclr.Text & "')"" " & vbCrLf & "onMouseUp=""LmUp('" & lnklst.ListItems.Item(i).ListSubItems(1).Text & "')"" bgcolor=""" & cclr.Text & """><A HREF=""" & lnklst.ListItems.Item(i).ListSubItems(1).Text & """ Class=""navlink"">&nbsp; " & lnklst.ListItems.Item(i).Text & "</a></td></tr>" & vbCrLf


i = i + 1

Loop
gentxt.Text = css & vbCrLf & script & vbclrf & tabel & urlgen.Text & etab & vbclrf & htmlbody


Open (App.Path & "\wise.html") For Output As #1

       Print #1, gentxt.Text
       Close #1

End Sub

Private Sub delink_Click()
delitem
End Sub
Private Sub delitem()
If lnklst.ListItems.Count > 0 Then
lnklst.ListItems.Remove (lnklst.SelectedItem.Index)
delink.Enabled = False

End If
End Sub
Private Sub Form_Load()
Frame1.Visible = False
delink.Enabled = False
List1.ListIndex = 1
cpd.ListIndex = 0
tbw.ListIndex = 0
Combo3.ListIndex = 4
Combo4.ListIndex = 0
chameleonButton2.Enabled = True
chameleonButton4.Enabled = False
chameleonButton3.Enabled = False
WebBrowser1.Navigate2 "about:blank"

End Sub

Private Sub Label2_Click()
visibler
Frame1.Visible = True
End Sub

Private Sub Label3_Click()
visibler
Frame3.Visible = True
End Sub

Private Sub linklst_Click()
Dim i As Integer
Dim strTitle As String

strStoreURL = ""
strStoreTarget = ""

i = linklst.ListIndex
If i = -1 Then
    delink.Enabled = False
Else
    delink.Enabled = True
End If
End Sub

Private Sub Picture1_Click()

End Sub


Private Sub visibler()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
End Sub

Private Sub Label1_Click()
abtfrm.Visible = True
End Sub

Private Sub lnklst_Click()
If lnklst.ListItems.Count > 0 Then
delink.Enabled = True
Else
delink.Enabled = False
End If
End Sub

Private Sub Option1_Click()
WebBrowser1.Visible = False
End Sub

Private Sub Option2_Click()
WebBrowser1.Visible = True
End Sub
