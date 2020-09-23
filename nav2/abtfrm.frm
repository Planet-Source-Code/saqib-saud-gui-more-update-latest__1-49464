VERSION 5.00
Begin VB.Form abtfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   4695
   ClientLeft      =   3555
   ClientTop       =   1500
   ClientWidth     =   4695
   LinkTopic       =   "Form2"
   ScaleHeight     =   4695
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   3480
      Top             =   4080
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   600
      Picture         =   "abtfrm.frx":0000
      ScaleHeight     =   2655
      ScaleWidth      =   3495
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BorderColor     =   &H00808080&
      Height          =   4695
      Left            =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail: Wisesabre@hotmail.com"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Developer: Saqib Saud"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"abtfrm.frx":2258A
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   3240
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Caution:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.2"
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
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   2880
      Width           =   1455
   End
End
Attribute VB_Name = "abtfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
If form1.Visible = False Then
form1.Visible = True
End If
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Label3_Click()
Unload Me
End Sub

Private Sub Label4_Click()
Unload Me
End Sub

Private Sub Label5_Click()
Unload Me
End Sub

Private Sub Picture1_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()

Unload Me
End Sub
