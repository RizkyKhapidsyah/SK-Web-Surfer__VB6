VERSION 5.00
Begin VB.Form FormAbout 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   Picture         =   "FormAbout.frx":0000
   ScaleHeight     =   6015
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Help 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version 3.1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   3735
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Web Surfer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
End
Attribute VB_Name = "FormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
FormAbout.Hide

End Sub

Private Sub Form_Load()
FormAbout.Left = 1000
FormAbout.Top = 1000
End Sub

Private Sub Help_Click()
FormAbout.Hide
Form1!data.Text = "help"

End Sub
