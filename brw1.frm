VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Web Surfer"
   ClientHeight    =   10380
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "brw1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10380
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame AskBox 
      BackColor       =   &H000000FF&
      Caption         =   "Advertisement Alert !"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   4695
      Left            =   3120
      TabIndex        =   75
      Top             =   1320
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CommandButton YesNo 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   4200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton YesNo 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   3840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton YesNo 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   3480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton YesNo 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   3120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton YesNo 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   2760
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton YesNo 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   2400
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton YesNo 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   2040
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton YesNo 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   1680
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton YesNo 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   1320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton YesNo 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   360
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton AskOK 
         BackColor       =   &H0000FF00&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaskColor       =   &H0000C000&
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label AskSiteName 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   9
         Left            =   1080
         TabIndex        =   97
         Top             =   4320
         Width           =   5655
      End
      Begin VB.Label AskSiteName 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   8
         Left            =   1080
         TabIndex        =   96
         Top             =   3960
         Width           =   5655
      End
      Begin VB.Label AskSiteName 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   7
         Left            =   1080
         TabIndex        =   95
         Top             =   3600
         Width           =   5655
      End
      Begin VB.Label AskSiteName 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   6
         Left            =   1080
         TabIndex        =   94
         Top             =   3240
         Width           =   5655
      End
      Begin VB.Label AskSiteName 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   93
         Top             =   2880
         Width           =   5655
      End
      Begin VB.Label AskSiteName 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   92
         Top             =   2520
         Width           =   5655
      End
      Begin VB.Label AskSiteName 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   91
         Top             =   2160
         Width           =   5655
      End
      Begin VB.Label AskSiteName 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   90
         Top             =   1800
         Width           =   5655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Do you want to close:    (Click to toggle Yes/No)"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   79
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label AskSiteName 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   78
         Top             =   1440
         Width           =   5655
      End
      Begin VB.Label AskSiteName 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   77
         Top             =   1080
         Width           =   5655
      End
   End
   Begin VB.TextBox BackBox 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   120
      MouseIcon       =   "brw1.frx":030A
      MousePointer    =   99  'Custom
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7320
      Top             =   7560
   End
   Begin VB.ComboBox ToolBarList 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   315
      Left            =   10560
      TabIndex        =   57
      Text            =   "Top"
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox UrlBox 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   1680
      TabIndex        =   39
      ToolTipText     =   "Shift+Home Select Line, Ctrl+Enter Autocomplete"
      Top             =   720
      Width           =   8775
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4680
      Top             =   7560
   End
   Begin VB.PictureBox Drop 
      Height          =   320
      Left            =   8280
      Picture         =   "brw1.frx":074C
      ScaleHeight     =   255
      ScaleWidth      =   435
      TabIndex        =   38
      Top             =   7560
      Width           =   495
   End
   Begin VB.TextBox RecentBox 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   1095
      Left            =   1680
      MouseIcon       =   "brw1.frx":0E06
      MousePointer    =   99  'Custom
      MultiLine       =   -1  'True
      TabIndex        =   37
      Top             =   1080
      Visible         =   0   'False
      Width           =   8895
   End
   Begin VB.PictureBox Fullscreen 
      Height          =   375
      Left            =   5280
      Picture         =   "brw1.frx":1248
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   36
      ToolTipText     =   "Full screen"
      Top             =   7560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   30
      Left            =   4560
      Picture         =   "brw1.frx":194A
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   35
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   29
      Left            =   4560
      Picture         =   "brw1.frx":2124
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   34
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   28
      Left            =   4560
      Picture         =   "brw1.frx":28FE
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   33
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   27
      Left            =   4560
      Picture         =   "brw1.frx":30D8
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   32
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   26
      Left            =   4560
      Picture         =   "brw1.frx":15A1A
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   31
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   25
      Left            =   4560
      Picture         =   "brw1.frx":161F4
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   30
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   24
      Left            =   4560
      Picture         =   "brw1.frx":169CE
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   29
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   23
      Left            =   4560
      Picture         =   "brw1.frx":171A8
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   28
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   22
      Left            =   4560
      Picture         =   "brw1.frx":17982
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   27
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   21
      Left            =   4560
      Picture         =   "brw1.frx":1815C
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   26
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   20
      Left            =   4560
      Picture         =   "brw1.frx":18936
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   25
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   19
      Left            =   4560
      Picture         =   "brw1.frx":19110
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   24
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   18
      Left            =   4560
      Picture         =   "brw1.frx":198EA
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   23
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   17
      Left            =   4560
      Picture         =   "brw1.frx":1A0C4
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   22
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   16
      Left            =   4560
      Picture         =   "brw1.frx":1A89E
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   21
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   15
      Left            =   4560
      Picture         =   "brw1.frx":1B078
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   20
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   14
      Left            =   4560
      Picture         =   "brw1.frx":1B852
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   19
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   13
      Left            =   4560
      Picture         =   "brw1.frx":1C02C
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   18
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   12
      Left            =   4560
      Picture         =   "brw1.frx":1C806
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   17
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   11
      Left            =   4560
      Picture         =   "brw1.frx":1CFE0
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   16
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   10
      Left            =   4560
      Picture         =   "brw1.frx":1D7BA
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   15
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   9
      Left            =   4560
      Picture         =   "brw1.frx":1DF94
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   14
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   8
      Left            =   4560
      Picture         =   "brw1.frx":1E76E
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   13
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   7
      Left            =   4560
      Picture         =   "brw1.frx":1EF48
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   12
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   6
      Left            =   4560
      Picture         =   "brw1.frx":1F722
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   11
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   5
      Left            =   4560
      Picture         =   "brw1.frx":1FEFC
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   10
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   2
      Left            =   4560
      Picture         =   "brw1.frx":206D6
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   8
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   4
      Left            =   4560
      Picture         =   "brw1.frx":20EB0
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   7
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   3
      Left            =   4560
      Picture         =   "brw1.frx":2168A
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   6
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   1
      Left            =   4560
      Picture         =   "brw1.frx":21E64
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6000
      Top             =   7560
   End
   Begin VB.PictureBox Surf 
      AutoRedraw      =   -1  'True
      Height          =   380
      Index           =   0
      Left            =   4560
      Picture         =   "brw1.frx":2263E
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   4
      Top             =   7100
      Width           =   615
   End
   Begin VB.TextBox data 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   7680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6720
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open HTML file"
      Filter          =   "*.html"
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6480
      Left            =   1680
      TabIndex        =   1
      Top             =   1080
      Width           =   9765
      ExtentX         =   17224
      ExtentY         =   11430
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
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
   Begin VB.CommandButton Command1 
      Caption         =   "Go!"
      Height          =   345
      Left            =   9480
      MouseIcon       =   "brw1.frx":22E18
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   7680
      Width           =   690
   End
   Begin VB.PictureBox HorizTool 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   240
      MouseIcon       =   "brw1.frx":2325A
      MousePointer    =   99  'Custom
      Picture         =   "brw1.frx":2369C
      ScaleHeight     =   525
      ScaleWidth      =   11505
      TabIndex        =   40
      Top             =   120
      Width           =   11535
      Begin VB.Label HorizButt 
         BackStyle       =   0  'Transparent
         Height          =   615
         Index           =   15
         Left            =   0
         TabIndex        =   56
         Top             =   0
         Width           =   495
      End
      Begin VB.Label HorizButt 
         BackStyle       =   0  'Transparent
         Height          =   615
         Index           =   14
         Left            =   0
         TabIndex        =   55
         Top             =   0
         Width           =   495
      End
      Begin VB.Label HorizButt 
         BackStyle       =   0  'Transparent
         Height          =   615
         Index           =   13
         Left            =   0
         TabIndex        =   54
         Top             =   0
         Width           =   495
      End
      Begin VB.Label HorizButt 
         BackStyle       =   0  'Transparent
         Height          =   615
         Index           =   12
         Left            =   0
         TabIndex        =   53
         Top             =   0
         Width           =   495
      End
      Begin VB.Label HorizButt 
         BackStyle       =   0  'Transparent
         Height          =   615
         Index           =   11
         Left            =   0
         TabIndex        =   52
         Top             =   0
         Width           =   495
      End
      Begin VB.Label HorizButt 
         BackStyle       =   0  'Transparent
         Height          =   615
         Index           =   10
         Left            =   0
         TabIndex        =   51
         Top             =   0
         Width           =   495
      End
      Begin VB.Label HorizButt 
         BackStyle       =   0  'Transparent
         Height          =   615
         Index           =   9
         Left            =   0
         TabIndex        =   50
         Top             =   0
         Width           =   495
      End
      Begin VB.Label HorizButt 
         BackStyle       =   0  'Transparent
         Height          =   615
         Index           =   8
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Width           =   495
      End
      Begin VB.Label HorizButt 
         BackStyle       =   0  'Transparent
         Height          =   615
         Index           =   7
         Left            =   0
         TabIndex        =   48
         Top             =   0
         Width           =   495
      End
      Begin VB.Label HorizButt 
         BackStyle       =   0  'Transparent
         Height          =   615
         Index           =   6
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Width           =   495
      End
      Begin VB.Label HorizButt 
         BackStyle       =   0  'Transparent
         Height          =   615
         Index           =   5
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Width           =   495
      End
      Begin VB.Label HorizButt 
         BackStyle       =   0  'Transparent
         Height          =   615
         Index           =   4
         Left            =   0
         TabIndex        =   45
         Top             =   0
         Width           =   495
      End
      Begin VB.Label HorizButt 
         BackStyle       =   0  'Transparent
         Height          =   615
         Index           =   3
         Left            =   0
         TabIndex        =   44
         Top             =   0
         Width           =   495
      End
      Begin VB.Label HorizButt 
         BackStyle       =   0  'Transparent
         Height          =   615
         Index           =   2
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   495
      End
      Begin VB.Label HorizButt 
         BackStyle       =   0  'Transparent
         Height          =   615
         Index           =   1
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label HorizButt 
         BackStyle       =   0  'Transparent
         Height          =   615
         Index           =   0
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox VerTool 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      MouseIcon       =   "brw1.frx":2AB5E
      MousePointer    =   99  'Custom
      Picture         =   "brw1.frx":2AFA0
      ScaleHeight     =   495
      ScaleWidth      =   975
      TabIndex        =   58
      Top             =   600
      Visible         =   0   'False
      Width           =   975
      Begin VB.Label VertButt 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   15
         Left            =   0
         TabIndex        =   74
         Top             =   0
         Width           =   735
      End
      Begin VB.Label VertButt 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   14
         Left            =   0
         TabIndex        =   73
         Top             =   0
         Width           =   735
      End
      Begin VB.Label VertButt 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   13
         Left            =   0
         TabIndex        =   72
         Top             =   0
         Width           =   735
      End
      Begin VB.Label VertButt 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   12
         Left            =   0
         TabIndex        =   71
         Top             =   0
         Width           =   735
      End
      Begin VB.Label VertButt 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   11
         Left            =   0
         TabIndex        =   70
         Top             =   0
         Width           =   735
      End
      Begin VB.Label VertButt 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   10
         Left            =   0
         TabIndex        =   69
         Top             =   0
         Width           =   735
      End
      Begin VB.Label VertButt 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   9
         Left            =   0
         TabIndex        =   68
         Top             =   0
         Width           =   735
      End
      Begin VB.Label VertButt 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   8
         Left            =   0
         TabIndex        =   67
         Top             =   0
         Width           =   735
      End
      Begin VB.Label VertButt 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   7
         Left            =   0
         TabIndex        =   66
         Top             =   0
         Width           =   735
      End
      Begin VB.Label VertButt 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   6
         Left            =   0
         TabIndex        =   65
         Top             =   0
         Width           =   735
      End
      Begin VB.Label VertButt 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   5
         Left            =   0
         TabIndex        =   64
         Top             =   0
         Width           =   735
      End
      Begin VB.Label VertButt 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   4
         Left            =   0
         TabIndex        =   63
         Top             =   0
         Width           =   735
      End
      Begin VB.Label VertButt 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   62
         Top             =   0
         Width           =   735
      End
      Begin VB.Label VertButt 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   61
         Top             =   0
         Width           =   735
      End
      Begin VB.Label VertButt 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   60
         Top             =   0
         Width           =   735
      End
      Begin VB.Label VertButt 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   59
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   405
      Left            =   705
      TabIndex        =   2
      Top             =   720
      Width           =   915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim url(1 To 37) As String
Dim UrlName(1 To 37) As String
Dim INeedaName As Boolean

Dim X As Integer
Dim numurl As Integer
Dim NewSite As String
Dim Found As Boolean
Dim RecentOn As Boolean
Dim fav(1 To 100) As String
Private Declare Function GetCursorPos Lib "user32" (lpPoint As _
   POINTAPI) As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type
Dim a As POINTAPI
Dim b As Long
Dim c As Long
Dim formtop As Long
Dim numfavs As Integer
Dim curl As String
Dim frame As Integer
Dim newframe As Integer
Dim Spam As String
Dim BackTrack(1 To 10) As String
Dim BackTrackPtr As Integer
Dim BackUrl(0 To 7) As String
Dim Full As Boolean
Dim LargeButt As Boolean
Dim Cmd As String
Dim tmpurl As String
Dim ToolEdge As String
Dim hite As Long
Dim wid As Long
Dim Theme As String
Dim TmpUrlName As String
Dim DontAsk(1 To 25) As String
Dim AskSitePtr As Integer
Dim AskSite(1 To 10) As String
Dim Handle(1 To 10) As Long
Dim DontAskPtr As Integer
Dim LeaveIt As Boolean
Dim SiteName As String
Dim WebStatus As String
Dim BrwHandle As Long

Const GW_HWNDFIRST = 0
Const GW_HWNDNEXT = 2
Const GW_OWNER = 4
Const GWL_STYLE = -16
Const WS_DISABLED = &H8000000
Const WS_CANCELMODE = &H1F
Const WM_CLOSE = &H10

Private Declare Function GetWindow Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal wCmd As Integer) As Integer
    
Private Declare Function GetWindowLong _
    Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hwnd As Long, _
        ByVal nIndex As Long) As Long

Private Declare Function PostMessage _
    Lib "user32" Alias "PostMessageA" ( _
        ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

Private Declare Function GetWindowTextLength _
    Lib "user32" Alias "GetWindowTextLengthA" ( _
        ByVal hwnd As Long) As Long

Private Declare Function GetWindowText _
    Lib "user32" Alias "GetWindowTextA" ( _
        ByVal hwnd As Long, _
        ByVal lpString As String, _
        ByVal cch As Long) As Long

Private Declare Function IsWindow Lib "user32" ( _
    ByVal hwnd As Integer) As Integer

Sub CloseWindow(ByVal partialWindowCaption$)
    Dim Whnd As Long
    Dim dummy As String
    
    Dim L&
    Dim nam As String
    
    
    partialWindowCaption = LCase$(partialWindowCaption)
   Whnd = GetWindow(Form1.hwnd, GW_HWNDFIRST)


    Do While Whnd <> 0
    '    If IsWindow(Whnd) Then
            L = GetWindowTextLength(Whnd)
          
            If L > 0 Then
                nam = Space$(L + 1)
                L = GetWindowText(Whnd, nam, L + 1)
                nam = LCase$(Left$(nam, Len(nam) - 1))


     
                If InStr(nam, partialWindowCaption) Then
                     Select Case url(6)
                     Case "KILL"
                      EndTask Whnd
                    'Exit Do
                     Case "ASK"
                      'Has this site already been left open by the dude?
                      LeaveIt = False
                      
                      If DontAskPtr > 0 Then
                       For X = 1 To DontAskPtr
                         If DontAsk(X) = nam Then LeaveIt = True
                       Next X
                      End If
                      
                      If LeaveIt = False Then
                      'Is ask box visible?
                      
                      If AskBox.Visible = False Then
                         AskBox.Visible = True
                      End If
                      
                      'Get ptr
                      If AskSitePtr > 0 Then
                         'Is the window already listed?
                       Found = False
                       
                       For X = 1 To AskSitePtr
                         If nam = AskSite(X) Then
                           'yes its there
                           Found = True
                         End If
                       Next X
                       
                      If Found = False And AskSitePtr < 10 Then
                      'no, add it
                        AskSitePtr = AskSitePtr + 1
                        AskSite(AskSitePtr) = nam
                        AskSiteName(AskSitePtr - 1).Caption = Left$(nam, 78)
                        YesNo(AskSitePtr - 1).Visible = True
                        Handle(AskSitePtr) = Whnd
                        AskBox.Height = 1800 + AskSitePtr * 320
                        
                      End If
    
                       'adjust the hite of the frame
                    Else
                       'first one
                        AskSitePtr = 1
                        AskSite(AskSitePtr) = nam
                        AskSiteName(AskSitePtr - 1).Caption = Left$(nam, 78)
                        YesNo(AskSitePtr - 1).Visible = True
                        Handle(AskSitePtr) = Whnd
                        AskBox.Height = 1800 + AskSitePtr * 320

                      
                    End If

                      'AskBox.Visible = False
                      
                   End If
                     'Case "NO"
                     
                     End Select
                     
                End If
            End If
     '   End If
        Whnd = GetWindow(Whnd, GW_HWNDNEXT)
        DoEvents
    Loop
End Sub

Private Sub EndTask(Whnd As Long)
    If Whnd = Form1.hwnd Or _
        GetWindow(Whnd, GW_OWNER) _
            = Form1.hwnd Then End
    
    If (GetWindowLong(Whnd, GWL_STYLE) _
        And WS_DISABLED) Then Exit Sub
    
    PostMessage Whnd, WS_CANCELMODE, 0, 0&
    PostMessage Whnd, WM_CLOSE, 0, 0&
End Sub



Private Sub AskOK_Click()
For X = 1 To AskSitePtr
 If YesNo(X - 1).Caption = "Yes" Then
   'close it
    EndTask Handle(X)
 Else
   'put it on the don't ask don't tell list
    DontAskPtr = DontAskPtr + 1
    DontAsk(DontAskPtr) = AskSite(X)
    
End If

Next X
AskSitePtr = 0
AskBox.Visible = False
For X = 0 To 9
 YesNo(X).Visible = False
 AskSiteName(X).Caption = ""
 YesNo(X).Caption = "Yes"
 
Next X
AskBox.Height = 1800

End Sub

Private Sub BackBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'retieve site at mouse cursor
Dim Yy As Integer
If RecentOn = True Then Drop_Click

formtop = Form1.Top

'get the mouse y position, convert to array location
'save current url
Call mousepos(b, c)


Select Case ToolEdge
Case "Top"
formtop = formtop + 60
formtop = formtop * 1.08

c = c - 48
c = c * 16
c = c - formtop
Yy = c / 236

Case "Bottom"
formtop = formtop / 15
w = BackBox.Top / 15
c = c - formtop - w
Yy = ((c - 17) / 13)

Case "Left"
formtop = formtop / 15
c = c - formtop
Yy = ((c - 198) / 16)




Case "Right"

formtop = formtop / 15
c = c - formtop
Yy = ((c - 198) / 16)


End Select


If Yy <= BackTrackPtr Then
 UrlBox.Text = BackUrl(Yy)
 Command1_Click
End If
BackBox.Visible = False
Timer3.Enabled = False


End Sub

     Private Sub Command1_Click()
BackBox.Visible = False
If RecentOn = True Then Drop_Click


         If UrlBox.Text <> "" Then
             
               WebBrowser1.Navigate2 UrlBox.Text

             
             If WebBrowser1.Visible = False Then
                WebBrowser1.Visible = True
           End If
 End If
End Sub

    




Private Sub Drop_Click()
If RecentOn = True Then
 RecentBox.Visible = False
 RecentOn = False
 Timer2.Enabled = False
 
Else
 RecentBox.Visible = True
 RecentOn = True
 Timer2.Enabled = True
 
End If
BackBox.Visible = False

End Sub

Private Sub Form_Click()
BackBox.Visible = False
If RecentOn = True Then Drop_Click

End Sub

Private Sub Form_Load()
 Dim exists
 Dim RecStr As String
 
hite = Form1.Height
wid = Form1.Width

            'open the recent url list
RecentOn = False
           'change to homepage
           
           
          exists = Dir("c:\websurfer.ini")
 
 
          Select Case exists
          Case "websurfer.ini"
             Call LoadRecent
             
            Case Else
            


            Open "c:\websurfer.ini" For Output As #1
            url(1) = "http://www.yahoo.com"
            url(2) = "http://www.dogpile.com"
            url(3) = "http://www.espn.com"
            url(4) = "http://www.weather.com"
            url(5) = "http://www.cnn.com"
            url(6) = "ASK***default{{{Top"
            url(7) = "http://www.yahoo.com"
            For X = 1 To 7
            Print #1, url(X)
            Next X
            
            Close #1
            numurl = 7
            ToolEdge = "Top"
            url(6) = "ASK"
            Theme = "default"
            
 End Select

frame = 0
  RecentBox.Height = (numurl - 6) * 200 + 700
                UrlBox.Text = url(1)
 WebBrowser1.RegisterAsBrowser = True
WebBrowser1.RegisterAsDropTarget = True

'See if theres anything on the command line
   Cmd = command()
If Cmd <> "" Then
If Left$(Cmd, 1) = "/" Then Cmd = Right$(Cmd, Len(Cmd) - 1)
' / indicates url or file on command line


  If Left$(Cmd, 1) = Chr$(34) Then
   Cmd = Right$(Cmd, Len(Cmd) - 1)
  End If
  If Right$(Cmd, 1) = Chr$(34) Then
    Cmd = Left$(Cmd, Len(Cmd) - 1)
  End If
  
   
 


X = InStr(Cmd, "C:")
If X > 0 Then
Cmd = "file:///" + Cmd
For X = 1 To Len(Cmd)
If Mid$(Cmd, X, 1) = "\" Then Mid$(Cmd, X, 1) = "/"
'convert spaces to %20
If Mid$(Cmd, X, 1) = " " Then
  Mid$(Cmd, X, 1) = "0"
  temp$ = Mid$(Cmd, X, Len(Cmd) - X + 1)
  Cmd = Left$(Cmd, X - 1)
  Cmd = Cmd + "%2" + temp$
End If
Next X
End If

WebBrowser1.GoHome
While WebBrowser1.Busy = True
DoEvents
Wend

UrlBox.Text = Cmd



End If

'setup toolbar
HorizButt(0).ToolTipText = "Home"
HorizButt(1).ToolTipText = "Refresh"
HorizButt(2).ToolTipText = "Stop"
HorizButt(3).ToolTipText = "Catalog"
HorizButt(4).ToolTipText = "Search"
HorizButt(5).ToolTipText = "Back (Hold To Step Back)"
HorizButt(6).ToolTipText = "Forward"
HorizButt(7).ToolTipText = "Sports"
HorizButt(8).ToolTipText = "Weather"
HorizButt(9).ToolTipText = "News"
HorizButt(10).ToolTipText = "About Web Surfer/ Help"
HorizButt(11).ToolTipText = "Options"
HorizButt(12).ToolTipText = "Open File"
HorizButt(13).ToolTipText = "Fullscreen"
HorizButt(14).ToolTipText = "Toolbar Location"
HorizButt(15).ToolTipText = "Exit"
VertButt(0).ToolTipText = "Home"
VertButt(1).ToolTipText = "Refresh"
VertButt(2).ToolTipText = "Stop"
VertButt(3).ToolTipText = "Catalog"
VertButt(4).ToolTipText = "Search"
VertButt(5).ToolTipText = "Back (Hold To Step Back)"
VertButt(6).ToolTipText = "Forward"
VertButt(7).ToolTipText = "Sports"
VertButt(8).ToolTipText = "Weather"
VertButt(9).ToolTipText = "News"
VertButt(10).ToolTipText = "About Web Surfer/ Help"
VertButt(11).ToolTipText = "Options"
VertButt(12).ToolTipText = "Open File"
VertButt(13).ToolTipText = "Fullscreen"
VertButt(14).ToolTipText = "Toolbar Location"
VertButt(15).ToolTipText = "Exit"

For X = 0 To 15
HorizButt(X).Width = 200
Next X
VerTool.Top = 0
VerTool.Left = 0
VerTool.Height = 9000
VerTool.Width = 800

Call DrawToolBar

ToolBarList.AddItem "Top"
ToolBarList.AddItem "Bottom"
ToolBarList.AddItem "Left"
ToolBarList.AddItem "Right"
Form1.Show

             If WebBrowser1.Visible = False Then
                 WebBrowser1.Visible = True
             End If
     WebBrowser1.Navigate2 UrlBox.Text


BackTrackPtr = 0

AskSitePtr = 0


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackBox.Visible = False
If RecentOn = True Then Drop_Click

End Sub

Private Sub Form_Resize()
Dim X As Integer

hite = Form1.Height
wid = Form1.Width

If Full = True Then
If Form1.Height > 400 And Form1.Width > 80 Then
WebBrowser1.Width = Form1.Width - 80
WebBrowser1.Height = Form1.Height - 400



Fullscreen.Left = Form1.Width - 760


Form1.Refresh
End If

Else
  Call DrawToolBar



End If

End Sub

Private Sub Form_Unload(Cancel As Integer)



Open "c:\websurfer.ini" For Output As #1
For X = 1 To 5
 Print #1, url(X)
Next X
Print #1, url(6); "***"; Theme; "{{{"; ToolEdge


For X = 7 To numurl
 Print #1, url(X); "***"; UrlName(X)
 
Next X
Close #1
WebBrowser1.Visible = False


End
End Sub

Private Sub Fullscreen_Click()
Full = False


Label1.Visible = True
Drop.Visible = True
UrlBox.Visible = True
Command1.Visible = True

Surf(frame).Visible = True
Fullscreen.Visible = False
DrawToolBar






End Sub



Private Sub HorizButt_Click(Index As Integer)
'stop timer
Timer3.Enabled = False

Call ButtHandler(Index)

End Sub

Private Sub HorizButt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'start timer
If Index = 5 Then
 Timer3.Enabled = True
End If
End Sub



Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
BackBox.Visible = False
If RecentOn = True Then Drop_Click

End Sub

Private Sub RecentBox_Click()
Timer2.Enabled = False

Call mousepos(b, c)
Call GetRow(c)


Select Case ToolEdge
Case "Top"
Case "Bottom"
c = c - 1

Case "Left"
c = c + 2
Case "Right"
c = c + 2

End Select
If c < 1 Then c = 1



BackBox.Visible = False
If c > 31 Then c = 31

'show tooltip
If RecentOn = True Then Drop_Click

UrlBox.Text = url(c + 6)

Command1_Click
'Move URL to top of list, lest it be forgotton, unless it already is at top.
If c > 1 Then
tmpurl = url(c + 6)
TmpUrlName = UrlName(c + 6)

For X = c + 6 To 8 Step -1
 url(X) = url(X - 1)
 UrlName(X) = UrlName(X - 1)
 
Next X
url(7) = tmpurl
UrlName(7) = TmpUrlName
ShowRecent
End If

End Sub

Private Sub RecentBox_LostFocus()
If RecentOn = True Then Drop_Click

End Sub


Private Sub Timer1_Timer()
newframe = frame + 1
If newframe = 27 Then newframe = 0

Surf(newframe).Visible = True

Surf(frame).Visible = False
frame = newframe

End Sub

Private Sub Timer2_Timer()
Dim b As Long
Dim c As Long

'get x,y
Call mousepos(b, c)
Call GetRow(c)
Select Case ToolEdge
Case "Top"

Case "Left"
c = c + 2
Case "Right"
c = c + 2

Case "Bottom"
c = c - 1
End Select
If c < 1 Then c = 1
If c > 30 Then Exit Sub


'show tooltip

RecentBox.ToolTipText = UrlName(c + 6)




End Sub


Private Sub Timer3_Timer()
BackBox.Visible = True

End Sub

Private Sub ToolbarList_Click()
ToolEdge = ToolBarList.Text
ToolBarList.Visible = False
DrawToolBar
End Sub

Private Sub UrlBox_Click()

BackBox.Visible = False
If RecentOn = True Then Drop_Click


End Sub

Private Sub UrlBox_GotFocus()
BackBox.Visible = False
If RecentOn = True Then Drop_Click

End Sub

Private Sub UrlBox_KeyPress(KeyAscii As Integer)
Dim tmpurl As String

If KeyAscii = 13 Then
  
 tmpurl = UrlBox.Text
   Command1_Click
End If
If KeyAscii = 10 Then
   UrlBox.Text = "http://www." + UrlBox.Text + ".com"
 
 
 
 tmpurl = UrlBox.Text
   Command1_Click


End If
If KeyAscii = 10 Or KeyAscii = 13 Then
INeedaName = True
For X = 37 To 8 Step -1
url(X) = url(X - 1)
UrlName(X) = UrlName(X - 1)
Next X
url(7) = tmpurl
If numurl < 37 Then numurl = numurl + 1
  RecentBox.Height = (numurl - 6) * 200 + 700

Call ShowRecent
End If


End Sub

Private Sub VertButt_Click(Index As Integer)
'stop timer
Timer3.Enabled = False

Call ButtHandler(Index)

End Sub

Private Sub VertButt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'start timer
If Index = 5 Then
 Timer3.Enabled = True
End If

End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, url As Variant)

Dim X As Integer
Dim doc As Object
Dim dummy As String



'UrlBox.Text = WebBrowser1.LocationURL
Form1.Caption = WebBrowser1.LocationName
dummy = WebBrowser1.LocationName
If INeedaName = True Then


'  If WebBrowser1.LocationName = WebBrowser1.LocationURL Then


    'extract the name form the url
    X = InStr(1, dummy, "www")
     If X = 0 Then
          X = InStr(1, dummy, "//")
    End If

 
 If X > 0 Then
  
    Y = InStr(X + 4, dummy, ".")
  
    If Y > X Then dummy = Mid$(dummy, X + 4, Y - X)
  
  
 End If
  If Len(dummy) > 60 Then dummy = Left$(dummy, 60)
'remove any Comma or LF from string
 For X = 1 To Len(dummy)
  If Mid$(dummy, X, 1) = "," Or Mid$(dummy, X, 1) = Chr$(13) Then
   Mid$(dummy, X, 1) = " "
  End If
 Next X


  UrlName(7) = dummy
  INeedaName = False
End If

'move each one down
If BackTrackPtr > 0 Then
 For X = BackTrackPtr To 1 Step -1

   BackUrl(X) = BackUrl(X - 1)
   

 Next X
End If

BackUrl(0) = WebBrowser1.LocationURL
UrlBox.Text = WebBrowser1.LocationURL

If BackTrackPtr < 7 Then BackTrackPtr = BackTrackPtr + 1
BackBox.Height = BackTrackPtr * 256
If ToolEdge = "Bottom" Then

BackBox.Top = hite - BackBox.Height - 800
End If

dummy = BackUrl(1)
If Len(dummy) > 42 Then
'truncate
dummy = Left$(dummy, 20) + "..." + Right$(dummy, 19)

End If
BackBox.Text = dummy + vbCrLf

For X = 2 To BackTrackPtr
 dummy = BackUrl(X)
 If Len(dummy) > 30 Then
  'truncate
    dummy = Left$(dummy, 20) + "..." + Right$(dummy, 19)

 End If

 BackBox.Text = BackBox.Text + dummy + vbCrLf
Next X





End Sub

Private Sub WebBrowser1_DownloadBegin()
BackBox.Visible = False

If Full = False Then

Timer1.Enabled = True


End If

End Sub

Private Sub WebBrowser1_DownloadComplete()
             CloseWindow "Internet Explorer"
             CloseWindow "Netscape"
             
Timer1.Enabled = False



End Sub
Private Sub mousepos(b, c)
Dim ret As Integer

ret = GetCursorPos(a)
b = a.X
c = a.Y

End Sub


Private Sub WebBrowser1_GotFocus()
BackBox.Visible = False
ToolBarList.Visible = False

End Sub
Private Sub GetRow(c As Long)


If ToolEdge = "Bottom" Then
  c = c - 7 - (RecentBox.Top / 15) - (Form1.Top / 15)

End If
If ToolEdge = "Top" Then

c = c - 74 - (Form1.Top / 15)

End If
If ToolEdge = "Left" Then

c = c - 70 - (Form1.Top / 15)

End If
If ToolEdge = "Right" Then

c = c - 70 - (Form1.Top / 15)

End If

c = c / 13




End Sub

Private Sub LoadRecent()
    Dim RecStr  As String
    Dim j As Integer
    Dim k As Integer
    
    X = 1
            Open "c:\websurfer.ini" For Input As #1
            While Not EOF(1)
            Input #1, RecStr
            
              If RecStr <> "" And X < 37 Then
                If X < 7 Then
                 url(X) = RecStr
             
                Else
                 ptr = InStr(RecStr, "***")
                 If ptr = 0 Then
                 'not found
                 url(X) = RecStr
                 UrlName(X) = ""
                 
                 Else
                 'found ***
               
                 url(X) = Left$(RecStr, ptr - 1)
                 UrlName(X) = Right$(RecStr, Len(RecStr) - ptr - 2)
                             
                 End If
                End If
            X = X + 1

              End If
            Wend
            Close #1
            
'seperate the theme if any
j = InStr(url(6), "***")
k = InStr(url(6), "{{{")

   On Error GoTo Amy
   dummy = Mid$(url(6), j + 3, k - j - 3)
   ToolEdge = Mid$(url(6), k + 3, Len(url(6)) - k + 3)

   url(6) = Left$(url(6), j - 1)

Amy:
If ToolEdge <> "Top" And ToolEdge <> "Bottom" And ToolEdge <> "Left" And ToolEdge <> "Right" Then


 'invalid ini entry, fix it!
 ToolEdge = "Right"
End If

If url(6) <> "ASK" And url(6) <> "KILL" And url(6) <> "NO" Then
 'invalid INI entry, fix it
 url(6) = "ASK"
End If


On Error GoTo Cathy
If dummy <> "default" And dummy <> "" Then
  Open dummy For Input As #1
   Input #1, hfile
   Input #1, vfile
  Close #1
  HorizTool.Picture = LoadPicture(hfile)
  VerTool.Picture = LoadPicture(vfile)
End If
  Theme = dummy
  
GoTo Hannah
Cathy:
Close


Theme = "default"
Hannah:

            numurl = X - 1


            If numurl > 37 Then numurl = 37
            Call ShowRecent
            
                                     

End Sub
Private Sub ShowRecent()
              tmpurl = url(7)
             If Len(tmpurl) > 70 Then tmpurl = Left$(tmpurl, 70)
           
            RecentBox.Text = tmpurl + vbCrLf
      
            For X = 8 To numurl
             tmpurl = url(X)
             If Len(tmpurl) > 70 Then tmpurl = Left$(tmpurl, 70)
              RecentBox.Text = RecentBox.Text + tmpurl + vbCrLf
              tmpurl = UrlName(X)
               
             
            Next X

End Sub
Private Sub DrawToolBar()
'draws toolbar at any screen edge

'if minimized, exit or there is an error
If hite < 500 Then Exit Sub

Select Case ToolEdge
Case "Top"
WebBrowser1.Top = 900
WebBrowser1.Left = 0
WebBrowser1.Height = hite - 1320
WebBrowser1.Width = wid - 60

VerTool.Visible = False
HorizTool.Visible = True

HorizTool.Left = 0
HorizTool.Top = 0
HorizTool.Height = 560
HorizTool.Width = 12000
For X = 0 To 15
HorizButt(X).Left = X * 750
HorizButt(X).Top = 0
HorizButt(X).Height = 560
HorizButt(X).Width = 700

Next X
UrlBox.Top = 580
UrlBox.Left = 1600
Drop.Left = 10700
UrlBox.Width = 9120

Drop.Top = 560
RecentBox.Top = 800
RecentBox.Left = 1600
RecentBox.Width = 9120

For X = 0 To 30
Surf(X).Top = 560
Surf(X).Left = 20
Next X


Label1.Top = 580
Label1.Left = 740

BackBox.Top = 560
BackBox.Left = 3000


Command1.Top = 560
Command1.Left = 11200

ToolBarList.Left = 10200
ToolBarList.Top = 500

Case "Bottom"
VerTool.Visible = False
HorizTool.Visible = True

HorizTool.Left = 0
HorizTool.Top = hite - 1000
HorizTool.Height = 560
HorizTool.Width = 12000
For X = 0 To 15
HorizButt(X).Left = X * 750
HorizButt(X).Top = 0
HorizButt(X).Height = 560
HorizButt(X).Width = 700

Next X
UrlBox.Top = hite - 1340
UrlBox.Left = 1600
UrlBox.Width = 9120

Drop.Top = hite - 1360
Drop.Left = 10700
RecentBox.Top = hite - 1360 - RecentBox.Height
RecentBox.Left = 1600
RecentBox.Width = 9120

For X = 0 To 30
Surf(X).Top = hite - 1380
Surf(X).Left = 20
Next X
ToolBarList.Top = hite - 1600
ToolBarList.Left = 10200

Label1.Top = hite - 1300
Label1.Left = 700

BackBox.Top = hite - 1100 - BackBox.Height
BackBox.Left = 3000
Command1.Top = hite - 1360
Command1.Left = 11200

WebBrowser1.Top = 0
WebBrowser1.Height = Form1.Height - 1400
WebBrowser1.Width = Form1.Width
WebBrowser1.Left = 0

Case "Left"
VerTool.Visible = True
HorizTool.Visible = False
VerTool.Left = 0
For X = 0 To 15
VertButt(X).Left = 0
VertButt(X).Top = X * 550
VertButt(X).Height = 580
VertButt(X).Width = 700

Next X
UrlBox.Top = 40
UrlBox.Left = 2400
UrlBox.Width = 9120

Drop.Top = 0
Drop.Left = 11500
RecentBox.Top = 340
RecentBox.Left = 2400
RecentBox.Width = 9120

WebBrowser1.Top = 360
WebBrowser1.Left = 760
Command1.Left = 12040
Command1.Top = 0
WebBrowser1.Height = hite - 740
WebBrowser1.Width = wid - 840
Label1.Top = 40
Label1.Left = 1540
ToolBarList.Left = 740
ToolBarList.Top = 7400
For X = 0 To 30
Surf(X).Top = 0
Surf(X).Left = 800
Next X
BackBox.Left = 740
BackBox.Top = 2800

Case "Right"
VerTool.Visible = True
HorizTool.Visible = False
VerTool.Left = wid - 880
For X = 0 To 15
VertButt(X).Left = 0
VertButt(X).Top = X * 550
VertButt(X).Height = 580
VertButt(X).Width = 700

Next X

UrlBox.Top = 40
UrlBox.Left = 1680
UrlBox.Width = 9120

Drop.Top = 0
Drop.Left = 10800
RecentBox.Top = 340
RecentBox.Left = 1680
RecentBox.Width = 9120

WebBrowser1.Top = 360
WebBrowser1.Left = 0
Command1.Top = 0
Command1.Left = 11400
WebBrowser1.Height = hite - 740
WebBrowser1.Width = wid - 880
Label1.Top = 40
Label1.Left = 840
ToolBarList.Left = wid - 2100
ToolBarList.Top = 7400
For X = 0 To 30
Surf(X).Top = 0
Surf(X).Left = 60
Next X
BackBox.Left = wid - 5000
BackBox.Top = 2800


End Select

End Sub
Private Sub ButtHandler(Index As Integer)
Dim FiN As String

BackBox.Visible = False

If RecentOn = True Then Drop_Click

ToolBarList.Visible = False

If Favorites.Visible = True Then Exit Sub

On Error GoTo berr

Select Case Index
Case 5
WebBrowser1.GoBack

Case 6
WebBrowser1.GoForward
Case 0
UrlBox.Text = url(1)
Command1_Click
Case 1
WebBrowser1.Refresh2
Case 2
WebBrowser1.Stop
Case 10
data.Text = ""
FormAbout.Show
While FormAbout.Visible = True
DoEvents
Wend
If data.Text = "help" Then Call Help

Case 11
Open "c:\websurfer.ini" For Output As #1
For X = 1 To 5
 Print #1, url(X)
Next X

Print #1, url(6); "***"; Theme; "{{{"; ToolEdge

For X = 7 To numurl
 Print #1, url(X); "***"; UrlName(X)
 
Next X
Close #1


FormOption!HomeDepot.Text = UrlBox.Text
FormOption.Refresh




FormOption.Show
'wait for the chump to close the option window
While FormOption.Visible = True
DoEvents

Wend





'process data return

If data.Text = "OK" Then
 'update ini
 Call LoadRecent
Else
 If data.Text <> "CANCEL" Then
  If data.Text <> "" Then
  On Error GoTo themerr
  dummy = data.Text


  Open dummy For Input As #1
   Input #1, hfile
   Input #1, vfile
  Close #1
  HorizTool.Picture = LoadPicture(hfile)
  VerTool.Picture = LoadPicture(vfile)
  Theme = dummy
  End If
 End If
End If





Case 3

curl = WebBrowser1.LocationURL


Favorites!CurrentUrl.Text = curl
Favorites!SiteNameBox.Text = WebBrowser1.LocationName

Favorites.Show

While Favorites.Visible = True
DoEvents

Wend
Command1_Click



Case 4

 WebBrowser1.Navigate url(2)
UrlBox.Text = url(2)



Case 8
             WebBrowser1.Navigate url(4)
UrlBox.Text = url(4)

Case 9
             WebBrowser1.Navigate url(5)
  UrlBox.Text = url(5)
           

Case 7
             WebBrowser1.Navigate url(3)
             UrlBox.Text = url(3)

Case 12
   CommonDialog1.Filter = "Internet Files (html)|*.html|All Files|*.*|"
   CommonDialog1.FilterIndex = 1
   
   CommonDialog1.ShowOpen
   FiN = CommonDialog1.FileName
   If FiN <> "" Then
   UrlBox.Text = FiN
   

   Command1_Click
   
   End If
Case 13
HorizTool.Visible = False
VerTool.Visible = False

Label1.Visible = False
Drop.Visible = False
UrlBox.Visible = False
For X = 0 To 30
Surf(X).Visible = False
Next X
Command1.Visible = False
Form1.WindowState = 2

WebBrowser1.Top = 0
WebBrowser1.Left = 0
WebBrowser1.Width = Form1.Width - 80
WebBrowser1.Height = Form1.Height - 400
Fullscreen.Top = 0
Fullscreen.Left = Form1.Width - 760

Fullscreen.Visible = True
Full = True

Case 14
ToolBarList.Visible = True

Case 15

 Form_Unload (1)
 

End

End Select
berr:

Exit Sub
themerr:
MsgBox "There was an error loading the toolbar theme."

End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)

WebStatus = Text

Form1.Caption = SiteName & "..." & WebStatus

End Sub

Private Sub WebBrowser1_TitleChange(ByVal Text As String)
SiteName = Text
Form1.Caption = SiteName & "..." & WebStatus

End Sub

Private Sub YesNo_Click(Index As Integer)
If YesNo(Index).Caption = "Yes" Then
  YesNo(Index).Caption = "No"
Else
  YesNo(Index).Caption = "Yes"
End If

End Sub
Private Sub Help()
UrlBox.Text = "file:///C:/web%20surfer/help.html"
Command1_Click

End Sub
