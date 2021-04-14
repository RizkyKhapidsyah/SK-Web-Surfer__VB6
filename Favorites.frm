VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Favorites 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catalog of Web Sites"
   ClientHeight    =   8535
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   11910
   ControlBox      =   0   'False
   Icon            =   "Favorites.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Favorites.frx":0442
   ScaleHeight     =   8535
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox SearchButt 
      Height          =   520
      Left            =   9000
      Picture         =   "Favorites.frx":75B84
      ScaleHeight     =   465
      ScaleWidth      =   600
      TabIndex        =   51
      ToolTipText     =   "Search"
      Top             =   2160
      Width           =   660
   End
   Begin VB.TextBox SearchText 
      Height          =   285
      Left            =   4800
      TabIndex        =   50
      ToolTipText     =   "Enter search string"
      Top             =   2280
      Width           =   3855
   End
   Begin VB.CommandButton Delete 
      Caption         =   "Delete from Catalog"
      Height          =   375
      Left            =   8640
      MouseIcon       =   "Favorites.frx":7649E
      MousePointer    =   99  'Custom
      TabIndex        =   49
      Top             =   7800
      Width           =   2655
   End
   Begin VB.PictureBox Cancelpress 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   9000
      MouseIcon       =   "Favorites.frx":768E0
      MousePointer    =   99  'Custom
      Picture         =   "Favorites.frx":76D22
      ScaleHeight     =   495
      ScaleWidth      =   1500
      TabIndex        =   13
      Top             =   1320
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox OKpress 
      BorderStyle     =   0  'None
      Height          =   530
      Left            =   9000
      MouseIcon       =   "Favorites.frx":77F10
      MousePointer    =   99  'Custom
      Picture         =   "Favorites.frx":78352
      ScaleHeight     =   525
      ScaleWidth      =   1500
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox Cancel 
      BorderStyle     =   0  'None
      Height          =   530
      Left            =   9000
      MouseIcon       =   "Favorites.frx":79540
      MousePointer    =   99  'Custom
      Picture         =   "Favorites.frx":79982
      ScaleHeight     =   525
      ScaleWidth      =   1500
      TabIndex        =   11
      Top             =   1320
      Width           =   1500
   End
   Begin VB.PictureBox OK 
      BorderStyle     =   0  'None
      Height          =   530
      Left            =   9000
      MouseIcon       =   "Favorites.frx":7AB70
      MousePointer    =   99  'Custom
      Picture         =   "Favorites.frx":7AFB2
      ScaleHeight     =   525
      ScaleWidth      =   1500
      TabIndex        =   10
      Top             =   480
      Width           =   1500
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   510
      Left            =   3968
      TabIndex        =   6
      Top             =   7800
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   900
      _Version        =   327682
      MousePointer    =   99
      MouseIcon       =   "Favorites.frx":7C1A0
      LargeChange     =   1
      Min             =   1
      Max             =   13
      SelStart        =   1
      Value           =   1
   End
   Begin VB.TextBox SiteNameBox 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   6975
   End
   Begin VB.TextBox CurrentUrl 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   6975
   End
   Begin VB.CommandButton Current 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Add the Current Site to the Catalog"
      Height          =   375
      Left            =   240
      MouseIcon       =   "Favorites.frx":7C4BA
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   6975
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4800
      Left            =   1235
      MouseIcon       =   "Favorites.frx":7C8FC
      MousePointer    =   99  'Custom
      Picture         =   "Favorites.frx":7CD3E
      ScaleHeight     =   4800
      ScaleWidth      =   9435
      TabIndex        =   8
      Top             =   2760
      Width           =   9440
      Begin VB.Label Fetch 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   48
         Top             =   4200
         Width           =   255
      End
      Begin VB.Label Fetch 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   47
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label Fetch 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   46
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label Fetch 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   45
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label Fetch 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   44
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label Fetch 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   43
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Fetch 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   42
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Fetch 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   41
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Fetch 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   40
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Fetch 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   39
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Fetch 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   38
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Fetch 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Favurl 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   11
         Left            =   5040
         TabIndex        =   36
         Top             =   4200
         Width           =   3975
      End
      Begin VB.Label Favurl 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   10
         Left            =   5040
         TabIndex        =   35
         Top             =   3840
         Width           =   3975
      End
      Begin VB.Label Favurl 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   9
         Left            =   5040
         TabIndex        =   34
         Top             =   3480
         Width           =   3975
      End
      Begin VB.Label Favurl 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   8
         Left            =   5040
         TabIndex        =   33
         Top             =   3120
         Width           =   3975
      End
      Begin VB.Label Favurl 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   7
         Left            =   5040
         TabIndex        =   32
         Top             =   2760
         Width           =   3975
      End
      Begin VB.Label Favurl 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   6
         Left            =   5040
         TabIndex        =   31
         Top             =   2400
         Width           =   3975
      End
      Begin VB.Label Favurl 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   5
         Left            =   5040
         TabIndex        =   30
         Top             =   2040
         Width           =   3975
      End
      Begin VB.Label Favurl 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   5040
         TabIndex        =   29
         Top             =   1680
         Width           =   3975
      End
      Begin VB.Label Favurl 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   5040
         TabIndex        =   28
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Label Favurl 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   5040
         TabIndex        =   27
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label Favurl 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   5040
         TabIndex        =   26
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Favurl 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   5040
         TabIndex        =   25
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label FavName 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   24
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label FavName 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   11
         Left            =   600
         TabIndex        =   23
         Top             =   4200
         Width           =   3975
      End
      Begin VB.Label FavName 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   10
         Left            =   600
         TabIndex        =   22
         Top             =   3840
         Width           =   3975
      End
      Begin VB.Label FavName 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   9
         Left            =   600
         TabIndex        =   21
         Top             =   3480
         Width           =   3975
      End
      Begin VB.Label FavName 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   8
         Left            =   600
         TabIndex        =   20
         Top             =   3120
         Width           =   3975
      End
      Begin VB.Label FavName 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   7
         Left            =   600
         TabIndex        =   19
         Top             =   2760
         Width           =   3975
      End
      Begin VB.Label FavName 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   18
         Top             =   2400
         Width           =   3975
      End
      Begin VB.Label FavName 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   5
         Left            =   600
         TabIndex        =   17
         Top             =   2040
         Width           =   3975
      End
      Begin VB.Label FavName 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   16
         Top             =   1680
         Width           =   3975
      End
      Begin VB.Label FavName 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   15
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Label FavName 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   14
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label FavName 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   600
         MouseIcon       =   "Favorites.frx":AE780
         TabIndex        =   9
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Label PageCount 
      BackStyle       =   0  'Transparent
      Caption         =   "Page 1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   3840
      Top             =   7680
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Site Name:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select from the Catalog, or search for:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   2280
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Site address:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Favorites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fav(1 To 156) As String
Dim FavNames(1 To 156) As String
Dim Page As Integer
Dim dummy As String
Dim numfavs As Integer
Dim SiteName As String

Option Explicit


Private Sub Cancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Cancelpress.Visible = True
Cancel.Visible = False

End Sub

Private Sub Cancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Favorites.Hide
Form_Unload (1)

End Sub

Private Sub Cancelpress_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Cancelpress.Visible = True
Cancel.Visible = False

End Sub

Private Sub Cancelpress_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Favorites.Hide
Cancelpress.Visible = False
Cancel.Visible = True

Form_Unload (1)

End Sub

Private Sub Current_Click()
Dim z As Integer
Dim X As Integer
Dim nam As String

'MsgBox numfavs

If numfavs < 156 And CurrentUrl.Text <> "" Then
nam = SiteNameBox.Text
If Len(nam) > 40 Then
 nam = Left$(nam, 40)
End If

Open "c:\favs.txt" For Output As #1
 

For X = 1 To numfavs
 Print #1, FavNames(X)
 Print #1, fav(X)
Next X
 'new page
Print #1, nam
Print #1, CurrentUrl.Text
Close #1
fav(X) = CurrentUrl.Text
FavNames(X) = nam
Page = ((X - 1) \ 12) + 1
Slider1.Value = Page

PageCount.Caption = "Page " + Str(Slider1.Value)

numfavs = numfavs + 1
'need to pass this to main

ShowFavs
End If

End Sub

Private Sub Delete_Click()
Dim X As Integer
Dim entry As Integer
'delete entry

entry = 12
'see if anyhting is checked

For X = 0 To 11
 If Fetch(X).Caption = "1" Then entry = X
Next X
If entry = 12 Then Exit Sub

'current location is (entry + 1 + (Page - 1) * 12)
For X = entry + 1 + (Page - 1) * 12 To numfavs
fav(X) = fav(X + 1)
FavNames(X) = FavNames(X + 1)
Next X
fav(X) = ""
FavNames(X) = ""

numfavs = numfavs - 1
ShowFavs



End Sub

Private Sub FavName_Click(Index As Integer)
If Favurl(Index).Caption <> "" Then

ClearFetch
CurrentUrl.Text = fav(Index + 1 + (Page - 1) * 12)
SiteNameBox.Text = FavNames(Index + 1 + (Page - 1) * 12)

Fetch(Index).Caption = "1"
End If

End Sub


Private Sub Favurl_Click(Index As Integer)
If Favurl(Index).Caption <> "" Then
ClearFetch
CurrentUrl.Text = fav(Index + 1 + (Page - 1) * 12)
SiteNameBox.Text = FavNames(Index + 1 + (Page - 1) * 12)

Fetch(Index).Caption = "1"
End If

End Sub

Private Sub Form_Load()
Dim exists
Dim X As Integer
Favorites.Top = 0
Favorites.Left = 0
For X = 0 To 11
FavName(X).Caption = ""
Favurl(X).Caption = ""
Fetch(X).Caption = ""
Next X


Page = 1
' show current favorites

    
 exists = Dir("c:\favs.txt")
 Select Case exists
 Case "favs.txt"
 
   Open "c:\favs.txt" For Input As #1
X = 1
   While Not EOF(1)
    Input #1, FavNames(X)
If Not EOF(1) Then Input #1, fav(X)
  
    
    X = X + 1
Wend
Close #1
numfavs = X - 1

Case Else


nofav:
Open "c:\favs.txt" For Output As #1
'set the default homepage
Print #1, "Yahoo!"

Print #1, "http://www.yahoo.com"
Close #1
numfavs = 1




End Select
ShowFavs
SiteName = Form1!WebBrowser1.LocationName
SiteNameBox.Text = SiteName

End Sub

Private Sub Form_Unload(Cancel As Integer)
Favorites.Visible = False

End Sub

Private Sub OK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
OKpress.Visible = True
OK.Visible = False

End Sub

Private Sub OK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1!UrlBox.Text = CurrentUrl.Text
'save favs
Open "c:\favs.txt" For Output As #1
For X = 1 To numfavs
Print #1, FavNames(X)
Print #1, fav(X)
Next X
Close #1
Favorites.Hide
Form_Unload (1)

End Sub

Private Sub ShowFavs()
Dim z As Integer
Dim X As Integer

ClearFetch
For X = 1 To 12

   FavName(X - 1).Caption = FavNames((Page - 1) * 12 + X)
    Favurl(X - 1).Caption = fav((Page - 1) * 12 + X)
Next X


End Sub

Private Sub OKpress_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
OKpress.Visible = True
OK.Visible = False

End Sub

Private Sub OKpress_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Favorites.Hide
OKpress.Visible = False
OK.Visible = True

Form_Unload (1)

End Sub

Private Sub SearchButt_Click()
Dim X As Integer
Dim Found As String
Dim UFound As String
Dim NoMatch As Boolean

Dim q As String
Dim ptr As Integer
NoMatch = True

dummy = SearchText.Text
If dummy = "" Then Exit Sub
X = 1

selop:
Found = InStr(fav(X), dummy)
UFound = InStr(FavNames(X), dummy)

If Found > 0 Or UFound > 0 Then
'turn to the page
  Page = Int((X - 1) / 12) + 1
  ShowFavs
  NoMatch = False
  

    ptr = X - (Page - 1) * 12

    Slider1.Value = Page
    PageCount.Caption = "Page " + Str(Slider1.Value)
    CurrentUrl.Text = fav(X)
    SiteNameBox.Text = FavNames(X)

    Fetch(ptr - 1).Caption = "1"

    q = MsgBox("Found match. Continue search?", vbYesNo, "Catalog Search for " & dummy, 1, 1)
    If q <> 6 Then Exit Sub
    Fetch(ptr - 1).Caption = ""

End If
X = X + 1
If X < 157 Then GoTo selop
If NoMatch = True Then
      q = MsgBox("No match found!", vbExclamation, "Catalog Search for " & dummy, 1, 1)
End If

End Sub

Private Sub SearchText_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 SearchButt_Click
End If

 
End Sub

Private Sub Slider1_Change()
'get the slider1 value, put it in page box
ClearFetch

Page = Slider1.Value
PageCount.Caption = "Page " + Str(Slider1.Value)
ShowFavs

End Sub

Private Sub ClearFetch()
Dim X As Integer
For X = 0 To 11
Fetch(X).Caption = ""
Next X

End Sub
