VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FormOption 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00A08000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11235
   ControlBox      =   0   'False
   Icon            =   "FormOption.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormOption.frx":0442
   ScaleHeight     =   7740
   ScaleWidth      =   11235
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      InitDir         =   "c:\web surfer"
   End
   Begin VB.TextBox SearchSite 
      Height          =   375
      Left            =   4080
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   5880
      Width           =   3975
   End
   Begin VB.TextBox NewSite 
      Height          =   375
      Left            =   4080
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   5280
      Width           =   3975
   End
   Begin VB.TextBox WeathSite 
      Height          =   375
      Left            =   4080
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   4680
      Width           =   3975
   End
   Begin VB.TextBox SportSite 
      Height          =   375
      Left            =   4080
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   4080
      Width           =   3975
   End
   Begin VB.PictureBox Cancel 
      BorderStyle     =   0  'None
      Height          =   530
      Left            =   9120
      MouseIcon       =   "FormOption.frx":75B84
      MousePointer    =   99  'Custom
      Picture         =   "FormOption.frx":75FC6
      ScaleHeight     =   525
      ScaleWidth      =   1500
      TabIndex        =   6
      Top             =   2040
      Width           =   1500
   End
   Begin VB.PictureBox OK 
      BorderStyle     =   0  'None
      Height          =   530
      Left            =   9120
      MouseIcon       =   "FormOption.frx":771B4
      MousePointer    =   99  'Custom
      Picture         =   "FormOption.frx":775F6
      ScaleHeight     =   525
      ScaleWidth      =   1500
      TabIndex        =   5
      Top             =   960
      Width           =   1500
   End
   Begin VB.PictureBox Cancelpress 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   9120
      MouseIcon       =   "FormOption.frx":787E4
      MousePointer    =   99  'Custom
      Picture         =   "FormOption.frx":78C26
      ScaleHeight     =   495
      ScaleWidth      =   1500
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox OKpress 
      BorderStyle     =   0  'None
      Height          =   530
      Left            =   9120
      MouseIcon       =   "FormOption.frx":79E14
      MousePointer    =   99  'Custom
      Picture         =   "FormOption.frx":7A256
      ScaleHeight     =   525
      ScaleWidth      =   1500
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox HomeDepot 
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   2040
      Width           =   4455
   End
   Begin VB.Label Theme 
      BackStyle       =   0  'Transparent
      Caption         =   "Load Toolbar Theme"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   960
      MouseIcon       =   "FormOption.frx":7B444
      MousePointer    =   99  'Custom
      TabIndex        =   32
      Top             =   6720
      Width           =   2775
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   11040
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label ChangeSearch 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Search Site to:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   960
      MouseIcon       =   "FormOption.frx":7B886
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   5880
      Width           =   2775
   End
   Begin VB.Label SearchBox 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   360
      MouseIcon       =   "FormOption.frx":7BCC8
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label AskBox 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   8760
      MouseIcon       =   "FormOption.frx":7C10A
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label AllowBox 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   8760
      MouseIcon       =   "FormOption.frx":7C54C
      MousePointer    =   99  'Custom
      TabIndex        =   28
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label KillBox 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   8760
      MouseIcon       =   "FormOption.frx":7C98E
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Allow 
      BackStyle       =   0  'Transparent
      Caption         =   "Allow New Windows"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   9360
      MouseIcon       =   "FormOption.frx":7CDD0
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Ask 
      BackStyle       =   0  'Transparent
      Caption         =   "Ask "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   9360
      MouseIcon       =   "FormOption.frx":7D212
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Kill 
      BackStyle       =   0  'Transparent
      Caption         =   "Exterminate all New Windows"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   9360
      MouseIcon       =   "FormOption.frx":7D654
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ad Control Options"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   9000
      TabIndex        =   22
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   8640
      X2              =   8640
      Y1              =   720
      Y2              =   6840
   End
   Begin VB.Label NewsBox 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   360
      MouseIcon       =   "FormOption.frx":7DA96
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label WeathBox 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   360
      MouseIcon       =   "FormOption.frx":7DED8
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label SportsBox 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   360
      MouseIcon       =   "FormOption.frx":7E31A
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label ChangeNews 
      BackStyle       =   0  'Transparent
      Caption         =   "Change News Site to:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   960
      MouseIcon       =   "FormOption.frx":7E75C
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   5280
      Width           =   2775
   End
   Begin VB.Label ChangeWeath 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Weather Site to:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   960
      MouseIcon       =   "FormOption.frx":7EB9E
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Label ChangeSport 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Sports Site to:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   960
      MouseIcon       =   "FormOption.frx":7EFE0
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Label SetClear 
      BackStyle       =   0  'Transparent
      Caption         =   "Clear History"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   960
      MouseIcon       =   "FormOption.frx":7F422
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label ClearBox 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   360
      MouseIcon       =   "FormOption.frx":7F864
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label HomeBox 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   360
      MouseIcon       =   "FormOption.frx":7FCA6
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label SetCurrentBox 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   360
      MouseIcon       =   "FormOption.frx":800E8
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label ChangeHome 
      BackStyle       =   0  'Transparent
      Caption         =   "Set the Home Page to:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   960
      MouseIcon       =   "FormOption.frx":8052A
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label SetCurrent 
      BackStyle       =   0  'Transparent
      Caption         =   "Set the Home Page to the Current Page"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   960
      MouseIcon       =   "FormOption.frx":8096C
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1320
      Width           =   4815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   11040
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   10920
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Home Page"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   7215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Web Surfer Options"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   3390
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "FormOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim url(1 To 37) As String
Dim command As String
Dim numurl As Integer

Dim X As Integer

Private Sub Allow_Click()
AllowBox.Caption = "="
KillBox.Caption = "1"
AskBox.Caption = "1"

End Sub

Private Sub AllowBox_Click()
Allow_Click

End Sub

Private Sub Ask_Click()
AllowBox.Caption = "1"
KillBox.Caption = "1"
AskBox.Caption = "="

End Sub

Private Sub AskBox_Click()
Ask_Click

End Sub

Private Sub Cancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Cancelpress.Visible = True
Cancel.Visible = False
SetCurrentBox.Caption = "1"
ClearBox.Caption = "1"
HomeBox.Caption = "1"
Form1!data.Text = "CANCEL"

End Sub

Private Sub Cancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)




Cancelpress.Visible = False
Cancel.Visible = True

FormOption.Hide
FormOption.Visible = False

End Sub

Private Sub Cancelpress_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Cancelpress.Visible = True
Cancel.Visible = False

End Sub

Private Sub Cancelpress_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1!data.Text = ""
Cancelpress.Visible = False
Cancel.Visible = True

FormOption.Hide
FormOption.Visible = False

End Sub

Private Sub ChangeHome_Click()
If HomeBox.Caption = "=" Then
 HomeBox.Caption = "1"
Else
 HomeBox.Caption = "="
End If
SetCurrentBox.Caption = "1"

End Sub

Private Sub ChangeNews_Click()
If NewsBox.Caption = "=" Then
 NewsBox.Caption = "1"
Else
 NewsBox.Caption = "="
End If

End Sub

Private Sub ChangeSearch_Click()
If SearchBox.Caption = "=" Then
 SearchBox.Caption = "1"
Else
 SearchBox.Caption = "="
End If


End Sub

Private Sub ChangeSport_Click()
If SportsBox.Caption = "=" Then
 SportsBox.Caption = "1"
Else
 SportsBox.Caption = "="
End If

End Sub

Private Sub ChangeWeath_Click()
If WeathBox.Caption = "=" Then
 WeathBox.Caption = "1"
Else
 WeathBox.Caption = "="
End If

End Sub

Private Sub ClearBox_Click()
SetClear_Click

End Sub


Private Sub Form_Activate()

 Open "c:\websurfer.ini" For Input As #1
           X = 1
           While Not EOF(1)
            Input #1, url(X)
             X = X + 1
             
           Wend
           
            Close #1
            numurl = X - 1
            
clearall

SportSite.Text = url(3)
NewSite.Text = url(5)
WeathSite.Text = url(4)
SearchSite.Text = url(2)
If InStr(url(6), "KILL") > 0 Then KillBox.Caption = "="
If InStr(url(6), "ASK") > 0 Then AskBox.Caption = "="
If InStr(url(6), "NO") > 0 Then Allow.Caption = "="

End Sub

Private Sub Form_Unload(Cancel As Integer)
'End


End Sub


Private Sub HomeBox_Click()
ChangeHome_Click

End Sub

Private Sub Kill_Click()
AllowBox.Caption = "1"
KillBox.Caption = "="
AskBox.Caption = "1"

End Sub

Private Sub KillBox_Click()
Kill_Click

End Sub


Private Sub NewsBox_Click()
ChangeNews_Click

End Sub

Private Sub OK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
OKpress.Visible = True
OK.Visible = False

End Sub

Private Sub OK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
OKpress.Visible = False
OK.Visible = True


Form1!data.Text = ""
If SetCurrentBox.Caption = "=" Then
url(1) = Form1!UrlBox.Text


End If
If HomeBox.Caption = "=" Then
url(1) = HomeDepot.Text

End If

If ClearBox.Caption = "=" Then
 command = "CLEAR"
 
End If

If NewsBox.Caption = "=" Then
url(5) = NewSite.Text

End If

If SportsBox.Caption = "=" Then
url(3) = SportSite.Text
End If

If WeathBox.Caption = "=" Then
url(4) = WeathSite.Text

End If
If SearchBox.Caption = "=" Then
url(2) = SearchSite.Text

End If
X = InStr(url(6), "***")
url(6) = Right$(url(6), Len(url(6)) - X + 1)



If AskBox.Caption = "=" Then url(6) = "ASK" + url(6)
If KillBox.Caption = "=" Then url(6) = "KILL" + url(6)
If AllowBox.Caption = "=" Then url(6) = "NO" + url(6)

'saveini
Open "C:\websurfer.ini" For Output As #1

For X = 1 To 6
Print #1, url(X)
Next X
If command <> "CLEAR" Then
For X = 7 To numurl
Print #1, url(X)
'this includes ***name but we wont use

Next X
End If
Close #1


Form1!data.Text = "OK"

FormOption.Hide


End Sub

Private Sub OKpress_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
OKpress.Visible = True
OK.Visible = False

End Sub

Private Sub OKpress_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
OKpress.Visible = False
OK.Visible = True



FormOption.Hide
FormOption.Visible = False

End Sub

Private Sub OptionCurrent_Click()

 If OptionSet.Value = True Then OptionSet.Value = False


End Sub

Private Sub OptionCurrent_DblClick()
OptionCurrent.Value = False
End Sub

Private Sub OptionSet_Click()
 If OptionCurrent.Value = True Then OptionCurrent.Value = False

End Sub

Private Sub OptionSet_DblClick()
 OptionSet.Value = False
 
End Sub

Private Sub SearchBox_Click()
ChangeSearch_Click

End Sub

Private Sub SetClear_Click()
If ClearBox.Caption = "=" Then
  ClearBox.Caption = "1"
Else
  ClearBox.Caption = "="
End If
End Sub

Private Sub SetCurrent_Click()
If SetCurrentBox.Caption = "=" Then
   SetCurrentBox.Caption = "1"
Else
   SetCurrentBox.Caption = "="
End If
HomeBox.Caption = "1"

End Sub
Private Sub clearall()
ClearBox.Caption = "1"
HomeBox.Caption = "1"
SetCurrentBox.Caption = "1"
NewsBox.Caption = "1"
WeathBox.Caption = "1"
SportsBox.Caption = "1"
AskBox.Caption = "1"
AllowBox.Caption = "1"
KillBox.Caption = "1"



Select Case url(6)


Case "NO"
AllowBox.Caption = "="
Case "ASK"
AskBox.Caption = "="
Case "KILL"
KillBox.Caption = "="
End Select

End Sub

Private Sub SetCurrentBox_Click()
SetCurrent_Click

End Sub

Private Sub SportsBox_Click()
ChangeSport_Click

End Sub

Private Sub Theme_Click()
'show common dialog open
   CommonDialog1.Filter = "Toolbar Themes|*.tbt|"
   CommonDialog1.FilterIndex = 1
   
   CommonDialog1.ShowOpen
   FiN = CommonDialog1.FileName
   If FiN <> "" Then
   ThemeFile = FiN
   Form1!data.Text = FiN

   FormOption.Hide

   End If

End Sub

Private Sub WeathBox_Click()
ChangeWeath_Click

End Sub
