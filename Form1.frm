VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Visual Basic Template Manager"
   ClientHeight    =   3945
   ClientLeft      =   300
   ClientTop       =   345
   ClientWidth     =   6600
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1680
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   38
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1224
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2448
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":366C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4890
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5AB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7EFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9120
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A344
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5741
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HotTracking     =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00E0E0E0&
      Visible         =   0   'False
      X1              =   3240
      X2              =   3240
      Y1              =   3480
      Y2              =   3840
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00E0E0E0&
      Visible         =   0   'False
      X1              =   4800
      X2              =   3240
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00808080&
      Visible         =   0   'False
      X1              =   3240
      X2              =   4800
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00808080&
      Visible         =   0   'False
      X1              =   4800
      X2              =   4800
      Y1              =   3480
      Y2              =   3840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Kill Template"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   3555
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      Visible         =   0   'False
      X1              =   4920
      X2              =   4920
      Y1              =   3480
      Y2              =   3840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      Visible         =   0   'False
      X1              =   6480
      X2              =   4920
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      Visible         =   0   'False
      X1              =   4920
      X2              =   6480
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      Visible         =   0   'False
      X1              =   6480
      X2              =   6480
      Y1              =   3480
      Y2              =   3840
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Add Template"
      Height          =   255
      Left            =   4920
      TabIndex        =   2
      Top             =   3555
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Raise Level"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   3555
      Width           =   1575
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      Visible         =   0   'False
      X1              =   3120
      X2              =   3120
      Y1              =   3480
      Y2              =   3840
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00808080&
      Visible         =   0   'False
      X1              =   1560
      X2              =   3120
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00E0E0E0&
      Visible         =   0   'False
      X1              =   3120
      X2              =   1560
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00E0E0E0&
      Visible         =   0   'False
      X1              =   1560
      X2              =   1560
      Y1              =   3480
      Y2              =   3840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelStat As Boolean
Dim myname, mypath As String
Sub BuildListView()
On Error Resume Next
Dim i As Integer
Dim ii As Integer
Screen.MousePointer = 11
ListView1.ListItems.Clear
myname = Dir(mypath & "/*.frm", vbNormal)
Do While myname <> ""
    If myname <> "." And myname <> ".." Then
        If (GetAttr(mypath & myname) And vbNormal) = vbNormal Then
            i = i + 1
            ListView1.ListItems.Add i, , myname, 1, 1
        End If
    End If
    myname = Dir
Loop
myname = Dir(mypath & "/*.dob", vbNormal)
Do While myname <> ""
    If myname <> "." And myname <> ".." Then
        If (GetAttr(mypath & myname) And vbNormal) = vbNormal Then
            i = i + 1
            ListView1.ListItems.Add i, , myname, 2, 2
        End If
    End If
    myname = Dir
Loop
myname = Dir(mypath & "/*.cls", vbNormal)
Do While myname <> ""
    If myname <> "." And myname <> ".." Then
        If (GetAttr(mypath & myname) And vbNormal) = vbNormal Then
            i = i + 1
            ListView1.ListItems.Add i, , myname, 3, 3
        End If
    End If
    myname = Dir
Loop
myname = Dir(mypath & "/*.ctl", vbNormal)
Do While myname <> ""
    If myname <> "." And myname <> ".." Then
        If (GetAttr(mypath & myname) And vbNormal) = vbNormal Then
            i = i + 1
            ListView1.ListItems.Add i, , myname, 4, 4
        End If
    End If
    myname = Dir
Loop
myname = Dir(mypath & "/*.dcr", vbNormal)
Do While myname <> ""
    If myname <> "." And myname <> ".." Then
        If (GetAttr(mypath & myname) And vbNormal) = vbNormal Then
            i = i + 1
            ListView1.ListItems.Add i, , myname, 5, 5
        End If
    End If
    myname = Dir
Loop
myname = Dir(mypath & "/*.dca", vbNormal)
Do While myname <> ""
    If myname <> "." And myname <> ".." Then
        If (GetAttr(mypath & myname) And vbNormal) = vbNormal Then
            i = i + 1
            ListView1.ListItems.Add i, , myname, 6, 6
        End If
    End If
    myname = Dir
Loop
myname = Dir(mypath & "/*.frx", vbNormal)
Do While myname <> ""
    If myname <> "." And myname <> ".." Then
        If (GetAttr(mypath & myname) And vbNormal) = vbNormal Then
            i = i + 1
            ListView1.ListItems.Add i, , myname, 7, 7
        End If
    End If
    myname = Dir
Loop
myname = Dir(mypath & "/*.prj", vbNormal)
Do While myname <> ""
    If myname <> "." And myname <> ".." Then
        If (GetAttr(mypath & myname) And vbNormal) = vbNormal Then
            i = i + 1
            ListView1.ListItems.Add i, , myname, 8, 8
        End If
    End If
    myname = Dir
Loop
myname = Dir(mypath & "/*.bas", vbNormal)
Do While myname <> ""
    If myname <> "." And myname <> ".." Then
        If (GetAttr(mypath & myname) And vbNormal) = vbNormal Then
            i = i + 1
            ListView1.ListItems.Add i, , myname, 9, 9
        End If
    End If
    myname = Dir
Loop
ListView1.Refresh
Screen.MousePointer = 0
End Sub
Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub
Private Sub Form_Load()
ListView1.ListItems.Clear
ListView1.ListItems.Add 1, , "Classes", 10, 10
ListView1.ListItems.Add 2, , "Code", 10, 10
ListView1.ListItems.Add 3, , "Controls", 10, 10
ListView1.ListItems.Add 4, , "Forms", 10, 10
ListView1.ListItems.Add 5, , "MDIForms", 10, 10
ListView1.ListItems.Add 6, , "Menus", 10, 10
ListView1.ListItems.Add 7, , "Projects", 10, 10
ListView1.ListItems.Add 8, , "PropPage", 10, 10
ListView1.ListItems.Add 9, , "UserCtls", 10, 10
ListView1.ListItems.Add 10, , "UserDocs", 10, 10
SelStat = True
Label5.Enabled = False
Label1.Enabled = False
Line1.BorderColor = &HE0E0E0
Line2.BorderColor = &HE0E0E0
Line3.BorderColor = &H808080
Line4.BorderColor = &H808080
Line1.Visible = True
Line2.Visible = True
Line3.Visible = True
Line4.Visible = True
Line5.BorderColor = &H808080
Line6.BorderColor = &H808080
Line7.BorderColor = &HE0E0E0
Line8.BorderColor = &HE0E0E0
Line5.Visible = True
Line6.Visible = True
Line7.Visible = True
Line8.Visible = True
Line9.BorderColor = &H808080
Line10.BorderColor = &H808080
Line11.BorderColor = &HE0E0E0
Line12.BorderColor = &HE0E0E0
Line9.Visible = True
Line10.Visible = True
Line11.Visible = True
Line12.Visible = True
End Sub
Private Sub Form_Resize()
On Error Resume Next
ListView1.Refresh
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line11.BorderColor = &H808080
Line12.BorderColor = &H808080
Line9.BorderColor = &HE0E0E0
Line10.BorderColor = &HE0E0E0
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Kill (mypath + ListView1.SelectedItem.Text)
Call Form_Load
End Sub

Private Sub Label4_Click()
Call Form_Load
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line5.BorderColor = &HE0E0E0
Line6.BorderColor = &HE0E0E0
Line7.BorderColor = &H808080
Line8.BorderColor = &H808080
End Sub
Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line5.BorderColor = &H808080
Line6.BorderColor = &H808080
Line7.BorderColor = &HE0E0E0
Line8.BorderColor = &HE0E0E0
End Sub
Private Sub Label5_Click()
Dim templine As String
CommonDialog1.Filter = "FRM Files (*.frm)|*.frm|DOB Files (*.dob)|*.dob|CLS Files (*.cls)|*.cls|CTL Files (*.ctl)|*.ctl|DCR Files (*.dcr)|*.dcr|DCA Files (*.dca)|*.dca|FRX Files (*.frx)|*.frx|PRJ Files (*.prj)|*.prj|BAS Files (*.bas)|*.bas|All Files (*.*)|*.*"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
    Name CommonDialog1.FileName As mypath + CommonDialog1.FileTitle
Else
    Exit Sub
End If
End Sub
Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line1.BorderColor = &H808080
Line2.BorderColor = &H808080
Line3.BorderColor = &HE0E0E0
Line4.BorderColor = &HE0E0E0
End Sub
Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_Load
End Sub
Private Sub ListView1_Click()
If SelStat = True Then
    mypath = "C:\Program Files\Microsoft Visual Studio\VB98\Template\" + ListView1.SelectedItem.Text + "\"
    Call BuildListView
    Label5.Enabled = True
    Label1.Enabled = True
    SelStat = False
End If
End Sub
Public Sub ClearListBoxes(frmTarget As Form)
    Dim i, j, ctrltarget
    For i = 0 To (frmTarget.Controls.Count - 1)
        Set ctrltarget = frmTarget.Controls(i)
        If TypeOf ctrltarget Is ListView Then
            For j = 0 To ctrltarget.ListCount
                ctrltarget.ListItems.Item(j) = ""
            Next j
        End If
    Next i
End Sub
