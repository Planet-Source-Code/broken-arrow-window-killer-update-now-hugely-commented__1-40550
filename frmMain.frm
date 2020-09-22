VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window Killer"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   360
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   522
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ListView lstWinTitle 
      Height          =   3135
      Left            =   120
      TabIndex        =   26
      Top             =   1200
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Window Title"
         Object.Width           =   6879
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Position"
         Object.Width           =   38100
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Case Sensitive"
         Object.Width           =   38100
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Status"
         Object.Width           =   38100
      EndProperty
   End
   Begin VB.PictureBox picAbout 
      BorderStyle     =   0  'None
      Height          =   2700
      Left            =   1680
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   2700
      ScaleWidth      =   4320
      TabIndex        =   25
      Top             =   1440
      Visible         =   0   'False
      Width           =   4320
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4080
      PasswordChar    =   "*"
      TabIndex        =   21
      Top             =   4485
      Width           =   1935
   End
   Begin VB.CheckBox chkStartMinimized 
      Caption         =   "Start minimized"
      Height          =   195
      Left            =   6360
      TabIndex        =   19
      ToolTipText     =   "Silent start up, resides in system tray without any notification"
      Top             =   4530
      Width           =   1455
   End
   Begin VB.Timer tmrCheck 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   5400
   End
   Begin VB.TextBox txtInterval 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      TabIndex        =   16
      Text            =   "1"
      ToolTipText     =   "Time interval in seconds to search for the windows listed"
      Top             =   4485
      Width           =   495
   End
   Begin VB.VScrollBar scrInterval 
      Height          =   285
      LargeChange     =   100
      Left            =   1800
      Max             =   999
      Min             =   1
      TabIndex        =   15
      Top             =   4485
      Value           =   999
      Width           =   255
   End
   Begin VB.CheckBox chkCase 
      Caption         =   "Case sensitive"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Check it if the search needs to be case sensitive"
      Top             =   750
      Width           =   1335
   End
   Begin VB.CommandButton cmdClearList 
      Caption         =   "&Clear List"
      Height          =   375
      Left            =   1200
      TabIndex        =   13
      ToolTipText     =   "Clear entire list"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdSetStatus 
      Caption         =   "Dis&able"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Disable/Enable entire search function"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.PictureBox picTray 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   120
      Picture         =   "frmMain.frx":262CC
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   5400
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.CommandButton cmdCaptureCurrentWinTitle 
      Caption         =   "..."
      Height          =   315
      Left            =   6000
      TabIndex        =   10
      ToolTipText     =   "Capture current active window title"
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      ToolTipText     =   "Exit Window Killer"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "&Hide"
      Default         =   -1  'True
      Height          =   375
      Left            =   5505
      TabIndex        =   8
      ToolTipText     =   "Send to system tray and work in background"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CheckBox chkStatus 
      Caption         =   "Active"
      Height          =   195
      Left            =   1680
      TabIndex        =   7
      ToolTipText     =   "Disable/Enable this item"
      Top             =   750
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   255
      Left            =   6840
      TabIndex        =   6
      ToolTipText     =   "Remove the current item"
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   255
      Left            =   6000
      TabIndex        =   5
      ToolTipText     =   "Update current search item"
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "&Add"
      Height          =   255
      Left            =   5160
      TabIndex        =   4
      ToolTipText     =   "Add this item in search list"
      Top             =   720
      Width           =   855
   End
   Begin VB.ComboBox cboPosition 
      Height          =   315
      ItemData        =   "frmMain.frx":265D6
      Left            =   6480
      List            =   "frmMain.frx":265E6
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Positional occurance of the search string in window titles"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtWinTitle 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "The window title/caption to search for"
      Top             =   360
      Width           =   5895
   End
   Begin ComctlLib.ImageList imgList 
      Left            =   960
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2660B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":26925
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":26C3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":26F59
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblJoySoftwares 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Joy Softwares"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   2775
      MouseIcon       =   "frmMain.frx":27273
      MousePointer    =   99  'Custom
      TabIndex        =   24
      ToolTipText     =   " http://www.joysoftwares.150m.com "
      Top             =   4935
      Width           =   2235
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Joy Softwares"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2790
      TabIndex        =   23
      Top             =   4950
      Width           =   2235
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Joy Softwares"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2760
      TabIndex        =   22
      Top             =   4920
      Width           =   2235
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   3240
      TabIndex        =   20
      Top             =   4530
      Width           =   690
   End
   Begin VB.Label Label4 
      Caption         =   "second(s)"
      Height          =   195
      Left            =   2160
      TabIndex        =   18
      Top             =   4530
      Width           =   705
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Check on every"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   4530
      Width           =   1125
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   8
      X2              =   512
      Y1              =   73
      Y2              =   73
   End
   Begin VB.Line Line1 
      X1              =   8
      X2              =   512
      Y1              =   72
      Y2              =   72
   End
   Begin VB.Label Label2 
      Caption         =   "Position"
      Height          =   195
      Left            =   6480
      TabIndex        =   1
      Top             =   120
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Window Title"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   915
   End
   Begin VB.Menu mnuSysTray 
      Caption         =   "SysTray"
      Visible         =   0   'False
      Begin VB.Menu mnuSysTrayShow 
         Caption         =   "&Show"
      End
      Begin VB.Menu mnuSysTrayDisable 
         Caption         =   "Dis&able"
      End
      Begin VB.Menu mnuSysTrayExit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu MBAR1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSysTrayAbout 
         Caption         =   "A&bout"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddNew_Click()
'Add the new record to the caption list
Dim Found As Boolean, a As Long
For a = 1 To lstWinTitle.ListItems.Count
    If lstWinTitle.ListItems(a).Text = txtWinTitle.Text Then Found = True
    If Found Then Exit For
Next
If Found Then Exit Sub

With lstWinTitle.ListItems
    .Add , , txtWinTitle
    .Item(.Count).SubItems(1) = cboPosition
    If chkCase Then .Item(.Count).SubItems(2) = "Yes" Else .Item(.Count).SubItems(2) = "No"
    If chkStatus Then .Item(.Count).SubItems(3) = "Active" Else .Item(.Count).SubItems(3) = "Inactive"
End With
End Sub

Private Sub cmdCaptureCurrentWinTitle_Click()
'Get the window title of the currently visible top most window under the WinKill window
Me.Hide 'Hide WinKill before getting the title, otherwise the WinKill's title will be returned!
DoEvents 'Let the operating system do it's job
txtWinTitle = GetForeWinTxt 'Get the window title of the top most window
Me.Show 'Show WinKill back
End Sub

Private Sub cmdClearList_Click()
'Clear up the current caption list
If MsgBox("Are you sure that you want to clear all search items from the list ?", vbQuestion + vbYesNo, "Reset List") = vbYes Then lstWinTitle.ListItems.Clear
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdHide_Click()
Me.Hide
If cmdSetStatus.Caption = "Dis&able" Then picTray.Picture = imgList.ListImages(1).Picture Else picTray.Picture = imgList.ListImages(4).Picture
CreateIcon
tmrCheck.Interval = 1000 * Val(txtInterval)
If cmdSetStatus.Caption = "Dis&able" Then tmrCheck.Enabled = True
mnuSysTrayDisable.Caption = cmdSetStatus.Caption
End Sub

Private Sub cmdRemove_Click()
'Remove the currently selected item from the list
lstWinTitle.ListItems.Remove (lstWinTitle.SelectedItem.Index)
End Sub

Private Sub cmdSetStatus_Click()
'Enable/disable WinKill
If cmdSetStatus.Caption = "Dis&able" Then cmdSetStatus.Caption = "En&able" Else cmdSetStatus.Caption = "Dis&able"
End Sub

Private Sub cmdUpdate_Click()
'Update current item in the list
cmdRemove_Click
cmdAddNew_Click
End Sub

Private Sub Form_Load()
If App.PrevInstance Then 'Prevent multiple instance of the application
    MsgBox "Window Killer is already running on this system! Clik on the system tray icon to open Window Killer.", vbInformation + vbOKOnly, "Window Killer running!"
    End
End If
'Set WinKill to the top of the other windows
SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / 15, Me.Top / 15, Me.Width / 15, Me.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW

INIFile = ChkPath(App.Path, True) & App.Title & ".INI" 'INI file to use
cboPosition.ListIndex = 0
LoadWindowList 'Load the caption list from the INI file
txtInterval.Text = GetINIKeyInt("Settings", "Interval") 'Interval to cycle for the window search
If Val(txtInterval) < 1 Then txtInterval = 1
scrInterval.Value = 999 - Val(txtInterval) + 1

'Check if the application is to be started minimized in the system tray
If GetINIKeyInt("Settings", "Start MiniMized") = vbChecked Then chkStartMinimized.Value = vbChecked Else chkStartMinimized.Value = vbUnchecked
If chkStartMinimized.Value = vbChecked Then cmdHide_Click

txtPassword = GetINIKeyStr("Settings", "Password") 'Load the password to restore the application

App.TaskVisible = False 'Don't show the application in the Task Manager's application list nor in the Task bar

'Open a log file for the history
Open ChkPath(App.Path, True) & "Log.txt" For Binary As #1
If LOF(1) > 1.38 * 1024 * 1024 Then 'Reset the log file if bigger
    Close #1
    Kill ChkPath(App.Path, True) & "Log.txt"
    Open ChkPath(App.Path, True) & "Log.txt" For Binary As #1
End If
Seek #1, LOF(1) + 1 'Set the file pointer to the very last byte for a new record
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Don't let the user close the application from the control box, hide instead
If UnloadMode = vbFormControlMenu Then
    Cancel = 1
    cmdHide_Click 'Hide the application
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveWindowList 'Save the current caption list to the INI file
SetINIKey "Settings", "Interval", txtInterval 'Save the interval to the INI file
SetINIKey "Settings", "Start Minimized", chkStartMinimized 'Save the Start Minimized value to the INI file
SetINIKey "Settings", "Password", txtPassword 'Save the password to the INI file

If Not Me.Visible Then DeleteIcon 'Remove the system tray icon
Close #1 'Close the INI file
End Sub

Private Sub lblJoySoftwares_Click()
On Error Resume Next
Shell "Explorer http://www.JoySoftwares.150m.com"
Shell "Explorer MailTo:SKabir_BGD@HotMail.com"
End Sub

Private Sub lstWinTitle_Click()
If lstWinTitle.ListItems.Count < 1 Then Exit Sub
txtWinTitle = lstWinTitle.SelectedItem.Text
cboPosition = lstWinTitle.SelectedItem.SubItems(1)
If lstWinTitle.SelectedItem.SubItems(2) = "Yes" Then chkCase.Value = vbChecked Else chkCase.Value = vbUnchecked
If lstWinTitle.SelectedItem.SubItems(3) = "Active" Then chkStatus.Value = vbChecked Else chkStatus.Value = vbUnchecked
End Sub

Private Sub mnuSysTrayAbout_Click()
mnuSysTrayShow_Click
picAbout.Visible = True
End Sub

Private Sub mnuSysTrayDisable_Click()
If InputBox("Please enter password (use blank if there is no password):", "Password") = txtPassword Then
    picTray_MouseMove 1, 0, Screen.TwipsPerPixelX * WM_LBUTTONDOWN, 0
    cmdSetStatus_Click
    cmdHide_Click
Else
    MsgBox "Sorry, password didn't match, you are not allowed to configure Window Killer.", vbCritical + vbOKOnly, "Password mismatch!"
End If
End Sub

Private Sub mnuSysTrayExit_Click()
If InputBox("Please enter password (use blank if there is no password):", "Password") = txtPassword Then
    cmdExit_Click
Else
    MsgBox "Sorry, password didn't match, you are not allowed to configure Window Killer.", vbCritical + vbOKOnly, "Password mismatch!"
End If
End Sub

Private Sub mnuSysTrayShow_Click()
picTray_MouseMove 1, 0, Screen.TwipsPerPixelX * WM_LBUTTONDOWN, 0
End Sub

Private Sub picAbout_Click()
picAbout.Visible = False
End Sub

Private Sub picTray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Handle the mouse activity on the system tray icon

If Me.Visible Then Exit Sub
If X / Screen.TwipsPerPixelX = WM_LBUTTONDOWN Then
    If InputBox("Please enter password (use blank if there is no password):", "Password") = txtPassword Then
        tmrCheck.Enabled = False
        txtWinTitle = GetForeWinTxt
        Me.Show
        Me.SetFocus
        DeleteIcon
    Else
        MsgBox "Sorry, password didn't match, you are not allowed to configure Window Killer.", vbCritical + vbOKOnly, "Password mismatch!"
    End If
End If
If X / Screen.TwipsPerPixelX = WM_RBUTTONDOWN Then PopupMenu mnuSysTray
End Sub

Sub SaveWindowList()
'Save the current caption list to the INI file
Dim a As Long

SetINIKey "Window List", "Total Windows", lstWinTitle.ListItems.Count 'Total captions
For a = 1 To lstWinTitle.ListItems.Count 'Each loop is a record
    SetINIKey "Window List", "Title    " & Format(a, "########"), lstWinTitle.ListItems.Item(a).Text
    SetINIKey "Window List", "Position " & Format(a, "########"), lstWinTitle.ListItems.Item(a).SubItems(1)
    SetINIKey "Window List", "Case     " & Format(a, "########"), lstWinTitle.ListItems.Item(a).SubItems(2)
    SetINIKey "Window List", "Active   " & Format(a, "########"), lstWinTitle.ListItems.Item(a).SubItems(3)
Next
End Sub

Sub LoadWindowList()
'Load the caption list from the INI file to the list view
Dim a As Long

For a = 1 To GetINIKeyInt("Window List", "Total Windows") 'Iterate through the INI file for the total number of records, each loop is a record
    With lstWinTitle.ListItems
        .Add , , GetINIKeyStr("Window List", "Title    " & Format(a, "########"))
        .Item(.Count).SubItems(1) = GetINIKeyStr("Window List", "Position " & Format(a, "########"))
        .Item(.Count).SubItems(2) = GetINIKeyStr("Window List", "Case     " & Format(a, "########"))
        .Item(.Count).SubItems(3) = GetINIKeyStr("Window List", "Active   " & Format(a, "########"))
    End With
Next
End Sub

Private Sub scrInterval_Change()
txtInterval = 999 - scrInterval.Value + 1
End Sub

Private Sub tmrCheck_Timer()
'This is the timer event that fires at every interval period specified to iterate through
'the visible windows on the desktop and take required action

If lstWinTitle.ListItems.Count < 1 Then Exit Sub

picTray.Picture = imgList.ListImages(2).Picture 'Change the tray icon to busy
ModifyIcon

tmrCheck.Enabled = False 'Don't let the same event to be fired until the pending task is complete

Dim hwnd As Long
'Call the EnumWindows API to itarate through the all visible windows on the desktop and
'call the ChkWindows subroutine with the window handle
Call EnumWindows(AddressOf ChkWindows, hwnd)

tmrCheck.Enabled = True 'Reenable the timer for the next trigger
picTray.Picture = imgList.ListImages(1).Picture 'Set the tray icon to normal
ModifyIcon
End Sub


Private Sub txtPassword_LostFocus()
If txtPassword <> InputBox("Please confirm the password:", "Password") Then
    MsgBox "Password didn't match, not password is active. Please retype password and press Hide to activate password.", vbCritical, "Password not active!"
    txtPassword = ""
End If
End Sub
