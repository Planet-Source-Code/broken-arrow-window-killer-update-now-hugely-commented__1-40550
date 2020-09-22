Attribute VB_Name = "modWinKill"
Option Explicit

'Windows 32 bit API declaration
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String)
Public Declare Function GetPrivateProfileInt& Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String)
Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String)
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Type declaration for the tray icon
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 13
End Type
'Constants for the tray icon
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
'Different window message constants
Public Const WM_CLOSE = &H10
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_RBUTTONDOWN = &H204
Public Const HWND_TOPMOST = -1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

Public INIFile As String 'Varibale to hold the path & INI filename

Public Sub CreateIcon()
'Create the system tray icon
Dim Tic As NOTIFYICONDATA, erg As Long
Tic.cbSize = Len(Tic)
Tic.hwnd = frmMain.picTray.hwnd 'Handle of the object to process the tray events
Tic.uID = 1&
Tic.uFlags = NIF_DOALL
Tic.uCallbackMessage = WM_MOUSEMOVE 'The event to throw the callback event
Tic.hIcon = frmMain.picTray.Picture 'Icon for the tray
Tic.szTip = "Window Killer" 'Tool tip for the tray icon
erg = Shell_NotifyIcon(NIM_ADD, Tic) 'Create the icon
End Sub

Public Sub ModifyIcon()
'Modify the image or the tooltip for the existing tray icon
Dim Tic As NOTIFYICONDATA, erg As Long
Tic.cbSize = Len(Tic)
Tic.hwnd = frmMain.picTray.hwnd 'Handle of the object to process the tray events
Tic.uID = 1&
Tic.uFlags = NIF_DOALL
Tic.uCallbackMessage = WM_MOUSEMOVE 'The event to throw the callback event
Tic.hIcon = frmMain.picTray.Picture 'Icon for the tray
Tic.szTip = "Window Killer" 'Tool tip for the tray icon
erg = Shell_NotifyIcon(NIM_MODIFY, Tic) 'Modify the icon
End Sub

Public Sub DeleteIcon()
'Remove the tray icon
Dim Tic As NOTIFYICONDATA, erg As Long
Tic.cbSize = Len(Tic)
Tic.hwnd = frmMain.picTray.hwnd 'Handle of the object to process the tray events
Tic.uID = 1&
erg = Shell_NotifyIcon(NIM_DELETE, Tic) 'Delete the icon
End Sub

Public Function GetForeWinTxt()
'Get the window title of the topmost window on the desktop
Dim ForeWin As Long, WinTxtLen As Long, WinTxt As String, fWin As Long
ForeWin = GetForegroundWindow 'Get the handle of the foreground window
If ForeWin > 0 Then 'Check if there was a window!
    WinTxtLen = GetWindowTextLength(ForeWin) + 1 'Get the length of the caption of the foreground window
    WinTxt = Space(WinTxtLen) 'Set the buffer
    fWin = GetWindowText(ForeWin, WinTxt, WinTxtLen) 'Get the window title of the foreground window
    GetForeWinTxt = Left(WinTxt, fWin) 'Pad the returned string
End If
End Function

Public Function GetINIKeyStr(strSection As String, strKey As String)
'Get a key string value from the INI file
    Dim strInfo As String
    Dim lngRet As Long
    strInfo = String(260, " ") 'Set the buffer
    GetPrivateProfileString strSection, strKey, "", strInfo, 260, INIFile 'Get the value
    GetINIKeyStr = Trim(strInfo) 'Pad the string
End Function

Public Function GetINIKeyInt(strSection As String, strKey As String)
'Get a key integer value from the INI file
    GetINIKeyInt = GetPrivateProfileInt(strSection, strKey, 0, INIFile) 'Get the value
End Function

Public Sub SetINIKey(strSection As String, strKey As String, strData As String)
'Save a key value in the INI file
    WritePrivateProfileString strSection, strKey, strData, INIFile 'Write the value
End Sub

Public Sub DeleteINISection(strSection As String)
'Delete a section from the INI file
    WritePrivateProfileString strSection, vbNullString, vbNullString, INIFile 'Write the section
End Sub

Public Sub DeleteINIKey(strSection As String, strKey As String)
'Delete a key from the INI file
    WritePrivateProfileString strSection, strKey, vbNullString, INIFile 'Delete the key
End Sub

Public Function ChkPath(strPath As String, Optional AddSlash As Boolean = True)
'Make sure that the path is always formatted with a BACKSLASH (important for root paths)
ChkPath = strPath
If AddSlash And Right(strPath, 1) <> "\" Then ChkPath = strPath & "\"
If Not AddSlash And Right(strPath, 1) = "\" Then ChkPath = Left(strPath, Len(strPath) - 1)
End Function

Public Function ChkWindows(ByVal hwnd As Long, ByVal lpData As Long) As Long
'This is the subroutine that is fired by the EnumWindows API for each visible window on the
'desktop with the window handle as the hwnd parameter

ChkWindows = 1  'Setting this value to 0 will stop processing any other window left on the
                'desktop. Setting the value to 1 tells the EnumWindows API to call this
                'routine again for the next visible window

Dim WindowCaption As String, Ret As Long
WindowCaption = Space(GetWindowTextLength(hwnd) + 1) 'Set the buffer
Ret = GetWindowText(hwnd, WindowCaption, GetWindowTextLength(hwnd) + 1) 'Get the window text of the top most window
WindowCaption = Left(WindowCaption, Ret) 'Pad the string

If ThisIsTheWindow(WindowCaption) Then 'Check if the window caption matches any of the list item
    frmMain.picTray.Picture = frmMain.imgList.ListImages(3).Picture 'Change the tray icon to Kill mode
    ModifyIcon
    
    PostMessage hwnd, WM_CLOSE, 0, 0 'Request the window to be closed
    'Here is an issue! I could have used the SendMessage API instead of PostMessage API
    'I found that SendMessage sometimes do bizzare with the windows and it cannot close
    'the Internet Explorer or the Explorer as these applications requires time to close.
    'So I used this API which simply ques a close request to the window but don't wait for
    'window to close, instead it heads for the next window. Thus I found this API to be
    'more stable than the SendMessage API!
    
    Put #1, , Format(Date, "dd/mm/yyyy- hh:nn:ss") & "- " & WindowCaption & vbCrLf 'Write the log record
    
    frmMain.picTray.Picture = frmMain.imgList.ListImages(2).Picture 'Set the tray icon to scan mode
    ModifyIcon
End If

End Function

Function ThisIsTheWindow(WindowCaption As String) As Boolean
'Check if the passed caption matches to the caption list in the exact matching formula
Dim a As Long, Found As Boolean

ThisIsTheWindow = 0
If WindowCaption = "" Then Exit Function

With frmMain.lstWinTitle.ListItems
    For a = 1 To .Count 'Iterate through the caption list
        If .Item(a).SubItems(3) = "Active" Then 'Check if the item is to be checked
            Select Case .Item(a).SubItems(1) 'Filter the matching formula
            Case "Exactly" 'True if the caption is exactly as found in the list
                If .Item(a).SubItems(2) = "Yes" Then 'Check if the item is case sensitive
                    If .Item(a).Text = WindowCaption Then ThisIsTheWindow = True
                Else
                    If LCase(.Item(a).Text) = LCase(WindowCaption) Then ThisIsTheWindow = True
                End If
            Case "Left" 'True if the caption is found to the left to the item
                If .Item(a).SubItems(2) = "Yes" Then 'Check if the item is case sensitive
                    If Left(WindowCaption, Len(.Item(a).Text)) = .Item(a).Text Then ThisIsTheWindow = True
                Else
                    If LCase(Left(WindowCaption, Len(.Item(a).Text))) = LCase(.Item(a).Text) Then ThisIsTheWindow = True
                End If
            Case "Contained" 'True if the caption is found anywhere in the item
                If .Item(a).SubItems(2) = "Yes" Then 'Check if the item is case sensitive
                    If InStr(WindowCaption, .Item(a).Text) > 0 Then ThisIsTheWindow = True
                Else
                    If InStr(LCase(WindowCaption), LCase(.Item(a).Text)) > 0 Then ThisIsTheWindow = True
                End If
            Case "Right" 'True if the caption is found to the right to the item
                If .Item(a).SubItems(2) = "Yes" Then 'Check if the item is case sensitive
                    If Right(.Item(a).Text, Len(WindowCaption)) = WindowCaption Then ThisIsTheWindow = True
                Else
                    If LCase(Right(WindowCaption, Len(.Item(a).Text))) = LCase(.Item(a).Text) Then ThisIsTheWindow = True
                End If
            End Select
        End If
    Next
End With
End Function

