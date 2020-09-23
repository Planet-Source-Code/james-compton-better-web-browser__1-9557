VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmBrowser 
   Caption         =   "Jim's Browser"
   ClientHeight    =   7395
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   11235
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3780
      Top             =   -90
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.PictureBox Load_Bar 
      Height          =   240
      Left            =   9090
      ScaleHeight     =   180
      ScaleWidth      =   2160
      TabIndex        =   10
      Top             =   7110
      Width           =   2220
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   615
      Left            =   10215
      Picture         =   "Main.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   1020
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   615
      Left            =   9180
      Picture         =   "Main.frx":0723
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   1020
   End
   Begin VB.ComboBox txtAddress 
      Height          =   315
      Left            =   0
      TabIndex        =   7
      Text            =   "http://"
      Top             =   270
      Width           =   4290
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   60
      Left            =   0
      TabIndex        =   6
      Top             =   630
      Width           =   11220
      ExtentX         =   19791
      ExtentY         =   106
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   615
      Left            =   8190
      Picture         =   "Main.frx":0AFA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   615
      Left            =   7155
      Picture         =   "Main.frx":0EB2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1020
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   "Forward"
      Height          =   615
      Left            =   6165
      Picture         =   "Main.frx":127B
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   615
      Left            =   5175
      Picture         =   "Main.frx":1604
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "GO!"
      Default         =   -1  'True
      Height          =   615
      Left            =   4365
      Picture         =   "Main.frx":198B
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   795
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jim's Browser - Blocking Pop-Up windows"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   7110
      Width           =   9075
   End
   Begin VB.Menu nmuFile 
      Caption         =   "&File"
      Begin VB.Menu menNEW 
         Caption         =   "&New instance"
         Shortcut        =   {F9}
      End
      Begin VB.Menu se 
         Caption         =   "-"
      End
      Begin VB.Menu menOffline 
         Caption         =   "&Work Offline"
         Shortcut        =   ^W
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsAllow 
         Caption         =   "Allow Pop-Up windows"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu menUpdate 
         Caption         =   "Update history"
         Checked         =   -1  'True
      End
      Begin VB.Menu menClear 
         Caption         =   "Clear history"
      End
   End
   Begin VB.Menu mentool 
      Caption         =   "&Tools"
      Begin VB.Menu mnuPage 
         Caption         =   "Get current address"
         Shortcut        =   {F12}
      End
      Begin VB.Menu menIP 
         Caption         =   "IP Address"
         Shortcut        =   ^I
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu menHistory 
         Caption         =   "View/Edit history"
      End
      Begin VB.Menu menHISTREFRESH 
         Caption         =   "Refresh history"
      End
      Begin VB.Menu sepp 
         Caption         =   "-"
      End
      Begin VB.Menu menSearch 
         Caption         =   "Search..."
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu menHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public AllowPopup As Boolean 'For the Pop-up windows
Sub Load_History()
On Error GoTo error
Open ProgPath & "History.dat" For Input As #1
txtAddress.Clear
Do Until EOF(1)
    Line Input #1, TempString
    txtAddress.AddItem TempString
Loop
Close #1
Exit Sub

error:
MsgBox "Error loading history folder. Please make sure it exists (even if it is blank)", vbCritical
End Sub
Sub Save_History()
On Error GoTo error
'SAVE IT!
Dim TempString As String
Open ProgPath & "History.dat" For Output As #1
For x = 0 To txtAddress.ListCount - 1
    txtAddress.ListIndex = x
    Print #1, txtAddress.Text
Next
Close #1
Exit Sub
error:
MsgBox "An error occured while saving the history file", vbCritical
End Sub
Private Function FindItem(xItem As String, Object As Object) As Boolean
'This function checks whether the page is already in the combo box
For x = 0 To Object.ListCount - 1
    If UCase(xItem) = UCase(Object.List(x)) Then
    FindItem = True
    Exit Function
    End If
Next x
    FindItem = False
End Function
Private Function ProgPath()
' If dragged file is in the root, append filename.
If Mid(App.Path, Len(App.Path)) = "\" Then
ProgPath = App.Path
' If dragged file is not in root, append "\" and filename.
Else
ProgPath = App.Path & "\"
End If
End Function

Private Sub cmdBack_Click()
On Error GoTo error
'Go back one page
WebBrowser1.GoBack
Exit Sub
error:
MsgBox "Can't go back any further!", vbCritical
End Sub

Private Sub cmdForward_Click()
On Error GoTo error
'go forward one page
WebBrowser1.GoForward
Exit Sub

error:
MsgBox "Can't go forward any further!", vbCritical
End Sub

Private Sub cmdGo_Click()
'Go to web page
WebBrowser1.Navigate txtAddress.Text
lblStatus.Caption = "Connecting to: " & txtAddress.Text

If menUpdate.Checked = True Then    'Update history?
    'Add item to combo box
    If FindItem(txtAddress.Text, txtAddress) = False Then
    'If FindItem = False Then
    txtAddress.AddItem (txtAddress.Text)
    Save_History
    Else
    End If
Else
End If
End Sub

Private Sub cmdHome_Click()
WebBrowser1.GoHome
End Sub

Private Sub cmdRefresh_Click()
'Refresh page
WebBrowser1.Refresh
End Sub

Private Sub cmdSearch_Click()
Form1.Show 1
End Sub

Private Sub cmdStop_Click()
'Stop loading
WebBrowser1.Stop
End Sub
Private Sub Form_Load()
txtAddress.Left = 50
cmdGo.Left = (txtAddress.Left + txtAddress.Width) + 20
'Load history file
Load_History
txtAddress.Text = "http://"

'Go home
WebBrowser1.GoHome
End Sub

Private Sub Form_Resize()
On Error Resume Next
'Resizes everything to fit to the form
With WebBrowser1
    .Width = Me.Width - 200
    .Left = 50
    .Height = Me.Height - 1620
End With
With lblStatus
    .Top = WebBrowser1.Top + WebBrowser1.Height + 20
    .Left = WebBrowser1.Left
    .Width = WebBrowser1.Width - Load_Bar.Width
End With
With Load_Bar
    .Top = WebBrowser1.Top + WebBrowser1.Height + 20
    .Left = lblStatus.Left + lblStatus.Width
    .Height = lblStatus.Height
End With
End Sub



Private Sub menClear_Click()
message = MsgBox("Are you sure you want to clear the history folder?", 68, "Clear history?")
If message = vbYes Then
txtAddress.Clear
Save_History
Else
End If
End Sub

Private Sub menHistory_Click()
On Error GoTo NT ' If notepad wasn't found in C:\Windows, the user
                 ' could have NT
x = Shell("C:\Windows\Notepad.exe " & ProgPath & "History.dat", vbMaximizedFocus)
Exit Sub

NT:
x = Shell("C:\WinNT\Notepad.exe " & ProgPath & "History.dat", vbMaximizedFocus)
End Sub

Private Sub menHISTREFRESH_Click()
Dim TEMP As String

TEMP = txtAddress.Text      'Don't want to lose what's in this box upon refreshing....so save it in a tempory string
Load_History
txtAddress.Text = TEMP      'Put it back
End Sub

Private Sub menIP_Click()
MsgBox "Your current IP address is " & Winsock1.LocalIP, vbInformation, "IP Check"
Clipboard.Clear
Clipboard.SetText Winsock1.LocalIP
End Sub

Private Sub menNEW_Click()
On Error GoTo error
x = Shell(ProgPath & App.EXEName, vbNormalFocus)
Exit Sub

error:
MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical, "Error"
End Sub

Private Sub menOffline_Click()
If menOffline.Checked = False Then
WebBrowser1.Offline = True
menOffline.Checked = True
Else
WebBrowser1.Offline = False
menOffline.Checked = False
End If
End Sub

Private Sub menSearch_Click()
Form1.Show 1
End Sub

Private Sub menUpdate_Click()
If menUpdate.Checked = False Then
menUpdate.Checked = True
Else
menUpdate.Checked = False
End If
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show 1
End Sub

Private Sub mnuExit_Click()
'Exit program
Unload Me
End Sub

Private Sub mnuOptionsAllow_Click()
'Turn on/off pop-up windows
If AllowPopup = True Then
    AllowPopup = False
    mnuOptionsAllow.Checked = False
    lblStatus.Caption = "Jim's Browser - Blocking Pop-Up windows"
ElseIf AllowPopup = False Then
    AllowPopup = True
    mnuOptionsAllow.Checked = True
    lblStatus.Caption = "Jim's Browser - Allowing Pop-Up windows"
End If
End Sub

Private Sub mnuPage_Click()
If WebBrowser1.LocationURL <> "" Then
MsgBox "Your current location is " & WebBrowser1.LocationURL, vbInformation, "Location Url"
Clipboard.Clear
Clipboard.SetText WebBrowser1.LocationURL
Else
MsgBox "You are not currently at a site.", vbInformation, "Location Url"
End If
End Sub

Private Sub txtAddress_Click()
cmdGo_Click
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
'shows done in the status bar
lblStatus.Caption = "Loading Complete"
Me.Caption = "Jim's Browser - " & WebBrowser1.LocationName
End Sub

Private Sub WebBrowser1_DownloadBegin()
'Starting download
lblStatus.Caption = "Starting Download"
End Sub

Private Sub WebBrowser1_DownloadComplete()
'Done downloading
lblStatus.Caption = "Download Complete"
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
'Loaded page
lblStatus.Caption = "Loading Complete"
Me.Caption = "Jim's Browser - " & WebBrowser1.LocationName  'Shows webpage in title bar
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
'This will allow a pop-up window to load or to be blocked!
If AllowPopup = True Then
    Cancel = False
    DoEvents
ElseIf AllowPopup = False Then
    Cancel = True
End If
End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
'Shows progress in status bar
lblStatus.Caption = "Reading " & Progress & " of " & ProgressMax
On Error Resume Next
If Progress > 0 Then Call PercentBar(Load_Bar, Progress, ProgressMax)
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
'shows new text in status bar
lblStatus.Caption = Text
End Sub

Function FileExist(vFile As String) As Boolean
    On Error Resume Next
    FileExist = False
    If Dir$(vFile) <> "" Then: FileExist = True
End Function

