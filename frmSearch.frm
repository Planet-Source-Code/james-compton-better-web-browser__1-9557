VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Jim's Search Engine"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   660
   ClientWidth     =   6975
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option1 
      Caption         =   "Tucows"
      Height          =   255
      Index           =   13
      Left            =   5280
      TabIndex        =   23
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Frame Frame7 
      Caption         =   "File Searches"
      Height          =   615
      Left            =   5160
      TabIndex        =   29
      Top             =   1080
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Astalavista"
      Height          =   255
      Index           =   14
      Left            =   3720
      TabIndex        =   27
      Top             =   2205
      Width           =   1455
   End
   Begin VB.Frame Frame5 
      Caption         =   "Crack Searches"
      Height          =   975
      Left            =   3600
      TabIndex        =   28
      Top             =   1800
      Width           =   3255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "AudioGalaxy"
      Height          =   255
      Index           =   16
      Left            =   240
      TabIndex        =   26
      Top             =   2400
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Audiofind"
      Height          =   255
      Index           =   15
      Left            =   240
      TabIndex        =   25
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Frame Frame6 
      Caption         =   "MP3 Searches"
      Height          =   975
      Left            =   120
      TabIndex        =   24
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Frame Frame4 
      Caption         =   "Search Engine"
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   6735
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1658
         TabIndex        =   18
         Top             =   285
         Width           =   3015
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4778
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Coded String:"
         Height          =   195
         Left            =   600
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Escaped 
         Height          =   255
         Left            =   1680
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Search &For:"
         Height          =   195
         Left            =   600
         TabIndex        =   17
         Top             =   300
         Width           =   825
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "SavySearch"
      Height          =   255
      Index           =   12
      Left            =   5280
      TabIndex        =   15
      Top             =   600
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "MetaCrawler"
      Height          =   255
      Index           =   11
      Left            =   5280
      TabIndex        =   14
      Top             =   240
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Excite"
      Height          =   255
      Index           =   6
      Left            =   2040
      TabIndex        =   9
      Top             =   1320
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Lycos"
      Height          =   255
      Index           =   5
      Left            =   2040
      TabIndex        =   8
      Top             =   960
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Open Text"
      Height          =   255
      Index           =   4
      Left            =   2040
      TabIndex        =   7
      Top             =   600
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Alta Vista"
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Inktomi"
      Height          =   255
      Index           =   10
      Left            =   3480
      TabIndex        =   13
      Top             =   1320
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "DejaNews"
      Height          =   255
      Index           =   9
      Left            =   3480
      TabIndex        =   12
      Top             =   960
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "InfoSeek"
      Height          =   255
      Index           =   8
      Left            =   3480
      TabIndex        =   11
      Top             =   600
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Webcrawler"
      Height          =   255
      Index           =   7
      Left            =   3480
      TabIndex        =   10
      Top             =   240
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Point"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Magellan"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Yahoo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Meta Searches"
      Height          =   975
      Left            =   5160
      TabIndex        =   2
      Top             =   0
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search Engines"
      Height          =   1695
      Left            =   1920
      TabIndex        =   1
      Top             =   0
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Directories"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   2640
      TabIndex        =   22
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSearch 
         Caption         =   "&Search"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu Space1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Api Functions Declaration
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Variable that contains the seleted Search Engine
Private Selected As Integer

Private Sub cmdSearch_Click()
    Dim S As String
'Generate the search command string for the selected
'Search Engine
    Select Case Selected
    Case 0
        S = "http://search.yahoo.com/bin/search?p=" & Escaped.Caption
    Case 1
        S = "http://www.mckinley.com/searcher.cgi?query=" & Escaped.Caption
    Case 2
        S = "http://www.lycos.co.uk/cgi-bin/pursuit?matchmode=and&mtemp=main.sites&etemp=error&query=" & Escaped.Caption & "&cat=brit&x=44&y=4"
    Case 3
        S = "http://www.altavista.digital.com/cgi-bin/query?pg=q&what=web&q=" & Escaped.Caption & "&mode=and"
    Case 4
        S = "http://search.opentext.com/omw/simplesearch?SearchFor=" & Escaped.Caption & "&mode=and"
    Case 5
        S = "http://www.lycos.com/cgi-bin/pursuit?query=" & Escaped.Caption + "&backlink=217&maxhits=25"
    Case 6
        S = "http://www.excite.com/search.gw?searchType=Concept&search=" & Escaped.Caption & "&category=default&mode=relevance&showqbe=1&display=html3,hb"
    Case 7
        S = "http://www.webcrawler.com/cgi-bin/WebQuery?searchText=" & Escaped.Caption & "&maxHits=25"
    Case 8
        S = "http://guide-p.infoseek.com/Titles?qt=" & Escaped.Caption & "&col=WW"
    Case 9
        S = "http://search.dejanews.com/nph-dnquery.xp?query=" & Escaped.Caption & "&defaultOp=AND&svcclass=dncurrent&maxhits=25"
    Case 10
        S = "http://204.161.74.8:1234/query/?query=" & Escaped.Caption & "&hits=25&disp=Text+Only"
    Case 11
        S = "http://search.go2net.com/crawler?general=" & Escaped.Caption & "&method=0&target=®ion=0&rpp=20&timeout=5&hpe=10"
    Case 12
        S = "http://guaraldi.cs.colostate.edu:2000/search?KW=" & Escaped.Caption & "&classic=on&t1=x&Boolean=AND&Hits=10&Mode=MakePlan&df=normal&AutoStep=on&AutoInt=on&lb=1"
    Case 13
        S = "http://www.tucows.com/perl/tucowsSearch?word=" & Escaped.Caption & "&key=all&platform=win95"
    Case 14
        S = "http://astalavista1.box.sk/cgi-bin/robot?srch=" & Escaped.Caption & "&submit=+search+&project=robot&gfx=robot"
    Case 15
        S = "http://www.audiofind.com:70/?audiofindsize=0&audiofindsearch=" & Escaped.Caption
    Case 16
        S = "http://www.audiogalaxy.com/search.php3?MP3Name=" & Escaped.Caption
        
    End Select
'Open the default Web Browser window with the selected
'location
frmBrowser.WebBrowser1.Navigate S
Me.Hide
'    ShellExecute Me.hwnd, "open", S, "", "", 1
End Sub

Private Sub cmdClose_Click()
'Close the program
Me.Hide
End Sub

Private Sub mnuAbout_Click()
'Generate a standard About Message Box
    MsgBox "Programmed by Pedro Lamas" & vbCrLf & "Modified by FLIBLO" & vbCrLf & vbCrLf & "Original version can be had at: www.terravista.pt/portosanto/3723/" & vbCrLf & "My (FLIBLO'S) site: http://members.tripod.com/yoda_jammies", vbApplicationModal + vbInformation, "Credits!"
End Sub

Private Sub mnuExit_Click()
'End the program
    End
End Sub

Private Sub mnuSearch_Click()
'Start searching
    cmdSearch_Click
End Sub

Private Sub Option1_Click(Index As Integer)
'Check if user selected a diferent Search Engine
    If Selected <> Index Then
'Set the selected option button FontBold property to True
'and the old one to False
        Option1(Selected).FontBold = False
        Option1(Index).FontBold = True
'Update the selected engine variable
        Selected = Index
    End If
End Sub

Private Sub Text1_Change()
'Declare the required local variables
    Dim I As Integer, Buffer As String, CBuffer As String
'Get the Text1 TextBox text
    Buffer = Text1.Text
'Check if it is empty
    If Buffer = "" Then
'If so, disable the Search CommandButton
        cmdSearch.Enabled = False
    Else
'If not, enable the Search CommandButton
        cmdSearch.Enabled = True
    End If
'Do for each letter of the Search String
    For I = 1 To Len(Buffer)
'Check the letters ASCII value
        Select Case Asc(Mid(Buffer, I, 1))
'Letters with no special encoding required, stay the same
        Case 42, 43, 45 To 57, 64 To 90, 95, 97 To 122
            CBuffer = CBuffer + Mid(Buffer, I, 1)
'Letters with special encoding required, are now coded
        Case Else
            CBuffer = CBuffer + "%" & Hex(Asc(Mid(Buffer, I, 1)))
        End Select
    Next I
'Show the encoded string on the Escaped Label
    Escaped.Caption = CBuffer
End Sub
