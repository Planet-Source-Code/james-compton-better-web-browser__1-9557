VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Jim's Browser"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4095
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton OKBUT 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   465
      Left            =   1215
      TabIndex        =   4
      Top             =   1980
      Width           =   1635
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   225
      Picture         =   "frmAbout.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   225
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "http://members.tripod.com/yoda_jammies"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   540
      MouseIcon       =   "frmAbout.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1620
      Width           =   2985
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "A small sample of some of the basic internet functions of the WebBrowser control in VB5. Don't forget to check out my web site at:"
      Height          =   600
      Left            =   225
      TabIndex        =   2
      Top             =   900
      Width           =   3660
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Jim's Browser"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   810
      TabIndex        =   1
      Top             =   225
      Width           =   2760
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label3_Click()
frmBrowser.WebBrowser1.Navigate "http://members.tripod.com/yoda_jammies"
Unload Me
End Sub

Private Sub OKBUT_Click()
Unload Me
End Sub

