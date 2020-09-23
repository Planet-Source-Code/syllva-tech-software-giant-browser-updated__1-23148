VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Giant Browser"
   ClientHeight    =   4605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6090
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4605
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4320
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5805
      ExtentX         =   10239
      ExtentY         =   7620
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Original code and GUI by Jason Theobald
'Code additions by Queen City Software
'GUI modifications by Queen City Software




Private Sub Form_Load()

    WebBrowser1.GoHome

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Form2

End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)

    Form2.txtAdd.Text = WebBrowser1.LocationURL

End Sub

Private Sub Form_Resize()

    On Error Resume Next
'###############################################
'Code improved by Queen City Software
'This puts in one line the resizing of the
'web browser to cover the entire form/screen
      WebBrowser1.Move 0, 0, ScaleWidth, ScaleHeight
'###############################################

End Sub
