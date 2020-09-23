VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   330
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   5865
   Icon            =   "frmToolbox.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   330
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAdd 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   35
      TabIndex        =   1
      Top             =   35
      Width           =   5000
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   0
      Width           =   5295
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go!"
      Height          =   315
      Left            =   5400
      TabIndex        =   0
      Top             =   0
      Width           =   420
   End
   Begin VB.Menu mnuBack 
      Caption         =   "Back"
   End
   Begin VB.Menu mnuForward 
      Caption         =   "Forward"
   End
   Begin VB.Menu mnuRefresh 
      Caption         =   "Refresh"
   End
   Begin VB.Menu mnuStop 
      Caption         =   "Stop"
   End
   Begin VB.Menu mnuHome 
      Caption         =   "Home"
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "Search"
   End
   Begin VB.Menu mnuMove 
      Caption         =   "Move"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Original code by Jason Theobald
'Code additions by Queen City Software
'GUI modifications by Queen City Software




'###############################################
'Code added by Queen City Software
'This part declares the Constants needed to
'keep the toolbox on top of the browser
'###############################################

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hwndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wflags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const flags = SWP_NOMOVE Or SWP_NOSIZE


Private Sub Combo1_Click()
'###############################################
'Code added by Queen City Software
'Let the web browser navigate according to the
'text in the Combo box
Form1.WebBrowser1.Navigate Combo1.Text
'###############################################

End Sub

Private Sub Form_Load()

'#####################################
    'Code added by Queen City Software
    'Put Toolbox above Browser so it won't
    'get lost...

    Call FormOnTop(Me)
'#####################################

    Load Form1
    Form1.Show

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Form1

End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

    txtAdd.Text = WebBrowser1.LocationURL
    'Form1.Caption = WebBrowser1.LocationName

End Sub

Private Sub cmdGo_Click()

    Form1.WebBrowser1.Navigate2 txtAdd.Text
'###############################################
'Code added by Queen City Software
'This adds the typed URL to the Combo box for
'later navigation
    Combo1.AddItem txtAdd.Text
'###############################################

End Sub

Public Sub FormOnTop(frm As Form)
'###############################################
'Code added by Queen City Software
'This is the actual call to make the toolbox
'form stay on top of the web browser
    Call SetWindowPos(frm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, flags)
'###############################################

End Sub

Private Sub mnuBack_Click()

    On Error Resume Next
      Form1.WebBrowser1.GoBack

End Sub

Private Sub mnuForward_Click()

    On Error Resume Next
      Form1.WebBrowser1.GoForward

End Sub

Private Sub mnuHome_Click()

    Form1.WebBrowser1.GoHome

End Sub

Private Sub mnuMove_Click()

    Me.Left = 0
    Me.Top = 0

End Sub

Private Sub mnuRefresh_Click()

    Form1.WebBrowser1.Refresh

End Sub

Private Sub mnuSearch_Click()

    Form1.WebBrowser1.Navigate2 "http://www.altavista.com"

End Sub

Private Sub mnuStop_Click()

    Form1.WebBrowser1.Stop

End Sub

Private Sub txtAdd_KeyPress(KeyAscii As Integer)
'###############################################
'This code added by Queen City Software
'pressing the Enter key makes the web browser
'go to the URL that you just typed
If KeyAscii = vbKeyReturn Then
cmdGo_Click
KeyAscii = 0
End If
'###############################################
End Sub
