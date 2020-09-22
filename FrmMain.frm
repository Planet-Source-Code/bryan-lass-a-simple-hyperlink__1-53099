VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1320
   ClientLeft      =   5220
   ClientTop       =   2235
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   MouseIcon       =   "FrmMain.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   1320
   ScaleWidth      =   4890
   Begin VB.Label LblLink 
      Alignment       =   2  'Center
      Caption         =   "Check It Out A working Hyperlink"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub LblLink_Click()
    ShellExecute hWnd, "open", "http://www.nocashzone.com/index.asp", vbNullString, vbNullString, conSwNormal
    Dim Ie As New InternetExplorer
    Ie.Visible = True
    Ie.Navigate "http://www.idavista.com"
End Sub
Private Sub LblLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LblLink.FontBold = True
    LblLink.FontUnderline = True
    LblLink.ForeColor = vbBlue
    Me.MousePointer = 99
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LblLink.FontBold = False
    LblLink.FontUnderline = False
    LblLink.ForeColor = vbBlack
    Me.MousePointer = 0
End Sub
