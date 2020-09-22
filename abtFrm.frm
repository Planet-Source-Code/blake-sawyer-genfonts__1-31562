VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About GenFonts"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   245
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   1733
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   840
      ItemData        =   "abtFrm.frx":0000
      Left            =   1080
      List            =   "abtFrm.frx":0010
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1500
      Left            =   668
      MousePointer    =   10  'Up Arrow
      Picture         =   "abtFrm.frx":0070
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   3345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "GenFonts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1718
      TabIndex        =   0
      Top             =   120
      Width           =   1245
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   120
      Picture         =   "abtFrm.frx":28E1
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
Form1.Show
Unload Me
End Sub

Private Sub Form_Load()
Unload Form1
End Sub

Private Sub Image2_Click()
ShellExecute Me.hwnd, vbNullString, "http://www.8op.com/yarsoft", vbNullString, "c:\", vbNormalFocus
End Sub

Private Sub List1_DblClick()
Select Case List1.Text
    Case "http://www.8op.com/yarsoft"
        Image2_Click
    Case "TeenageRiot309@att.net"
        ShellExecute Me.hwnd, vbNullString, "mailto:TeenageRiot309@att.net?subject=GenFonts", vbNullString, "c:\", vbNormalFocus
End Select
End Sub
