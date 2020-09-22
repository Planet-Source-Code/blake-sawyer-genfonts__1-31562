VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "GenFonts"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "genFonts.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   112
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2160
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "GenFonts"
      Filter          =   "HTML Files(*.htm,*.html)|*.htm;*.html"
      InitDir         =   "C:\"
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   495
      Left            =   1736
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3191
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
      Height          =   495
      Left            =   274
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Waiting..."
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   4680
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   135
      Top             =   495
      Width           =   15
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   120
      Top             =   480
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FontCount As Integer
Dim Counter, A As Integer

Private Sub Command1_Click()
CommonDialog1.ShowSave
Counter = 0
Me.Refresh
If CommonDialog1.FileName <> "" Then
Open CommonDialog1.FileName For Output As #1
    For A = FontCount To 0 Step -1
        Print #1, "<table border=1 cellspacing=1 cellpadding=1><tr><td>" + Printer.Fonts(A) + "</td></tr>"
        Print #1, "<tr><td><font face=""" + Printer.Fonts(A) + """ size=24>ABCDEFGHIJKLMNOPQRSTUVWXYZ</font></td></tr>"
        Print #1, "<tr><td><font face=""" + Printer.Fonts(A) + """ size=24>abcdefghijklmnopqtstuvwxyz</font></td></tr>"
        Print #1, "<tr><td><font face=""" + Printer.Fonts(A) + """ size=24>1234567890</font></td></tr>"
        Print #1, "</table><br>"
        Counter = Counter + 1
        Shape2.Width = 286 * (Counter / FontCount)
        Label1.Caption = Str(Int(100 * (Counter / FontCount))) + "%"
        DoEvents
    Next
Close #1
Label1.Caption = "Done."
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Form2.Show
End Sub

Private Sub Form_Load()
Counter = 0
A = 0
FontCount = 0
FontCount = Printer.FontCount - 1
End Sub
