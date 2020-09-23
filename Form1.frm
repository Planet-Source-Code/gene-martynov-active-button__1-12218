VERSION 5.00
Object = "{DC1B9E15-A466-11D4-90D4-006097935401}#6.0#0"; "ActiveButton.ocx"
Begin VB.Form Form1 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin ActiveButtonPr.ActiveButton ActiveButton1 
      Height          =   390
      Index           =   0
      Left            =   2280
      TabIndex        =   7
      Top             =   1920
      Width           =   390
      _extentx        =   688
      _extenty        =   688
      imagedown       =   "Form1.frx":0000
      imagehot        =   "Form1.frx":0872
      imagedisabled   =   "Form1.frx":10E6
      style           =   1
      imageup         =   "Form1.frx":1958
      backstyle       =   0
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Transparent"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get hWnd"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   360
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Style"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1575
      Begin VB.OptionButton Option1 
         Caption         =   "Check"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Standard"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Enable"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin ActiveButtonPr.ActiveButton ActiveButton1 
      Height          =   390
      Index           =   1
      Left            =   2880
      TabIndex        =   8
      Top             =   1920
      Width           =   390
      _extentx        =   688
      _extenty        =   688
      imagedown       =   "Form1.frx":21CA
      imagehot        =   "Form1.frx":2A3C
      imagedisabled   =   "Form1.frx":32B0
      imageup         =   "Form1.frx":3B22
      backstyle       =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetCapture Lib "user32" () As Long


Private Sub ActiveButton1_Click(Index As Integer)
If Index = 1 Then
    ActiveButton1(0).Value = abUnPressed
End If

End Sub

Private Sub Check1_Click()
If Check1.Value = 0 Then
    ActiveButton1(0).Enabled = False
    ActiveButton1(1).Enabled = False
Else
    ActiveButton1(0).Enabled = True
    ActiveButton1(1).Enabled = True
End If

End Sub

Private Sub Check2_Click()
If Check2.Value = vbChecked Then
    ActiveButton1(0).BackStyle = abTransparent
    ActiveButton1(1).BackStyle = abTransparent
Else
    ActiveButton1(0).BackStyle = abOpaque
    ActiveButton1(1).BackStyle = abOpaque
End If

End Sub

Private Sub Command1_Click()
Text1.Text = ActiveButton1(0).hWnd & "  " & ActiveButton1(1).hWnd

End Sub

Private Sub Form_Load()
If ActiveButton1(0).BackStyle = abOpaque Then
    Check2.Value = 0
Else
    Check2.Value = 1
End If

End Sub

Private Sub Option1_Click(Index As Integer)
ActiveButton1(0).Style = Index
ActiveButton1(1).Style = Index
End Sub

Private Sub UserControl11_Click()
MsgBox "Was clicked"

End Sub
