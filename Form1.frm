VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Check Form 3"
      Height          =   510
      Left            =   3240
      TabIndex        =   1
      Top             =   1665
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check Form 2"
      Height          =   510
      Left            =   540
      TabIndex        =   0
      Top             =   1665
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim m_bLoaded As Boolean

Public Property Get Loaded() As Boolean
Loaded = m_bLoaded
End Property

Public Property Let Loaded(ByVal bLoaded As Boolean)
m_bLoaded = bLoaded
End Property

Private Sub Command1_Click()
If Form2.Loaded Then
    MsgBox "Form 2 is loaded"
Else
    MsgBox "Form 2 is Un loaded"
End If
End Sub
Private Sub Command2_Click()
If Form3.Loaded Then
    MsgBox "Form 3 is loaded"
Else
    MsgBox "Form 3 is Un loaded"
End If
End Sub

Private Sub Form_Load()
Form2.Show
Form3.Show
End Sub
