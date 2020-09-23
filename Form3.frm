VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form3"
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

Private Sub Form_Load()
Loaded = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
Loaded = False
End Sub
