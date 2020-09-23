VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1800
   LinkTopic       =   "Form2"
   ScaleHeight     =   1635
   ScaleWidth      =   1800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   30
      TabIndex        =   1
      Top             =   1275
      Width           =   1755
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   1740
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Lil coding done by
'Syed Imran Rais (emnuu@hotmail.com)
'BlackAvenger
'India


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub List1_Click()
If List1.ListIndex < 0 Then Exit Sub
nxt = 1
For i = 1 To nedges

edges(nxt) = edges(i)
If i <> selnodes(List1.ListIndex + 1) Then nxt = nxt + 1

Next i
nedges = nedges - 1
update
Unload Me
End Sub

'Lil coding done by
'Syed Imran Rais (emnuu@hotmail.com)
'BlackAvenger
'India


