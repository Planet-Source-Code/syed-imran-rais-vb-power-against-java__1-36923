VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Java Copied AI"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9015
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Shake"
      Height          =   285
      Left            =   7995
      TabIndex        =   13
      Top             =   270
      Width           =   945
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Captions"
      Height          =   225
      Left            =   8010
      TabIndex        =   12
      Top             =   15
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin VB.OptionButton MoveNode 
      Caption         =   " Move Node"
      Height          =   435
      Left            =   1995
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   75
      Width           =   930
   End
   Begin VB.OptionButton RemoveNode 
      Caption         =   "Remove Node"
      Height          =   435
      Left            =   5925
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   75
      Width           =   930
   End
   Begin VB.OptionButton DisconnectNode 
      Caption         =   "Disconnect Node"
      Height          =   435
      Left            =   4935
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   75
      Width           =   930
   End
   Begin VB.OptionButton ConnectNode 
      Caption         =   "Connect Node"
      Height          =   435
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   75
      Width           =   930
   End
   Begin VB.OptionButton NewNode 
      Caption         =   "New Node"
      Height          =   435
      Left            =   2970
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   75
      Width           =   930
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   4260
      Top             =   3300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      Height          =   435
      Left            =   1245
      TabIndex        =   6
      Top             =   75
      Width           =   525
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load"
      Height          =   435
      Left            =   660
      TabIndex        =   5
      Top             =   75
      Width           =   525
   End
   Begin VB.CommandButton Command2 
      Caption         =   "New"
      Height          =   435
      Left            =   75
      TabIndex        =   4
      Top             =   75
      Width           =   525
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   6945
      TabIndex        =   3
      Top             =   -75
      Width           =   75
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   1845
      TabIndex        =   2
      Top             =   -75
      Width           =   75
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   435
      Left            =   7080
      TabIndex        =   1
      Top             =   75
      Width           =   885
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5625
      Top             =   45
   End
   Begin VB.PictureBox Picturebox1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   6480
      Left            =   30
      ScaleHeight     =   400
      ScaleMode       =   0  'User
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   570
      Width           =   8955
   End
   Begin VB.Menu addn 
      Caption         =   "add node"
      Visible         =   0   'False
      Begin VB.Menu nrml 
         Caption         =   "Add new"
      End
      Begin VB.Menu cncl 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Lil coding done by
'Syed Imran Rais (emnuu@hotmail.com)
'BlackAvenger
'India


Dim active As Integer
Dim curx, cury, sfrom


Private Sub Check1_Click()
update
End Sub

Private Sub Command1_Click()
If Command1.Caption = "Go" Then
Timer1.Enabled = True
Command1.Caption = "Stop"

NewNode.Value = 0
ConnectNode.Value = 0
RemoveNode.Value = 0
DisconnectNode.Value = 0

NewNode.Enabled = False
ConnectNode.Enabled = False
RemoveNode.Enabled = False
DisconnectNode.Enabled = False

Else
Timer1.Enabled = False
Command1.Caption = "Go"

NewNode.Enabled = True
ConnectNode.Enabled = True
RemoveNode.Enabled = True
DisconnectNode.Enabled = True


End If
End Sub

Private Sub Command2_Click()
If MsgBox("Are you sure you want to make a new layout.", vbYesNo, "New Confirm") = vbNo Then Exit Sub
If Command1.Caption = "Stop" Then Command1_Click
nnodes = 0
nedges = 0
Picturebox1.Cls

End Sub

Private Sub Command3_Click()
On Error GoTo ed
Dim frm As String, sto As String, l As Integer
If MsgBox("Are you sure you want to load a layout.", vbYesNo, "Load Confirm") = vbNo Then Exit Sub
If Command1.Caption = "Stop" Then Command1_Click
dlg.Filter = "Java AI|*.jai"
dlg.ShowOpen

nedges = 0
nnodes = 0

ff = FreeFile
Open dlg.FileName For Input As ff

Do Until EOF(ff)
Input #ff, frm, sto, l
Call addEdge(frm, sto, l)
Loop
Close #ff

run

ed:
End Sub

Private Sub Command4_Click()
On Error GoTo ed
If Command1.Caption = "Stop" Then Command1_Click
dlg.Filter = "Java AI|*.jai"
dlg.ShowSave
Dim ff
ff = FreeFile
Open dlg.FileName For Output As ff

For i = 1 To nedges
frm = nodes(edges(i).from).lbl
sto = nodes(edges(i).sto).lbl
Write #ff, frm, sto, edges(i).l

Next i

Close #ff
ed:
End Sub


Private Sub Command5_Click()
Dim n As Node
For i = 1 To nnodes
n = nodes(i)
  Randomize
     n.x = 10 + 380 * Rnd
    Randomize
    n.y = 10 + 380 * Rnd
    
    nodes(i) = n
Next i
update
End Sub

Private Sub Form_Load()
init
run
'Timer1.Enabled = True
End Sub

Private Sub nrml_Click()
On Error GoTo ed
Dim ax As String
redo:
ax = InputBox("Enter a name for new node", "Add Node", "Node" & nnodes + 1)

Dim nop As Boolean
For i = 1 To nnodes
    If ax = nodes(i).lbl Then nop = True
Next i

If ax = "" Or ax = Null Or nop Then
    If MsgBox("Node name already exists or invalid node name. Do you wish to assign a new name(Yes) or a default name must be assigned(No).", vbYesNo, "Invalid Node") = vbYes Then GoTo redo
    ax = "Node" & nnodes + 1
End If

Call addNode(ax)
nodes(nnodes).x = curx
nodes(nnodes).y = cury
update

ed:
End Sub

Private Sub Picturebox1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
active = 0
curx = x
cury = y

If NewNode Then

PopupMenu addn
Exit Sub
End If

For i = 1 To nnodes
If x > nodes(i).x - 5 And x < nodes(i).x + 5 And y > nodes(i).y - 5 And y < nodes(i).y + 5 Then active = i: Exit For
Next i

If DisconnectNode Then
Form2.List1.Clear
For i = 1 To nedges
If edges(i).from = active Then Form2.List1.AddItem (nodes(edges(i).sto).lbl): c = 1 + c: selnodes(c) = i
If edges(i).sto = active Then Form2.List1.AddItem (nodes(edges(i).from).lbl): c = 1 + c: selnodes(c) = i
Next i

If c > 0 Then
Dim pa As POINTAPI
Call GetCursorPos(pa)

Form2.Move pa.x * Screen.TwipsPerPixelX, pa.y * Screen.TwipsPerPixelY
Form2.Show 1

active = 0
End If
End If

If RemoveNode Then
For i = 1 To nedges
If edges(i).from = active Then Form2.List1.AddItem (nodes(edges(i).sto).lbl): c = 1: Exit For
If edges(i).sto = active Then Form2.List1.AddItem (nodes(edges(i).from).lbl): c = 1: Exit For
Next i

If c = 1 Then MsgBox "Disconnect the node from other nodes.", vbOKOnly, "Cannot remove": active = 0: Exit Sub

nxt = 1

For i = 1 To nnodes

For j = 1 To nedges
If edges(j).from = i Then edges(j).from = nxt
If edges(j).sto = i Then edges(j).sto = nxt
Next j


nodes(nxt) = nodes(i)
If i <> active Then nxt = nxt + 1

Next i
nnodes = nnodes - 1
update

End If

If ConnectNode Then
sfrom = active
End If


End Sub



Private Sub Picturebox1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If active > 0 And ConnectNode.Value = False Then
nodes(active).x = x
nodes(active).y = y
update
End If
End Sub

Private Sub Picturebox1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
active = 0

Dim l As Integer
If ConnectNode Then
    For i = 1 To nnodes
        If x > nodes(i).x - 5 And x < nodes(i).x + 5 And y > nodes(i).y - 5 And y < nodes(i).y + 5 Then active = i: Exit For
    Next i
    
    If active = 0 Then Exit Sub
    l = Val(Trim(InputBox("Enter for ce between selected nodes (1-100)", "Connect Nodes", "50")))
    If l <= 0 Or l > 100 Then l = 50
    
    Call addEdge(nodes(sfrom).lbl, nodes(active).lbl, l)
    update
End If

active = 0
End Sub

Private Sub Timer1_Timer()
run
End Sub

'Lil coding done by
'Syed Imran Rais (emnuu@hotmail.com)
'BlackAvenger
'India


