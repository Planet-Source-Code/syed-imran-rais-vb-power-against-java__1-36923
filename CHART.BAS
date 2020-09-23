Attribute VB_Name = "modAI"
'Lil coding done by
'Syed Imran Rais (emnuu@hotmail.com)
'BlackAvenger
'India


Type Node
 x As Double
 y As Double
 dx As Double
 dy As Double
 fixed As Boolean
 lbl As String
End Type


Type Edge
 from As Integer
 sto As Integer
 l As Integer
End Type

   Global nodes(100) As Node
   Global nnodes, nedges As Integer
   Global edges(100) As Edge
   Global stress As Boolean
   Global random As Boolean
   Global selnodes(100) As Integer
    
    Dim pickfixed As Boolean
      

    Const fixedColor = vbRed
    Const selectColor = vbMagenta
    Const edgeColor = vbBlack
    Const nodeColor = vbCyan
    Const stressColor = vbGreen
    Const arcColor1 = vbBlack
    Const arcColor2 = selectColor
    Const arcColor3 = fixedColor
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
        x As Long
        y As Long
End Type






Function findNode(lbl As String) As Integer
For i = 1 To nnodes
If nodes(i).lbl = lbl Then findNode = i: Exit Function
Next
findNode = addNode(lbl)

End Function
        
Function addNode(lbl As String) As Integer
    Dim n As Node
    Randomize
    
    n.x = 10 + 380 * Rnd
    Randomize
    n.y = 10 + 380 * Rnd
    n.lbl = lbl

    nnodes = nnodes + 1
    nodes(nnodes) = n
    addNode = nnodes
    
   End Function
   
Sub addEdge(from As String, sto As String, l As Integer)
    Dim e As Edge
    e.from = findNode(from)
    e.sto = findNode(sto)
    e.l = l
    nedges = nedges + 1
    edges(nedges) = e
    
End Sub

Sub run()

        Call relax
        Randomize

        If random = True And Rnd < 0.03 Then
        Dim n As Node
        Randomize
        dc = Int(Rnd(nnodes))
         n = nodes(dc)
         Randomize
        If Not n.fixed Then
            n.x = n.x + 100 * Rnd - 50
            n.y = n.y + 100 * Rnd - 50
        End If
        nodes(dc) = n
        End If
               
End Sub

Sub relax()

'GoTo up
    Dim d As PictureBox
    Dim dlen As Double
    Dim n2 As Node
    Set d = Form1.Picturebox1
    Dim vx, vy, l, f, dx, dy As Double
    
    For i = 1 To nedges
        Dim e As Edge
        e = edges(i)
        
        vx = nodes(e.sto).x - nodes(e.from).x
        vy = nodes(e.sto).y - nodes(e.from).y

        l = Sqr(vx * vx + vy * vy)
        f = (edges(i).l - l) / (l * 3)
        dx = f * vx
        dy = f * vy

        nodes(e.sto).dx = nodes(e.sto).dx + dx
        nodes(e.sto).dy = nodes(e.sto).dy + dy
        nodes(e.from).dx = nodes(e.from).dx - dx
        nodes(e.from).dy = nodes(e.from).dy - dy

    Next i

    Dim n1 As Node
    For i = 1 To nnodes

        n1 = nodes(i)
        dx = 0
        dy = 0

        For j = 1 To nnodes
            If i = j Then GoTo dn
        
            n2 = nodes(j)
            vx = n1.x - n2.x
            vy = n1.y - n2.y
            l = vx * vx + vy * vy
            Randomize
            If (l = 0) Then
                dx = dx + Rnd
                dy = dy + Rnd
            ElseIf l < (100 * 100) Then
                dx = dx + vx / l
                dy = dy + vy / l
            End If
        
        Next j

        dlen = dx * dx + dy * dy
        
        If dlen > 0 Then
            dlen = Sqr(dlen) / 2
            n1.dx = n1.dx + dx / dlen
            n1.dy = n1.dy + dy / dlen
            nodes(i) = n1

        End If
dn:

    Next i
    
    
        

     Dim n As Node
    For i = 1 To nnodes

        n = nodes(i)
        If Not n.fixed Then

        n.x = n.x + MAX(-2, Min(2, n.dx))
        n.y = n.y + MAX(-2, Min(2, n.dy))


        If n.x < 0 Then
            n.x = 0
        ElseIf (n.x > d.ScaleWidth) Then
            n.x = d.ScaleWidth
        End If
        
        If n.y < 0 Then
            n.y = 0
        ElseIf (n.y > d.ScaleHeight) Then
            n.y = d.ScaleHeight
        End If
        
        End If
        n.dx = n.dx / 2
        n.dy = n.dy / 2
        nodes(i) = n

    Next i
up:
    update
    End Sub
    

    Sub paintNode(n As Node)
    
    Dim g As PictureBox
    Set g = Form1.Picturebox1
    x = n.x
    y = n.y
    g.Circle (x, y), 2
    
    Dim pick As Node
    If checknode(n, pick) Then colo = selectColor Else If n.fixed Then colo = fixedColor Else colo = nodeColor
    g.ForeColor = vbBlue
    If Form1.Check1 Then g.PSet (x, y):       g.Print n.lbl
    

    End Sub

    Sub update()
    Dim g As PictureBox
    Set g = Form1.Picturebox1
      
    g.Cls
    For i = 1 To nedges
    Dim e As Edge
        e = edges(i)
        X1 = nodes(e.from).x
        Y1 = nodes(e.from).y
        X2 = nodes(e.sto).x
        Y2 = nodes(e.sto).y

 

        l = (Sqr((X1 - X2) * (X1 - X2) + (Y1 - Y2) * (Y1 - Y2)) - e.l)
        If l < 10 Then colo = arcColor1 Else If l < 20 Then colo = arcColor2 Else colo = arcColor3
        g.FillColor = colo
        g.FillStyle = 0
       
        g.Line (X2, Y2)-(X1, Y1), colo
             
    Next i
    
    For i = 1 To nnodes
        Call paintNode(nodes(i))
    Next i
    
    End Sub

    

    Sub init()

    Call addEdge("a1", "a2", 50)
    Call addEdge("a2", "a3", 50)
    Call addEdge("a3", "b1", 50)
    Call addEdge("b1", "b2", 50)
    Call addEdge("b2", "b3", 50)
    Call addEdge("b3", "c1", 50)
    Call addEdge("c1", "c2", 50)
    Call addEdge("c2", "c3", 50)
    Call addEdge("c3", "a1", 50)
    Call addEdge("x", "a1", 60)
    Call addEdge("x", "b1", 60)
    Call addEdge("x", "c1", 60)
    'Call addEdge("x", "a6", 50)
    'Call addEdge("x", "b6", 20)
    'Call addEdge("x", "c6", 50)
       
        
    End Sub


Function MAX(a, b)
If a > b Then MAX = a Else MAX = b
End Function

Function Min(a, b)
If a < b Then Min = a Else Min = b
End Function

Function checknode(a As Node, b As Node) As Boolean
If a.dx = b.dx And a.dy = b.dy And a.fixed = b.fixed And a.lbl = b.lbl And a.x = b.x And a.y = b.y Then checknode = True
End Function

'Lil coding done by
'Syed Imran Rais (emnuu@hotmail.com)
'BlackAvenger
'India

