Attribute VB_Name = "Triangulate"
' Module for polygon triangulation.
' I split up an array of convex polygons
' into another composed only of triangles

Type Trianrs
  a As Integer
  b As Integer
  C As Integer
End Type

Function Triangul(pol() As Integer, n As Integer, nrs() As Trianrs, Ori As Integer) As Integer
'   Triangulation of a polygon with consecutive vertices numbers
'      pol[0],..., pol[n-1], counterclockwise, with 3 numbered
'      verteices P,Q,R <- SetOrient function must determine
'      their orientation.
'      ---------------------------
'      Negative = clockwise
'      Zero     = on same line
'      Positive = counterclockwise
'      ---------------------------
'      If the triangulation is possible, the returnned triangles
'      are stored in the array 'nrs'. Triangle j as vertex nrs[j].A, nrs[j].B, nrs[j].C.
'      ---------------------------
'      This Returns...
'            Number of triangles found or:
'            -1  if the polygon as less 3 vertex or is clockwise.
'            -2  generic error.

   Dim Ptr() As Integer, ort() As Integer
   Dim q As Integer, qA As Integer, qB As Integer, qC As Integer, r As Integer '  -1 usato come 'NULL'
   Dim I As Integer, i1 As Integer, i2 As Integer, j As Integer, K As Integer, m As Integer, ok As Integer, ortB As Integer, Polconvex  As Integer
   Dim a As Integer, b As Integer, C As Integer, p As Integer, collinear As Integer
    
   r = -1
   Polconvex = True
    
    If n < 3 Then
       Triangul = -1 '  No polygons
       Exit Function
    End If

    If n = 3 Then
      nrs(0).a = pol(0)
      nrs(0).b = pol(1)
      nrs(0).C = pol(2)
      Triangul = 1   '  Only a triangle
      Exit Function
    End If
    
    ReDim ort(n) '  ort[i] = 1 if is a convex vertex
    
    Do
      collinear = False
       For I = 0 To n - 1
          If I < n - 1 Then i1 = I + 1 Else i1 = 0
          If i1 < n - 1 Then i2 = i1 + 1 Else i2 = 0
          ort(i1) = SetOrient(pol(I), pol(i1), pol(i2))
          If ort(i1) = 0 Then
            collinear = True
            For j = i1 To n - 1
                pol(j) = pol(j + 1)
            Next j
            n = n - 1
            Exit For
          End If
          If ort(i1) < 1 Then Polconvex = False
        Next I
    Loop While collinear
    
    If n < 3 Then
       Triangul = -1
       Exit Function
    End If

    If Polconvex Then
       For j = 0 To n - 2
          nrs(j).a = pol(0)
          nrs(j).b = pol(j + 1)
          nrs(j).C = pol(j + 2)
       Next
       
       Erase ort
       Triangul = n - 2
       Exit Function
    End If

    ReDim Ptr(n)

' Build a circular list chained with number vertex
    For I = 1 To n - 1: Ptr(I - 1) = I: Next I
    
    Ptr(n - 1) = 0
    q = 0
    qA = Ptr(q)
    qB = Ptr(qA)
    qC = Ptr(qB)
    j = 0            '  j stored triangle up to now
    
    For m = n To 3 Step -1 '  m= remaining node on circular list.
      For K = 0 To m
       ' try with triangle ABC:
          ortB = ort(qB)
          ok = False
       '   B is candidate, if convex:
          If (ortB > 0) Then
             a = pol(qA)
             b = pol(qB)
             C = pol(qC)
             ok = True
             r = Ptr(qC)
             Do While r <> qA And ok
                p = pol(r)     ' ABC counterclockwise:
                ok = p = a Or p = b Or p = C Or SetOrient(a, b, p) < 0 Or SetOrient(b, C, p) < 0 Or SetOrient(C, a, p) < 0
                r = Ptr(r)
             Loop
          '    ok: P coincide with A, B o C
          '    or is external to ABC
             If ok Then
               nrs(j).a = pol(qA)
               nrs(j).b = pol(qB)
               nrs(j).C = pol(qC)
               j = j + 1
             End If
          End If
          
          If (ok Or ortB = 0) Then
         ' delete triangle from polygon ABC
             Ptr(qA) = qC
             qB = qC
             qC = Ptr(qC)
             If ort(qA) < 1 Then ort(qA) = SetOrient(pol(q), pol(qA), pol(qB))
             If ort(qB) < 1 Then ort(qB) = SetOrient(pol(qA), pol(qB), pol(qC))
             Do While ort(qA) = 0 And m > 2
                Ptr(q) = qB
                qA = qB
                qB = qC
                qC = Ptr(qC)
                m = m - 1
             Loop
             Do While ort(qB) = 0 And m > 2
               Ptr(qA) = qC
               qB = qC
               qC = Ptr(qC)
               m = m - 1
             Loop
             Exit For
          End If
          
          q = qA
          qA = qB
          qB = qC
          qC = Ptr(qC)
       Next
   Next
   Triangul = j '  total triangles
End Function
