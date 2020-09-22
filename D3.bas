Attribute VB_Name = "D3"
' Module for rapresenting and calculating vectors in 3D
Type Vec3
    X As Single
    Y As Single
    Z As Single
End Type

' Variables in Coeff
Public v11 As Double
Public v12 As Double
Public v13 As Double
Public v21 As Double
Public v22 As Double
Public v23 As Double
Public v32 As Double
Public v33 As Double
Public v43 As Double


' 3D Global Variables
Public ObjPoint As Vec3

Sub Coeff(rho As Double, Theta As Double, Phi As Double)
'Calculate Sin/etc radians from R, T, P to plot on graph
 Dim costh As Double
 Dim sinth As Double
 Dim cosph As Double
 Dim sinph As Double
   
'   Radians:
 costh = Cos(Theta)
 sinth = Sin(Theta)
 cosph = Cos(Phi)
 sinph = Sin(Phi)
 v11 = -sinth
 v12 = -cosph * costh
 v13 = -sinph * costh
 v21 = costh
 v22 = -cosph * sinth
 v23 = -sinph * sinth
 v32 = sinph
 v33 = -cosph
 v43 = rho
End Sub

Function CopyVec3(v As Vec3) As Vec3
    'Copy the v (vector 3 points) into a new one
    CopyVec3 = v
End Function

Function AssignVec3(X As Double, Y As Double, Z As Double) As Vec3
    'Change values to vector
    AssignVec3.X = X
    AssignVec3.Y = Y
    AssignVec3.Z = Z
End Function

Function DotProduct(a As Vec3, b As Vec3) As Double
   DotProduct = a.X * b.X + a.Y * b.Y + a.Z * b.Z
End Function

Sub Eyecoord(pw As Vec3, pe As Vec3)
  pe.X = v11 * pw.X + v21 * pw.Y
  pe.Y = v12 * pw.X + v22 * pw.Y + v32 * pw.Z
  pe.Z = v13 * pw.X + v23 * pw.Y + v33 * pw.Z + v43
End Sub

Function SommaVec3(u As Vec3, v As Vec3) As Vec3
   SommaVec3.X = u.X + v.X
   SommaVec3.Y = u.Y + v.Y
   SommaVec3.Z = u.Z + v.Z
End Function

Function IncrVec3(u As Vec3, v As Vec3) As Vec3
   u.X = u.X + v.X
   u.Y = u.Y + v.Y
   u.Z = u.Z + v.Z
   IncrVec3 = u
End Function

Function DecrVec3(u As Vec3, v As Vec3) As Vec3
   u.X = u.X - v.X
   u.Y = u.Y - v.Y
   u.Z = u.Z - v.Z
   DecrVec3 = u
End Function

Function MultIncVec3(v As Vec3, C As Double) As Vec3
   v.X = C * v.X
   v.Y = C * v.Y
   v.Z = C * v.Z
   MultIncVec3 = v
End Function

Function MoltiplicaVec3(C As Double, v As Vec3) As Vec3
   MoltiplicaVec3.X = C * v.X
   MoltiplicaVec3.Y = C * v.Y
   MoltiplicaVec3.Z = C * v.Z
End Function

Function SottraiVec3(u As Vec3, v As Vec3) As Vec3
   SottraiVec3.X = u.X - v.X
   SottraiVec3.Y = u.Y - v.Y
   SottraiVec3.Z = u.Z - v.Z
End Function
