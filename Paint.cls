VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rcircle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public omiga1 As Double, omiga2 As Double, _
d As Double, t As Double, interval As Double, _
fai1 As Double, fai2 As Double, py As Double



Public Sub initialize(ByVal omiga1s As Double, ByVal omiga2s As Double, _
    ByVal ds As Double, ByVal intervals As Double, _
    ByVal fai1s As Double, ByVal fai2s As Double, _
     ByVal pys As Double)
omiga1 = omiga1s
omiga2 = omiga2s
d = ds
t = 0
interval = intervals
fai1 = fai1s
fai2 = fai2s
py = pys
End Sub
Public Sub getnext(ByRef x As Double, y As Double)
t = t + interval
x = (1 - d) * Cos(omiga1 * t) - d * Cos(omiga2 * t - py)
y = (1 - d) * Cos(omiga1 * t - fai1) + d * Cos(omiga2 * t - fai2 - py)

End Sub
Public Sub setd(ds As Double)
d = ds
End Sub
Public Sub setomiga1(o As Double)
omiga1 = o
End Sub
Public Sub setomiga2(o As Double)
omiga2 = o
End Sub
Public Sub setpy(p As Double)
py = p
End Sub
Public Sub setfai1(f As Double)
fai1 = f
End Sub
Public Sub setfai2(f As Double)
fai2 = f
End Sub
