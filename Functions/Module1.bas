Attribute VB_Name = "Module1"
Option Explicit

Public X As Integer
Public y As Integer
Public count As Integer
Public functions() As String
Public functionsColor() As String
Public functionsEnable() As Integer

Public Sub addFunctions(ByVal anew As String)
Module1.count = Module1.count + 1
ReDim Preserve functions(Module1.count)
ReDim Preserve functionsColor(Module1.count)
ReDim Preserve functionsEnable(Module1.count)
functions(Module1.count - 1) = anew
functionsColor(Module1.count - 1) = vbBlack
functionsEnable(Module1.count - 1) = 1
End Sub

Public Sub removeFunctions(ByVal aold As String)
Dim index As Integer
Dim i
index = getIndex(aold)
For i = index To Module1.count - 1
functionsColor(i) = functionsColor(i + 1)
functionsEnable(i) = functionsEnable(i + 1)
functions(i) = functions(i + 1)
Next i
Module1.count = Module1.count - 1
ReDim functionsColor(Module1.count)
ReDim functionsEnable(Module1.count)
ReDim functions(Module1.count)
End Sub

Public Sub setColor(ByVal expression As String, ByVal color As String)
functionsColor(getIndex(expression)) = color
End Sub

Public Sub setEnable(ByVal expression As String, ByVal enable As Integer)
functionsEnable(getIndex(expression)) = enable
End Sub

Public Function getIndex(ByVal arg As String) As Integer
Dim index As Integer
Dim find As Boolean
Dim i
For i = 0 To Module1.count - 1
If functions(i) = arg Then
index = i
find = True
getIndex = index
Exit For
End If
Next i
If find = False Then
getIndex = -1
End If
End Function
