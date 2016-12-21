Attribute VB_Name = "Data"
Option Explicit

Public x1 As Integer
Public y1 As Integer
Public x2 As Integer
Public y2 As Integer
Public count As Integer
Public functions() As String
Public functionsColor() As String
Public functionsEnable() As Integer

Public Sub addFunctions(ByVal anew As String)
Data.count = Data.count + 1
ReDim Preserve functions(Data.count)
ReDim Preserve functionsColor(Data.count)
ReDim Preserve functionsEnable(Data.count)
functions(Data.count - 1) = anew
functionsColor(Data.count - 1) = vbBlack
functionsEnable(Data.count - 1) = 1
End Sub

Public Sub removeFunctions(ByVal aold As String)
Dim index As Integer
Dim i
index = getIndex(aold)
For i = index To Data.count - 1
functionsColor(i) = functionsColor(i + 1)
functionsEnable(i) = functionsEnable(i + 1)
functions(i) = functions(i + 1)
Next i
Data.count = Data.count - 1
ReDim functionsColor(Data.count)
ReDim functionsEnable(Data.count)
ReDim functions(Data.count)
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
For i = 0 To Data.count - 1
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
