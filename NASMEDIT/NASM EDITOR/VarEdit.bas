Attribute VB_Name = "VarEdit"
Function GetVar(Var As String)
Dim MyString As String
Open "C:\Windows\AstonVariableFile_" & App.Title & Var & ".var" For Binary As #1
MyString = Space(LOF(1))
Get 1, , MyString
Close #1
GetVar = MyString
End Function

Sub SetVar(Var As String, Value)
On Error Resume Next
Dim MyString As String
MyString = Value
Kill "C:\Windows\AstonVariableFile_" & App.Title & Var & ".var"
Open "C:\Windows\AstonVariableFile_" & App.Title & Var & ".var" For Binary As #1
Put 1, , MyString
Close #1
End Sub
