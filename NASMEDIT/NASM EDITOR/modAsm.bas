Attribute VB_Name = "modAsm"
Option Explicit

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal ptrMC As Long, ByVal P1 As Long, ByVal P2 As Long, ByVal P3 As Long, ByVal P4 As Long) As Long
Private Bytes() As Byte
Private ptrMC As Long

Public Function asmRun(P1 As Long, P2 As Long, P3 As Long, P4 As Long) As Long
    If ptrMC = 0 Then Exit Function
    Dim ret As Long
    ret = CallWindowProc(ptrMC, P1, P2, P3, P4)
    asmRun = ret
End Function

Public Sub asmInject(str As String)
    Dim i As Long
    If Len(str) = 0 Then Exit Sub
    ReDim Bytes(Len(str) \ 2 - 1)
    For i = 0 To Len(str) \ 2 - 1
        Bytes(i) = CByte("&H" & Mid(str, i * 2 + 1, 2))
    Next i
    ptrMC = VarPtr(Bytes(0))
End Sub

Public Sub asmInjectRaw(str As String)
    Dim i As Long
    If Len(str) = 0 Then Exit Sub
    ReDim Bytes(Len(str) - 1)
    For i = 0 To Len(str) - 1
        Bytes(i) = CByte(Asc(Mid(str, i + 1, 1)))
    Next i
    ptrMC = VarPtr(Bytes(0))
End Sub
