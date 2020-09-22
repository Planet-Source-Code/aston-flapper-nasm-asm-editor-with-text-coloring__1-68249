Attribute VB_Name = "modAsm"
Option Explicit

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal ptrMC As Long, ByVal P1 As Long, ByVal P2 As Long, ByVal P3 As Long, ByVal P4 As Long) As Long
Private Bytes() As Byte ' Byte array to hold Machine Code
Dim blen As Long
Private ptrMC As Long ' Pointer to begin of Bytes

Public Function asmRun(P1 As Long, P2 As Long, P3 As Long, P4 As Long) As Long
    asmRun = CallWindowProc(ptrMC, P1, P2, P3, P4) 'Goes to the assembly code
End Function

Public Sub asmInject(str As String)
    Dim i As Long
    If Len(str) = 0 Then Exit Sub
    ReDim Bytes(Len(str) \ 2 - 1)
    For i = 0 To Len(str) \ 2 - 1
        Bytes(i) = CByte("&H" & Mid(str, i * 2 + 1, 2))
    Next i
    blen = Len(str) \ 2
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
