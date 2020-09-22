VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Asm editor"
   ClientHeight    =   7020
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8355
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Text            =   "Form1.frx":0E42
      Top             =   810
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3705
      Left            =   2610
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Text            =   "Form1.frx":10AD
      Top             =   1575
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.TextBox Text7 
      Height          =   6765
      Left            =   5625
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Text            =   "Form1.frx":16FC
      Top             =   45
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3705
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Text            =   "Form1.frx":1B42
      Top             =   675
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3705
      Left            =   2610
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Text            =   "Form1.frx":23A6
      Top             =   495
      Visible         =   0   'False
      Width           =   5055
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   6690
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text4 
      Height          =   5730
      Left            =   3060
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.TextBox Text3 
      Height          =   6765
      Left            =   765
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Text            =   "Form1.frx":29F5
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "Form1.frx":2E20
      Top             =   1215
      Visible         =   0   'False
      Width           =   5055
   End
   Begin MSComDlg.CommonDialog Common 
      Left            =   1710
      Top             =   1980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2790
      Top             =   1215
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   1680
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   2963
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form1.frx":3393
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu rdheuy 
         Caption         =   "New project..."
         Begin VB.Menu emptysa 
            Caption         =   "Empty project"
         End
         Begin VB.Menu af 
            Caption         =   "-"
         End
         Begin VB.Menu gsdta 
            Caption         =   "nAsm projects"
            Begin VB.Menu sfda 
               Caption         =   "CallWindowProc Project"
            End
            Begin VB.Menu mulf 
               Caption         =   "Multiply function example"
            End
         End
         Begin VB.Menu hetydfg 
            Caption         =   "fAsm projects"
            Begin VB.Menu mulf2 
               Caption         =   "CallWindowProc Project"
            End
         End
      End
      Begin VB.Menu hrddya 
         Caption         =   "-"
      End
      Begin VB.Menu opN 
         Caption         =   "Open"
      End
      Begin VB.Menu hdae 
         Caption         =   "Open previsiously opened"
         Begin VB.Menu prevopen 
            Caption         =   "<empty>"
            Index           =   0
         End
         Begin VB.Menu prevopen 
            Caption         =   "<empty>"
            Index           =   1
         End
         Begin VB.Menu prevopen 
            Caption         =   "<empty>"
            Index           =   2
         End
         Begin VB.Menu prevopen 
            Caption         =   "<empty>"
            Index           =   3
         End
         Begin VB.Menu prevopen 
            Caption         =   "<empty>"
            Index           =   4
         End
      End
      Begin VB.Menu reyf 
         Caption         =   "-"
      End
      Begin VB.Menu sag 
         Caption         =   "Save"
      End
      Begin VB.Menu sav 
         Caption         =   "Save as..."
      End
      Begin VB.Menu hrwayt 
         Caption         =   "-"
      End
      Begin VB.Menu extbye 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu asetg 
      Caption         =   "Edit"
      Begin VB.Menu cac 
         Caption         =   "Copy as Const..."
      End
      Begin VB.Menu grsat 
         Caption         =   "-"
      End
      Begin VB.Menu th 
         Caption         =   "Calculate to hex"
      End
      Begin VB.Menu td 
         Caption         =   "Calculate to decimall"
      End
      Begin VB.Menu sdgfset 
         Caption         =   "-"
      End
      Begin VB.Menu eeln 
         Caption         =   "Set every line even length"
      End
      Begin VB.Menu addcomm 
         Caption         =   "Add comments"
      End
      Begin VB.Menu remcom 
         Caption         =   "Remove all comments"
      End
   End
   Begin VB.Menu inst 
      Caption         =   "Insert"
      Begin VB.Menu defin 
         Caption         =   "DEFINE"
      End
      Begin VB.Menu rhadf 
         Caption         =   "-"
      End
      Begin VB.Menu tuj 
         Caption         =   "Constants"
         Begin VB.Menu messag 
            Caption         =   "String"
         End
      End
      Begin VB.Menu ate 
         Caption         =   "Registers"
         Begin VB.Menu sgest 
            Caption         =   "Basic registers"
            Begin VB.Menu regeax 
               Caption         =   "eax - Accumulator"
            End
            Begin VB.Menu regebx 
               Caption         =   "ebx - Base register"
            End
            Begin VB.Menu regecx 
               Caption         =   "ecx - Counting register"
            End
            Begin VB.Menu regedx 
               Caption         =   "edx - Data register"
            End
            Begin VB.Menu ayrr 
               Caption         =   "-"
            End
            Begin VB.Menu regal 
               Caption         =   "al - eax first byte"
            End
            Begin VB.Menu regbl 
               Caption         =   "bl - ebx first byte"
            End
            Begin VB.Menu regcl 
               Caption         =   "cl - ecx first byte"
            End
            Begin VB.Menu regdl 
               Caption         =   "dl - edx first byte"
            End
            Begin VB.Menu dgahr 
               Caption         =   "-"
            End
            Begin VB.Menu regah 
               Caption         =   "ah - eax last byte"
            End
            Begin VB.Menu regbh 
               Caption         =   "bh - ebx last byte"
            End
            Begin VB.Menu regch 
               Caption         =   "ch - ecx last byte"
            End
            Begin VB.Menu regdh 
               Caption         =   "dh - edx last byte"
            End
         End
         Begin VB.Menu sdgert 
            Caption         =   "-"
         End
         Begin VB.Menu regds 
            Caption         =   "ds - Data segment register"
         End
         Begin VB.Menu reges 
            Caption         =   "es - Extra segment register"
         End
         Begin VB.Menu est 
            Caption         =   "ss - Battery segment register"
         End
         Begin VB.Menu tes 
            Caption         =   "cs - Code segment register"
         End
         Begin VB.Menu dafgr 
            Caption         =   "-"
         End
         Begin VB.Menu regebp 
            Caption         =   "ebp - Base pointers register"
         End
         Begin VB.Menu regesp 
            Caption         =   "esp - Battery pointer register"
         End
         Begin VB.Menu regedi 
            Caption         =   "edi - Destiny index register"
         End
         Begin VB.Menu regesi 
            Caption         =   "esi - Source index register"
         End
      End
      Begin VB.Menu sdge 
         Caption         =   "Operators"
         Begin VB.Menu opAND 
            Caption         =   "AND"
         End
         Begin VB.Menu opOR 
            Caption         =   "OR"
         End
         Begin VB.Menu opXOR 
            Caption         =   "XOR"
         End
         Begin VB.Menu opNOT 
            Caption         =   "NOT"
         End
         Begin VB.Menu agtweat 
            Caption         =   "-"
         End
         Begin VB.Menu opNEG 
            Caption         =   "NEG - Negative"
         End
         Begin VB.Menu opSUB 
            Caption         =   "SUB - Substract"
         End
         Begin VB.Menu opADD 
            Caption         =   "ADD - Add"
         End
         Begin VB.Menu opINC 
            Caption         =   "INC - Increase"
         End
         Begin VB.Menu opDEC 
            Caption         =   "DEC - Decrease"
         End
         Begin VB.Menu opMUL 
            Caption         =   "MUL - Multiply"
         End
         Begin VB.Menu opIMUL 
            Caption         =   "IMUL - Multiply"
         End
         Begin VB.Menu opDIV 
            Caption         =   "DIV - Divide"
         End
         Begin VB.Menu opIDIV 
            Caption         =   "IDIV - Divide"
         End
         Begin VB.Menu dhre 
            Caption         =   "-"
         End
         Begin VB.Menu opCMP 
            Caption         =   "CMP - Compare"
         End
         Begin VB.Menu opTEST 
            Caption         =   "TEST - Test"
         End
      End
      Begin VB.Menu jmpsdf 
         Caption         =   "Jump"
         Begin VB.Menu jmpJUMP 
            Caption         =   "JMP - Jump"
         End
         Begin VB.Menu sdfe 
            Caption         =   "-"
         End
         Begin VB.Menu jmpJE 
            Caption         =   "JE - Jump if equal"
         End
         Begin VB.Menu jmpJG 
            Caption         =   "JG - Jump if greater"
         End
         Begin VB.Menu jmpJL 
            Caption         =   "JL - Jump if lower"
         End
         Begin VB.Menu dsafge 
            Caption         =   "-"
         End
         Begin VB.Menu jmpJNE 
            Caption         =   "JNE - Jump if not equal"
         End
         Begin VB.Menu jmpJNG 
            Caption         =   "JNG - Jump if not greater"
         End
         Begin VB.Menu jmpJNL 
            Caption         =   "JNL - Jump if not lower"
         End
         Begin VB.Menu sdgewat 
            Caption         =   "-"
         End
         Begin VB.Menu jmpJGE 
            Caption         =   "JGE - Jump if equal or greater"
         End
         Begin VB.Menu jmpJLE 
            Caption         =   "JLE - Jump if equal or lower"
         End
         Begin VB.Menu sdfgg 
            Caption         =   "-"
         End
         Begin VB.Menu jmpJC 
            Caption         =   "JC - Jump if Carry"
         End
         Begin VB.Menu jmpJO 
            Caption         =   "JO - Jump if Overflow"
         End
         Begin VB.Menu jmpJP 
            Caption         =   "JP - Jump if Parity"
         End
         Begin VB.Menu jmpJS 
            Caption         =   "JS - Jump if Signed"
         End
         Begin VB.Menu jmpJZ 
            Caption         =   "JZ - Jump if Zero"
         End
         Begin VB.Menu sgeta 
            Caption         =   "-"
         End
         Begin VB.Menu jmpJNC 
            Caption         =   "JNC - Jump if not Carry"
         End
         Begin VB.Menu jmpJNO 
            Caption         =   "JNO - Jump if not Overflow"
         End
         Begin VB.Menu jmpJNP 
            Caption         =   "JNP - Jump if not Parity"
         End
         Begin VB.Menu jmpJNS 
            Caption         =   "JNS - Jump if not Signed"
         End
         Begin VB.Menu jmpJNZ 
            Caption         =   "JNZ - Jump if not Zero"
         End
      End
      Begin VB.Menu flginst 
         Caption         =   "Flag instructions"
         Begin VB.Menu flgCLC 
            Caption         =   "CLC - CF Flag on"
         End
         Begin VB.Menu flgCLD 
            Caption         =   "CLD - DF Flag on"
         End
         Begin VB.Menu flgCLI 
            Caption         =   "CLI - IF Flag on"
         End
         Begin VB.Menu flgCMC 
            Caption         =   "CMC - CF Flag is not CF Flag"
         End
         Begin VB.Menu flgSTC 
            Caption         =   "STC - CF Flag off"
         End
         Begin VB.Menu flgSTD 
            Caption         =   "STD - DF Flag off"
         End
         Begin VB.Menu flgSTI 
            Caption         =   "STI - IF Flag off"
         End
      End
      Begin VB.Menu hear 
         Caption         =   "-"
      End
      Begin VB.Menu intkls 
         Caption         =   "Interrupts"
         Begin VB.Menu int33d 
            Caption         =   "INT 33 (21h)"
            Begin VB.Menu int33_1charScreen 
               Caption         =   "DL(1 character) -> screen"
            End
            Begin VB.Menu int33_mchar_screen 
               Caption         =   "DS:DX - ""$"" -> screen"
            End
            Begin VB.Menu int33inputkey 
               Caption         =   "AL <- keyboard"
            End
         End
      End
   End
   Begin VB.Menu bulksakdlfj 
      Caption         =   "Build"
      Begin VB.Menu sds 
         Caption         =   "Show dos screen when compiling"
      End
      Begin VB.Menu sfgfs 
         Caption         =   "-"
      End
      Begin VB.Menu CreateEXE 
         Caption         =   "Create COM..."
      End
      Begin VB.Menu transhex 
         Caption         =   "Translate to HexData"
      End
      Begin VB.Menu dfhgr 
         Caption         =   "-"
      End
      Begin VB.Menu testvvb 
         Caption         =   "Test if compatible with VB (will crash if not)"
      End
      Begin VB.Menu dbg 
         Caption         =   "Debug for errors"
      End
   End
   Begin VB.Menu fhar 
      Caption         =   "Options"
      Begin VB.Menu adrhgr 
         Caption         =   "Build with"
         Begin VB.Menu bnasm 
            Caption         =   "nAsm"
            Checked         =   -1  'True
         End
         Begin VB.Menu bfasm 
            Caption         =   "fAsm"
         End
      End
   End
   Begin VB.Menu sdf 
      Caption         =   "Info"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
                
Dim chcur As Boolean
Dim last1 As String
Dim last2 As String
Dim filp As String
Dim chng As Boolean
Dim Opens(4) As String

Public Function AppPath() As String
    If Right$(App.Path, 1) = "\" Then
        AppPath = App.Path
    Else
        AppPath = App.Path & "\"
    End If
End Function

Private Sub addcomm_Click()
    remcom_Click
    Dim Txt As String
    Dim t() As String
    Dim a As Long
    Dim l As Long
    l = -1
    Txt = Text1.Text
    t = Split(Txt, vbCrLf)
    Txt = ""
    If l = -1 Then
        For a = 0 To UBound(t)
            If Len(t(a)) > l Then l = Len(t(a))
        Next
    End If
    For a = 0 To UBound(t)
        Txt = Txt & StrSetLength(t(a), l) & vbCrLf
    Next
    t = Split(Txt, vbCrLf)
    Txt = ""
    For a = 0 To UBound(t)
        If info(t(a)) <> "" Then
            Txt = Txt & t(a) & "   ;" & info(t(a)) & vbCrLf
        Else
            Txt = Txt & t(a) & vbCrLf
        End If
    Next
    Text1.Text = Txt
End Sub

Private Sub bfasm_Click()
    bnasm.Checked = False
    bfasm.Checked = True
    Text1_Change
    SetVar "optasm", 2
End Sub

Private Sub bnasm_Click()
    bnasm.Checked = True
    bfasm.Checked = False
    Text1_Change
    SetVar "optasm", 1
End Sub

Private Sub cac_Click()
    Dialog.Text1.Text = StrToHex(Text1.Text, "")
    If Len(Dialog.Text1.Text) <= 900 Then
        Dialog.Text2.Text = "Const myconst = " & Chr$(34) & Dialog.Text1.Text & Chr$(34)
    Else
        Dim a As Long
        Dialog.Text2.Text = "Const myconst = " & Chr$(34) & Mid$(Dialog.Text1.Text, 1, 900) & Chr$(34) & " & _" & vbCrLf
        
        For a = 901 To Len(Dialog.Text1) Step 900
            Dialog.Text2.Text = Dialog.Text2.Text & vbTab & vbTab & vbTab & vbTab & Chr$(34) & Mid$(Dialog.Text1.Text, a, 900) & Chr$(34) & " & _" & vbCrLf
        Next
        Dialog.Text2.Text = Left$(Dialog.Text2.Text, Len(Dialog.Text2.Text) - 6)
    End If
    Dialog.DialogOpen Form1
End Sub

Private Sub Command1_Click()

End Sub

Private Sub CreateEXE_Click()
    On Error GoTo exits
    Common.CancelError = True
    Common.Filter = ".Com files|*.com"
    Common.ShowSave
    StatusBar1.SimpleText = "Please wait..."
    On Error Resume Next
    Dim rndlng As Long
    Dim rndstr As String
    Dim ff As Long
    Dim g As Long
    Dim dat As String
    Dim ret As Long
    Dim str As String
    Dim fpathf As String
    str = Common.FileName
    Common.FileName = ""
    Randomize Timer
    rndlng = Int(Rnd * 50000) + 1
    rndstr = "build" & rndlng & ".asm"
    ff = FreeFile
    Open AppPath & rndstr For Binary As ff
        Put ff, , Text1.Text
    Close ff
    ff = 0
    ff = FreeFile
    Open AppPath & rndlng & ".bat" For Binary As ff
        dat = "cd " & AppPath & vbCrLf
        dat = dat & "@ECHO OFF" & vbCrLf
        If bnasm.Checked Then
            dat = dat & "nasmw.exe " & rndstr & " -o " & "build" & rndlng & ".bin" & vbCrLf
        Else
            dat = dat & "fasm.exe " & rndstr & vbCrLf
        End If
        dat = dat & "ECHO Done>don" & rndlng & ".txt" & vbCrLf
        'dat = dat & "pause"
        Put ff, , dat
    Close ff
    ff = 0
    If sds.Checked Then
        Shell AppPath & rndlng & ".bat", vbNormalFocus
    Else
        Shell AppPath & rndlng & ".bat", vbHide
    End If
    Do
        DoEvents
        g = Timer
        ff = FreeFile
        Open AppPath & "don" & rndlng & ".txt" For Binary As ff
            dat = Space$(LOF(ff))
            Get ff, , dat
        Close ff
        Do: DoEvents: Loop Until Timer >= g + 0.25
    Loop Until dat <> ""
        ff = FreeFile

        Open AppPath & "build" & rndlng & ".bin" For Binary As ff
            dat = Space$(LOF(ff))
            Get ff, , dat
        Close ff
        If dat = "" Then
            ret = MsgBox("One or more errors occurred!" & vbCrLf & "Do you want to debug for errors?", vbYesNo)
            If ret = vbYes Then
                Kill AppPath & rndlng & ".bat"
                Open AppPath & (rndlng + 1) & ".bat" For Binary As ff
                    dat = "cd " & AppPath & vbCrLf
                    dat = dat & "@ECHO OFF" & vbCrLf
                    If bnasm.Checked Then
                        dat = dat & "nasmw.exe " & rndstr & " -o " & "build" & rndlng & ".bin" & vbCrLf
                    Else
                        dat = dat & "fasm.exe " & rndstr & vbCrLf
                    End If
                    dat = dat & "pause" & vbCrLf
                    dat = dat & "del " & (rndlng + 1) & ".bat" & vbCrLf
                    Put ff, , dat
                Close ff
                Shell AppPath & (rndlng + 1) & ".bat", vbNormalFocus
            Else
            End If
        End If
        Kill AppPath & rndlng & ".bat"
        Kill AppPath & rndstr
        Kill AppPath & "don" & rndlng & ".txt"
        FileMove AppPath & "build" & rndlng & ".bin", str
        StatusBar1.SimpleText = "Stopped compiling."
exits:
    rndlng = 0
End Sub

Private Sub dbg_Click()
    On Error Resume Next
    Dim rt As Long
    Dim rndlng As Long
    Dim rndstr As String
    Dim ff As Long
    Dim g As Long
    Dim dat As String
    Dim ret As Long
    Dim str As String
    StatusBar1.SimpleText = "Please wait..."
    str = Common.FileName
    Common.FileName = ""
    Randomize Timer
    rndlng = Int(Rnd * 50000) + 1
    rndstr = "build" & rndlng & ".asm"
    ff = FreeFile
    Open AppPath & rndstr For Binary As ff
        Put ff, , Text1.Text
    Close ff
    ff = 0
    ff = FreeFile
    Open AppPath & rndlng & ".bat" For Binary As ff
        dat = "cd " & AppPath & vbCrLf
        dat = dat & "@ECHO OFF" & vbCrLf
        If bnasm.Checked Then
            dat = dat & "nasmw.exe " & rndstr & vbCrLf
        Else
            dat = dat & "fasm.exe " & rndstr & vbCrLf
        End If
        dat = dat & "ECHO Done>don" & rndlng & ".txt" & vbCrLf
        'dat = dat & "pause"
        Put ff, , dat
    Close ff
    ff = 0

                Kill AppPath & rndlng & ".bat"
                ff = FreeFile
                Open AppPath & (rndlng + 1) & ".bat" For Binary As ff
                    dat = "cd " & AppPath & vbCrLf
                    dat = dat & "@ECHO OFF" & vbCrLf
                    If bnasm.Checked Then
                        dat = dat & "nasmw.exe " & rndstr & vbCrLf
                    Else
                        dat = dat & "fasm.exe " & rndstr & vbCrLf
                    End If
                    dat = dat & "pause" & vbCrLf
                    dat = dat & "del " & (rndlng + 1) & ".bat" & vbCrLf
                    Put ff, , dat
                Close ff
                Shell AppPath & (rndlng + 1) & ".bat", vbNormalFocus
        rndlng = 0
        StatusBar1.SimpleText = "Stopped compiling."
End Sub

Private Sub defin_Click()
    Text1.SelText = "%DEFINE"
End Sub

Private Sub eeln_Click()
    On Error GoTo ers
    Dim l As Long
        l = InputBox("What length must every line be?" & vbCrLf & "-1 for longest line", , "-1")
    Dim Txt As String
    Dim t() As String
    Dim a As Long
    Txt = Text1.Text
    t = Split(Txt, vbCrLf)
    Txt = ""
    If l = -1 Then
        For a = 0 To UBound(t)
                If Len(t(a)) > l Then l = Len(t(a))
        Next
    End If
    For a = 0 To UBound(t)
            Txt = Txt & StrSetLength(t(a), l) & vbCrLf
    Next
    Text1.Text = Txt
ers:
End Sub

Private Sub emptysa_Click()
    If chng Then
        Dim askkk As Long
        askkk = MsgBox("Are you sure?", vbYesNo)
        If askkk = vbNo Then Exit Sub
    End If
    Text1.Text = ""
    filp = ""
    chng = False
End Sub

Private Sub est_Click()
    Text1.SelText = "ss"
End Sub

Private Sub extbye_Click()
    If chng Then
        Dim askkk As Long
        askkk = MsgBox("Are you sure?", vbYesNo)
        If askkk = vbNo Then Exit Sub
    End If
    Dim h As String
    Dim a As Long
    For a = 0 To 4
        h = h & Opens(a) & ";"
    Next
    h = Left$(h, Len(h) - 1)
    SetVar "jslkeujir", h
    End
End Sub

Private Sub flgCLC_Click()
    Text1.SelText = "CLC"
End Sub

Private Sub flgCLD_Click()
    Text1.SelText = "CLD"
End Sub

Private Sub flgCLI_Click()
    Text1.SelText = "CLI"
End Sub

Private Sub flgCMC_Click()
    Text1.SelText = "CMC"
End Sub

Private Sub flgSTC_Click()
    Text1.SelText = "STC"
End Sub

Private Sub flgSTD_Click()
    Text1.SelText = "STD"
End Sub

Private Sub flgSTI_Click()
    Text1.SelText = "STI"
End Sub

Private Sub Form_Load()
    If GetVar("optasm") = 1 Then
        bnasm.Checked = True
        bfasm.Checked = False
    Else
        bnasm.Checked = False
        bfasm.Checked = True
    End If
    
    If Command = "" Then
        Text1.Text = vbCrLf
    Else
        Dim ff As Long
        Dim dat As String
        ff = FreeFile
        dat = Command
        If Left$(dat, 1) = Chr$(34) Then dat = StrBetween(dat, Chr$(34), Chr$(34))
        Open dat For Binary As ff
            dat = Space$(LOF(ff))
            Get ff, , dat
        Close ff
        Text1.Text = dat
        filp = Command
        RemOpen Command
        AddOpen Command
        LoadOpen
    End If
    Dim a As Long
    If GetVar("jslkeujir") <> "" Then
        For a = 0 To UBound(Split(GetVar("jslkeujir"), ";"))
            Opens(a) = Split(GetVar("jslkeujir"), ";")(a)
        Next
    End If
    LoadOpen
    chng = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If chng Then
        Dim askkk As Long
        askkk = MsgBox("Are you sure?", vbYesNo)
        If askkk = vbNo Then Cancel = True: Exit Sub
    End If
    Dim h As String
    Dim a As Long
    For a = 0 To 4
        h = h & Opens(a) & ";"
    Next
    h = Left$(h, Len(h) - 1)
    SetVar "jslkeujir", h
    End
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Text1.Left = 0
    Text1.Top = 0
    Text1.Width = Form1.Width - 120
    Text1.Height = Form1.Height - 810 - StatusBar1.Height
End Sub

Private Sub int33_1charScreen_Click()
    Text1.SelText = "MOV ah, 2" & vbCrLf & "INT  33"
End Sub

Private Sub int33_mchar_screen_Click()
    Text1.SelText = "MOV ah, 9" & vbCrLf & "INT  33"
End Sub

Private Sub int33inputkey_Click()
    Text1.SelText = "MOV ah, 1" & vbCrLf & "INT  33"
End Sub

Private Sub jmpJC_Click()
    Text1.SelText = "JC"
End Sub

Private Sub jmpJE_Click()
    Text1.SelText = "JE"
End Sub

Private Sub jmpJG_Click()
    Text1.SelText = "JG"
End Sub

Private Sub jmpJGE_Click()
    Text1.SelText = "JGE"
End Sub

Private Sub jmpJL_Click()
    Text1.SelText = "JL"
End Sub

Private Sub jmpJLE_Click()
    Text1.SelText = "JLE"
End Sub

Private Sub jmpJNC_Click()
    Text1.SelText = "JNC"
End Sub

Private Sub jmpJNE_Click()
    Text1.SelText = "JNE"
End Sub

Private Sub jmpJNG_Click()
    Text1.SelText = "JNG"
End Sub

Private Sub jmpJNL_Click()
    Text1.SelText = "JNL"
End Sub

Private Sub jmpJNO_Click()
    Text1.SelText = "JNO"
End Sub

Private Sub jmpJNP_Click()
    Text1.SelText = "JNP"
End Sub

Private Sub jmpJNS_Click()
    Text1.SelText = "JNS"
End Sub

Private Sub jmpJNZ_Click()
    Text1.SelText = "JNZ"
End Sub

Private Sub jmpJO_Click()
    Text1.SelText = "JO"
End Sub

Private Sub jmpJP_Click()
    Text1.SelText = "JP"
End Sub

Private Sub jmpJS_Click()
    Text1.SelText = "JS"
End Sub

Private Sub jmpJUMP_Click()
    Text1.SelText = "JMP"
End Sub

Private Sub jmpJZ_Click()
    Text1.SelText = "JZ"
End Sub

Private Sub messag_Click()
    Text1.SelText = "MyConst db " & Chr$(34) & "MyValue" & Chr$(34)
End Sub

Private Sub mulf_Click()
    If chng Then
        Dim askkk As Long
        askkk = MsgBox("Are you sure?", vbYesNo)
        If askkk = vbNo Then Exit Sub
    End If
    Text1.Text = Text5.Text
    filp = ""
    chng = False
End Sub

Private Sub mulf2_Click()
    If chng Then
        Dim askkk As Long
        askkk = MsgBox("Are you sure?", vbYesNo)
        If askkk = vbNo Then Exit Sub
    End If
    Text1.Text = Text9.Text
    filp = ""
    chng = False
End Sub

Private Sub opADD_Click()
    Text1.SelText = "ADD"
End Sub

Private Sub opAND_Click()
    Text1.SelText = "AND"
End Sub

Private Sub opCMP_Click()
    Text1.SelText = "CMP"
End Sub

Private Sub opDEC_Click()
    Text1.SelText = "DEC"
End Sub

Private Sub opDIV_Click()
    Text1.SelText = "DIV"
End Sub

Private Sub opIDIV_Click()
    Text1.SelText = "IDIV"
End Sub

Private Sub opIMUL_Click()
    Text1.SelText = "IMUL"
End Sub

Private Sub opINC_Click()
    Text1.SelText = "INC"
End Sub

Private Sub opMUL_Click()
    Text1.SelText = "MUL"
End Sub

Private Sub opn_Click()
    If chng Then
        Dim askkk As Long
        askkk = MsgBox("Are you sure?", vbYesNo)
        If askkk = vbNo Then Exit Sub
    End If
    On Error GoTo exits
    Dim ff As Long
    Dim dat As String
    Common.CancelError = True
    Common.Filter = ".Asm files|*.asm|All files|*.*"
    Common.ShowOpen
    ff = FreeFile
    Open Common.FileName For Binary As ff
        dat = Space$(LOF(ff))
        Get ff, , dat
    Close ff
    Text1.Text = dat
    dat = ""
    ff = 0
    filp = Common.FileName
    chng = False
    RemOpen Common.FileName
    AddOpen Common.FileName
    LoadOpen
exits:
End Sub

Private Sub opNEG_Click()
    Text1.SelText = "NEG"
End Sub

Private Sub opNOT_Click()
    Text1.SelText = "NOT"
End Sub

Private Sub opOR_Click()
    Text1.SelText = "OR"
End Sub

Private Sub opSUB_Click()
    Text1.SelText = "SUB"
End Sub

Private Sub opTEST_Click()
    Text1.SelText = "TEST"
End Sub

Private Sub opXOR_Click()
    Text1.SelText = "XOR"
End Sub

Private Sub prevopen_Click(Index As Integer)
    If chng Then
        Dim askkk As Long
        askkk = MsgBox("Are you sure?", vbYesNo)
        If askkk = vbNo Then Exit Sub
    End If
    On Error GoTo exits
    Dim ff As Long
    Dim dat As String
    ff = FreeFile
    Open prevopen(Index).Caption For Binary As ff
        dat = Space$(LOF(ff))
        Get ff, , dat
    Close ff
    Text1.Text = dat
    dat = ""
    ff = 0
    filp = prevopen(Index).Caption
    chng = False
    RemOpen prevopen(Index).Caption
    AddOpen prevopen(Index).Caption
    LoadOpen
exits:
End Sub

Private Sub regah_Click()
    Text1.SelText = "ah"
End Sub

Private Sub regal_Click()
    Text1.SelText = "al"
End Sub

Private Sub regbh_Click()
    Text1.SelText = "bh"
End Sub

Private Sub regbl_Click()
    Text1.SelText = "bl"
End Sub

Private Sub regch_Click()
    Text1.SelText = "ch"
End Sub

Private Sub regcl_Click()
    Text1.SelText = "cl"
End Sub

Private Sub regdh_Click()
    Text1.SelText = "dh"
End Sub

Private Sub regdl_Click()
    Text1.SelText = "dl"
End Sub

Private Sub regds_Click()
    Text1.SelText = "ds"
End Sub

Private Sub regeax_Click()
    Text1.SelText = "eax"
End Sub

Private Sub regebp_Click()
    Text1.SelText = "ebp"
End Sub

Private Sub regebx_Click()
    Text1.SelText = "ebx"
End Sub

Private Sub regecx_Click()
    Text1.SelText = "ecx"
End Sub

Private Sub regedi_Click()
    Text1.SelText = "edi"
End Sub

Private Sub regedx_Click()
    Text1.SelText = "edx"
End Sub

Private Sub reges_Click()
    Text1.SelText = "es"
End Sub

Private Sub regesi_Click()
    Text1.SelText = "esi"
End Sub

Private Sub regesp_Click()
    Text1.SelText = "esp"
End Sub

Private Sub remcom_Click()
    Dim Txt As String
    Dim t() As String
    Dim a As Long
    Dim dat As String
    Txt = Text1.Text
    t = Split(Txt, vbCrLf)
    Txt = ""
    For a = 0 To UBound(t)
            dat = StrTo(t(a), ";")
            Do Until Right$(dat, 1) <> " "
                dat = Left$(dat, Len(dat) - 1)
            Loop
            Txt = Txt & dat & vbCrLf
    Next
    Text1.Text = Txt
End Sub

Private Sub sag_Click()
    If filp <> "" Then
        Dim ff As Long
        ff = FreeFile
        On Error Resume Next
        Kill Common.FileName
        Open Common.FileName For Binary As ff
            Put ff, , Text1.Text
        Close ff
        ff = 0
        chng = False
    Else
    sav_Click
    End If
End Sub

Private Sub sav_Click()
    On Error GoTo exits
    Dim ff As Long
    Common.CancelError = True
    Common.Filter = ".Asm files|*.asm"
    Common.ShowSave
    ff = FreeFile
    On Error Resume Next: Kill Common.FileName: On Error GoTo exits
    Open Common.FileName For Binary As ff
        Put ff, , Text1.Text
    Close ff
    ff = 0
    chng = False
exits:
End Sub

Private Sub sdf_Click()
    MsgBox "This program was created by Aston Flapper", vbInformation
End Sub

Private Sub sds_Click()
    sds.Checked = Not sds.Checked
End Sub

Private Sub sfda_Click()
    If chng Then
        Dim askkk As Long
        askkk = MsgBox("Are you sure?", vbYesNo)
        If askkk = vbNo Then Exit Sub
    End If
    Text1.Text = Text2.Text
    filp = ""
    chng = False
End Sub



Private Sub td_Click()
    Text1.SelText = XHexToDecimall(Text1.SelText)
End Sub

Private Sub tes_Click()
    Text1.SelText = "cs"
End Sub

Private Sub testvvb_Click()
    On Error Resume Next
    Dim rt As Long
    rt = MsgBox("Are you sure you want to test if this script is VB compatible? this program will crash if not!", vbExclamation + vbYesNo)
    If rt = vbNo Then Exit Sub
    Dim rndlng As Long
    Dim rndstr As String
    Dim ff As Long
    Dim g As Long
    Dim dat As String
    Dim ret As Long
    Dim str As String
    StatusBar1.SimpleText = "Please wait..."
    str = Common.FileName
    Common.FileName = ""
    Randomize Timer
    rndlng = Int(Rnd * 50000) + 1
    rndstr = "build" & rndlng & ".asm"
    ff = FreeFile
    Open AppPath & rndstr For Binary As ff
        Put ff, , Text1.Text
    Close ff
    ff = 0
    ff = FreeFile
    Open AppPath & rndlng & ".bat" For Binary As ff
        dat = "cd " & AppPath & vbCrLf
        dat = dat & "@ECHO OFF" & vbCrLf
        If bnasm.Checked Then
            dat = dat & "nasmw.exe " & rndstr & " -o " & "build" & rndlng & ".bin" & vbCrLf
        Else
            dat = dat & "fasm.exe " & rndstr & vbCrLf
        End If
        dat = dat & "ECHO Done>don" & rndlng & ".txt" & vbCrLf
        dat = dat & "pause"
        Put ff, , dat
    Close ff
    ff = 0
    If sds.Checked Then
        Shell AppPath & rndlng & ".bat", vbNormalFocus
    Else
        Shell AppPath & rndlng & ".bat", vbHide
    End If
    Do
        DoEvents
        g = Timer
        ff = FreeFile
        Open AppPath & "don" & rndlng & ".txt" For Binary As ff
            dat = Space$(LOF(ff))
            Get ff, , dat
        Close ff
        Do: DoEvents: Loop Until Timer >= g + 0.25
    Loop Until dat <> ""
        ff = FreeFile
        Open AppPath & "build" & rndlng & ".bin" For Binary As ff
            dat = Space$(LOF(ff))
            Get ff, , dat
        Close ff
        If dat = "" Then
            ret = MsgBox("One or more errors occurred!" & vbCrLf & "Do you want to debug for errors?", vbYesNo)
            If ret = vbYes Then
                Kill AppPath & rndlng & ".bat"
                ff = FreeFile
                Open AppPath & (rndlng + 1) & ".bat" For Binary As ff
                    dat = "cd " & AppPath & vbCrLf
                    dat = dat & "@ECHO OFF" & vbCrLf
                    If bnasm.Checked Then
                        dat = dat & "nasmw.exe " & rndstr & vbCrLf
                    Else
                        dat = dat & "fasm.exe " & rndstr & vbCrLf
                    End If
                    dat = dat & "pause" & vbCrLf
                    dat = dat & "del " & (rndlng + 1) & ".bat" & vbCrLf
                    Put ff, , dat
                Close ff
                Shell AppPath & (rndlng + 1) & ".bat", vbNormalFocus
            Else
            End If
        Else
            Dim P1 As Long
            Dim P2 As Long
            Dim P3 As Long
            Dim P4 As Long
            Dim l1 As String
            Dim l2 As String
            Dim l3 As String
            Dim l4 As String
            Dim s1 As Long
            Dim s2 As Long
            Dim s3 As Long
            Dim s4 As Long
            
            Dim retx As Long
            asmInjectRaw dat
            l1 = InputBox("Param1?" & vbCrLf & "Put $ at te beginning to point out to an string")
            l2 = InputBox("Param2?" & vbCrLf & "Put $ at te beginning to point out to an string")
            l3 = InputBox("Param3?" & vbCrLf & "Put $ at te beginning to point out to an string")
            l4 = InputBox("Param4?" & vbCrLf & "Put $ at te beginning to point out to an string")
            If Left$(l1, 1) = "$" Then l1 = Mid$(l1, 2): P1 = StrPtr(l1)
            If Left$(l2, 1) = "$" Then l2 = Mid$(l2, 2): P2 = StrPtr(l2)
            If Left$(l3, 1) = "$" Then l3 = Mid$(l3, 2): P3 = StrPtr(l3)
            If Left$(l4, 1) = "$" Then l4 = Mid$(l4, 2): P4 = StrPtr(l4)
            If Left$(l1, 1) <> "$" Then P1 = Val(l1)
            If Left$(l2, 1) <> "$" Then P2 = Val(l2)
            If Left$(l3, 1) <> "$" Then P3 = Val(l3)
            If Left$(l4, 1) <> "$" Then P4 = Val(l4)
            s1 = P1
            s2 = P2
            s3 = P3
            s4 = P4
            Dim gtim1 As Double
            Dim txtt As String
            Dim gtim2 As Double
            gtim1 = Timer
            retx = asmRun(P1, P2, P3, P4)
            gtim2 = Timer
            txtt = CStr((gtim2 - gtim1) * 1000)
            MsgBox "Good news! VB will at least not crash if it runs this script!" & vbCrLf & vbCrLf & "Return value: " & retx & vbCrLf & "P1: " & s1 & " to " & P1 & vbCrLf & "P2: " & s2 & " to " & P2 & vbCrLf & "P3: " & s3 & " to " & P3 & vbCrLf & "P4: " & s4 & " to " & P4 & vbCrLf & vbCrLf & "That all in " & txtt & " milliseconds"
            Kill AppPath & rndstr
        End If
        Kill AppPath & rndlng & ".bat"
        Kill AppPath & "don" & rndlng & ".txt"
        Kill AppPath & "build" & rndlng & ".bin"
        ff = FreeFile
        rndlng = 0
        StatusBar1.SimpleText = "Stopped compiling."
End Sub

Private Sub Text1_Change()
    Dim dat As String
    Dim str As String
    Dim g As Long
    Dim p As Long
    Dim c() As String
    Dim a As Long
    Dim s As String
    Dim b As Long
    Dim dat1 As String
    Dim dat2 As String
    Dim dat3 As String
    Dim datt As String
    If bnasm.Checked Then
        chng = True
        g = Text1.SelStart
        p = Text1.SelLength
        dat = Text1.TextRTF
        dat = StrFrom(dat, vbCrLf)
        dat = StrFrom(dat, "\fs17")
        dat = "\par" & dat
        dat = Replace$(dat, "\cf4 ", "")
        dat = Replace$(dat, "\cf3 ", "")
        dat = Replace$(dat, "\cf2 ", "")
        dat = Replace$(dat, "\cf1 ", "")
        dat = Replace$(dat, "\cf0 ", "")
        dat = Replace$(dat, "\tab ", " ")
        dat = Replace$(dat, Chr$(9), "")
        Do
            DoEvents
            dat = Replace$(dat, "\par  ", "\par ")
        Loop Until InStr(dat, "\par  ") = 0
        dat = Replace$(dat, vbCrLf, vbCrLf & "\cf0")
        dat = Replace$(dat, "[", "\cf1 [\cf0 ")
        dat = Replace$(dat, "]", "\cf1 ]\cf0 ")
        dat = Replace$(dat, "+", "\cf1 +\cf0 ")
        dat = Replace$(dat, "-", "\cf1 -\cf0 ")
        dat = Replace$(dat, ":", "\cf3 :\cf0 ")
        dat = Replace$(dat, " db ", " \cf1 db \cf0 ")
        dat = Replace$(dat, " dw ", " \cf1 dw \cf0 ")
        dat = Replace$(dat, " dl ", " \cf1 dl \cf0 ")
        dat = Replace$(dat, " dd ", " \cf1 dd \cf0 ")
        dat = Replace$(dat, " dq ", " \cf1 dq \cf0 ")
        dat = Replace$(dat, " dt ", " \cf1 dt \cf0 ")
        dat = Replace$(dat, " equ ", " \cf1 equ \cf0 ")
        
        dat = Replace$(dat, "\par .", "\par \cf4 .")
        c = Split(Text3.Text, vbCrLf)
        For a = 0 To UBound(c)
            s = c(a)
            If UBound(Split(s, " ")) >= 0 Then s = Split(s, " ")(0)
            If s <> "" Then dat = Replace$(dat, "\par " & s, "\par \cf1 " & s & "\cf0 ", , , vbTextCompare)
            DoEvents
        Next
        c = Split(Text4.Text, vbCrLf)
        For a = 0 To UBound(c)
            If c(a) <> "" Then dat = Replace$(dat, c(a), "\cf3 " & c(a) & "\cf0 ", , , vbTextCompare)
            DoEvents
        Next
        b = 1
        Do
            datt = dat
asda:
            If InStr(b, dat, "'") = 0 Then
                GoTo pta
            ElseIf Right$(Mid$(dat, 1, InStr(b, dat, "'")), 6) <> "\cf3 " & "'" Then
                dat1 = Mid$(dat, 1, InStr(b, dat, "'") - 1)
                dat2 = Mid$(dat, 1 + InStr(b, dat, "'"))
                dat3 = StrFrom(dat2, "'")
                dat2 = "\cf3 " & "'" & StrTo(dat2, "'") & "'" & "\cfo "
                dat2 = Replace$(dat2, "\cf0 ", "\cf3 ")
                dat2 = Replace$(dat2, "\cfo ", "\cf0 ")
                dat2 = Replace$(dat2, "\cf1 ", "\cf3 ")
                dat2 = Replace$(dat2, "\cf3 ", "\cf3 ")
                dat = dat1 & dat2 & dat3
                If datt = dat Then Exit Do
            Else
                b = InStr(b, dat, "'") + 2
                If InStr(b, dat, "'") = 0 Then Exit Do
                b = InStr(b, dat, "'") + 2
                GoTo asda
            End If
pta:
        Loop Until InStr(b, dat, "'") = 0
            b = 1
        Do
            datt = dat
asdaf:
            If InStr(b, dat, Chr$(34)) = 0 Then
                GoTo ptaf
            ElseIf Right$(Mid$(dat, 1, InStr(b, dat, Chr$(34))), 6) <> "\cf3 " & Chr$(34) Then
                dat1 = Mid$(dat, 1, InStr(b, dat, Chr$(34)) - 1)
                dat2 = Mid$(dat, 1 + InStr(b, dat, Chr$(34)))
                dat3 = StrFrom(dat2, Chr$(34))
                
                dat2 = "\cf3 " & Chr$(34) & StrTo(dat2, Chr$(34)) & Chr$(34) & "\cfo "
                dat2 = Replace$(dat2, "\cf0 ", "\cf3 ")
                dat2 = Replace$(dat2, "\cfo ", "\cf0 ")
                dat2 = Replace$(dat2, "\cf1 ", "\cf3 ")
                dat2 = Replace$(dat2, "\cf3 ", "\cf3 ")
                dat = dat1 & dat2 & dat3
                If datt = dat Then Exit Do
            Else
                b = InStr(b, dat, Chr$(34)) + 2
                If InStr(b, dat, Chr$(34)) = 0 Then Exit Do
                b = InStr(b, dat, Chr$(34)) + 2
                GoTo asdaf
            End If
ptaf:
        Loop Until datt = dat
        b = 1
        Do
            datt = dat
asd:
            If InStr(b, dat, ";") = 0 Then
                GoTo pt
            ElseIf Right$(Mid$(dat, 1, InStr(b, dat, ";")), 6) <> "\cf2 ;" Then
                dat1 = Mid$(dat, 1, InStr(b, dat, ";") - 1)
                dat2 = "\cf2 " & Mid$(dat, InStr(b, dat, ";"))
                dat3 = StrFrom(dat2, vbCrLf)
                dat2 = StrTo(dat2, vbCrLf) & vbCrLf
                dat2 = Replace$(dat2, "\cf0", "\cf2")
                dat2 = Replace$(dat2, "\cfo", "\cf0")
                dat2 = Replace$(dat2, "\cf1", "\cf2")
                dat2 = Replace$(dat2, "\cf3", "\cf2")
                dat = dat1 & dat2 & dat3
            Else
                b = InStr(b, dat, ";") + 1
                GoTo asd
            End If
pt:
        Loop Until datt = dat
            b = 1
            
        
        
        Do Until InStr(dat, ",  ") = 0
        dat = Replace$(dat, ",  ", ", ")
        Loop
        dat = Mid$(dat, 5)
        str = "{\rtf1\ansi\ansicpg1252\deff0\deflang1043{\fonttbl{\f0\fnil\fcharset0 Courier new;}}" & vbCrLf
        str = str & "{\colortbl ;\red0\green0\blue200;\red0\green150\blue0;\red150\green0\blue0;\red150\green0\blue150;}" & vbCrLf
        str = str & "\viewkind4\uc1\pard\cf0\f0\fs17"
        chcur = True
        Text1.TextRTF = str & dat
        Text1.SelStart = g
        Text1.SelLength = p
        chcur = False
    Else
        chng = True
        g = Text1.SelStart
        p = Text1.SelLength
        dat = Text1.TextRTF
        dat = StrFrom(dat, vbCrLf)
        dat = StrFrom(dat, "\fs17")
        dat = "\par" & dat
        dat = Replace$(dat, "\cf4 ", "")
        dat = Replace$(dat, "\cf3 ", "")
        dat = Replace$(dat, "\cf2 ", "")
        dat = Replace$(dat, "\cf1 ", "")
        dat = Replace$(dat, "\cf0 ", "")
        dat = Replace$(dat, "\tab ", " ")
        dat = Replace$(dat, Chr$(9), "")
        Do
            DoEvents
            dat = Replace$(dat, "\par  ", "\par ")
        Loop Until InStr(dat, "\par  ") = 0
        dat = Replace$(dat, vbCrLf, vbCrLf & "\cf0")
        dat = Replace$(dat, "[", "\cf1 [\cf0 ")
        dat = Replace$(dat, "]", "\cf1 ]\cf0 ")
        dat = Replace$(dat, "+", "\cf1 +\cf0 ")
        dat = Replace$(dat, "-", "\cf1 -\cf0 ")
        dat = Replace$(dat, ":", "\cf3 :\cf0 ")
        dat = Replace$(dat, " db ", " \cf1 db \cf0 ")
        dat = Replace$(dat, " dw ", " \cf1 dw \cf0 ")
        dat = Replace$(dat, " dl ", " \cf1 dl \cf0 ")
        dat = Replace$(dat, " dd ", " \cf1 dd \cf0 ")
        dat = Replace$(dat, " dq ", " \cf1 dq \cf0 ")
        dat = Replace$(dat, " dt ", " \cf1 dt \cf0 ")
        dat = Replace$(dat, " equ ", " \cf1 equ \cf0 ")
        
        dat = Replace$(dat, "\par .", "\par \cf4 .")
        c = Split(Text7.Text, vbCrLf)
        For a = 0 To UBound(c)
            s = c(a)
            If UBound(Split(s, " ")) >= 0 Then s = Split(s, " ")(0)
            If s <> "" Then dat = Replace$(dat, "\par " & s, "\par \cf1 " & s & "\cf0 ", , , vbTextCompare)
            DoEvents
        Next
        c = Split(Text4.Text, vbCrLf)
        For a = 0 To UBound(c)
            If c(a) <> "" Then dat = Replace$(dat, c(a), "\cf3 " & c(a) & "\cf0 ", , , vbTextCompare)
            DoEvents
        Next
        b = 1
        Do
            datt = dat
asdaae:
            If InStr(b, dat, "'") = 0 Then
                GoTo ptaae
            ElseIf Right$(Mid$(dat, 1, InStr(b, dat, "'")), 6) <> "\cf3 " & "'" Then
                dat1 = Mid$(dat, 1, InStr(b, dat, "'") - 1)
                dat2 = Mid$(dat, 1 + InStr(b, dat, "'"))
                dat3 = StrFrom(dat2, "'")
                dat2 = "\cf3 " & "'" & StrTo(dat2, "'") & "'" & "\cfo "
                dat2 = Replace$(dat2, "\cf0 ", "\cf3 ")
                dat2 = Replace$(dat2, "\cfo ", "\cf0 ")
                dat2 = Replace$(dat2, "\cf1 ", "\cf3 ")
                dat2 = Replace$(dat2, "\cf3 ", "\cf3 ")
                dat = dat1 & dat2 & dat3
                If datt = dat Then Exit Do
            Else
                b = InStr(b, dat, "'") + 2
                If InStr(b, dat, "'") = 0 Then Exit Do
                b = InStr(b, dat, "'") + 2
                GoTo asdaae
            End If
ptaae:
        Loop Until InStr(b, dat, "'") = 0
            b = 1
        Do
            datt = dat
asdafae:
            If InStr(b, dat, Chr$(34)) = 0 Then
                GoTo ptafae
            ElseIf Right$(Mid$(dat, 1, InStr(b, dat, Chr$(34))), 6) <> "\cf3 " & Chr$(34) Then
                dat1 = Mid$(dat, 1, InStr(b, dat, Chr$(34)) - 1)
                dat2 = Mid$(dat, 1 + InStr(b, dat, Chr$(34)))
                dat3 = StrFrom(dat2, Chr$(34))
                
                dat2 = "\cf3 " & Chr$(34) & StrTo(dat2, Chr$(34)) & Chr$(34) & "\cfo "
                dat2 = Replace$(dat2, "\cf0 ", "\cf3 ")
                dat2 = Replace$(dat2, "\cfo ", "\cf0 ")
                dat2 = Replace$(dat2, "\cf1 ", "\cf3 ")
                dat2 = Replace$(dat2, "\cf3 ", "\cf3 ")
                dat = dat1 & dat2 & dat3
                If datt = dat Then Exit Do
            Else
                b = InStr(b, dat, Chr$(34)) + 2
                If InStr(b, dat, Chr$(34)) = 0 Then Exit Do
                b = InStr(b, dat, Chr$(34)) + 2
                GoTo asdafae
            End If
ptafae:
        Loop Until datt = dat
        b = 1
        Do
            datt = dat
asdae:
            If InStr(b, dat, ";") = 0 Then
                GoTo ptae
            ElseIf Right$(Mid$(dat, 1, InStr(b, dat, ";")), 6) <> "\cf2 ;" Then
                dat1 = Mid$(dat, 1, InStr(b, dat, ";") - 1)
                dat2 = "\cf2 " & Mid$(dat, InStr(b, dat, ";"))
                dat3 = StrFrom(dat2, vbCrLf)
                dat2 = StrTo(dat2, vbCrLf) & vbCrLf
                dat2 = Replace$(dat2, "\cf0", "\cf2")
                dat2 = Replace$(dat2, "\cfo", "\cf0")
                dat2 = Replace$(dat2, "\cf1", "\cf2")
                dat2 = Replace$(dat2, "\cf3", "\cf2")
                dat = dat1 & dat2 & dat3
            Else
                b = InStr(b, dat, ";") + 1
                GoTo asdae
            End If
ptae:
        Loop Until datt = dat
            b = 1
            
        
        
        Do Until InStr(dat, ",  ") = 0
        dat = Replace$(dat, ",  ", ", ")
        Loop
        dat = Mid$(dat, 5)
        str = "{\rtf1\ansi\ansicpg1252\deff0\deflang1043{\fonttbl{\f0\fnil\fcharset0 Courier new;}}" & vbCrLf
        str = str & "{\colortbl ;\red0\green0\blue200;\red0\green150\blue0;\red150\green0\blue0;\red150\green0\blue150;}" & vbCrLf
        str = str & "\viewkind4\uc1\pard\cf0\f0\fs17"
        chcur = True
        Text1.TextRTF = str & dat
        Text1.SelStart = g
        Text1.SelLength = p
        chcur = False
    End If
End Sub


Private Sub Text1_SelChange()
If Not chcur Then
        Dim t As Long
        Dim p As String
        Dim c As String
        Dim a As Long
        t = Text1.SelStart
        c = Text1.Text
        p = Mid$(c, 1, t)
        For a = t To 1 Step -1
            If Mid$(c, a, 1) = vbCr Then Exit For
            DoEvents
        Next
        If a = 0 Then
            If InStr(c, vbCrLf) <> 0 Then c = StrTo(c, vbCrLf)
        Else
            c = Mid$(c, a)
            c = StrBetween(c, vbCrLf, vbCrLf)
        End If
    If Text1.SelText = "" Then
        StatusBar1.SimpleText = StrSetLength(cntin(p, vbCrLf), 6, "0", 1) & ": " & info(c)
    Else
        StatusBar1.SimpleText = StrSetLength(cntin(p, vbCrLf), 6, "0", 1) & ": " & ginfo(Text1.SelText)
    End If
End If
End Sub

Public Function cntin(Str1 As String, Str2 As String)
    On Error GoTo ers
    Dim b As Long
    Dim g As Long
    b = 1
    Do Until InStr(b, Str1, Str2) = 0
        b = InStr(b, Str1, Str2) + Len(Str2)
        g = g + 1
    Loop
ers:
    cntin = g + 1
End Function

Private Sub th_Click()
    On Error Resume Next
    Text1.SelText = Hex(Val(Text1.SelText))
End Sub

Private Sub transhex_Click()
    On Error Resume Next
    Dim rndlng As Long
    Dim rndstr As String
    Dim ff As Long
    Dim g As Long
    Dim dat As String
    Dim ret As Long
    Dim str As String
    StatusBar1.SimpleText = "Please wait..."
    str = Common.FileName
    Common.FileName = ""
    Randomize Timer
    rndlng = Int(Rnd * 50000) + 1
    rndstr = "build" & rndlng & ".asm"
    ff = FreeFile
    Open AppPath & rndstr For Binary As ff
        Put ff, , Text1.Text
    Close ff
    ff = 0
    ff = FreeFile
    Open AppPath & rndlng & ".bat" For Binary As ff
        dat = "cd " & AppPath & vbCrLf
        dat = dat & "@ECHO OFF" & vbCrLf
        If bnasm.Checked Then
            dat = dat & "nasmw.exe " & rndstr & " -o " & "build" & rndlng & ".bin" & vbCrLf
        Else
            dat = dat & "fasm.exe " & rndstr & vbCrLf
        End If
        dat = dat & "ECHO Done>don" & rndlng & ".txt" & vbCrLf
        'dat = dat & "pause"
        Put ff, , dat
    Close ff
    ff = 0
    If sds.Checked Then
        Shell AppPath & rndlng & ".bat", vbNormalFocus
    Else
        Shell AppPath & rndlng & ".bat", vbHide
    End If
    Do
        DoEvents
        g = Timer
        ff = FreeFile
        Open AppPath & "don" & rndlng & ".txt" For Binary As ff
            dat = Space$(LOF(ff))
            Get ff, , dat
        Close ff
        Do: DoEvents: Loop Until Timer >= g + 0.25
    Loop Until dat <> ""
        ff = FreeFile
        Open AppPath & "build" & rndlng & ".bin" For Binary As ff
            dat = Space$(LOF(ff))
            Get ff, , dat
        Close ff
        If dat = "" Then
            ret = MsgBox("One or more errors occurred!" & vbCrLf & "Do you want to debug for errors?", vbYesNo)
            If ret = vbYes Then
                Kill AppPath & rndlng & ".bat"
                ff = FreeFile
                Open AppPath & (rndlng + 1) & ".bat" For Binary As ff
                    dat = "cd " & AppPath & vbCrLf
                    dat = dat & "@ECHO OFF" & vbCrLf
                    If bnasm.Checked Then
                        dat = dat & "nasmw.exe " & rndstr & vbCrLf
                    Else
                        dat = dat & "fasm.exe " & rndstr & vbCrLf
                    End If
                    dat = dat & "pause" & vbCrLf
                    dat = dat & "del " & (rndlng + 1) & ".bat" & vbCrLf
                    Put ff, , dat
                Close ff
                Shell AppPath & (rndlng + 1) & ".bat", vbNormalFocus
            Else
            End If
        Else
            Dialog.Text1.Text = StrToHex(dat, "")
            If Len(Dialog.Text1.Text) <= 900 Then
                Dialog.Text2.Text = "Const myconst = " & Chr$(34) & Dialog.Text1.Text & Chr$(34)
            Else
                Dim a As Long
                Dialog.Text2.Text = "Const myconst = " & Chr$(34) & Mid$(Dialog.Text1.Text, a, 900) & Chr$(34) & " & _" & vbCrLf
                For a = 901 To Len(Dialog.Text1) Step 900
                    Dialog.Text2.Text = Dialog.Text2.Text & vbTab & vbTab & vbTab & vbTab & Chr$(34) & Mid$(Dialog.Text1.Text, 1, 900) & Chr$(34) & " & _" & vbCrLf
                Next
                Dialog.Text2.Text = Left$(Dialog.Text2.Text, Len(Dialog.Text2.Text) - 6)
            End If
            Dialog.DialogOpen Form1
            Kill AppPath & rndstr
        End If
        Kill AppPath & rndlng & ".bat"
        Kill AppPath & "don" & rndlng & ".txt"
        Kill AppPath & "build" & rndlng & ".bin"
        ff = FreeFile
        rndlng = 0
        StatusBar1.SimpleText = "Stopped compiling."
End Sub

Public Function info(Txt As String)
    On Error Resume Next
    Dim t() As String
    Dim Str1 As String
    Dim Str2 As String
    Dim p As Long
    Dim cmp As Boolean
    Dim cm As Boolean
    p = -1
    If Left$(Txt, 4) = "MOV " Then info = "V1 = V2": p = 3: cm = True
    If Left$(Txt, 4) = "INT " Then info = "Interrupt V1": p = 3
    If Left$(Txt, 4) = "SUB " Then info = "V1 = V1 - V2": p = 3: cm = True
    If Left$(Txt, 4) = "ADD " Then info = "V1 = V1 + V2": p = 3: cm = True
    If Left$(Txt, 4) = "INC " Then info = "V1 = V1 + 1": p = 3: cm = True
    If Left$(Txt, 4) = "AND " Then info = "V1 = V1 And V2": p = 3: cm = True
    If Left$(Txt, 4) = "NOT " Then info = "V1 = Not V1": p = 3: cm = True
    If Left$(Txt, 3) = "OR " Then info = "V1 = V1 Or V2": p = 2: cm = True
    If Left$(Txt, 4) = "XOR " Then info = "V1 = V1 Xor V2": p = 3: cm = True
    
    If Left$(Txt, 4) = "DEC " Then info = "V1 = V1 - 1": p = 3: cm = True
    If Left$(Txt, 4) = "RET " Then info = "Return V1": p = 3
    If Left$(Txt, 5) = "PUSH " Then info = "Push stack into V1.": p = 4: cm = True
    If Left$(Txt, 4) = "POP " Then info = "Pop V1 back into stack.": p = 3: cm = True
    If Left$(Txt, 4) = "CMP " Then info = "Compare V1 and V2.": p = 3: cmp = True
    If Left$(Txt, 5) = "TEST " Then info = "Test V1 and V2.": p = 4: cmp = True
    If Left$(Txt, 8) = "%DEFINE " Then info = "Define variable."
    If Left$(Txt, 9) = "[BITS 32]" Then info = "Indicate that registers are 32 bits."
    If Left$(Txt, 9) = "[BITS 16]" Then info = "Indicate that registers are 16 bits."
    If Left$(Txt, 8) = "[BITS 8]" Then info = "Indicate that registers are 8 bits."
    If Left$(Txt, 4) = "JMP " Then info = "Goto V1": p = 3
    
    If Left$(Txt, 3) = "JNL" Then info = "If " & last1 & " >= " & last2 & " Then goto V1": p = 3
    If Left$(Txt, 3) = "JNG" Then info = "If " & last1 & " <= " & last2 & " Then goto V1": p = 3

    If Left$(Txt, 4) = "JAE " Then info = "If " & last1 & " <= " & last2 & " Then goto V1": p = 3
    If Left$(Txt, 3) = "JA " Then info = "?If " & last1 & " < " & last2 & " Then goto V1": p = 2
    If Left$(Txt, 3) = "JB " Then info = "?If " & last1 & " > " & last2 & " Then goto V1": p = 2
    If Left$(Txt, 4) = "JBE " Then info = "?If " & last1 & " >= " & last2 & " Then goto V1": p = 3
    If Left$(Txt, 3) = "JE " Then info = "If " & last1 & " = " & last2 & " Then goto V1": p = 2
    If Left$(Txt, 4) = "JNE " Then info = "If " & last1 & " <> " & last2 & " Then goto V1": p = 3
    If Left$(Txt, 3) = "JG " Then info = "If " & last1 & " > " & last2 & " Then goto V1": p = 2
    If Left$(Txt, 4) = "JGE " Then info = "If " & last1 & " >= " & last2 & " Then goto V1": p = 3
    If Left$(Txt, 3) = "JL " Then info = "If " & last1 & " < " & last2 & " Then goto V1": p = 2
    If Left$(Txt, 4) = "JLE " Then info = "If " & last1 & " <= " & last2 & " Then goto V1": p = 3
    If Left$(Txt, 4) = "JNC " Then info = "If " & last1 & " <> Carry Then goto V1": p = 3
    If Left$(Txt, 4) = "JNO " Then info = "If " & last1 & " <> Overflow Then goto V1": p = 3
    If Left$(Txt, 4) = "JNP " Then info = "If " & last1 & " <> Parity Then goto V1": p = 3
    If Left$(Txt, 4) = "JNS " Then info = "If " & last1 & " <> 0" & " Then goto V1": p = 3
    If Left$(Txt, 3) = "JC " Then info = "If " & last1 & " = Carry Then goto V1": p = 2
    If Left$(Txt, 3) = "JP " Then info = "If " & last1 & " = Parity Then goto V1": p = 2
    If Left$(Txt, 3) = "JS " Then info = "If " & last1 & " = Signed Then goto V1": p = 2
    If Left$(Txt, 3) = "JZ " Then info = "If " & last1 & " = 0 Then goto V1": p = 2
    If Left$(Txt, 4) = "JNZ " Then info = "If " & last1 & " <> 0 Then goto V1": p = 3
    If Left$(Txt, 4) = "NEG " Then info = "V1 = 0 - V1": p = 3: cm = True
    
    If Left$(Txt, 4) = "DIV " Then info = "edx:eax = edx:eax / V1": p = 3: cm = True
    If Left$(Txt, 4) = "MUL " Then info = "edx:eax = eax * V1": p = 3: cm = True
    If Left$(Txt, 4) = "SBB " Then info = "V1 = V1 - V2": p = 3: cm = True
    If Left$(Txt, 4) = "SHL " Then info = "V1 = V1 * 2": p = 3: cm = True
    
    If Left$(Txt, 4) = "CLC " Then info = "CF = 0"
    If Left$(Txt, 4) = "CLD " Then info = "DF = 0"
    If Left$(Txt, 4) = "CLI " Then info = "IF = 0"
    If Left$(Txt, 4) = "CMC " Then info = "CF = Not CF"
    If Left$(Txt, 4) = "STC " Then info = "CF = 1"
    If Left$(Txt, 4) = "STD " Then info = "DF = 1"
    If Left$(Txt, 4) = "STI " Then info = "IF = 1"
    If p = -1 Then Exit Function
    t = Split(Mid$(Txt, p + 2), ",")
    Str1 = "V1"
    Str2 = "V2"
    'If Left$(t(0), 1) = " " Then t(0) = Mid$(t(0), 2)
    If InStr(t(0), ";") <> 0 Then t(0) = StrTo(t(0), ";")
    Do
    If Left$(t(0), 1) = " " Then t(0) = Mid$(t(0), 2)
    Loop Until Left$(t(0), 1) <> " "
    Do
    If Right$(t(0), 1) = " " Then t(0) = Left$(t(0), Len(t(0)) - 1)
    Loop Until Right$(t(0), 1) <> " "
    Str1 = t(0)
    If UBound(t) = 1 Then
        If InStr(t(1), ";") <> 0 Then t(1) = StrTo(t(1), ";")
        Do
        If Left$(t(1), 1) = " " Then t(1) = Mid$(t(1), 2)
        Loop Until Left$(t(1), 1) <> " "
        Do
        If Right$(t(1), 1) = " " Then t(1) = Left$(t(1), Len(t(1)) - 1)
        Loop Until Right$(t(1), 1) <> " "
        Str2 = t(1)
    End If
    If Str1 = "" Then Str1 = "V1"
    If Str2 = "" Then Str2 = "V2"
    If cmp Then last1 = Str1: last2 = Str2
    If cm Then last1 = Str1
    info = Replace$(info, "V1", Str1)
    info = Replace$(info, "V2", Str2)
End Function

Public Function ginfo(Txt As String)
    Dim c() As String
    Dim a As Long
    Txt = LCase$(Txt)
    Txt = Replace$(Txt, " ", "")
    If Txt = "eax" Then ginfo = "eax: Accumulator register"
    If Txt = "ebx" Then ginfo = "ebx: Base register"
    If Txt = "ecx" Then ginfo = "ecx: Counting register"
    If Txt = "edx" Then ginfo = "edx: Data register"
    
    If Txt = "ah" Then ginfo = "ah: Accumulator register, 2nd byte"
    If Txt = "bh" Then ginfo = "bh: Base register, 2nd byte"
    If Txt = "ch" Then ginfo = "ch Counting register, 2nd byte"
    If Txt = "dh" Then ginfo = "dh: Data register, 2nd byte"
    
    If Txt = "al" Then ginfo = "al: Accumulator register, 1st byte"
    If Txt = "bl" Then ginfo = "bl: Base register, 1st byte"
    If Txt = "cl" Then ginfo = "cl: Counting register, 1st byte"
    If Txt = "dl" Then ginfo = "dl: Data register, 1st byte"
    
    If Txt = "ds" Then ginfo = "ds: Data segment register"
    If Txt = "es" Then ginfo = "es: Extra segment register"
    If Txt = "ss" Then ginfo = "ss: Battery segment register"
    If Txt = "cs" Then ginfo = "cs: Code segment register"
    
    If Txt = "ebp" Then ginfo = "ebp: Base pointer register"
    If Txt = "esi" Then ginfo = "esi: Source index register"
    If Txt = "edi" Then ginfo = "edi: Destiny index register"
    If Txt = "esp" Then ginfo = "esp: Battery pointer register"
    c = Split(Text3.Text, vbCrLf)
    For a = 0 To UBound(c)
        If InStr(c(a), " ") <> 0 Then
            If Txt = LCase$(StrTo(c(a), " ")) Then ginfo = StrTo(c(a), " ") & ": " & StrFrom(c(a), " ")
        End If
    Next
End Function

Private Sub AddOpen(str As String)
    Dim a As Long
    For a = 3 To 0 Step -1
        Opens(a + 1) = Opens(a)
    Next
    Opens(0) = str
End Sub

Private Sub RemOpen(str As String)
    Dim a As Long
    Dim b As Long
    For a = 0 To 4
        If Opens(a) = str Then
            Opens(a) = ""
            For b = a To 3
                Opens(b) = Opens(b + 1)
            Next
            Opens(4) = ""
        End If
    Next
End Sub

Private Sub LoadOpen()
    Dim a As Long
    For a = 0 To 4
        prevopen(a).Caption = Opens(a)
        If Opens(a) = "" Then prevopen(a).Caption = "<empty>"
    Next
End Sub
