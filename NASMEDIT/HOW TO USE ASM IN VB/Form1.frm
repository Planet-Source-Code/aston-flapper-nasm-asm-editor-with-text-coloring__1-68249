VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   510
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Top             =   135
      Width           =   1905
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   285
      Left            =   1530
      TabIndex        =   1
      Top             =   135
      Width           =   825
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Text            =   "5"
      Top             =   135
      Width           =   1230
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
                'Compiled ASM code in hex \/
Const MyConst = "5589E55356578B5D08B801000000B90100000081C10100000001C839D975F45F5E5B89EC5DC21000"

'How to compile ASM code into hex
'It is very easy with my nAsm editor
'Go to menu Build->Translate to hexdata
'Copy the bottom text, that is an precreated Const
'If you have different methods of using ASM with vb copy the upper text


Private Sub Command1_Click()
    Text2.Text = asmRun(CLng(Text1.Text), 0, 0, 0)
    'Run your asm code through VB!
End Sub

Private Sub Form_Load()
    asmInject MyConst   'Inject your compiled asm code into
                        'this program
End Sub
