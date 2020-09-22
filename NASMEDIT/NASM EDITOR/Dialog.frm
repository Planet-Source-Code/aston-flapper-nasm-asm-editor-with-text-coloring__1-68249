VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   1365
      Left            =   45
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1485
      Width           =   4830
   End
   Begin VB.TextBox Text1 
      Height          =   1365
      Left            =   45
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   45
      Width           =   4830
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RetVal As String
Dim SetVal As Boolean
Dim Formx As Form

Public Function DialogOpen(Formy As Form) As String: Set Formx = Formy: Formx.Enabled = False: Me.Enabled = True: Me.Show: Me.SetFocus: SetVal = False: Do: DoEvents: Loop Until Me.Visible = False: DialogOpen = RetVal: End Function
Private Sub ExitOk(Txt As String): SetVal = True: RetVal = Txt: Unload Me: Formx.Enabled = True: End Sub
Private Sub Command1_Click(): SetVal = False: Unload Me: End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer): If Not SetVal Then RetVal = "": Cancel = True: Formx.Enabled = True: Me.Visible = False: Formx.SetFocus
End Sub

Private Sub Command2_Click()
ExitOk "Text!"
End Sub

