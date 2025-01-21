VERSION 5.00
Begin VB.Form CapturaDatos 
   Caption         =   "Form1"
   ClientHeight    =   6195
   ClientLeft      =   90
   ClientTop       =   480
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Examinar 
      Caption         =   "Command1"
      Height          =   375
      Index           =   0
      Left            =   5160
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "CapturaDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
        
Private Sub Form_Load()
   Dim W
     Label1(0).Move 50, 700, 1700, 700
     Label1(0).Alignment = 1
     Label1(0).FontBold = True
     Text1(0).Move 2200, 700, 2700, 300
     Text1(0).Text = ""
     Examinar(0).Move 5000, 700, 1700, 300
     For W = 1 To 1
         Load Examinar(W)
         Examinar(W).Move 5000, (W + 1) * 700, 1700, 300
         Examinar(W).Visible = True
     Next W
        For W = 1 To Tot_Dat
        Load Label1(W)
        Load Text1(W)
        Label1(W).Move 50, (W + 1) * 700, 1700, 700
        Text1(W).Move 2200, (W + 1) * 700, 2700, 300
        Label1(W).Visible = True
        Text1(W).Visible = True
        Text1(W).Text = ""
        Label1(W).Alignment = 1
        Label1(W).FontBold = True
     Next W
    
End Sub
Private Sub Examinar_Click(Index As Integer)
   Select Case Index
      Case 0
      DirecT_arifas = UCase(RTrim(Text1(Index).Text))
      Case 1
        Close
        Unload CapturaDatos
   End Select
End Sub

Private Sub Form_Resize()
    CapturaDatos.Height = (700 * Tot_Dat) + 3000
End Sub
