VERSION 5.00
Begin VB.Form sudio 
   Caption         =   "Captura subsidios aplicables en el 2001"
   ClientHeight    =   2700
   ClientLeft      =   195
   ClientTop       =   765
   ClientWidth     =   4875
   Icon            =   "subdio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3480
      TabIndex        =   9
      Top             =   2145
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3480
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3480
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3480
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "R.F.C. de la Empresa"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Aplicado en el año 2000"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Aplicable Por el año de 2001"
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Aplicable Noviembre Diciembre"
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Aplicable Enero-Octubre"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Menu archsub 
      Caption         =   "&Archivo"
      Begin VB.Menu GuarSub 
         Caption         =   "&Guardar"
         Shortcut        =   ^G
      End
      Begin VB.Menu sepsub1 
         Caption         =   "-"
      End
      Begin VB.Menu SalSub 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "sudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
     Open "Datsub.nom" For Random As 13 Len = Len(sub_dio)
     finsub = LOF(13) / Len(sub_dio)
     If finsub > 0 Then
        Text1.Text = sub_dio.s1
        Text2.Text = sub_dio.s2
        Text3.Text = sub_dio.s3
        Text4.Text = sub_dio.s0
        Text5.Text = sub_dio.rfc
     End If
     
End Sub


Private Sub Form_Unload(Cancel As Integer)
   
   Close 13
End Sub

Private Sub GuarSub_Click()
   Put 13, 1, sub_dio
   psub1 = sub_dio.s1
   psub2 = sub_dio.s2
   psub3 = sub_dio.s3
   psub0 = sub_dio.s0

End Sub

Private Sub SalSub_Click()
   Close 13
   Unload sudio
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   
   If Text1.Text <> "" And KeyAscii = 13 Then
       sub_dio.s1 = Text1.Text
       If sub_dio.s1 > 1 Then sub_dio.s1 = sub_dio.s1 / 100
       Text1.Text = Format(sub_dio.s1, " ##0.0000")
   End If
End Sub
Private Sub Text1_LostFocus()
     Text1_KeyPress 13
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
  If Text2.Text <> "" And KeyAscii = 13 Then
       sub_dio.s2 = Text2.Text
       If sub_dio.s2 > 1 Then sub_dio.s2 = sub_dio.s2 / 100
       Text2.Text = Format(sub_dio.s2, " ##0.0000")
   End If
End Sub
Private Sub Text2_LostFocus()
     Text2_KeyPress 13
End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
     If Text3.Text <> "" And KeyAscii = 13 Then
       sub_dio.s3 = Text3.Text
       If sub_dio.s3 > 1 Then sub_dio.s3 = sub_dio.s3 / 100
       Text3.Text = Format(sub_dio.s3, " ##0.0000")
    End If

End Sub

Private Sub Text3_LostFocus()
     Text3_KeyPress 13
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   If Text4.Text <> "" And KeyAscii = 13 Then
    
       sub_dio.s0 = Text4.Text
       If sub_dio.s0 > 1 Then sub_dio.s0 = sub_dio.s0 / 100
       Text4.Text = Format(sub_dio.s0, " ##0.0000")
    End If

End Sub

Private Sub Text4_LostFocus()
     Text4_KeyPress 13
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
    If Text5.Text <> "" And KeyAscii = 13 Then
       sub_dio.rfc = UCase(Text5.Text)
       Text5.Text = sub_dio.rfc
    End If

End Sub

Private Sub Text5_LostFocus()
     Text5_KeyPress 13
End Sub

