VERSION 5.00
Begin VB.Form Datempr 
   Caption         =   "Datos de la empresa:"
   ClientHeight    =   2676
   ClientLeft      =   96
   ClientTop       =   660
   ClientWidth     =   6432
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9.6
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Datempr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2676
   ScaleWidth      =   6432
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Pago Mensual"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4800
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   1800
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Fecha actualiz.  dd/mm/aaaa"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3120
      TabIndex        =   8
      Top             =   1140
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Subsidio aplicable :  "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Salario Minimo :  "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Año :   "
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Empresa :"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Menu EmpArch 
      Caption         =   "&Archivo"
      Index           =   0
      Begin VB.Menu EmprArch 
         Caption         =   "&Archivar"
         Index           =   1
      End
      Begin VB.Menu EmprSal 
         Caption         =   "&Salida"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Datempr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EmprArch_Click(Index As Integer)
   cm = cm + 1
   empresa.Quincenal = Check1.Value
   Put 1, cm, empresa
End Sub

Private Sub EmprSal_Click(Index As Integer)
   Unload Datempr
End Sub

Private Sub Form_Load()
   Close 1: Open "EMPRESA.NOM" For Random As #1 Len = Len(empresa)
    cm = LOF(1) / Len(empresa)
    
    If cm > 0 Then
            Get 1, cm, empresa
            Text1.Text = empresa.name
            Text2.Text = empresa.ao
            Text3.Text = Format(empresa.sm, "##,##0.00")
            Text4.Text = Format(empresa.psub, "0.0000")
            Text5.Text = Format(empresa.fecha, "dd/mm/yyyy")
            Check1.Value = empresa.Quincenal
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Select Case (KeyAscii)
            Case 27
             Text1.Text = empresa.name
             
            Case 13
             empresa.name = Text1.Text
             Text2.SetFocus
    End Select
End Sub

Private Sub Text1_LostFocus()
   Text1_KeyPress 13
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
     Select Case (KeyAscii)
            Case 27
             Text2.Text = empresa.ao
             
            Case 13
             empresa.ao = Text2.Text
             Text3.SetFocus
    End Select
End Sub

Private Sub Text2_LostFocus()
    Text2_KeyPress 13
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
 Select Case (KeyAscii)
            Case 27
             Text3.Text = Format(empresa.sm, "##,##0.00")
            Case 13
             empresa.sm = Text3.Text
             Text3.Text = Format(empresa.sm, "##,##0.00")
             Text4.SetFocus
    End Select

End Sub

Private Sub Text3_LostFocus()
  Text3_KeyPress 13
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
 Select Case (KeyAscii)
            Case 27
             Text4.Text = Format(empresa.psub, "0.0000")
             
            Case 13
            If Text4.Text > 1 Then
             empresa.psub = Text4.Text / 100
             Text4.Text = Format(Text4.Text, "0.0000")
             
             Else
             empresa.psub = Text4.Text
             Text4.Text = Format(Text4.Text, "0.0000")
             
            End If
            Text5.SetFocus
 End Select

End Sub

Private Sub Text4_LostFocus()
      Text4_KeyPress 13
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
       Select Case (KeyAscii)
        Case 27
             Text5.Text = Format(empresa.fecha, "dd/mm/yyyy")
             
            Case 13
            
             empresa.fecha = Text5.Text
             Text5.Text = Format(Text5.Text, "dd/mm/yyyy")
        End Select
            
End Sub

Private Sub Text5_LostFocus()
 Text5_KeyPress 13
End Sub
