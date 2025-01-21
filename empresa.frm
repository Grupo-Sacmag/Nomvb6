VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Nomina: Datos de la empresa"
   ClientHeight    =   5475
   ClientLeft      =   2070
   ClientTop       =   2130
   ClientWidth     =   7245
   LinkTopic       =   "Form3"
   ScaleHeight     =   5475
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   9
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Archivar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   6615
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   6
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   4
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4560
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6735
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Salario MInimo Vigente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   10
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha del ultimo cambio: dd/mm/aaaa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   7
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "UMA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   5
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Año"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Empresa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Menu empr 
      Caption         =   "&Empresa"
      Index           =   0
      Begin VB.Menu empsal 
         Caption         =   "&Salida"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error GoTo manejadorError
       empresa.name = Text1.Text
       empresa.ao = Text2.Text
       empresa.sm = Text3.Text
       empresa.fecha = Text4.Text
       empresa.psub = Text5.Text
       
       If cm = 0 Then
          Open "EMPRESA.dno" For Random As 1 Len = Len(empresa)
       End If
       
       cm = cm + 1
       Put 1, cm, empresa
    
    Form1.Label2.Caption = ("Año : " + Str$(empresa.ao) + Chr(13) + " Salario Minimo $" + Str$(empresa.psub) + Chr(13) + "UMA x dia :$ " + Format(empresa.sm, z1$))
    
    Exit Sub

manejadorError:
    
End Sub

Private Sub empsal_Click(Index As Integer)
        Unload Form3
        Load Form1
        Form1.Show
End Sub

Private Sub Form_Load()
    If cm > 0 Then
       Get 1, cm, empresa
       Text1.Text = empresa.name
       Text2.Text = empresa.ao
       Text3.Text = empresa.sm
       Text4.Text = empresa.fecha
       Text5.Text = empresa.psub
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        mientras = Text3.Text
        If mientras > 0 Then
              Text3.Text = Format(mientras, z1$)
        End If
  End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
        mientras = Text5.Text
        If mientras > 0 Then
              Text5.Text = Format(mientras, "#0.0000")
        End If
  End If

End Sub


