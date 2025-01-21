VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "NOMINA CAPTURA DE DATOS DE PERSONAL"
   ClientHeight    =   7425
   ClientLeft      =   2370
   ClientTop       =   2910
   ClientWidth     =   7860
   ClipControls    =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7425
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text14 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      TabIndex        =   33
      Top             =   480
      Width           =   5775
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2640
      TabIndex        =   31
      Top             =   5265
      Width           =   1695
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1680
      TabIndex        =   30
      Top             =   5265
      Width           =   855
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5640
      TabIndex        =   27
      Top             =   4650
      Width           =   2055
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5640
      TabIndex        =   26
      Top             =   5265
      Width           =   2055
   End
   Begin VB.Frame frame2 
      Caption         =   "Anterior"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   23
      Top             =   5760
      Width           =   3495
      Begin VB.Label Label12 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   24
         Top             =   -1800
         Width           =   3255
      End
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   22
      Top             =   4650
      Width           =   1815
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4680
      TabIndex        =   20
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "SALIDA"
      Height          =   615
      Left            =   4080
      MaskColor       =   &H008080FF&
      TabIndex        =   18
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5640
      TabIndex        =   17
      Top             =   3810
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   16
      Top             =   3810
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5640
      TabIndex        =   15
      Top             =   2970
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   14
      Top             =   2970
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   13
      Top             =   2250
      Width           =   6015
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   12
      Top             =   1680
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   11
      Top             =   1005
      Width           =   6015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BORRAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   10
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ARCHIVAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   5880
      Width           =   3615
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "CURP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   32
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "Obra: num, Porcentaje (100%)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   29
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "REGISTRO IMSS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   28
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "SALARIO INTEGRADO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   25
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "OTROS  DIARIOS INGRESOS FIJOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   21
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "R.F.C."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3480
      TabIndex        =   19
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "INGRESO  DIARIO NORMAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "FECHA DE BAJA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "FECHA DE INGRESO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "APELLIDO MATERNO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   4
      Top             =   2250
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "APELLIDO PATERNO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "NOMBRE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "NUMERO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "VIATICOS DIARIOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   8
      Top             =   3720
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim toTingr

Private Sub Command1_Click()
On Error GoTo SalEdit
    integrar
    archivar
    Label12.Caption = Text1.Text + " " + Text2.Text + " " + Text3.Text
    If rgtro > 0 Then
         Put #2, rgtro, personal
         maestro.por_1 = Text13.Text
         Put #8, rgtro, maestro
         Put #3, rgtro, Otros_Rgtros
         borrar
         Close
         Unload Form2
         
         Else
         
         Put #2, dm, personal
         If Text13.Text = "" Then Text13.Text = 100: maestro.por_1 = Text13.Text
         Put #8, dm, maestro
         Put #3, dm, Otros_Rgtros
         dm = dm + 1
         Label2.Caption = dm
         borrar
     End If
SalEdit:
End Sub
Sub carga()
    Text1.Text = LTrim(personal.nom)
    Text2.Text = LTrim(personal.ape1)
    Text3.Text = LTrim(personal.ape2)
    Text8.Text = LTrim(personal.rfc)
    Text4.Text = LTrim(personal.fal)
    Text5.Text = LTrim(personal.fab)
    Text6.Text = Format(personal.ingr, z1)
    Text7.Text = Format(personal.viat, z1)
    Text9.Text = Format(personal.otras, z1)
    Text10.Text = Format(personal.integrado, z1)
    Text11.Text = personal.imss
    Text12.Text = maestro.O_1
    Text13.Text = maestro.por_1
    Text14.Text = Otros_Rgtros.curp

End Sub
Private Sub archivar()
    personal.nom = Text1.Text
    personal.ape1 = Text2.Text
    personal.ape2 = Text3.Text
    personal.rfc = Text8.Text
    personal.fal = Text4.Text
    personal.fab = Text5.Text
    Otros_Rgtros.curp = Text14.Text
    If Val(Text6.Text) = 0 Then
        personal.ingr = 0
        Else
           personal.ingr = Text6.Text
     End If
     If Val(Text7.Text) = 0 Then
        personal.viat = 0
        Else
           personal.viat = Text7.Text
     End If
     If Val(Text9.Text) = 0 Then
        personal.otras = 0
        Else
           personal.otras = Text9.Text
     End If
     
     If Val(Text10.Text) = 0 Then
        personal.integrado = 0
        Else
           personal.integrado = Text10.Text
     End If
        personal.imss = Text11.Text
        If Text12.Text = "" Then Text12.Text = 9000
        maestro.O_1 = Text12.Text
        
End Sub
Sub integrar()
    aoingr = Val(Mid$(Text4.Text, 7, 4))
    If aoingr < 1900 Then
         MsgBox "La fecha de ingreso no es correcta, no es posible calcular el salario integrado"
         Else
         antig = empresa.ao + 1 - aoingr
         facto = 0: toTingr = 0
         If IsNumeric(Text6.Text) Then toTingr = toTingr + Text6.Text
         If IsNumeric(Text7.Text) Then toTingr = toTingr + Text7.Text
         If IsNumeric(Text9.Text) Then toTingr = toTingr + Text9.Text
         factor antig, facto
         Text10.Text = Format((toTingr * facto), z1$)
    End If
End Sub

Private Sub Command2_Click()
   Rem Print "BORRAR"
   Label12.Caption = Text1.Text + " " + Text2.Text + " " + Text3.Text
    If rgtro > 0 Then
         Put #2, rgtro, personal
         Put #8, rgtro, obra
         borrar
         Close
         Unload Form2
         Form4.Show
         Else: borrar
    End If
End Sub

Private Sub Command3_Click()
    Close
    Unload Form2
    If rgtro > 0 Then
          Rem Form4.Show
          Else:
          Load Form1
          Form1.Show
     End If
End Sub
   Private Sub borrar()
      Text1.Text = ""
      Text2.Text = ""
      Text3.Text = ""
      Text4.Text = ""
      Text5.Text = ""
      Text6.Text = ""
      Text7.Text = ""
      Text8.Text = ""
      Text9.Text = ""
      Text10.Text = ""
      Text11.Text = ""
      Text12.Text = ""
      Text13.Text = ""
    End Sub

Private Sub Form_Load()
  Open "personal.dno" For Random As 2 Len = Len(personal)
  dm = LOF(2) / Len(personal)
  Close 3
  Open "PerOtre.dno" For Random As 3 Len = Len(Otros_Rgtros)
  dmper = LOF(3) / Len(personal)

  Close 8
  Open "maestro.dno" For Random As 8 Len = Len(maestro)
  ddm = LOF(8) / Len(maestro)
  If rgtro > 0 Then
       Get #2, rgtro, personal
       Get #8, rgtro, maestro
       Get #3, rgtro, Otros_Rgtros
       carga
       Label2.Caption = Format(rgtro, "           #####.-")
       Else
        If dm > 0 Then
            Get #2, dm, personal
            Get #8, dm, maestro
            Label12.Caption = RTrim$(personal.nom) + " " + RTrim$(personal.ape1) + " " + RTrim$(personal.ape2)
        End If
        dm = dm + 1
        Label2.Caption = Format(dm, "           #####.-")
   End If
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
          Text1.Text = LTrim$(Text1.Text)
          Text14.SetFocus
    End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
          Text11.Text = LTrim$(Text11.Text)
          Command1.SetFocus
End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
          Text14.Text = UCase(LTrim(Text14.Text))
          Text1.SetFocus
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
          Text2.Text = LTrim$(Text2.Text)
          Text3.SetFocus
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
          Text3.Text = LTrim$(Text3.Text)
          Text4.SetFocus
    End If

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
          Text4.Text = LTrim$(Text4.Text)
          aoingr = Val(Mid$(Text4.Text, 7, 4))
    If aoingr < 1900 Then
         MsgBox "La fecha de ingreso no es correcta,el formato es dd/mm/aaaa"
    End If
         Text6.SetFocus
    End If

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         Text5.Text = LTrim$(Text5.Text)
         aobaj = Val(Mid$(Text5.Text, 7, 4))
    If aobaj < 1900 Then
         MsgBox "La fecha de baja no es correcta,el formato es dd/mm/aaaa"
    End If
         Text6.SetFocus
    End If

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
          Text6.Text = Format(Text6.Text, "         ###,##0.0000")
          integrar
          Text7.SetFocus
    End If

End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
             
          Text7.Text = Format(Text7.Text, "         ###,##0.0000")
          integrar
          Text9.SetFocus
    End If
End Sub


Private Sub Text8_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
          Text8.Text = LTrim$(Text8.Text)
          Text1.SetFocus
      End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
          Text9.Text = Format(Text9.Text, "         ###,##0.0000")
          integrar
          Text11.SetFocus
        End If
End Sub

