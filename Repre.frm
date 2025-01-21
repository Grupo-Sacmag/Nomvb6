VERSION 5.00
Begin VB.Form Repre 
   Caption         =   "Datos Complementarios empresa y Representante"
   ClientHeight    =   4695
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   2160
      TabIndex        =   18
      Top             =   120
      Width           =   4215
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Pago 2da.quincena 16 dias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   15
      Top             =   2400
      Width           =   3255
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3120
      TabIndex        =   14
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5640
      TabIndex        =   9
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   1200
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label Label7 
      Caption         =   "Registro Patronal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "No.Cta.Banamex Sucursal ,  No. y Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   12
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Curp Rep.Legal : "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Rfc Rep.Legal : "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Representate Legal A. Paterno,  Mat. y Nom."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Dirección : "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Rfc Empresa : "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
   Begin VB.Menu repAr 
      Caption         =   "&Archivo"
      Begin VB.Menu Argua 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu RepSal 
         Caption         =   "&Salida"
      End
   End
End
Attribute VB_Name = "Repre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub carga()
   Text1.Text = Dat_ide.Emp_Rfc
   Text2.Text = Dat_ide.Emp_Dom
   Text3.Text = Dat_ide.Rep_Legapm
   Text4.Text = Dat_ide.Rep_Legapn
   End Sub


Private Sub Argua_Click()
 salvar
End Sub

Private Sub Check1_Click()
   If Check1 = 0 Then
        
        Dat_ide.dias = 0
        Else
        
        Dat_ide.dias = 1
   End If
End Sub

Private Sub Form_Load()
  Repre.Caption = "Datos complementarios y representate de " + empresa.name
  recuperar
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Unload Repre
End Sub
Sub salvar()
   Open "Empcomp.dno" For Random As 11 Len = Len(Dat_ide)
   
    Dat_ide.Emp_Rfc = Text1.Text
    Dat_ide.Emp_Dom = Text2.Text
    Dat_ide.Rep_Legapp = Text3.Text
    Dat_ide.Rep_Legapm = Text4.Text
    Dat_ide.Rep_Legapn = Text5.Text
    Dat_ide.Rep_Rfc = Text6.Text
    Dat_ide.Rep_Curp = Text7.Text
    Dat_ide.suc = Text8.Text
    Dat_ide.cta = Text9.Text
    Dat_ide.clte = Text10.Text
    EmpCFDI.RegPatr = Text11.Text
   
   Put 11, 1, Dat_ide
   Put 10, 1, EmpCFDI
   Close 11

End Sub
Sub recuperar()
    Close 13, 10
   Open "EMP_CFDI.DNO" For Random As 10 Len = Len(EmpCFDI)
   Open "Empcomp.dno" For Random As 13 Len = Len(Dat_ide)
   Get 13, 1, Dat_ide
   Get 10, 1, EmpCFDI
   
   Text1.Text = Dat_ide.Emp_Rfc
   Text2.Text = Dat_ide.Emp_Dom
   Text3.Text = Dat_ide.Rep_Legapp
   Text4.Text = Dat_ide.Rep_Legapm
   Text5.Text = Dat_ide.Rep_Legapn
   Text6.Text = Dat_ide.Rep_Rfc
   Text7.Text = Dat_ide.Rep_Curp
   Text8.Text = Dat_ide.suc
   Text9.Text = Dat_ide.cta
   Text10.Text = Dat_ide.clte
   Text11.Text = EmpCFDI.RegPatr
   If Dat_ide.dias = 0 Then
        Check1 = 0
        Else
        Check1 = 1
   End If
   
   
   Close 13
End Sub

Private Sub RepSal_Click()
    Form_Unload 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text1.BackColor = vbYellow
        Text1.Text = UCase(Text1.Text)
    End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text11.BackColor = vbYellow
        Text11.Text = UCase(Text11.Text)
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text2.BackColor = vbYellow
        Text2.Text = UCase(Text2.Text)
    End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text3.BackColor = vbYellow
        Text3.Text = UCase(Text3.Text)
    End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text4.BackColor = vbYellow
        Text4.Text = UCase(Text4.Text)
    End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text5.BackColor = vbYellow
        Text5.Text = UCase(Text5.Text)
    End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text6.BackColor = vbYellow
        Text6.Text = UCase(Text6.Text)
    End If
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text7.BackColor = vbYellow
        Text7.Text = UCase(Text7.Text)
    End If
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text8.BackColor = vbYellow
        Text8.Text = UCase(Text8.Text)
    End If
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text9.BackColor = vbYellow
        Text9.Text = UCase(Text9.Text)
    End If
End Sub

