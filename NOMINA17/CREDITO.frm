VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form7 
   Caption         =   "CREDITO AL SALARIO MENSUAL"
   ClientHeight    =   5475
   ClientLeft      =   1185
   ClientTop       =   2565
   ClientWidth     =   7245
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   7245
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid connom1 
      Height          =   3495
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6165
      _Version        =   393216
      Rows            =   15
      Cols            =   4
   End
   Begin VB.Menu archcre 
      Caption         =   "&Archivo"
      Index           =   0
      Begin VB.Menu ArGdr1 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu ArSep2 
         Caption         =   "-"
      End
      Begin VB.Menu creimp 
         Caption         =   "&Imprimir"
         Index           =   2
      End
      Begin VB.Menu ArSep1 
         Caption         =   "-"
      End
      Begin VB.Menu salcre 
         Caption         =   "&Salida"
         Index           =   1
      End
   End
   Begin VB.Menu Ed 
      Caption         =   "&Edicion"
      Begin VB.Menu EdTd 
         Caption         =   "&Seleccionar Todo"
      End
      Begin VB.Menu EdSep1 
         Caption         =   "-"
      End
      Begin VB.Menu EdCop 
         Caption         =   "&Copiar"
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ArGdr1_Click()
   yavas = 0
    For I = 1 To 14
       connom1.Row = I
       connom1.Col = 1
       If Val(connom1.Text) > 0 Then
             credito.crede = connom1.Text
             yavas = 1
             Else
             subsidio.liminfs = 0
       End If
       connom1.Col = 2
       If Val(connom1.Text) > 0 Then
            credito.crea = connom1.Text
            yavas = 1
            Else
            subsidio.liminfs = 0
       End If
       connom1.Col = 3
       
       If Val(connom1.Text) > 0 Then
            credito.cresam = connom1.Text
            
            yavas = 1
            Else
            subsidio.cuotafs = 0
       End If
       If yavas > 0 Then
             Put #5, I, credito: yavas = 0
             Else: I = 14
       End If
     Next I
End Sub

Private Sub creimp_Click(Index As Integer)
   tam_largo = Printer.Height
    tam_ancho = Printer.Width
    
    Printer.Font.Size = 12
    Printer.CurrentY = tam_largo * 0.33
    Rem Printer.Print Tab(40); "Credito al Salario."
    Printer.Print Tab(30); Form7.Caption
    Printer.Print
    guarda = Printer.CurrentY
    Printer.Line ((2500 + 1200), Printer.CurrentY)-((2500 + 5000), (Printer.CurrentY + 600)), , B
    Printer.Font.Size = 8
    Printer.CurrentY = guarda
    Printer.Print Tab(50); "Lim.Inf";
    Printer.Print Tab(65); "Lim.Sup";
    Printer.Print Tab(80); "Credito"
    Printer.Print
    Printer.Print
    For r = 1 To 14
      For l = 1 To 3
         If connom1.TextMatrix(r, l) <> "" Then
                bala = connom1.TextMatrix(r, l)
                valor$ = Format(bala, "##,###,##0.00"): uso$ = "##,###,##0.00"
                pone = 0: colocar pone, valor$, uso$
                Printer.CurrentX = (2500 + (1200 * l)) + pone
                Printer.Print valor$;
         End If
      Next l
      Printer.Print
    Next r
    Printer.EndDoc
End Sub

Private Sub Form_Load()
    Width = Screen.Width * 0.65 ' Establecer el ancho del formulario.
    Height = Screen.Height * 0.78   ' Establecer el alto del formulario.
    Left = (Screen.Width - Width) / 2   ' Centrar el formulario horizontalmente.
    Top = (Screen.Height - Height) / 2  ' Centrar el formulario verticalmente.

If ta_r = 0 Then
      Close 5:  Open SubT_Mes + "TABLA.CRE" For Random As #5 Len = Len(credito)
      
   Else
        Open arch_tr For Random As #5 Len = Len(credito)
        ta_r = 0
End If
Form7.Caption = SubT_Mes + "Tarifa articulo Credito al salario"
cem = LOF(5) / Len(credito)
connom1.ColWidth(0) = 330
connom1.Col = 0
connom1.Row = 0
connom1.Text = "#"
connom1.CellAlignment = 4
connom1.Col = 1
connom1.CellAlignment = 4
connom1.Text = "Lim.Inferior"
connom1.ColWidth(1) = 1500
connom1.Col = 2
connom1.CellAlignment = 4
connom1.Text = "Lim.Superior"
connom1.ColWidth(2) = 1500
connom1.Col = 3
connom1.CellAlignment = 4
connom1.Text = "Cuota Fija"
connom1.ColWidth(3) = 1500
If cem > 0 Then
    For r = 1 To cem: Get #5, r, credito
      connom1.Col = 0
      connom1.Row = r
      connom1.Text = r
      connom1.Col = 1
      connom1.CellAlignment = 7
      connom1.Text = Format(credito.crede, "###,##0.00")
      connom1.Col = 2
      connom1.CellAlignment = 7
      connom1.Text = Format(credito.crea, "###,##0.00")
      connom1.Col = 3
      connom1.CellAlignment = 7
      connom1.Text = Format(credito.cresam, "###,##0.00")

    Next r
      Else
        For r = 1 To 14
           connom1.Col = 0
           connom1.Row = r
           connom1.Text = r
        Next r
    End If
     connom1.Row = 1
     connom1.Col = 1
     connom1.Appearance = flex3D

End Sub



Private Sub Form_Resize()
    Form7.Width = connom1.Width + 1400
    Form7.Height = connom1.Height + 2400

End Sub

Private Sub salcre_Click(Index As Integer)
    
     Close:
     Unload Form7
     Rem Load Form1
     Rem Form1.Show
End Sub
Private Sub ConNom1_EnterCell()
    If connom1.Col > 0 And connom1.Row > 0 Then
        connom1.CellBackColor = &H80FF80
    End If
End Sub


Private Sub ConNom1_KeyPress(KeyAscii As Integer)
            Text1.Text = Chr$(KeyAscii)
            Text1.SetFocus
End Sub

Private Sub ConNom1_LeaveCell()
   If connom1.Col > 0 And connom1.Row > 0 Then
        connom1.CellBackColor = vbWhite
   End If
End Sub

Private Sub ConNom1_RowColChange()
   Text1.Text = connom1.Text
End Sub
Private Sub Text1_Change()
    connom1.Text = Text1.Text
End Sub


Private Sub Text1_GotFocus()
     SendKeys "{end}"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
       Case 13
         connom1.Text = Format(Val(connom1.Text), "###,##0.00")
         connom1.SetFocus
       Case 27
         connom1.Text = ""
         connom1.SetFocus
   End Select
End Sub


