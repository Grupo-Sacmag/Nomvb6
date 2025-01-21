VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form6 
   Caption         =   "Tabla del subsidio"
   ClientHeight    =   5475
   ClientLeft      =   1770
   ClientTop       =   2415
   ClientWidth     =   7245
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   7245
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin MSFlexGridLib.MSFlexGrid ConNom1 
      Height          =   3615
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   6376
      _Version        =   393216
      Rows            =   15
      Cols            =   5
   End
   Begin VB.Menu arcsub 
      Caption         =   "&Archivo"
      Index           =   0
      Begin VB.Menu ArGdr 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu Arsep1 
         Caption         =   "-"
      End
      Begin VB.Menu subimp 
         Caption         =   "&Imprimir"
         Index           =   2
      End
      Begin VB.Menu ArSep2 
         Caption         =   "-"
      End
      Begin VB.Menu subsal 
         Caption         =   "&Salida"
         Index           =   1
      End
   End
   Begin VB.Menu Ed 
      Caption         =   "&Edicion"
      Begin VB.Menu EdTd 
         Caption         =   "&Seleccionar Todo"
      End
      Begin VB.Menu Sep 
         Caption         =   "-"
      End
      Begin VB.Menu EdCop 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu EdSep2 
         Caption         =   "-"
      End
      Begin VB.Menu EdPg 
         Caption         =   "&Pegar"
         Shortcut        =   ^V
      End
      Begin VB.Menu EdSep4 
         Caption         =   "-"
      End
      Begin VB.Menu EdEl 
         Caption         =   "&Eliminar"
      End
      Begin VB.Menu EdSep3 
         Caption         =   "-"
      End
      Begin VB.Menu EdIns 
         Caption         =   "&Insertar"
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ArGdr_Click()
   yavas = 0
    For i = 1 To 14
       ConNom1.Row = i
       ConNom1.Col = 1
       If Val(ConNom1.Text) > 0 Then
             subsidio.liminfs = ConNom1.Text
             yavas = 1
             Else
             subsidio.liminfs = 0
       End If
       ConNom1.Col = 2
       If Val(ConNom1.Text) > 0 Then
            subsidio.limsups = ConNom1.Text
            yavas = 1
            Else
            subsidio.liminfs = 0
       End If
       ConNom1.Col = 3
       If Val(ConNom1.Text) > 0 Then
            subsidio.cuotafs = ConNom1.Text
            yavas = 1
            Else
            subsidio.cuotafs = 0
       End If
       ConNom1.Col = 4
       If Val(ConNom1.Text) > 0 Then
            subsidio.porcsls = ConNom1.Text
            yavas = 1
            Else
            subsidio.porcsls = 0
       End If
       If yavas > 0 Then
                Put #4, i, subsidio: yavas = 0
                Else: i = 14
       End If
     Next i

End Sub

Private Sub EdCop_Click()
       Dim Temporal1
 Clipboard.Clear
   
   difer = ConNom1.RowSel - ConNom1.Row
   For i = ConNom1.Row To ConNom1.RowSel
      
      For f = ConNom1.Col To ConNom1.ColSel
            Temporal1 = Temporal1 + ConNom1.TextMatrix(i, f)
            If f < ConNom1.ColSel Then
                Temporal1 = Temporal1 & Chr(9)
            End If
      Next f
      Temporal1 = Temporal1 & Chr(13) & Chr(10)
      
   Next i
    Clipboard.SetText Temporal1

End Sub

Private Sub EdPg_Click()
          Dim DeAqui As Integer, RetornoCarro As Long, InicioCopia As Long, Temp_oral
  Temp_oral1 = ""
  Temp_oral = Clipboard.GetText(vbCFText)

  RetornoCarro = ConNom1.Col
  InicioCopia = ConNom1.Row: DeAqui = InicioCopia
  InIPGr = InicioCopia: InIPGc = RetornoCarro
If Temp_oral <> "" Then
  Clipboard.Clear
  DeAqui = 1
For i = 1 To Len(Temp_oral)
    Select Case Mid(Temp_oral, i, 1)
          Case Chr(9)
          Rem LTXT.Text = Mid(Temp_oral, DeAqui, (i - DeAqui))
          Temp_oral1 = Temp_oral1 + ConNom1.TextMatrix(ConNom1.Row, ConNom1.Col) & Chr(9)
          ConNom1.Text = Mid(Temp_oral, DeAqui, (i - DeAqui))
          ConNom1.Col = ConNom1.Col + 1
          DeAqui = i + 1
          Case Chr(13)
          Temp_oral1 = Temp_oral1 + ConNom1.TextMatrix(ConNom1.Row, ConNom1.Col) & Chr(13)
          Rem LTXT.Text = Mid(Temp_oral, DeAqui, (i - DeAqui))
          ConNom1.Text = Mid(Temp_oral, DeAqui, (i - DeAqui))
          If (ConNom1.Rows - 1) <= ConNom1.Row Then ConNom1.Rows = ConNom1.Rows + 100:
          ConNom1.Row = ConNom1.Row + 1: ConNom1.TopRow = 1
          DeAqui = i + 1
          Case Chr(10)
          Temp_oral1 = Temp_oral1 & Chr(10)
          ConNom1.Col = RetornoCarro
          DeAqui = i + 1
          Case Else
          Rem nada
    End Select
 Next i
 Rem CONNOM1.Row = InicioCopia:
 ConNom1.Col = RetornoCarro
End If
 
 Clipboard.Clear
 Clipboard.SetText Clipboard.GetText + Temp_oral1

End Sub

Private Sub EdTd_Click()
  ConNom1.Row = 1
    ConNom1.Col = 1
    ConNom1.ColSel = ConNom1.Cols - 1
    ConNom1.RowSel = ConNom1.Rows - 1

End Sub

Private Sub Form_Load()

Width = Screen.Width * 0.65 ' Establecer el ancho del formulario.
Height = Screen.Height * 0.78   ' Establecer el alto del formulario.
Left = (Screen.Width - Width) / 2   ' Centrar el formulario horizontalmente.
Top = (Screen.Height - Height) / 2  ' Centrar el formulario verticalmente.

Close 4
Select Case a_opago
    Case Is < 2008
        Open SubT_Mes + "TABLA.SUB" For Random As #4 Len = Len(subsidio)
        Form6.Caption = SubT_Mes + "Tarifa Mensual Subsidio Aplicable"
        
    Case Is > 2007
        Select Case Kincenal
            Case 0
            Open SubT_Mes + "Tab08Mes.SUB" For Random As #4 Len = Len(subsidio)
            Form6.Caption = SubT_Mes + "Tarifa Mensual Subsidio para empleo"
            Case 1
            Open SubT_Mes + "Tab08kin.SUB" For Random As #4 Len = Len(subsidio)
            Form6.Caption = SubT_Mes + "Tarifa Quincenal Subsidio para empleo"
            
            Case 2
            Open SubT_Mes + "Tab08Sem.SUB" For Random As #4 Len = Len(subsidio)
            Form6.Caption = SubT_Mes + "Tarifa Semanal Subsidio para empleo"
        End Select
            
End Select
em = LOF(4) / Len(subsidio)
ConNom1.ColWidth(0) = 330
ConNom1.Col = 0
ConNom1.Row = 0
ConNom1.Text = "#"
ConNom1.CellAlignment = 4
ConNom1.Col = 1
ConNom1.CellAlignment = 4
ConNom1.Text = "Lim.Inferior"
ConNom1.ColWidth(1) = 1500
ConNom1.Col = 2
ConNom1.CellAlignment = 4
ConNom1.Text = "Lim.Superior"
ConNom1.ColWidth(2) = 1500
ConNom1.Col = 3
ConNom1.CellAlignment = 4
ConNom1.Text = "Cuota Fija"
ConNom1.ColWidth(3) = 1500
ConNom1.Col = 4
ConNom1.CellAlignment = 4
ConNom1.Text = "%/excedente"
ConNom1.ColWidth(4) = 1500
If em > 0 Then
    For r = 1 To em: Get #4, r, subsidio
      ConNom1.Col = 0
      ConNom1.Row = r
      ConNom1.Text = r
      ConNom1.Col = 1
      ConNom1.CellAlignment = 7
      ConNom1.Text = Format(subsidio.liminfs, "###,##0.00")
      ConNom1.Col = 2
      ConNom1.CellAlignment = 7
      ConNom1.Text = Format(subsidio.limsups, "###,##0.00")
      ConNom1.Col = 3
      ConNom1.CellAlignment = 7
      ConNom1.Text = Format(subsidio.cuotafs, "###,##0.00")
      ConNom1.Col = 4
      ConNom1.CellAlignment = 7
      ConNom1.Text = Format(subsidio.porcsls, "###,##0.00")
    Next r
      Else
        For r = 1 To 14
           ConNom1.Col = 0
           ConNom1.Row = r
           ConNom1.Text = r
        Next r
    End If
     ConNom1.Row = 1
     ConNom1.Col = 1
     ConNom1.Appearance = flex3D
End Sub

Private Sub ConNom1_EnterCell()
    If ConNom1.Col > 0 And ConNom1.Row > 0 Then
        ConNom1.CellBackColor = &H80FF80
    End If
End Sub


Private Sub ConNom1_KeyPress(KeyAscii As Integer)
            Text1.Text = Chr$(KeyAscii)
            Text1.SetFocus
End Sub

Private Sub ConNom1_LeaveCell()
   If ConNom1.Col > 0 And ConNom1.Row > 0 Then
        ConNom1.CellBackColor = vbWhite
   End If
End Sub

Private Sub ConNom1_RowColChange()
   Text1.Text = ConNom1.Text
End Sub

Private Sub Form_Resize()
    Form6.Width = ConNom1.Width + 1400
    Form6.Height = ConNom1.Height + 2400
End Sub

Private Sub subimp_Click(Index As Integer)
  tam_largo = Printer.Height
    tam_ancho = Printer.Width
    
    Printer.Font.Size = 12
    Printer.CurrentY = tam_largo * 0.33
    Rem Printer.Print Tab(40); "Subsidio Impuesto sobre la renta."
    Printer.Print Tab(40); Form6.Caption
    Printer.Print
    guarda = Printer.CurrentY
    Printer.Line ((2500 + 1200), Printer.CurrentY)-((2500 + 6000), (Printer.CurrentY + 600)), , B
    Printer.Font.Size = 8
    Printer.CurrentY = guarda
    Printer.Print Tab(50); "Lim.Inf";
    Printer.Print Tab(65); "Lim.Sup";
    Printer.Print Tab(80); "Cuota fija.";
    Printer.Print Tab(95); " % "
    Printer.Print
    Printer.Print
    
    For r = 1 To 14
      For l = 1 To 4
         If ConNom1.TextMatrix(r, l) <> "" Then
                bala = ConNom1.TextMatrix(r, l)
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

Private Sub SubSal_Click(Index As Integer)
     Close
     Unload Form6
     Rem Load Form1
     Rem Form1.Show

End Sub

Private Sub Text1_Change()
    ConNom1.Text = Text1.Text
End Sub
Private Sub Text1_GotFocus()
SendKeys "{END}"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
       Case 13
         ConNom1.Text = Format(Val(ConNom1.Text), "###,##0.00")
         ConNom1.SetFocus
       Case 27
         ConNom1.Text = ""
         ConNom1.SetFocus
         
         
   End Select
End Sub



