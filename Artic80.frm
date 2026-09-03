VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form5 
   Caption         =   "TABLA ARTICULO 80"
   ClientHeight    =   5475
   ClientLeft      =   1770
   ClientTop       =   2715
   ClientWidth     =   7245
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   7245
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin MSFlexGridLib.MSFlexGrid connom1 
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5318
      _Version        =   393216
      Rows            =   15
      Cols            =   5
   End
   Begin VB.Menu indart 
      Caption         =   "&Archivo"
      Index           =   0
      Begin VB.Menu ArGuard 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu ArSep1 
         Caption         =   "-"
      End
      Begin VB.Menu tarimp 
         Caption         =   "&Imprimir"
         Index           =   2
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu indsaltar 
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
      Begin VB.Menu EdIns 
         Caption         =   "&Insertar"
      End
      Begin VB.Menu EdSep3 
         Caption         =   "-"
      End
      Begin VB.Menu EdEl 
         Caption         =   "&Eliminar"
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mensaje, titulo, valorpred, mivalor, temporal

Private Sub ArGuard_Click()
    yavas = 0
    For i = 1 To 14:
       connom1.Row = i
       connom1.Col = 1
       If Val(connom1.Text) > 0 Then
             articulo.liminf = connom1.Text
             yavas = 1
             Else
             articulo.liminf = 0
       End If
       connom1.Col = 2
       If Val(connom1.Text) > 0 Then
            articulo.limsup = connom1.Text
            yavas = 1
            Else
            articulo.liminf = 0
       End If
       connom1.Col = 3
       If Val(connom1.Text) > 0 Then
            articulo.cuotaf = connom1.Text
            yavas = 1
            Else
            articulo.cuotaf = 0
       End If
       connom1.Col = 4
       If Val(connom1.Text) > 0 Then
            articulo.porcsl = connom1.Text
            yavas = 1
            Else
            articulo.porcsl = 0
       End If
       If yavas > 0 Then
                Put #3, i, articulo: yavas = 0
                Else: i = 14
       End If
     Next i
     

End Sub

Private Sub EdCop_Click()
       Dim Temporal1
 Clipboard.Clear
   
   difer = connom1.RowSel - connom1.Row
   For i = connom1.Row To connom1.RowSel
      
      For f = connom1.Col To connom1.ColSel
            Temporal1 = Temporal1 + connom1.TextMatrix(i, f)
            If f < connom1.ColSel Then
                Temporal1 = Temporal1 & Chr(9)
            End If
      Next f
      Temporal1 = Temporal1 & Chr(13) & Chr(10)
      
   Next i
    Clipboard.SetText Temporal1

End Sub

Private Sub EdPg_Click()
       Dim DeAqui As Integer, RetornoCarro As Long, InicioCopia As Long
  Temporal1 = ""
  temporal = Clipboard.GetText(vbCFText)

  RetornoCarro = connom1.Col
  InicioCopia = connom1.Row: DeAqui = InicioCopia
  InIPGr = InicioCopia: InIPGc = RetornoCarro
If temporal <> "" Then
  Clipboard.Clear
  DeAqui = 1
For i = 1 To Len(temporal)
    Select Case Mid(temporal, i, 1)
          Case Chr(9)
          Rem LTXT.Text = Mid(TemPOraL, DeAqui, (i - DeAqui))
          Temporal1 = Temporal1 + connom1.TextMatrix(connom1.Row, connom1.Col) & Chr(9)
          connom1.Text = Mid(temporal, DeAqui, (i - DeAqui))
          connom1.Col = connom1.Col + 1
          DeAqui = i + 1
          Case Chr(13)
          Temporal1 = Temporal1 + connom1.TextMatrix(connom1.Row, connom1.Col) & Chr(13)
          Rem LTXT.Text = Mid(TemPOraL, DeAqui, (i - DeAqui))
          connom1.Text = Mid(temporal, DeAqui, (i - DeAqui))
          If (connom1.Rows - 1) <= connom1.Row Then connom1.Rows = connom1.Rows + 100:
          connom1.Row = connom1.Row + 1: connom1.TopRow = 1
          DeAqui = i + 1
          Case Chr(10)
          Temporal1 = Temporal1 & Chr(10)
          connom1.Col = RetornoCarro
          DeAqui = i + 1
          Case Else
          Rem nada
    End Select
 Next i
 Rem CONNOM1.Row = InicioCopia:
 connom1.Col = RetornoCarro
End If
 
 Clipboard.Clear
 Clipboard.SetText Clipboard.GetText + Temporal1

End Sub

Private Sub EdTd_Click()
    connom1.Row = 1
    connom1.Col = 1
    connom1.ColSel = connom1.Cols - 1
    connom1.RowSel = connom1.Rows - 1
End Sub

Private Sub Form_Load()
    Width = Screen.Width * 0.65 ' Establecer el ancho del formulario.
    Height = Screen.Height * 0.78   ' Establecer el alto del formulario.
    Left = (Screen.Width - Width) / 2   ' Centrar el formulario horizontalmente.
    Top = (Screen.Height - Height) / 2  ' Centrar el formulario verticalmente.
    AoFi = Val(empresa.ao)
    
If ta_r = 0 Then
    
    Close 3
    Select Case a_opago
    Case Is < 2008
        Open SubT_Mes + "TABLA.ISR" For Random As #3 Len = Len(articulo)
        Form5.Caption = SubT_Mes + "Tarifa Mensual retencion ISR"
    Case Is > 2007
        Select Case Kincenal
            Case 0
            Open SubT_Mes + "Tab08Mes.ISR" For Random As #3 Len = Len(articulo)
            Form5.Caption = SubT_Mes + "Tarifa Mensual retencion ISR"
            Case 1
            Open Trim(SubT_Mes) + "Tab08Kin.ISR" For Random As #3 Len = Len(articulo)
            Form5.Caption = SubT_Mes + "Tarifa Quincenal retencion ISR"
            Case 2
            Open Trim(SubT_Mes) + "Tab08Sem.ISR" For Random As #3 Len = Len(articulo)
            Form5.Caption = SubT_Mes + "Tarifa Semanal retencion ISR"
        End Select
      End Select
    Else
    Close 3
    Open arch_tr For Random As #3 Len = Len(articulo)
    If ta_r = 1 Then Form5.Caption = "Tarifa ISR Anual 177"
    If ta_r = 2 Then Form5.Caption = "Tarifa Subisidio Anual 178"
    
    ta_r = 0
End If
Dem = LOF(3) / Len(articulo)
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
connom1.Col = 4
connom1.CellAlignment = 4
connom1.Text = "%/excedente"
connom1.ColWidth(4) = 1500
If Dem > 0 Then
    For r = 1 To Dem: Get #3, r, articulo
      connom1.Col = 0
      connom1.Row = r
      connom1.Text = r
      connom1.Col = 1
      connom1.CellAlignment = 7
      connom1.Text = Format(articulo.liminf, "###,##0.00")
      connom1.Col = 2
      connom1.CellAlignment = 7
      connom1.Text = Format(articulo.limsup, "###,##0.00")
      connom1.Col = 3
      connom1.CellAlignment = 7
      connom1.Text = Format(articulo.cuotaf, "###,##0.00")
      connom1.Col = 4
      connom1.CellAlignment = 7
      connom1.Text = Format(articulo.porcsl, "###,##0.00")
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
   Me.Width = connom1.Width + 400
   Me.Height = connom1.Height + 2200
End Sub


Private Sub indsaltar_Click(Index As Integer)
    Close
    Unload Form5
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

Private Sub tarimp_Click(Index As Integer)
    
    tam_largo = Printer.Height
    tam_ancho = Printer.Width
    
    Printer.Font.Size = 12
    Printer.CurrentY = tam_largo * 0.33
    Rem Printer.Print Tab(40); "Retencion Impuesto sobre la renta."
    Printer.Print Tab(40); Form5.Caption
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
      For L = 1 To 4
         If connom1.TextMatrix(r, L) <> "" Then
                bala = connom1.TextMatrix(r, L)
                valor$ = Format(bala, "##,###,##0.00"): uso$ = "##,###,##0.00"
                pone = 0: colocar pone, valor$, uso$
                Printer.CurrentX = (2500 + (1200 * L)) + pone
                Printer.Print valor$;
         End If
      Next L
      Printer.Print
    Next r
    Printer.EndDoc
End Sub

Private Sub Text1_Change()
    connom1.Text = Text1.Text
  
End Sub


Private Sub Text1_GotFocus()
SendKeys "{END}"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
       Case 13
         connom1.Text = Format(Val(connom1.Text), "###,##0.00")
         connom1.SetFocus
       Case 27
         connom1.Text = mientras$
         connom1.Text = ""
         connom1.SetFocus
         
         
   End Select
End Sub

' comentario
