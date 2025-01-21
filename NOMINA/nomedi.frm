VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form4 
   BackColor       =   &H8000000A&
   Caption         =   "NOMINA: EDICION DE PERSONAL"
   ClientHeight    =   7920
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   9840
   Icon            =   "nomedi.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   7920
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TexEdicion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9375
   End
   Begin MSFlexGridLib.MSFlexGrid ListPer 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   12091
      _Version        =   393216
      Cols            =   6
      FixedCols       =   2
   End
   Begin VB.Menu editarch 
      Caption         =   "&Archivo"
      Index           =   0
      Begin VB.Menu ArOrd 
         Caption         =   "&Ordenar"
         Begin VB.Menu ArOrAlf 
            Caption         =   "&Alfabeticamente"
         End
         Begin VB.Menu ArOrNum 
            Caption         =   "&Numericamente"
         End
      End
      Begin VB.Menu ArSep1 
         Caption         =   "-"
      End
      Begin VB.Menu editsale 
         Caption         =   "&Salida"
         Index           =   1
      End
   End
   Begin VB.Menu EditAr 
      Caption         =   "&Edicion"
      Begin VB.Menu EdCopiar 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu EdiSep1 
         Caption         =   "-"
      End
      Begin VB.Menu EdPegar 
         Caption         =   "&Pegar"
         Shortcut        =   ^V
      End
      Begin VB.Menu EdiSep2 
         Caption         =   "-"
      End
      Begin VB.Menu EditSelTot 
         Caption         =   "&Seleccionar Todo"
      End
      Begin VB.Menu EdSep3 
         Caption         =   "-"
      End
      Begin VB.Menu EdSup 
         Caption         =   "&Suprimir titulos"
      End
   End
   Begin VB.Menu edim 
      Caption         =   "&Imprimir"
      Index           =   0
      Begin VB.Menu edimalf 
         Caption         =   "&Pantalla"
         Index           =   1
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu edimnum 
         Caption         =   "&Total"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ValorAnt
Sub RecPra()
    ListPer.TextMatrix(ListPer.Row, 1) = Trim(personal.ape1) + " " + Trim(personal.ape2) + " " + Trim(personal.nom)
    ListPer.TextMatrix(ListPer.Row, 2) = Trim(personal.rfc)
    ListPer.TextMatrix(ListPer.Row, 3) = Trim(Otros_Rgtros.curp)
    ListPer.TextMatrix(ListPer.Row, 4) = " " + Trim(personal.imss)
    ListPer.TextMatrix(ListPer.Row, 5) = Trim(personal.fal)
    ListPer.TextMatrix(ListPer.Row, 6) = Trim(personal.fab)
    ListPer.TextMatrix(ListPer.Row, 7) = Format(personal.ingr, z1)
    ListPer.TextMatrix(ListPer.Row, 8) = Format(personal.viat, z1)
    ListPer.TextMatrix(ListPer.Row, 9) = Format(personal.otras, z1)
    ListPer.TextMatrix(ListPer.Row, 10) = Format(personal.integrado, z2)
    ListPer.TextMatrix(ListPer.Row, 11) = Format(Clbnx.Q1, "<")
End Sub
Sub ArchCorr()
  Close 1, 3
  Open "personal.dno" For Random As 1 Len = Len(personal)
  rgtro = ListPer.TextMatrix(ListPer.Row, 0)
  Open "PerOtre.dno" For Random As 3 Len = Len(Otros_Rgtros)
  Get 1, rgtro, personal
  Get 3, rgtro, Otros_Rgtros
  
  Select Case ListPer.Col
      Case 0, 1
      Rem nada
      Case 2
        personal.rfc = ListPer.Text
        Put 1, rgtro, personal
      Case 3
        Otros_Rgtros.curp = ListPer.Text
        Put 3, rgtro, Otros_Rgtros
        
      Case 4
        personal.imss = LTrim(ListPer.Text)
        Put 1, rgtro, personal
      Case 5
      If ListPer.Text <> "" Then
        Mes = Mid(ListPer.Text, 4, 2)
        yea = Mid(ListPer.Text, 7, 4)
        dia = Left(ListPer.Text, 2)
        If (Mes < 1) Or (Mes > 12) Then
            MsgBox "Mes invalido "
            Tex_Edicion = ValorAnt
            ListPer.Text = ValorAnt
            GoTo salte5
        End If
        If (dia < 1) Or (dia > dd(Mes)) Then
            MsgBox "dia invalido "
            Tex_Edicion = ValorAnt
            ListPer.Text = ValorAnt
            GoTo salte5
        End If
        If (yea < 0) Or (yea > empresa.ao) Then
             MsgBox "Año invalido "
             Tex_Edicion = ValorAnt
             ListPer.Text = ValorAnt
             GoTo salte5
        End If
      End If
        personal.fal = ListPer.Text
        Put 1, rgtro, personal
salte5:
     Case 6
     If ListPer.Text <> "" Then
        Mes = Mid(ListPer.Text, 4, 2)
        yea = Mid(ListPer.Text, 7, 4)
        dia = Left(ListPer.Text, 2)
        If (Mes < 1) Or (Mes > 12) Then
            MsgBox "Mes invalido "
            Tex_Edicion = ValorAnt
            ListPer.Text = ValorAnt
            GoTo salte6
        End If
        If (dia < 1) Or (dia > dd(Mes)) Then
            MsgBox "dia invalido "
            Tex_Edicion = ValorAnt
            ListPer.Text = ValorAnt
            GoTo salte6
        End If
        If (yea < 1900) Or (yea > empresa.ao) Then
             MsgBox "Año invalido "
             Tex_Edicion = ValorAnt
             ListPer.Text = ValorAnt
             GoTo salte6
        End If
     End If
    
        personal.fab = ListPer.Text
        Put 1, rgtro, personal
salte6:
     Case 7
        If ListPer.Text = "" Then ListPer.Text = 0
        personal.ingr = ListPer.Text
        ListPer.Text = Format(ListPer.Text, z2$)
    If personal.ingr > 0 Then
        yea = Mid(personal.fal, 7, 4)
    If yea < 1900 Then
         MsgBox "La fecha de ingreso no es correcta, no es posible calcular el salario integrado"
         Else
         antig = empresa.ao + 1 - yea
         facto = 0
         toTingr = personal.ingr + personal.viat + personal.otras
         factor antig, facto
         personal.integrado = toTingr * facto
         If personal.integrado > (empresa.sm * 25) Then personal.integrado = (empresa.sm * 25)
         ListPer.TextMatrix(ListPer.Row, 10) = Format(personal.integrado, z2$)
    End If
   End If
   Put 1, rgtro, personal
     Case 8
        If ListPer.Text = "" Then ListPer.Text = 0
        personal.viat = ListPer.Text
        ListPer.Text = Format(ListPer.Text, z2$)
        yea = Mid(personal.fal, 7, 4)
        If yea < 1900 Then
                MsgBox "La fecha de ingreso no es correcta, no es posible calcular el salario integrado"
            Else
                antig = empresa.ao + 1 - yea
                facto = 0
                toTingr = personal.ingr + personal.viat + personal.otras
                factor antig, facto
                personal.integrado = toTingr * facto
                If personal.integrado > (empresa.sm * 25) Then personal.integrado = (empresa.sm * 25)
                ListPer.TextMatrix(ListPer.Row, 10) = Format(personal.integrado, z2$)
        End If
        Put 1, rgtro, personal
     Case 9
        If ListPer.Text = "" Then ListPer.Text = 0
        personal.otras = ListPer.Text
        ListPer.Text = Format(ListPer.Text, z2$)
        yea = Mid(personal.fal, 7, 4)
        If yea < 1900 Then
                MsgBox "La fecha de ingreso no es correcta, no es posible calcular el salario integrado"
            Else
                antig = empresa.ao + 1 - yea
                facto = 0
                toTingr = personal.ingr + personal.viat + personal.otras
                factor antig, facto
                personal.integrado = toTingr * facto
                If personal.integrado > (empresa.sm * 25) Then personal.integrado = (empresa.sm * 25)
                ListPer.TextMatrix(ListPer.Row, 10) = Format(personal.integrado, z2$)
        End If
        Put 1, rgtro, personal
     Case 10
       MsgBox "No es posible modificar el salario Integrado "
       Tex_Edicion = ValorAnt
       ListPer.Text = ValorAnt
  End Select
End Sub
Sub pertit()
    fontviejo = Printer.FontSize
    Printer.FontSize = 10
    Printer.FontBold = True
    ancho1 = Int(Printer.TextWidth(empresa.name) / 2)
    Printer.CurrentX = (65 * 120) - ancho1
    Printer.Print empresa.name;
    Printer.Print
    titulo$ = "LISTADO DE PERSONAL " + LTrim$(Str$(empresa.ao))
    ancho1 = Int(Printer.TextWidth(titulo$) / 2)
    ancho2 = (55 * 120) - ancho1
    Printer.CurrentX = ancho2
    Printer.Print titulo$
    Printer.Print
    Printer.Line (0, Printer.CurrentY)-((130 * 120), Printer.CurrentY + 50), , BF
    Printer.FontSize = fontviejo
    Printer.Print
    Printer.CurrentX = (120 * 3)
    Printer.Print "No.";
    Printer.CurrentX = (120 * 10)
    Printer.Print " N o m b r e";
    Printer.CurrentX = (120 * 32)
    Printer.Print "R.F.C.";
    Printer.CurrentX = (120 * 42)
    Printer.Print " I.M.S.S.";
    Printer.CurrentX = (120 * 50)
    Printer.Print "F.Alta";
    Printer.CurrentX = (120 * 59)
    Printer.Print "F.baja";
    Printer.CurrentX = (120 * 68)
    Printer.Print "Diario";
    Printer.CurrentX = (120 * 76)
    Printer.Print "Viaticos";
    Printer.CurrentX = (120 * 84)
    Printer.Print "Otros";
    Printer.CurrentX = (120 * 90)
    Printer.Print "Integrado"
    Printer.Print
    Printer.Line (0, Printer.CurrentY)-((130 * 120), Printer.CurrentY + 50), , BF
    Printer.FontBold = False
    Printer.Print
    Printer.Print
    Printer.Print
End Sub
Sub linea()
        apelativo1$ = RTrim$(personal.ape1) + " " + RTrim$(personal.ape2) + " " + RTrim$(personal.nom)
        Printer.CurrentX = (120 * 7)
        Printer.Print apelativo1$;
        Printer.CurrentX = (120 * 30)
        Printer.Print personal.rfc;
        Printer.CurrentX = (120 * 40)
        Printer.Print personal.imss;
        Printer.CurrentX = (120 * 50)
        Printer.Print personal.fal;
        Printer.CurrentX = (120 * 58)
        Printer.Print personal.fab;
        Rem GoTo proximo
        valor$ = Format(personal.ingr, z2$): uso$ = z2$
        pone = 0: colocar pone, valor$, uso$
        Printer.CurrentX = (120 * 68) + pone
        Printer.Print valor$;
        Rem GoTo proximo
        If personal.viat > o Then
            valor$ = Format(personal.viat, z2$)
            pone = 0: colocar pone, valor$, uso$
            Printer.CurrentX = (120 * 76) + pone
            Printer.Print valor$;
        End If
        If personal.otras > 0 Then
            valor$ = Format(personal.otras, z2$)
            pone = 0: colocar pone, valor$, uso$
            Printer.CurrentX = (120 * 84) + pone
            Printer.Print valor$;
        End If
        valor$ = Format(personal.integrado, z2$)
        pone = 0: colocar pone, valor$, uso$
        Printer.CurrentX = (120 * 90) + pone
        Printer.Print valor$
proximo:
End Sub

Private Sub ArOrAlf_Click()
    colanti = ListPer.Col
    renati = ListPer.Row
    ListPer.Row = 1
    ListPer.Col = 1
    ListPer.RowSel = ListPer.Rows - 1
    ListPer.Sort = 1
    ListPer.Col = colanti
    ListPer.Row = renati

End Sub

Private Sub ArOrNum_Click()
    colanti = ListPer.Col
    renati = ListPer.Row
    ListPer.Row = 1
    ListPer.Col = 0
    ListPer.RowSel = ListPer.Rows - 1
    ListPer.Sort = 3
    ListPer.Col = colanti
    ListPer.Row = renati
    ListPer.SetFocus
End Sub

Private Sub EdCopiar_Click()
   Dim Temporal1
   Clipboard.Clear
   Rem For i = 0 To 0
      Rem For f = 0 To ListPer.ColSel
            Rem Temporal1 = Temporal1 + ListPer.TextMatrix(i, f)
            Rem If f < ListPer.ColSel Then
                Rem Temporal1 = Temporal1 & Chr(9)
            Rem End If
      Rem Next f
      Rem Temporal1 = Temporal1 & Chr(13) & Chr(10)
      
   Rem Next i

   
   For i = ListPer.Row To ListPer.RowSel
      Rem For f = 0 To ListPer.ColSel
      For f = ListPer.Col To ListPer.ColSel
            Temporal1 = Temporal1 + ListPer.TextMatrix(i, f)
            If f < ListPer.ColSel Then
                Temporal1 = Temporal1 & Chr(9)
            End If
      Next f
      Temporal1 = Temporal1 & Chr(13) & Chr(10)
      
   Next i
    Clipboard.SetText Temporal1
    ListPer.FixedCols = 2
    ListPer.FixedRows = 1
    ListPer.Row = 1: ListPer.Col = 2
    TexEdicion.Text = ListPer.Text
End Sub

Private Sub edimalf_Click(Index As Integer)
     Close 2
     Open "personal.dno" For Random As 2 Len = Len(personal)
     Dm = LOF(2) / Len(personal)
     Close 8
     Open "MAESTRO" For Random As 8 Len = Len(maestro)
     ddm = LOF(8) / Len(maestro)
 If Dm > 0 Then
     Printer.FontName = "Courier New"
     Printer.FontSize = 6
     Printer.Orientation = 1
     pertit
     For r = 1 To (ListPer.Rows - 1)
        rgtro = ListPer.TextMatrix(r, 0)
        valor = rgtro: uso$ = "####0.": ancho2 = 0
        colocar ancho2, valor, uso$
        Printer.CurrentX = 0 + ancho2
        Printer.Print rgtro;
        Get 2, rgtro, personal
        linea
        contador = contador + 1
        If contador > 83 Then
           contador = 0
           Printer.NewPage
           pertit
        End If
     Next r
     Printer.EndDoc
  End If
  Close 2
End Sub

Private Sub edimnum_Click(Index As Integer)
     Close 2, 8
     Open "personal.dno" For Random As 2 Len = Len(personal)
     Dm = LOF(2) / Len(personal)
     Open "maestro" For Random As 8 Len = Len(maestro)
     ddm = LOF(8) / Len(maestro)
 If Dm > 0 Then
     Printer.FontName = "Courier New"
     Printer.FontSize = 6
     Printer.Orientation = 1
     pertit
     For r = 1 To Dm
        rgtro = r
        valor = r: uso$ = "####0.": ancho2 = 0
        colocar ancho2, valor, uso$
        Printer.CurrentX = 0 + ancho2
        Printer.Print valor;
        Get 2, r, personal
        linea
        contador = contador + 1
        If contador > 83 Then
           contador = 0
           Printer.NewPage
           pertit
        End If
     Next r
     Printer.EndDoc
  End If
  Close 2
    
End Sub

Private Sub editsale_Click(Index As Integer)
 Unload Form4
 Form1.Show
End Sub

Private Sub EditSelTot_Click()
    Clipboard.Clear
    ListPer.Row = 0
    ListPer.Col = 0
    ListPer.RowSel = ListPer.Rows - 1
    ListPer.ColSel = ListPer.Cols - 1
End Sub

Private Sub EdPegar_Click()
  Dim temporal, DeAqui As Integer, RetornoCarro As Long, InicioCopia As Long
  temporal = Clipboard.GetText(vbCFText)
  RetornoCarro = ListPer.Col
  InicioCopia = ListPer.Row
If temporal <> "" Then
  Clipboard.Clear
  DeAqui = 1
For i = 1 To Len(temporal)
    Select Case Mid(temporal, i, 1)
          Case Chr(9)
          TexEdicion.Text = Mid(temporal, DeAqui, (i - DeAqui))
          ListPer.Text = Mid(temporal, DeAqui, (i - DeAqui))
          ListPer.Col = ListPer.Col + 1
          DeAqui = i + 1
          Case Chr(13)
          TexEdicion.Text = Mid(temporal, DeAqui, (i - DeAqui))
          ListPer.Text = Mid(temporal, DeAqui, (i - DeAqui))
          ArchCorr
          ListPer.Row = ListPer.Row + 1
          DeAqui = i + 1
          Case Chr(10)
          ListPer.Col = RetornoCarro
          DeAqui = i + 1
          Case Else
          Rem nada
    End Select
 Next i
 ListPer.Row = InicioCopia: ListPer.Col = RetornoCarro
End If
End Sub

Private Sub EdSup_Click()
 ListPer.FixedCols = 0
 ListPer.FixedRows = 0
End Sub

Private Sub Form_Load()
   z2$ = "#,###,##0.0000"
   Close 2
   Open "personal.dno" For Random As 2 Len = Len(personal)
   Dm = LOF(2) / Len(personal)
   
   Close 3
   Open "PerOtre.dno" For Random As 3 Len = Len(Otros_Rgtros)
   dmper = LOF(3) / Len(Otros_Rgtros)
   Close 4
   Open "Bnxcla.dno" For Random As 4 Len = Len(Clbnx)
   Rem  dmper = LOF(3) / Len(Otros_Rgtros)
   
   
   If dmper < Dm Then
      If dmper < 1 Then dmper = 1
      For r = (dmper + 1) To Dm: Get 3, r, Otros_Rgtros
              Otros_Rgtros.curp = "": Otros_Rgtros.otra = ""
              Otros_Rgtros.yotra = "": Otros_Rgtros.yporsi = ""
              Put 3, r, Otros_Rgtros
      Next r
   End If
   ListPer.Cols = 12: ListPer.Rows = 1: ListPer.Row = 0
   ListPer.Col = 0: ListPer.CellAlignment = 4: ListPer.ColWidth(0) = 400: ListPer.Text = "#"
   ListPer.Col = 1: ListPer.CellAlignment = 4: ListPer.ColWidth(1) = 3200: ListPer.Text = "Nombre"
   ListPer.Col = 2: ListPer.CellAlignment = 4: ListPer.ColWidth(2) = 2200: ListPer.Text = "RFC"
   ListPer.Col = 3: ListPer.CellAlignment = 4: ListPer.ColWidth(3) = 2200: ListPer.Text = "CURP"
   ListPer.Col = 4: ListPer.CellAlignment = 4: ListPer.ColWidth(4) = 1600: ListPer.Text = "IMSS"
   ListPer.Col = 5: ListPer.CellAlignment = 4: ListPer.ColWidth(5) = 1200: ListPer.Text = "Fcha.Alta"
   ListPer.Col = 6: ListPer.CellAlignment = 4: ListPer.ColWidth(6) = 1200: ListPer.Text = "Fcha.Baja"
   ListPer.Col = 7: ListPer.CellAlignment = 4: ListPer.ColWidth(7) = 1200: ListPer.Text = "Salario dia"
   ListPer.Col = 8: ListPer.CellAlignment = 4: ListPer.ColWidth(8) = 1200: ListPer.Text = "Viaticos dia"
   ListPer.Col = 9: ListPer.CellAlignment = 4: ListPer.ColWidth(9) = 1200: ListPer.Text = "Otros diario"
   ListPer.Col = 10: ListPer.CellAlignment = 4: ListPer.ColWidth(10) = 1200: ListPer.Text = "Integrado"
   ListPer.Col = 11: ListPer.CellAlignment = 4: ListPer.ColWidth(11) = 2200: ListPer.Text = "Tarjeta Bnx"
   If Dm > 0 Then
        For r = 1 To Dm: Get 2, r, personal: Get 3, r, Otros_Rgtros: Get 4, r, Clbnx
             apelativo1$ = "": abaja = Val(Mid(personal.fab, 7, 4))
             'If (abaja < 1900 Then abaja = abaja + 2000
             If (abaja > 0) And (abaja < empresa.ao - 1) Then GoTo Sig_te
             If (personal.ape1 >= "A") Or (personal.ape2 >= "A") Then
               apelativo1$ = RTrim$(personal.ape1) + " " + RTrim$(personal.ape2) + " " + RTrim$(personal.nom)
               apelativo$ = Left(apelativo1$, 59) + String$(60 - Len(Left(apelativo1$, 59)), " ")
               ListPer.AddItem Format(r, "####0") & Chr(9) & apelativo$ _
               & Chr(9) & RTrim(personal.rfc) & Chr(9) & RTrim(Otros_Rgtros.curp) & Chr(9) & (" " + RTrim(personal.imss)) _
               & Chr(9) & RTrim(personal.fal) & Chr(9) & RTrim(personal.fab) _
               & Chr(9) & Format(personal.ingr, z2) & Chr(9) & Format(personal.viat, z2) _
               & Chr(9) & Format(personal.otras, z2) & Chr(9) & Format(personal.integrado, z2) _
               & Chr(9) & Format(Clbnx.Q1, "<")
             End If
Sig_te:
        Next r
        Close 2, 3
        Else
        MsgBox "No existe personal para la edicion "
        Close 2, 3
        Rem Unload Form4
        Load Form1
        Form1.Show
   End If
End Sub
Sub REPONE()
 Dim Repuse
         Get 2, r, personal: Get 3, r, Otros_Rgtros: Get 4, r, Clbnx
           Repuse = Format(r, "####0") & Chr(9) & apelativo$ _
               & Chr(9) & RTrim(personal.rfc) & Chr(9) & RTrim(Otros_Rgtros.curp) & Chr(9) & (" " + RTrim(personal.imss)) _
               & Chr(9) & RTrim(personal.fal) & Chr(9) & RTrim(personal.fab) _
               & Chr(9) & Format(personal.ingr, z2) & Chr(9) & Format(personal.viat, z2) _
               & Chr(9) & Format(personal.otras, z2) & Chr(9) & Format(personal.integrado, z2) _
               & Chr(9) & Format(Clbnx.Q1, "<")
End Sub
Private Sub list1_DblClick()
       list1_KeyPress 13
End Sub

Private Sub list1_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        rgtro = Val(Mid$(list1.List(list1.ListIndex), 61, 6))
        Load Form2
        Form2.Show
     End If
End Sub

Private Sub Form_Resize()
    ListPer.Width = ScaleWidth - 200
End Sub

Private Sub listper_EnterCell()
    If ListPer.Col > 1 And ListPer.Row > 0 Then
        ListPer.CellBackColor = &H80FF80
    End If
    
End Sub


Private Sub listper_KeyDown(KeyCode As Integer, Shift As Integer)
      Select Case KeyCode
            Case vbKeyDelete
                ListPer.Text = ""
                Tex_Edicion = ListPer.Text
                ArchCorr
            Case vbKeyF2
                TexEdicion.Text = ListPer.Text
                TexEdicion.SetFocus
       End Select
End Sub

Private Sub ListPer_KeyPress(KeyAscii As Integer)
     Select Case KeyAscii
         Case 13:
           rgtro = ListPer.TextMatrix(ListPer.Row, 0)
           Load Form2
           Form2.Show 1
             Open "personal.dno" For Random As 2 Len = Len(personal)
             Get 2, rgtro, personal
             'maestro.por_1 = Text13.Text
             Rem Get #8, rgtro, maestro
              Close 3
             Open "PerOtre.dno" For Random As 3 Len = Len(Otros_Rgtros)
             dmper = LOF(3) / Len(personal)
             RecPra
             TexEdicion.Text = ListPer.Text
             Close 2, 3
             Rem BORRAR

           'Select Case ListPer.Col
               'Case 2
                
                 'ListPer.TextMatrix(ListPer.Row, ListPer.Col) = personal.rfc
                 'TexEdicion.Text = personal.rfc
               'Case 3
                 'ListPer.TextMatrix(ListPer.Row, ListPer.Col) = Otros_Rgtros.curp
                 'TexEdicion.Text = Otros_Rgtros.curp
                 
               'Case 4
                 'ListPer.TextMatrix(ListPer.Row, ListPer.Col) = personal.fal
                 'TexEdicion.Text = personal.fal
               'Case 5
                 'ListPer.TextMatrix(ListPer.Row, ListPer.Col) = personal.fab
                 'TexEdicion.Text = personal.fab
               'Case Else
                 ' nada
                  
           'End Select
         Case 27
             Rem nada
         Case Else
           'TexEdicion.Text = Chr(KeyAscii)
           'TexEdicion.SetFocus
     End Select
     Rem TexEdicion.Text = ListPer.Text
     Rem TexEdicion.SetFocus
      
End Sub

Private Sub listper_LeaveCell()
    
   If ListPer.Col > 1 And ListPer.Row > 0 Then
         ListPer.Text = UCase(TexEdicion.Text)
         ListPer.CellBackColor = vbWhite
   End If
End Sub

Private Sub listper_RowColChange()
       ValorAnt = ListPer.Text
       TexEdicion.Text = ListPer.Text
End Sub


Private Sub TexEdicion_Change()
    ListPer.Text = TexEdicion.Text
End Sub

Private Sub TexEdicion_GotFocus()
On Error Resume Next
    SendKeys "{end}"
End Sub

Private Sub TexEdicion_KeyPress(KeyAscii As Integer)
     Select Case KeyAscii
       Case 13
        
        TexEdicion.Text = UCase(TexEdicion.Text)
        ListPer.Text = TexEdicion.Text
        ArchCorr
        ListPer.SetFocus
       Case 27
        TexEdicion.Text = ListPer.Text
       
       
     End Select
End Sub

Private Sub TexEdicion_LostFocus()
    
    TexEdicion_KeyPress 13
End Sub

