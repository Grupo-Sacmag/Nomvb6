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
' ============================================================
' GUARDADO DE SUBSIDIO
' ============================================================

Private Sub ArGdr_Click()
    yavas = 0

    For i = 1 To 14
        connom1.Row = i

        connom1.col = 1
        If Val(connom1.Text) > 0 Then
            subsidio.liminfs = connom1.Text
            yavas = 1
        Else
            subsidio.liminfs = 0
        End If

        connom1.col = 2
        If Val(connom1.Text) > 0 Then
            subsidio.limsups = connom1.Text
            yavas = 1
        Else
            subsidio.liminfs = 0
        End If

        connom1.col = 3
        If Val(connom1.Text) > 0 Then
            subsidio.cuotafs = connom1.Text
            yavas = 1
        Else
            subsidio.cuotafs = 0
        End If

        connom1.col = 4
        If Val(connom1.Text) > 0 Then
            subsidio.porcsls = connom1.Text
            yavas = 1
        Else
            subsidio.porcsls = 0
        End If

        If yavas > 0 Then
            Put #4, i, subsidio
            yavas = 0
        Else
            i = 14
        End If
    Next i
End Sub

' ============================================================
' COPIAR / PEGAR / SELECCIONAR
' ============================================================

Private Sub EdCop_Click()
    Dim Temporal1

    Clipboard.Clear

    difer = connom1.RowSel - connom1.Row

    For i = connom1.Row To connom1.RowSel
        For f = connom1.col To connom1.ColSel
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
    Dim DeAqui As Integer
    Dim RetornoCarro As Long
    Dim InicioCopia As Long
    Dim Temp_oral

    Temp_oral1 = ""
    Temp_oral = Clipboard.GetText(vbCFText)

    RetornoCarro = connom1.col
    InicioCopia = connom1.Row
    DeAqui = InicioCopia
    InIPGr = InicioCopia
    InIPGc = RetornoCarro

    If Temp_oral <> "" Then
        Clipboard.Clear
        DeAqui = 1

        For i = 1 To Len(Temp_oral)
            Select Case Mid(Temp_oral, i, 1)
                Case Chr(9)
                    Temp_oral1 = Temp_oral1 + connom1.TextMatrix(connom1.Row, connom1.col) & Chr(9)
                    connom1.Text = Mid(Temp_oral, DeAqui, (i - DeAqui))
                    connom1.col = connom1.col + 1
                    DeAqui = i + 1

                Case Chr(13)
                    Temp_oral1 = Temp_oral1 + connom1.TextMatrix(connom1.Row, connom1.col) & Chr(13)
                    connom1.Text = Mid(Temp_oral, DeAqui, (i - DeAqui))

                    If (connom1.Rows - 1) <= connom1.Row Then connom1.Rows = connom1.Rows + 100

                    connom1.Row = connom1.Row + 1
                    connom1.TopRow = 1
                    DeAqui = i + 1

                Case Chr(10)
                    Temp_oral1 = Temp_oral1 & Chr(10)
                    connom1.col = RetornoCarro
                    DeAqui = i + 1

                Case Else
                    Rem nada
            End Select
        Next i

        connom1.col = RetornoCarro
    End If

    Clipboard.Clear
    Clipboard.SetText Clipboard.GetText + Temp_oral1
End Sub

Private Sub EdTd_Click()
    connom1.Row = 1
    connom1.col = 1
    connom1.ColSel = connom1.Cols - 1
    connom1.RowSel = connom1.Rows - 1
End Sub

' ============================================================
' CARGA DEL FORMULARIO
' ============================================================

Private Sub Form_Load()
    PrepararFormulario
    AbrirArchivoSubsidio
    ConfigurarGrid
    CargarSubsidioEnGrid

    connom1.Row = 1
    connom1.col = 1
    connom1.Appearance = flex3D
End Sub

Private Sub PrepararFormulario()
    Width = Screen.Width * 0.65
    Height = Screen.Height * 0.78
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
End Sub

Private Sub AbrirArchivoSubsidio()
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

    EM = LOF(4) / Len(subsidio)
End Sub

Private Sub ConfigurarGrid()
    connom1.ColWidth(0) = 330

    connom1.col = 0
    connom1.Row = 0
    connom1.Text = "#"
    connom1.CellAlignment = 4

    connom1.col = 1
    connom1.CellAlignment = 4
    connom1.Text = "Lim.Inferior"
    connom1.ColWidth(1) = 1500

    connom1.col = 2
    connom1.CellAlignment = 4
    connom1.Text = "Lim.Superior"
    connom1.ColWidth(2) = 1500

    connom1.col = 3
    connom1.CellAlignment = 4
    connom1.Text = "Cuota Fija"
    connom1.ColWidth(3) = 1500

    connom1.col = 4
    connom1.CellAlignment = 4
    connom1.Text = "%/excedente"
    connom1.ColWidth(4) = 1500
End Sub

Private Sub CargarSubsidioEnGrid()
    If EM > 0 Then
        For r = 1 To EM
            Get #4, r, subsidio

            connom1.col = 0
            connom1.Row = r
            connom1.Text = r

            connom1.col = 1
            connom1.CellAlignment = 7
            connom1.Text = Format(subsidio.liminfs, "###,##0.00")

            connom1.col = 2
            connom1.CellAlignment = 7
            connom1.Text = Format(subsidio.limsups, "###,##0.00")

            connom1.col = 3
            connom1.CellAlignment = 7
            connom1.Text = Format(subsidio.cuotafs, "###,##0.00")

            connom1.col = 4
            connom1.CellAlignment = 7
            connom1.Text = Format(subsidio.porcsls, "###,##0.00")
        Next r
    Else
        For r = 1 To 14
            connom1.col = 0
            connom1.Row = r
            connom1.Text = r
        Next r
    End If
End Sub

Private Sub Form_Resize()
    Form6.Width = connom1.Width + 1400
    Form6.Height = connom1.Height + 2400
End Sub

Private Sub SubSal_Click(Index As Integer)
    Close
    Unload Form6
End Sub

' ============================================================
' EVENTOS DEL GRID
' ============================================================

Private Sub ConNom1_EnterCell()
    If connom1.col > 0 And connom1.Row > 0 Then
        connom1.CellBackColor = &H80FF80
    End If
End Sub

Private Sub ConNom1_KeyPress(KeyAscii As Integer)
    Text1.Text = Chr$(KeyAscii)
    Text1.SetFocus
End Sub

Private Sub ConNom1_LeaveCell()
    If connom1.col > 0 And connom1.Row > 0 Then
        connom1.CellBackColor = vbWhite
    End If
End Sub

Private Sub ConNom1_RowColChange()
    Text1.Text = connom1.Text
End Sub

' ============================================================
' IMPRESION
' ============================================================

Private Sub subimp_Click(Index As Integer)
    tam_largo = Printer.Height
    tam_ancho = Printer.Width

    Printer.Font.Size = 12
    Printer.CurrentY = tam_largo * 0.33
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
        For L = 1 To 4
            If connom1.TextMatrix(r, L) <> "" Then
                bala = connom1.TextMatrix(r, L)
                valor$ = Format(bala, "##,###,##0.00")
                uso$ = "##,###,##0.00"
                pone = 0

                colocar pone, valor$, uso$

                Printer.CurrentX = (2500 + (1200 * L)) + pone
                Printer.Print valor$;
            End If
        Next L

        Printer.Print
    Next r

    Printer.EndDoc
End Sub

' ============================================================
' EDICION CON TEXT1
' ============================================================

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
            connom1.Text = ""
            connom1.SetFocus
    End Select
End Sub

