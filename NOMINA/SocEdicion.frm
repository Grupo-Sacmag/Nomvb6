VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form SocEdicion 
   BackColor       =   &H8000000B&
   Caption         =   "Edición Datos Socios"
   ClientHeight    =   7695
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   14280
   Icon            =   "SocEdicion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   14280
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Socios1 
      Height          =   6615
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   11668
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   2
      BackColorFixed  =   -2147483643
      BackColorSel    =   -2147483639
      BackColorBkg    =   14737632
      Appearance      =   0
   End
   Begin VB.Menu Ord 
      Caption         =   "&Ordenar"
      Begin VB.Menu OrdAlf 
         Caption         =   "&Alfabeticamente"
      End
      Begin VB.Menu OrdNum 
         Caption         =   "&Numericamente"
      End
   End
   Begin VB.Menu EdiC 
      Caption         =   "&Editar"
      Begin VB.Menu EdiCop 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu EdSelT 
         Caption         =   "&Seleccionar Todo"
      End
   End
End
Attribute VB_Name = "SocEdicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub EdiCop_Click()
            Clipboard.Clear
   Rem Clipboard.SetText Clipboard.GetText + Poliza1.Caption & Chr(13)
   Rem Clipboard.SetText Clipboard.GetText + Label1.Caption & Chr(13)
   Rem socios1.RowSel = socios1.Rows - 1
   Rem socios1.ColSel = socios1.Cols - 1
   Clipboard.SetText Clipboard.GetText + Socios1.TextMatrix(0, 0) & Chr(9)
   For F = Socios1.Col To Socios1.ColSel
         Clipboard.SetText Clipboard.GetText + Socios1.TextMatrix(0, F) & Chr(9)
   Next F
        Clipboard.SetText Clipboard.GetText + Chr(13)
   For I = Socios1.Row To Socios1.RowSel
           Clipboard.SetText Clipboard.GetText + Socios1.TextMatrix(I, 0) & Chr(9)
      For F = Socios1.Col To Socios1.ColSel
            Clipboard.SetText Clipboard.GetText + Socios1.TextMatrix(I, F) & Chr(9)
      Next F
      Clipboard.SetText Clipboard.GetText + Chr(13)
   Next I
   difer = Socios1.RowSel - Socios1.Row

End Sub

Private Sub EdSelT_Click()
   Socios1.Col = 1: Socios1.Row = 1
    Socios1.RowSel = Socios1.Rows - 1
    Socios1.ColSel = Socios1.Cols - 1
End Sub

Private Sub Form_Load()
   Socios1.Cols = 12
   Socios1.Row = 0: Socios1.Col = 0
   Socios1.ColWidth(0) = 400: Socios1.CellAlignment = 4: Socios1.CellFontBold = True: Socios1.Text = "No."
   Socios1.Col = 1: Socios1.CellAlignment = 4: Socios1.ColWidth(1) = 3500: Socios1.CellFontBold = True: Socios1.Text = "Nombre"
   Socios1.ColWidth(2) = 2000: Socios1.Col = 2: Socios1.CellAlignment = 4: Socios1.CellFontBold = True: Socios1.Text = "Rfc"
   Socios1.ColWidth(3) = 2000: Socios1.Col = 3: Socios1.CellAlignment = 4: Socios1.CellFontBold = True: Socios1.Text = "Curp"
   Socios1.ColWidth(4) = 1000: Socios1.Col = 4: Socios1.CellAlignment = 4: Socios1.CellFontBold = True: Socios1.Text = "F.Alta"
   Socios1.ColWidth(5) = 1000: Socios1.Col = 5: Socios1.CellAlignment = 4: Socios1.CellFontBold = True: Socios1.Text = "F.Baja"
   Socios1.ColWidth(6) = 1000: Socios1.Col = 6: Socios1.CellAlignment = 4: Socios1.CellFontBold = True: Socios1.Text = "Ingr.Norm."
   Socios1.ColWidth(7) = 1000: Socios1.Col = 7: Socios1.CellAlignment = 4: Socios1.CellFontBold = True: Socios1.Text = "Ingr.Adic."
   Socios1.ColWidth(8) = 1000: Socios1.Col = 8: Socios1.CellAlignment = 4: Socios1.CellFontBold = True: Socios1.Text = "Sub.Adic"
   Socios1.ColWidth(9) = 2000: Socios1.Col = 9: Socios1.CellAlignment = 4:  Socios1.CellFontBold = True: Socios1.Text = "Calle"
   Socios1.ColWidth(10) = 2000: Socios1.Col = 10: Socios1.CellAlignment = 4: Socios1.CellFontBold = True: Socios1.Text = "Colonia"
   Socios1.ColWidth(11) = 2000: Socios1.Col = 11: Socios1.CellAlignment = 4: Socios1.CellFontBold = True: Socios1.Text = "Ciudad"
   
   Close 2: Open "pers.nom" For Random As 2 Len = Len(Socios)
   Dm = LOF(2) / Len(Socios)
   Socios1.Row = 0: Socios1.Col = 0
   
    For r = 1 To Dm: Get 2, r, Socios
       If (Socios.ape1) > " " Or (Socios.ape2 > " ") Then
          Socios1.AddItem Format(r, "###0") & Chr(9) & RTrim(Socios.ape1) + " " + RTrim(Socios.ape2) + " " + RTrim(Socios.Nom) _
          & Chr(9) & Socios.rfc & Chr(9) & Socios.curp & Chr(9) & Socios.fal & Chr(9) & Socios.fab _
          & Chr(9) & Format(Socios.ingr, "#,##0.00;(#,##0.00)") & Chr(9) & Format(Socios.otroing, "#,##0.00;(#,##0.00)") _
          & Chr(9) & Format(Socios.otrsub, "#0.0000") & Chr(9) & Socios.calle & Chr(9) & Socios.colonia & Chr(9) & Socios.ciudad
       End If
    Next r
    Socios1.Col = 1
    'rem socios1_LeaveCell
    colanti = Socios1.Col
    renati = Socios1.Row
    Socios1.RowSel = Socios1.Rows - 1
    Socios1.Sort = 1
    Socios1.Col = colanti
    Socios1.Row = renati
    Rem socios1_LeaveCell
    Rem socios1_EnterCell
    Rem Socios1.SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Close 2
End Sub

Private Sub OrdAlf_Click()
    Socios1.Row = 1
    Socios1.Col = 1
    colanti = Socios1.Col
    renati = Socios1.Row
    Socios1.RowSel = Socios1.Rows - 2
    Socios1.Sort = 1
    Socios1.Col = colanti
    Socios1.Row = renati
    Socios1.SetFocus

End Sub

Private Sub OrdNum_Click()
    Socios1.Col = 1
    Socios1.Row = 1
    
    colanti = Socios1.Col
    renati = Socios1.Row
    Socios1.Col = 0
    Socios1.RowSel = Socios1.Rows - 2
    Socios1.Sort = 1
    Socios1.Col = colanti
    Socios1.Row = renati
    
    
    Socios1.SetFocus

End Sub

Private Sub socios1_Click()
    Rem Socios1.SelectionMode = flexSelectionByRow
End Sub

Private Sub Socios1_DblClick()
              ddm = Socios1.TextMatrix(Socios1.Row, 0)
              DatSocios.Show 1

End Sub

Private Sub Socios1_KeyDown(KeyCode As Integer, Shift As Integer)
          Select Case KeyCode
            Case vbKeyDelete
              respuesta = MsgBox("DESEA ELIMINAR ESTE RENGLON ", vbYesNo + vbCritical + vbDefaultButton2)
              If respuesta = vbYes Then
                Rem ANTPO1.Text = ""
                borreg
                Socios1.RemoveItem Socios1.Row
                Socios1.Clear
                Socios1.Rows = 1
                Form_Load
              End If
          End Select
End Sub
Sub borreg()
   Open "peso.tem" For Random As 8 Len = Len(Socios)
   ddm = Socios1.TextMatrix(Socios1.Row, 0)
   Get 2, ddm, Socios
   
   
        Socios.rfc = " ": Socios.curp = " "
        Socios.ingr = 0: Socios.ape1 = " "
        Socios.ape2 = " ": Socios.Nom = " "
        Socios.fal = " ": Socios.fab = " "
        Socios.otroing = 0
   Put 2, ddm, Socios
   fi_n = LOF(2) / Len(Socios)
   
   If ddm < fi_n Then
            For r = (ddm + 1) To fi_n: Get 2, r, Socios
                Put 2, r - 1, Socios
            Next r
   End If
   For r = 1 To (fi_n - 1): Get 2, r, Socios
        Put 8, r, Socios
   Next r
   Close 8, 2
   FileCopy "peso.tem", "personal.nom"
   Kill "peso.tem"
   
End Sub
Private Sub Socios1_KeyPress(KeyAscii As Integer)
       Select Case (KeyAscii)
              Case 27
              Rem Unload SocEdicion
              Case 13
              ddm = Socios1.TextMatrix(Socios1.Row, 0)
              DatSocios.Show 1
       End Select
End Sub
