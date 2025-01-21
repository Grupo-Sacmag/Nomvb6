VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form AJTECH 
   Caption         =   "Ajuste en MILIMETROS  impresión de cheques"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid AJTE 
      Height          =   3720
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   6562
      _Version        =   393216
      Rows            =   15
      Cols            =   3
      BackColorFixed  =   -2147483634
      FocusRect       =   2
      GridLinesFixed  =   1
   End
   Begin VB.Menu ach 
      Caption         =   "&Archivo"
      Index           =   0
      Begin VB.Menu achArch 
         Caption         =   "&Guardar"
         Index           =   1
      End
      Begin VB.Menu ArcSep1 
         Caption         =   "-"
      End
      Begin VB.Menu achImpr 
         Caption         =   "&Impresion datos"
         Index           =   2
      End
      Begin VB.Menu achPrue 
         Caption         =   "&Prueba Impresion"
         Index           =   3
      End
      Begin VB.Menu ArSep2 
         Caption         =   "-"
      End
      Begin VB.Menu achSal 
         Caption         =   "&Salida"
         Index           =   4
      End
   End
   Begin VB.Menu EdAj 
      Caption         =   "&Edición"
      Begin VB.Menu EdAjCop 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "AJTECH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim valcelant
Dim fila(50) As Integer, ubica(50) As Integer, rubro(50) As String * 30
Sub depura(elemento)
    If IsNumeric(elemento) Then
        Rem Nada
        Else
        MsgBox "Solo se aceptan valores numericos"
        Text1.Text = valcelant
    End If
End Sub

Private Sub achArch_Click(Index As Integer)
   archaj
End Sub

Private Sub achImpr_Click(Index As Integer)
   AJTECH.PrintForm
End Sub
Private Sub achSal_Click(Index As Integer)
  Close 8
  Unload AJTECH
End Sub

Private Sub AJTE_EnterCell()
   If AJTE.Row = 1 And AJTE.Col = 2 Then AJTE.Col = 1: AJTE_EnterCell
   If AJTE.Row = 7 And AJTE.Col = 2 Then AJTE.Col = 1: AJTE_EnterCell
   If AJTE.Row = 8 And AJTE.Col = 2 Then AJTE.Col = 1: AJTE_EnterCell
   If AJTE.Row > 8 And AJTE.Col = 1 Then AJTE.Col = 2: AJTE_EnterCell
    AJTE.CellBackColor = vbGreen
End Sub

Private Sub EdAjCop_Click()
   Clipboard.Clear
   Clipboard.SetText Clipboard.GetText + AJTE.TextMatrix(0, 0) & Chr(9)
   For f = 0 To AJTE.Cols - 1
         Clipboard.SetText Clipboard.GetText + AJTE.TextMatrix(0, f) & Chr(9)
   Next f
        Clipboard.SetText Clipboard.GetText + Chr(13)
        
   For i = 1 To AJTE.Rows - 1
           Clipboard.SetText Clipboard.GetText + AJTE.TextMatrix(i, 0) & Chr(9)
      For f = 1 To AJTE.Cols - 1
            Clipboard.SetText Clipboard.GetText + AJTE.TextMatrix(i, f) & Chr(9)
      Next f
      Clipboard.SetText Clipboard.GetText + Chr(13)
   Next i
   difer = AJTE.RowSel - AJTE.Row

End Sub

Private Sub Text1_Change()
    AJTE.Text = Text1.Text
End Sub
Private Sub Text1_GotFocus()
    SendKeys "{end}"
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Rem AJTE.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
       
       Case 13
        
        depura Text1.Text
        AJTE.Text = Text1.Text
        AJTE.SetFocus
        Rem nada
       Case 27
        Text1.Text = valcelant
        AJTE.SetFocus
       Rem nada
       
    End Select
End Sub

Private Sub AJTE_LeaveCell()
  AJTE.CellBackColor = vbWhite
End Sub

Private Sub AJTE_RowColChange()
     Rem AJTE.BackColor = vbWhite
     If AJTE.Text <> "" Then valcelant = AJTE.Text
     Text1.Text = AJTE.Text
End Sub
Private Sub AJTE_KeyDown(KeyCode As Integer, Shift As Integer)
      Select Case KeyCode
            Case vbKeyDelete
                AJTE.Text = ""
                Text1.Text = AJTE.Text
            Case vbKeyF2
                If AJTE.Text <> "" Then valcelant = AJTE.Text
                Text1.Text = AJTE.Text
                Text1.SetFocus
               
       End Select
End Sub

Private Sub AJTE_KeyPress(KeyAscii As Integer)
         Rem ajte.SelectionMode = flexSelectionFree
         If AJTE.Text <> "" Then valcelant = AJTE.Text
         Text1.Text = Chr$(KeyAscii)
         Text1.SetFocus
End Sub

Private Sub Form_Load()
    apeajte
    cargaaj
    AJTE.Row = 0
    AJTE.ColWidth(0) = 3600: AJTE.Col = 0: AJTE.CellAlignment = 4: AJTE.Text = "CONCEPTO"
    AJTE.ColWidth(1) = 600: AJTE.TextMatrix(0, 1) = "Reng."
    AJTE.ColWidth(2) = 600: AJTE.TextMatrix(0, 2) = "Col."
    AJTE.TextMatrix(1, 0) = "Total renglones de la forma"
    AJTE.TextMatrix(2, 0) = "Fecha"
    AJTE.TextMatrix(3, 0) = "Beneficiario"
    AJTE.TextMatrix(4, 0) = "Importe en número"
    AJTE.TextMatrix(5, 0) = "Importe en letra"
    AJTE.TextMatrix(6, 0) = "Concepto"
    AJTE.TextMatrix(7, 0) = "Inicio aplicación"
    AJTE.TextMatrix(8, 0) = "Sumas iguales"
    AJTE.TextMatrix(9, 0) = "Cuenta"
    AJTE.TextMatrix(10, 0) = "Subcuenta"
    AJTE.TextMatrix(11, 0) = "Nombre Cta-Subcta"
    AJTE.TextMatrix(12, 0) = "Parcial"
    AJTE.TextMatrix(13, 0) = "Debe"
    AJTE.TextMatrix(14, 0) = "Haber"
    AJTE.Row = 1: AJTE.Col = 1
    AJTE_RowColChange
    AJTE.GridColor = vbRed
End Sub

