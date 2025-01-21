VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form TREXC 
   Caption         =   "TRASPASO NOMINA BANAMEX EN EXCEL"
   ClientHeight    =   5925
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   10125
   Icon            =   "TREXC.frx":0000
   LinkTopic       =   "Form9"
   ScaleHeight     =   5925
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Ex 
      Height          =   4812
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   9372
      _ExtentX        =   16536
      _ExtentY        =   8493
      _Version        =   393216
   End
   Begin VB.Menu Arc 
      Caption         =   "&Archivo"
      Begin VB.Menu ArcSale 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu Edt 
      Caption         =   "&Edicion"
      Begin VB.Menu EdSel 
         Caption         =   "&Seleccionar todo"
      End
      Begin VB.Menu EdSep1 
         Caption         =   "-"
      End
      Begin VB.Menu EdCopiar 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "TREXC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Col_Def()
   Ex.Clear: Ex.Rows = 1: Ex.Cols = 10: Ex.FixedCols = 0
   Ex.Text = "No."
   Ex.Col = 0: Ex.CellAlignment = 4: Ex.ColWidth(0) = 1800: Ex.Text = "Num.Cta"
   Ex.Col = 1: Ex.CellAlignment = 4: Ex.ColWidth(1) = 1200: Ex.Text = "Nombre"
   Ex.Col = 2: Ex.CellAlignment = 4: Ex.ColWidth(2) = 1200: Ex.Text = "Apell.Paterno"
   Ex.Col = 3: Ex.CellAlignment = 4: Ex.ColWidth(3) = 1200: Ex.Text = "Apell.Materno"
   Ex.Col = 4: Ex.CellAlignment = 4: Ex.ColWidth(4) = 800: Ex.Text = "Importe"
   Ex.Col = 5: Ex.CellAlignment = 4: Ex.ColWidth(5) = 800: Ex.Text = "Trabajador"
   Ex.Col = 6: Ex.CellAlignment = 4: Ex.ColWidth(6) = 800: Ex.Text = "Ref.Alfanum"
   Ex.Col = 7: Ex.CellAlignment = 4: Ex.ColWidth(7) = 800: Ex.Text = "Concepto Pago"
   Ex.Col = 8: Ex.CellAlignment = 4: Ex.ColWidth(8) = 800: Ex.Text = "MIsmoDia"
   Ex.Col = 9: Ex.CellAlignment = 4: Ex.ColWidth(9) = 600: Ex.Text = "Cons."
   
   
End Sub

Private Sub EdCopiar_Click()
 Dim Temporal1
 Clipboard.Clear
   
   difer = Ex.RowSel - Ex.Row
   For i = Ex.Row To Ex.RowSel
      
      For f = Ex.Col To Ex.ColSel
            Temporal1 = Temporal1 + Ex.TextMatrix(i, f)
            If f < Ex.ColSel Then
                Temporal1 = Temporal1 & Chr(9)
            End If
      Next f
      Temporal1 = Temporal1 & Chr(13) & Chr(10)
      
   Next i
    Clipboard.SetText Temporal1

End Sub

Private Sub EdSel_Click()
Clipboard.Clear
    Ex.Row = 1: Ex.Col = 0
    Ex.RowSel = Ex.Rows - 1
    Ex.ColSel = Ex.Cols - 2
End Sub

Private Sub Form_Load()
    Col_Def
End Sub
