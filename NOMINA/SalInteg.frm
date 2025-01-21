VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form SalInteg 
   Caption         =   "SalarioInteg"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   7830
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid SalIn1 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   12303
      _Version        =   393216
   End
End
Attribute VB_Name = "SalInteg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PrTnes(50) As String * 64, Z1 As String
Sub CatCpt()
    Rem ***************** Como se integra salario mas prestaciones integrables ********************
    PrTnes(1) = "Sueldo"
    Rem Sueldo
    PrTnes(2) = "Viaticos"
    Rem Viaticos
    PrTnes(3) = "Horas extras"
    Rem Horas extras
    PrTnes(4) = "Aguinaldo"
    Rem Aguinaldo(Dias)
    PrTnes(5) = "Prima dominical"
    Rem Prima dominical
    PrTnes(6) = "Prima Vacacional"
    Rem Prima Vacacional (Porcentaje)
    PrTnes(7) = "Reembolso de gastos medicos"
    Rem Reembolso de gastos medicos
    PrTnes(8) = "Fondo de ahorro"
    Rem Fondo de ahorro
    PrTnes(9) = "Caja de ahorro"
    Rem Caja de ahorro
    PrTnes(10) = "Vales de despensa"
    Rem Vales de despensa
    PrTnes(11) = "Ayuda gtos. funeral"
    Rem Ayuda gtos. funeral
    PrTnes(12) = "Contribuciones a cargo del trabajadores"
    Rem Contribuciones a cargo del trabajadores
    PrTnes(13) = "Premio de puntualidad"
    Rem Premio de puntualidad
    PrTnes(14) = "Prima de seguro de Vida"
    Rem Prima de seguro de Vida
    PrTnes(15) = "Seguro de gastos médicos mayores"
    Rem Seguro de gastos médicos mayores
    Rem Salario minimo
    PrTnes(16) = "Vales para restaurante"
    Rem Vales para restaurante
    PrTnes(17) = "Vales para gasolina"
    Rem Vales para gasolina
    PrTnes(18) = "Vales para ropa"
    Rem Vales para ropa
    PrTnes(19) = "Ayuda para renta"
    Rem Ayuda para renta
    PrTnes(20) = "Ayuda para articulos escolares"
    Rem Ayuda para articulos escolares
    PrTnes(21) = "Ayuda para anteojos"
    Rem Ayuda para anteojos
    PrTnes(22) = "Ayuda para transportes"
    Rem Ayuda para transportes
    PrTnes(23) = "Cuotas sindicales pagadas por el patron"
    Rem Cuotas sindicales pagadas por el patron
    PrTnes(24) = "Subsidios por incapacidad"
    Rem Subsidios por incapacidad
    PrTnes(25) = "Becas para trabajadores"
    Rem Becas para trabajadores
    PrTnes(26) = "Otros ingresos"
    Rem Otros ingresos
    PrTnes(27) = "Pagos efectuados por otros patrones"
    Rem ******************* Pagos efectuados por otros patrones ****************************
    PrTnes(28) = "Retenciones efectuadas por otros patrones"

End Sub
Sub ARRANKE()
   Z1 = "#,##0.00;(#,##0.00)"
   SalIn1.Cols = 4
   SalIn1.Rows = 1: SalIn1.Height = SalIn1.RowHeight(0) + 100
   Rem SalIn1.FixedCols = 0
   SalIn1.Row = 0: SalIn1.Col = 0
   SalIn1.CellFontName = "Comic Sans MS": Rem SalIn1.CellFontBold = True
   SalIn1.ColWidth(0) = TextWidth(Prestciones.Conc) + 300: SalIn1.CellBackColor = vbCyan:
    SalIn1.Text = "Concepto":  SalIn1.CellFontSize = 10: SalIn1.CellAlignment = 4
   
   SalIn1.Width = SalIn1.ColWidth(0)
   SalIn1.ColWidth(1) = TextWidth(Z1)
   
   SalIn1.Col = 1: SalIn1.CellFontName = "Comic Sans MS": Rem SalIn1.CellFontBold = True
   SalIn1.CellBackColor = vbCyan: SalIn1.Text = "Total": SalIn1.CellFontSize = 10: SalIn1.CellAlignment = 4
   SalIn1.Width = SalIn1.Width + SalIn1.ColWidth(1)
   SalIn1.Col = 2: SalIn1.CellFontName = "Comic Sans MS": Rem SalIn1.CellFontBold = True
   SalIn1.CellBackColor = vbCyan:  SalIn1.Text = "Gravado": SalIn1.CellFontSize = 10: SalIn1.CellAlignment = 4
   SalIn1.ColWidth(2) = TextWidth(Z1)
   SalIn1.Width = SalIn1.Width + SalIn1.ColWidth(2)
   SalIn1.ColWidth(3) = TextWidth(Z1)
   SalIn1.Width = SalIn1.Width + SalIn1.ColWidth(3) + 100
   SalIn1.Col = 3: SalIn1.CellFontName = "Comic Sans MS": Rem SalIn1.CellFontBold = True
   SalIn1.CellBackColor = vbCyan:  SalIn1.Text = "Integrable": SalIn1.CellFontSize = 10: SalIn1.CellAlignment = 4
   SalIn1.ColWidth(3) = TextWidth(Z1)
   CatCpt
   For r = 1 To 28
       SalIn1.AddItem Trim(PrTnes(r)) & Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
       SalIn1.Height = SalIn1.Height + SalIn1.RowHeight(r)
   Next r
   Rem SalIn1.Height = SalIn1.Height
   For r = 1 To 28
       SalIn1.Col = 0: SalIn1.Row = r
       Rem SalIn1.CellBackColor = vbCyan
       Rem SalIn1.CellAlignment = 7
       SalIn1.CellFontBold = True
   Next r
End Sub

Private Sub Form_Load()
    ARRANKE
    Form_Resize
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
       Me.Width = SalIn1.Width + 600
       Me.Height = SalIn1.Height + 800
    End If
End Sub
