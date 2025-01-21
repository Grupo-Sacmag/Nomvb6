Option Strict Off
Option Explicit On
Friend Class SalInteg
	Inherits System.Windows.Forms.Form
	'UPGRADE_ISSUE: Tipo declarativo no admitido: Matriz de cadenas con largo fijo. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
	Dim PrTnes(50) As String*64
	Dim Z1 As String
	Sub CatCpt()
		' ***************** Como se integra salario mas prestaciones integrables ********************
		PrTnes(1) = "Sueldo"
		' Sueldo
		PrTnes(2) = "Viaticos"
		' Viaticos
		PrTnes(3) = "Horas extras"
		' Horas extras
		PrTnes(4) = "Aguinaldo"
		' Aguinaldo(Dias)
		PrTnes(5) = "Prima dominical"
		' Prima dominical
		PrTnes(6) = "Prima Vacacional"
		' Prima Vacacional (Porcentaje)
		PrTnes(7) = "Reembolso de gastos medicos"
		' Reembolso de gastos medicos
		PrTnes(8) = "Fondo de ahorro"
		' Fondo de ahorro
		PrTnes(9) = "Caja de ahorro"
		' Caja de ahorro
		PrTnes(10) = "Vales de despensa"
		' Vales de despensa
		PrTnes(11) = "Ayuda gtos. funeral"
		' Ayuda gtos. funeral
		PrTnes(12) = "Contribuciones a cargo del trabajadores"
		' Contribuciones a cargo del trabajadores
		PrTnes(13) = "Premio de puntualidad"
		' Premio de puntualidad
		PrTnes(14) = "Prima de seguro de Vida"
		' Prima de seguro de Vida
		PrTnes(15) = "Seguro de gastos médicos mayores"
		' Seguro de gastos médicos mayores
		' Salario minimo
		PrTnes(16) = "Vales para restaurante"
		' Vales para restaurante
		PrTnes(17) = "Vales para gasolina"
		' Vales para gasolina
		PrTnes(18) = "Vales para ropa"
		' Vales para ropa
		PrTnes(19) = "Ayuda para renta"
		' Ayuda para renta
		PrTnes(20) = "Ayuda para articulos escolares"
		' Ayuda para articulos escolares
		PrTnes(21) = "Ayuda para anteojos"
		' Ayuda para anteojos
		PrTnes(22) = "Ayuda para transportes"
		' Ayuda para transportes
		PrTnes(23) = "Cuotas sindicales pagadas por el patron"
		' Cuotas sindicales pagadas por el patron
		PrTnes(24) = "Subsidios por incapacidad"
		' Subsidios por incapacidad
		PrTnes(25) = "Becas para trabajadores"
		' Becas para trabajadores
		PrTnes(26) = "Otros ingresos"
		' Otros ingresos
		PrTnes(27) = "Pagos efectuados por otros patrones"
		' ******************* Pagos efectuados por otros patrones ****************************
		PrTnes(28) = "Retenciones efectuadas por otros patrones"
		
	End Sub
	Sub ARRANKE()
		Dim r As Object
		Z1 = "#,##0.00;(#,##0.00)"
		SalIn1.Cols = 4
		SalIn1.Rows = 1 : SalIn1.Height = VB6.TwipsToPixelsY(SalIn1.get_RowHeight(0) + 100)
		' SalIn1.FixedCols = 0
		SalIn1.Row = 0 : SalIn1.Col = 0
		SalIn1.CellFontName = "Comic Sans MS" ' SalIn1.CellFontBold = True
		'UPGRADE_ISSUE: Form método SalInteg.TextWidth no se actualizó. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SalIn1.set_ColWidth(0, TextWidth(Prestciones.Conc) + 300) : SalIn1.CellBackColor = System.Drawing.Color.Cyan
		SalIn1.Text = "Concepto" : SalIn1.CellFontSize = 10 : SalIn1.CellAlignment = 4
		
		SalIn1.Width = VB6.TwipsToPixelsX(SalIn1.get_ColWidth(0))
		'UPGRADE_ISSUE: Form método SalInteg.TextWidth no se actualizó. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SalIn1.set_ColWidth(1, TextWidth(Z1))
		
		SalIn1.Col = 1 : SalIn1.CellFontName = "Comic Sans MS" ' SalIn1.CellFontBold = True
		SalIn1.CellBackColor = System.Drawing.Color.Cyan : SalIn1.Text = "Total" : SalIn1.CellFontSize = 10 : SalIn1.CellAlignment = 4
		SalIn1.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(SalIn1.Width) + SalIn1.get_ColWidth(1))
		SalIn1.Col = 2 : SalIn1.CellFontName = "Comic Sans MS" ' SalIn1.CellFontBold = True
		SalIn1.CellBackColor = System.Drawing.Color.Cyan : SalIn1.Text = "Gravado" : SalIn1.CellFontSize = 10 : SalIn1.CellAlignment = 4
		'UPGRADE_ISSUE: Form método SalInteg.TextWidth no se actualizó. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SalIn1.set_ColWidth(2, TextWidth(Z1))
		SalIn1.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(SalIn1.Width) + SalIn1.get_ColWidth(2))
		'UPGRADE_ISSUE: Form método SalInteg.TextWidth no se actualizó. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SalIn1.set_ColWidth(3, TextWidth(Z1))
		SalIn1.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(SalIn1.Width) + SalIn1.get_ColWidth(3) + 100)
		SalIn1.Col = 3 : SalIn1.CellFontName = "Comic Sans MS" ' SalIn1.CellFontBold = True
		SalIn1.CellBackColor = System.Drawing.Color.Cyan : SalIn1.Text = "Integrable" : SalIn1.CellFontSize = 10 : SalIn1.CellAlignment = 4
		'UPGRADE_ISSUE: Form método SalInteg.TextWidth no se actualizó. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SalIn1.set_ColWidth(3, TextWidth(Z1))
		CatCpt()
		For r = 1 To 28
			'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto r. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SalIn1.AddItem(Trim(PrTnes(r)) & Chr(9) & "" & Chr(9) & "" & Chr(9) & "")
			'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto r. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SalIn1.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(SalIn1.Height) + SalIn1.get_RowHeight(r))
		Next r
		' SalIn1.Height = SalIn1.Height
		For r = 1 To 28
			SalIn1.Col = 0
			'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto r. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SalIn1.Row = r
			' SalIn1.CellBackColor = vbCyan
			' SalIn1.CellAlignment = 7
			SalIn1.CellFontBold = True
		Next r
	End Sub
	
	Private Sub SalInteg_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		ARRANKE()
		SalInteg_Resize(Me, New System.EventArgs())
	End Sub
	
	'UPGRADE_WARNING: El evento SalInteg.Resize se puede desencadenar cuando se inicializa el formulario. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub SalInteg_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		If Me.WindowState <> 1 Then
			Me.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(SalIn1.Width) + 600)
			Me.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(SalIn1.Height) + 800)
		End If
	End Sub
End Class