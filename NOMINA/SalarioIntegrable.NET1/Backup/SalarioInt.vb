Option Strict Off
Option Explicit On
Module SalarioInt
	Structure Presta
		'UPGRADE_WARNING: El tamaño de la cadena de longitud fija debe caber en el búfer. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(64),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=64)> Public Conc() As Char
		Dim Tot As Decimal
		Dim Acum As Decimal
		Dim Integ As Decimal
	End Structure
	Public Prestciones As Presta
End Module