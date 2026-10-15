Attribute VB_Name = "modGridNomina"
'==============================================================
' Indices de columnas de ConNom1 (MSFlexGrid) - UNICO lugar donde
' se definen. En todo el codigo usar estas constantes, nunca numeros.
'
' Layout fijo para TODOS los tipos de nomina. Las columnas 9 y 10
' solo se ven en Liquidacion/Finiquito (ancho 0 en las demas).
'
' Equivalencia con el layout anterior (0 a 8 no cambian):
'   viejo 9  Otras      -> nuevo 11      viejo 17 Pension  -> nuevo 19
'   viejo 10 Exenta     -> nuevo 12      viejo 18 Infonavit-> nuevo 20
'   viejo 11 Tot.Ingr.  -> nuevo 13      viejo 19 Tot.Ded. -> nuevo 21
'   viejo 12 Ispt       -> nuevo 14      viejo 20 Neto     -> nuevo 22
'   viejo 13 Subsidio   -> nuevo 15      viejo 21 flag ISR -> nuevo 23
'   viejo 14 Imss       -> nuevo 16      viejo 22 flag IMSS-> nuevo 24
'   viejo 15 Prestamos  -> nuevo 17      viejo 23 flag esp.-> nuevo 25
'   viejo 16 Fonacot    -> nuevo 18      viejo 24 Banamex  -> nuevo 26
'==============================================================
Public Const C_NUM = 0
Public Const C_NOMBRE = 1
Public Const C_DIAS = 2
Public Const C_SUELDO = 3
Public Const C_CONC4 = 4        ' hs.Norm.  / Compens.
Public Const C_CONC5 = 5        ' hs.Dobles / Aguinaldo / P.Antig.
Public Const C_CONC6 = 6        ' hs.Triples / Ptu / Indemn.
Public Const C_VIAT = 7         ' O.F. / BONO
Public Const C_PVAC = 8         ' P.Vacac. / Sdo.Vac.
Public Const C_AGUILIQ = 9      ' Aguinaldo (solo liquidacion)  -> nomina.hsnor
Public Const C_PVACLIQ = 10     ' P.Vac.Liq. (solo liquidacion) -> nomina.hstri
Public Const C_OTRAS = 11
Public Const C_EXENTA = 12
Public Const C_TOTING = 13
Public Const C_ISPT = 14
Public Const C_SUBS = 15
Public Const C_IMSS = 16
Public Const C_PREST = 17
Public Const C_FONACOT = 18
Public Const C_PENSION = 19
Public Const C_INFONAVIT = 20
Public Const C_TOTDED = 21
Public Const C_NETO = 22
Public Const C_FLAGISR = 23     ' oculta: 1 = ISR capturado a mano
Public Const C_FLAGIMSS = 24    ' oculta: 1 = IMSS capturado a mano
Public Const C_FLAGESP = 25     ' oculta: 1 = renglon de calculo especial
Public Const C_BANCO = 26
Public Const C_NUMCOLS = 27     ' total de columnas (0 a 26)

Public Function EsLiq() As Boolean
    EsLiq = (g_TipoNominaActiva = tnLiquidacionFiniquito)
End Function
