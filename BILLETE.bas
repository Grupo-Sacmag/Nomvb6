Attribute VB_Name = "Module2"
Public Feria As String, gt_bi#, GTBI$
Dim B_ILL1$(10), B_ILL2$(10), B_ILL3$(10)
Dim BILL1, BILL2, BILL3, BILL4, BILL5, BILL6, BILL7, BILL8, BILL9
Sub Dinero()
   Rem entra valor gt_bi# y sale GTBI$
   GoSub moneda: GoTo saleferia
moneda:
5020 MONA$ = "/100 M.N.)": B_ILL1$(1) = "UN": B_ILL1$(2) = "DOS": B_ILL1$(3) = "TRES"
5030 B_ILL1$(4) = "CUATRO": B_ILL1$(5) = "CINCO": B_ILL1$(6) = "SEIS": B_ILL1$(7) = "SIETE"
5040 B_ILL1$(8) = "OCHO": B_ILL1$(9) = "NUEVE": B_ILL2$(1) = "DIEZ": B_ILL2$(2) = "VEINTE"
5050 B_ILL2$(3) = "TREINTA": B_ILL2$(4) = "CUARENTA": B_ILL2$(5) = "CINCUENTA"
5055 B_ILL2$(6) = "SESENTA": B_ILL2$(7) = "SETENTA": B_ILL2$(8) = "OCHENTA"
5060 B_ILL2$(9) = "NOVENTA"
5070 B_ILL3$(1) = "CIENTO": B_ILL3$(2) = "DOSCIENTOS": B_ILL3$(3) = "TRESCIENTOS"
5080 B_ILL3$(4) = "CUATROCIENTOS": B_ILL3$(5) = "QUINIENTOS": B_ILL3$(6) = "SEISCIENTOS"
5090 B_ILL3$(7) = "SETECIENTOS": B_ILL3$(8) = "OCHOCIENTOS": B_ILL3$(9) = "NOVECIENTOS"
5100 PESO$ = " PESOS "
5110 BILL1 = 0: BILL2 = 0: BILL3 = 0: BILL4 = 0: BILL5 = 0: BILL6 = 0: BILL7 = 0: BILL8 = 0: BILL9 = 0
5120 If gt_bi# >= 10 Then GoTo 5180
5130 If gt_bi# < 1 Then NBIL$ = " ": Return
5140 NOBI = 0: GoSub 5670: BI_LL1$ = Mid$(GTBI$, 2, 1): BILL1 = Val(BI_LL1$): GoSub 5710
5150 GoSub 5680: If BILL1 = 1 Then BI_LL1$ = BI_LL1$ + " PESO": GoTo 5170
5160 BI_LL1$ = BI_LL1$ + PESO$
5170 NBIL$ = "(" + BI_LL1$ + " " + MNBI$ + MONA$: Return
5180 If gt_bi# >= 100 Then GoTo 5220
5190 NOBI = 0: GoSub 5670: BI_LL1$ = Mid$(GTBI$, 3, 1): BILL1 = Val(BI_LL1$): GoSub 5710
5200 BI_LL2$ = Mid$(GTBI$, 2, 1): BILL2 = Val(BI_LL2$): GoSub 5730: GoSub 5800: GoSub 5680
5210 NBIL$ = "(" + BI_LL2$ + PESO$ + MNBI$ + MONA$: Return
5220 If gt_bi# >= 1000 Then GoTo 5260
5230 NOBI = 0: GoSub 5670: YBIL$ = Mid$(GTBI$, 2, 3): GoSub 5970: GoSub 6010: GoSub 5680
5240 If BI_LL3$ = "CIEN " Then BI_LL3$ = "CIEN"
5250 NBIL$ = "(" + BI_LL3$ + PESO$ + MNBI$ + MONA$: Return
5260 If gt_bi# >= 10000 Then GoTo 5300
5270 NOBI = 0: GoSub 5670: YBIL$ = Mid$(GTBI$, 3, 3): GoSub 5970: GoSub 6010
5280 BI_LL4$ = Mid$(GTBI$, 2, 1): BILL4 = Val(BI_LL4$): GoSub 5770: BI_LL4$ = BI_LL4$ + " MIL"
5290 GoSub 5680: NBIL$ = "(" + BI_LL4$ + " " + BI_LL3$ + PESO$ + MNBI$ + MONA$: Return
5300 If gt_bi# >= 100000! Then GoTo 5340
5310 NOBI = 0: GoSub 5670: YBIL$ = Mid$(GTBI$, 4, 3): GoSub 5970: GoSub 6010: NO_BI$ = BI_LL3$
5320 YBIL$ = Mid$(GTBI$, 1, 3): GoSub 5970: GoSub 6010
5330 GoSub 5680: NBIL$ = "(" + BI_LL2$ + " MIL " + NO_BI$ + PESO$ + MNBI$ + MONA$: Return
5340 If gt_bi# >= 1000000! Then GoTo 5380
5350 NOBI = 0: GoSub 5670: YBIL$ = Mid$(GTBI$, 5, 3): GoSub 5970: GoSub 6010: NO_BI$ = BI_LL3$
5360 YBIL$ = Mid$(GTBI$, 2, 3): GoSub 5970: GoSub 6010
5370 GoSub 5680: NBIL$ = "(" + BI_LL3$ + " MIL " + NO_BI$ + PESO$ + MNBI$ + MONA$: Return
5380 If gt_bi# >= 10000000# Then GoTo 5470
5390 NOBI = 0: GoSub 5670: YBIL$ = Mid$(GTBI$, 6, 3): GoSub 5970: GoSub 6010: NO_BI$ = BI_LL3$
5400 YBIL$ = Mid$(GTBI$, 3, 3): GoSub 5970: GoSub 6010
5410 If BI_LL3$ = "" Then GoTo 5430
5420 NO_BI$ = BI_LL3$ + " MIL " + NO_BI$
5430 BI_LL4$ = Mid$(GTBI$, 2, 1): BILL4 = Val(BI_LL4$): GoSub 5770
5440 If BILL4 = 1 And NO_BI$ = "" Then BI_LL4$ = BI_LL4$ + " MILLON DE ": GoTo 5460
5450 If BILL4 = 1 Then BI_LL4$ = BI_LL4$ + " MILLON " Else BI_LL4$ = BI_LL4$ + " MILLONES "
5460 GoSub 5680: NBIL$ = "(" + BI_LL4$ + NO_BI$ + PESO$ + MNBI$ + MONA$: Return
5470 If gt_bi# >= 100000000# Then GoTo 5530
5480 NOBI = 0: GoSub 5670: GoSub 5680: YBIL$ = Mid$(GTBI$, 7, 3): GoSub 5970: GoSub 6010
5490 NO_BI$ = BI_LL3$: YBIL$ = Mid$(GTBI$, 4, 3): GoSub 5970: GoSub 6010: If BI_LL3$ = "" Then GoTo 5510
5500 NO_BI$ = BI_LL3$ + " MIL " + NO_BI$
5510 NOBI = 2: YBIL$ = Mid$(GTBI$, 1, 3): GoSub 5970: GoSub 5800: BI_LL2$ = BI_LL2$ + " MILLONES "
5520 GoSub 5680: NBIL$ = "(" + BI_LL2$ + NO_BI$ + PESO$ + MNBI$ + MONA$: Return
5530 If gt_bi# >= 1000000000# Then GoTo 5590
5540 NOBI = 0: GoSub 5670: YBIL$ = Mid$(GTBI$, 8, 3): GoSub 5970: GoSub 6010: NO_BI$ = BI_LL3$
5550 YBIL$ = Mid$(GTBI$, 5, 3): GoSub 5970: GoSub 6010: If BI_LL3$ = "" Then GoTo 5570
5560 NO_BI$ = BI_LL3$ + " MIL " + NO_BI$
5570 YBIL$ = Mid$(GTBI$, 2, 3): GoSub 5970: GoSub 6010: BI_LL3$ = BI_LL3$ + " MILLONES " + NO_BI$
5580 GoSub 5680: NBIL$ = "(" + BI_LL3$ + PESO$ + MNBI$ + MONA$: Return
5590 If gt_bi# >= 10000000000# Then Return
5600 NOBI = 0: GoSub 5670: YBIL$ = Mid$(GTBI$, 9, 3): GoSub 5970: GoSub 6010: NO_BI$ = BI_LL3$
5610 YBIL$ = Mid$(GTBI$, 6, 3): GoSub 5970: GoSub 6010: If BI_LL3$ = "" Then GoTo 5630
5620 NO_BI$ = BI_LL3$ + " MIL " + NO_BI$
5630 YBIL$ = Mid$(GTBI$, 3, 3): GoSub 5970: GoSub 6010: BI_LL3$ = BI_LL3$ + " MILLONES " + NO_BI$
5640 BI_LL4$ = Mid$(GTBI$, 2, 1): BILL4 = Val(BI_LL4$): GoSub 5770: BI_LL4$ = "MIL"
5650 If BILL4 = 1 And BI_LL3$ = " MILLONES " Then PESO$ = " PESOS "
5660 GoSub 5680: NBIL$ = "(" + BI_LL4$ + BI_LL3$ + PESO$ + MNBI$ + MONA$: Return
5670 FRB = gt_bi# - Int(gt_bi#): FRB = 0.01 * Int(FRB * 100 + 0.5): GTBI$ = Str$(Int(gt_bi#)) ': RETURN
5680 If FRB = 0 Then MNBI$ = "00": Return
5690 MNBI$ = LTrim$(Str$(Int((FRB * 100)))): Return
5700 'MNBI$ = "0" + LTRIM$(MNBI$, 2, 1): RETURN
5710 If BILL1 = 0 Then BI_LL1$ = "": Return
5720 BI_LL1$ = B_ILL1$(BILL1): Return
5730 If BILL2 = 0 Then BI_LL2$ = "": Return
5740 BI_LL2$ = B_ILL2$(BILL2): Return
5750 If BILL3 = 0 Then BI_LL3$ = "": Return
5760 BI_LL3$ = B_ILL3$(BILL3): Return
5770 If BILL4 = 0 Then BI_LL4$ = "": Return
5780 BI_LL4$ = B_ILL1$(BILL4): Return
5790 Return
5800 If BILL2 = 0 And BILL1 <> 0 Then BI_LL2$ = BI_LL1$: Return
5810 If BILL2 = 1 And BILL1 <> 0 Then GoSub 5860: Return
5820 If BILL2 = 2 And BILL1 <> 0 Then BI_LL2$ = "VEINTI" + BI_LL1$: Return
5830 If BILL2 <> 0 And BILL1 = 0 Then BI_LL2$ = BI_LL2$: Return
5840 If BILL2 <> 0 And BILL1 <> 0 Then BI_LL2$ = BI_LL2$ + " Y " + BI_LL1$: Return
5850 BI_LL2$ = "": Return
5860 If BILL2 = 1 And BILL1 = 1 Then BI_LL2$ = "ONCE": GoTo 5960
5870 If BILL2 = 1 And BILL1 = 2 Then BI_LL2$ = "DOCE": GoTo 5960
5880 If BILL2 = 1 And BILL1 = 3 Then BI_LL2$ = "TRECE": GoTo 5960
5890 If BILL2 = 1 And BILL1 = 4 Then BI_LL2$ = "CATORCE": GoTo 5960
5900 If BILL2 = 1 And BILL1 = 5 Then BI_LL2$ = "QUINCE": GoTo 5960
5910 If BILL2 = 1 And BILL1 = 6 Then BI_LL2$ = "DIECISEIS": GoTo 5960
5920 If BILL2 = 1 And BILL1 = 7 Then BI_LL2$ = "DIECISIETE": GoTo 5960
5930 If BILL2 = 1 And BILL1 = 8 Then BI_LL2$ = "DIECIOCHO": GoTo 5960
5940 If BILL2 = 1 And BILL1 = 9 Then BI_LL2$ = "DIECINUEVE": GoTo 5960
5950 Return
5960 BI_LL1$ = "": Return
5970 BILL1 = Val(Mid$(YBIL$, 3, 1)): GoSub 5710: If NOBI = 1 Then Return
5980 BILL2 = Val(Mid$(YBIL$, 2, 1)): GoSub 5730: If NOBI = 2 Then Return
5990 BILL3 = Val(Mid$(YBIL$, 1, 1)): GoSub 5750: If NOBI = 3 Then Return
6000 Return
6010 If BILL3 = 1 And BILL2 = 0 And BILL1 = 0 Then BI_LL3$ = "CIEN": Return
6020 GoSub 5800: If BILL3 = 0 Then BI_LL3$ = BI_LL2$: Return
6025 If BILL3 <> 0 And BI_LL2$ = "" Then Return
6030 If BILL3 <> 0 Then BI_LL3$ = BI_LL3$ + " " + BI_LL2$: Return
6040 BI_LL3$ = "": Return
saleferia:
Feria = NBIL$

End Sub

