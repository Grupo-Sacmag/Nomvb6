Attribute VB_Name = "Module3"
Type acu
     dias As Currency
     ispt As Currency
     crdsal As Currency
     imss As Currency
     sueldo As Currency
     hs_ext As Currency
     viaticos As Currency
     pvac As Currency
     otras As Currency
     exentos As Currency
     aguin As Currency
     ptu As Currency
     prestamos As Currency
     fonacot As Currency
     telefono As Currency
     otraded As Currency
  End Type
  Type co_Acu
     No_arch As String * 64
     largo As Integer
  End Type
  
  Public acumulado As acu, sdo As Currency
  Public N_ormal As Integer
  Public contr_Acu As co_Acu
  Public AR1 As String, nq_1 As Integer, nq_2 As Integer
  Public nq_3, inter_clave As Integer, nominilla(26), t As Integer
  Sub Acumu_lar()
        Form10.ProgressBar1.Max = nq_1
        Form10.ProgressBar1.Min = 0
        Form10.ProgressBar1.Value = comp_ara
        For r = 1 To nq_1: Get 2, r, acumulado
             Get 1, r, nomina
             Rem If LCase(Mid(AR1, 1, 3)) = LCase("AGU") Or LCase(Mid(AR1, 1, 3)) = LCase("COM") Or LCase(Mid(AR1, 1, 3)) = LCase("PTU") Or LCase(Mid(AR1, 1, 3)) = LCase("GRA") Then
             
              Rem sdo = nomina.sueldo + nomina.viaticos + nomina.otras
              
              Rem If LCase(Mid(AR1, 1, 3)) = LCase("PTU") Then
                      
                      Rem acumulado.ispt = acumulado.ispt + nomina.ispt
                      Rem acumulado.ptu = acumulado.ptu + sdo
                      Rem sdo = 0
                      Rem Put #2, r, acumulado
                      Rem Else
                      Rem acumulado.ispt = acumulado.ispt + nomina.ispt
                      Rem acumulado.aguin = acumulado.aguin + sdo
                      Rem acumulado.pvac = acumulado.pvac + nomina.pvac
                      Rem acumulado.exentos = acumulado.exentos + nomina.exentos
                      Rem sdo = 0
                      Rem Put #2, r, acumulado
                Rem End If
              Rem Else
                acumulado.dias = acumulado.dias + nomina.dias
                acumulado.ispt = acumulado.ispt + nomina.ispt
                acumulado.crdsal = acumulado.crdsal + nomina.crdsal
                acumulado.imss = acumulado.imss + nomina.imss
                acumulado.sueldo = acumulado.sueldo + nomina.sueldo
                Total# = Total# + acumulado.sueldo
                acumulado.hs_ext = acumulado.hs_ext + nomina.hs_nor + nomina.hs_dbl + nomina.hs_tri
                Total# = Total# + acumulado.hs_ext
                acumulado.viaticos = acumulado.viaticos + nomina.viaticos
                Total# = Total# + acumulado.viaticos
                acumulado.pvac = acumulado.pvac + nomina.pvac
                Total# = Total# + acumulado.pvac
                acumulado.otras = acumulado.otras + nomina.otras
                Total# = Total# + acumulado.otras
                acumulado.aguin = acumulado.aguin + nomina.aguin
                Total# = Total# + acumulado.aguin
                acumulado.ptu = acumulado.ptu + nomina.ptu
                Total# = Total# + acumulado.ptu
                acumulado.exentos = acumulado.exentos + nomina.exentos
                Total# = Total# + acumulado.exentos
                acumulado.prestamos = acumulado.prestamos + nomina.prestamos
                acumulado.fonacot = acumulado.fonacot + nomina.fonacot
                acumulado.telefono = acumulado.telefono + nomina.telefono
                acumulado.otraded = acumulado.otraded + nomina.otraded
                
                Put #2, r, acumulado
                'PRINT USING "###,###,###.##"; total#
           Rem End If
           Form10.ProgressBar1.Value = r
        Next r
        
End Sub

  
Sub blanc0()
   For r = 1 To nq_3: Get 2, r, acumulado
   acumulado.dias = 0
   acumulado.ispt = 0
   acumulado.crdsal = 0
   acumulado.imss = 0
   acumulado.sueldo = 0
   acumulado.hs_ext = 0
   acumulado.viaticos = 0
   acumulado.pvac = 0
   acumulado.otras = 0
   acumulado.exentos = 0
   acumulado.aguin = 0
   acumulado.ptu = 0
   acumulado.prestamos = 0
   acumulado.fonacot = 0
   acumulado.telefono = 0
   acumulado.otraded = 0
 Put 2, r, acumulado
   Next r


End Sub

Sub convierte()
obra(1) = maestro.O_1: obra(2) = maestro.O_2: obra(3) = maestro.O_3: obra(4) = maestro.O_4
obra(5) = maestro.O_5: obra(6) = maestro.O_6: obra(7) = maestro.O_7: obra(8) = maestro.O_8
obra(9) = maestro.O_9: obra(10) = maestro.O_10: obra(11) = maestro.O_11: obra(12) = maestro.O_12
obra(13) = maestro.O_13: obra(14) = maestro.O_14: obra(15) = maestro.O_15: obra(16) = maestro.O_16
obra(17) = maestro.O_17: obra(18) = maestro.O_18: obra(19) = maestro.O_19: obra(20) = maestro.O_20
porcentaje(1) = maestro.por_1: porcentaje(2) = maestro.por_2: porcentaje(3) = maestro.por_3: porcentaje(4) = maestro.por_4
porcentaje(5) = maestro.por_5: porcentaje(6) = maestro.por_6: porcentaje(7) = maestro.por_7: porcentaje(8) = maestro.por_8
porcentaje(9) = maestro.por_9: porcentaje(10) = maestro.por_10: porcentaje(11) = maestro.por_11: porcentaje(12) = maestro.por_12
porcentaje(13) = maestro.por_13: porcentaje(14) = maestro.por_14: porcentaje(15) = maestro.por_15: porcentaje(16) = maestro.por_16
porcentaje(17) = maestro.por_17: porcentaje(18) = maestro.por_18: porcentaje(19) = maestro.por_19: porcentaje(20) = maestro.por_20

   
End Sub
Sub grabamaestro()
    maestro.O_1 = obra(1): maestro.O_2 = obra(2): maestro.O_3 = obra(3): maestro.O_4 = obra(4)
    maestro.O_5 = obra(5): maestro.O_6 = obra(6): maestro.O_7 = obra(7): maestro.O_8 = obra(8)
    maestro.O_9 = obra(9): maestro.O_10 = obra(10): maestro.O_11 = obra(11): maestro.O_12 = obra(12)
    maestro.O_13 = obra(13): maestro.O_14 = obra(14): maestro.O_15 = obra(15): maestro.O_16 = obra(16)
    maestro.O_17 = obra(17): maestro.O_18 = obra(18): maestro.O_19 = obra(19): maestro.O_20 = obra(20)
    maestro.por_1 = porcentaje(1): maestro.por_2 = porcentaje(2): maestro.por_3 = porcentaje(3): maestro.por_4 = porcentaje(4)
    maestro.por_5 = porcentaje(5): maestro.por_6 = porcentaje(6): maestro.por_7 = porcentaje(7): maestro.por_8 = porcentaje(8)
    maestro.por_9 = porcentaje(9): maestro.por_10 = porcentaje(10): maestro.por_11 = porcentaje(11): maestro.por_12 = porcentaje(12)
    maestro.por_13 = porcentaje(13): maestro.por_14 = porcentaje(14): maestro.por_15 = porcentaje(15): maestro.por_16 = porcentaje(16)
    maestro.por_17 = porcentaje(17): maestro.por_18 = porcentaje(18): maestro.por_19 = porcentaje(19): maestro.por_20 = porcentaje(20)
End Sub

