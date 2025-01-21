Attribute VB_Name = "ModuloAbril"
Public VARIABLELARGA
Sub MdAbr()
    '0 1 SUCURSAL
        NOMCF2.NOMCF.TextMatrix(I7, 0) = 1:
    '1 2 FOLIO
        NOMCF2.NOMCF.TextMatrix(I7, 1) = Folio
    '2 3 SERIE
        NOMCF2.NOMCF.TextMatrix(I7, 2) = serie:
    '3 4 NOMBRE
        Get 2, NumerodePersonal, personal:
        MiFechaAlta = Trim(personal.fal)
        Nombrey = Trim(personal.nom) + " " + Trim(personal.ape1) + " " + Trim(personal.ape2)
        NOMCF2.NOMCF.TextMatrix(I7, 3) = Nombrey:
        Get 7, NumerodePersonal, Empleado_1
    '4 5 DIRECCION
        NOMCF2.NOMCF.TextMatrix(I7, 4) = Trim(Empleado_1.Direccion):
    '5 6 COLONIA
        NOMCF2.NOMCF.TextMatrix(I7, 5) = Trim(Empleado_1.Colonia):
    '6 7 CIUDAD
        NOMCF2.NOMCF.TextMatrix(I7, 6) = Trim(Empleado_1.Ciudad):
    '7 8 ESTADO
        NOMCF2.NOMCF.TextMatrix(I7, 7) = Trim(Empleado_1.Estado):
    '8 9 DELEGACION
        NOMCF2.NOMCF.TextMatrix(I7, 8) = Trim(Empleado_1.Delegacion):
    '9 10 CP
        NOMCF2.NOMCF.TextMatrix(I7, 9) = Trim(Empleado_1.Cpostal):
    '10 11 RFC
        NOMCF2.NOMCF.TextMatrix(I7, 10) = Trim(personal.RFC):
    '11 12 PAIS
        NOMCF2.NOMCF.TextMatrix(I7, 11) = "MEXICO":
    '12 13 CORREO
        NOMCF2.NOMCF.TextMatrix(I7, 12) = Trim(Empleado_1.correo):
    '13 14 OBSERVACIONES
        NOMCF2.NOMCF.TextMatrix(I7, 13) = ""
    '14 15 MONEDA
        NOMCF2.NOMCF.TextMatrix(I7, 14) = "PESOS"
    '15 16 TIPOCAMBIO
        NOMCF2.NOMCF.TextMatrix(I7, 15) = 1
'******************************************LLAMADO AL MODULO ***********************************************
    reng
    '16 17 TOTAL
        NOMCF2.NOMCF.TextMatrix(I7, 16) = Format(T_neto, "#,##0.00")
    '17 18 SUBTOTAL
        NOMCF2.NOMCF.TextMatrix(I7, 17) = Format(t_per, "#,##0.00")
    '18 19 DESCUENTO
        NOMCF2.NOMCF.TextMatrix(I7, 18) = Format(t_ded, "#,##0.00")
    '19 20 TOTALGRAVADOPERCEPCIONES
        NOMCF2.NOMCF.TextMatrix(I7, 19) = Format(t_grav, "#,##0.00")
    '20 21 TOTALEXENTOPERCEPCIONES
        NOMCF2.NOMCF.TextMatrix(I7, 20) = Format(t_ext, "#,##0.00")
    '21 22 TOTALPERCEPCIONES
        NOMCF2.NOMCF.TextMatrix(I7, 21) = Format(t_per, "#,##0.00")
    '22 23 TOTALDEDUCCIONES
        NOMCF2.NOMCF.TextMatrix(I7, 22) = Format(t_ded, "#,##0.00")
    '23 24 TOTALOTROSPAGOS
        NOMCF2.NOMCF.TextMatrix(I7, 23) = Format(t_oi, "#,##0.00")
    '24 25 TOTALSUELDOS
        NOMCF2.NOMCF.TextMatrix(I7, 24) = Format(sue3, "#,##0.00")
    '25 26 TOTALSEPARACIONINDEMNIZACION
        NOMCF2.NOMCF.TextMatrix(I7, 25) = Format(0, "#,##0.00")
    '26 27 TOTALJUBILACIONPENSIONRETIRO
        NOMCF2.NOMCF.TextMatrix(I7, 26) = Format(0, "#,##0.00")
    '27 28 TOTALOTRASDEDUCCIONES
        NOMCF2.NOMCF.TextMatrix(I7, 27) = Format(t_oded, "#,##0.00")
    '28 29 TOTALIMPUESTOSRETENIDOS
        NOMCF2.NOMCF.TextMatrix(I7, 28) = Format(isr12, "#,##0.00")
    '29 30 VALORUNITARIO
        NOMCF2.NOMCF.TextMatrix(I7, 29) = Format(t_per, "#,##0.00")
    '30 31 IMPORTE
        NOMCF2.NOMCF.TextMatrix(I7, 30) = Format(T_neto, "#,##0.00")
    Aletra (T_neto)
    '31 32 TOTALCONLETRA
        NOMCF2.NOMCF.TextMatrix(I7, 31) = LetrVal
    '32 33 TIPODENOMINA
        NOMCF2.NOMCF.TextMatrix(I7, 32) = "O"
    '33 34 SINDICALIZADO
        NOMCF2.NOMCF.TextMatrix(I7, 33) = "No"
    '34 35 METODOPAGO
        NOMCF2.NOMCF.TextMatrix(I7, 34) = Met_pago
    '35 36 LUGAREXPEDICION
        NOMCF2.NOMCF.TextMatrix(I7, 35) = "DIF"
    '36 37 REGIMEN
        NOMCF2.NOMCF.TextMatrix(I7, 36) = "601"
    '37 38 NUMCTAPAG
        NOMCF2.NOMCF.TextMatrix(I7, 37) = CStr(Form8.ConNom1.TextMatrix(I7, 24))
    '38 39 REGISTROPATRONAL
        NOMCF2.NOMCF.TextMatrix(I7, 38) = Reg_Patr
    '39 40 NUMEMPLEADO
        NOMCF2.NOMCF.TextMatrix(I7, 39) = NumerodePersonal
    '40 41 CURP
        NOMCF2.NOMCF.TextMatrix(I7, 40) = Trim(Otros_Rgtros.curp)
    '41 42 TIPOREGIMEN
        NOMCF2.NOMCF.TextMatrix(I7, 41) = "02 SUELDOS"
    '42 43 NUMSEGURIDADSOCIAL
        NOMCF2.NOMCF.TextMatrix(I7, 42) = Trim(personal.imss)
        
        DDINI = Left(MiFecha, 2)
             If DDINI < 16 Then
                  DDINI = 1
                  MiFecha1 = "01" + Mid(MiFecha, 3)
                  Else
                  DDINI = 16
                  MiFecha1 = "16" + Mid(MiFecha, 3)
             End If
        VMiFecha = Mid(Trim(MiFecha), 7, 4) + "-" + Mid(Trim(MiFecha), 4, 2) + "-" + Mid(Trim(MiFecha), 1, 2)
        VMiFecha1 = Mid(Trim(MiFecha1), 7, 4) + "-" + Mid(Trim(MiFecha1), 4, 2) + "-" + Mid(Trim(MiFecha1), 1, 2)
        
    '43 44 FECHAPAGO
        NOMCF2.NOMCF.TextMatrix(I7, 43) = VMiFecha
    '44 45 FECHAINICIALPAGO
        NOMCF2.NOMCF.TextMatrix(I7, 44) = VMiFecha1
    '45 46 FECHAFINALPAGO
        NOMCF2.NOMCF.TextMatrix(I7, 45) = VMiFecha
    '46 47 NUMDIASPAGADOS
        NOMCF2.NOMCF.TextMatrix(I7, 46) = Form8.ConNom1.TextMatrix(I7, 2)
    '47 48 DEPARTAMENTO
        NOMCF2.NOMCF.TextMatrix(I7, 47) = "ADMINISTRACION"
    '48 49 CLABE
        NOMCF2.NOMCF.TextMatrix(I7, 48) = CStr(Form8.ConNom1.TextMatrix(I7, 24))
    '49 50 BANCO
        NOMCF2.NOMCF.TextMatrix(I7, 49) = "002"
    '50 51 FECHAINICIORELLABORAL
        NOMCF2.NOMCF.TextMatrix(I7, 50) = VFal
    '51 52 ANTIGUEDAD
        NOMCF2.NOMCF.TextMatrix(I7, 51) = Year(MiFecha) - Year(MiFechaAlta)
    '52 53 PUESTO
        NOMCF2.NOMCF.TextMatrix(I7, 52) = "ADMINISTRACION"
    '53 54 TIPOCONTRATO
        NOMCF2.NOMCF.TextMatrix(I7, 53) = "01 CONTRATO DE TRABAJO POR TIEMPO INDETERMINADO"
    '54 55 TIPOJORNADA
        NOMCF2.NOMCF.TextMatrix(I7, 54) = "01 DIURNO"
    '55 56 PERIODICIDADPAGO
        If N_ormal = 1 Then
                    NOMCF2.NOMCF.TextMatrix(I7, 55) = "ANUAL"
                    Else
                    NOMCF2.NOMCF.TextMatrix(I7, 55) = "04 QUINCENAL":
        End If
    '56 57 SALARIOBASECOTAPOR
        NOMCF2.NOMCF.TextMatrix(I7, 56) = Format(personal.ingr, "#,##0.00")
    '57 58 RIESGOPUESTO
        NOMCF2.NOMCF.TextMatrix(I7, 57) = RiesgoImss
    '58 59 SALARIODIARIOINTEGRADO
        NOMCF2.NOMCF.TextMatrix(I7, 58) = Format(personal.ingr, "#,##0.00")
    '59 60 ENTIDADFEDERATIVA
        NOMCF2.NOMCF.TextMatrix(I7, 59) = "DIF"
    '60 61 RFCLABORA
        NOMCF2.NOMCF.TextMatrix(I7, 60) = RFC_EXTERNO
    '61 62 PORCENTAJETIEMPO
        If RFC_EXTERNO <> "" Then
            NOMCF2.NOMCF.TextMatrix(I7, 61) = 100
        Else
            NOMCF2.NOMCF.TextMatrix(I7, 61) = ""
        End If
'***********************************************************PERCEPCIONES***************************************
    '62 63 P001
        NOMCF2.NOMCF.TextMatrix(I7, 62) = Format(sue3, "#,##0.00")
    '63 64 P019HED
        NOMCF2.NOMCF.TextMatrix(I7, 63) = Format(0, "#,##0.00")
    '64 65 P019HET
        NOMCF2.NOMCF.TextMatrix(I7, 64) = Format(0, "#,##0.00")
    '65 66 P019HS
        NOMCF2.NOMCF.TextMatrix(I7, 65) = Format(0, "#,##0.00")
    '66 67 P020
        NOMCF2.NOMCF.TextMatrix(I7, 66) = Format(0, "#,##0.00")
    '67 68 P021G
        NOMCF2.NOMCF.TextMatrix(I7, 67) = Format(0, "#,##0.00")
    '68 69 P021E
        NOMCF2.NOMCF.TextMatrix(I7, 68) = Format(0, "#,##0.00")
    '69 70 P002G AGUINALDO
    '70 71 P002E AGUINALDO
    If N_ormal = 1 Then
        If IsNumeric(Form8.ConNom1.TextMatrix(I7, 5)) Then
            NOMCF2.NOMCF.TextMatrix(I7, 69) = Format(Form8.ConNom1.TextMatrix(I7, 5), "###0.00"):  Rem 64 P002 GGratificación Anual (Aguinaldo)
            NOMCF2.NOMCF.TextMatrix(I7, 70) = Format(Form8.ConNom1.TextMatrix(I7, 10), "###0.00"): Rem 65 P002 GGratificación Anual (Aguinaldo exento)
        Else
            NOMCF2.NOMCF.TextMatrix(I7, 69) = Format(0, "#,##0.00")
            NOMCF2.NOMCF.TextMatrix(I7, 70) = Format(0, "#,##0.00")
        End If
    End If
   
    '71 72 P003G PTU
    '72 73 P003E PTU
    If N_ormal = 1 Then
        If IsNumeric(Form8.ConNom1.TextMatrix(I7, 6)) Then
            NOMCF2.NOMCF.TextMatrix(I7, 71) = Format(Form8.ConNom1.TextMatrix(I7, 6), "###0.00"):          Rem 66 P003G PTU
            NOMCF2.NOMCF.TextMatrix(I7, 72) = Format(Form8.ConNom1.TextMatrix(I7, 10), "###0.00"):          Rem 67 P003G PTU
            
        Else
            NOMCF2.NOMCF.TextMatrix(I7, 71) = Format(0, "#,##0.00")
            NOMCF2.NOMCF.TextMatrix(I7, 72) = Format(0, "#,##0.00")
            
        End If
    End If
    
    '73 74 P004GTOS.MEDICOS
        NOMCF2.NOMCF.TextMatrix(I7, 73) = Format(0, "#,##0.00")
    '74 75 P005 FONDO DE AHORRO
        NOMCF2.NOMCF.TextMatrix(I7, 74) = Format(0, "#,##0.00")
    '76 P006 CAJA DE AHORRO
        NOMCF2.NOMCF.TextMatrix(I7, 75) = Format(0, "#,##0.00")
    '77 P009 CONT.A CARGO DEL TRAB.
        NOMCF2.NOMCF.TextMatrix(I7, 76) = Format(0, "#,##0.00")
    '78 P010 PREMIO DE PUNTUALIDAD
    
    If N_ormal = 1 Then
        If IsNumeric(Form8.ConNom1.TextMatrix(I7, 7)) Then
            NOMCF2.NOMCF.TextMatrix(I7, 73) = Format(Form8.ConNom1.TextMatrix(I7, 7) + Form8.ConNom1.TextMatrix(I7, 10), "###0.00"):           Rem 74 P003GPREMIO DE PUNTUALIDAD
        End If
    End If
    
    '78 79 P011 SEGURO DE VIDA
        NOMCF2.NOMCF.TextMatrix(I7, 78) = Format(0, "#,##0.00")
    '80 P012 Seguro de Gastos Medicos Mayores
        NOMCF2.NOMCF.TextMatrix(I7, 79) = Format(0, "#,##0.00")
    '81 P013 Cuotas Sindicales Pagadas por el Patrón
        NOMCF2.NOMCF.TextMatrix(I7, 80) = Format(0, "#,##0.00")
    '82 P014I Subsidios por incapacidad
        NOMCF2.NOMCF.TextMatrix(I7, 81) = Format(0, "#,##0.00")
    '83 P014R Subsidios por incapacidad
        NOMCF2.NOMCF.TextMatrix(I7, 82) = Format(0, "#,##0.00")
    '84 P014M Subsidios por incapacidad
        NOMCF2.NOMCF.TextMatrix(I7, 83) = Format(0, "#,##0.00")
    '85 P015 Becas para trabajadores y/o hijos
        NOMCF2.NOMCF.TextMatrix(I7, 84) = Format(0, "#,##0.00")
    '86 P022G Prima por antigüedad
        NOMCF2.NOMCF.TextMatrix(I7, 85) = Format(0, "#,##0.00")
    '87 P022E Prima por antigüedad
        NOMCF2.NOMCF.TextMatrix(I7, 86) = Format(0, "#,##0.00")
    '88 P023G Pagos por separación
        NOMCF2.NOMCF.TextMatrix(I7, 87) = Format(0, "#,##0.00")
    '89 P023E Pagos por separación
        NOMCF2.NOMCF.TextMatrix(I7, 88) = Format(0, "#,##0.00")
    '90 P024G Seguro de retiro
        NOMCF2.NOMCF.TextMatrix(I7, 89) = Format(0, "#,##0.00")
    '91 P024E Seguro de retiro
        NOMCF2.NOMCF.TextMatrix(I7, 90) = Format(0, "#,##0.00")
    '92 P025G Indeminizaciones
        NOMCF2.NOMCF.TextMatrix(I7, 91) = Format(0, "#,##0.00")
    '93 P025E Indeminizaciones
        NOMCF2.NOMCF.TextMatrix(I7, 92) = Format(0, "#,##0.00")
    '94 P026 Reembolso por funeral
        NOMCF2.NOMCF.TextMatrix(I7, 93) = Format(0, "#,##0.00")
    '95 P027 Cuotas de seguridad social pagadas por el patrón
        NOMCF2.NOMCF.TextMatrix(I7, 94) = Format(0, "#,##0.00")
    '96 P028 Comisiones
        NOMCF2.NOMCF.TextMatrix(I7, 95) = Format(0, "#,##0.00")
    '97 P029G Vales de despensa
        NOMCF2.NOMCF.TextMatrix(I7, 96) = Format(0, "#,##0.00")
    '98 P029E Vales de despensa
        NOMCF2.NOMCF.TextMatrix(I7, 97) = Format(0, "#,##0.00")
    '99 P030 Vales de restaurante
        NOMCF2.NOMCF.TextMatrix(I7, 98) = Format(0, "#,##0.00")
    '100 P031 Vales de gasolina
        NOMCF2.NOMCF.TextMatrix(I7, 99) = Format(0, "#,##0.00")
    '101 P032 Vales de ropa
        NOMCF2.NOMCF.TextMatrix(I7, 100) = Format(0, "#,##0.00")
    '102 P033 Ayuda para renta
        NOMCF2.NOMCF.TextMatrix(I7, 101) = Format(0, "#,##0.00")
    '103 P034 Ayuda para artículos escolares
        NOMCF2.NOMCF.TextMatrix(I7, 102) = Format(0, "#,##0.00")
    '104 P035 Ayuda para anteojos
        NOMCF2.NOMCF.TextMatrix(I7, 103) = Format(0, "#,##0.00")
    '105 P036 Ayuda para transporte
        NOMCF2.NOMCF.TextMatrix(I7, 104) = Format(0, "#,##0.00")
    '106 P037 Ayuda para gastos de funeral
        NOMCF2.NOMCF.TextMatrix(I7, 105) = Format(0, "#,##0.00")
    '107 P038 Otros ingresos por salarios
        NOMCF2.NOMCF.TextMatrix(I7, 106) = Format(0, "#,##0.00")
    '108 P039 Jubilaciones, pensiones o haberes de retiro
        NOMCF2.NOMCF.TextMatrix(I7, 107) = Format(0, "#,##0.00")
    '109 P044 Jubilaciones, pensiones o haberes de retiro en parcialidades
        NOMCF2.NOMCF.TextMatrix(I7, 108) = Format(0, "#,##0.00")
    '110 P045 Ingresos en acciones o títulos valor que representan bienes
        NOMCF2.NOMCF.TextMatrix(I7, 109) = Format(0, "#,##0.00")
    '111 P046 Ingresos asimilados a salarios
        NOMCF2.NOMCF.TextMatrix(I7, 110) = Format(0, "#,##0.00")
    '112 P047 Alimentación
        NOMCF2.NOMCF.TextMatrix(I7, 111) = Format(0, "#,##0.00")
    '113 P048 Habitación
        NOMCF2.NOMCF.TextMatrix(I7, 112) = Format(0, "#,##0.00")
    '114 P049 Premios por asistencia
        NOMCF2.NOMCF.TextMatrix(I7, 113) = Format(0, "#,##0.00")
    '115 P050 Viáticos
        NOMCF2.NOMCF.TextMatrix(I7, 114) = Format(0, "#,##0.00")

'**************************************************DEDUCCIONES******************************************
    '116 D001 Seguridad social
         NOMCF2.NOMCF.TextMatrix(I7, 115) = Format(ims14, "#,##0.00")
    '117 D002 ISR
        NOMCF2.NOMCF.TextMatrix(I7, 116) = Format(isr12, "#,##0.00")
    '118 D003 Aportaciones a retiro, cesantía en edad avanzada y vejez
        NOMCF2.NOMCF.TextMatrix(I7, 117) = Format(0, "#,##0.00")
    '119 D004 Otros
        NOMCF2.NOMCF.TextMatrix(I7, 118) = Format(pre15, "#,##0.00")
    '120 D005 Aportaciones a Fondo de vivienda
        NOMCF2.NOMCF.TextMatrix(I7, 119) = Format(0, "#,##0.00")
    '121 D006M Descuento por incapacidad
        NOMCF2.NOMCF.TextMatrix(I7, 120) = Format(0, "#,##0.00")
    '122 D006I Descuento por incapacidad
        NOMCF2.NOMCF.TextMatrix(I7, 121) = Format(0, "#,##0.00")
    '123 D006R Descuento por incapacidad
        NOMCF2.NOMCF.TextMatrix(I7, 122) = Format(0, "#,##0.00")
    '124 D007 Pensión alimenticia
        NOMCF2.NOMCF.TextMatrix(I7, 123) = Format(pea17, "#,##0.00")
    '125 D008 Renta
        NOMCF2.NOMCF.TextMatrix(I7, 124) = Format(0, "#,##0.00")
    '126 D009 Préstamos provenientes del Fondo Nacional de la vivienda para los trab
        NOMCF2.NOMCF.TextMatrix(I7, 125) = Format(0, "#,##0.00")
    '127 D010 Pago por crédito de vivienda
        NOMCF2.NOMCF.TextMatrix(I7, 126) = Format(ifv18, "#,##0.00")
    '128 D011 Pago de abonos INFONACOT
        NOMCF2.NOMCF.TextMatrix(I7, 127) = Format(fon16, "#,##0.00")
    '129 D012 Anticipo de salarios
        NOMCF2.NOMCF.TextMatrix(I7, 128) = Format(0, "#,##0.00")
    '130 D013 Pagos hechos con exceso al trabajador
        NOMCF2.NOMCF.TextMatrix(I7, 129) = Format(0, "#,##0.00")
    '131 D014 Errores
        NOMCF2.NOMCF.TextMatrix(I7, 130) = Format(0, "#,##0.00")
    '132 D015 Pérdidas
        NOMCF2.NOMCF.TextMatrix(I7, 131) = Format(0, "#,##0.00")
    '133 D016 Averías
        NOMCF2.NOMCF.TextMatrix(I7, 132) = Format(0, "#,##0.00")
    '134 D017 Adquisición de artículos producidos por la empresa o establecimiento
        NOMCF2.NOMCF.TextMatrix(I7, 133) = Format(0, "#,##0.00")
    '135 D018 Cuotas para la constitución y fomento de sociedades cooperativas y de
        NOMCF2.NOMCF.TextMatrix(I7, 134) = Format(0, "#,##0.00")
    '136 D019 Cuotas sindicales
        NOMCF2.NOMCF.TextMatrix(I7, 135) = Format(0, "#,##0.00")
    '137 D020 Ausencia (Ausentismo)
        NOMCF2.NOMCF.TextMatrix(I7, 136) = Format(0, "#,##0.00")
    '138 D021 Cuotas obrero patronales
        NOMCF2.NOMCF.TextMatrix(I7, 137) = Format(0, "#,##0.00")
    '139 D022 Impuestos Locales
        NOMCF2.NOMCF.TextMatrix(I7, 138) = Format(0, "#,##0.00")
    '140 D023 Aportaciones voluntarias
        NOMCF2.NOMCF.TextMatrix(I7, 139) = Format(0, "#,##0.00")
    '141 D024 Ajuste en Gratificación Anual (Aguinaldo) Exento
        NOMCF2.NOMCF.TextMatrix(I7, 140) = Format(0, "#,##0.00")
    '142 D025 Ajuste en Gratificación Anual (Aguinaldo) Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 141) = Format(0, "#,##0.00")
    '143 D026 Ajuste en Participación de los Trabajadores en las Utilidades PTU Exento
        NOMCF2.NOMCF.TextMatrix(I7, 142) = Format(0, "#,##0.00")
    '144 D027 Ajuste en Participación de los Trabajadores en las Utilidades PTU Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 143) = Format(0, "#,##0.00")
    '145 D028 Ajuste en Reembolso de Gastos Médicos Dentales y Hospitalarios Exento
        NOMCF2.NOMCF.TextMatrix(I7, 144) = Format(0, "#,##0.00")
    '146 D029 Ajuste en Fondo de ahorro Exento
        NOMCF2.NOMCF.TextMatrix(I7, 145) = Format(0, "#,##0.00")
    '147 D030 Ajuste en Caja de ahorro Exento
        NOMCF2.NOMCF.TextMatrix(I7, 146) = Format(0, "#,##0.00")
    '148 D031 Ajuste en Contribuciones a Cargo del Trabajador Pagadas por el Patrón Exento
        NOMCF2.NOMCF.TextMatrix(I7, 147) = Format(0, "#,##0.00")
    '149 D032 Ajuste en Premios por puntualidad Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 148) = Format(0, "#,##0.00")
    '150 D033 Ajuste en Prima de Seguro de vida Exento
        NOMCF2.NOMCF.TextMatrix(I7, 149) = Format(0, "#,##0.00")
    '151 D034 Ajuste en Seguro de Gastos Médicos Mayores Exento
        NOMCF2.NOMCF.TextMatrix(I7, 150) = Format(0, "#,##0.00")
    '152 D035 Ajuste en Cuotas Sindicales Pagadas por el Patrón Exento
        NOMCF2.NOMCF.TextMatrix(I7, 151) = Format(0, "#,##0.00")
    '153 D036 Ajuste en Subsidios por incapacidad Exento
        NOMCF2.NOMCF.TextMatrix(I7, 152) = Format(0, "#,##0.00")
    '154 D037 Ajuste en Becas para trabajadores y/o hijos Exento
        NOMCF2.NOMCF.TextMatrix(I7, 153) = Format(0, "#,##0.00")
    '155 D038 Ajuste en Horas extra Exento
        NOMCF2.NOMCF.TextMatrix(I7, 154) = Format(0, "#,##0.00")
    '156 D039 Ajuste en Horas extra Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 155) = Format(0, "#,##0.00")
    '157 D040 Ajuste en Prima dominical Exento
        NOMCF2.NOMCF.TextMatrix(I7, 156) = Format(0, "#,##0.00")
    '158 D041 Ajuste en Prima dominical Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 157) = Format(0, "#,##0.00")
    '159 D042 Ajuste en Prima vacacional Exento
        NOMCF2.NOMCF.TextMatrix(I7, 158) = Format(0, "#,##0.00")
    '160 D043 Ajuste en Prima vacacional Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 159) = Format(0, "#,##0.00")
    '161 D044 Ajuste en Prima por antigüedad Exento
        NOMCF2.NOMCF.TextMatrix(I7, 160) = Format(0, "#,##0.00")
    '162 D045 Ajuste en Prima por antigüedad Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 161) = Format(0, "#,##0.00")
    '163 D046 Ajuste en Pagos por separación Exento
        NOMCF2.NOMCF.TextMatrix(I7, 162) = Format(0, "#,##0.00")
    '164 D047 Ajuste en Pagos por separación Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 163) = Format(0, "#,##0.00")
    '165 D048 Ajuste en Seguro de retiro Exento
        NOMCF2.NOMCF.TextMatrix(I7, 164) = Format(0, "#,##0.00")
    '166 D049 Ajuste en Indemnizaciones Exento
        NOMCF2.NOMCF.TextMatrix(I7, 165) = Format(0, "#,##0.00")
    '167 D050 Ajuste en Indemnizaciones Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 166) = Format(0, "#,##0.00")
    '168 D051 Ajuste en Reembolso por funeral Exento
        NOMCF2.NOMCF.TextMatrix(I7, 167) = Format(0, "#,##0.00")
    '169 D052 Ajuste en Cuotas de seguridad social pagadas por el patrón Exento
        NOMCF2.NOMCF.TextMatrix(I7, 168) = Format(0, "#,##0.00")
    '170 D053 Ajuste en Comisiones Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 169) = Format(0, "#,##0.00")
    '171 D054 Ajuste en Vales de despensa Exento
        NOMCF2.NOMCF.TextMatrix(I7, 170) = Format(0, "#,##0.00")
    '172 D055 Ajuste en Vales de restaurante Exento
        NOMCF2.NOMCF.TextMatrix(I7, 171) = Format(0, "#,##0.00")
    '173 D056 Ajuste en Vales de gasolina Exento
        NOMCF2.NOMCF.TextMatrix(I7, 172) = Format(0, "#,##0.00")
    '174 D057 Ajuste en Vales de ropa Exento
        NOMCF2.NOMCF.TextMatrix(I7, 173) = Format(0, "#,##0.00")
    '175 D058 Ajuste en Ayuda para renta Exento
        NOMCF2.NOMCF.TextMatrix(I7, 174) = Format(0, "#,##0.00")
    '176 D059 Ajuste en Ayuda para artículos escolares Exento
        NOMCF2.NOMCF.TextMatrix(I7, 175) = Format(0, "#,##0.00")
    '177 D060 Ajuste en Ayuda para anteojos Exento
        NOMCF2.NOMCF.TextMatrix(I7, 176) = Format(0, "#,##0.00")
    '178 D061 Ajuste en Ayuda para transporte Exento
        NOMCF2.NOMCF.TextMatrix(I7, 177) = Format(0, "#,##0.00")
    '179 D062 Ajuste en Ayuda para gastos de funeral Exento
        NOMCF2.NOMCF.TextMatrix(I7, 178) = Format(0, "#,##0.00")
    '180 D063 Ajuste en Otros ingresos por salarios Exento
        NOMCF2.NOMCF.TextMatrix(I7, 179) = Format(0, "#,##0.00")
    '181 D064 Ajuste en Otros ingresos por salarios Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 180) = Format(0, "#,##0.00")
    '182 D065 Ajuste en Jubilaciones, pensiones o haberes de retiro Exento
        NOMCF2.NOMCF.TextMatrix(I7, 181) = Format(0, "#,##0.00")
    '183 D066 Ajuste en Jubilaciones, pensiones o haberes de retiro Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 182) = Format(0, "#,##0.00")
    '184 D067 Ajuste en Pagos por separación Acumulable
        NOMCF2.NOMCF.TextMatrix(I7, 183) = Format(0, "#,##0.00")
    '185 D068 Ajuste en Pagos por separación No acumulable
        NOMCF2.NOMCF.TextMatrix(I7, 184) = Format(0, "#,##0.00")
    '186 D069 Ajuste en Jubilaciones, pensiones o haberes de retiro Acumulable
        NOMCF2.NOMCF.TextMatrix(I7, 185) = Format(0, "#,##0.00")
    '187 D070 Ajuste en Jubilaciones, pensiones o haberes de retiro No acumulable
        NOMCF2.NOMCF.TextMatrix(I7, 186) = Format(0, "#,##0.00")
    '188 D071 Ajuste en Subsidio para el empleo (efectivamente entregado al trabajador)
        NOMCF2.NOMCF.TextMatrix(I7, 187) = Format(0, "#,##0.00")
    '189 D072 Ajuste en Ingresos en acciones o títulos valor que representan bienes Exento
        NOMCF2.NOMCF.TextMatrix(I7, 188) = Format(0, "#,##0.00")
    '190 D073 Ajuste en Ingresos en acciones o títulos valor que representan bienes Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 189) = Format(0, "#,##0.00")
    '191 D074 Ajuste en Alimentación Exento
        NOMCF2.NOMCF.TextMatrix(I7, 190) = Format(0, "#,##0.00")
    '192 D075 Ajuste en Alimentación Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 191) = Format(0, "#,##0.00")
    '193 D076 Ajuste en Habitación Exento
        NOMCF2.NOMCF.TextMatrix(I7, 192) = Format(0, "#,##0.00")
    '194 D077 Ajuste en Habitación Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 193) = Format(0, "#,##0.00")
    '195 D078 Ajuste en Premios por asistencia
        NOMCF2.NOMCF.TextMatrix(I7, 194) = Format(0, "#,##0.00")
    '196 D079 Ajuste en Pagos distintos a los listados y que no deben considerarse como ingreso por sueldos, salarios o ingresos asimilados.
        NOMCF2.NOMCF.TextMatrix(I7, 195) = Format(0, "#,##0.00")
    '197 D080 Ajuste en Viáticos gravados
        NOMCF2.NOMCF.TextMatrix(I7, 196) = Format(0, "#,##0.00")
    '198 D081 Ajuste en Viáticos (entregados al trabajador)
        NOMCF2.NOMCF.TextMatrix(I7, 197) = Format(0, "#,##0.00")
    '199 D082 Ajuste en Fondo de ahorro Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 198) = Format(0, "#,##0.00")
    '200 D083 Ajuste en Caja de ahorro Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 199) = Format(0, "#,##0.00")
    '201 D084 Ajuste en Prima de Seguro de vida Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 200) = Format(0, "#,##0.00")
    '202 D085 Ajuste en Seguro de Gastos Médicos Mayores Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 201) = Format(0, "#,##0.00")
    '203 D086 Ajuste en Subsidios por incapacidad Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 202) = Format(0, "#,##0.00")
    '204 D087 Ajuste en Becas para trabajadores y/o hijos Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 203) = Format(0, "#,##0.00")
    '205 D088 Ajuste en Seguro de retiro Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 204) = Format(0, "#,##0.00")
    '206 D089 Ajuste en Vales de despensa Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 205) = Format(0, "#,##0.00")
    '207 D090 Ajuste en Vales de restaurante Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 206) = Format(0, "#,##0.00")
    '208 D091 Ajuste en Vales de gasolina Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 207) = Format(0, "#,##0.00")
    '209 D092 Ajuste en Vales de ropa Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 208) = Format(0, "#,##0.00")
    '210 D093 Ajuste en Ayuda para renta Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 209) = Format(0, "#,##0.00")
    '211 D094 Ajuste en Ayuda para artículos escolares Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 210) = Format(0, "#,##0.00")
    '212 D095 Ajuste en Ayuda para anteojos Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 211) = Format(0, "#,##0.00")
    '213 D096 Ajuste en Ayuda para transporte Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 212) = Format(0, "#,##0.00")
    '214 D097 Ajuste en Ayuda para gastos de funeral Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 213) = Format(0, "#,##0.00")
    '215 D098 Ajuste a ingresos asimilados a salarios gravados
        NOMCF2.NOMCF.TextMatrix(I7, 214) = Format(0, "#,##0.00")
    '216 D099 Ajuste a ingresos por sueldos y salarios gravados
        NOMCF2.NOMCF.TextMatrix(I7, 215) = Format(0, "#,##0.00")
    '217 D100 Ajuste en Viáticos exentos
        NOMCF2.NOMCF.TextMatrix(I7, 216) = Format(0, "#,##0.00")


    '218 OP001 Reintegro de ISR pagado en exceso (siempre que no haya sido enterado al SAT).
        NOMCF2.NOMCF.TextMatrix(I7, 217) = Format(0, "#,##0.00")
    '219 OP002 Subsidio para el empleo (efectivamente entregado al trabajador).
        NOMCF2.NOMCF.TextMatrix(I7, 218) = Format(sub13, "#,##0.00")
    '220 OP003 Viáticos (entregados al trabajador).
        NOMCF2.NOMCF.TextMatrix(I7, 219) = Format(0, "#,##0.00")
    '221 OP004 Aplicación de saldo a favor por compensación anual.
        NOMCF2.NOMCF.TextMatrix(I7, 220) = Format(0, "#,##0.00")
    '222 OP999 Pagos distintos a los listados y que no deben considerarse como ingreso por sueldos, salarios o ingresos asimilados.
        NOMCF2.NOMCF.TextMatrix(I7, 221) = Format(0, "#,##0.00")
    '223 DINCAPACIDADM
        NOMCF2.NOMCF.TextMatrix(I7, 222) = Format(0, "#,##0.00")
    '224 DINCAPACIDADI
        NOMCF2.NOMCF.TextMatrix(I7, 223) = Format(0, "#,##0.00")
    '225 DINCAPACIDADR
        NOMCF2.NOMCF.TextMatrix(I7, 224) = Format(0, "#,##0.00")
    '226 P014DINCAM
        NOMCF2.NOMCF.TextMatrix(I7, 225) = Format(0, "#,##0.00")
    '227 P014DINCAI
        NOMCF2.NOMCF.TextMatrix(I7, 226) = Format(0, "#,##0.00")
    '228 P014DINCAR
        NOMCF2.NOMCF.TextMatrix(I7, 227) = Format(0, "#,##0.00")
    '229 DIASHE
        NOMCF2.NOMCF.TextMatrix(I7, 228) = Format(0, "#,##0.00")
    '230 HORASDOBLES
        NOMCF2.NOMCF.TextMatrix(I7, 229) = Format(0, "#,##0.00")
    '231 DIASHETRIPLES
        NOMCF2.NOMCF.TextMatrix(I7, 230) = Format(0, "#,##0.00")
    '232 HORASTRIPLES
        NOMCF2.NOMCF.TextMatrix(I7, 231) = Format(0, "#,##0.00")
    '233 DIASHESIMPLES
        NOMCF2.NOMCF.TextMatrix(I7, 232) = Format(0, "#,##0.00")
    '234 HORASSIMPLES
        NOMCF2.NOMCF.TextMatrix(I7, 233) = Format(0, "#,##0.00")
    '235 AT_VMERCADO
        NOMCF2.NOMCF.TextMatrix(I7, 234) = Format(0, "#,##0.00")
    '236 AT_POTORGARSE
        NOMCF2.NOMCF.TextMatrix(I7, 235) = Format(0, "#,##0.00")
    '237 JB_TEXHIBICION
        NOMCF2.NOMCF.TextMatrix(I7, 236) = Format(0, "#,##0.00")
    '238 JB_TPARCIAL
        NOMCF2.NOMCF.TextMatrix(I7, 237) = Format(0, "#,##0.00")
    '239 JB_MDIARIO
        NOMCF2.NOMCF.TextMatrix(I7, 238) = Format(0, "#,##0.00")
    '240 JB_IACUM
        NOMCF2.NOMCF.TextMatrix(I7, 239) = Format(0, "#,##0.00")
    '241 JB_INOACUM
        NOMCF2.NOMCF.TextMatrix(I7, 240) = Format(0, "#,##0.00")
    '242 SI_TPAGADO
        NOMCF2.NOMCF.TextMatrix(I7, 241) = Format(0, "#,##0.00")
    '243 SI_ASERVICIO
        NOMCF2.NOMCF.TextMatrix(I7, 242) = Format(0, "#,##0.00")
    '244 SI_USUELDO
        NOMCF2.NOMCF.TextMatrix(I7, 243) = Format(0, "#,##0.00")
    '245 SI_IACUM
        NOMCF2.NOMCF.TextMatrix(I7, 244) = Format(0, "#,##0.00")
    '246 SI_INOACUM
        NOMCF2.NOMCF.TextMatrix(I7, 245) = Format(0, "#,##0.00")
    '247 CSF_SALFAV
        NOMCF2.NOMCF.TextMatrix(I7, 246) = Format(0, "#,##0.00")
    '248 CSF_ANIO
        NOMCF2.NOMCF.TextMatrix(I7, 247) = Format(0, "#,##0.00")
    '249 CSF_RSFAV
        NOMCF2.NOMCF.TextMatrix(I7, 248) = Format(0, "#,##0.00")
    '250 BANDERA
        NOMCF2.NOMCF.TextMatrix(I7, 249) = Format(0, "#,##0.00")
End Sub
