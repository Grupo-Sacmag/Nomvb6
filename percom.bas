Attribute VB_Name = "Module1"
Type sueldosDEON
    sueldoDeon As Currency
    sueldoFactura As Currency
End Type

Type saldia
    Salario As Double
End Type

Type per
    nom As String * 20
    ape1 As String * 20
    ape2 As String * 20
    RFC As String * 18
    imss As String * 18
    fal As String * 12
    fab As String * 12
    ingr As Currency
    viat As Currency
    otras As Currency
    integrado As Currency
 End Type
 
 Type basini
     datoarch As String * 255
 End Type

 Type nom
     dias As Currency
     hsnor As Currency
     hs_no As Currency
     hsdbl As Currency
     hs_db As Currency
     hstri As Currency
     hs_tr As Currency
     ispt As Currency
     crdsal As Currency
     imss As Currency
     sueldo As Currency
     hs_nor As Currency
     hs_dbl As Currency
     hs_tri As Currency
     viaticos As Currency
     pvac As Currency
     otras As Currency
     aguin As Currency
     ptu As Currency
     exentos As Currency
     prestamos As Currency
     fonacot As Currency
     telefono As Currency
     otraded As Currency
  End Type
   Type nom_dos
     dias As Currency
     hsnor As Currency
     hs_no As Currency
     hsdbl As Currency
     hs_db As Currency
     hstri As Currency
     hs_tr As Currency
     ispt As Currency
     crdsal As Currency
     imss As Currency
     sueldo As Currency
     hs_nor As Currency
     hs_dbl As Currency
     hs_tri As Currency
     viaticos As Currency
     pvac As Currency
     otras As Currency
     aguin As Currency
     ptu As Currency
     exentos As Currency
     prestamos As Currency
     fonacot As Currency
     telefono As Currency
     otraded As Currency
  End Type
 Type art
     liminf As Currency
     limsup As Currency
     cuotaf As Currency
     porcsl As Currency
  End Type
  Type subs
     liminfs As Currency
     limsups As Currency
     cuotafs As Currency
     porcsls As Currency
  End Type
   Type cred
     crede As Currency
     crea As Currency
     cresam As Currency
  End Type
  
    Type empre
       name As String * 60
       ao As Integer
       sm As Currency
       psub As Currency
       fecha As String * 14
    End Type
  
    Type ob
    O_1 As Integer
    por_1 As Integer
    im_1 As Currency
    O_2 As Integer
    por_2 As Integer
    im_2 As Currency
    O_3 As Integer
    por_3 As Integer
    im_3 As Currency
    O_4 As Integer
    por_4 As Integer
    im_4 As Currency
    O_5 As Integer
    por_5 As Integer
    im_5 As Currency
    O_6 As Integer
    por_6 As Integer
    im_6 As Currency
    O_7 As Integer
    por_7 As Integer
    im_7 As Currency
    O_8 As Integer
    por_8 As Integer
    im_8 As Currency
    O_9 As Currency
    por_9 As Integer
    im_9 As Currency
    O_10 As Integer
    por_10 As Integer
    im_10 As Currency
    O_11 As Integer
    por_11 As Integer
    im_11 As Currency
    O_12 As Integer
    por_12 As Integer
    im_12 As Currency
    O_13 As Integer
    por_13 As Integer
    im_13 As Currency
    O_14 As Integer
    por_14 As Integer
    im_14 As Currency
    O_15 As Integer
    por_15 As Integer
    im_15 As Currency
    O_16 As Integer
    por_16 As Integer
    im_16 As Currency
    O_17 As Integer
    por_17 As Integer
    im_17 As Currency
    O_18 As Integer
    por_18 As Integer
    im_18 As Currency
    O_19 As Integer
    por_19 As Integer
    im_19 As Currency
    O_20 As Integer
    por_20 As Integer
    im_20 As Currency

 End Type
 Type CAT_MA
    B1 As String * 6
    B2 As String * 32
    B3 As String * 16
    B4 As String * 5
    B5 As String * 5
End Type
Type CAT_AX
    C1 As String * 6
    C2 As String * 32
    C3 As String * 16
    C4 As String * 5
    C5 As String * 5
End Type
Type Clabnx
    Q1 As String * 16
 End Type
 Type da_id
       Emp_Rfc As String * 25
       Emp_Dom As String * 70
       Rep_Legapp As String * 20
       Rep_Legapm As String * 20
       Rep_Legapn As String * 20
       Rep_Rfc As String * 25
       Rep_Curp As String * 25
       suc As String * 4
       cta As String * 12
       dias As Integer
       clte As String * 12
  End Type
 Type ult
     num As Long
     ubi As Integer
     renglon As Long
     texto As String
     Poliza As Integer
     Impresion As Integer
 End Type
 Type refo
      Mes As Integer
      dia As Integer
 End Type
 
 Type imptosp
        ImpTIni As Currency
        ImpTFin As Currency
        SubAIni As Currency
        SubAFin As Currency
        SubNIni As Currency
        SubNFin As Currency
        CalcDoble As Integer
 End Type
 Type NominaXFuera
     Cordinador As Integer
     Empleado As Integer
     DiasTrab As Integer
     Comentarios As String * 100
     Aplicado As Integer
 End Type
 
 Public DEON As sueldosDEON
 Public con As New ADODB.Connection
 Public strconnect As String
 Public emp As String
 Public NxF As NominaXFuera, NxFTmP As NominaXFuera
 Public ReferOper As refo
 Public li As Long, Imto_deotra As Currency, P_sub1 As Currency
 Public DifImptos As imptosp, Tot_dias As Currency, Crd_deotra As Currency, Crpag_deotra As Currency
 Public Dat_ide As da_id, SubT_Mes As String
 Public Clbnx As Clabnx
 Public CATAUX As CAT_AX
 Public salarioDI As saldia
 Public basico As basini
 Public obras As ob
 Public maestro As ob, creere As Currency
 Public empresa As empre, ingresos As Currency
 Public nomina As nom
 Public nominaAnterior As nom_dos
 Public deducciones As Currency
 Public personal As per, neto As Currency
 Public rgtro As Double, CREDITO_PROV As Currency
 Public articulo As art, final As Long, APLICAR
 Public subsidio As subs, impto As Currency
 Public credito As cred, base As Currency, Sub_Aplic1 As Currency, ImtoTo_otra As Currency
 Public z1$, mm(12) As String * 20, cm, Dm, z2$, ddm
 Public subdirectorio$, valor$, dd(12) As Integer
 Public Dir_imptos As String, regtro As Integer
 Public arch_tr As String, tar As Integer, Direc_torio
 Public dir_obras As String, baseanual As Currency, cal_anual As Integer, baseor As Currency
 Public obra(22) As Currency, porcentaje(22) As Currency, Base_anual1 As Currency, Psub_Extra As Currency
 Public estadoDelaConexion
 
Sub veridir()
    On Error GoTo corrdire
    Dir_imptos = DirecT_arifas
    Open "C:\GconTA\perma.dno" For Random As #7 Len = Len(basico)
    fin_basico = LOF(7) / Len(basico)
    
    If fin_basico > 0 Then
        Get 7, 1, basico
        Direc_torio = RTrim(basico.datoarch)
        If Direc_torio <> "" Then
            Form1.Drive1.Drive = Left(Direc_torio, 2)
            ChDir Direc_torio
            Form1.Dir1 = Direc_torio
            Close 7
        End If

        Open "C:\GconTA\perma.dno" For Random As #7 Len = Len(basico)
        fin_basico = LOF(7) / Len(basico)

        If fin_basico > 1 Then
            Get 7, 2, basico
            dir_obras = RTrim(basico.datoarch)
        End If
    End If

    Close 7
    Exit Sub

corrdire:
    Close 7
    ChDir "C:\GconTa\"
End Sub
  Sub detbase()
     Close 10: SUMA_CREDITO_MES = 0
     Open "AcuNom2.dno" For Random As 10 Len = Len(ArAcum)
     lq_2 = LOF(10) / Len(ArAcum)
     If lq_2 = 0 Then
          MsgBox "No Existe archivo acumulado para calcular la ultima nomina", vbCritical + vbDefaultButton1, "Captura Nomina"
          Else
          Get 10, rgtro, ArAcum
         Rem ******* PRIMA VACACIONAL  ***************************************************
          Acumul_virtual
          vaca = ArAcum.Pvaca
          'If vaca < (empresa.sm * 15) Then
                'vaca = 0
                'Else
                'vaca = ArAcum.Pvaca - (empresa.sm * 15)
          'End If
          Rem ******************* AGUINALDO ANUAL  ***************************************
          agu_i = ArAcum.Pagui
          'If agu_i < (empresa.sm * 30) Then
                'agu_i = 0
                'Else
                'agu_i = ArAcum.Pagui - (empresa.sm * 30)
                
          'End If
        Rem ************* PTU  *************************************************************
          pt_u = ArAcum.PPTU
          'If pt_u < (empresa.sm * 15) Then
                'pt_u = 0
                'Else
                'pt_u = ArAcum.PPTU - (empresa.sm * 15)
          'End If
'          Stop
          baseanual = ArAcum.Pnormal + ArAcum.Pextra + ArAcum.Potras + ArAcum.Pviaticos + vaca + agu_i + pt_u
          
          SUMA_CREDITO_MES = ArAcum.DCrApl
      End If
       Rem P_sub1 = (Sub_Aplic1 + ArAcum.DSubioAp) / (psub_extra + ArAcum.DSubioAp + ArAcum.DSubNoap)
       
End Sub
Sub Acumul_virtual()
If li = 2165 Then
    Stop
End If


If Form8.ConNom1.TextMatrix(li, 0) <> "" Then regtro = Form8.ConNom1.TextMatrix(li, 0) Else regtro = 0
If Form8.ConNom1.TextMatrix(li, 2) <> "" Then nomina.dias = Form8.ConNom1.TextMatrix(li, 2) Else nomina.dias = 0
If Form8.ConNom1.TextMatrix(li, 3) <> "" Then nomina.sueldo = Form8.ConNom1.TextMatrix(li, 3) Else nomina.sueldo = 0
If Form8.ConNom1.TextMatrix(li, 4) <> "" Then nomina.hs_nor = Form8.ConNom1.TextMatrix(li, 4) Else nomina.hs_nor = 0
If N_ormal = 1 Then
     If Form8.ConNom1.TextMatrix(li, 5) <> "" Then nomina.aguin = Form8.ConNom1.TextMatrix(li, 5) Else nomina.aguin = 0
        Else
        If Form8.ConNom1.TextMatrix(li, 5) <> "" Then nomina.hs_dbl = Form8.ConNom1.TextMatrix(li, 5) Else nomina.hs_dbl = 0
   End If
   If N_ormal = 1 Then
        If Form8.ConNom1.TextMatrix(li, 6) <> "" Then nomina.ptu = Form8.ConNom1.TextMatrix(li, 6) Else nomina.ptu = 0
        Else
        If Form8.ConNom1.TextMatrix(li, 6) <> "" Then nomina.hs_tri = Form8.ConNom1.TextMatrix(li, 6) Else nomina.hs_tri = 0
   End If
   If Form8.ConNom1.TextMatrix(li, 7) <> "" Then
                nomina.viaticos = Form8.ConNom1.TextMatrix(li, 7)
                Else
                nomina.viaticos = 0
   End If
   If Form8.ConNom1.TextMatrix(li, 8) <> "" Then
                nomina.pvac = Form8.ConNom1.TextMatrix(li, 8)
                Else
                nomina.pvac = 0
   End If
   If Form8.ConNom1.TextMatrix(li, 9) <> "" Then
                nomina.otras = Form8.ConNom1.TextMatrix(li, 9)
                Else
                nomina.otras = 0
   End If
   If Form8.ConNom1.TextMatrix(li, 10) <> "" Then
                nomina.exentos = Form8.ConNom1.TextMatrix(li, 10)
                Else
                nomina.exentos = 0
   End If
'Stop
ArAcum.Pdias = ArAcum.Pdias + nomina.dias
ArAcum.Pnormal = ArAcum.Pnormal + nomina.sueldo
ArAcum.Pextra = ArAcum.Pextra + (nomina.hs_nor + nomina.hs_dbl + nomina.hs_tri)
ArAcum.Pviaticos = ArAcum.Pviaticos + nomina.viaticos
ArAcum.Pagui = ArAcum.Pagui + nomina.aguin
ArAcum.Pvaca = ArAcum.Pvaca + nomina.pvac
ArAcum.Potras = ArAcum.Potras + nomina.otras
ArAcum.PPTU = ArAcum.PPTU + nomina.ptu
ArAcum.Pexenta = ArAcum.Pexenta + nomina.exentos
Rem ArAcum.DImpto = ArAcum.DImpto + nom_com.ImpTot
SUMA_CREDITO_MES = ArAcum.DCrPag
If Sub_Aplic1 > 0 Then
    P_sub1 = (ArAcum.DSubioAp + Sub_Aplic1) / (Psub_Extra + ArAcum.DSubioAp + ArAcum.DSubNoap)
End If
'Stop
End Sub
Sub Acum_Doble()
     Close 10
     Psub_Extra = 0
     Base_anual1 = 0
     Tot_dias = 0
     P_sub1 = 0
     Sub_Aplic1 = 0
     Crd_deotra = 0
     Crpag_deotra = 0
     Open "AcuNom1.dno" For Random As 10 Len = Len(ArAcum)
     lq_2 = LOF(10) / Len(ArAcum)
      
     If lq_2 = 0 Then
          Rem MsgBox "No Existe archivo acumulado para calcular la ultima nomina", vbCritical + vbDefaultButton1, "Captura Nomina"
          Rem Exit Sub
          Else
          Get 10, rgtro, ArAcum
          
          vaca = ArAcum.Pvaca
          'If vaca < (empresa.sm * 15) Then
                'vaca = 0
                'Else
                'vaca = ArAcum.Pvaca - (empresa.sm * 15)
          'End If
          agu_i = ArAcum.Pagui
          'If agu_i < (empresa.sm * 30) Then
                'agu_i = 0
                'Else
                'agu_i = ArAcum.Pagui - (empresa.sm * 30)
                
          'End If
          pt_u = ArAcum.PPTU
          'If pt_u < (empresa.sm * 15) Then
                'pt_u = 0
                'Else
                'pt_u = ArAcum.PPTU - (empresa.sm * 15)
          'End If
          Tot_dias = ArAcum.Pdias
          Psub_Extra = ArAcum.DSubioAp + ArAcum.DSubNoap
          Sub_Aplic1 = ArAcum.DSubioAp
          Base_anual1 = ArAcum.Pnormal + ArAcum.Pextra + ArAcum.Potras + ArAcum.Pviaticos + vaca + agu_i + pt_u
          
          ImtoTo_otra = ArAcum.DImpto
          Imto_deotra = ArAcum.DImpret
          Crd_deotra = ArAcum.DCrApl
          Crpag_deotra = ArAcum.DCrPag
      End If
   Close 10
     Open "AcuNom2.dno" For Random As 10 Len = Len(ArAcum)
     lq_2 = LOF(10) / Len(ArAcum)
     Get 10, rgtro, ArAcum
     Tot_dias = Tot_dias + ArAcum.Pdias
     Close 10
     
End Sub
  Sub calc_anual(base, impto, psub)
  
    If psub <> 1 Then
    Stop
    End If
    Dim ISR_1 As Currency, SUB_1 As Currency, CRED_1 As Currency
    'calculo base, impto, psub
    ISR_1 = nom_com.ImpTot: SUB_1 = nom_com.subapl: CRED_1 = nom_com.CredNe
    Close 3: Close 4: Close 5
    Open (Trim(Dir_imptos) + "ISR177.03") For Random As #3 Len = Len(articulo)
    Dem = LOF(3) / Len(articulo)
    Open (Trim(Dir_imptos) + "SUB178.03") For Random As #4 Len = Len(subsidio)
    EM = LOF(4) / Len(subsidio)
    Open (Trim(Dir_imptos) + "CRE116.03") For Random As #5 Len = Len(credito)
    eem = LOF(5) / Len(credito)
    Rem  **** CALCULAR IMPUESTO ****
    
    baseor = base
    baseanual = 0
    
    detbase
    
    Rem If regtro = 587 Then Stop
    base = baseanual + Base_anual1
    
    For i = 1 To Dem: Get 3, i, articulo
     If base > (articulo.liminf) And base < (articulo.limsup) Then
           marginal = ((base - articulo.liminf) * (articulo.porcsl / 100))
           impto = marginal + articulo.cuotaf
           nom_com.ImpTot = impto - ArAcum.DImpto - ImtoTo_otra: Rem   *********************
           Imp_mag = (articulo.porcsl / 100)
           i = Dem
     End If
    Next i
'    Stop
    mientras = psub
GoTo SuBsidio_Salto
SuBsidio_Salto:
     subdio = ArAcum.DSubioAp + SUB_1
    psub = mientras
Rem ************************** ELIMINA CREDITO ANUAL *********************************
    'For i = 1 To eem: Get 5, i, credito
     'If base > (credito.crede) And base < credito.crea Then
         'creere = (credito.cresam)
         'nom_com.CreTot = creere - ArAcum.DCrApl - Crd_deotra
         Rem If nom_com.CreTot < 0 Then nom_com.CreTot = 0
         
     'End If
    'Next i
Rem *************************  ELIMINA CREDITO ANUAL *********************************
    
    creere = SUMA_CREDITO_MES + CREDITO_PROV
    
    Rem impto = impto - subdio - creere
    
    impto = impto - subdio
    Rem ***********************************************************************************************
    Rem SI EL  IMPUESTO ANUAL ES MAYOR AL SUBSIDIO ENTONCES SE CALCULA EL IMPUESTO DE LA ULTIMA NOMINA*
    Rem ***********************************************************************************************
    
    If impto > 0 Then
      Rem impto = impto - ArAcum.DImpret - ArAcum.DCrPag - Imto_deotra - Crpag_deotra
      
      impto = impto - ArAcum.DImpret - Imto_deotra + ArAcum.DCrApl
      Else
      impto = CRED_1
      Rem AQUI VA EL OTRO
    End If
    
    Rem If (impto < 0) And (ArAcum.DSubioAp > 1) Then impto = 0
    Rem impto = impto - ArAcum.DImpret + ArAcum.DCrPag - Imto_deotra + Crpag_deotra
    Rem If impto < 0 Then
        Rem If ArAcum.DCrApl = 0 Then
                Rem nom_com.CredNe = 0
                Rem Else
            Rem Stop
            Rem impto = CRED_1
            
     Rem End If
        Rem Else
        Rem nom_com.CredNe = 0
    Rem End If
    Put 14, regtro, nom_com
    Rem impto = impto - subdio
    base = baseor
    Close 10
'    Stop
End Sub
  
Sub calculo_compl(base, impto, psub)
    Close 3: Close 4: Close 5
    Open (Dir_imptos + "TABLA.ISR") For Random As #3 Len = Len(articulo)
    Dem = LOF(3) / Len(articulo)
    Open (Dir_imptos + "TABLA.SUB") For Random As #4 Len = Len(subsidio)
    EM = LOF(4) / Len(subsidio)
    Open (Dir_imptos + "TABLA.CRE") For Random As #5 Len = Len(credito)
    eem = LOF(5) / Len(credito)
    
    Rem  **** CALCULAR IMPUESTO ****
    For i = 1 To Dem: Get 3, i, articulo
     If base > (articulo.liminf) And base < (articulo.limsup) Then
           marginal = ((base - articulo.liminf) * (articulo.porcsl / 100))
           impto = marginal + articulo.cuotaf
           Imp_mag = (articulo.porcsl / 100)
           nom_com.ImpTot = impto: Rem **********************************
           i = Dem
     End If
    Next i
    For i = 1 To EM: Get 4, i, subsidio
     If base > (subsidio.liminfs) And base < (subsidio.limsups) Then
           marginal2 = ((base - subsidio.liminfs) * Imp_mag)
           subdio = (marginal2 * subsidio.porcsls / 100) + (subsidio.cuotafs)
           nom_com.subdio = subdio
           nom_com.PSubDi = psub
           subdio = (subdio * psub)
           nom_com.subapl = subdio: Rem ************************************
           nom_com.subNap = nom_com.subdio - nom_com.subapl: Rem ************
           i = EM
     End If
    Next i
    For i = 1 To eem: Get 5, i, credito
     If base > (credito.crede) And base < credito.crea Then
         creere = (credito.cresam)
         nom_com.CreTot = creere: Rem ****************************************
     End If
    Next i
    
    impto = impto - subdio - creere
    If impto > 0 Then
        impto = impto + creere
        Else
        If (impto < 0) And (DifImptos.CalcDoble = 0) Then impto = 0
    End If
    
    Rem impto = impto - subdio
 End Sub
Sub calculo(base, impto, psub)
    Close 3: Close 4: Close 5
    If Form8.Option3.Value = True Then ' Retencion Quincenal
        Close 3: Close 4: Close 5
        
        Open (Trim(Dir_imptos) + "Tab08Kin.ISR") For Random As #3 Len = Len(articulo)
            Dem = LOF(3) / Len(articulo)
        Open (Trim(Dir_imptos) + "Tab08Kin.SUB") For Random As #4 Len = Len(subsidio)
            EM = LOF(4) / Len(subsidio)
        
        For i = 1 To Dem
            Get 3, i, articulo
            If base > (articulo.liminf) And base < (articulo.limsup) Then
                marginal = ((base - articulo.liminf) * (articulo.porcsl / 100)) 'IMP.S/EXEC.
                Imp_mag = (articulo.porcsl / 100) '% APLIC.
                impto = marginal + articulo.cuotaf '1er impuesto
                nom_com.ImpTot = impto
                Exit For
            End If
        Next i
        
        For i = 1 To EM
            Get 4, i, subsidio
            If base > (subsidio.liminfs) And base < (subsidio.limsups) Then
                subdio = subsidio.cuotafs
                nom_com.subdio = subdio
                subdio = subdio
                nom_com.subapl = subdio 'subsidio causado
                nom_com.subNap = 0
                Exit For
            End If
        Next i
        
        impto = impto - subdio
        If impto < 0 Then
            impto = 0
        Else
            impto = impto
        End If
        nom_com.CredNe = 0
        Put 14, regtro, nom_com
    End If
     
    If Form8.Option4.Value = True Then 'Retencion mensual
        mesNomina = Trim(Left(Form8.Combo1.Text, 3))
        AñoNomina = Right(Trim(Form8.Label7.Caption), 4)
        Close 35
        Open CStr(Trim(Form1.Dir1)) + "\" + mesNomina + "1" + AñoNomina + ".NOM" For Random As 35 Len = Len(nomina)
        Get 35, regtro, nomina
        IngresosAnteriores = nomina.sueldo + nomina.hs_nor + nomina.hs_dbl + nomina.hs_tri + nomina.aguin + nomina.ptu + nomina.viaticos + nomina.pvac + nomina.otras '+ nomina.exentos
        ImpuestoAnteriorRetenido = nomina.ispt
        Close 35: Close 36
        Open CStr(Trim(Form1.Dir1)) + "\" + mesNomina + "1" + AñoNomina + ".cmp" For Random As 36 Len = Len(nom_com)
        Get 36, regtro, nom_com
        ImpuestoAnteriorTotal = nom_com.ImpTot
        subsidioCausado = nom_com.subapl
        Close 36: Close 3: Close 4: Close 5
        Open (Trim(Dir_imptos) + "Tab08Mes.ISR") For Random As #3 Len = Len(articulo)
            Dem = LOF(3) / Len(articulo)
        Open (Trim(Dir_imptos) + "Tab08Mes.SUB") For Random As #4 Len = Len(subsidio)
            EM = LOF(4) / Len(subsidio)
        
        baseFinal = base + IngresosAnteriores
        
        For i = 1 To Dem
            Get 3, i, articulo
            If baseFinal > (articulo.liminf) And baseFinal < (articulo.limsup) Then
                marginal = ((baseFinal - articulo.liminf) * (articulo.porcsl / 100)) 'IMP.S/EXEC.
                Imp_mag = (articulo.porcsl / 100) '% APLIC.
                impto = marginal + articulo.cuotaf '1er impuesto
                nom_com.ImpTot = impto - ImpuestoAnteriorTotal
                Exit For
            End If
        Next i
        
        For i = 1 To EM
            Get 4, i, subsidio
            If baseFinal > (subsidio.liminfs) And baseFinal < (subsidio.limsups) Then
                subdio = subsidio.cuotafs 'subsidio causado
                nom_com.subdio = subdio
                subdio = subdio
                nom_com.subapl = subdio - subsidioCausado
                nom_com.subNap = 0
                Exit For
            End If
        Next i
        
        impto = impto - subdio

        If impto < 0 Then
            nom_com.CredNe = impto 'subsidio pagado
        Else
            nom_com.CredNe = 0
        End If
        Put 14, regtro, nom_com
    End If


End Sub
Sub imss(integrado As Currency, seguro As Currency, diaseg As Currency)
    imss2 integrado, seguro, diaseg
End Sub

Sub imss2(integrado As Currency, seguro As Currency, diaseg As Currency)
    Dim Anual1 As Integer
        salint1 = 0: salint2 = 0
        enfymat = 0: otrascu = 0
        salint1 = integrado: salint2 = 0
        seguro = 0
        Anual1 = empresa.ao
        
        If Anual1 >= 2008 Then
                salint1 = 0: salint2 = 0
                enfymat = 0: otrascu = 0
                salint1 = integrado
                
                If integrado > (empresa.sm * 3) Then
                    
                    If integrado > (empresa.sm * 25) Then
                        salint1 = (empresa.sm * 25)
                    Else
                        salint1 = integrado
                    End If
                
                    enfymat = (salint1 - (empresa.sm * 3)) * (0.4 / 100): Rem EXCEDENTE
                End If
                
                enfymat = enfymat + (salint1 * ((0.25 + 0.375) / 100)): Rem PRESTACIONES EN ESPECIE Y DINERO
                
                If integrado > (empresa.sm * 25) Then
                    salint2 = (empresa.sm * 25)
                Else
                    salint2 = integrado
                End If
                
                If integrado > (empresa.sm * 1.0561 + 0.01) Then
                    otrascu = salint2 * ((0.625 + 1.125) / 100): Rem INVALIDEZ VIDA Y CESANTIA Y VEJEZ
                End If
                                
                seguro = diaseg * (enfymat + otrascu)
                
                If (p_vacacional <= 0 And nomina.aguin <= 0 And nomina.ptu <= 0) Then
                    Put #2, rgtro, personal 'Guarda el registro si se captura una nomina normal
                End If
    End If
End Sub
 Sub colocar(ancho2, valor$, us_o As String)
     ancho2 = 0
     ancho = Printer.TextWidth(valor$)
     ancho1 = Printer.TextWidth(us_o)
     ancho2 = ancho1 - ancho
     Rem Printer.CurrentX = Printer.currex + ancho2
 End Sub
Sub centrar(ancho2, micadena As String, anchototal As Long)
     ancho2 = 0
     ancho = Printer.TextWidth(micadena) / 2
     ancho1 = anchototal / 2
     ancho2 = ancho1 - ancho
End Sub

Sub factor(antig, facto) 'vacaciones 2023
      
      Select Case antig
        Case 1
            vac = 1.0493
        Case 2
            vac = 1.0507
        Case 3
            vac = 1.0521
        Case 4
            vac = 1.0534
        Case 5
            vac = 1.0548
        Case 6 To 10
            vac = 1.0562
        Case 11 To 15
            vac = 1.0575
        Case 16 To 20
            vac = 1.0589
        Case 21 To 25
            vac = 1.0603
        Case 26 To 30
            vac = 1.0616
        Case 31 To 35
            vac = 1.063
        Case Else
            vac = 1.063
      End Select
      facto = vac

End Sub
Sub calculo_compl2(base, impto, psub)
'   Stop
     Close 3: Close 4: Close 5
  If empresa.ao < 2008 Then
    Open (Dir_imptos + "TABLA.ISR") For Random As #3 Len = Len(articulo)
    Dem = LOF(3) / Len(articulo)
    Open (Dir_imptos + "TABLA.SUB") For Random As #4 Len = Len(subsidio)
    EM = LOF(4) / Len(subsidio)
    Open (Dir_imptos + "TABLA.CRE") For Random As #5 Len = Len(credito)
    eem = LOF(5) / Len(credito)
    
    Rem  **** CALCULAR IMPUESTO ****
    For i = 1 To Dem: Get 3, i, articulo
     If base > (articulo.liminf) And base < (articulo.limsup) Then
           marginal = ((base - articulo.liminf) * (articulo.porcsl / 100))
           impto = marginal + articulo.cuotaf
           Imp_mag = (articulo.porcsl / 100)
           nom_com.ImpTot = impto: Rem **********************************
           i = Dem
     End If
    Next i
    For i = 1 To EM: Get 4, i, subsidio
     If base > (subsidio.liminfs) And base < (subsidio.limsups) Then
           marginal2 = ((base - subsidio.liminfs) * Imp_mag)
           subdio = 0
           subdio = (marginal2 * subsidio.porcsls / 100) + (subsidio.cuotafs)
           nom_com.subdio = subdio
           nom_com.PSubDi = psub
           subdio = (subdio * psub)
           
           nom_com.subapl = subdio: Rem ************************************
           nom_com.subNap = nom_com.subdio - nom_com.subapl: Rem ************
           i = EM
     End If
    Next i
    
    For i = 1 To eem: Get 5, i, credito
     If base > (credito.crede) And base < credito.crea Then
         creere = (credito.cresam)
         nom_com.CreTot = creere: Rem ****************************************
     End If
    Next i
    
    impto = impto - subdio - creere
    If impto > 0 Then
        impto = impto + creere
        Else
        If (impto < 0) And (DifImptos.CalcDoble = 0) Then impto = 0
    End If
    Rem ********************************************************************************************
    Rem *********************************  Aqui es 2008 ********************************************
    Rem ********************************************************************************************
    Rem impto = impto - subdio
Else
     Close 3: Close 4: Close 5
        Open (Dir_imptos + "Tab08Mes.ISR") For Random As #3 Len = Len(articulo)
        Dem = LOF(3) / Len(articulo)
        Open (Dir_imptos + "Tab08Mes.SUB") For Random As #4 Len = Len(subsidio)
        EM = LOF(4) / Len(subsidio)
        
        For i = 1 To Dem: Get 3, i, articulo
        If base > (articulo.liminf) And base < (articulo.limsup) Then
           marginal = ((base - articulo.liminf) * (articulo.porcsl / 100))
           Imp_mag = (articulo.porcsl / 100)
           impto = marginal + articulo.cuotaf
           nom_com.ImpTot = impto
           Exit For
        End If
        Next i
        For i = 1 To EM: Get 4, i, subsidio
        If base > (subsidio.liminfs) And base < (subsidio.limsups) Then

           subdio = subsidio.cuotafs
           subdio = 0
           nom_com.subdio = subdio
           subdio = subdio
           nom_com.subapl = subdio
           nom_com.subNap = 0
           Exit For
        End If
        Next i
    impto = impto - subdio
    
    If impto < 0 Then
        nom_com.CredNe = impto
    Else
        nom_com.CredNe = 0
    End If
        Put 14, regtro, nom_com
    End If
  
  End Sub
  


