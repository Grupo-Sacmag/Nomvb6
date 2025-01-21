Attribute VB_Name = "Module1"
Type per
    nom As String * 20
    ape1 As String * 20
    ape2 As String * 20
    rfc As String * 18
    imss As String * 18
    fal As String * 12
    fab As String * 12
    ingr As Currency
    viat As Currency
    otras As Currency
    integrado As Currency
 End Type
 Type basini
     datoarch As String * 64
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
       CTA As String * 12
       dias As Integer
       clte As String * 12
  End Type
 Type ult
     num As Long
     ubi As Integer
     renglon As Long
     texto As String
     poliza As Integer
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
 Public NxF As NominaXFuera, NxFTmP As NominaXFuera
 Public ReferOper As refo
 Public li As Long, Imto_deotra As Currency, P_sub1 As Currency
 Public DifImptos As imptosp, Tot_dias As Currency, Crd_deotra As Currency, Crpag_deotra As Currency
 Public Dat_ide As da_id, SubT_Mes As String
 Public Clbnx As Clabnx
 Public CATAUX As CAT_AX
 Public basico As basini
  Public obras As ob
 Public maestro As ob, creere As Currency
 Public empresa As empre, ingresos As Currency
 Public nomina As nom, deducciones As Currency
 Public personal As per, neto As Currency
 Public rgtro As Integer, CREDITO_PROV As Currency
 Public articulo As art, final As Long, APLICAR
 Public subsidio As subs, impto As Currency
 Public credito As cred, base As Currency, Sub_Aplic1 As Currency, ImtoTo_otra As Currency
 Public z1$, mm(12) As String * 20, cm, Dm, z2$, ddm
 Public subdirectorio$, valor$, dd(12) As Integer
 Public Dir_imptos As String, regtro As Integer
 Public arch_tr As String, tar As Integer, Direc_torio
 Public dir_obras As String, baseanual As Currency, cal_anual As Integer, baseor As Currency
 Public obra(22) As Currency, porcentaje(22) As Currency, Base_anual1 As Currency, Psub_Extra As Currency
 Sub veridir()
 On Error GoTo corrdire
    Dir_imptos = DirecT_arifas
    Open "C:\Archivos de programa\NOMINA1\perma.dno" For Random As #7 Len = Len(basico)
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
         Open "C:\Archivos de programa\NOMINA1\perma.dno" For Random As #7 Len = Len(basico)
         fin_basico = LOF(7) / Len(basico)
         If fin_basico > 1 Then
           Get 7, 2, basico
           dir_obras = RTrim(basico.datoarch)
         End If
      End If
      Close 7
      
      GoTo saltoerror
corrdire:
   Close 7
   ChDir "C:\Archivos de programa\NOMINA1"
   Exit Sub
saltoerror:
   
 End Sub
  Sub detbase()
     Close 10: SUMA_CREDITO_MES = 0
     Open "AcuNom2.dno" For Random As 10 Len = Len(ArAcum)
     lq_2 = LOF(10) / Len(ArAcum)
     If lq_2 = 0 Then
          MsgBox "No Existe archivo acumulado para calcular la ultima nomina", vbCritical + vbDefaultButton1, "Captura Nomina"
          Else
          Get 10, rgtro, ArAcum
          
          Acumul_virtual
          vaca = ArAcum.Pvaca
          If vaca < (empresa.sm * 15) Then
                vaca = 0
                Else
                vaca = ArAcum.Pvaca - (empresa.sm * 15)
          End If
          agu_i = ArAcum.Pagui
          If agu_i < (empresa.sm * 30) Then
                agu_i = 0
                Else
                agu_i = ArAcum.Pagui - (empresa.sm * 30)
                
          End If
          pt_u = ArAcum.PPTU
          If pt_u < (empresa.sm * 15) Then
                pt_u = 0
                Else
                pt_u = ArAcum.PPTU - (empresa.sm * 15)
          End If
          
          baseanual = ArAcum.Pnormal + ArAcum.Pextra + ArAcum.Potras + ArAcum.Pviaticos + vaca + agu_i + pt_u
          
          SUMA_CREDITO_MES = ArAcum.DCrApl
      End If
       Rem P_sub1 = (Sub_Aplic1 + ArAcum.DSubioAp) / (psub_extra + ArAcum.DSubioAp + ArAcum.DSubNoap)
       
End Sub
Sub Acumul_virtual()
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
          If vaca < (empresa.sm * 15) Then
                vaca = 0
                Else
                vaca = ArAcum.Pvaca - (empresa.sm * 15)
          End If
          agu_i = ArAcum.Pagui
          If agu_i < (empresa.sm * 30) Then
                agu_i = 0
                Else
                agu_i = ArAcum.Pagui - (empresa.sm * 30)
                
          End If
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
    Dim ISR_1 As Currency, SUB_1 As Currency, CRED_1 As Currency
    calculo base, impto, psub
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
    mientras = psub
GoTo SuBsidio_Salto
    For i = 1 To EM: Get 4, i, subsidio
     If P_sub1 > 0 Then psub = P_sub1
     If base > (subsidio.liminfs) And base < (subsidio.limsups) Then
           marginal2 = ((base - subsidio.liminfs) * Imp_mag)
           subdio = (marginal2 * subsidio.porcsls / 100) + (subsidio.cuotafs)
           Rem subdio = (marginal * subsidio.porcsls / 100) + (subsidio.cuotafs)
           subdiono = subdio
           subdio = (subdio * psub)
           subdiono = subdiono - subdio
           nom_com.subapl = subdio - ArAcum.DSubioAp - Sub_Aplic1
           nom_com.subNap = subdiono - ArAcum.DSubNoap - (Psub_Extra - Sub_Aplic1)
           nom_com.subdio = nom_com.subapl + nom_com.subNap
           nom_com.PSubDi = psub
           i = EM
     End If
    Next i
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
  If empresa.ao < 2008 Then
        Open (Dir_imptos + "TABLA.ISR") For Random As #3 Len = Len(articulo)
        Dem = LOF(3) / Len(articulo)
        Open (Dir_imptos + "TABLA.SUB") For Random As #4 Len = Len(subsidio)
        EM = LOF(4) / Len(subsidio)
        Open (Dir_imptos + "TABLA.CRE") For Random As #5 Len = Len(credito)
        eem = LOF(5) / Len(credito)
        nom_com.ArchImp = Dir_imptos
        Rem  **** CALCULAR IMPUESTO ****
    For i = 1 To Dem: Get 3, i, articulo
     If base > (articulo.liminf / 2) And base < (articulo.limsup / 2) Then
           marginal = ((base - articulo.liminf / 2) * (articulo.porcsl / 100))
           Imp_mag = (articulo.porcsl / 100)
           impto = marginal + articulo.cuotaf / 2
           nom_com.ImpTot = impto
           i = Dem
     End If
    Next i
    For i = 1 To EM: Get 4, i, subsidio
     If base > (subsidio.liminfs / 2) And base < (subsidio.limsups / 2) Then
           marginal2 = ((base - subsidio.liminfs / 2) * Imp_mag)
           subdio = (marginal2 * subsidio.porcsls / 100) + (subsidio.cuotafs / 2)
           nom_com.subdio = subdio
           nom_com.PSubDi = psub
           subdio = (subdio * psub)
           nom_com.subapl = subdio
           nom_com.subNap = nom_com.subdio - nom_com.subapl
           i = EM
     End If
    Next i
    For i = 1 To eem: Get 5, i, credito
     If base > (credito.crede / 2) And base < credito.crea / 2 Then
         creere = (credito.cresam / 2)
         nom_com.CreTot = creere
         
     End If
    Next i
    

    CREDITO_PROV = creere
    impto = impto - subdio - creere
    If impto < 0 Then
        nom_com.CredNe = impto
        Rem nom_com.CredNe = 0
        Else
        nom_com.CredNe = 0
    End If
    
    
    Put 14, regtro, nom_com
   Rem *****************************************************************************************************
   Rem ********************************  A PARTIR DE AQUI ES 2008 ******************************************
   Rem *****************************************************************************************************
   Else
        Close 3: Close 4: Close 5
        Open (Trim(Dir_imptos) + "Tab08Kin.ISR") For Random As #3 Len = Len(articulo)
        Dem = LOF(3) / Len(articulo)
        Open (Trim(Dir_imptos) + "Tab08Kin.SUB") For Random As #4 Len = Len(subsidio)
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
           Rem marginal2 = ((base - subsidio.liminfs) * Imp_mag)
           Rem subdio = (marginal2 * subsidio.porcsls / 100) + (subsidio.cuotafs)
           subdio = subsidio.cuotafs
           nom_com.subdio = subdio
           Rem nom_com.PSubDi = psub
           subdio = subdio
           nom_com.subapl = subdio
           nom_com.subNap = 0
           Exit For
        End If
        Next i
        Rem CREDITO_PROV = creere
    impto = impto - subdio
    If impto < 0 Then
        nom_com.CredNe = impto
        Rem nom_com.CredNe = 0
        Else
        nom_com.CredNe = 0
    End If
     Put 14, regtro, nom_com
    End If
    
End Sub
Sub imss(integrado As Currency, seguro As Currency, diaseg As Currency)
    imss2 integrado, seguro, diaseg
    'salint1 = 0: salint2 = 0
    'enfymat = 0: otrascu = 0
    'If integrado > (empresa.sm * 3) Then
       'If integrado > (empresa.sm * 25) Then
                'salint1 = (empresa.sm * 25)
                'Else
                'salint1 = integrado
       'End If
       'enfymat = (salint1 - (empresa.sm * 3)) * (0.72 / 100): Rem EXCEDENTE
       Rem *** DE TRES SMM
    'End If
    'enfymat = enfymat + (salint1 * ((0.25 + 0.375) / 100)): Rem PRESTACIONES EN ESPECIE Y DINERO
    'If integrado > (empresa.sm * 24) Then salint2 = (empresa.sm * 24) Else salint2 = integrado
    'If integrado > (empresa.sm * 1.0561 + 0.01) Then
              'otrascu = salint2 * ((0.625 + 1.125) / 100): Rem INVALIDEZ VIDA Y CESANTIA Y VEJEZ
              Rem SALARIO INTEGRADO
    'End If
    'seguro = diaseg * (enfymat + otrascu)
    
End Sub
Sub imss2(integrado As Currency, seguro As Currency, diaseg As Currency)
    Dim Anual1 As Integer
        salint1 = 0: salint2 = 0
        enfymat = 0: otrascu = 0
        salint1 = integrado: salint2 = 0
        seguro = 0
        Anual1 = empresa.ao
        
        Select Case Anual1
            Case 1997
                salint1 = 0: salint2 = 0
                enfymat = 0: otrascu = 0
                salint1 = integrado: salint2 = 0
                If integrado > (empresa.sm * 3) Then
                    If integrado > (empresa.sm * 25) Then
                        salint1 = (empresa.sm * 25)
                        Else
                        salint1 = integrado
                    End If
                    enfymat = (salint1 - (empresa.sm * 3)) * (2 / 100): Rem EXCEDENTE
                    Rem *** DE TRES SMM
                End If
                enfymat = enfymat + (salint1 * ((0.25 + 0.375) / 100)): Rem PRESTACIONES EN ESPECIE Y DINERO
                If integrado > (empresa.sm * 15) Then salint2 = (empresa.sm * 15) Else salint2 = integrado
                If integrado > (empresa.sm * 1.0561 + 0.01) Then
                    otrascu = salint2 * ((0.625 + 1.125) / 100): Rem INVALIDEZ VIDA Y CESANTIA Y VEJEZ
              Rem SALARIO INTEGRADO
                End If
                seguro = diaseg * (enfymat + otrascu)

            Case 1998
                salint1 = 0: salint2 = 0
                enfymat = 0: otrascu = 0
                salint1 = integrado
                If integrado > (empresa.sm * 3) Then
                    If integrado > (empresa.sm * 25) Then
                        salint1 = (empresa.sm * 25)
                        Else
                        salint1 = integrado
                    End If
                    enfymat = (salint1 - (empresa.sm * 3)) * (2 / 100): Rem EXCEDENTE
                    Rem *** DE TRES SMM
                End If
                enfymat = enfymat + (salint1 * ((0.25 + 0.375) / 100)): Rem PRESTACIONES EN ESPECIE Y DINERO
                If integrado > (empresa.sm * 16) Then salint2 = (empresa.sm * 16) Else salint2 = integrado
                If integrado > (empresa.sm * 1.0561 + 0.01) Then
                    otrascu = salint2 * ((0.625 + 1.125) / 100): Rem INVALIDEZ VIDA Y CESANTIA Y VEJEZ
              Rem SALARIO INTEGRADO
                End If
                seguro = diaseg * (enfymat + otrascu)
            Case 1999
                salint1 = 0: salint2 = 0
                enfymat = 0: otrascu = 0
                salint1 = integrado
                If integrado > (empresa.sm * 3) Then
                    If integrado > (empresa.sm * 25) Then
                        salint1 = (empresa.sm * 25)
                        Else
                        salint1 = integrado
                    End If
                    enfymat = (salint1 - (empresa.sm * 3)) * (1.84 / 100): Rem EXCEDENTE
                    Rem *** DE TRES SMM
                End If
                enfymat = enfymat + (salint1 * ((0.25 + 0.375) / 100)): Rem PRESTACIONES EN ESPECIE Y DINERO
                If integrado > (empresa.sm * 17) Then salint2 = (empresa.sm * 17) Else salint2 = integrado
                If integrado > (empresa.sm * 1.0561 + 0.01) Then
                    otrascu = salint2 * ((0.625 + 1.125) / 100): Rem INVALIDEZ VIDA Y CESANTIA Y VEJEZ
              Rem SALARIO INTEGRADO
                End If
                seguro = diaseg * (enfymat + otrascu)
            Case 2000
                salint1 = 0: salint2 = 0
                enfymat = 0: otrascu = 0
                salint1 = integrado
                If integrado > (empresa.sm * 3) Then
                    If integrado > (empresa.sm * 25) Then
                        salint1 = (empresa.sm * 25)
                        Else
                        salint1 = integrado
                    End If
                    enfymat = (salint1 - (empresa.sm * 3)) * (1.66 / 100): Rem EXCEDENTE
                    Rem *** DE TRES SMM
                End If
                enfymat = enfymat + (salint1 * ((0.25 + 0.375) / 100)): Rem PRESTACIONES EN ESPECIE Y DINERO
                If integrado > (empresa.sm * 18) Then salint2 = (empresa.sm * 18) Else salint2 = integrado
                If integrado > (empresa.sm * 1.0561 + 0.01) Then
                    otrascu = salint2 * ((0.625 + 1.125) / 100): Rem INVALIDEZ VIDA Y CESANTIA Y VEJEZ
              Rem SALARIO INTEGRADO
                End If
                seguro = diaseg * (enfymat + otrascu)

            Case 2001
                salint1 = 0: salint2 = 0
                enfymat = 0: otrascu = 0
                salint1 = integrado: seguro = 0
                
                If integrado > (empresa.sm * 3) Then
                    If integrado > (empresa.sm * 25) Then
                        salint1 = (empresa.sm * 25)
                        Else
                        salint1 = integrado
                    End If
                    enfymat = (salint1 - (empresa.sm * 3)) * (1.52 / 100): Rem EXCEDENTE
                    Rem *** DE TRES SMM
                End If
                enfymat = enfymat + (salint1 * ((0.25 + 0.375) / 100)): Rem PRESTACIONES EN ESPECIE Y DINERO
                If integrado > (empresa.sm * 19) Then salint2 = (empresa.sm * 19) Else salint2 = integrado
                If integrado > (empresa.sm * 1.0561 + 0.01) Then
                    otrascu = salint2 * ((0.625 + 1.125) / 100): Rem INVALIDEZ VIDA Y CESANTIA Y VEJEZ
              Rem SALARIO INTEGRADO
                End If
                seguro = diaseg * (enfymat + otrascu)

            Case 2002
                salint1 = 0: salint2 = 0
                enfymat = 0: otrascu = 0
                salint1 = integrado
                If integrado > (empresa.sm * 3) Then
                    If integrado > (empresa.sm * 25) Then
                        salint1 = (empresa.sm * 25)
                        Else
                        salint1 = integrado
                    End If
                    enfymat = (salint1 - (empresa.sm * 3)) * (1.36 / 100): Rem EXCEDENTE
                    Rem *** DE TRES SMM
                End If
                enfymat = enfymat + (salint1 * ((0.25 + 0.375) / 100)): Rem PRESTACIONES EN ESPECIE Y DINERO
                If integrado > (empresa.sm * 20) Then salint2 = (empresa.sm * 20) Else salint2 = integrado
                If integrado > (empresa.sm * 1.0561 + 0.01) Then
                    otrascu = salint2 * ((0.625 + 1.125) / 100): Rem INVALIDEZ VIDA Y CESANTIA Y VEJEZ
              Rem SALARIO INTEGRADO
                End If
                seguro = diaseg * (enfymat + otrascu)

            Case 2003
                salint1 = 0: salint2 = 0
                enfymat = 0: otrascu = 0
                salint1 = integrado
                If integrado > (empresa.sm * 3) Then
                    If integrado > (empresa.sm * 25) Then
                        salint1 = (empresa.sm * 25)
                        Else
                        salint1 = integrado
                    End If
                    enfymat = (salint1 - (empresa.sm * 3)) * (1.2 / 100): Rem EXCEDENTE
                    Rem *** DE TRES SMM
                End If
                enfymat = enfymat + (salint1 * ((0.25 + 0.375) / 100)): Rem PRESTACIONES EN ESPECIE Y DINERO
                If integrado > (empresa.sm * 21) Then salint2 = (empresa.sm * 21) Else salint2 = integrado
                If integrado > (empresa.sm * 1.0561 + 0.01) Then
                    otrascu = salint2 * ((0.625 + 1.125) / 100): Rem INVALIDEZ VIDA Y CESANTIA Y VEJEZ
              Rem SALARIO INTEGRADO
                End If
                seguro = diaseg * (enfymat + otrascu)

            Case 2004
                salint1 = 0: salint2 = 0
                enfymat = 0: otrascu = 0
                salint1 = integrado
                If integrado > (empresa.sm * 3) Then
                    If integrado > (empresa.sm * 25) Then
                        salint1 = (empresa.sm * 25)
                        Else
                        salint1 = integrado
                    End If
                    enfymat = (salint1 - (empresa.sm * 3)) * (1.04 / 100): Rem EXCEDENTE
                    Rem *** DE TRES SMM
                End If
                enfymat = enfymat + (salint1 * ((0.25 + 0.375) / 100)): Rem PRESTACIONES EN ESPECIE Y DINERO
                If integrado > (empresa.sm * 22) Then salint2 = (empresa.sm * 22) Else salint2 = integrado
                If integrado > (empresa.sm * 1.0561 + 0.01) Then
                    otrascu = salint2 * ((0.625 + 1.125) / 100): Rem INVALIDEZ VIDA Y CESANTIA Y VEJEZ
              Rem SALARIO INTEGRADO
                End If
                seguro = diaseg * (enfymat + otrascu)
            Case 2005
                salint1 = 0: salint2 = 0
                enfymat = 0: otrascu = 0
                salint1 = integrado
                If integrado > (empresa.sm * 3) Then
                    If integrado > (empresa.sm * 25) Then
                        salint1 = (empresa.sm * 25)
                        Else
                        salint1 = integrado
                    End If
                    enfymat = (salint1 - (empresa.sm * 3)) * (0.88 / 100): Rem EXCEDENTE
                    Rem *** DE TRES SMM
                End If
                enfymat = enfymat + (salint1 * ((0.25 + 0.375) / 100)): Rem PRESTACIONES EN ESPECIE Y DINERO
                If integrado > (empresa.sm * 23) Then salint2 = (empresa.sm * 23) Else salint2 = integrado
                If integrado > (empresa.sm * 1.0561 + 0.01) Then
                    otrascu = salint2 * ((0.625 + 1.125) / 100): Rem INVALIDEZ VIDA Y CESANTIA Y VEJEZ
              Rem SALARIO INTEGRADO
                End If
                seguro = diaseg * (enfymat + otrascu)

            Case 2006
             salint1 = 0: salint2 = 0
             salint1 = integrado
                enfymat = 0: otrascu = 0
                If integrado > (empresa.sm * 3) Then
                    If integrado > (empresa.sm * 25) Then
                        salint1 = (empresa.sm * 25)
                        Else
                        salint1 = integrado
                    End If
                    enfymat = (salint1 - (empresa.sm * 3)) * (0.72 / 100): Rem EXCEDENTE
                    Rem *** DE TRES SMM
                End If
                enfymat = enfymat + (salint1 * ((0.25 + 0.375) / 100)): Rem PRESTACIONES EN ESPECIE Y DINERO
                If integrado > (empresa.sm * 24) Then salint2 = (empresa.sm * 24) Else salint2 = integrado
                If integrado > (empresa.sm * 1.0561 + 0.01) Then
                    otrascu = salint2 * ((0.625 + 1.125) / 100): Rem INVALIDEZ VIDA Y CESANTIA Y VEJEZ
              Rem SALARIO INTEGRADO
                End If
                seguro = diaseg * (enfymat + otrascu)
                
            Case 2007
                
                salint1 = 0: salint2 = 0
                enfymat = 0: otrascu = 0
                salint1 = integrado
                If integrado > (empresa.sm * 3) Then
                    If integrado > (empresa.sm * 25) Then
                        salint1 = (empresa.sm * 25)
                        Else
                        salint1 = integrado
                    End If
                    enfymat = (salint1 - (empresa.sm * 3)) * (0.56 / 100): Rem EXCEDENTE
                    Rem *** DE TRES SMM
                End If
                enfymat = enfymat + (salint1 * ((0.25 + 0.375) / 100)): Rem PRESTACIONES EN ESPECIE Y DINERO
                If integrado > (empresa.sm * 25) Then salint2 = (empresa.sm * 25) Else salint2 = integrado
                
                If integrado > (empresa.sm * 1.0561 + 0.01) Then
                    otrascu = salint2 * ((0.625 + 1.125) / 100): Rem INVALIDEZ VIDA Y CESANTIA Y VEJEZ
              Rem SALARIO INTEGRADO
                End If
                seguro = diaseg * (enfymat + otrascu)
                
            Case Else
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
                    Rem *** DE TRES SMM
                End If
                enfymat = enfymat + (salint1 * ((0.25 + 0.375) / 100)): Rem PRESTACIONES EN ESPECIE Y DINERO
                If integrado > (empresa.sm * 25) Then salint2 = (empresa.sm * 25) Else salint2 = integrado
                If integrado > (empresa.sm * 1.0561 + 0.01) Then
                    otrascu = salint2 * ((0.625 + 1.125) / 100): Rem INVALIDEZ VIDA Y CESANTIA Y VEJEZ
              Rem SALARIO INTEGRADO
                End If
                seguro = diaseg * (enfymat + otrascu)
        End Select
      'Debug.Print integrado; seguro; diaseg
      
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

Sub factor(antig, facto)
      Select Case antig
        Case 1
         vac = 1.0458333
        Case 2
         vac = 1.047222
        Case 3
         vac = 1.0486111
        Case 4
         vac = 1.05
        Case 5 To 9
         vac = 1.05139
        Case 10 To 14
         vac = 1.05278
        Case 15 To 19
         vac = 1.05417
        Case 20 To 24
         vac = 1.05556
        Case 25 To 29
         vac = 1.05694
        Case 31 To 34
            vac = 1.05833
        Case 35 To 39
            vac = 1.05972
        Case Else
            vac = 1.05972
      End Select
      facto = vac
End Sub
Sub calculo_compl2(base, impto, psub)
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
           Rem marginal2 = ((base - subsidio.liminfs) * Imp_mag)
           Rem subdio = (marginal2 * subsidio.porcsls / 100) + (subsidio.cuotafs)
           
           subdio = subsidio.cuotafs
           subdio = 0
           nom_com.subdio = subdio
           Rem nom_com.PSubDi = psub
           subdio = subdio
           nom_com.subapl = subdio
           nom_com.subNap = 0
           Exit For
        End If
        Next i
        Rem CREDITO_PROV = creere
    impto = impto - subdio
    If impto < 0 Then
        nom_com.CredNe = impto
        Rem nom_com.CredNe = 0
        Else
        nom_com.CredNe = 0
    End If
     Put 14, regtro, nom_com
    
  End If
  End Sub

