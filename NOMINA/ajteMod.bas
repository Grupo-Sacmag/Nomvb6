Attribute VB_Name = "Module4"
Type ajform
     totr As Integer
     fechar As Integer
     fechac As Integer
     benefr As Integer
     benefc As Integer
     impnumr As Integer
     impnumc As Integer
     impletr As Integer
     impletc As Integer
     concepr As Integer
     concepc As Integer
     inicopr As Integer
     cta As Integer
     scta As Integer
     name As Integer
     aplicr As Integer
     parcialc As Integer
     debec As Integer
     haberc As Integer
     sumasr As Integer
End Type
Type nomco
     ArchImp As String * 50
     PSubDi As Currency
     subdio As Currency
     subapl As Currency
     subNap As Currency
     CreTot As Currency
     CredNe As Currency
     ImpTot As Currency
End Type
Type antnom
   inggrav As Currency
   imptoret As Currency
   credcalc As Currency
   credpag As Currency
   subapl As Currency
   subtotal As Currency
   subnoapl As Currency
End Type
Public NomAnterior As antnom
Public nom_com As nomco, Arch1 As String * 20, ArchAnterior, Arch1Anterior
Public formajte As ajform, cmaj, QUIN As Integer, Arch$
Sub apeajte()
    Close 8
    Open "cheque.ajt" For Random As 8 Len = Len(formajte)
    cmaj = LOF(8) / Len(formajte)
End Sub
Sub cargaaj()
    If cmaj < 1 Then Exit Sub
    Get 8, 1, formajte
    AJTECH.AJTE.TextMatrix(1, 1) = formajte.totr
    AJTECH.AJTE.TextMatrix(2, 1) = formajte.fechar
    AJTECH.AJTE.TextMatrix(2, 2) = formajte.fechac
    AJTECH.AJTE.TextMatrix(3, 1) = formajte.benefr
    AJTECH.AJTE.TextMatrix(3, 2) = formajte.benefc
    AJTECH.AJTE.TextMatrix(4, 1) = formajte.impnumr
    AJTECH.AJTE.TextMatrix(4, 2) = formajte.impnumc
    AJTECH.AJTE.TextMatrix(5, 1) = formajte.impletr
    AJTECH.AJTE.TextMatrix(5, 2) = formajte.impletc
    AJTECH.AJTE.TextMatrix(6, 1) = formajte.concepr
    AJTECH.AJTE.TextMatrix(6, 2) = formajte.concepc
    AJTECH.AJTE.TextMatrix(7, 1) = formajte.inicopr
    AJTECH.AJTE.TextMatrix(8, 1) = formajte.sumasr
    AJTECH.AJTE.TextMatrix(9, 2) = formajte.cta
    AJTECH.AJTE.TextMatrix(10, 2) = formajte.scta
    AJTECH.AJTE.TextMatrix(11, 2) = formajte.name
    AJTECH.AJTE.TextMatrix(12, 2) = formajte.parcialc
    AJTECH.AJTE.TextMatrix(13, 2) = formajte.debec
    AJTECH.AJTE.TextMatrix(14, 2) = formajte.haberc
End Sub
Sub archaj()
    formajte.totr = AJTECH.AJTE.TextMatrix(1, 1)
    formajte.fechar = AJTECH.AJTE.TextMatrix(2, 1)
    formajte.fechac = AJTECH.AJTE.TextMatrix(2, 2)
    formajte.benefr = AJTECH.AJTE.TextMatrix(3, 1)
    formajte.benefc = AJTECH.AJTE.TextMatrix(3, 2)
    formajte.impnumr = AJTECH.AJTE.TextMatrix(4, 1)
    formajte.impnumc = AJTECH.AJTE.TextMatrix(4, 2)
    formajte.impletr = AJTECH.AJTE.TextMatrix(5, 1)
    formajte.impletc = AJTECH.AJTE.TextMatrix(5, 2)
    formajte.concepr = AJTECH.AJTE.TextMatrix(6, 1)
    formajte.concepc = AJTECH.AJTE.TextMatrix(6, 2)
    formajte.inicopr = AJTECH.AJTE.TextMatrix(7, 1)
    formajte.sumasr = AJTECH.AJTE.TextMatrix(8, 1)
    formajte.cta = AJTECH.AJTE.TextMatrix(9, 2)
    formajte.scta = AJTECH.AJTE.TextMatrix(10, 2)
    formajte.name = AJTECH.AJTE.TextMatrix(11, 2)
    formajte.parcialc = AJTECH.AJTE.TextMatrix(12, 2)
    formajte.debec = AJTECH.AJTE.TextMatrix(13, 2)
    formajte.haberc = AJTECH.AJTE.TextMatrix(14, 2)
    Put 8, 1, formajte
End Sub
Sub todoslosarchivos()
  Rem  buffer 1 "Empresa.dno"
  Rem  buffer 2 "Personal.dno"
  Rem  buffer 3  dir_imptos + "Tabla.Isr" "Isr141"
  Rem  buffer 4  dir_imptos + "Tabla.Sub" "Isr141A"
  Rem  buffer 5  dir_imptos + "Tabla.Cre" "Isr141A"
  Rem  buffer 6  "NOMINA.NOM" ARCH$
  Rem  buffer 7  "perma.dno"
  Rem  buffer 8  "Maestro.dno" y "cheque.ajt"
  Rem  buffer 9  "Cataux"
  Rem  buffer 10 "Acumul.dno"
  Rem  buffer 11 "quin.txt"
  Rem  buffer 12 "bnxcla.dno"
  Rem  buffer 13 "Empcomp.dno"
  Rem  buffer 14 "com.nom" Arch1
  Rem buffer 15 "ArchAnterior"
  Rem buffer 15 "Arch1Anterior"
End Sub
Sub calculo_mes()
   
End Sub
Sub mescompleto()
  Rem Open ArchAnterior For Random As 15 Len = Len(nomina)
  Rem Open Arch1Anterior For Random As 16 Len = Len(nom_com)
  Rem Get 15, regtro, nomina
  Rem NomAnterior.inggrav
  
  

End Sub
