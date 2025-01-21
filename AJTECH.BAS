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
Public formajte As ajform, cmaj
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

