Attribute VB_Name = "Convertidor"
Public LetrVal
Function Aletra(Rcantidad As Double) As String
Dim Rcant As String
Dim cAux As String
Dim Runi As String
Dim Rdec As String
Dim rdecs As String
Dim rcen As String
Dim riter As Integer
Dim rnum As String
Dim cDecim As String

'Runi$ , Rdec$, Rdecs$, Rcen$, Rnum$, Riter$

Rcant = ""
Runi = "      UN    DOS   TRES  CUATROCINCO SEIS  SIETE OCHO  NUEVE "
Rdec = "DIEZ      ONCE      DOCE      TRECE     CATORCE   QUINCE    DIECISEIS DIECISIETEDIECIOCHO DIECINUEVE"
rdecs = "                  VEINTE   TREINTA  CUARENTA CINCUENTASESENTA  SETENTA  OCHENTA  NOVENTA "
rcen = "            DOS   TRES  CUATRO      SEIS  SETE  OCHO  NOVE  "

Rcant = Trim(Str(Rcantidad))
If InStr(1, Rcant, ".") > 0 Then
  'cAux = Left(Rcant, InStr(1, Rcant, ".") + 2)
  cAux = cRound(Rcant, 2)
  Rcant = cAux
  If Mid(Rcant, Len(Rcant) - 1, 1) = "." Then
    Rcant = Rcant + "0"
    Rcant = Space(12 - Len(Left(Rcant, Len(Rcant) - 3))) + Rcant
  Else
    Rcant = Space(12 - Len(Left(Rcant, Len(Rcant) - 3))) + Rcant
    cDecim = Right(Rcant, 2)
  End If
Else
  cDecim = "00"
End If
rnum = Mid(Rcant, 1, 12)
Rcant = ""
If Len(rnum) < 12 Then
  rnum = Space(12 - Len(rnum)) + rnum
End If
If Val(rnum) = 0 Then
  Rcant = "CERO PESOS "
Else
  riter = 1
  While riter < 13

    If Mid(rnum, riter, 1) <> " " And Mid(rnum, riter, 1) <> "0" Then
      Select Case Mid(rnum, riter, 1)
        Case "1"
          If Mid(rnum, riter + 1, 2) = "00" Then
            Rcant = Rcant + "CIEN "
          Else
            Rcant = Rcant + "CIENTO "
          End If
        Case "5"
          Rcant = Rcant + "QUINIENTOS "
        Case Else
          Rcant = Rcant + RTrim(Mid(rcen, Val(Mid(rnum, riter, 1)) * 6 + 1, 6)) + "CIENTOS "
      End Select
    End If

    If Mid(rnum, riter + 1, 1) <> " " And Mid(rnum, riter + 1, 1) <> "0" Then
      Select Case Mid(rnum, riter + 1, 1)
        Case "1"
          Rcant = Rcant + RTrim(Mid(Rdec, Val(Mid(rnum, riter + 2, 1)) * 10 + 1, 10)) + " "
        Case "2"
          If Mid(rnum, riter + 2, 1) = "0" Then
            Rcant = Rcant + "VEINTE "
          Else
            Rcant = Rcant + "VEINTI" + RTrim(Mid(Runi, Val(Mid(rnum, riter + 2, 1)) * 6 + 1, 6)) + " "
          End If
        Case Else
          Rcant = Rcant + RTrim(Mid(rdecs, Val(Mid(rnum, riter + 1, 1)) * 9 + 1, 9))
          If Mid(rnum, riter + 2, 1) > "0" Then
            Rcant = Rcant + " Y " + RTrim(Mid(Runi, Val(Mid(rnum, riter + 2, 1)) * 6 + 1, 6)) + " "
          Else
            Rcant = Rcant + " "
          End If
      End Select
    End If

    If Mid(rnum, riter + 2, 1) <> " " And Mid(rnum, riter + 1, 1) < "1" And Mid(rnum, riter + 1, 2) <> "00" Then
      Rcant = Rcant + RTrim(Mid(Runi, Val(Mid(rnum, riter + 2, 1)) * 6 + 1, 6)) + " "
    End If

    Select Case riter
      Case 1
        If Mid(rnum, 1, 3) <> Space(3) And Mid(rnum, 1, 3) <> "000" Then
          Rcant = Rcant + "MIL "
        End If
      Case 4
        If Mid(rnum, 1, 6) <> Space(6) And Mid(rnum, 1, 6) <> "000000" Then
          If Mid(rnum, 1, 6) <> Space(5) + "1" Then
            Rcant = Rcant + "MILLONES "
          Else
            Rcant = Rcant + "MILLON "
          End If
        End If
      Case 7
        If Mid(rnum, 1, 9) <> Space(9) And Mid(rnum, 7, 3) <> "000" Then
          Rcant = Rcant + "MIL "
        End If
    End Select
    riter = riter + 3
  Wend

  If rnum = Space(11) + "1" Then
    Rcant = Rcant + "PESO "
  Else
    If Mid(rnum, 7, 6) = "000000" Then
      Rcant = Rcant + "DE PESOS "
    Else
      Rcant = Rcant + "PESOS "
    End If
  End If
End If

Rcant = LTrim(RTrim((Rcant + cDecim + "/100 M. N.")))
Aletra = Rcant
LetrVal = Rcant
End Function

Function cRound(ByVal cVal, ByVal nDec) As String
Dim cAux, cRet As String
Dim nI, nPos, nAcum, nCurVal, nNextVal As Integer
    nAcum = 0
    nCurVal = 0
    nNextVal = 0
    cRet = ""
    nPos = InStr(1, cVal, ".")
    If nPos = 0 Then
        'cAux = Padc("", nDec, "0")
        cRet = cVal + "." + "00"
    Else
        cAux = Right(cVal, Len(cVal) - nPos)
        If Len(cAux) > nDec Then
            nPos = Len(cAux) - 1
            For nI = nPos To nDec Step -1
                nCurVal = Int(Val(Mid(cAux, nI + 1, 1)))
                nNextVal = Int(Val(Mid(cAux, nI, 1)))
                If nCurVal < 5 Then
                    nAcum = nNextVal
                Else
                    nAcum = nNextVal + 1
                End If
                cRet = Mid(cAux, 1, nI - 1) + Trim(Str(nAcum))
            Next
            nPos = InStr(1, cVal, ".")
            cRet = Left(cVal, nPos) + cRet
        Else
            nAcum = nDec - Len(Right(cVal, Len(cVal) - nPos))
            cRet = cVal
            For nI = 1 To nAcum
                cRet = cRet + "0"
            Next
        End If
    End If

    cRound = cRet
End Function
