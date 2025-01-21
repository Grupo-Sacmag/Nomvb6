Attribute VB_Name = "percepdeduc"
Public sue3 As Currency 'sueldos
Public via7 As Currency 'viaticos
'esta madre ni se ocupa lo capturan en otras
Public pva8 As Currency 'prima vacacional
Public otr9 As Currency 'otros ingresos
Public pee10 As Currency 'percepcion exenta
Public sub13 As Currency ' subsidio pagado
Public agui5 As Currency ' aguinaldo

Public subc13 As Currency 'subsidio caisado

Public isr12 As Currency 'isr
Public ims14 As Currency 'imss
Public ifv18 As Currency 'infonavit
Public pea17 As Currency 'pension alimenticia
Public fon16 As Currency 'fonacot
Public pre15 As Currency 'prestamos otros


Public t_oi As Currency 'total otros ingresos

Public t_per As Currency 'total de percepciones
Public t_ded As Currency 'total de deducciones
Public T_neto As Currency 'salario neto

Public t_grav As Currency   'total gravado percepciones
Public t_ext As Currency  'total  exento percepciones
Public t_gravded As Currency 'total gravado deducciones
Public t_extded As Currency 'total exento deducciones
Public t_oded As Currency 'total otras deducciones
Public ptu_1 As Currency, ptu_2 As Currency, ptu_3 As Currency
Public I7 As Long



Sub reng()
Dim c_per As Integer 'contador percepciones
Dim c_ded As Integer 'contador de dedducciones

t_per = 0
t_ded = 0
c_per = 0
c_ded = 0

 sue3 = 0
 via7 = 0
 pva8 = 0
 otr9 = 0
 pee10 = 0
 sub13 = 0
 ptu_1 = 0
 ptu_2 = 0
 ptu_3 = 0
 isr12 = 0
 ims14 = 0
 pre15 = 0
 fon16 = 0
 pea17 = 0
 ifv18 = 0
 
 t_oi = 0
 T_neto = 0
 
 t_grav = 0
 t_ext = 0
 t_gravded = 0
 t_extded = 0
 
 t_oded = 0
 
 'hibrido---------------------------------------------------------

If IsNumeric(Form8.ConNom1.TextMatrix(I7, 13)) Then
    sub13 = (Form8.ConNom1.TextMatrix(I7, 13) * -1)
Else
    sub13 = 0
End If

If IsNumeric(Form8.ConNom1.TextMatrix(I7, 0)) Then
 Get 14, Form8.ConNom1.TextMatrix(I7, 0), nom_com
            subc13 = nom_com.subdio
End If
 
 'If (subc13 > 0) And (sub13 = 0) Then
    'sub13 = 0.01
    't_per = t_per + sub13
    't_oi = t_oi '+ subc13
'End If

'percepciones---------------------------------------------------
If IsNumeric(Form8.ConNom1.TextMatrix(I7, 3)) Then
    sue3 = Form8.ConNom1.TextMatrix(I7, 3)
    c_per = c_per + 1:  t_per = t_per + sue3
Else
    sue3 = 0
End If

'Parche Aguinaldp 19/12/2017------------------------------------------
If N_ormal = 1 Then
    If IsNumeric(Form8.ConNom1.TextMatrix(I7, 5)) Then
        agui5 = Form8.ConNom1.TextMatrix(I7, 5)
        c_per = c_per + 1:  t_per = t_per + agui5
    Else
        agui5 = 0
    End If
End If
'---------------------------------------------------------------------

If IsNumeric(Form8.ConNom1.TextMatrix(I7, 9)) Then
    otr9 = Form8.ConNom1.TextMatrix(I7, 9)
    sue3 = sue3 ' + otr9
    t_per = t_per + otr9
    'otr9 = 0
Else
    otr9 = 0
End If

If IsNumeric(Form8.ConNom1.TextMatrix(I7, 7)) Then
    
    via7 = Form8.ConNom1.TextMatrix(I7, 7)
    sue3 = sue3 '+ via7
    t_per = t_per + via7
    'via7 = 0
Else
    via7 = 0
End If

'If (via7 > 0) Or (otr9 > 0) Then
    'c_per = c_per + 1:  t_per = t_per + otr9 + via7
    't_oi = t_oi + otr9 + via7
'End If



If IsNumeric(Form8.ConNom1.TextMatrix(I7, 8)) Then
    pva8 = Form8.ConNom1.TextMatrix(I7, 8)
    c_per = c_per + 1:
Else
    pva8 = 0
End If

If IsNumeric(Form8.ConNom1.TextMatrix(I7, 10)) Then
    pee10 = Form8.ConNom1.TextMatrix(I7, 10)
    Rem c_per = c_per + 1: t_per = t_per + pee10 **********   aqui se duplica el exento ***********************************
Else
    pee10 = 0
End If


If pva8 > 0 Then
    If pee10 > 0 Then
        t_per = t_per + pva8 + pee10
    Else
        If pva8 > (empresa.sm * 15) Then
            pee10 = (empresa.sm * 15)
            pva8 = pva8 - pee10
            t_per = t_per + pva8 + pee10
        Else
            pee10 = (pva8 - 0.01)
            pva8 = 0.01
            t_per = t_per + pva8 + pee10
        End If
    End If
    Else
    t_per = t_per + pee10
End If
Rem ******************* PARCHE DE PTU DEL 6/6/17 **********************************
If N_ormal = 1 Then
  If IsNumeric(Form8.ConNom1.TextMatrix(I7, 6)) Then
    ptu_1 = Form8.ConNom1.TextMatrix(I7, 6)
     If IsNumeric(Form8.ConNom1.TextMatrix(I7, 10)) Then
        ptu_2 = Form8.ConNom1.TextMatrix(I7, 10)
        Else
        ptu_2 = empresa.sm * 15
        If ptu_2 >= ptu_1 Then
            ptu_2 = ptu_1
            ptu_1 = 0
           Else
            ptu_1 = ptu_1 - ptu_2
        End If
     End If
     ptu_3 = ptu_1 + ptu_2
     t_per = ptu_3
     pee10 = ptu_2
  End If
End If
Rem ******************************************************************************

t_grav = t_per - pee10
t_ext = pee10

'deducciones-----------------------------------------------------
If IsNumeric(Form8.ConNom1.TextMatrix(I7, 12)) Then
    isr12 = Form8.ConNom1.TextMatrix(I7, 12)
    c_ded = c_ded + 1: t_ded = t_ded + isr12
Else
    If sub13 > 0 Then
        sub13 = sub13 - 0.01: t_per = t_per + sub13
        isr12 = 0.01: t_ded = t_ded + isr12
    End If
End If
Rem MODIFICADO CON EL PARCHE DEL 6/6/17 ******************
If isr12 = 0 Then isr12 = 0.01: t_ded = t_ded + isr12
Rem ******************************************************
If IsNumeric(Form8.ConNom1.TextMatrix(I7, 14)) Then
    ims14 = Form8.ConNom1.TextMatrix(I7, 14)
    c_ded = c_ded + 1: t_ded = t_ded + ims14
Else
    ims14 = 0
End If

If IsNumeric(Form8.ConNom1.TextMatrix(I7, 15)) Then
    pre15 = Form8.ConNom1.TextMatrix(I7, 15)
    c_ded = c_ded + 1: t_ded = t_ded + pre15
Else
    pre15 = 0
End If
    
If IsNumeric(Form8.ConNom1.TextMatrix(I7, 16)) Then
    fon16 = Form8.ConNom1.TextMatrix(I7, 16)
    c_ded = c_ded + 1: t_ded = t_ded + fon16
Else
    fon16 = 0
End If

If IsNumeric(Form8.ConNom1.TextMatrix(I7, 17)) Then
    pea17 = Form8.ConNom1.TextMatrix(I7, 17)
    c_ded = c_ded + 1: t_ded = t_ded + pea17
Else
    pea17 = 0
End If

If IsNumeric(Form8.ConNom1.TextMatrix(I7, 18)) Then
    ifv18 = Form8.ConNom1.TextMatrix(I7, 18)
    c_ded = c_ded + 1: t_ded = t_ded + ifv18
Else
    ifv18 = 0
End If

t_oded = t_ded - isr12
T_neto = t_per - t_ded
t_extded = t_ded - isr12


End Sub
