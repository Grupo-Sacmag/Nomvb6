Attribute VB_Name = "contabilidad"
 Type CAT_MA
    B1 As String * 6
    B2 As String * 32
    B3 As String * 16
    B4 As String * 5
    B5 As String * 5
End Type

Type DAT_OS
    D1 As String * 64
    D2 As String * 60
    D3 As String * 45
    No_arch As String * 15
    a_o As String * 5
    others1  As String * 25
    UltimaPol As String * 5
    UltimoReg As String * 5
    others As String * 12
End Type

Type oper_aciones
    CTA As String * 6
    descr As String * 30
    fe As String * 2
    impte As String * 16
    identi As String * 1
    real As String * 9
End Type

Type cheques
         num As Integer
         beneficiario As String * 40
         importe As Currency
         Clave As String * 1
         numreal As Integer
         refer As Integer
         conta As Integer
End Type

Type tra_cta
       num As Integer
       Nombre As String * 32
       donde As Integer
       inicia As Integer
       termina As Integer
       Clave As String * 1
 End Type

Type tra_Scta
       num As Integer
       Nombre As String * 32
       donde As Integer
       refer As Integer
       Clave As String * 1
 End Type
 
 Type ult
     num As Long
     ubi As Integer
     renglon As Long
     texto As String
     poliza As Integer
     Impresion As Integer
     TipoCap As Integer
     redaccion As String * 32
 End Type
 
 Type ContCat
    g1 As String * 5
    g2 As String * 5
    g3 As String * 6
 End Type
 
 Type Mvtos
     Inc As Currency
     Ene As Currency
     Feb As Currency
     Mar As Currency
     Abr As Currency
     May As Currency
     Jun As Currency
     Jul As Currency
     Ago As Currency
     Sep As Currency
     Oct As Currency
     Nov As Currency
     Dic As Currency
 End Type
 
 Type sc
    guarda As String * 64
End Type

Type Su
   Parcial As Currency
   Debe As Currency
   Haber As Currency
End Type

Type Al
    Cos As Integer
    Rda As String
    Imt As Currency
    Gto As Integer
    Dep As Integer
    Clt As Integer
    Otr As Integer
 End Type
 
 Public Alarma As Al
 Public Su_ma As Su, Ruta_Acceso, Ruta_Acceso_Contr, Arch_Oper As String
 Public Dir_Costos As String
 Public MvDebe As Mvtos, MvHaber As Mvtos, Sub_dir As String
 Public ContrCatal As ContCat, Arc_FinaL As Integer
 Public SCont As sc, CONS_ULTA As Integer
 Public ultimo As ult, ultimo1 As ult, BALANZON As Integer
 Public trcta As tra_cta
 Public trscta As tra_Scta
 Public cheque As cheques
 Public Datos As DAT_OS
 Public CATMAY As CAT_MA
 Public Rango_Inic As Long, Rango_Final As Long
 Public oper As oper_aciones, Arch_act
 Public valcelant, fin_oper As Long, Mes_Act As Integer
 Public EM As Integer, qm As Integer, Mes As Integer
 Public m_m As Integer, dia As Integer

