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
  
  Type empre
       name As String * 60
       ao As Integer
       sm As Currency
       psub As Currency
       fecha As String * 14
  End Type
 Public empresa As empre
 Public personal As per, neto As Currency
 Public rgtro As Integer
 Public z1$, mm(12) As String * 20, cm, Dm, z2$
 Public subdirectorio$, valor$, dd(12) As Integer
 Public arch_tr As String * 20, tar As Integer

