Attribute VB_Name = "Moneda"

Public Function NumLetra(Número As Double, Optional NumDecimales As Integer, _
                    Optional Unidad As String, Optional UdFracc As String, _
                    Optional Conexión As String, Optional Cero As Boolean, _
                    Optional UD_un_uno_a As Integer, Optional Frac_un_uno_a As Integer, _
                    Optional UnMil As Boolean) As String
' función para convertir números a letra.
' se emplea un número tipo double, 8 bytes
' (Máximo NÚMEROS DE 15 DÍGITOS, a partir de ahí son números en coma flotante,
' y deberían introducirse como texto para poder mantener todas las cifras)

' Número       el número a convertir, OBLIGATORIO
' NumDecimales  número de decimales a considerar para pasar a texto (por defecto: cero)
'           Si el número de decimales es negativo, se redondea el número según lo indicado
' Unidad    nombre de la unidad principal, se pondrá detrás de la parte entera
' UdFracc   nombre de la unidad fraccionaria, se pondrá detrás de los decimales
' Conexión  texto que separará la parte entera de la decimal(por ejemplo, tres euros "CON" quince céntimos)
' Cero      verdadero 0->"cero"   falso 0->"" (por defecto: falso)
' UD_un_uno_a  para la unidad principal     1: 1->"un"  2: 1->"uno" 3: 1->"una" (por defecto: 1)
' Fracc_un_uno_a  para la unidad principal  1: 1->"un"  2: 1->"uno" 3: 1->"una" (por defecto: 1)
' UnMil   para que aparezca "un mil" en casos en que hay una unidad de millar (por defecto: falso)
    
    '''Valores por defecto:
    If NumDecimales = 0 Then NumDecimales = 0
    If Unidad = "" Then Unidad = ""
    If UdFracc = "" Then UdFracc = ""
    If Conexión = "" Then Conexión = ""
    'If IsMissing(Cero) Then Cero = False
    If UD_un_uno_a = 0 Then UD_un_uno_a = 1
    If Frac_un_uno_a = 0 Then Frac_un_uno_a = 1
    
    Dim intSigno As Integer     ' signo del número
    Dim dblNúmero As Double     ' número a transformar, en positivo

    Dim intParteEntera As Double
    Dim intParteDecimal As Double
    
    
    dblNúmero = Abs(Número)
    intSigno = Sgn(Número)
    
    
    If NumDecimales >= 0 Then
      Número = Round(Número, NumDecimales)
      intParteEntera = Int(Abs(Número))
      intParteDecimal = Int(Round((Abs(Número) - intParteEntera) * 10 ^ NumDecimales, 0))
    ElseIf NumDecimales < 0 Then
      intParteDecimal = 0
      intParteEntera = Round((Abs(Número) / 10 ^ Abs(NumDecimales)), 0) * 10 ^ Abs(NumDecimales)
    End If
    If (intParteEntera = 0 And Not Cero) Or (intParteDecimal = 0 And Not Cero) Then Conexión = "" ' si no hay parte entera o decimal no ponemos el texto intermedio
    NumLetra = Trim(Entero_Letra(intParteEntera, Unidad, Cero, UD_un_uno_a, UnMil) & " " & _
                     Conexión & " con  " & formato(intParteDecimal) & "/100 ")
                    
    ' cuando el número sea negativo
    If intSigno = -1 Then
        NumLetra = "menos " & NumLetra
    End If
    NumLetra = UCase(NumLetra)
End Function
Function formato(v)
    If v < 10 Then
        formato = "0" & v
    Else
        formato = v
    End If
End Function

Private Function Entero_Letra(Número As Double, Unidad As String, _
    Cero As Boolean, un_o_a As Integer, UnMil As Boolean) As String
' función para convertir en texto números enteros
' lo pongo como variable double en vez de entero largo, por si es un número muy grande
    Const intMaxGrupo As Integer = 3    ' grupos de 6 cifras considerados. con 3 llegamos hasta Trillones.
    Dim lngAuxNum As Long    ' número auxiliar para convertir el número por partes, hasta seis cifras
    Dim IntUnidades As Integer, IntMillares As Integer  ' millares y unidades de lngAuxNum
    Dim dblAuxResto As Double  ' número auxiliar, la parte todavía por convertir
    Dim strAuxUnidad As String
    Dim Resultado As String
    Dim i As Integer
 
    Resultado = ""
    
    'si tenemos un cero
    If Número = 0 Then
        Select Case Cero
            Case True
                Resultado = "cero " & Unidad
            Case False
                Resultado = ""
        End Select
        Entero_Letra = Trim(Resultado)
        Exit Function
    End If
    
    dblAuxResto = Abs(Número)
    
    For i = 0 To intMaxGrupo
        ' tomamos el número de 6 en 6 cifras
        lngAuxNum = Round(((dblAuxResto / 10 ^ 6 - Int(dblAuxResto / 10 ^ 6)) * 10 ^ 6), 0)
        IntUnidades = Round((lngAuxNum / 10 ^ 3 - Int(lngAuxNum / 10 ^ 3)) * 10 ^ 3, 0)
        IntMillares = (lngAuxNum - IntUnidades) / 10 ^ 3
        If lngAuxNum <> 0 Or i = 0 Then
            Select Case i
                Case 0  'unidades
                    'unidades
                    strAuxUnidad = RTrim(" " & Unidad)
                    Resultado = cifra_3(IntUnidades, False, un_o_a) & strAuxUnidad
                    ' millares
                    If IntMillares <> 0 Then
                      strAuxUnidad = " mil"
                      If IntMillares = 1 Then
                          Resultado = Trim(cifra_3(IntMillares, False, IIf(UnMil, 1, 0)) & strAuxUnidad & " " & Resultado)
                      Else
                          Resultado = Trim(cifra_3(IntMillares, False, un_o_a) & strAuxUnidad & " " & Resultado)
                      End If
                    End If

                    
                Case 1 ' millones
                    ' millones
                    If lngAuxNum = 1 Then
                        strAuxUnidad = " millón"
                    Else
                        strAuxUnidad = " millones"
                    End If
                    Resultado = Trim(cifra_3(IntUnidades, False, 1) & strAuxUnidad & " " & Resultado)
                
                    ' mil millones
                    If IntMillares <> 0 Then
                      strAuxUnidad = " mil"
                      If IntMillares = 1 Then
                          Resultado = Trim(cifra_3(IntMillares, False, 0) & strAuxUnidad & " " & Resultado)
                      Else
                          Resultado = Trim(cifra_3(IntMillares, False, 1) & strAuxUnidad & " " & Resultado)
                      End If
                    End If
                    
                    
                Case 2 ' billones
                    ' billones
                    If lngAuxNum = 1 Then
                        strAuxUnidad = " billón"
                    Else
                        strAuxUnidad = " billones"
                    End If
                    Resultado = Trim(cifra_3(IntUnidades, False, 1) & strAuxUnidad & " " & Resultado)
                    ' mil billones
                    If IntMillares <> 0 Then
                      strAuxUnidad = " mil"
                      If IntMillares = 1 Then
                          Resultado = Trim(cifra_3(IntMillares, False, 0) & strAuxUnidad & " " & Resultado)
                      Else
                          Resultado = Trim(cifra_3(IntMillares, False, 1) & strAuxUnidad & " " & Resultado)
                      End If
                    End If
                    
                Case 3 ' Trillones
                    ' Trillones
                    If lngAuxNum = 1 Then
                        strAuxUnidad = " trillón"
                    Else
                        strAuxUnidad = " trillones"
                    End If
                    Resultado = Trim(cifra_3(IntUnidades, False, 1) & strAuxUnidad & " " & Resultado)
                    ' mil trillones
                    If IntMillares <> 0 Then
                      strAuxUnidad = " mil"
                      If IntMillares = 1 Then
                          Resultado = Trim(cifra_3(IntMillares, False, 0) & strAuxUnidad & " " & Resultado)
                      Else
                          Resultado = Trim(cifra_3(IntMillares, False, 1) & strAuxUnidad & " " & Resultado)
                      End If
                    End If
                    
            End Select
        End If
        dblAuxResto = Int(dblAuxResto / 10 ^ 6) 'para el siguiente ciclo
        If dblAuxResto = 0 Then Exit For    ' si ya hemos acabado
    Next

    Entero_Letra = Trim(Replace(Resultado, "  ", " "))

End Function

Private Function cifra_1(num As Integer, Cero As Boolean, un_o_a As Integer) As String
' función para convertir en texto números de una cifra
' num       el númeroa convertir
' cero      verdadero 0->"cero"   falso 0->""
' un_o_a    0: 1->""    1: 1->"un"  2: 1->"uno" 3: 1->"una"
    Select Case num
        Case 0
            If Cero Then
                cifra_1 = "cero"
            Else
                cifra_1 = ""
            End If
        Case 1
            Select Case un_o_a
                Case 0
                    cifra_1 = ""
                Case 1
                    cifra_1 = "un"
                Case 2
                    cifra_1 = "uno"
                Case 3
                    cifra_1 = "una"
            End Select
        Case 2
            cifra_1 = "dos"
        Case 3
            cifra_1 = "tres"
        Case 4
            cifra_1 = "cuatro"
        Case 5
            cifra_1 = "cinco"
        Case 6
            cifra_1 = "seis"
        Case 7
            cifra_1 = "siete"
        Case 8
            cifra_1 = "ocho"
        Case 9
            cifra_1 = "nueve"
    End Select
End Function

Private Function cifra_2(num As Integer, Cero As Boolean, un_o_a As Integer) As String
' función para convertir en texto números de una cifra
' num       el númeroa convertir
' cero      verdadero 0->"cero"   falso 0->""
' un_o_a    0: 1->""    1: 1->"un"  2: 1->"uno" 3: 1->"una"
    Select Case num
        Case Is < 10
            cifra_2 = cifra_1(num, Cero, un_o_a)
        Case 10
            cifra_2 = "diez"
        Case 11
            cifra_2 = "once"
        Case 12
            cifra_2 = "doce"
        Case 13
            cifra_2 = "trece"
        Case 14
            cifra_2 = "catorce"
        Case 15
            cifra_2 = "quince"
        Case 16 To 19
            cifra_2 = "dieci" & cifra_1(num - 10, False, un_o_a)
        Case 20
            cifra_2 = "veinte"
        Case 21 To 29
            cifra_2 = "veinti" & cifra_1(num - 20, False, un_o_a)
        Case 30
            cifra_2 = "treinta"
        Case 31 To 39
            cifra_2 = "treinta y " & cifra_1(num - 30, False, un_o_a)
        Case 40
            cifra_2 = "cuarenta"
        Case 41 To 49
            cifra_2 = "cuarenta y " & cifra_1(num - 40, False, un_o_a)
        Case 50
            cifra_2 = "cincuenta"
        Case 51 To 59
            cifra_2 = "cincuenta y " & cifra_1(num - 50, False, un_o_a)
        Case 60
            cifra_2 = "sesenta"
        Case 61 To 69
            cifra_2 = "sesenta y " & cifra_1(num - 60, False, un_o_a)
        Case 70
            cifra_2 = "setenta"
        Case 71 To 79
            cifra_2 = "setenta y " & cifra_1(num - 70, False, un_o_a)
        Case 80
            cifra_2 = "ochenta"
        Case 81 To 89
            cifra_2 = "ochenta y " & cifra_1(num - 80, False, un_o_a)
        Case 90
            cifra_2 = "noventa"
        Case 91 To 99
            cifra_2 = "noventa y " & cifra_1(num - 90, False, un_o_a)
    End Select
End Function

Private Function cifra_3(num As Integer, Cero As Boolean, un_o_a As Integer) As String
' función para convertir en texto números de una cifra
' num       el númeroa convertir
' cero      verdadero 0->"cero"   falso 0->""
' un_o_a    0: 1->""    1: 1->"un"  2: 1->"uno" 3: 1->"una"
' se realizan llamadas a la funcion cifra_2 para valores inferiores a 100
    
    Select Case num
        Case Is < 100
            cifra_3 = cifra_2(num, Cero, un_o_a)
        Case 100
            cifra_3 = "cien"
        Case 101 To 199
            cifra_3 = "ciento " & cifra_2(num - 100, False, un_o_a)
        Case 200 To 299
            If un_o_a = 3 Then
                cifra_3 = "doscientas " & cifra_2(num - 200, False, un_o_a)
            Else
                cifra_3 = "doscientos " & cifra_2(num - 200, False, un_o_a)
            End If
        Case 300 To 399
            If un_o_a = 3 Then
                cifra_3 = "trescientas " & cifra_2(num - 300, False, un_o_a)
            Else
                cifra_3 = "trescientos " & cifra_2(num - 300, False, un_o_a)
            End If
        Case 400 To 499
            If un_o_a = 3 Then
                cifra_3 = "cuatrocientas " & cifra_2(num - 400, False, un_o_a)
            Else
                cifra_3 = "cuatrocientos " & cifra_2(num - 400, False, un_o_a)
            End If
        Case 500 To 599
            If un_o_a = 3 Then
                cifra_3 = "quinientas " & cifra_2(num - 500, False, un_o_a)
            Else
                cifra_3 = "quinientos " & cifra_2(num - 500, False, un_o_a)
            End If
        Case 600 To 699
            If un_o_a = 3 Then
                cifra_3 = "seiscientas " & cifra_2(num - 600, False, un_o_a)
            Else
                cifra_3 = "seiscientos " & cifra_2(num - 600, False, un_o_a)
            End If
        Case 700 To 799
            If un_o_a = 3 Then
                cifra_3 = "setecientas " & cifra_2(num - 700, False, un_o_a)
            Else
                cifra_3 = "setecientos " & cifra_2(num - 700, False, un_o_a)
            End If
        Case 800 To 899
            If un_o_a = 3 Then
                cifra_3 = "ochocientas " & cifra_2(num - 800, False, un_o_a)
            Else
                cifra_3 = "ochocientos " & cifra_2(num - 800, False, un_o_a)
            End If
        Case 900 To 999
            If un_o_a = 3 Then
                cifra_3 = "novecientas " & cifra_2(num - 900, False, un_o_a)
            Else
                cifra_3 = "novecientos " & cifra_2(num - 900, False, un_o_a)
            End If
    End Select
End Function



