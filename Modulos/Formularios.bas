Attribute VB_Name = "Formularios"
Sub cargarStatusPago(combo As ComboBox)
    combo.Clear
    combo.AddItem "Abierta"
    combo.AddItem "Debe"
    combo.AddItem "Cerrada"
    combo.AddItem "Anulado"
End Sub
Sub cargarStatusTrabajo(combo As ComboBox)
    combo.Clear
    combo.AddItem "Pendiente"
    combo.AddItem "En Ejecucion"
    combo.AddItem "Terminado"
    combo.AddItem "Entregado"
End Sub

Public Sub modificar(formulario As Form)
    formulario.bloquear True
    formulario.Text2.SetFocus
    Botones formulario, 2

End Sub
Public Sub Nuevo(formulario As Form)
    Botones formulario, 2
    formulario.Limpiar
    formulario.bloquear True
    formulario.Text2.SetFocus
    formulario.Text1.Text = Datos.generarCodigo(formulario.Tabla, formulario.campoClave)
End Sub
Public Sub Guardar(formulario As Form)
    Conexion.Execute formulario.SqlActualizacion
    Botones formulario, 1
    formulario.bloquear False
End Sub
Public Sub Primero(formulario As Form)
    Dim n As Integer
    Dim Maximo As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "select max([" & formulario.campoClave & "])from [" & formulario.Tabla & "]", Conexion
    If IsNumeric(rs(0)) Then
       Maximo = rs(0)
    Else
        Maximo = 0
    End If
    n = 0
    rs.Close
    
    
    

2:
    n = n + 1
    
    rs.Open "select *from [" & formulario.Tabla & "] where [" & formulario.campoClave & "]=" & n, Conexion
    If rs.EOF Then
        If n < Maximo Then
            rs.Close
            GoTo 2
        End If
        MsgBox "El Registro no Existe"
    Else
        On Error Resume Next
        Botones formulario, 3
        formulario.bloquear False
        formulario.Mostrar rs
    
    
    End If
End Sub
Public Sub Siguiente(formulario As Form)
    Dim n As Integer
    Dim rs As New ADODB.Recordset
    
    rs.Open "select max([" & formulario.campoClave & "])from [" & formulario.Tabla & "]", Conexion
    If IsNumeric(rs(0)) Then
       Maximo = rs(0)
    Else
        Maximo = 0
    End If
    rs.Close
    
    
    n = Val(formulario.Text1.Text)
2:
    n = n + 1
    rs.Open "select *from [" & formulario.Tabla & "] where [" & formulario.campoClave & "]=" & n, Conexion
    If rs.EOF Then
    If n < Maximo Then
        rs.Close
        GoTo 2
    End If
              MsgBox "El Registro no Existe"
    Else
        On Error Resume Next
        Botones formulario, 3
        formulario.Mostrar rs
        formulario.bloquear False
    
    End If
End Sub
Public Sub Anterior(formulario As Form)
    Dim n As Integer
    Dim rs As New ADODB.Recordset
    n = Val(formulario.Text1.Text)
2:
    n = n - 1
    rs.Open "select *from [" & formulario.Tabla & "] where [" & formulario.campoClave & "]=" & n, Conexion
    If rs.EOF Then
        If n > 1 Then
            rs.Close
            GoTo 2
        End If
            
        MsgBox "El Registro no Existe"
    Else
        On Error Resume Next
        Botones formulario, 3
        formulario.Mostrar rs
        formulario.bloquear False
    
    End If
End Sub
Public Sub Ultimo(formulario As Form)
    Dim n As Integer
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    rs2.Open "select max([" & formulario.campoClave & "]) from [" & formulario.Tabla & "]", Conexion
    If IsNumeric(rs2(0)) Then
        n = rs2(0)
        rs.Open "select *from [" & formulario.Tabla & "] where [" & formulario.campoClave & "]=" & n, Conexion
        If rs.EOF Then
            MsgBox "El Registro no Existe"
        Else
        On Error Resume Next
            Botones formulario, 3
            formulario.Mostrar rs
            
            formulario.bloquear False
        End If
    Else
            MsgBox "El Registro no Existe"
    End If

End Sub


Public Sub cancelar(formulario As Form)
    Botones formulario, 1
    formulario.Limpiar
    formulario.bloquear False
End Sub
Public Sub Eliminar(formulario As Form)
    If MsgBox("Esta Seguro de que Desea Eliminar Este Registro ", vbYesNo) = vbYes Then
        Conexion.Execute "delete from [" & formulario.Tabla & "] where [" & formulario.campoClave & "]=" & Val(formulario.Text1.Text)
        cancelar formulario
    End If
End Sub
Sub Buscar(formulario As Form)
    Dim rs As New ADODB.Recordset
    rs.Open "Select *from [" & formulario.Tabla & "] where [" & formulario.campoClave & "] =" & Val(formulario.Text1.Text), Conexion
    If rs.EOF Then
        formulario.Limpiar
        Botones formulario, 1
    Else
        formulario.Mostrar rs
        Botones formulario, 3
        formulario.bloquear False
        'formulario.Text1.Enabled = False
    End If

End Sub

Sub Botones(formulario As Form, es As Byte)
    If es = 1 Then
        formulario.Command1.Enabled = True 'Nuevo
        formulario.Command2.Enabled = False 'Guardar
        formulario.Command3.Enabled = False 'Cancelar
        formulario.Command4.Enabled = False 'Modificar
        formulario.Command5.Enabled = False 'Eliminar
    End If
    If es = 2 Then
        formulario.Command1.Enabled = True 'Nuevo
        formulario.Command2.Enabled = True 'Guardar
        formulario.Command3.Enabled = True 'Cancelar
        formulario.Command4.Enabled = False 'Modificar
        formulario.Command5.Enabled = False 'Eliminar
    End If
    If es = 3 Then
        formulario.Command1.Enabled = False 'Nuevo
        formulario.Command2.Enabled = False 'Guardar
        formulario.Command3.Enabled = True 'Cancelar
        formulario.Command4.Enabled = True 'Modificar
        formulario.Command5.Enabled = True 'Eliminar
    End If
End Sub
Sub ColoresTextos(texto As Variant, Fondo As Variant, f As Form)
   On Error Resume Next
    f.Text1.ForeColor = texto
    f.Text2.ForeColor = texto
    f.Text3.ForeColor = texto
    f.Text4.ForeColor = texto
    f.Text5.ForeColor = texto
    f.Text6.ForeColor = texto
    f.Text7.ForeColor = texto
    f.Text8.ForeColor = texto
    f.Text9.ForeColor = texto
    f.Text10.ForeColor = texto
    f.Text11.ForeColor = texto
    f.Text12.ForeColor = texto
    f.Text13.ForeColor = texto
    f.Text14.ForeColor = texto
    f.Text15.ForeColor = texto
    f.Text16.ForeColor = texto
    f.Text17.ForeColor = texto
    f.Text18.ForeColor = texto
    f.Text19.ForeColor = texto
    f.Text20.ForeColor = texto
    f.Text21.ForeColor = texto
    f.Text22.ForeColor = texto
    f.Text23.ForeColor = texto
    f.Text24.ForeColor = texto
    f.Text25.ForeColor = texto
    f.Text26.ForeColor = texto
    f.Text27.ForeColor = texto
    f.Text28.ForeColor = texto
    f.Text29.ForeColor = texto
    f.Text30.ForeColor = texto
    f.Text31.ForeColor = texto
    f.Text32.ForeColor = texto
    f.Text33.ForeColor = texto
    f.Text32.ForeColor = texto
    
    f.Text1.BackColor = Fondo
    f.Text2.BackColor = Fondo
    f.Text3.BackColor = Fondo
    f.Text4.BackColor = Fondo
    f.Text5.BackColor = Fondo
    f.Text6.BackColor = Fondo
    f.Text7.BackColor = Fondo
    f.Text8.BackColor = Fondo
    f.Text9.BackColor = Fondo
    f.Text10.BackColor = Fondo
    f.Text11.BackColor = Fondo
    f.Text12.BackColor = Fondo
    f.Text13.BackColor = Fondo
    f.Text14.BackColor = Fondo
    f.Text15.BackColor = Fondo
    f.Text16.BackColor = Fondo
    f.Text17.BackColor = Fondo
    f.Text18.BackColor = Fondo
    f.Text19.BackColor = Fondo
    f.Text20.BackColor = Fondo
    f.Text21.BackColor = Fondo
    f.Text22.BackColor = Fondo
    f.Text23.BackColor = Fondo
    f.Text24.BackColor = Fondo
    f.Text25.BackColor = Fondo
    f.Text26.BackColor = Fondo
    f.Text27.BackColor = Fondo
    f.Text28.BackColor = Fondo
    f.Text29.BackColor = Fondo
    f.Text30.BackColor = Fondo
    f.Text31.BackColor = Fondo
    f.Text32.BackColor = Fondo
    f.Text33.BackColor = Fondo
    f.Text32.BackColor = Fondo





    f.Combo1.ForeColor = texto
    f.Combo2.ForeColor = texto
    f.Combo3.ForeColor = texto
    f.Combo4.ForeColor = texto
    f.Combo5.ForeColor = texto
    f.Combo6.ForeColor = texto
    f.Combo7.ForeColor = texto
    f.Combo8.ForeColor = texto
    f.Combo9.ForeColor = texto
    f.Combo10.ForeColor = texto
    f.Combo11.ForeColor = texto
    f.Combo12.ForeColor = texto
    f.Combo13.ForeColor = texto
    f.Combo14.ForeColor = texto
    f.Combo15.ForeColor = texto
    f.Combo16.ForeColor = texto
    f.Combo17.ForeColor = texto
    f.Combo18.ForeColor = texto
    f.Combo19.ForeColor = texto
    f.Combo20.ForeColor = texto
    f.Combo21.ForeColor = texto
    f.Combo22.ForeColor = texto
    f.Combo23.ForeColor = texto
    f.Combo24.ForeColor = texto
    f.Combo25.ForeColor = texto
    f.Combo26.ForeColor = texto
    f.Combo27.ForeColor = texto
    f.Combo28.ForeColor = texto
    f.Combo29.ForeColor = texto
    f.Combo30.ForeColor = texto
    f.Combo31.ForeColor = texto
    f.Combo32.ForeColor = texto
    f.Combo33.ForeColor = texto
    f.Combo32.ForeColor = texto
    
    f.Combo1.BackColor = Fondo
    f.Combo2.BackColor = Fondo
    f.Combo3.BackColor = Fondo
    f.Combo4.BackColor = Fondo
    f.Combo5.BackColor = Fondo
    f.Combo6.BackColor = Fondo
    f.Combo7.BackColor = Fondo
    f.Combo8.BackColor = Fondo
    f.Combo9.BackColor = Fondo
    f.Combo10.BackColor = Fondo
    f.Combo11.BackColor = Fondo
    f.Combo12.BackColor = Fondo
    f.Combo13.BackColor = Fondo
    f.Combo14.BackColor = Fondo
    f.Combo15.BackColor = Fondo
    f.Combo16.BackColor = Fondo
    f.Combo17.BackColor = Fondo
    f.Combo18.BackColor = Fondo
    f.Combo19.BackColor = Fondo
    f.Combo20.BackColor = Fondo
    f.Combo21.BackColor = Fondo
    f.Combo22.BackColor = Fondo
    f.Combo23.BackColor = Fondo
    f.Combo24.BackColor = Fondo
    f.Combo25.BackColor = Fondo
    f.Combo26.BackColor = Fondo
    f.Combo27.BackColor = Fondo
    f.Combo28.BackColor = Fondo
    f.Combo29.BackColor = Fondo
    f.Combo30.BackColor = Fondo
    f.Combo31.BackColor = Fondo
    f.Combo32.BackColor = Fondo
    f.Combo33.BackColor = Fondo
    f.Combo32.BackColor = Fondo
    
    f.List1.ForeColor = texto
    f.List2.ForeColor = texto
    f.List3.ForeColor = texto
    f.List4.ForeColor = texto
    f.List5.ForeColor = texto
    f.List6.ForeColor = texto
    f.List7.ForeColor = texto
    f.List8.ForeColor = texto
    f.List9.ForeColor = texto
    f.List10.ForeColor = texto



    f.List1.BackColor = Fondo
    f.List2.BackColor = Fondo
    f.List3.BackColor = Fondo
    f.List4.BackColor = Fondo
    f.List5.BackColor = Fondo
    f.List6.BackColor = Fondo
    f.List7.BackColor = Fondo
    f.List8.BackColor = Fondo
    f.List9.BackColor = Fondo
    f.List10.BackColor = Fondo


    f.DTPicker1.BackColor = Fondo
    f.DTPicker2.BackColor = Fondo
    f.DTPicker3.BackColor = Fondo
    f.DTPicker4.BackColor = Fondo
    
    f.DTPicker1.ForeColor = texto
    f.DTPicker2.ForeColor = texto
    f.DTPicker3.ForeColor = texto
    f.DTPicker4.ForeColor = texto
    

           
End Sub
Sub ColorLabels(c As ColorConstants, f As Form)
On Error Resume Next
    f.BackColor = Proyecto.ColorFondo
    ColoresTextos Proyecto.ColorTextoLetras, Proyecto.ColorFondoLetras, f
    f.Label1.ForeColor = c
    f.Label2.ForeColor = c
    f.Label3.ForeColor = c
    f.Label4.ForeColor = c
    f.Label5.ForeColor = c
    f.Label6.ForeColor = c
    f.Label7.ForeColor = c
    f.Label8.ForeColor = c
    f.Label9.ForeColor = c
    f.Label10.ForeColor = c
    f.Label11.ForeColor = c
    f.Label12.ForeColor = c
    f.Label13.ForeColor = c
    f.Label14.ForeColor = c
    f.Label15.ForeColor = c
    f.Label16.ForeColor = c
    f.Label17.ForeColor = c
    f.Label18.ForeColor = c
    f.Label19.ForeColor = c
    f.Label20.ForeColor = c
    f.Label21.ForeColor = c
    f.Label22.ForeColor = c
    f.Label23.ForeColor = c
    f.Label24.ForeColor = c
    f.Label25.ForeColor = c
    f.Label26.ForeColor = c
    f.Label27.ForeColor = c
    f.Label28.ForeColor = c
    f.Label29.ForeColor = c
    f.Label30.ForeColor = c
    f.Label31.ForeColor = c
    f.Label32.ForeColor = c
    f.Label33.ForeColor = c
    f.Label34.ForeColor = c
    f.Label35.ForeColor = c
    f.Label36.ForeColor = c
    f.Label37.ForeColor = c
    f.Label38.ForeColor = c
    f.Label39.ForeColor = c
End Sub
Function sinIva(Monto As Double)
sinIva = Val(Monto) / (1 + PIVa)
End Function
Function conIva(Monto As Double)
Dim Montos As Double
    Montos = (Val(Monto) * PIVa) + Val(Monto)
    conIva = Round(Montos, 2)
End Function

Public Function Num2Text(ByVal value As Double) As String
    Select Case value
        Case 0: Num2Text = "CERO"
        Case 1: Num2Text = "UN"
        Case 2: Num2Text = "DOS"
        Case 3: Num2Text = "TRES"
        Case 4: Num2Text = "CUATRO"
        Case 5: Num2Text = "CINCO"
        Case 6: Num2Text = "SEIS"
        Case 7: Num2Text = "SIETE"
        Case 8: Num2Text = "OCHO"
        Case 9: Num2Text = "NUEVE"
        Case 10: Num2Text = "DIEZ"
        Case 11: Num2Text = "ONCE"
        Case 12: Num2Text = "DOCE"
        Case 13: Num2Text = "TRECE"
        Case 14: Num2Text = "CATORCE"
        Case 15: Num2Text = "QUINCE"
        Case Is < 20: Num2Text = "DIECI" & Num2Text(value - 10)
        Case 20: Num2Text = "VEINTE"
        Case Is < 30: Num2Text = "VEINTI" & Num2Text(value - 20)
        Case 30: Num2Text = "TREINTA"
        Case 40: Num2Text = "CUARENTA"
        Case 50: Num2Text = "CINCUENTA"
        Case 60: Num2Text = "SESENTA"
        Case 70: Num2Text = "SETENTA"
        Case 80: Num2Text = "OCHENTA"
        Case 90: Num2Text = "NOVENTA"
        Case Is < 100: Num2Text = Num2Text(Int(value \ 10) * 10) & " Y " & Num2Text(value Mod 10)
        Case 100: Num2Text = "CIEN"
        Case Is < 200: Num2Text = "CIENTO " & Num2Text(value - 100)
        Case 200, 300, 400, 600, 800: Num2Text = Num2Text(Int(value \ 100)) & "CIENTOS"
        Case 500: Num2Text = "QUINIENTOS"
        Case 700: Num2Text = "SETECIENTOS"
        Case 900: Num2Text = "NOVECIENTOS"
        Case Is < 1000: Num2Text = Num2Text(Int(value \ 100) * 100) & " " & Num2Text(value Mod 100)
        Case 1000: Num2Text = "MIL"
        Case Is < 2000: Num2Text = "MIL " & Num2Text(value Mod 1000)
        Case Is < 1000000: Num2Text = Num2Text(Int(value \ 1000)) & " MIL"
            If value Mod 1000 Then Num2Text = Num2Text & " " & Num2Text(value Mod 1000)
        Case 1000000: Num2Text = "UN MILLON"
        Case Is < 2000000: Num2Text = "UN MILLON " & Num2Text(value Mod 1000000)
        Case Is < 1000000000000#: Num2Text = Num2Text(Int(value / 1000000)) & " MILLONES "
            If (value - Int(value / 1000000) * 1000000) Then Num2Text = Num2Text & " " & Num2Text(value - Int(value / 1000000) * 1000000)
        Case 1000000000000#: Num2Text = "UN BILLON"
        Case Is < 2000000000000#: Num2Text = "UN BILLON " & Num2Text(value - Int(value / 1000000000000#) * 1000000000000#)
        Case Else: Num2Text = Num2Text(Int(value / 1000000000000#)) & " BILLONES"
            If (value - Int(value / 1000000000000#) * 1000000000000#) Then Num2Text = Num2Text & " " & Num2Text(value - Int(value / 1000000000000#) * 1000000000000#)
    End Select


End Function

