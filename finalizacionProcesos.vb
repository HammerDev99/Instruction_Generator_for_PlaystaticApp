Public Sub concatenaRadicado()
    
    '("A1:A552")
    Dim UltimaFila As Long
    UltimaFila = 0
    Dim concatena, radicado As String
    For n = 1 To 530 'Cantidad de filas del excel que contiene los radicados
        
'Ingresar instruccion 1 por cada radicado
        concatena = ""
        radicado = Hoja2.Cells(n, 1)
        Dim celda, caracter(23) As String
        
        'Llena vector con cada caracter del radicado
        For l = 1 To 23
            caracter(l) = Mid(radicado, l, 1)
        Next
        
        'Hoja7.Cells(1, 13) = caracter(22)
        'Hoja7.Cells(2, 13) = caracter(23)
        'Hoja7.Cells(3, 6) = caracter(3)
        'Hoja7.Cells(4, 6) = caracter(4)
        'Hoja7.Cells(5, 6) = caracter(5)
        
'Primer paquete
        'UltimaFila = Hoja7.Range("A1").End(xlDown).Offset(1, 0).Select
        UltimaFila = CLng(UltimaFila) + 1
        'UltimaFila = UltimaFila + 1
        Hoja7.Cells(UltimaFila, 1) = Hoja3.Cells(1, 12)
        'Concatena las instrucciones que estanen las columnas de
        'la fila del primer radicado que se encuentre
        For i = 2 To 11 'Filas
            concatena = ""
            For j = 1 To 6 'Columnas
                
                If j = 3 Then
                    'Las condiciones concatenan cada caracter del
                    'radicado segun instrucción
                    Select Case i - 1
                        Case 1
                            concatena = concatena & "NUMPAD" & caracter(1) & ";"
                        Case 2
                            concatena = concatena & "NUMPAD" & caracter(1) & ";"
                        Case 3
                            concatena = concatena & "NUMPAD" & caracter(2) & ";"
                        Case 4
                            concatena = concatena & "NUMPAD" & caracter(2) & ";"
                        Case 5
                            concatena = concatena & "NUMPAD" & caracter(3) & ";"
                        Case 6
                            concatena = concatena & "NUMPAD" & caracter(3) & ";"
                        Case 7
                            concatena = concatena & "NUMPAD" & caracter(4) & ";"
                        Case 8
                            concatena = concatena & "NUMPAD" & caracter(4) & ";"
                        Case 9
                            concatena = concatena & "NUMPAD" & caracter(5) & ";"
                        Case 10
                            concatena = concatena & "NUMPAD" & caracter(5) & ";"
                    End Select
                Else
                    'Concatena el resto de columnas de las 10 filas con instrucción
                    If j = 6 Then
                        concatena = concatena & Hoja3.Cells(i, j)
                    Else
                        concatena = concatena & Hoja3.Cells(i, j) & ";"
                    End If
                End If
            Next
            'Agrega fila que concatenó el radicado
            UltimaFila = CLng(UltimaFila) + 1
        'UltimaFila = UltimaFila + 1
            Hoja7.Cells(UltimaFila, 1) = concatena
        Next
        
'Segundo paquete
        UltimaFila = CLng(UltimaFila) + 1
        'UltimaFila = UltimaFila + 1
        Hoja7.Cells(UltimaFila, 1) = Hoja3.Cells(12, 12)
        'Concatena las instrucciones que estanen las columnas de
        'la fila del primer radicado que se encuentre
        For i = 13 To 16 'Filas
            concatena = ""
            For j = 1 To 6 'Columnas
                
                If j = 3 Then
                    'Las condiciones concatenan cada caracter del
                    'radicado segun instrucción
                    Select Case i - 1
                        Case 12
                            concatena = concatena & "NUMPAD" & caracter(6) & ";"
                        Case 13
                            concatena = concatena & "NUMPAD" & caracter(6) & ";"
                        Case 14
                            concatena = concatena & "NUMPAD" & caracter(7) & ";"
                        Case 15
                            concatena = concatena & "NUMPAD" & caracter(7) & ";"
                    End Select
                Else
                    'Concatena el resto de columnas de las 10 filas con instrucción
                    If j = 6 Then
                        concatena = concatena & Hoja3.Cells(i, j)
                    Else
                        concatena = concatena & Hoja3.Cells(i, j) & ";"
                    End If
                End If
            Next
            'Agrega fila que concatenó el radicado
            UltimaFila = CLng(UltimaFila) + 1
        'UltimaFila = UltimaFila + 1
            Hoja7.Cells(UltimaFila, 1) = concatena
        Next

'Tercer paquete
        UltimaFila = CLng(UltimaFila) + 1
        'UltimaFila = UltimaFila + 1
        Hoja7.Cells(UltimaFila, 1) = Hoja3.Cells(17, 12)
        'Concatena las instrucciones que estanen las columnas de
        'la fila del primer radicado que se encuentre
        For i = 18 To 21 'Filas
            concatena = ""
            For j = 1 To 6 'Columnas
                
                If j = 3 Then
                    'Las condiciones concatenan cada caracter del
                    'radicado segun instrucción
                    Select Case i - 1
                        Case 17
                            concatena = concatena & "NUMPAD" & caracter(8) & ";"
                        Case 18
                            concatena = concatena & "NUMPAD" & caracter(8) & ";"
                        Case 19
                            concatena = concatena & "NUMPAD" & caracter(9) & ";"
                        Case 20
                            concatena = concatena & "NUMPAD" & caracter(9) & ";"
                    End Select
                Else
                    'Concatena el resto de columnas de las 10 filas con instrucción
                    If j = 6 Then
                        concatena = concatena & Hoja3.Cells(i, j)
                    Else
                        concatena = concatena & Hoja3.Cells(i, j) & ";"
                    End If
                End If
            Next
            'Agrega fila que concatenó el radicado
            UltimaFila = CLng(UltimaFila) + 1
        'UltimaFila = UltimaFila + 1
            Hoja7.Cells(UltimaFila, 1) = concatena
        Next
        
'Cuarto paquete
        UltimaFila = CLng(UltimaFila) + 1
        'UltimaFila = UltimaFila + 1
        Hoja7.Cells(UltimaFila, 1) = Hoja3.Cells(22, 12)
        'Concatena las instrucciones que estanen las columnas de
        'la fila del primer radicado que se encuentre
        For i = 23 To 28 'Filas
            concatena = ""
            For j = 1 To 6 'Columnas
                
                If j = 3 Then
                    'Las condiciones concatenan cada caracter del
                    'radicado segun instrucción
                    Select Case i - 1
                        Case 22
                            concatena = concatena & "NUMPAD" & caracter(10) & ";"
                        Case 23
                            concatena = concatena & "NUMPAD" & caracter(10) & ";"
                        Case 24
                            concatena = concatena & "NUMPAD" & caracter(11) & ";"
                        Case 25
                            concatena = concatena & "NUMPAD" & caracter(11) & ";"
                        Case 26
                            concatena = concatena & "NUMPAD" & caracter(12) & ";"
                        Case 27
                            concatena = concatena & "NUMPAD" & caracter(12) & ";"
                    End Select
                Else
                    'Concatena el resto de columnas de las 10 filas con instrucción
                    If j = 6 Then
                        concatena = concatena & Hoja3.Cells(i, j)
                    Else
                        concatena = concatena & Hoja3.Cells(i, j) & ";"
                    End If
                End If
            Next
            'Agrega fila que concatenó el radicado
            UltimaFila = CLng(UltimaFila) + 1
        'UltimaFila = UltimaFila + 1
            Hoja7.Cells(UltimaFila, 1) = concatena
        Next
        
'Quinto paquete
        'UltimaFila = Hoja7.Range("A1").End(xlDown).Offset(1, 0).Select
        UltimaFila = CLng(UltimaFila) + 1
        'UltimaFila = UltimaFila + 1
        Hoja7.Cells(UltimaFila, 1) = Hoja3.Cells(29, 12)
        'Concatena las instrucciones que estanen las columnas de
        'la fila del primer radicado que se encuentre
        For i = 30 To 37 'Filas
            concatena = ""
            For j = 1 To 6 'Columnas
                
                If j = 3 Then
                    'Las condiciones concatenan cada caracter del
                    'radicado segun instrucción
                    Select Case i - 1
                        Case 29
                            concatena = concatena & "NUMPAD" & caracter(13) & ";"
                        Case 30
                            concatena = concatena & "NUMPAD" & caracter(13) & ";"
                        Case 31
                            concatena = concatena & "NUMPAD" & caracter(14) & ";"
                        Case 32
                            concatena = concatena & "NUMPAD" & caracter(14) & ";"
                        Case 33
                            concatena = concatena & "NUMPAD" & caracter(15) & ";"
                        Case 34
                            concatena = concatena & "NUMPAD" & caracter(15) & ";"
                        Case 35
                            concatena = concatena & "NUMPAD" & caracter(16) & ";"
                        Case 36
                            concatena = concatena & "NUMPAD" & caracter(16) & ";"
                    End Select
                Else
                    'Concatena el resto de columnas de las 10 filas con instrucción
                    If j = 6 Then
                        concatena = concatena & Hoja3.Cells(i, j)
                    Else
                        concatena = concatena & Hoja3.Cells(i, j) & ";"
                    End If
                End If
            Next
            'Agrega fila que concatenó el radicado
            UltimaFila = CLng(UltimaFila) + 1
        'UltimaFila = UltimaFila + 1
            Hoja7.Cells(UltimaFila, 1) = concatena
        Next
        
'Sexto paquete
        'UltimaFila = Hoja7.Range("A1").End(xlDown).Offset(1, 0).Select
        UltimaFila = CLng(UltimaFila) + 1
        'UltimaFila = UltimaFila + 1
        Hoja7.Cells(UltimaFila, 1) = Hoja3.Cells(38, 12)
        'Concatena las instrucciones que estanen las columnas de
        'la fila del primer radicado que se encuentre
        For i = 39 To 48 'Filas
            concatena = ""
            For j = 1 To 6 'Columnas
                
                If j = 3 Then
                    'Las condiciones concatenan cada caracter del
                    'radicado segun instrucción
                    Select Case i - 1
                        Case 38
                            concatena = concatena & "NUMPAD" & caracter(17) & ";"
                        Case 39
                            concatena = concatena & "NUMPAD" & caracter(17) & ";"
                        Case 40
                            concatena = concatena & "NUMPAD" & caracter(18) & ";"
                        Case 41
                            concatena = concatena & "NUMPAD" & caracter(18) & ";"
                        Case 42
                            concatena = concatena & "NUMPAD" & caracter(19) & ";"
                        Case 43
                            concatena = concatena & "NUMPAD" & caracter(19) & ";"
                        Case 44
                            concatena = concatena & "NUMPAD" & caracter(20) & ";"
                        Case 45
                            concatena = concatena & "NUMPAD" & caracter(20) & ";"
                        Case 46
                            concatena = concatena & "NUMPAD" & caracter(21) & ";"
                        Case 47
                            concatena = concatena & "NUMPAD" & caracter(21) & ";"
                    End Select
                Else
                    'Concatena el resto de columnas de las 10 filas con instrucción
                    If j = 6 Then
                        concatena = concatena & Hoja3.Cells(i, j)
                    Else
                        concatena = concatena & Hoja3.Cells(i, j) & ";"
                    End If
                End If
            Next
            'Agrega fila que concatenó el radicado
            UltimaFila = CLng(UltimaFila) + 1
        'UltimaFila = UltimaFila + 1
            Hoja7.Cells(UltimaFila, 1) = concatena
        Next
        
'Septimo paquete
        UltimaFila = CLng(UltimaFila) + 1
        'UltimaFila = UltimaFila + 1
        Hoja7.Cells(UltimaFila, 1) = Hoja3.Cells(49, 12)
        'Concatena las instrucciones que estanen las columnas de
        'la fila del primer radicado que se encuentre
        For i = 50 To 53 'Filas
            concatena = ""
            For j = 1 To 6 'Columnas
                
                If j = 3 Then
                    'Las condiciones concatenan cada caracter del
                    'radicado segun instrucción
                    Select Case i - 1
                        Case 49
                            concatena = concatena & "NUMPAD" & caracter(22) & ";"
                        Case 50
                            concatena = concatena & "NUMPAD" & caracter(22) & ";"
                        Case 51
                            concatena = concatena & "NUMPAD" & caracter(23) & ";"
                        Case 52
                            concatena = concatena & "NUMPAD" & caracter(23) & ";"
                    End Select
                Else
                    'Concatena el resto de columnas de las 10 filas con instrucción
                    If j = 6 Then
                        concatena = concatena & Hoja3.Cells(i, j)
                    Else
                        concatena = concatena & Hoja3.Cells(i, j) & ";"
                    End If
                End If
            Next
            'Agrega fila que concatenó el radicado
            UltimaFila = CLng(UltimaFila) + 1
        'UltimaFila = UltimaFila + 1
            Hoja7.Cells(UltimaFila, 1) = concatena
        Next
        
'Agregar las instrucciones siguientes
        For k = 54 To 73
            'UltimaFila = Hoja7.Range("A1").End(xlDown).Offset(1, 0).Select
            UltimaFila = CLng(UltimaFila) + 1
        'UltimaFila = UltimaFila + 1
            Hoja7.Cells(UltimaFila, 1) = Hoja3.Cells(k, 12)
        Next
    Next
    Hoja7.Select
    MsgBox ("Ejecución finalizada")
End Sub
