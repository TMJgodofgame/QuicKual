Private Sub QuicKual(control As IRibbonControl)
    Dim selectedLanguage
    selectedLanguage = InputBox("Select the language you want to use:" & vbNewLine & "1) Castellano" & vbNewLine & "2) English" & vbNewLine & "3) Gaulois")
    
    Select Case selectedLanguage
        Case "1"
            Dim curso
            curso = InputBox("En que curso estas:" & vbNewLine & "1) 1 y 2 de primaria" & vbNewLine & "2) 3 y 4 de primaria" & vbNewLine & "3) A partir de 5 primaria")
    
            Select Case curso
                Case "1"
                    Call Castellano_1_y_2
                Case "2"
                    Call Castellano_3_y_4
                Case "3"
                    Call Castellano_5
                Case Else
                    MsgBox "Curso no valido"
            End Select
    
        Case "2"
            Dim course
            course = InputBox("Select your school academic:" & vbNewLine & "1) 1st and 2nd grade" & vbNewLine & "2) 3rd and 4th grade" & vbNewLine & "3) 5th grade and above")
            
            Select Case course
                Case "1"
                    Call English_1st_2nd
                Case "2"
                   Call English_3rd_4th
                Case "3"
                    Call English_from_5th_grade
                Case Else
                MsgBox "Invalid course"
            End Select
        Case "3"
            Dim cours
            cours = InputBox("Dans quelle classe es-tu ?" & vbNewLine & "1) CP et CE1" & vbNewLine & "2) CE2 et CM1" & vbNewLine & "3) � partir du CM2")

            Select Case cours
                Case "1"
                    Call Gaulois_CP_CE1
                Case "2"
                    Call Gaulois_CE2_CM1
                Case "3"
                    Call Gaulois_CM2_et_plus
                Case Else
                    MsgBox "Classe non valide"
            End Select

        Case Else
            MsgBox "language not valid"
    End Select
End Sub

Function Castellano_1_y_2()
    Dim opcion As String
    opcion = InputBox("Seleccione la operacion que desea realizar:" & vbCrLf & "1. Suma" & vbCrLf & "2. Resta" & vbCrLf & "3. Multiplicacion" & vbCrLf & "4. Division Larga" & vbCrLf & "5. Descomposicion factorial", "Menu de Operaciones")
    
    Select Case opcion
        Case "1"
            Call Realizar_Suma
        Case "2"
            Call Realizar_Resta
        Case "3"
            Call Realizar_Multiplicacion
        Case "4"
            Call Hacer_divisiones_desarrolladas
        Case "5"
            Call Hacer_descomposicion_factorial
        Case Else
            MsgBox "Opcion no valida. Por favor, seleccione una operacion valida."
    End Select
End Function

Function Realizar_Suma()
    Dim PrimerSumando As String
    Dim SegundoSumando As String
    
    ' Solicita el primer sumando
    PrimerSumando = InputBox("Ingrese el primer sumando:", "Suma")
    
    ' Verifica si el usuario ingreso un numero valido
    If IsNumeric(PrimerSumando) Then
        ' Solicita el segundo sumando
        SegundoSumando = InputBox("Ingrese el segundo sumando:", "Suma")
        
        ' Verifica si el usuario ingreso un numero valido
        If IsNumeric(SegundoSumando) Then
            ' Crear la tabla en Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)
            
            ' Ocultar los bordes de las celdas
            tbl.Borders.Enable = True
            
            ' Llena la tabla desde la ultima celda hacia atras
            FillTableCells tbl.cell(2, tbl.Columns.Count), PrimerSumando
            FillTableCells tbl.cell(3, tbl.Columns.Count - 14), "+" ' Coloca el s�mbolo '+' aqu�
            FillTableCells tbl.cell(3, tbl.Columns.Count), SegundoSumando
            
            ' Modificar el grosor del borde inferior de la fila 3
            tbl.Rows(3).Borders(wdBorderBottom).LineWidth = wdLineWidth150pt
            tbl.Rows(1).Borders(wdBorderBottom).LineWidth = wdLineWidth150pt
            tbl.cell(3, 1).Borders(wdBorderRight).LineWidth = wdLineWidth150pt
        Else
            MsgBox "El segundo sumando no es un numero valido."
        End If
    Else
        MsgBox "El primer sumando no es un numero valido."
    End If
End Function

Sub FillTableCells(cell As cell, value As String)
    Dim i As Integer
    Dim lenValue As Integer
    lenValue = Len(value)
    
    For i = lenValue To 1 Step -1
        cell.Range.Text = Mid(value, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Realizar_Resta()
    Dim minuendo As String
    Dim Sustraendo As String
    
    ' Solicita el primer minuendo
    minuendo = InputBox("Ingrese el minuendo:", "minuendo")
    
    ' Verifica si el usuario ingreso un numero valido
    If IsNumeric(minuendo) Then
        ' Solicita el sustraendo
        Sustraendo = InputBox("Ingrese el sustraendo:", "Sustraendo")
        
        ' Verifica si el usuario ingreso un numero valido
        If IsNumeric(Sustraendo) Then
            ' Crear la tabla en Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)
            
            ' Establecer los bordes de las celdas
            tbl.Borders.Enable = True
            
            ' Modificar el grosor del borde inferior de la fila 1
            tbl.Rows(1).Borders(wdBorderBottom).LineWidth = wdLineWidth150pt
    
            ' Llena la tabla desde la ultima celda hacia atras
            FillTableCells tbl.cell(2, tbl.Columns.Count), minuendo
            FillTableCells tbl.cell(3, tbl.Columns.Count - 14), "-" ' Coloca el s�mbolo 'X' aqu�
            FillTableCells tbl.cell(3, tbl.Columns.Count), Sustraendo
    
            ' Modificar el grosor del borde inferior de la fila 3
            tbl.Rows(3).Borders(wdBorderBottom).LineWidth = wdLineWidth150pt
    
            tbl.cell(3, 1).Borders(wdBorderRight).LineWidth = wdLineWidth150pt
        Else
            MsgBox "El sustraendo no es un numero valido."
        End If
    Else
        MsgBox "El minuendo no es un numero valido."
    End If
End Function

Sub FillTableCell(cell As cell, value As String)
    Dim i As Integer
    Dim lenValue As Integer
    lenValue = Len(value)
    
    For i = lenValue To 1 Step -1
        cell.Range.Text = Mid(value, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Realizar_Multiplicacion()
    Dim Multiplicando As String
    Dim Multiplicador As String
    
    ' Solicita el Multiplicador
    Multiplicador = InputBox("Ingrese el Multiplicador:", "Multiplicador")

    ' Verifica si el usuario ingreso un numero valido
    If IsNumeric(Multiplicador) Then
        ' Crear la tabla en Word
        Dim doc As Document
        Set doc = ActiveDocument
        Dim tbl As table
        Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)

        ' Establecer los bordes de las celdas
        tbl.Borders.Enable = True

        ' Modificar el grosor del borde inferior de la fila 1
        tbl.Rows(1).Borders(wdBorderBottom).LineWidth = wdLineWidth150pt

        ' Llena la tabla desde la ultima celda hacia atras
        FillTableCells tbl.cell(2, tbl.Columns.Count), Multiplicando
        FillTableCells tbl.cell(3, tbl.Columns.Count - 14), "X" ' Coloca el s�mbolo 'X' aqu�
        FillTableCells tbl.cell(3, tbl.Columns.Count), Multiplicador

        ' Modificar el grosor del borde inferior de la fila 3
        tbl.Rows(3).Borders(wdBorderBottom).LineWidth = wdLineWidth150pt

        tbl.cell(3, 1).Borders(wdBorderRight).LineWidth = wdLineWidth150pt
    Else
        MsgBox "El Multiplicador no es un numero valido."
    End If
End Function

Sub FillTableCel(cell As cell, value As String)
    Dim i As Integer
    Dim lenValue As Integer
    lenValue = Len(value)

    For i = lenValue To 1 Step -1
        cell.Range.Text = Mid(value, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Hacer_divisiones_desarrolladas()
    Dim Dividendo As Double
    Dim Divisor As Double

    ' Ingresa el dividendo y el divisor
    Dividendo = InputBox("Ingrese el dividendo:", "Division Larga")
    Divisor = InputBox("Ingrese el divisor:", "Division Larga")

    ' Crea una tabla de 2 filas y 2 columnas
    Selection.Tables.Add Range:=Selection.Range, NumRows:=2, NumColumns:=2

    ' Establece el ancho de la primera columna para el dividendo
    Selection.Tables(1).Columns(1).Width = InchesToPoints(2.5) ' Ajusta el ancho segun tus necesidades

    ' Establece el ancho de la segunda columna para el divisor (5 cm)
    Selection.Tables(1).Columns(2).Width = CentimetersToPoints(5)

    ' Configura el espaciado entre celdas y el espaciado del parrafo
    Selection.Tables(1).cell(1, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(1, 2).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(2, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(2, 2).Range.Paragraphs.SpaceBefore = 0

    ' Establece la altura de las celdas
    Selection.Tables(1).Rows.HeightRule = wdRowHeightExactly
    Selection.Tables(1).Rows.Height = InchesToPoints(0.25) ' Ajusta la altura segun tus necesidades

    ' Oculta los bordes de la tabla
    Selection.Tables(1).Borders.Enable = False

    ' Agrega el dividendo en la primera celda de la primera fila
    Selection.Tables(1).cell(1, 1).Range.Text = Format(Dividendo, "0")

    ' Agrega el divisor en la segunda celda de la primera fila
    Selection.Tables(1).cell(1, 2).Range.Text = CStr(Divisor)

    ' Cambia el estilo del borde inferior de la celda 1,2 a un estilo de l�nea discontinua
    Selection.Tables(1).cell(1, 2).Borders(wdBorderBottom).LineStyle = wdLineStyleSingle  ' Establece el estilo de l�nea deseado

    ' Agrega el borde lateral central solo en la primera columna de la primera fila
    Selection.Tables(1).cell(1, 1).Borders(wdBorderRight).LineStyle = wdLineStyleSingle

    ' Puedes personalizar la tabla y su apariencia segun tus necesidades
End Function

Sub FilTableCell(cell As cell, value As String)
    Dim i As Integer
    Dim lenValue As Integer
    lenValue = Len(value)

    For i = lenValue To 1 Step -1
        cell.Range.Text = Mid(value, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Hacer_descomposicion_factorial()
    Dim Numero As Double

    ' Ingresa el Numero y el divisor
    Numero = InputBox("Ingrese el numero que deseas descomponer:", "Descomposicion factorial")

    ' Crea una tabla de 1 fila y 2 columnas
    Selection.Tables.Add Range:=Selection.Range, NumRows:=1, NumColumns:=2

    ' Establece el ancho de la primera columna para el Numero
    Selection.Tables(1).Columns(1).Width = InchesToPoints(1) ' Ajusta el ancho segun tus necesidades

    ' Configura el espaciado entre celdas y el espaciado del parrafo
    Selection.Tables(1).cell(1, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(1, 2).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).Rows.SpaceBetweenColumns = 0

    ' Establece la altura de las celdas
    Selection.Tables(1).Rows.HeightRule = wdRowHeightExactly
    Selection.Tables(1).Rows.Height = InchesToPoints(0.25) ' Ajusta la altura segun tus necesidades

    ' Oculta los bordes de la tabla
    Selection.Tables(1).Borders.Enable = False

    ' Agrega el Numero y el divisor en la tabla con espacios en blanco antes del divisor
    Selection.Tables(1).cell(1, 1).Range.Text = CStr(Numero)
    Selection.Tables(1).cell(1, 2).Range.Text = "   "

    ' Establece el ancho de la primera columna para el Numero
    Selection.Tables(1).Columns(2).Width = InchesToPoints(1.25)
    Selection.Tables(1).Columns(2).Width = InchesToPoints(1) ' Ajusta el ancho segun tus necesidades

    ' Agrega el borde lateral central solo en
    Selection.Tables(1).Columns(1).Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
    Selection.Tables(1).Columns(1).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
    Selection.Tables(1).Columns(2).Borders(wdBorderRight).LineStyle = wdLineStyleSingle

    ' Puedes personalizar la tabla y su apariencia segun tus necesidades
End Function

Function Castellano_3_y_4()
    Dim opcion As String
    opcion = InputBox("Seleccione la operacion que desea realizar:" & vbCrLf & "1. Suma" & vbCrLf & "2. Resta" & vbCrLf & "3. Multiplicacion" & vbCrLf & "4. Division Larga" & vbCrLf & "5. Descomposicion factorial", "Menu de Operaciones")

    Select Case opcion
        Case "1"
            Hacer_Suma
        Case "2"
            Hacer_Resta
        Case "3"
            Hacer_Multiplicacion
        Case "4"
            Realizar_divisiones_desarrolladas
        Case "5"
            Realizar_descomposicion_factorial
        Case Else
            MsgBox "Opcion no valida. Por favor, seleccione una operacion valida."
    End Select
End Function

Function Hacer_Suma()
    Dim PrimerSumando As String
    Dim SegundoSumando As String
    
    ' Solicita el primer sumando
    PrimerSumando = InputBox("Ingrese el primer sumando:", "Suma")
    
    ' Verifica si el usuario ingresÃ³ un nÃºmero vÃ¡lido
    If IsNumeric(PrimerSumando) Then
        ' Solicita el segundo sumando
        SegundoSumando = InputBox("Ingrese el segundo sumando:", "Suma")
        
        ' Verifica si el usuario ingresÃ³ un nÃºmero vÃ¡lido
        If IsNumeric(SegundoSumando) Then
            ' Crear la tabla en Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)
            
            ' Ocultar los bordes de las celdas
            tbl.Borders.Enable = False
            
            ' Llena la tabla desde la Ãºltima celda hacia atrÃ¡s
            FillTableCells tbl.cell(2, tbl.Columns.Count), PrimerSumando
            FillTableCells tbl.cell(3, tbl.Columns.Count - 14), "+" ' Coloca el sÃ­mbolo 'X' aquÃ­
            FillTableCells tbl.cell(3, tbl.Columns.Count), SegundoSumando
            
            ' Mostrar los bordes laterales y el borde inferior de la primera y tercera fila
            For Each cell In tbl.Rows(1).Cells
                cell.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
            Next cell
            
            For Each cell In tbl.Rows(3).Cells
            cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell
            
            For Each cell In tbl.Rows(4).Cells
            cell.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
            cell.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
            cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell
        Else
            MsgBox "El segundo sumando no es un numero valido."
        End If
    Else
        MsgBox "El primer sumando no es un numero valido."
    End If
End Function

Sub FillTableCels(cell As cell, value As String)
    Dim i As Integer
    Dim lenValue As Integer
    lenValue = Len(value)
    
    For i = lenValue To 1 Step -1
        cell.Range.Text = Mid(value, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Hacer_Resta()
    Dim minuendo As String
    Dim Sustraendo As String
    
    ' Solicita el primer minuendo
    minuendo = InputBox("Ingrese el minuendo:", "minuendo")
    
    ' Verifica si el usuario ingreso un numero valido
    If IsNumeric(minuendo) Then
        ' Solicita el sustraendo
        Sustraendo = InputBox("Ingrese el sustraendo:", "Sustraendo")
        
        ' Verifica si el usuario ingreso un numero valido
        If IsNumeric(Sustraendo) Then
            ' Crear la tabla en Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)
            
            ' Establecer los bordes de las celdas
            tbl.Borders.Enable = False
            
            ' Llena la tabla desde la Ãºltima celda hacia atrÃ¡s
            FillTableCells tbl.cell(2, tbl.Columns.Count), minuendo
            FillTableCells tbl.cell(3, tbl.Columns.Count - 14), "-" ' Coloca el sÃ­mbolo 'X' aquÃ­
            FillTableCells tbl.cell(3, tbl.Columns.Count), Sustraendo

           ' Mostrar los bordes laterales y el borde inferior de la primera y tercera fila
            For Each cell In tbl.Rows(1).Cells
            cell.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
            cell.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
            cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            cell.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
            Next cell
            
            For Each cell In tbl.Rows(3).Cells
            cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell
            
            For Each cell In tbl.Rows(4).Cells
            cell.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
            cell.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
            cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell
        Else
            MsgBox "El sustraendo no es un numero valido."
        End If
    Else
        MsgBox "El minuendo no es un numero valido."
    End If
End Function

Sub FllTableCell(cell As cell, value As String)
    Dim i As Integer
    Dim lenValue As Integer
    lenValue = Len(value)
    
    For i = lenValue To 1 Step -1
        cell.Range.Text = Mid(value, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Hacer_Multiplicacion()
    Dim Multiplicando As String
    Dim Multiplicador As String
    
    ' Solicita el primer sumando
    Multiplicando = InputBox("Ingrese el Multiplicando:", "Multiplicacion")
    
    ' Verifica si el usuario ingresÃ³ un nÃºmero vÃ¡lido
    If IsNumeric(Multiplicando) Then
        ' Solicita el segundo sumando
        Multiplicador = InputBox("Ingrese el segundo Multiplicador:", "Multiplicacion")
        
        ' Verifica si el usuario ingresÃ³ un nÃºmero vÃ¡lido
        If IsNumeric(Multiplicador) Then
            ' Crear la tabla en Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)
            
            ' Ocultar los bordes de las celdas
            tbl.Borders.Enable = False
            
            ' Llena la tabla desde la Ãºltima celda hacia atrÃ¡s
            FillTableCells tbl.cell(2, tbl.Columns.Count), Multiplicando
            FillTableCells tbl.cell(3, tbl.Columns.Count - 14), "X" ' Coloca el sÃ­mbolo 'X' aquÃ­
            FillTableCells tbl.cell(3, tbl.Columns.Count), Multiplicador
            
            ' Mostrar los bordes laterales y el borde inferior de la primera y tercera fila
            For Each cell In tbl.Rows(1).Cells
                cell.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
            Next cell
            
            tbl.cell(3, tbl.Columns.Count).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
            
            For Each cell In tbl.Rows(3).Cells
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell
            
            For Each cell In tbl.Rows(4).Cells
                cell.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell
        Else
            MsgBox "El multiplicando no es un numero valido."
        End If
    Else
        MsgBox "El Multiplicador es un numero valido."
    End If
End Function

Sub FillTbleCells(cell As cell, value As String)
    Dim i As Integer
    Dim lenValue As Integer
    lenValue = Len(value)
    
    For i = lenValue To 1 Step -1
        cell.Range.Text = Mid(value, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Realizar_divisiones_desarrolladas()
    Dim Dividendo As Double
    Dim Divisor As Double
    
    ' Ingresa el dividendo y el divisor
    Dividendo = InputBox("Ingrese el dividendo:", "Division Larga")
    Divisor = InputBox("Ingrese el divisor:", "Division Larga")

    ' Crea una tabla de 2 filas y 2 columnas
    Selection.Tables.Add Range:=Selection.Range, NumRows:=2, NumColumns:=2
    
    ' Establece el ancho de la primera columna para el dividendo
    Selection.Tables(1).Columns(1).Width = InchesToPoints(2.5) ' Ajusta el ancho segun tus necesidades
    
    ' Establece el ancho de la segunda columna para el divisor (5 cm)
    Selection.Tables(1).Columns(2).Width = CentimetersToPoints(5)
    
    ' Configura el espaciado entre celdas y el espaciado del parrafo
    Selection.Tables(1).cell(1, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(1, 2).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(2, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(2, 2).Range.Paragraphs.SpaceBefore = 0
    
    ' Establece la altura de las celdas
    Selection.Tables(1).Rows.HeightRule = wdRowHeightExactly
    Selection.Tables(1).Rows.Height = InchesToPoints(0.25) ' Ajusta la altura segun tus necesidades
    
    ' Oculta los bordes de la tabla
    Selection.Tables(1).Borders.Enable = False
    
    ' Agrega el dividendo en la primera celda de la primera fila
    Selection.Tables(1).cell(1, 1).Range.Text = Format(Dividendo, "0")
    
    ' Agrega el divisor en la segunda celda de la primera fila
    Selection.Tables(1).cell(1, 2).Range.Text = CStr(Divisor)
    
    ' Cambia el estilo del borde inferior de la celda 1,2 a un estilo de linea discontinua
    Selection.Tables(1).cell(1, 2).Borders(wdBorderBottom).LineStyle = wdLineStyleSingle  ' Establece el estilo de linea deseado
    
    ' Agrega el borde lateral central solo en la primera columna de la primera fila
    Selection.Tables(1).cell(1, 1).Borders(wdBorderRight).LineStyle = wdLineStyleSingle

    ' Puedes personalizar la tabla y su apariencia segun tus necesidades
End Function

Sub FillTabeCells(cell As cell, number As String)
    Dim i As Integer
    For i = 1 To Len(number)
        cell.Range.Text = Mid(number, i, 1)
        Set cell = cell.Next
    Next i
End Sub


Function Realizar_descomposicion_factorial()
    Dim Numero As Double
    
    ' Ingresa el Numero y el divisor
    Numero = InputBox("Ingrese el numero que deseas descomponer:", "Descomposicion factorial")
    
    ' Crea una tabla de 1 fila y 2 columnas
    Selection.Tables.Add Range:=Selection.Range, NumRows:=1, NumColumns:=2
    
    ' Establece el ancho de la primera columna para el Numero
    Selection.Tables(1).Columns(1).Width = InchesToPoints(1) ' Ajusta el ancho segun tus necesidades
    
    ' Configura el espaciado entre celdas y el espaciado del parrafo
    Selection.Tables(1).cell(1, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(1, 2).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).Rows.SpaceBetweenColumns = 0
    
    ' Establece la altura de las celdas
    Selection.Tables(1).Rows.HeightRule = wdRowHeightExactly
    Selection.Tables(1).Rows.Height = InchesToPoints(0.25) ' Ajusta la altura segun tus necesidades

    ' Oculta los bordes de la tabla
    Selection.Tables(1).Borders.Enable = False
    
    ' Agrega el Numero y el divisor en la tabla con espacios en blanco antes del divisor
    Selection.Tables(1).cell(1, 1).Range.Text = CStr(Numero)
    Selection.Tables(1).cell(1, 2).Range.Text = "   "
    
    ' Establece el ancho de la primera columna para el Numero
    Selection.Tables(1).Columns(2).Width = InchesToPoints(1.25)
    Selection.Tables(1).Columns(2).Width = InchesToPoints(1) ' Ajusta el ancho segun tus necesidades

    ' Agrega el borde lateral central solo en la primera columna
    Selection.Tables(1).Columns(1).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
    
    ' Puedes personalizar la tabla y su apariencia segun tus necesidades
End Function

Function Castellano_5()
    Dim opcion As String
    opcion = InputBox("Seleccione la operacion que desea realizar:" & vbCrLf & "1. Suma" & vbCrLf & "2. Resta" & vbCrLf & "3. Multiplicacion" & vbCrLf & "4. Division Larga" & vbCrLf & "5. Descomposicion factorial" & vbCrLf & "6. Hacer raices", "Menu de Operaciones")

    Select Case opcion
        Case "1"
            Suma
        Case "2"
            Resta
        Case "3"
            Multiplicacion
        Case "4"
            divisiones_desarrolladas
        Case "5"
            descomposicion_factorial
        Case "6"
            Hacer_raices
        Case Else
            MsgBox "Opcion no valida. Por favor, seleccione una operacion valida."
    End Select
End Function

Function Suma()
    Dim PrimerSumando As String
    Dim SegundoSumando As String
    
    ' Solicita el primer sumando
    PrimerSumando = InputBox("Ingrese el primer sumando:", "Suma")
    
    ' Verifica si el usuario ingresÃƒÂ³ un nÃƒÂºmero vÃƒÂ¡lido
    If IsNumeric(PrimerSumando) Then
        ' Solicita el segundo sumando
        SegundoSumando = InputBox("Ingrese el segundo sumando:", "Suma")
        
        ' Verifica si el usuario ingresÃƒÂ³ un nÃƒÂºmero vÃƒÂ¡lido
        If IsNumeric(SegundoSumando) Then
            ' Crear la tabla en Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)
            
            ' Ocultar los bordes de las celdas
            tbl.Borders.Enable = False
            
            ' Llena la tabla desde la ÃƒÂºltima celda hacia atrÃƒÂ¡s
            FillTableCells tbl.cell(2, tbl.Columns.Count), PrimerSumando
            FillTableCells tbl.cell(3, tbl.Columns.Count - 14), "+" ' Coloca el sÃƒÂ­mbolo 'X' aquÃƒÂ­
            FillTableCells tbl.cell(3, tbl.Columns.Count), SegundoSumando
                        
            For Each cell In tbl.Rows(3).Cells
            cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell
            
        Else
            MsgBox "El segundo sumando no es un nÃƒÂºmero vÃƒÂ¡lido."
        End If
    Else
        MsgBox "El primer sumando no es un nÃƒÂºmero vÃƒÂ¡lido."
    End If
End Function

Sub FillTableCe(cell As cell, value As String)
    Dim i As Integer
    Dim lenValue As Integer
    lenValue = Len(value)
    
    For i = lenValue To 1 Step -1
        cell.Range.Text = Mid(value, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Resta()
    Dim minuendo As String
    Dim Sustraendo As String
    
    ' Solicita el primer minuendo
    minuendo = InputBox("Ingrese el minuendo:", "minuendo")
    
    ' Verifica si el usuario ingreso un numero valido
    If IsNumeric(minuendo) Then
        ' Solicita el sustraendo
        Sustraendo = InputBox("Ingrese el sustraendo:", "Sustraendo")
        
        ' Verifica si el usuario ingreso un numero valido
        If IsNumeric(Sustraendo) Then
            ' Crear la tabla en Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)
            
            ' Establecer los bordes de las celdas
            tbl.Borders.Enable = False
            
           
            ' Llena la tabla desde la ÃƒÂºltima celda hacia atrÃƒÂ¡s
            FillTableCells tbl.cell(2, tbl.Columns.Count), minuendo
            FillTableCells tbl.cell(3, tbl.Columns.Count - 14), "-" ' Coloca el sÃƒÂ­mbolo 'X' aquÃƒÂ­
            FillTableCells tbl.cell(3, tbl.Columns.Count), Sustraendo
            
            For Each cell In tbl.Rows(3).Cells
            cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell
            
        Else
            MsgBox "El sustraendo no es un numero valido."
        End If
    Else
        MsgBox "El minuendo no es un numero valido."
    End If
End Function

Sub FillTabeCell(cell As cell, value As String)
    Dim i As Integer
    Dim lenValue As Integer
    lenValue = Len(value)
    
    For i = lenValue To 1 Step -1
        cell.Range.Text = Mid(value, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Multiplicacion()
    Dim Multiplicando As String
    Dim Multiplicador As String
    
    ' Solicita el primer sumando
    Multiplicando = InputBox("Ingrese multiplicando:", "Multiplicacion")
    
    ' Verifica si el usuario ingresÃƒÂ³ un nÃƒÂºmero vÃƒÂ¡lido
    If IsNumeric(Multiplicando) Then
        ' Solicita el segundo sumando
        Multiplicador = InputBox("Ingrese el segundo Multiplicador:", "Multiplicacion")
        
        ' Verifica si el usuario ingresÃƒÂ³ un nÃƒÂºmero vÃƒÂ¡lido
        If IsNumeric(Multiplicador) Then
            ' Crear la tabla en Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)
            
            ' Ocultar los bordes de las celdas
            tbl.Borders.Enable = False
            
            ' Llena la tabla desde la ÃƒÂºltima celda hacia atrÃƒÂ¡s
            FillTableCells tbl.cell(2, tbl.Columns.Count), Multiplicando
            FillTableCells tbl.cell(3, tbl.Columns.Count - 14), "X" ' Coloca el sÃƒÂ­mbolo 'X' aquÃƒÂ­
            FillTableCells tbl.cell(3, tbl.Columns.Count), Multiplicador
            
            ' Mostrar los bordes laterales y el borde inferior de la primera y tercera fila
            
            
            For Each cell In tbl.Rows(3).Cells
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell

        Else
            MsgBox "El multiplicando no es un nÃƒÂºmero vÃƒÂ¡lido."
        End If
    Else
        MsgBox "El Multiplicador es un nÃƒÂºmero vÃƒÂ¡lido."
    End If
End Function

Sub FillTableClls(cell As cell, value As String)
    Dim i As Integer
    Dim lenValue As Integer
    lenValue = Len(value)
    
    For i = lenValue To 1 Step -1
        cell.Range.Text = Mid(value, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function divisiones_desarrolladas()
    Dim Dividendo As Double
    Dim Divisor As Double
    
    ' Ingresa el dividendo y el divisor
    Dividendo = InputBox("Ingrese el dividendo:", "Division Larga")
    Divisor = InputBox("Ingrese el divisor:", "Division Larga")

    ' Crea una tabla de 2 filas y 2 columnas
    Selection.Tables.Add Range:=Selection.Range, NumRows:=2, NumColumns:=2
    
    ' Establece el ancho de la primera columna para el dividendo
    Selection.Tables(1).Columns(1).Width = InchesToPoints(2.5) ' Ajusta el ancho segun tus necesidades
    
    ' Establece el ancho de la segunda columna para el divisor (5 cm)
    Selection.Tables(1).Columns(2).Width = CentimetersToPoints(5)
    
    ' Configura el espaciado entre celdas y el espaciado del parrafo
    Selection.Tables(1).cell(1, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(1, 2).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(2, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(2, 2).Range.Paragraphs.SpaceBefore = 0
    
    ' Establece la altura de las celdas
    Selection.Tables(1).Rows.HeightRule = wdRowHeightExactly
    Selection.Tables(1).Rows.Height = InchesToPoints(0.25) ' Ajusta la altura segun tus necesidades
    
    ' Oculta los bordes de la tabla
    Selection.Tables(1).Borders.Enable = False
    
    ' Agrega el dividendo en la primera celda de la primera fila
    Selection.Tables(1).cell(1, 1).Range.Text = Format(Dividendo, "0")
    
    ' Agrega el divisor en la segunda celda de la primera fila
    Selection.Tables(1).cell(1, 2).Range.Text = CStr(Divisor)
    
    ' Cambia el estilo del borde inferior de la celda 1,2 a un estilo de linea discontinua
    Selection.Tables(1).cell(1, 2).Borders(wdBorderBottom).LineStyle = wdLineStyleSingle  ' Establece el estilo de linea deseado
    
    ' Agrega el borde lateral central solo en la primera columna de la primera fila
    Selection.Tables(1).cell(1, 1).Borders(wdBorderRight).LineStyle = wdLineStyleSingle

    ' Puedes personalizar la tabla y su apariencia segun tus necesidades
End Function

Sub FillTabCells(cell As cell, number As String)
    Dim i As Integer
    For i = 1 To Len(number)
        cell.Range.Text = Mid(number, i, 1)
        Set cell = cell.Next
    Next i
End Sub


Function descomposicion_factorial()
    Dim Numero As Double
    
    ' Ingresa el Numero y el divisor
    Numero = InputBox("Ingrese el numero que deseas descomponer:", "Descomposicion factorial")
    
    ' Crea una tabla de 1 fila y 2 columnas
    Selection.Tables.Add Range:=Selection.Range, NumRows:=1, NumColumns:=2
    
    ' Establece el ancho de la primera columna para el Numero
    Selection.Tables(1).Columns(1).Width = InchesToPoints(1) ' Ajusta el ancho segun tus necesidades
    
    ' Configura el espaciado entre celdas y el espaciado del parrafo
    Selection.Tables(1).cell(1, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(1, 2).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).Rows.SpaceBetweenColumns = 0
    
    ' Establece la altura de las celdas
    Selection.Tables(1).Rows.HeightRule = wdRowHeightExactly
    Selection.Tables(1).Rows.Height = InchesToPoints(0.25) ' Ajusta la altura segun tus necesidades

    ' Oculta los bordes de la tabla
    Selection.Tables(1).Borders.Enable = False
    
    ' Agrega el Numero y el divisor en la tabla con espacios en blanco antes del divisor
    Selection.Tables(1).cell(1, 1).Range.Text = CStr(Numero)
    Selection.Tables(1).cell(1, 2).Range.Text = "   "
    
    ' Establece el ancho de la primera columna para el Numero
    Selection.Tables(1).Columns(2).Width = InchesToPoints(1.25)
    Selection.Tables(1).Columns(2).Width = InchesToPoints(1) ' Ajusta el ancho segun tus necesidades

    ' Agrega el borde lateral central solo en la primera columna
    Selection.Tables(1).Columns(1).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
    
    ' Puedes personalizar la tabla y su apariencia segun tus necesidades
End Function

Function Hacer_raices()
    ' Crear una tabla de 2x3
    Dim tabla As table
    Set tabla = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=2, NumColumns:=3)
    
    ' Desactivar los bordes de la tabla por defecto
    tabla.Borders.Enable = False

    ' Establecer el ancho de la primera columna en 1 cm
    tabla.Columns(1).Width = CentimetersToPoints(1)

    ' Establecer el ancho de la segunda columna en 5 cm
    tabla.Columns(2).Width = CentimetersToPoints(2.5)
    tabla.Columns(3).Width = CentimetersToPoints(2.5)

    ' Activar los bordes que deseas en la primera fila
    With tabla.Rows(1)
        .Cells(1).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
        .Cells(2).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
        .Cells(2).Borders(wdBorderTop).LineStyle = wdLineStyleSingle
    End With
    With tabla.Rows(2)
        .Cells(2).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
    End With
    
    ' Preguntar al usuario por el índice de la raíz
    Dim indiceRaiz As String
    indiceRaiz = InputBox("¿Que indice tiene la raiz?")
    
    ' Colocar el índice en la celda 1,1 como superíndice
    With tabla.cell(1, 1).Range
        .Text = indiceRaiz
        .Font.Superscript = True
        .ParagraphFormat.Alignment = wdAlignParagraphRight
    End With
    
    ' Preguntar al usuario por el radicando
    Dim radicando As String
    radicando = InputBox("¿Cual es el radicando?")
    
    ' Verificar si se proporcionó un valor para el radicando
    If radicando = "" Then
        MsgBox "Debes ingresar un valor en el radicando", vbExclamation, "Error"
        Exit Function
    End If
    
    ' Colocar el radicando en la celda 1,2
    tabla.cell(1, 2).Range.Text = radicando
        
    With tabla.Rows(2)
        .Cells(2).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
    End With

End Function

Function English_1st_2nd()
    Dim choice As String
    choice = InputBox("Select the math operation you want to perform:" & vbCrLf & "1. Addition" & vbCrLf & "2. Subtraction" & vbCrLf & "3. Multiplication" & vbCrLf & "4. Long Division" & vbCrLf & "5. Factorial Decomposition", "Operation Menu")

    Select Case choice
        Case "1"
            Perform_Addition
        Case "2"
            Perform_Subtraction
        Case "3"
            Perform_Multiplication
        Case "4"
            Perform_LongDivision
        Case "5"
            Perform_Factorial_Decomposition
        Case Else
            MsgBox "Invalid option. Please select a valid operation."
    End Select
End Function

Function Perform_Addition()
    Dim FirstOperand As String
    Dim SecondOperand As String
    
    ' Request the first operand
    FirstOperand = InputBox("Enter the first operand:", "Addition")
    
    ' Check if the user entered a valid number
    If IsNumeric(FirstOperand) Then
        ' Request the second operand
        SecondOperand = InputBox("Enter the second operand:", "Addition")
        
        ' Check if the user entered a valid number
        If IsNumeric(SecondOperand) Then
            ' Create a table in Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)
            
            ' Hide cell borders
            tbl.Borders.Enable = True
            
            ' Fill the table from the last cell backwards
            FillTableCells tbl.cell(2, tbl.Columns.Count), FirstOperand
            FillTableCells tbl.cell(3, tbl.Columns.Count - 14), "+" ' Place the '+' symbol here
            FillTableCells tbl.cell(3, tbl.Columns.Count), SecondOperand
            
            ' Modify the thickness of the bottom border of row 3
            tbl.Rows(3).Borders(wdBorderBottom).LineWidth = wdLineWidth150pt

            tbl.cell(3, 1).Borders(wdBorderRight).LineWidth = wdLineWidth150pt
        Else
            MsgBox "The second operand is not a valid number."
        End If
    Else
        MsgBox "The first operand is not a valid number."
    End If
End Function

Sub FillTableCes(cell As cell, value As String)
    Dim i As Integer
    Dim lenValue As Integer
    lenValue = Len(value)
    
    For i = lenValue To 1 Step -1
        cell.Range.Text = Mid(value, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Perform_Subtraction()
    Dim minuend As String
    Dim Subtrahend As String
    
    ' Request the first minuend
    minuend = InputBox("Enter the minuend:", "minuend")
    
    ' Check if the user entered a valid number
    If IsNumeric(minuend) Then
        ' Request the Subtrahend
        Subtrahend = InputBox("Enter the Subtrahend:", "Subtrahend")
        
        ' Check if the user entered a valid number
        If IsNumeric(Subtrahend) Then
            ' Create a table in Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)
            
            ' Set cell borders
            tbl.Borders.Enable = True
            
            ' Modify the thickness of the bottom border of row 1
            tbl.Rows(1).Borders(wdBorderBottom).LineWidth = wdLineWidth150pt

            ' Fill the table from the last cell backwards
            FillTableCel tbl.cell(2, tbl.Columns.Count), minuend
            FillTableCel tbl.cell(3, tbl.Columns.Count - 14), "-" ' Place the '-' symbol here
            FillTableCel tbl.cell(3, tbl.Columns.Count), Subtrahend

            ' Modify the thickness of the bottom border of row 3
            tbl.Rows(3).Borders(wdBorderBottom).LineWidth = wdLineWidth150pt

            tbl.cell(3, 1).Borders(wdBorderRight).LineWidth = wdLineWidth150pt
        Else
            MsgBox "The Subtrahend is not a valid number."
        End If
    Else
        MsgBox "The minuend is not a valid number."
    End If
End Function

Sub FillTableCl(cell As cell, value As String)
    Dim i As Integer
    Dim lenValue As Integer
    lenValue = Len(value)
    
    For i = lenValue To 1 Step -1
        cell.Range.Text = Mid(value, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Perform_Multiplication()
    Dim Multiplicand As String
    Dim Multiplier As String
    
    ' Request the Multiplicand
    Multiplicand = InputBox("Enter the Multiplicand:", "Multiplicand")
    
    ' Check if the user entered a valid number
    If IsNumeric(Multiplicand) Then
        ' Request the Multiplier
        Multiplier = InputBox("Enter the Multiplier:", "Multiplier")
        
        ' Check if the user entered a valid number
        If IsNumeric(Multiplier) Then
            ' Create a table in Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)
            
            ' Set cell borders
            tbl.Borders.Enable = True
            
            ' Modify the thickness of the bottom border of row 1
            tbl.Rows(1).Borders(wdBorderBottom).LineWidth = wdLineWidth150pt
            
            ' Fill the table from the last cell backwards
            FillTableCels tbl.cell(2, tbl.Columns.Count), Multiplicand
            FillTableCels tbl.cell(3, tbl.Columns.Count - 14), "X" ' Place the 'X' symbol here
            FillTableCels tbl.cell(3, tbl.Columns.Count), Multiplier

            ' Modify the thickness of the bottom border of row 3
            tbl.Rows(3).Borders(wdBorderBottom).LineWidth = wdLineWidth150pt

            tbl.cell(3, 1).Borders(wdBorderRight).LineWidth = wdLineWidth150pt
        Else
            MsgBox "The Multiplier is not a valid number."
        End If
    Else
        MsgBox "The Multiplicand is not a valid number."
    End If
End Function

Sub FillTablels(cell As cell, value As String)
    Dim i As Integer
    Dim lenValue As Integer
    lenValue = Len(value)
    
    For i = lenValue To 1 Step -1
        cell.Range.Text = Mid(value, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Perform_LongDivision()
    Dim Dividend As Double
    Dim Divisor As Double
    
    ' Enter the dividend and divisor
    Dividend = InputBox("Enter the dividend:", "Long Division")
    Divisor = InputBox("Enter the divisor:", "Long Division")

    ' Create a table with 2 rows and 2 columns
    Selection.Tables.Add Range:=Selection.Range, NumRows:=2, NumColumns:=2
    
    ' Set the width of the first column for the dividend
    Selection.Tables(1).Columns(1).Width = InchesToPoints(2.5) ' Adjust the width as needed
    
    ' Set the width of the second column for the divisor (5 cm)
    Selection.Tables(1).Columns(2).Width = CentimetersToPoints(5)
    
    ' Configure cell and paragraph spacing
    Selection.Tables(1).cell(1, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(1, 2).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(2, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(2, 2).Range.Paragraphs.SpaceBefore = 0
    
    ' Set cell heights
    Selection.Tables(1).Rows.HeightRule = wdRowHeightExactly
    Selection.Tables(1).Rows.Height = InchesToPoints(0.25) ' Adjust the height as needed
    
    ' Hide table borders
    Selection.Tables(1).Borders.Enable = False
    
    ' Add the dividend in the first cell of the first row
    Selection.Tables(1).cell(1, 1).Range.Text = Format(Dividend, "0")
    
    ' Add the divisor in the second cell of the first row
    Selection.Tables(1).cell(1, 2).Range.Text = CStr(Divisor)
    
    ' Change the style of the bottom border of cell 1,2 to a dashed line style
    Selection.Tables(1).cell(1, 2).Borders(wdBorderBottom).LineStyle = wdLineStyleSingle ' Set the desired line style
    
    ' Add the central side border only in the first column of the first row
    Selection.Tables(1).cell(1, 1).Borders(wdBorderRight).LineStyle = wdLineStyleSingle

    ' You can customize the table and its appearance as needed
End Function

Sub FillleCell(cell As cell, value As String)
    Dim i As Integer
    Dim lenValue As Integer
    lenValue = Len(value)
    
    For i = lenValue To 1 Step -1
        cell.Range.Text = Mid(value, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Perform_Factorial_Decomposition()
    Dim number As Double
    
    ' Enter the Number you want to decompose
    number = InputBox("Enter the number you want to decompose:", "Factorial Decomposition")
    
    ' Create a table with 1 row and 2 columns
    Selection.Tables.Add Range:=Selection.Range, NumRows:=1, NumColumns:=2
    
    ' Set the width of the first column for the Number
    Selection.Tables(1).Columns(1).Width = InchesToPoints(1) ' Adjust the width as needed
    
    ' Configure cell and paragraph spacing
    Selection.Tables(1).cell(1, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(1, 2).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).Rows.SpaceBetweenColumns = 0
    
    ' Set cell heights
    Selection.Tables(1).Rows.HeightRule = wdRowHeightExactly
    Selection.Tables(1).Rows.Height = InchesToPoints(0.25) ' Adjust the height as needed

    ' Hide table borders
    Selection.Tables(1).Borders.Enable = False
    
    ' Add the Number and some spaces before the divisor
    Selection.Tables(1).cell(1, 1).Range.Text = CStr(number)
    Selection.Tables(1).cell(1, 2).Range.Text = "   "
    
    ' Set the width of the second column for the Number
    Selection.Tables(1).Columns(2).Width = InchesToPoints(1.25) ' Adjust the width as needed
    
    ' Add the central side border only in the first column
    Selection.Tables(1).Columns(1).Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
    Selection.Tables(1).Columns(1).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
    Selection.Tables(1).Columns(2).Borders(wdBorderRight).LineStyle = wdLineStyleSingle

    ' You can customize the table and its appearance as needed
End Function

Function English_3rd_4th()
    Dim choice As String
    choice = InputBox("Select the math operation you want to perform:" & vbCrLf & "1. Addition" & vbCrLf & "2. Subtraction" & vbCrLf & "3. Multiplication" & vbCrLf & "4. Long Division" & vbCrLf & "5. Factorial Decomposition", "Operation Menu")

    Select Case choice
        Case "1"
            Make_Addition
        Case "2"
            Make_Subtraction
        Case "3"
            Make_Multiplication
        Case "4"
            Make_LongDivision
        Case "5"
            Make_FactorialDecomposition
        Case Else
            MsgBox "Invalid option. Please select a valid operation."
    End Select
End Function

Function Make_Addition()
    Dim FirstOperand As String
    Dim SecondOperand As String
    
    ' Request the first operand
    FirstOperand = InputBox("Enter the first operand:", "Addition")
    
    ' Check if the user entered a valid number
    If IsNumeric(FirstOperand) Then
        ' Request the second operand
        SecondOperand = InputBox("Enter the second operand:", "Addition")
        
        ' Check if the user entered a valid number
        If IsNumeric(SecondOperand) Then
            ' Create a table in Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)
            
            ' Hide cell borders
            tbl.Borders.Enable = False
            
            ' Fill the table from the last cell backward
            FillTableCells tbl.cell(2, tbl.Columns.Count), FirstOperand
            FillTableCells tbl.cell(3, tbl.Columns.Count - 14), "+" ' Place the symbol '+' here
            FillTableCells tbl.cell(3, tbl.Columns.Count), SecondOperand
            
            ' Show side borders and bottom border for the first and third rows
            For Each cell In tbl.Rows(1).Cells
                cell.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
            Next cell
            
            For Each cell In tbl.Rows(3).Cells
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell
            
            For Each cell In tbl.Rows(4).Cells
                cell.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell
        Else
            MsgBox "The second operand is not a valid number."
        End If
    Else
        MsgBox "The first operand is not a valid number."
    End If
End Function

Sub illTableCells(cell As cell, value As String)
    Dim i As Integer
    Dim lenValue As Integer
    lenValue = Len(value)
    
    For i = lenValue To 1 Step -1
        cell.Range.Text = Mid(value, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Make_Subtraction()
    Dim minuend As String
    Dim Subtrahend As String
    
    ' Request the first minuend
    minuend = InputBox("Enter the minuend:", "Subtraction")
    
    ' Check if the user entered a valid number
    If IsNumeric(minuend) Then
        ' Request the subtrahend
        Subtrahend = InputBox("Enter the Subtrahend:", "Subtraction")
        
        ' Check if the user entered a valid number
        If IsNumeric(Subtrahend) Then
            ' Create a table in Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)
            
            ' Hide cell borders
            tbl.Borders.Enable = False
            
            ' Fill the table from the last cell backward
            FillTableCell tbl.cell(2, tbl.Columns.Count), minuend
            FillTableCell tbl.cell(3, tbl.Columns.Count - 14), "-" ' Place the symbol '-' here
            FillTableCell tbl.cell(3, tbl.Columns.Count), Subtrahend

            ' Show side borders and bottom border for the first and third rows
            For Each cell In tbl.Rows(1).Cells
                cell.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
            Next cell
            
            For Each cell In tbl.Rows(3).Cells
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell
            
            For Each cell In tbl.Rows(4).Cells
                cell.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell
        Else
            MsgBox "The subtrahend is not a valid number."
        End If
    Else
        MsgBox "The minuend is not a valid number."
    End If
End Function

Sub FiTableCell(cell As cell, value As String)
    Dim i As Integer
    Dim lenValue As Integer
    lenValue = Len(value)
    
    For i = lenValue To 1 Step -1
        cell.Range.Text = Mid(value, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Make_Multiplication()
    Dim Multiplicand As String
    Dim Multiplier As String
    
    ' Request the first multiplicand
    Multiplicand = InputBox("Enter the Multiplicand:", "Multiplication")
    
    ' Check if the user entered a valid number
    If IsNumeric(Multiplicand) Then
        ' Request the second multiplier
        Multiplier = InputBox("Enter the second Multiplier:", "Multiplication")
        
        ' Check if the user entered a valid number
        If IsNumeric(Multiplier) Then
            ' Create a table in Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)
            
            ' Hide cell borders
            tbl.Borders.Enable = False
            
            ' Fill the table from the last cell backward
            FillableCells tbl.cell(2, tbl.Columns.Count), Multiplicand
            FillableCells tbl.cell(3, tbl.Columns.Count - 14), "X" ' Place the symbol 'X' here
            FillableCells tbl.cell(3, tbl.Columns.Count), Multiplier

            ' Show side borders and the bottom border for the first and third rows
            For Each cell In tbl.Rows(1).Cells
                cell.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
            Next cell
            
            tbl.cell(3, tbl.Columns.Count).Borders(wdBorderRight).LineStyle = wdLineStyleSingle

            For Each cell In tbl.Rows(3).Cells
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell
            
            For Each cell In tbl.Rows(4).Cells
                cell.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell
        Else
            MsgBox "The multiplier is not a valid number."
        End If
    Else
        MsgBox "The Multiplicand is not a valid number."
    End If
End Function

Sub FillableCells(cell As cell, value As String)
    Dim i As Integer
    Dim lenValue As Integer
    lenValue = Len(value)
    
    For i = lenValue To 1 Step -1
        cell.Range.Text = Mid(value, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Make_LongDivision()
    Dim Dividend As Double
    Dim Divisor As Double
    
    ' Enter the dividend and divisor
    Dividend = InputBox("Enter the dividend:", "Long Division")
    Divisor = InputBox("Enter the divisor:", "Long Division")

    ' Create a 2x2 table
    Selection.Tables.Add Range:=Selection.Range, NumRows:=2, NumColumns:=2
    
    ' Set the width of the first column for the dividend
    Selection.Tables(1).Columns(1).Width = InchesToPoints(2.5) ' Adjust the width as needed
    
    ' Set the width of the second column for the divisor (5 cm)
    Selection.Tables(1).Columns(2).Width = CentimetersToPoints(5)
    
    ' Configure cell spacing and paragraph spacing
    Selection.Tables(1).cell(1, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(1, 2).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(2, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(2, 2).Range.Paragraphs.SpaceBefore = 0
    
    ' Set cell heights
    Selection.Tables(1).Rows.HeightRule = wdRowHeightExactly
    Selection.Tables(1).Rows.Height = InchesToPoints(0.25) ' Adjust the height as needed
    
    ' Hide table borders
    Selection.Tables(1).Borders.Enable = False
    
    ' Add the dividend to the first cell in the first row
    Selection.Tables(1).cell(1, 1).Range.Text = Format(Dividend, "0")
    
    ' Add the divisor to the second cell in the first row
    Selection.Tables(1).cell(1, 2).Range.Text = CStr(Divisor)
    
    ' Change the style of the bottom border of cell 1,2 to a dashed line
    Selection.Tables(1).cell(1, 2).Borders(wdBorderBottom).LineStyle = wdLineStyleSingle  ' Set the desired line style
    
    ' Add the central side border only in the first column of the first row
    Selection.Tables(1).cell(1, 1).Borders(wdBorderRight).LineStyle = wdLineStyleSingle

    ' Customize the table and its appearance as needed
End Function

Function Make_FactorialDecomposition()
    Dim number As Double
    
    ' Enter the number you want to decompose
    number = InputBox("Enter the number you want to decompose:", "Factorial Decomposition")
    
    ' Create a 1x2 table
    Selection.Tables.Add Range:=Selection.Range, NumRows:=1, NumColumns:=2
    
    ' Set the width of the first column for the number
    Selection.Tables(1).Columns(1).Width = InchesToPoints(1) ' Adjust the width as needed
    
    ' Configure cell spacing and paragraph spacing
    Selection.Tables(1).cell(1, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(1, 2).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).Rows.SpaceBetweenColumns = 0
    
    ' Set cell heights
    Selection.Tables(1).Rows.HeightRule = wdRowHeightExactly
    Selection.Tables(1).Rows.Height = InchesToPoints(0.25) ' Adjust the height as needed

    ' Hide table borders
    Selection.Tables(1).Borders.Enable = False
    
    ' Add the Number and the divisor to the table with spaces before the divisor
    Selection.Tables(1).cell(1, 1).Range.Text = CStr(number)
    Selection.Tables(1).cell(1, 2).Range.Text = "   "
    
    ' Set the width of the second column for the Number
    Selection.Tables(1).Columns(2).Width = InchesToPoints(1.25)
    Selection.Tables(1).Columns(2).Width = InchesToPoints(1) ' Adjust the width as needed

    ' Add the central side border only in the first column
    Selection.Tables(1).Columns(1).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
    
    ' Customize the table and its appearance as needed
End Function

Function English_from_5th_grade()
    Dim choice As String
    choice = InputBox("Select the math operation you want to perform:" & vbCrLf & "1. Addition" & vbCrLf & "2. Subtraction" & vbCrLf & "3. Multiplication" & vbCrLf & "4. Long Division" & vbCrLf & "5. Factorial Decomposition" & vbCrLf & "6. Square roots", "Operation Menu")

    Select Case choice
        Case "1"
            Addition
        Case "2"
            Subtraction
            Case "3"
            Multiplication
        Case "4"
            LongDivision
        Case "5"
            FactorialDecomposition
        Case "6"
            roots
        Case Else
            MsgBox "Invalid option. Please select a valid operation."
    End Select
End Function

Function Addition()
    Dim FirstOperand As String
    Dim SecondOperand As String
    
    ' Request the first operand
    FirstOperand = InputBox("Enter the first operand:", "Addition")
    
    ' Check if the user entered a valid number
    If IsNumeric(FirstOperand) Then
        ' Request the second operand
        SecondOperand = InputBox("Enter the second operand:", "Addition")
        
        ' Check if the user entered a valid number
        If IsNumeric(SecondOperand) Then
            ' Create a table in Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)
            
            ' Hide cell borders
            tbl.Borders.Enable = False
            
            ' Fill the table from the last cell backward
            FillTableCels tbl.cell(2, tbl.Columns.Count), FirstOperand
            FillTableCels tbl.cell(3, tbl.Columns.Count - 14), "+" ' Place the '+' symbol here
            FillTableCels tbl.cell(3, tbl.Columns.Count), SecondOperand
                        
            For Each cell In tbl.Rows(3).Cells
            cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell
            
        Else
            MsgBox "The second operand is not a valid number."
        End If
    Else
        MsgBox "The first operand is not a valid number."
    End If
End Function

Sub FillTable(cell As cell, value As String)
    Dim i As Integer
    Dim lenValue As Integer
    lenValue = Len(value)
    
    For i = lenValue To 1 Step -1
        cell.Range.Text = Mid(value, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Subtraction()
    Dim minuend As String
    Dim Subtrahend As String
    
    ' Request the first minuend
    minuend = InputBox("Enter the minuend:", "minuend")
    
    ' Check if the user entered a valid number
    If IsNumeric(minuend) Then
        ' Request the Subtrahend
        Subtrahend = InputBox("Enter the Subtrahend:", "Subtrahend")
        
        ' Check if the user entered a valid number
        If IsNumeric(Subtrahend) Then
            ' Create a table in Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)
            
            ' Set cell borders
            tbl.Borders.Enable = False
            
            ' Fill the table from the last cell backward
            FillTableCell tbl.cell(2, tbl.Columns.Count), minuend
            FillTableCell tbl.cell(3, tbl.Columns.Count - 14), "-" ' Place the '-' symbol here
            FillTableCell tbl.cell(3, tbl.Columns.Count), Subtrahend
            
            For Each cell In tbl.Rows(3).Cells
            cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell
            
        Else
            MsgBox "The Subtrahend is not a valid number."
        End If
    Else
        MsgBox "The minuend is not a valid number."
    End If
End Function

Sub FillTableCellx(cell As cell, value As String)
    Dim i As Integer
    Dim lenValue As Integer
    lenValue = Len(value)
    
    For i = lenValue To 1 Step -1
        cell.Range.Text = Mid(value, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Multiplication()
    Dim Multiplicand As String
    Dim Multiplier As String
    
    ' Request the first operand
    Multiplicand = InputBox("Enter the Multiplicand:", "Multiplication")
    
    ' Check if the user entered a valid number
    If IsNumeric(Multiplicand) Then
        ' Request the second operand
        Multiplier = InputBox("Enter the Multiplier:", "Multiplication")
        
        ' Check if the user entered a valid number
        If IsNumeric(Multiplier) Then
            ' Create a table in Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)
            
            ' Hide cell borders
            tbl.Borders.Enable = False
            
            ' Fill the table from the last cell backward
            FillTableCells tbl.cell(2, tbl.Columns.Count), Multiplicand
            FillTableCells tbl.cell(3, tbl.Columns.Count - 14), "X" ' Place the 'X' symbol here
            FillTableCells tbl.cell(3, tbl.Columns.Count), Multiplier
            
            ' Show the bottom border of the first and third row
            For Each cell In tbl.Rows(3).Cells
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell

        Else
            MsgBox "The Multiplier is not a valid number."
        End If
    Else
        MsgBox "The Multiplicand is not a valid number."
    End If
End Function

Sub FillTableCello(cell As cell, value As String)
    Dim i As Integer
    Dim lenValue As Integer
    lenValue = Len(value)
    
    For i = lenValue To 1 Step -1
        cell.Range.Text = Mid(value, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function LongDivision()
    Dim Dividend As Double
    Dim Divisor As Double
    
    ' Enter the dividend and divisor
    Dividend = InputBox("Enter the dividend:", "Long Division")
    Divisor = InputBox("Enter the divisor:", "Long Division")

    ' Create a table with 2 rows and 2 columns
    Selection.Tables.Add Range:=Selection.Range, NumRows:=2, NumColumns:=2
    
    ' Set the width of the first column for the dividend
    Selection.Tables(1).Columns(1).Width = InchesToPoints(2.5) ' Adjust the width as needed
    
    ' Set the width of the second column for the divisor (5 cm)
    Selection.Tables(1).Columns(2).Width = CentimetersToPoints(5)
    
    ' Configure cell spacing and paragraph spacing
    Selection.Tables(1).cell(1, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(1, 2).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(2, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(2, 2).Range.Paragraphs.SpaceBefore = 0
    
    ' Set the height of the cells
    Selection.Tables(1).Rows.HeightRule = wdRowHeightExactly
    Selection.Tables(1).Rows.Height = InchesToPoints(0.25) ' Adjust the height as needed
    
    ' Hide the table borders
    Selection.Tables(1).Borders.Enable = False
    
    ' Add the dividend to the first cell of the first row
    Selection.Tables(1).cell(1, 1).Range.Text = Format(Dividend, "0")
    
    ' Add the divisor to the second cell of the first row
    Selection.Tables(1).cell(1, 2).Range.Text = CStr(Divisor)
    
    ' Change the style of the bottom border of cell 1,2 to a dashed line style
    Selection.Tables(1).cell(1, 2).Borders(wdBorderBottom).LineStyle = wdLineStyleSingle  ' Set the desired line style
    
    ' Add the central right border only in the first column of the first row
    Selection.Tables(1).cell(1, 1).Borders(wdBorderRight).LineStyle = wdLineStyleSingle

    ' You can customize the table and its appearance as needed
End Function

Sub FillTablCells(cell As cell, number As String)
    Dim i As Integer
    For i = 1 To Len(number)
        cell.Range.Text = Mid(number, i, 1)
        Set cell = cell.Next
    Next i
End Sub

Function FactorialDecomposition()
    Dim number As Double
    
    ' Enter the number you want to decompose
    number = InputBox("Enter the number you want to decompose:", "Factorial Decomposition")
    
    ' Create a table with 1 row and 2 columns
    Selection.Tables.Add Range:=Selection.Range, NumRows:=1, NumColumns:=2
    
    ' Set the width of the first column for the Number
    Selection.Tables(1).Columns(1).Width = InchesToPoints(1) ' Adjust the width as needed
    
    ' Configure cell spacing and paragraph spacing
    Selection.Tables(1).cell(1, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(1, 2).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).Rows.SpaceBetweenColumns = 0
    
    ' Set the height of the cells
    Selection.Tables(1).Rows.HeightRule = wdRowHeightExactly
    Selection.Tables(1).Rows.Height = InchesToPoints(0.25) ' Adjust the height as needed

    ' Hide the table borders
    Selection.Tables(1).Borders.Enable = False
    
    ' Add the Number and some blank space before the divisor
    Selection.Tables(1).cell(1, 1).Range.Text = CStr(number)
    Selection.Tables(1).cell(1, 2).Range.Text = "   "
    
    ' Set the width of the second column for the Number
    Selection.Tables(1).Columns(2).Width = InchesToPoints(1.25)
    Selection.Tables(1).Columns(2).Width = InchesToPoints(1) ' Adjust the width as needed

    ' Add the central right border only in the first column
    Selection.Tables(1).Columns(1).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
    
    ' You can customize the table and its appearance as needed
End Function

Function roots()
    ' Create a table with 2 rows and 3 columns
    Dim table As table
    Set table = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=2, NumColumns:=3)
    
    ' Deactivate the default table borders
    table.Borders.Enable = False

    ' Set the width of the first column to 1 cm
    table.Columns(1).Width = CentimetersToPoints(1)

    ' Set the width of the second column to 5 cm
    table.Columns(2).Width = CentimetersToPoints(2.5)
    table.Columns(3).Width = CentimetersToPoints(2.5)

    ' Activate the desired borders in the first row
    With table.Rows(1)
        .Cells(1).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
        .Cells(2).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
        .Cells(2).Borders(wdBorderTop).LineStyle = wdLineStyleSingle
    End With
    With table.Rows(2)
        .Cells(2).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
    End With

    ' Ask the user for the root index
    Dim rootIndex As String
    rootIndex = InputBox("What is the root index?")

    ' Place the root index as a superscript in cell 1,1
    With table.cell(1, 1).Range
        .Text = rootIndex
        .Font.Superscript = True
        .ParagraphFormat.Alignment = wdAlignParagraphRight
    End With

    ' Ask the user for the radicand
    Dim radicand As String
    radicand = InputBox("What is the radicand?")

    ' Check if a value for the radicand was provided
    If radicand = "" Then
        MsgBox "You must enter a value for the radicand", vbExclamation, "Error"
        Exit Function
    End If

    ' Add the radicand in cell 1,2
    table.cell(1, 2).Range.Text = radicand

End Function

Function Gaulois_CP_CE1()
    Dim choix As String
    choix = InputBox("Selectionnez l'operation que vous souhaitez effectuer:" & vbCrLf & "1. Addition" & vbCrLf & "2. Soustraction" & vbCrLf & "3. Multiplication" & vbCrLf & "4. Division Longue" & vbCrLf & "5. Decomposition factorielle", "Menu des Operations")

    Select Case choix
        Case "1"
            Realiser_Addition
        Case "2"
            Realiser_Soustraction
        Case "3"
            Realiser_Multiplication
        Case "4"
            Faire_divisions_developpees
        Case "5"
            Faire_decomposition_factorielle
        Case Else
            MsgBox "Option invalide. Veuillez selectionner une operation valide."
    End Select
End Function

Function Realiser_Addition()
    Dim PremierAddend As String
    Dim DeuxiemeAddend As String
    
    ' Demande le premier addend
    PremierAddend = InputBox("Entrez le premier addend:", "Addition")
    
    ' Verifie si l'utilisateur a saisi un nombre valide
    If IsNumeric(PremierAddend) Then
        ' Demande le deuxieme addend
        DeuxiemeAddend = InputBox("Entrez le deuxieme addend:", "Addition")
        
        ' Verifie si l'utilisateur a saisi un nombre valide
        If IsNumeric(DeuxiemeAddend) Then
            ' Creer le tableau dans Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)
            
            ' Masquer les bordures des cellules
            tbl.Borders.Enable = True
            
            ' Remplir le tableau depuis la derniere cellule vers l'arriere
            RemplirCellules tbl.cell(2, tbl.Columns.Count), PremierAddend
            RemplirCellules tbl.cell(3, tbl.Columns.Count - 14), "+" ' Placez le symbole '+' ici
            RemplirCellules tbl.cell(3, tbl.Columns.Count), DeuxiemeAddend
            
            ' Modifier l'epaisseur de la bordure inferieure de la ligne 3
            tbl.Rows(3).Borders(wdBorderBottom).LineWidth = wdLineWidth150pt

            tbl.cell(3, 1).Borders(wdBorderRight).LineWidth = wdLineWidth150pt
        Else
            MsgBox "Le deuxieme addend n'est pas un nombre valide."
        End If
    Else
        MsgBox "Le premier addend n'est pas un nombre valide."
    End If
End Function

Sub RemplirCellules(cell As cell, value As String)
    Dim i As Integer
    Dim lenValue As Integer
    lenValue = Len(value)
    
    For i = lenValue To 1 Step -1
        cell.Range.Text = Mid(value, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Realiser_Soustraction()
    Dim minuend As String
    Dim Subtrahend As String
    
    ' Demande le premier minuend
    minuend = InputBox("Entrez le minuend:", "minuend")
    
    ' Verifie si l'utilisateur a saisi un nombre valide
    If IsNumeric(minuend) Then
        ' Demande le subtrahend
        Subtrahend = InputBox("Entrez le subtrahend:", "Subtrahend")
        
        ' Verifie si l'utilisateur a saisi un nombre valide
        If IsNumeric(Subtrahend) Then
            ' Creer le tableau dans Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)
            
            ' Definir les bordures des cellules
            tbl.Borders.Enable = True
            
            ' Modifier l'epaisseur de la bordure inferieure de la ligne 1
            tbl.Rows(1).Borders(wdBorderBottom).LineWidth = wdLineWidth150pt

            ' Remplir le tableau depuis la derniere cellule vers l'arriere
            RemplirCellules tbl.cell(2, tbl.Columns.Count), minuend
            RemplirCellules tbl.cell(3, tbl.Columns.Count - 14), "-" ' Placez le symbole '-' ici
            RemplirCellules tbl.cell(3, tbl.Columns.Count), Subtrahend

            ' Modifier l'epaisseur de la bordure inferieure de la ligne 3
            tbl.Rows(3).Borders(wdBorderBottom).LineWidth = wdLineWidth150pt

            tbl.cell(3, 1).Borders(wdBorderRight).LineWidth = wdLineWidth150pt

        Else
            MsgBox "Le subtrahend n'est pas un nombre valide."
        End If
    Else
        MsgBox "Le minuend n'est pas un nombre valide."
    End If
End Function

Sub RemplirCell(cell As cell, value As String)
    Dim i As Integer
    Dim lenValue As Integer
    lenValue = Len(value)
    
    For i = lenValue To 1 Step -1
        cell.Range.Text = Mid(value, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Realiser_Multiplication()
    Dim Multiplicand As String
    Dim Multiplier As String
    
    ' Demande le Multiplicand
    Multiplicand = InputBox("Entrez le Multiplicand:", "Multiplicand")
    
    ' Verifie si l'utilisateur a saisi un nombre valide
    If IsNumeric(Multiplicand) Then
        ' Demande le Multiplier
        Multiplier = InputBox("Entrez le Multiplier:", "Multiplier")
        
        ' Verifie si l'utilisateur a saisi un nombre valide
        If IsNumeric(Multiplier) Then
            ' Creer le tableau dans Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)
            
            ' Definir les bordures des cellules
            tbl.Borders.Enable = True
            
            ' Modifier l'epaisseur de la bordure inferieure de la ligne 1
            tbl.Rows(1).Borders(wdBorderBottom).LineWidth = wdLineWidth150pt
            
            ' Remplir le tableau depuis la derniere cellule vers l'arriere
            RemplirCellules tbl.cell(2, tbl.Columns.Count), Multiplicand
            RemplirCellules tbl.cell(3, tbl.Columns.Count - 14), "X" ' Placez le symbole 'X' ici
            RemplirCellules tbl.cell(3, tbl.Columns.Count), Multiplier

            ' Modifier l'epaisseur de la bordure inferieure de la ligne 3
            tbl.Rows(3).Borders(wdBorderBottom).LineWidth = wdLineWidth150pt

            tbl.cell(3, 1).Borders(wdBorderRight).LineWidth = wdLineWidth150pt
        Else
            MsgBox "Le Multiplier n'est pas un nombre valide."
        End If
    Else
        MsgBox "Le Multiplicand n'est pas un nombre valide."
    End If
End Function

Sub RemplirTableCel(cell As cell, value As String)
    Dim i As Integer
    Dim lenValue As Integer
    lenValue = Len(value)
    
    For i = lenValue To 1 Step -1
        cell.Range.Text = Mid(value, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Faire_divisions_developpees()
    Dim Dividende As Double
    Dim Diviseur As Double
    
    ' Entrez le dividende et le diviseur
    Dividende = InputBox("Entrez le dividende:", "Division Longue")
    Diviseur = InputBox("Entrez le diviseur:", "Division Longue")

    ' Cree un tableau de 2 lignes et 2 colonnes
    Selection.Tables.Add Range:=Selection.Range, NumRows:=2, NumColumns:=2
    
    ' Definir la largeur de la premiere colonne pour le dividende
    Selection.Tables(1).Columns(1).Width = InchesToPoints(2.5) ' Ajustez la largeur en fonction de vos besoins
    
    ' Definir la largeur de la deuxieme colonne pour le diviseur (5 cm)
    Selection.Tables(1).Columns(2).Width = CentimetersToPoints(5)
    
    ' Configurer l'espacement entre les cellules et l'espacement de paragraphe
    Selection.Tables(1).cell(1, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(1, 2).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(2, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(2, 2).Range.Paragraphs.SpaceBefore = 0
    
    ' Definir la hauteur des lignes
    Selection.Tables(1).Rows.HeightRule = wdRowHeightExactly
    Selection.Tables(1).Rows.Height = InchesToPoints(0.25) ' Ajustez la hauteur en fonction de vos besoins
    
    ' Masquer les bordures du tableau
    Selection.Tables(1).Borders.Enable = False
    
    ' Ajouter le dividende dans la premiere cellule de la premiere ligne
    Selection.Tables(1).cell(1, 1).Range.Text = Format(Dividende, "0")
    
    ' Ajouter le diviseur dans la deuxieme cellule de la premiere ligne
    Selection.Tables(1).cell(1, 2).Range.Text = CStr(Diviseur)
    
    ' Changer le style de la bordure inferieure de la cellule 1,2 en un style de ligne discontinue
    Selection.Tables(1).cell(1, 2).Borders(wdBorderBottom).LineStyle = wdLineStyleSingle  ' Definissez le style de ligne souhaite
    
    ' Ajouter la bordure laterale centrale uniquement dans la premiere colonne de la premiere ligne
    Selection.Tables(1).cell(1, 1).Borders(wdBorderRight).LineStyle = wdLineStyleSingle

    ' Vous pouvez personnaliser le tableau et son apparence selon vos besoins
End Function

Sub FilTableCells(cell As cell, number As String)
    Dim i As Integer
    For i = 1 To Len(number)
        cell.Range.Text = Mid(number, i, 1)
        Set cell = cell.Next
    Next i
End Sub

Function Faire_decomposition_factorielle()
    Dim Nombre As Double
    
    ' Entrez le Nombre que vous souhaitez decomposer
    Nombre = InputBox("Entrez le nombre que vous souhaitez decomposer:", "Decomposition factorielle")
    
    ' Cree un tableau de 1 ligne et 2 colonnes
    Selection.Tables.Add Range:=Selection.Range, NumRows:=1, NumColumns:=2
    
    ' Definir la largeur de la premiere colonne pour le Nombre
    Selection.Tables(1).Columns(1).Width = InchesToPoints(1) ' Ajustez la largeur en fonction de vos besoins
    
    ' Configurer l'espacement entre les cellules et l'espacement de paragraphe
    Selection.Tables(1).cell(1, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(1, 2).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).Rows.SpaceBetweenColumns = 0
    
    ' Definir la hauteur des lignes
    Selection.Tables(1).Rows.HeightRule = wdRowHeightExactly
    Selection.Tables(1).Rows.Height = InchesToPoints(0.25) ' Ajustez la hauteur en fonction de vos besoins

    ' Masquer les bordures du tableau
    Selection.Tables(1).Borders.Enable = False
    
    ' Ajouter le Nombre et un espace blanc avant le diviseur dans le tableau
    Selection.Tables(1).cell(1, 1).Range.Text = CStr(Nombre)
    Selection.Tables(1).cell(1, 2).Range.Text = "   "
    
    ' Definir la largeur de la premiere colonne pour le Nombre
    Selection.Tables(1).Columns(2).Width = InchesToPoints(1.25)
    Selection.Tables(1).Columns(2).Width = InchesToPoints(1) ' Ajustez la largeur en fonction de vos besoins

    ' Ajouter la bordure laterale centrale uniquement dans la premiere colonne
    Selection.Tables(1).Columns(1).Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
    Selection.Tables(1).Columns(1).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
    Selection.Tables(1).Columns(2).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
    ' Vous pouvez personnaliser le tableau et son apparence selon vos besoins
End Function

Function Gaulois_CE2_CM1()
    Dim choix As String
    choix = InputBox("Selectionnez l'operation que vous souhaitez effectuer:" & vbCrLf & "1. Addition" & vbCrLf & "2. Soustraction" & vbCrLf & "3. Multiplication" & vbCrLf & "4. Division Longue" & vbCrLf & "5. Decomposition factorielle", "Menu des operations")

    Select Case choix
        Case "1"
            Additions
        Case "2"
            Soustractions
        Case "3"
            Multiplications
        Case "4"
            divisions_developpees
        Case "5"
            decomposition_factorielle
        Case Else
            MsgBox "Option non valide. Veuillez selectionner une operation valide."
    End Select
End Function

Function Additions()
    Dim PremierAddend As String
    Dim DeuxiemeAddend As String

    ' Demande le premier addend
    PremierAddend = InputBox("Entrez le premier addend :", "Addition")

    ' Verifie si l'utilisateur a entre un nombre valide
    If IsNumeric(PremierAddend) Then
        ' Demande le deuxieme addend
        DeuxiemeAddend = InputBox("Entrez le deuxieme addend :", "Addition")

        ' Verifie si l'utilisateur a entre un nombre valide
        If IsNumeric(DeuxiemeAddend) Then
            ' Cree le tableau dans Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)

            ' Masque les bordures des cellules
            tbl.Borders.Enable = False

            ' Remplit le tableau de la derniere cellule vers l'arriere
            RemplirCellules(tbl.cell(2, tbl.Columns.Count), PremierAddend)
            RemplirCellules tbl.cell(3, tbl.Columns.Count - 14), "+" ' Place le symbole '+' ici
            RemplirCellules tbl.cell(3, tbl.Columns.Count), DeuxiemeAddend

            ' Affiche les bordures laterales et la bordure inferieure de la premiere et de la troisieme ligne
            For Each cell In tbl.Rows(1).Cells
                cell.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
            Next cell

            For Each cell In tbl.Rows(3).Cells
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell

            For Each cell In tbl.Rows(4).Cells
                cell.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell
        Else
            MsgBox "Le deuxieme addend n'est pas un nombre valide."
        End If
    Else
        MsgBox "Le premier addend n'est pas un nombre valide."
    End If
End Function

Sub RemplirCellulex(cell As cell, valeur As String)
    Dim i As Integer
    Dim longValeur As Integer
    longValeur = Len(valeur)

    For i = longValeur To 1 Step -1
        cell.Range.Text = Mid(valeur, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Soustraction()
    Dim minuend As String
    Dim Subtrahend As String

    ' Demande le minuend
    minuend = InputBox("Entrez le minuend :", "Soustraction")

    ' Verifie si l'utilisateur a entre un nombre valide
    If IsNumeric(minuend) Then
        ' Demande le subtrahend
        Subtrahend = InputBox("Entrez le subtrahend :", "Soustraction")

        ' Verifie si l'utilisateur a entre un nombre valide
        If IsNumeric(Subtrahend) Then
            ' Cree le tableau dans Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)

            ' Masque les bordures des cellules
            tbl.Borders.Enable = False

            ' Remplit le tableau de la derniere cellule vers l'arriere
            RemplirCellules(tbl.cell(2, tbl.Columns.Count), minuend)
            RemplirCellules tbl.cell(3, tbl.Columns.Count - 14), "-" ' Place le symbole '-' ici
            RemplirCellules tbl.cell(3, tbl.Columns.Count), Subtrahend

            ' Affiche les bordures laterales et la bordure inferieure de la premiere et de la troisieme ligne
            For Each cell In tbl.Rows(1).Cells
                cell.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
            Next cell

            For Each cell In tbl.Rows(3).Cells
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell

            For Each cell In tbl.Rows(4).Cells
                cell.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell
        Else
            MsgBox "Le subtrahend n'est pas un nombre valide."
        End If
    Else
        MsgBox "Le minuend n'est pas un nombre valide."
    End If
End Function

Function Multiplications()
    Dim Multiplicand As String
    Dim Multiplier As String

    ' Demande le multiplicand
    Multiplicand = InputBox("Entrez le multiplicand :", "Multiplication")

    ' Verifie si l'utilisateur a entre un nombre valide
    If IsNumeric(Multiplicand) Then
        ' Demande le multiplier
        Multiplier = InputBox("Entrez le multiplier :", "Multiplication")

        ' Verifie si l'utilisateur a entre un nombre valide
        If IsNumeric(Multiplier) Then
            ' Cree le tableau dans Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)

            ' Masque les bordures des cellules
            tbl.Borders.Enable = False

            ' Remplit le tableau de la derniere cellule vers l'arriere
            RemplirCellules(tbl.cell(2, tbl.Columns.Count), Multiplicand)
            RemplirCellules tbl.cell(3, tbl.Columns.Count - 14), "X" ' Place le symbole 'X' ici
            RemplirCellules tbl.cell(3, tbl.Columns.Count), Multiplier

            ' Affiche les bordures laterales et la bordure inferieure de la premiere et de la troisieme ligne
            For Each cell In tbl.Rows(1).Cells
                cell.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
            Next cell

            tbl.cell(3, tbl.Columns.Count).Borders(wdBorderRight).LineStyle = wdLineStyleSingle

            For Each cell In tbl.Rows(3).Cells
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell

            For Each cell In tbl.Rows(4).Cells
                cell.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell
        Else
            MsgBox "Le multiplier n'est pas un nombre valide."
        End If
    Else
        MsgBox "Le multiplicand n'est pas un nombre valide."
    End If
End Function

Sub RemplirCellulea(cell As cell, valeur As String)
    Dim i As Integer
    Dim longValeur As Integer
    longValeur = Len(valeur)

    For i = longValeur To 1 Step -1
        cell.Range.Text = Mid(valeur, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function divisions_developpees()
    Dim Dividende As Double
    Dim Diviseur As Double

    ' Saisit le dividende et le diviseur
    Dividende = InputBox("Entrez le dividende :", "Division longue")
    Diviseur = InputBox("Entrez le diviseur :", "Division longue")

    ' Cree un tableau de 2 lignes et 2 colonnes
    Selection.Tables.Add Range:=Selection.Range, NumRows:=2, NumColumns:=2

    ' Definit la largeur de la premiere colonne pour le dividende
    Selection.Tables(1).Columns(1).Width = InchesToPoints(2.5) ' Ajustez la largeur en fonction de vos besoins

    ' Definit la largeur de la deuxieme colonne pour le diviseur (5 cm)
    Selection.Tables(1).Columns(2).Width = CentimetersToPoints(5)

    ' Configure l'espacement entre les cellules et l'espacement du paragraphe
    Selection.Tables(1).cell(1, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(1, 2).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(2, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(2, 2).Range.Paragraphs.SpaceBefore = 0

    ' Definit la hauteur des lignes
    Selection.Tables(1).Rows.HeightRule = wdRowHeightExactly
    Selection.Tables(1).Rows.Height = InchesToPoints(0.25) ' Ajustez la hauteur en fonction de vos besoins

    ' Masque les bordures de la table
    Selection.Tables(1).Borders.Enable = False

    ' Ajoute le dividende dans la premiere cellule de la premiere ligne
    Selection.Tables(1).cell(1, 1).Range.Text = Format(Dividende, "0")

    ' Ajoute le diviseur dans la deuxieme cellule de la premiere ligne
    Selection.Tables(1).cell(1, 2).Range.Text = CStr(Diviseur)

    ' Change le style de la bordure inferieure de la cellule 1,2 en une ligne en pointilles
    Selection.Tables(1).cell(1, 2).Borders(wdBorderBottom).LineStyle = wdLineStyleSingle  ' Definissez le style de ligne souhaite

    ' Ajoute la bordure laterale centrale uniquement dans la premiere colonne de la premiere ligne
    Selection.Tables(1).cell(1, 1).Borders(wdBorderRight).LineStyle = wdLineStyleSingle

    ' Vous pouvez personnaliser la table et son apparence en fonction de vos besoins
End Function

Sub RemplirCelluleas(cell As cell, Nombre As String)
    Dim i As Integer
    For i = 1 To Len(Nombre)
        cell.Range.Text = Mid(Nombre, i, 1)
        Set cell = cell.Next
    Next i
End Sub

Function decomposition_factorielle()
    Dim Numero As Double
    
    ' Saisissez le numero que vous souhaitez decomposer
    Numero = InputBox("Entrez le numero que vous souhaitez decomposer :", "Decomposition factorielle")
    
    ' Creez un tableau d'1 ligne et 2 colonnes
    Selection.Tables.Add Range:=Selection.Range, NumRows:=1, NumColumns:=2
    
    ' Definissez la largeur de la premiere colonne pour le Numero
    Selection.Tables(1).Columns(1).Width = InchesToPoints(1) ' Ajustez la largeur selon vos besoins
    
    ' Configurez l'espacement entre les cellules et l'espacement du paragraphe
    Selection.Tables(1).cell(1, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(1, 2).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).Rows.SpaceBetweenColumns = 0
    
    ' Definissez la hauteur des lignes
    Selection.Tables(1).Rows.HeightRule = wdRowHeightExactly
    Selection.Tables(1).Rows.Height = InchesToPoints(0.25) ' Ajustez la hauteur selon vos besoins

    ' Masquez les bordures du tableau
    Selection.Tables(1).Borders.Enable = False
    
    ' Ajoutez le Numero et un espace dans la deuxieme colonne
    Selection.Tables(1).cell(1, 1).Range.Text = CStr(Numero)
    Selection.Tables(1).cell(1, 2).Range.Text = "   "
    
    ' Definissez la largeur de la deuxieme colonne
    Selection.Tables(1).Columns(2).Width = InchesToPoints(1.25)
    Selection.Tables(1).Columns(2).Width = InchesToPoints(1) ' Ajustez la largeur selon vos besoins

    ' Ajoutez la bordure laterale centrale uniquement dans la premiere colonne
    Selection.Tables(1).Columns(1).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
    
    ' Vous pouvez personnaliser le tableau et son apparence selon vos besoins
End Function

Function Gaulois_CM2_et_plus()
    Dim choix As String
    choix = InputBox("Selectionnez l'operation que vous souhaitez effectuer:" & vbCrLf & "1. Addition" & vbCrLf & "2. Soustraction" & vbCrLf & "3. Multiplication" & vbCrLf & "4. Division longue" & vbCrLf & "5. Decomposition factorielle" & vbCrLf & "6. Faire des racines", "Menu des operations")

    Select Case choix
        Case "1"
            Realiser_Additions
        Case "2"
            Realiser_Soustractions
        Case "3"
            Realiser_Multiplications
        Case "4"
            Faire_divisions_developpeess
        Case "5"
            Faire_decomposition_factorielles
        Case "6"
            Faire_des_racines
        Case Else
            MsgBox "Option invalide. Veuillez selectionner une operation valide."
    End Select
End Function

Function Realiser_Additions()
    Dim PremierAddend As String
    Dim DeuxiemeAddend As String
    
    ' Demande le premier addend
    PremierAddend = InputBox("Entrez le premier addend :", "Addition")
    
    ' Verifie si l'utilisateur a entre un numero valide
    If IsNumeric(PremierAddend) Then
        ' Demande le deuxieme addend
        DeuxiemeAddend = InputBox("Entrez le deuxieme addend :", "Addition")
        
        ' Verifie si l'utilisateur a entre un numero valide
        If IsNumeric(DeuxiemeAddend) Then
            ' Cree le tableau dans Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)
            
            ' Masquer les bordures des cellules
            tbl.Borders.Enable = False
            
            ' Remplir le tableau depuis la derniere cellule vers l'arriere
            RemplirCellules tbl.cell(2, tbl.Columns.Count), PremierAddend
            RemplirCellules tbl.cell(3, tbl.Columns.Count - 14), "+" ' Place le symbole '+' ici
            RemplirCellules tbl.cell(3, tbl.Columns.Count), DeuxiemeAddend
                        
            For Each cell In tbl.Rows(3).Cells
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell
            
        Else
            MsgBox "Le deuxieme addend n'est pas un numero valide."
        End If
    Else
        MsgBox "Le premier addend n'est pas un numero valide."
    End If
End Function

Sub RemplirCellulesl(cell As cell, valeur As String)
    Dim i As Integer
    Dim lenValeur As Integer
    lenValeur = Len(valeur)
    
    For i = lenValeur To 1 Step -1
        cell.Range.Text = Mid(valeur, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Realiser_Soustractions()
    Dim minuend As String
    Dim Subtrahend As String
    
    ' Demande le minuend
    minuend = InputBox("Entrez le minuend :", "Soustraction")
    
    ' Verifie si l'utilisateur a entre un numero valide
    If IsNumeric(minuend) Then
        ' Demande le subtrahend
        Subtrahend = InputBox("Entrez le subtrahend :", "Soustraction")
        
        ' Verifie si l'utilisateur a entre un numero valide
        If IsNumeric(Subtrahend) Then
            ' Cree le tableau dans Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)
            
            ' Masquer les bordures des cellules
            tbl.Borders.Enable = False
            
            ' Remplir le tableau depuis la derniere cellule vers l'arriere
            RemplirCellules tbl.cell(2, tbl.Columns.Count), minuend
            RemplirCellules tbl.cell(3, tbl.Columns.Count - 14), "-" ' Place le symbole '-' ici
            RemplirCellules tbl.cell(3, tbl.Columns.Count), Subtrahend
            
            For Each cell In tbl.Rows(3).Cells
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell
            
        Else
            MsgBox "Le subtrahend n'est pas un numero valide."
        End If
    Else
        MsgBox "Le minuend n'est pas un numero valide."
    End If
End Function

Function Realiser_Multiplications()
    Dim Multiplicand As String
    Dim Multiplier As String
    
    ' Demande le premier multiplicand
    Multiplicand = InputBox("Entrez le premier multiplicand :", "Multiplication")
    
    ' Verifie si l'utilisateur a entre un numero valide
    If IsNumeric(Multiplicand) Then
        ' Demande le deuxieme multiplicand
        Multiplier = InputBox("Entrez le deuxieme multiplicand :", "Multiplication")
        
        ' Verifie si l'utilisateur a entre un numero valide
        If IsNumeric(Multiplier) Then
            ' Cree le tableau dans Word
            Dim doc As Document
            Set doc = ActiveDocument
            Dim tbl As table
            Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=15)
            
            ' Masquer les bordures des cellules
            tbl.Borders.Enable = False
            
            ' Remplir le tableau depuis la derniere cellule vers l'arriere
            RemplirCellules tbl.cell(2, tbl.Columns.Count), Multiplicand
            RemplirCellules tbl.cell(3, tbl.Columns.Count - 14), "X" ' Place le symbole 'X' ici
            RemplirCellules tbl.cell(3, tbl.Columns.Count), Multiplier
            
            For Each cell In tbl.Rows(3).Cells
                cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            Next cell

        Else
            MsgBox "Le multiplicand n'est pas un numero valide."
        End If
    Else
        MsgBox "Le multiplicateur n'est pas un numero valide."
    End If
End Function

Sub RemplirCellulesq(cell As cell, valeur As String)
    Dim i As Integer
    Dim lenValeur As Integer
    lenValeur = Len(valeur)
    
    For i = lenValeur To 1 Step -1
        cell.Range.Text = Mid(valeur, i, 1)
        If i > 1 Then
            Set cell = cell.Previous
        End If
    Next i
End Sub

Function Faire_divisions_developpeess()
    Dim Dividende As Double
    Dim Diviseur As Double
    
    ' Entrez le dividende et le diviseur
    Dividende = InputBox("Entrez le dividende :", "Division Longue")
    Diviseur = InputBox("Entrez le diviseur :", "Division Longue")

    ' Cree un tableau de 2 lignes et 2 colonnes
    Selection.Tables.Add Range:=Selection.Range, NumRows:=2, NumColumns:=2
    
    ' Definir la largeur de la premiere colonne pour le dividende
    Selection.Tables(1).Columns(1).Width = InchesToPoints(2.5) ' Ajustez la largeur selon vos besoins
    
    ' Definir la largeur de la deuxieme colonne pour le diviseur (5 cm)
    Selection.Tables(1).Columns(2).Width = CentimetersToPoints(5)
    
    ' Configurez l'espacement entre les cellules et l'espacement du paragraphe
    Selection.Tables(1).cell(1, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(1, 2).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(2, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(2, 2).Range.Paragraphs.SpaceBefore = 0
    
    ' Definir la hauteur des cellules
    Selection.Tables(1).Rows.HeightRule = wdRowHeightExactly
    Selection.Tables(1).Rows.Height = InchesToPoints(0.25) ' Ajustez la hauteur selon vos besoins
    
    ' Masquer les bordures du tableau
    Selection.Tables(1).Borders.Enable = False
    
    ' Ajoutez le dividende dans la premiere cellule de la premiere ligne
    Selection.Tables(1).cell(1, 1).Range.Text = Format(Dividende, "0")
    
    ' Ajoutez le diviseur dans la deuxieme cellule de la premiere ligne
    Selection.Tables(1).cell(1, 2).Range.Text = CStr(Diviseur)
    
    ' Changez le style de la bordure inferieure de la cellule 1,2 en un style de ligne discontinue
    Selection.Tables(1).cell(1, 2).Borders(wdBorderBottom).LineStyle = wdLineStyleSingle  ' Definissez le style de ligne souhaite
    
    ' Ajoutez la bordure laterale centrale uniquement dans la premiere colonne de la premiere ligne
    Selection.Tables(1).cell(1, 1).Borders(wdBorderRight).LineStyle = wdLineStyleSingle

    ' Vous pouvez personnaliser le tableau et son apparence selon vos besoins
End Function

Sub RemplirCellulesw(cell As cell, Nombre As String)
    Dim i As Integer
    For i = 1 To Len(Nombre)
        cell.Range.Text = Mid(Nombre, i, 1)
        Set cell = cell.Next
    Next i
End Sub

Function Faire_decomposition_factorielles()
    Dim Nombre As Double
    
    ' Entrez le nombre que vous souhaitez decomposer
    Nombre = InputBox("Entrez le nombre que vous souhaitez decomposer :", "Decomposition factorielle")
    
    ' Creez un tableau de 1 ligne et 2 colonnes
    Selection.Tables.Add Range:=Selection.Range, NumRows:=1, NumColumns:=2
    
    ' Definissez la largeur de la premiere colonne pour le Nombre
    Selection.Tables(1).Columns(1).Width = InchesToPoints(1) ' Ajustez la largeur selon vos besoins
    
    ' Configurez l'espacement entre les cellules et l'espacement du paragraphe
    Selection.Tables(1).cell(1, 1).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).cell(1, 2).Range.Paragraphs.SpaceBefore = 0
    Selection.Tables(1).Rows.SpaceBetweenColumns = 0
    
    ' Definissez la hauteur des cellules
    Selection.Tables(1).Rows.HeightRule = wdRowHeightExactly
    Selection.Tables(1).Rows.Height = InchesToPoints(0.25) ' Ajustez la hauteur selon vos besoins

    ' Masquez les bordures du tableau
    Selection.Tables(1).Borders.Enable = False
    
    ' Ajoutez le Nombre et un espace dans la deuxieme cellule
    Selection.Tables(1).cell(1, 1).Range.Text = CStr(Nombre)
    Selection.Tables(1).cell(1, 2).Range.Text = "   "
    
    ' Definissez la largeur de la premiere colonne pour le Nombre
    Selection.Tables(1).Columns(2).Width = InchesToPoints(1.25)
    
    ' Ajoutez la bordure laterale centrale uniquement dans la premiere colonne
    Selection.Tables(1).Columns(1).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
    
    ' Vous pouvez personnaliser le tableau et son apparence selon vos besoins
End Function

Function Faire_des_racines()
    ' Creez un tableau de 2x3
    Dim tableau As table
    Set tableau = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=2, NumColumns:=3)
    
    ' Desactivez les bordures du tableau par defaut
    tableau.Borders.Enable = False

    ' Definissez la largeur de la premiere colonne a 1 cm
    tableau.Columns(1).Width = CentimetersToPoints(1)

    ' Definissez la largeur de la deuxieme colonne a 5 cm
    tableau.Columns(2).Width = CentimetersToPoints(2.5)
    tableau.Columns(3).Width = CentimetersToPoints(2.5)

    ' Activez les bordures que vous souhaitez dans la premiere ligne
    With tableau.Rows(1)
        .Cells(1).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
        .Cells(2).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
        .Cells(2).Borders(wdBorderTop).LineStyle = wdLineStyleSingle
    End With
    With tableau.Rows(2)
        .Cells(2).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
    End With
    
    ' Demandez a l'utilisateur l'indice de la racine
    Dim indiceRacine As String
    indiceRacine = InputBox("Quel est l'indice de la racine ?")
    
    ' Placez l'indice dans la cellule 1,1 comme exposant
    With tableau.cell(1, 1).Range
        .Text = indiceRacine
        .Font.Superscript = True
        .ParagraphFormat.Alignment = wdAlignParagraphRight
    End With
    
    ' Demandez a l'utilisateur le radicande
    Dim radicande As String
    radicande = InputBox("Quel est le radicande ?")
    
    ' Verifiez si une valeur a ete fournie pour le radicande
    If radicande = "" Then
        MsgBox "Vous devez entrer une valeur pour le radicande", vbExclamation, "Erreur"
        Exit Function
    End If
    
    ' Placez le radicande dans la cellule 1,2
    tableau.cell(1, 2).Range.Text = radicande
        
    With tableau.Rows(2)
        .Cells(2).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
    End With
End Function

