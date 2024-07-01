Sub Leer_Nombre()

    ' Aquí vamos a escribir
    ' Nuestro código
    
    ' Declarar las variables
    Dim nombre As String
    Dim i As Long

    ' El valor que quiero grabar debe estar en la celda C7 de la hoja activa
    ' Aquí deben leer una variable por cada valor que desean grabar
    nombre = Range("C7").Value
 
    ' Borramos el contenido de la celda
    ' Se puede borrar porque el contenido de la celda ahora esta en la variable nombre
    Range("C7").Clear

    ' Utilizando una estructura ciclica, en este caso el ciclo Do..Loop
    ' Busco la siquiente celda vacia, para este ejemplo partimos en la fila 9
    ' Ajustar al ejercicio
    i = 9
    Do

        ' Verificamos si la celda esta vacia
        ' Si esta vacia salgo del ciclo para guardar el valor
        If IsEmpty(Cells(i, 4).Value) Then
            Exit Do
        Else
            ' Vamos a la fila siguiente
            i = i + 1
        End If
    
    Loop While i < 100

    ' JUstar al ejercicio
    ' Guardamos el valor del nombre en la celda disponible
    MsgBox i
    Cells(i, 4).Value = nombre

End Sub
