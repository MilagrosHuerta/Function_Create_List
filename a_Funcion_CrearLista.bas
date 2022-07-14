Attribute VB_Name = "a_Funcion_CrearLista"
' ------------------------------------------------------------ '
' ---                Funcion creada por                    --- '
' ---         MILAGROS HUERTA GÓMEZ DE MERODIO             --- '
' ------------------------------------------------------------ '
' ---                    Crear_Lista                       --- '
' ------------------------------------------------------------ '
' ---    Puedes usarla libremente en tus aplicaciones,     --- '
' ---    pero no asignarte la autoría.                     --- '
' ---    Sirve para enviar mensajes con otro formato       --- '
' ---    y poder posicionarlo donde quieras                --- '
' ------------------------------------------------------------ '
Function Crear_Lista(datoEntrada As Variant, Separador As String)
Dim Cadena As Variant
Dim Tipo As String
Dim Dato_Inicial As Variant
Dim Dato_Final As Variant
Dim NuevaLista As Variant
Dim i As Integer
Dim Columna As String
Dim Fila_Inicial As Integer
Dim Fila_Final As Long
' ------------------------------------------------------------ '
' --- Crea una lista utomáticamente si el Dato de Estrada  --- '
' --- empieza por NUM_ CAR_ o CEL_                         --- '
' ------------------------------------------------------------ '
' ---   NUM_ => Para crear lista de números                --- '
' ---           Para números positivos                     --- '
' ---           ej:  NUM_1-10 y separador "-"              --- '
' ---   CAR_ => Para crear lista de caracteres de 1 dígito --- '
' ---           Diferencia mayúsculas de minúsculas        --- '
' ---           ej:  CAR_a_e y separador "_"               --- '
' ---   CEL_ => Para crear lista de un rango de celdas     --- '
' ---           Columna de la celda: fila inicial y fila   --- '
' ---           Final. Si se omite la fila final, toma los --- '
' ---           datos de las celdas que no estén vacías    --- '
' ---           ej:  CEL_A:10,14 y separador ","           --- '
' ------------------------------------------------------------ '

    Tipo = UCase(Left(datoEntrada, 4))
    Cadena = Mid(datoEntrada, 5, Len(datoEntrada))
    Dato_Inicial = Left(Cadena, InStr(Cadena, Separador) - 1)
    Dato_Final = Mid(Cadena, InStr(Cadena, Separador) + 1, Len(Cadena))
    Select Case Tipo
        Case "NUM_"     ' Lista de números
            NuevaLista = Dato_Inicial
            For i = Dato_Inicial * 1 + 1 To Dato_Final
                NuevaLista = NuevaLista & Separador & i
            Next i
        Case "CAR_"     'Lista de caracteres
            NuevaLista = Dato_Inicial
            If Len(Dato_Inicial) = 1 And Len(Dato_Final) = 1 Then
                For i = Asc(Dato_Inicial) + 1 To Asc(Dato_Final)
                    NuevaLista = NuevaLista & Separador & Chr(i)
                Next i
            Else        ' Para cuando la lista sea de más de un caracter, PENDIENTE DE PREPARAR
                ''codigo ascii  65 = A ( Letra A mayúscula )
                ''codigo ascii  90 = Z ( Letra Z mayúscula )
                
                ''codigo ascii  97 = a ( Letra a minúscula )
                ''codigo ascii 122 = z ( Letra z minúscula )
            End If
        Case "CEL_"     ' La lista está en un rango de celdas
            Columna = Left(Dato_Inicial, InStr(Dato_Inicial, ":") - 1)
            Fila_Inicial = Mid(Dato_Inicial, InStr(Dato_Inicial, ":") + 1, Len(Dato_Inicial)) * 1
            i = Fila_Inicial + 1
            NuevaLista = Range(Columna & Fila_Inicial)
            
            If Right(datoEntrada, 1) >= 0 And Right(datoEntrada, 1) <= 9 Then
                Fila_Final = Mid(datoEntrada, InStr(datoEntrada, Separador) + 1, Len(datoEntrada)) * 1
                For i = Fila_Inicial + 1 To Fila_Final
                    NuevaLista = NuevaLista & Separador & Range(Columna & i)
                Next i
            Else
                Do While Range(Columna & i) <> ""
                    NuevaLista = NuevaLista & Separador & Range(Columna & i)
                    i = i + 1
                Loop
            End If
    End Select
    Crear_Lista = NuevaLista
End Function
