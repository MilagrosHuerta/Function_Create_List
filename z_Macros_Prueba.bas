Attribute VB_Name = "z_Macros_Prueba"
Sub Prueba_Lista()
Dim DatoLista As Variant

    DatoLista = "NUM_1_7"
    Lista_Combo = Crear_Lista(DatoLista, "_")
    frmComboBox "Has seleccionado que quieres crear una lista con los siguientes parámetros:" & vbNewLine & vbNewLine & DatoLista, _
                "TITULO", Lista_Combo, "_", "AGATA"
    
    DatoLista = InputBox("Introduce los datos con los que quieres crear la lista")
    Separador_Dato = InputBox("Introduce el separador que vas a usar en la lista")
    Lista_Combo = Crear_Lista(DatoLista, Separador_Dato)
    frmComboBox "Has seleccionado que quieres crear una lista con los siguientes parámetros:" & vbNewLine & vbNewLine & DatoLista, _
                "TITULO", Lista_Combo, Separador_Dato, "AGATA"
    
    DatoLista = "CAR_A,F"
    Lista_Combo = Crear_Lista(DatoLista, ",")
    frmComboBox "Has seleccionado que quieres crear una lista con los siguientes parámetros:" & vbNewLine & vbNewLine & DatoLista, _
                "TITULO", Lista_Combo, ",", "AGATA"
    
    DatoLista = InputBox("Introduce los datos con los que quieres crear la lista")
    Separador_Dato = InputBox("Introduce el separador que vas a usar en la lista")
    Lista_Combo = Crear_Lista(DatoLista, Separador_Dato)
    frmComboBox "Has seleccionado que quieres crear una lista con los siguientes parámetros:" & vbNewLine & vbNewLine & DatoLista, _
                "TITULO", Lista_Combo, Separador_Dato, "AGATA"
    
    DatoLista = "CEL_A:19,22"
    Lista_Combo = Crear_Lista(DatoLista, ",")
    frmComboBox "Has seleccionado que quieres crear una lista con los siguientes parámetros:" & vbNewLine & vbNewLine & DatoLista, _
                "TITULO", Lista_Combo, ",", "AGATA"
    
    DatoLista = "CEL_A:19-"
    Lista_Combo = Crear_Lista(DatoLista, "-")
    frmComboBox "Has seleccionado que quieres crear una lista con los siguientes parámetros:" & vbNewLine & vbNewLine & DatoLista, _
                "TITULO", Lista_Combo, "-", "AGATA"

    DatoLista = InputBox("Introduce los datos con los que quieres crear la lista")
    Separador_Dato = InputBox("Introduce el separador que vas a usar en la lista")
    Lista_Combo = Crear_Lista(DatoLista, Separador_Dato)
    frmComboBox "Has seleccionado que quieres crear una lista con los siguientes parámetros:" & vbNewLine & vbNewLine & DatoLista, _
                "TITULO", Lista_Combo, Separador_Dato, "AGATA"


End Sub

