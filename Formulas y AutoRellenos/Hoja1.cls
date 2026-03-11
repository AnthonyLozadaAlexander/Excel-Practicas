Private Sub btnIngresar_Click()

    Dim filaVacia As Long
    Dim i As Long
    Dim hayEspacio As Boolean

    hayEspacio = false

    For i = 9 to 33

    If Hoja1.Cells(i, 6).Value = "" Then

    filaVacia = i
    hayEspacio = true

    Exit For

    End If
    
    Next i

    If txtCantidad.Text = "" Or txtDescripcion.Text = "" Or txtVUnitario.Text = "" Then
    
    MsgBox "Porfavor Llene Los Campos", VbExclamation, "Error: Campos Vacio"

    end If

    If hayEspacio = True Then

      Hoja1.Cells(filaVacia, 6).Value = txtCantidad.Text
      Hoja1.Cells(filaVacia, 7).Value = txtDescripcion.Text
      Hoja1.Cells(filaVacia, 10).Value = txtVUnitario.Text

      Call LimpiarCajas

    Else

     MsgBox "La Tabla Se Encuentra LLena"

    End If

End Sub

Private Sub RemoverFila_Click()

Call EliminarFila

End Sub
