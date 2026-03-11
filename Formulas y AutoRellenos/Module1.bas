Attribute VB_Name = "Module1"

Public Sub LimpiarCajas()

Hoja1.txtCantidad.Text = ""
Hoja1.txtDescripcion.Text = ""
Hoja1.txtVUnitario.Text = ""

end sub

' Funcion Para Eliminar Una Fila De La Factura

Public Sub EliminarFila()

Dim FilaActual As Long

FilaActual = ActiveCell.Row ' Fila Seleccionada Con El Cursor


If FilaActual >= 9 And FilaActua <= 33 Then

if MsgBox("¿Desea Eliminar La Fila Seleccionada?" & FilaActual, vbQuestion + vbYesNo, "Eliminar Fila") = vbYes Then

Hoja1.Cells(FilaActual, 6).Value = ""
Hoja1.Cells(FilaActual, 7).Value = ""
Hoja1.Cells(FilaActual,10).Value = ""

MsgBox "Fila Limpiada Exitosamente", vbInformation, "Eliminar Fila"

end if

Else
    
    MsgBox "Seleccione Una Fila Dentro Del Rango De La Factura", vbExclamation, "Error"

End If

end Sub