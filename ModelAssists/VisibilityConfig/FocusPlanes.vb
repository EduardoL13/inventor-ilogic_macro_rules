Dim doc As PartDocument = ThisDoc.Document
Dim compDef As PartComponentDefinition = doc.ComponentDefinition

' 1. Pedir al usuario el número de planos a seleccionar (1 a 3)
Dim cantidadInput As String = InputBox("Ingresa el número de planos a seleccionar (1, 2 o 3):", "Cantidad de Planos", "1")

' Validar la entrada del usuario
Dim numPlanes As Integer
If Not Integer.TryParse(cantidadInput, numPlanes) OrElse numPlanes < 1 OrElse numPlanes > 3 Then
    MessageBox.Show("Entrada inválida. Debes ingresar un número entre 1 y 3.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    Return
End If

' Lista para almacenar los nombres de los planos seleccionados
Dim selectedPlanesNames As New List(Of String)

' 2. Bucle para seleccionar la cantidad de planos solicitada
For i As Integer = 1 To numPlanes
    Dim selectedPlane As Object
    Try
        ' Filtro específico para WorkPlanes (Planos de trabajo y de origen)
        selectedPlane = ThisApplication.CommandManager.Pick(
            SelectionFilterEnum.kWorkPlaneFilter, 
            "Selecciona el Plano " & i & " de " & numPlanes & ":")
    Catch
        ' Si el usuario presiona ESC o cancela la selección
        Return 
    End Try
    
    ' Si no se seleccionó nada, salir
    If selectedPlane Is Nothing Then Return
    
    ' Guardar el nombre del plano seleccionado
    selectedPlanesNames.Add(selectedPlane.Name)
Next

' 3. Ocultar todos los planos de trabajo, excepto los seleccionados
For Each wp As WorkPlane In compDef.WorkPlanes
    If selectedPlanesNames.Contains(wp.Name) Then
        wp.Visible = True ' Mantener visible si fue seleccionado
    Else
        wp.Visible = False ' Ocultar si no fue seleccionado
    End If
Next

' NOTA: Los bloques de código para sketches se eliminaron 
' para asegurar que NO se altere la visibilidad de ningún boceto.

' Actualizar la vista de Inventor
ThisApplication.ActiveView.Update()
