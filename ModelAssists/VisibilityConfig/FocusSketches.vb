Dim doc As PartDocument = ThisDoc.Document
    Dim compDef As PartComponentDefinition = doc.ComponentDefinition

' 1. Pedir al usuario el número de sketches a seleccionar (1 a 3)
	Dim cantidadInput As String = InputBox("Enter desired sketches (1, 2 or 3):", "Quantity of sketches", "1")

' Validar la entrada del usuario
	Dim numSketches As Integer
	If Not Integer.TryParse(cantidadInput, numSketches) OrElse numSketches < 1 OrElse numSketches > 3 Then
    	MessageBox.Show("Invalid Input. Number of sketches must be between 1 and 3.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    Return
	End If

' Lista para almacenar los nombres de los sketches seleccionados
	Dim selectedNames As New List(Of String)

' 2. Bucle para seleccionar la cantidad de sketches solicitada
	For i As Integer = 1 To numSketches
    	Dim selectedSketch As Object
    	Try
        	selectedSketch = ThisApplication.CommandManager.Pick(
            	SelectionFilterEnum.kSketchObjectFilter, 
            	"Select a Sketch " & i & " de " & numSketches & ":")
    	Catch
        	' Si el usuario presiona ESC o cancela la selección
        Return 
    	End Try
    
    ' Si no se seleccionó nada, salir
    	If selectedSketch Is Nothing Then Return
    
    ' Guardar el nombre del sketch seleccionado
    	selectedNames.Add(selectedSketch.Name)
	Next

' 3. Ocultar todos los planos de trabajo (Work Planes)
	For Each wp As WorkPlane In compDef.WorkPlanes
    	wp.Visible = False
	Next

' 4. Procesar bocetos 2D (PlanarSketches)
	For Each sk2d As PlanarSketch In compDef.Sketches
    If selectedNames.Contains(sk2d.Name) Then
        sk2d.Visible = True ' Mantener visible si fue seleccionado
    Else
        sk2d.Visible = False ' Ocultar si no fue seleccionado
    End If
Next

' 5. Procesar bocetos 3D (Sketches3D)
For Each sk3d As Sketch3D In compDef.Sketches3D
    If selectedNames.Contains(sk3d.Name) Then
        sk3d.Visible = True ' Mantener visible si fue seleccionado
    Else
        sk3d.Visible = False ' Ocultar si no fue seleccionado
    End If
Next

' Actualizar la vista de Inventor
ThisApplication.ActiveView.Update()
