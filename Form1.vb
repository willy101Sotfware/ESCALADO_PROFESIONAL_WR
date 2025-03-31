Imports System.Windows.Forms
Imports System.Drawing
Imports System.Text.RegularExpressions

Public Class ThisMacroStorage_EscalarCalzado
    ' Constantes de CorelDRAW
    Private Const cdrTextShape As Integer = 6
    Private Const cdrGroupShape As Integer = 7
    Private Const cdrLayerShape As Integer = 0
    Private Const cdrRectangleShape As Integer = 2
    Private Const cdrEllipseShape As Integer = 3
    Private Const cdrCurveShape As Integer = 4
    Private Const cdrArrowShape As Integer = 5
    Private Const cdrPolygonShape As Integer = 9
    Private Const cdrBitmapShape As Integer = 10

    Private corelApp As Object

    Public Sub New()
        Try
            corelApp = GetObject(, "CorelDRAW.Application")
        Catch
            corelApp = CreateObject("CorelDRAW.Application")
        End Try
        corelApp.Visible = True
        
        ' Asegurarse de que haya un documento activo
        If corelApp.Documents.Count = 0 Then
            corelApp.CreateDocument()
        End If
    End Sub

    Private Function ExtraerNumeros(texto As String) As String
        Dim resultado As String = ""
        For i As Integer = 1 To Len(texto)
            If IsNumeric(Mid(texto, i, 1)) Then
                resultado &= Mid(texto, i, 1)
            End If
        Next
        Return resultado
    End Function

    Private Function ContieneNumero(texto As String, numero As Integer) As Boolean
        Return InStr(1, texto, CStr(numero)) > 0
    End Function

    Private Function ObtenerTallaDesdeForma(shape As Object) As Integer
        Try
            ' Verificar si es forma de texto
            If shape.Type = cdrTextShape Then
                Dim texto As String = shape.Text.Story.Text
                Dim numeros As String = ExtraerNumeros(texto)
                If Not String.IsNullOrEmpty(numeros) Then
                    Return CInt(numeros)
                End If
            End If

            ' Buscar recursivamente en sub-formas
            For Each subShape As Object In shape.Shapes
                If subShape.Type = cdrTextShape Then
                    Dim texto As String = subShape.Text.Story.Text
                    Dim numeros As String = ExtraerNumeros(texto)
                    If Not String.IsNullOrEmpty(numeros) Then
                        Return CInt(numeros)
                    End If
                ElseIf subShape.Shapes.Count > 0 Then
                    Dim talla As Integer = ObtenerTallaDesdeForma(subShape)
                    If talla > 0 Then Return talla
                End If
            Next
        Catch ex As Exception
            ' Ignorar errores y continuar buscando
        End Try
        Return 0
    End Function

    Private Sub ActualizarNumerosEnForma(shape As Object, tallaVieja As Integer, tallaNueva As Integer)
        Try
            ' Actualizar texto en forma principal
            If shape.Type = cdrTextShape Then
                Dim texto As String = shape.Text.Story.Text
                If ContieneNumero(texto, tallaVieja) Then
                    shape.Text.Story.Text = Replace(texto, CStr(tallaVieja), CStr(tallaNueva))
                End If
            End If

            ' Actualizar en sub-formas
            For Each subShape As Object In shape.Shapes
                If subShape.Type = cdrTextShape Then
                    Dim texto As String = subShape.Text.Story.Text
                    If ContieneNumero(texto, tallaVieja) Then
                        subShape.Text.Story.Text = Replace(texto, CStr(tallaVieja), CStr(tallaNueva))
                    End If
                ElseIf subShape.Shapes.Count > 0 Then
                    ActualizarNumerosEnForma(subShape, tallaVieja, tallaNueva)
                End If
            Next
        Catch ex As Exception
            ' Ignorar errores y continuar actualizando
        End Try
    End Sub

    Private Function InputTallaManual(mensaje As String) As Integer
        Dim inputValor As String = InputBox(mensaje, "Talla", "")
        
        If String.IsNullOrEmpty(inputValor) Then
            Return 0
        End If
        
        If Not IsNumeric(inputValor) Then
            MessageBox.Show("Debe ingresar un número válido", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return 0
        End If
        
        Return CInt(inputValor)
    End Function

    Public Sub EscalarPatrones()
        Try
            ' Verificar selección
            If corelApp.ActiveDocument.Selection.Shapes.Count <> 1 Then
                MessageBox.Show("Seleccione EXACTAMENTE UN grupo o molde base", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim baseShape As Object = corelApp.ActiveDocument.Selection.Shapes(1)
            
            ' Detectar talla base
            Dim tallaBase As Integer = ObtenerTallaDesdeForma(baseShape)
            
            ' Si no se detectó talla, pedirla manualmente
            If tallaBase = 0 Then
                tallaBase = InputTallaManual("No se detectó talla. Ingrese la talla BASE del molde:")
                If tallaBase = 0 Then Exit Sub
            End If

            ' Formulario para rango de tallas
            Using form As New Form()
                form.Text = "Escalar Patrón - Rango de Tallas"
                form.Size = New Size(300, 250)
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.StartPosition = FormStartPosition.CenterScreen
                form.MaximizeBox = False
                form.MinimizeBox = False

                ' Mostrar talla base
                Dim lblBase As New Label()
                lblBase.Text = $"Talla Base Detectada: {tallaBase}"
                lblBase.Location = New Point(20, 20)
                lblBase.AutoSize = True

                ' Input para talla menor
                Dim lblMenor As New Label()
                lblMenor.Text = "Talla Menor:"
                lblMenor.Location = New Point(20, 60)
                lblMenor.AutoSize = True

                Dim txtMenor As New TextBox()
                txtMenor.Location = New Point(20, 80)
                txtMenor.Size = New Size(100, 20)

                ' Input para talla mayor
                Dim lblMayor As New Label()
                lblMayor.Text = "Talla Mayor:"
                lblMayor.Location = New Point(20, 120)
                lblMayor.AutoSize = True

                Dim txtMayor As New TextBox()
                txtMayor.Location = New Point(20, 140)
                txtMayor.Size = New Size(100, 20)

                ' Botón OK
                Dim btnOK As New Button()
                btnOK.Text = "Generar Tallas"
                btnOK.DialogResult = DialogResult.OK
                btnOK.Location = New Point(20, 180)

                ' Agregar controles
                form.Controls.AddRange({lblBase, lblMenor, txtMenor, lblMayor, txtMayor, btnOK})

                If form.ShowDialog() = DialogResult.OK Then
                    Dim tallaMenor As Integer
                    Dim tallaMayor As Integer

                    If Integer.TryParse(txtMenor.Text, tallaMenor) AndAlso 
                       Integer.TryParse(txtMayor.Text, tallaMayor) Then
                        
                        If tallaMenor > tallaMayor Then
                            MessageBox.Show("ERROR: La talla menor no puede ser mayor que la talla mayor", "Error", 
                                          MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        End If

                        ' Iniciar proceso de escalado
                        corelApp.Optimization = True
                        Dim desplazamiento As Double = 0

                        ' Generar cada talla
                        For talla As Integer = tallaMenor To tallaMayor
                            If talla <> tallaBase Then
                                ' Calcular incrementos según sistema francés
                                Dim diferenciaTallas As Integer = talla - tallaBase
                                ' 6.667mm por talla en largo (2/3 cm)
                                Dim factorLargo As Double = 1 + ((6.667 * diferenciaTallas) / 100)
                                ' 2.5mm por talla en ancho (1/4 cm)
                                Dim factorAncho As Double = 1 + ((2.5 * diferenciaTallas) / 100)

                                ' Duplicar y escalar con factores diferentes para largo y ancho
                                Dim newShape As Object = baseShape.Duplicate()
                                newShape.Stretch(factorLargo, factorAncho)

                                ' Actualizar números en el nuevo molde
                                ActualizarNumerosEnForma(newShape, tallaBase, talla)

                                ' Posicionar
                                desplazamiento += newShape.SizeWidth + 25
                                newShape.Move(desplazamiento, 0)
                            End If
                        Next

                        corelApp.Optimization = False
                        corelApp.Refresh()
                        MessageBox.Show($"Patrones escalados correctamente:{Environment.NewLine}" & _
                                      $"Talla base: {tallaBase}{Environment.NewLine}" & _
                                      $"Tallas generadas: {tallaMenor} a {tallaMayor}", 
                                      "Proceso Completado", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("Por favor ingrese tallas válidas", "Error", 
                                      MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End If
            End Using

        Catch ex As Exception
            corelApp.Optimization = False
            MessageBox.Show($"Error: {ex.Message}{Environment.NewLine}" & _
                          "Verifique que el molde base contenga números visibles", 
                          "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class

Public Class Form1
    Inherits Form

    Private WithEvents btnEjecutar As New Button()
    Private macroStorage As ThisMacroStorage_EscalarCalzado
    Private formBackgroundImage As Image

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        ' Configurar el botón
        btnEjecutar.Text = "Escalar Patrón"
        btnEjecutar.Location = New Point(100, 80)
        btnEjecutar.Size = New Size(200, 50)
        btnEjecutar.Font = New Font("Arial", 12, FontStyle.Bold)
        btnEjecutar.BackColor = Color.FromArgb(240, 240, 240)
        btnEjecutar.FlatStyle = FlatStyle.Flat

        ' Configurar el formulario
        Me.Text = "Escalador de Patrones Profesional"
        Me.Size = New Size(400, 250)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Controls.Add(btnEjecutar)

        ' Configurar la imagen de fondo
        Try
            formBackgroundImage = Image.FromFile(System.IO.Path.Combine(Application.StartupPath, "www.root", "Images", "LOGO.png"))
            Me.BackgroundImage = formBackgroundImage
            Me.BackgroundImageLayout = ImageLayout.Stretch
        Catch ex As Exception
            MessageBox.Show("No se pudo cargar la imagen de fondo: " & ex.Message)
        End Try
    End Sub

    Protected Overrides Sub OnPaintBackground(e As PaintEventArgs)
        If formBackgroundImage IsNot Nothing Then
            Using brush As New SolidBrush(Color.FromArgb(200, Color.White))
                e.Graphics.DrawImage(formBackgroundImage, 0, 0, Me.Width, Me.Height)
                e.Graphics.FillRectangle(brush, 0, 0, Me.Width, Me.Height)
            End Using
        Else
            MyBase.OnPaintBackground(e)
        End If
    End Sub

    Private Sub btnEjecutar_Click(sender As Object, e As EventArgs) Handles btnEjecutar.Click
        Try
            If macroStorage Is Nothing Then
                macroStorage = New ThisMacroStorage_EscalarCalzado()
            End If
            macroStorage.EscalarPatrones()
        Catch ex As Exception
            MessageBox.Show("Error al ejecutar la macro: " & vbCrLf & ex.Message, "Error", 
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If formBackgroundImage IsNot Nothing Then
            formBackgroundImage.Dispose()
        End If
        MyBase.OnFormClosing(e)
    End Sub
End Class 