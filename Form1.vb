Imports System.Windows.Forms
Imports System.Drawing
Imports System.Text.RegularExpressions
Imports System.IO

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
                ' Configuración básica del formulario
                form.Text = "Escalar Patrón - Rango de Tallas"
                form.Size = New Size(400, 350)
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.StartPosition = FormStartPosition.CenterScreen
                form.MaximizeBox = False
                form.MinimizeBox = False

                ' Cargar y configurar la imagen de fondo usando el mismo método que funciona
                Dim formBackgroundImage2 As Image = Nothing
                Try
                    ' Usar la ruta directa desde el directorio del proyecto
                    formBackgroundImage2 = Image.FromFile("C:\Users\Willian\Desktop\ESCALADO_PROFESIONAL_WR\www.root\Images\LOGO.png")
                    form.BackgroundImage = formBackgroundImage2
                    form.BackgroundImageLayout = ImageLayout.Stretch

                    ' Agregar el mismo evento de pintura que funciona en la ventana principal
                    AddHandler form.Paint, Sub(sender As Object, e As PaintEventArgs)
                                               If formBackgroundImage2 IsNot Nothing Then
                                                   e.Graphics.DrawImage(formBackgroundImage2, 0, 0, form.Width, form.Height)
                                                   Using brush As New SolidBrush(Color.FromArgb(100, Color.White))
                                                       e.Graphics.FillRectangle(brush, 0, 0, form.Width, form.Height)
                                                   End Using
                                               End If
                                           End Sub
                Catch ex As Exception
                    MessageBox.Show("No se pudo cargar la imagen de fondo: " & ex.Message)
                End Try

                ' Mostrar talla base con estilo
                Dim lblBase As New Label()
                lblBase.Text = $"Talla Base Detectada: {tallaBase}"
                lblBase.Location = New Point(20, 20)
                lblBase.Size = New Size(360, 30)
                lblBase.Font = New Font("Arial", 12, FontStyle.Bold)
                lblBase.TextAlign = ContentAlignment.MiddleCenter
                lblBase.BackColor = Color.FromArgb(200, 0, 120, 215)
                lblBase.ForeColor = Color.White

                ' Input para talla menor con estilo
                Dim lblMenor As New Label()
                lblMenor.Text = "Talla Menor:"
                lblMenor.Location = New Point(20, 70)
                lblMenor.Size = New Size(150, 25)
                lblMenor.Font = New Font("Arial", 10, FontStyle.Bold)
                lblMenor.BackColor = Color.FromArgb(200, 255, 255, 255)
                lblMenor.TextAlign = ContentAlignment.MiddleLeft

                Dim txtMenor As New TextBox()
                txtMenor.Location = New Point(180, 70)
                txtMenor.Size = New Size(100, 25)
                txtMenor.Font = New Font("Arial", 12, FontStyle.Regular)
                txtMenor.TextAlign = HorizontalAlignment.Center

                ' Input para talla mayor con estilo
                Dim lblMayor As New Label()
                lblMayor.Text = "Talla Mayor:"
                lblMayor.Location = New Point(20, 110)
                lblMayor.Size = New Size(150, 25)
                lblMayor.Font = New Font("Arial", 10, FontStyle.Bold)
                lblMayor.BackColor = Color.FromArgb(200, 255, 255, 255)
                lblMayor.TextAlign = ContentAlignment.MiddleLeft

                Dim txtMayor As New TextBox()
                txtMayor.Location = New Point(180, 110)
                txtMayor.Size = New Size(100, 25)
                txtMayor.Font = New Font("Arial", 12, FontStyle.Regular)
                txtMayor.TextAlign = HorizontalAlignment.Center

                ' Checkbox para plantillas
                Dim chkPlantilla As New CheckBox()
                chkPlantilla.Text = "Escalar Plantilla"
                chkPlantilla.Location = New Point(20, 150)
                chkPlantilla.Size = New Size(200, 25)
                chkPlantilla.Font = New Font("Arial", 10, FontStyle.Bold)
                chkPlantilla.BackColor = Color.FromArgb(200, 255, 255, 255)

                ' Botón OK con estilo
                Dim btnOK As New Button()
                btnOK.Text = "Generar Tallas"
                btnOK.DialogResult = DialogResult.OK
                btnOK.Location = New Point(100, 190)
                btnOK.Size = New Size(200, 40)
                btnOK.Font = New Font("Arial", 12, FontStyle.Bold)
                btnOK.BackColor = Color.FromArgb(0, 120, 215)
                btnOK.ForeColor = Color.White
                btnOK.FlatStyle = FlatStyle.Flat
                btnOK.Cursor = Cursors.Hand

                ' Agregar controles directamente al formulario
                form.Controls.AddRange({lblBase, lblMenor, txtMenor, lblMayor, txtMayor, chkPlantilla, btnOK})

                ' Asegurarse de liberar la imagen cuando se cierre el formulario
                AddHandler form.FormClosing, Sub(sender As Object, e As FormClosingEventArgs)
                                                 If formBackgroundImage2 IsNot Nothing Then
                                                     formBackgroundImage2.Dispose()
                                                 End If
                                             End Sub

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

                        ' Obtener dimensiones del cuadro escaneado base
                        Dim baseBox As Object = baseShape.BoundingBox
                        Dim baseHeight As Double = baseBox.Height
                        Dim baseWidth As Double = baseBox.Width

                        ' Generar cada talla
                        For talla As Integer = tallaMenor To tallaMayor
                            If talla <> tallaBase Then
                                ' Calcular incrementos según sistema francés
                                Dim diferenciaTallas As Integer = talla - tallaBase

                                ' Duplicar el cuadro completo
                                Dim newShape As Object = baseShape.Duplicate()

                                If chkPlantilla.Checked Then
                                    ' Para plantillas: igual que moldes pero para el cuadro escaneado
                                    Dim factorLargo As Double = 1 + ((3.33 * diferenciaTallas) / 100)  ' 6.67mm/2 para el cuadro
                                    Dim factorAncho As Double = 1 + ((1.25 * diferenciaTallas) / 100)  ' 2.5mm/2 para el cuadro
                                    newShape.Stretch(factorAncho, factorLargo)
                                Else
                                    ' Para moldes: sistema francés reducido para cuadro completo
                                    Dim factorLargo As Double = 1 + ((3.33 * diferenciaTallas) / 100)  ' 6.67mm/2
                                    Dim factorAncho As Double = 1 + ((1.25 * diferenciaTallas) / 100)  ' 2.5mm/2
                                    newShape.Stretch(factorAncho, factorLargo)
                                End If

                                ' Actualizar números en el nuevo molde
                                ActualizarNumerosEnForma(newShape, tallaBase, talla)

                                ' Posicionar
                                desplazamiento += newShape.SizeWidth + 25
                                newShape.Move(desplazamiento, 0)
                            End If
                        Next

                        corelApp.Optimization = False
                        corelApp.Refresh()
                        MessageBox.Show($"Patrones escalados correctamente:{Environment.NewLine}" &
                                      $"Talla base: {tallaBase}{Environment.NewLine}" &
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
            MessageBox.Show($"Error: {ex.Message}{Environment.NewLine}" &
                          "Verifique que el molde base contenga números visibles",
                          "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class

Public Class Form1
    Inherits Form

    Private WithEvents btnEjecutar As New Button()
    Private WithEvents btnAdmin As New Button()
    Private formBackgroundImage As Image
    Private currentLicense As LicenseManager.LicenseInfo

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "Escalado Profesional"
        Me.Size = New Size(800, 650)
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.MaximizeBox = True
        Me.MinimizeBox = True

        ' Configurar el fondo
        Try
            formBackgroundImage = Image.FromFile("C:\Users\Willian\Desktop\ESCALADO_PROFESIONAL_WR\www.root\Images\Logo2.png")
            Me.BackgroundImage = formBackgroundImage
            Me.BackgroundImageLayout = ImageLayout.Stretch
        Catch ex As Exception
            MessageBox.Show("No se pudo cargar la imagen de fondo: " & ex.Message)
        End Try

        ' Panel para centrar los botones
        Dim buttonPanel As New Panel()
        buttonPanel.Dock = DockStyle.Fill
        buttonPanel.AutoSize = False
        buttonPanel.BackColor = Color.Transparent

        ' Botón Ejecutar
        btnEjecutar.Text = "Ejecutar Escalado"
        btnEjecutar.Size = New Size(200, 40)
        btnEjecutar.BackColor = Color.FromArgb(0, 120, 215)
        btnEjecutar.ForeColor = Color.White
        btnEjecutar.FlatStyle = FlatStyle.Flat
        btnEjecutar.Anchor = AnchorStyles.None
        btnEjecutar.Location = New Point((buttonPanel.Width - btnEjecutar.Width) \ 2, (buttonPanel.Height - 100) \ 2)

        ' Botón Admin
        btnAdmin.Text = "Administrar Licencia"
        btnAdmin.Size = New Size(200, 40)
        btnAdmin.BackColor = Color.FromArgb(0, 120, 215)
        btnAdmin.ForeColor = Color.White
        btnAdmin.FlatStyle = FlatStyle.Flat
        btnAdmin.Anchor = AnchorStyles.None
        btnAdmin.Location = New Point(btnEjecutar.Left, btnEjecutar.Bottom + 10)

        ' Agregar los botones al panel
        buttonPanel.Controls.AddRange({btnEjecutar, btnAdmin})

        ' Pie de página
        Dim lblFooter As New Label()
        lblFooter.Text = "ATENCIÓN AL CLIENTE" & vbCrLf & _
                        "CORREO: wilyd2@hotmail.com" & vbCrLf & _
                        "TEL: +573147743846 - 3178625955"
        lblFooter.Dock = DockStyle.Bottom
        lblFooter.Height = 80
        lblFooter.TextAlign = ContentAlignment.MiddleCenter
        lblFooter.Font = New Font("Arial", 12, FontStyle.Bold)
        lblFooter.BackColor = Color.FromArgb(0, 120, 215)
        lblFooter.ForeColor = Color.White

        ' Agregar los controles al formulario
        Me.Controls.AddRange({buttonPanel, lblFooter})

        ' Manejar el evento Resize para mantener los botones centrados
        AddHandler Me.Resize, Sub()
            btnEjecutar.Location = New Point((buttonPanel.Width - btnEjecutar.Width) \ 2, (buttonPanel.Height - 100) \ 2)
            btnAdmin.Location = New Point(btnEjecutar.Left, btnEjecutar.Bottom + 10)
        End Sub
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            ' Verificar licencia al iniciar
            currentLicense = LicenseManager.ValidateLicense()
            If Not currentLicense.IsValid Then
                If MessageBox.Show("No hay una licencia válida instalada. ¿Desea generar una nueva licencia ahora?",
                                 "Licencia Requerida",
                                 MessageBoxButtons.YesNo,
                                 MessageBoxIcon.Question) = DialogResult.Yes Then
                    currentLicense = LicenseManager.ShowLicenseDialog()
                    If currentLicense Is Nothing OrElse Not currentLicense.IsValid Then
                        MessageBox.Show("No se pudo generar una licencia válida. La aplicación se cerrará.",
                                      "Error de Licencia",
                                      MessageBoxButtons.OK,
                                      MessageBoxIcon.Error)
                        Me.Close()
                    End If
                Else
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show($"Error al iniciar la aplicación: {ex.Message}",
                          "Error",
                          MessageBoxButtons.OK,
                          MessageBoxIcon.Error)
            Me.Close()
        End Try
    End Sub

    Private Sub btnEjecutar_Click(sender As Object, e As EventArgs) Handles btnEjecutar.Click
        Try
            If Not currentLicense.IsValid Then
                MessageBox.Show("La licencia no es válida o ha expirado. Por favor use el botón de Administración para generar una nueva licencia.",
                              "Error de Licencia", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Dim macro As New ThisMacroStorage_EscalarCalzado()
            macro.EscalarPatrones()
        Catch ex As Exception
            MessageBox.Show($"Error al ejecutar el escalado: {ex.Message}",
                          "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnAdmin_Click(sender As Object, e As EventArgs) Handles btnAdmin.Click
        Try
            Dim message As String

            If currentLicense IsNot Nothing Then
                message = $"Estado de la licencia actual:{Environment.NewLine}" &
                         $"Válida hasta: {currentLicense.ExpirationDate:dd/MM/yyyy}{Environment.NewLine}" &
                         $"Tipo: {GetLicenseTypeText(currentLicense.LicenseType)}{Environment.NewLine}" &
                         $"Estado: {If(currentLicense.IsValid, "Activa", "Expirada")}"
            Else
                message = "No hay licencia instalada."
            End If

            If MessageBox.Show($"{message}{Environment.NewLine}{Environment.NewLine}¿Desea generar una nueva licencia?",
                             "Administración de Licencia",
                             MessageBoxButtons.YesNo,
                             MessageBoxIcon.Question) = DialogResult.Yes Then

                Dim newLicense = LicenseManager.ShowLicenseDialog()
                If newLicense IsNot Nothing Then
                    currentLicense = newLicense
                    MessageBox.Show($"Nueva licencia generada correctamente:{Environment.NewLine}" &
                                  $"Válida hasta: {newLicense.ExpirationDate:dd/MM/yyyy}{Environment.NewLine}" &
                                  $"Tipo: {GetLicenseTypeText(newLicense.LicenseType)}",
                                  "Licencia Generada",
                                  MessageBoxButtons.OK,
                                  MessageBoxIcon.Information)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show($"Error al gestionar la licencia: {ex.Message}",
                          "Error",
                          MessageBoxButtons.OK,
                          MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function GetLicenseTypeText(licenseType As LicenseManager.LicenseTypes) As String
        Select Case licenseType
            Case LicenseManager.LicenseTypes.Trial
                Return "Prueba (1 semana)"
            Case LicenseManager.LicenseTypes.SixMonths
                Return "Semestral (6 meses)"
            Case LicenseManager.LicenseTypes.OneYear
                Return "Anual (1 año)"
            Case Else
                Return "Desconocido"
        End Select
    End Function

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If formBackgroundImage IsNot Nothing Then
            formBackgroundImage.Dispose()
        End If
        MyBase.OnFormClosing(e)
    End Sub
End Class
