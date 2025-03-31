Imports System.Windows.Forms
Imports System.Drawing

Public Class AdminLicenseForm
    Inherits Form

    Private txtPassword As TextBox
    Private btnGenerate As Button
    Private lblStatus As Label
    Private lblExpiration As Label
    Private lblDaysLeft As Label
    Private lblLicenseKey As Label

    Public Sub New()
        InitializeComponent()
        LoadLicenseInfo()
    End Sub

    Private Sub InitializeComponent()
        Me.Size = New Size(400, 300)
        Me.Text = "Administración de Licencia"
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        ' Password Label
        Dim lblPassword As New Label()
        lblPassword.Text = "Contraseña:"
        lblPassword.Location = New Point(20, 20)
        lblPassword.Size = New Size(100, 20)

        ' Password TextBox
        txtPassword = New TextBox()
        txtPassword.Location = New Point(120, 20)
        txtPassword.Size = New Size(200, 20)
        txtPassword.UseSystemPasswordChar = True

        ' Generate Button
        btnGenerate = New Button()
        btnGenerate.Text = "Generar Nueva Licencia"
        btnGenerate.Location = New Point(120, 60)
        btnGenerate.Size = New Size(200, 30)
        btnGenerate.Enabled = False
        AddHandler btnGenerate.Click, AddressOf GenerateNewLicense

        ' Status Label
        lblStatus = New Label()
        lblStatus.Location = New Point(20, 100)
        lblStatus.Size = New Size(360, 25)
        lblStatus.TextAlign = ContentAlignment.MiddleCenter

        ' License Key Label
        lblLicenseKey = New Label()
        lblLicenseKey.Location = New Point(20, 130)
        lblLicenseKey.Size = New Size(360, 25)
        lblLicenseKey.TextAlign = ContentAlignment.MiddleCenter
        lblLicenseKey.Font = New Font(lblLicenseKey.Font, FontStyle.Bold)

        ' Expiration Label
        lblExpiration = New Label()
        lblExpiration.Location = New Point(20, 160)
        lblExpiration.Size = New Size(360, 25)
        lblExpiration.TextAlign = ContentAlignment.MiddleCenter

        ' Days Left Label
        lblDaysLeft = New Label()
        lblDaysLeft.Location = New Point(20, 190)
        lblDaysLeft.Size = New Size(360, 25)
        lblDaysLeft.TextAlign = ContentAlignment.MiddleCenter
        lblDaysLeft.Font = New Font(lblDaysLeft.Font, FontStyle.Bold)

        ' Add controls
        Me.Controls.AddRange({lblPassword, txtPassword, btnGenerate, lblStatus, lblLicenseKey, lblExpiration, lblDaysLeft})

        ' Add password validation
        AddHandler txtPassword.TextChanged, AddressOf ValidatePassword
    End Sub

    Private Sub LoadLicenseInfo()
        Dim license = LicenseManager.ValidateLicense()
        If license.IsValid Then
            lblStatus.Text = "Estado: Licencia Válida"
            lblStatus.ForeColor = Color.Green
            lblLicenseKey.Text = $"Clave de Licencia: {license.LicenseKey}"
            lblExpiration.Text = $"Fecha de Expiración: {license.ExpirationDate:dd/MM/yyyy}"
            
            Dim diasRestantes As Integer = (license.ExpirationDate - Date.Now).Days
            lblDaysLeft.Text = $"Días Restantes: {diasRestantes}"
            If diasRestantes <= 30 Then
                lblDaysLeft.ForeColor = Color.Red
            ElseIf diasRestantes <= 60 Then
                lblDaysLeft.ForeColor = Color.Orange
            Else
                lblDaysLeft.ForeColor = Color.Green
            End If
        Else
            lblStatus.Text = "Estado: Licencia No Válida"
            lblStatus.ForeColor = Color.Red
            lblLicenseKey.Text = "Clave de Licencia: No disponible"
            lblExpiration.Text = "Fecha de Expiración: No disponible"
            lblDaysLeft.Text = "Días Restantes: 0"
            lblDaysLeft.ForeColor = Color.Red
        End If
    End Sub

    Private Sub ValidatePassword(sender As Object, e As EventArgs)
        btnGenerate.Enabled = (txtPassword.Text = "EscaladoProfesionalWR2025")
    End Sub

    Private Sub GenerateNewLicense(sender As Object, e As EventArgs)
        Try
            Dim license = LicenseManager.GenerateNewLicense()
            Dim mensaje As String = $"Nueva licencia generada correctamente:" & Environment.NewLine & Environment.NewLine & _
                                  $"Clave: {license.LicenseKey}" & Environment.NewLine & _
                                  $"Expira el: {license.ExpirationDate:dd/MM/yyyy}" & Environment.NewLine & _
                                  $"Días de validez: 365" & Environment.NewLine & Environment.NewLine & _
                                  $"IMPORTANTE: Guarde esta información en caso de necesitarla más adelante."
            
            MessageBox.Show(mensaje, "Licencia Generada", MessageBoxButtons.OK, MessageBoxIcon.Information)
            LoadLicenseInfo()
            Me.DialogResult = DialogResult.OK
        Catch ex As Exception
            MessageBox.Show("Error al generar la licencia: " & ex.Message, 
                          "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
