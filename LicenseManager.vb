Imports System.IO
Imports System.Linq
Imports System.Security.Cryptography
Imports System.Text
Imports System.Windows.Forms

Public Class LicenseManager
    Public Enum LicenseTypes
        Trial = 0    ' 1 semana
        SixMonths = 1
        OneYear = 2
    End Enum

    Public Class LicenseInfo
        Public Property LicenseKey As String
        Public Property ExpirationDate As Date
        Public Property IsValid As Boolean
        Public Property LicenseType As LicenseTypes
    End Class

    Private Shared ReadOnly LicenseFile As String = "license.dat"
    Private Shared ReadOnly Key As String = "EscaladoProfesionalWR2025"

    Public Shared Function ValidateLicense() As LicenseInfo
        Try
            Dim defaultLicense As New LicenseInfo()
            defaultLicense.LicenseType = LicenseTypes.Trial
            defaultLicense.ExpirationDate = Date.MinValue
            defaultLicense.IsValid = False

            If Not File.Exists(LicenseFile) Then
                Return defaultLicense
            End If

            Dim encryptedData As String = File.ReadAllText(LicenseFile)
            Dim decryptedData As String = Decrypt(encryptedData, Key)
            Dim parts As String() = decryptedData.Split("|"c)

            If parts.Length <> 3 Then
                Return defaultLicense
            End If

            Dim expirationDate As Date = Date.Parse(parts(1))
            Dim licenseType As LicenseTypes = CType(Integer.Parse(parts(2)), LicenseTypes)

            Dim license As New LicenseInfo()
            license.LicenseKey = parts(0)
            license.ExpirationDate = expirationDate
            license.IsValid = (expirationDate > Date.Now)
            license.LicenseType = licenseType
            Return license

        Catch ex As Exception
            Return New LicenseInfo With {
                .IsValid = False,
                .ExpirationDate = Date.MinValue,
                .LicenseType = LicenseTypes.Trial
            }
        End Try
    End Function

    Public Shared Function GenerateNewLicense(licenseType As LicenseTypes) As LicenseInfo
        Dim key As String = GenerateRandomKey()
        Dim expirationDate As Date = GetExpirationDateForLicenseType(licenseType)

        Dim license As New LicenseInfo()
        license.LicenseKey = key
        license.ExpirationDate = expirationDate
        license.IsValid = True
        license.LicenseType = licenseType

        SaveLicense(license)
        Return license
    End Function

    Private Shared Function GetExpirationDateForLicenseType(licenseType As LicenseTypes) As Date
        Select Case licenseType
            Case LicenseTypes.Trial
                Return Date.Now.AddDays(7)  ' 1 semana
            Case LicenseTypes.SixMonths
                Return Date.Now.AddMonths(6)
            Case LicenseTypes.OneYear
                Return Date.Now.AddYears(1)
            Case Else
                Return Date.Now.AddDays(7)
        End Select
    End Function

    Private Shared Sub SaveLicense(license As LicenseInfo)
        Dim data As String = $"{license.LicenseKey}|{license.ExpirationDate:yyyy-MM-dd HH:mm:ss}|{CInt(license.LicenseType)}"
        Dim encryptedData As String = Encrypt(data, Key)
        File.WriteAllText(LicenseFile, encryptedData)
    End Sub

    Public Shared Function ShowLicenseDialog() As LicenseInfo
        Using form As New Form()
            form.Text = "Generar Nueva Licencia"
            form.Size = New Size(400, 300)
            form.FormBorderStyle = FormBorderStyle.FixedDialog
            form.StartPosition = FormStartPosition.CenterScreen
            form.MaximizeBox = False
            form.MinimizeBox = False

            ' Título
            Dim lblTitle As New Label()
            lblTitle.Text = "Seleccione el tipo de licencia:"
            lblTitle.Location = New Point(20, 20)
            lblTitle.Size = New Size(360, 30)
            lblTitle.Font = New Font("Arial", 12, FontStyle.Bold)
            lblTitle.TextAlign = ContentAlignment.MiddleCenter

            ' Radio buttons para tipos de licencia
            Dim rbTrial As New RadioButton()
            rbTrial.Text = "Licencia de Prueba (1 semana)"
            rbTrial.Location = New Point(40, 70)
            rbTrial.Size = New Size(300, 30)
            rbTrial.Checked = True

            Dim rbSixMonths As New RadioButton()
            rbSixMonths.Text = "Licencia Semestral (6 meses)"
            rbSixMonths.Location = New Point(40, 110)
            rbSixMonths.Size = New Size(300, 30)

            Dim rbOneYear As New RadioButton()
            rbOneYear.Text = "Licencia Anual (1 año)"
            rbOneYear.Location = New Point(40, 150)
            rbOneYear.Size = New Size(300, 30)

            ' Botón generar
            Dim btnGenerate As New Button()
            btnGenerate.Text = "Generar Licencia"
            btnGenerate.Location = New Point(100, 200)
            btnGenerate.Size = New Size(200, 40)
            btnGenerate.BackColor = Color.FromArgb(0, 120, 215)
            btnGenerate.ForeColor = Color.White
            btnGenerate.FlatStyle = FlatStyle.Flat

            Dim result As LicenseInfo = Nothing

            AddHandler btnGenerate.Click, Sub()
                                              Dim selectedType As LicenseTypes
                                              If rbTrial.Checked Then
                                                  selectedType = LicenseTypes.Trial
                                              ElseIf rbSixMonths.Checked Then
                                                  selectedType = LicenseTypes.SixMonths
                                              Else
                                                  selectedType = LicenseTypes.OneYear
                                              End If

                result = GenerateNewLicense(selectedType)
                form.DialogResult = DialogResult.OK
                form.Close()
            End Sub

            form.Controls.AddRange({lblTitle, rbTrial, rbSixMonths, rbOneYear, btnGenerate})
            
            If form.ShowDialog() = DialogResult.OK Then
                Return result
            End If

            Return Nothing
        End Using
    End Function

    Private Shared Function GenerateRandomKey() As String
        Dim chars As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
        Dim random As New Random()
        Return New String(Enumerable.Repeat(chars, 16).Select(Function(s) s(random.Next(s.Length))).ToArray())
    End Function

    Private Shared Function Encrypt(text As String, key As String) As String
        Using aes As Aes = Aes.Create()
            aes.Key = GetKey(key)
            aes.IV = New Byte(15) {}

            Dim encryptor As ICryptoTransform = aes.CreateEncryptor(aes.Key, aes.IV)
            Using msEncrypt As New MemoryStream()
                Using csEncrypt As New CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write)
                    Using swEncrypt As New StreamWriter(csEncrypt)
                        swEncrypt.Write(text)
                    End Using
                    Return Convert.ToBase64String(msEncrypt.ToArray())
                End Using
            End Using
        End Using
    End Function

    Private Shared Function Decrypt(cipherText As String, key As String) As String
        Using aes As Aes = Aes.Create()
            aes.Key = GetKey(key)
            aes.IV = New Byte(15) {}

            Dim decryptor As ICryptoTransform = aes.CreateDecryptor(aes.Key, aes.IV)
            Using msDecrypt As New MemoryStream(Convert.FromBase64String(cipherText))
                Using csDecrypt As New CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read)
                    Using srDecrypt As New StreamReader(csDecrypt)
                        Return srDecrypt.ReadToEnd()
                    End Using
                End Using
            End Using
        End Using
    End Function

    Private Shared Function GetKey(key As String) As Byte()
        Using sha256 As SHA256 = SHA256.Create()
            Return sha256.ComputeHash(Encoding.UTF8.GetBytes(key)).Take(16).ToArray()
        End Using
    End Function
End Class
