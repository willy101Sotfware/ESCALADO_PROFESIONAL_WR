Imports System.IO
Imports System.Linq
Imports System.Security.Cryptography
Imports System.Text

Public Class LicenseManager
    Public Class LicenseInfo
        Public Property LicenseKey As String
        Public Property ExpirationDate As Date
        Public Property IsValid As Boolean
    End Class
    'clave para la nueva licencia 
    Private Shared ReadOnly LicenseFile As String = "license.dat"
    Private Shared ReadOnly Key As String = "EscaladoProfesionalWR2025"

    Public Shared Function ValidateLicense() As LicenseInfo
        Try
            If Not File.Exists(LicenseFile) Then
                Return New LicenseInfo With {
                    .IsValid = False,
                    .ExpirationDate = Date.MinValue
                }
            End If

            Dim encryptedData As String = File.ReadAllText(LicenseFile)
            Dim decryptedData As String = Decrypt(encryptedData, Key)
            Dim parts As String() = decryptedData.Split("|"c)

            If parts.Length <> 2 Then
                Return New LicenseInfo With {
                    .IsValid = False,
                    .ExpirationDate = Date.MinValue
                }
            End If

            Dim expirationDate As Date = Date.Parse(parts(1))
            Return New LicenseInfo With {
                .LicenseKey = parts(0),
                .ExpirationDate = expirationDate,
                .IsValid = (expirationDate > Date.Now)
            }
        Catch
            Return New LicenseInfo With {
                .IsValid = False,
                .ExpirationDate = Date.MinValue
            }
        End Try
    End Function

    Public Shared Function GenerateNewLicense() As LicenseInfo
        Dim key As String = GenerateRandomKey()
        Dim expirationDate As Date = Date.Now.AddYears(1)
        Dim license As New LicenseInfo With {
            .LicenseKey = key,
            .ExpirationDate = expirationDate,
            .IsValid = True
        }
        SaveLicense(license)
        Return license
    End Function

    Private Shared Sub SaveLicense(license As LicenseInfo)
        Dim data As String = $"{license.LicenseKey}|{license.ExpirationDate:yyyy-MM-dd HH:mm:ss}"
        Dim encryptedData As String = Encrypt(data, Key)
        File.WriteAllText(LicenseFile, encryptedData)
    End Sub

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
