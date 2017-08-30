Imports System.Collections.Generic
Imports System.IO
Imports System.IO.Packaging
Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates
Imports Microsoft.Office.Interop.Excel
Imports CAPICOM
Imports System.Windows.Forms
Imports Ionic.Zip

Public Class ClsDigitalSignature
    Private _filePath As String = String.Empty
    Public c As String = Nothing
    Public s As String = Nothing
    Public IsDigitalSignatureAttachedFromUSBToken As Boolean = False
    Dim Obj_clsRprtGnrtnTool As New clsRprtGnrtnTool

    Public Property FilePath() As String
        Get
            Return _filePath
        End Get
        Set(value As String)
            _filePath = value
        End Set
    End Property

    Public Sub AddDigitalSignature(ByVal PFXFilePath As String, ByVal PFXFilePassword As String, ByVal TargetFilePath As String)

        Try
            Dim certificate = New X509Certificate(PFXFilePath, PFXFilePassword)

            ' Open the Package copy in the target directory.
            Dim e As String = Path.GetExtension(TargetFilePath)
            If e = ".xlsx" Then
                Using package__1 = Package.Open(TargetFilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None)

                    If ValidateSignatures(package__1) Then
                        Console.WriteLine("Already signed")
                        Return
                    End If

                    SignAllParts(package__1, certificate)
                    ' Validate the Package Signatures.
                    ' If all package signatures are valid, go ahead and unpack.

                    If ValidateSignatures(package__1) Then
                    End If
                    package__1.Flush()
                    package__1.Close()
                End Using
            Else
                Dim MyStore As New Store()
                Dim Signobj As New SignedData()
                Dim Signer As New Signer()
                Dim excel As New Microsoft.Office.Interop.Excel.Application()
                Dim objExcel As Workbook = excel.Workbooks.Open(TargetFilePath)
                Dim SigningTime As New CAPICOM.Attribute()
                Dim SubjectNameCn As String = Nothing
                SubjectNameCn = "CN = IDRBTTESTCLASS3_SIGNING_NOPAN"
                MyStore.Open(CAPICOM_STORE_LOCATION.CAPICOM_CURRENT_USER_STORE, "MY", CAPICOM_STORE_OPEN_MODE.CAPICOM_STORE_OPEN_READ_ONLY)
                Signer.Load(PFXFilePath, PFXFilePassword)
                Signobj.Content = objExcel.ToString()
                Signobj.Sign(Signer, False, CAPICOM_ENCODING_TYPE.CAPICOM_ENCODE_BASE64)
                MyStore.Close()

            End If
        Catch Ex As Exception
            Console.WriteLine("Digital Signature Error " & vbLf & Ex.Message)
        End Try


    End Sub

    ''' <summary>
    '''   Validates all the digital signatures of a given package.</summary>
    ''' <param name="package">
    '''   The package for validating digital signatures.</param>
    ''' <returns>
    '''   true if all digital signatures are valid; otherwise false if the
    '''   package is unsigned or any of the signatures are invalid.</returns>
    ''' 

    Private Function ValidateSignatures(ByVal package As Package) As Boolean
        If package Is Nothing Then
            Throw New ArgumentNullException("package")
        End If

        ' Create a PackageDigitalSignatureManager for the given Package.
        Dim dsm = New PackageDigitalSignatureManager(package)

        ' Check to see if the package contains any signatures.
        If Not dsm.IsSigned Then
            Return False
        End If
        ' The package is not signed.
        ' Verify that all signatures are valid.
        Dim result = dsm.VerifySignatures(False)
        Return result = VerifyResult.Success
    End Function

    Private Sub SignAllParts(package As Package, certificate As X509Certificate)
        Dim prgrmName As String = "ClsDigitalSignature->SignAllParts()"
        If package Is Nothing Then
            Throw New ArgumentNullException("package")
        End If
        If certificate Is Nothing Then
            Throw New ArgumentNullException("certificate")
        End If

        ' Create the DigitalSignature Manager
        Dim dsm = New PackageDigitalSignatureManager(package) With {.CertificateOption = CertificateEmbeddingOption.InSignaturePart}

        ' Create a list of all the part URIs in the package to sign
        ' (GetParts() also includes PackageRelationship parts).
        Dim toSign = New List(Of Uri)()
        For Each packagePart As Object In package.GetParts()
            ' Add all package parts to the list for signing.
            toSign.Add(packagePart.Uri)
        Next

        ' Add the URI for SignatureOrigin PackageRelationship part.
        ' The SignatureOrigin relationship is created when Sign() is called.
        ' Signing the SignatureOrigin relationship disables counter-signatures.
        toSign.Add(PackUriHelper.GetRelationshipPartUri(dsm.SignatureOrigin))

        ' Also sign the SignatureOrigin part.
        toSign.Add(dsm.SignatureOrigin)

        ' Add the package relationship to the signature origin to be signed.
        toSign.Add(PackUriHelper.GetRelationshipPartUri(New Uri("/", UriKind.RelativeOrAbsolute)))

        Try

            dsm.Sign(toSign, certificate)
        Catch ex As CryptographicException
            Obj_clsRprtGnrtnTool.logError(ex, prgrmName)
        End Try
    End Sub

    Public Sub FileZip(ZipFilePath As String, targetFilename As String, Password As String)
        Dim zipPath As String = ZipFilePath + ".zip"
        Dim prgrmName As String = "ClsDigitalSignature->FileZip()"
        Try
            If Not [String].IsNullOrEmpty(targetFilename) Then
                Using zip As New ZipFile()
                    zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestSpeed
                    zip.Password = Password
                    zip.AddFile(targetFilename, String.Empty)
                    zip.Save(zipPath)
                End Using
            End If
        Catch Ex As Exception
            Obj_clsRprtGnrtnTool.logError(Ex, prgrmName)
        End Try

    End Sub

    ''To Get Certificate Details from Token based on Serial Key
    Public Shared Function GetSignerCert() As X509Certificate2

        Dim Obj_clsRprtGnrtnTool As New clsRprtGnrtnTool
        Dim prgrmName As String = "ClsDigitalSignature->GetSignerCert() with USBToken Parameter"
        Dim functionReturnValue As X509Certificate2 = Nothing
        Dim certColl As X509Certificate2Collection = Nothing

        Try
            Dim Mystore As New X509Store(StoreName.My, StoreLocation.CurrentUser)
            Mystore.Open(OpenFlags.[ReadOnly] Or OpenFlags.OpenExistingOnly)

            ''Added by KP for ST
            '' certColl = Mystore.Certificates.Find(X509FindType.FindBySerialNumber, "440DF33F6B76D8DD68A9", False)
            ''Added by KP for UAT
            'certColl = Mystore.Certificates.Find(X509FindType.FindBySerialNumber, "51F5752F08E2B4C781", False)
            ''Added by KP for Nandan Token
            ''      certColl = Mystore.Certificates.Find(X509FindType.FindBySerialNumber, "5315A9474D8BE7132404", False)
            ''Added by KP for Nandan Token wef Certificate Validity - ‎ 24 ‎February ‎2016  To 23 ‎February ‎2018 ‎
            certColl = Mystore.Certificates.Find(X509FindType.FindBySerialNumber, "142226906A66879ACEDF", False)



            '' 51F5752F08E2B4C781
            If certColl.Count = 0 Then
                prgrmName = "ClsDigitalSignature->GetSignerCert()->The Certificate is not found in the certificate store." & Environment.NewLine & Environment.NewLine & "Please check if valid Serial Number is supplied in Configuration text file"
                Dim Ex As New Exception
                Obj_clsRprtGnrtnTool.logError(Ex, prgrmName)
            End If
            Mystore.Close()
        Catch ex As Exception
            Obj_clsRprtGnrtnTool.logError(ex, prgrmName)
        End Try
        Return certColl(0)

    End Function

    Public Function AddDigitalSignature(ByVal PFXFilePath As String, ByVal PFXFilePassword As String, ByVal TargetFilePath As String, ByVal USBToken As Boolean) As Boolean

        Dim prgrmName As String = "ClsDigitalSignature->AddDigitalSignature() with USBToken Parameter"
        Dim certificate = New X509Certificate2()

        Try
            certificate = GetSignerCert()
            certificate.SetPinForPrivateKey(Configuration.ConfigurationManager.AppSettings("usbpassword").ToString)

            ' Open the Package copy in the target directory.
            Dim e As String = Path.GetExtension(TargetFilePath)
            If e = ".xlsx" Then
                Using package__1 = Package.Open(TargetFilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
                    If ValidateSignatures(package__1) Then
                        IsDigitalSignatureAttachedFromUSBToken = True
                        Return IsDigitalSignatureAttachedFromUSBToken
                    End If
                    SignAllParts(package__1, certificate)

                    ' Validate the Package Signatures.
                    ' If all package signatures are valid, go ahead and unpack.
                    If ValidateSignatures(package__1) Then
                    End If
                    package__1.Flush()
                    package__1.Close()
                    IsDigitalSignatureAttachedFromUSBToken = True
                End Using
            End If
        Catch Ex As Exception
            Obj_clsRprtGnrtnTool.logError(Ex, prgrmName)
            IsDigitalSignatureAttachedFromUSBToken = False
        End Try
        Return IsDigitalSignatureAttachedFromUSBToken
    End Function


End Class

