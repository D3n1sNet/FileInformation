Imports System.IO
Imports System.Security.Cryptography

Public Class FileInformation
    Public Shared Function IsNetFile(ByVal Path As String) As Boolean
        Dim FileInfo As String = My.Computer.FileSystem.ReadAllText(Path)
        If FileInfo.Contains("mscorlib") Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Shared Function GetFileSize(ByVal path As String) As String
        Dim myFile As FileInfo
        Dim mySize As Single
        Try
            myFile = New FileInfo(path)

            If Not myFile.Exists Then
                mySize = 0
            Else
                mySize = myFile.Length
            End If

            Select Case mySize
                Case 0 To 1023
                    Return mySize & " Bytes"

                Case 1024 To 1048575
                    Return Format(mySize / 1024, "###0.00") & " Kilobytes"

                Case 1048576 To 1043741824
                    Return Format(mySize / 1024 ^ 2, "###0.00") & " Megabytes"

                Case Is > 1043741824
                    Return Format(mySize / 1024 ^ 3, "###0.00") & " Gigabytes"
            End Select
            Return "0 bytes"
        Catch ex As Exception
            Return "0 bytes"
        End Try
    End Function

    Public Shared Function GetCRC32(ByVal sFileName As String) As String
        Try
            Dim FS As FileStream = New FileStream(sFileName, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
            Dim CRC32Result As Integer = &HFFFFFFFF
            Dim Buffer(4096) As Byte
            Dim ReadSize As Integer = 4096
            Dim Count As Integer = FS.Read(Buffer, 0, ReadSize)
            Dim CRC32Table(256) As Integer
            Dim DWPolynomial As Integer = &HEDB88320
            Dim DWCRC As Integer
            Dim i As Integer, j As Integer, n As Integer
            For i = 0 To 255
                DWCRC = i
                For j = 8 To 1 Step -1
                    If (DWCRC And 1) Then
                        DWCRC = ((DWCRC And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                        DWCRC = DWCRC Xor DWPolynomial
                    Else
                        DWCRC = ((DWCRC And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                    End If
                Next j
                CRC32Table(i) = DWCRC
            Next i

            Do While (Count > 0)
                For i = 0 To Count - 1
                    n = (CRC32Result And &HFF) Xor Buffer(i)
                    CRC32Result = ((CRC32Result And &HFFFFFF00) \ &H100) And &HFFFFFF
                    CRC32Result = CRC32Result Xor CRC32Table(n)
                Next i
                Count = FS.Read(Buffer, 0, ReadSize)
            Loop
            Return Hex(Not (CRC32Result))
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Shared Function MD5FileHash(ByVal sFile As String) As String
        Dim MD5 As New MD5CryptoServiceProvider
        Dim Hash As Byte()
        Dim Result As String = ""
        Dim Tmp As String = ""

        Dim FN As New FileStream(sFile, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
        MD5.ComputeHash(FN)
        FN.Close()

        Hash = MD5.Hash

        For i As Integer = 0 To Hash.Length - 1
            Tmp = Hex(Hash(i))
            If Len(Tmp) = 1 Then Tmp = "0" & Tmp
            Result += Tmp
        Next
        Return Result
    End Function

    Public Shared Function SHA1FileHash(ByVal sFile As String) As String
        Dim SHA1 As New SHA1CryptoServiceProvider
        Dim Hash As Byte()
        Dim Result As String = ""
        Dim Tmp As String = ""

        Dim FN As New FileStream(sFile, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
        SHA1.ComputeHash(FN)
        FN.Close()

        Hash = SHA1.Hash

        For i As Integer = 0 To Hash.Length - 1
            Tmp = Hex(Hash(i))
            If Len(Tmp) = 1 Then Tmp = "0" & Tmp
            Result += Tmp
        Next
        Return Result
    End Function
End Class
