Imports System.Security.Cryptography
Imports System.IO
Imports System.Text

Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.ShowDialog()
        Dim md5code As String
        Dim md5 As MD5CryptoServiceProvider = New MD5CryptoServiceProvider
        Dim f As FileStream = New FileStream(OpenFileDialog1.FileName, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
        f = New FileStream(OpenFileDialog1.FileName, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
        md5.ComputeHash(f)
        Dim ObjFSO As Object = CreateObject("Scripting.FileSystemObject")
        Dim objFile = ObjFSO.GetFile(OpenFileDialog1.FileName)

        Dim hash As Byte() = md5.Hash
        Dim buff As StringBuilder = New StringBuilder
        Dim hashByte As Byte
        For Each hashByte In hash
            buff.Append(String.Format("{0:X1}", hashByte))
        Next

        md5code = buff.ToString()
        TextBox1.Text = md5code
    End Sub
End Class
