Imports System.IO
Imports System.Reflection
Public Class Form1
    Dim upk_dll As Assembly
    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Button3.Enabled = False
        If IO.File.Exists(TextBox1.Text) Then
            Try
                System.Reflection.AssemblyName.GetAssemblyName(TextBox1.Text)
            Catch ex As Exception
                TextBox1.BackColor = ColorTranslator.FromHtml("#E9C083")
                Label3.Text = "Not an assembly file."
                Return
            End Try
            TextBox1.BackColor = ColorTranslator.FromHtml("#AAE895")
            Label3.Text = "File is valid."
            Button3.Enabled = True
        Else
            TextBox1.BackColor = ColorTranslator.FromHtml("#EB8181")
            Label3.Text = "Invalid path."
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Panel2.Hide()
        SplitContainer1.Enabled = True
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If OpenFileDialog1.ShowDialog = vbOK Then
            TextBox1.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Panel2_DragDrop(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Panel2.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        TextBox1.Text = files(0)
    End Sub

    Private Sub Panel2_DragEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Panel2.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        RichTextBox1.Clear()
        Dim asm As Assembly = Assembly.LoadFile(TextBox1.Text)
        RichTextBox1.Text = "Resources of """ & IO.Path.GetFileName(TextBox1.Text) & """ (" & asm.GetManifestResourceNames.Count & ")" & vbCrLf & "- %Name% | %HEXSize%" & vbCrLf & vbCrLf
        For x = 0 To asm.GetManifestResourceNames.Count - 1
            RichTextBox1.Text += "- " & asm.GetManifestResourceNames(x) & " | 0x" & Hex(asm.GetManifestResourceStream(asm.GetManifestResourceNames(x)).Length.ToString) & vbCrLf
        Next
    End Sub
    Function decompress(ByVal data As Byte(), ByVal compressionType As Integer)
        Dim asm As Assembly = Assembly.LoadFile(TextBox1.Text)
        Dim method As MethodInfo = upk_dll.GetType("bytepress.lib.Main").GetMethod("Decompress")
        Return CType(method.Invoke(Nothing, New Object() {data, compressionType}), Byte())
    End Function
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        RichTextBox1.Clear()
        IO.Directory.CreateDirectory("BytePress_Unpacker\" & IO.Path.GetFileNameWithoutExtension(TextBox1.Text))
        Using manifestResourceStream As Stream = Assembly.LoadFile(TextBox1.Text).GetManifestResourceStream("bytepress.lib.dll")
            Using memoryStream As MemoryStream = New MemoryStream()
                manifestResourceStream.CopyTo(memoryStream)
                upk_dll = Assembly.Load(memoryStream.ToArray())
                If CheckBox2.Checked Then
                    RichTextBox1.Text += "Extracting ""bytepress.lib.dll""..." & vbCrLf
                    IO.File.WriteAllBytes("BytePress_Unpacker\" & IO.Path.GetFileNameWithoutExtension(TextBox1.Text) & "\bytepress.lib.dll", memoryStream.ToArray())
                    RichTextBox1.Text += """bytepress.lib.dll"" successfully extracted!" & vbCrLf & vbCrLf
                End If
            End Using
        End Using
        Using manifestResourceStream2 As Stream = Assembly.LoadFile(TextBox1.Text).GetManifestResourceStream("data")
            Using memoryStream2 As MemoryStream = New MemoryStream()
                manifestResourceStream2.CopyTo(memoryStream2)
                Dim array As Byte() = decompress(memoryStream2.ToArray(), 2)
                If array Is Nothing OrElse array.Length = 0 Then
                    Throw New Exception("Failed to decompress file")
                End If
                RichTextBox1.Text += "Extracting ""data.bin""..." & vbCrLf
                IO.File.WriteAllBytes("BytePress_Unpacker\" & IO.Path.GetFileNameWithoutExtension(TextBox1.Text) & "\data.bin", array)
                RichTextBox1.Text += """data.bin"" successfully extracted!"
            End Using
        End Using
        If MessageBox.Show("Files have been extracted, do you want to open the output folder?", "BytePress Unpacker", MessageBoxButtons.YesNo, MessageBoxIcon.Information) = Windows.Forms.DialogResult.Yes Then
            Process.Start("explorer.exe", "BytePress_Unpacker\" & IO.Path.GetFileNameWithoutExtension(TextBox1.Text))
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        TopMost = CheckBox1.Checked
    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Process.Start("https://github.com/roachadam/bytepress")
    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click
        Process.Start("https://github.com/miso-xyz/BytePressUnpacker")
    End Sub

    Private Sub Form1_HelpButtonClicked(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.HelpButtonClicked
        ContextMenu1.Show(Me, New Point(Size.Width - 80, 0))
    End Sub
End Class
