Imports System.ComponentModel
Imports System.Text.RegularExpressions
Imports System.ServiceProcess

Public Class Form1

    'Button głowny
    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click

        'Czyszczenie
        ProgressBar1.Value = 0
        RichTextBox1.Clear()

        Dim dopasowanie As Match = Regex.Match(TextBox1.Text, "FlexiBRSvc.exe")

        If dopasowanie.Success = 0 Then
            MsgBox("Brak wskazanego pliku wykonalnego!")
        Else
            'Start CMD
            Dim CMDThread As New Threading.Thread(AddressOf CMDAutomate)
            CMDThread.Start()
            time.Enabled = False
        End If
    End Sub

    'Liczniki
    Dim Count As Integer = 120
    Dim Count2 As Integer = 60

    Public WithEvents time As Timer = New Timer()
    Public WithEvents timee As Timer = New Timer()

    'Licznik główny
    Private Sub TimeTicked() Handles time.Tick
        Count -= 1
        Label2.Text = Count
        ProgressBar1.Maximum = 1000
        'progressBar1.Value = ProgressBar1.Value + 1
        If (Count = 0) Then
            time.Enabled = False
            Label2.Text = " "
            Dim CMDThread As New Threading.Thread(AddressOf CMDAutomate)
            CMDThread.Start()
        End If
    End Sub

    'Ładowanie formy
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Timer
        time.Interval = 1000
        time.Enabled = True
        timee.Enabled = False

        'Czyszczenie progressbar-a
        ProgressBar1.Value = 0
        'ToolStripProgressBar1.Value = 0

        'Tworzenie daty do loga
        Dokumenty = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        Dim czas As String
        czas = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")

        If My.Computer.FileSystem.FileExists(Dokumenty + "\FC_UPDATER.log") Then
            My.Computer.FileSystem.WriteAllText(Dokumenty + "\FC_UPDATER.log", vbCrLf + "[" + czas + "]" + vbCrLf, True)
            'MsgBox(fi.Length.ToString())
        Else
            My.Computer.FileSystem.WriteAllText(Dokumenty + "\FC_UPDATER.log", "[" + czas + "]" + vbCrLf, True)
            'MsgBox(fi.Length.ToString())
        End If

        'Ustawienia formy
        MaximizeBox = False
        Button7.Enabled = False
        Button6.Enabled = False
        RichTextBox2.Visible = False
        ListBox1.Items.Clear()
        RichTextBox1.Clear()

        'Wczytywnie danych z app-config
        Dim contents As Specialized.StringCollection = My.Settings.Lista
        For Each item As String In contents
            ListBox1.Items.Add(item)
        Next

        'MsgBox(DateTime.Now.ToString("yyyy-MM-dd"))

    End Sub

    'Button "Przeglądaj
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        OpenFileDialog1.InitialDirectory = "c:\"
        OpenFileDialog1.Filter = "Exe|*.exe"
        OpenFileDialog1.Title = "Wskaż plik: FlexiBRSvc.exe"
        OpenFileDialog1.FileName = "FlexiBRSvc.exe"
        OpenFileDialog1.RestoreDirectory = True

        If OpenFileDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            If OpenFileDialog1.SafeFileName = "FlexiBRSvc.exe" Then
                Dim path As String = OpenFileDialog1.FileName
                TextBox1.Text = path
            Else
                MessageBox.Show("Zła nazwa pliku!")
                Dim path As String = OpenFileDialog1.FileName
                TextBox1.Text = " "
            End If
        End If

    End Sub

    'Button "Anuluj"
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        time.Enabled = False
        Label2.Text = " "
    End Sub

    'Button "Dodaj"
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ListBox1.Items.Add(TextBox6.Text)
        TextBox6.Clear()
    End Sub

    'Button "Usuń"
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
    End Sub

    'Stare - zapis do configu
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        'If My.Computer.FileSystem.FileExists("C:\Users\pmazurkiewicz\Desktop\updater_config.ini") = True Then
        '    My.Computer.FileSystem.DeleteFile("C:\Users\pmazurkiewicz\Desktop\updater_config.ini")
        'End If

        'If String.IsNullOrEmpty(TextBox5.Text) Then
        '    MessageBox.Show("ID nie może być puste.",
        '                    "Error!",
        '                    MessageBoxButtons.OK,
        '                    MessageBoxIcon.Exclamation)
        'End If

        'My.Computer.FileSystem.WriteAllText("C:\Users\pmazurkiewicz\Desktop\updater_config.ini", "[ID] = " + TextBox5.Text + vbCrLf, True)

        'For i = 0 To ListBox1.Items.Count - 1
        '    My.Computer.FileSystem.WriteAllText("C:\Users\pmazurkiewicz\Desktop\updater_config.ini", "[DataSet] = " + ListBox1.Items(i) + vbCrLf, True)
        'Next

    End Sub

    'Zamykanie formy, zapis do app-config
    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Dim contents As New Specialized.StringCollection
        For Each elem As String In ListBox1.Items
            contents.Add(elem)
        Next

        My.Settings.Lista = contents
        My.Settings.Save()
    End Sub

    'Kasowanie data-set
    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        If TextBox5.Text = "TextBox5.Text" Then
            TextBox5.Text = " "
        End If
    End Sub

    'Odpalanie CMD i funkcji głównej programu
    Private Results As String
    Private Baza As String
    Private Dokumenty As String
    Private Delegate Sub delUpdate()
    Private Finished As New delUpdate(AddressOf UpdateText)

    Private Sub UpdateText()

        Label2.Text = " "

        RichTextBox2.Text = Results

        Threading.Thread.Sleep(1000)

        Dim keywords() As String = {"succeeded"}
        Dim message

        If keywords.Count(Function(w) RichTextBox2.Text.ToLower.Contains(w)) > 0 Then
            message = "Baza: " + Chr(34) + Baza + Chr(34) + " została zaktualizowana."
            RichTextBox1.Text = RichTextBox1.Text + message + vbCrLf
            ProgressBar1.Value = ProgressBar1.Value + ProgressBar1.Maximum / ListBox1.Items.Count
            'ToolStripProgressBar1.Value = ToolStripProgressBar1.Value + ToolStripProgressBar1.Maximum / ListBox1.Items.Count
            Dokumenty = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            My.Computer.FileSystem.WriteAllText(Dokumenty + "\FC_UPDATER.log", "Baza: " + Chr(34) + Baza + Chr(34) + " została zaktualizowana." + vbCrLf, True)
            Button7.Enabled = True
            timee.Enabled = True
        Else
            message = "Baza: " + Chr(34) + Baza + Chr(34) + " nie została zaktualizowana!"
            RichTextBox1.Text = RichTextBox1.Text + message + vbCrLf
            ProgressBar1.Value = ProgressBar1.Value + ProgressBar1.Maximum / ListBox1.Items.Count
            'ToolStripProgressBar1.Value = ToolStripProgressBar1.Value + ToolStripProgressBar1.Maximum / ListBox1.Items.Count
            Dokumenty = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            My.Computer.FileSystem.WriteAllText(Dokumenty + "\FC_UPDATER.log", "Baza: " + Chr(34) + Baza + Chr(34) + " nie została zaktualizowana!" + vbCrLf, True)
            Button7.Enabled = False
        End If

    End Sub

    'Licznik 2
    Private Sub Timer2_Timer() Handles timee.Tick
        timee.Interval = 1000
        Count2 -= 1
        Button7.Text = "OK (" + Str(Count2) + " )"

        If (Count2 = 0) Then
            Close()
        End If

    End Sub

    'Button "OK"
    Private Sub Button7_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button7.Click
        Close()
    End Sub

    'Odpalanie CMD
    Private Sub CMDAutomate()

        For Each elem As String In ListBox1.Items
            Dim myprocess As New Process
            Dim StartInfo As New System.Diagnostics.ProcessStartInfo
            StartInfo.FileName = "cmd"
            StartInfo.RedirectStandardInput = True
            StartInfo.RedirectStandardOutput = True
            StartInfo.UseShellExecute = False
            StartInfo.CreateNoWindow = True
            myprocess.StartInfo = StartInfo
            myprocess.Start()
            Dim SR As System.IO.StreamReader = myprocess.StandardOutput
            Dim SW As System.IO.StreamWriter = myprocess.StandardInput
            SW.WriteLine(Chr(34) + TextBox1.Text + Chr(34) + " please update dataset ""http://localhost/" + LTrim(TextBox5.Text) + Chr(34) + " " + Chr(34) + LTrim(TextBox2.Text) + Chr(34) + " " + Chr(34) + LTrim(elem) + Chr(34))
            My.Computer.FileSystem.WriteAllText(Dokumenty + "\FC_UPDATER.log", Chr(34) + TextBox1.Text + Chr(34) + " please update dataset ""http://localhost/" + LTrim(TextBox5.Text) + Chr(34) + " " + Chr(34) + LTrim(TextBox2.Text) + Chr(34) + " " + Chr(34) + LTrim(elem) + Chr(34), True)
            SW.WriteLine("exit")
            Results = SR.ReadToEnd
            Threading.Thread.Sleep(2000)
            Baza = elem
            SW.Close()
            SR.Close()
            Invoke(Finished)
        Next

    End Sub

    'Czyszczenie
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "TextBox1.Text" Then
            TextBox1.Text = " "
        End If
    End Sub

    'Czyszczenie
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text = "TextBox2.Text" Then
            TextBox2.Text = " "
        End If
    End Sub
End Class