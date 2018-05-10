Option Strict On

Imports Microsoft.Win32
Imports System.Text
Imports System.Text.RegularExpressions

'v0.2: Took a list for folder ids instead of using Trim and/or Split. We just have to strFolderIDList(ListBox1.SelectedIndex)
'v0.1: Initial release

Public Class Form1

    Dim strProgramPath As String
    Dim strSteamPath As String
    Dim strFolderIDList As New List(Of String)()

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        strSteamPath = getSteamPath()

        If strSteamPath <> String.Empty Then

            For Each directory In System.IO.Directory.GetDirectories(strSteamPath)

                'Get user ids and steam names
                Dim strUserConfigSource As String
                Dim strUserIDSteamName As String

                strUserConfigSource = System.IO.File.ReadAllText(directory & "\config\localconfig.vdf")
                strUserIDSteamName = GetToken("""PersonaName""" & vbTab & vbTab & """(.+?)""", 1, strUserConfigSource)

                ListBox1.Items.Add(strUserIDSteamName & " [" & Replace(directory, strSteamPath & "\", "") & "]")
                strFolderIDList.Add(IO.Path.GetFileNameWithoutExtension(directory)) 'GetFileNameWithoutExtension is no joke. GetDirectoryName returns the part before the directory name

            Next

            strProgramPath = Application.StartupPath & "\"

        Else

            ListBox1.Enabled = False
            Button1.Enabled = False
            MessageBox.Show("Steam isn't installed")

        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If Not ListBox1.SelectedItem Is Nothing Then

            Dim strUserIDPath As String

            'ID -> Path
            strUserIDPath = strSteamPath & "\" & strFolderIDList(ListBox1.SelectedIndex) & "\"

            For Each file In System.IO.Directory.GetFiles(strProgramPath & "files")

                Dim strtmp As String = IO.Path.GetFileName(file)

                If System.IO.File.Exists(strUserIDPath & "730\local\cfg\" & strtmp) Then
                    System.IO.File.Delete(strUserIDPath & "730\local\cfg\" & strtmp)
                End If

                System.IO.File.Copy(file, strUserIDPath & "730\local\cfg\" & System.IO.Path.GetFileName(file))

            Next

            MessageBox.Show("Work done!")

        Else

            MessageBox.Show("Please select an account")

        End If


    End Sub

    Private Function getSteamPath() As String

        'Search steam in standard path (x64 windows)
        If IO.File.Exists("C:\Program Files\Steam\Steam.exe") Then

            Return "C:\Program Files\Steam\userdata"

        Else

            'Search steam in standard path (x86 windows)
            If IO.File.Exists("C:\Program Files (x86)\Steam\Steam.exe") Then

                Return "C:\Program Files (x86)\Steam\userdata"

            Else

                'Search steam path in x86 registry (cpu dependent)
                Dim RegTyp32 As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)
                Dim Key32 As RegistryKey = RegTyp32.OpenSubKey("SOFTWARE\Valve\Steam", False)

                If Key32 IsNot Nothing Then

                    Return CStr(Key32.GetValue("InstallPath")) & "\userdata"

                Else

                    'Search steam path in x64 registry (cpu dependent)
                    Dim RegTyp64 As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64)
                    Dim Key64 As RegistryKey = RegTyp64.OpenSubKey("SOFTWARE\Valve\Steam", False)

                    If Key64 IsNot Nothing Then

                        Return CStr(Key64.GetValue("InstallPath")) & "\userdata"

                    Else

                        'Steam is not installed
                        Return ""

                    End If
                End If
            End If
        End If

    End Function

    Private Function GetToken(ByRef pattern As String, ByRef value As Integer, ByRef source As String) As String

        Dim RegexObj As New Regex(pattern, RegexOptions.IgnoreCase)
        Dim RegexMatch As Match = RegexObj.Match(source)

        Return RegexMatch.Groups(value).ToString

    End Function

End Class
