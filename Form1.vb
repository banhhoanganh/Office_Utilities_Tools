Imports System.IO
Imports System.Reflection
Imports Microsoft.Win32
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel


Public Class Form1
    ' read and write registry
    Private Const RegistryPath As String = "HKEY_CURRENT_USER\Software\Office_Utilities_Tools"
    Private Const AutoStartup As String = "AutoStartup"


    Private startupFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.Startup)
    'Private currentFolder As String = Directory.GetCurrentDirectory()
    Private currentFilePath As String = Assembly.GetExecutingAssembly().Location
    Private fileName As String = Path.GetFileName(currentFilePath)

    Private isDragging As Boolean = False
    Private startPoint As Point

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim processes As Process() = Process.GetProcessesByName("Office_Utilities_Tools")
        If processes.Length > 1 Then
            For Each proc As Process In processes
                If proc.Id <> Process.GetCurrentProcess().Id Then
                    proc.Kill()
                End If
            Next
            'Me.Close()
        End If

        LoadRegistery()

        '' Configure the NotifyIcon
        AddHandler Me.EndTaskMicrosoftExelToolStripMenuItem.Click, AddressOf ExitExel
        AddHandler Me.ExitToolStripMenuItem.Click, AddressOf ExitApplication

        'Me.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub AddStartup_Click(sender As Object, e As EventArgs) Handles AddStartup.Click
        SaveRegistery()
        ShortcutStartup(AddStartup.Checked)
    End Sub

    Private Sub LoadRegistery()
        Dim autoStartupValue As String = Microsoft.Win32.Registry.GetValue(RegistryPath, AutoStartup, Nothing)
        If autoStartupValue IsNot Nothing Then
            ' MsgBox(autoStartupValue.ToString())
            If autoStartupValue.ToString() = "True" Then
                ShortcutStartup(True)
            Else
                ShortcutStartup(False)
            End If
        Else
            ' If the registry value doesn't exist, create it with a default value
            Registry.SetValue(RegistryPath, AutoStartup, "True")
            ShortcutStartup(True)
        End If
    End Sub

    Private Sub SaveRegistery()
        Dim autoStartupValue As String = Microsoft.Win32.Registry.GetValue(RegistryPath, AutoStartup, Nothing)
        Registry.SetValue(RegistryPath, AutoStartup, AddStartup.Checked.ToString())
    End Sub

    ' Handle double-click on the NotifyIcon to restore the form
    Private Sub NotifyIcon_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        Me.Show() ' Show the form
        'Me.WindowState = FormWindowState.Normal ' Restore the form
    End Sub

    ' Exit the application from the NotifyIcon context menu
    Private Sub ExitApplication(sender As Object, e As EventArgs)
        NotifyIcon1.Visible = False
        Application.Exit()
    End Sub

    Private Sub ExitExel(sender As Object, e As EventArgs)
        KillAPP("Excel")
    End Sub

    Private Sub ShortcutStartup(Creat As Boolean)
        Dim shortcutPath As String = Path.Combine(startupFolder, Path.GetFileNameWithoutExtension(currentFilePath) & ".lnk")
        If File.Exists(shortcutPath) Then
            File.Delete(shortcutPath)
        End If
        AddStartup.Checked = Creat
        If Creat Then
            Try
                ' Create a shortcut using Windows Script Host
                Dim wsh As Object = CreateObject("WScript.Shell")
                Dim shortcut As Object = wsh.CreateShortcut(shortcutPath)

                ' Set the shortcut target to the current file
                shortcut.TargetPath = currentFilePath
                shortcut.WorkingDirectory = Path.GetDirectoryName(currentFilePath)
                shortcut.Save() ' Save the shortcut

                Console.WriteLine("Shortcut created successfully at: " & shortcutPath)
            Catch ex As Exception
                ' Handle any errors
                Console.WriteLine("Error: " & ex.Message)
            End Try
        End If
        AddStartup.Checked = Creat
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        KillAPP("Excel")
    End Sub

    Private Sub KillAPP(processName As String)
        Dim processes As Process() = Process.GetProcessesByName(processName)

        ' Loop through the processes and kill each one
        For Each proc As Process In processes
            Try
                proc.Kill()
                proc.WaitForExit() ' Wait for the process to exit
                'MsgBox($"Process {proc.ProcessName} with ID {proc.Id} has been terminated.")
            Catch ex As Exception
                'MsgBox($"Failed to terminate process {proc.ProcessName} with ID {proc.Id}. Error: {ex.Message}")
            End Try
        Next
    End Sub

    Private Function CheckProgramOpen(processName As String) As Integer
        Dim processes() As Process = Process.GetProcessesByName(processName)

        If processes.Length > 0 Then
            Console.WriteLine($"{processName} is running.")
        Else
            Console.WriteLine($"{processName} is not running.")
        End If

        Return processes.Length
    End Function


    Private Sub BTexit_Click(sender As Object, e As EventArgs) Handles BTexit.Click
        Me.Hide()
        'Me.Close()
    End Sub

    Private Sub BTSetting_Click(sender As Object, e As EventArgs) Handles BTSetting.Click

    End Sub

    Private Sub MainForm_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles BTheader.MouseDown
        If e.Button = MouseButtons.Left Then
            isDragging = True
            startPoint = New Point(e.X, e.Y)
        End If
    End Sub

    Private Sub MainForm_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles BTheader.MouseMove
        If isDragging Then
            Dim p As Point = PointToScreen(e.Location)
            Location = New Point(p.X - startPoint.X, p.Y - startPoint.Y)
        End If
    End Sub

    Private Sub MainForm_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs) Handles BTheader.MouseUp
        If e.Button = MouseButtons.Left Then
            isDragging = False
        End If
    End Sub

    Private Sub MainForm_MouseEnter(ByVal sender As Object, ByVal e As EventArgs)

    End Sub

    Private Sub MainForm_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles BTheader.MouseLeave
        Dim clientRect As New Rectangle(0, 0, Me.ClientSize.Width, Me.ClientSize.Height)
        If Not clientRect.Contains(Me.PointToClient(Cursor.Position)) Then

        End If
    End Sub

    ' Form closing event to ensure cleanup
    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If NotifyIcon1 IsNot Nothing Then
            NotifyIcon1.Visible = False
            NotifyIcon1.Dispose()
        End If
    End Sub

End Class
