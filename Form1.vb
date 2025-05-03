Imports System.IO
Imports System.Reflection

Public Class Form1
    ' Create a NotifyIcon
    Private WithEvents notifyIcon As New NotifyIcon()

    Private startupFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.Startup)
    'Private currentFolder As String = Directory.GetCurrentDirectory()
    Private currentFilePath As String = Assembly.GetExecutingAssembly().Location
    Private fileName As String = Path.GetFileName(currentFilePath)

    Private isDragging As Boolean = False
    Private startPoint As Point

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Initialize the form
        ' InitializeComponent()

        ' Configure the NotifyIcon
        ' NotifyIcon1.Icon = SystemIcons.Application ' Use a default system icon
        NotifyIcon1.Text = "Office_Utilities_Tools" ' Text displayed when hovering over the icon
        NotifyIcon1.Visible = True

        ' Add a context menu to the NotifyIcon (optional)
        Dim contextMenu As New ContextMenuStrip()
        Dim exitItem As New ToolStripMenuItem("Exit", Nothing, AddressOf ExitApplication)
        Dim exitItem1 As New ToolStripMenuItem("End Task Exel", Nothing, AddressOf ExitExel)
        contextMenu.Items.Add(exitItem1)
        contextMenu.Items.Add(exitItem)
        NotifyIcon1.ContextMenuStrip = contextMenu

        CreatShortcutStartup()
        'Me.StartPosition = FormStartPosition.CenterScreen
    End Sub

    ' Handle double-click on the NotifyIcon to restore the form
    Private Sub notifyIcon_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
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

    Private Sub CreatShortcutStartup()
        ' Define the shortcut file name and path
        Dim shortcutPath As String = Path.Combine(startupFolder, Path.GetFileNameWithoutExtension(currentFilePath) & ".lnk")
        If Not File.Exists(shortcutPath) Then
            'File.Delete(shortcutPath) ' Delete the existing shortcut if it exists
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

End Class
