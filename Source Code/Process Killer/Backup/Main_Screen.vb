Imports System.IO
Imports Microsoft.Win32

Public Class Main_Screen
    Inherits System.Windows.Forms.Form

    Dim WithEvents Worker1 As Worker

    Private workerbusy As Boolean = False
    Private steps As Integer = 0
    'steps 0: process not launched
    'steps 1: line count

    Public Delegate Sub WorkerComplete_h()
    Public Delegate Sub WorkerError_h(ByVal Message As Exception, ByVal identifier As String)
    Public Delegate Sub WorkerFileCount_h(ByVal Result As Long)
    Public Delegate Sub WorkerStatusMessage_h(ByVal message As String, ByVal statustag As Integer)
    Public Delegate Sub WorkerFileProcessing_h(ByVal filename As String, ByVal queue As Integer)

    Private application_exit As Boolean = False
    Private shutting_down As Boolean = False
    Private splash_loader As Splash_Screen
    Public dataloaded As Boolean = False
    Private error_reporting_level

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerCompleteHandler
        AddHandler Worker1.WorkerError, AddressOf WorkerErrorHandler
        AddHandler Worker1.WorkerFileCount, AddressOf WorkerFileCountHandler
        AddHandler Worker1.WorkerStatusMessage, AddressOf WorkerStatusMessageHandler
        AddHandler Worker1.WorkerFileProcessing, AddressOf WorkerFileProcessingHandler

    End Sub

    Public Sub New(ByVal splash As Splash_Screen)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        splash_loader = splash
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerCompleteHandler
        AddHandler Worker1.WorkerError, AddressOf WorkerErrorHandler
        AddHandler Worker1.WorkerFileCount, AddressOf WorkerFileCountHandler
        AddHandler Worker1.WorkerStatusMessage, AddressOf WorkerStatusMessageHandler
        AddHandler Worker1.WorkerFileProcessing, AddressOf WorkerFileProcessingHandler
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtBaseProcess As System.Windows.Forms.TextBox
    Friend WithEvents FullError As System.Windows.Forms.CheckBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents txtProcessLaunched As System.Windows.Forms.Label
    Friend WithEvents ButtonOperationLaunch As System.Windows.Forms.Button
    Friend WithEvents txtStatus As System.Windows.Forms.TextBox
    Friend WithEvents txtFileCount As System.Windows.Forms.Label
    Friend WithEvents txtFolderCount As System.Windows.Forms.Label
    Friend WithEvents txtStatusBar1 As System.Windows.Forms.TextBox
    Friend WithEvents Button_Pause1 As System.Windows.Forms.Button
    Friend WithEvents Button_ExitThread As System.Windows.Forms.Button
    Friend WithEvents txtFileCount2 As System.Windows.Forms.Label
    Friend WithEvents txtFileCount3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Main_Screen))
        Me.txtBaseProcess = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.FullError = New System.Windows.Forms.CheckBox
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ButtonOperationLaunch = New System.Windows.Forms.Button
        Me.Button_Pause1 = New System.Windows.Forms.Button
        Me.Button_ExitThread = New System.Windows.Forms.Button
        Me.txtProcessLaunched = New System.Windows.Forms.Label
        Me.txtStatus = New System.Windows.Forms.TextBox
        Me.txtFileCount = New System.Windows.Forms.Label
        Me.txtFolderCount = New System.Windows.Forms.Label
        Me.txtStatusBar1 = New System.Windows.Forms.TextBox
        Me.txtFileCount2 = New System.Windows.Forms.Label
        Me.txtFileCount3 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'txtBaseProcess
        '
        Me.txtBaseProcess.BackColor = System.Drawing.Color.White
        Me.txtBaseProcess.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtBaseProcess.ForeColor = System.Drawing.Color.Black
        Me.txtBaseProcess.Location = New System.Drawing.Point(8, 56)
        Me.txtBaseProcess.Name = "txtBaseProcess"
        Me.txtBaseProcess.Size = New System.Drawing.Size(312, 20)
        Me.txtBaseProcess.TabIndex = 4
        Me.txtBaseProcess.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtBaseProcess, "The Process Names to be searched for")
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(8, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(304, 32)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "ENTER THE PROCESS NAMES BELOW AND HIT THE KILL BUTTON TO PROCEED. MULTIPLE PROCES" & _
        "SES CAN BE SPECIFIED USING THE ';;' CHAR SEQUENCE. * ACTS AS A WILDCARD CHARACTE" & _
        "R. "
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(8, 72)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(136, 16)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "PROCESSES TO BE KILLED"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'FullError
        '
        Me.FullError.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.FullError.Location = New System.Drawing.Point(312, 14)
        Me.FullError.Name = "FullError"
        Me.FullError.Size = New System.Drawing.Size(1, 1)
        Me.FullError.TabIndex = 9
        Me.ToolTip1.SetToolTip(Me.FullError, "If Checked, Error Handling Routine enters Full Exception Mode")
        Me.FullError.Visible = False
        '
        'ButtonOperationLaunch
        '
        Me.ButtonOperationLaunch.BackColor = System.Drawing.Color.MediumSeaGreen
        Me.ButtonOperationLaunch.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonOperationLaunch.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonOperationLaunch.Location = New System.Drawing.Point(8, 96)
        Me.ButtonOperationLaunch.Name = "ButtonOperationLaunch"
        Me.ButtonOperationLaunch.Size = New System.Drawing.Size(88, 20)
        Me.ButtonOperationLaunch.TabIndex = 10
        Me.ButtonOperationLaunch.Text = "Kill"
        Me.ToolTip1.SetToolTip(Me.ButtonOperationLaunch, "Launches Process Killer Operation")
        '
        'Button_Pause1
        '
        Me.Button_Pause1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.Button_Pause1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button_Pause1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_Pause1.Location = New System.Drawing.Point(104, 96)
        Me.Button_Pause1.Name = "Button_Pause1"
        Me.Button_Pause1.Size = New System.Drawing.Size(56, 20)
        Me.Button_Pause1.TabIndex = 17
        Me.Button_Pause1.Text = "PAUSE"
        Me.ToolTip1.SetToolTip(Me.Button_Pause1, "Pauses Process Killer Operation")
        Me.Button_Pause1.Visible = False
        '
        'Button_ExitThread
        '
        Me.Button_ExitThread.BackColor = System.Drawing.Color.LightSeaGreen
        Me.Button_ExitThread.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button_ExitThread.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_ExitThread.Location = New System.Drawing.Point(168, 96)
        Me.Button_ExitThread.Name = "Button_ExitThread"
        Me.Button_ExitThread.Size = New System.Drawing.Size(56, 20)
        Me.Button_ExitThread.TabIndex = 18
        Me.Button_ExitThread.Text = "EXIT"
        Me.ToolTip1.SetToolTip(Me.Button_ExitThread, "Exits Process Killer Operation")
        Me.Button_ExitThread.Visible = False
        '
        'txtProcessLaunched
        '
        Me.txtProcessLaunched.BackColor = System.Drawing.Color.MediumTurquoise
        Me.txtProcessLaunched.Location = New System.Drawing.Point(8, 128)
        Me.txtProcessLaunched.Name = "txtProcessLaunched"
        Me.txtProcessLaunched.Size = New System.Drawing.Size(312, 16)
        Me.txtProcessLaunched.TabIndex = 12
        Me.txtProcessLaunched.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtStatus
        '
        Me.txtStatus.BackColor = System.Drawing.Color.MediumTurquoise
        Me.txtStatus.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStatus.ForeColor = System.Drawing.Color.White
        Me.txtStatus.Location = New System.Drawing.Point(72, 176)
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.ReadOnly = True
        Me.txtStatus.Size = New System.Drawing.Size(248, 12)
        Me.txtStatus.TabIndex = 15
        Me.txtStatus.Text = ""
        Me.txtStatus.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtFileCount
        '
        Me.txtFileCount.Location = New System.Drawing.Point(288, 160)
        Me.txtFileCount.Name = "txtFileCount"
        Me.txtFileCount.Size = New System.Drawing.Size(32, 16)
        Me.txtFileCount.TabIndex = 16
        Me.txtFileCount.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtFolderCount
        '
        Me.txtFolderCount.BackColor = System.Drawing.Color.MediumTurquoise
        Me.txtFolderCount.ForeColor = System.Drawing.Color.Black
        Me.txtFolderCount.Location = New System.Drawing.Point(8, 160)
        Me.txtFolderCount.Name = "txtFolderCount"
        Me.txtFolderCount.Size = New System.Drawing.Size(304, 16)
        Me.txtFolderCount.TabIndex = 13
        Me.txtFolderCount.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtStatusBar1
        '
        Me.txtStatusBar1.BackColor = System.Drawing.Color.MediumTurquoise
        Me.txtStatusBar1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtStatusBar1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStatusBar1.ForeColor = System.Drawing.Color.White
        Me.txtStatusBar1.Location = New System.Drawing.Point(10, 144)
        Me.txtStatusBar1.Name = "txtStatusBar1"
        Me.txtStatusBar1.ReadOnly = True
        Me.txtStatusBar1.Size = New System.Drawing.Size(310, 13)
        Me.txtStatusBar1.TabIndex = 14
        Me.txtStatusBar1.Text = ""
        '
        'txtFileCount2
        '
        Me.txtFileCount2.BackColor = System.Drawing.Color.MediumTurquoise
        Me.txtFileCount2.Location = New System.Drawing.Point(8, 144)
        Me.txtFileCount2.Name = "txtFileCount2"
        Me.txtFileCount2.Size = New System.Drawing.Size(312, 16)
        Me.txtFileCount2.TabIndex = 19
        Me.txtFileCount2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtFileCount3
        '
        Me.txtFileCount3.Location = New System.Drawing.Point(280, 168)
        Me.txtFileCount3.Name = "txtFileCount3"
        Me.txtFileCount3.Size = New System.Drawing.Size(40, 16)
        Me.txtFileCount3.TabIndex = 20
        Me.txtFileCount3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(244, 1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 9)
        Me.Label1.TabIndex = 21
        Me.Label1.Text = "BUILD 20051220.1"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.ToolTip1.SetToolTip(Me.Label1, "Build Number")
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(8, 176)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 16)
        Me.Label2.TabIndex = 23
        Me.Label2.Text = "ACTIVITY LOGS"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ToolTip1.SetToolTip(Me.Label2, "View Activity Logs")
        '
        'Main_Screen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.MediumTurquoise
        Me.ClientSize = New System.Drawing.Size(328, 198)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtStatus)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtFileCount3)
        Me.Controls.Add(Me.txtFileCount2)
        Me.Controls.Add(Me.Button_ExitThread)
        Me.Controls.Add(Me.Button_Pause1)
        Me.Controls.Add(Me.txtFileCount)
        Me.Controls.Add(Me.txtStatusBar1)
        Me.Controls.Add(Me.txtBaseProcess)
        Me.Controls.Add(Me.txtFolderCount)
        Me.Controls.Add(Me.txtProcessLaunched)
        Me.Controls.Add(Me.ButtonOperationLaunch)
        Me.Controls.Add(Me.FullError)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.ForeColor = System.Drawing.Color.Black
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(336, 232)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(336, 232)
        Me.Name = "Main_Screen"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Process Killer"
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message("The Application encountered the following problem: " & vbCrLf & identifier_msg & ":" & ex.ToString)
                Display_Message1.ShowDialog()
                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg & ":" & ex.ToString)
                filewriter.Flush()
                filewriter.Close()
            End If
        Catch exc As Exception
            MsgBox("An error occurred in Process Killer's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub Activity_Logger(ByVal message As String)
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & Format(Now(), "yyyyMMdd") & "_Activity_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & message)
            filewriter.Flush()
            filewriter.Close()
        Catch ex As Exception
            Error_Handler(ex, "Activity Logger")
        End Try
    End Sub






    Private Sub SendMessage(ByVal labelname As String, ByVal message As String)
        Try
            Dim controllist As ControlCollection = Me.Controls
            Dim cont As Control

            For Each cont In controllist
                If cont.Name = labelname Then
                    cont.Text = message
                    cont.Refresh()
                    Exit For
                End If
            Next
        Catch ex As Exception
            If (Not Worker1.WorkerThread.ThreadState.ToString.IndexOf("Aborted") > -1) And (Not Worker1.WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1) Then
                Error_Handler(ex, "Send Message")
            End If
        End Try
    End Sub

    Private Sub ButtonOperationLaunch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonOperationLaunch.Click
        ButtonOperationLaunch.Enabled = False
        SendMessage("txtProcessLaunched", "")
        SendMessage("txtStatus", "")
        SendMessage("txtStatusBar1", "")
        SendMessage("txtFolderCount", "")
        SendMessage("txtFileCount", "")

        SendMessage("txtProcessLaunched", "Launched: " & Format(Now(), "dd/MM/yyyy HH:mm:ss"))
        Worker1.processname1 = txtBaseProcess.Text

        steps = 1
        workerbusy = True
        Worker1.ChooseThreads(1)
        Button_Pause1.Visible = True
        Button_Pause1.Refresh()

    End Sub

    Public Sub WorkerStatusMessageHandler(ByVal message As String, ByVal statustag As Integer)
        Try
            If statustag = 1 Then
                SendMessage("txtStatus", message)
            Else
                SendMessage("txtStatusBar1", message)
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerErrorHandler(ByVal Message As Exception, ByVal identifier As String)
        Try
            Error_Handler(Message, "Worker: " & identifier)
        Catch ex As Exception
            Error_Handler(ex, "Worker Error Handler")
        End Try
    End Sub


    Public Sub WorkerFileCountHandler(ByVal Result As Long, ByVal line As Integer)
        Try
            Select Case line
                Case 1
                    If (Not Worker1.WorkerThread.ThreadState.ToString.IndexOf("Aborted") > -1) And (Not Worker1.WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1) Then
                        SendMessage("txtFileCount", "Total Line Count: " & Result.ToString)
                    End If
                Case 2
                    If (Not Worker1.WorkerThread.ThreadState.ToString.IndexOf("Aborted") > -1) And (Not Worker1.WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1) Then
                        SendMessage("txtFileCount2", "Processes Killed: " & Result.ToString)
                    End If
                Case 3
                    If (Not Worker1.WorkerThread.ThreadState.ToString.IndexOf("Aborted") > -1) And (Not Worker1.WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1) Then
                        SendMessage("txtFileCount3", "Blank Line Count: " & Result.ToString)
                    End If
            End Select

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerCompleteHandler(ByVal queue As Integer)
        Try
            Dim eventhandled As Boolean = False

            Button_Pause1.Visible = False
            Button_Pause1.Refresh()
            workerbusy = False
            ButtonOperationLaunch.Enabled = True
            SendMessage("txtFolderCount", "Completed: " & Format(Now(), "dd/MM/yyyy HH:mm:ss"))
           
            eventhandled = True
            txtStatusBar1.Select(0, 0)
            txtBaseProcess.Select(0, 0)
            txtBaseProcess.Select()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerFileProcessingHandler(ByVal filename As String, ByVal queue As Integer)
        Try


        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Main_Screen_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Try
            Worker1.Dispose()
        Catch ex As Exception
            Error_Handler(ex)
        End Try

    End Sub

    Private Sub Button_Pause_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Pause1.Click
        Try

            If workerbusy = True Then
                Worker1.WorkerThread.Suspend()
                Button_Pause1.Text = "RESUME"
                Button_Pause1.Refresh()

                Button_ExitThread.Visible = True
                Button_ExitThread.Refresh()
                workerbusy = False
            Else

                Worker1.WorkerThread.Resume()
                Button_Pause1.Text = "PAUSE"
                Button_Pause1.Refresh()

                Button_ExitThread.Visible = False
                Button_ExitThread.Refresh()
                workerbusy = True
            End If

        Catch ex As Exception

            Error_Handler(ex)

        End Try
    End Sub

    Private Sub exit_threads()
        Try

            If Worker1.WorkerThread.ThreadState.ToString.IndexOf("Suspended") > -1 Or Worker1.WorkerThread.ThreadState.ToString.IndexOf("SuspendRequested") > -1 Then
                Worker1.WorkerThread.Resume()
            End If

            If Worker1.WorkerThread.ThreadState.ToString.IndexOf("WaitSleepJoin") > -1 Then
                Worker1.WorkerThread.Interrupt()
            End If

            If Worker1.WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1 Then
                Worker1.WorkerThread.ResetAbort()
            End If

            ' SendMessage("txtStatus", "Aborting worker thread")
            Worker1.WorkerThread.Abort()
            SendMessage("txtFolderCount", "Aborted: " & Format(Now(), "dd/MM/yyyy HH:mm:ss"))

            SendMessage("txtStatus", "Worker thread aborted")
            Button_ExitThread.Visible = False
            Button_ExitThread.Refresh()
            Button_Pause1.Visible = False
            Button_Pause1.Text = "PAUSE"
            ButtonOperationLaunch.Enabled = True

            'Worker1.ChooseThreads(2)

            'Worker1.Dispose()
            'Worker1 = New Worker
            'AddHandler Worker1.WorkerComplete, AddressOf WorkerCompleteHandler
            'AddHandler Worker1.WorkerError, AddressOf WorkerErrorHandler
            'AddHandler Worker1.WorkerFileCount, AddressOf WorkerFileCountHandler
            'AddHandler Worker1.WorkerStatusMessage, AddressOf WorkerStatusMessageHandler
            'AddHandler Worker1.WorkerFileProcessing, AddressOf WorkerFileProcessingHandler

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Button_ExitThread_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_ExitThread.Click
        exit_threads()
    End Sub

    Private Sub Main_Screen_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Save_Registry_Values()
        Worker1.Dispose()
        Application.Exit()
    End Sub

    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Load_Registry_Values()
        dataloaded = True
        splash_loader.Visible = False
    End Sub

    Public Sub Load_Registry_Values()
        Try
            Dim configflag As Boolean
            configflag = False
            Dim str As String
            Dim keyflag1 As Boolean = False
            Dim oReg As RegistryKey = Registry.LocalMachine
            Dim keys() As String = oReg.GetSubKeyNames()
            System.Array.Sort(keys)

            For Each str In keys
                If str.Equals("Software\Process Killer") = True Then
                    keyflag1 = True
                    Exit For
                End If
            Next str

            If keyflag1 = False Then
                oReg.CreateSubKey("Software\Process Killer")
            End If

            keyflag1 = False

            Dim oKey As RegistryKey = oReg.OpenSubKey("Software\Process Killer", True)

            str = oKey.GetValue("processname")
            If Not IsNothing(str) And Not (str = "") Then
                txtBaseProcess.Text = str
            Else
                configflag = True
                oKey.SetValue("processname", "")
                txtBaseProcess.Text = ""
            End If

            oKey.Close()
            oReg.Close()

        Catch ex As Exception
            Error_Handler(ex, "Load Registry Values")
        End Try
    End Sub

    Private Sub Save_Registry_Values()
        Try
            Dim oReg As RegistryKey = Registry.LocalMachine
            Dim oKey As RegistryKey = oReg.OpenSubKey("Software\Process Killer", True)

            oKey.SetValue("processname", txtBaseProcess.Text)

            oKey.Close()
            oReg.Close()
        Catch ex As Exception

            Error_Handler(ex, "Save Registry Values")

        End Try
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            DosShellCommand("explorer """ & (Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs" & """")
        Catch ex As Exception
            Error_Handler(ex, "View Activity Logs")
        End Try
    End Sub

    

    Private Function DosShellCommand(ByVal AppToRun As String) As String
        Dim s As String = ""
        Try
            Dim myProcess As Process = New Process

            myProcess.StartInfo.FileName = "cmd.exe"
            myProcess.StartInfo.UseShellExecute = False
            myProcess.StartInfo.CreateNoWindow = True

            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.StartInfo.RedirectStandardError = True
            myProcess.Start()
            Dim sIn As StreamWriter = myProcess.StandardInput
            sIn.AutoFlush = True

            Dim sOut As StreamReader = myProcess.StandardOutput
            Dim sErr As StreamReader = myProcess.StandardError
            sIn.Write(AppToRun & _
               System.Environment.NewLine)
            sIn.Write("exit" & System.Environment.NewLine)
            s = sOut.ReadToEnd()
            If Not myProcess.HasExited Then
                myProcess.Kill()
            End If

            'MessageBox.Show("The 'dir' command window was closed at: " & myProcess.ExitTime & "." & System.Environment.NewLine & "Exit Code: " & myProcess.ExitCode)

            sIn.Close()
            sOut.Close()
            sErr.Close()
            myProcess.Close()
            myProcess.Dispose()
            'MessageBox.Show(s)
        Catch ex As Exception
            Error_Handler(ex, "DosShellCommand")
        End Try
        Return s
    End Function


End Class
