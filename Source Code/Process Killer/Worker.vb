Imports System.IO
Imports System.Text


Public Class Worker

    Inherits System.ComponentModel.Component

    ' Declares the variables you will use to hold your thread objects.

    Public WorkerThread As System.Threading.Thread



    Public processname1 As String
    'Private filecount As Long
    'Private blankcount As Long
    Private fullcount As Long
    Private filereader As StreamReader
    Private filewriter As StreamWriter

    Public Event WorkerFileProcessing(ByVal filename As String, ByVal queue As Integer)
    Public Event WorkerStatusMessage(ByVal message As String, ByVal statustag As Integer)
    Public Event WorkerError(ByVal Message As Exception, ByVal identifier As String)
    Public Event WorkerFileCount(ByVal Result As Long, ByVal count As Integer)
    Public Event WorkerComplete(ByVal queue As Integer)




#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)

    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub Error_Handler(ByVal message As Exception, ByVal identifier As String)
        Try
            If (Not WorkerThread.ThreadState.ToString.IndexOf("Aborted") > -1) And (Not WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1) Then
                RaiseEvent WorkerError(message, identifier)
            End If
        Catch ex As Exception
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


    Public Sub ChooseThreads(ByVal threadNumber As Integer)
        Try
            ' Determines which thread to start based on the value it receives.
            Select Case threadNumber
                Case 1
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerExecute)
                    WorkerThread.Start()
                    'Case 2
                    'filereader.Close()
            End Select
        Catch ex As Exception
            Error_Handler(ex, "Choose Threads")
        End Try
    End Sub

    Private Sub WorkerExecute()
        RaiseEvent WorkerStatusMessage("Running Process Killer", 1)

        fullcount = 0



        RaiseEvent WorkerFileCount(fullcount, 2)


        Dim processes As String() = processname1.Split(";;")
        Dim proc As String
        For Each proc In processes
            If Not proc = "" And Not proc Is Nothing Then
                Dim existing As Process() = Process.GetProcesses
                Dim eproc As Process
                For Each eproc In existing
                    Try

                    
                    Dim eprocname As String = eproc.ProcessName
                    Dim handled As Boolean = False
                    If proc.EndsWith("*") And handled = False Then
                        If eproc.ProcessName.ToLower.StartsWith(proc.ToLower.Remove(proc.Length - 1, 1)) = True Then
                            If eproc.CloseMainWindow = False Then
                                eproc.Kill()
                            End If
                            handled = True
                        End If
                    End If
                    If proc.StartsWith("*") And handled = False Then
                        If eproc.ProcessName.ToLower.EndsWith(proc.ToLower.Remove(0, 1)) = True Then
                            If eproc.CloseMainWindow = False Then
                                eproc.Kill()
                            End If
                            handled = True
                        End If
                    End If
                    If handled = False Then
                        If eproc.ProcessName.ToLower.Equals(proc.ToLower) = True Then
                            If eproc.CloseMainWindow = False Then
                                eproc.Kill()
                            End If
                            handled = True
                        End If
                    End If
                    eproc.Dispose()
                    If handled = True Then
                        Activity_Logger("Closed " & eprocname)
                        fullcount = fullcount + 1
                        RaiseEvent WorkerFileCount(fullcount, 2)
                    End If
                    Catch ex As Exception
                        Error_Handler(ex, "Attempting Process Killing")
                    End Try
                Next

            End If
        Next

        RaiseEvent WorkerFileCount(fullcount, 2)


        RaiseEvent WorkerStatusMessage("Process Killing Complete", 1)
        RaiseEvent WorkerComplete(0)
    End Sub


   

   

End Class
