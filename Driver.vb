'tabs=4
' --------------------------------------------------------------------------------
'
' ASCOM SafetyMonitor driver V2 for TriStar Observatory
'
' Description:	New safety driver to read the output from Christian Kardach's
'				AllSky AI, an AI that makes Cloud/Rain determinations from AllSky images.
'               https://allskyai.kardach.se/ for more.
'
'
' Implements:	ASCOM SafetyMonitor interface version: 1.0
' Author:		(EOR) eorequis@tristarobservatory.com
'
' Edit Log:
'
' Date			Who	Vers	Description
' -----------	---	-----	-------------------------------------------------------
' 2022-12-25	EOR	2.0.0	Initial edit, from SafetyMonitor template
' 2023-01-27	EOR	2.0.1	Clean up, remove all weather references, AllSkyAI only
' ---------------------------------------------------------------------------------

#Const Device = "SafetyMonitor"

Imports System.Net


Imports ASCOM.DeviceInterface
Imports ASCOM.Utilities

Imports Newtonsoft.Json


<Guid("9d044a62-f0fc-413f-b1f7-e80f51cbff4d")>
<ClassInterface(ClassInterfaceType.None)>
Public Class SafetyMonitor

    ' The Guid attribute sets the CLSID for ASCOM.TriStar.SafetyMonitor
    ' The ClassInterface/None attribute prevents an empty interface called
    ' _TriStar from being created and used as the [default] interface

    ' TODO Replace the not implemented exceptions with code to implement the function or
    ' throw the appropriate ASCOM exception.
    '
    Implements ISafetyMonitor

    ' Driver info
    Friend Shared driverID As String = "ASCOM.TriStar.SafetyMonitor"
    Private Shared driverDescription As String = "TriStar SafetyMonitor"
    Private majorVersion As String = "2"
    Private minorVersion As String = "0"
    Private buildVersion As String = "1"

    'Constants used for Profile persistence
    Friend Shared URLProfileName As String = "URL"
    Friend Shared traceStateProfileName As String = "Trace Level"
    Friend Shared URLDefault As String = "https://allskyai.kardach.se/allskyai/v1/live?allsky=tristar"
    Friend Shared traceStateDefault As String = "False"

    ' Variables to hold the current driver configuration
    Friend Shared URL As String
    Friend Shared traceState As Boolean

    Private connectedState As Boolean
    Private TL As TraceLogger

    Private AllSky As AllSkyAI
    Private SafetyTimer As System.Timers.Timer
    Private csFailCount As Integer = 0

    Private LastWrite As DateTime
    Private Alert As Integer
    Private AllSkyAISky As String
    Private AllSkyAIConfidence As Double
    Private UTC As Double
    Private LastAI As DateTime
    Private nDateTime As System.DateTime = New System.DateTime(1970, 1, 1, 0, 0, 0, 0)
    Private AISkySafe As Boolean


    '
    ' Constructor - Must be public for COM registration!
    '
    Public Sub New()

        ReadProfile() ' Read device configuration from the ASCOM Profile store
        TL = New TraceLogger("", "TriStar")
        TL.Enabled = traceState
        TL.LogMessage("SafetyMonitor", "Starting initialisation")

        connectedState = False ' Initialise connected to false
        SafetyTimer = New Timers.Timer
        SafetyTimer.Interval = 30000
        SafetyTimer.AutoReset = True
        AddHandler SafetyTimer.Elapsed, AddressOf Timer_Tick

        TL.LogMessage("SafetyMonitor", "Completed initialisation")
    End Sub

    '
    ' PUBLIC COM INTERFACE ISafetyMonitor IMPLEMENTATION
    '

#Region "Common properties and methods"
    ''' <summary>
    ''' Displays the Setup Dialog form.
    ''' If the user clicks the OK button to dismiss the form, then
    ''' the new settings are saved, otherwise the old values are reloaded.
    ''' THIS IS THE ONLY PLACE WHERE SHOWING USER INTERFACE IS ALLOWED!
    ''' </summary>
    Public Sub SetupDialog() Implements ISafetyMonitor.SetupDialog
        ' consider only showing the setup dialog if not connected
        ' or call a different dialog if connected
        If IsConnected Then
            System.Windows.Forms.MessageBox.Show("Already connected, just press OK")
        End If

        Using F As SetupDialogForm = New SetupDialogForm()
            Dim result As System.Windows.Forms.DialogResult = F.ShowDialog()
            If result = DialogResult.OK Then
                WriteProfile() ' Persist device configuration values to the ASCOM Profile store
            End If
        End Using
    End Sub

    Public ReadOnly Property SupportedActions() As ArrayList Implements ISafetyMonitor.SupportedActions
        Get
            TL.LogMessage("SupportedActions Get", "Returning empty arraylist")
            Return New ArrayList()
        End Get
    End Property

    Public Function Action(ByVal ActionName As String, ByVal ActionParameters As String) As String Implements ISafetyMonitor.Action
        Throw New ActionNotImplementedException("Action " & ActionName & " is not supported by this driver")
    End Function

    Public Sub CommandBlind(ByVal Command As String, Optional ByVal Raw As Boolean = False) Implements ISafetyMonitor.CommandBlind
        CheckConnected("CommandBlind")
        ' TODO The optional CommandBlind method should either be implemented OR throw a MethodNotImplementedException
        ' If implemented, CommandBlind must send the supplied command to the mount And return immediately without waiting for a response

        Throw New MethodNotImplementedException("CommandBlind")
    End Sub

    Public Function CommandBool(ByVal Command As String, Optional ByVal Raw As Boolean = False) As Boolean _
        Implements ISafetyMonitor.CommandBool
        CheckConnected("CommandBool")
        ' TODO The optional CommandBool method should either be implemented OR throw a MethodNotImplementedException
        ' If implemented, CommandBool must send the supplied command to the mount, wait for a response and parse this to return a True Or False value

        ' Dim retString as String = CommandString(command, raw) ' Send the command And wait for the response
        ' Dim retBool as Boolean = XXXXXXXXXXXXX ' Parse the returned string And create a boolean True / False value
        ' Return retBool ' Return the boolean value to the client

        Throw New MethodNotImplementedException("CommandBool")
    End Function

    Public Function CommandString(ByVal Command As String, Optional ByVal Raw As Boolean = False) As String _
        Implements ISafetyMonitor.CommandString
        CheckConnected("CommandString")
        ' TODO The optional CommandString method should either be implemented OR throw a MethodNotImplementedException
        ' If implemented, CommandString must send the supplied command to the mount and wait for a response before returning this to the client

        Throw New MethodNotImplementedException("CommandString")
    End Function

    Public Property Connected() As Boolean Implements ISafetyMonitor.Connected
        Get
            TL.LogMessage("Connected Get", IsConnected.ToString())
            Return IsConnected
        End Get
        Set(value As Boolean)
            TL.LogMessage("Connected Set", value.ToString())
            If value = IsConnected Then
                Return
            End If

            If value Then
                If CheckFile(URL) = True Then
                    AllSky = New AllSkyAI
                    checkSafety()
                    connectedState = True
                    SafetyTimer.Enabled = True

                Else
                    Throw New DriverException("Connection to weather file failed")
                    connectedState = False
                    SafetyTimer.Enabled = False
                    AllSky = Nothing
                    TL.LogMessage("Connected Set ERROR", "Failed to connect")
                End If
            Else
                connectedState = False
                TL.LogMessage("Connected Set", "Disconnecting")
                ' TODO disconnect from the device
            End If
        End Set
    End Property

    Public ReadOnly Property Description As String Implements ISafetyMonitor.Description
        Get
            ' this pattern seems to be needed to allow a public property to return a private field
            Dim d As String = driverDescription
            TL.LogMessage("Description Get", d)
            Return d
        End Get
    End Property

    Public ReadOnly Property DriverInfo As String Implements ISafetyMonitor.DriverInfo
        Get
            Dim m_version As Version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
            ' TODO customise this driver description
            Dim s_driverInfo As String = "Information about the driver itself. Version: " + m_version.Major.ToString() + "." + m_version.Minor.ToString()
            TL.LogMessage("DriverInfo Get", s_driverInfo)
            Return s_driverInfo
        End Get
    End Property

    Public ReadOnly Property DriverVersion() As String Implements ISafetyMonitor.DriverVersion
        Get
            ' Get our own assembly and report its version number
            TL.LogMessage("DriverVersion Get", Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString(2))
            Return Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString(2)
        End Get
    End Property

    Public ReadOnly Property InterfaceVersion() As Short Implements ISafetyMonitor.InterfaceVersion
        Get
            TL.LogMessage("InterfaceVersion Get", "1")
            Return 1
        End Get
    End Property

    Public ReadOnly Property Name As String Implements ISafetyMonitor.Name
        Get
            Dim s_name As String = "Short driver name - please customise"
            TL.LogMessage("Name Get", s_name)
            Return s_name
        End Get
    End Property

    Public Sub Dispose() Implements ISafetyMonitor.Dispose
        ' Clean up the trace logger and util objects
        TL.Enabled = False
        TL.Dispose()
        TL = Nothing
    End Sub

#End Region

#Region "ISafetyMonitor Implementation"
    Public ReadOnly Property IsSafe() As Boolean Implements ISafetyMonitor.IsSafe
        Get
            If AISkySafe Then
                TL.LogMessage("IsSafe Get", "True")
                Return True
            Else
                TL.LogMessage("IsSafe Get", "False")
                Return False
            End If
        End Get
    End Property

#End Region

#Region "Private properties and methods"
    ' here are some useful properties and methods that can be used as required
    ' to help with

#Region "ASCOM Registration"

    Private Shared Sub RegUnregASCOM(ByVal bRegister As Boolean)

        Using P As New Profile() With {.DeviceType = "SafetyMonitor"}
            If bRegister Then
                P.Register(driverID, driverDescription)
            Else
                P.Unregister(driverID)
            End If
        End Using

    End Sub

    <ComRegisterFunction()>
    Public Shared Sub RegisterASCOM(ByVal T As Type)

        RegUnregASCOM(True)

    End Sub

    <ComUnregisterFunction()>
    Public Shared Sub UnregisterASCOM(ByVal T As Type)

        RegUnregASCOM(False)

    End Sub

#End Region

    ''' <summary>
    ''' Returns true if there is a valid connection to the driver hardware
    ''' </summary>
    Private ReadOnly Property IsConnected As Boolean
        Get
            ' TODO check that the driver hardware connection exists and is connected to the hardware
            Return connectedState
        End Get
    End Property

    ''' <summary>
    ''' Use this function to throw an exception if we aren't connected to the hardware
    ''' </summary>
    ''' <param name="message"></param>
    Private Sub CheckConnected(ByVal message As String)
        If Not IsConnected Then
            Throw New NotConnectedException(message)
        End If
    End Sub

    ''' <summary>
    ''' Read the device configuration from the ASCOM Profile store
    ''' </summary>
    Friend Sub ReadProfile()
        Using driverProfile As New Profile()
            driverProfile.DeviceType = "SafetyMonitor"
            traceState = Convert.ToBoolean(driverProfile.GetValue(driverID, traceStateProfileName, String.Empty, traceStateDefault))
            URL = driverProfile.GetValue(driverID, URLProfileName, String.Empty, URLDefault)
        End Using
    End Sub

    ''' <summary>
    ''' Write the device configuration to the  ASCOM  Profile store
    ''' </summary>
    Friend Sub WriteProfile()
        Using driverProfile As New Profile()
            driverProfile.DeviceType = "SafetyMonitor"
            driverProfile.WriteValue(driverID, traceStateProfileName, traceState.ToString())
            driverProfile.WriteValue(driverID, URLProfileName, URL)
        End Using

    End Sub

#End Region

#Region "My private functions and methods"
    Private Function CheckFile(AIUrl As String) As Boolean
        ' Checks to see if we can successfully retrieve the JSON
        Dim webclient As New WebClient, result As String = ""
        Try
            result = webclient.DownloadString(AIUrl)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub Timer_Tick(source As Object, e As EventArgs)
        checkSafety()
    End Sub

    Private Sub checkSafety()
        Try
            Dim webclient As New WebClient
            Dim response As String = webclient.DownloadString(URL)
            AllSky = JsonConvert.DeserializeObject(Of AllSkyAI)(response)
            AllSkyAISky = AllSky.AllSkyAISky
            AllSkyAIConfidence = CDbl(AllSky.AllSkyAIConfidence)
            UTC = CDbl(AllSky.UTC)
            LastAI = nDateTime.AddSeconds(UTC)

            TL.LogMessage("checkSafety", "Read JSON at " + DateTime.Now.ToUniversalTime.ToString("yyyy-MM-dd HH:mm:ss") + " UTC")
            TL.LogMessage("checkSafety", "LastAI is " + LastAI.ToString("yyyy-MM-dd HH:mm:ss") + " UTC")
            TL.LogMessage("checkSafety", "AllSkyAISky is " + AllSkyAISky)
            TL.LogMessage("checkSafety", "AllSkyAIConfidence is " + AllSkyAIConfidence.ToString)
            If AllSkyAISky = "clear" OrElse AllSkyAISky = "light_clouds" Then
                AISkySafe = True
            Else
                AISkySafe = False
            End If
            If DateDiff(DateInterval.Minute, LastAI, DateTime.Now.ToUniversalTime) > 5 Then
                AISkySafe = False
                TL.LogMessage("checkSafety ERROR", "Unsafe set, LastAI too old!")
            End If
            TL.LogMessage("checkSafety", "AI Sky Safe is " + AISkySafe.ToString)
            csFailCount = 0
        Catch ex As Exception
            csFailCount = csFailCount + 1
            If csFailCount > 1 Then
                AISkySafe = False
                TL.LogMessage("checkSafety ERROR", ex.Message)
            End If
        End Try
    End Sub

#End Region
End Class
