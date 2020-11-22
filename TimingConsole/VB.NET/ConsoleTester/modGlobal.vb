Imports System
Imports Microsoft.Win32
Imports System.IO

Module modGlobal

    Public Const REGPATH As String = "Software\Gary Wood Software\ConsoleTester"

    Public InterfaceFields As WinswimCTS.ColoradoCTS.Interface_Renamed

    Public Function RegSetValue(ByVal tPath As String, ByVal tKey As String, ByVal tVal As Integer) As Boolean
        Dim regKey As RegistryKey
        RegSetValue = False
        Try
            regKey = Registry.CurrentUser.OpenSubKey(tPath, True)
            If regKey Is Nothing Then
                RegSetValue = False
            Else
                RegSetValue = True
                regKey.SetValue(tKey, tVal)
                regKey.Close()
            End If
        Catch ex As Exception
            WriteConsole(ex.Message)
        End Try
    End Function

    Public Function RegGetValue(ByVal tPath As String, ByVal tKey As String, ByVal tDefault As Integer) As Integer
        Dim regKey As RegistryKey
        Try
            regKey = Registry.CurrentUser.OpenSubKey(tPath, False)
            If regKey Is Nothing Then
                RegGetValue = -1
            Else
                RegGetValue = regKey.GetValue(tKey, tDefault)
                regKey.Close()
            End If
        Catch ex As Exception
            WriteConsole(ex.Message)
        End Try
    End Function
    Public Function SubKeyExist(ByVal tPath As String) As Boolean
        Dim regKey As RegistryKey
        SubKeyExist = False
        Try
            regKey = Registry.CurrentUser.OpenSubKey(tPath, False)
            If regKey Is Nothing Then
                SubKeyExist = False
            Else
                SubKeyExist = True
                regKey.Close()
            End If
        Catch ex As Exception
            WriteConsole(ex.Message)
        End Try
    End Function

    Public Function CreateSubKey(ByVal tPath As String, ByVal tKey As String) As Boolean
        Dim regKey As RegistryKey
        Try
            regKey = Registry.CurrentUser.OpenSubKey(tPath, True)
            If regKey Is Nothing Then
                CreateSubKey = False
            Else
                CreateSubKey = True
                regKey.CreateSubKey(tKey)
                regKey.Close()
            End If
        Catch ex As Exception
            WriteConsole(ex.Message)
        End Try
    End Function

    Public Function CSngTime(ByRef tStr As String) As Single
        Dim i As Short

        On Error GoTo CSngErr

        For i = 1 To Len(tStr)
            If InStr(1, " 0123456789:.", Mid(tStr, i, 1), CompareMethod.Binary) = 0 Then
                CSngTime = 0
                Exit Function
            End If
        Next i

        If InStr(1, tStr, ":") <> 0 Then
            CSngTime = CSng(Left(tStr, InStr(1, tStr, ":") - 1)) * 100
            CSngTime = CSngTime + CSng(Right(tStr, Len(tStr) - InStr(1, tStr, ":")))
        Else
            CSngTime = CSng(tStr)
        End If
        Exit Function

CSngErr:
        WriteConsole("CSngTime(" & tStr & ")")
        CSngTime = 0

    End Function

    Public Sub ClearConsole()
        frmConsole.txtConsole.Clear()
    End Sub

    Public Sub WriteConsole(ByVal tstr As String, Optional ByVal bCRLF As Boolean = True)
        frmConsole.txtConsole.AppendText(tstr & IIf((bCRLF = True), vbCrLf, ""))
    End Sub

    Public Function CTSWhoAreYou() As String
        CTSWhoAreYou = ""
        Try
            ClearConsole()
            Dim _cts As WinswimCTS.ColoradoCTS = New WinswimCTS.ColoradoCTS()
            If _cts Is Nothing Then
                WriteConsole(String.Format("Verify that {0} is registered using REGSVR32.EXE.", Path.Combine(My.Application.Info.DirectoryPath, "WinSwimCTS.dll")))
            Else
                WriteConsole("Opening console...")
                WriteConsole(String.Format("Connection String: COM{0} {1}", CStr(InterfaceFields.Port + 1), _cts.CTSGetSerialString(InterfaceFields)))
                If _cts.CTSOpen(InterfaceFields) = 0 Then
                    CTSWhoAreYou = _cts.CTSWhoAreYou()
                    If CTSWhoAreYou.Length > 0 Then
                        WriteConsole("Successful... ")
                        WriteConsole("CTS Version is " & CTSWhoAreYou)
                    Else
                        WriteConsole("Failed connection on COM" & InterfaceFields.Port + 1)
                    End If
                    _cts.CTSClose(InterfaceFields)
                End If
            End If

        Catch ex As Exception
            WriteConsole("Open failed : " & ex.Message)
        End Try
    End Function

    Public Sub SetInterface(ByRef tInf As WinswimCTS.ColoradoCTS.Interface_Renamed)
        tInf.Port = RegGetValue(REGPATH, "Port", 0)
        tInf.Speed = RegGetValue(REGPATH, "Speed", 1)
        tInf.Parity = RegGetValue(REGPATH, "Parity", 3)
        tInf.DataBits = RegGetValue(REGPATH, "Databits", 4)
        tInf.StopBits = RegGetValue(REGPATH, "Stopbits", 1)
        tInf.Sleep = RegGetValue(REGPATH, "Wait", 1)
        tInf.EventSequence = RegGetValue(REGPATH, "EventSeqType", 7)
    End Sub
End Module
