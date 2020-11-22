Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Buffer

<System.Runtime.InteropServices.ProgId("CommIO_NET.CommIO")> Public Class ColoradoCTS
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)

#Region "Structures"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Enum CTSEventType
        swimmerType = 0
        AgeGroup = 1
        EventType = 2
        bracket = 3
    End Enum

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Structure TYPE_MEETMGMT_HEADER
        'UPGRADE_NOTE: event was upgraded to event_Renamed. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim event_Renamed As Byte
        Dim heat As Byte
        Dim include_backup As Byte
        Dim include_splits As Byte
        Dim race_number As Short
        Dim date_seconds As Byte
        Dim date_minutes As Byte
        Dim date_hours As Byte
        Dim date_weekday As Byte
        Dim date_month As Byte
        Dim date_day As Byte
        Dim date_year As Short
        Dim race_lengths As Byte
        Dim lanes_in_pool As Byte
        Dim num_times_for_each_lane As Byte
        Dim date_long_year As Short
        Dim event_16_bit As Short
        Dim include_individual_buttons As Byte
        Dim include_relay_judging As Byte
        Dim number_of_buttons As Byte
        Dim date_date As Date
    End Structure

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Enum CTSLabelType
        EventNo = 0
        Distance = 1
        Sex = 2
        Stroke = 3
        AgeGroup = 4
        Session = 5
    End Enum

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Structure CTS_LABELS
        Dim checked As Boolean
        Dim value As Integer
        Dim label As CTSLabelType
    End Structure

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Structure TYPE_LANE_DATA
        Dim place As Short '  0=DNF -1=DQ 1-10=Place
        Dim time As Integer
        Dim mins As Integer
        Dim secs As Integer
        Dim hsecs As Integer
    End Structure

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Structure RACE_LIST_DATA
        Dim Race As Integer
        Dim EventNo As Integer
        Dim Heat As Integer
    End Structure

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Structure Interface_Renamed
        Dim Interface_Renamed As Short
        Dim Port As Short
        Dim Speed As Short
        Dim Parity As Short
        Dim DataBits As Short
        Dim StopBits As Short
        Dim Precision As Short
        Dim SerialString As String
        Dim Sleep As Short
        Dim EventSequence As Short
    End Structure

#End Region

#Region "Publics"

    Public SerialOb As WinswimCTS.CommIO
    Public Const CTS_EVENT_SEQ As Short = 7
    Public Const CTS_SLEEP_PAD As Short = 2

    Const MAXSTRINGSIZE As Short = 255
    Const SIZEOFTYPE_MEETMGMT_HEADER As Short = 49
    Const SIZEOFTYPE_LANE_DATA As Short = 7
    Const HEADER_DATE_OFFSET As Short = 8
    Const EVENT_MAX_EVENTS As Short = 500
    Const _CTSNORMAL As Short = 6
    Const _CTSERROR As Short = -1

    Dim CTSVERSION As String

#End Region
    ''' <summary>
    ''' Maximum number of events
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ReadOnly Property MaxEvents() As Integer
        Get
            MaxEvents = EVENT_MAX_EVENTS
        End Get
    End Property
    ''' <summary>
    ''' CTS Error Code
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ReadOnly Property GetCTSError() As Integer
        Get
            GetCTSError = _CTSERROR
        End Get
    End Property
    ''' <summary>
    ''' Normal CTS Return Code
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ReadOnly Property GetCTSNormal() As Integer
        Get
            GetCTSNormal = _CTSNORMAL
        End Get
    End Property

    Private _interface As Interface_Renamed

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="tInf"></param>
    ''' <remarks></remarks>
    Public Sub New(ByRef tInf As Interface_Renamed)
        Me.CTSOpen(tInf)
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub Finalize()
        Me.CTSClose(_interface)
        MyBase.Finalize()
    End Sub

#Region "Basic IO"
    ''' <summary>
    ''' Open the Colorado Console
    ''' </summary>
    ''' <param name="tInf"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CTSOpen(ByRef tInf As Interface_Renamed) As Integer
        _interface = tInf
        CTSOpen = 0
        Try
            If SerialOb Is Nothing Then
                SerialOb = New WinswimCTS.CommIO
            End If
            If SerialOb.PortStatus(tInf.Port + 1).blnPortOpen = False Then
                tInf.SerialString = CTSGetSerialString(tInf)
                CTSOpen = SerialOb.CommOpen(tInf.Port + 1, "COM" & CStr(tInf.Port + 1), tInf.SerialString)
                If CTSOpen <> 0 Then
                    Throw New System.Exception(CTSCommErrorMessage("CTSOpen : Open error on COM"))
                End If
            End If
        Catch ex As Exception
            Throw New System.Exception(CTSNotInstalled("CTSOpen", CTSOpen))
            CTSOpen = -1
        End Try
    End Function
    ''' <summary>
    ''' Change return code to string message
    ''' </summary>
    ''' <param name="tFun"></param>
    ''' <param name="errCode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CTSNotInstalled(ByVal tFun As String, ByVal errCode As Integer) As String
        Select Case errCode
            Case 5
                CTSNotInstalled = tFun & " : Port Already Open!"
            Case Else
                CTSNotInstalled = tFun & String.Format(" : Colorado Timing Console is not installed on COM{0} port!", _interface.Port + 1)
        End Select
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="tMessage"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CTSCommErrorMessage(ByVal tMessage As String) As String
        Dim strError As String = ""
        CTSCommErrorMessage = tMessage & CStr(_interface.Port + 1) & " : " & SerialOb.CommGetError(strError)
    End Function
    ''' <summary>
    ''' Close CTS Console
    ''' </summary>
    ''' <param name="tInf"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CTSClose(ByRef tInf As Interface_Renamed) As Integer

        If SerialOb Is Nothing Then
            CTSClose = -1
            Exit Function
        End If

        Try
            CTSClose = 0
            If SerialOb.PortStatus(tInf.Port + 1).blnPortOpen = True Then
                CTSClose = SerialOb.CommClose(tInf.Port + 1)
            End If
            SerialOb = Nothing
        Catch ex As Exception
            Throw New System.Exception(CTSNotInstalled("CTSClose", CTSClose))
        End Try

    End Function
    ''' <summary>
    ''' Return a communications string for the communications port
    ''' </summary>
    ''' <param name="tInf"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CTSGetSerialString(ByRef tInf As Interface_Renamed) As String

        CTSGetSerialString = "baud="
        Select Case tInf.Speed
            Case 0
                CTSGetSerialString = CTSGetSerialString & "2400"
            Case 1
                CTSGetSerialString = CTSGetSerialString & "9600"
            Case 2
                CTSGetSerialString = CTSGetSerialString & "19200"
            Case 3
                CTSGetSerialString = CTSGetSerialString & "38400"
        End Select

        CTSGetSerialString = CTSGetSerialString & " parity="
        Select Case tInf.Parity
            Case 0
                CTSGetSerialString = CTSGetSerialString & "E"
            Case 1
                CTSGetSerialString = CTSGetSerialString & "M"
            Case 2
                CTSGetSerialString = CTSGetSerialString & "N"
            Case 3
                CTSGetSerialString = CTSGetSerialString & "O"
            Case 4
                CTSGetSerialString = CTSGetSerialString & "S"
        End Select

        CTSGetSerialString = CTSGetSerialString & " data="
        Select Case tInf.DataBits
            Case 0
                CTSGetSerialString = CTSGetSerialString & "4"
            Case 1
                CTSGetSerialString = CTSGetSerialString & "5"
            Case 2
                CTSGetSerialString = CTSGetSerialString & "6"
            Case 3
                CTSGetSerialString = CTSGetSerialString & "7"
            Case 4
                CTSGetSerialString = CTSGetSerialString & "8"
        End Select

        CTSGetSerialString = CTSGetSerialString & " stop="
        Select Case tInf.StopBits
            Case 0
                CTSGetSerialString = CTSGetSerialString & "1"
            Case 1
                CTSGetSerialString = CTSGetSerialString & "1.5"
            Case 2
                CTSGetSerialString = CTSGetSerialString & "2"
        End Select

    End Function
    ''' <summary>
    ''' Write byte array to communications port
    ''' </summary>
    ''' <param name="tInf"></param>
    ''' <param name="strByte"></param>
    ''' <param name="bSleep"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CTSCommWrite(ByRef tInf As Interface_Renamed, ByRef strByte As Byte(), Optional ByVal bSleep As Boolean = False) As Integer
        Dim lngstatus As Integer

        Try
            CTSCommWrite = GetCTSNormal
            SerialOb.CommFlush(tInf.Port + 1)
            Dim lngSize As Integer = strByte.Length
            lngstatus = SerialOb.CommWriteByte(tInf.Port + 1, strByte)
            If lngStatus <> lngSize Then
                Throw New System.Exception(CTSCommErrorMessage("CTSCommWrite : Write error on COM"))
            ElseIf bSleep = True Then
                Sleep((tInf.Sleep + CTS_SLEEP_PAD) * 100)
            End If
        Catch ex As Exception
            CTSCommWrite = GetCTSError
            Throw New System.Exception(CTSNotInstalled("CTSCommWrite", lngstatus))
        End Try

    End Function
    ''' <summary>
    ''' Write string to communications port
    ''' </summary>
    ''' <param name="tInf"></param>
    ''' <param name="strData"></param>
    ''' <param name="bSleep"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CTSCommWriteString(ByRef tInf As Interface_Renamed, ByRef strData As String, Optional ByVal bSleep As Boolean = False) As Integer

        Try
            Dim OutByte() As Byte = SerialOb.ConvertStringToByteArray(strData)
            Dim OutByteLength As Integer = OutByte.Length
            ReDim Preserve OutByte(OutByteLength + 1)
            GetDICString(strData, OutByte(OutByteLength + 1), OutByte(OutByteLength))
            CTSCommWriteString = CTSCommWrite(tInf, OutByte, bSleep)
        Catch ex As Exception
            Throw New System.Exception("CTSCommWriteString : " & ex.Message)
            CTSCommWriteString = GetCTSError
        End Try

    End Function
    ''' <summary>
    ''' Read string from communications port
    ''' </summary>
    ''' <param name="tInf"></param>
    ''' <param name="strData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CTSCommReadString(ByRef tInf As Interface_Renamed, ByRef strData As String) As Integer
        Try
            CTSCommReadString = SerialOb.CommReadString(tInf.Port + 1, strData, MAXSTRINGSIZE)
            If CTSCommReadString = -1 Then
                CTSCommReadString = 0
                Throw New System.Exception(CTSCommErrorMessage("CTSCommReadString : Read Error"))
            Else
                Dim lngStatus As Integer = CTSErrorString(CTSCommReadString, strData)
                If lngStatus <> GetCTSNormal Then
                    CTSCommReadString = lngStatus * -1
                End If
            End If
        Catch ex As Exception
            Throw New System.Exception("CTSCommReadString : " & ex.Message)
            CTSCommReadString = GetCTSError
        End Try

    End Function
    ''' <summary>
    ''' Read byte array from communications port
    ''' </summary>
    ''' <param name="tInf"></param>
    ''' <param name="strByte"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CTSCommRead(ByRef tInf As Interface_Renamed, ByRef strByte As Byte()) As Integer

        Try
            CTSCommRead = SerialOb.CommRead(tInf.Port + 1, strByte, MAXSTRINGSIZE)
            If CTSCommRead = -1 Then
                CTSCommRead = 0
                Throw New System.Exception(CTSCommErrorMessage("CTSCommRead : Read Error"))
            Else
                Dim lngStatus As Integer = CTSError(strByte)
                If lngStatus <> GetCTSNormal Then
                    CTSCommRead = lngStatus * -1
                End If
            End If
        Catch ex As Exception
            Throw New System.Exception("CTSCommRead : " & ex.Message)
            CTSCommRead = GetCTSError
        End Try

    End Function

#End Region

#Region "GetSetSequence"
    ''' <summary>
    ''' Get the event sequence type from CTS console 0x52 0x74
    ''' </summary>
    ''' <param name="tType"></param>
    ''' <param name="tSeq"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CTSGetSequenceType(ByRef tType As CTSEventType, ByRef tSeq As Short) As String

        CTSGetSequenceType = ""
        Try
            Dim OutByte(7) As Byte
            OutByte(0) = &H8
            OutByte(1) = &H0
            OutByte(2) = &H52
            OutByte(3) = &H74
            OutByte(4) = tType
            OutByte(5) = tSeq
            OutByte(6) = &H31 - (tType + tSeq)
            OutByte(7) = &HFF
            If CTSCommWrite(_interface, OutByte, True) = GetCTSNormal Then
                Dim OutString As String = ""
                If CTSCommReadString(_interface, OutString) > 0 Then
                    CTSGetSequenceType = sTrim(Mid(OutString, 5))
                End If
            End If
        Catch ex As Exception
            Throw New System.Exception("CTSGetSequenceType : " & ex.Message)
        End Try

    End Function
    ''' <summary>
    ''' Put event type sequence to CTS console 0x49 0x74
    ''' </summary>
    ''' <param name="tType"></param>
    ''' <param name="tSeq"></param>
    ''' <param name="tValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CTSPutSequenceType(ByRef tType As CTSEventType, ByRef tSeq As Short, ByRef tValue As String) As Integer
        Try
            CTSPutSequenceType = CTSCommWriteString(_interface, Hex2ASCII(ToHex(Len(tValue) + 9) & "004974" & ToHex(CShort(tType)) & ToHex(tSeq)) & tValue & Chr(0), True)
        Catch ex As Exception
            Throw New System.Exception("CTSPutSequenceType : " & ex.Message)
            CTSPutSequenceType = GetCTSError
        End Try
    End Function
    ''' <summary>
    ''' Read Sequence 0x52 0x73
    ''' </summary>
    ''' <param name="tSeq"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CTSGetSequence(ByRef tSeq As Short) As String
        CTSGetSequence = ""
        Try
            Dim OutByte(8) As Byte
            OutByte(0) = &H9
            OutByte(1) = &H0
            OutByte(2) = &H52
            OutByte(3) = &H73
            OutByte(4) = tSeq
            OutByte(5) = &H0
            OutByte(6) = &H0
            OutByte(7) = &H31 - tSeq
            OutByte(8) = &HFF
            If CTSCommWrite(_interface, OutByte, True) = GetCTSNormal Then
                Dim OutString As String = ""
                If CTSCommReadString(_interface, OutString) > 0 Then
                    CTSGetSequence = sTrim(Mid(OutString, 4))
                End If
            End If
        Catch ex As Exception
            Throw New System.Exception("CTSGetSequence : " & ex.Message)
        End Try

    End Function
    ''' <summary>
    ''' Set the event sequence select 0x49 0x72
    ''' </summary>
    ''' <param name="tSeq"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CTSSetSequence(ByRef tSeq As Short) As Integer
        Try
            Dim OutByte(6) As Byte
            OutByte(0) = &H7
            OutByte(1) = &H0
            OutByte(2) = &H49
            OutByte(3) = &H72
            OutByte(4) = tSeq
            OutByte(5) = &H3D - tSeq
            OutByte(6) = &HFF
            CTSSetSequence = CTSCommWrite(_interface, OutByte, True)
        Catch ex As Exception
            Throw New System.Exception("CTSSetSequence : " & ex.Message)
            CTSSetSequence = GetCTSError
        End Try

    End Function

#End Region

#Region "WhoAreYou"
    ''' <summary>
    ''' Returns the CTS console version string 0x57
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CTSWhoAreYou() As String
        CTSWhoAreYou = ""
        Try
            Dim OutByte(4) As Byte
            OutByte(0) = &H5
            OutByte(1) = &H0
            OutByte(2) = &H57
            OutByte(3) = &HA3
            OutByte(4) = &HFF
            If CTSCommWrite(_interface, OutByte, True) = GetCTSNormal Then
                Dim OutString As String = ""
                If CTSCommReadString(_interface, OutString) > 0 Then
                    CTSWhoAreYou = Mid(OutString, 3, Len(OutString) - 4) ' TODO: Fix to find the "3.1.x" version number in string
                End If
                CTSVERSION = CTSWhoAreYou
            End If
        Catch ex As Exception
            Throw New System.Exception("CTSWhoAreYou : " & ex.Message)
        End Try

    End Function

#End Region

#Region "Meet"
    ''' <summary>
    ''' Returns the current, next or previous meet into Type_Meet_Header 
    ''' tDirection = 0 then current meet
    ''' tDirection = 1 then next meet
    ''' tDirection = -1 then previous meet
    ''' </summary>
    ''' <param name="tDirection"></param>
    ''' <param name="tMeetHeader"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CTSGetMeet(ByRef tDirection As Short, ByRef tMeetHeader As TYPE_MEETMGMT_HEADER) As Integer
        Dim tByte(10) As Byte

        Try
            Select Case tDirection
                Case 0 'T
                    tByte(0) = &H5
                    tByte(1) = &H0
                    tByte(2) = &H54
                    tByte(3) = &HA6
                    tByte(4) = &HFF

                Case 1 'M+
                    tByte(0) = &H6
                    tByte(1) = &H0
                    tByte(2) = &H4D
                    tByte(3) = &H2B
                    tByte(4) = &H81
                    tByte(5) = &HFF

                Case -1 'M-
                    tByte(0) = &H6
                    tByte(1) = &H0
                    tByte(2) = &H4D
                    tByte(3) = &H2D
                    tByte(4) = &H7F
                    tByte(5) = &HFF

            End Select
            ReDim Preserve tByte(tByte(0) - 1)
            Dim iret As Integer = CTSCommWrite(_interface, tByte, True)
            If iret <> GetCTSNormal Then
                CTSGetMeet = iret
                Exit Function
            End If
            Try
                Dim OutByte(MAXSTRINGSIZE) As Byte
                CTSGetMeet = CTSCommRead(_interface, OutByte)
                If CTSGetMeet = 0 Then
                    CTSGetMeet = GetCTSError
                ElseIf CTSGetMeet >= 50 Then
                    If CTSGetHeaderDate(tMeetHeader, OutByte) = -1 Then
                        CTSGetMeet = GetCTSError
                    End If
                End If
            Catch ex As Exception
                Throw New System.Exception("CTSGetMeet : Read Error : " & ex.Message)
            End Try
        Catch ex As Exception
            Throw New System.Exception("CTSGetMeet : Write Error : " & ex.Message)
        End Try

    End Function
#End Region

#Region "MeetData"
    ''' <summary>
    ''' Return array of races with eventno and heats for current meet
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CTSGetRaceList() As RACE_LIST_DATA()
        Dim iSavEvent As Integer = -1
        Dim iSavHeat As Integer = -1

        Try
            Dim i As Integer = 0
            Dim tRace(500) As RACE_LIST_DATA
            tRace(i).Race = 1
            While CTSGetRaceData(tRace(i)) = True
                If tRace(i).EventNo = iSavEvent And tRace(i).Heat = iSavHeat Then
                    Exit While
                End If
                iSavEvent = tRace(i).EventNo
                iSavHeat = tRace(i).Heat
                i += 1
                tRace(i).Race = i + 1
                tRace(i).EventNo = 0
                tRace(i).Heat = 0
            End While
            ReDim Preserve tRace(i - 1)
            CTSGetRaceList = tRace
        Catch ex As Exception
            Throw New System.Exception("CTSGetRaceList : " & ex.Message)
        End Try

    End Function
    ''' <summary>
    ''' Get race event and heat for current meet
    ''' </summary>
    ''' <param name="tRace"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CTSGetRaceData(ByRef tRace As RACE_LIST_DATA) As Boolean
        Dim block(7) As Byte
        Dim OutByte(200) As Byte

        CTSGetRaceData = False
        Try
            Dim InByte(7) As Byte
            InByte(0) = &H8
            InByte(1) = &H0
            InByte(2) = &H53
            InByte(3) = &H52
            uwSplit(CLng(tRace.Race), block(0), block(1), block(2), block(3))
            InByte(4) = block(3)
            InByte(5) = block(2)
            InByte(6) = &H52 - (InByte(4) + InByte(5))
            InByte(7) = &HFF
            If CTSCommWrite(_interface, InByte, True) = GetCTSNormal Then
                If CTSCommRead(_interface, OutByte) > 6 Then
                    tRace.EventNo = OutByte(2)
                    tRace.Heat = OutByte(3)
                    CTSGetRaceData = True
                End If
            End If
        Catch ex As Exception
            Throw New System.Exception("CTSGetRaceData : " & ex.Message)
        End Try

    End Function
    ''' <summary>
    ''' Get Swim Data Header
    ''' </summary>
    ''' <param name="tType"></param>
    ''' <param name="tMeetHeader"></param>
    ''' <param name="tLaneData"></param>
    ''' <param name="tEvent"></param>
    ''' <param name="tHeat"></param>
    ''' <param name="tBackup"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CTSGetSwimDataHeader(ByRef tType As String, ByRef tMeetHeader As TYPE_MEETMGMT_HEADER, ByRef tLaneData() As TYPE_LANE_DATA, Optional ByRef tEvent As Short = 1, Optional ByRef tHeat As Short = 1, Optional ByRef tBackup As Short = 0) As Integer
        Dim OutString As String
        Dim tbyte() As Byte

        CTSGetSwimDataHeader = GetCTSError
        Select Case tType
            Case "C"
                OutString = Hex2ASCII("09005353424943")
            Case "L"
                OutString = Hex2ASCII("0900535342494C")
            Case "N"
                OutString = Hex2ASCII("0900535342494E")
            Case "E"
                OutString = Hex2ASCII("0B005353424945" & ToHex(tHeat) & ToHex(tEvent))
            Case "R"
                OutString = Hex2ASCII("0B00535342495200" & ToHex(Val(CStr(tEvent))))
            Case Else
                Throw New System.Exception("CTSGetSwimDataHeader : Invalid command : Sub-Command=" & tType)
                Exit Function
        End Select

        Try
            CTSGetSwimDataHeader = CTSCommWriteString(_interface, OutString, True)
            Sleep((_interface.Sleep + CTS_SLEEP_PAD) * 100)
            If CTSGetSwimDataHeader = GetCTSNormal Then
                Try
                    ReDim tbyte(MAXSTRINGSIZE)
                    CTSGetSwimDataHeader = CTSCommRead(_interface, tbyte)
                    If CTSGetSwimDataHeader > 8 Then
                        tMeetHeader.event_Renamed = tbyte(2)
                        tMeetHeader.heat = tbyte(3)
                        tMeetHeader.lanes_in_pool = tbyte(16)
                        tMeetHeader.num_times_for_each_lane = tbyte(17)
                        tMeetHeader.number_of_buttons = tbyte(48)
                        tMeetHeader.include_relay_judging = tbyte(47)
                        tMeetHeader.include_individual_buttons = tbyte(46)
                        tMeetHeader.race_number = MakeWord(tbyte(6), tbyte(7))
                        tMeetHeader.date_long_year = MakeWord(tbyte(40), tbyte(41))
                        tMeetHeader.event_16_bit = MakeWord(tbyte(44), tbyte(45))
                        Dim cLen As Integer = (51 + ((tMeetHeader.lanes_in_pool * tMeetHeader.num_times_for_each_lane * 4) + (tMeetHeader.lanes_in_pool * 1)))
                        If CTSGetHeaderDate(tMeetHeader, tbyte) = -1 Then
                            CTSGetSwimDataHeader = -1
                        ElseIf CTSGetSwimDataHeader <> cLen Then
                            Throw New System.Exception("CTSGetSwimDataHeader : Invalid Swim Data Length " & cLen.ToString & " <> " & CTSGetSwimDataHeader.ToString)
                            CTSGetSwimDataHeader = -2
                        Else
                            CTSGetLane(tbyte, CShort(tMeetHeader.num_times_for_each_lane), CShort(tMeetHeader.number_of_buttons), CShort(tMeetHeader.lanes_in_pool), tLaneData, tBackup)
                        End If
                    End If
                Catch ex As Exception
                    Throw New System.Exception("CTSGetSwimDataHeader : Read Error : " & ex.Message)
                End Try
            End If
        Catch ex As Exception
            Throw New System.Exception("CTSGetSwimDataHeader : Write Error : " & ex.Message)
        End Try

    End Function
    ''' <summary>
    ''' Get meet header
    ''' </summary>
    ''' <param name="tMeetHeader"></param>
    ''' <param name="tbyte"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CTSGetHeaderDate(ByRef tMeetHeader As TYPE_MEETMGMT_HEADER, ByRef tbyte() As Byte) As Short

        CTSGetHeaderDate = 0
        Try
            tMeetHeader.date_seconds = tbyte(HEADER_DATE_OFFSET)
            tMeetHeader.date_minutes = tbyte(HEADER_DATE_OFFSET + 1)
            tMeetHeader.date_hours = tbyte(HEADER_DATE_OFFSET + 2)
            tMeetHeader.date_month = tbyte(HEADER_DATE_OFFSET + 4)
            tMeetHeader.date_day = tbyte(HEADER_DATE_OFFSET + 5)
            tMeetHeader.date_year = tbyte(HEADER_DATE_OFFSET + 6)
            tMeetHeader.date_weekday = tbyte(HEADER_DATE_OFFSET + 3)
            If (tMeetHeader.date_seconds < 1 Or tMeetHeader.date_seconds > 59) Or (tMeetHeader.date_minutes < 1 Or tMeetHeader.date_minutes > 59) Or (tMeetHeader.date_hours < 1 Or tMeetHeader.date_hours > 24) Then
                CTSGetHeaderDate = -1
            ElseIf (tMeetHeader.date_month < 1 Or tMeetHeader.date_month > 12) Or (tMeetHeader.date_day < 1 Or tMeetHeader.date_day > 31) Or (tMeetHeader.date_year < 1 Or tMeetHeader.date_year > 99) Then
                CTSGetHeaderDate = -1
            Else
                If tMeetHeader.date_year < 20 Then
                    tMeetHeader.date_year = tMeetHeader.date_year + 2000
                Else
                    tMeetHeader.date_year = tMeetHeader.date_year + 1900
                End If
                tMeetHeader.date_date = System.DateTime.FromOADate(DateSerial(tMeetHeader.date_year, tMeetHeader.date_month, tMeetHeader.date_day).ToOADate + TimeSerial(tMeetHeader.date_hours, tMeetHeader.date_minutes, tMeetHeader.date_seconds).ToOADate)
            End If
        Catch ex As Exception
            Throw New System.Exception("CTSGetHeaderDate : " & ex.Message)
            CTSGetHeaderDate = -1
        End Try

    End Function
    ''' <summary>
    ''' Return lane information
    ''' </summary>
    ''' <param name="szL">ddd</param>
    ''' <param name="iNbrTimesPerLane"></param>
    ''' <param name="iNbrButtonsPerLane"></param>
    ''' <param name="iNbrLanes"></param>
    ''' <param name="tLaneHeader"></param>
    ''' <param name="tBackup"></param>
    ''' <remarks>A remark</remarks>
    Private Sub CTSGetLane(ByRef szL() As Byte, ByRef iNbrTimesPerLane As Short, ByRef iNbrButtonsPerLane As Short, ByRef iNbrLanes As Short, ByRef tLaneHeader() As TYPE_LANE_DATA, ByRef tBackup As Short)
        Dim iB3Disp, iB1Disp, iFDisp, iB2Disp, iBKDisp As Short

        ' determine the displacements for the final time, buttons and backup time
        iFDisp = 999
        iB1Disp = 999
        iB2Disp = 999
        iB3Disp = 999
        iBKDisp = 999

        If (iNbrButtonsPerLane = 0) Then '/sssf 4/0
            iFDisp = iNbrTimesPerLane - iNbrButtonsPerLane - 1 '//finaltime
            iB1Disp = 999 '//button1
            iB2Disp = 999 '//button2
            iB3Disp = 999 '//button3
            iBKDisp = iNbrTimesPerLane - 1 '//backup
        End If
        If (iNbrButtonsPerLane = 1) Then '/sssf 4/0
            iFDisp = iNbrTimesPerLane - iNbrButtonsPerLane - 1 '//finaltime
            iB1Disp = iNbrTimesPerLane - iNbrButtonsPerLane '//button1
            iB2Disp = 999 '//button2
            iB3Disp = 999 '//button3
            iBKDisp = iNbrTimesPerLane - 1 '//backup
        End If
        If (iNbrButtonsPerLane = 2) Then '/sssf 4/0
            iFDisp = iNbrTimesPerLane - iNbrButtonsPerLane - 2 '//finaltime
            iB1Disp = iNbrTimesPerLane - iNbrButtonsPerLane - 1 '//button1
            iB2Disp = iNbrTimesPerLane - iNbrButtonsPerLane - 0 '//button2
            iB3Disp = 999 '//button3
            iBKDisp = iNbrTimesPerLane - 1 '//backup
        End If
        If (iNbrButtonsPerLane = 3) Then '/sssf 4/0
            iFDisp = iNbrTimesPerLane - iNbrButtonsPerLane - 2 '//finaltime
            iB1Disp = iNbrTimesPerLane - iNbrButtonsPerLane - 1 '//button1
            iB2Disp = iNbrTimesPerLane - iNbrButtonsPerLane - 0 '//button2
            iB3Disp = iNbrTimesPerLane - iNbrButtonsPerLane + 1 '//button3
            iBKDisp = iNbrTimesPerLane - 1 '//backup
        End If

        Try
            Dim i, i2 As Short
            Dim dwTime As Integer
            Dim place As Byte
            Dim iDisp As Short = IIf((tBackup = 1), iBKDisp, iFDisp)
            Dim cP As Short = SIZEOFTYPE_MEETMGMT_HEADER + 1 '/point at the first time, skip over the meet header and lane 1 place

            For i = 0 To iNbrLanes - 1
                place = szL(cP - 1)
                For i2 = 0 To iNbrTimesPerLane - 1
                    dwTime = MakeDWord(MakeWord(szL(cP + 0), szL(cP + 1)), MakeWord(szL(cP + 2), szL(cP + 3)))
                    If (i2 = iDisp) Then
                        tLaneHeader(i).place = place
                        tLaneHeader(i).time = dwTime
                        tLaneHeader(i).mins = Fix(dwTime / 60000)
                        tLaneHeader(i).secs = dwTime - (tLaneHeader(i).mins * 60000)
                        tLaneHeader(i).hsecs = (tLaneHeader(i).secs Mod 1000)
                        tLaneHeader(i).secs = Fix(tLaneHeader(i).secs / 1000)
                    End If
                    cP = cP + 4
                Next i2
                cP = cP + 1 ' Skip the next place holder
            Next i
        Catch ex As Exception
            Throw New System.Exception("CTSGetLane : " & ex.Message)
        End Try

    End Sub

#End Region

#Region "Events"
    ''' <summary>
    ''' Get an event from CTS console
    ''' </summary>
    ''' <param name="iEvent"></param>
    ''' <param name="tByte"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CTSGetEvent(ByRef iEvent As Short, ByRef tByte As Byte()) As Integer
        CTSGetEvent = GetCTSError
        Try
            If CTSCommWriteString(_interface, Hex2ASCII("09005275" & ToHex(_interface.EventSequence) & ToHex(iEvent, 2)), True) = GetCTSNormal Then
                ReDim tByte(MAXSTRINGSIZE)
                CTSGetEvent = SerialOb.CommRead(_interface.Port + 1, tByte, MAXSTRINGSIZE)
                ReDim Preserve tByte(CTSGetEvent)
            End If
        Catch ex As Exception
            Throw New System.Exception("CTSGetEvent : " & ex.Message)
        End Try

    End Function

    ''' <summary>
    ''' Store event into CTS console
    ''' </summary>
    ''' <param name="iEvent"></param>
    ''' <param name="iDistance"></param>
    ''' <param name="chkArray"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CTSPutEvent(ByRef iEvent As Short, ByRef iDistance As Short, ByRef chkArray() As CTS_LABELS) As Integer
        Dim i As Integer
        Dim j As Integer
        Dim OutByte(50) As Byte
        Dim block(4) As Byte
        Const EVENT_FIXED_LENGTH As Integer = 9

        Try
            OutByte(0) = &H0
            OutByte(1) = &H0
            OutByte(2) = &H49
            OutByte(3) = &H75
            OutByte(4) = _interface.EventSequence
            uwSplit(CLng(iEvent), block(0), block(1), block(2), block(3))
            OutByte(5) = block(3)
            OutByte(6) = block(2)
            uwSplit(CLng(iDistance), block(0), block(1), block(2), block(3))
            OutByte(7) = block(3)
            OutByte(8) = block(2)

            j = EVENT_FIXED_LENGTH
            For i = 0 To UBound(chkArray)
                If chkArray(i).checked = True Then
                    Select Case chkArray(i).label
                        Case CTSLabelType.EventNo ' event
                        Case CTSLabelType.Distance ' distance
                        Case CTSLabelType.Sex ' sex
                            OutByte(j) = CByte(chkArray(i).value)
                            j = j + 1
                        Case CTSLabelType.Stroke ' stroke
                            OutByte(j) = CByte(chkArray(i).value)
                            j = j + 1
                        Case CTSLabelType.AgeGroup ' age group
                            OutByte(j) = CByte(chkArray(i).value)
                            j = j + 1
                        Case CTSLabelType.Session ' session
                            OutByte(j) = CByte(chkArray(i).value)
                            j = j + 1
                    End Select
                End If
            Next i

            uwSplit(CLng(j + 3), block(0), block(1), block(2), block(3))
            OutByte(0) = block(3)
            OutByte(1) = block(2)
            ReDim Preserve OutByte(j + 2)
            Dim hash As Integer = -1
            For i = 0 To j
                hash = hash - OutByte(i)
            Next i

            uwSplit(CLng(hash), block(0), block(1), block(2), block(3))
            OutByte(j + 1) = block(3)
            OutByte(j + 2) = block(2)

            CTSPutEvent = CTSCommWrite(_interface, OutByte, True)
        Catch ex As Exception
            Throw New System.Exception("CTSPutEvent : " & ex.Message)
            CTSPutEvent = GetCTSError
        End Try

    End Function
    ''' <summary>
    ''' Clear the event 
    ''' </summary>
    ''' <param name="iEvent"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CTSClearEvent(ByRef iEvent As Short) As Integer
        Try
            Dim OutString As String = Hex2ASCII("11004975" & ToHex(_interface.EventSequence) & ToHex(iEvent, 2)) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0)
            CTSClearEvent = CTSCommWriteString(_interface, OutString, True)
        Catch ex As Exception
            Throw New System.Exception("CTSClearEvent : " & ex.Message)
            CTSClearEvent = GetCTSError
        End Try
    End Function

#End Region

#Region "Error"
    ''' <summary>
    ''' Return the CTS error code from the CTS response in a string
    ''' </summary>
    ''' <param name="tLen"></param>
    ''' <param name="strData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CTSErrorString(ByRef tLen As Integer, ByRef strData As String) As Integer
        Dim strByte(MAXSTRINGSIZE) As Byte
        strByte = SerialOb.ConvertStringToByteArray(strData)
        If strByte(0) = 6 Then
            CTSErrorString = MakeWord(strByte(2), strByte(3))
        Else
            CTSErrorString = GetCTSNormal
        End If
    End Function

    ''' <summary>
    ''' Return the CTS error code from the CTS response in a byte()
    ''' </summary>
    ''' <param name="strByte"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CTSError(ByRef strByte As Byte()) As Integer
        If strByte(0) = 6 Then
            CTSError = MakeWord(strByte(2), strByte(3))
        Else
            CTSError = GetCTSNormal
        End If
    End Function

#End Region

#Region "String"
    ''' <summary>
    ''' Convert a short integer to hex string
    ''' </summary>
    ''' <param name="num"></param>
    ''' <param name="iLen"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ToHex(ByRef num As Short, Optional ByRef iLen As Short = 1) As String
        Dim theTemp As String
        Dim i As Short

        ToHex = ""
        theTemp = Hex(num)

        If Len(theTemp) Mod 2 = 1 Then theTemp = "0" & theTemp
        While Len(theTemp) < (iLen * 2)
            theTemp = theTemp & "00"
        End While

        If num > 255 Then ' reverse the bytes
            For i = iLen To 1 Step -1
                ToHex = ToHex & Mid(theTemp, i * 2 - 1, 2)
            Next i
        Else
            ToHex = theTemp
        End If

    End Function

    ''' <summary>
    ''' Make short integer from two bytes
    ''' </summary>
    ''' <param name="LoByte"></param>
    ''' <param name="HiByte"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function MakeWord(ByVal LoByte As Byte, ByVal HiByte As Byte) As Short
        If HiByte < &H80S Then
            MakeWord = HiByte * &H100S Or LoByte
        Else
            MakeWord = CShort(HiByte Or &HFF00S) * &H100 Or LoByte
        End If
    End Function
    ''' <summary>
    ''' Make long integer from two words
    ''' </summary>
    ''' <param name="LoWord"></param>
    ''' <param name="HiWord"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function MakeDWord(ByVal LoWord As Short, ByVal HiWord As Short) As Integer
        MakeDWord = HiWord * &H10000 Or (LoWord And &HFFFF)
    End Function
    ''' <summary>
    ''' This function trims a string of right and left spaces
    ''' </summary>
    ''' <param name="s"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function sTrim(ByRef s As String) As String
        ' this function trims a string of right and left spaces
        ' it recognizes 0 as a string terminator
        Dim i As Short
        i = InStr(s, Chr(0))
        If (i > 0) Then
            sTrim = Trim(Left(s, i - 1))
        Else
            sTrim = Trim(s)
        End If
    End Function
#End Region
    ''' <summary>
    ''' Calculate the checksum for the string and put into lobyte and hibyte
    ''' </summary>
    ''' <param name="tStr"></param>
    ''' <param name="hibyte"></param>
    ''' <param name="lobyte"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetDICString(ByRef tStr As String, ByRef hibyte As Byte, ByRef lobyte As Byte) As Integer
        Dim i As Short
        Dim hash As Integer = -1
        Dim block(3) As Byte

        Try
            For i = 0 To Len(tStr) - 1
                hash = hash - CInt(Asc(Mid(tStr, i + 1, 1)))
            Next i
            uwSplit(CLng(hash), block(0), block(1), block(2), block(3))
            GetDICString = hash
            hibyte = block(2)
            lobyte = block(3)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function
    ''' <summary>
    ''' Split a Long Integer into 4 bytes
    ''' </summary>
    ''' <param name="w"></param>
    ''' <param name="a"></param>
    ''' <param name="b"></param>
    ''' <param name="c"></param>
    ''' <param name="d"></param>
    ''' <remarks></remarks>
    Private Sub uwSplit(ByVal w As Long, ByRef a As Byte, ByRef b As Byte, ByRef c As Byte, ByRef d As Byte)
        ' Split 32-bit word w into 4 x 8-bit bytes
        a = CByte(((w And &HFF000000) \ &H1000000) And &HFF)
        b = CByte(((w And &HFF0000) \ &H10000) And &HFF)
        c = CByte(((w And &HFF00) \ &H100) And &HFF)
        d = CByte((w And &HFF) And &HFF)
    End Sub
    ''' <summary>
    ''' Convert hex notation to a hex string
    ''' </summary>
    ''' <param name="hextext"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function Hex2ASCII(ByVal hextext As String) As String
        Dim y As Short
        Dim num As String
        Dim Value As String = ""

        For y = 1 To Len(hextext)
            num = Mid(hextext, y, 2)
            Value = Value & Chr(Val("&h" & num))
            y = y + 1
        Next y

        Hex2ASCII = Value
    End Function

End Class