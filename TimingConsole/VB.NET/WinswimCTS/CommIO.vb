Option Strict Off
Option Explicit On

Imports System.text
<System.Runtime.InteropServices.ProgId("CommIO_NET.CommIO")> Public Class CommIO

	'-------------------------------------------------------------------------------
	' modCOMM - Written by: David M. Hitchner
	'
	' This VB module is a collection of routines to perform serial port I/O without
	' using the Microsoft Comm Control component.  This module uses the Windows API
	' to perform the overlapped I/O operations necessary for serial communications.
	'
	' The routine can handle up to 4 serial ports which are identified with a
	' Port ID.
	'
	' All routines (with the exception of CommRead and CommWrite) return an error
	' code or 0 if no error occurs.  The routine CommGetError can be used to get
	' the complete error message.
	'-------------------------------------------------------------------------------
	
	'-------------------------------------------------------------------------------
	' Public Constants
	'-------------------------------------------------------------------------------
	
	' Output Control Lines (CommSetLine)
	Private Const LINE_BREAK As Short = 1
	Private Const LINE_DTR As Short = 2
	Private Const LINE_RTS As Short = 3
	
	' Input Control Lines  (CommGetLine)
	Private Const LINE_CTS As Integer = &H10
	Private Const LINE_DSR As Integer = &H20
	Private Const LINE_RING As Integer = &H40
	Private Const LINE_RLSD As Integer = &H80
	Private Const LINE_CD As Integer = &H80
	
	'-------------------------------------------------------------------------------
	' System Constants
	'-------------------------------------------------------------------------------
	Private Const ERROR_IO_INCOMPLETE As Short = 996
	Private Const ERROR_IO_PENDING As Short = 997
	Private Const GENERIC_READ As Integer = &H80000000
	Private Const GENERIC_WRITE As Integer = &H40000000
	Private Const FILE_ATTRIBUTE_NORMAL As Short = &H80s
	Private Const FILE_FLAG_OVERLAPPED As Integer = &H40000000
	Private Const FORMAT_MESSAGE_FROM_SYSTEM As Short = &H1000s
	Private Const OPEN_EXISTING As Short = 3
	
	' COMM Functions
	Private Const MS_CTS_ON As Integer = &H10
	Private Const MS_DSR_ON As Integer = &H20
	Private Const MS_RING_ON As Integer = &H40
	Private Const MS_RLSD_ON As Integer = &H80
	Private Const PURGE_RXABORT As Short = &H2s
	Private Const PURGE_RXCLEAR As Short = &H8s
	Private Const PURGE_TXABORT As Short = &H1s
	Private Const PURGE_TXCLEAR As Short = &H4s
	
	' COMM Escape Functions
	Private Const CLRBREAK As Short = 9
	Private Const CLRDTR As Short = 6
	Private Const CLRRTS As Short = 4
	Private Const SETBREAK As Short = 8
	Private Const SETDTR As Short = 5
	Private Const SETRTS As Short = 3
	
	'-------------------------------------------------------------------------------
	' System Structures
	'-------------------------------------------------------------------------------
	Private Structure COMSTAT
		Dim fBitFields As Integer ' See Comment in Win32API.Txt
		Dim cbInQue As Integer
		Dim cbOutQue As Integer
	End Structure
	
	Private Structure COMMTIMEOUTS
		Dim ReadIntervalTimeout As Integer
		Dim ReadTotalTimeoutMultiplier As Integer
		Dim ReadTotalTimeoutConstant As Integer
		Dim WriteTotalTimeoutMultiplier As Integer
		Dim WriteTotalTimeoutConstant As Integer
	End Structure
	
	'
	' The DCB structure defines the control setting for a serial
	' communications device.
	'
	Public Structure DCB
		Dim DCBlength As Integer
		Dim BaudRate As Integer
		Dim fBitFields As Integer ' See Comments in Win32API.Txt
		Dim wReserved As Short
		Dim XonLim As Short
		Dim XoffLim As Short
		Dim ByteSize As Byte
		Dim Parity As Byte
		Dim StopBits As Byte
		Dim XonChar As Byte
		Dim XoffChar As Byte
		Dim ErrorChar As Byte
		Dim EofChar As Byte
		Dim EvtChar As Byte
		Dim wReserved1 As Short 'Reserved; Do Not Use
	End Structure
	
	Private Structure OVERLAPPED
		Dim Internal As Integer
		Dim InternalHigh As Integer
		Dim offset As Integer
		Dim OffsetHigh As Integer
		Dim hEvent As Integer
	End Structure
	
	Private Structure SECURITY_ATTRIBUTES
		Dim nLength As Integer
		Dim lpSecurityDescriptor As Integer
		Dim bInheritHandle As Integer
	End Structure


    '-------------------------------------------------------------------------------
    ' System Functions
    '-------------------------------------------------------------------------------
    '
    ' Fills a specified DCB structure with values specified in
    ' a device-control string.
    '
    'UPGRADE_WARNING: Structure DCB may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
    Private Declare Function BuildCommDCB Lib "kernel32" Alias "BuildCommDCBA" (ByVal lpDef As String, ByRef lpDCB As DCB) As Integer
    '
    ' Retrieves information about a communications error and reports
    ' the current status of a communications device. The function is
    ' called when a communications error occurs, and it clears the
    ' device's error flag to enable additional input and output
    ' (I/O) operations.
    '
    'UPGRADE_WARNING: Structure COMSTAT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
    Private Declare Function ClearCommError Lib "kernel32" (ByVal hFile As Integer, ByRef lpErrors As Integer, ByRef lpStat As COMSTAT) As Integer
    '
    ' Closes an open communications device or file handle.
    '
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Integer) As Integer
    '
    ' Creates or opens a communications resource and returns a handle
    ' that can be used to access the resource.
    '
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Integer, ByVal dwShareMode As Integer, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Integer, ByVal dwFlagsAndAttributes As Integer, ByVal hTemplateFile As Integer) As Integer
    '
    ' Directs a specified communications device to perform a function.
    '
    Private Declare Function EscapeCommFunction Lib "kernel32" (ByVal nCid As Integer, ByVal nFunc As Integer) As Integer
    '
    ' Formats a message string such as an error string returned
    ' by anoher function.
    '
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Integer, ByRef lpSource As Long, ByVal dwMessageId As Integer, ByVal dwLanguageId As Integer, ByVal lpBuffer As String, ByVal nSize As Integer, ByRef Arguments As Integer) As Integer
    '
    ' Retrieves modem control-register values.
    '
    Private Declare Function GetCommModemStatus Lib "kernel32" (ByVal hFile As Integer, ByRef lpModemStat As Integer) As Integer
    '
    ' Retrieves the current control settings for a specified
    ' communications device.
    '
    'UPGRADE_WARNING: Structure DCB may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
    Private Declare Function GetCommState Lib "kernel32" (ByVal nCid As Integer, ByRef lpDCB As DCB) As Integer
    '
    ' Retrieves the calling thread's last-error code value.
    '
    Private Declare Function GetLastError Lib "kernel32" () As Integer
    '
    ' Retrieves the results of an overlapped operation on the
    ' specified file, named pipe, or communications device.
    '
    'UPGRADE_WARNING: Structure OVERLAPPED may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
    Private Declare Function GetOverlappedResult Lib "kernel32" (ByVal hFile As Integer, ByRef lpOverlapped As OVERLAPPED, ByRef lpNumberOfBytesTransferred As Integer, ByVal bWait As Integer) As Integer
    '
    ' Discards all characters from the output or input buffer of a
    ' specified communications resource. It can also terminate
    ' pending read or write operations on the resource.
    '
    Private Declare Function PurgeComm Lib "kernel32" (ByVal hFile As Integer, ByVal dwFlags As Integer) As Integer
    '
    ' Reads data from a file, starting at the position indicated by the
    ' file pointer. After the read operation has been completed, the
    ' file pointer is adjusted by the number of bytes actually read,
    ' unless the file handle is created with the overlapped attribute.
    ' If the file handle is created for overlapped input and output
    ' (I/O), the application must adjust the position of the file pointer
    ' after the read operation.
    '
    'UPGRADE_WARNING: Structure OVERLAPPED may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
    Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Integer, ByVal lpBuffer As Byte(), ByVal nNumberOfBytesToRead As Integer, ByRef lpNumberOfBytesRead As Integer, ByRef lpOverlapped As OVERLAPPED) As Integer
    '
    ' Configures a communications device according to the specifications
    ' in a device-control block (a DCB structure). The function
    ' reinitializes all hardware and control settings, but it does not
    ' empty output or input queues.
    '
    'UPGRADE_WARNING: Structure DCB may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
    Private Declare Function SetCommState Lib "kernel32" (ByVal hCommDev As Integer, ByRef lpDCB As DCB) As Integer
    '
    ' Sets the time-out parameters for all read and write operations on a
    ' specified communications device.
    '
    'UPGRADE_WARNING: Structure COMMTIMEOUTS may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
    Private Declare Function SetCommTimeouts Lib "kernel32" (ByVal hFile As Integer, ByRef lpCommTimeouts As COMMTIMEOUTS) As Integer
    '
    ' Initializes the communications parameters for a specified
    ' communications device.
    '
    Private Declare Function SetupComm Lib "kernel32" (ByVal hFile As Integer, ByVal dwInQueue As Integer, ByVal dwOutQueue As Integer) As Integer
    '
    ' Writes data to a file and is designed for both synchronous and a
    ' synchronous operation. The function starts writing data to the file
    ' at the position indicated by the file pointer. After the write
    ' operation has been completed, the file pointer is adjusted by the
    ' number of bytes actually written, except when the file is opened with
    ' FILE_FLAG_OVERLAPPED. If the file handle was created for overlapped
    ' input and output (I/O), the application must adjust the position of
    ' the file pointer after the write operation is finished.
    '
    'UPGRADE_WARNING: Structure OVERLAPPED may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
    Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Integer, ByVal lpBuffer As Byte(), ByVal nNumberOfBytesToWrite As Integer, ByRef lpNumberOfBytesWritten As Integer, ByRef lpOverlapped As OVERLAPPED) As Integer

    '-------------------------------------------------------------------------------
    ' Program Constants
    '-------------------------------------------------------------------------------

    Private Const MAX_PORTS As Short = 7

    '-------------------------------------------------------------------------------
    ' Program Structures
    '-------------------------------------------------------------------------------

    Private Structure COMM_ERROR
        Dim lngErrorCode As Integer
        Dim strFunction As String
        Dim strErrorMessage As String
    End Structure

    Public Structure COMM_PORT
        Dim lngHandle As Integer
        Dim blnPortOpen As Boolean
        Dim udtDCB As DCB
    End Structure

    '-------------------------------------------------------------------------------
    ' Program Storage
    '-------------------------------------------------------------------------------

    Private udtCommOverlap As OVERLAPPED
    Private udtCommError As COMM_ERROR
    Private udtPorts(MAX_PORTS) As COMM_PORT

    ''' <summary>
    ''' Return Port Status
    ''' </summary>
    ''' <param name="intPortID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function PortStatus(ByRef intPortID As Short) As COMM_PORT
        PortStatus = udtPorts(intPortID)
    End Function

    ''' <summary>
    ''' GetSystemMessage - Gets system error text for the specified error code.
    ''' </summary>
    ''' <param name="lngErrorCode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetSystemMessage(ByRef lngErrorCode As Integer) As String
        Dim intPos As Short
        Dim strMessage As String
        Dim strMsgBuff As New VB6.FixedLengthString(256)

        Call FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, lngErrorCode, 0, strMsgBuff.Value, 255, 0)

        intPos = InStr(1, strMsgBuff.Value, vbNullChar)
        If intPos > 0 Then
            strMessage = Trim(Left(strMsgBuff.Value, intPos - 1))
        Else
            strMessage = Trim(strMsgBuff.Value)
        End If

        GetSystemMessage = strMessage

    End Function

    ''' <summary>
    ''' CommOpen - Opens/Initializes serial port.
    '''
    '''
    ''' Parameters:
    '''   intPortID   - Port ID used when port was opened.
    '''   strPort     - COM port name. (COM1, COM2, COM3, COM4)
    '''   strSettings - Communication settings.
    '''                 Example: "baud=9600 parity=N data=8 stop=1"
    '''
    ''' Returns:
    ''' Error Code  - 0 = No Error.
    ''' </summary>
    ''' <param name="intPortID"></param>
    ''' <param name="strPort"></param>
    ''' <param name="strSettings"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CommOpen(ByRef intPortID As Short, ByRef strPort As String, ByRef strSettings As String) As Integer

        Dim lngStatus As Integer
        Dim udtCommTimeOuts As COMMTIMEOUTS
        Dim udtSecAttributes As SECURITY_ATTRIBUTES

        On Error GoTo Routine_Error

        ' See if port already in use.
        If udtPorts(intPortID).blnPortOpen Then
            lngStatus = -1
            With udtCommError
                .lngErrorCode = lngStatus
                .strFunction = "CommOpen"
                .strErrorMessage = "Port in use."
            End With

            GoTo Routine_Exit
        End If

        ' Open serial port.
        udtPorts(intPortID).lngHandle = CreateFile(strPort, GENERIC_READ Or GENERIC_WRITE, 0, udtSecAttributes, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)

        If udtPorts(intPortID).lngHandle = -1 Then
            lngStatus = SetCommError("CommOpen (CreateFile)")
            GoTo Routine_Exit
        End If

        udtPorts(intPortID).blnPortOpen = True

        ' Setup device buffers (1K each).
        lngStatus = SetupComm(udtPorts(intPortID).lngHandle, 1024, 1024)

        If lngStatus = 0 Then
            lngStatus = SetCommError("CommOpen (SetupComm)")
            GoTo Routine_Exit
        End If

        ' Purge buffers.
        lngStatus = PurgeComm(udtPorts(intPortID).lngHandle, PURGE_TXABORT Or PURGE_RXABORT Or PURGE_TXCLEAR Or PURGE_RXCLEAR)

        If lngStatus = 0 Then
            lngStatus = SetCommError("CommOpen (PurgeComm)")
            GoTo Routine_Exit
        End If

        ' Set serial port timeouts.
        With udtCommTimeOuts
            .ReadIntervalTimeout = -1
            .ReadTotalTimeoutMultiplier = 0
            .ReadTotalTimeoutConstant = 1000
            .WriteTotalTimeoutMultiplier = 0
            .WriteTotalTimeoutMultiplier = 1000
        End With

        lngStatus = SetCommTimeouts(udtPorts(intPortID).lngHandle, udtCommTimeOuts)

        If lngStatus = 0 Then
            lngStatus = SetCommError("CommOpen (SetCommTimeouts)")
            GoTo Routine_Exit
        End If

        ' Get the current state (DCB).
        lngStatus = GetCommState(udtPorts(intPortID).lngHandle, udtPorts(intPortID).udtDCB)

        If lngStatus = 0 Then
            lngStatus = SetCommError("CommOpen (GetCommState)")
            GoTo Routine_Exit
        End If

        ' Modify the DCB to reflect the desired settings.
        lngStatus = BuildCommDCB(strSettings, udtPorts(intPortID).udtDCB)

        If lngStatus = 0 Then
            lngStatus = SetCommError("CommOpen (BuildCommDCB)")
            GoTo Routine_Exit
        End If

        ' Set the new state.
        lngStatus = SetCommState(udtPorts(intPortID).lngHandle, udtPorts(intPortID).udtDCB)

        If lngStatus = 0 Then
            lngStatus = SetCommError("CommOpen (SetCommState)")
            GoTo Routine_Exit
        End If

        lngStatus = 0

Routine_Exit:
        CommOpen = lngStatus
        Exit Function

Routine_Error:
        lngStatus = Err.Number
        With udtCommError
            .lngErrorCode = lngStatus
            .strFunction = "CommOpen"
            .strErrorMessage = Err.Description
        End With
        Resume Routine_Exit
    End Function

    ''' <summary>
    ''' Set communications error
    ''' </summary>
    ''' <param name="strFunction"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SetCommError(ByRef strFunction As String) As Integer

        With udtCommError
            .lngErrorCode = Err.LastDllError
            .strFunction = strFunction
            .strErrorMessage = GetSystemMessage(.lngErrorCode)
            SetCommError = .lngErrorCode
        End With

    End Function

    ''' <summary>
    ''' Set extended error message
    ''' </summary>
    ''' <param name="strFunction"></param>
    ''' <param name="lngHnd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SetCommErrorEx(ByRef strFunction As String, ByRef lngHnd As Integer) As Integer
        Dim lngErrorFlags As Integer
        Dim udtCommStat As COMSTAT

        With udtCommError
            .lngErrorCode = GetLastError
            .strFunction = strFunction
            .strErrorMessage = GetSystemMessage(.lngErrorCode)

            Call ClearCommError(lngHnd, lngErrorFlags, udtCommStat)

            .strErrorMessage = .strErrorMessage & "  COMM Error Flags = " & Hex(lngErrorFlags)

            SetCommErrorEx = .lngErrorCode
        End With

    End Function

    '-------------------------------------------------------------------------------
    ' CommSet - Modifies the serial port settings.
    '
    ' Parameters:
    '   intPortID   - Port ID used when port was opened.
    '   strSettings - Communication settings.
    '                 Example: "baud=9600 parity=N data=8 stop=1"
    '
    ' Returns:
    '   Error Code  - 0 = No Error.
    '-------------------------------------------------------------------------------
    ''' <summary>
    ''' CommSet - Modifies the serial port settings.
    ''' </summary>
    ''' <param name="intPortID"></param>
    ''' <param name="strSettings"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CommSet(ByRef intPortID As Short, ByRef strSettings As String) As Integer

        Dim lngStatus As Integer

        On Error GoTo Routine_Error

        lngStatus = GetCommState(udtPorts(intPortID).lngHandle, udtPorts(intPortID).udtDCB)

        If lngStatus = 0 Then
            lngStatus = SetCommError("CommSet (GetCommState)")
            GoTo Routine_Exit
        End If

        lngStatus = BuildCommDCB(strSettings, udtPorts(intPortID).udtDCB)

        If lngStatus = 0 Then
            lngStatus = SetCommError("CommSet (BuildCommDCB)")
            GoTo Routine_Exit
        End If

        lngStatus = SetCommState(udtPorts(intPortID).lngHandle, udtPorts(intPortID).udtDCB)

        If lngStatus = 0 Then
            lngStatus = SetCommError("CommSet (SetCommState)")
            GoTo Routine_Exit
        End If

        lngStatus = 0

Routine_Exit:
        CommSet = lngStatus
        Exit Function

Routine_Error:
        lngStatus = Err.Number
        With udtCommError
            .lngErrorCode = lngStatus
            .strFunction = "CommSet"
            .strErrorMessage = Err.Description
        End With
        Resume Routine_Exit
    End Function

    '-------------------------------------------------------------------------------
    ' CommClose - Close the serial port.
    '
    ' Parameters:
    '   intPortID   - Port ID used when port was opened.
    '
    ' Returns:
    '   Error Code  - 0 = No Error.
    '-------------------------------------------------------------------------------
    ''' <summary>
    ''' CommClose - Close the serial port.
    ''' </summary>
    ''' <param name="intPortID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CommClose(ByRef intPortID As Short) As Integer

        Dim lngStatus As Integer

        On Error GoTo Routine_Error

        If udtPorts(intPortID).blnPortOpen Then
            lngStatus = CloseHandle(udtPorts(intPortID).lngHandle)

            If lngStatus = 0 Then
                lngStatus = SetCommError("CommClose (CloseHandle)")
                GoTo Routine_Exit
            End If

            udtPorts(intPortID).blnPortOpen = False
        End If

        lngStatus = 0

Routine_Exit:
        CommClose = lngStatus
        Exit Function

Routine_Error:
        lngStatus = Err.Number
        With udtCommError
            .lngErrorCode = lngStatus
            .strFunction = "CommClose"
            .strErrorMessage = Err.Description
        End With
        Resume Routine_Exit
    End Function

    '-------------------------------------------------------------------------------
    ' CommFlush - Flush the send and receive serial port buffers.
    '
    ' Parameters:
    '   intPortID   - Port ID used when port was opened.
    '
    ' Returns:
    '   Error Code  - 0 = No Error.
    '-------------------------------------------------------------------------------

    ''' <summary>
    ''' CommFlush - Flush the send and receive serial port buffers.
    ''' </summary>
    ''' <param name="intPortID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CommFlush(ByRef intPortID As Short) As Integer

        Dim lngStatus As Integer

        On Error GoTo Routine_Error

        lngStatus = PurgeComm(udtPorts(intPortID).lngHandle, PURGE_TXABORT Or PURGE_RXABORT Or PURGE_TXCLEAR Or PURGE_RXCLEAR)

        If lngStatus = 0 Then
            lngStatus = SetCommError("CommFlush (PurgeComm)")
            GoTo Routine_Exit
        End If

        lngStatus = 0

Routine_Exit:
        CommFlush = lngStatus
        Exit Function

Routine_Error:
        lngStatus = Err.Number
        With udtCommError
            .lngErrorCode = lngStatus
            .strFunction = "CommFlush"
            .strErrorMessage = Err.Description
        End With
        Resume Routine_Exit
    End Function
    '-------------------------------------------------------------------------------
    ' CommReadString - Read serial port input buffer.
    '
    ' Parameters:
    '   intPortID   - Port ID used when port was opened.
    '   strData     - Data buffer.
    '   lngSize     - Maximum number of bytes to be read.
    '
    ' Returns:
    '   Error Code  - 0 = No Error.
    '-------------------------------------------------------------------------------
    ''' <summary>
    ''' CommReadString - Read serial port input buffer.
    ''' </summary>
    ''' <param name="intPortID"></param>
    ''' <param name="strData"></param>
    ''' <param name="lngSize"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CommReadString(ByRef intPortID As Short, ByRef strData As String, ByRef lngSize As Integer) As Integer
        Dim tByte(lngSize) As Byte

        CommReadString = CommRead(intPortID, tByte, lngSize)
        'Dim enc As New System.Text.ASCIIEncoding()
        'strData = enc.GetString(tByte, 0, lngSize)
        Dim utf As New System.Text.UTF8Encoding()
        strData = utf.GetString(tByte)

    End Function
    '-------------------------------------------------------------------------------
    ' CommRead - Read serial port input buffer.
    '
    ' Parameters:
    '   intPortID   - Port ID used when port was opened.
    '   strData     - Data buffer.
    '   lngSize     - Maximum number of bytes to be read.
    '
    ' Returns:
    '   Error Code  - 0 = No Error.
    '-------------------------------------------------------------------------------
    ''' <summary>
    ''' CommRead - Read serial port input buffer.
    ''' </summary>
    ''' <param name="intPortID"></param>
    ''' <param name="strData"></param>
    ''' <param name="lngSize"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CommRead(ByRef intPortID As Short, ByRef strData As Byte(), ByRef lngSize As Integer) As Integer

        Dim lngStatus As Integer
        Dim lngRdSize, lngBytesRead As Integer
        Dim lngRdStatus As Integer
        Dim strRdBuffer As New VB6.FixedLengthString(1024)
        Dim lngErrorFlags As Integer
        Dim udtCommStat As COMSTAT

        On Error GoTo Routine_Error

        ReDim Preserve strData(lngSize + 1)
        lngBytesRead = 0
        System.Windows.Forms.Application.DoEvents()

        ' Clear any previous errors and get current status.
        lngStatus = ClearCommError(udtPorts(intPortID).lngHandle, lngErrorFlags, udtCommStat)

        If lngStatus = 0 Then
            lngBytesRead = -1
            lngStatus = SetCommError("CommRead (ClearCommError)")
            GoTo Routine_Exit
        End If

        If udtCommStat.cbInQue > 0 Then
            If udtCommStat.cbInQue > lngSize Then
                lngRdSize = udtCommStat.cbInQue
            Else
                lngRdSize = lngSize
            End If
        Else
            lngRdSize = 0
        End If

        If lngRdSize Then
            lngRdStatus = ReadFile(udtPorts(intPortID).lngHandle, strData, lngRdSize, lngBytesRead, udtCommOverlap)

            If lngRdStatus = 0 Then
                lngStatus = GetLastError
                If lngStatus = ERROR_IO_PENDING Then
                    ' Wait for read to complete.
                    ' This function will timeout according to the
                    ' COMMTIMEOUTS.ReadTotalTimeoutConstant variable.
                    ' Every time it times out, check for port errors.

                    ' Loop until operation is complete.
                    While GetOverlappedResult(udtPorts(intPortID).lngHandle, udtCommOverlap, lngBytesRead, True) = 0

                        lngStatus = GetLastError

                        If lngStatus <> ERROR_IO_INCOMPLETE Then
                            lngBytesRead = -1
                            lngStatus = SetCommErrorEx("CommRead (GetOverlappedResult)", udtPorts(intPortID).lngHandle)
                            GoTo Routine_Exit
                        End If
                    End While
                Else
                    ' Some other error occurred.
                    lngBytesRead = -1
                    lngStatus = SetCommErrorEx("CommRead (ReadFile)", udtPorts(intPortID).lngHandle)
                    GoTo Routine_Exit

                End If
            End If
            ReDim Preserve strData(lngBytesRead)
        End If

Routine_Exit:
        CommRead = lngBytesRead
        Exit Function

Routine_Error:
        lngBytesRead = -1
        lngStatus = Err.Number
        With udtCommError
            .lngErrorCode = lngStatus
            .strFunction = "CommRead"
            .strErrorMessage = Err.Description
        End With
        Resume Routine_Exit
    End Function

    ''' <summary>
    ''' Convert a string to a byte array
    ''' </summary>
    ''' <param name="stringToConvert"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ConvertStringToByteArray(ByVal stringToConvert As String) As Byte()
        Return (New System.Text.ASCIIEncoding()).GetBytes(stringToConvert)
    End Function

    ''' <summary>
    ''' CommWriteByte - Output data to the serial port.
    ''' </summary>
    ''' <param name="intPortID"></param>
    ''' <param name="strData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CommWriteByte(ByRef intPortID As Short, ByRef strData As Byte()) As Integer

        Dim i As Short
        Dim lngStatus, lngSize As Integer
        Dim lngWrSize, lngWrStatus As Integer

        On Error GoTo Routine_Error

        ' Get the length of the data.
        lngSize = strData.Length

        ' Output the data.
        lngWrStatus = WriteFile(udtPorts(intPortID).lngHandle, strData, lngSize, lngWrSize, udtCommOverlap)

        ' Note that normally the following code will not execute because the driver
        ' caches write operations. Small I/O requests (up to several thousand bytes)
        ' will normally be accepted immediately and WriteFile will return true even
        ' though an overlapped operation was specified.

        System.Windows.Forms.Application.DoEvents()

        If lngWrStatus = 0 Then
            lngStatus = GetLastError
            If lngStatus = 0 Then
                GoTo Routine_Exit
            ElseIf lngStatus = ERROR_IO_PENDING Then
                ' We should wait for the completion of the write operation so we know
                ' if it worked or not.
                '
                ' This is only one way to do this. It might be beneficial to place the
                ' writing operation in a separate thread so that blocking on completion
                ' will not negatively affect the responsiveness of the UI.
                '
                ' If the write takes long enough to complete, this function will timeout
                ' according to the CommTimeOuts.WriteTotalTimeoutConstant variable.
                ' At that time we can check for errors and then wait some more.

                ' Loop until operation is complete.
                While GetOverlappedResult(udtPorts(intPortID).lngHandle, udtCommOverlap, lngWrSize, True) = 0

                    lngStatus = GetLastError

                    If lngStatus <> ERROR_IO_INCOMPLETE Then
                        lngStatus = SetCommErrorEx("CommWrite (GetOverlappedResult)", udtPorts(intPortID).lngHandle)
                        GoTo Routine_Exit
                    End If
                End While
            Else
                ' Some other error occurred.
                lngWrSize = -1

                lngStatus = SetCommErrorEx("CommWrite (WriteFile)", udtPorts(intPortID).lngHandle)
                GoTo Routine_Exit

            End If
        End If

        For i = 1 To 10
            System.Windows.Forms.Application.DoEvents()
        Next

Routine_Exit:
        CommWriteByte = lngWrSize
        Exit Function

Routine_Error:
        lngStatus = Err.Number
        With udtCommError
            .lngErrorCode = lngStatus
            .strFunction = "CommWrite"
            .strErrorMessage = Err.Description
        End With
        Resume Routine_Exit
    End Function

    '-------------------------------------------------------------------------------
    ' CommGetLine - Get the state of selected serial port control lines.
    '
    ' Parameters:
    '   intPortID   - Port ID used when port was opened.
    '   intLine     - Serial port line. CTS, DSR, RING, RLSD (CD)
    '   blnState    - Returns state of line (Cleared or Set).
    '
    ' Returns:
    '   Error Code  - 0 = No Error.
    '-------------------------------------------------------------------------------
    ''' <summary>
    ''' CommGetLine - Get the state of selected serial port control lines.
    ''' </summary>
    ''' <param name="intPortID"></param>
    ''' <param name="intLine"></param>
    ''' <param name="blnState"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CommGetLine(ByRef intPortID As Short, ByRef intLine As Short, ByRef blnState As Boolean) As Integer

        Dim lngStatus As Integer
        Dim lngComStatus, lngModemStatus As Integer

        On Error GoTo Routine_Error

        lngStatus = GetCommModemStatus(udtPorts(intPortID).lngHandle, lngModemStatus)

        If lngStatus = 0 Then
            lngStatus = SetCommError("CommReadCD (GetCommModemStatus)")
            GoTo Routine_Exit
        End If

        If (lngModemStatus And intLine) Then
            blnState = True
        Else
            blnState = False
        End If

        lngStatus = 0

Routine_Exit:
        CommGetLine = lngStatus
        Exit Function

Routine_Error:
        lngStatus = Err.Number
        With udtCommError
            .lngErrorCode = lngStatus
            .strFunction = "CommReadCD"
            .strErrorMessage = Err.Description
        End With
        Resume Routine_Exit
    End Function

    '-------------------------------------------------------------------------------
    ' CommSetLine - Set the state of selected serial port control lines.
    '
    ' Parameters:
    '   intPortID   - Port ID used when port was opened.
    '   intLine     - Serial port line. BREAK, DTR, RTS
    '                 Note: BREAK actually sets or clears a "break" condition on
    '                 the transmit data line.
    '   blnState    - Sets the state of line (Cleared or Set).
    '
    ' Returns:
    '   Error Code  - 0 = No Error.
    '-------------------------------------------------------------------------------
    ''' <summary>
    ''' CommSetLine - Set the state of selected serial port control lines.
    ''' </summary>
    ''' <param name="intPortID"></param>
    ''' <param name="intLine"></param>
    ''' <param name="blnState"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CommSetLine(ByRef intPortID As Short, ByRef intLine As Short, ByRef blnState As Boolean) As Integer

        Dim lngStatus As Integer
        Dim lngNewState As Integer

        On Error GoTo Routine_Error

        If intLine = LINE_BREAK Then
            If blnState Then
                lngNewState = SETBREAK
            Else
                lngNewState = CLRBREAK
            End If

        ElseIf intLine = LINE_DTR Then
            If blnState Then
                lngNewState = SETDTR
            Else
                lngNewState = CLRDTR
            End If

        ElseIf intLine = LINE_RTS Then
            If blnState Then
                lngNewState = SETRTS
            Else
                lngNewState = CLRRTS
            End If
        End If

        lngStatus = EscapeCommFunction(udtPorts(intPortID).lngHandle, lngNewState)

        If lngStatus = 0 Then
            lngStatus = SetCommError("CommSetLine (EscapeCommFunction)")
            GoTo Routine_Exit
        End If

        lngStatus = 0

Routine_Exit:
        CommSetLine = lngStatus
        Exit Function

Routine_Error:
        lngStatus = Err.Number
        With udtCommError
            .lngErrorCode = lngStatus
            .strFunction = "CommSetLine"
            .strErrorMessage = Err.Description
        End With
        Resume Routine_Exit
    End Function

    ''' <summary>
    ''' CommGetError - Get the last serial port error message.
    ''' </summary>
    ''' <param name="strMessage"></param>
    ''' <returns>Error Code  - Last serial port error code.</returns>
    ''' <remarks></remarks>
    Public Function CommGetError(ByRef strMessage As String) As Integer

        With udtCommError
            CommGetError = .lngErrorCode
            strMessage = "Error (" & CStr(.lngErrorCode) & "): " & .strFunction & " - " & .strErrorMessage
        End With

    End Function

    
End Class