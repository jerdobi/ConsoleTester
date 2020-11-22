Imports DataGridViewPrinter
Imports CustomColumnAndCell
Imports WinswimCTS.ColoradoCTS
Imports System.IO

Public Class frmConsole

    Const MEET_FORWARD As Short = 1
    Const MEET_CURRENT As Short = 0
    Const MEET_BACKWARD As Short = -1
    Const gcMAXEVENTPLACES As Integer = 30
    Dim myRaces() As RACE_LIST_DATA
    Dim myMeetSelectedIndex As Integer
    Dim MyDataGridView As DataGridView

    Dim barrLoaded As Boolean = False
    Dim arrSwimmerType(10) As String
    Dim arrEventType(10) As String
    Dim arrAgeGroup(10) As String
    Dim arrBracket(10) As String
    Dim bStop As Boolean = False

    Private Sub ExitToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        AboutBox1.Show()
    End Sub

    Private Sub frmConsole_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If SubKeyExist("Software\Gary Wood Software") = False Then
            CreateSubKey("Software", "Gary Wood Software")
        End If
        If SubKeyExist(REGPATH) = False Then
            CreateSubKey("Software\Gary Wood Software", "ConsoleTester")
            RegSetValue(REGPATH, "Port", 0)
            RegSetValue(REGPATH, "Speed", 1)
            RegSetValue(REGPATH, "Parity", 3)
            RegSetValue(REGPATH, "Stopbits", 0)
            RegSetValue(REGPATH, "Databits", 4)
            RegSetValue(REGPATH, "Wait", 1)
            RegSetValue(REGPATH, "EventSeqType", 7)
        End If

        txtConsole.Top = 0
        txtConsole.Left = SplitContainer1.Panel2.Left
        txtConsole.Width = SplitContainer1.Panel2.Width
        txtConsole.Height = SplitContainer1.Panel2.Height
        Me.Text = My.Application.Info.ProductName
        Me.Show()

    End Sub

    Public Function ResetConnection() As Boolean
        bStop = False
        ResetConnection = False
        ClearConsole()
        SetInterface(InterfaceFields)
        EnablePrinter(False)
        MyDataGridView = InsertEventGrid("")
        If CTSWhoAreYou().Length > 0 Then
            If bStop = False Then
                LoadcboEventSequence()
            End If
            If bStop = False Then
                If CTSLoadcboMeet() = True Then
                    myMeetSelectedIndex = cboMeet.SelectedIndex
                End If
            End If
            If bStop = False Then
                Dim _cts As WinswimCTS.ColoradoCTS = New WinswimCTS.ColoradoCTS()
                If _cts.CTSOpen(InterfaceFields) = 0 Then
                    Dim i As Integer = CTSLoadcboEvent(_cts)
                    _cts.CTSClose(InterfaceFields)
                    cboEvent.SelectedIndex = i
                End If
            End If
            ProgressBar1.Value = 0
            ResetConnection = True
            'Else
            '    MsgBox("No console present on COM" & InterfaceFields.Port + 1, MsgBoxStyle.Critical)
        End If
        bStop = False
    End Function

    Private Sub EnablePrinter(ByVal tflag As Boolean)
        mnuPrint2.Enabled = tflag
        mnuPrintPreview2.Enabled = tflag
    End Sub

    Public Function CTSLoadcboMeet() As Boolean
        Dim tstr As String

        ProgressBar1.Value = 0
        CTSLoadcboMeet = False
        cboMeet.Items.Clear()
        Try
            Dim _cts As WinswimCTS.ColoradoCTS = New WinswimCTS.ColoradoCTS()
            If _cts.CTSOpen(InterfaceFields) = 0 Then
                Dim tMeetHeader As TYPE_MEETMGMT_HEADER
                WriteConsole(vbCrLf & "Loading Meets...Please wait...")
                Dim lngStatus As Integer = _cts.CTSGetMeet(MEET_CURRENT, tMeetHeader)
                cboMeet.SelectedIndex = cboMeet.Items.Add(Format(tMeetHeader.date_date, "Short Date") & Space(1) & Format(tMeetHeader.date_date, "Long Time"))
                Dim tCurrDate As Date = tMeetHeader.date_date
                Do Until lngStatus = _cts.GetCTSError
                    lngStatus = _cts.CTSGetMeet(MEET_BACKWARD, tMeetHeader)
                    If tMeetHeader.date_date = tCurrDate Then
                        Exit Do
                    End If
                    tstr = Format(tMeetHeader.date_date, "Short Date") & Space(1) & Format(tMeetHeader.date_date, "Long Time")
                    WriteConsole(tstr)
                    cboMeet.Items.Add(tstr)
                    AddProgressBar()
                    If bStop = True Then
                        Exit Do
                    End If
                Loop
                _cts.CTSClose(InterfaceFields)
                ProgressBar1.Value = 0
                CTSLoadcboMeet = True
            End If
        Catch ex As Exception
            MsgBox("CTSLoadcboMeet : " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
        End Try

    End Function

    Public Function CTSLoadcboEvent(ByVal _cts As WinswimCTS.ColoradoCTS, Optional ByVal iEvent As Integer = 0, Optional ByVal iHeat As Integer = 0) As Integer
        Dim i As Integer
        Dim tStr As String

        Try
            cboEvent.Items.Clear()
            CTSLoadcboEvent = 0
            ProgressBar1.Value = 0
            WriteConsole(vbCrLf & "Loading Races...Please wait...")
            myRaces = _cts.CTSGetRaceList()
            Dim myRace As RACE_LIST_DATA
            For Each myRace In myRaces
                tStr = "E=" & myRace.EventNo & " H=" & myRace.Heat
                WriteConsole(tStr & ", ", False)
                i = cboEvent.Items.Add(tStr)
                If myRace.EventNo = iEvent And myRace.Heat = iHeat Then
                    CTSLoadcboEvent = i
                End If
                AddProgressBar()
                If bStop = True Then
                    Exit For
                End If
            Next
            ProgressBar1.Value = 0
            WriteConsole(" Done!", False)
        Catch ex As Exception
            MsgBox("CTSLoadcboEvent : " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
        End Try

    End Function

    Private Function LoadcboEventSequence() As Boolean
        Dim tStr As String
        Dim i As Integer = 2

        Try
            ProgressBar1.Value = 0
            LoadcboEventSequence = False
            WriteConsole(vbCrLf & "Loading Event Sequences...Please wait...")
            Dim _cts As WinswimCTS.ColoradoCTS = New WinswimCTS.ColoradoCTS()
            If _cts.CTSOpen(InterfaceFields) = 0 Then
                tStr = _cts.CTSGetSequence(1)
                Do While Len(tStr) > 0
                    WriteConsole(tStr)
                    cboEventSequence.Items.Add(tStr)
                    tStr = _cts.CTSGetSequence(i)
                    i = i + 1
                    AddProgressBar()
                    If bStop = True Then
                        Exit Do
                    End If
                Loop
                If bStop = False Then
                    cboEventSequence.SelectedIndex = InterfaceFields.EventSequence - 1
                    LoadcboEventSequence = True
                End If
            Else
                WriteConsole("CTS console not responding on COM" & CStr(InterfaceFields.Port + 1))
            End If
            _cts.CTSClose(InterfaceFields)
            ProgressBar1.Value = ProgressBar1.Maximum
        Catch ex As Exception
            MsgBox("LoadcboEventSequence : " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
        End Try

    End Function


    Private Sub cboMeet_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboMeet.SelectedIndexChanged
        Dim i As Integer = 0

        bStop = False
        If cboMeet.SelectedIndex = -1 Or _
            cboMeet.SelectedIndex = myMeetSelectedIndex Then
            Exit Sub
        End If

        Dim iMove As Integer
        Dim iDiff As Integer
        If cboMeet.SelectedIndex > myMeetSelectedIndex Then
            iMove = MEET_FORWARD
            iDiff = cboMeet.SelectedIndex - myMeetSelectedIndex
        Else
            iMove = MEET_BACKWARD
            iDiff = myMeetSelectedIndex - cboMeet.SelectedIndex
        End If
        WriteConsole(vbCrLf & "Moving meet pointer " & IIf((iMove = MEET_FORWARD), "+", "-") & iDiff.ToString)
        Try
            Dim _cts As WinswimCTS.ColoradoCTS = New WinswimCTS.ColoradoCTS()
            If _cts.CTSOpen(InterfaceFields) = 0 Then
                Dim tMeetHeader As TYPE_MEETMGMT_HEADER
                Do Until iDiff = 0
                    _cts.CTSGetMeet(iMove, tMeetHeader)
                    iDiff -= 1
                    If bStop = True Then
                        Exit Do
                    End If
                Loop
                If bStop = False Then
                    myMeetSelectedIndex = cboMeet.SelectedIndex
                    i = CTSLoadcboEvent(_cts)
                End If
                _cts.CTSClose(InterfaceFields)
                cboEvent.SelectedIndex = i
            End If
        Catch ex As Exception
            MsgBox("cboMeet_SelectedIndexChanged : " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
        End Try

    End Sub

    Private Sub cboEvent_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboEvent.SelectedIndexChanged
        bStop = False
        If cboEvent.SelectedIndex <> -1 Then
            CTSGetRace("E", myRaces(cboEvent.SelectedIndex).EventNo, myRaces(cboEvent.SelectedIndex).Heat)
        End If
    End Sub

    Private Function CTSGetRace(ByRef tType As String, Optional ByRef tEvent As Short = 0, Optional ByRef tHeat As Short = 0) As Short
        Dim tHeader As TYPE_MEETMGMT_HEADER
        Dim tLaneData(gcMAXEVENTPLACES) As TYPE_LANE_DATA
        Dim lngStatus As Integer
        Dim i As Short

        Try
            Dim _cts As WinswimCTS.ColoradoCTS = New WinswimCTS.ColoradoCTS()
            If _cts.CTSOpen(InterfaceFields) = 0 Then
                _cts.CTSOpen(InterfaceFields)
                lngStatus = _cts.CTSGetSwimDataHeader(tType, tHeader, tLaneData, tEvent, tHeat, (chkBackup.CheckState))
                _cts.CTSClose(InterfaceFields)
            End If

            If lngStatus < 8 Then
                WriteConsole("Event : " & tEvent.ToString & " not found!")
                CTSGetRace = -1
                Exit Function
            End If

            MyDataGridView = InsertRaceGrid(tEvent.ToString & " Heat " & tHeat.ToString)

            For i = 0 To gcMAXEVENTPLACES - 1
                If tLaneData(i).place = 0 Then
                    Exit For
                End If
                MyDataGridView.Rows.Add(i + 1, tLaneData(i).place, CSngTime(tLaneData(i).mins & ":" & tLaneData(i).secs & "." & tLaneData(i).hsecs))
            Next i

            SplitContainer1.Panel1.Refresh()

            'If tHeader.event_Renamed <> SaveGetTimeInfo.giEvent Or tHeader.heat <> SaveGetTimeInfo.giHeat Then
            '    MsgBox(GetLang(Lang_839), MsgBoxStyle.Critical, My.Application.Info.Title)
            '    CTSGetRace = -1
            'End If

        Catch ex As Exception
            MsgBox("CTSGetRace : " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
        End Try

    End Function

    Private Sub CreateLabel(ByVal tTitle As String, ByVal tPanel As Windows.Forms.Panel)
        ' Create a label
        Dim lbl As New Label()

        lbl.Text = tTitle
        lbl.Dock = DockStyle.Top
        lbl.BackColor = Color.DarkSlateBlue
        lbl.ForeColor = Color.White
        lbl.Dock = System.Windows.Forms.DockStyle.Top
        lbl.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        tPanel.Controls.Add(lbl)
    End Sub

    Private Sub AddProgressBar(Optional ByVal iVal As Integer = 1)
        Try
            ProgressBar1.Value += 1
        Catch ex As Exception
            ProgressBar1.Value = 0
        End Try
    End Sub

    Private Sub chkBackup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBackup.Click
        bStop = False
        If cboEvent.SelectedIndex <> -1 Then
            CTSGetRace("E", myRaces(cboEvent.SelectedIndex).EventNo, myRaces(cboEvent.SelectedIndex).Heat)
        End If
    End Sub


    Private Sub DisplayEventTypeSeq()
        Dim K As Short
        Dim tStr As String

        Try
            ProgressBar1.Value = 0
            ClearConsole()
            Dim _cts As WinswimCTS.ColoradoCTS = New WinswimCTS.ColoradoCTS()
            If _cts.CTSOpen(InterfaceFields) = 0 Then
                WriteConsole(vbCrLf & "Swimmer Type:")
                '\nCTS index &1 = "&2"\n--> Set "&3" CTSKey to "&1" for record specified for "&2"
                For K = 0 To 9 - 1
                    tStr = _cts.CTSGetSequenceType(WinswimCTS.ColoradoCTS.CTSEventType.swimmerType, K)
                    arrSwimmerType(K) = tStr
                    If Len(tStr) = 0 Then Exit For
                    WriteConsole(tStr)
                    AddProgressBar()
                    If bStop = True Then
                        Exit For
                    End If
                Next K
                WriteConsole(vbCrLf & "Event Type:")
                For K = 0 To 9 - 1
                    tStr = _cts.CTSGetSequenceType(WinswimCTS.ColoradoCTS.CTSEventType.EventType, K)
                    arrEventType(K) = tStr
                    If Len(tStr) = 0 Then Exit For
                    WriteConsole(tStr)
                    AddProgressBar()
                    If bStop = True Then
                        Exit For
                    End If
                Next K
                WriteConsole(vbCrLf & "Age Group:")
                For K = 0 To 20 - 1
                    tStr = _cts.CTSGetSequenceType(WinswimCTS.ColoradoCTS.CTSEventType.AgeGroup, K)
                    arrAgeGroup(K) = tStr
                    If Len(tStr) = 0 Then Exit For
                    WriteConsole(tStr)
                    AddProgressBar()
                    If bStop = True Then
                        Exit For
                    End If
                Next K
                WriteConsole(vbCrLf & "Bracket:")
                For K = 0 To 9 - 1
                    tStr = _cts.CTSGetSequenceType(WinswimCTS.ColoradoCTS.CTSEventType.bracket, K)
                    arrBracket(K) = tStr
                    If Len(tStr) = 0 Then Exit For
                    WriteConsole(tStr)
                    AddProgressBar()
                    If bStop = True Then
                        Exit For
                    End If
                Next K
                _cts.CTSClose(InterfaceFields)
                ProgressBar1.Value = 0
                If bStop = False Then
                    barrLoaded = True
                End If
            End If
        Catch ex As Exception
            MsgBox("DisplayEventTypeSeq : " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        bStop = False
        mnuView.Enabled = ResetConnection()
    End Sub

    Private Sub DisplayEventsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DisplayEventsToolStripMenuItem.Click

        Dim i As Integer
        Dim iEventType As Integer
        Dim iSwimmerType As Integer
        Dim iBracket As Integer
        Dim iAgeGroup As Integer
        Dim tStr As String
        Dim iEvent As Integer
        Dim myWriter As StreamWriter
        Dim myStream As FileStream

        bStop = False

        ClearConsole()

        If barrLoaded = False Then
            DisplayEventTypeSeq()
        End If

        If bStop = True Then
            Exit Sub
        End If

        MyDataGridView = InsertEventGrid(cboEventSequence.Items(cboEventSequence.SelectedIndex))

        Try
            ProgressBar1.Value = 0
            Dim tPath As String = Path.GetTempFileName()
            myStream = New FileStream(tPath, FileMode.Create)
            myWriter = New StreamWriter(myStream)
            ' TAGS Short Course Championships;03/08/2001;03/11/2001;03/08/2001;Y;Loos Natatorium, Dallas, TX;;Hy-Tek Sports Software;1.4i;A;4153G
            'meetdesc;startdate;enddate;startdate;Y|L|S;teamname;city;state;software
            With myWriter
                .WriteLine(cboEventSequence.Items(cboEventSequence.SelectedIndex) _
                & ";" & Format(cboMeet.Items(cboMeet.SelectedIndex), "Short Date") _
                & ";" & Format(cboMeet.Items(cboMeet.SelectedIndex), "Short Date") _
                & ";" & Format(cboMeet.Items(cboMeet.SelectedIndex), "Short Date") _
                & ";" & "Y" _
                & ";" _
                & ";" _
                & ";" _
                & ";ConsoleTester")
            End With
            WriteConsole(vbCrLf & vbCrLf & "Listing Events:")
            Dim _cts As WinswimCTS.ColoradoCTS = New WinswimCTS.ColoradoCTS()
            If _cts.CTSOpen(InterfaceFields) = 0 Then
                _cts.CTSSetSequence(InterfaceFields.EventSequence)
                Dim tByte(1) As Byte

                For iEvent = 1 To _cts.MaxEvents
                    Dim iLen As Integer = _cts.CTSGetEvent(iEvent, tByte)
                    If iLen = _cts.GetCTSError Then
                        Exit For
                    End If
                    tStr = tByte(3).ToString & " - " & tByte(5).ToString & Space(1)
                    iEventType = -1
                    For i = 7 To iLen
                        If tByte(i) = 0 Then
                            Exit For
                        ElseIf tByte(i) < 10 Then
                            iSwimmerType = tByte(i)
                            tStr = tStr & arrSwimmerType(iSwimmerType) & " "
                        ElseIf tByte(i) < 20 Then
                            iEventType = tByte(i) - 10
                            tStr = tStr & arrEventType(iEventType) & " "
                        ElseIf tByte(i) < 30 Then
                            iAgeGroup = tByte(i) - 20
                            tStr = tStr & arrAgeGroup(iAgeGroup) & " "
                        ElseIf tByte(i) < 40 Then
                            iBracket = tByte(i) - 30
                            tStr = tStr & arrBracket(iBracket) & " "
                        End If
                    Next i
                    If iEventType = -1 Then
                        Exit For
                    End If
                    'event;bracket;sex;eventtype;minage;maxage;distance;strokekey;lowtime;hightime;;entryfee
                    Dim tarrAge() As String = getMaxAge(arrAgeGroup(iAgeGroup))
                    Dim iMinAge As Integer = 0
                    Dim iMaxAge As Integer = 0
                    Select Case tarrAge.Length
                        Case 0
                            iMinAge = 0
                            iMaxAge = 0
                        Case 1
                            iMinAge = 0
                            iMaxAge = CInt(tarrAge(0))
                        Case 2
                            iMinAge = CInt(tarrAge(0))
                            iMaxAge = CInt(tarrAge(1))
                    End Select
                    With myWriter
                        .WriteLine(iEvent.ToString _
                        & ";" & arrBracket(iBracket).Substring(0, 1) _
                        & ";" & arrSwimmerType(iSwimmerType).Substring(0, 1) _
                        & ";" & arrEventType(iEventType).Substring(0, 1) _
                        & ";" & iMinAge.ToString _
                        & ";" & iMaxAge.ToString _
                        & ";" & tByte(5).ToString _
                        & ";" & (iEventType + 1).ToString _
                        & ";0;;0;0")
                    End With
                    MyDataGridView.Rows.Add(iEvent, tByte(5).ToString, arrSwimmerType(iSwimmerType), arrAgeGroup(iAgeGroup), arrEventType(iEventType), arrBracket(iBracket))
                    SplitContainer1.Panel1.Refresh()
                    AddProgressBar()
                    WriteConsole(tStr)
                    If bStop = True Then
                        Exit For
                    End If
                Next iEvent

                _cts.CTSClose(InterfaceFields)
                ProgressBar1.Value = 0
                myWriter.Close()
                myStream.Close()
                If bStop = False Then
                    If MsgBox("Export to HYV?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        Dim saveFileDialog As New SaveFileDialog()
                        saveFileDialog.Title = "Save HYV File"
                        saveFileDialog.AddExtension = True
                        saveFileDialog.DefaultExt = "hyv"
                        saveFileDialog.OverwritePrompt = True
                        saveFileDialog.Filter = "Hy-tek HYV (*.hyv)|*.hyv|All Files (*.*)|*.*"
                        If saveFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                            My.Computer.FileSystem.CopyFile(tPath, saveFileDialog.FileName, True)
                            MsgBox(saveFileDialog.FileName & " Saved!", MsgBoxStyle.Information)
                        End If
                        saveFileDialog.Dispose()
                    End If
                End If
                My.Computer.FileSystem.DeleteFile(tPath)
            End If
        Catch ex As Exception
            MsgBox("DisplayEventsToolStripMenuItem_Click : " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
        End Try

    End Sub

    Private Function getMaxAge(ByVal tAgeGroup As String) As String()
        Dim i As Integer = 0
        Dim x As Integer = 0
        Dim j As Integer = 0
        Dim tstr(10) As String

        j = 0
        Dim tchar As Char() = tAgeGroup.ToCharArray
        For i = 0 To tchar.Length - 1
            If IsNumeric(tchar(i)) Then
                tstr(j) = ""
                For x = i To tchar.Length - 1
                    If IsNumeric(tchar(x)) Then
                        tstr(j) = tstr(j) & tchar(x)
                    Else
                        Exit For
                    End If
                Next x
                j += 1
                i = x
            End If
        Next i
        ReDim Preserve tstr(j - 1)
        getMaxAge = tstr

    End Function

    Private Function InsertRaceGrid(ByVal tLabel As String) As DataGridView
        EnablePrinter(True)
        SplitContainer1.Panel1.Controls.Clear()

        Dim dgv As DataGridView = New DataGridView()
        dgv.Dock = DockStyle.Fill
        dgv.ColumnCount = 0
        SplitContainer1.Panel1.Controls.Add(dgv)
        If tLabel.Length > 0 Then
            CreateLabel("Event " & tLabel, SplitContainer1.Panel1)
        End If

        Dim laneColumn As New DataGridViewTextBoxColumn
        With laneColumn
            .HeaderText = "Lane"
            .Name = "Lane"
        End With
        Dim placeColumn As New DataGridViewTextBoxColumn
        With placeColumn
            .HeaderText = "Place"
            .Name = "Place"
        End With
        'Dim timeColumn As New DataGridViewTextBoxColumn
        'With timeColumn
        '    .HeaderText = "Time"
        '    .Name = "Time"
        'End With
        Dim timeColumn As New CustomColumnAndCell.TimeColumn
        With timeColumn
            .HeaderText = "Time"
            .Name = "Time"
        End With
        
        dgv.Columns.Insert(0, placeColumn)
        dgv.Columns.Insert(1, laneColumn)
        dgv.Columns.Insert(2, timeColumn)

        dgv.ReadOnly = True
        InsertRaceGrid = dgv
    End Function

    Private Function InsertEventGrid(ByVal tLabel As String) As DataGridView
        EnablePrinter(True)
        SplitContainer1.Panel1.Controls.Clear()
        Dim dgv As DataGridView = New DataGridView()
        dgv.Dock = DockStyle.Fill
        dgv.ColumnCount = 0
        SplitContainer1.Panel1.Controls.Add(dgv)

        If tLabel.Length > 0 Then
            CreateLabel("Events : " & tLabel, SplitContainer1.Panel1)
        End If

        Dim eventColumn As New DataGridViewTextBoxColumn
        With eventColumn
            .HeaderText = "Event"
            .Name = "Event"
        End With
        Dim distanceColumn As New DataGridViewTextBoxColumn
        With distanceColumn
            .HeaderText = "Distance"
            .Name = "Distance"
        End With
        Dim SwimmerTypeColumn As New DataGridViewTextBoxColumn
        With SwimmerTypeColumn
            .HeaderText = "SwimmerType"
            .Name = "SwimmerType"
        End With
        Dim AgeGroupColumn As New DataGridViewTextBoxColumn
        With AgeGroupColumn
            .HeaderText = "AgeGroup"
            .Name = "AgeGroup"
        End With
        Dim EventTypeColumn As New DataGridViewTextBoxColumn
        With EventTypeColumn
            .HeaderText = "EventType"
            .Name = "EventType"
        End With
        Dim BracketColumn As New DataGridViewTextBoxColumn
        With BracketColumn
            .HeaderText = "Bracket"
            .Name = "Bracket"
        End With
       
        dgv.Columns.Insert(0, eventColumn)
        dgv.Columns.Insert(1, distanceColumn)
        dgv.Columns.Insert(2, SwimmerTypeColumn)
        dgv.Columns.Insert(3, AgeGroupColumn)
        dgv.Columns.Insert(4, EventTypeColumn)
        dgv.Columns.Insert(5, BracketColumn)

        dgv.ReadOnly = True
        dgv.Dock = DockStyle.Fill

        InsertEventGrid = dgv
    End Function

    Private Sub DisplayEventSeqTypesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DisplayEventSeqTypesToolStripMenuItem.Click
        bStop = False
        DisplayEventTypeSeq()
    End Sub

    Private Sub ToolStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ToolStrip1.ItemClicked
        bStop = True
    End Sub

    Private Sub mnuOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOptions.Click
        bStop = False
        frmOptions.Show()
    End Sub

    Private Sub SplitContainer1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles SplitContainer1.Resize
        txtConsole.Top = 0
        txtConsole.Left = SplitContainer1.Panel2.Left
        txtConsole.Width = SplitContainer1.Panel2.Width
        txtConsole.Height = SplitContainer1.Panel2.Height
        txtConsole.Refresh()
    End Sub

    Private Sub SplitContainer1_SplitterMoved(ByVal sender As Object, ByVal e As System.Windows.Forms.SplitterEventArgs) Handles SplitContainer1.SplitterMoved
        txtConsole.Top = 0
        txtConsole.Left = SplitContainer1.Panel2.Left
        txtConsole.Width = SplitContainer1.Panel2.Width
        txtConsole.Height = SplitContainer1.Panel2.Height
        txtConsole.Refresh()
    End Sub

    Private Function SetupThePrinting() As Boolean
        Dim MyPrintDialog As PrintDialog = New PrintDialog()
        MyPrintDialog.AllowCurrentPage = False
        MyPrintDialog.AllowPrintToFile = False
        MyPrintDialog.AllowSelection = False
        MyPrintDialog.AllowSomePages = False
        MyPrintDialog.PrintToFile = False
        MyPrintDialog.ShowHelp = False
        MyPrintDialog.ShowNetwork = False

        If MyPrintDialog.ShowDialog() <> Windows.Forms.DialogResult.OK Then
            Return False
        End If

        MyPrintDocument.DocumentName = "Console Report"
        MyPrintDocument.PrinterSettings = MyPrintDialog.PrinterSettings
        MyPrintDocument.DefaultPageSettings = MyPrintDialog.PrinterSettings.DefaultPageSettings
        MyPrintDocument.DefaultPageSettings.Margins = New System.Drawing.Printing.Margins(40, 40, 40, 40)

        'If MsgBox("Do you want the report to be centered on the page", MsgBoxStyle.YesNo, "InvoiceManager - Center on Page") = MsgBoxResult.Yes Then
        '    Dim MyDataGridViewPrinter As DataGridViewPrinter.DataGridViewPrinter = New DataGridViewPrinter.DataGridViewPrinter(MyDataGridView, MyPrintDocument, True, True, My.Application.Info.ProductName, New Font("Tahoma", 18, FontStyle.Bold, GraphicsUnit.Point), Color.Black, True)
        'Else
        Dim MyDataGridViewPrinter As DataGridViewPrinter.DataGridViewPrinter = New DataGridViewPrinter.DataGridViewPrinter(MyDataGridView, MyPrintDocument, False, True, My.Application.Info.ProductName, New Font("Tahoma", 18, FontStyle.Bold, GraphicsUnit.Point), Color.Black, True)
        'End If
        Return True
    End Function

    Private Sub mnuPrint2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPrint2.Click
        If (SetupThePrinting()) Then
            MyPrintDocument.Print()
        End If
    End Sub

    Private Sub mnuPrintPreview2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPrintPreview2.Click
        If (SetupThePrinting()) Then
            Dim myPrintPreviewDialog As PrintPreviewDialog = New PrintPreviewDialog()
            myPrintPreviewDialog.Document = MyPrintDocument
            myPrintPreviewDialog.ShowDialog()
        End If
    End Sub

    Private Sub OpenConnectionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenConnectionToolStripMenuItem.Click
        mnuView.Enabled = ResetConnection()
    End Sub

    Private Sub SplitContainer1_Panel2_Paint(sender As Object, e As PaintEventArgs) Handles SplitContainer1.Panel2.Paint

    End Sub
End Class

