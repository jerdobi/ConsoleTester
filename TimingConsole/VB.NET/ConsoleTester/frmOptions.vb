Public Class frmOptions
    Dim bEnable As Boolean = False

    Private Sub cboPort_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPort.SelectedIndexChanged
        If bEnable = True Then
            InterfaceFields.Port = cboPort.SelectedIndex
            RegSetValue(REGPATH, "Port", cboPort.SelectedIndex)
        End If
    End Sub

    Private Sub cboSpeed_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSpeed.SelectedIndexChanged
        If bEnable = True Then
            InterfaceFields.Speed = cboSpeed.SelectedIndex
            RegSetValue(REGPATH, "Speed", cboSpeed.SelectedIndex)
        End If
    End Sub

    Private Sub cboParity_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboParity.SelectedIndexChanged
        If bEnable = True Then
            InterfaceFields.Parity = cboParity.SelectedIndex
            RegSetValue(REGPATH, "Parity", cboParity.SelectedIndex)
        End If
    End Sub

    Private Sub cboDatabits_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDatabits.SelectedIndexChanged
        If bEnable = True Then
            InterfaceFields.DataBits = cboDatabits.SelectedIndex
            RegSetValue(REGPATH, "Databits", cboDatabits.SelectedIndex)
        End If
    End Sub

    Private Sub cboStopbits_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboStopbits.SelectedIndexChanged
        If bEnable = True Then
            InterfaceFields.StopBits = cboStopbits.SelectedIndex
            RegSetValue(REGPATH, "Stopbits", cboStopbits.SelectedIndex)
        End If
    End Sub

    Private Sub cboWait_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboWait.SelectedIndexChanged
        If bEnable = True Then
            InterfaceFields.Sleep = cboWait.SelectedIndex
            RegSetValue(REGPATH, "Wait", cboWait.SelectedIndex)
        End If
    End Sub

    Private Sub cboEventSeqType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboEventSeqType.SelectedIndexChanged
        If bEnable = True Then
            InterfaceFields.EventSequence = cboEventSeqType.SelectedIndex
            RegSetValue(REGPATH, "EventSeqType", cboEventSeqType.SelectedIndex)
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmOptions_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        bEnable = False
        cboPort.SelectedIndex = RegGetValue(REGPATH, "Port", 0)
        cboSpeed.SelectedIndex = RegGetValue(REGPATH, "Speed", 1)
        cboParity.SelectedIndex = RegGetValue(REGPATH, "Parity", 3)
        cboDatabits.SelectedIndex = RegGetValue(REGPATH, "Databits", 4)
        cboStopbits.SelectedIndex = RegGetValue(REGPATH, "Stopbits", 1)
        cboWait.SelectedIndex = RegGetValue(REGPATH, "Wait", 1)
        cboEventSeqType.SelectedIndex = RegGetValue(REGPATH, "EventSeqType", 7)
        bEnable = True
    End Sub

End Class