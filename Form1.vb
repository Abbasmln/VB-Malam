Public Class Form1


    Private Sub BtnClear_Click(sender As Object, e As EventArgs) Handles BtnClear.Click
        txtNama.Text = " "
        txtNPM.Text = " "
        txtKehadiran.Text = " "
        txtTugas.Text = " "
        txtUTS.Text = " "
        txtUAS.Text = " "
        txtNilai.Text = " "
        txtGrade.Text = " "
        txtBobot.Text = " "
    End Sub

    Private Sub BtnExit_Click(sender As Object, e As EventArgs) Handles BtnExit.Click
        End
    End Sub

    Private Sub BtnHitung_Click(sender As Object, e As EventArgs) Handles BtnHitung.Click
        Dim Nama As String
        Dim NPM As String
        Dim Kehadiran As Double
        Dim Tugas As Double
        Dim UTS As Double
        Dim UAS As Double
        Dim Nilai As Double
        Dim Grade As String

        Nama = txtNama.Text
        NPM = txtNPM.Text
        Kehadiran = Val(txtKehadiran.Text)
        Tugas = Val(txtTugas.Text)
        UTS = Val(txtUTS.Text)
        UAS = Val(txtUAS.Text)
        Nilai = ((Kehadiran / 16) * 0.1 + Tugas * 0.2 + UTS * 0.3 + UAS * 0.4)
        txtNilai.Text = Nilai

        If Nilai >= 80 Then
            Grade = "A"
            txtBobot.Text = 4
        ElseIf Nilai >= 70 Then
            Grade = "B"
            txtBobot.Text = 3
        ElseIf Nilai >= 60 Then
            Grade = "C"
            txtBobot.Text = 2
        ElseIf Nilai >= 50 Then
            Grade = "D"
            txtBobot.Text = 1
        Else
            Grade = "E"
            txtBobot.Text = 0
        End If
        txtGrade.Text = Grade
    End Sub

    Private Sub BtnIPS_Click(sender As Object, e As EventArgs) Handles BtnIPS.Click
        Dim Total As Double
        Dim SKS As Double
        Dim IPS As Double
        Total = Val(txtTotal.Text)
        SKS = Val(txtSKS.Text)
        IPS = (Total / SKS)
        txtIPS.Text = IPS
    End Sub

    Private Sub BtnIPK_Click(sender As Object, e As EventArgs) Handles BtnIPK.Click
        Dim T_IPS As Double
        Dim Semester As Double
        Dim H_IPK As Double
        T_IPS = Val(txtjumIPS.Text)
        Semester = Val(txtSemester.Text)
        H_IPK = (T_IPS / Semester)
        txtHasilIPK.Text = H_IPK
    End Sub
End Class
