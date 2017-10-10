Public Class NG_FORM

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If Me.BackColor = Color.White Then
            Me.BackColor = Color.Tomato
            Label1.ForeColor = Color.Yellow
            Lb_inform_NG.ForeColor = Color.Yellow
        Else
            Me.BackColor = Color.White
            Label1.ForeColor = Color.Red
            Lb_inform_NG.ForeColor = Color.Red
        End If
    End Sub

    Private Sub NG_FORM_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed

        Control.Show()
        AppActivate(My.Application.ToString)
    End Sub


End Class