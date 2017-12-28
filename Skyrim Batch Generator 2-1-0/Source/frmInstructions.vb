Public Class frmInstructions
    Private Shared OneInstance As frmInstructions

    Public Shared ReadOnly Property Instance As frmInstructions
        Get
            If OneInstance Is Nothing Then
                OneInstance = New frmInstructions
            End If
            Return OneInstance
        End Get
    End Property

    Private Sub frmInstructions_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        OneInstance = Nothing
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub
End Class