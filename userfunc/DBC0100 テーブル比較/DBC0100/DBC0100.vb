Imports skysystem.common


Public Class DBC0100
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim dbc As New DBconnection

            dbCommon.sqlsvrSqlUtil.GetViewDifine(dbc, "V_SYS_USER")

        Catch ex As Exception
            MsgBox("エラーダヨ")
        End Try
    End Sub
End Class
