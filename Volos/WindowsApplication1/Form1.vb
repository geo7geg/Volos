Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        con = CreateObject("ADODB.Connection")
        con.CursorLocation = 3
        con.ConnectionString = "Provider=WinCCOLEDBProvider.1;Catalog=CC_TEST_SQL_15_03_17_12_02_10R;Data Source=VMVER1PC\WINCC"
        con.Open()
    End Sub
End Class
