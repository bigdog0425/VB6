Public Class Form1
    Dim s(50, 3)
    Dim sno

    Private Sub Form1_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated
        Call rdata()
        For i = 1 To sno
            If s(i, 0) = "" Then Call sp1(i)
            If s(i, 0) = "" Then Call sp2(i)
            If s(i, 0) = "" Then Call sp3(i)
        Next
        Call wdata()
    End Sub
   
    Sub rdata()
        FileOpen(1, "c:\丙設磁片\940306.sm", OpenMode.Input)
        sno = 0
        Do While Not EOF(1)
            sno = sno + 1
            For j = 1 To 3
                Input(1, s(sno, j))
            Next
        Loop
        FileClose(1)
        TextBox1.Text = sno & "," & s(sno, 1)    
    End Sub

    Sub sp1(ByVal i)
        If Not s(i, 1) Like "[A-Z]#########" Then s(i, 0) = "FORMAT ERROR"
    End Sub
    
    Sub sp2(ByVal i)
        Dim sexcode
        sexcode = Mid(s(i, 1), 2, 1) & s(i, 3)
        If sexcode <> "1M" And sexcode <> "2F" Then s(i, 0) = "SEX CODE ERROR"
    End Sub
    
    Sub sp3(ByVal i)
        Dim x, x1, x2, SS
        Dim L1 As String = Mid(s(i, 1), 1, 1)
        Dim s26 As String = "ABCDEFGHJKLMNPQRSTUVXYWZIO"
        Dim d(9)
        x = InStr(s26, L1)
        x1 = x \ 10
        x2 = x Mod 10
        For i = 1 To 9
            d(i) = Mid(s(i, 1), i + 1, 1)
        Next
        SS = x1 + x2 * 9 + d(1) * 8 + d(2) * 7 + d(3) * 6 + d(4) * 5 + d(5) * 4 + d(6) * 3 + d(7) * 2 + d(8) + d(9)
        If SS Mod 10 <> 0 Then s(i, 0) = "CHECK SUM ERROR"
    End Sub
    
    Sub wdata()
        Dim dt As New DataTable
        dt.Columns.Add("ID_NO")
        dt.Columns.Add("NAME")
        dt.Columns.Add("SEX")
        dt.Columns.Add("ERROR")
        For i = 1 To sno
            Dim dr As DataRow = dt.NewRow
            dr(0) = s(i, 1)
            dr(1) = s(i, 2)
            dr(2) = s(i, 3)
            dr(3) = s(i, 0)
            dt.Rows.Add(dr)
        Next
        DGV.DataSource = dt
        DGV.Sort(DGV.Columns(0), 0)
        DGV.Columns(3).Width = 150
    End Sub
End Class
