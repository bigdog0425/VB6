Public Class Form1

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim b(10), a(10), op(10), x(10), y(10), ans(10)
        Dim sno, m1, m2 As Integer

        '建立分數四則運算資料表
        Dim dt As New DataTable
        dt.Columns.Add("VALUE1")
        dt.Columns.Add("OP")
        dt.Columns.Add("VALUE2")
        dt.Columns.Add("ANSWER")

        '讀取檔案資料
        FileOpen(1, "C:\丙設磁片\940308.sm", OpenMode.Input)
        sno = 0
        Do While Not EOF(1)
            sno = sno + 1
            Input(1, b(sno)) : Input(1, a(sno)) : Input(1, op(sno)) : Input(1, y(sno)) : Input(1, x(sno))
        Loop
        FileClose(1)

        '分數四則運算--以運算元做條件
        For i = 1 To sno
            Select Case op(i)
                Case "+"
                    m1 = b(i) * x(i) + a(i) * y(i)
                    m2 = a(i) * x(i)
                Case "-"
                    m1 = b(i) * x(i) - a(i) * y(i)
                    m2 = a(i) * x(i)
                Case "*"
                    m1 = b(i) * y(i)
                    m2 = a(i) * x(i)
                Case "/"
                    m1 = b(i) * x(i)
                    m2 = a(i) * y(i)
            End Select

            '約分
            For j = 2 To m1
                If m1 Mod j = 0 And m2 Mod j = 0 Then
                    m1 = m1 / j : m2 = m2 / j
                    j = 1
                End If
            Next
            ans(i) = m1 & "/" & m2

            '檢查是否整除
            If m1 Mod m2 = 0 Then
                ans(i) = m1 / m2
            End If

            '寫入運算結果
            Dim dr As DataRow = dt.NewRow
            dr(0) = b(i) & "/" & a(i)
            dr(1) = op(i)
            dr(2) = y(i) & "/" & x(i)
            dr(3) = ans(i)
            dt.Rows.Add(dr)
        Next

        '顯示資料表
        DataGridView1.DataSource = dt

        '資料表預設資料欄位對齊中間
        DataGridView1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

    End Sub
End Class
