Public Class Form1
    
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim s(50, 3), d(9)
        Dim L1, S26, x1, x2, sno, m1, m
        FileOpen(1, "C:\丙設磁片\940306.T01", OpenMode.Input)
        sno = 0
        
        '檢查檔案內每一筆資料
        Do While Not EOF(1)
            sno = sno + 1
            Input(1, s(sno, 0)) : Input(1, s(sno, 1)) : Input(1, s(sno, 2)) '讀出身分證、姓名、性別

            '檢查格式錯誤
            If Not s(sno, 0) Like "[A-Z]#########" Then
                s(sno, 3) = "FORMAT ERROR"

            '檢查性別碼錯誤
            ElseIf Mid(s(sno, 0), 2, 1) & s(sno, 2) <> "1M" And Mid(s(sno, 0), 2, 1) & s(sno, 2) <> "2F" Then
                s(sno, 3) = "SEX CODE ERROR"
                
            '檢查身分證號碼錯誤
            Else               
                '取出第一碼英文並轉成Ｘ1,X2
                L1 = Mid(s(sno, 0), 1, 1)'取出第一個碼
                S26 = "ABCDEFGHJKLMNPQRSTUVXYWZIO" '英文排序
                m1 = InStr(S26, L1) + 9 '取出第一碼的序號
                x1 = m1 \ 10 'x1=十位數
                x2 = m1 Mod 10 'x2=個位數

                '取出D1~D9
                For i = 1 To 9
                    d(i) = Mid(s(sno, 0), i + 1, 1)
                Next

                '計算身分證檢核碼
                m = x1 + 9 * x2 + 8 * d(1) + 7 * d(2) + 6 * d(3) + 5 * d(4) + 4 * d(5) + 3 * d(6) + 2 * d(7) + d(8) + d(9)
            
                '檢查檢核碼是否錯誤
                If m Mod 10 <> 0 Then
                    s(sno, 3) = "CHECK SUM ERROR"
                Else
                    s(sno, 3) = ""
                End If
            End If
        Loop
        FileClose(1)
    
        '建立一個資料表（ID_NO，NAME，SEX，ERROR）
        Dim dt As New DataTable
        dt.Columns.Add("ID_NO") '增加ID_NO欄位
        dt.Columns.Add("NAME") '增加NAME欄位
        dt.Columns.Add("SEX")'增加SEX欄位
        dt.Columns.Add("ERROR")'增加ERROR欄位

        '將陣列內資料寫入資料表
        For i = 1 To sno
            Dim dr As DataRow = dt.NewRow
            dr(0) = s(i, 0) '寫入身分證
            dr(1) = s(i, 1) '寫入姓名
            dr(2) = s(i, 2) '寫入性別
            dr(3) = s(i, 3) '寫入檢查碼
            dt.Rows.Add(dr) 'dt加入一筆資料
        Next

        '依身分證號碼排序資料表
        DGV.DataSource = dt '指定資料來源為dt
        DGV.Sort(DGV.Columns(0), 0) '排序:依身分證字號(欄位0)

        '指定欄位寬度(非必備)
        DGV.Columns(1).Width = 70
        DGV.Columns(2).Width = 50
        DGV.Columns(3).Width = 120
    End Sub
End Class
