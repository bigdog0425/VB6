Public Class Form1
    Dim car(50, 5)
    Dim C(7, 3)
    Dim carno, weekall



    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        FileOpen(1, "c:\丙設磁片\940307.sm", OpenMode.Input)
        carno = 0
        Do While Not EOF(1)
            carno = carno + 1
            Input(1, car(carno, 1)) : Input(1, car(carno, 2)) : Input(1, car(carno, 3)) : Input(1, car(carno, 4)) : Input(1, car(carno, 5))
            car(carno, 0) = car(carno, 2) + car(carno, 3) + car(carno, 4) + car(carno, 5)
            car(0, 2) = car(0, 2) + car(carno, 2)
            car(0, 3) = car(0, 3) + car(carno, 3)
            car(0, 4) = car(0, 4) + car(carno, 4)
            car(0, 5) = car(0, 5) + car(carno, 5)
            weekall = weekall + car(carno, 0)
        Loop

        C(1, 1) = C11 : C(1, 2) = C12 : C(1, 3) = C13
        C(2, 1) = C21 : C(2, 2) = C22 : C(2, 3) = C23
        C(3, 1) = C31 : C(3, 2) = C32 : C(3, 3) = C33
        C(4, 1) = C41 : C(4, 2) = C42 : C(4, 3) = C43
        C(5, 1) = C51 : C(5, 2) = C52 : C(5, 3) = C53
        C(6, 1) = C61 : C(6, 2) = C62 : C(6, 3) = C63
        C(7, 1) = C71 : C(7, 2) = C72 : C(7, 3) = C73
        FileClose(1)
        GBox1.Hide()

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim weekfl(7), weekname(7)
        weekname(1) = "星期一"
        weekname(2) = "星期二"
        weekname(3) = "星期三"
        weekname(4) = "星期四"
        weekname(5) = "星期五"
        weekname(6) = "星期六"
        weekname(7) = "星期日"
        GBox1.Text = "依星期別"
        GBox1.Show()

        For i = 1 To carno
            Select Case car(i, 1)
                Case "MONDAY"
                    weekfl(1) = weekfl(1) + car(i, 0)
                Case "TUESDAY"
                    weekfl(2) = weekfl(2) + car(i, 0)
                Case "WEDNESDAY"
                    weekfl(3) = weekfl(3) + car(i, 0)
                Case "THURSDAY"
                    weekfl(4) = weekfl(4) + car(i, 0)
                Case "FRIDAY"
                    weekfl(5) = weekfl(5) + car(i, 0)
                Case "SATURDAY"
                    weekfl(6) = weekfl(6) + car(i, 0)
                Case "SUNDAY"
                    weekfl(7) = weekfl(7) + car(i, 0)
            End Select
        Next


        For i = 1 To 7
            C(i, 1).text = weekname(i)
            C(i, 2).Width = 200 * weekfl(i) / weekall
            C(i, 3).text = Format(weekfl(i), "###,###,###")
            C(i, 3).left = C(i, 2).left + C(i, 2).Width + 5

            '顯示1~7排資料
            For j = 1 To 3
                C(i, j).show()
            Next

        Next

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Dim carsort(4, 2)
        Dim temp1, temp2
        carsort(1, 1) = "大型車" : carsort(1, 2) = car(0, 2) '車輛排序:第一欄車型，第2欄流量
        carsort(2, 1) = "中型車" : carsort(2, 2) = car(0, 3)
        carsort(3, 1) = "小型車" : carsort(3, 2) = car(0, 4)
        carsort(4, 1) = "公務車" : carsort(4, 2) = car(0, 5)

        GBox1.Text = "依車輛種類"
        GBox1.Show()

        '依流量大小排序
        For j = 1 To 3
            For i = 1 To 3
                If carsort(i, 2) > carsort(i + 1, 2) Then
                    temp1 = carsort(i, 1) : temp2 = carsort(i, 2)
                    carsort(i, 1) = carsort(i + 1, 1) : carsort(i, 2) = carsort(i + 1, 2)
                    carsort(i + 1, 1) = temp1 : carsort(i + 1, 2) = temp2
                End If
            Next
        Next

        '隱藏5~7排資料
        For i = 5 To 7
            For j = 1 To 3
                C(i, j).hide()
            Next
        Next

        '顯示第1~4排資料
        For i = 1 To 4
            C(i, 1).text = carsort(i, 1)
            C(i, 2).Width = 200 * carsort(i, 2) / weekall
            C(i, 3).text = Format(carsort(i, 2), "###,###,###")
            C(i, 3).left = C(i, 2).left + C(i, 2).Width + 5

        Next
    End Sub
End Class
