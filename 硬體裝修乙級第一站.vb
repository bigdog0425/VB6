'USB連接與輸出函式庫
Private Declare Function OpenUsbDevice Lib "USBIO.dll" (ByVal a As Integer, ByVal b As Integer) As Boolean
Private Declare Sub OutDataCtrl Lib "USBIO.dll" (ByVal a As Byte, ByVal b As Byte)

'公用變數宣告
Dim a, c As Integer    'a為按鈕編號(Green:1,Red:2,Exit:3)，c為記錄燈號狀態
Dim r(20) As Byte, g(20) As Byte    'r(20)：存放紅燈的燈號狀態 g(20)：存放綠燈的燈號狀態

'載入程式：依題號預設燈號狀態，採10進位計算(以下範例為試題一)
Private Sub Form_Load()
    r(0) = 1: r(1) = 2: r(2) = 4: r(3) = 8: r(4) = 16: r(5) = 32: r(6) = 64: r(7) = 128
    g(0) = 1: g(1) = 3: g(2) = 7: g(3) = 15: g(4) = 31: g(5) = 63: g(6) = 127: g(7) = 255
End Sub

'程式結束時清除電路板燈號
Private Sub Form_Unload(Cancel As Integer)
    OutDataCtrl 0, &H0: OutDataCtrl 0, &H30    '&H0：控制板子上的綠燈，&H30：控制板子上的紅燈，0：代表熄燈
End Sub

'紀錄按鈕編號，並把燈號狀態設為0
Private Sub Command1_Click(Index As Integer)
    c = 0
    a = Index + 1
'按EXIT按鈕 則結束程式
    If a = 3 Then End  
End Sub

'定時執行：每秒(1,000 ms)
Private Sub Timer1_Timer()
'顯示系統時間
Label1.Caption = "Current Time：" & Time$

'取得連線結果
Dim conn As Boolean
conn = OpenUsbDevice(&H1234, &H6789)

'電路板連線測試，若連結成功則程式亮暗色燈
For i = 0 To 15
    Shape1(i).FillStyle = IIf(conn, 0, 1)
    Shape1(i).FillColor = IIf(i < 8, RGB(128, 0, 0), RGB(0, 128, 0))
Next

'依按鈕控制電路板亮燈
If a = 1 Then OutDataCtrl 0, &H30: OutDataCtrl g(c), 0
If a = 2 Then OutDataCtrl 0, 0: OutDataCtrl r(c), &H30

'依按鈕控制畫面亮燈控制
For i = 0 To 7
    If a = 1 And (g(c) And 2 ^ i) Then Shape1(i + 8).FillColor = RGB(0, 255, 0)
    If a = 2 And (r(c) And 2 ^ i) Then Shape1(i).FillColor = RGB(255, 0, 0)
Next

'往下一個狀態，最多15種狀態--試題9,10
c = c + 1: If c > 14 Then c = 14
End Sub
