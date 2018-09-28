Imports System.Data.SqlClient

Module MATCHFILE
    Dim S1(12) As String
    Dim c01(), c02(), c03(), c04(), c05(), c06(), c07(), c08(), c09(), c10(), c11(), c12() As String
    Dim m_planet As String = "MA"
    Dim m_key As String
    Dim InsertQuery = ""
    Dim InsertCounter = 0
    'Dim connstr = "data source=49.50.103.132;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=False;User Id=sa;password=pSI)TA1t0K[)"
    Dim connstr = "data source=WIN-KSTUPT6CJRC;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=True;multipleactiveresultsets=True;"
    'Dim connstr = "data source=DESKTOP-JBRFH9E;initial catalog=testdb;integrated security=True;"
    Dim SelectF2PLANET = "SELECT * FROM F2PLANETS WHERE p0 = 'MA'"
    Sub Main()

        Dim starttime As DateTime = DateTime.Now
        EmailNotify.SendEmail("MATCHFILE_MA (NEW sql WAY) 50 threads Started", starttime, "start")
        Console.WriteLine(starttime)
        MakeCusp_checkHouse()
        EmailNotify.SendEmail("MATCHFILE_MA (NEW sql WAY) 50 threads Ended", starttime, "end")
        Console.WriteLine("TOTAL TIME TAKEN IS : " + DateTime.Now.Subtract(starttime).ToString())
        'Console.ReadKey()
    End Sub
    Sub MakeCusp_checkHouse()
        Dim M_uid As String = "XXXXXXXXXX"
        Dim M_Hid As String = "1000000001"
        Dim SelectCUSP = "SELECT * FROM CUSP"
        Dim con As New SqlConnection(connstr)
        Dim connection As SqlConnection = New SqlConnection(connstr)
        connection.Open()
        Dim cmd As New SqlCommand(SelectCUSP, con)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        da.Fill(ds)

        For Each row As DataRow In ds.Tables(0).Rows
            Dim l As Integer = row.Item(3).ToString().Trim.Length / 2
            Select Case row.Item(2).ToString()
                Case "01"
                    c01 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c01(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next

                Case "02"
                    c02 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c02(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next
                Case "03"
                    c03 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c03(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next
                Case "04"
                    c04 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c04(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next
                Case "05"
                    c05 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c05(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next
                Case "06"
                    c06 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c06(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next
                Case "07"
                    c07 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c07(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next
                Case "08"
                    c08 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c08(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next
                Case "09"
                    c09 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c09(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next
                Case "10"
                    c10 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c10(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next
                Case "11"
                    c11 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c11(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next
                Case "12"
                    c12 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c12(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next
                Case Else
            End Select

        Next

        connection.Close()

        Dim a8(8) As String
        Dim con0 As New SqlConnection(connstr)
        Dim connection0 As SqlConnection = New SqlConnection(connstr)
        connection0.Open()
        Dim cmd0 As New SqlCommand(SelectF2PLANET, con0)
        Dim da0 As New SqlDataAdapter(cmd0)
        Dim ds0 As New DataSet()
        Console.WriteLine("Select from F2PLANETS FOR ME STARTED AT: " + DateTime.Now.ToString())
        da0.Fill(ds0)
        Console.WriteLine("Select from F2PLANETS ENDED AT: " + DateTime.Now.ToString())
        connection.Close()
        Dim TotalRows = ds0.Tables(0).Rows.Count
        Dim Rows1_50 = Convert.ToInt32(TotalRows * 0.02)
        Dim Rows2_50 = Convert.ToInt32(TotalRows * 0.04)
        Dim Rows3_50 = Convert.ToInt32(TotalRows * 0.06)
        Dim Rows4_50 = Convert.ToInt32(TotalRows * 0.08)
        Dim Rows5_50 = Convert.ToInt32(TotalRows * 0.1)
        Dim Rows6_50 = Convert.ToInt32(TotalRows * 0.12)
        Dim Rows7_50 = Convert.ToInt32(TotalRows * 0.14)
        Dim Rows8_50 = Convert.ToInt32(TotalRows * 0.16)
        Dim Rows9_50 = Convert.ToInt32(TotalRows * 0.18)
        Dim Rows10_50 = Convert.ToInt32(TotalRows * 0.2)
        Dim Rows11_50 = Convert.ToInt32(TotalRows * 0.22)
        Dim Rows12_50 = Convert.ToInt32(TotalRows * 0.24)
        Dim Rows13_50 = Convert.ToInt32(TotalRows * 0.26)
        Dim Rows14_50 = Convert.ToInt32(TotalRows * 0.28)
        Dim Rows15_50 = Convert.ToInt32(TotalRows * 0.3)
        Dim Rows16_50 = Convert.ToInt32(TotalRows * 0.32)
        Dim Rows17_50 = Convert.ToInt32(TotalRows * 0.34)
        Dim Rows18_50 = Convert.ToInt32(TotalRows * 0.36)
        Dim Rows19_50 = Convert.ToInt32(TotalRows * 0.38)
        Dim Rows20_50 = Convert.ToInt32(TotalRows * 0.4)
        Dim Rows21_50 = Convert.ToInt32(TotalRows * 0.42)
        Dim Rows22_50 = Convert.ToInt32(TotalRows * 0.44)
        Dim Rows23_50 = Convert.ToInt32(TotalRows * 0.46)
        Dim Rows24_50 = Convert.ToInt32(TotalRows * 0.48)
        Dim Rows25_50 = Convert.ToInt32(TotalRows * 0.5)
        Dim Rows26_50 = Convert.ToInt32(TotalRows * 0.52)
        Dim Rows27_50 = Convert.ToInt32(TotalRows * 0.54)
        Dim Rows28_50 = Convert.ToInt32(TotalRows * 0.56)
        Dim Rows29_50 = Convert.ToInt32(TotalRows * 0.58)
        Dim Rows30_50 = Convert.ToInt32(TotalRows * 0.6)
        Dim Rows31_50 = Convert.ToInt32(TotalRows * 0.62)
        Dim Rows32_50 = Convert.ToInt32(TotalRows * 0.64)
        Dim Rows33_50 = Convert.ToInt32(TotalRows * 0.66)
        Dim Rows34_50 = Convert.ToInt32(TotalRows * 0.68)
        Dim Rows35_50 = Convert.ToInt32(TotalRows * 0.7)
        Dim Rows36_50 = Convert.ToInt32(TotalRows * 0.72)
        Dim Rows37_50 = Convert.ToInt32(TotalRows * 0.74)
        Dim Rows38_50 = Convert.ToInt32(TotalRows * 0.76)
        Dim Rows39_50 = Convert.ToInt32(TotalRows * 0.78)
        Dim Rows40_50 = Convert.ToInt32(TotalRows * 0.8)
        Dim Rows41_50 = Convert.ToInt32(TotalRows * 0.82)
        Dim Rows42_50 = Convert.ToInt32(TotalRows * 0.84)
        Dim Rows43_50 = Convert.ToInt32(TotalRows * 0.86)
        Dim Rows44_50 = Convert.ToInt32(TotalRows * 0.88)
        Dim Rows45_50 = Convert.ToInt32(TotalRows * 0.9)
        Dim Rows46_50 = Convert.ToInt32(TotalRows * 0.92)
        Dim Rows47_50 = Convert.ToInt32(TotalRows * 0.94)
        Dim Rows48_50 = Convert.ToInt32(TotalRows * 0.96)
        Dim Rows49_50 = Convert.ToInt32(TotalRows * 0.98)
        Dim Rows50_50 = Convert.ToInt32(TotalRows * 1)
        Parallel.Invoke(
            Sub() Process_match_Key_set1(0, Rows1_50, ds0),
            Sub() Process_match_Key_set1(Rows1_50, Rows2_50, ds0),
            Sub() Process_match_Key_set1(Rows2_50, Rows3_50, ds0),
            Sub() Process_match_Key_set1(Rows3_50, Rows4_50, ds0),
            Sub() Process_match_Key_set1(Rows4_50, Rows5_50, ds0),
            Sub() Process_match_Key_set1(Rows5_50, Rows6_50, ds0),
            Sub() Process_match_Key_set1(Rows6_50, Rows7_50, ds0),
            Sub() Process_match_Key_set1(Rows7_50, Rows8_50, ds0),
            Sub() Process_match_Key_set1(Rows8_50, Rows9_50, ds0),
            Sub() Process_match_Key_set1(Rows9_50, Rows10_50, ds0),
            Sub() Process_match_Key_set1(Rows10_50, Rows11_50, ds0),
            Sub() Process_match_Key_set1(Rows11_50, Rows12_50, ds0),
            Sub() Process_match_Key_set1(Rows12_50, Rows13_50, ds0),
            Sub() Process_match_Key_set1(Rows13_50, Rows14_50, ds0),
            Sub() Process_match_Key_set1(Rows14_50, Rows15_50, ds0),
            Sub() Process_match_Key_set1(Rows15_50, Rows16_50, ds0),
            Sub() Process_match_Key_set1(Rows16_50, Rows17_50, ds0),
            Sub() Process_match_Key_set1(Rows17_50, Rows18_50, ds0),
            Sub() Process_match_Key_set1(Rows18_50, Rows19_50, ds0),
            Sub() Process_match_Key_set1(Rows19_50, Rows20_50, ds0),
            Sub() Process_match_Key_set1(Rows20_50, Rows21_50, ds0),
            Sub() Process_match_Key_set1(Rows21_50, Rows22_50, ds0),
            Sub() Process_match_Key_set1(Rows22_50, Rows23_50, ds0),
            Sub() Process_match_Key_set1(Rows23_50, Rows24_50, ds0),
            Sub() Process_match_Key_set1(Rows24_50, Rows25_50, ds0),
            Sub() Process_match_Key_set1(Rows25_50, Rows26_50, ds0),
            Sub() Process_match_Key_set1(Rows26_50, Rows27_50, ds0),
            Sub() Process_match_Key_set1(Rows27_50, Rows28_50, ds0),
            Sub() Process_match_Key_set1(Rows28_50, Rows29_50, ds0),
            Sub() Process_match_Key_set1(Rows29_50, Rows30_50, ds0),
            Sub() Process_match_Key_set1(Rows30_50, Rows31_50, ds0),
            Sub() Process_match_Key_set1(Rows31_50, Rows32_50, ds0),
            Sub() Process_match_Key_set1(Rows32_50, Rows33_50, ds0),
            Sub() Process_match_Key_set1(Rows33_50, Rows34_50, ds0),
            Sub() Process_match_Key_set1(Rows34_50, Rows35_50, ds0),
            Sub() Process_match_Key_set1(Rows35_50, Rows36_50, ds0),
            Sub() Process_match_Key_set1(Rows36_50, Rows37_50, ds0),
            Sub() Process_match_Key_set1(Rows37_50, Rows38_50, ds0),
            Sub() Process_match_Key_set1(Rows38_50, Rows39_50, ds0),
            Sub() Process_match_Key_set1(Rows39_50, Rows40_50, ds0),
            Sub() Process_match_Key_set1(Rows40_50, Rows41_50, ds0),
            Sub() Process_match_Key_set1(Rows41_50, Rows42_50, ds0),
            Sub() Process_match_Key_set1(Rows42_50, Rows43_50, ds0),
            Sub() Process_match_Key_set1(Rows43_50, Rows44_50, ds0),
            Sub() Process_match_Key_set1(Rows44_50, Rows45_50, ds0),
            Sub() Process_match_Key_set1(Rows45_50, Rows46_50, ds0),
            Sub() Process_match_Key_set1(Rows46_50, Rows47_50, ds0),
            Sub() Process_match_Key_set1(Rows47_50, Rows48_50, ds0),
            Sub() Process_match_Key_set1(Rows48_50, Rows49_50, ds0),
            Sub() Process_match_Key_set1(Rows49_50, Rows50_50, ds0)
            )
        Dim counter As Integer = 0

    End Sub

    Sub Process_match_Key_set1(ByVal InitialPosition As Integer, ByVal RowPosition As Integer, ByVal RowsData As DataSet)
        Console.WriteLine("Process_match_Key_set1 started at: " + DateTime.Now)
        Dim a8(8) As String
        Dim counter As Integer = 0
        For i As Integer = InitialPosition To RowPosition - 1
            a8(0) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key1(a8, c01, "01"),
                Sub() Match_Key1(a8, c02, "02"),
                Sub() Match_Key1(a8, c03, "03"),
                Sub() Match_Key1(a8, c04, "04"),
                Sub() Match_Key1(a8, c05, "05"),
                Sub() Match_Key1(a8, c06, "06"),
                Sub() Match_Key1(a8, c07, "07"),
                Sub() Match_Key1(a8, c08, "08"),
                Sub() Match_Key1(a8, c09, "09"),
                Sub() Match_Key1(a8, c10, "10"),
                Sub() Match_Key1(a8, c11, "11"),
                Sub() Match_Key1(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key1(a8, c01, "01"),
                Sub() Match_Key1(a8, c02, "02"),
                Sub() Match_Key1(a8, c03, "03"),
                Sub() Match_Key1(a8, c04, "04"),
                Sub() Match_Key1(a8, c05, "05"),
                Sub() Match_Key1(a8, c06, "06"),
                Sub() Match_Key1(a8, c07, "07"),
                Sub() Match_Key1(a8, c08, "08"),
                Sub() Match_Key1(a8, c09, "09"),
                Sub() Match_Key1(a8, c10, "10"),
                Sub() Match_Key1(a8, c11, "11"),
                Sub() Match_Key1(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key1(a8, c01, "01"),
                Sub() Match_Key1(a8, c02, "02"),
                Sub() Match_Key1(a8, c03, "03"),
                Sub() Match_Key1(a8, c04, "04"),
                Sub() Match_Key1(a8, c05, "05"),
                Sub() Match_Key1(a8, c06, "06"),
                Sub() Match_Key1(a8, c07, "07"),
                Sub() Match_Key1(a8, c08, "08"),
                Sub() Match_Key1(a8, c09, "09"),
                Sub() Match_Key1(a8, c10, "10"),
                Sub() Match_Key1(a8, c11, "11"),
                Sub() Match_Key1(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key1(a8, c01, "01"),
                Sub() Match_Key1(a8, c02, "02"),
                Sub() Match_Key1(a8, c03, "03"),
                Sub() Match_Key1(a8, c04, "04"),
                Sub() Match_Key1(a8, c05, "05"),
                Sub() Match_Key1(a8, c06, "06"),
                Sub() Match_Key1(a8, c07, "07"),
                Sub() Match_Key1(a8, c08, "08"),
                Sub() Match_Key1(a8, c09, "09"),
                Sub() Match_Key1(a8, c10, "10"),
                Sub() Match_Key1(a8, c11, "11"),
                Sub() Match_Key1(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key1(a8, c01, "01"),
                Sub() Match_Key1(a8, c02, "02"),
                Sub() Match_Key1(a8, c03, "03"),
                Sub() Match_Key1(a8, c04, "04"),
                Sub() Match_Key1(a8, c05, "05"),
                Sub() Match_Key1(a8, c06, "06"),
                Sub() Match_Key1(a8, c07, "07"),
                Sub() Match_Key1(a8, c08, "08"),
                Sub() Match_Key1(a8, c09, "09"),
                Sub() Match_Key1(a8, c10, "10"),
                Sub() Match_Key1(a8, c11, "11"),
                Sub() Match_Key1(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key1(a8, c01, "01"),
                Sub() Match_Key1(a8, c02, "02"),
                Sub() Match_Key1(a8, c03, "03"),
                Sub() Match_Key1(a8, c04, "04"),
                Sub() Match_Key1(a8, c05, "05"),
                Sub() Match_Key1(a8, c06, "06"),
                Sub() Match_Key1(a8, c07, "07"),
                Sub() Match_Key1(a8, c08, "08"),
                Sub() Match_Key1(a8, c09, "09"),
                Sub() Match_Key1(a8, c10, "10"),
                Sub() Match_Key1(a8, c11, "11"),
                Sub() Match_Key1(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key1(a8, c01, "01"),
                Sub() Match_Key1(a8, c02, "02"),
                Sub() Match_Key1(a8, c03, "03"),
                Sub() Match_Key1(a8, c04, "04"),
                Sub() Match_Key1(a8, c05, "05"),
                Sub() Match_Key1(a8, c06, "06"),
                Sub() Match_Key1(a8, c07, "07"),
                Sub() Match_Key1(a8, c08, "08"),
                Sub() Match_Key1(a8, c09, "09"),
                Sub() Match_Key1(a8, c10, "10"),
                Sub() Match_Key1(a8, c11, "11"),
                Sub() Match_Key1(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key1(a8, c01, "01"),
                Sub() Match_Key1(a8, c02, "02"),
                Sub() Match_Key1(a8, c03, "03"),
                Sub() Match_Key1(a8, c04, "04"),
                Sub() Match_Key1(a8, c05, "05"),
                Sub() Match_Key1(a8, c06, "06"),
                Sub() Match_Key1(a8, c07, "07"),
                Sub() Match_Key1(a8, c08, "08"),
                Sub() Match_Key1(a8, c09, "09"),
                Sub() Match_Key1(a8, c10, "10"),
                Sub() Match_Key1(a8, c11, "11"),
                Sub() Match_Key1(a8, c12, "12")
            )
        Next
        Console.WriteLine("Process_match_Key_set1 ended- at: " + DateTime.Now)
    End Sub
    Sub Process_match_Key_set2(ByVal InitialPosition As Integer, ByVal RowPosition As Integer, ByVal RowsData As DataSet)
        Console.WriteLine("Process_match_Key_set2 started at: " + DateTime.Now)
        Dim a8(8) As String
        Dim counter As Integer = 0
        For i As Integer = InitialPosition To RowPosition - 1
            a8(0) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key2(a8, c01, "01"),
                Sub() Match_Key2(a8, c02, "02"),
                Sub() Match_Key2(a8, c03, "03"),
                Sub() Match_Key2(a8, c04, "04"),
                Sub() Match_Key2(a8, c05, "05"),
                Sub() Match_Key2(a8, c06, "06"),
                Sub() Match_Key2(a8, c07, "07"),
                Sub() Match_Key2(a8, c08, "08"),
                Sub() Match_Key2(a8, c09, "09"),
                Sub() Match_Key2(a8, c10, "10"),
                Sub() Match_Key2(a8, c11, "11"),
                Sub() Match_Key2(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key2(a8, c01, "01"),
                Sub() Match_Key2(a8, c02, "02"),
                Sub() Match_Key2(a8, c03, "03"),
                Sub() Match_Key2(a8, c04, "04"),
                Sub() Match_Key2(a8, c05, "05"),
                Sub() Match_Key2(a8, c06, "06"),
                Sub() Match_Key2(a8, c07, "07"),
                Sub() Match_Key2(a8, c08, "08"),
                Sub() Match_Key2(a8, c09, "09"),
                Sub() Match_Key2(a8, c10, "10"),
                Sub() Match_Key2(a8, c11, "11"),
                Sub() Match_Key2(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key2(a8, c01, "01"),
                Sub() Match_Key2(a8, c02, "02"),
                Sub() Match_Key2(a8, c03, "03"),
                Sub() Match_Key2(a8, c04, "04"),
                Sub() Match_Key2(a8, c05, "05"),
                Sub() Match_Key2(a8, c06, "06"),
                Sub() Match_Key2(a8, c07, "07"),
                Sub() Match_Key2(a8, c08, "08"),
                Sub() Match_Key2(a8, c09, "09"),
                Sub() Match_Key2(a8, c10, "10"),
                Sub() Match_Key2(a8, c11, "11"),
                Sub() Match_Key2(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key2(a8, c01, "01"),
                Sub() Match_Key2(a8, c02, "02"),
                Sub() Match_Key2(a8, c03, "03"),
                Sub() Match_Key2(a8, c04, "04"),
                Sub() Match_Key2(a8, c05, "05"),
                Sub() Match_Key2(a8, c06, "06"),
                Sub() Match_Key2(a8, c07, "07"),
                Sub() Match_Key2(a8, c08, "08"),
                Sub() Match_Key2(a8, c09, "09"),
                Sub() Match_Key2(a8, c10, "10"),
                Sub() Match_Key2(a8, c11, "11"),
                Sub() Match_Key2(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key2(a8, c01, "01"),
                Sub() Match_Key2(a8, c02, "02"),
                Sub() Match_Key2(a8, c03, "03"),
                Sub() Match_Key2(a8, c04, "04"),
                Sub() Match_Key2(a8, c05, "05"),
                Sub() Match_Key2(a8, c06, "06"),
                Sub() Match_Key2(a8, c07, "07"),
                Sub() Match_Key2(a8, c08, "08"),
                Sub() Match_Key2(a8, c09, "09"),
                Sub() Match_Key2(a8, c10, "10"),
                Sub() Match_Key2(a8, c11, "11"),
                Sub() Match_Key2(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key2(a8, c01, "01"),
                Sub() Match_Key2(a8, c02, "02"),
                Sub() Match_Key2(a8, c03, "03"),
                Sub() Match_Key2(a8, c04, "04"),
                Sub() Match_Key2(a8, c05, "05"),
                Sub() Match_Key2(a8, c06, "06"),
                Sub() Match_Key2(a8, c07, "07"),
                Sub() Match_Key2(a8, c08, "08"),
                Sub() Match_Key2(a8, c09, "09"),
                Sub() Match_Key2(a8, c10, "10"),
                Sub() Match_Key2(a8, c11, "11"),
                Sub() Match_Key2(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key2(a8, c01, "01"),
                Sub() Match_Key2(a8, c02, "02"),
                Sub() Match_Key2(a8, c03, "03"),
                Sub() Match_Key2(a8, c04, "04"),
                Sub() Match_Key2(a8, c05, "05"),
                Sub() Match_Key2(a8, c06, "06"),
                Sub() Match_Key2(a8, c07, "07"),
                Sub() Match_Key2(a8, c08, "08"),
                Sub() Match_Key2(a8, c09, "09"),
                Sub() Match_Key2(a8, c10, "10"),
                Sub() Match_Key2(a8, c11, "11"),
                Sub() Match_Key2(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key2(a8, c01, "01"),
                Sub() Match_Key2(a8, c02, "02"),
                Sub() Match_Key2(a8, c03, "03"),
                Sub() Match_Key2(a8, c04, "04"),
                Sub() Match_Key2(a8, c05, "05"),
                Sub() Match_Key2(a8, c06, "06"),
                Sub() Match_Key2(a8, c07, "07"),
                Sub() Match_Key2(a8, c08, "08"),
                Sub() Match_Key2(a8, c09, "09"),
                Sub() Match_Key2(a8, c10, "10"),
                Sub() Match_Key2(a8, c11, "11"),
                Sub() Match_Key2(a8, c12, "12")
            )
        Next
        Console.WriteLine("Process_match_Key_set2 ended- at: " + DateTime.Now)
    End Sub
    Sub Process_match_Key_set3(ByVal InitialPosition As Integer, ByVal RowPosition As Integer, ByVal RowsData As DataSet)
        Console.WriteLine("Process_match_Key_set3 started at: " + DateTime.Now)
        Dim a8(8) As String
        Dim counter As Integer = 0
        For i As Integer = InitialPosition To RowPosition - 1
            a8(0) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key3(a8, c01, "01"),
                Sub() Match_Key3(a8, c02, "02"),
                Sub() Match_Key3(a8, c03, "03"),
                Sub() Match_Key3(a8, c04, "04"),
                Sub() Match_Key3(a8, c05, "05"),
                Sub() Match_Key3(a8, c06, "06"),
                Sub() Match_Key3(a8, c07, "07"),
                Sub() Match_Key3(a8, c08, "08"),
                Sub() Match_Key3(a8, c09, "09"),
                Sub() Match_Key3(a8, c10, "10"),
                Sub() Match_Key3(a8, c11, "11"),
                Sub() Match_Key3(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key3(a8, c01, "01"),
                Sub() Match_Key3(a8, c02, "02"),
                Sub() Match_Key3(a8, c03, "03"),
                Sub() Match_Key3(a8, c04, "04"),
                Sub() Match_Key3(a8, c05, "05"),
                Sub() Match_Key3(a8, c06, "06"),
                Sub() Match_Key3(a8, c07, "07"),
                Sub() Match_Key3(a8, c08, "08"),
                Sub() Match_Key3(a8, c09, "09"),
                Sub() Match_Key3(a8, c10, "10"),
                Sub() Match_Key3(a8, c11, "11"),
                Sub() Match_Key3(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key3(a8, c01, "01"),
                Sub() Match_Key3(a8, c02, "02"),
                Sub() Match_Key3(a8, c03, "03"),
                Sub() Match_Key3(a8, c04, "04"),
                Sub() Match_Key3(a8, c05, "05"),
                Sub() Match_Key3(a8, c06, "06"),
                Sub() Match_Key3(a8, c07, "07"),
                Sub() Match_Key3(a8, c08, "08"),
                Sub() Match_Key3(a8, c09, "09"),
                Sub() Match_Key3(a8, c10, "10"),
                Sub() Match_Key3(a8, c11, "11"),
                Sub() Match_Key3(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key3(a8, c01, "01"),
                Sub() Match_Key3(a8, c02, "02"),
                Sub() Match_Key3(a8, c03, "03"),
                Sub() Match_Key3(a8, c04, "04"),
                Sub() Match_Key3(a8, c05, "05"),
                Sub() Match_Key3(a8, c06, "06"),
                Sub() Match_Key3(a8, c07, "07"),
                Sub() Match_Key3(a8, c08, "08"),
                Sub() Match_Key3(a8, c09, "09"),
                Sub() Match_Key3(a8, c10, "10"),
                Sub() Match_Key3(a8, c11, "11"),
                Sub() Match_Key3(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key3(a8, c01, "01"),
                Sub() Match_Key3(a8, c02, "02"),
                Sub() Match_Key3(a8, c03, "03"),
                Sub() Match_Key3(a8, c04, "04"),
                Sub() Match_Key3(a8, c05, "05"),
                Sub() Match_Key3(a8, c06, "06"),
                Sub() Match_Key3(a8, c07, "07"),
                Sub() Match_Key3(a8, c08, "08"),
                Sub() Match_Key3(a8, c09, "09"),
                Sub() Match_Key3(a8, c10, "10"),
                Sub() Match_Key3(a8, c11, "11"),
                Sub() Match_Key3(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key3(a8, c01, "01"),
                Sub() Match_Key3(a8, c02, "02"),
                Sub() Match_Key3(a8, c03, "03"),
                Sub() Match_Key3(a8, c04, "04"),
                Sub() Match_Key3(a8, c05, "05"),
                Sub() Match_Key3(a8, c06, "06"),
                Sub() Match_Key3(a8, c07, "07"),
                Sub() Match_Key3(a8, c08, "08"),
                Sub() Match_Key3(a8, c09, "09"),
                Sub() Match_Key3(a8, c10, "10"),
                Sub() Match_Key3(a8, c11, "11"),
                Sub() Match_Key3(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key3(a8, c01, "01"),
                Sub() Match_Key3(a8, c02, "02"),
                Sub() Match_Key3(a8, c03, "03"),
                Sub() Match_Key3(a8, c04, "04"),
                Sub() Match_Key3(a8, c05, "05"),
                Sub() Match_Key3(a8, c06, "06"),
                Sub() Match_Key3(a8, c07, "07"),
                Sub() Match_Key3(a8, c08, "08"),
                Sub() Match_Key3(a8, c09, "09"),
                Sub() Match_Key3(a8, c10, "10"),
                Sub() Match_Key3(a8, c11, "11"),
                Sub() Match_Key3(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key3(a8, c01, "01"),
                Sub() Match_Key3(a8, c02, "02"),
                Sub() Match_Key3(a8, c03, "03"),
                Sub() Match_Key3(a8, c04, "04"),
                Sub() Match_Key3(a8, c05, "05"),
                Sub() Match_Key3(a8, c06, "06"),
                Sub() Match_Key3(a8, c07, "07"),
                Sub() Match_Key3(a8, c08, "08"),
                Sub() Match_Key3(a8, c09, "09"),
                Sub() Match_Key3(a8, c10, "10"),
                Sub() Match_Key3(a8, c11, "11"),
                Sub() Match_Key3(a8, c12, "12")
            )
        Next
        Console.WriteLine("Process_match_Key_set3 ended- at: " + DateTime.Now)
    End Sub
    Sub Process_match_Key_set4(ByVal InitialPosition As Integer, ByVal RowPosition As Integer, ByVal RowsData As DataSet)
        Console.WriteLine("Process_match_Key_set4 started at: " + DateTime.Now)
        Dim a8(8) As String
        Dim counter As Integer = 0
        For i As Integer = InitialPosition To RowPosition - 1
            a8(0) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key4(a8, c01, "01"),
                Sub() Match_Key4(a8, c02, "02"),
                Sub() Match_Key4(a8, c03, "03"),
                Sub() Match_Key4(a8, c04, "04"),
                Sub() Match_Key4(a8, c05, "05"),
                Sub() Match_Key4(a8, c06, "06"),
                Sub() Match_Key4(a8, c07, "07"),
                Sub() Match_Key4(a8, c08, "08"),
                Sub() Match_Key4(a8, c09, "09"),
                Sub() Match_Key4(a8, c10, "10"),
                Sub() Match_Key4(a8, c11, "11"),
                Sub() Match_Key4(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key4(a8, c01, "01"),
                Sub() Match_Key4(a8, c02, "02"),
                Sub() Match_Key4(a8, c03, "03"),
                Sub() Match_Key4(a8, c04, "04"),
                Sub() Match_Key4(a8, c05, "05"),
                Sub() Match_Key4(a8, c06, "06"),
                Sub() Match_Key4(a8, c07, "07"),
                Sub() Match_Key4(a8, c08, "08"),
                Sub() Match_Key4(a8, c09, "09"),
                Sub() Match_Key4(a8, c10, "10"),
                Sub() Match_Key4(a8, c11, "11"),
                Sub() Match_Key4(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key4(a8, c01, "01"),
                Sub() Match_Key4(a8, c02, "02"),
                Sub() Match_Key4(a8, c03, "03"),
                Sub() Match_Key4(a8, c04, "04"),
                Sub() Match_Key4(a8, c05, "05"),
                Sub() Match_Key4(a8, c06, "06"),
                Sub() Match_Key4(a8, c07, "07"),
                Sub() Match_Key4(a8, c08, "08"),
                Sub() Match_Key4(a8, c09, "09"),
                Sub() Match_Key4(a8, c10, "10"),
                Sub() Match_Key4(a8, c11, "11"),
                Sub() Match_Key4(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key4(a8, c01, "01"),
                Sub() Match_Key4(a8, c02, "02"),
                Sub() Match_Key4(a8, c03, "03"),
                Sub() Match_Key4(a8, c04, "04"),
                Sub() Match_Key4(a8, c05, "05"),
                Sub() Match_Key4(a8, c06, "06"),
                Sub() Match_Key4(a8, c07, "07"),
                Sub() Match_Key4(a8, c08, "08"),
                Sub() Match_Key4(a8, c09, "09"),
                Sub() Match_Key4(a8, c10, "10"),
                Sub() Match_Key4(a8, c11, "11"),
                Sub() Match_Key4(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()

            Parallel.Invoke(
                Sub() Match_Key4(a8, c01, "01"),
                Sub() Match_Key4(a8, c02, "02"),
                Sub() Match_Key4(a8, c03, "03"),
                Sub() Match_Key4(a8, c04, "04"),
                Sub() Match_Key4(a8, c05, "05"),
                Sub() Match_Key4(a8, c06, "06"),
                Sub() Match_Key4(a8, c07, "07"),
                Sub() Match_Key4(a8, c08, "08"),
                Sub() Match_Key4(a8, c09, "09"),
                Sub() Match_Key4(a8, c10, "10"),
                Sub() Match_Key4(a8, c11, "11"),
                Sub() Match_Key4(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key4(a8, c01, "01"),
                Sub() Match_Key4(a8, c02, "02"),
                Sub() Match_Key4(a8, c03, "03"),
                Sub() Match_Key4(a8, c04, "04"),
                Sub() Match_Key4(a8, c05, "05"),
                Sub() Match_Key4(a8, c06, "06"),
                Sub() Match_Key4(a8, c07, "07"),
                Sub() Match_Key4(a8, c08, "08"),
                Sub() Match_Key4(a8, c09, "09"),
                Sub() Match_Key4(a8, c10, "10"),
                Sub() Match_Key4(a8, c11, "11"),
                Sub() Match_Key4(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key4(a8, c01, "01"),
                Sub() Match_Key4(a8, c02, "02"),
                Sub() Match_Key4(a8, c03, "03"),
                Sub() Match_Key4(a8, c04, "04"),
                Sub() Match_Key4(a8, c05, "05"),
                Sub() Match_Key4(a8, c06, "06"),
                Sub() Match_Key4(a8, c07, "07"),
                Sub() Match_Key4(a8, c08, "08"),
                Sub() Match_Key4(a8, c09, "09"),
                Sub() Match_Key4(a8, c10, "10"),
                Sub() Match_Key4(a8, c11, "11"),
                Sub() Match_Key4(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key4(a8, c01, "01"),
                Sub() Match_Key4(a8, c02, "02"),
                Sub() Match_Key4(a8, c03, "03"),
                Sub() Match_Key4(a8, c04, "04"),
                Sub() Match_Key4(a8, c05, "05"),
                Sub() Match_Key4(a8, c06, "06"),
                Sub() Match_Key4(a8, c07, "07"),
                Sub() Match_Key4(a8, c08, "08"),
                Sub() Match_Key4(a8, c09, "09"),
                Sub() Match_Key4(a8, c10, "10"),
                Sub() Match_Key4(a8, c11, "11"),
                Sub() Match_Key4(a8, c12, "12")
            )
        Next
        Console.WriteLine("Process_match_Key_set4 ended- at: " + DateTime.Now)
    End Sub
    Sub Process_match_Key_set5(ByVal InitialPosition As Integer, ByVal RowPosition As Integer, ByVal RowsData As DataSet)
        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter("c:\test.txt", True)
        Console.WriteLine("Process_match_Key_set5 started at: " + DateTime.Now)
        Dim a8(8) As String
        Dim counter As Integer = 0
        For i As Integer = InitialPosition To RowPosition - 1
            a8(0) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key5(a8, c01, "01"),
                Sub() Match_Key5(a8, c02, "02"),
                Sub() Match_Key5(a8, c03, "03"),
                Sub() Match_Key5(a8, c04, "04"),
                Sub() Match_Key5(a8, c05, "05"),
                Sub() Match_Key5(a8, c06, "06"),
                Sub() Match_Key5(a8, c07, "07"),
                Sub() Match_Key5(a8, c08, "08"),
                Sub() Match_Key5(a8, c09, "09"),
                Sub() Match_Key5(a8, c10, "10"),
                Sub() Match_Key5(a8, c11, "11"),
                Sub() Match_Key5(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key5(a8, c01, "01"),
                Sub() Match_Key5(a8, c02, "02"),
                Sub() Match_Key5(a8, c03, "03"),
                Sub() Match_Key5(a8, c04, "04"),
                Sub() Match_Key5(a8, c05, "05"),
                Sub() Match_Key5(a8, c06, "06"),
                Sub() Match_Key5(a8, c07, "07"),
                Sub() Match_Key5(a8, c08, "08"),
                Sub() Match_Key5(a8, c09, "09"),
                Sub() Match_Key5(a8, c10, "10"),
                Sub() Match_Key5(a8, c11, "11"),
                Sub() Match_Key5(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key5(a8, c01, "01"),
                Sub() Match_Key5(a8, c02, "02"),
                Sub() Match_Key5(a8, c03, "03"),
                Sub() Match_Key5(a8, c04, "04"),
                Sub() Match_Key5(a8, c05, "05"),
                Sub() Match_Key5(a8, c06, "06"),
                Sub() Match_Key5(a8, c07, "07"),
                Sub() Match_Key5(a8, c08, "08"),
                Sub() Match_Key5(a8, c09, "09"),
                Sub() Match_Key5(a8, c10, "10"),
                Sub() Match_Key5(a8, c11, "11"),
                Sub() Match_Key5(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key5(a8, c01, "01"),
                Sub() Match_Key5(a8, c02, "02"),
                Sub() Match_Key5(a8, c03, "03"),
                Sub() Match_Key5(a8, c04, "04"),
                Sub() Match_Key5(a8, c05, "05"),
                Sub() Match_Key5(a8, c06, "06"),
                Sub() Match_Key5(a8, c07, "07"),
                Sub() Match_Key5(a8, c08, "08"),
                Sub() Match_Key5(a8, c09, "09"),
                Sub() Match_Key5(a8, c10, "10"),
                Sub() Match_Key5(a8, c11, "11"),
                Sub() Match_Key5(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key5(a8, c01, "01"),
                Sub() Match_Key5(a8, c02, "02"),
                Sub() Match_Key5(a8, c03, "03"),
                Sub() Match_Key5(a8, c04, "04"),
                Sub() Match_Key5(a8, c05, "05"),
                Sub() Match_Key5(a8, c06, "06"),
                Sub() Match_Key5(a8, c07, "07"),
                Sub() Match_Key5(a8, c08, "08"),
                Sub() Match_Key5(a8, c09, "09"),
                Sub() Match_Key5(a8, c10, "10"),
                Sub() Match_Key5(a8, c11, "11"),
                Sub() Match_Key5(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key5(a8, c01, "01"),
                Sub() Match_Key5(a8, c02, "02"),
                Sub() Match_Key5(a8, c03, "03"),
                Sub() Match_Key5(a8, c04, "04"),
                Sub() Match_Key5(a8, c05, "05"),
                Sub() Match_Key5(a8, c06, "06"),
                Sub() Match_Key5(a8, c07, "07"),
                Sub() Match_Key5(a8, c08, "08"),
                Sub() Match_Key5(a8, c09, "09"),
                Sub() Match_Key5(a8, c10, "10"),
                Sub() Match_Key5(a8, c11, "11"),
                Sub() Match_Key5(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key5(a8, c01, "01"),
                Sub() Match_Key5(a8, c02, "02"),
                Sub() Match_Key5(a8, c03, "03"),
                Sub() Match_Key5(a8, c04, "04"),
                Sub() Match_Key5(a8, c05, "05"),
                Sub() Match_Key5(a8, c06, "06"),
                Sub() Match_Key5(a8, c07, "07"),
                Sub() Match_Key5(a8, c08, "08"),
                Sub() Match_Key5(a8, c09, "09"),
                Sub() Match_Key5(a8, c10, "10"),
                Sub() Match_Key5(a8, c11, "11"),
                Sub() Match_Key5(a8, c12, "12")
            )
            a8(0) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
            a8(1) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
            a8(2) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
            a8(3) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
            a8(4) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
            a8(5) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
            a8(6) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
            a8(7) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
            Parallel.Invoke(
                Sub() Match_Key5(a8, c01, "01"),
                Sub() Match_Key5(a8, c02, "02"),
                Sub() Match_Key5(a8, c03, "03"),
                Sub() Match_Key5(a8, c04, "04"),
                Sub() Match_Key5(a8, c05, "05"),
                Sub() Match_Key5(a8, c06, "06"),
                Sub() Match_Key5(a8, c07, "07"),
                Sub() Match_Key5(a8, c08, "08"),
                Sub() Match_Key5(a8, c09, "09"),
                Sub() Match_Key5(a8, c10, "10"),
                Sub() Match_Key5(a8, c11, "11"),
                Sub() Match_Key5(a8, c12, "12")
            )
        Next
        Console.WriteLine("Process_match_Key_set5 ended- at: " + DateTime.Now)
        file.Close()
    End Sub
    Sub Match_Key1(ByVal combo() As String, ByVal Cuspp() As String, ByVal cloc As String)
        Dim fl1 As String = "N"
        Dim fl2 As String = "N"
        Dim fl3 As String = "N"
        Dim fl4 As String = "N"
        Dim fl5 As String = "N"
        Dim fl6 As String = "N"
        Dim fl7 As String = "N"
        Dim fl8 As String = "N"
        Dim m_str As String = ""
        Dim xcusp(Cuspp.Length) As String
        Dim xcombo(combo.Length) As String

        Dim pstr1 As String = ""
        For i As Integer = 0 To combo.Length - 1
            pstr1 = pstr1 + combo(i)
        Next

        Dim pstr2 As String = ""
        For i As Integer = 0 To Cuspp.Length - 1
            pstr2 = pstr2 + Cuspp(i)
        Next

        'Check first position
        Array.Copy(Cuspp, xcusp, Cuspp.Length)
        Array.Copy(combo, xcombo, Cuspp.Length)

        For i As Integer = 0 To xcusp.Length - 2
            If xcombo(0) = xcusp(i) Then
                m_str = m_str + xcombo(0)
                xcusp(i) = "XX"
                fl1 = "Y"
                m_str = m_str
                Exit For
            End If
            If xcombo(0) = Cuspp(i) Then
                m_str = m_str + xcombo(0)
                fl1 = "Y"
            End If
        Next

        If fl1 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(1) = xcusp(i) Then
                    m_str = m_str + xcombo(1)
                    xcusp(i) = "XX"
                    fl2 = "Y"
                    Exit For
                End If
                If xcombo(1) = Cuspp(i) Then
                    m_str = m_str + xcombo(1)
                    fl2 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(2) = xcusp(i) Then
                    m_str = m_str + xcombo(2)
                    xcusp(i) = "XX"
                    fl3 = "Y"
                    Exit For
                End If
                If xcombo(2) = Cuspp(i) Then
                    m_str = m_str + xcombo(2)
                    fl3 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(3) = xcusp(i) Then
                    m_str = m_str + xcombo(3)
                    xcusp(i) = "XX"
                    fl4 = "Y"
                    Exit For
                End If
                If xcombo(3) = Cuspp(i) Then
                    m_str = m_str + xcombo(3)
                    fl4 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" And fl4 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(4) = xcusp(i) Then
                    m_str = m_str + xcombo(4)
                    xcusp(i) = "XX"
                    fl5 = "Y"
                    Exit For
                End If
                If xcombo(4) = Cuspp(i) Then
                    m_str = m_str + xcombo(4)
                    fl5 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" And fl4 = "Y" And fl5 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(5) = xcusp(i) Then
                    m_str = m_str + xcombo(5)
                    xcusp(i) = "XX"
                    fl6 = "Y"
                    Exit For
                End If
                If xcombo(5) = Cuspp(i) Then
                    m_str = m_str + xcombo(5)
                    fl6 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" And fl4 = "Y" And fl5 = "Y" And fl6 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(6) = xcusp(i) Then
                    m_str = m_str + xcombo(6)
                    xcusp(i) = "XX"
                    fl7 = "Y"
                    Exit For
                End If
                If xcombo(6) = Cuspp(i) Then
                    m_str = m_str + xcombo(6)
                    fl7 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" And fl4 = "Y" And fl5 = "Y" And fl6 = "Y" And fl7 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(7) = xcusp(i) Then
                    m_str = m_str + xcombo(7)
                    xcusp(i) = "XX"
                    fl8 = "Y"
                    Exit For
                End If
                If xcombo(7) = Cuspp(i) Then
                    m_str = m_str + xcombo(7)
                    fl8 = "Y"
                End If
            Next
        End If

        Dim pattern As Boolean = True
        For i As Integer = 0 To xcusp.Length - 2
            If xcusp(i) IsNot "XX" Then
                pattern = False
            End If
        Next

        If pattern = True Then
            Dim con As New SqlConnection
            Dim cmd As New SqlCommand
            Dim uid As String = "XXXXXXXXXX"
            Dim hid As String = "100001"
            Try
                con.ConnectionString = "data source=WIN-KSTUPT6CJRC;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=True;multipleactiveresultsets=True;"
                con.Open()
                cmd.Connection = con
                cmd.CommandText = "INSERT INTO MATCHFILE VALUES ('" + uid + "','" + hid + "','" + m_planet + "','" + m_key + "','" + pstr1 + "','" + cloc + "','" + pstr2 + "','" + m_str + "');"
                cmd.ExecuteNonQuery()

            Catch ex As Exception
                Console.WriteLine("Error while inserting record on table..." & ex.Message, "Insert Records")
            Finally
                con.Close()
            End Try
        End If
    End Sub
    Sub Match_Key2(ByVal combo() As String, ByVal Cuspp() As String, ByVal cloc As String)
        Dim fl1 As String = "N"
        Dim fl2 As String = "N"
        Dim fl3 As String = "N"
        Dim fl4 As String = "N"
        Dim fl5 As String = "N"
        Dim fl6 As String = "N"
        Dim fl7 As String = "N"
        Dim fl8 As String = "N"
        Dim m_str As String = ""
        Dim xcusp(Cuspp.Length) As String
        Dim xcombo(combo.Length) As String

        Dim pstr1 As String = ""
        For i As Integer = 0 To combo.Length - 1
            pstr1 = pstr1 + combo(i)
        Next

        Dim pstr2 As String = ""
        For i As Integer = 0 To Cuspp.Length - 1
            pstr2 = pstr2 + Cuspp(i)
        Next

        'Check first position
        Array.Copy(Cuspp, xcusp, Cuspp.Length)
        Array.Copy(combo, xcombo, Cuspp.Length)

        For i As Integer = 0 To xcusp.Length - 2
            If xcombo(0) = xcusp(i) Then
                m_str = m_str + xcombo(0)
                xcusp(i) = "XX"
                fl1 = "Y"
                m_str = m_str
                Exit For
            End If
            If xcombo(0) = Cuspp(i) Then
                m_str = m_str + xcombo(0)
                fl1 = "Y"
            End If
        Next

        If fl1 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(1) = xcusp(i) Then
                    m_str = m_str + xcombo(1)
                    xcusp(i) = "XX"
                    fl2 = "Y"
                    Exit For
                End If
                If xcombo(1) = Cuspp(i) Then
                    m_str = m_str + xcombo(1)
                    fl2 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(2) = xcusp(i) Then
                    m_str = m_str + xcombo(2)
                    xcusp(i) = "XX"
                    fl3 = "Y"
                    Exit For
                End If
                If xcombo(2) = Cuspp(i) Then
                    m_str = m_str + xcombo(2)
                    fl3 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(3) = xcusp(i) Then
                    m_str = m_str + xcombo(3)
                    xcusp(i) = "XX"
                    fl4 = "Y"
                    Exit For
                End If
                If xcombo(3) = Cuspp(i) Then
                    m_str = m_str + xcombo(3)
                    fl4 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" And fl4 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(4) = xcusp(i) Then
                    m_str = m_str + xcombo(4)
                    xcusp(i) = "XX"
                    fl5 = "Y"
                    Exit For
                End If
                If xcombo(4) = Cuspp(i) Then
                    m_str = m_str + xcombo(4)
                    fl5 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" And fl4 = "Y" And fl5 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(5) = xcusp(i) Then
                    m_str = m_str + xcombo(5)
                    xcusp(i) = "XX"
                    fl6 = "Y"
                    Exit For
                End If
                If xcombo(5) = Cuspp(i) Then
                    m_str = m_str + xcombo(5)
                    fl6 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" And fl4 = "Y" And fl5 = "Y" And fl6 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(6) = xcusp(i) Then
                    m_str = m_str + xcombo(6)
                    xcusp(i) = "XX"
                    fl7 = "Y"
                    Exit For
                End If
                If xcombo(6) = Cuspp(i) Then
                    m_str = m_str + xcombo(6)
                    fl7 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" And fl4 = "Y" And fl5 = "Y" And fl6 = "Y" And fl7 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(7) = xcusp(i) Then
                    m_str = m_str + xcombo(7)
                    xcusp(i) = "XX"
                    fl8 = "Y"
                    Exit For
                End If
                If xcombo(7) = Cuspp(i) Then
                    m_str = m_str + xcombo(7)
                    fl8 = "Y"
                End If
            Next
        End If

        Dim pattern As Boolean = True
        For i As Integer = 0 To xcusp.Length - 2
            If xcusp(i) IsNot "XX" Then
                pattern = False
            End If
        Next

        If pattern = True Then
            Dim con As New SqlConnection
            Dim cmd As New SqlCommand
            Dim uid As String = "XXXXXXXXXX"
            Dim hid As String = "100001"
            Try
                con.ConnectionString = "data source=WIN-KSTUPT6CJRC;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=True;multipleactiveresultsets=True;"
                con.Open()
                cmd.Connection = con
                cmd.CommandText = "INSERT INTO MATCHFILE VALUES ('" + uid + "','" + hid + "','" + m_planet + "','" + m_key + "','" + pstr1 + "','" + cloc + "','" + pstr2 + "','" + m_str + "');"
                cmd.ExecuteNonQuery()

            Catch ex As Exception
                Console.WriteLine("Error while inserting record on table..." & ex.Message, "Insert Records")
            Finally
                con.Close()
            End Try
        End If
    End Sub
    Sub Match_Key3(ByVal combo() As String, ByVal Cuspp() As String, ByVal cloc As String)
        Dim fl1 As String = "N"
        Dim fl2 As String = "N"
        Dim fl3 As String = "N"
        Dim fl4 As String = "N"
        Dim fl5 As String = "N"
        Dim fl6 As String = "N"
        Dim fl7 As String = "N"
        Dim fl8 As String = "N"
        Dim m_str As String = ""
        Dim xcusp(Cuspp.Length) As String
        Dim xcombo(combo.Length) As String

        Dim pstr1 As String = ""
        For i As Integer = 0 To combo.Length - 1
            pstr1 = pstr1 + combo(i)
        Next

        Dim pstr2 As String = ""
        For i As Integer = 0 To Cuspp.Length - 1
            pstr2 = pstr2 + Cuspp(i)
        Next

        'Check first position
        Array.Copy(Cuspp, xcusp, Cuspp.Length)
        Array.Copy(combo, xcombo, Cuspp.Length)

        For i As Integer = 0 To xcusp.Length - 2
            If xcombo(0) = xcusp(i) Then
                m_str = m_str + xcombo(0)
                xcusp(i) = "XX"
                fl1 = "Y"
                m_str = m_str
                Exit For
            End If
            If xcombo(0) = Cuspp(i) Then
                m_str = m_str + xcombo(0)
                fl1 = "Y"
            End If
        Next

        If fl1 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(1) = xcusp(i) Then
                    m_str = m_str + xcombo(1)
                    xcusp(i) = "XX"
                    fl2 = "Y"
                    Exit For
                End If
                If xcombo(1) = Cuspp(i) Then
                    m_str = m_str + xcombo(1)
                    fl2 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(2) = xcusp(i) Then
                    m_str = m_str + xcombo(2)
                    xcusp(i) = "XX"
                    fl3 = "Y"
                    Exit For
                End If
                If xcombo(2) = Cuspp(i) Then
                    m_str = m_str + xcombo(2)
                    fl3 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(3) = xcusp(i) Then
                    m_str = m_str + xcombo(3)
                    xcusp(i) = "XX"
                    fl4 = "Y"
                    Exit For
                End If
                If xcombo(3) = Cuspp(i) Then
                    m_str = m_str + xcombo(3)
                    fl4 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" And fl4 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(4) = xcusp(i) Then
                    m_str = m_str + xcombo(4)
                    xcusp(i) = "XX"
                    fl5 = "Y"
                    Exit For
                End If
                If xcombo(4) = Cuspp(i) Then
                    m_str = m_str + xcombo(4)
                    fl5 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" And fl4 = "Y" And fl5 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(5) = xcusp(i) Then
                    m_str = m_str + xcombo(5)
                    xcusp(i) = "XX"
                    fl6 = "Y"
                    Exit For
                End If
                If xcombo(5) = Cuspp(i) Then
                    m_str = m_str + xcombo(5)
                    fl6 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" And fl4 = "Y" And fl5 = "Y" And fl6 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(6) = xcusp(i) Then
                    m_str = m_str + xcombo(6)
                    xcusp(i) = "XX"
                    fl7 = "Y"
                    Exit For
                End If
                If xcombo(6) = Cuspp(i) Then
                    m_str = m_str + xcombo(6)
                    fl7 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" And fl4 = "Y" And fl5 = "Y" And fl6 = "Y" And fl7 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(7) = xcusp(i) Then
                    m_str = m_str + xcombo(7)
                    xcusp(i) = "XX"
                    fl8 = "Y"
                    Exit For
                End If
                If xcombo(7) = Cuspp(i) Then
                    m_str = m_str + xcombo(7)
                    fl8 = "Y"
                End If
            Next
        End If

        Dim pattern As Boolean = True
        For i As Integer = 0 To xcusp.Length - 2
            If xcusp(i) IsNot "XX" Then
                pattern = False
            End If
        Next

        If pattern = True Then
            Dim con As New SqlConnection
            Dim cmd As New SqlCommand
            Dim uid As String = "XXXXXXXXXX"
            Dim hid As String = "100001"
            Try
                con.ConnectionString = "data source=WIN-KSTUPT6CJRC;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=True;multipleactiveresultsets=True;"
                con.Open()
                cmd.Connection = con
                cmd.CommandText = "INSERT INTO MATCHFILE VALUES ('" + uid + "','" + hid + "','" + m_planet + "','" + m_key + "','" + pstr1 + "','" + cloc + "','" + pstr2 + "','" + m_str + "');"
                cmd.ExecuteNonQuery()

            Catch ex As Exception
                Console.WriteLine("Error while inserting record on table..." & ex.Message, "Insert Records")
            Finally
                con.Close()
            End Try
        End If
    End Sub
    Sub Match_Key4(ByVal combo() As String, ByVal Cuspp() As String, ByVal cloc As String)
        Dim fl1 As String = "N"
        Dim fl2 As String = "N"
        Dim fl3 As String = "N"
        Dim fl4 As String = "N"
        Dim fl5 As String = "N"
        Dim fl6 As String = "N"
        Dim fl7 As String = "N"
        Dim fl8 As String = "N"
        Dim m_str As String = ""
        Dim xcusp(Cuspp.Length) As String
        Dim xcombo(combo.Length) As String

        Dim pstr1 As String = ""
        For i As Integer = 0 To combo.Length - 1
            pstr1 = pstr1 + combo(i)
        Next

        Dim pstr2 As String = ""
        For i As Integer = 0 To Cuspp.Length - 1
            pstr2 = pstr2 + Cuspp(i)
        Next

        'Check first position
        Array.Copy(Cuspp, xcusp, Cuspp.Length)
        Array.Copy(combo, xcombo, Cuspp.Length)

        For i As Integer = 0 To xcusp.Length - 2
            If xcombo(0) = xcusp(i) Then
                m_str = m_str + xcombo(0)
                xcusp(i) = "XX"
                fl1 = "Y"
                m_str = m_str
                Exit For
            End If
            If xcombo(0) = Cuspp(i) Then
                m_str = m_str + xcombo(0)
                fl1 = "Y"
            End If
        Next

        If fl1 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(1) = xcusp(i) Then
                    m_str = m_str + xcombo(1)
                    xcusp(i) = "XX"
                    fl2 = "Y"
                    Exit For
                End If
                If xcombo(1) = Cuspp(i) Then
                    m_str = m_str + xcombo(1)
                    fl2 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(2) = xcusp(i) Then
                    m_str = m_str + xcombo(2)
                    xcusp(i) = "XX"
                    fl3 = "Y"
                    Exit For
                End If
                If xcombo(2) = Cuspp(i) Then
                    m_str = m_str + xcombo(2)
                    fl3 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(3) = xcusp(i) Then
                    m_str = m_str + xcombo(3)
                    xcusp(i) = "XX"
                    fl4 = "Y"
                    Exit For
                End If
                If xcombo(3) = Cuspp(i) Then
                    m_str = m_str + xcombo(3)
                    fl4 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" And fl4 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(4) = xcusp(i) Then
                    m_str = m_str + xcombo(4)
                    xcusp(i) = "XX"
                    fl5 = "Y"
                    Exit For
                End If
                If xcombo(4) = Cuspp(i) Then
                    m_str = m_str + xcombo(4)
                    fl5 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" And fl4 = "Y" And fl5 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(5) = xcusp(i) Then
                    m_str = m_str + xcombo(5)
                    xcusp(i) = "XX"
                    fl6 = "Y"
                    Exit For
                End If
                If xcombo(5) = Cuspp(i) Then
                    m_str = m_str + xcombo(5)
                    fl6 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" And fl4 = "Y" And fl5 = "Y" And fl6 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(6) = xcusp(i) Then
                    m_str = m_str + xcombo(6)
                    xcusp(i) = "XX"
                    fl7 = "Y"
                    Exit For
                End If
                If xcombo(6) = Cuspp(i) Then
                    m_str = m_str + xcombo(6)
                    fl7 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" And fl4 = "Y" And fl5 = "Y" And fl6 = "Y" And fl7 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(7) = xcusp(i) Then
                    m_str = m_str + xcombo(7)
                    xcusp(i) = "XX"
                    fl8 = "Y"
                    Exit For
                End If
                If xcombo(7) = Cuspp(i) Then
                    m_str = m_str + xcombo(7)
                    fl8 = "Y"
                End If
            Next
        End If

        Dim pattern As Boolean = True
        For i As Integer = 0 To xcusp.Length - 2
            If xcusp(i) IsNot "XX" Then
                pattern = False
            End If
        Next

        If pattern = True Then
            Dim con As New SqlConnection
            Dim cmd As New SqlCommand
            Dim uid As String = "XXXXXXXXXX"
            Dim hid As String = "100001"
            Try
                con.ConnectionString = "data source=WIN-KSTUPT6CJRC;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=True;multipleactiveresultsets=True;"
                con.Open()
                cmd.Connection = con
                cmd.CommandText = "INSERT INTO MATCHFILE VALUES ('" + uid + "','" + hid + "','" + m_planet + "','" + m_key + "','" + pstr1 + "','" + cloc + "','" + pstr2 + "','" + m_str + "');"
                cmd.ExecuteNonQuery()

            Catch ex As Exception
                Console.WriteLine("Error while inserting record on table..." & ex.Message, "Insert Records")
            Finally
                con.Close()
            End Try
        End If
    End Sub
    Sub Match_Key5(ByVal combo() As String, ByVal Cuspp() As String, ByVal cloc As String)
        Dim fl1 As String = "N"
        Dim fl2 As String = "N"
        Dim fl3 As String = "N"
        Dim fl4 As String = "N"
        Dim fl5 As String = "N"
        Dim fl6 As String = "N"
        Dim fl7 As String = "N"
        Dim fl8 As String = "N"
        Dim m_str As String = ""
        Dim xcusp(Cuspp.Length) As String
        Dim xcombo(combo.Length) As String

        Dim pstr1 As String = ""
        For i As Integer = 0 To combo.Length - 1
            pstr1 = pstr1 + combo(i)
        Next

        Dim pstr2 As String = ""
        For i As Integer = 0 To Cuspp.Length - 1
            pstr2 = pstr2 + Cuspp(i)
        Next

        'Check first position
        Array.Copy(Cuspp, xcusp, Cuspp.Length)
        Array.Copy(combo, xcombo, Cuspp.Length)

        For i As Integer = 0 To xcusp.Length - 2
            If xcombo(0) = xcusp(i) Then
                m_str = m_str + xcombo(0)
                xcusp(i) = "XX"
                fl1 = "Y"
                m_str = m_str
                Exit For
            End If
            If xcombo(0) = Cuspp(i) Then
                m_str = m_str + xcombo(0)
                fl1 = "Y"
            End If
        Next

        If fl1 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(1) = xcusp(i) Then
                    m_str = m_str + xcombo(1)
                    xcusp(i) = "XX"
                    fl2 = "Y"
                    Exit For
                End If
                If xcombo(1) = Cuspp(i) Then
                    m_str = m_str + xcombo(1)
                    fl2 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(2) = xcusp(i) Then
                    m_str = m_str + xcombo(2)
                    xcusp(i) = "XX"
                    fl3 = "Y"
                    Exit For
                End If
                If xcombo(2) = Cuspp(i) Then
                    m_str = m_str + xcombo(2)
                    fl3 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(3) = xcusp(i) Then
                    m_str = m_str + xcombo(3)
                    xcusp(i) = "XX"
                    fl4 = "Y"
                    Exit For
                End If
                If xcombo(3) = Cuspp(i) Then
                    m_str = m_str + xcombo(3)
                    fl4 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" And fl4 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(4) = xcusp(i) Then
                    m_str = m_str + xcombo(4)
                    xcusp(i) = "XX"
                    fl5 = "Y"
                    Exit For
                End If
                If xcombo(4) = Cuspp(i) Then
                    m_str = m_str + xcombo(4)
                    fl5 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" And fl4 = "Y" And fl5 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(5) = xcusp(i) Then
                    m_str = m_str + xcombo(5)
                    xcusp(i) = "XX"
                    fl6 = "Y"
                    Exit For
                End If
                If xcombo(5) = Cuspp(i) Then
                    m_str = m_str + xcombo(5)
                    fl6 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" And fl4 = "Y" And fl5 = "Y" And fl6 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(6) = xcusp(i) Then
                    m_str = m_str + xcombo(6)
                    xcusp(i) = "XX"
                    fl7 = "Y"
                    Exit For
                End If
                If xcombo(6) = Cuspp(i) Then
                    m_str = m_str + xcombo(6)
                    fl7 = "Y"
                End If
            Next
        End If

        If fl1 = "Y" And fl2 = "Y" And fl3 = "Y" And fl4 = "Y" And fl5 = "Y" And fl6 = "Y" And fl7 = "Y" Then
            For i As Integer = 0 To xcusp.Length - 2
                If xcombo(7) = xcusp(i) Then
                    m_str = m_str + xcombo(7)
                    xcusp(i) = "XX"
                    fl8 = "Y"
                    Exit For
                End If
                If xcombo(7) = Cuspp(i) Then
                    m_str = m_str + xcombo(7)
                    fl8 = "Y"
                End If
            Next
        End If

        Dim pattern As Boolean = True
        For i As Integer = 0 To xcusp.Length - 2
            If xcusp(i) IsNot "XX" Then
                pattern = False
            End If
        Next
        'Dim file As System.IO.StreamWriter
        'file = My.Computer.FileSystem.OpenTextFileWriter("c:\test.txt", True)
        'file.WriteLine("Here is the first string.")
        'file.Close()
        If pattern = True Then
            Dim con As New SqlConnection
            Dim cmd As New SqlCommand
            Dim uid As String = "XXXXXXXXXX"
            Dim hid As String = "100001"
            Try
                con.ConnectionString = "data source=WIN-KSTUPT6CJRC;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=True;multipleactiveresultsets=True;"
                con.Open()
                cmd.Connection = con
                cmd.CommandText = "INSERT INTO MATCHFILE VALUES ('" + uid + "','" + hid + "','" + m_planet + "','" + m_key + "','" + pstr1 + "','" + cloc + "','" + pstr2 + "','" + m_str + "');"
                cmd.ExecuteNonQuery()

            Catch ex As Exception
                Console.WriteLine("Error while inserting record on table..." & ex.Message, "Insert Records")
            Finally
                con.Close()
            End Try
        End If
    End Sub
End Module
