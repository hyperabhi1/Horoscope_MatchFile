Imports System.Data.SqlClient

Module MaxDepthMatch_file
    Dim S1(12) As String
    Dim c01(), c02(), c03(), c04(), c05(), c06(), c07(), c08(), c09(), c10(), c11(), c12() As String
    Dim m_planet As String = "MA"
    Dim m_key As String
    Dim InsertQuery = ""
    Dim InsertCounter = 0
    Dim counter = 0
    'Dim connstr = "data source=49.50.103.132;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=False;User Id=sa;password=pSI)TA1t0K[)"
    Dim connstr = "data source=WIN-KSTUPT6CJRC;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=True;multipleactiveresultsets=True;"
    'Dim connstr = "data source=DESKTOP-1JBRFH9E;initial catalog=testdb;integrated security=True;"
    Dim SelectF2PLANET = "SELECT * FROM F2PLANETS;"
    Sub Main()

        Dim starttime As DateTime = DateTime.Now
        EmailNotify.SendEmail("Match_File COUNTER-- Started", starttime, "start")
        Console.WriteLine(starttime)
        MakeCusp_checkHouse()
        EmailNotify.SendEmail("Match_File COUNTER-- Ended", starttime, "end")
        Console.WriteLine("TOTAL TIME TAKEN IS : " + DateTime.Now.Subtract(starttime).ToString())
        Console.ReadKey()
    End Sub
    Sub MakeCusp_checkHouse()
        Dim M_uid As String = "XXXXXXXXXX"
        Dim M_Hid As String = "1000000001"
        Dim SelectCUSP = "SELECT * FROM CUSP;"
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
        counter = ds0.Tables(0).Rows.Count - 1
        Dim ListF2Planets As List(Of F2Planets) = New List(Of F2Planets)
        connection.Close()

        Parallel.Invoke(
                Sub() Process_match_Key_set1(ds0),
                Sub() Process_match_Key_set2(ds0),
                Sub() Process_match_Key_set3(ds0),
                Sub() Process_match_Key_set4(ds0),
                Sub() Process_match_Key_set5(ds0),
                Sub() Process_match_Key_set1(ds0),
                Sub() Process_match_Key_set2(ds0),
                Sub() Process_match_Key_set3(ds0),
                Sub() Process_match_Key_set4(ds0),
                Sub() Process_match_Key_set5(ds0),
                Sub() Process_match_Key_set1(ds0),
                Sub() Process_match_Key_set2(ds0),
                Sub() Process_match_Key_set3(ds0),
                Sub() Process_match_Key_set4(ds0),
                Sub() Process_match_Key_set5(ds0),
                Sub() Process_match_Key_set1(ds0),
                Sub() Process_match_Key_set2(ds0),
                Sub() Process_match_Key_set3(ds0),
                Sub() Process_match_Key_set4(ds0),
                Sub() Process_match_Key_set5(ds0))


    End Sub

    Sub Process_match_Key_set1(ByVal RowsData As DataSet)
        Console.WriteLine("1 started")
        Dim a8(8) As String
        While counter > -1
            counter -= 1
            Dim i = counter
            Try
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
            Catch ex As Exception
                Exit While
            End Try
        End While
    End Sub
    Sub Process_match_Key_set2(ByVal RowsData As DataSet)
        Console.WriteLine("2 started")
        Dim a8(8) As String
        While counter > -1
            counter -= 1
            Dim i = counter
            Try
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
            Catch ex As Exception
                Exit While
            End Try
        End While

    End Sub
    Sub Process_match_Key_set3(ByVal RowsData As DataSet)
        Console.WriteLine("3 started")
        Dim a8(8) As String
        While counter > -1
            counter -= 1
            Dim i = counter
            Try
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
            Catch ex As Exception
                Exit While
            End Try
        End While
    End Sub
    Sub Process_match_Key_set4(ByVal RowsData As DataSet)
        Console.WriteLine("4 started")
        Dim a8(8) As String
        While counter > -1
            counter -= 1
            Dim i = counter
            Try
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
            Catch ex As Exception
                Exit While
            End Try
        End While
    End Sub
    Sub Process_match_Key_set5(ByVal RowsData As DataSet)
        Console.WriteLine("5 started")
        Dim a8(8) As String
        While counter > -1
            counter -= 1
            Dim i = counter
            Try
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
                a8(0) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
                a8(1) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
                a8(2) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
                a8(3) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
                a8(4) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
                a8(5) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
                a8(6) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
                a8(7) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
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
            Catch ex As Exception
                Exit While
            End Try
        End While
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
                cmd.CommandText = "INSERT INTO MATCH_FILE VALUES ('" + uid + "','" + hid + "','" + m_planet + "','" + m_key + "','" + pstr1 + "','" + cloc + "','" + pstr2 + "','" + m_str + "');"
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
                cmd.CommandText = "INSERT INTO MATCH_FILE VALUES ('" + uid + "','" + hid + "','" + m_planet + "','" + m_key + "','" + pstr1 + "','" + cloc + "','" + pstr2 + "','" + m_str + "');"
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
                cmd.CommandText = "INSERT INTO MATCH_FILE VALUES ('" + uid + "','" + hid + "','" + m_planet + "','" + m_key + "','" + pstr1 + "','" + cloc + "','" + pstr2 + "','" + m_str + "');"
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
                cmd.CommandText = "INSERT INTO MATCH_FILE VALUES ('" + uid + "','" + hid + "','" + m_planet + "','" + m_key + "','" + pstr1 + "','" + cloc + "','" + pstr2 + "','" + m_str + "');"
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
                cmd.CommandText = "INSERT INTO MATCH_FILE VALUES ('" + uid + "','" + hid + "','" + m_planet + "','" + m_key + "','" + pstr1 + "','" + cloc + "','" + pstr2 + "','" + m_str + "');"
                cmd.ExecuteNonQuery()

            Catch ex As Exception
                Console.WriteLine("Error while inserting record on table..." & ex.Message, "Insert Records")
            Finally
                con.Close()
            End Try
        End If
    End Sub

End Module
Class F2Planets
    Public Property key As Integer
    Public Property p0 As String
    Public Property p1 As String
    Public Property p2 As String
    Public Property p3 As String
    Public Property p4 As String
    Public Property p5 As String
    Public Property p6 As String
    Public Property p7 As String
End Class
