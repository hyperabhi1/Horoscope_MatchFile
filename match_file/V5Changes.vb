Imports System.Data.SqlClient
Module V5Changes
    Dim c01(), c02(), c03(), c04(), c05(), c06(), c07(), c08(), c09(), c10(), c11(), c12() As String
    Dim counter = 0
    Dim ThreadStartCounter = 0
    Dim ThreadEndCounter = 0
    'Dim connstr = "data source=49.50.103.132;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=False;User Id=sa;password=pSI)TA1t0K[)"
    Dim connstr = "data source=WIN-KSTUPT6CJRC;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=True;multipleactiveresultsets=True;"
    'Dim connstr = "data source=DESKTOP-T751K49;initial catalog=testdb;integrated security=True;"
    Sub Main()
        Dim ST As DateTime = DateTime.Now
        ConsoleLogs = ConsoleLogs + "Sending Start Mail at: " + DateTime.Now.ToString()
        EmailNotify.SendEmail("NEWMATCH + COUNTER-- FOR (MA PLANET) Logic Test " + ST.ToString(), ST, "start")
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "Start Mail Sent at: " + DateTime.Now.ToString()
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "Fetching CUSP Details at: " + DateTime.Now.ToString()
        FillCusp()
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "CUSP Details Fetched at: " + DateTime.Now.ToString()
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "Initiating MatchFile Program at: " + DateTime.Now.ToString()
        'Parallel.Invoke(Sub() Match_File_MA(), Sub() Match_File_ME(), Sub() Match_File_MO(), Sub() Match_File_JU(), Sub() Match_File_SA(), Sub() Match_File_SU(), Sub() Match_File_VE())
        Match_File_MA()
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "MatchFile Program Finished at: " + DateTime.Now.ToString()
        NumberOfRecordsInMatchFile()
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "Total Time taken : " + DateTime.Now.Subtract(ST).ToString()
        Console.WriteLine("TOTAL TIME TAKEN IS : " + DateTime.Now.Subtract(ST).ToString())
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "Sending Finish Mail at: " + DateTime.Now.ToString()
        EmailNotify.SendEmail("NEWMATCH + COUNTER-- FOR (MA PLANET) Logic Test " + ST.ToString(), ST, "end")
        ConsoleLogs = ConsoleLogs + "Finish Mail Sent."
        'Console.ReadKey()
    End Sub
    Sub FillCusp()
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
    End Sub
    Sub Match_File_MA()
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "MatchFile_MA Program Started at: " + DateTime.Now.ToString()
        Dim m_planet = "MA"
        Dim a8(8) As String
        Dim con As New SqlConnection(connstr)
        Dim connection As SqlConnection = New SqlConnection(connstr)
        connection.Open()
        Dim cmd As New SqlCommand("TRUNCATE TABLE MATCH_FILE; SELECT * FROM F2PLANETS WHERE p0 = 'MA';", con)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "Select * From F2PLANETS For (MA) Started at: " + DateTime.Now.ToString()
        Console.WriteLine("Select * From F2PLANETS For (MA) Started at: " + DateTime.Now.ToString())
        da.Fill(ds)
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "Select * From F2PLANETS For (MA) Ended at: " + DateTime.Now.ToString()
        Console.WriteLine("Select * From F2PLANETS For (MA) Ended at: " + DateTime.Now.ToString())
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "<b>Total Records in F2PLANETS-(MA) are: " + ds.Tables(0).Rows.Count.ToString() + "</b>"
        counter = ds.Tables(0).Rows.Count - 1
        connection.Close()
        Parallel.Invoke(
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet))
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "MatchFile_MA Program finished at: " + DateTime.Now.ToString()
    End Sub
    Sub Match_File_JU()
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "MatchFile_JU Program Started at: " + DateTime.Now.ToString()
        Dim m_planet = "JU"
        Dim a8(8) As String
        Dim con As New SqlConnection(connstr)
        Dim connection As SqlConnection = New SqlConnection(connstr)
        connection.Open()
        Dim cmd As New SqlCommand("TRUNCATE TABLE MATCH_FILE; SELECT * FROM F2PLANETS WHERE p0 = 'JU';", con)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "Select * From F2PLANETS For (JU) Started at: " + DateTime.Now.ToString()
        Console.WriteLine("Select * From F2PLANETS For (JU) Started at: " + DateTime.Now.ToString())
        da.Fill(ds)
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "Select * From F2PLANETS For (JU) Ended at: " + DateTime.Now.ToString()
        Console.WriteLine("Select * From F2PLANETS For (JU) Ended at: " + DateTime.Now.ToString())
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "<b>Total Records in F2PLANETS-(JU) are: " + ds.Tables(0).Rows.Count.ToString() + "</b>"
        counter = ds.Tables(0).Rows.Count - 1
        connection.Close()
        Parallel.Invoke(
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet))
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "MatchFile_JU Program finished at: " + DateTime.Now.ToString()
    End Sub
    Sub Match_File_SA()
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "MatchFile_SA Program Started at: " + DateTime.Now.ToString()
        Dim m_planet = "SA"
        Dim a8(8) As String
        Dim con As New SqlConnection(connstr)
        Dim connection As SqlConnection = New SqlConnection(connstr)
        connection.Open()
        Dim cmd As New SqlCommand("TRUNCATE TABLE MATCH_FILE; SELECT * FROM F2PLANETS WHERE p0 = 'SA';", con)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "Select * From F2PLANETS For (SA) Started at: " + DateTime.Now.ToString()
        Console.WriteLine("Select * From F2PLANETS For (SA) Started at: " + DateTime.Now.ToString())
        da.Fill(ds)
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "Select * From F2PLANETS For (SA) Ended at: " + DateTime.Now.ToString()
        Console.WriteLine("Select * From F2PLANETS For (SA) Ended at: " + DateTime.Now.ToString())
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "<b>Total Records in F2PLANETS-(SA) are: " + ds.Tables(0).Rows.Count.ToString() + "</b>"
        counter = ds.Tables(0).Rows.Count - 1
        connection.Close()
        Parallel.Invoke(
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet))
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "MatchFile_SA Program finished at: " + DateTime.Now.ToString()
    End Sub
    Sub Match_File_ME()
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "MatchFile_ME Program Started at: " + DateTime.Now.ToString()
        Dim m_planet = "ME"
        Dim a8(8) As String
        Dim con As New SqlConnection(connstr)
        Dim connection As SqlConnection = New SqlConnection(connstr)
        connection.Open()
        Dim cmd As New SqlCommand("TRUNCATE TABLE MATCH_FILE; SELECT * FROM F2PLANETS WHERE p0 = 'ME';", con)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "Select * From F2PLANETS For (ME) Started at: " + DateTime.Now.ToString()
        Console.WriteLine("Select * From F2PLANETS For (ME) Started at: " + DateTime.Now.ToString())
        da.Fill(ds)
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "Select * From F2PLANETS For (ME) Ended at: " + DateTime.Now.ToString()
        Console.WriteLine("Select * From F2PLANETS For (ME) Ended at: " + DateTime.Now.ToString())
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "<b>Total Records in F2PLANETS-(ME) are: " + ds.Tables(0).Rows.Count.ToString() + "</b>"
        counter = ds.Tables(0).Rows.Count - 1
        connection.Close()
        Parallel.Invoke(
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet))
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "MatchFile_ME Program finished at: " + DateTime.Now.ToString()
    End Sub
    Sub Match_File_MO()
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "MatchFile_MO Program Started at: " + DateTime.Now.ToString()
        Dim m_planet = "MO"
        Dim a8(8) As String
        Dim con As New SqlConnection(connstr)
        Dim connection As SqlConnection = New SqlConnection(connstr)
        connection.Open()
        Dim cmd As New SqlCommand("TRUNCATE TABLE MATCH_FILE; SELECT * FROM F2PLANETS WHERE p0 = 'MO';", con)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "Select * From F2PLANETS For (MO) Started at: " + DateTime.Now.ToString()
        Console.WriteLine("Select * From F2PLANETS For (MO) Started at: " + DateTime.Now.ToString())
        da.Fill(ds)
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "Select * From F2PLANETS For (MO) Ended at: " + DateTime.Now.ToString()
        Console.WriteLine("Select * From F2PLANETS For (MO) Ended at: " + DateTime.Now.ToString())
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "<b>Total Records in F2PLANETS-(MO) are: " + ds.Tables(0).Rows.Count.ToString() + "</b>"
        counter = ds.Tables(0).Rows.Count - 1
        connection.Close()
        Parallel.Invoke(
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet))
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "MatchFile_MO Program finished at: " + DateTime.Now.ToString()
    End Sub
    Sub Match_File_VE()
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "MatchFile_VE Program Started at: " + DateTime.Now.ToString()
        Dim m_planet = "VE"
        Dim a8(8) As String
        Dim con As New SqlConnection(connstr)
        Dim connection As SqlConnection = New SqlConnection(connstr)
        connection.Open()
        Dim cmd As New SqlCommand("TRUNCATE TABLE MATCH_FILE; SELECT * FROM F2PLANETS WHERE p0 = 'VE';", con)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "Select * From F2PLANETS For (VE) Started at: " + DateTime.Now.ToString()
        Console.WriteLine("Select * From F2PLANETS For (VE) Started at: " + DateTime.Now.ToString())
        da.Fill(ds)
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "Select * From F2PLANETS For (VE) Ended at: " + DateTime.Now.ToString()
        Console.WriteLine("Select * From F2PLANETS For (VE) Ended at: " + DateTime.Now.ToString())
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "<b>Total Records in F2PLANETS-(VE) are: " + ds.Tables(0).Rows.Count.ToString() + "</b>"
        counter = ds.Tables(0).Rows.Count - 1
        connection.Close()
        Parallel.Invoke(
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet))
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "MatchFile_VE Program finished at: " + DateTime.Now.ToString()
    End Sub
    Sub Match_File_SU()
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "MatchFile_SU Program Started at: " + DateTime.Now.ToString()
        Dim m_planet = "SU"
        Dim a8(8) As String
        Dim con As New SqlConnection(connstr)
        Dim connection As SqlConnection = New SqlConnection(connstr)
        connection.Open()
        Dim cmd As New SqlCommand("TRUNCATE TABLE MATCH_FILE; SELECT * FROM F2PLANETS WHERE p0 = 'SU';", con)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "Select * From F2PLANETS For (SU) Started at: " + DateTime.Now.ToString()
        Console.WriteLine("Select * From F2PLANETS For (SU) Started at: " + DateTime.Now.ToString())
        da.Fill(ds)
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "Select * From F2PLANETS For (SU) Ended at: " + DateTime.Now.ToString()
        Console.WriteLine("Select * From F2PLANETS For (SU) Ended at: " + DateTime.Now.ToString())
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "<b>Total Records in F2PLANETS-(SU) are: " + ds.Tables(0).Rows.Count.ToString() + "</b>"
        counter = ds.Tables(0).Rows.Count - 1
        connection.Close()
        Parallel.Invoke(
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet),
                Sub() Process_match_Key_set(ds, m_planet))
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "MatchFile_SU Program finished at: " + DateTime.Now.ToString()
    End Sub
    Sub Process_match_Key_set(ByVal RowsData As DataSet, ByRef m_planet As String)
        ThreadStartCounter += 1
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "Thread #" + ThreadStartCounter.ToString() + " Started at : " + DateTime.Now.ToString()
        Console.WriteLine("Thread #" + ThreadStartCounter.ToString() + " Started at " + DateTime.Now.ToString())
        Dim a8(63) As String
        While counter > -1
            Dim i = counter
            counter -= 1
            Try
                Dim pa = m_planet
                Dim m_key = RowsData.Tables(0).Rows(i)(0).Trim.ToString() + RowsData.Tables(0).Rows(i)(1).Trim.ToString() + RowsData.Tables(0).Rows(i)(2).Trim.ToString() + RowsData.Tables(0).Rows(i)(3).Trim.ToString() + RowsData.Tables(0).Rows(i)(4).Trim.ToString() + RowsData.Tables(0).Rows(i)(5).Trim.ToString() + RowsData.Tables(0).Rows(i)(6).Trim.ToString() + RowsData.Tables(0).Rows(i)(7).Trim.ToString()

                a8(0) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
                a8(1) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
                a8(2) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
                a8(3) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
                a8(4) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
                a8(5) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
                a8(6) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
                a8(7) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()

                a8(9) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
                a8(10) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
                a8(11) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
                a8(12) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
                a8(13) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
                a8(14) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
                a8(14) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
                a8(15) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()

                a8(16) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
                a8(17) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
                a8(18) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
                a8(19) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
                a8(20) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
                a8(21) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
                a8(22) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
                a8(23) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()

                a8(24) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
                a8(25) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
                a8(26) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
                a8(27) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
                a8(28) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
                a8(29) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
                a8(30) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
                a8(31) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()

                a8(32) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
                a8(33) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
                a8(34) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
                a8(35) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
                a8(36) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
                a8(37) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
                a8(38) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
                a8(39) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()

                a8(40) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
                a8(41) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
                a8(42) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
                a8(43) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
                a8(44) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
                a8(45) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
                a8(46) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
                a8(47) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()

                a8(48) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
                a8(49) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
                a8(50) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
                a8(51) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
                a8(52) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
                a8(53) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
                a8(54) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
                a8(55) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()

                a8(56) = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
                a8(57) = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
                a8(58) = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
                a8(59) = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
                a8(60) = RowsData.Tables(0).Rows(i)(7).Trim.ToString()
                a8(61) = RowsData.Tables(0).Rows(i)(6).Trim.ToString()
                a8(62) = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
                a8(63) = RowsData.Tables(0).Rows(i)(4).Trim.ToString()

                Parallel.Invoke(
                Sub() Match_Key(pa, m_key, a8, c01, "01"),
                Sub() Match_Key(pa, m_key, a8, c02, "02"),
                Sub() Match_Key(pa, m_key, a8, c03, "03"),
                Sub() Match_Key(pa, m_key, a8, c04, "04"),
                Sub() Match_Key(pa, m_key, a8, c05, "05"),
                Sub() Match_Key(pa, m_key, a8, c06, "06"),
                Sub() Match_Key(pa, m_key, a8, c07, "07"),
                Sub() Match_Key(pa, m_key, a8, c08, "08"),
                Sub() Match_Key(pa, m_key, a8, c09, "09"),
                Sub() Match_Key(pa, m_key, a8, c10, "10"),
                Sub() Match_Key(pa, m_key, a8, c11, "11"),
                Sub() Match_Key(pa, m_key, a8, c12, "12"))
            Catch ex As Exception
                Exit While
            End Try
        End While
        ThreadEndCounter += 1
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "Thread #" + ThreadEndCounter.ToString() + " End at : " + DateTime.Now.ToString()
        Console.WriteLine("Thread #" + ThreadEndCounter.ToString() + " ended at " + DateTime.Now.ToString())
    End Sub
    Sub Match_Key(ByRef m_planet As String, ByVal m_key As String, ByVal combo() As String, ByVal Cuspp() As String, ByVal cloc As String)
        Dim y2(63) As String
        Dim c = Cuspp.Length - 1
        Dim z1(c) As String
        Dim mstr = ""
        For i As Integer = 0 To 63
            For j As Integer = 0 To c
                If combo(i) = Cuspp(j) Then
                    y2(i) = "XX"
                End If
            Next
        Next
        For mega As Integer = 0 To 7
            Dim pstr1 = ""
            Dim pstr2 = ""
            mstr = ""
            Array.Copy(Cuspp, z1, Cuspp.Length)
            Dim start = (mega - 0) * 8
            Dim finish = start + 7
            For i = start To finish
                If y2(i) IsNot "XX" Then
                    Exit For
                End If
                For j = 0 To c
                    If z1(j) = combo(i) Then
                        z1(j) = "XX"
                        mstr = mstr + combo(i)
                        Exit For
                    Else
                        If Cuspp(j) = combo(i) Then
                            mstr = mstr + combo(i)
                        End If
                    End If
                Next
            Next
            Dim pattern As Boolean = True
            For m As Integer = 0 To c
                If z1(m) IsNot "XX" Then
                    pattern = False
                End If
            Next
            If pattern = True Then
                For i As Integer = start To finish
                    pstr1 = pstr1 + combo(i)
                Next
                For i As Integer = 0 To c
                    pstr2 = pstr2 + Cuspp(i)
                Next
                Dim con As New SqlConnection
                Dim cmd As New SqlCommand
                Try
                    Dim uid As String = "XXXXXXXXXX"
                    Dim hid As String = "100001"
                    con.ConnectionString = connstr
                    con.Open()
                    cmd.Connection = con
                    cmd.CommandText = "INSERT INTO MATCH_FILE VALUES ('" + uid + "','" + hid + "','" + m_planet + "','" + m_key + "','" + pstr1 + "','" + cloc + "','" + pstr2 + "','" + If(mstr.Length > 16, mstr.Substring(0, 16), mstr) + "');"
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    Console.WriteLine("Error while inserting record on table..." & ex.Message, "Insert Records")
                Finally
                    con.Close()
                End Try
            End If
        Next
    End Sub
    Sub NumberOfRecordsInMatchFile()
        Dim con As New SqlConnection(connstr)
        Dim connection As SqlConnection = New SqlConnection(connstr)
        connection.Open()
        Dim cmd As New SqlCommand("SELECT * FROM MATCH_FILE;", con)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "SELECT COUNT(*) FROM MATCH_FILE Started at: " + DateTime.Now.ToString()
        Console.WriteLine("SELECT COUNT(*) FROM MATCH_FILE Started at: " + DateTime.Now.ToString())
        da.Fill(ds)
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "SELECT COUNT(*) FROM MATCH_FILE Ended at: " + DateTime.Now.ToString()
        Console.WriteLine("SELECT COUNT(*) FROM MATCH_FILE Ended at: " + DateTime.Now.ToString())
        ConsoleLogs = ConsoleLogs + Environment.NewLine + "<br>" + "<b>Total Records in MATCH_FILE are: " + ds.Tables(0).Rows.Count.ToString() + "</b>"
        connection.Close()
    End Sub
End Module
