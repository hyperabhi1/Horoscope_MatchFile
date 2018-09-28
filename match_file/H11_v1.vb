Imports System.Data.SqlClient

Module H11_v1
    Dim c01(), c02(), c03(), c04(), c05(), c06(), c07(), c08(), c09(), c10(), c11(), c12() As String
    Dim m_planet = "JU"
    Dim m_key = ""
    Dim counter_JU = 0
    Dim primary = ""
    Dim connstr = "data source=49.50.103.132;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=False;User Id=sa;password=pSI)TA1t0K[)"
    Dim a_keystr1() As String
    Sub Main()
        FillCusp()
    End Sub
    Sub FillCusp(Optional ByRef uid As String = "asppdeo@gmail.com", Optional ByRef hid As String = "1")
        Dim getCUSP = "SELECT * FROM HEADLETTERS_ENGINE.DBO.CUSP WHERE UID = '" + uid + "' AND HID = '" + hid + "';"
        Dim con As New SqlConnection(connstr)
        Dim connection As SqlConnection = New SqlConnection(connstr)
        Dim cmd As New SqlCommand(getCUSP, con)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        Try
            connection.Open()
            da.Fill(ds)
        Catch ex As Exception
        Finally
            connection.Close()
        End Try
        If ds.Tables(0).Rows.Count <> 0 Then
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
            MatchFile(uid, hid)
        Else
        End If
    End Sub
    Sub MatchFile(ByRef uid As String, ByRef hid As String)
        Dim m_planet = "JU"
        Dim a8(8) As String
        Dim con As New SqlConnection(connstr)
        Dim cmd As New SqlCommand("SELECT * FROM HEADLETTERS_ENGINE.DBO.F2BASE;", con)
        Dim da As New SqlDataAdapter(cmd)
        Dim RowsData As New DataSet()
        Console.WriteLine("Select * From F2BASE For (JU) Started at: " + DateTime.Now.ToString())
        Try
            con.Open()
            da.Fill(RowsData)
        Catch ex As Exception
        Finally
            con.Close()
        End Try
        Console.WriteLine("Select * From F2BASE For (JU) Ended at: " + DateTime.Now.ToString())
        Process_match_Key_set(uid, hid, RowsData, m_planet)
    End Sub
    Sub Process_match_Key_set(ByRef uid As String, ByRef hid As String, ByVal RowsData As DataSet, ByRef m_planet As String)
        Dim a8(48) As String

        While counter_JU > -1
            Dim x = counter_JU
            counter_JU -= 1
            Try
                Dim keystr1 = ""
                primary = "N"
                Dim pa = m_planet
                Dim p1 = RowsData.Tables(0).Rows(0)(0).Trim.ToString()
                Dim p2 = RowsData.Tables(0).Rows(0)(1).Trim.ToString()
                Dim p3 = RowsData.Tables(0).Rows(0)(2).Trim.ToString()
                Dim p4 = RowsData.Tables(0).Rows(0)(3).Trim.ToString()
                Dim p5 = RowsData.Tables(0).Rows(0)(4).Trim.ToString()
                Dim p6 = RowsData.Tables(0).Rows(0)(5).Trim.ToString()
                Dim m_key = pa + p1 + p2 + p3 + p4 + p5 + p6

                a8(0) = pa
                a8(1) = p1
                a8(2) = p2
                a8(3) = p3
                a8(4) = p4
                a8(5) = p5
                a8(6) = p6

                a8(7) = p1
                a8(8) = p2
                a8(9) = p3
                a8(10) = p4
                a8(11) = p5
                a8(12) = p6
                a8(13) = pa

                a8(14) = p2
                a8(15) = p3
                a8(16) = p4
                a8(17) = p5
                a8(18) = p6
                a8(19) = pa
                a8(20) = p1

                a8(21) = p3
                a8(22) = p4
                a8(23) = p5
                a8(24) = p6
                a8(25) = pa
                a8(26) = p1
                a8(27) = p2

                a8(28) = p4
                a8(29) = p5
                a8(30) = p6
                a8(31) = pa
                a8(32) = p1
                a8(33) = p2
                a8(34) = p3

                a8(35) = p5
                a8(36) = p6
                a8(37) = pa
                a8(38) = p1
                a8(39) = p2
                a8(40) = p3
                a8(41) = p4

                a8(42) = p6
                a8(43) = pa
                a8(44) = p1
                a8(45) = p2
                a8(46) = p3
                a8(47) = p4
                a8(48) = p5

                MatchKey(keystr1, uid, hid, pa, m_key, a8, c01, "01")
                MatchKey(keystr1, uid, hid, pa, m_key, a8, c02, "02")
                MatchKey(keystr1, uid, hid, pa, m_key, a8, c03, "03")
                MatchKey(keystr1, uid, hid, pa, m_key, a8, c04, "04")
                MatchKey(keystr1, uid, hid, pa, m_key, a8, c05, "05")
                MatchKey(keystr1, uid, hid, pa, m_key, a8, c06, "06")
                MatchKey(keystr1, uid, hid, pa, m_key, a8, c07, "07")
                MatchKey(keystr1, uid, hid, pa, m_key, a8, c08, "08")
                MatchKey(keystr1, uid, hid, pa, m_key, a8, c09, "09")
                MatchKey(keystr1, uid, hid, pa, m_key, a8, c10, "10")
                MatchKey(keystr1, uid, hid, pa, m_key, a8, c11, "11")
                MatchKey(keystr1, uid, hid, pa, m_key, a8, c12, "12")
                Dim a_keystr1() As String = {"", "", "", "", "", "", "", "", "", "", "", ""}

                Dim pointer = 0
                If Len(keystr1) > 2 Then
                    For i = 1 To (Len(keystr1) / 2)
                        a_keystr1(i) = keystr1.Substring(pointer, 2)
                        pointer = pointer + 2
                    Next
                    Dim noelements = a_keystr1.Length
                    Array.Sort(a_keystr1)
                    For i As Integer = 0 To noelements - 2
                        Dim startstr = a_keystr1(i)
                        For j = i + 1 To noelements - 1
                            Dim con As New SqlConnection
                            Dim cmd As New SqlCommand
                            Try
                                con.ConnectionString = connstr
                                con.Open()
                                cmd.Connection = con
                                cmd.CommandText = "INSERT INTO HEADLETTERS_ENGINE.DBO.PROMISE VALUES ('" + uid + "','" + hid + "','" + startstr + a_keystr1(j) + "','" + primary + "','" + m_planet + "');"
                                cmd.ExecuteNonQuery()
                            Catch ex As Exception
                            Finally
                                con.Close()
                            End Try
                        Next
                    Next
                End If
            Catch ex As Exception
                Exit While
            End Try
        End While
    End Sub
    Sub MatchKey(ByRef keystr1 As String, ByRef uid As String, ByRef hid As String, ByRef m_planet As String, ByVal m_key As String, ByVal combo() As String, ByVal Cuspp() As String, ByVal cloc As String)
        Dim y2() As String = {"YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY"}
        Dim c = Cuspp.Length - 1
        Dim z1(c) As String
        Dim mstr = ""
        For i As Integer = 0 To 48
            For j As Integer = 0 To c
                If combo(i) = Cuspp(j) Then
                    y2(i) = "XX"
                End If
            Next
        Next
        For mega As Integer = 0 To 6
            Dim pstr2 = ""
            mstr = ""
            Array.Copy(Cuspp, z1, c + 1)
            Dim start = (mega) * 7
            Dim finish = start + 6
            For i = start To finish
                If y2(i) <> "XX" Then
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
                If z1(m) <> "XX" Then
                    pattern = False
                End If
            Next
            Dim position = ""
            If pattern = True Then
                If mega = 0 Then
                    position = "Y"
                Else
                    position = "X"
                End If
                keystr1 = keystr1 + cloc
                For i As Integer = 0 To c
                    pstr2 = pstr2 + Cuspp(i)
                Next
                Dim con As New SqlConnection
                Dim cmd As New SqlCommand
                Try
                    con.ConnectionString = connstr
                    con.Open()
                    cmd.Connection = con
                    cmd.CommandText = "INSERT INTO MATCHFILE VALUES ('" + uid + "','" + hid + "','" + m_planet + "','" + m_key + "','" + cloc + "','" + pstr2 + "', '" + mega + "');"
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                Finally
                    con.Close()
                End Try
            End If
        Next
    End Sub
End Module