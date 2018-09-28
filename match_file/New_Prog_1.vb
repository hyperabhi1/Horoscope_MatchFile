Imports System.Data.SqlClient

Module New_Prog_1
    Dim c01(), c02(), c03(), c04(), c05(), c06(), c07(), c08(), c09(), c10(), c11(), c12() As String
    Dim SelectCUSP = "SELECT * FROM CUSP;"
    Dim m_planet = "ME"
    Dim SelectF2BASE = "TRUNCATE TABLE MATCHFILE_ME; SELECT * FROM F2BASE;"
    Dim connstr = "data source=WIN-KSTUPT6CJRC;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=True;multipleactiveresultsets=True;"
    'Dim connstr = "data source=DESKTOP-JBRFH9E;initial catalog=testdb;integrated security=True;"
    'Dim connstr = "data source=49.50.103.132;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=False;User Id=sa;password=pSI)TA1t0K[)"
    Sub Main()
        FillCuspStrings()
        Dim con0 As New SqlConnection(connstr)
        Dim connection0 As SqlConnection = New SqlConnection(connstr)
        connection0.Open()
        Dim cmd0 As New SqlCommand(SelectF2BASE, con0)
        Dim da0 As New SqlDataAdapter(cmd0)
        Dim ds0 As New DataSet()
        Console.WriteLine("select query fired at :" + DateTime.Now.ToString())
        da0.Fill(ds0)
        Console.WriteLine("select query execs at :" + DateTime.Now.ToString())
        Dim counter As Integer = 0
        Console.WriteLine("row match start :" + DateTime.Now.ToString())
        Dim dt As DateTime = DateTime.Now
        EmailNotify.SendEmail("MATCHFILE_Me started filling", DateTime.Now, "start")
        For Each Row As DataRow In ds0.Tables(0).Rows
            MatchKey(m_planet, Row.Item(1).ToString(), Row.Item(2).ToString(), Row.Item(3).ToString(), Row.Item(4).ToString(), Row.Item(5).ToString(), Row.Item(6).ToString(), Row.Item(7).ToString(), c01, "01")
            MatchKey(m_planet, Row.Item(1).ToString(), Row.Item(2).ToString(), Row.Item(3).ToString(), Row.Item(4).ToString(), Row.Item(5).ToString(), Row.Item(6).ToString(), Row.Item(7).ToString(), c02, "02")
            MatchKey(m_planet, Row.Item(1).ToString(), Row.Item(2).ToString(), Row.Item(3).ToString(), Row.Item(4).ToString(), Row.Item(5).ToString(), Row.Item(6).ToString(), Row.Item(7).ToString(), c03, "03")
            MatchKey(m_planet, Row.Item(1).ToString(), Row.Item(2).ToString(), Row.Item(3).ToString(), Row.Item(4).ToString(), Row.Item(5).ToString(), Row.Item(6).ToString(), Row.Item(7).ToString(), c04, "04")
            MatchKey(m_planet, Row.Item(1).ToString(), Row.Item(2).ToString(), Row.Item(3).ToString(), Row.Item(4).ToString(), Row.Item(5).ToString(), Row.Item(6).ToString(), Row.Item(7).ToString(), c05, "05")
            MatchKey(m_planet, Row.Item(1).ToString(), Row.Item(2).ToString(), Row.Item(3).ToString(), Row.Item(4).ToString(), Row.Item(5).ToString(), Row.Item(6).ToString(), Row.Item(7).ToString(), c06, "06")
            MatchKey(m_planet, Row.Item(1).ToString(), Row.Item(2).ToString(), Row.Item(3).ToString(), Row.Item(4).ToString(), Row.Item(5).ToString(), Row.Item(6).ToString(), Row.Item(7).ToString(), c07, "07")
            MatchKey(m_planet, Row.Item(1).ToString(), Row.Item(2).ToString(), Row.Item(3).ToString(), Row.Item(4).ToString(), Row.Item(5).ToString(), Row.Item(6).ToString(), Row.Item(7).ToString(), c08, "08")
            MatchKey(m_planet, Row.Item(1).ToString(), Row.Item(2).ToString(), Row.Item(3).ToString(), Row.Item(4).ToString(), Row.Item(5).ToString(), Row.Item(6).ToString(), Row.Item(7).ToString(), c09, "09")
            MatchKey(m_planet, Row.Item(1).ToString(), Row.Item(2).ToString(), Row.Item(3).ToString(), Row.Item(4).ToString(), Row.Item(5).ToString(), Row.Item(6).ToString(), Row.Item(7).ToString(), c10, "10")
            MatchKey(m_planet, Row.Item(1).ToString(), Row.Item(2).ToString(), Row.Item(3).ToString(), Row.Item(4).ToString(), Row.Item(5).ToString(), Row.Item(6).ToString(), Row.Item(7).ToString(), c11, "11")
            MatchKey(m_planet, Row.Item(1).ToString(), Row.Item(2).ToString(), Row.Item(3).ToString(), Row.Item(4).ToString(), Row.Item(5).ToString(), Row.Item(6).ToString(), Row.Item(7).ToString(), c12, "12")
        Next
        EmailNotify.SendEmail("MATCHFILE_Me filled up", dt, "end")
        connection0.Close()
    End Sub
    Sub MatchKey(ByVal m_planet As String, ByVal p1 As String, ByVal p2 As String, ByVal p3 As String, ByVal p4 As String, ByVal p5 As String, ByVal p6 As String, ByVal p7 As String, ByVal cusp() As String, ByVal p_cusp As String)
        Dim combo_len = 8
        Dim cusp_len = cusp.Length
        Dim xcusp(cusp.Length - 1) As String
        Array.Copy(cusp, xcusp, cusp.Length)
        Dim cusp_str = ""

        Dim fl1 = ""
        Dim fl2 = ""
        Dim fl3 = ""
        Dim fl4 = ""
        Dim fl5 = ""
        Dim fl6 = ""
        Dim fl7 = ""
        Dim fl8 = ""

        Dim f1 = 0
        Dim f2 = 0
        Dim f3 = 0
        Dim f4 = 0
        Dim f5 = 0
        Dim f6 = 0
        Dim f7 = 0
        Dim f8 = 0

        For i As Integer = 0 To cusp.Length - 1
            cusp_str = cusp_str + cusp(i)
        Next

        Dim m_combo = ""

        Dim combo(8) As String

        combo(0) = m_planet
        combo(1) = p1
        combo(2) = p2
        combo(3) = p3
        combo(4) = p4
        combo(5) = p5
        combo(6) = p6
        combo(7) = p7

        Dim m_comb As String = ""
        For i As Integer = 0 To combo.Length - 1
            m_comb = m_comb + combo(i)
        Next

        For i As Integer = 0 To combo.Length - 1
            For k As Integer = 0 To cusp_len - 1
                If combo(i) = xcusp(k) Then
                    f1 = i
                    xcusp(k) = "XX"
                    fl1 = "Y"
                    Exit For
                End If
                If fl1 = "Y" Then
                    Exit For
                End If
            Next
        Next

        f2 = f1 + 1
        If f2 > 8 Then
            f2 = 1
        End If

        If fl1 = "Y" Then
            For k = 0 To cusp_len - 1
                If combo(f2) = xcusp(k) Then
                    xcusp(k) = "XX"
                    fl1 = "Y"
                    Exit For
                End If
            Next
        End If

        f3 = f2 + 1
        If f3 > 8 Then
            f3 = 1
        End If

        If fl2 = "Y" Then
            For k = 0 To cusp_len - 1
                If combo(f3) = xcusp(k) Then
                    xcusp(k) = "XX"
                    fl3 = "Y"
                    Exit For
                End If
            Next
        End If

        f4 = f3 + 1
        If f4 > 8 Then
            f4 = 1
        End If

        If fl3 = "Y" Then
            For k = 0 To cusp_len - 1
                If combo(f4) = xcusp(k) Then
                    xcusp(k) = "XX"
                    fl4 = "Y"
                    Exit For
                End If
            Next
        End If

        f5 = f4 + 1
        If f5 > 8 Then
            f5 = 1
        End If

        If fl4 = "Y" Then
            For k = 0 To cusp_len - 1
                If combo(f5) = xcusp(k) Then
                    xcusp(k) = "XX"
                    fl5 = "Y"
                    Exit For
                End If
            Next
        End If

        f6 = f5 + 1
        If f6 > 8 Then
            f6 = 1
        End If

        If fl5 = "Y" Then
            For k = 0 To cusp_len - 1
                If combo(f6) = xcusp(k) Then
                    xcusp(k) = "XX"
                    fl6 = "Y"
                    Exit For
                End If
            Next
        End If

        f7 = f6 + 1
        If f7 > 8 Then
            f7 = 1
        End If

        If fl6 = "Y" Then
            For k = 0 To cusp_len - 1
                If combo(f7) = xcusp(k) Then
                    xcusp(k) = "XX"
                    fl7 = "Y"
                    Exit For
                End If
            Next
        End If

        f8 = f7 + 1
        If f8 > 8 Then
            f8 = 1
        End If

        If fl7 = "Y" Then
            For k = 0 To cusp_len - 1
                If combo(f8) = xcusp(k) Then
                    xcusp(k) = "XX"
                    fl8 = "Y"
                    Exit For
                End If
            Next
        End If

        Dim pattern = True
        For i = 0 To xcusp.Length - 1
            If xcusp(i) IsNot "XX" Then
                pattern = False
            End If
        Next

        If pattern = True Then
            Dim uid As String = "XXXXXXXXXX"
            Dim hid As String = "1000000001"

            Dim con As New SqlConnection(connstr)
            Dim connection As SqlConnection = New SqlConnection(connstr)
            connection.Open()
            Dim query As String

            query = "INSERT INTO MATCHFILE_ME VALUES ('" + uid + "','" + hid + "','" + m_planet + "','" + m_comb + "','" + p_cusp + "','" + cusp_str + "')"
            Dim cmd As New SqlCommand(query, con)
            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet()
            da.Fill(ds)
            connection.Close()
        End If
    End Sub
    Sub FillCuspStrings()
        Dim M_uid As String = "XXXXXXXXXX"
        Dim M_Hid As String = "1000000001"
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
End Module
