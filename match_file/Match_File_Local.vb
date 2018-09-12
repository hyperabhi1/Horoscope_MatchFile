Imports System.Data.SqlClient

Module Match_File_Local
    Dim S1(12) As String
    Dim c01(), c02(), c03(), c04(), c05(), c06(), c07(), c08(), c09(), c10(), c11(), c12() As String
    Dim m_planet As String
    Dim f2prow As String
    Dim connstr = "data source=DESKTOP-JBRFH9E;initial catalog=testdb;integrated security=True;"
    Dim SelectF2PLANET = "truncate table match_file;SELECT top 100000 * FROM F2PLANETS;"

    Sub Main()
        Dim starttime As DateTime = DateTime.Now
        'EmailNotify.SendEmail("Match_File_Local Started", starttime, "start")
        Console.WriteLine(starttime)
        MakeCusp_checkHouse()
        'EmailNotify.SendEmail("Match_File_Local Ended", starttime, "end")
        Console.WriteLine(DateTime.Now.Subtract(starttime))
        Console.ReadKey()
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
        'Console.WriteLine(String.Join("-", c01))
        'Console.WriteLine(String.Join("-", c02))
        'Console.WriteLine(String.Join("-", c03))
        'Console.WriteLine(String.Join("-", c04))
        'Console.WriteLine(String.Join("-", c05))
        'Console.WriteLine(String.Join("-", c06))
        'Console.WriteLine(String.Join("-", c07))
        'Console.WriteLine(String.Join("-", c08))
        'Console.WriteLine(String.Join("-", c09))
        'Console.WriteLine(String.Join("-", c10))
        'Console.WriteLine(String.Join("-", c11))
        'Console.WriteLine(String.Join("-", c12))

        connection.Close()

        Dim a8(7) As String
        '       Dim SelectF2PLANET = "SELECT * FROM F2PLANETS WHERE p0 = 'SU'"
        Dim con0 As New SqlConnection(connstr)
        Dim connection0 As SqlConnection = New SqlConnection(connstr)
        connection0.Open()
        Dim cmd0 As New SqlCommand(SelectF2PLANET, con0)
        Dim da0 As New SqlDataAdapter(cmd0)
        Dim ds0 As New DataSet()
        Console.WriteLine("select query fired at :" + DateTime.Now.ToString())
        da0.Fill(ds0)
        Console.WriteLine("select query execs at :" + DateTime.Now.ToString())
        Dim counter As Integer = 0
        Console.WriteLine("row match start :" + DateTime.Now.ToString())
        For Each Row As DataRow In ds0.Tables(0).Rows
            f2prow = Row.Item(0).ToString() + Row.Item(1).ToString() + Row.Item(2).ToString() + Row.Item(3).ToString()
            f2prow = f2prow + Row.Item(4).ToString() + Row.Item(5).ToString() + Row.Item(6).ToString() + Row.Item(7).ToString()
            m_planet = Row.Item(0).ToString()
            If c01.Length = 3 Then

            End If

            'p0
            a8(0) = Row.Item(0).ToString()
            a8(1) = Row.Item(1).ToString()
            a8(2) = Row.Item(2).ToString()
            a8(3) = Row.Item(3).ToString()
            a8(4) = Row.Item(4).ToString()
            a8(5) = Row.Item(5).ToString()
            a8(6) = Row.Item(6).ToString()
            a8(7) = Row.Item(7).ToString()
            Match_Key(a8, c01, "01")
            Match_Key(a8, c02, "02")
            Match_Key(a8, c03, "03")
            Match_Key(a8, c04, "04")
            Match_Key(a8, c05, "05")
            Match_Key(a8, c06, "06")
            Match_Key(a8, c07, "07")
            Match_Key(a8, c08, "08")
            Match_Key(a8, c09, "09")
            Match_Key(a8, c10, "10")
            Match_Key(a8, c11, "11")
            Match_Key(a8, c12, "12")




            'p1
            a8(0) = Row.Item(7).ToString()
            a8(1) = Row.Item(0).ToString()
            a8(2) = Row.Item(1).ToString()
            a8(3) = Row.Item(2).ToString()
            a8(4) = Row.Item(3).ToString()
            a8(5) = Row.Item(4).ToString()
            a8(6) = Row.Item(5).ToString()
            a8(7) = Row.Item(6).ToString()
            Match_Key(a8, c01, "01")
            Match_Key(a8, c02, "02")
            Match_Key(a8, c03, "03")
            Match_Key(a8, c04, "04")
            Match_Key(a8, c05, "05")
            Match_Key(a8, c06, "06")
            Match_Key(a8, c07, "07")
            Match_Key(a8, c08, "08")
            Match_Key(a8, c09, "09")
            Match_Key(a8, c10, "10")
            Match_Key(a8, c11, "11")
            Match_Key(a8, c12, "12")

            'p2
            a8(0) = Row.Item(6).ToString()
            a8(1) = Row.Item(7).ToString()
            a8(2) = Row.Item(0).ToString()
            a8(3) = Row.Item(1).ToString()
            a8(4) = Row.Item(2).ToString()
            a8(5) = Row.Item(3).ToString()
            a8(6) = Row.Item(4).ToString()
            a8(7) = Row.Item(5).ToString()
            Match_Key(a8, c01, "01")
            Match_Key(a8, c02, "02")
            Match_Key(a8, c03, "03")
            Match_Key(a8, c04, "04")
            Match_Key(a8, c05, "05")
            Match_Key(a8, c06, "06")
            Match_Key(a8, c07, "07")
            Match_Key(a8, c08, "08")
            Match_Key(a8, c09, "09")
            Match_Key(a8, c10, "10")
            Match_Key(a8, c11, "11")
            Match_Key(a8, c12, "12")

            'p3
            a8(0) = Row.Item(5).ToString()
            a8(1) = Row.Item(6).ToString()
            a8(2) = Row.Item(7).ToString()
            a8(3) = Row.Item(0).ToString()
            a8(4) = Row.Item(1).ToString()
            a8(5) = Row.Item(2).ToString()
            a8(6) = Row.Item(3).ToString()
            a8(7) = Row.Item(4).ToString()
            Match_Key(a8, c01, "01")
            Match_Key(a8, c02, "02")
            Match_Key(a8, c03, "03")
            Match_Key(a8, c04, "04")
            Match_Key(a8, c05, "05")
            Match_Key(a8, c06, "06")
            Match_Key(a8, c07, "07")
            Match_Key(a8, c08, "08")
            Match_Key(a8, c09, "09")
            Match_Key(a8, c10, "10")
            Match_Key(a8, c11, "11")
            Match_Key(a8, c12, "12")

            'p4
            a8(0) = Row.Item(4).ToString()
            a8(1) = Row.Item(5).ToString()
            a8(2) = Row.Item(6).ToString()
            a8(3) = Row.Item(7).ToString()
            a8(4) = Row.Item(0).ToString()
            a8(5) = Row.Item(1).ToString()
            a8(6) = Row.Item(2).ToString()
            a8(7) = Row.Item(3).ToString()
            Match_Key(a8, c01, "01")
            Match_Key(a8, c02, "02")
            Match_Key(a8, c03, "03")
            Match_Key(a8, c04, "04")
            Match_Key(a8, c05, "05")
            Match_Key(a8, c06, "06")
            Match_Key(a8, c07, "07")
            Match_Key(a8, c08, "08")
            Match_Key(a8, c09, "09")
            Match_Key(a8, c10, "10")
            Match_Key(a8, c11, "11")
            Match_Key(a8, c12, "12")

            'p5
            a8(0) = Row.Item(3).ToString()
            a8(1) = Row.Item(4).ToString()
            a8(2) = Row.Item(5).ToString()
            a8(3) = Row.Item(6).ToString()
            a8(4) = Row.Item(7).ToString()
            a8(5) = Row.Item(0).ToString()
            a8(6) = Row.Item(1).ToString()
            a8(7) = Row.Item(2).ToString()
            Match_Key(a8, c01, "01")
            Match_Key(a8, c02, "02")
            Match_Key(a8, c03, "03")
            Match_Key(a8, c04, "04")
            Match_Key(a8, c05, "05")
            Match_Key(a8, c06, "06")
            Match_Key(a8, c07, "07")
            Match_Key(a8, c08, "08")
            Match_Key(a8, c09, "09")
            Match_Key(a8, c10, "10")
            Match_Key(a8, c11, "11")
            Match_Key(a8, c12, "12")

            'p6
            a8(0) = Row.Item(2).ToString()
            a8(1) = Row.Item(3).ToString()
            a8(2) = Row.Item(4).ToString()
            a8(3) = Row.Item(5).ToString()
            a8(4) = Row.Item(6).ToString()
            a8(5) = Row.Item(7).ToString()
            a8(6) = Row.Item(0).ToString()
            a8(7) = Row.Item(1).ToString()
            Match_Key(a8, c01, "01")
            Match_Key(a8, c02, "02")
            Match_Key(a8, c03, "03")
            Match_Key(a8, c04, "04")
            Match_Key(a8, c05, "05")
            Match_Key(a8, c06, "06")
            Match_Key(a8, c07, "07")
            Match_Key(a8, c08, "08")
            Match_Key(a8, c09, "09")
            Match_Key(a8, c10, "10")
            Match_Key(a8, c11, "11")
            Match_Key(a8, c12, "12")

            'p7
            a8(0) = Row.Item(1).ToString()
            a8(1) = Row.Item(2).ToString()
            a8(2) = Row.Item(3).ToString()
            a8(3) = Row.Item(4).ToString()
            a8(4) = Row.Item(5).ToString()
            a8(5) = Row.Item(6).ToString()
            a8(6) = Row.Item(7).ToString()
            a8(7) = Row.Item(0).ToString()
            Match_Key(a8, c01, "01")
            Match_Key(a8, c02, "02")
            Match_Key(a8, c03, "03")
            Match_Key(a8, c04, "04")
            Match_Key(a8, c05, "05")
            Match_Key(a8, c06, "06")
            Match_Key(a8, c07, "07")
            Match_Key(a8, c08, "08")
            Match_Key(a8, c09, "09")
            Match_Key(a8, c10, "10")
            Match_Key(a8, c11, "11")
            Match_Key(a8, c12, "12")


        Next
        Console.WriteLine("row match stend :" + DateTime.Now.ToString())
    End Sub
    Sub Match_Key(ByVal combo() As String, ByVal Cuspp() As String, ByVal cusp_id As String)
        Dim fl1 As String = "N"
        Dim fl2 As String = "N"
        Dim fl3 As String = "N"
        Dim fl4 As String = "N"
        Dim fl5 As String = "N"
        Dim fl6 As String = "N"
        Dim fl7 As String = "N"
        Dim fl8 As String = "N"
        Dim m_str As String = ""
        Dim xcusp(Cuspp.Length - 1) As String
        Dim xcombo(combo.Length - 1) As String

        Dim f2rw_comb As String = ""
        For i As Integer = 0 To combo.Length - 1
            f2rw_comb = f2rw_comb + combo(i)
        Next

        Dim cusp_str As String = ""
        For i As Integer = 0 To Cuspp.Length - 1
            cusp_str = cusp_str + Cuspp(i)
        Next

        'Check first position
        Array.Copy(Cuspp, xcusp, Cuspp.Length)
        Array.Copy(combo, xcombo, Cuspp.Length)

        For i As Integer = 0 To xcusp.Length - 1
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
            For i As Integer = 0 To xcusp.Length - 1
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
            For i As Integer = 0 To xcusp.Length - 1
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
            For i As Integer = 0 To xcusp.Length - 1
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
            For i As Integer = 0 To xcusp.Length - 1
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
            For i As Integer = 0 To xcusp.Length - 1
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
            For i As Integer = 0 To xcusp.Length - 1
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
            For i As Integer = 0 To xcusp.Length - 1
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
        For i As Integer = 0 To xcusp.Length - 1
            If xcusp(i) IsNot "XX" Then
                pattern = False
            End If
        Next

        If pattern = True Then
            Dim uid As String = "XXXXXXXXXX"
            Dim hid As String = "100001"

            Dim con As New SqlConnection(connstr)
            Dim connection As SqlConnection = New SqlConnection(connstr)
            connection.Open()
            Dim query As String

            query = "INSERT INTO MATCH_FILE VALUES ('" + uid + "','" + hid + "','" + m_planet + "','" + f2prow + "','" + f2rw_comb + "','" + cusp_id + "','" + cusp_str + "','" + m_str + "')"
            Dim cmd As New SqlCommand(query, con)
            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet()
            da.Fill(ds)
            connection.Close()
        End If
    End Sub
End Module
