Imports System.Security.Cryptography
Imports System.Data.SqlClient
Imports System.Text

Module Match_file_modified
    Dim S1(12) As String
    Dim c01(), c02(), c03(), c04(), c05(), c06(), c07(), c08(), c09(), c10(), c11(), c12() As String
    Dim m_planet As String
    Dim f2prow As String
    Dim connstr = "data source=DESKTOP-JBRFH9E;initial catalog=testdb;integrated security=True;"
    Dim SelectF2PLANET = "truncate table match_file1;SELECT top 100000 * FROM F2PLANETS;"
    Dim F2P_RwCsp01, F2P_RwCsp02, F2P_RwCsp03, F2P_RwCsp04, F2P_RwCsp05, F2P_RwCsp06, F2P_RwCsp07, F2P_RwCsp08, F2P_RwCsp09, F2P_RwCsp10, F2P_RwCsp11, F2P_RwCsp12 As String

    Sub Main()
        Dim str As String = "kirti.vashishtha1994@gmail.com"
        Console.WriteLine(Encrypt(str, "sblw-3hn8-sqoy19"))
        Dim st As String = Encrypt("mP0Bnz8Ii/spHwdPZuHJYFv3Wvbz3LGPBi+W1mfmjIw=", "sblw-3hn8-sqoy19")
        Console.WriteLine(Decrypt("mP0Bnz8Ii/spHwdPZuHJYFv3Wvbz3LGPBi+W1mfmjIw=", "sblw-3hn8-sqoy19"))
        Dim starttime As DateTime = DateTime.Now
        'EmailNotify.SendEmail("Match_File_Local Started", starttime, "start")
        Console.WriteLine(starttime)
        MakeCusp_checkHouse()
        'EmailNotify.SendEmail("Match_File_Local Ended", starttime, "end")
        Console.WriteLine(DateTime.Now.Subtract(starttime))
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
        '        Dim SelectF2PLANET = "SELECT * FROM F2PLANETS WHERE p0 = 'SU'"
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

            F2P_RwCsp01 = MakeStr4Cusps(c01, Row)
            F2P_RwCsp02 = MakeStr4Cusps(c02, Row)
            F2P_RwCsp03 = MakeStr4Cusps(c03, Row)
            F2P_RwCsp04 = MakeStr4Cusps(c04, Row)
            F2P_RwCsp05 = MakeStr4Cusps(c05, Row)
            F2P_RwCsp06 = MakeStr4Cusps(c06, Row)
            F2P_RwCsp07 = MakeStr4Cusps(c07, Row)
            F2P_RwCsp08 = MakeStr4Cusps(c08, Row)
            F2P_RwCsp09 = MakeStr4Cusps(c09, Row)
            F2P_RwCsp10 = MakeStr4Cusps(c10, Row)
            F2P_RwCsp11 = MakeStr4Cusps(c11, Row)
            F2P_RwCsp12 = MakeStr4Cusps(c12, Row)

            Permutations.ForAllPermutation(c01, Function(vals)
                                                    'Console.WriteLine(String.Join("", vals))
                                                    If F2P_RwCsp01.Contains(String.Join("", c01)) Then
                                                        InsertMATCHFILE(F2P_RwCsp01, "01", String.Join("", c01))
                                                    End If
                                                    Return False
                                                End Function)

            Permutations.ForAllPermutation(c02, Function(vals)
                                                    'Console.WriteLine(String.Join("", vals))
                                                    If F2P_RwCsp02.Contains(String.Join("", c02)) Then
                                                        InsertMATCHFILE(F2P_RwCsp02, "02", String.Join("", c02))
                                                    End If
                                                    Return False
                                                End Function)

            Permutations.ForAllPermutation(c03, Function(vals)
                                                    'Console.WriteLine(String.Join("", vals))
                                                    If F2P_RwCsp03.Contains(String.Join("", c03)) Then
                                                        InsertMATCHFILE(F2P_RwCsp03, "03", String.Join("", c03))
                                                    End If
                                                    Return False
                                                End Function)

            Permutations.ForAllPermutation(c04, Function(vals)
                                                    'Console.WriteLine(String.Join("", vals))
                                                    If F2P_RwCsp04.Contains(String.Join("", c04)) Then
                                                        InsertMATCHFILE(F2P_RwCsp04, "04", String.Join("", c04))
                                                    End If
                                                    Return False
                                                End Function)

            Permutations.ForAllPermutation(c05, Function(vals)
                                                    'Console.WriteLine(String.Join("", vals))
                                                    If F2P_RwCsp05.Contains(String.Join("", c05)) Then
                                                        InsertMATCHFILE(F2P_RwCsp05, "05", String.Join("", c05))
                                                    End If
                                                    Return False
                                                End Function)

            Permutations.ForAllPermutation(c06, Function(vals)
                                                    'Console.WriteLine(String.Join("", vals))
                                                    If F2P_RwCsp06.Contains(String.Join("", c06)) Then
                                                        InsertMATCHFILE(F2P_RwCsp06, "06", String.Join("", c06))
                                                    End If
                                                    Return False
                                                End Function)

            Permutations.ForAllPermutation(c07, Function(vals)
                                                    'Console.WriteLine(String.Join("", vals))
                                                    If F2P_RwCsp07.Contains(String.Join("", c07)) Then
                                                        InsertMATCHFILE(F2P_RwCsp07, "07", String.Join("", c07))
                                                    End If
                                                    Return False
                                                End Function)

            Permutations.ForAllPermutation(c08, Function(vals)
                                                    'Console.WriteLine(String.Join("", vals))
                                                    If F2P_RwCsp08.Contains(String.Join("", c08)) Then
                                                        InsertMATCHFILE(F2P_RwCsp08, "08", String.Join("", c08))
                                                    End If
                                                    Return False
                                                End Function)

            Permutations.ForAllPermutation(c09, Function(vals)
                                                    'Console.WriteLine(String.Join("", vals))
                                                    If F2P_RwCsp09.Contains(String.Join("", c09)) Then
                                                        InsertMATCHFILE(F2P_RwCsp09, "09", String.Join("", c09))
                                                    End If
                                                    Return False
                                                End Function)

            Permutations.ForAllPermutation(c10, Function(vals)
                                                    'Console.WriteLine(String.Join("", vals))
                                                    If F2P_RwCsp10.Contains(String.Join("", c10)) Then
                                                        InsertMATCHFILE(F2P_RwCsp10, "10", String.Join("", c10))
                                                    End If
                                                    Return False
                                                End Function)

            Permutations.ForAllPermutation(c11, Function(vals)
                                                    'Console.WriteLine(String.Join("", vals))
                                                    If F2P_RwCsp11.Contains(String.Join("", c11)) Then
                                                        InsertMATCHFILE(F2P_RwCsp11, "11", String.Join("", c11))
                                                    End If
                                                    Return False
                                                End Function)

            Permutations.ForAllPermutation(c12, Function(vals)
                                                    'Console.WriteLine(String.Join("", vals))
                                                    If F2P_RwCsp12.Contains(String.Join("", c12)) Then
                                                        InsertMATCHFILE(F2P_RwCsp12, "12", String.Join("", c12))
                                                    End If
                                                    Return False
                                                End Function)

            'While (NextPermutation(c01))
            '    Console.WriteLine(String.Join("", c01))
            'End While


            'If F2P_RwCsp01.Contains(String.Join("", c01)) Then
            '    InsertMATCHFILE(F2P_RwCsp01, "01", String.Join("", c01))
            'End If
            'If F2P_RwCsp02.Contains(String.Join("", c02)) Then
            '    InsertMATCHFILE(F2P_RwCsp02, "02", String.Join("", c02))
            'End If
            'If F2P_RwCsp03.Contains(String.Join("", c03)) Then
            '    InsertMATCHFILE(F2P_RwCsp03, "03", String.Join("", c03))
            'End If
            'If F2P_RwCsp04.Contains(String.Join("", c04)) Then
            '    InsertMATCHFILE(F2P_RwCsp04, "04", String.Join("", c04))
            'End If
            'If F2P_RwCsp05.Contains(String.Join("", c05)) Then
            '    InsertMATCHFILE(F2P_RwCsp05, "05", String.Join("", c05))
            'End If
            'If F2P_RwCsp06.Contains(String.Join("", c06)) Then
            '    InsertMATCHFILE(F2P_RwCsp06, "06", String.Join("", c06))
            'End If
            'If F2P_RwCsp07.Contains(String.Join("", c07)) Then
            '    InsertMATCHFILE(F2P_RwCsp07, "07", String.Join("", c07))
            'End If
            'If F2P_RwCsp08.Contains(String.Join("", c08)) Then
            '    InsertMATCHFILE(F2P_RwCsp08, "08", String.Join("", c08))
            'End If
            'If F2P_RwCsp09.Contains(String.Join("", c09)) Then
            '    InsertMATCHFILE(F2P_RwCsp09, "09", String.Join("", c09))
            'End If
            'If F2P_RwCsp10.Contains(String.Join("", c10)) Then
            '    InsertMATCHFILE(F2P_RwCsp10, "10", String.Join("", c10))
            'End If
            'If F2P_RwCsp11.Contains(String.Join("", c11)) Then
            '    InsertMATCHFILE(F2P_RwCsp11, "11", String.Join("", c11))
            'End If
            'If F2P_RwCsp12.Contains(String.Join("", c12)) Then
            '    InsertMATCHFILE(F2P_RwCsp12, "12", String.Join("", c12))
            'End If





            'f2prow = Row.Item(0).ToString() + Row.Item(1).ToString() + Row.Item(2).ToString() + Row.Item(3).ToString()
            'f2prow = f2prow + Row.Item(4).ToString() + Row.Item(5).ToString() + Row.Item(6).ToString() + Row.Item(7).ToString() + Row.Item(0).ToString()
            'f2prow = f2prow + Row.Item(0).ToString() + Row.Item(1).ToString() + Row.Item(2).ToString() + Row.Item(3).ToString()
            'f2prow = f2prow + Row.Item(4).ToString() + Row.Item(5).ToString() + Row.Item(6).ToString()


            'Dim c01_str As String = String.Join("", c01).Replace("ME", "m").Replace("SU", "s").Replace("KE", "k").Replace("SA", "a").Replace("MA", "r").Replace("VE", "v").Replace("RA", "r").Replace("MO", "o").Replace("JU", "j")
            'Dim c02_str As String = String.Join("", c02).Replace("ME", "m").Replace("SU", "s").Replace("KE", "k").Replace("SA", "a").Replace("MA", "r").Replace("VE", "v").Replace("RA", "r").Replace("MO", "o").Replace("JU", "j")
            'Dim c03_str As String = String.Join("", c03).Replace("ME", "m").Replace("SU", "s").Replace("KE", "k").Replace("SA", "a").Replace("MA", "r").Replace("VE", "v").Replace("RA", "r").Replace("MO", "o").Replace("JU", "j")
            'Dim c04_str As String = String.Join("", c04).Replace("ME", "m").Replace("SU", "s").Replace("KE", "k").Replace("SA", "a").Replace("MA", "r").Replace("VE", "v").Replace("RA", "r").Replace("MO", "o").Replace("JU", "j")
            'Dim c05_str As String = String.Join("", c05).Replace("ME", "m").Replace("SU", "s").Replace("KE", "k").Replace("SA", "a").Replace("MA", "r").Replace("VE", "v").Replace("RA", "r").Replace("MO", "o").Replace("JU", "j")
            'Dim c06_str As String = String.Join("", c06).Replace("ME", "m").Replace("SU", "s").Replace("KE", "k").Replace("SA", "a").Replace("MA", "r").Replace("VE", "v").Replace("RA", "r").Replace("MO", "o").Replace("JU", "j")
            'Dim c07_str As String = String.Join("", c07).Replace("ME", "m").Replace("SU", "s").Replace("KE", "k").Replace("SA", "a").Replace("MA", "r").Replace("VE", "v").Replace("RA", "r").Replace("MO", "o").Replace("JU", "j")
            'Dim c08_str As String = String.Join("", c08).Replace("ME", "m").Replace("SU", "s").Replace("KE", "k").Replace("SA", "a").Replace("MA", "r").Replace("VE", "v").Replace("RA", "r").Replace("MO", "o").Replace("JU", "j")
            'Dim c09_str As String = String.Join("", c09).Replace("ME", "m").Replace("SU", "s").Replace("KE", "k").Replace("SA", "a").Replace("MA", "r").Replace("VE", "v").Replace("RA", "r").Replace("MO", "o").Replace("JU", "j")
            'Dim c10_str As String = String.Join("", c10).Replace("ME", "m").Replace("SU", "s").Replace("KE", "k").Replace("SA", "a").Replace("MA", "r").Replace("VE", "v").Replace("RA", "r").Replace("MO", "o").Replace("JU", "j")
            'Dim c11_str As String = String.Join("", c11).Replace("ME", "m").Replace("SU", "s").Replace("KE", "k").Replace("SA", "a").Replace("MA", "r").Replace("VE", "v").Replace("RA", "r").Replace("MO", "o").Replace("JU", "j")
            'Dim c12_str As String = String.Join("", c12).Replace("ME", "m").Replace("SU", "s").Replace("KE", "k").Replace("SA", "a").Replace("MA", "r").Replace("VE", "v").Replace("RA", "r").Replace("MO", "o").Replace("JU", "j")



            'If f2prow.Contains(String.Join("", c01)) Then
            '    InsertMATCHFILE(f2prow, "01", String.Join("", c01))
            'End If
            'If f2prow.Contains(String.Join("", c02)) Then
            '    InsertMATCHFILE(f2prow, "02", String.Join("", c02))
            'End If
            'If f2prow.Contains(String.Join("", c03)) Then
            '    InsertMATCHFILE(f2prow, "03", String.Join("", c03))
            'End If
            'If f2prow.Contains(String.Join("", c04)) Then
            '    InsertMATCHFILE(f2prow, "04", String.Join("", c04))
            'End If
            'If f2prow.Contains(String.Join("", c05)) Then
            '    InsertMATCHFILE(f2prow, "05", String.Join("", c05))
            'End If
            'If f2prow.Contains(String.Join("", c06)) Then
            '    InsertMATCHFILE(f2prow, "06", String.Join("", c06))
            'End If
            'If f2prow.Contains(String.Join("", c07)) Then
            '    InsertMATCHFILE(f2prow, "07", String.Join("", c07))
            'End If
            'If f2prow.Contains(String.Join("", c08)) Then
            '    InsertMATCHFILE(f2prow, "08", String.Join("", c08))
            'End If
            'If f2prow.Contains(String.Join("", c09)) Then
            '    InsertMATCHFILE(f2prow, "09", String.Join("", c09))
            'End If
            'If f2prow.Contains(String.Join("", c10)) Then
            '    InsertMATCHFILE(f2prow, "10", String.Join("", c10))
            'End If
            'If f2prow.Contains(String.Join("", c11)) Then
            '    InsertMATCHFILE(f2prow, "11", String.Join("", c11))
            'End If
            'If f2prow.Contains(String.Join("", c12)) Then
            '    InsertMATCHFILE(f2prow, "12", String.Join("", c12))
            'End If

        Next
        Console.WriteLine("row match stend :" + DateTime.Now.ToString())
    End Sub

    Private Function MakeStr4Cusps(Cusp() As String, Row As DataRow) As String
        Dim str As String = ""
        If Cusp.Length = 3 Then
            str = Row.Item(0).ToString() + Row.Item(1).ToString() + Row.Item(2).ToString() + Row.Item(3).ToString() + Row.Item(4).ToString() + Row.Item(5).ToString()
            str = str + Row.Item(6).ToString() + Row.Item(7).ToString() + Row.Item(0).ToString() + Row.Item(1).ToString()
        End If
        If Cusp.Length = 4 Then
            str = Row.Item(0).ToString() + Row.Item(1).ToString() + Row.Item(2).ToString() + Row.Item(3).ToString() + Row.Item(4).ToString() + Row.Item(5).ToString()
            str = str + Row.Item(6).ToString() + Row.Item(7).ToString() + Row.Item(0).ToString() + Row.Item(1).ToString() + Row.Item(2).ToString()
        End If
        If Cusp.Length = 5 Then
            str = Row.Item(0).ToString() + Row.Item(1).ToString() + Row.Item(2).ToString() + Row.Item(3).ToString() + Row.Item(4).ToString() + Row.Item(5).ToString()
            str = str + Row.Item(6).ToString() + Row.Item(7).ToString() + Row.Item(0).ToString() + Row.Item(1).ToString() + Row.Item(2).ToString() + Row.Item(3).ToString()
        End If
        If Cusp.Length = 6 Then
            str = Row.Item(0).ToString() + Row.Item(1).ToString() + Row.Item(2).ToString() + Row.Item(3).ToString() + Row.Item(4).ToString() + Row.Item(5).ToString()
            str = str + Row.Item(6).ToString() + Row.Item(7).ToString() + Row.Item(0).ToString() + Row.Item(1).ToString() + Row.Item(2).ToString() + Row.Item(3).ToString()
            str = str + Row.Item(4).ToString()
        End If
        If Cusp.Length = 7 Then
            str = Row.Item(0).ToString() + Row.Item(1).ToString() + Row.Item(2).ToString() + Row.Item(3).ToString() + Row.Item(4).ToString() + Row.Item(5).ToString()
            str = str + Row.Item(6).ToString() + Row.Item(7).ToString() + Row.Item(0).ToString() + Row.Item(1).ToString() + Row.Item(2).ToString() + Row.Item(3).ToString()
            str = str + Row.Item(4).ToString() + Row.Item(5).ToString()
        End If
        If Cusp.Length = 8 Then
            str = Row.Item(0).ToString() + Row.Item(1).ToString() + Row.Item(2).ToString() + Row.Item(3).ToString() + Row.Item(4).ToString() + Row.Item(5).ToString()
            str = str + Row.Item(6).ToString() + Row.Item(7).ToString() + Row.Item(0).ToString() + Row.Item(1).ToString() + Row.Item(2).ToString() + Row.Item(3).ToString()
            str = str + Row.Item(4).ToString() + Row.Item(5).ToString() + Row.Item(6).ToString()
        End If
        Return str
    End Function

    Sub InsertMATCHFILE(ByVal f2prow As String, ByVal CuspId As String, ByVal Cusp_str As String)
        Dim uid As String = "XXXXXXXXXX"
        Dim hid As String = "100001"

        Dim con As New SqlConnection(connstr)
        Dim connection As SqlConnection = New SqlConnection(connstr)
        connection.Open()
        Dim query As String

        query = "INSERT INTO MATCH_FILE1 VALUES ('" + uid + "','" + hid + "','" + "mp" + "','" + f2prow.Substring(0, 16) + "','XXXXXXXXXXXXXXXX','" + CuspId + "','" + Cusp_str + "','XXXXXX')"
        Dim cmd As New SqlCommand(query, con)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        da.Fill(ds)
        connection.Close()
    End Sub


    Function NextPermutation(Of T As IComparable(Of T))(ByVal elements As T()) As Boolean
        Dim count = elements.Length
        Dim done = True

        For i = count - 1 To 0 + 1
            Dim curr = elements(i)

            If curr.CompareTo(elements(i - 1)) < 0 Then
                Continue For
            End If

            done = False
            Dim prev = elements(i - 1)
            Dim currIndex = i
            Dim j = count - 1
            For j = i + 1 To count - 1
                Dim tmp = elements(j)

                If tmp.CompareTo(curr) < 0 AndAlso tmp.CompareTo(prev) > 0 Then
                    curr = tmp
                    currIndex = j
                End If
            Next

            elements(currIndex) = prev
            elements(i - 1) = curr


            While j > i
                Dim tmp = elements(j)
                elements(j) = elements(i)
                elements(i) = tmp
                j -= 1
                i += 1
            End While

            Exit For
        Next

        Return done
    End Function
    Public Function Encrypt(ByVal input As String, ByVal key As String) As String
        Dim inputArray As Byte() = UTF8Encoding.UTF8.GetBytes(input)
        Dim tripleDES As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider()
        tripleDES.Key = UTF8Encoding.UTF8.GetBytes(key)
        tripleDES.Mode = CipherMode.ECB
        tripleDES.Padding = PaddingMode.PKCS7
        Dim cTransform As ICryptoTransform = tripleDES.CreateEncryptor()
        Dim resultArray As Byte() = cTransform.TransformFinalBlock(inputArray, 0, inputArray.Length)
        tripleDES.Clear()
        Return Convert.ToBase64String(resultArray, 0, resultArray.Length)
    End Function

    Public Function Decrypt(ByVal input As String, ByVal key As String) As String
        Dim inputArray As Byte() = Convert.FromBase64String(input)
        Dim tripleDES As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider()
        tripleDES.Key = UTF8Encoding.UTF8.GetBytes(key)
        tripleDES.Mode = CipherMode.ECB
        tripleDES.Padding = PaddingMode.PKCS7
        Dim cTransform As ICryptoTransform = tripleDES.CreateDecryptor()
        Dim resultArray As Byte() = cTransform.TransformFinalBlock(inputArray, 0, inputArray.Length)
        tripleDES.Clear()
        Return UTF8Encoding.UTF8.GetString(resultArray)
    End Function
End Module