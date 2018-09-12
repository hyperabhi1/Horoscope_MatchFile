Imports System.Data.SqlClient

Module Chart_Gen1
    Dim SelectHCUSP = "SELECT * FROM TABLE_HCUSP"
    Dim SelectHPLANET = "SELECT * FROM TABLE_HPLANET"
    Dim connstr = "data source=49.50.103.132;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=False;User Id=sa;password=pSI)TA1t0K[)"

    Sub Main()
        Process_HCUSP()
        Process_HPLANET()
    End Sub

    Sub Process_HCUSP()
        Dim con As New SqlConnection(connstr)
        Dim connection As SqlConnection = New SqlConnection(connstr)
        connection.Open()
        Dim cmd As New SqlCommand(SelectHCUSP, con)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        da.Fill(ds)
        con.Close()
        con.ConnectionString = connstr
        con.Open()
        cmd.Connection = con
        Try
            For Each Row As DataRow In ds.Tables(0).Rows
                Dim M_uid As String = Row.Item(0).ToString()
                Dim M_hid As String = Row.Item(1).ToString()
                Dim M_cusp As String = Row.Item(2).ToString()
                Dim M_sign As String = Row.Item(3).ToString()
                Dim M_DMS As String = Row.Item(4).ToString()
                Dim M_cp1 As String = Row.Item(5).ToString()
                Dim M_cp2 As String = Row.Item(6).ToString()
                Dim M_cp3 As String = Row.Item(7).ToString()
                Dim M_order As String = getorder(Row.Item(3).ToString())
                Dim M_roman As String = getroman(Row.Item(2).ToString())
                cmd.CommandText = "INSERT INTO HCHART VALUES ('" + M_uid + "','" + M_hid + "','" + M_sign + "','" + M_order + "','" + M_DMS + "','" + M_cusp + "','??','" + M_roman + "','?','????????????????????????????????????????')"
                cmd.ExecuteNonQuery()
            Next
        Catch ex As Exception
            Console.WriteLine("Error while inserting record on table..." & ex.Message, "Insert Records")
        Finally
            con.Close()
        End Try
    End Sub
    Sub Process_HPLANET()
        Dim con As New SqlConnection(connstr)
        Dim connection As SqlConnection = New SqlConnection(connstr)
        connection.Open()
        Dim cmd As New SqlCommand(SelectHPLANET, con)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        da.Fill(ds)
        con.Close()
        con.ConnectionString = connstr
        con.Open()
        cmd.Connection = con
        Try
            For Each Row As DataRow In ds.Tables(0).Rows
                Dim m_uid As String = Row.Item(0).ToString()
                Dim m_hid As String = Row.Item(1).ToString()
                Dim m_planet As String = Row.Item(2).ToString()
                Dim m_sign As String = Row.Item(3).ToString()
                Dim m_DMS As String = Row.Item(4).ToString()
                Dim m_cuspid As String = Row.Item(10).ToString()
                Dim m_order As String = getorder(Row.Item(3).ToString())
                cmd.CommandText = "INSERT INTO HCHART VALUES ('" + m_uid + "','" + m_hid + "','" + m_sign + "','" + m_order + "','" + m_DMS + "','" + m_cuspid + "','" + m_planet + "','????','Y','????????????????????????????????????????')"
                cmd.ExecuteNonQuery()
            Next
        Catch ex As Exception
            Console.WriteLine("Error while inserting record on table..." & ex.Message, "Insert Records")
        Finally
            con.Close()
        End Try
    End Sub
    Function getorder(ByVal s As String)
        Dim ov = "00"
        Select Case s
            Case "ARI"
                ov = "01"
            Case "TAU"
                ov = "02"
            Case "GEM"
                ov = "03"
            Case "CAN"
                ov = "04"
            Case "LEO"
                ov = "05"
            Case "VIR"
                ov = "06"
            Case "LIB"
                ov = "07"
            Case "SCO"
                ov = "08"
            Case "SAG"
                ov = "09"
            Case "CAP"
                ov = "10"
            Case "AQU"
                ov = "11"
            Case "PIS"
                ov = "12"
        End Select
        Return ov
    End Function
    Function getroman(ByVal s As String)
        Dim ov = "XX"
        Select Case s
            Case "01"
                ov = "I"
            Case "02"
                ov = "II"
            Case "03"
                ov = "III"
            Case "04"
                ov = "IV"
            Case "05"
                ov = "V"
            Case "06"
                ov = "VI"
            Case "07"
                ov = "VII"
            Case "08"
                ov = "VIII"
            Case "09"
                ov = "IX"
            Case "10"
                ov = "X"
            Case "11"
                ov = "XI"
            Case "12"
                ov = "XII"
            Case Else
        End Select
        Return ov
    End Function
End Module
