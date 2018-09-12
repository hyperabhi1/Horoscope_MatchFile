Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Linq
Imports System.Runtime.CompilerServices
Module Permutations
    Function ForAllPermutation(Of T)(ByVal items As T(), ByVal funcExecuteAndTellIfShouldStop As Func(Of T(), Boolean)) As Boolean
        Dim countOfItem As Integer = items.Length

        If countOfItem <= 1 Then
            Return funcExecuteAndTellIfShouldStop(items)
        End If

        Dim indexes = New Integer(countOfItem - 1) {}

        For i As Integer = 0 To countOfItem - 1
            indexes(i) = 0
        Next

        If funcExecuteAndTellIfShouldStop(items) Then
            Return True
        End If

        Dim ii As Integer = 1

        While ii < countOfItem

            If indexes(ii) < ii Then

                If (ii And 1) = 1 Then
                    Swap(items(ii), items(indexes(ii)))
                Else
                    Swap(items(ii), items(0))
                End If

                If funcExecuteAndTellIfShouldStop(items) Then
                    Return True
                End If

                indexes(ii) += 1
                ii = 1
            Else
                indexes(Math.Min(System.Threading.Interlocked.Increment(ii), ii - 1)) = 0
            End If
        End While

        Return False
    End Function

    Private Function GetPermutations(Of T)(ByVal list As IEnumerable(Of T), ByVal length As Integer) As IEnumerable(Of IEnumerable(Of T))
        If length = 1 Then Return list.[Select](Function(ttt) New T() {ttt})
        Return GetPermutations(list, length - 1).SelectMany(Function(tt) list.Where(Function(e) Not tt.Contains(e)), Function(t1, t2) t1.Concat(New T() {t2}))
    End Function

    <MethodImpl(MethodImplOptions.AggressiveInlining)>
    Private Sub Swap(Of T)(ByRef a As T, ByRef b As T)
        Dim temp As T = a
        a = b
        b = temp
    End Sub

    Sub Test()
        ForAllPermutation("123".ToCharArray(), Function(vals)
                                                   Console.WriteLine(String.Join("", vals))
                                                   Return False
                                               End Function)
        Dim values As Integer() = New Integer() {0, 1, 2, 4}
        Console.WriteLine("Ouellet heap's algorithm implementation")
        ForAllPermutation(values, Function(vals)
                                      Console.WriteLine(String.Join("", vals))
                                      Return False
                                  End Function)
        Console.WriteLine("Linq algorithm")

        For Each v In GetPermutations(values, values.Length)
            Console.WriteLine(String.Join("", v))
        Next

        Dim count As Integer = 0
        values = New Integer(9) {}

        For n As Integer = 0 To values.Length - 1
            values(n) = n
        Next

        Dim stopWatch As Stopwatch = New Stopwatch()
        ForAllPermutation(values, Function(vals)

                                      For Each v In vals
                                          count += 1
                                      Next

                                      Return False
                                  End Function)
        stopWatch.[Stop]()
        Console.WriteLine($"Ouellet heap's algorithm implementation {count} items in {stopWatch.ElapsedMilliseconds} millisecs")
        count = 0
        stopWatch.Reset()
        stopWatch.Start()

        For Each vals In GetPermutations(values, values.Length)

            For Each v In vals
                count += 1
            Next
        Next

        stopWatch.[Stop]()
        Console.WriteLine($"Linq {count} items in {stopWatch.ElapsedMilliseconds} millisecs")
    End Sub
End Module
