Attribute VB_Name = "Main"
Sub test1()
        Dim emp() As clsEmp
        ReDim emp(1 To 1)
        Set emp(1) = New clsEmp
        Dim entryCount As Long
        entryCount = Application.WorksheetFunction.CountA(Sheets("recordList").Range("C:C"))
        Dim rawData() As Variant
        rawData = Sheets("recordList").Range("A2:E" & entryCount).value
        Dim empList() As Long
        Sheets("Arrive").Activate
        empList = getEmpList(Sheets("Arrive"))
        Sheets("recordList").Activate
        For i = LBound(rawData) To UBound(rawData)
                If isWorker(CStr(rawData(i, 3)), empList) Then
                        Dim nme As String, uid As Long, dte As Long, tme As Double
                        nme = rawData(i, 1)
                        uid = CLng(Right(rawData(i, 3), 7))
                        dte = CLng(rawData(i, 4))
                        tme = rawData(i, 5)
                        If Not alreadyCreate(emp, uid, dte) Then
                                Call createEmp(emp, nme, uid, dte, tme)
                        ElseIf Not thisEmp(emp, uid, dte).have_near_time(tme) Then
                                thisEmp(emp, uid, dte).tme = tme
                        End If
                End If
        Next i
        Call get_Wage(emp)
        Call get_Arrive(emp)
        Call get_Leave(emp)
        Call completeTime(emp)
        Call cal_work_time(emp)
        Call printThem(emp)
End Sub
Sub printThem(emp() As clsEmp)
        Dim dl() As Long
        Dim el() As Long
        ReDim dl(1 To 1)
        ReDim el(1 To 1)
        Dim dchk As Boolean
        Dim echk As Boolean
        For i = LBound(emp) To UBound(emp)
                dchk = False
                echk = False
                For j = LBound(dl) To UBound(dl)
                        If emp(i).dte = dl(j) Then
                                dchk = True
                                j = UBound(dl)
                        End If
                Next j
                For k = LBound(el) To UBound(el)
                        If emp(i).uid = el(k) Then
                                echk = True
                                k = UBound(el)
                        End If
                Next k
                If dchk = False Then
                        ReDim Preserve dl(LBound(dl) To UBound(dl) + 1)
                        dl(UBound(dl) - 1) = emp(i).dte
                End If
                If echk = False Then
                        ReDim Preserve el(LBound(el) To UBound(el) + 1)
                        el(UBound(el) - 1) = emp(i).uid
                End If
        Next i
        ReDim Preserve dl(LBound(dl) To UBound(dl) - 1)
        ReDim Preserve el(LBound(el) To UBound(el) - 1)
        With Sheets("wageResult")
                .Select
                .UsedRange.Clear
                For i = LBound(dl) To UBound(dl)
                        .Cells(1, 2 + i).value = dl(i)
                Next i
                For j = LBound(el) To UBound(el)
                        .Cells(2 * j, 1).value = el(j)
                        .Cells(2 * j, 2).value = nameFromID(emp, el(j))
                Next j
                For i = LBound(dl) To UBound(dl)
                        For j = LBound(el) To UBound(el)
                                If alreadyCreate(emp, el(j), dl(i)) Then
                                        .Cells(2 * j, 2 + i).value = thisEmp(emp, el(j), dl(i)).normalWage
                                        .Cells(2 * j + 1, 2 + i).value = thisEmp(emp, el(j), dl(i)).overWage
                                End If
                        Next j
                Next i
        End With
        With Sheets("timeResult")
                .Select
                .UsedRange.Clear
                Row = 2
                For i = LBound(dl) To UBound(dl)
                        For j = LBound(el) To UBound(el)
                                If alreadyCreate(emp, el(j), dl(i)) Then
                                        Dim e As clsEmp
                                        Set e = thisEmp(emp, el(j), dl(i))
                                        .Cells(Row, 1).value = e.nme
                                        .Cells(Row, 2).value = e.uid
                                        .Cells(Row, 3).value = e.dte
                                        Dim t() As Double
                                        t = e.ttime
                                        .Range(Cells(Row, 4), Cells(Row, 3 + (UBound(t) - LBound(t)))).value = t
                                        Row = Row + 1
                                End If
                        Next j
                Next i
        End With
End Sub
Function nameFromID(emp() As clsEmp, uid As Long) As String
        For i = LBound(emp) To UBound(emp)
                If emp(i).uid = uid Then
                        nameFromID = emp(i).nme
                        i = UBound(emp)
                End If
        Next i
        If Left(nameFromID, 1) = "'" Then nameFromID = Right(nameFromID, Len(nameFromID) - 1)
End Function
Sub cal_work_time(emp() As clsEmp)
        For i = LBound(emp) To UBound(emp)
                emp(i).work_time
        Next i
End Sub
Sub get_Arrive(emp() As clsEmp)
        Dim sht As Object
        Set sht = Sheets("Arrive")
        sht.Select
        Dim el() As Long
        el = getEmpList(sht)
        Dim dl() As Long
        dl = getDateList(sht)
        For e = LBound(el) To UBound(el)
                For d = LBound(dl) To UBound(dl)
                        If alreadyCreate(emp, el(e), dl(d)) Then
                                thisEmp(emp, el(e), dl(d)).arrive = sht.Cells(2 + d, 1 + e).value
                        End If
                Next d
        Next e
End Sub
Sub get_Leave(emp() As clsEmp)
        Dim sht As Object
        Set sht = Sheets("Leave")
        sht.Select
        Dim el() As Long
        el = getEmpList(sht)
        Dim dl() As Long
        dl = getDateList(sht)
        For e = LBound(el) To UBound(el)
                For d = LBound(dl) To UBound(dl)
                        If alreadyCreate(emp, el(e), dl(d)) Then
                                thisEmp(emp, el(e), dl(d)).leave = sht.Cells(2 + d, 1 + e).value
                        End If
                Next d
        Next e
End Sub
Sub get_Wage(emp() As clsEmp)
        Dim sht As Object
        Set sht = Sheets("Wage")
        sht.Select
        Dim el() As Long
        el = getEmpList(sht)
        Dim dl() As Long
        dl = getDateList(sht)
        For e = LBound(el) To UBound(el)
                For d = LBound(dl) To UBound(dl)
                        If alreadyCreate(emp, el(e), dl(d)) Then
                                thisEmp(emp, el(e), dl(d)).wage = sht.Cells(2 + d, 1 + e).value
                        End If
                Next d
        Next e
End Sub
Function getEmpList(sht As Worksheet) As Long()
        Dim ans() As Long
        Dim tmp() As Variant
        ReDim ans(1 To 1)
        Dim i As Long
        i = Application.WorksheetFunction.CountA(sht.Range("2:2"))
        tmp = sht.Range(Cells(2, 2), Cells(2, 1 + i)).value
        For i = LBound(tmp, 2) To UBound(tmp, 2)
                ReDim Preserve ans(LBound(ans) To UBound(ans) + 1)
                ans(UBound(ans) - 1) = CLng(tmp(1, i))
        Next i
        ReDim Preserve ans(LBound(ans) To UBound(ans) - 1)
        getEmpList = ans
End Function
Function getDateList(sht As Worksheet) As Long()
        Dim ans() As Long
        Dim tmp() As Variant
        ReDim ans(1 To 1)
        Dim i As Long
        i = Application.WorksheetFunction.CountA(sht.Range("A:A"))
        tmp = sht.Range("A3:A" & (2 + i)).value
        For i = LBound(tmp) To UBound(tmp)
                ReDim Preserve ans(LBound(ans) To UBound(ans) + 1)
                ans(UBound(ans) - 1) = CLng(tmp(i, 1))
        Next i
        ReDim Preserve ans(LBound(ans) To UBound(ans) - 1)
        getDateList = ans
End Function
Sub completeTime(emp() As clsEmp)
        For i = LBound(emp) To UBound(emp)
                If emp(i).time_isnot_complete() Then
                        emp(i).fill_time
                End If
        Next i
End Sub
Function thisEmp(emp() As clsEmp, uid As Long, dte As Long) As clsEmp
        For i = LBound(emp) To UBound(emp)
                If emp(i).uid = uid And emp(i).dte = dte Then
                        Set thisEmp = emp(i)
                        i = UBound(emp)
                End If
        Next i
End Function
Sub createEmp(emp() As clsEmp, nme As String, uid As Long, dte As Long, tme As Double)
        ReDim Preserve emp(LBound(emp) To UBound(emp) + 1)
        Set emp(UBound(emp)) = New clsEmp
        emp(UBound(emp) - 1).nme = nme
        emp(UBound(emp) - 1).uid = uid
        emp(UBound(emp) - 1).dte = dte
        emp(UBound(emp) - 1).tme = tme
End Sub
Function alreadyCreate(emp() As clsEmp, uid As Long, dte As Long) As Boolean
        alreadyCreate = False
        For i = LBound(emp) To UBound(emp)
                If emp(i).uid = uid And emp(i).dte = dte Then
                        alreadyCreate = True
                        i = UBound(emp)
                End If
        Next i
End Function
Function isWorker(val As String, empList() As Long) As Boolean
        isWorker = False
        For i = LBound(empList) To UBound(empList)
                If CLng(Right(val, Len(val) - 1)) = empList(i) Then
                        isWorker = True
                        i = UBound(empList)
                End If
        Next i
End Function

Sub tt()
        Dim a() As Integer
        ReDim a(1 To 2)
        a(1) = 1
        a(2) = 2
        Cells(1, 1).value = a
End Sub
