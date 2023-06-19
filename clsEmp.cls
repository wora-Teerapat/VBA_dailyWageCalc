VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private lng_id As Long
Private lng_date As Long
Private str_name As String
Private dbl_time() As Double
'_____________________________'
Private dbl_wage As Double
Private dbl_arrive As Double
Private dbl_leave As Double
'____________________________'
Private normal_time As Double
Private over_time As Double
Private normal_wage As Double
Private over_wage As Double
Const DTIME As Double = 9 / 24 + 5 / 24 / 60  'Default time
Const MINUTE As Double = 1 / 24 / 60
Const HOUR As Double = 1 / 24
Const OTFACTOR As Double = 1.5

Public Sub class_test()
        If Me.time_isnot_complete() Then Me.fill_time
        For i = LBound(dbl_time) To UBound(dbl_time) - 1
                Debug.Print timeDbltoStr(dbl_time(i))
        Next i
End Sub
Public Sub work_time()
        Dim i As Integer
        i = LBound(dbl_time)
        Dim chk As Boolean
        chk = False
        Dim nt As Double, ot As Double
        nt = 0
        ot = 0
        If dbl_time(LBound(dbl_time)) < dbl_arrive Then dbl_time(LBound(dbl_time)) = dbl_arrive
        If dbl_time(LBound(dbl_time)) > dbl_leave Then chk = True
        While i < UBound(dbl_time) - 1 And chk = False
                If dbl_time(i) < dbl_time(i + 1) And dbl_time(i + 1) < dbl_leave Then
                        nt = nt + dbl_time(i + 1) - dbl_time(i)
                ElseIf dbl_time(i) < dbl_time(i + 1) And dbl_time(i + 1) > dbl_leave Then
                        nt = nt + dbl_leave - dbl_time(i)
                        ot = dbl_time(i + 1) - dbl_leave
                        chk = True
                Else
                        ot = ot + dbl_time(i + 1) - dbl_time(i)
                        chk = True
                End If
                i = i + 2
        Wend
        While i < UBound(dbl_time)
                ot = dbl_time(i + 1) - dbl_time(i)
                i = i + 2
        Wend
        If nt <= 8 * HOUR Then
                normal_time = nt
        Else
                normal_time = 8 * HOUR
                ot = ot + nt - 8 * HOUR
        End If
        over_time = ot
        normal_wage = normal_time * dbl_wage
        over_wage = over_time * dbl_wage * OTFACTOR
        
End Sub
Public Function have_near_time(var As Double) As Boolean
        have_near_time = False
        If Abs(dbl_time(UBound(dbl_time) - 1) - var) < 5 * MINUTE Then have_near_time = True
End Function
Public Function time_isnot_complete() As Boolean
        Dim chk As Boolean
        chk = True
        If UBound(dbl_time) Mod 2 = 1 Then chk = False
        time_isnot_complete = chk
End Function
Public Sub fill_time()
        Dim bdy As String
        bdy = "Name: " & str_name & vbNewLine & " ID: " & lng_id & ", Date: " & CDate(lng_date) & vbNewLine
        Dim i As Integer
        i = 1
        While i < UBound(dbl_time)
                bdy = bdy & i & " -> " & timeDbltoStr(dbl_time(i)) & vbNewLine
                i = i + 1
        Wend
        Call extend_time
        Dim a As String
        bdy = bdy & "Enter time entry number " & i & " in HH:MM format"
        Dim dm As String 'default time for next inputbox
        dm = timeDbltoStr(dbl_leave)
        a = InputBox(bdy, , timeDbltoStr(dbl_leave))
        dbl_time(UBound(dbl_time) - 1) = timeStrtoDbl(a)
End Sub
Private Function timeDbltoStr(val As Double) As String
        Dim ans As String
        Dim hour_val As Double
        hour_val = Application.WorksheetFunction.RoundDown(24 * val, 0)
        Dim min_val As Double
        min_val = (24 * val - hour_val) * 60
        min_val = Application.WorksheetFunction.RoundDown(min_val, 0)
        ans = "" & Right("00" & hour_val, 2)
        ans = ans & ":" & Right("00" & min_val, 2)
        timeDbltoStr = ans

End Function
Private Function timeStrtoDbl(val As String) As Double
        Dim ans As Double
        ans = CDbl(Left(val, 2)) / 24
        ans = ans + CDbl(Right(val, 2)) / 24 / 60
        timeStrtoDbl = ans
End Function
Private Sub extend_time()
        ReDim Preserve dbl_time(LBound(dbl_time) To UBound(dbl_time) + 1)
        dbl_time(UBound(dbl_time)) = DTIME
End Sub
Private Sub Class_Initialize()
        ReDim dbl_time(1 To 1)
        dbl_time(1) = DTIME
End Sub
Public Property Let uid(val As Long)
        lng_id = val
End Property
Public Property Get uid() As Long
        uid = lng_id
End Property
Public Property Let dte(val As Long)
        lng_date = val
End Property
Public Property Get dte() As Long
        dte = lng_date
End Property
Public Property Let nme(val As String)
        str_name = val
End Property
Public Property Get nme() As String
        nme = str_name
End Property
Public Property Let tme(val As Double)
        Call extend_time
        dbl_time(UBound(dbl_time) - 1) = val
End Property
Public Property Get ttime() As Double()
        Dim ans() As Double
        ans = dbl_time
        'ReDim Preserve ans(LBound(ans) To UBound(ans) - 1)
        ttime = ans
End Property
Public Property Let wage(val As Double)
        dbl_wage = val
End Property
Public Property Let arrive(val As Double)
        dbl_arrive = val
End Property
Public Property Let leave(val As Double)
        dbl_leave = val
End Property
Public Property Get normalTime()
        normalTime = normal_time
End Property
Public Property Get overTime()
        overTime = over_time
End Property
Public Property Get normalWage()
        normalWage = normal_wage
End Property
Public Property Get overWage()
        overWage = over_wage
End Property
