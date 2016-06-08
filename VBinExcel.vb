Function sqlcall() As String
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection

    Dim connectionString As String
    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\**\Desktop\Dairy.accdb;Persist Security Info=False;"
    conn.Open connectionString
    
    Dim c As String
    c = ActiveCell.Value
        
    Dim strSql As String
     
    strSql = "SELECT * FROM [Slot Assignments] "
    strSql = strSql & "INNER JOIN Slot ON Slot.SlotId = [Slot Assignments].SlotId AND Slot.Warehouse = [Slot Assignments].Warehouse "
    strSql = strSql & "WHERE Slot.Warehouse = 4 AND [Slot Assignments].Warehouse = 4 AND Slot.OpenClosed = 'O' AND Slot.SlotType = 1 AND [Slot Assignments].SlotId = " & "'" & c & "'"
    


    Dim RS As ADODB.Recordset
    Dim str As Variant
    Dim out As String
    Dim s As String
    
    ' Open up a recordset / run query
    Set RS = New ADODB.Recordset
    RS.Open strSql, conn, adOpenStatic, adLockReadOnly, adCmdText

    If Not RS.EOF Then s = RS.GetString(, , ",", ",")
    Dim strArray() As String
    strArray = Split(s, ",")
    
    If UBound(strArray) > 1 Then
    s = strArray(2)
    Selection.Offset(-4, 0).Select
    ActiveCell.Value = s
    Else
    MsgBox ("Item not In Database")
    End If
    
    
    conn.Close
    sqlcall = s
    End Function

Sub Update()
Dim cell(0 To 96) As String
    'row 7
    cell(0) = "X7"
    cell(1) = "Z7"
    cell(2) = "AB7"
    cell(3) = "AD7"
    cell(4) = "AF7"
    cell(5) = "AH7"
    cell(6) = "AJ7"
    cell(7) = "AL7"
    cell(8) = "AN7"
    cell(9) = "AP7"
    cell(10) = "AR7"
    cell(11) = "AT7"
    'row 14
    cell(12) = "X14"
    cell(13) = "Z14"
    cell(14) = "AB14"
    cell(15) = "AD14"
    cell(16) = "AF14"
    cell(17) = "AH14"
    cell(18) = "AJ14"
    cell(19) = "AL14"
    cell(20) = "AN14"
    cell(21) = "AP14"
    cell(22) = "AR14"
    cell(23) = "AT14"
    'row 21
    cell(24) = "X21"
    cell(25) = "Z21"
    cell(26) = "AB21"
    cell(27) = "AD21"
    cell(28) = "AF21"
    cell(29) = "AH21"
    cell(30) = "AJ21"
    cell(31) = "AL21"
    cell(32) = "AN21"
    cell(33) = "AP21"
    cell(34) = "AR21"
    cell(35) = "AT21"
    'row 35
    cell(36) = "X35"
    cell(37) = "Z35"
    cell(38) = "AB35"
    cell(39) = "AD35"
    cell(40) = "AF35"
    cell(41) = "AH35"
    cell(42) = "AJ35"
    cell(43) = "AL35"
    cell(44) = "AN35"
    cell(45) = "AP35"
    cell(46) = "AR35"
    cell(47) = "AT35"
    cell(48) = "AP42"
    'row 66
    cell(49) = "X66"
    cell(50) = "Z66"
    cell(51) = "AB66"
    cell(52) = "AD66"
    cell(53) = "AF66"
    cell(54) = "AH66"
    cell(55) = "AJ66"
    cell(56) = "AL66"
    'row 73
    cell(57) = "X73"
    cell(58) = "Z73"
    cell(59) = "AB73"
    cell(60) = "AD73"
    cell(61) = "AF73"
    cell(62) = "AH73"
    cell(63) = "AJ73"
    cell(64) = "AL73"
    'row 80
    cell(65) = "X80"
    cell(66) = "Z80"
    cell(67) = "AB80"
    cell(68) = "AD80"
    cell(69) = "AF80"
    cell(70) = "AH80"
    cell(71) = "AJ80"
    cell(72) = "AL80"
    cell(73) = "AB87"
    cell(74) = "AD87"
    cell(75) = "AF87"
    cell(76) = "AH87"
    cell(77) = "AJ87"
    cell(78) = "AL87"
    'row 94
    cell(79) = "AB94"
    cell(80) = "AD94"
    cell(81) = "AH94"
    cell(82) = "AJ94"
    cell(83) = "AL94"
    'row 28
    cell(84) = "X28"
    cell(85) = "Z28"
    cell(86) = "AB28"
    cell(87) = "AD28"
    cell(88) = "AF28"
    cell(89) = "AF28"
    cell(90) = "AH28"
    cell(91) = "AJ28"
    cell(92) = "AL28"
    cell(93) = "AN28"
    cell(94) = "AP28"
    cell(95) = "AR28"
    cell(96) = "AT28"



    For i = LBound(cell) To UBound(cell)
        ActiveSheet.Range(cell(i)).Select
        sqlcall
    Next i
    
    Dim String0 As String
    Dim String1 As String
    Dim String2 As String
    Dim String3 As String
    Dim String4 As String
    
    
    Dim StringA(100) As String
    
    
        
    Dim loc(0 To 100) As String
    loc(0) = "X5"
    loc(1) = "X12"
    loc(2) = "X19"
    loc(3) = "X26"
    loc(4) = "X33"
    
    loc(5) = "Z5"
    loc(6) = "Z12"
    loc(7) = "Z19"
    loc(8) = "Z26"
    loc(9) = "Z33"
    
    loc(10) = "AB5"
    loc(11) = "AB12"
    loc(12) = "AB19"
    loc(13) = "AB26"
    loc(14) = "AB33"
    
    loc(15) = "AD5"
    loc(16) = "AD12"
    loc(17) = "AD19"
    loc(18) = "AD26"
    loc(19) = "AD33"
    
    loc(20) = "AF5"
    loc(21) = "AF12"
    loc(22) = "AF19"
    loc(23) = "AF26"
    loc(24) = "AF33"
    
    loc(25) = "AH5"
    loc(26) = "AH12"
    loc(27) = "AH19"
    loc(28) = "AH26"
    loc(29) = "AH33"
    
    loc(30) = "AJ5"
    loc(31) = "AJ12"
    loc(32) = "AJ19"
    loc(33) = "AJ26"
    loc(34) = "AJ33"
    
    loc(35) = "Al5"
    loc(36) = "Al12"
    loc(37) = "AL19"
    loc(38) = "AL26"
    loc(39) = "AL33"
    
    
    loc(40) = "AN5"
    loc(41) = "AN12"
    loc(42) = "AN19"
    loc(43) = "AN26"
    loc(44) = "AN33"
    
    loc(45) = "AP5"
    loc(46) = "AP12"
    loc(47) = "AP19"
    loc(48) = "AP26"
    loc(49) = "AP33"
    
    loc(50) = "AR5"
    loc(51) = "AR12"
    loc(52) = "AR19"
    loc(53) = "AR26"
    loc(54) = "AR33"
    
    'section 2 start
    loc(55) = "X64"
    loc(56) = "X71"
    loc(57) = "X78"
    loc(58) = "X85"
    loc(59) = "X92"
    
    loc(60) = "z64"
    loc(61) = "z71"
    loc(62) = "z78"
    loc(63) = "z85"
    loc(64) = "z92"
    
    loc(65) = "AB64"
    loc(66) = "AB71"
    loc(67) = "AB78"
    loc(68) = "AB85"
    loc(69) = "AB92"
    
    loc(70) = "AD64"
    loc(71) = "AD71"
    loc(72) = "AD78"
    loc(73) = "AD85"
    loc(74) = "AD92"
    
    loc(70) = "AF64"
    loc(71) = "AF71"
    loc(72) = "AF78"
    loc(73) = "AF85"
    loc(74) = "AF92"
    
    loc(75) = "AH64"
    loc(76) = "AH71"
    loc(77) = "AH78"
    loc(78) = "AH85"
    loc(79) = "AH92"
    
    loc(80) = "AJ64"
    loc(81) = "AJ71"
    loc(82) = "AJ78"
    loc(83) = "AJ85"
    loc(84) = "AJ92"
    
    loc(85) = "AL64"
    loc(86) = "AL71"
    loc(87) = "AL78"
    loc(88) = "AL85"
    loc(89) = "AL92"

    For co = 0 To 89
    Range(loc(co)).Interior.Color = RGB(255, 255, 255)
    Next co


'populating values
For i = 0 To 89
    If Not IsError(Range(loc(i)).Value) Then
        If Range(loc(i)).Value = "NA" Then
            String1 = ""
        Else
            String1 = Range(loc(i)).Value
        End If
    Else
    String1 = ""
    End If
    StringA(i) = String1
Next i



Dim isTrue As Boolean
Dim StringTemp(5) As String
Dim x As Integer
Dim c As Integer
Dim cx As Integer
x = 0

For o = 0 To 18
    isTrue = True
    cx = x
     For cx = x To cx + 4
        If Not (StrComp(StringA(cx), "") = 0) Then
            StringTemp(c) = StringA(cx)
            'Debug.Print StringTemp(c)
            c = c + 1
        End If
    Next cx
    Debug.Print (x + 5)
 
    
    
    If c > 2 Then
        For cxb = 0 To c - 1
            If Not (StrComp(StringTemp(cxb), StringTemp(cxb + 1)) = 0) Then
                If Not (StrComp(StringTemp(cxb + 1), "") = 0) Then
                    If cxb < 5 Then
                        strArray1 = Split(StringTemp(cxb), ",")
                        strArray2 = Split(StringTemp(cxb + 1), ",")
                        If UBound(strArray1) > UBound(strArray2) Then
                            If strArray1(0) = strArray2(0) Or strArray1(1) = strArray2(0) Then
                            End If
                        Else
                            isTrue = False
                        End If
                    End If
            End If
            End If
        Next cxb
    End If
    
    If c = 2 Then
        If Not (StrComp(StringTemp(0), StringTemp(1)) = 0) Then
            strArray1 = Split(StringTemp(0), ",")
            strArray2 = Split(StringTemp(1), ",")
            If UBound(strArray1) > UBound(strArray2) Then
                If strArray1(0) = strArray2(0) Or strArray1(1) = strArray2(0) Then
                End If
            Else
                isTrue = False
            End If
        End If
    End If
             
  
    
    If isTrue = False Then
    Range(loc(x)).Interior.Color = RGB(200, 160, 35)
    End If
        
Erase StringTemp
x = x + 5
c = 0
Next o


End Sub
