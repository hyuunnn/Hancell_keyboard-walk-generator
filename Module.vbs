Option Explicit
Function StringArray(ByRef str)
    Dim str_tmp, output, i
  
    ' 한셀에서는 ByRef만 사용가능하여 tmp 변수 생성 (ByVal로 지정해도 ByRef로 동작하므로 값이 변경된다.)
    str_tmp = str

    For i = 1 To Len(str_tmp)
        output = output + Left(str_tmp, 1) + " "
        str_tmp = Right(str_tmp, Len(str_tmp) - 1)
    Next
    StringArray = Split(output)
End Function

Function findArrayData(ByRef str, ByRef keyArr)
  Dim check_boolean, i
  check_boolean = False
  For Each i in keyArr
    If i = str Then
      check_boolean = True
    End If
  Next
  findArrayData = check_boolean
End Function

Function getArrayIdx(ByRef str, ByRef keyArr)
  Dim idx, i
  idx = 0
  For i = 0 To UBound(keyArr)
    If keyArr(i) = str Then
      idx = i
    End If
  Next
  getArrayIdx = idx
End Function

' 자바스크립트의 호이스팅처럼 코드가 이상하게 동작한다.
' getIdx(j) 부분을 j로 가져오면 j의 마지막 루프 num 값이 저장된다.
Function getIdx(ByRef num)
    getIdx = num
End Function

Function getPosArrayOutput(ByRef posArr, ByRef keyArr, ByRef num)
    Dim output, i
    For Each i in posArr
        output = output + keyArr(i(0))(i(1) + num)
    Next
    getPosArrayOutput = output
End Function

Function reverse(ByRef str)
    Dim str_tmp, output, i
    str_tmp = StringArray(str)
    For i = UBound(str_tmp) To 0 Step -1
        output = output + str_tmp(i)
    Next
    reverse = output
End Function

Function reverseStr(ByRef str)
    Dim first_pattern, mid_pattern, last_pattern
    first_pattern = Left(str, 2)
    mid_pattern = reverse(Mid(str, 3, Len(str) - 4)) ' 처음 2글자, 마지막 2글자의 길이를 제거해야 하기 때문에 4를 뺀다.
    last_pattern = Right(str, 2)
    reverseStr = first_pattern + mid_pattern + last_pattern
End Function

Function twoBytesReverseStr(ByRef str)
    Dim first_pattern, mid1_pattern, mid2_pattern, mid3_pattern, last_pattern
    first_pattern = Left(str, 2)
    mid1_pattern = reverse(Mid(str, 3, 2)) ' 처음 2글자를 역방향 패턴으로 저장
    mid2_pattern = reverse(Mid(str, 5, 2)) ' 그 다음 2글자를 역방향 패턴으로 저장
    mid3_pattern = reverse(Mid(str, 7, 2))
    last_pattern = Right(str, 2) ' 나머지 패턴을 저장
    twoBytesReverseStr = first_pattern + mid1_pattern + mid2_pattern + mid3_pattern + last_pattern
End Function

Function setReversePattern(ByRef str)
    Dim check_pattern1, check_pattern2
    check_pattern1 = ActiveSheet.CheckBoxes().Item(1).Value
    check_pattern2 = ActiveSheet.CheckBoxes().Item(2).Value
    
    If (check_pattern1 = 1 And check_pattern2 = 1) Then
        setReversePattern = ""
    ElseIf (check_pattern1 = 1) Then
        setReversePattern = reverseStr(str)
    ElseIf (check_pattern2 = 1) Then
        setReversePattern = twoBytesReverseStr(str)
    End if
End Function

Sub 확인()
    Dim keyArray(7)
    keyArray(0) = Array("1","2","3","4","5","6","7","8","9","0","-","=","\")
    keyArray(1) = Array("q","w","e","r","t","y","u","i","o","p","[","]")
    keyArray(2) = Array("a","s","d","f","g","h","j","k","l",";","'")
    keyArray(3) = Array("z","x","c","v","b","n","m",",",".","/")
    keyArray(4) = Array("!","@","#","$","%","^","&","*","(",")","_","+","|")
    keyArray(5) = Array("Q","W","E","R","T","Y","U","I","O","P","{","}")
    keyArray(6) = Array("A","S","D","F","G","H","J","K","L",":","""")
    keyArray(7) = Array("Z","X","C","V","B","N","M","<",">","?")

    Dim inputData, inputArray
    inputData = Range("F2:F2").Value

    If Len(inputData) = 10 Then
        inputArray = StringArray(inputData)

        'Dim posArray(Len(inputData)) ' array length 값을 직접 입력하면 되는데 값을 가져와서 사용하면 에러가 발생한다.
        ' 구글링해서 배열 길이의 변수 지정하는 방법 찾아보기
        Dim posArray(9)
    
        Dim pos_idx, i, j
        pos_idx = 0
        For Each i in inputArray
            For j = 0 To 7 ' keyArray 배열 7개 for loop
                If findArrayData(i, keyArray(j)) = True Then
                    'MsgBox(getIdx(j) & " " & getArrayIdx(i, keyArray(j)))

                    ' 몇 번째 keyArray에 있는지, 해당 keyArray에서 몇 번째에 있는지
                    posArray(pos_idx) = Array(getIdx(j), getArrayIdx(i, keyArray(j)))

                    pos_idx = pos_idx + 1
                End If
            Next
        Next

        ' 결과 저장
        Range("B3:B3").Value = getPosArrayOutput(posArray, keyArray, 0)
        Range("B4:B4").Value = getPosArrayOutput(posArray, keyArray, 1)
        Range("B5:B5").Value = getPosArrayOutput(posArray, keyArray, 2)
        Range("B6:B6").Value = getPosArrayOutput(posArray, keyArray, 3)
        Range("B7:B7").Value = getPosArrayOutput(posArray, keyArray, 4)
        Range("B8:B8").Value = getPosArrayOutput(posArray, keyArray, 5)
        Range("B9:B9").Value = getPosArrayOutput(posArray, keyArray, 6)
        Range("B10:B10").Value = getPosArrayOutput(posArray, keyArray, 7)
        
        Range("C3:C3").Value = setReversePattern(getPosArrayOutput(posArray, keyArray, 0))
        Range("C4:C4").Value = setReversePattern(getPosArrayOutput(posArray, keyArray, 1))
        Range("C5:C5").Value = setReversePattern(getPosArrayOutput(posArray, keyArray, 2))
        Range("C6:C6").Value = setReversePattern(getPosArrayOutput(posArray, keyArray, 3))
        Range("C7:C7").Value = setReversePattern(getPosArrayOutput(posArray, keyArray, 4))
        Range("C8:C8").Value = setReversePattern(getPosArrayOutput(posArray, keyArray, 5))
        Range("C9:C9").Value = setReversePattern(getPosArrayOutput(posArray, keyArray, 6))
        Range("C10:C10").Value = setReversePattern(getPosArrayOutput(posArray, keyArray, 7))
    Else
        MsgBox("입력한 패턴의 글자 수가 맞지 않습니다.")
    End If
End Sub
