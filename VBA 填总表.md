
```
'####更改记录####
'在32位和64位上运行有一点点不同，如在32位系统运行，会重新打开总表。
'20190516 行29 行44 行63 行39 将 for 2 改为for 2500，提高效率10s，目前运行约30s，主要是vlookup慢
'20190702 行23 因为QC的送检结果含过多的使用过的空单元格，导致整体时长大于60s经常无响应，补充删除空行，实测感觉没啥用，自己动手删吧。
'####


Sub 送检单DNA总表()
Dim sht0, sht1 As Worksheet 'sht0为总表，sht1为送检结果
Dim i, j, L As Integer
Dim rng As Range
If MsgBox("此操作可能会重新打开总表，请确认刚才保存了再点击确定", vbOKCancel, "HBB") = vbCancel Then Exit Sub
Application.ScreenUpdating = False
Application.DisplayAlerts = False
'判定为送检结果
If ActiveSheet.Range("a1") <> "提取后DNA检测交接表" Then
    MsgBox "此表非送检结果，退出程序"
    Exit Sub
End If
Set sht1 = ActiveSheet
'删除UsedRange
'For i = 1 To 16
'    sht1.Cells(sht1.Range("a65536").End(xlUp).Row + 1, 1).Value = 1
'    sht1.Rows(65536).Select
'    Range(Selection, Selection.End(xlUp)).Delete
'Next
'打开总表
Workbooks.Open Filename:="\\172.16.20.23\样本制备（共享）\1.DNA提取\1、2019年提取统计汇总(1月始）.xlsx"
Set sht0 = ActiveWorkbook.Sheets("汇总")
sht0.Range("Z:Z").Clear
sht0.Range("AZ:AZ").Clear
'填入结果
On Error Resume Next
For i = 2500 To 10000
sht0.Range("Q" & i) = WorksheetFunction.VLookup(sht0.Range("d" & i), sht1.Range("e:az"), WorksheetFunction.Match("酶标仪浓度（ng/ul）", sht1.Range("e3:az3"), 0), 0)
sht0.Range("R" & i) = WorksheetFunction.VLookup(sht0.Range("d" & i), sht1.Range("e:az"), WorksheetFunction.Match("质量浓度（ng/ul）", sht1.Range("e3:az3"), 0), 0)
sht0.Range("S" & i) = WorksheetFunction.VLookup(sht0.Range("d" & i), sht1.Range("e:az"), WorksheetFunction.Match("A260/280", sht1.Range("e3:az3"), 0), 0)
sht0.Range("T" & i) = WorksheetFunction.VLookup(sht0.Range("d" & i), sht1.Range("e:az"), WorksheetFunction.Match("A260/230", sht1.Range("e3:az3"), 0), 0)
sht0.Range("az" & i) = WorksheetFunction.VLookup(sht0.Range("d" & i), sht1.Range("e:az"), WorksheetFunction.Match("Qubit浓度（ng/ul）", sht1.Range("e3:az3"), 0), 0)
sht0.Range("ay" & i) = WorksheetFunction.VLookup(sht0.Range("d" & i), sht1.Range("e:az"), WorksheetFunction.Match("电泳结果", sht1.Range("e2:az2"), 0) + 6, 0)

Next
'Qubit与酶标结果合并为一列
For j = 2500 To 100000
    If Range("az" & j) <> "" Then
        sht0.Range("Q" & j) = sht0.Range("az" & j)
    End If
Next
For L = 2500 To Range("a65536").End(xlUp).Row
    If VarType(sht0.Range("Q" & L)) <> 5 And VarType(sht0.Range("R" & L)) <> 0 Then
        sht0.Range("Q" & L) = "/"
    End If
    
    If VarType(sht0.Range("Q" & L)) = 0 Then
        sht0.Range("U" & L) = ""
    ElseIf VarType(sht0.Range("Q" & L)) <> 8 Then
        sht0.Range("U" & L) = sht0.Range("Q" & L) * sht0.Range("m" & L) / 1000
    Else
        sht0.Range("U" & L) = "/"
    End If
    
Next

'结果判定
Dim jg1, jg2, jg3, jg4 As String
Dim k As Integer

For k = 2500 To Range("a65536").End(xlUp).Row
jg1 = ""
jg2 = ""
jg3 = ""
jg4 = ""
jg5 = ""

If sht0.Range("d" & k).Value <> WorksheetFunction.VLookup(sht0.Range("d" & k), sht1.Range("e:az"), 1, 0) Then
    sht0.Range("z" & k) = ""
    Else
    sht0.Range("z" & k) = "送检结果填写"
    
    If sht0.Range("F" & k) = "Fluidigm" Then
        If VarType(sht0.Range("Q" & k)) <> 5 _
        And VarType(sht0.Range("R" & k)) <> 5 _
        And VarType(sht0.Range("S" & k)) <> 5 _
        And VarType(sht0.Range("T" & k)) <> 5 _
        And VarType(sht0.Range("ay" & k)) <> 8 Then
            sht0.Range("X" & k) = "未检"
        Else
            If VarType(sht0.Range("Q" & k)) <> 0 And 8 Then
                If sht0.Range("Q" & k) < 1 Then
                    sht0.Range("V" & k) = "不合格"
                    sht0.Range("a" & k & ":V" & k).Interior.Color = 255
                    jg1 = "Qubit/酶标浓度低；"
                End If
            End If 'Qubit判定
            If sht0.Range("R" & k) < 6 Then
                sht0.Range("V" & k) = "不合格"
                sht0.Range("a" & k & ":V" & k).Interior.Color = 255
                jg3 = "Nano浓度低；"
            End If 'Nano判定
            If sht0.Range("S" & k) < 1.5 Then
                sht0.Range("V" & k) = "不合格"
                sht0.Range("a" & k & ":V" & k).Interior.Color = 255
                jg4 = "260/280低；"
            End If 'Nano纯度判定
            If sht0.Range("T" & k) < 0.8 Then
                sht0.Range("V" & k) = "不合格"
                sht0.Range("a" & k & ":V" & k).Interior.Color = 255
                jg5 = "260/230低；"
            End If 'Nano纯度判定
                If sht0.Range("V" & k) <> "不合格" Then
                sht0.Range("V" & k) = "合格"
            End If '合格判定，异常判定
            sht0.Range("X" & k).Value = jg1 & jg2 & jg3 & jg4 & jg5
        End If '判空，非空继续判定
    End If '判建库类型Fluidigm
    
    If UCase(sht0.Range("F" & k).Value) Like "*META*" _
    Then
         If VarType(sht0.Range("Q" & k)) <> 5 _
        And VarType(sht0.Range("R" & k)) <> 5 _
        And VarType(sht0.Range("S" & k)) <> 5 _
        And VarType(sht0.Range("T" & k)) <> 5 _
        And VarType(sht0.Range("ay" & k)) <> 8 Then
            sht0.Range("X" & k) = "未检"
        Else
            If VarType(sht0.Range("Q" & k)) <> 0 Then
                If sht0.Range("Q" & k) < 1.5 Then
                    sht0.Range("V" & k) = "不合格"
                    sht0.Range("a" & k & ":V" & k).Interior.Color = 255
                    jg1 = "Qubit/酶标浓度低；"
                End If
            End If 'Qubit判定
                If sht0.Range("V" & k) <> "不合格" Then
                sht0.Range("V" & k) = "合格"
            End If '合格判定，异常判定
            sht0.Range("X" & k).Value = jg1 & jg2 & jg3 & jg4 & jg5
        End If '判空，非空继续判定
    End If 'META
    
    If sht0.Range("F" & k).Value = "肠道菌群在高胆固醇血症中的作用" _
    Then
         If VarType(sht0.Range("Q" & k)) <> 5 _
        And VarType(sht0.Range("R" & k)) <> 5 _
        And VarType(sht0.Range("S" & k)) <> 5 _
        And VarType(sht0.Range("T" & k)) <> 5 _
        And VarType(sht0.Range("ay" & k)) <> 8 Then
            sht0.Range("X" & k) = "未检"
        Else
            If VarType(sht0.Range("Q" & k)) <> 0 Then
                If sht0.Range("Q" & k) < 6 Then
                    sht0.Range("V" & k) = "不合格"
                    sht0.Range("a" & k & ":V" & k).Interior.Color = 255
                    jg1 = "Qubit/酶标浓度低；"
                End If
            End If 'Qubit判定
                If sht0.Range("V" & k) <> "不合格" Then
                sht0.Range("V" & k) = "合格"
            End If '合格判定，异常判定
            sht0.Range("X" & k).Value = jg1 & jg2 & jg3 & jg4 & jg5
        End If '判空，非空继续判定
    End If 'META
    
    If UCase(sht0.Range("F" & k).Value) Like "*1K*" _
    Or UCase(sht0.Range("F" & k).Value) Like "*WGS*" _
    Or UCase(sht0.Range("F" & k).Value) Like "*WES*" _
    Or sht0.Range("F" & k) Like "*基因组*" Then
         If VarType(sht0.Range("Q" & k)) <> 5 _
        And VarType(sht0.Range("R" & k)) <> 5 _
        And VarType(sht0.Range("S" & k)) <> 5 _
        And VarType(sht0.Range("T" & k)) <> 5 _
        And VarType(sht0.Range("ay" & k)) <> 8 Then
            sht0.Range("X" & k) = "未检"
        Else
            If VarType(sht0.Range("Q" & k)) <> 0 Then
                If sht0.Range("Q" & k) < 6 Then
                    sht0.Range("V" & k) = "不合格"
                    sht0.Range("a" & k & ":V" & k).Interior.Color = 255
                    jg1 = "Qubit/酶标浓度低；"
                End If
            End If 'Qubit判定
            If sht0.Range("ay" & k) Like "*中度降解*" _
            Or sht0.Range("ay" & k) Like "*重降解*" _
            Or sht0.Range("ay" & k) Like "*完全降解*" _
            Or sht0.Range("ay" & k) Like "*无条带*" Then
                sht0.Range("V" & k) = "不合格"
                sht0.Range("a" & k & ":V" & k).Interior.Color = 255
                jg4 = "降解；"
            End If '电泳判定
                If sht0.Range("V" & k) <> "不合格" Then
                sht0.Range("V" & k) = "合格"
            End If '合格判定，异常判定
            sht0.Range("X" & k).Value = jg1 & jg2 & jg3 & jg4 & jg5
        End If '判空，非空继续判定
    End If '判建库类型1K,WGS,WES
    
    If UCase(sht0.Range("F" & k).Value) Like "*QPCR*" _
    Or UCase(sht0.Range("F" & k).Value) Like "端粒定量" Then
        If VarType(sht0.Range("Q" & k)) <> 5 _
        And VarType(sht0.Range("R" & k)) <> 5 _
        And VarType(sht0.Range("S" & k)) <> 5 _
        And VarType(sht0.Range("T" & k)) <> 5 _
        And VarType(sht0.Range("ay" & k)) <> 8 Then
            sht0.Range("X" & k) = "未检"
        Else
            If VarType(sht0.Range("Q" & k)) <> 0 Then
                If sht0.Range("Q" & k) < 0.5 Then
                    sht0.Range("V" & k) = "不合格"
                    sht0.Range("a" & k & ":V" & k).Interior.Color = 255
                    jg1 = "Qubit/酶标浓度低；"
                End If
            End If 'Qubit判定
            If sht0.Range("ay" & k) Like "*中度降解*" _
            Or sht0.Range("ay" & k) Like "*重降解*" _
            Or sht0.Range("ay" & k) Like "*无条带*" Then
                sht0.Range("V" & k) = "不合格"
                sht0.Range("a" & k & ":V" & k).Interior.Color = 255
                jg4 = "降解；"
            End If '电泳判定降解
            If sht0.Range("ay" & k) Like "*蛋白*" _
            Or sht0.Range("ay" & k) Like "*污染*" _
            Or UCase(sht0.Range("l" & k).Value) Like "*RNA*" Then
                sht0.Range("n" & k) = "不合格"
                jg5 = "污染；"
            End If '电泳判定污染
                If sht0.Range("V" & k) <> "不合格" Then
                sht0.Range("V" & k) = "合格"
            End If '合格判定，异常判定
            sht0.Range("X" & k).Value = jg1 & jg2 & jg3 & jg4 & jg5
        End If '判空，非空继续判定
    End If '判建库类型QPCR
    
'    If UCase(sht0.Range("F" & k).Value) = "16S" Then
'        If VarType(sht0.Range("Q" & k)) <> 5 _
'        And VarType(sht0.Range("R" & k)) <> 5 _
'        And VarType(sht0.Range("S" & k)) <> 5 _
'        And VarType(sht0.Range("T" & k)) <> 5 _
'        And VarType(sht0.Range("ay" & k)) <> 8 Then
'            sht0.Range("X" & k) = "未检"
'        ElseIf sht0.Range("f" & k) = "皮肤拭子" Or "皮肤微生物" Then
'            sht0.Range("V" & k) = "合格"
'            sht0.Range("X" & k) = ""
'        Else
'            If VarType(sht0.Range("Q" & k)) <> 0 Then
'                If sht0.Range("Q" & k) < 1 Then
'                    sht0.Range("V" & k) = "不合格"
'                    sht0.Range("a" & k & ":V" & k).Interior.Color = 255
'                    jg1 = "Qubit/酶标浓度低；"
'                End If
'            End If 'Qubit判定
'            If sht0.Range("ay" & k) Like "*中度降解*" _
'            Or sht0.Range("ay" & k) Like "*重降解*" _
'            Or sht0.Range("ay" & k) Like "*完全降解*" _
'            Or sht0.Range("ay" & k) Like "*无条带*" Then
'                sht0.Range("V" & k) = "不合格"
'                sht0.Range("a" & k & ":V" & k).Interior.Color = 255
'
'                jg5 = "降解；"
'            End If '电泳判定
'                If sht0.Range("V" & k) <> "不合格" Then
'                sht0.Range("V" & k) = "合格"
'            End If '合格判定，异常判定
'            sht0.Range("X" & k).Value = jg1 & jg2 & jg3 & jg4 & jg5
'        End If '判空，非空继续判定
'    End If '判建库类型16S
    If UCase(sht0.Range("F" & k).Value) = "16S" Then
        If VarType(sht0.Range("Q" & k)) <> 5 _
        And VarType(sht0.Range("R" & k)) <> 5 _
        And VarType(sht0.Range("S" & k)) <> 5 _
        And VarType(sht0.Range("T" & k)) <> 5 _
        And VarType(sht0.Range("ay" & k)) <> 8 Then
            sht0.Range("X" & k) = "未检"
        Else
            If VarType(sht0.Range("Q" & k)) <> 0 Then
                If sht0.Range("Q" & k) < 1 Then
                    sht0.Range("V" & k) = "不合格"
                    sht0.Range("a" & k & ":V" & k).Interior.Color = 255
                    jg1 = "Qubit/酶标浓度低；"
                End If
            End If 'Qubit判定
            If sht0.Range("ay" & k) Like "*中度降解*" _
            Or sht0.Range("ay" & k) Like "*重降解*" _
            Or sht0.Range("ay" & k) Like "*完全降解*" _
            Or sht0.Range("ay" & k) Like "*无条带*" Then
                sht0.Range("V" & k) = "不合格"
                sht0.Range("a" & k & ":V" & k).Interior.Color = 255

                jg5 = "降解；"
            End If '电泳判定
                If sht0.Range("V" & k) <> "不合格" Then
                sht0.Range("V" & k) = "合格"
            End If '合格判定，异常判定
            sht0.Range("X" & k).Value = jg1 & jg2 & jg3 & jg4 & jg5
        End If '判空，非空继续判定
    End If '判建库类型16S
    
    If sht0.Range("e" & k).Value = "阴性对照" Then
        If VarType(sht0.Range("Q" & k)) <> 5 _
        And VarType(sht0.Range("R" & k)) <> 5 _
        And VarType(sht0.Range("S" & k)) <> 5 _
        And VarType(sht0.Range("T" & k)) <> 5 _
        And VarType(sht0.Range("ay" & k)) <> 8 Then
            sht0.Range("X" & k) = "未检"
        Else
            If VarType(sht0.Range("Q" & k)) <> 0 Then
            jg1 = ""
            jg2 = ""
            jg3 = ""
            jg4 = ""
            jg5 = ""
                If sht0.Range("Q" & k) <= 0.19 Then
                    sht0.Range("V" & k) = "合格"
                    sht0.Range("a" & k & ":V" & k).Interior.Pattern = xlNone
                ElseIf sht0.Range("Q" & k) = "/" Then
                    sht0.Range("V" & k) = "合格"
                    sht0.Range("a" & k & ":V" & k).Interior.Pattern = xlNone
                ElseIf sht0.Range("Q" & k) > 0.19 Then
                    sht0.Range("V" & k) = "不合格"
                    sht0.Range("a" & k & ":V" & k).Interior.Color = 255
                    jg1 = "Qubit/酶标有浓度；"
                End If
            End If 'Qubit判定
            sht0.Range("X" & k).Value = jg1 & jg2 & jg3 & jg4 & jg5
        End If '判空，非空继续判定
    End If '判阴性对照####2
     
End If
Next

'确认结果
MsgBox "完毕，请查收！"
sht0.Range(sht0.Cells.Find(sht1.Range("e4")).Address).Select
Selection.Offset(0, 23).Select

Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
```
