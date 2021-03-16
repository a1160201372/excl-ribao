VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4365
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    '运行过程中禁止点击
    TextBox1.Enabled = False
    CommandButton1.Enabled = False
    CommandButton2.Enabled = False
    

    
    '定义变量
    '底板信息
    Dim NegativeName1, NegativePath1 As String '新家客日报_模板文件位置，文件名称
    Dim NegativeName2, NegativePath2 As String '日报专用模板_模板文件位置，文件名称
    
    Dim HomeName, HomePath As String '家客处理文件位置，文件名称
    Dim UseName, UsePath As String '日报专用模板处理文件位置，文件名称
    '原始位置信息
    Dim AddName, AddPath, AddNameNull As String '新增量位置信息
    Dim InstallName, InstallPath, InstallNameNull As String '装机量位置信息
    Dim RunName, RunPath, RunNameNull As String '装机量位置信息
    '三个池子
    Dim DispatchName, DispatchPath, DispatchNameNull As String '待装池
    Dim ExpansionName, ExpansionPath, ExpansionNameNull As String '扩容池
    Dim AppointmentName, AppointmentPath, AppointmentNameNull As String '预约池
    
    

    
    Dim pth As String '文件夹位置
    
    Dim Year, Month, Day As String '文件夹位置
    
    
    
    '变量赋值
    
    pth = ThisWorkbook.Path
    pth = Left(pth, Len(pth) - 2)
    
    HomeName = "新家客日报" & TextBox1.Text & ".xls"
    HomePath = pth & "处理\" & HomeName
    
    UseName = "日报专用模板" & TextBox1.Text & ".xls"
    UsePath = pth & "处理\" & UseName
    
    NegativeName1 = "新家客日报.xls"
    NegativePath1 = pth & "底版\" & NegativeName1
    NegativeName2 = "日报专用模板.xls"
    NegativePath2 = pth & "底版\" & NegativeName2
    
    AddNameNull = TextBox1.Text & "_新增量"
    AddName = AddNameNull & ".csv"
    AddPath = pth & "处理\" & AddName
    
    InstallNameNull = TextBox1.Text & "_装机量"
    InstallName = InstallNameNull & ".csv"
    InstallPath = pth & "处理\" & InstallName
    
    RunNameNull = TextBox1.Text & "_运行中"
    RunName = RunNameNull & ".csv"
    RunPath = pth & "处理\" & RunName
    
    DispatchNameNull = TextBox1.Text & "_待装池"
    DispatchName = DispatchNameNull & ".xls"
    DispatchPath = pth & "处理\" & DispatchName
    
    ExpansionNameNull = TextBox1.Text & "_扩容池"
    ExpansionName = ExpansionNameNull & ".xls"
    ExpansionPath = pth & "处理\" & ExpansionName
    
    AppointmentNameNull = TextBox1.Text & "_预约池"
    AppointmentName = AppointmentNameNull & ".xls"
    AppointmentPath = pth & "处理\" & AppointmentName
    
    Year = TextBox1.Text / 10000
    Year = Format(Year, 0)
    Month = TextBox1.Text Mod 10000
    Month = Month / 100
    Month = Format(Month, 0)
    Day = TextBox1.Text Mod 100
   
  
   
    
   'MsgBox (Month & "/" & Day - 1 & "/" & Year)
    
   
    
    
'

'    '判断是否存在文件

    If Dir(NegativePath1, vbDirectory) = vbNullString Then
        MsgBox ("未找到模板文件_新家客日报！")
        TextBox1.Enabled = True
        CommandButton1.Enabled = True
        CommandButton2.Enabled = True
        Exit Sub
    End If
    If Dir(NegativePath2, vbDirectory) = vbNullString Then
        MsgBox ("未找到模板文件_日报专用模板！")
        TextBox1.Enabled = True
        CommandButton1.Enabled = True
        CommandButton2.Enabled = True
        Exit Sub
    End If
    
    If Dir(pth & "原始数据\" & "新增量\" & TextBox1.Text & ".csv", vbDirectory) = vbNullString Then
        MsgBox ("未找到新增量文件！")
        TextBox1.Enabled = True
        CommandButton1.Enabled = True
        CommandButton2.Enabled = True
        Exit Sub
    End If
    If Dir(pth & "原始数据\" & "装机量\" & TextBox1.Text & ".csv", vbDirectory) = vbNullString Then
        MsgBox ("未找到装机量文件！")
        TextBox1.Enabled = True
        CommandButton1.Enabled = True
        CommandButton2.Enabled = True
        Exit Sub
    End If
    If Dir(pth & "原始数据\" & "运行中\" & TextBox1.Text & ".csv", vbDirectory) = vbNullString Then
        MsgBox ("未找到运行中文件！")
        TextBox1.Enabled = True
        CommandButton1.Enabled = True
        CommandButton2.Enabled = True
        Exit Sub
    End If
    
    If Dir(pth & "原始数据\" & "待装池\" & TextBox1.Text & ".xls", vbDirectory) = vbNullString Then
        MsgBox ("未找到待装池文件！")
        TextBox1.Enabled = True
        CommandButton1.Enabled = True
        CommandButton2.Enabled = True
        Exit Sub
    End If
    If Dir(pth & "原始数据\" & "扩容池\" & TextBox1.Text & ".xls", vbDirectory) = vbNullString Then
        MsgBox ("未找到扩容池文件！")
        TextBox1.Enabled = True
        CommandButton1.Enabled = True
        CommandButton2.Enabled = True
        Exit Sub
    End If
    If Dir(pth & "原始数据\" & "预约池\" & TextBox1.Text & ".xls", vbDirectory) = vbNullString Then
        MsgBox ("未找到预约池文件！")
        TextBox1.Enabled = True
        CommandButton1.Enabled = True
        CommandButton2.Enabled = True
        Exit Sub
    End If

    '复制底板和原始数据
    FileCopy NegativePath1, HomePath   '
    FileCopy NegativePath2, UsePath    '
    FileCopy pth & "原始数据\" & "新增量\" & TextBox1.Text & ".csv", AddPath
    FileCopy pth & "原始数据\" & "装机量\" & TextBox1.Text & ".csv", InstallPath
    FileCopy pth & "原始数据\" & "运行中\" & TextBox1.Text & ".csv", RunPath
    FileCopy pth & "原始数据\" & "待装池\" & TextBox1.Text & ".xls", DispatchPath
    FileCopy pth & "原始数据\" & "扩容池\" & TextBox1.Text & ".xls", ExpansionPath
    FileCopy pth & "原始数据\" & "预约池\" & TextBox1.Text & ".xls", AppointmentPath


   ' 打开所需文件
    Workbooks.Open Filename:=HomePath, AddToMru:=True
    Workbooks.Open Filename:=UsePath, AddToMru:=True

    Workbooks.Open Filename:=AddPath, AddToMru:=True
    Workbooks.Open Filename:=InstallPath, AddToMru:=True
    Workbooks.Open Filename:=RunPath, AddToMru:=True
    Workbooks.Open Filename:=DispatchPath, AddToMru:=True
    Workbooks.Open Filename:=ExpansionPath, AddToMru:=True
    Workbooks.Open Filename:=AppointmentPath, AddToMru:=True

    

    Call 初始化(HomeName, UseName, Year, Month, Day)
    Call 新增量_装机量_去重(AddName, AddNameNull, InstallName, InstallNameNull)
    Call 处理数据_新增量_装机量(AddName, AddNameNull, HomeName, InstallName, InstallNameNull, Year, Month, Day)
    Call 复制数据_日报_宽带装移机进展(HomeName, UseName, Day)
    Call 处理_专用日报_有效带宽_三日预约量(InstallName, InstallNameNull, UseName, RunName, RunNameNull, Year, Month, Day)
    Call 处理_专用日报_待装工单(DispatchName, ExpansionName, AppointmentName, RunName, RunNameNull, UseName, HomeName)
    Call 统计_新家客_待装天数_ImsOtt(RunName, RunNameNull, HomeName, UseName)
    Call 复制_专用日报_待装天数模板(HomeName, UseName)
    
    
    '关闭文件
    
    
    
    '复制文件
    
    
    '恢复使能
    TextBox1.Enabled = True
    CommandButton1.Enabled = True
    CommandButton2.Enabled = True
     MsgBox "结束"
     
    
    
    
    
    
    
    
End Sub
Function 初始化(HomeName, UseName, Year, Month, Day)
    Call 家客_家客_初始化(HomeName, Month, Day)
    Call 家客_待装天数_新__初始化(HomeName, Day)
    Call 日报专用_有线宽带_初始化(UseName, Month, Day)
End Function
Function 家客_家客_初始化(HomeName, Month, Day)
    Windows(HomeName).Activate
    Sheets("新家客日报（模板）").Activate
    
    Range("K1:N11").Select
    Selection.Copy
    Range("M1:P11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    Range("K1").Formula = Month & "月" & Day - 1 & "日新增工单数"
    
    Range("Q1:T11").Select
    Selection.Copy
    Range("S1:V11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    Range("Q1").Formula = Month & "月" & Day - 1 & "日完成工单数"

    Range("K14:N24").Select
    Selection.Copy
    Range("M14:P24").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    Range("K14").Formula = Month & "月" & Day - 1 & "日新增装、移机工单数"
    
    Range("Q14:T24").Select
    Selection.Copy
    Range("S14:V24").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    Range("Q14").Formula = Month & "月" & Day - 1 & "日完成装、移机工单数"

    Range("K3:L11").Select
    Application.CutCopyMode = False
    
End Function
Function 日报专用_有线宽带_初始化(UseName, Month, Day)
    Windows(UseName).Activate
    Sheets("有线宽带报表").Activate
    Range("B2").Formula = Month & "." & Day & "日预约量"
    Range("C2").Formula = Month & "." & Day + 1 & "日预约量"
    Range("D2").Formula = Month & "." & Day + 2 & "日预约量"

    

End Function
Function 家客_待装天数_新__初始化(HomeName, Day)
    Windows(HomeName).Activate
    Sheets("待装天数（新）").Activate
    
    Range("F11").Formula = Day - 3 & "日完成数"
    Range("G11").Formula = Day - 2 & "日完成数"
    Range("H11").Formula = Day - 1 & "日完成数"
    
    
    
End Function
Function 处理数据_新增量_装机量(AddName, AddNameNull, HomeName, InstallName, InstallNameNull, Year, Month, Day)

    '删除 特定字符
    Windows(AddName).Activate
    Dim EndH, i As Long
    EndH = Range("B65536").End(xlUp).Row
    
    For i = EndH To 1 Step -1
     XX = Range("B" & i).Value
     If InStr(XX, "【智能组网】") <> 0 Then '@就是要找到指定的特定字符 可以改成你指定的其他字符
          Range("B" & i).EntireRow.Delete 'E列是指定要查找的列.狂野改成你指定的其他列
     End If
     Next

    '处理新增量文件
    '新建工作表
    Worksheets.Add before:=Sheets(AddNameNull)
    Sheets(AddNameNull).Activate
    '第三个
    Cells.Select
    Cells.AutoFilter
    ActiveWorkbook.Names.Add Name:="'" & AddNameNull & "'!_FilterDatabase", RefersTo:="='" & AddNameNull & "'!$A$1:$BW$65536", Visible:=False
    Range("N4").Select
    ActiveWindow.ScrollColumn = 26
    Columns("AI:AI").Select
    Range("A1:BW65536").AutoFilter Field:=35, Criteria1:=Array("已归档", "已开通", "执行中"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & AddNameNull & "'!_FilterDatabase", RefersTo:="='" & AddNameNull & "'!$A$1:$BW$65536", Visible:=False
    ActiveWindow.ScrollColumn = 16
    Range("A1:BW65536").AutoFilter Field:=19, Operator:=xlFilterValues, Criteria2:=Array(2, Month & "/" & Day - 1 & "/" & Year) '"3/8/2021"
    ActiveWorkbook.Names.Add Name:="'" & AddNameNull & "'!_FilterDatabase", RefersTo:="='" & AddNameNull & "'!$A$1:$BW$65536", Visible:=False
    ActiveWindow.ScrollColumn = 1
    Columns("L:L").Select
    Selection.Copy
    Range("J4201").Select
    Sheets("Sheet1").Activate
    Columns("E:E").Select
    ActiveSheet.Paste
    Range("G4").Select
    Sheets(AddNameNull).Activate
    '第一个
    Range("A1:BW65536").AutoFilter Field:=4, Criteria1:=Array("业务融合开机流程"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & AddNameNull & "'!_FilterDatabase", RefersTo:="='" & AddNameNull & "'!$A$1:$BW$65536", Visible:=False
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("A:A").Select
    ActiveSheet.Paste
    Sheets(AddNameNull).Activate
    Range("A1:BW65536").AutoFilter Field:=4
    ActiveWorkbook.Names.Add Name:="'" & AddNameNull & "'!_FilterDatabase", RefersTo:="='" & AddNameNull & "'!$A$1:$BW$65536", Visible:=False
    '第二个
    Range("A1:BW65536").AutoFilter Field:=4, Criteria1:=Array("业务融合开机流程"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & AddNameNull & "'!_FilterDatabase", RefersTo:="='" & AddNameNull & "'!$A$1:$BW$65536", Visible:=False
    Range("A1:BW65536").AutoFilter Field:=5, Criteria1:=Array("家庭有线宽带"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & AddNameNull & "'!_FilterDatabase", RefersTo:="='" & AddNameNull & "'!$A$1:$BW$65536", Visible:=False
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("C:C").Select
    ActiveSheet.Paste
    Sheets(AddNameNull).Activate
    Range("A1:BW65536").AutoFilter Field:=4
    ActiveWorkbook.Names.Add Name:="'" & AddNameNull & "'!_FilterDatabase", RefersTo:="='" & AddNameNull & "'!$A$1:$BW$65536", Visible:=False
    Range("A1:BW65536").AutoFilter Field:=5
    ActiveWorkbook.Names.Add Name:="'" & AddNameNull & "'!_FilterDatabase", RefersTo:="='" & AddNameNull & "'!$A$1:$BW$65536", Visible:=False
    '第四个
    Range("A1:BW65536").AutoFilter Field:=5, Criteria1:=Array("家庭有线宽带"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & AddNameNull & "'!_FilterDatabase", RefersTo:="='" & AddNameNull & "'!$A$1:$BW$65536", Visible:=False
    Selection.Copy
    Range("K4200").Select
    Sheets("Sheet1").Activate
    Columns("G:G").Select
    ActiveSheet.Paste

    '赋值
    Windows(HomeName).Activate
    Sheets("新家客日报（模板）").Activate
    '第一个
    Range("K3").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$A$1:$A$65536," & """" & "滨北区" & """" & ")"
    Range("K4").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$A$1:$A$65536," & """" & "滨城区" & """" & ")"
    Range("K5").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$A$1:$A$65536," & """" & "博兴县" & """" & ")"
    Range("K6").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$A$1:$A$65536," & """" & "惠民县" & """" & ")"
    Range("K7").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$A$1:$A$65536," & """" & "无棣县" & """" & ")"
    Range("K8").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$A$1:$A$65536," & """" & "阳信县" & """" & ")"
    Range("K9").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$A$1:$A$65536," & """" & "沾化区" & """" & ")"
    Range("K10").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$A$1:$A$65536," & """" & "邹平市" & """" & ")"
    '第二个
    Range("L3").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$C$1:$C$65536," & """" & "滨北区" & """" & ")"
    Range("L4").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$C$1:$C$65536," & """" & "滨城区" & """" & ")"
    Range("L5").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$C$1:$C$65536," & """" & "博兴县" & """" & ")"
    Range("L6").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$C$1:$C$65536," & """" & "惠民县" & """" & ")"
    Range("L7").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$C$1:$C$65536," & """" & "无棣县" & """" & ")"
    Range("L8").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$C$1:$C$65536," & """" & "阳信县" & """" & ")"
    Range("L9").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$C$1:$C$65536," & """" & "沾化区" & """" & ")"
    Range("L10").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$C$1:$C$65536," & """" & "邹平市" & """" & ")"
     '第三个
    Range("K16").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$E$1:$E$65536," & """" & "滨北区" & """" & ")"
    Range("K17").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$E$1:$E$65536," & """" & "滨城区" & """" & ")"
    Range("K18").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$E$1:$E$65536," & """" & "博兴县" & """" & ")"
    Range("K19").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$E$1:$E$65536," & """" & "惠民县" & """" & ")"
    Range("K20").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$E$1:$E$65536," & """" & "无棣县" & """" & ")"
    Range("K21").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$E$1:$E$65536," & """" & "阳信县" & """" & ")"
    Range("K22").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$E$1:$E$65536," & """" & "沾化区" & """" & ")"
    Range("K23").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$E$1:$E$65536," & """" & "邹平市" & """" & ")"
    '第二个
    Range("L16").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$G$1:$G$65536," & """" & "滨北区" & """" & ")"
    Range("L17").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$G$1:$G$65536," & """" & "滨城区" & """" & ")"
    Range("L18").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$G$1:$G$65536," & """" & "博兴县" & """" & ")"
    Range("L19").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$G$1:$G$65536," & """" & "惠民县" & """" & ")"
    Range("L20").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$G$1:$G$65536," & """" & "无棣县" & """" & ")"
    Range("L21").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$G$1:$G$65536," & """" & "阳信县" & """" & ")"
    Range("L22").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$G$1:$G$65536," & """" & "沾化区" & """" & ")"
    Range("L23").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$G$1:$G$65536," & """" & "邹平市" & """" & ")"
     Cells.Select
     Cells.AutoFilter

    '选中装机量
    Windows(InstallName).Activate
    '删除特定行
    EndH = Range("B65536").End(xlUp).Row

    For i = EndH To 1 Step -1
     XX = Range("B" & i).Value
     If InStr(XX, "【智能组网】") <> 0 Then '@就是要找到指定的特定字符 可以改成你指定的其他字符
          Range("B" & i).EntireRow.Delete 'E列是指定要查找的列.狂野改成你指定的其他列
     End If
     Next

    Application.DisplayAlerts = False
    Application.DisplayAlerts = True
    Cells.Select
    Cells.AutoFilter
    ActiveWorkbook.Names.Add Name:="'" & InstallNameNull & "'!_FilterDatabase", RefersTo:="='" & InstallNameNull & "'!$A$1:$BW$4034", Visible:=False
    Range("M7").Select
    ActiveWindow.ScrollColumn = 23
    Range("AJ1").Select
    Range("A1:BW4034").AutoFilter Field:=35, Criteria1:=Array("已归档", "已开通"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & InstallNameNull & "'!_FilterDatabase", RefersTo:="='" & InstallNameNull & "'!$A$1:$BW$4034", Visible:=False
    Range("AG1").Select
    Range("A1:BW4034").AutoFilter Field:=33, Operator:=xlFilterValues, Criteria2:=Array(2, Month & "/" & Day - 1 & "/" & Year) '"3/8/2021"
    ActiveWorkbook.Names.Add Name:="'" & InstallNameNull & "'!_FilterDatabase", RefersTo:="='" & InstallNameNull & "'!$A$1:$BW$4034", Visible:=False
    ActiveWindow.ScrollColumn = 1

    Worksheets.Add before:=Sheets(InstallNameNull)
    Sheets(InstallNameNull).Activate
    '第三个
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("E:E").Select
    ActiveSheet.Paste
    Sheets(InstallNameNull).Activate
    '第一个
    Range("A1:BW65536").AutoFilter Field:=4, Criteria1:=Array("业务融合开机流程"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & InstallNameNull & "'!_FilterDatabase", RefersTo:="='" & InstallNameNull & "'!$A$1:$BW$65536", Visible:=False
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("A:A").Select
    ActiveSheet.Paste
    Sheets(InstallNameNull).Activate
     '第二个
    Range("A1:BW65536").AutoFilter Field:=5, Criteria1:=Array("家庭有线宽带"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & InstallNameNull & "'!_FilterDatabase", RefersTo:="='" & InstallNameNull & "'!$A$1:$BW$65536", Visible:=False
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("C:C").Select
    ActiveSheet.Paste
    Sheets(InstallNameNull).Activate
     '第四个
    Range("A1:BW65536").AutoFilter Field:=4
    ActiveWorkbook.Names.Add Name:="'" & InstallNameNull & "'!_FilterDatabase", RefersTo:="='" & InstallNameNull & "'!$A$1:$BW$65536", Visible:=False
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("G:G").Select
    ActiveSheet.Paste
    Sheets(InstallNameNull).Activate

    '赋值
    Windows(HomeName).Activate
    '第一个
    Range("Q3").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$A$1:$A$65536," & """" & "滨北区" & """" & ")"
    Range("Q4").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$A$1:$A$65536," & """" & "滨城区" & """" & ")"
    Range("Q5").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$A$1:$A$65536," & """" & "博兴县" & """" & ")"
    Range("Q6").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$A$1:$A$65536," & """" & "惠民县" & """" & ")"
    Range("Q7").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$A$1:$A$65536," & """" & "无棣县" & """" & ")"
    Range("Q8").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$A$1:$A$65536," & """" & "阳信县" & """" & ")"
    Range("Q9").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$A$1:$A$65536," & """" & "沾化区" & """" & ")"
    Range("Q10").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$A$1:$A$65536," & """" & "邹平市" & """" & ")"
    '第二个
    Range("R3").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$C$1:$C$65536," & """" & "滨北区" & """" & ")"
    Range("R4").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$C$1:$C$65536," & """" & "滨城区" & """" & ")"
    Range("R5").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$C$1:$C$65536," & """" & "博兴县" & """" & ")"
    Range("R6").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$C$1:$C$65536," & """" & "惠民县" & """" & ")"
    Range("R7").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$C$1:$C$65536," & """" & "无棣县" & """" & ")"
    Range("R8").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$C$1:$C$65536," & """" & "阳信县" & """" & ")"
    Range("R9").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$C$1:$C$65536," & """" & "沾化区" & """" & ")"
    Range("R10").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$C$1:$C$65536," & """" & "邹平市" & """" & ")"
     '第三个
    Range("Q16").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$E$1:$E$65536," & """" & "滨北区" & """" & ")"
    Range("Q17").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$E$1:$E$65536," & """" & "滨城区" & """" & ")"
    Range("Q18").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$E$1:$E$65536," & """" & "博兴县" & """" & ")"
    Range("Q19").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$E$1:$E$65536," & """" & "惠民县" & """" & ")"
    Range("Q20").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$E$1:$E$65536," & """" & "无棣县" & """" & ")"
    Range("Q21").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$E$1:$E$65536," & """" & "阳信县" & """" & ")"
    Range("Q22").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$E$1:$E$65536," & """" & "沾化区" & """" & ")"
    Range("Q23").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$E$1:$E$65536," & """" & "邹平市" & """" & ")"
    '第四个
    Range("R16").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$G$1:$G$65536," & """" & "滨北区" & """" & ")"
    Range("R17").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$G$1:$G$65536," & """" & "滨城区" & """" & ")"
    Range("R18").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$G$1:$G$65536," & """" & "博兴县" & """" & ")"
    Range("R19").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$G$1:$G$65536," & """" & "惠民县" & """" & ")"
    Range("R20").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$G$1:$G$65536," & """" & "无棣县" & """" & ")"
    Range("R21").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$G$1:$G$65536," & """" & "阳信县" & """" & ")"
    Range("R22").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$G$1:$G$65536," & """" & "沾化区" & """" & ")"
    Range("R23").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$G$1:$G$65536," & """" & "邹平市" & """" & ")"
    
    '全月求和
    Range("K11").Select
    Selection.Formula = "=sum(K3:K10)"
    Range("L11").Activate
    Selection.Formula = "=sum(L3:L10)"
    Range("Q11").Activate
    Selection.Formula = "=sum(Q3:Q10)"
    Range("R11").Activate
    Selection.Formula = "=sum(R3:R10)"
    
    Range("K24").Select
    Selection.Formula = "=sum(K16:K23)"
    Range("L24").Activate
    Selection.Formula = "=sum(L16:L23)"
    Range("Q24").Activate
    Selection.Formula = "=sum(Q16:Q23)"
    Range("R24").Activate
    Selection.Formula = "=sum(R16:R23)"
    
    
    Range("K3:V11").Select
    Selection.Copy
    Range("K3:V11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
       
    Range("K16:V24").Select
    Selection.Copy
    Range("K16:V24").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
  
   
     Cells.Select
     Cells.AutoFilter

End Function

Function 复制数据_日报_宽带装移机进展(HomeName, UseName, Day)
    Dim CopyTmp As String '文件夹位置
    
    If Day <= 16 Then
        CopyTmp = Chr(Asc("A") + Day - 1)
    Else
        CopyTmp = Chr(Asc("A") + Day - 15 - 1)
    End If
    
    '复制数据
    '一
    Windows(HomeName).Activate
    Range("L3:L11").Select
    Selection.Copy
    Windows(UseName).Activate
    Sheets("3月宽带装移机进展").Activate
    Range("K3:K11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    If Day <= 16 Then
        Range(CopyTmp & "18:" & CopyTmp & "26").Select
    Else
        Range(CopyTmp & "29:" & CopyTmp & "37").Select
    End If
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

    '二
    Windows(HomeName).Activate
    Range("R3:R11").Select
    Selection.Copy
    Windows(UseName).Activate
    Sheets("3月宽带装移机进展").Activate
    Range("M3:M11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    If Day <= 16 Then
        Range(CopyTmp & "41:" & CopyTmp & "49").Select
    Else
        Range(CopyTmp & "52:" & CopyTmp & "60").Select
    End If
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

    '三(七)
    Windows(HomeName).Activate
    Range("L16:L24").Select
    Selection.Copy
    Windows(UseName).Activate
    Sheets("有线宽带报表").Activate
    Range("G4:G12").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

    '四（八）
    Windows(HomeName).Activate
    Range("R16:R24").Select
    Selection.Copy
    Sheets("待装天数（新）").Activate
    Range("H13:H21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    
    Windows(UseName).Activate
    Sheets("有线宽带报表").Activate
    Range("H4:H12").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    
    '
    Windows(HomeName).Activate
    Sheets("新家客日报（模板）").Activate
    Range("T16:T24").Select
    Selection.Copy
    Sheets("待装天数（新）").Activate
    Range("G13:G21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

    Sheets("新家客日报（模板）").Activate
    Range("V16:V24").Select
    Selection.Copy
    Sheets("待装天数（新）").Activate
    Range("F13:F21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

    Windows(UseName).Activate
    Sheets("3月宽带装移机进展").Activate
    Range("R29:R37").Select
    Selection.Copy
    Range("J3:J11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

    Range("R52:R60").Select
    Selection.Copy
    Range("L3:L11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

    Range("J3:N11").Select
    Selection.Copy
    Range("B3:F11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

    
    
    


End Function

Function 处理_专用日报_有效带宽_三日预约量(InstallName, InstallNameNull, UseName, RunName, RunNameNull, Year, Month, Day)
    '装机量(在后面)
    Windows(InstallName).Activate
    Sheets(InstallNameNull).Activate
    Cells.Select
    Cells.AutoFilter
    ActiveWorkbook.Names.Add Name:="'" & InstallNameNull & "'!_FilterDatabase", RefersTo:="='" & InstallNameNull & "'!$A$1:$BW$4034", Visible:=False
    ActiveWindow.ScrollColumn = 16
    Range("A1:BW4034").AutoFilter Field:=35, Criteria1:=Array("已归档", "已开通"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & InstallNameNull & "'!_FilterDatabase", RefersTo:="='" & InstallNameNull & "'!$A$1:$BW$4034", Visible:=False
    Range("A1:BW4034").AutoFilter Field:=33
    ActiveWorkbook.Names.Add Name:="'" & InstallNameNull & "'!_FilterDatabase", RefersTo:="='" & InstallNameNull & "'!$A$1:$BW$4034", Visible:=False
    ActiveWindow.ScrollColumn = 3
    Range("A1:BW4034").AutoFilter Field:=5, Criteria1:=Array("家庭有线宽带"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & InstallNameNull & "'!_FilterDatabase", RefersTo:="='" & InstallNameNull & "'!$A$1:$BW$4034", Visible:=False
    ActiveWindow.ScrollColumn = 1
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("I:I").Select
    ActiveSheet.Paste

    Windows(UseName).Activate
    Sheets("有线宽带报表").Activate
    '赋值
    Range("I4").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$I$1:$I$65536," & """" & "滨北区" & """" & ")"
    Range("I5").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$I$1:$I$65536," & """" & "滨城区" & """" & ")"
    Range("I6").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$I$1:$I$65536," & """" & "博兴县" & """" & ")"
    Range("I7").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$I$1:$I$65536," & """" & "惠民县" & """" & ")"
    Range("I8").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$I$1:$I$65536," & """" & "无棣县" & """" & ")"
    Range("I9").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$I$1:$I$65536," & """" & "阳信县" & """" & ")"
    Range("I10").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$I$1:$I$65536," & """" & "沾化区" & """" & ")"
    Range("I11").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$I$1:$I$65536," & """" & "邹平市" & """" & ")"

    '处理运行中的数据
    Windows(RunName).Activate
    '新建工作表
    'Worksheets.Add before:=Sheets("20210309_新增量")
    Worksheets.Add before:=Sheets(RunNameNull)
    Sheets(RunNameNull).Activate
    '处理数据
    '今天
    Cells.Select
    Range("H1").Activate
    Cells.AutoFilter
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$BW$3625", Visible:=False
    Range("A1:BW3625").AutoFilter Field:=22, Operator:=xlFilterValues, Criteria2:=Array(2, Month & "/" & Day & "/" & Year) '今天
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$BW$3625", Visible:=False
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("A:A").Select
    ActiveSheet.Paste
    '明天
    Sheets(RunNameNull).Activate
    Range("H1").Activate
    Cells.AutoFilter
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$BW$3625", Visible:=False
    Range("A1:BW3625").AutoFilter Field:=22, Operator:=xlFilterValues, Criteria2:=Array(2, Month & "/" & Day + 1 & "/" & Year)
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$BW$3625", Visible:=False
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("B:B").Select
    ActiveSheet.Paste
    
    '后天
    Sheets(RunNameNull).Activate
    Range("H1").Activate
    Cells.AutoFilter
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$BW$3625", Visible:=False
    Range("A1:BW3625").AutoFilter Field:=22, Operator:=xlFilterValues, Criteria2:=Array(2, Month & "/" & Day + 2 & "/" & Year)
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$BW$3625", Visible:=False
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("C:C").Select
    ActiveSheet.Paste
    
    '赋值
    Windows(UseName).Activate
    Sheets("有线宽带报表").Activate
    '赋值
    '今天
    Range("B4").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$A$1:$A$65536," & """" & "滨北区" & """" & ")"
    Range("B5").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$A$1:$A$65536," & """" & "滨城区" & """" & ")"
    Range("B6").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$A$1:$A$65536," & """" & "博兴县" & """" & ")"
    Range("B7").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$A$1:$A$65536," & """" & "惠民县" & """" & ")"
    Range("B8").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$A$1:$A$65536," & """" & "无棣县" & """" & ")"
    Range("B9").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$A$1:$A$65536," & """" & "阳信县" & """" & ")"
    Range("B10").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$A$1:$A$65536," & """" & "沾化区" & """" & ")"
    Range("B11").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$A$1:$A$65536," & """" & "邹平市" & """" & ")"
   '明天
    Range("C4").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$B$1:$B$65536," & """" & "滨北区" & """" & ")"
    Range("C5").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$B$1:$B$65536," & """" & "滨城区" & """" & ")"
    Range("C6").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$B$1:$B$65536," & """" & "博兴县" & """" & ")"
    Range("C7").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$B$1:$B$65536," & """" & "惠民县" & """" & ")"
    Range("C8").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$B$1:$B$65536," & """" & "无棣县" & """" & ")"
    Range("C9").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$B$1:$B$65536," & """" & "阳信县" & """" & ")"
    Range("C10").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$B$1:$B$65536," & """" & "沾化区" & """" & ")"
    Range("C11").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$B$1:$B$65536," & """" & "邹平市" & """" & ")"
   '后天
    Range("D4").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$C$1:$C$65536," & """" & "滨北区" & """" & ")"
    Range("D5").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$C$1:$C$65536," & """" & "滨城区" & """" & ")"
    Range("D6").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$C$1:$C$65536," & """" & "博兴县" & """" & ")"
    Range("D7").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$C$1:$C$65536," & """" & "惠民县" & """" & ")"
    Range("D8").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$C$1:$C$65536," & """" & "无棣县" & """" & ")"
    Range("D9").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$C$1:$C$65536," & """" & "阳信县" & """" & ")"
    Range("D10").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$C$1:$C$65536," & """" & "沾化区" & """" & ")"
    Range("D11").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$C$1:$C$65536," & """" & "邹平市" & """" & ")"

    '求和
    Range("B12").Select
    Selection.Formula = "=sum(B4:B11)"
    Range("C12").Select
    Selection.Formula = "=sum(C4:C11)"
    Range("D12").Select
    Selection.Formula = "=sum(D4:D11)"
    Range("I12").Select
    Selection.Formula = "=sum(I4:I11)"
    
    '转为数字格式
    
    Range("B4:D11").Select
    Selection.Copy
    Range("B4:D11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    
    Range("I4:I11").Select
    Selection.Copy
    Range("I4:I11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
   

End Function

Function 处理_专用日报_待装工单(DispatchName, ExpansionName, AppointmentName, RunName, RunNameNull, UseName, HomeName)
    Dim EndH, MyH As Long
'
    Windows(RunName).Activate
    Sheets(RunNameNull).Activate
    Cells.Select
    Range("D1").Activate
    Cells.AutoFilter
    Cells.AutoFilter
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='20210315_运行中'!$A$1:$BX$3579", Visible:=False




    MyH = 0
    EndH = 0
    '处理三个池子
    Windows(AppointmentName).Activate
    Worksheets.Add before:=Sheets("待调度工单")
    Sheets("待调度工单").Activate

    MyH = Range("G65536").End(xlUp).Row - 1
    EndH = EndH + MyH
    Range("G2:G" & (MyH + 1)).Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Range("A1:A" & EndH).Select
    ActiveSheet.Paste
    Range("B1").Formula = "预约池"
    Range("B1").Select
    Selection.AutoFill Destination:=Range("B1:B" & EndH), Type:=xlFillDefault


    Windows(ExpansionName).Activate
    Sheets("待调度工单").Activate
    MyH = Range("G65536").End(xlUp).Row - 1
    Range("G2:G" & (MyH + 1)).Select
    Selection.Copy

    Windows(AppointmentName).Activate
    Sheets("Sheet1").Activate
    Range("A" & EndH + 1 & ":A" & (EndH + MyH)).Select

    ActiveSheet.Paste
    Range("B" & EndH + 1).Formula = "扩容池"
    Range("B" & EndH + 1).Select
    Selection.AutoFill Destination:=Range("B" & (EndH + 1) & ":B" & (EndH + MyH)), Type:=xlFillDefault
    EndH = EndH + MyH
'
'
    Windows(DispatchName).Activate
    Sheets("待调度工单").Activate
    MyH = Range("E65536").End(xlUp).Row - 1
    Range("E2:E" & (MyH + 1)).Select
    Selection.Copy

    Windows(AppointmentName).Activate
    Sheets("Sheet1").Activate
    Range("A" & EndH + 1 & ":A" & (EndH + MyH)).Select
    ActiveSheet.Paste
    Range("B" & EndH + 1).Formula = "待装池"
    Range("B" & EndH + 1).Select
    Selection.AutoFill Destination:=Range("B" & (EndH + 1) & ":B" & (EndH + MyH)), Type:=xlFillDefault
'    EndH = EndH + MyH


'    Cells.Select
'    Range("BD1").Activate
'    Cells.AutoFilter
'    ActiveWorkbook.Names.Add Name:="'20210309_运行中'!_FilterDatabase", RefersTo:="='20210309_运行中'!$A$1:$BW$3625", Visible:=False
'    ActiveWindow.ScrollColumn = 1
'    Range("A1:BW3625").AutoFilter Field:=5, Criteria1:=Array("驳回1"), Operator:=xlFilterValues
'    ActiveWorkbook.Names.Add Name:="'20210309_运行中'!_FilterDatabase", RefersTo:="='20210309_运行中'!$A$1:$BW$3625", Visible:=False
'    ActiveWindow.ScrollColumn = 32
'    Range("A1:BW3625").AutoFilter Field:=36, Criteria1:=Array("前台营业确认"), Operator:=xlFilterValues
'    ActiveWorkbook.Names.Add Name:="'20210309_运行中'!_FilterDatabase", RefersTo:="='20210309_运行中'!$A$1:$BW$3625", Visible:=False
'    Range("AK939").Select
'    Selection.Formula = "驳回"
'    Range("AK939").Select
'    Selection.AutoFill Destination:=Range("AK939:AK941"), Type:=xlFillDefault
'    Range("AK939:AK941").Select
'    Selection.AutoFill Destination:=Range("AK939:AK945"), Type:=xlFillDefault
'    Range("AK939:AK945").Select
'    With Range("AK939:AK3468")
'        .FillDown
'        .Select
'    End With
'    Selection.AutoFill Destination:=Range("AK939:AK3625"), Type:=xlFillDefault
'    Range("AK939:AK3625").Select

    '处理运行中的数据
    Windows(RunName).Activate
    Sheets(RunNameNull).Activate
    '插入行
    Columns("AK:AK").Select
    Selection.Insert Shift:=xlShiftToRight
    
'    '循环处理
        EndH = Range("A65536").End(xlUp).Row
'      '遍历表格
'      '判断是否为空，为空则跳过
'       ' EndH = 20
        For i = EndH To 2 Step -1
        XX = Range("AJ" & i).Value
        If IsEmpty(Cells(i, 5)) Then
        
        ElseIf InStr(XX, "前台营业确认") <> 0 Then '
          Range("AK" & i).Formula = "驳回"
        ElseIf Not InStr(XX, "资源回填") <> 0 And (InStr(XX, "设备回收/线路拆除") <> 0 Or InStr(XX, "施工预约") <> 0 Or InStr(XX, "现场开通") <> 0 Or InStr(XX, "异常处理") <> 0) Then     '
            Range("AK" & i).Formula = "待装池"
        ElseIf InStr(XX, "资源回填") <> 0 Then

        Else
        '套公式
            Range("AK" & i).Select
            Selection.Formula = "=VLOOKUP(Q" & i & "&"""",[" & AppointmentName & "]Sheet1!$A:$B,2,FALSE)"
            'Selection.Formula = "=VLOOKUP(Q532,[20210309_预约池.xls]Sheet1!$A:$B,2,FALSE)"
            If IsError(Cells(i, 37)) Then
                Range("AK" & i).Formula = "驳回"
            Else
            
            End If
        End If
     Next

    '处理数据
    '新建工作表
    Worksheets.Add before:=Sheets(RunNameNull)
    Sheets(RunNameNull).Activate
    '整体筛选
    Cells.Select
    Range("BA1").Activate
    Cells.AutoFilter
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CB$3625", Visible:=False
    Application.DisplayAlerts = False
    Application.DisplayAlerts = True
    ActiveWindow.ScrollColumn = 1
    Range("A1:CB3625").AutoFilter Field:=5, Criteria1:=Array("家庭有线宽带"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CB$3625", Visible:=False
    ActiveWindow.ScrollColumn = 29
    Range("A1:CB3625").AutoFilter Field:=36, Criteria1:=Array("前台营业确认", "设备回收/线路拆除", "施工预约", "现场开通", "异常处理", "装机调度"), Operator:=xlFilterValues
    
    '分开导出数据
    '扩容池
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CB$3625", Visible:=False
    Range("A1:CB3625").AutoFilter Field:=37, Criteria1:=Array("扩容池"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CB$3625", Visible:=False
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("O:O").Select
    ActiveSheet.Paste
    '预约池
    Sheets(RunNameNull).Activate
     ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CB$3625", Visible:=False
    Range("A1:CB3625").AutoFilter Field:=37, Criteria1:=Array("预约池"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CB$3625", Visible:=False
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("P:P").Select
    ActiveSheet.Paste
    '驳回
    Sheets(RunNameNull).Activate
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CB$3625", Visible:=False
    Range("A1:CB3625").AutoFilter Field:=37, Criteria1:=Array("驳回"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CB$3625", Visible:=False
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("Q:Q").Select
    ActiveSheet.Paste
    '待装池
    Sheets(RunNameNull).Activate
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CB$3625", Visible:=False
    Range("A1:CB3625").AutoFilter Field:=37, Criteria1:=Array("待装池"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CB$3625", Visible:=False
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("R:R").Select
    ActiveSheet.Paste
    
    '赋值
    Windows(UseName).Activate
    Sheets("待装工单").Activate
    
    '扩容池
    Range("D3").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$O$1:$O$65536," & """" & "滨北区" & """" & ")"
    Range("D4").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$O$1:$O$65536," & """" & "滨城区" & """" & ")"
    Range("D5").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$O$1:$O$65536," & """" & "博兴县" & """" & ")"
    Range("D6").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$O$1:$O$65536," & """" & "惠民县" & """" & ")"
    Range("D7").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$O$1:$O$65536," & """" & "无棣县" & """" & ")"
    Range("D8").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$O$1:$O$65536," & """" & "阳信县" & """" & ")"
    Range("D9").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$O$1:$O$65536," & """" & "沾化区" & """" & ")"
    Range("D10").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$O$1:$O$65536," & """" & "邹平市" & """" & ")"
    '预约池
    Range("E3").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$P$1:$P$65536," & """" & "滨北区" & """" & ")"
    Range("E4").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$P$1:$P$65536," & """" & "滨城区" & """" & ")"
    Range("E5").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$P$1:$P$65536," & """" & "博兴县" & """" & ")"
    Range("E6").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$P$1:$P$65536," & """" & "惠民县" & """" & ")"
    Range("E7").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$P$1:$P$65536," & """" & "无棣县" & """" & ")"
    Range("E8").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$P$1:$P$65536," & """" & "阳信县" & """" & ")"
    Range("E9").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$P$1:$P$65536," & """" & "沾化区" & """" & ")"
    Range("E10").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$P$1:$P$65536," & """" & "邹平市" & """" & ")"
      '驳回
    Range("F3").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$Q$1:$Q$65536," & """" & "滨北区" & """" & ")"
    Range("F4").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$Q$1:$Q$65536," & """" & "滨城区" & """" & ")"
    Range("F5").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$Q$1:$Q$65536," & """" & "博兴县" & """" & ")"
    Range("F6").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$Q$1:$Q$65536," & """" & "惠民县" & """" & ")"
    Range("F7").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$Q$1:$Q$65536," & """" & "无棣县" & """" & ")"
    Range("F8").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$Q$1:$Q$65536," & """" & "阳信县" & """" & ")"
    Range("F9").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$Q$1:$Q$65536," & """" & "沾化区" & """" & ")"
    Range("F10").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$Q$1:$Q$65536," & """" & "邹平市" & """" & ")"
      '待装
    Range("G3").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$R$1:$R$65536," & """" & "滨北区" & """" & ")"
    Range("G4").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$R$1:$R$65536," & """" & "滨城区" & """" & ")"
    Range("G5").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$R$1:$R$65536," & """" & "博兴县" & """" & ")"
    Range("G6").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$R$1:$R$65536," & """" & "惠民县" & """" & ")"
    Range("G7").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$R$1:$R$65536," & """" & "无棣县" & """" & ")"
    Range("G8").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$R$1:$R$65536," & """" & "阳信县" & """" & ")"
    Range("G9").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$R$1:$R$65536," & """" & "沾化区" & """" & ")"
    Range("G10").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$R$1:$R$65536," & """" & "邹平市" & """" & ")"
    
    '求和
    Range("D11").Select
    Selection.Formula = "=sum(D3:D10)"
    Range("E11").Select
    Selection.Formula = "=sum(E3:E10)"
    Range("F11").Select
    Selection.Formula = "=sum(F3:F10)"
    Range("G11").Select
    Selection.Formula = "=sum(G3:G10)"
    
    Range("B3").Select
    Selection.Formula = "=sum(C3:G3)"
    Range("B4").Select
    Selection.Formula = "=sum(C4:G4)"
    Range("B5").Select
    Selection.Formula = "=sum(C5:G5)"
    Range("B6").Select
    Selection.Formula = "=sum(C6:G6)"
    Range("B7").Select
    Selection.Formula = "=sum(C7:G7)"
    Range("B8").Select
    Selection.Formula = "=sum(C8:G8)"
    Range("B9").Select
    Selection.Formula = "=sum(C9:G9)"
    Range("B10").Select
    Selection.Formula = "=sum(C10:G10)"
    Range("B11").Select
    Selection.Formula = "=sum(C11:G11)"
    
    Range("D3:G11").Select
    Selection.Copy
    Range("D3:G11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
  
    
    
    
    '复制数据
    '选中数据
    Range("G3:G11").Select
    Selection.Copy
    '复制数据
    Sheets("有线宽带报表").Activate
    Range("F4:F12").Select
    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    
    Windows(HomeName).Activate
    Sheets("待装天数模板").Activate
    Range("C3:C11").Select
    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    Sheets("待装天数（新）").Activate
    Range("B13:B21").Select
    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    
    '
    Sheets("新家客日报（模板）").Activate
    Range("R16:R24").Select
    Selection.Copy
   
    Sheets("待装天数（新）").Activate
    Range("H13:H21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

 
    Range("I13:I21").Select
    Selection.Copy
    Sheets("待装天数模板").Activate
    Range("B3:B11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

   



End Function
Function 统计_新家客_待装天数_ImsOtt(RunName, RunNameNull, HomeName, UseName)


    Windows(RunName).Activate
    Sheets(RunNameNull).Activate
    
    
    Cells.Select
    Cells.AutoFilter
    Cells.AutoFilter
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CM$3625", Visible:=False
    Range("A1:CM3625").AutoFilter Field:=7, Criteria1:=Array("IMS"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CM$3625", Visible:=False
    ActiveWindow.ScrollColumn = 27
    Range("A1:CM3625").AutoFilter Field:=37, Criteria1:=Array("待装池"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CM$3625", Visible:=False
    ActiveWindow.ScrollColumn = 1
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("T:T").Select
    ActiveSheet.Paste
    
    Sheets(RunNameNull).Activate
    Range("A1:CM3625").AutoFilter Field:=7
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="&R unNameNull &='""'!$A$1:$CM$3625", Visible:=False
    Range("A1:CM3625").AutoFilter Field:=6, Criteria1:=Array("OTT"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CM$3625", Visible:=False
    ActiveWindow.ScrollColumn = 4
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("U:U").Select
    ActiveSheet.Paste
    
    '赋值
    Windows(HomeName).Activate
    Sheets("待装天数模板").Activate
    
    'IMS
    Range("E3").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$T$1:$T$65536," & """" & "滨北区" & """" & ")"
    Range("E4").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$T$1:$T$65536," & """" & "滨城区" & """" & ")"
    Range("E5").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$T$1:$T$65536," & """" & "博兴县" & """" & ")"
    Range("E6").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$T$1:$T$65536," & """" & "惠民县" & """" & ")"
    Range("E7").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$T$1:$T$65536," & """" & "无棣县" & """" & ")"
    Range("E8").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$T$1:$T$65536," & """" & "阳信县" & """" & ")"
    Range("E9").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$T$1:$T$65536," & """" & "沾化区" & """" & ")"
    Range("E10").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$T$1:$T$65536," & """" & "邹平市" & """" & ")"
      'OTT
    Range("D3").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$U$1:$U$65536," & """" & "滨北区" & """" & ")"
    Range("D4").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$U$1:$U$65536," & """" & "滨城区" & """" & ")"
    Range("D5").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$U$1:$U$65536," & """" & "博兴县" & """" & ")"
    Range("D6").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$U$1:$U$65536," & """" & "惠民县" & """" & ")"
    Range("D7").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$U$1:$U$65536," & """" & "无棣县" & """" & ")"
    Range("D8").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$U$1:$U$65536," & """" & "阳信县" & """" & ")"
    Range("D9").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$U$1:$U$65536," & """" & "沾化区" & """" & ")"
    Range("D10").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$U$1:$U$65536," & """" & "邹平市" & """" & ")"
    '求和
    
    Range("D11").Select
    Selection.Formula = "=sum(D3:D10)"
    Range("E11").Select
    Selection.Formula = "=sum(E3:E10)"
       
    Range("D3:E11").Select
    Selection.Copy
    Range("D3:E11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    
    
    '复制
    '扩容池
    Windows(UseName).Activate
    Sheets("待装工单").Activate
    Range("D3:D11").Select
    Selection.Copy
    Windows(HomeName).Activate
    Sheets("待装天数模板").Activate
    Range("G3:G11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    
    '可待装量
    Sheets("待装天数（新）").Activate
    Range("B13:B21").Select
    Selection.Copy
    Windows(UseName).Activate
    Sheets("装机进度通报").Activate
    Range("K4:K12").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    '三日平均
    Windows(HomeName).Activate
    Sheets("待装天数（新）").Activate
    Range("E13:E21").Select
    Selection.Copy
    Windows(UseName).Activate
    Range("L4:L12").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

    Windows(HomeName).Activate
    Range("I13:I21").Select
    Selection.Copy
    Windows(UseName).Activate
    Range("M4:M12").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

    '统计
    Range("K12:M12").Select
    Selection.Copy
    Range("E12:G12").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    
     For i = 11 To 4 Step -1
        XX = Range("A" & i).Value
        For y = 11 To 4 Step -1
           YY = Range("J" & y).Value
           If InStr(XX, YY) <> 0 Then '@就是要找到指定的特定字符 可以改成你指定的其他字符
             '  MsgBox (XX & "结束" & YY & "DD" & y & ":" & i)
               Range("K" & i & ":M" & i).Select
               Selection.Copy
               Range("E" & y & ":G" & y).Select
               Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
           End If
        Next
     Next


     
     
    

    
    
    
    
    
    


End Function
Function 复制_专用日报_待装天数模板(HomeName, UseName)
    Windows(HomeName).Activate
    Sheets("待装天数模板").Activate
    Range("B3:G11").Select
    Selection.Copy
    
    Windows(UseName).Activate
    Sheets("待装天数模板").Activate
    Range("B3:G11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    
    
    

End Function

Function 新增量_装机量_去重(AddName, AddNameNull, InstallName, InstallNameNull)

    Windows(AddName).Activate
    Sheets(AddNameNull).Activate
    Cells.Select
    Range("A291").Activate
    Range("A2:BW7263").Select
    ActiveWindow.ScrollRow = 1
    Range("A1:BW7263").RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75), Header:=xlYes
    Cells.Select
    Range("A291").Activate
    
    Windows(InstallName).Activate
    Sheets(InstallNameNull).Activate
    
    Cells.Select
    Range("A2:BW6902").Select
    Range("A1:BW6902").RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75), Header:=xlYes
    Cells.Select
    
    
    
End Function
    

Private Sub CommandButton3_Click()


End Sub


Private Sub CommandButton2_Click()

End Sub
