VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4365
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    '���й����н�ֹ���
    TextBox1.Enabled = False
    CommandButton1.Enabled = False
    CommandButton2.Enabled = False
    

    
    '�������
    '�װ���Ϣ
    Dim NegativeName1, NegativePath1 As String '�¼ҿ��ձ�_ģ���ļ�λ�ã��ļ�����
    Dim NegativeName2, NegativePath2 As String '�ձ�ר��ģ��_ģ���ļ�λ�ã��ļ�����
    
    Dim HomeName, HomePath As String '�ҿʹ����ļ�λ�ã��ļ�����
    Dim UseName, UsePath As String '�ձ�ר��ģ�崦���ļ�λ�ã��ļ�����
    'ԭʼλ����Ϣ
    Dim AddName, AddPath, AddNameNull As String '������λ����Ϣ
    Dim InstallName, InstallPath, InstallNameNull As String 'װ����λ����Ϣ
    Dim RunName, RunPath, RunNameNull As String 'װ����λ����Ϣ
    '��������
    Dim DispatchName, DispatchPath, DispatchNameNull As String '��װ��
    Dim ExpansionName, ExpansionPath, ExpansionNameNull As String '���ݳ�
    Dim AppointmentName, AppointmentPath, AppointmentNameNull As String 'ԤԼ��
    
    

    
    Dim pth As String '�ļ���λ��
    
    Dim Year, Month, Day As String '�ļ���λ��
    
    
    
    '������ֵ
    
    pth = ThisWorkbook.Path
    pth = Left(pth, Len(pth) - 2)
    
    HomeName = "�¼ҿ��ձ�" & TextBox1.Text & ".xls"
    HomePath = pth & "����\" & HomeName
    
    UseName = "�ձ�ר��ģ��" & TextBox1.Text & ".xls"
    UsePath = pth & "����\" & UseName
    
    NegativeName1 = "�¼ҿ��ձ�.xls"
    NegativePath1 = pth & "�װ�\" & NegativeName1
    NegativeName2 = "�ձ�ר��ģ��.xls"
    NegativePath2 = pth & "�װ�\" & NegativeName2
    
    AddNameNull = TextBox1.Text & "_������"
    AddName = AddNameNull & ".csv"
    AddPath = pth & "����\" & AddName
    
    InstallNameNull = TextBox1.Text & "_װ����"
    InstallName = InstallNameNull & ".csv"
    InstallPath = pth & "����\" & InstallName
    
    RunNameNull = TextBox1.Text & "_������"
    RunName = RunNameNull & ".csv"
    RunPath = pth & "����\" & RunName
    
    DispatchNameNull = TextBox1.Text & "_��װ��"
    DispatchName = DispatchNameNull & ".xls"
    DispatchPath = pth & "����\" & DispatchName
    
    ExpansionNameNull = TextBox1.Text & "_���ݳ�"
    ExpansionName = ExpansionNameNull & ".xls"
    ExpansionPath = pth & "����\" & ExpansionName
    
    AppointmentNameNull = TextBox1.Text & "_ԤԼ��"
    AppointmentName = AppointmentNameNull & ".xls"
    AppointmentPath = pth & "����\" & AppointmentName
    
    Year = TextBox1.Text / 10000
    Year = Format(Year, 0)
    Month = TextBox1.Text Mod 10000
    Month = Month / 100
    Month = Format(Month, 0)
    Day = TextBox1.Text Mod 100
   
  
   
    
   'MsgBox (Month & "/" & Day - 1 & "/" & Year)
    
   
    
    
'

'    '�ж��Ƿ�����ļ�

    If Dir(NegativePath1, vbDirectory) = vbNullString Then
        MsgBox ("δ�ҵ�ģ���ļ�_�¼ҿ��ձ���")
        TextBox1.Enabled = True
        CommandButton1.Enabled = True
        CommandButton2.Enabled = True
        Exit Sub
    End If
    If Dir(NegativePath2, vbDirectory) = vbNullString Then
        MsgBox ("δ�ҵ�ģ���ļ�_�ձ�ר��ģ�壡")
        TextBox1.Enabled = True
        CommandButton1.Enabled = True
        CommandButton2.Enabled = True
        Exit Sub
    End If
    
    If Dir(pth & "ԭʼ����\" & "������\" & TextBox1.Text & ".csv", vbDirectory) = vbNullString Then
        MsgBox ("δ�ҵ��������ļ���")
        TextBox1.Enabled = True
        CommandButton1.Enabled = True
        CommandButton2.Enabled = True
        Exit Sub
    End If
    If Dir(pth & "ԭʼ����\" & "װ����\" & TextBox1.Text & ".csv", vbDirectory) = vbNullString Then
        MsgBox ("δ�ҵ�װ�����ļ���")
        TextBox1.Enabled = True
        CommandButton1.Enabled = True
        CommandButton2.Enabled = True
        Exit Sub
    End If
    If Dir(pth & "ԭʼ����\" & "������\" & TextBox1.Text & ".csv", vbDirectory) = vbNullString Then
        MsgBox ("δ�ҵ��������ļ���")
        TextBox1.Enabled = True
        CommandButton1.Enabled = True
        CommandButton2.Enabled = True
        Exit Sub
    End If
    
    If Dir(pth & "ԭʼ����\" & "��װ��\" & TextBox1.Text & ".xls", vbDirectory) = vbNullString Then
        MsgBox ("δ�ҵ���װ���ļ���")
        TextBox1.Enabled = True
        CommandButton1.Enabled = True
        CommandButton2.Enabled = True
        Exit Sub
    End If
    If Dir(pth & "ԭʼ����\" & "���ݳ�\" & TextBox1.Text & ".xls", vbDirectory) = vbNullString Then
        MsgBox ("δ�ҵ����ݳ��ļ���")
        TextBox1.Enabled = True
        CommandButton1.Enabled = True
        CommandButton2.Enabled = True
        Exit Sub
    End If
    If Dir(pth & "ԭʼ����\" & "ԤԼ��\" & TextBox1.Text & ".xls", vbDirectory) = vbNullString Then
        MsgBox ("δ�ҵ�ԤԼ���ļ���")
        TextBox1.Enabled = True
        CommandButton1.Enabled = True
        CommandButton2.Enabled = True
        Exit Sub
    End If

    '���Ƶװ��ԭʼ����
    FileCopy NegativePath1, HomePath   '
    FileCopy NegativePath2, UsePath    '
    FileCopy pth & "ԭʼ����\" & "������\" & TextBox1.Text & ".csv", AddPath
    FileCopy pth & "ԭʼ����\" & "װ����\" & TextBox1.Text & ".csv", InstallPath
    FileCopy pth & "ԭʼ����\" & "������\" & TextBox1.Text & ".csv", RunPath
    FileCopy pth & "ԭʼ����\" & "��װ��\" & TextBox1.Text & ".xls", DispatchPath
    FileCopy pth & "ԭʼ����\" & "���ݳ�\" & TextBox1.Text & ".xls", ExpansionPath
    FileCopy pth & "ԭʼ����\" & "ԤԼ��\" & TextBox1.Text & ".xls", AppointmentPath


   ' �������ļ�
    Workbooks.Open Filename:=HomePath, AddToMru:=True
    Workbooks.Open Filename:=UsePath, AddToMru:=True

    Workbooks.Open Filename:=AddPath, AddToMru:=True
    Workbooks.Open Filename:=InstallPath, AddToMru:=True
    Workbooks.Open Filename:=RunPath, AddToMru:=True
    Workbooks.Open Filename:=DispatchPath, AddToMru:=True
    Workbooks.Open Filename:=ExpansionPath, AddToMru:=True
    Workbooks.Open Filename:=AppointmentPath, AddToMru:=True

    

    Call ��ʼ��(HomeName, UseName, Year, Month, Day)
    Call ������_װ����_ȥ��(AddName, AddNameNull, InstallName, InstallNameNull)
    Call ��������_������_װ����(AddName, AddNameNull, HomeName, InstallName, InstallNameNull, Year, Month, Day)
    Call ��������_�ձ�_���װ�ƻ���չ(HomeName, UseName, Day)
    Call ����_ר���ձ�_��Ч����_����ԤԼ��(InstallName, InstallNameNull, UseName, RunName, RunNameNull, Year, Month, Day)
    Call ����_ר���ձ�_��װ����(DispatchName, ExpansionName, AppointmentName, RunName, RunNameNull, UseName, HomeName)
    Call ͳ��_�¼ҿ�_��װ����_ImsOtt(RunName, RunNameNull, HomeName, UseName)
    Call ����_ר���ձ�_��װ����ģ��(HomeName, UseName)
    
    
    '�ر��ļ�
    
    
    
    '�����ļ�
    
    
    '�ָ�ʹ��
    TextBox1.Enabled = True
    CommandButton1.Enabled = True
    CommandButton2.Enabled = True
     MsgBox "����"
     
    
    
    
    
    
    
    
End Sub
Function ��ʼ��(HomeName, UseName, Year, Month, Day)
    Call �ҿ�_�ҿ�_��ʼ��(HomeName, Month, Day)
    Call �ҿ�_��װ����_��__��ʼ��(HomeName, Day)
    Call �ձ�ר��_���߿��_��ʼ��(UseName, Month, Day)
End Function
Function �ҿ�_�ҿ�_��ʼ��(HomeName, Month, Day)
    Windows(HomeName).Activate
    Sheets("�¼ҿ��ձ���ģ�壩").Activate
    
    Range("K1:N11").Select
    Selection.Copy
    Range("M1:P11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    Range("K1").Formula = Month & "��" & Day - 1 & "������������"
    
    Range("Q1:T11").Select
    Selection.Copy
    Range("S1:V11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    Range("Q1").Formula = Month & "��" & Day - 1 & "����ɹ�����"

    Range("K14:N24").Select
    Selection.Copy
    Range("M14:P24").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    Range("K14").Formula = Month & "��" & Day - 1 & "������װ���ƻ�������"
    
    Range("Q14:T24").Select
    Selection.Copy
    Range("S14:V24").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    Range("Q14").Formula = Month & "��" & Day - 1 & "�����װ���ƻ�������"

    Range("K3:L11").Select
    Application.CutCopyMode = False
    
End Function
Function �ձ�ר��_���߿��_��ʼ��(UseName, Month, Day)
    Windows(UseName).Activate
    Sheets("���߿������").Activate
    Range("B2").Formula = Month & "." & Day & "��ԤԼ��"
    Range("C2").Formula = Month & "." & Day + 1 & "��ԤԼ��"
    Range("D2").Formula = Month & "." & Day + 2 & "��ԤԼ��"

    

End Function
Function �ҿ�_��װ����_��__��ʼ��(HomeName, Day)
    Windows(HomeName).Activate
    Sheets("��װ�������£�").Activate
    
    Range("F11").Formula = Day - 3 & "�������"
    Range("G11").Formula = Day - 2 & "�������"
    Range("H11").Formula = Day - 1 & "�������"
    
    
    
End Function
Function ��������_������_װ����(AddName, AddNameNull, HomeName, InstallName, InstallNameNull, Year, Month, Day)

    'ɾ�� �ض��ַ�
    Windows(AddName).Activate
    Dim EndH, i As Long
    EndH = Range("B65536").End(xlUp).Row
    
    For i = EndH To 1 Step -1
     XX = Range("B" & i).Value
     If InStr(XX, "������������") <> 0 Then '@����Ҫ�ҵ�ָ�����ض��ַ� ���Ըĳ���ָ���������ַ�
          Range("B" & i).EntireRow.Delete 'E����ָ��Ҫ���ҵ���.��Ұ�ĳ���ָ����������
     End If
     Next

    '�����������ļ�
    '�½�������
    Worksheets.Add before:=Sheets(AddNameNull)
    Sheets(AddNameNull).Activate
    '������
    Cells.Select
    Cells.AutoFilter
    ActiveWorkbook.Names.Add Name:="'" & AddNameNull & "'!_FilterDatabase", RefersTo:="='" & AddNameNull & "'!$A$1:$BW$65536", Visible:=False
    Range("N4").Select
    ActiveWindow.ScrollColumn = 26
    Columns("AI:AI").Select
    Range("A1:BW65536").AutoFilter Field:=35, Criteria1:=Array("�ѹ鵵", "�ѿ�ͨ", "ִ����"), Operator:=xlFilterValues
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
    '��һ��
    Range("A1:BW65536").AutoFilter Field:=4, Criteria1:=Array("ҵ���ںϿ�������"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & AddNameNull & "'!_FilterDatabase", RefersTo:="='" & AddNameNull & "'!$A$1:$BW$65536", Visible:=False
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("A:A").Select
    ActiveSheet.Paste
    Sheets(AddNameNull).Activate
    Range("A1:BW65536").AutoFilter Field:=4
    ActiveWorkbook.Names.Add Name:="'" & AddNameNull & "'!_FilterDatabase", RefersTo:="='" & AddNameNull & "'!$A$1:$BW$65536", Visible:=False
    '�ڶ���
    Range("A1:BW65536").AutoFilter Field:=4, Criteria1:=Array("ҵ���ںϿ�������"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & AddNameNull & "'!_FilterDatabase", RefersTo:="='" & AddNameNull & "'!$A$1:$BW$65536", Visible:=False
    Range("A1:BW65536").AutoFilter Field:=5, Criteria1:=Array("��ͥ���߿��"), Operator:=xlFilterValues
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
    '���ĸ�
    Range("A1:BW65536").AutoFilter Field:=5, Criteria1:=Array("��ͥ���߿��"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & AddNameNull & "'!_FilterDatabase", RefersTo:="='" & AddNameNull & "'!$A$1:$BW$65536", Visible:=False
    Selection.Copy
    Range("K4200").Select
    Sheets("Sheet1").Activate
    Columns("G:G").Select
    ActiveSheet.Paste

    '��ֵ
    Windows(HomeName).Activate
    Sheets("�¼ҿ��ձ���ģ�壩").Activate
    '��һ��
    Range("K3").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$A$1:$A$65536," & """" & "������" & """" & ")"
    Range("K4").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$A$1:$A$65536," & """" & "������" & """" & ")"
    Range("K5").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$A$1:$A$65536," & """" & "������" & """" & ")"
    Range("K6").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$A$1:$A$65536," & """" & "������" & """" & ")"
    Range("K7").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$A$1:$A$65536," & """" & "�����" & """" & ")"
    Range("K8").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$A$1:$A$65536," & """" & "������" & """" & ")"
    Range("K9").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$A$1:$A$65536," & """" & "մ����" & """" & ")"
    Range("K10").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$A$1:$A$65536," & """" & "��ƽ��" & """" & ")"
    '�ڶ���
    Range("L3").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$C$1:$C$65536," & """" & "������" & """" & ")"
    Range("L4").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$C$1:$C$65536," & """" & "������" & """" & ")"
    Range("L5").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$C$1:$C$65536," & """" & "������" & """" & ")"
    Range("L6").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$C$1:$C$65536," & """" & "������" & """" & ")"
    Range("L7").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$C$1:$C$65536," & """" & "�����" & """" & ")"
    Range("L8").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$C$1:$C$65536," & """" & "������" & """" & ")"
    Range("L9").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$C$1:$C$65536," & """" & "մ����" & """" & ")"
    Range("L10").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$C$1:$C$65536," & """" & "��ƽ��" & """" & ")"
     '������
    Range("K16").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$E$1:$E$65536," & """" & "������" & """" & ")"
    Range("K17").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$E$1:$E$65536," & """" & "������" & """" & ")"
    Range("K18").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$E$1:$E$65536," & """" & "������" & """" & ")"
    Range("K19").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$E$1:$E$65536," & """" & "������" & """" & ")"
    Range("K20").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$E$1:$E$65536," & """" & "�����" & """" & ")"
    Range("K21").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$E$1:$E$65536," & """" & "������" & """" & ")"
    Range("K22").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$E$1:$E$65536," & """" & "մ����" & """" & ")"
    Range("K23").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$E$1:$E$65536," & """" & "��ƽ��" & """" & ")"
    '�ڶ���
    Range("L16").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$G$1:$G$65536," & """" & "������" & """" & ")"
    Range("L17").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$G$1:$G$65536," & """" & "������" & """" & ")"
    Range("L18").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$G$1:$G$65536," & """" & "������" & """" & ")"
    Range("L19").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$G$1:$G$65536," & """" & "������" & """" & ")"
    Range("L20").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$G$1:$G$65536," & """" & "�����" & """" & ")"
    Range("L21").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$G$1:$G$65536," & """" & "������" & """" & ")"
    Range("L22").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$G$1:$G$65536," & """" & "մ����" & """" & ")"
    Range("L23").Select
    Selection.Formula = "=COUNTIFS([" & AddName & "]Sheet1!$G$1:$G$65536," & """" & "��ƽ��" & """" & ")"
     Cells.Select
     Cells.AutoFilter

    'ѡ��װ����
    Windows(InstallName).Activate
    'ɾ���ض���
    EndH = Range("B65536").End(xlUp).Row

    For i = EndH To 1 Step -1
     XX = Range("B" & i).Value
     If InStr(XX, "������������") <> 0 Then '@����Ҫ�ҵ�ָ�����ض��ַ� ���Ըĳ���ָ���������ַ�
          Range("B" & i).EntireRow.Delete 'E����ָ��Ҫ���ҵ���.��Ұ�ĳ���ָ����������
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
    Range("A1:BW4034").AutoFilter Field:=35, Criteria1:=Array("�ѹ鵵", "�ѿ�ͨ"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & InstallNameNull & "'!_FilterDatabase", RefersTo:="='" & InstallNameNull & "'!$A$1:$BW$4034", Visible:=False
    Range("AG1").Select
    Range("A1:BW4034").AutoFilter Field:=33, Operator:=xlFilterValues, Criteria2:=Array(2, Month & "/" & Day - 1 & "/" & Year) '"3/8/2021"
    ActiveWorkbook.Names.Add Name:="'" & InstallNameNull & "'!_FilterDatabase", RefersTo:="='" & InstallNameNull & "'!$A$1:$BW$4034", Visible:=False
    ActiveWindow.ScrollColumn = 1

    Worksheets.Add before:=Sheets(InstallNameNull)
    Sheets(InstallNameNull).Activate
    '������
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("E:E").Select
    ActiveSheet.Paste
    Sheets(InstallNameNull).Activate
    '��һ��
    Range("A1:BW65536").AutoFilter Field:=4, Criteria1:=Array("ҵ���ںϿ�������"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & InstallNameNull & "'!_FilterDatabase", RefersTo:="='" & InstallNameNull & "'!$A$1:$BW$65536", Visible:=False
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("A:A").Select
    ActiveSheet.Paste
    Sheets(InstallNameNull).Activate
     '�ڶ���
    Range("A1:BW65536").AutoFilter Field:=5, Criteria1:=Array("��ͥ���߿��"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & InstallNameNull & "'!_FilterDatabase", RefersTo:="='" & InstallNameNull & "'!$A$1:$BW$65536", Visible:=False
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("C:C").Select
    ActiveSheet.Paste
    Sheets(InstallNameNull).Activate
     '���ĸ�
    Range("A1:BW65536").AutoFilter Field:=4
    ActiveWorkbook.Names.Add Name:="'" & InstallNameNull & "'!_FilterDatabase", RefersTo:="='" & InstallNameNull & "'!$A$1:$BW$65536", Visible:=False
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("G:G").Select
    ActiveSheet.Paste
    Sheets(InstallNameNull).Activate

    '��ֵ
    Windows(HomeName).Activate
    '��һ��
    Range("Q3").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$A$1:$A$65536," & """" & "������" & """" & ")"
    Range("Q4").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$A$1:$A$65536," & """" & "������" & """" & ")"
    Range("Q5").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$A$1:$A$65536," & """" & "������" & """" & ")"
    Range("Q6").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$A$1:$A$65536," & """" & "������" & """" & ")"
    Range("Q7").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$A$1:$A$65536," & """" & "�����" & """" & ")"
    Range("Q8").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$A$1:$A$65536," & """" & "������" & """" & ")"
    Range("Q9").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$A$1:$A$65536," & """" & "մ����" & """" & ")"
    Range("Q10").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$A$1:$A$65536," & """" & "��ƽ��" & """" & ")"
    '�ڶ���
    Range("R3").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$C$1:$C$65536," & """" & "������" & """" & ")"
    Range("R4").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$C$1:$C$65536," & """" & "������" & """" & ")"
    Range("R5").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$C$1:$C$65536," & """" & "������" & """" & ")"
    Range("R6").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$C$1:$C$65536," & """" & "������" & """" & ")"
    Range("R7").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$C$1:$C$65536," & """" & "�����" & """" & ")"
    Range("R8").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$C$1:$C$65536," & """" & "������" & """" & ")"
    Range("R9").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$C$1:$C$65536," & """" & "մ����" & """" & ")"
    Range("R10").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$C$1:$C$65536," & """" & "��ƽ��" & """" & ")"
     '������
    Range("Q16").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$E$1:$E$65536," & """" & "������" & """" & ")"
    Range("Q17").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$E$1:$E$65536," & """" & "������" & """" & ")"
    Range("Q18").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$E$1:$E$65536," & """" & "������" & """" & ")"
    Range("Q19").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$E$1:$E$65536," & """" & "������" & """" & ")"
    Range("Q20").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$E$1:$E$65536," & """" & "�����" & """" & ")"
    Range("Q21").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$E$1:$E$65536," & """" & "������" & """" & ")"
    Range("Q22").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$E$1:$E$65536," & """" & "մ����" & """" & ")"
    Range("Q23").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$E$1:$E$65536," & """" & "��ƽ��" & """" & ")"
    '���ĸ�
    Range("R16").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$G$1:$G$65536," & """" & "������" & """" & ")"
    Range("R17").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$G$1:$G$65536," & """" & "������" & """" & ")"
    Range("R18").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$G$1:$G$65536," & """" & "������" & """" & ")"
    Range("R19").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$G$1:$G$65536," & """" & "������" & """" & ")"
    Range("R20").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$G$1:$G$65536," & """" & "�����" & """" & ")"
    Range("R21").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$G$1:$G$65536," & """" & "������" & """" & ")"
    Range("R22").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$G$1:$G$65536," & """" & "մ����" & """" & ")"
    Range("R23").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$G$1:$G$65536," & """" & "��ƽ��" & """" & ")"
    
    'ȫ�����
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

Function ��������_�ձ�_���װ�ƻ���չ(HomeName, UseName, Day)
    Dim CopyTmp As String '�ļ���λ��
    
    If Day <= 16 Then
        CopyTmp = Chr(Asc("A") + Day - 1)
    Else
        CopyTmp = Chr(Asc("A") + Day - 15 - 1)
    End If
    
    '��������
    'һ
    Windows(HomeName).Activate
    Range("L3:L11").Select
    Selection.Copy
    Windows(UseName).Activate
    Sheets("3�¿��װ�ƻ���չ").Activate
    Range("K3:K11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    If Day <= 16 Then
        Range(CopyTmp & "18:" & CopyTmp & "26").Select
    Else
        Range(CopyTmp & "29:" & CopyTmp & "37").Select
    End If
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

    '��
    Windows(HomeName).Activate
    Range("R3:R11").Select
    Selection.Copy
    Windows(UseName).Activate
    Sheets("3�¿��װ�ƻ���չ").Activate
    Range("M3:M11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    If Day <= 16 Then
        Range(CopyTmp & "41:" & CopyTmp & "49").Select
    Else
        Range(CopyTmp & "52:" & CopyTmp & "60").Select
    End If
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

    '��(��)
    Windows(HomeName).Activate
    Range("L16:L24").Select
    Selection.Copy
    Windows(UseName).Activate
    Sheets("���߿������").Activate
    Range("G4:G12").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

    '�ģ��ˣ�
    Windows(HomeName).Activate
    Range("R16:R24").Select
    Selection.Copy
    Sheets("��װ�������£�").Activate
    Range("H13:H21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    
    Windows(UseName).Activate
    Sheets("���߿������").Activate
    Range("H4:H12").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    
    '
    Windows(HomeName).Activate
    Sheets("�¼ҿ��ձ���ģ�壩").Activate
    Range("T16:T24").Select
    Selection.Copy
    Sheets("��װ�������£�").Activate
    Range("G13:G21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

    Sheets("�¼ҿ��ձ���ģ�壩").Activate
    Range("V16:V24").Select
    Selection.Copy
    Sheets("��װ�������£�").Activate
    Range("F13:F21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

    Windows(UseName).Activate
    Sheets("3�¿��װ�ƻ���չ").Activate
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

Function ����_ר���ձ�_��Ч����_����ԤԼ��(InstallName, InstallNameNull, UseName, RunName, RunNameNull, Year, Month, Day)
    'װ����(�ں���)
    Windows(InstallName).Activate
    Sheets(InstallNameNull).Activate
    Cells.Select
    Cells.AutoFilter
    ActiveWorkbook.Names.Add Name:="'" & InstallNameNull & "'!_FilterDatabase", RefersTo:="='" & InstallNameNull & "'!$A$1:$BW$4034", Visible:=False
    ActiveWindow.ScrollColumn = 16
    Range("A1:BW4034").AutoFilter Field:=35, Criteria1:=Array("�ѹ鵵", "�ѿ�ͨ"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & InstallNameNull & "'!_FilterDatabase", RefersTo:="='" & InstallNameNull & "'!$A$1:$BW$4034", Visible:=False
    Range("A1:BW4034").AutoFilter Field:=33
    ActiveWorkbook.Names.Add Name:="'" & InstallNameNull & "'!_FilterDatabase", RefersTo:="='" & InstallNameNull & "'!$A$1:$BW$4034", Visible:=False
    ActiveWindow.ScrollColumn = 3
    Range("A1:BW4034").AutoFilter Field:=5, Criteria1:=Array("��ͥ���߿��"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & InstallNameNull & "'!_FilterDatabase", RefersTo:="='" & InstallNameNull & "'!$A$1:$BW$4034", Visible:=False
    ActiveWindow.ScrollColumn = 1
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("I:I").Select
    ActiveSheet.Paste

    Windows(UseName).Activate
    Sheets("���߿������").Activate
    '��ֵ
    Range("I4").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$I$1:$I$65536," & """" & "������" & """" & ")"
    Range("I5").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$I$1:$I$65536," & """" & "������" & """" & ")"
    Range("I6").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$I$1:$I$65536," & """" & "������" & """" & ")"
    Range("I7").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$I$1:$I$65536," & """" & "������" & """" & ")"
    Range("I8").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$I$1:$I$65536," & """" & "�����" & """" & ")"
    Range("I9").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$I$1:$I$65536," & """" & "������" & """" & ")"
    Range("I10").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$I$1:$I$65536," & """" & "մ����" & """" & ")"
    Range("I11").Select
    Selection.Formula = "=COUNTIFS([" & InstallName & "]Sheet1!$I$1:$I$65536," & """" & "��ƽ��" & """" & ")"

    '���������е�����
    Windows(RunName).Activate
    '�½�������
    'Worksheets.Add before:=Sheets("20210309_������")
    Worksheets.Add before:=Sheets(RunNameNull)
    Sheets(RunNameNull).Activate
    '��������
    '����
    Cells.Select
    Range("H1").Activate
    Cells.AutoFilter
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$BW$3625", Visible:=False
    Range("A1:BW3625").AutoFilter Field:=22, Operator:=xlFilterValues, Criteria2:=Array(2, Month & "/" & Day & "/" & Year) '����
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$BW$3625", Visible:=False
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("A:A").Select
    ActiveSheet.Paste
    '����
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
    
    '����
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
    
    '��ֵ
    Windows(UseName).Activate
    Sheets("���߿������").Activate
    '��ֵ
    '����
    Range("B4").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$A$1:$A$65536," & """" & "������" & """" & ")"
    Range("B5").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$A$1:$A$65536," & """" & "������" & """" & ")"
    Range("B6").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$A$1:$A$65536," & """" & "������" & """" & ")"
    Range("B7").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$A$1:$A$65536," & """" & "������" & """" & ")"
    Range("B8").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$A$1:$A$65536," & """" & "�����" & """" & ")"
    Range("B9").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$A$1:$A$65536," & """" & "������" & """" & ")"
    Range("B10").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$A$1:$A$65536," & """" & "մ����" & """" & ")"
    Range("B11").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$A$1:$A$65536," & """" & "��ƽ��" & """" & ")"
   '����
    Range("C4").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$B$1:$B$65536," & """" & "������" & """" & ")"
    Range("C5").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$B$1:$B$65536," & """" & "������" & """" & ")"
    Range("C6").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$B$1:$B$65536," & """" & "������" & """" & ")"
    Range("C7").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$B$1:$B$65536," & """" & "������" & """" & ")"
    Range("C8").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$B$1:$B$65536," & """" & "�����" & """" & ")"
    Range("C9").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$B$1:$B$65536," & """" & "������" & """" & ")"
    Range("C10").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$B$1:$B$65536," & """" & "մ����" & """" & ")"
    Range("C11").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$B$1:$B$65536," & """" & "��ƽ��" & """" & ")"
   '����
    Range("D4").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$C$1:$C$65536," & """" & "������" & """" & ")"
    Range("D5").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$C$1:$C$65536," & """" & "������" & """" & ")"
    Range("D6").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$C$1:$C$65536," & """" & "������" & """" & ")"
    Range("D7").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$C$1:$C$65536," & """" & "������" & """" & ")"
    Range("D8").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$C$1:$C$65536," & """" & "�����" & """" & ")"
    Range("D9").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$C$1:$C$65536," & """" & "������" & """" & ")"
    Range("D10").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$C$1:$C$65536," & """" & "մ����" & """" & ")"
    Range("D11").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$C$1:$C$65536," & """" & "��ƽ��" & """" & ")"

    '���
    Range("B12").Select
    Selection.Formula = "=sum(B4:B11)"
    Range("C12").Select
    Selection.Formula = "=sum(C4:C11)"
    Range("D12").Select
    Selection.Formula = "=sum(D4:D11)"
    Range("I12").Select
    Selection.Formula = "=sum(I4:I11)"
    
    'תΪ���ָ�ʽ
    
    Range("B4:D11").Select
    Selection.Copy
    Range("B4:D11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    
    Range("I4:I11").Select
    Selection.Copy
    Range("I4:I11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
   

End Function

Function ����_ר���ձ�_��װ����(DispatchName, ExpansionName, AppointmentName, RunName, RunNameNull, UseName, HomeName)
    Dim EndH, MyH As Long
'
    Windows(RunName).Activate
    Sheets(RunNameNull).Activate
    Cells.Select
    Range("D1").Activate
    Cells.AutoFilter
    Cells.AutoFilter
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='20210315_������'!$A$1:$BX$3579", Visible:=False




    MyH = 0
    EndH = 0
    '������������
    Windows(AppointmentName).Activate
    Worksheets.Add before:=Sheets("�����ȹ���")
    Sheets("�����ȹ���").Activate

    MyH = Range("G65536").End(xlUp).Row - 1
    EndH = EndH + MyH
    Range("G2:G" & (MyH + 1)).Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Range("A1:A" & EndH).Select
    ActiveSheet.Paste
    Range("B1").Formula = "ԤԼ��"
    Range("B1").Select
    Selection.AutoFill Destination:=Range("B1:B" & EndH), Type:=xlFillDefault


    Windows(ExpansionName).Activate
    Sheets("�����ȹ���").Activate
    MyH = Range("G65536").End(xlUp).Row - 1
    Range("G2:G" & (MyH + 1)).Select
    Selection.Copy

    Windows(AppointmentName).Activate
    Sheets("Sheet1").Activate
    Range("A" & EndH + 1 & ":A" & (EndH + MyH)).Select

    ActiveSheet.Paste
    Range("B" & EndH + 1).Formula = "���ݳ�"
    Range("B" & EndH + 1).Select
    Selection.AutoFill Destination:=Range("B" & (EndH + 1) & ":B" & (EndH + MyH)), Type:=xlFillDefault
    EndH = EndH + MyH
'
'
    Windows(DispatchName).Activate
    Sheets("�����ȹ���").Activate
    MyH = Range("E65536").End(xlUp).Row - 1
    Range("E2:E" & (MyH + 1)).Select
    Selection.Copy

    Windows(AppointmentName).Activate
    Sheets("Sheet1").Activate
    Range("A" & EndH + 1 & ":A" & (EndH + MyH)).Select
    ActiveSheet.Paste
    Range("B" & EndH + 1).Formula = "��װ��"
    Range("B" & EndH + 1).Select
    Selection.AutoFill Destination:=Range("B" & (EndH + 1) & ":B" & (EndH + MyH)), Type:=xlFillDefault
'    EndH = EndH + MyH


'    Cells.Select
'    Range("BD1").Activate
'    Cells.AutoFilter
'    ActiveWorkbook.Names.Add Name:="'20210309_������'!_FilterDatabase", RefersTo:="='20210309_������'!$A$1:$BW$3625", Visible:=False
'    ActiveWindow.ScrollColumn = 1
'    Range("A1:BW3625").AutoFilter Field:=5, Criteria1:=Array("����1"), Operator:=xlFilterValues
'    ActiveWorkbook.Names.Add Name:="'20210309_������'!_FilterDatabase", RefersTo:="='20210309_������'!$A$1:$BW$3625", Visible:=False
'    ActiveWindow.ScrollColumn = 32
'    Range("A1:BW3625").AutoFilter Field:=36, Criteria1:=Array("ǰ̨Ӫҵȷ��"), Operator:=xlFilterValues
'    ActiveWorkbook.Names.Add Name:="'20210309_������'!_FilterDatabase", RefersTo:="='20210309_������'!$A$1:$BW$3625", Visible:=False
'    Range("AK939").Select
'    Selection.Formula = "����"
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

    '���������е�����
    Windows(RunName).Activate
    Sheets(RunNameNull).Activate
    '������
    Columns("AK:AK").Select
    Selection.Insert Shift:=xlShiftToRight
    
'    'ѭ������
        EndH = Range("A65536").End(xlUp).Row
'      '�������
'      '�ж��Ƿ�Ϊ�գ�Ϊ��������
'       ' EndH = 20
        For i = EndH To 2 Step -1
        XX = Range("AJ" & i).Value
        If IsEmpty(Cells(i, 5)) Then
        
        ElseIf InStr(XX, "ǰ̨Ӫҵȷ��") <> 0 Then '
          Range("AK" & i).Formula = "����"
        ElseIf Not InStr(XX, "��Դ����") <> 0 And (InStr(XX, "�豸����/��·���") <> 0 Or InStr(XX, "ʩ��ԤԼ") <> 0 Or InStr(XX, "�ֳ���ͨ") <> 0 Or InStr(XX, "�쳣����") <> 0) Then     '
            Range("AK" & i).Formula = "��װ��"
        ElseIf InStr(XX, "��Դ����") <> 0 Then

        Else
        '�׹�ʽ
            Range("AK" & i).Select
            Selection.Formula = "=VLOOKUP(Q" & i & "&"""",[" & AppointmentName & "]Sheet1!$A:$B,2,FALSE)"
            'Selection.Formula = "=VLOOKUP(Q532,[20210309_ԤԼ��.xls]Sheet1!$A:$B,2,FALSE)"
            If IsError(Cells(i, 37)) Then
                Range("AK" & i).Formula = "����"
            Else
            
            End If
        End If
     Next

    '��������
    '�½�������
    Worksheets.Add before:=Sheets(RunNameNull)
    Sheets(RunNameNull).Activate
    '����ɸѡ
    Cells.Select
    Range("BA1").Activate
    Cells.AutoFilter
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CB$3625", Visible:=False
    Application.DisplayAlerts = False
    Application.DisplayAlerts = True
    ActiveWindow.ScrollColumn = 1
    Range("A1:CB3625").AutoFilter Field:=5, Criteria1:=Array("��ͥ���߿��"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CB$3625", Visible:=False
    ActiveWindow.ScrollColumn = 29
    Range("A1:CB3625").AutoFilter Field:=36, Criteria1:=Array("ǰ̨Ӫҵȷ��", "�豸����/��·���", "ʩ��ԤԼ", "�ֳ���ͨ", "�쳣����", "װ������"), Operator:=xlFilterValues
    
    '�ֿ���������
    '���ݳ�
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CB$3625", Visible:=False
    Range("A1:CB3625").AutoFilter Field:=37, Criteria1:=Array("���ݳ�"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CB$3625", Visible:=False
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("O:O").Select
    ActiveSheet.Paste
    'ԤԼ��
    Sheets(RunNameNull).Activate
     ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CB$3625", Visible:=False
    Range("A1:CB3625").AutoFilter Field:=37, Criteria1:=Array("ԤԼ��"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CB$3625", Visible:=False
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("P:P").Select
    ActiveSheet.Paste
    '����
    Sheets(RunNameNull).Activate
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CB$3625", Visible:=False
    Range("A1:CB3625").AutoFilter Field:=37, Criteria1:=Array("����"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CB$3625", Visible:=False
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("Q:Q").Select
    ActiveSheet.Paste
    '��װ��
    Sheets(RunNameNull).Activate
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CB$3625", Visible:=False
    Range("A1:CB3625").AutoFilter Field:=37, Criteria1:=Array("��װ��"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CB$3625", Visible:=False
    Columns("L:L").Select
    Selection.Copy
    Sheets("Sheet1").Activate
    Columns("R:R").Select
    ActiveSheet.Paste
    
    '��ֵ
    Windows(UseName).Activate
    Sheets("��װ����").Activate
    
    '���ݳ�
    Range("D3").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$O$1:$O$65536," & """" & "������" & """" & ")"
    Range("D4").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$O$1:$O$65536," & """" & "������" & """" & ")"
    Range("D5").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$O$1:$O$65536," & """" & "������" & """" & ")"
    Range("D6").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$O$1:$O$65536," & """" & "������" & """" & ")"
    Range("D7").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$O$1:$O$65536," & """" & "�����" & """" & ")"
    Range("D8").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$O$1:$O$65536," & """" & "������" & """" & ")"
    Range("D9").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$O$1:$O$65536," & """" & "մ����" & """" & ")"
    Range("D10").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$O$1:$O$65536," & """" & "��ƽ��" & """" & ")"
    'ԤԼ��
    Range("E3").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$P$1:$P$65536," & """" & "������" & """" & ")"
    Range("E4").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$P$1:$P$65536," & """" & "������" & """" & ")"
    Range("E5").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$P$1:$P$65536," & """" & "������" & """" & ")"
    Range("E6").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$P$1:$P$65536," & """" & "������" & """" & ")"
    Range("E7").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$P$1:$P$65536," & """" & "�����" & """" & ")"
    Range("E8").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$P$1:$P$65536," & """" & "������" & """" & ")"
    Range("E9").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$P$1:$P$65536," & """" & "մ����" & """" & ")"
    Range("E10").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$P$1:$P$65536," & """" & "��ƽ��" & """" & ")"
      '����
    Range("F3").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$Q$1:$Q$65536," & """" & "������" & """" & ")"
    Range("F4").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$Q$1:$Q$65536," & """" & "������" & """" & ")"
    Range("F5").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$Q$1:$Q$65536," & """" & "������" & """" & ")"
    Range("F6").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$Q$1:$Q$65536," & """" & "������" & """" & ")"
    Range("F7").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$Q$1:$Q$65536," & """" & "�����" & """" & ")"
    Range("F8").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$Q$1:$Q$65536," & """" & "������" & """" & ")"
    Range("F9").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$Q$1:$Q$65536," & """" & "մ����" & """" & ")"
    Range("F10").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$Q$1:$Q$65536," & """" & "��ƽ��" & """" & ")"
      '��װ
    Range("G3").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$R$1:$R$65536," & """" & "������" & """" & ")"
    Range("G4").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$R$1:$R$65536," & """" & "������" & """" & ")"
    Range("G5").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$R$1:$R$65536," & """" & "������" & """" & ")"
    Range("G6").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$R$1:$R$65536," & """" & "������" & """" & ")"
    Range("G7").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$R$1:$R$65536," & """" & "�����" & """" & ")"
    Range("G8").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$R$1:$R$65536," & """" & "������" & """" & ")"
    Range("G9").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$R$1:$R$65536," & """" & "մ����" & """" & ")"
    Range("G10").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$R$1:$R$65536," & """" & "��ƽ��" & """" & ")"
    
    '���
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
  
    
    
    
    '��������
    'ѡ������
    Range("G3:G11").Select
    Selection.Copy
    '��������
    Sheets("���߿������").Activate
    Range("F4:F12").Select
    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    
    Windows(HomeName).Activate
    Sheets("��װ����ģ��").Activate
    Range("C3:C11").Select
    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    Sheets("��װ�������£�").Activate
    Range("B13:B21").Select
    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    
    '
    Sheets("�¼ҿ��ձ���ģ�壩").Activate
    Range("R16:R24").Select
    Selection.Copy
   
    Sheets("��װ�������£�").Activate
    Range("H13:H21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

 
    Range("I13:I21").Select
    Selection.Copy
    Sheets("��װ����ģ��").Activate
    Range("B3:B11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

   



End Function
Function ͳ��_�¼ҿ�_��װ����_ImsOtt(RunName, RunNameNull, HomeName, UseName)


    Windows(RunName).Activate
    Sheets(RunNameNull).Activate
    
    
    Cells.Select
    Cells.AutoFilter
    Cells.AutoFilter
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CM$3625", Visible:=False
    Range("A1:CM3625").AutoFilter Field:=7, Criteria1:=Array("IMS"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'" & RunNameNull & "'!_FilterDatabase", RefersTo:="='" & RunNameNull & "'!$A$1:$CM$3625", Visible:=False
    ActiveWindow.ScrollColumn = 27
    Range("A1:CM3625").AutoFilter Field:=37, Criteria1:=Array("��װ��"), Operator:=xlFilterValues
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
    
    '��ֵ
    Windows(HomeName).Activate
    Sheets("��װ����ģ��").Activate
    
    'IMS
    Range("E3").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$T$1:$T$65536," & """" & "������" & """" & ")"
    Range("E4").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$T$1:$T$65536," & """" & "������" & """" & ")"
    Range("E5").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$T$1:$T$65536," & """" & "������" & """" & ")"
    Range("E6").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$T$1:$T$65536," & """" & "������" & """" & ")"
    Range("E7").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$T$1:$T$65536," & """" & "�����" & """" & ")"
    Range("E8").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$T$1:$T$65536," & """" & "������" & """" & ")"
    Range("E9").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$T$1:$T$65536," & """" & "մ����" & """" & ")"
    Range("E10").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$T$1:$T$65536," & """" & "��ƽ��" & """" & ")"
      'OTT
    Range("D3").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$U$1:$U$65536," & """" & "������" & """" & ")"
    Range("D4").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$U$1:$U$65536," & """" & "������" & """" & ")"
    Range("D5").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$U$1:$U$65536," & """" & "������" & """" & ")"
    Range("D6").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$U$1:$U$65536," & """" & "������" & """" & ")"
    Range("D7").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$U$1:$U$65536," & """" & "�����" & """" & ")"
    Range("D8").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$U$1:$U$65536," & """" & "������" & """" & ")"
    Range("D9").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$U$1:$U$65536," & """" & "մ����" & """" & ")"
    Range("D10").Select
    Selection.Formula = "=COUNTIFS([" & RunName & "]Sheet1!$U$1:$U$65536," & """" & "��ƽ��" & """" & ")"
    '���
    
    Range("D11").Select
    Selection.Formula = "=sum(D3:D10)"
    Range("E11").Select
    Selection.Formula = "=sum(E3:E10)"
       
    Range("D3:E11").Select
    Selection.Copy
    Range("D3:E11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    
    
    '����
    '���ݳ�
    Windows(UseName).Activate
    Sheets("��װ����").Activate
    Range("D3:D11").Select
    Selection.Copy
    Windows(HomeName).Activate
    Sheets("��װ����ģ��").Activate
    Range("G3:G11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    
    '�ɴ�װ��
    Sheets("��װ�������£�").Activate
    Range("B13:B21").Select
    Selection.Copy
    Windows(UseName).Activate
    Sheets("װ������ͨ��").Activate
    Range("K4:K12").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    '����ƽ��
    Windows(HomeName).Activate
    Sheets("��װ�������£�").Activate
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

    'ͳ��
    Range("K12:M12").Select
    Selection.Copy
    Range("E12:G12").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    
     For i = 11 To 4 Step -1
        XX = Range("A" & i).Value
        For y = 11 To 4 Step -1
           YY = Range("J" & y).Value
           If InStr(XX, YY) <> 0 Then '@����Ҫ�ҵ�ָ�����ض��ַ� ���Ըĳ���ָ���������ַ�
             '  MsgBox (XX & "����" & YY & "DD" & y & ":" & i)
               Range("K" & i & ":M" & i).Select
               Selection.Copy
               Range("E" & y & ":G" & y).Select
               Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
           End If
        Next
     Next


     
     
    

    
    
    
    
    
    


End Function
Function ����_ר���ձ�_��װ����ģ��(HomeName, UseName)
    Windows(HomeName).Activate
    Sheets("��װ����ģ��").Activate
    Range("B3:G11").Select
    Selection.Copy
    
    Windows(UseName).Activate
    Sheets("��װ����ģ��").Activate
    Range("B3:G11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    
    
    

End Function

Function ������_װ����_ȥ��(AddName, AddNameNull, InstallName, InstallNameNull)

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
