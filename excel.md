

# excel编程

(Alt+F11)打开代码编辑器

Cells(row,column)代表单个单元格，其中row为行号，column为列号；

Cells(row, column)  表示row行第column列；

Trim(Cells(row, column))  取该单元格当中的内容，除去两侧空格；

Range(arg)来引用单元格或单元格区域，其中arg可为单元格号、单元

格号范围、单元格区域名称;

Cells(row_MTK, custom_mnc).Interior.ColorIndex得到当前单元格的

颜色值；

Sheets(“Sheet2”).Select   选中当前工作表；

MsgBox “XXXXXX”  消息弹出框。

```
Sub 表改名()
For i = 1 To ThisWorkbook.Sheets.Count
    If Sheets(i).Name = “我的数据" then 
       Sheets(i-1).Name = “数据透视"
    end If
Next
End Sub
```

```
Private Sub CommandButton1_Click()

Sheets("Sheet1").Select
If MsgBox("Confirm to reset?", vbYesNo) <> vbYes Then Exit Sub
For j = 6 To 42

Cells(j, 3) = Null

Next j

MsgBox "qingkong"

End Sub

```







```
Sub CommandButton2_Click()
Dim cnn As Object, SQL$, sh As Worksheet
Set cnn = CreateObject("ADODB.Connection")
cnn.Open "Provider = Microsoft.Jet.Oledb.4.0;Extended Properties ='Excel 8.0;hdr=no';Data Source =" & ThisWorkbook.Path & "\test2.xls"
For Each sh In Sheets
SQL = "Select c6,c7 from [" & sh.Name & "$a2:f]"
sh.Range("A2:C65536").ClearContents
sh.[a2].CopyFromRecordset cnn.Execute(SQL)
Next
cnn.Close
Set cnn = Nothing

End Sub




Sub HideApplication()
    Dim Sht As Worksheet
    Dim Temp As String
    Temp = ThisWorkbook.Path & "\abcd.xls"
    Set Sht = Workbooks.Open(Temp).Sheets(1)
    With Cells.Select
    Cells.Copy
    End With
ThisWorkbook.Activate
    Cells.Select
    ActiveSheet.Paste
    Application.DisplayAlerts = False
    Windows("abcd.xls").Activate
   ActiveWindow.Close
End Sub


Dim  wb As Workbook
Application.ScreenUpdating = False
Set wb = Workbooks.Open(要调用的工作薄的路径及名称）
‘路径及名称格式如下   ThisWorkbook.Path & "\Back.xlsx")  
With wb.Sheets("表名  不是工作薄名").range(要调用的单元格）
对调用单无格的操作
End With
wb.Close 1
Application.ScreenUpdating = True
```





```
Sub ReadB()
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim sheet As Excel.Worksheet
Dim i as Integer
Dim j as Integer
Set xlApp = New Excel.Application
Set xlBook = xlApp.Workbooks.Open("d:\a.xlsx")
Set sheet = xlBook.Worksheets(1)
For i=1 to 22
For j=1 to 110
sheets(1).cells(i,j)=sheet.cells(i,j)
Next j
Next
xlBook.Close
End Sub





Private Sub CommandButton3_Click()
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim sheet As Excel.Worksheet
Dim i As Integer
Dim j As Integer
Set xlApp = New Excel.Application
Set xlBook = xlApp.Workbooks.Open("C:\Users\ldx\Desktop\test\test2.xlsx")
Set sheet = xlBook.Worksheets(1)
For i = 1 To 20
For j = 1 To 110
Sheets(1).Cells(i, j) = sheet.Cells(i, j)
Next j
Next i
xlBook.Close
End Sub


Private Sub CommandButton3_Click()
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim sheet As Excel.Worksheet
Dim i As Integer
Dim j As Integer
Set xlApp = New Excel.Application
Set xlBook = xlApp.Workbooks.Open("C:\Users\ldx\Desktop\test\test2.xlsx")
Set sheet = xlBook.Worksheets(1)
For j = 1 To 100
For i = 1 To 100
If Sheets(1).Cells(i, 5) = sheet.Cells(j, 1) Then
Sheets(1).Cells(i, 7) = sheet.Cells(j, 3)
End If
Next i
Next j
xlBook.Close
End Sub

Private Sub CommandButton1_Click()
If MsgBox("Confirm to reset?", vbYesNo) <> vbYes Then Exit Sub
For j = 32 To 1183
Sheets(1).Cells(j, 3) = Null
Next j
MsgBox "qingkong"
End Sub


Private Sub CommandButton1_Click()
Sheets("Sheet1").Select
If MsgBox("Confirm to reset?", vbYesNo) <> vbYes Then Exit Sub
For j = 2 To 1158
Cells(j, 3) = Null
Next j
MsgBox "qingkong"
End Sub

Private Sub CommandButton2_Click()
If MsgBox("Confirm to input?", vbYesNo) <> vbYes Then Exit Sub
For i = 2 To 1183
For j = 2 To 1142
If (Sheets(1).Cells(i, 1) = Sheets(1).Cells(j, 7)) Then
Sheets(1).Cells(i, 3) = Sheets(1).Cells(j, 9)
End If
Next j
Next i
End Sub

```





```
Option Explicit

Private Const FirstRow_Nr = 16
Private Const DelimiterChar_Cell = "C8"
Private Const AllowedCol_Cell = "C6"
Private Const SupportedCol_Cell = "C7"

Public Function tokenize(Value As String, Delimiter As String) As Collection
    '1.0 | 2007-07-21 | WRU
    On Error GoTo Catch
    Dim i As Long, s As Long, c As String
    Set tokenize = New Collection
    For i = 1 To Len(Value)
        c = Mid(Value, i, 1)
        If c = Delimiter Then
            tokenize.Add Trim(Mid(Value, s + 1, i - s - 1))
            s = i
        End If
    Next
    If Right(Value, 1) = Delimiter Then
        tokenize.Add ""
    Else
        tokenize.Add Trim(Right(Value, i - s - 1))
    End If
    Exit Function
Catch:
    Set tokenize = New Collection
End Function

Private Function isAllowed(Allowed As String, Supported As String, Delimiter As String, Value As String) As Boolean
    '1.0 | 2007-07-21 | WRU
    On Error GoTo Final
    Dim colAllowed As Collection, colSupported As Collection, SupportedLC As String
    Value = ""
    If Allowed = "" Or LCase(Allowed) = "yes | no" Then
        If Supported = "" Then
            isAllowed = True
        ElseIf Supported = "1" Or Supported = "+" Or LCase(Supported) = "yes" Or LCase(Supported) = "y" Or LCase(Supported) = "true" Then
            Value = "yes"
            isAllowed = True
        ElseIf Supported = "0" Or Supported = "-" Or LCase(Supported) = "no" Or LCase(Supported) = "n" Or LCase(Supported) = "false" Then
            Value = "no"
            isAllowed = True
        End If
    Else
        Dim a As Long, s As Long
        Set colAllowed = tokenize(Allowed, Delimiter)
        Set colSupported = tokenize(Supported, Delimiter)
        For s = 1 To colSupported.Count
            For a = 1 To colAllowed.Count
                If colSupported.Item(s) = colAllowed.Item(a) Then Exit For
            Next
            If a > colAllowed.Count Then GoTo Final
        Next
        Value = Supported
        isAllowed = True
    End If
Final:
    Set colSupported = Nothing
    Set colAllowed = Nothing
End Function

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    '1.3 | 2007-09-24 | WRU
    
    On Error GoTo Final
    Dim Value As String
    
    Static Atomic As Boolean
    If Atomic Then Exit Sub
    Atomic = Not Atomic
    
    If Target.Columns.Count <> 1 Then GoTo Final
    If Target.Text = "" Then GoTo Final
    If Target.Column <> CLng(Sh.Range(SupportedCol_Cell).Text) Then GoTo Final
    If Target.Row < FirstRow_Nr Then GoTo Final
    
    Dim Cell As Object
    For Each Cell In Target.Cells
        If isAllowed(Sh.Cells(Cell.Row, CLng(Sh.Range(AllowedCol_Cell).Text)).Text, Cell.Text, Sh.Range(DelimiterChar_Cell).Text, Value) Then
            Cell.Value = Value
        Else
            MsgBox """" & Cell.Text & """ is not valid, please enter allowed values only, separated by pipe char (""|"").", vbExclamation, _
                "Invalid supported value(s) in row " & Cell.Row
            Cell.Value = ""
            Cell.Select
        End If
    Next
Final:
    Atomic = Not Atomic
End Sub
```





```
Private Sub CommandButton1_Click()

If TextBox1.Text<>"郭轶凡"Then '判断用户名是否正确

MsgBox"用户登录名错误，您无权登录!" '不正确给出提示

With TextBox1

.SelStart=0 '设置选择文字的开始字符

.SelLength=Len(TextBox1.Text) '设置选择文本的长度

.SetFocus '文本框获得焦点

End With

ElseIf TextBox2.Text<>"abcdef "Then '如果密码错误

MsgBox"密码输入错误，请重新输入!" '给出提示

With TextBox2

.SelStart=0 '设置选择文本的开始字符

.SelLength=Len(TextBox2.Text) '设置选择文本的长度

.SetFocus '获得焦点

End With

Else

MsgBox"登录成功，欢迎你的到来!" '登录成功提示

Unload Me '卸载窗体

End If

End Sub

使用Excel制作用户登录窗口的方法
5
步骤五：接着在“代码”窗口中输入程序代码，为“取消”按钮添加Click事件代码，具体程序如下所示：

Private Sub CommandButton2_Click()

Unload Me '卸载窗体

ThisWorkbook.Close '关闭工作簿

End Sub








Private Sub 创建下拉列表框()
    OLEObjects.Add ClassType:="Forms.ComboBox.1", Link:=True, DisplayAsIcon:=False, Left:=0, Top:=29, Width:=55, Height:=20
End Sub
 
Sub 添加列表数据()
    Dim brr()
    n = Range("IV3").End(xlToLeft).Column
    ReDim brr(1 To n / 2)
    For i = 2 To n Step 2
        brr(i / 2) = Cells(3, i)
    Next i
    ComboBox1.List = brr
End Sub
 
Private Sub ComboBox1_Change()
    Range("B3:IV3").Find(What:=ComboBox1.Value).Activate
End Sub




Sub CreateList()
    ''创建下拉列表
    Dim i As Long, w1 As String
    w1 = ""
    With Sheet1
        ''首先创建下拉列表数据
        For i = 2 To 6 Step 2 ''至于最后一列请自定义
            w1 = w1 & IIf(w1 <> "", ",", "")
            w1 = w1 & Trim$(.Cells(3, i)) & "(" & Trim$(.Cells(4, i)) & ")"
        Next
        ''添加数据有效性
        With .Range("a3").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=w1
            .InCellDropdown = True
        End With
    End With
End Sub
'''下面是excel的事件
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address(0, 0) = "A3" Then ''将事件限制在单元格a3的改变上
        Dim w1 As String
        w1 = Split(Target.Value, "(")(0) ''分解出人名
        Range("B3:Z3").Find(What:=w1).Activate ''利用excel的自动搜索功能
    End If
End Sub



  '首先建立一个全局变量，用于存储是否为代码做出的更改   
  Public isMe As Boolean      1  2  3  4  5  6  7  8  9  10  11  12  13  14  15  16  17  18  19  20  21  22  23  24  25  26  27  28  29  30  31  32  33   '在工作簿如下事件中，输入如下代码，即可实现感知效果   
  Private Sub Workbook_SheetChange(ByVal Sh As Object, _               ByVal Target As Range)     Sub tst()
Dim arr(1 To 10)
Dim i%, p As String
p = InputBox("tst")
For i = 1 To 10
arr(i) = i & "个"
If p = arr(i) Then MsgBox i
Next
End Sub
  If isMe Then      
  isMe = False      
  Exit Sub    
  End If    
  On Error GoTo ExitMe     
  Dim Arr     Dim i As Integer    
  Dim IsFind As Boolean    
  IsFind = False    
  If Target.Validation.Type = 3 Then      
  Arr = Split(Target.Validation.Formula1, ",")       
  For i = 0 To UBound(Arr)        
  If InStr(1, Arr(i), Target.Value, vbTextCompare) <> 0 Then         
  isMe = True         Target.Value = Arr(i)          IsFind = True         
  Exit For      
  End If     
  Next     
  If Not IsFind Then      
  MsgBox "数据输入错误！", vbCritical        
  Target.Select    
  End If  
  End If 
  ExitMe:   
  End Sub 
  
  Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, _               ByVal Target As Range)   
  isMe = False  
  End Sub
  
  
  
 
  
  
  
  
####待验证  
  Private Sub Worksheet_SelectionChange(ByVal Target As Range)
If Target.Column = 2 And Cells(Target.Row, 1) <> "" Then
    Sheet1.Range("C:C").Delete
    j = 1
    For i = 1 To Sheet1.UsedRange.Rows.Count
        If Sheet1.Cells(i, 1) = Cells(Target.Row, 1) Then
            Sheet1.Cells(j, 3) = Sheet1.Cells(i, 2)
            j = j + 1
        End If
    Next i
     
    Dim cnum
    cnum = Application.WorksheetFunction.CountA(Sheet1.Range("c:c"))
    If cnum >= 1 Then
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=Sheet1!C1:C" & Application.WorksheetFunction.CountA(Sheet1.Range("c:c"))
        End With
    End If
End If
End Sub





########给窗口下拉框赋值
Private Sub UserForm_Initialize()
ComboBox1.List = Array("收入", "支也")
End Sub

Private Sub UserForm_Initialize()
    Dim k%
    'ComboBox1.AddItem "收入"
    'ComboBox1.AddItem "支出"
    For k = 1 To 10
        ComboBox1.AddItem Range("a" & k)
    Next k
End Sub

Private Sub UserForm_Initialize()
Dim rq
rq = Date
With Me.ComboBox1
    .AddItem rq - 2
    .AddItem rq - 1
    .AddItem Date
    .Value = rq - 1
End With
End Sub
Sub tst()
Dim arr(1 To 10)
Dim i%, p As String
p = InputBox("tst")
For i = 1 To 10
arr(i) = i & "个"
If p = arr(i) Then MsgBox i
Next
End Sub
Private Sub UserForm_Initialize()
    d1 = VBA.DateAdd("d", -1, Date)
    d2 = VBA.DateAdd("d", -2, Date)
    d3 = VBA.DateAdd("d", -3, Date)
    ComboBox1.List = Array(d1, d2, d3)
    ComboBox1.Text = d1 ’设置默认选项值
End Sub




#######自定义两个下拉列表，数据从工作表中获取
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Dim arr, s
Dim Rng As Range
Dim row_begin As Long
Dim row_end As Long

row_begin = 1  '下拉备选菜单选择项开始和接受的行数，根据需要自行修改
row_end = 10

For i = row_begin To row_end
           s = Sheets("Sheet1").Range("a" & i)
          's = Sheets(1).Range("A" & i)     '选择A列的内容作为下拉备选项，根据需要自行修改
          If s <> "" Then arr = arr & "," & s
Next i

''''''''''''''''''''第一个下拉框'''''''''''''''''''''''''''
Set Rng = Sheets("Sheet1").Range("H3")   '下拉框是放在H13单元，可以根据需要自行修改
With Rng.Validation
.Delete
.Add Type:=xlValidateList, Formula1:=arr
End With
Set Rng = Nothing 		'下拉框的初始值为空
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''第二个下拉框'''''''''''''''''''''''''''
Set Rng = Sheets("Sheet2").Range("I4")      '如需需要选择sheet，通过括号中修改
With Rng.Validation
.Delete
.Add Type:=xlValidateList, Formula1:=arr
End With
Set Rng = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''后面如需增加，自行负责''''''''''
End Sub

###判断字符串是否存在与数组
Function InArr( v,  a) as Boolean
    dim t
    InArr=true
    for each t in a
        if v=t then exit function
    next t
    InArr=false
End Function

Sub tst()
Dim arr(1 To 10)
Dim i%, p As String
p = InputBox("tst")
For i = 1 To 10
arr(i) = i & "个"
If p = arr(i) Then MsgBox i
Next
End Sub

 
  模块1：
Option Explicit'声明它，以后第个变量都必须显式声明
Public arr() As String
Sub AAA()
Dim i   '不能再Dim arr...
ReDim arr(1 To 3)
For i = 1 To 3
arr(i) = Sheet1.Cells(i + 1, 1)
Next
End Sub
模块2
Sub Macro1()
Call AAA'必须调用里面的redim，否则没有初始化
For i = 1 To 3
MsgBox (arr(i))
Next
End Sub

###获取OptionButton的值
Private Sub CommandButton1_Click() '工作表代码区
    For i = 1 To 20
        If Me.OLEObjects("OptionButton" & i).Object.Value Then
            Cells(3, 3) = "OptionButton" & i & "被选中"
            Exit Sub
        End If
    Next i
End Sub


类模块(名:clsOpt):
Public WithEvents Optbox As MSForms.OptionButton

Private Sub Optbox_Click()
Cells(3, 3) = Optbox.Name
End Sub

工作表代码区:
Dim Opt(1 To 20) As New clsOpt


Private Sub 结束_Click()
    Erase Opt
End Sub

Private Sub 开始_Click()
    For i = 1 To 20
        Set Opt(i).Optbox = Me.OLEObjects("OptionButton" & i).Object
    Next i
End Sub




Sub TEST()
ARR = Array("A", "V", "C")
If VBA.FLITER(ARR, "D", FALSE) Then
  MSGBOX "NO EXISTS"
End If
End Sub



```

