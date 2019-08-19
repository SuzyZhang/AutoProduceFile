'
'
'使用须知:
'1. 需引用 Microsoft Word Object Library 和 Microsoft Scripting Runtime 对象库。路径：工具-->引用
'2. Word 模板对应位置，添加 DocVariable 域，设置的域名必须和配置表一致
'3. Word 域的使用方法，请查阅 https://support.office.com/en-us/article/insert-edit-and-view-fields-in-word-c429bbb0-8669-48a7-bd24-bab6ba6b06bb?ui=en-US&rs=en-US&ad=US
'4. 模板路径，是相对路径写法
'5. 操作方法：确定要打印的行-->选择该行的任意单元格-->运行宏
'6. 运行宏前，不要打开 Word 模板！否则，程序报错！
'
'

Sub Main()

    Dim confSheetName As String
    Dim dataSheetName As String


'================================================================

    confSheetName = "信封_配置"     '  在此处修改配置表
    dataSheetName = "参会方信息"    '  在此处修改数据源表
    
'================================================================


    Dim confSheet As Worksheet  ' 配置表
    Dim dataSheet As Worksheet   ' 数据源表
    Dim dict As Dictionary  ' DocVariable 域名和数据源表列的对应关系
    Dim appWD As Word.Application  '  Word 的 Application 对象，生成 doc
    Dim doc As Word.Document  '  Word 的 Document 对象，控制模板
    Dim row As Integer  '  需要打印的行号

    Set confSheet = ActiveWorkbook.Worksheets(confSheetName)
    Set dataSheet = ActiveWorkbook.Worksheets(dataSheetName)
    Set dict = GetDict(confSheet)
    Set doc = GetTemplate(appWD, confSheet)
    row = ActiveCell.row

    Call AutoProduceFile(doc, dict, dataSheet, row)

End Sub

'
' 获取 DocVariable 域名和数据源表列的对应关系
'
' @param { Worksheet } sheet 配置表
'
' @return Name-Value 键值对

Public Function GetDict(ByRef sheet As Worksheet) As Dictionary

    Dim erow  ' 存放DocVariable 域名列的最后一行
    Dim startCell As Range
    Dim endCell As Range
    Dim nameRng  ' DocVariable 域名存放区域
    
    erow = sheet.Range("A65536").End(xlUp).row
    Set startCell = sheet.Cells(2, 1)
    Set endCell = sheet.Cells(erow, 1)
    Set nameRng = sheet.Range(startCell, endCell)
    
    Dim dict As Dictionary
    Dim r As Range
    Dim name As String  ' dict的 key
    Dim value ' dict的 item
    
    Set dict = New Dictionary
       
'  将 Name-Value 对应关系写入 dict
    For Each r In nameRng
        name = r.value
        value = r.Offset(0, 1).value
        dict.Add key:=name, item:=value
    Next
    
    Set GetDict = dict

End Function

'
' 根据模板的存放路径，返回 Document 实例
'
' @param { Word.Application } appWD Word的顶层对象
' @param { Worksheet } sheet 配置表
'
' @return 指定路径的 Document 实例

Public Function GetTemplate(ByRef appWD As Word.Application, ByRef sheet As Worksheet) As Word.Document

    Dim path As String
    Dim docName As String
    Dim doc As Word.Document

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True
    
'  目标模板路径组成：Excel 所在的路径 + E1.Value
    path = ActiveWorkbook.path & sheet.Range("E1").value
    docName = sheet.Range("G1").value
    
    Set doc = appWD.Documents.Open(path + "\" + docName)

    Set GetTemplate = doc
    
End Function


'
' 把数据源中的值，写入 Word 对应的 DocVariable域
'
' @param { Object } doc Word格式的模板
' @param { Dictionary } dict 存储 DocVariable 域名和数据源表列的对应关系
' @param { Worksheet } sheet 数据源表
' @param { Integer } row 要打印的行

' @return 无
'

Public Function AutoProduceFile(ByRef doc As Word.Document, ByVal dict As Dictionary, ByRef sheet As Worksheet, ByVal row As Integer)

    Application.ScreenUpdating = False
    
    Dim keys  ' 存放 DocVariable 域名的数组
    Dim key As String  ' DocVariable 域名
    Dim item As String  ' DocVariable 域值，必须是 String；否则，强制转换
    Dim i As Integer  '  计数器
    
    keys = dict.keys
    Items = dict.Items
    
'  更新 Word 模板的 DocVariable 域
'  判断 Word 里是否有相同的域名
'  如果没有，新增；如果有，变更
    For i = 0 To dict.Count - 1
    
        key = keys(i)
        colum = dict.item(key)
        item_temp = sheet.Cells(row, colum).value
        '  若不是字符串，强制转换
        myCheck = TypeName(item_temp) Like "String"
        If myCheck = False Then
            item = CStr(item_temp)
        Else
            item = item_temp
        End If
        
        For Each aVar In doc.Variables
            If aVar.name = key Then
                Num = aVar.Index
                Exit For
            End If
        Next aVar
        
        If Num = 0 Then
            doc.Variables.Add name:=key, value:=item
        Else
            doc.Variables(Num).value = item
        End If
        
    Next
    
'  更新域
'  文本框的域，要单独更新
    doc.Fields.Update
    
    For Each sp In doc.Shapes
        If sp.Type = msoTextBox Then
            sp.TextFrame.TextRange.Fields.Update
        End If
    Next
         
    Application.ScreenUpdating = True

End Function
