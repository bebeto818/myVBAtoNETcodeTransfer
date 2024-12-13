# myVBAtoNETcodeTransfer
## 项目概述
本项目旨在将VBA代码迁移到VS.NET环境中，以便于维护和扩展。当前已完成用户窗口的VBA代码迁移，剩余主程序部分尚未迁移。

## 迁移目标
- 将VBA代码转换为VB.NET代码，确保功能一致性。
- 优化代码结构，提高可读性和可维护性。
- 利用VB.NET的特性（如类型安全、面向对象编程）来增强代码的性能和稳定性。

## 迁移原则（特殊性）
在本例中优先考虑以下迁移原则：

### 1. 对象引用方式
- **原生态VSTO**：
  - 直接访问range等对象，都要加很长的前缀“Globals.ThisAddIn.Application."。例如：
    ```vb.net
    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        'A1单元格写入当前时间
        Globals.ThisAddIn.Application.Range("A1").Value = Now
        '对话框弹出当前选定区域的地址
        MsgBox(Globals.ThisAddIn.Application.Selection.Address)
        'A2单元格写入当前工作表的名称
        Globals.ThisAddIn.Application.Range("A2").Value = Globals.ThisAddIn.Application.ActiveSheet.Name
    End Sub
    ```
  - 原生态VSTO对Excel对象的引用，无法摆脱必须在对象前面加上前缀的局面，书写较为繁杂，浪费时间。

- **使用Excel880VSTO框架**：
  - 引用对象代码精简，不用繁杂的前缀，仅仅用加前缀 "Excel."即可。例如：
    ```vb.net
    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        'EXCEL880VSTO框架加持后和VBA写法几乎一致
        'A1单元格写入当前时间
        Range("A1").Value = Now
        '对话框弹出当前选定区域的地址
        MsgBox(Selection.Address)
        'A2单元格写入当前工作表的名称
        Range("A2").Value = ActiveSheet.Name
    End Sub
    ```

### 2. 数组的迁移原则
- **VBA数组**：
  - 数组的下标从1开始，使用 `Dim` 语句声明数组。
  - 例如：
    ```vba
    Dim myArray(1 To 5) As Integer
    ```

- **VB.NET数组**：
  - 数组的下标从0开始，使用 `Dim` 语句声明数组。
  - 例如：
    ```vb.net
    Dim myArray(0 To 4) As Integer  ' 创建一个包含5个元素的数组，索引从0到4
    ```

- **主要区别**：
  - VBA数组的下标从1开始，而VB.NET数组的下标从0开始。
  - VB.NET支持更复杂的数组操作和类型安全，适合处理更复杂的数据结构。

- **数组迁移主要难点**：
  - 原生态VSTO对数组迁移：VB.Net中数组都是0下标开始，无法直接声明1下标的数组了。
  - 如果要保持原有VBA代码的逻辑，就必须对一系列参数进行联动更改，而且容易出错。特别是很多历史VBA代码需要转换的时候，会花费很多的时间。
  - 使用Excel880VSTO框架进行数组迁移：Excel880VSTO框架提供了更便捷的解决方案，仅需要做简单的代码改动，就可以在VB.NET中继续使用1下标的二维数组，原来的VBA代码非零下标数组可直接移植使用，并不用修改相关的参数。

### 二维数组操作情况

#### 1. VBA模式
```vba
Sub 二维数组VBA模式()
    Dim arr
    Sheet1.Activate
    arr = Range("A1:B10")
    ReDim brr(1 To UBound(arr), 1 To 3)
    For i = 1 To UBound(brr)
        For j = 1 To UBound(arr, 2)
            brr(i, j) = arr(i, j) * 2 'Arr的数据*2
        Next
        brr(i, 3) = brr(i, 1) + brr(i, 2) '第3列等于前两列之和
    Next
    
    Sheet1.Range("D1").Resize(UBound(brr), UBound(brr, 2)) = brr
    ReDim Preserve brr(1 To 10, 1 To 2) '重定义保留值
End Sub
```

#### 2. 原生态VSTO模式
```vb.net
Imports Excel = Microsoft.Office.Interop.Excel

Sub 二维数组原生态VSTO模式()
    Dim arr As Object(,)
    Dim i, j As Integer
    Dim sht As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
    
    sht.Activate()
    arr = sht.Range("A1:B10").Value2  '从Excel读取的数组保持1基
    
    'VB.NET中重定义数组需要考虑0基
    Dim brr(UBound(arr, 1) - 1, 2) As Object  '第二维是3-1=2，因为是0基
    
    '注意：从Excel读取的arr仍然是1基的，所以循环保持1基
    For i = 1 To UBound(arr, 1)
        For j = 1 To UBound(arr, 2)
            brr(i - 1, j - 1) = arr(i, j) * 2  'brr需要减1以适应0基
        Next
        brr(i - 1, 2) = brr(i - 1, 0) + brr(i - 1, 1)  '第3列(索引2)等于前两列之和
    Next
    
    'Resize时需要考虑实际大小
    sht.Range("D1").Resize(UBound(brr, 1) + 1, UBound(brr, 2) + 1).Value2 = brr
    
    'VB.NET只能保留最后一维
    ReDim Preserve brr(9, 1)  '10行2列，需要考虑0基
End Sub
```

#### 3. 使用Excel880VSTO框架模式
```vb.net
Imports Excel = Microsoft.Office.Interop.Excel
Imports Excel880VSTO.ArrayHelper

Sub 二维数组Excel880VSTO框架模式()
    Dim arr As Object(,)
    Dim i, j As Integer
    Dim sht As Excel.Worksheet = Sheets("二维数组")
    
    sht.Activate()
    arr = sht.Range("A1:B10").Value2
    
    '使用框架函数DimArray，保持1基数组
    Dim brr(,) = DimArray(1, UBound(arr, 1), 1, 3)
    
    '保持与VBA完全相同的循环逻辑
    For i = 1 To UBound(brr, 1)
        For j = 1 To UBound(arr, 2)
            brr(i, j) = arr(i, j) * 2  'Arr的数据*2
        Next
        brr(i, 3) = brr(i, 1) + brr(i, 2)  '第3列等于前两列之和
    Next
    
    '保持与VBA完全相同的Resize逻辑
    sht.Range("D1").Resize(UBound(brr, 1), UBound(brr, 2)).Value2 = brr
    
    '使用框架函数ReDimArrayPreserve，可以保留任意维度的值
    ReDimArrayPreserve(brr, 1, 10, 1, 2)
End Sub
```

### Excel区域数组的特殊性
- **VB.NET中的Excel区域读取特性**：
  - 这是一个很特殊的情况
    ```vb.net
    Dim arr As Object(,) = Range("A1:B10").Value
    '此时arr是一个从1开始的数组，而不是标准的0基数组
    '即 arr(1,1) 对应A1单元格
    '   arr(10,2) 对应B10单元格
    ```

- **具体示例对比**：
```vb.net
Sub 演示Excel区域数组特性()
    ' 1. 普通VB.NET数组
    Dim normalArr(9, 1) As Object '10行2列，0基数组
    ' 访问方式：normalArr(0,0) 到 normalArr(9,1)
    
    ' 2. Excel读取的数组
    Dim excelArr As Object(,) = Range("A1:B10").Value
    ' 访问方式：excelArr(1,1) 到 excelArr(10,2)
    
    ' 3. 这时可能会有趣的现象
    Debug.Print(excelArr.GetLowerBound(0)) '输出1
    Debug.Print(excelArr.GetUpperBound(0)) '输出10
    Debug.Print(excelArr.GetLowerBound(1)) '输出1
    Debug.Print(excelArr.GetUpperBound(1)) '输出2
End Sub
```

### 原因解释
- 这是Office COM互操作性的特殊设计。
- 为了保持与VBA的兼容性，Excel通过COM接口返回的数组保持1基索引。
- 这不是VB.NET的特性，而是Excel COM对象模型的特性。

### 实际应用中的注意事项
```vb.net
Sub 处理Excel区域数组()
    Dim sht As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
    Dim arr As Object(,) = sht.Range("A1:B10").Value
    ' ...
End Sub
```

### 关键总结
- Excel通过COM接口返回的数组是特殊的1基数组。
- 在处理Excel区域数组时，必须使用1基索引。
- 与普通VB.NET数组（0基）混合使用时需要特别注意。
- 使用Excel880VSTO框架可以统一处理这种差异，简化开发。

### 关键点总结
- **VBA模式**：直接使用1基数组，所有操作与VBA环境一致。
- **原生态VSTO模式**：需要将数组转换为0基，手动调整索引。
- **Excel880VSTO框架模式**：使用框架提供的函数，保持1基数组，代码改动最小，保持与VBA逻辑一致。



	 
## 迁移原则（普遍性）

1. **代码结构**：
   - **VBA**：
     - 使用模块化设计，功能分解为多个小模块，确保每个模块的职责单一。
     - 代码结构相对简单，适合快速开发。
   - **VB.NET**：
     - **模块化设计**：
       - 在VB.NET中，采用更严格的模块化设计，将代码分解为多个独立的模块或类。
       - 每个模块或类应专注于完成特定的任务或功能，避免将多个功能混合在一起。
       - 例如，一个模块可以专门处理数据访问，另一个模块可以处理用户界面逻辑。
     - **单一职责原则（SRP）**：
       - 遵循单一职责原则（SRP），意味着每个模块或类只负责一个功能或任务。
       - 这样做的好处是，当需要修改或扩展某个功能时，只需关注相关的模块，而不必担心影响到其他功能。
       - 例如，如果有一个处理用户登录的模块，它只负责验证用户的凭据，而不涉及数据存储或界面显示的逻辑。
       - 这种设计使得代码更易于维护、测试和理解，因为每个模块的功能清晰且独立。

2. **数组的差异**：
   - **VBA数组**：
     - 数组的下标从1开始，使用 `Dim` 语句声明数组。
     - 例如：
       ```vba
       Dim myArray(1 To 5) As Integer
       ```
   - **VB.NET数组**：
     - 数组的下标从0开始，使用 `Dim` 语句声明数组。
     - 例如：
       ```vb.net
       Dim myArray(4) As Integer  ' 创建一个包含5个元素的数组，索引从0到4
       ```
   - **主要区别**：
     - VBA数组的下标从1开始，而VB.NET数组的下标从0开始。
     - VB.NET支持更复杂的数组操作和类型安全，适合处理更复杂的数据结构。

3. **数据类型**：
   - **VBA**：
     - 在VBA中，数据类型的使用较为灵活，常常使用`Variant`类型来存储不同类型的数据。
     - 这种灵活性虽然方便，但可能导致性能问题和类型不安全。
     - 变量声明不强制，可能导致运行时错误。
   - **VB.NET**：
     - **明确的数据类型**：
       - 在VB.NET中，使用明确的数据类型是最佳实践，避免使用`Variant`类型，以提高性能和类型安全。
       - 例如，使用 `Integer`、`String`、`Boolean` 等具体类型，确保数据的准确性和操作的安全性。
       - 明确的数据类型有助于编译器在编译时检测错误，减少运行时错误。
     - **变量声明**：
       - 在VB.NET中，所有变量在使用前必须声明，并使用`Option Strict On`来强制类型安全。
       - 这意味着在编写代码时，必须明确指定变量的类型，减少类型转换错误。
       - 例如：
         ```vb.net
         Dim count As Integer
         ```
       - 这种严格的类型检查有助于提高代码的可靠性和可维护性。
     - **集合类型**：
       - 使用合适的集合类型（如`List<T>`、`Dictionary<TKey, TValue>`）来替代VBA中的`Collection`，提高数据操作的效率。
       - 例如，使用 `List(Of String)` 来存储字符串列表，提供更好的性能和类型安全。
       - 这些集合类型提供了丰富的方法和属性，支持更复杂的数据操作和查询。

4. **错误处理**：
   - **VBA**：
     - 在VBA中，错误处理通常使用`On Error`语句，这种方式可能导致错误处理逻辑不清晰。
     - 错误处理代码与业务逻辑混杂在一起，可能导致代码难以维护。
     - 例如：
       ```vba
       On Error GoTo ErrorHandler
       ' 代码逻辑
       Exit Sub
       ErrorHandler:
       ' 错误处理逻辑
       ```
   - **VB.NET**：
     - **Try...Catch结构**：
       - 在VB.NET中，使用`Try...Catch`结构替代VBA中的`On Error`语句，提供更清晰的错误处理机制。
       - 这种结构将错误处理代码与业务逻辑分开，使代码更易于阅读和维护。
       - 例如：
         ```vb.net
         Try
             ' 可能引发错误的代码
         Catch ex As Exception
             ' 错误处理代码
         End Try
         ```
       - 这种方式允许捕获特定类型的异常，并根据需要进行不同的处理。
     - **错误记录**：
       - 在VB.NET中，记录错误信息并提供用户友好的错误提示是最佳实践。
       - 可以使用日志记录机制（如`log4net`或`NLog`）来记录错误信息，便于后续分析和调试。
       - 例如，记录错误到日志文件：
         ```vb.net
         Log.Error(ex.Message)
         ```
       - 这种方式有助于在生产环境中监控和诊断问题。
     - **资源管理**：
       - 确保在错误处理代码中释放资源，避免内存泄漏，使用`Using`语句来自动管理资源。
       - `Using`语句确保在代码块结束时自动释放资源，无论是否发生异常。
       - 例如：
         ```vb.net
         Using connection As New SqlConnection(connectionString)
             ' 使用连接
         End Using  ' 自动释放资源
         ```
       - 这种方式简化了资源管理，减少了手动释放资源的错误风险。

5. **对象模型**：
   - **理解VBA和VB.NET之间的对象模型差异**：
     - **VBA对象模型**：
       - VBA的对象模型与Excel紧密集成，使用简单，适合快速开发。对象的访问和操作通常是直接的，代码简洁。
       - 例如，使用 `Application` 对象可以直接访问工作簿、工作表和单元格等：
         ```vba
         Dim ws As Worksheet
         Set ws = Application.Worksheets("Sheet1")
         ```

     - **VB.NET对象模型**：
       - VB.NET的对象模型更为复杂，支持面向对象编程的特性，如继承、封装和多态。
       - 在VB.NET中，访问Excel对象需要通过 `Microsoft.Office.Interop.Excel` 命名空间，通常需要添加对Excel互操作库的引用。
       - 例如：
         ```vb.net
         Dim excelApp As Microsoft.Office.Interop.Excel.Application
         excelApp = Globals.ThisAddIn.Application
         Dim ws As Microsoft.Office.Interop.Excel.Worksheet
         ws = excelApp.Worksheets("Sheet1")
         ```
     - **使用VB.NET的集合和LINQ等特性来简化代码**：
       - VB.NET提供了强大的集合类（如 `List<T>` 和 `Dictionary<TKey, TValue>`），可以替代VBA中的 `Collection` 对象，提供更好的性能和类型安全。
       - 使用LINQ（语言集成查询）可以简化数据操作和查询，使代码更简洁易读。例如，使用LINQ查询Excel数据：
         ```vb.net
         Dim values = From cell In ws.Range("A1:A10")
                      Where cell.Value IsNot Nothing
                      Select cell.Value
         ```

   - **在处理Excel对象时，使用`Microsoft.Office.Interop.Excel`命名空间**：
     - 在VB.NET中，处理Excel对象时，必须引用 `Microsoft.Office.Interop.Excel` 命名空间，以便访问Excel的对象模型。
     - 确保正确管理Excel应用程序的生命周期，避免内存泄漏。使用 `Using` 语句可以确保资源的自动释放：
       ```vb.net
       Using excelApp As New Microsoft.Office.Interop.Excel.Application()
           Dim wb As Microsoft.Office.Interop.Excel.Workbook
           wb = excelApp.Workbooks.Open("C:\path\to\file.xlsx")
           ' 进行操作
           wb.Close()
       End Using
       ```
     - 通过这种方式，可以确保Excel应用程序在使用后被正确关闭，释放资源。

6. **用户界面**：
   - **VBA**：
     - 在VBA中，用户界面通常通过Excel的内置表单和控件实现，灵活但功能有限。
     - 用户界面设计较为简单，适合快速开发和原型设计。
   - **VB.NET**：
     - **用户体验一致性**：
       - 在迁移用户界面时，确保保持用户体验的一致性，考虑用户的操作习惯。
       - 通过一致的设计风格和交互模式，提升用户的使用体验。
     - **界面实现**：
       - 使用Windows Forms或WPF来实现用户界面，推荐使用MVVM模式来组织代码，提升可维护性。
       - Windows Forms适合传统的桌面应用程序，而WPF提供了更丰富的界面设计能力。
     - **事件处理**：
       - 确保界面元素的事件处理逻辑清晰，避免复杂的嵌套结构，使用事件和委托来处理用户交互。
       - 例如，使用事件处理程序来响应按钮点击事件：
         ```vb.net
         AddHandler button.Click, AddressOf Button_Click
         ```

7. **性能优化**：
   - **VBA**：
     - 性能优化相对有限，主要依赖于VBA的执行效率。
     - 代码执行速度可能受到VBA环境的限制。
   - **VB.NET**：
     - **性能关注**：
       - 在迁移过程中，关注性能优化，使用合适的数据结构和算法，避免不必要的计算和内存分配。
       - 通过分析和优化代码，提高应用程序的响应速度。
     - **使用LINQ**：
       - 使用LINQ进行数据查询和操作，简化代码并提高可读性。
       - LINQ提供了强大的数据操作能力，使代码更简洁。
     - **性能测试**：
       - 定期进行性能测试，确保迁移后的代码在性能上满足需求。
       - 使用性能分析工具识别和解决性能瓶颈。

8. **文档和注释**：
   - **VBA**：
     - 文档和注释的维护依赖于开发者的习惯，可能不够规范。
     - 注释通常是手动添加的，缺乏系统性。
   - **VB.NET**：
     - **良好的文档**：
       - 在迁移过程中，保持良好的文档和注释，以便后续维护和团队协作。
       - 详细的文档有助于新成员快速理解项目。
     - **XML注释**：
       - 使用XML注释来生成文档，确保代码的可读性和可维护性。
       - XML注释可以自动生成API文档，提升文档质量。

9. **版本控制**：
   - **VBA**：
     - 版本控制通常依赖于手动管理，缺乏系统性。
     - 版本历史可能不完整，难以追溯。
   - **VB.NET**：
     - **使用版本控制系统**：
       - 使用版本控制系统（如Git）来管理代码的变更，确保代码的历史记录和版本管理。
       - 版本控制系统提供了分支和合并功能，便于团队协作。
     - **定期提交**：
       - 定期提交代码，保持代码库的整洁和可追溯性。
       - 通过频繁提交，减少合并冲突，保持代码库的稳定性。

10. **单元测试**：
    - **VBA**：
      - 单元测试支持有限，通常依赖手动测试。
      - 测试覆盖率可能不足，难以保证代码质量。
    - **VB.NET**：
      - **编写单元测试**：
        - 在迁移过程中，编写单元测试以验证每个模块的功能，确保代码的正确性。
        - 单元测试有助于在代码变更时快速发现问题。
      - **使用测试框架**：
        - 使用测试框架（如NUnit或MSTest）进行自动化测试，提升代码的可靠性。
        - 自动化测试提高了测试效率和覆盖率。

11. **代码审查**：
    - **VBA**：
      - 代码审查通常不够系统，依赖于团队的自发行为。
      - 可能缺乏正式的审查流程。
    - **VB.NET**：
      - **定期审查**：
        - 定期进行代码审查，确保代码符合最佳实践和设计原则。
        - 通过审查，发现和解决潜在问题，提升代码质量。
      - **团队反馈**：
        - 鼓励团队成员之间的反馈和讨论，提升代码质量。
        - 通过团队协作，分享知识和经验，促进团队成长。

12. **安全性**：
    - **VBA**：
      - 安全性考虑相对较少，容易受到攻击。
      - 代码可能缺乏必要的安全防护。
    - **VB.NET**：
      - **考虑安全性**：
        - 在编写代码时，考虑安全性，避免潜在的安全漏洞（如SQL注入、跨站脚本等）。
        - 通过安全编码实践，减少安全风险。
      - **参数化查询**：
        - 使用参数化查询和输入验证来增强应用程序的安全性。
        - 参数化查询防止SQL注入攻击，提升数据安全。

13. **可扩展性**：
    - **VBA**：
      - 可扩展性有限，难以适应复杂的需求变化。
      - 代码结构可能不支持灵活扩展。
    - **VB.NET**：
      - **未来扩展需求**：
        - 设计代码时，考虑未来的扩展需求，确保代码易于修改和扩展。
        - 通过良好的设计，支持功能的快速迭代和扩展。
      - **接口和抽象类**：
        - 使用接口和抽象类来定义可扩展的架构，便于后续功能的添加。
        - 接口和抽象类提供了灵活的扩展点，支持多态和代码重用。

14. **日志记录**：
    - **VBA**：
      - 日志记录功能有限，通常依赖手动记录或简单的文件写入。
      - 日志格式和内容不统一，难以进行系统化的日志管理。
    - **VB.NET**：
      - **使用日志库**：
        - 在VB.NET中，使用成熟的日志库（如log4net或NLog）来记录日志信息，确保日志记录的规范性和一致性。
        - 这些日志库提供了丰富的功能，如日志级别、日志格式、日志输出目标等，便于进行灵活的日志管理。
      - **日志级别**：
        - 使用不同的日志级别（如Debug、Info、Warn、Error、Fatal）来区分日志的重要性，便于后续分析和调试。
        - 例如，在代码中记录调试信息和错误信息：
          ```vb.net
          Log.Debug("This is a debug message.")
          Log.Error("An error occurred.")
          ```
      - **日志输出**：
        - 配置日志输出目标（如文件、数据库、控制台等），确保日志信息能够被有效存储和查看。
        - 例如，将日志输出到文件：
          ```vb.net
          Dim logFilePath As String = "C:\logs\app.log"
          Dim fileAppender As New log4net.Appender.FileAppender()
          fileAppender.File = logFilePath
          fileAppender.AppendToFile = True
          fileAppender.Layout = New log4net.Layout.PatternLayout("%date [%thread] %-5level %logger - %message%newline")
          log4net.Config.BasicConfigurator.Configure(fileAppender)
          ```
      - **日志分析**：
        - 定期分析日志信息，识别系统中的潜在问题和性能瓶颈，提升系统的稳定性和性能。
        - 使用日志分析工具（如ELK Stack）进行日志的集中管理和分析，便于快速定位问题。

15. **多线程**：
    - **VBA**：
      - VBA本身不支持多线程，所有代码在单线程中执行。这意味着在处理长时间运行的任务时，Excel界面可能会变得无响应。
      - 需要通过VBA进行异步操作时，通常依赖于外部组件或使用Excel的后台刷新功能。
    - **VB.NET**：
      - **多线程支持**：
        - VB.NET提供了丰富的多线程支持，可以使用`System.Threading`命名空间中的类来创建和管理线程。
        - 通过多线程，可以在后台执行长时间运行的任务，同时保持用户界面的响应。
      - **异步编程**：
        - 使用`Async`和`Await`关键字可以简化异步编程，使代码更易于编写和维护。
        - 例如，异步调用一个长时间运行的操作：
          ```vb.net
          Public Async Function LoadDataAsync() As Task
              Await Task.Run(Sub() 
                  ' 长时间运行的操作
              End Sub)
          End Function
          ```
      - **任务并行库（TPL）**：
        - 使用任务并行库（TPL）可以更高效地管理并发任务，提供更好的性能和可扩展性。
        - TPL简化了并行编程，自动管理线程池和任务调度。
      - **线程安全**：
        - 在多线程环境中，确保数据访问的线程安全非常重要。可以使用锁（如`SyncLock`）来保护共享资源。
        - 例如：
          ```vb.net
          SyncLock lockObject
              ' 线程安全的代码块
          End SyncLock
          ```
        - 通过这种方式，避免数据竞争和不一致性。


## 迁移步骤
1. **准备VBA代码**：
   - 确保所有VBA模块、表单和类已导出并备份。
   - 进行代码审查，识别需要迁移的功能和模块。

2. **在VS.NET中导入VBA代码**：
   - 创建新项目并打开工作区。
   - 将未迁移的VBA代码文件添加到项目中，确保文件结构清晰。

3. **逐步迁移代码**：
   - 按模块逐步迁移VBA代码，确保每个模块的功能在VB.NET中正常工作。
   - 进行单元测试，确保功能一致性，使用单元测试框架（如NUnit或MSTest）进行测试。
   - 在迁移过程中，记录每个模块的功能和迁移细节，以便后续参考。

4. **文档记录**：
   - 在迁移过程中，记录每个模块的功能和迁移细节，以便后续参考。
   - 使用Markdown或其他文档格式记录迁移过程中的决策和注意事项。

## 注意事项
- VS.NET不支持直接运行VBA代码，需在Excel中测试。
- 在迁移过程中，保持与团队成员的沟通，确保代码的一致性和质量。
- 定期进行代码审查，确保代码符合最佳实践和设计原则。

## 贡献
如有任何问题或建议，请联系项目维护者。

## 许可证
本项目遵循MIT许可证。
