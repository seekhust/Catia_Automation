## 基础设施

### 访问文件和文件夹

本宏演示了如何使用 CAA V5 的可移植自动化对象来访问文件和文件夹。
它创建了一个文本文件，在其中写入数据，复制该文件，然后从副本中重新提取数据。

`CAAInfFileAccess` 在 CATIA 中启动 [1]。无需打开任何文档。
`CAAInfFileAccess.CATScript` 位于 `CAAScdInfUseCases` 模块中。可以执行宏（直接执行宏的功能仅在 Windows 上有效）。

`CAAInfFileAccess` 包含以下五个步骤：

1. 序言 (Prolog)
2. 必要时删除现有文件
3. 创建输入文件
4. 复制输入文件
5. 从输出文件提取数据

---

#### 序言 (Prolog)

```vb
  ...
    Dim sLF as String
    sLF = Chr(10) 
    Dim sMessage as String
    sMessage = InputBox ("Enter a message", "Message", "Hello World")

    ' ------------------------------------------
    ' 获取文件系统对象 (Get the file system object)
    Dim oFileSys as FileSystem
    Set oFileSys = CATIA.FileSystem

    ' ------------------------------------------
    ' 获取用于临时文件的文件夹 (Retrieve a folder for temporary files)
    Dim sTmpPath as String 
    sTmpPath=CATIA.SystemService.Environ("CATTemp")
    If (Not oFileSys.FolderExists(sTmpPath)) Then
      Err.Raise 9999,,"No Tmp Path Defined"
    End If
  ...

```

为换行符 (LineFeed) 定义了一个变量。`InputBox` 函数提示用户输入一条消息字符串，其默认值为 "Hello World"。

`FileSystem` 对象是所有文件/文件夹访问的根对象。输入和输出文件将在专用于存放临时 CATIA 数据的文件夹中创建。我们使用 `SystemService` 对象，通过读取 `CATTemp` 环境变量来获取该文件夹的完整路径。
使用 `FileSystem` 对象的 `FolderExists` 方法检查此文件夹是否存在。

---

#### 必要时删除现有文件 (Deleting Existing Files if Needed)

```vb
  ...
    ' ------------------------------------------
    ' 删除可能存在的输入和输出文件 (Delete possibly existing input and output files)
    Dim sFilOu As String ' 输出文件的完整路径
    sFilOu = sTmpPath & "/caatmpfilou.txt"
    If (oFileSys.FileExists(sFilou)) Then 
        oFileSys.DeleteFile sFilOu
    End If
  ...

```

使用 `FileSystem` 对象的两个方法 `FileExists` 和 `DeleteFile` 来检查输出文件是否存在，并在必要时将其删除。随后对输入文件执行相同的操作（此处代码省略了输入文件的删除部分）。

---

#### 创建输入文件 (Creating the Input File)

```vb
...
     ' ---------------------------------------
    ' 创建文件 FilIn (Create file FilIn)
      Dim oFilIn As File    
      Set oFilIn = oFileSys.CreateFile(sFilIn, FALSE)
    Dim oStream As TextStream
    Set oStream = oFilIn.OpenAsTextStream("ForWriting")
    oStream.Write "<MESSAGE>"  & sLF
    oStream.Write "<VALUE>"
    oStream.Write sMessage 
    oStream.Write "</VALUE>"   & sLF
    oStream.Write "</MESSAGE>" & sLF
    oStream.Close
  ...

```

变量 `sFilIn` 包含输入文件的完整路径。使用 `FileSystem` 对象的 `CreateFile` 方法创建此文件，该方法返回一个 `File` 对象。`OpenAsTextStream` 方法允许从 `File` 对象返回一个 `TextStream`（文本流）对象。`"ForWriting"` 选项指定该文件的内容将被修改。这个 `TextStream` 对象使得以文本流的方式操作文件内容成为可能。

连续使用 `Write` 方法来填充输入文件。在这个例子中，消息被类似 XML 的标签包裹起来。`sLF` 变量用于指定行尾的换行。然后使用 `Close` 方法关闭文本流。

---

#### 复制输入文件 (Duplicating the Input File)

```vb
...
     ' ---------------------------------------
    ' 将 FilIn 复制到 FilOu (Duplicate FilIn in FilOu)
    oFileSys.CopyFile sFilIn, sFilOu, FALSE
... 

```

`FileSystem` 对象的 `CopyFile` 方法创建了输入文件的一个副本。它使用包含在 `sFilIn` 变量中的完整文件名作为源，并使用 `sFilOu` 变量中的内容作为副本的完整文件名。

---

#### 从输出文件提取数据 (Extracting Data from the Output File)

```vb
...
    ' ---------------------------------------
    ' 从输出文件中获取结果 (Get the result from the output file)
      Dim oFilOu As File
    Set oFilOu = oFileSys.GetFile(sFilOu)
    Set oStream = oFilOu.OpenAsTextStream("ForReading")
    Dim sBuffer As String
    sMessage = ""
    sBuffer = oStream.ReadLine
    Do  Until oStream.AtEndOfStream
        sMessage = sMessage & sBuffer
        sBuffer = oStream.ReadLine
    Loop

    oStream.Close

    msgbox sMessage 

```

使用 `GetFile` 方法获取输出文件的 `File` 对象。然后使用 `OpenAsTextStream` 获取文本流，但这次指定的是 `"ForReading"`（只读）模式。在循环中使用 `ReadLine` 方法读取输出文件的每一行。每读取一行，就将其连接到 `sMessage` 字符串上。使用 `AtEndOfStream` 方法来检测是否到达文件末尾。

此时，`sMessage` 变量包含了输出文件的全部内容。使用 `MsgBox` 函数在消息框中显示这些内容。

![Acess_Files](image-7.png)

---

### 总结 (In Short)

本用例演示了如何使用 CAA V5 的 `FileSystem`、`SystemService`、`File` 和 `TextStream` 自动化对象来访问文件和文件夹。