## 基础设施 (Infrastructure)

### 打开现有的 CATIA 文档

本宏示例演示了如何打开一个现有的 CATIA 文档。它详细介绍了与菜单栏中 **文件 (File) -> 打开 (Open)** 命令等效的编程操作。

CAAInfOpenDocument 在 CATIA [1] 中启动。在 CATDocView 路径下必须能找到一个名为 "CAAInfReadDocument.CATPart" 的现有文档。
`CAAInfOpenDocument.CATScript` 位于 CAAScdInfUseCases 模块中。可以通过**执行宏**（仅限 Windows）运行。

CAAInfOpenDocument 包含两个步骤：

1. 获取 CATDocView 环境变量并检查其有效性
2. 打开现有的 CATIA 文档

---

### 1. 获取 CATDocView 环境变量并检查其有效性

```vb
' 可选：允许在安装路径下找到示例文件
     Dim sDocPath As String
     sDocPath=CATIA.SystemService.Environ("CATDocView")

     If (Not CATIA.FileSystem.FolderExists(sDocPath)) Then
        Err.Raise 9999,,"No Doc Path Defined"
     End If

```

定义 `sDocPath` 变量以检索 CATDocView 环境变量的值。该变量保存了存储后续所用 Part 文档的文档路径。同时，检查该变量是否已赋值且文件夹是否存在。

[回到顶部]

---

### 2. 打开现有的 CATIA 文档

```vb
  ...
    Dim iPartDoc As Document
    Set iPartDoc = CATIA.Documents.Open(sDocPath & _
       "\online\CAAScdInfUseCases\samples\CAAInfReadDocument.CATPart")
  ...

```

使用 **Documents** 集合（Documents 对象）的 **Open** 方法打开现有的 CATIA 文档。该方法的唯一参数必须包含指定文档所在位置的完整存储路径和文件名。

这一条语句会打开存储在文件中的文档，并将其添加到文档集合的末尾。它还会为该文档创建并显示一个新窗口。文档和窗口都会被同时激活。

**注意：** 如果对打开的文档进行了任何更改，应在会话结束时保存文档。有关如何保存已打开文档的说明，请参阅 CAAInfSaveDocument [2] 示例。在任何情况下，已打开的文档都应在会话结束时关闭。有关如何关闭文档的说明，请参阅 CAAInfCloseDocument [3] 示例。

[回到顶部]

---

### 简结

本案例展示了如何在交互式会话期间打开一个现有的 CATIA 文档。