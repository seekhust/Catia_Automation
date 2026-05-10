## 基础设施 (Infrastructure)

### 关闭 CATIA 文档

本宏示例演示了如何关闭当前会话中现有的 CATIA 文档。它详细介绍了与 **文件 (File) -> 关闭 (Close)** 命令等效的编程操作。

CAAInfCloseDocument 在 CATIA [1] 中启动。在 CATDocView 路径下必须能找到一个名为 "CAAInfReadDocument.CATPart" 的现有文档。
`CAAInfCloseDocument.CATScript` 位于 CAAScdInfUseCases 模块中。可通过**执行宏**（仅限 Windows）运行。

CAAInfCloseDocument 包含四个步骤：

1. 获取 CATDocView 环境变量并检查其有效性
2. 关闭当前活动 (Active) 的 CATIA 文档
3. 通过对象变量关闭指定的 CATIA 文档
4. 通过文档名称关闭指定的 CATIA 文档

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

### 2. 关闭当前活动的 CATIA 文档

```vb
  ...
    ' 打开文档
     Dim iPartDoc As Document
     Set iPartDoc = CATIA.Documents.Open(sDocPath & _
       "\online\CAAScdInfUseCases\samples\CAAInfReadDocument.CATPart")

    ' 关闭活动文档，即刚刚打开的文档
     CATIA.ActiveDocument.Close()

  ...

```

使用 **Documents** 集合（Documents 对象）的 **Open** 方法打开现有的 CATIA 文档。有关如何打开文档的描述，请参阅 CAAInfOpenDocument [2] 示例。由于 `Open` 方法会自动激活所打开的文档，因此你可以使用 CATIA 应用程序 **ActiveDocument** 属性的 **Close** 方法将其关闭：文档将从 `Documents` 集合中移除，包含该文档的所有窗口也将被关闭并从 `Windows` 集合中移除。

[回到顶部]

---

### 3. 通过对象变量关闭指定的 CATIA 文档

```vb
  ...
    ' 再次打开同一个文档
     Set iPartDoc = CATIA.Documents.Open(sDocPath & _
       "\online\CAAScdInfUseCases\samples\CAAInfReadDocument.CATPart")

    ' 使用为该文档定义的变量来关闭它
     iPartDoc.Close()

  ...

```

关闭现有 CATIA 文档的第二种方法是：直接在已分配用于处理该文档的**对象变量**上执行 **Close** 方法。文档将从 `Documents` 集合中移除，所有相关窗口也将被关闭并移除。

[回到顶部]

---

### 4. 通过文档名称关闭指定的 CATIA 文档

```vb
  ...
     ' 第三次打开同一个文档
      Set iPartDoc = CATIA.Documents.Open(sDocPath & _
        "\online\CAAScdInfUseCases\samples\CAAInfReadDocument.CATPart")

     ' 通过指定文档名称来关闭它
      CATIA.Documents.Item("CAAInfReadDocument.CATPart").Close()
  ...

```

最后，第三种关闭 CATIA 文档的方法是对文档的**名称**执行 **Close** 方法。该名称通过 `Documents` 集合的 **Item** 方法检索。同样，文档将从 `Documents` 集合中移除，相关窗口也将被关闭。

[回到顶部]

---

### 简结

本案例展示了在交互式会话中关闭现有 CATIA 文档的三种方式：

* 使用 CATIA 应用程序的 **ActiveDocument** 属性。
* 使用分配给该文档的**对象变量**。
* 通过 **Item** 方法从 `Documents` 集合中根据**文档名称**检索并关闭。