## 基础设施

### 创建一个新的 CATIA 文档

此宏向您展示如何创建一个新的 CATIA 文档。它详细说明了等同于交互式操作中 **文件 -> 新建 (File -> New)** 命令的编程实现方法。

`CAAInfCreateDocument` 在 CATIA [1] 中启动。不需要事先打开任何文档。
`CAAInfCreateDocument.CATScript` 位于 `CAAScdInfUseCases` 模块中。可以执行该宏（仅限 Windows 系统）。

`CAAInfCreateDocument` 包含一个步骤：

* 创建一个新的 CATIA 文档

---

#### 创建一个新的 CATIA 文档

```vb
  ...
    'Create a new part document.
    'Add the new document to the end of the collection of documents.
    'Create and display a new window for the new document.
    'Activate the new document and the window.
     Dim oNewPartDoc As Document
     Set oNewPartDoc = CATIA.Documents.Add("Part")
  ...

```

使用 `Documents` 集合（`Documents` 对象）的 `Add` 方法创建一个新的 CATIA 文档。必须向 `Add` 方法指定要创建的文档类型。此类型可以是零件 (part)、装配体 (assembly) 或工程图 (drawing)。因此，您需要分别使用 `Part`、`Product` 或 `Drawing` 类型作为 `Add` 方法的参数。

在本例中，创建的是一个 `Part`（零件）文档。新文档被添加到文档集合的末尾。同时，会为该文档创建并显示一个新窗口。文档和窗口都会被自动激活。

请注意，新文档的生命周期仅限于当前的 CATIA 会话。若要永久保留，必须保存该文档。有关如何保存新文档的演示，请参见 `CAAInfSaveAsDocument` [2] 示例。同时请注意，新创建的文档也应在会话结束时关闭。有关如何关闭新文档的演示，请参见 `CAAInfCloseDocument` [3] 示例。

---

### 总结

本用例演示了如何在交互式会话期间创建一个新的 CATIA 文档。