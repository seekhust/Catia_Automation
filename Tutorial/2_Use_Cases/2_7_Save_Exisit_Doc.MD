## 基础设施

### 保存现有的 CATIA 文档

本宏指南将向您展示如何保存当前会话中已存在的 CATIA 文档。如需保存新文档，请参阅 CAAInfSaveAsDocument [1] 示例以获取详细说明。本示例详细介绍了与 **文件 -> 保存 (File -> Save)** 命令等效的编程操作。

CAAInfSaveDocument 在 CATIA [2] 中启动。必须在 CATDocView 中找到名为 "CAAInfReadDocument.CATPart" 的现有文档。
CAAInfSaveDocument.CATScript 位于 CAAScdInfUseCases 模块中。执行宏（仅限 Windows）。

---

### CAAInfSaveDocument 包含四个步骤：

1. 检索 CATDocView 环境变量并检查其赋值情况
2. 保存当前的活动 CATIA 文档
3. 通过对象变量保存指定的 CATIA 文档
4. 通过名称保存指定的 CATIA 文档

---

### 1. 检索 CATDocView 环境变量并检查其赋值情况

'可选：允许在任何安装位置找到示例

```vb
      Dim sDocPath  As String
      sDocPath=CATIA.SystemService.Environ("CATDocView")

      If (Not CATIA.FileSystem.FolderExists(sDocPath))  Then
         Err.Raise 9999,,"No Doc Path Defined"
      End If

```

定义 `sDocPath` 变量以检索 CATDocView 环境变量的值，该变量保存了存储下文所用 Part 文档的路径。同时，检查该变量是否已赋值。

[回到顶部]

---

### 2. 保存当前的活动 CATIA 文档

...

```vb
    '打开文档并将其作为文档集合的最后一项添加。
    '为文档创建并显示一个新窗口。
    '激活文档及其窗口。
     Dim iPartDoc As Document
     Set iPartDoc = CATIA.Documents.Open(sDocPath & _
       "\online\CAAScdInfUseCases\samples\CAAInfReadDocument.CATPart")

    '刚刚打开的文档即为活动文档（Active Document）。
    '保存活动文档，然后将其关闭。
     CATIA.ActiveDocument.Save()
     CATIA.ActiveDocument.Close()

```

...
使用 `Documents` 集合（`Documents` 对象）的 `Open` 方法打开一个现有的 CATIA 文档。有关如何打开文档的说明，请参阅 CAAInfOpenDocument [3] 示例。由于 `Open` 方法会自动激活所打开的文档，因此您可以使用 CATIA 应用程序 `ActiveDocument` 属性的 `Save` 方法来保存它：文档将按原名称保存到原存储位置。请注意，在退出会话之前，还应关闭该文档。有关如何关闭文档的说明，请参阅 CAAInfCloseDocument [4] 示例。

[回到顶部]

---

### 3. 通过对象变量保存指定的 CATIA 文档

...

```vb
    '再次打开同一个文档。
     Set iPartDoc = CATIA.Documents.Open(sDocPath & _
       "\online\CAAScdInfUseCases\samples\CAAInfReadDocument.CATPart")

    '使用为该文档定义的变量名保存刚刚打开的文档。
    '然后，以同样的方式关闭文档。
     iPartDoc.Save()
     iPartDoc.Close()

```

...
第二次打开同一个文档。保存现有 CATIA 文档的另一种方法是直接对已分配用于处理该文档的**变量**执行 `Save` 方法：文档将按原名称保存到原存储位置。同样，在退出会话前应关闭文档。

[回到顶部]

---

### 4. 通过名称保存指定的 CATIA 文档

...

```vb
     '第三次打开同一个文档。
     Set iPartDoc = CATIA.Documents.Open(sDocPath & _
       "\online\CAAScdInfUseCases\samples\CAAInfReadDocument.CATPart")

    '通过指定文档名称来保存文档，然后以同样方式关闭。
     CATIA.Documents.Item("CAAInfReadDocument.CATPart").Save()
     CATIA.Documents.Item("CAAInfReadDocument.CATPart").Close()

```

...
最后一次打开同一个文档。第三种保存 CATIA 文档的方法是对从 `Documents` 集合中检索到的文档执行 `Save` 方法，检索时需向 `Item` 方法的参数提供**文档名称**：文档将按原名称保存到原存储位置。请注意，退出前应关闭文档。

[回到顶部]

---

### 简而言之

本用例展示了在交互式会话中保存现有 CATIA 文档的三种方式：

* 使用 CATIA 应用程序的 **ActiveDocument** 属性
* 使用分配给该文档的 **对象变量**
* 使用通过 **Item** 方法从文档集合中按 **名称** 检索到的文档对象

在以上三种情况下，文档都会以相同的路径和名称进行保存。