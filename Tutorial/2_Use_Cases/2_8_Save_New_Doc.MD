## 基础设施

### 保存新的 CATIA 文档

本宏指南将向您展示如何保存当前会话中新建的 CATIA 文档。若要将现有的 CATIA 文档保存在同一路径并使用原文件名，请参阅 CAAInfSaveDocument [1] 示例以获取详细说明。

CAAInfSaveAsDocument 在 CATIA [2] 中启动。必须在 CATDocView 中找到名为 "CAAInfReadDocument.CATPart" 的现有文档。
CAAInfSaveAsDocument.CATScript 位于 CAAScdInfUseCases 模块中。执行宏（仅限 Windows）。

---

### CAAInfSaveAsDocument 包含五个步骤：

1. 检索 CATDocView 环境变量并检查其赋值情况
2. 检索 CATTEMP 环境变量并检查其赋值情况
3. 保存当前的活动 CATIA 文档
4. 通过对象变量保存指定的 CATIA 文档
5. 通过名称保存指定的 CATIA 文档

---

### 1. 检索 CATDocView 环境变量并检查其赋值情况

' 可选：允许在任何安装位置找到示例

```vb
    Dim sDocPath  As String
    sDocPath=CATIA.SystemService.Environ("CATDocView")
    If (Not CATIA.FileSystem.FolderExists(sDocPath)) Then
            Err.Raise 9999,,"No Doc Path Defined"
    End If

```

定义 `sDocPath` 变量以检索 CATDocView 环境变量的值，该变量保存了下文所用 Part 文档的存储路径。同时，检查该变量是否已赋值。

---

### 2. 检索 CATTEMP 环境变量并检查其赋值情况

' 可选：允许指定文档的保存位置

```vb
    Dim sTmpPath  As String
    sTmpPath=CATIA.SystemService.Environ("CATTemp")
    If (Not CATIA.FileSystem.FolderExists(sTmpPath)) Then
        Err.Raise 9999,,"No Tmp Path Defined"
    End If

```

为当前的 CATTEMP 定义一个变量，用于确定保存文档的存储路径。同时，验证该变量是否已赋值。

[回到顶部]

---

### 3. 保存当前的活动 CATIA 文档

...

```vb
    '创建一个新的 Part 文档。
    '将新文档添加到文档集合的末尾。
    '为新文档创建并显示一个新窗口。
    '激活新文档及其窗口。
     Dim oFirstNewPartDoc As Document
     Set oFirstNewPartDoc = CATIA.Documents.Add("Part")

    '刚刚创建的文档即为活动文档。
    '保存新文档。
     CATIA.ActiveDocument.SaveAs(sTmpPath & _
       "\CAAInfWriteDocument1.CATPart")

```

...
使用 `Documents` 集合（`Documents` 对象）的 `Add` 方法创建一个新的 CATIA 文档。有关如何创建新文档的说明，请参阅 CAAInfCreateDocument [3] 示例。由于 `Add` 方法会自动激活新文档，因此您可以对 CATIA 应用程序的 `ActiveDocument` 属性执行 `SaveAs` 方法来保存它：文档将保存到 `SaveAs` 方法参数中指定的存储路径和位置。如果已存在同名文档，系统将提示您是否要覆盖该文档。请注意，在退出会话之前，还应关闭该文档。

[回到顶部]

---

### 4. 通过对象变量保存指定的 CATIA 文档

...

```vb
   '创建第二个新的 Part 文档。
     Dim oSecondNewPartDoc As Document
     Set oSecondNewPartDoc = CATIA.Documents.Add("Part")

    '使用为该文档定义的变量名保存新文档。
     oSecondNewPartDoc.SaveAs(sTmpPath & _
       "\CAAInfWriteDocument2.CATPart")

```

...
保存新 CATIA 文档的第二种方法是直接对已分配用于处理该文档的**变量**执行 `SaveAs` 方法：文档将保存到 `SaveAs` 方法参数中指定的路径和位置。如果同名文档已存在，将弹出覆盖确认提示。同样，在退出会话前应关闭文档。

[回到顶部]

---

### 5. 通过名称保存指定的 CATIA 文档

...

```vb
     '打开一个现有文档。
      Dim iThirdPartDoc As Document
      Set iThirdPartDoc = CATIA.Documents.Open(sDocPath & _
        "\online\CAAScdInfUseCases\samples\CAAInfReadDocument.CATPart")

     '通过指定名称保存新文档。
      CATIA.Documents.Item("CAAInfReadDocument.CATPart").SaveAs(sTmpPath & _
        "\CAAInfWriteDocument3.CATPart")

```

...
最后，第三种保存方式是对从 `Documents` 集合中检索到的现有文档（通过 `Item` 方法指定名称）执行 `SaveAs` 方法：一个与现有文档完全相同的新文档将根据 `SaveAs` 参数中指定的路径和位置进行保存。如果同名文件已存在，会提示是否覆盖。请注意，退出会话前应将其关闭。

[回到顶部]

---

### 简而言之

本用例展示了在交互式会话中保存新 CATIA 文档的三种方式：

* 使用 CATIA 应用程序的 **ActiveDocument** 属性
* 使用分配给该文档的 **对象变量**
* 使用通过 **Item** 方法从文档集合中检索到的现有文档的 **名称**

在上述所有三种情况下，都必须明确指定新文档存储的路径名和位置。