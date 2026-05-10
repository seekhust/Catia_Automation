# CATIA CAA V5 自动化对象模型指南

## 1. 架构概览 (Object Architecture)

CAA V5 的自动化对象模型采用层级化的结构，主要通过以下三种方式进行描述：
* **对象图 (Object Diagrams)**：描述对象间的整体连接结构。
* **对象描述 (Object Descriptions)**：解释如何使用主要对象。
* **对象参考 (Object References)**：包含继承、属性和方法的完整参考手册。

### 对象类型说明
| 类型 | 图例颜色 | 说明 |
| :--- | :---: | :--- |
| **对象 (Object)** | 蓝色框 | 代表具体的功能实体（如 Pad, Line）。 |
| **抽象对象 (Abstract)** | 蓝色斜体 | 定义共有属性，不能被直接创建（如 Document）。 |
| **集合 (Collection)** | 黄色框 | 包含同类对象的容器（如 Documents）。 |

---

## 2. 核心关系：继承与聚合

### 继承 (Inheritance)
子对象继承基对象的属性和方法。这有助于在保持通用的同时实现专业化。
* **示例**：`PartDocument` 继承自 `Document`。它不仅拥有 `Save` 方法（通用），还拥有 `PartNumber` 属性（零件专用）。

### 聚合 (Aggregation)
一个对象包含另一个对象的能力，通常用于描述层级树。
* **示例**：`Application` 聚合了 `Documents` 集合，而 `Documents` 集合又聚合了多个 `Document` 对象。
* **父子溯源**：由于聚合是双向的，你可以通过 `Parent` 属性找到任何对象的上级。

---

## 3. 基础组件：属性与方法

### 属性 (Property)
用于刻画对象的特征。
* **访问语法**：`对象.属性`
* **示例**：`CATIA.ActiveDocument.FullName`（获取当前文档的全路径）。
* **读写性**：分为“可读写”和“只读”。

### 方法 (Method)
请求对象执行的动作。
* **Function**：执行动作并返回一个值（或对象）。
* **Sub**：仅执行动作，不返回值。
* **示例**：`CATIA.ActiveDocument.Save()`。

---

## 4. 集合操作技巧 (About Collections)

集合是自动化开发中最常处理的对象类型之一。

### 访问元素 (Item Method)
你可以通过**排名索引**或**对象名称**来获取元素：
* **注意**：在 COM 接口中，集合排名从 **1** 开始。

```vb
' 方式 A：通过排名 (获取第6个文档)
Dim Doc As Document
Set Doc = CATIA.Documents.Item(6)

' 方式 B：通过名称
Dim Doc As Document
Set Doc = CATIA.Documents.Item("MyPart_V1")
```

### 扫描集合 (Scanning)
以下是三种常用的遍历集合方式：

方式 1：For 循环 (利用 Count 属性)
```VB.Net
For I = 1 To CATIA.Documents.Count
    MsgBox CATIA.Documents.Item(I).Name
Next
```
方式 2：Do Until 循环
```VB.Net
I = 1
Do
    MsgBox CATIA.Documents.Item(I).Name
    I = I + 1
Loop Until I > CATIA.Documents.Count
```
方式 3：For Each 遍历 (推荐)
这是最简洁的方式，无需手动维护计数器：

```VB.Net
For Each Doc In CATIA.Documents
    MsgBox Doc.Name
Next
```
5. 补充：Python (pycatia) 实现参考
如果你使用 Python 进行开发，上述逻辑将变得更加简洁：

```Python
from pycatia import catia

# 初始化连接
caa = catia()
documents = caa.documents

# 遍历集合
for doc in documents:
    print(doc.name)

# 通过名称获取特定项
my_doc = documents.item("MyPart_V1")
```