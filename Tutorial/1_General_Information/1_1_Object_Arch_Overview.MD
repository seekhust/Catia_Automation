## 基础设施 (Infrastructure)
### 对象架构概览 (Object Architecture Overview)
在本节中，你将了解 CAA V5 公开的自动化对象是如何描述的，并获取以下信息：

-对象图、对象描述和对象参考： 展示对象在何处以及如何被描述。

-关于对象、集合、属性和方法： 帮助你理解脚本编写所需的对象技术基础。

-关于继承、聚合和对象模型： 描述编写宏（Macros）必须掌握的两种对象机制。

#### 对象图、对象描述和对象参考
我们使用对象图（Object Diagrams）和对象描述来定义对象及其属性、方法和对象间的关系。对象图描述了对象是如何连接在一起的总体结构：

图例说明：

集合 (Collection)： 黄色框表示。

抽象对象 (Abstract object)： 蓝色框加斜体表示。

对象 (Object)： 蓝色框表示。

箭头图标用于在不同层级的图表间展开或返回。

从上图中你可以了解到：
Application 对象聚合了一个 Documents 集合。Documents 集合可以包含任意数量的文档（Document）。Document 对象是一个抽象对象，只有它的派生类型才能被实际创建，即：

PartDocument：用于零件设计 (Part Design)

ProductDocument：用于装配设计 (Assembly Design)

DrawingDocument：用于工程制图 (Drafting)

以及 AnalysisDocument (分析)、ProcessDocument (工艺)、MaterialDocument (材料) 等。

所有这些文档都继承了由 Document 对象聚合的 Selection（选择）对象。但只有 PartDocument 拥有聚合的 Part 对象，只有 DrawingDocument 拥有聚合的 DrawingSheets 集合。

#### 关于对象、集合、属性和方法
像 Visual Basic 这样的脚本语言依赖于对象。在 CATIA 或 DELMIA 中，几乎所有访问的数据都是对象（文档、窗口、零件、草图、凸台甚至线条）。

对象 (Object)： 在图中用蓝色框表示。

集合 (Collection)： 是一种包含其他对象的对象（通常是同类型对象）。

例如：Documents 集合包含多个文档。

命名规则： 集合通常使用复数形式。

索引： 集合索引从 1 开始（而非 0）。可以通过索引或名称访问其中的元素。

属性 (Property)： 对象的特征。

例如：Document.FullName 存储文件的完整路径。

读写性： 有些属性是只读的，有些可以修改。

访问方式： 对象.属性。例如：DocName = CATIA.ActiveDocument.FullName。

方法 (Method)： 请求对象执行的动作。

例如：CATIA.ActiveDocument.Save。

参数： 有些方法需要参数，如 SaveAs "NewName"。

Function vs Sub： 在参考文档中，返回值的称为 Function（函数），不返回值的称为 Sub（子程序）。

#### 关于继承、聚合和对象模型
继承（Inheritance）和聚合（Aggregation）是对象之间最主要的两种关系。

1. 继承 (Inheritance)
一个对象继承自另一个“基对象”（Base Object），它会自动拥有基对象的所有属性和方法，并能添加自己的特性。

作用： 提取共性，实现专业化。

符号： 在图中用特定的继承符号表示。

示例： PartDocument 继承自 Document。它拥有 Document 的 Save 方法，同时也拥有自己特有的 PartNumber 属性。

2. 抽象对象 (Abstract Object)
图中以蓝色斜体显示的框（如 Document）是抽象对象。它定义了通用属性和方法，但不能被直接创建。你只能创建它的非抽象派生类（如 PartDocument）。

3. 聚合 (Aggregation)
聚合是一个对象包含另一个对象的能力。

示例： Application 聚合了 Documents 集合。

角色 (Role)： 有时聚合线上会有描述。例如，一个“棱柱体”（Prism）对象聚合了两个“极限”（Limit）对象（第一极限和第二极限）。

双向性： 聚合通常是双向的，因为子对象可以通过 Parent 属性找到它的父对象（聚合它的对象）。


## 总结
API 的核心： 通过 CATIA 根对象，层层深入访问属性和方法。

结构逻辑： 所有的操作都遵循“对象模型图”定义的层级关系。对于 Senior Engineer 来说，理解这一套“对象树”是编写高效自动化脚本（VBA/Python）的关键。