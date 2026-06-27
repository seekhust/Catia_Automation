# README.MD

## 简单介绍


简单来说，**CATIA V5 Automation**（自动化）是达索系统提供的一种编程接口（API），允许用户通过编写脚本或外部程序来操控 CATIA 软件。

它的核心目的在于：**让计算机代替人工完成重复性、繁琐的操作，从而提高设计效率和减少人为错误。**

以下是关于 CATIA V5 Automation 的核心要点介绍：

### 1. 它能做什么？

* **自动执行重复任务**：例如一键批量重命名成百上千个零件、自动生成标准的工程图纸、或者批量转换文件格式（如将 .CATPart 转为 .stp）。
* **建立参数化设计模板**：通过外部输入（如 Excel 表格里的尺寸数据），驱动 CATIA 自动生成对应的三维模型。
* **定制专用工具**：根据公司特定的设计规范，开发出专门的工具栏或功能按钮。
* **与其他软件联动**：实现 CATIA 与 Excel、Word、数据库或 ERP 系统的双向数据交互（例如自动导出 BOM 表到 Excel）。

### 2. 核心架构：COM 接口

CATIA V5 的自动化是基于微软的 **COM (Component Object Model)** 技术。这意味着 CATIA 就像一个“服务提供者”，它把内部的各种功能（如零件、草图、约束等）封装成一个个对象（Objects）暴露给外部，程序员可以通过控制这些对象来指挥 CATIA 工作。

### 3. 支持的编程语言

你可以根据复杂程度选择不同的方式：

* **脚本语言 (VBScript/CATScript)**：简单易学，通常用于编写宏（Macros）。CATIA 内置了录制功能，可以将你的操作直接录制成脚本。
* **VBA (Visual Basic for Applications)**：集成在 CATIA 内部的开发环境，带界面设计功能，适合开发有交互界面的小工具。
* **外部编程语言 (VB.NET, C#)**：通过 Visual Studio 开发。适合开发大型、高性能的复杂系统，或需要与企业内部其他软件集成的场景。

### 4. 关键概念：对象树（Object Tree）

在 Automation 中，CATIA 的一切都被看作是层级结构。要操作一个特征，你需要顺着树往下找。例如：
`Application (CATIA软件) -> Documents (文档集合) -> PartDocument (零件文档) -> Part (零件) -> Bodies (几何体) -> ...`

### 5. Automation vs. CAA 的区别

这是初学者最容易混淆的点：

* **Automation (自动化)**：门槛低、开发快，主要使用 VB 系列语言。适合 90% 的日常自动化需求。
* **CAA (Component Application Architecture)**：基于 C++ 的底层开发。门槛极高，可以修改 CATIA 的底层行为，开发全新的功能模块。通常只有专业的软件开发人员或大型企业的核心团队才会使用。

**总结：**
如果你经常觉得在 CATIA 里做某些操作很“烦”、很“重复”，那么 **Automation** 就是为你准备的工具。它能把原本需要几个小时的手工劳动缩短到几秒钟完成。

## 更新记录 Update Log:

2026-5-10 为了了解Catia的自动化编程而创建
 [官方介绍：]https://catiadesign.org/_doc/V5Automation/

 https://catiadesign.org/_doc/V5Automation/online/CAAScdPriUseCases/CAAPriCamShaft.htm

CATIA的版本为Catia V5-6 Release 2020
![Catia Version](image-6.png)

[pyCatia官方网站]https://pycatia.readthedocs.io/en/latest/introduction.html#introduction

### 自动化对象总览

![自动化对象框架总览](image.png)
![零件自动化对象](image-1.png)
![产品自动化对象](image-2.png)
![放样自动化对象](image-3.png)
![材料自动化对象](image-4.png)
![部分设置控制器自动化对象](image-5.png)