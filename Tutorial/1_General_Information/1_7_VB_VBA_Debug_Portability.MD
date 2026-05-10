## 基础设施

### 关于 VB、VBA、调试与可移植性

#### VB 和 VBA 的技巧

自 V5R7 起，基于 V5 平台的 Dassault Systemes 产品在 Unix 平台上不再使用 Basic Script 脚本引擎。**CATScript** 语言被保留用于兼容性，在移除类型信息后由 VBScript 脚本引擎处理。

这对宏本身没有影响，但在将以 **CATScript** 语言录制的宏复制/粘贴到 VB/VBA 项目中时，您可能会遇到以下问题。

Basic Script 代码和 VB 代码的**类型要求**不兼容。当一个方法返回抽象对象时（例如 `ActiveDocument` 方法返回一个 `Document`），Basic Script 期望存储返回值的变量类型与方法签名声明的类型一致：

```vb
Dim oDoc As Document
Set oDoc = CATIA.ActiveDocument ' 返回一个 Document

```

如果该文档是一个 `PartDocument`，您随后可以使用特定的方法或属性（例如 `Part` 属性），这些属性在通用的 `Document` 中不存在，仅存在于 `PartDocument` 中：

```vb
Dim oPart As Part
Set oPart = oDoc.Part

```

在使用 VB 时，您需要一个 `PartDocument` 对象才能使用 `Part` 属性：

```vb
Dim oDoc As PartDocument
Set oDoc = CATIA.ActiveDocument ' 返回一个 Document
Dim oPart As Part
Set oPart = oDoc.Part

```

对于文档示例，我们尽可能选择 VBScript/Basic Script 通用的语法。因此，当从文档中复制示例到 VB 项目时，您需要留意这些**虚拟对象**的处理。

您还需要注意**带有数组参数的方法**。例如以下在 VBScript 或 BasicScript 中提取并显示 `StringParameter` 多个可能值的示例：

```vb
Dim strParam1 As StrParam
Set strParam1 = parameters1.Item("STRING")

iSize = strParam1.GetEnumerateValuesSize()
Redim myArray(iSize-1)
strParam.GetEnumerateValues myArray
For i = 0 To iSize-1
    msgbox myArray(i)
Next

```

根据您的 VB/VBA 级别，将这段代码复制到 VB 项目可能会导致编译错误，因为 VB 可能无法处理我们为数组类型使用的签名，并发出以下消息：

> "Function or interface marked as restricted, or the function uses an automation type not supported by Visual Basic."（函数或接口被标记为受限，或函数使用了 Visual Basic 不支持的自动化类型。）

一个简单的解决方法是**取消应用该方法的变量的类型定义**：

```vb
Dim strParam1 ' As StrParam 
Set strParam1 = parameters1.Item("STRING")

iSize = strParam1.GetEnumerateValuesSize()
Redim myArray(iSize-1)
strParam.GetEnumerateValues myArray
For i = 0 To iSize-1
    msgbox myArray(i)
Next

```

当数组是方法的唯一参数时，应避免使用带括号的语法：

```vb
Redim myArray(15)
strParam.GetEnumerateValues (myArray)

```

对于非函数的方法（即没有返回值的方法），这种括号语法要求通过引用传递参数，这在某些情况下可能无法正常工作。正确的语法是**不带括号**：

```vb
strParam.GetEnumerateValues myArray

```

或者使用 **Call** 关键字：

```vb
call strParam.GetEnumerateValues (myArray)

```

---

#### 脚本调试器 (Script Debugger)

在开发进程内（In-process）宏时，您可以使用 Microsoft (R) Script Debugger。安装后，脚本中的错误或 `Stop` 命令将移交给调试器：

```vb
For i = 0 To Ubound(myArray)
    Stop
    msgbox myArray(i)
Next

```

---

#### 可移植性 (Portability)

在 Windows 平台上，CAA V5 脚本功能依赖于平台提供的 VBScript 引擎，要求最低版本为 VBScript 5.0。在 Unix 平台上，安装 V5 时会安装 VBScript 3.0 引擎。

因此，**可移植宏必须避免使用 VBScript 3.0 之后出现的新特性**。

---

#### 版本间兼容性 (Inter releases compatibility)

**虚函数表兼容性 (Virtual Function Tables compatibility)**
CAA V5 自动化对象模型随每个新版本演进。即使只是添加了新方法，也可能影响您的应用程序。

例如以下代码：

```vb
Dim oObject
oObject.DoThis

```

由于 `oObject` 未定义类型，将执行**后期绑定 (Late binding)**。而以下代码：

```vb
Dim oObject As SpecificTypeOfObject
oObject.DoThat

```

将执行**早期绑定 (Early binding)**。如果在版本更新中，虚函数表中方法的顺序发生了变化，应用程序可能会调用错误的方法，导致难以调试。

如果重新编译项目无法解决问题，请尝试：

1. 打开 VB 或 VBA 项目。
2. 移除所有对 V5 类型库（Type Libraries）的引用。
3. 保存并关闭项目。
4. 重新打开项目并再次添加所需的引用。

**过时的类型库 (Obsolete Typelibs)**
某些类型库可能会在新版本中消失。如果无法自动移除这些引用，运行宏时可能会出现 `ERR_1000: The script entry point could not be found` 错误。此时需在 VBA 编辑器（Alt-F11）的 **Tools/References** 中手动取消勾选过时的引用。

**CATVBA 中布尔（Boolean）类型的支持**
VBA 调用 V5 方法时对 Boolean 类型存在已知限制。V5 应用中 Boolean 定义为 `unsigned char`（1 字节），而 VBA 定义为 `short`（2 字节）。
当 V5 方法返回 `True` 时，返回值为 **1**，而 VBA 期望的是 **-1**。因此，以下代码在 VBA 中可能失效：

```vb
If myObj.boolMethod() = True Then // 此测试将失败
...
End If

```

**纠正方法：**

```vb
If myObj.boolMethod() Then 
' 或者
If myObj.boolMethod() <> False Then

```

同理，不能直接对返回值使用 `Not` 运算符。需先存入变量：

```vb
Dim myBool
myBool = myObj.boolMethod()
Not(myBool) ' 此时将正确返回 False

```

*注意：此限制仅针对 VBA，不涉及 VBScript。*

**64 位环境下声明函数的限制**
在 64 位环境下，如果 `My.dll` 是 64 位 DLL，则无法进行如下声明：

```vb
Private Declare Function MyFunctionNameInDll Lib "E:\My.dll" (...) As ...

```

这是因为 V5 的 64 位 VBA 集成是基于一个专门的 32 位应用程序。因此，VBA 宏引用的 DLL 必须是 **32 位**的。