## 基础设施

### 关于子程序 (Subs) 和函数 (Functions)

在 Visual Basic 中，对象所公开的方法被称为：

* **Sub (子程序)**：如果不返回任何值。
* **Function (函数)**：如果返回一个值。

当方法需要参数时请务必小心。在 Visual Basic Script 中，向 **Sub** 传递参数时**不要使用括号**，示例如下：

```vb
Object.Sub arg1, arg2, arg3

```

但在调用 **Function** 时必须**使用括号**：

```vb
Dim ReturnedObject As AnyObject
Set ReturnedObject = Object.Function (arg1, arg2, arg3)

```

仅当返回值为**对象**时才必须使用 `Set` 关键字，如果返回值是字符串或数字则不需要。然而，被定义为 CATIA 字面量（Literals）的字符串和数字也是对象，如果函数返回的是字面量对象，则必须使用 `Set`。

最后，如果你将返回值存储在一个**属性 (Property)** 中，则不需要使用 `Set`：

```vb
myObject.aggregatedObject = Object.Function (arg1, arg2, arg3)

```

这是因为实际上并不存在 `aggregatedObject` 变量，属性只是**访问器方法 (Accessor Methods)** 的语法简写（在此处对应 `get_aggregatedObject` 和 `set_aggregatedObject`），允许将这些方法呈现为对象的属性。因此，上述语法等同于：

```vb
myObject.set_aggregatedObject( Object.Function (arg1, arg2, arg3) )

```

在这种情况下不需要使用 `Set`。