## 基础设施

### 关于 CATSafeArrayVariant

许多返回数值的方法并不只返回单个值，而是返回一个值数组。例如，`Viewpoint3D` 对象的 `GetOrigin` 属性以包含三个坐标的数组形式返回一个 3D 点。该数组作为输出参数返回。

这些参数在 VB/VBA 对象浏览器中显示为 **Variant** 类型，但在参考文档中被列为 **CATSafeArrayVariant**。

您可以按照以下方式在 `MyVPOrigin` 变量中获取当前窗口中当前视图的 3D 视点原点：

```vb
ReDim MyVPOrigin(2)
CATIA.ActiveWindow.ActiveViewer.Viewpoint3D.GetOrigin MyVPOrigin

```

您应该直接使用 `ReDim` 语句为数组分配大小（同时将其声明为可变大小数组），然后调用 `GetOrigin` 方法。**分配大小为 2 意味着分配了一个包含三个值的数组。**

要访问每个坐标，您需要深入了解：

* **x** 坐标位于 `MyVPOrigin(0)`
* **y** 坐标位于 `MyVPOrigin(1)`
* **z** 坐标位于 `MyVPOrigin(2)`

与集合（Collections）不同，**CATSafeArrayVariant 的索引是从 0 开始的**。若要设置一组新的三元组数值，可以这样写：

```vb
MyVPOrigin = Array(150, 200, 50)
CATIA.ActiveWindow.ActiveViewer.Viewpoint3D.PutOrigin MyVPOrigin

```

或者：

```vb
CATIA.ActiveWindow.ActiveViewer.Viewpoint3D.PutOrigin Array(150, 200, 50)

```

但如果仅需修改 **y** 坐标，请编写：

```vb
MyVPOrigin(1) = 200
CATIA.ActiveWindow.ActiveViewer.Viewpoint3D.PutOrigin MyVPOrigin

```

当您不知道数组的大小时，请使用 `UBound` 函数。它会返回数组中最高元素的秩（索引）。例如，以下示例将在 `Highestrank` 中返回 2：

```vb
Highestrank = Ubound(MyVPOrigin)

```