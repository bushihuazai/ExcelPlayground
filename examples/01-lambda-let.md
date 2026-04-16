# LAMBDA + LET：封装可读函数

> LAMBDA 让 Excel 拥有了"定义函数"的能力；LET 让公式拥有了"局部变量"。
> 二者配合是所有高级公式的基石。

---

## 基础：LET 消除重复计算

**示例：多次引用同一中间结果**

```excel
=LET(
  avg, AVERAGE(A2:A100),
  stdev, STDEV(A2:A100),
  (A2 - avg) / stdev
)
```

> 📌 不使用 LET 时 `AVERAGE(A2:A100)` 和 `STDEV(A2:A100)` 可能要写多次；LET 只计算一次并命名。

---

## 基础：LAMBDA 定义可复用函数

**示例：计算折后价（含最小为 0 保护）**

```excel
=LAMBDA(price, discount,
  LET(result, price * (1 - discount), MAX(0, result))
)(A2, B2)
```

**示例：安全除法 `SAFE_DIV`（在名称管理器中定义）**

```excel
=LAMBDA(a, b, IF(b = 0, "", a / b))
```

使用时直接 `=SAFE_DIV(A2, B2)`。

---

## 进阶：LET + SEQUENCE 替代辅助列

传统做法需要辅助列来生成行号，现在可以用 `SEQUENCE` 内联：

**示例：为数据添加序号列**

```excel
=LET(
  data, A2:C10,
  n, ROWS(data),
  seq, SEQUENCE(n),
  HSTACK(seq, data)
)
```

> 📌 `SEQUENCE(n)` 生成 1..n 的列向量，无需占用额外列。

---

## 进阶：LAMBDA 递归（自引用）

通过在名称管理器中定义递归 LAMBDA，可以实现循环逻辑：

**示例：`FACTORIAL` — 阶乘**

```excel
=LAMBDA(n, IF(n <= 1, 1, n * FACTORIAL(n - 1)))
```

**示例：`GCD_FUNC` — 最大公约数（辗转相除）**

```excel
=LAMBDA(a, b, IF(b = 0, a, GCD_FUNC(b, MOD(a, b))))
```

---

## 技巧：LET 多层嵌套构建流水线

```excel
=LET(
  raw, A2:E200,
  cleaned, FILTER(raw, INDEX(raw,,1) <> ""),
  sorted, SORT(cleaned, 3, -1),
  top10, INDEX(sorted, SEQUENCE(MIN(10, ROWS(sorted))), SEQUENCE(1, COLUMNS(sorted))),
  top10
)
```

> 📌 每一步结果都有名字，调试时可以把最后一行改为任意中间变量名来检查。

---

[← 返回目录](../README.md)
