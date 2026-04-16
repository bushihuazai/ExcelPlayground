# 06 — REDUCE 累计折叠：把数组压缩成单个结果

> **适用版本**：Microsoft 365 / Excel 2021+

## 简介

`REDUCE` 类似编程语言中的 `fold` / `reduce` / `aggregate`。它从一个初始值出发，依次与数组中每个元素执行 LAMBDA 运算，最终返回**一个结果**。

语法：`=REDUCE(初始值, array, LAMBDA(累加器, 当前元素, 表达式))`

---

## 示例 1 ⭐ — 累加求和

```excel
=REDUCE(0, A2:A20, LAMBDA(acc, x, acc + x))
```

**说明**：`acc` 从 0 出发，依次加上 A2、A3 … A20，最终返回总和。效果等同于 `SUM(A2:A20)`，但展示了 REDUCE 的基本模式。

---

## 示例 2 ⭐⭐ — 累计连乘（复合收益率）

```excel
=REDUCE(1, B2:B13, LAMBDA(acc, x, acc * (1 + x)))
```

| 参数 | 说明 |
|------|------|
| `1` | 初始本金系数（代表 100%） |
| `B2:B13` | 每月收益率（如 0.05 表示 5%） |

**说明**：逐月将上期累计乘以 `(1+当月收益率)`，最终得到复合增长倍数。减去 1 即为年化收益率。

---

## 示例 3 ⭐⭐⭐ — 文本拼接（带分隔符）

```excel
=REDUCE("", A2:A30, LAMBDA(acc, x,
  IF(acc="", TEXT(x,"@"), acc & "、" & TEXT(x,"@"))
))
```

**说明**：第一个元素直接作为起始文本，后续元素用 "、" 分隔拼接。`TEXT(x,"@")` 确保数值也能正确转换为文本。

> 💡 简单场景推荐直接用 `TEXTJOIN`，REDUCE 更适合需要**自定义逻辑**的拼接。

---

## 示例 4 ⭐⭐⭐⭐ — 求最大连续正数个数

```excel
=REDUCE(
  HSTACK(0, 0),
  A2:A100,
  LAMBDA(state, x,
    LET(
      cur,    IF(x>0, INDEX(state,1,1)+1, 0),
      best,   MAX(INDEX(state,1,2), cur),
      HSTACK(cur, best)
    )
  )
)
```

取结果的第 2 列：

```excel
=INDEX(上述公式, 1, 2)
```

**说明**：

- `state` 是一个 1×2 数组：`{当前连续计数, 历史最大}`
- 每遇到正数，计数 +1；遇到非正数，计数归零
- 每步更新历史最大值

这是 REDUCE 的**状态机**用法——用数组作为累加器，维护多个状态字段。

---

## 示例 5 ⭐⭐⭐⭐⭐ — 去除连续重复项

```excel
=LET(
  arr, A2:A50,
  result, REDUCE("", arr, LAMBDA(acc, x,
    IF(acc="", x,
      IF(INDEX(TEXTSPLIT(acc,"||"), 1, LEN(acc)-LEN(SUBSTITUTE(acc,"||",""))+1) = TEXT(x,"@"),
        acc,
        acc & "||" & TEXT(x,"@")
      )
    )
  )),
  TEXTSPLIT(result, "||")
)
```

> **更简洁的写法**（利用辅助列判断前项）：

```excel
=LET(
  arr, A2:A50,
  n,   ROWS(arr),
  idx, FILTER(SEQUENCE(n), (SEQUENCE(n)=1) + (INDEX(arr,SEQUENCE(n),1)<>INDEX(arr,SEQUENCE(n)-1,1))),
  INDEX(arr, idx)
)
```

**说明**：连续重复去除（Run-Length Encoding 的前半段）在数据清洗中常见。REDUCE 方案通过拼接/拆分实现；第二种方案利用 FILTER + 位移比较更高效。

---

## 应用注意点

| 要点 | 说明 |
|------|------|
| 初始值类型 | 初始值的类型应与最终结果一致（数值用 `0`、文本用 `""`、数组用 `HSTACK(...)` 等） |
| 累加器可以是数组 | REDUCE 的"状态"不必是标量——可以用数组维护多个中间值 |
| 与 SCAN 对比 | REDUCE 只返回**最终结果**；SCAN 返回**每一步的中间结果** |
| 性能 | REDUCE 逐元素调用 LAMBDA，对 10,000+ 元素可能较慢 |
| 调试 | 用 SCAN 替换 REDUCE 可以查看中间状态，方便排错 |
