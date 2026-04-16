# REDUCE / SCAN：折叠与累计

> `REDUCE` 把数组折叠成一个结果；`SCAN` 保留每一步中间值。
> 这是 Excel 中最接近"循环"的纯公式实现。

---

## 1. REDUCE：把数组折叠成单个值

### 语法

```
REDUCE(initial_value, array, LAMBDA(acc, x, ...))
```

### 示例：累加求和

```excel
=REDUCE(0, A2:A10, LAMBDA(acc, x, acc + x))
```

### 示例：累计连乘（几何增长）

```excel
=REDUCE(1, A2:A10, LAMBDA(acc, x, acc * x))
```

### 示例：拼接字符串（TEXTJOIN 替代方案）

```excel
=REDUCE("", A2:A10, LAMBDA(acc, x, IF(acc = "", x, acc & "、" & x)))
```

### 示例：计数满足条件的元素

```excel
=REDUCE(0, A2:A100, LAMBDA(acc, x, acc + (x > 60)))
```

> 📌 `(x > 60)` 返回 TRUE(1) 或 FALSE(0)，累加即计数。

---

## 2. REDUCE 进阶

### 示例：找到最大值（手动实现 MAX）

```excel
=REDUCE(-1E+308, A2:A100, LAMBDA(acc, x, IF(x > acc, x, acc)))
```

### 示例：条件累加（只累加正数）

```excel
=REDUCE(0, A2:A100, LAMBDA(acc, x, acc + IF(x > 0, x, 0)))
```

### 示例：REDUCE + VSTACK 动态构建表

逐行判断是否保留，构建结果集：

```excel
=REDUCE(
  "",
  SEQUENCE(ROWS(A2:E100)),
  LAMBDA(acc, i,
    LET(
      row, INDEX(A2:E100, i, SEQUENCE(1, 5)),
      cond, INDEX(row,, 3) > 1000,
      IF(cond, IF(acc = "", row, VSTACK(acc, row)), acc)
    )
  )
)
```

> 📌 这相当于一个带条件的"逐行累积器"，是纯公式模拟 for-loop + if 的经典模式。

---

## 3. SCAN：输出每一步中间值

### 语法

```
SCAN(initial_value, array, LAMBDA(acc, x, ...))
```

> 与 REDUCE 的唯一区别：SCAN 返回每一步的 acc，而非只返回最终值。

### 示例：运行累计和（Running Total）

```excel
=SCAN(0, A2:A10, LAMBDA(acc, x, acc + x))
```

### 示例：账户余额轨迹

```excel
=SCAN(1000, B2:B20, LAMBDA(balance, cashflow, balance + cashflow))
```

### 示例：运行累积最大值

```excel
=SCAN(-1E+308, A2:A50, LAMBDA(acc, x, MAX(acc, x)))
```

### 示例：运行累积计数（满足条件）

```excel
=SCAN(0, A2:A100, LAMBDA(acc, x, acc + (x > 0)))
```

---

## 4. SCAN 进阶

### 示例：运行移动平均（滑动窗口 N=5）

利用 `SCAN` + `SEQUENCE` + `INDEX` 实现：

```excel
=LET(
  data, A2:A100,
  n, ROWS(data),
  window, 5,
  idx, SEQUENCE(n),
  MAP(idx, LAMBDA(i,
    LET(
      start_row, MAX(1, i - window + 1),
      length, i - start_row + 1,
      AVERAGE(INDEX(data, SEQUENCE(length, 1, start_row)))
    )
  ))
)
```

> 📌 虽然这里用了 MAP 而非 SCAN，但核心思路是 SEQUENCE 生成滑动窗口索引。

### 示例：带上限的余额追踪

```excel
=SCAN(0, A2:A20, LAMBDA(acc, x, MIN(10000, MAX(0, acc + x))))
```

> 📌 余额不低于 0、不超过 10000，模拟有上下限约束的累积。

---

## 5. REDUCE vs SCAN 对比

| 特性 | REDUCE | SCAN |
|------|--------|------|
| 返回值 | 单个最终值 | 每一步的数组 |
| 适用场景 | 汇总/折叠 | 过程可视化、运行累计 |
| 与 SUM 对比 | 自定义版 SUM | 自定义版 Running Total |

---

## 6. 实用组合

### REDUCE + SEQUENCE：自定义阶乘

```excel
=REDUCE(1, SEQUENCE(10), LAMBDA(acc, x, acc * x))
```

> 📌 `SEQUENCE(10)` 生成 1~10，REDUCE 逐个相乘 = 10!

### SCAN + HSTACK：余额轨迹表

```excel
=LET(
  dates, A2:A20,
  flows, B2:B20,
  balance, SCAN(1000, flows, LAMBDA(acc, x, acc + x)),
  HSTACK(dates, flows, balance)
)
```

> 📌 一个公式输出日期、现金流、余额三列完整表格。

---

[← 返回目录](../README.md)
