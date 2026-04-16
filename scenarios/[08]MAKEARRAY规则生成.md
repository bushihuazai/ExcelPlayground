# 08 — MAKEARRAY 规则生成：按公式构造任意维度数组

> **适用版本**：Microsoft 365 / Excel 2021+

## 简介

`MAKEARRAY` 根据指定的行数、列数和生成规则，创建一个二维数组。LAMBDA 接收 `(行号, 列号)` 两个参数，返回该位置的值。

语法：`=MAKEARRAY(rows, cols, LAMBDA(r, c, 表达式))`

---

## 示例 1 ⭐ — 九九乘法表

```excel
=MAKEARRAY(9, 9, LAMBDA(r, c, IF(c<=r, r&"×"&c&"="&r*c, "")))
```

**说明**：当列号 ≤ 行号时输出乘法算式，否则为空。溢出后即可看到完整的三角形乘法表。

---

## 示例 2 ⭐⭐ — 棋盘格（0/1 交替矩阵）

```excel
=MAKEARRAY(8, 8, LAMBDA(r, c, MOD(r+c, 2)))
```

**说明**：`MOD(r+c, 2)` 使相邻单元格值交替为 0 和 1。配合条件格式即可呈现国际象棋棋盘效果。

---

## 示例 3 ⭐⭐⭐ — 单位矩阵（Identity Matrix）

```excel
=MAKEARRAY(5, 5, LAMBDA(r, c, IF(r=c, 1, 0)))
```

**说明**：对角线上为 1，其余为 0。在矩阵运算（如 `MMULT`）中，单位矩阵是基础构件。

---

## 示例 4 ⭐⭐⭐⭐ — 日历矩阵（生成某月日历）

```excel
=LET(
  year,   2025,
  month,  3,
  first,  DATE(year, month, 1),
  wday1,  WEEKDAY(first, 2),
  days,   DAY(EOMONTH(first, 0)),
  MAKEARRAY(6, 7, LAMBDA(r, c,
    LET(
      d, (r-1)*7 + c - wday1 + 1,
      IF(AND(d>=1, d<=days), d, "")
    )
  ))
)
```

**说明**：

1. 计算该月第一天是星期几（`wday1`，周一=1）
2. 计算该月总天数（`days`）
3. 6 行 7 列矩阵，每格根据位置推算日期编号
4. 超出范围的格子为空

> 💡 修改 `year` 和 `month` 即可动态切换月份。

---

## 示例 5 ⭐⭐⭐⭐⭐ — 帕斯卡三角形（杨辉三角）

```excel
=MAKEARRAY(10, 10, LAMBDA(r, c,
  IF(c>r, "",
    IF(OR(c=1, c=r), 1,
      COMBIN(r-1, c-1)
    )
  )
))
```

**说明**：

- 第 `r` 行第 `c` 列的值 = C(r−1, c−1)（组合数）
- 当 `c > r` 时留空，形成三角形
- `COMBIN` 是 Excel 内置组合数函数

> ⚠️ 行数不宜过大（建议 ≤20），否则数值会增长到很大。

---

## 应用注意点

| 要点 | 说明 |
|------|------|
| 行列号起始 | `r` 和 `c` 从 **1** 开始（不是 0） |
| 静态生成 | MAKEARRAY 生成后为静态溢出数组，修改公式参数后自动刷新 |
| 数据填充 | 常用于生成测试数据、日历、矩阵、查表等 |
| 性能 | 行列数的乘积即 LAMBDA 调用次数，100×100 = 10,000 次，注意规模 |
| 与 SEQUENCE 对比 | `SEQUENCE` 只能生成等差数列；`MAKEARRAY` 可实现任意规则 |
