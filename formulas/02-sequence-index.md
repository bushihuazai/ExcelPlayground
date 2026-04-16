# SEQUENCE 与 INDEX：动态数组的核心技巧

> `SEQUENCE` 是动态数组时代最强大的"生成器"——它能创建任意形状的数字序列。
> 搭配 `INDEX`，可以实现"按位置选取"的几乎一切需求。

---

## 1. SEQUENCE 基础

### 语法

```
SEQUENCE(rows, [cols], [start], [step])
```

| 参数 | 说明 | 默认值 |
|------|------|--------|
| rows | 行数 | 必填 |
| cols | 列数 | 1 |
| start | 起始值 | 1 |
| step | 步长 | 1 |

---

### 示例：生成 1~10 的列向量

```excel
=SEQUENCE(10)
```

### 示例：生成 1~10 的行向量

```excel
=SEQUENCE(1, 10)
```

### 示例：从 0 开始、步长为 2 的偶数序列

```excel
=SEQUENCE(10, 1, 0, 2)
```

> 结果：0, 2, 4, 6, 8, 10, 12, 14, 16, 18

### 示例：生成 5×5 矩阵（1~25）

```excel
=SEQUENCE(5, 5)
```

---

## 2. SEQUENCE 生成日期/时间

### 示例：生成连续 7 天日期

```excel
=SEQUENCE(7, 1, TODAY(), 1)
```

### 示例：生成本月每一天

```excel
=LET(
  m_start, DATE(YEAR(TODAY()), MONTH(TODAY()), 1),
  m_days, DAY(EOMONTH(TODAY(), 0)),
  SEQUENCE(m_days, 1, m_start, 1)
)
```

### 示例：生成每小时时间点（24 小时制）

```excel
=SEQUENCE(24, 1, 0, 1/24)
```

> 📌 单元格格式设为"时间"即可显示 0:00, 1:00, …, 23:00。

---

## 3. INDEX 动态选列/选行

### 语法

```
INDEX(array, row_num, [col_num])
```

- `row_num = 0` 或省略：返回**整列**
- `col_num = 0` 或省略：返回**整行**

### 示例：提取第 3 列

```excel
=INDEX(A2:E100,, 3)
```

### 示例：提取第 2 行

```excel
=INDEX(A2:E100, 2,)
```

### 示例：提取多列（第 1、3、5 列）

```excel
=INDEX(A2:E100, SEQUENCE(ROWS(A2:E100)), {1,3,5})
```

> 📌 `{1,3,5}` 是横向常量数组，配合 `SEQUENCE(n)` 的纵向行号，INDEX 自动展开为多列结果。

---

## 4. SEQUENCE + INDEX 组合技

### 4.1 取前 N 行（Top-N）

```excel
=LET(
  data, A2:E100,
  n, 10,
  INDEX(data, SEQUENCE(n), SEQUENCE(1, COLUMNS(data)))
)
```

> 📌 `SEQUENCE(n)` 生成行号 1~n，`SEQUENCE(1, COLUMNS(data))` 生成列号 1~cols，INDEX 交叉选取。

### 4.2 取最后 N 行

```excel
=LET(
  data, A2:E100,
  n, 5,
  total, ROWS(data),
  INDEX(data, SEQUENCE(n, 1, total - n + 1), SEQUENCE(1, COLUMNS(data)))
)
```

### 4.3 每隔 k 行取样

```excel
=LET(
  data, A2:E100,
  k, 3,
  total, ROWS(data),
  row_idx, SEQUENCE(INT(total / k), 1, 1, k),
  INDEX(data, row_idx, SEQUENCE(1, COLUMNS(data)))
)
```

### 4.4 逆序排列

```excel
=LET(
  data, A2:A20,
  n, ROWS(data),
  INDEX(data, SEQUENCE(n, 1, n, -1))
)
```

> 📌 `SEQUENCE(n, 1, n, -1)` 生成 n, n-1, …, 1。

### 4.5 生成乘法表（SEQUENCE 替代 MAKEARRAY）

```excel
=LET(
  n, 9,
  rows, SEQUENCE(n),
  cols, SEQUENCE(1, n),
  rows * cols
)
```

> 📌 利用 SEQUENCE 的行向量 × 列向量自动广播，比 MAKEARRAY 更简洁。

---

## 5. SEQUENCE 生成辅助索引

### 示例：对 UNIQUE 结果编号

```excel
=LET(
  u, UNIQUE(A2:A100),
  HSTACK(SEQUENCE(ROWS(u)), u)
)
```

### 示例：给 FILTER 结果附加原始行号

```excel
=LET(
  data, A2:E100,
  cond, INDEX(data,,3) > 1000,
  all_rows, SEQUENCE(ROWS(data)),
  matched_rows, FILTER(all_rows, cond),
  HSTACK(matched_rows, FILTER(data, cond))
)
```

> 📌 通过 `SEQUENCE` 先生成行号数组，再用同一条件 FILTER，即可保留原始行号。

---

## 6. SEQUENCE 生成重复模式

### 示例：循环重复 1,2,3,1,2,3,…

```excel
=LET(
  n, 12,
  cycle, 3,
  MOD(SEQUENCE(n) - 1, cycle) + 1
)
```

### 示例：生成分组编号（每 k 个一组）

```excel
=LET(
  n, 20,
  k, 4,
  INT((SEQUENCE(n) - 1) / k) + 1
)
```

> 结果：1,1,1,1, 2,2,2,2, 3,3,3,3, 4,4,4,4, 5,5,5,5

---

[← 返回目录](../README.md)
