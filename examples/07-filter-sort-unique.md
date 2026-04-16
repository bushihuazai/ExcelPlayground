# FILTER / SORT / SORTBY / UNIQUE / XLOOKUP / XMATCH：声明式数据清洗

> 这些函数构成了 Excel 中"声明式数据处理"的核心——你描述**你想要什么**，而非**怎么做**。

---

## 1. FILTER：条件筛选

### 基础：单条件

```excel
=FILTER(A2:E100, B2:B100 = "华东", "无结果")
```

### 多条件（AND — 同时满足）

```excel
=FILTER(A2:E100, (B2:B100 = "华东") * (E2:E100 > 10000), "无结果")
```

### 多条件（OR — 满足任一）

```excel
=FILTER(A2:E100, (B2:B100 = "华东") + (B2:B100 = "华南"), "无结果")
```

### 使用 INDEX 简化多列引用

```excel
=LET(
  data, A2:E100,
  region, INDEX(data,, 2),
  amount, INDEX(data,, 5),
  FILTER(data, (region = "华东") * (amount > 10000), "无结果")
)
```

> 📌 `INDEX(data,, 2)` 提取整列，避免独立写 `B2:B100`，当数据列变化时只需改 data 引用。

---

## 2. FILTER + SEQUENCE：保留原始行号

```excel
=LET(
  data, A2:E100,
  cond, INDEX(data,, 5) > 10000,
  row_idx, SEQUENCE(ROWS(data)),
  HSTACK(FILTER(row_idx, cond), FILTER(data, cond))
)
```

> 📌 在 FILTER 之前用 SEQUENCE 创建行号列，同条件筛选后拼接。

---

## 3. SORT / SORTBY：排序

### SORT：按指定列排序

```excel
=SORT(A2:E100, 5, -1)
```

> 按第 5 列降序排列。

### SORTBY：按外部数组排序

```excel
=SORTBY(A2:E100, E2:E100, -1, B2:B100, 1)
```

> 先按金额（E 列）降序，再按区域（B 列）升序。

### SORT + INDEX：只取排序后的部分列

```excel
=LET(
  sorted, SORT(A2:E100, 5, -1),
  CHOOSECOLS(sorted, 1, 2, 5)
)
```

> 📌 使用 CHOOSECOLS 从排序结果中选取需要的列。

---

## 4. UNIQUE：去重

### 基础去重

```excel
=UNIQUE(C2:C100)
```

### 去重后排序

```excel
=SORT(UNIQUE(C2:C100))
```

### 多列去重（按多列组合判断唯一性）

```excel
=UNIQUE(B2:C100)
```

### 找出只出现一次的值（`exactly_once` 参数）

```excel
=UNIQUE(C2:C100,, TRUE)
```

---

## 5. XLOOKUP / XMATCH：增强查找

### XLOOKUP 基础

```excel
=XLOOKUP(E2, A2:A100, B2:B100, "未找到")
```

### XLOOKUP 返回多列

```excel
=XLOOKUP(E2, A2:A100, B2:D100, "未找到")
```

### XLOOKUP 逆向查找（从下往上）

```excel
=XLOOKUP(E2, A2:A100, B2:B100, "未找到", 0, -1)
```

### XMATCH：返回位置

```excel
=XMATCH("目标值", A2:A100, 0)
```

### XLOOKUP + INDEX + SEQUENCE：批量查找

对一组查找值批量查询：

```excel
=LET(
  lookup_vals, E2:E10,
  source_key, A2:A100,
  source_data, B2:D100,
  MAP(lookup_vals, LAMBDA(v, TEXTJOIN(", ", TRUE, XLOOKUP(v, source_key, source_data, ""))))
)
```

---

## 6. 组合实战

### 去重 + 计数 + 排序 = 频率表

```excel
=LET(
  col, C2:C100,
  u, UNIQUE(col),
  cnt, MAP(u, LAMBDA(v, SUM((col = v) * 1))),
  result, HSTACK(u, cnt),
  SORT(result, 2, -1)
)
```

> 📌 纯公式生成按频率降序排列的频率表。

### FILTER + SORT + TAKE = Top-N

```excel
=LET(
  data, A2:E100,
  region, INDEX(data,, 2),
  amount, INDEX(data,, 5),
  filtered, FILTER(data, region = "华东"),
  sorted, SORT(filtered, 5, -1),
  TAKE(sorted, 10)
)
```

> 📌 筛选 → 排序 → 取前 10，三步流水线。

### 每组 Top-1（分组取最大值行）

```excel
=LET(
  data, A2:E100,
  group_col, INDEX(data,, 2),
  val_col, INDEX(data,, 5),
  groups, UNIQUE(group_col),
  MAP(SEQUENCE(ROWS(groups)), LAMBDA(i,
    LET(
      g, INDEX(groups, i),
      subset, FILTER(data, group_col = g),
      best, SORT(subset, 5, -1),
      TEXTJOIN(" | ", TRUE, INDEX(best, 1, SEQUENCE(1, COLUMNS(data))))
    )
  ))
)
```

> 📌 利用 SEQUENCE 生成索引遍历每个分组，分别 FILTER → SORT → 取首行。

---

[← 返回目录](../README.md)
