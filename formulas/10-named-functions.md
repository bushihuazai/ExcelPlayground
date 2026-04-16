# 命名函数库：在名称管理器中沉淀可复用公式

> 在 **公式 → 名称管理器** 中定义 LAMBDA 函数，团队共享，无需 VBA。

---

## 使用方法

1. 打开 **公式 → 名称管理器 → 新建**
2. 名称：填函数名（如 `SAFE_DIV`）
3. 引用位置：粘贴 LAMBDA 公式
4. 确定后即可在任意单元格中像内置函数一样使用

---

## 推荐函数库

### 1. SAFE_DIV — 安全除法

```excel
=LAMBDA(a, b, IF(b = 0, "", a / b))
```

使用：`=SAFE_DIV(A2, B2)`

---

### 2. RUNNING_TOTAL — 运行累计

```excel
=LAMBDA(arr, SCAN(0, arr, LAMBDA(acc, x, acc + x)))
```

使用：`=RUNNING_TOTAL(B2:B100)`

---

### 3. GROUP_SUM — 分组求和

```excel
=LAMBDA(keys, vals,
  LET(
    u, UNIQUE(keys),
    totals, MAP(u, LAMBDA(k, SUM(FILTER(vals, keys = k)))),
    SORT(HSTACK(u, totals), 2, -1)
  )
)
```

使用：`=GROUP_SUM(B2:B100, E2:E100)`

---

### 4. GROUP_COUNT — 分组计数

```excel
=LAMBDA(keys,
  LET(
    u, UNIQUE(keys),
    counts, MAP(u, LAMBDA(k, SUM((keys = k) * 1))),
    SORT(HSTACK(u, counts), 2, -1)
  )
)
```

使用：`=GROUP_COUNT(C2:C100)`

---

### 5. TOP_N — 取前 N 行

```excel
=LAMBDA(data, n, sort_col, [order],
  LET(
    _order, IF(ISOMITTED(order), -1, order),
    sorted, SORT(data, sort_col, _order),
    TAKE(sorted, MIN(n, ROWS(sorted)))
  )
)
```

使用：`=TOP_N(A2:E100, 10, 5)` 或 `=TOP_N(A2:E100, 10, 5, 1)` 升序

---

### 6. REVERSE — 逆序数组

```excel
=LAMBDA(arr,
  LET(
    n, ROWS(arr),
    INDEX(arr, SEQUENCE(n, 1, n, -1), SEQUENCE(1, COLUMNS(arr)))
  )
)
```

使用：`=REVERSE(A2:C20)`

> 📌 `SEQUENCE(n, 1, n, -1)` 生成逆序行号，INDEX 按此选取。

---

### 7. PAGINATE — 分页取数

```excel
=LAMBDA(data, page, page_size,
  LET(
    n, ROWS(data),
    start_row, (page - 1) * page_size + 1,
    actual_size, MIN(page_size, n - start_row + 1),
    IF(start_row > n, "无数据",
      INDEX(data, SEQUENCE(actual_size, 1, start_row), SEQUENCE(1, COLUMNS(data)))
    )
  )
)
```

使用：`=PAGINATE(A2:E1000, 3, 20)` — 第 3 页，每页 20 条

> 📌 核心是 `SEQUENCE(size, 1, start_row)` 动态生成行号切片。

---

### 8. MOVING_AVG — 移动平均

```excel
=LAMBDA(arr, window,
  LET(
    n, ROWS(arr),
    idx, SEQUENCE(n),
    MAP(idx, LAMBDA(i,
      LET(
        start_row, MAX(1, i - window + 1),
        length, i - start_row + 1,
        AVERAGE(INDEX(arr, SEQUENCE(length, 1, start_row)))
      )
    ))
  )
)
```

使用：`=MOVING_AVG(B2:B100, 7)` — 7 期移动平均

---

### 9. UNPIVOT — 逆透视（宽表转长表）

```excel
=LAMBDA(row_labels, col_headers, values,
  LET(
    n_r, ROWS(values),
    n_c, COLUMNS(values),
    total, n_r * n_c,
    idx, SEQUENCE(total),
    r, INT((idx - 1) / n_c) + 1,
    c, MOD(idx - 1, n_c) + 1,
    label, INDEX(row_labels, r),
    header, INDEX(col_headers,, c),
    val, INDEX(values, r, c),
    HSTACK(label, header, val)
  )
)
```

使用：`=UNPIVOT(A2:A10, B1:F1, B2:F10)`

> 📌 `SEQUENCE(total)` + 整除/取余 = 行列双索引展开，经典 SEQUENCE 技巧。

---

### 10. PERCENTILE_RANK — 百分位排名

```excel
=LAMBDA(arr,
  LET(
    n, ROWS(arr),
    MAP(arr, LAMBDA(x, SUM((arr <= x) * 1) / n))
  )
)
```

使用：`=PERCENTILE_RANK(C2:C100)`

---

## 函数速查表

| 函数名 | 用途 | 核心技巧 |
|--------|------|----------|
| SAFE_DIV | 安全除法 | IF 防零 |
| RUNNING_TOTAL | 运行累计 | SCAN |
| GROUP_SUM | 分组求和 | UNIQUE + MAP + FILTER |
| GROUP_COUNT | 分组计数 | UNIQUE + MAP |
| TOP_N | 前 N 行 | SORT + TAKE |
| REVERSE | 逆序 | SEQUENCE(n,1,n,-1) + INDEX |
| PAGINATE | 分页 | SEQUENCE(size,1,start) + INDEX |
| MOVING_AVG | 移动平均 | SEQUENCE 滑动窗口 + INDEX |
| UNPIVOT | 逆透视 | SEQUENCE + 整除/取余索引 |
| PERCENTILE_RANK | 百分位排名 | MAP + 向量化比较 |

---

[← 返回目录](../README.md)
