# MAP / BYROW / BYCOL：映射函数

> 这三个函数是 Excel 的"高阶函数三兄弟"，让你对数组的每个元素、每一行、每一列执行自定义操作。

---

## 1. MAP：逐元素映射

### 单数组映射

**示例：批量转大写并去除前后空格**

```excel
=MAP(A2:A10, LAMBDA(x, UPPER(TRIM(x))))
```

**示例：将空值替换为默认文本**

```excel
=MAP(A2:A20, LAMBDA(x, IF(ISBLANK(x), "（未填写）", x)))
```

### 多数组映射

**示例：两列对应相乘后保留两位小数**

```excel
=MAP(A2:A10, B2:B10, LAMBDA(price, qty, ROUND(price * qty, 2)))
```

**示例：三列拼接地址**

```excel
=MAP(A2:A10, B2:B10, C2:C10, LAMBDA(province, city, district,
  province & city & district
))
```

---

## 2. BYROW：按行计算

`BYROW` 对二维区域的**每一行**执行 LAMBDA，返回一列结果。

**示例：按行求和**

```excel
=BYROW(B2:F10, LAMBDA(r, SUM(r)))
```

**示例：每行非空单元格计数**

```excel
=BYROW(A2:F10, LAMBDA(r, COUNTA(r)))
```

**示例：每行最大值与最小值之差（极差）**

```excel
=BYROW(B2:F10, LAMBDA(r, MAX(r) - MIN(r)))
```

### 进阶：BYROW + INDEX + SEQUENCE 提取每行指定列

```excel
=LET(
  data, A2:F100,
  col_idx, {1,3,5},
  BYROW(
    SEQUENCE(ROWS(data)),
    LAMBDA(i, TEXTJOIN(",", TRUE, INDEX(data, i, col_idx)))
  )
)
```

> 📌 利用 `SEQUENCE` 生成行号配合 INDEX 精准定位。

---

## 3. BYCOL：按列计算

`BYCOL` 对二维区域的**每一列**执行 LAMBDA，返回一行结果。

**示例：按列取最大值**

```excel
=BYCOL(B2:F10, LAMBDA(c, MAX(c)))
```

**示例：每列的非零值平均**

```excel
=BYCOL(B2:F10, LAMBDA(c,
  LET(nz, FILTER(c, c <> 0, 0), IF(ROWS(nz) = 1, 0, AVERAGE(nz)))
))
```

**示例：按列计算变异系数（CV）**

```excel
=BYCOL(B2:F10, LAMBDA(c,
  LET(
    avg, AVERAGE(c),
    sd, STDEV(c),
    IF(avg = 0, "", sd / avg)
  )
))
```

---

## 4. MAP vs BYROW vs BYCOL 对比

| 函数 | 输入粒度 | 返回形状 | 典型场景 |
|------|----------|----------|----------|
| MAP | 单个元素 | 与输入同形 | 逐元素转换、多数组对应计算 |
| BYROW | 一行 | 单列（每行一个结果）| 行级汇总、行级判断 |
| BYCOL | 一列 | 单行（每列一个结果）| 列级统计、列级指标 |

---

## 5. 实用组合

### MAP + SEQUENCE：给结果编号

```excel
=LET(
  data, UNIQUE(A2:A50),
  n, ROWS(data),
  idx, SEQUENCE(n),
  MAP(idx, data, LAMBDA(i, v, i & ". " & v))
)
```

### BYROW + SEQUENCE：构建行级汇总表

```excel
=LET(
  data, B2:M10,
  headers, B1:M1,
  BYROW(data, LAMBDA(r,
    LET(
      max_val, MAX(r),
      max_idx, XMATCH(max_val, r),
      INDEX(headers,, max_idx) & ":" & max_val
    )
  ))
)
```

> 📌 为每一行找到最大值所在的列名及其值。

---

[← 返回目录](../README.md)
