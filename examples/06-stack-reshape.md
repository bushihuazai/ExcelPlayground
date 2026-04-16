# VSTACK / HSTACK / TOCOL / TOROW / WRAPCOLS / WRAPROWS / CHOOSECOLS / CHOOSEROWS / DROP / TAKE：拼接与重塑

> 动态数组时代的"表操作"函数族，让你在纯公式中实现拼接、选取、截取和重塑。

---

## 1. VSTACK / HSTACK：纵向/横向拼接

### 上下合并两块数据

```excel
=VSTACK(A2:C10, A15:C20)
```

### 左右合并

```excel
=HSTACK(A2:C10, E2:G10)
```

### 为不同来源补齐列后再纵向汇总

```excel
=LET(
  t1, A2:C10,
  t2, H2:I8,
  cols_diff, COLUMNS(t1) - COLUMNS(t2),
  filler, SEQUENCE(ROWS(t2), cols_diff, 0, 0),
  VSTACK(t1, HSTACK(t2, filler))
)
```

> 📌 使用 `SEQUENCE(rows, cols, 0, 0)` 生成全 0 填充矩阵，比 MAKEARRAY + LAMBDA 更简洁。

### HSTACK + SEQUENCE 添加序号列

```excel
=LET(
  data, A2:E100,
  HSTACK(SEQUENCE(ROWS(data)), data)
)
```

---

## 2. CHOOSECOLS / CHOOSEROWS：按位置选取

### 选取第 1、3、5 列

```excel
=CHOOSECOLS(A2:F100, 1, 3, 5)
```

### 使用 SEQUENCE 动态选取前 N 列

```excel
=CHOOSECOLS(A2:Z100, SEQUENCE(1, 5))
```

> 📌 `SEQUENCE(1, 5)` = `{1,2,3,4,5}`，动态替代手写列号。

### 选取第 1、5、10 行

```excel
=CHOOSEROWS(A2:F100, 1, 5, 10)
```

### 使用 SEQUENCE 选取每隔 k 行

```excel
=LET(
  data, A2:F100,
  k, 3,
  n, INT(ROWS(data) / k),
  CHOOSEROWS(data, SEQUENCE(n, 1, 1, k))
)
```

---

## 3. DROP / TAKE：截取数组的头部/尾部

### 取前 5 行

```excel
=TAKE(A2:F100, 5)
```

### 取最后 5 行

```excel
=TAKE(A2:F100, -5)
```

### 去掉标题行（第 1 行）

```excel
=DROP(A1:F100, 1)
```

### 去掉最后 2 行

```excel
=DROP(A2:F100, -2)
```

### 取前 3 列

```excel
=TAKE(A2:F100,, 3)
```

### 去掉最后 1 列

```excel
=DROP(A2:F100,, -1)
```

### 组合：取前 10 行、前 3 列

```excel
=TAKE(A2:F100, 10, 3)
```

---

## 4. TOCOL / TOROW：展平数组

### 将二维区域展平为一列

```excel
=TOCOL(A2:E10)
```

### 将二维区域展平为一行

```excel
=TOROW(A2:E10)
```

### 展平并去除空值和错误

```excel
=TOCOL(A2:E10, 3)
```

> 📌 第二参数：1 = 忽略空值，2 = 忽略错误，3 = 忽略空值和错误。

---

## 5. WRAPCOLS / WRAPROWS：重塑一维数组

### 将 1~12 排列为 3 行 4 列

```excel
=WRAPROWS(SEQUENCE(12), 4)
```

> 结果：
> | 1 | 2 | 3 | 4 |
> | 5 | 6 | 7 | 8 |
> | 9 | 10 | 11 | 12 |

### 将 1~12 排列为 4 行 3 列（按列填充）

```excel
=WRAPCOLS(SEQUENCE(12), 4)
```

> 结果：
> | 1 | 5 | 9 |
> | 2 | 6 | 10 |
> | 3 | 7 | 11 |
> | 4 | 8 | 12 |

### 实用：将长列表分成 N 列展示

```excel
=LET(
  data, TOCOL(A2:A50, 1),
  cols, 5,
  WRAPROWS(data, cols, "")
)
```

> 📌 第三参数 `""` 指定当数据不够整除时用空字符串填充。

---

## 6. 组合实战

### 转置表（行列互换）

```excel
=LET(
  data, A2:E10,
  n_rows, ROWS(data),
  n_cols, COLUMNS(data),
  flat, TOCOL(data),
  WRAPCOLS(flat, n_rows)
)
```

> 📌 先展平再按列包裹 = 转置（等效于 TRANSPOSE，但更灵活）。

### 交错合并两列

将 A 列和 B 列交错合并为一列（A1,B1,A2,B2,…）：

```excel
=LET(
  a, A2:A10,
  b, B2:B10,
  n, ROWS(a),
  idx, SEQUENCE(n * 2),
  row_num, INT((idx - 1) / 2) + 1,
  col_flag, MOD(idx - 1, 2) + 1,
  MAP(row_num, col_flag, LAMBDA(r, c, IF(c = 1, INDEX(a, r), INDEX(b, r))))
)
```

> 📌 用 SEQUENCE 生成交替索引，MAP 按索引拼合。

---

[← 返回目录](../README.md)
