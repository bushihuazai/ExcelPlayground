# MAKEARRAY：规则生成数组

> `MAKEARRAY` 按行号和列号生成任意形状的二维数组。
> 很多场景下可以用 `SEQUENCE` 的行列广播替代，更简洁。

---

## 1. 基础语法

```
MAKEARRAY(rows, cols, LAMBDA(r, c, ...))
```

---

## 2. 经典示例

### 生成 5×5 乘法表

**MAKEARRAY 写法：**

```excel
=MAKEARRAY(5, 5, LAMBDA(r, c, r * c))
```

**SEQUENCE 简洁写法（推荐）：**

```excel
=SEQUENCE(5) * SEQUENCE(1, 5)
```

> 📌 SEQUENCE 行向量 × 列向量利用**广播**自动生成二维乘法表，更简洁。

---

### 生成棋盘格 0/1

**MAKEARRAY 写法：**

```excel
=MAKEARRAY(8, 8, LAMBDA(r, c, MOD(r + c, 2)))
```

**SEQUENCE 简洁写法：**

```excel
=MOD(SEQUENCE(8) + SEQUENCE(1, 8), 2)
```

---

### 生成单位矩阵（Identity Matrix）

```excel
=MAKEARRAY(5, 5, LAMBDA(r, c, IF(r = c, 1, 0)))
```

**SEQUENCE 简洁写法：**

```excel
=(SEQUENCE(5) = SEQUENCE(1, 5)) * 1
```

> 📌 `SEQUENCE(5) = SEQUENCE(1,5)` 产生 TRUE/FALSE 矩阵，乘以 1 转为 0/1。

---

### 生成上三角矩阵

```excel
=MAKEARRAY(5, 5, LAMBDA(r, c, IF(r <= c, 1, 0)))
```

**SEQUENCE 简洁写法：**

```excel
=(SEQUENCE(5) <= SEQUENCE(1, 5)) * 1
```

---

## 3. MAKEARRAY 的不可替代场景

当 LAMBDA 内的逻辑较复杂、无法用简单广播实现时，MAKEARRAY 仍然是最佳选择：

### 示例：生成杨辉三角（帕斯卡三角）

```excel
=MAKEARRAY(10, 10, LAMBDA(r, c,
  IF(c > r, "",
    REDUCE(1, SEQUENCE(MAX(1, c) - 1), LAMBDA(acc, i, acc * (r - i) / i))
  )
))
```

### 示例：距离矩阵

假设 A2:A6 为 X 坐标，B2:B6 为 Y 坐标：

```excel
=LET(
  x, A2:A6,
  y, B2:B6,
  n, ROWS(x),
  MAKEARRAY(n, n, LAMBDA(i, j,
    SQRT((INDEX(x, i) - INDEX(x, j))^2 + (INDEX(y, i) - INDEX(y, j))^2)
  ))
)
```

### 示例：生成随机稀疏矩阵（约 30% 非零）

```excel
=MAKEARRAY(10, 10, LAMBDA(r, c,
  IF(RANDARRAY(1, 1) < 0.3, RANDBETWEEN(1, 100), 0)
))
```

---

## 4. MAKEARRAY vs SEQUENCE 对照表

| 场景 | MAKEARRAY | SEQUENCE 广播 |
|------|-----------|---------------|
| 乘法表 | `MAKEARRAY(n,n,LAMBDA(r,c,r*c))` | `SEQUENCE(n)*SEQUENCE(1,n)` ✅ |
| 棋盘格 | `MAKEARRAY(n,n,LAMBDA(r,c,MOD(r+c,2)))` | `MOD(SEQUENCE(n)+SEQUENCE(1,n),2)` ✅ |
| 单位矩阵 | `MAKEARRAY(n,n,LAMBDA(r,c,(r=c)*1))` | `(SEQUENCE(n)=SEQUENCE(1,n))*1` ✅ |
| 距离矩阵 | ✅ 需要 INDEX 引用外部数据 | ❌ 无法直接实现 |
| 杨辉三角 | ✅ 逻辑复杂需逐元素控制 | ❌ 无法直接实现 |

> 📌 **经验法则**：如果 LAMBDA 体内只用 `r` 和 `c` 做简单算术，优先用 SEQUENCE 广播；如果需要引用外部数据或复杂逻辑，用 MAKEARRAY。

---

[← 返回目录](../README.md)
