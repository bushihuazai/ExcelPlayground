# TEXTSPLIT / TEXTJOIN / TEXTBEFORE / TEXTAFTER：文本结构化处理

> 文本处理是 Excel 中最常见的需求之一。新函数让"拆/合/截/替"变得极其简洁。

---

## 1. TEXTSPLIT：拆分文本

### 基础：按分隔符拆为多列

```excel
=TEXTSPLIT(A2, "-")
```

> "北京-朝阳-望京" → 三列：北京 | 朝阳 | 望京

### 按行分隔符拆为多行

```excel
=TEXTSPLIT(A2,, CHAR(10))
```

> 📌 `CHAR(10)` 是换行符，将单元格内多行文本拆成多行。

### 同时按列和行拆分

```excel
=TEXTSPLIT(A2, ",", ";")
```

> "a,b;c,d" → 2×2 矩阵：
> | a | b |
> | c | d |

---

## 2. TEXTJOIN：合并文本

### 基础合并

```excel
=TEXTJOIN(",", TRUE, A2:A10)
```

> 📌 第二参数 TRUE 表示忽略空单元格。

### 合并筛选后的结果

```excel
=TEXTJOIN(",", TRUE, FILTER(B2:B100, C2:C100 = "VIP"))
```

### 带编号的合并

```excel
=LET(
  data, FILTER(A2:A50, A2:A50 <> ""),
  n, ROWS(data),
  idx, SEQUENCE(n),
  numbered, MAP(idx, data, LAMBDA(i, v, i & ". " & v)),
  TEXTJOIN(CHAR(10), TRUE, numbered)
)
```

> 📌 用 SEQUENCE 生成序号，MAP 拼接后用 TEXTJOIN + 换行符合并。

---

## 3. TEXTBEFORE / TEXTAFTER：截取文本

### 取分隔符之前的部分

```excel
=TEXTBEFORE(A2, "@")
```

> "user@example.com" → "user"

### 取分隔符之后的部分

```excel
=TEXTAFTER(A2, "@")
```

> "user@example.com" → "example.com"

### 取第 N 次出现的分隔符之后

```excel
=TEXTAFTER(A2, "/", 2)
```

> "a/b/c/d" → "c/d"（第 2 个 `/` 之后）

### 批量提取域名

```excel
=MAP(A2:A100, LAMBDA(email,
  TEXTBEFORE(TEXTAFTER(email, "@"), ".")
))
```

> "user@example.com" → "example"

---

## 4. VALUETOTEXT / ARRAYTOTEXT：值转文本

### 显示公式中间结果（调试用）

```excel
=ARRAYTOTEXT(SEQUENCE(5), 1)
```

> 返回 `"{1;2;3;4;5}"`——直观看到数组内容。

---

## 5. 组合实战

### 解析 CSV 行并转为表格

```excel
=LET(
  csv_text, A2,
  lines, TEXTSPLIT(csv_text,, CHAR(10)),
  n_rows, ROWS(lines),
  MAP(SEQUENCE(n_rows), LAMBDA(i,
    TEXTSPLIT(INDEX(lines, i), ",")
  ))
)
```

> 📌 先按换行拆行，再按逗号拆列。

### 批量替换（多模式）

将多个旧值替换为新值：

```excel
=LET(
  text, A2,
  old_vals, {"kg","m","cm"},
  new_vals, {"千克","米","厘米"},
  REDUCE(text, SEQUENCE(COLUMNS(old_vals)), LAMBDA(acc, i,
    SUBSTITUTE(acc, INDEX(old_vals,, i), INDEX(new_vals,, i))
  ))
)
```

> 📌 REDUCE + SEQUENCE 遍历替换对，逐个 SUBSTITUTE，模拟多轮替换循环。

### 提取文本中所有数字

```excel
=LET(
  text, A2,
  len, LEN(text),
  chars, MID(text, SEQUENCE(len), 1),
  digits, FILTER(chars, ISNUMBER(VALUE(chars)), ""),
  TEXTJOIN("", TRUE, digits)
)
```

> 📌 `MID + SEQUENCE` 将文本拆为单字符数组，FILTER 保留数字字符，TEXTJOIN 拼回。

---

[← 返回目录](../README.md)
