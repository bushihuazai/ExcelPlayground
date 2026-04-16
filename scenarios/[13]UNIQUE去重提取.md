# 13 — UNIQUE 去重提取：获取不重复值列表

> **适用版本**：Microsoft 365 / Excel 2021+

## 简介

`UNIQUE` 从数组中提取不重复的值，返回动态数组。

语法：`=UNIQUE(array, [by_col], [exactly_once])`

| 参数 | 默认 | 说明 |
|------|------|------|
| `by_col` | FALSE | TRUE = 按列去重；FALSE = 按行去重 |
| `exactly_once` | FALSE | TRUE = 只返回**仅出现一次**的值 |

---

## 示例 1 ⭐ — 基本去重

```excel
=UNIQUE(B2:B100)
```

**说明**：从 B 列提取所有不重复的值。结果为动态数组，自动溢出。

---

## 示例 2 ⭐⭐ — 去重并排序

```excel
=SORT(UNIQUE(B2:B200))
```

**说明**：先去重再按字母/数值升序排列。这是生成"下拉选项列表"的常用手法。

---

## 示例 3 ⭐⭐⭐ — 多列联合去重

```excel
=UNIQUE(B2:C100)
```

**说明**：以 B、C 两列的**组合**为去重依据。例如 B 列是"部门"、C 列是"职位"，结果是所有不重复的"部门-职位"组合。

---

## 示例 4 ⭐⭐⭐⭐ — 仅出现一次的值（exactly_once）

```excel
=UNIQUE(B2:B200, FALSE, TRUE)
```

**说明**：第三参数 `TRUE` 表示只返回在原数据中**仅出现过一次**的值。适用于查找唯一客户、异常记录等。

---

## 示例 5 ⭐⭐⭐⭐⭐ — UNIQUE + FILTER + MAP：分组计数报表

```excel
=LET(
  category, C2:C500,
  amount,   E2:E500,
  u,        SORT(UNIQUE(FILTER(category, category<>""))),
  cnt,      MAP(u, LAMBDA(c, ROWS(FILTER(category, category=c)))),
  total,    MAP(u, LAMBDA(c, SUM(FILTER(amount, category=c)))),
  avg,      MAP(u, LAMBDA(c, ROUND(AVERAGE(FILTER(amount, category=c)),2))),
  header,   {"类别","笔数","合计","均值"},
  VSTACK(header, HSTACK(u, cnt, total, avg))
)
```

**说明**：

1. `UNIQUE` 提取非空类别列表
2. `MAP` 对每个类别分别计数、求和、均值
3. `HSTACK + VSTACK` 组装成带表头的报表

这是"纯公式透视表"的典型实现。

---

## 应用注意点

| 要点 | 说明 |
|------|------|
| 空值处理 | UNIQUE 会保留空值，建议先 `FILTER` 去空再去重 |
| 大小写 | UNIQUE 默认**不区分**大小写（"ABC" = "abc"），如需区分需额外处理 |
| exactly_once | 第三参数 TRUE 不是"去重"，而是"找唯一项"，语义不同 |
| 多列去重 | 传入多列时，去重基于所有列的组合值 |
| 与 COUNTIF 搭配 | `COUNTIF(range, UNIQUE(range))` 可快速得到每个唯一值的出现次数 |
