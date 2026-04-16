# 09 — VSTACK 纵向拼接：上下合并多个数据集

> **适用版本**：Microsoft 365 / Excel 2021+

## 简介

`VSTACK` 将多个数组**纵向（上下）堆叠**，返回一个合并后的数组。它是数据整合中最常用的工具之一。

语法：`=VSTACK(array1, [array2, ...])`

---

## 示例 1 ⭐ — 合并两个月份的数据

```excel
=VSTACK(A2:D20, A25:D40)
```

**说明**：将第 2–20 行和第 25–40 行的 4 列数据上下拼接。两个区域的**列数必须相同**。

---

## 示例 2 ⭐⭐ — 合并时添加表头

```excel
=VSTACK(
  {"姓名","部门","工资","入职日期"},
  Sheet1!A2:D50,
  Sheet2!A2:D30
)
```

**说明**：第一个参数是**常量数组**作为表头行，后面跟两个工作表的数据，一次性输出带表头的完整表。

---

## 示例 3 ⭐⭐⭐ — 多表合并 + 添加来源列

```excel
=LET(
  t1, Sheet1!A2:C50,
  t2, Sheet2!A2:C30,
  t3, Sheet3!A2:C40,
  tag1, MAKEARRAY(ROWS(t1), 1, LAMBDA(r,c, "门店A")),
  tag2, MAKEARRAY(ROWS(t2), 1, LAMBDA(r,c, "门店B")),
  tag3, MAKEARRAY(ROWS(t3), 1, LAMBDA(r,c, "门店C")),
  VSTACK(
    HSTACK(tag1, t1),
    HSTACK(tag2, t2),
    HSTACK(tag3, t3)
  )
)
```

**说明**：先用 `MAKEARRAY` 生成标识列（来源标签），再用 `HSTACK` 拼到每个表左侧，最后 `VSTACK` 上下合并。合并后可以清楚区分数据来源。

---

## 示例 4 ⭐⭐⭐⭐ — 列数不同时先补齐再合并

表 1 有 3 列（A:C），表 2 只有 2 列（A:B），需要补齐。

```excel
=LET(
  t1, A2:C20,
  t2, E2:F15,
  padding, MAKEARRAY(ROWS(t2), 1, LAMBDA(r,c, "")),
  t2_fixed, HSTACK(t2, padding),
  VSTACK(t1, t2_fixed)
)
```

**说明**：用 `MAKEARRAY` 生成空列，`HSTACK` 补到 t2 右侧使列数对齐，再 `VSTACK` 合并。

> ⚠️ 如果列数不一致直接 VSTACK，Excel 会用 `#N/A` 填充缺失列。

---

## 示例 5 ⭐⭐⭐⭐⭐ — 动态合并：跳过空表

```excel
=LET(
  t1, IF(ROWS(Sheet1!A2:A1000)>0, Sheet1!A2:C1000, ""),
  t2, IF(ROWS(Sheet2!A2:A1000)>0, Sheet2!A2:C1000, ""),
  raw, VSTACK(t1, t2),
  FILTER(raw, INDEX(raw,,1)<>"")
)
```

**说明**：合并后通过 `FILTER` 去掉首列为空的行，实现"动态裁剪"——只保留实际有数据的行。

> 💡 在实际场景中，建议配合 `ROWS(FILTER(...))` 判断各表是否为空，再决定是否参与合并。

---

## 应用注意点

| 要点 | 说明 |
|------|------|
| 列数一致 | 各数组列数应相同；不同时 Excel 会用 `#N/A` 补齐 |
| 数据类型 | 混合数值和文本列时注意，VSTACK 不做类型转换 |
| 表头处理 | 合并多表时注意去除重复表头，或手动添加统一表头 |
| 与 HSTACK | VSTACK 上下拼，HSTACK 左右拼，常配合使用 |
| 溢出区域 | 结果为动态数组，会自动溢出；确保目标区域有足够空间 |
