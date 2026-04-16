# 10 — HSTACK 横向拼接：左右合并多个数据集

> **适用版本**：Microsoft 365 / Excel 2021+

## 简介

`HSTACK` 将多个数组**横向（左右）并排**，返回一个合并后的数组。常用于为数据添加计算列、合并维度表或拼装报表。

语法：`=HSTACK(array1, [array2, ...])`

---

## 示例 1 ⭐ — 并排展示两个区域

```excel
=HSTACK(A2:B20, D2:E20)
```

**说明**：将两个 2 列区域左右拼成 4 列。两个区域的**行数必须相同**。

---

## 示例 2 ⭐⭐ — 为数据表添加序号列

```excel
=HSTACK(SEQUENCE(ROWS(A2:D50)), A2:D50)
```

**说明**：`SEQUENCE` 自动生成 1, 2, 3 … 的序号列，拼接到数据表最左侧。

---

## 示例 3 ⭐⭐⭐ — 添加计算列：合计与占比

```excel
=LET(
  data, B2:D50,
  rowSum, BYROW(data, LAMBDA(r, SUM(r))),
  total, SUM(data),
  pct, MAP(rowSum, LAMBDA(s, TEXT(s/total, "0.0%"))),
  HSTACK(data, rowSum, pct)
)
```

**说明**：

1. `rowSum` — 每行合计
2. `pct` — 每行占总合计的百分比
3. `HSTACK` 将原始数据、合计列、占比列拼在一起

---

## 示例 4 ⭐⭐⭐⭐ — XLOOKUP 关联维度表

主表 A:C（订单号、产品ID、数量），维度表 F:G（产品ID、产品名称）。

```excel
=LET(
  orders,   A2:C100,
  prodID,   INDEX(orders,,2),
  prodName, MAP(prodID, LAMBDA(id, XLOOKUP(id, F:F, G:G, "未知"))),
  HSTACK(orders, prodName)
)
```

**说明**：通过 `MAP + XLOOKUP` 逐行查找产品名称，再 `HSTACK` 拼到主表右侧。相当于 SQL 中的 `LEFT JOIN`。

---

## 示例 5 ⭐⭐⭐⭐⭐ — 动态拼接多级汇总报表

```excel
=LET(
  raw,       A2:E200,
  region,    INDEX(raw,,2),
  amount,    INDEX(raw,,5),
  u,         SORT(UNIQUE(FILTER(region, region<>""))),
  sumCol,    MAP(u, LAMBDA(r, SUM(FILTER(amount, region=r)))),
  avgCol,    MAP(u, LAMBDA(r, ROUND(AVERAGE(FILTER(amount, region=r)),2))),
  cntCol,    MAP(u, LAMBDA(r, ROWS(FILTER(amount, region=r)))),
  header,    {"区域","合计","均值","笔数"},
  body,      HSTACK(u, sumCol, avgCol, cntCol),
  VSTACK(header, body)
)
```

**说明**：

1. 提取唯一区域并排序
2. 分别计算合计、均值、笔数
3. `HSTACK` 横向拼成 4 列
4. `VSTACK` 加上表头

一个纯公式就能生成完整的分组统计报表。

---

## 应用注意点

| 要点 | 说明 |
|------|------|
| 行数一致 | 各数组行数应相同；不同时 Excel 会用 `#N/A` 补齐 |
| 与 VSTACK 搭配 | 先 HSTACK 补列/加列 → 再 VSTACK 合并行，是常见组合 |
| 常量数组 | 可以直接写 `{"A","B","C"}` 作为参数（单行常量数组） |
| 动态列 | 用 HSTACK 拼接的列会随数据变化自动更新 |
| 溢出区域 | 注意目标右侧不要有其他数据，否则溢出会报 `#SPILL!` 错误 |
