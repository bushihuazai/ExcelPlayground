# 05 — BYCOL 按列计算：将每一列压缩为一个值

> **适用版本**：Microsoft 365 / Excel 2021+

## 简介

`BYCOL` 将二维区域的**每一列**作为一维数组传入 LAMBDA，返回单个值。结果是一个 1×M 的行向量，常用于汇总统计。

语法：`=BYCOL(array, LAMBDA(col, 表达式))`

---

## 示例 1 ⭐ — 每列求最大值

```excel
=BYCOL(B2:F50, LAMBDA(c, MAX(c)))
```

**说明**：返回 1 行 5 列结果，分别是 B–F 五列的最大值。适合在汇总行展示列级统计。

---

## 示例 2 ⭐⭐ — 每列非空单元格占比

```excel
=BYCOL(B2:F100, LAMBDA(c,
  TEXT(COUNTA(c) / ROWS(c), "0.0%")
))
```

**说明**：`COUNTA(c)` 统计非空个数，除以总行数得到填充率，再格式化为百分比字符串。

---

## 示例 3 ⭐⭐⭐ — 每列标准差（忽略错误值）

```excel
=BYCOL(B2:G200, LAMBDA(c,
  LET(
    clean, FILTER(c, ISNUMBER(c)),
    IF(ROWS(clean)>1, STDEV(clean), "")
  )
))
```

**说明**：先用 `FILTER + ISNUMBER` 过滤非数值（文本、错误值），再求标准差。当有效值不足 2 个时返回空。

---

## 示例 4 ⭐⭐⭐⭐ — 动态生成列标题 + 统计摘要

```excel
=LET(
  data, B2:F100,
  headers, B1:F1,
  sums, BYCOL(data, LAMBDA(c, SUM(c))),
  avgs, BYCOL(data, LAMBDA(c, ROUND(AVERAGE(c),2))),
  VSTACK(
    HSTACK("指标", headers),
    HSTACK("合计", sums),
    HSTACK("均值", avgs)
  )
)
```

**说明**：利用 BYCOL 分别计算合计和均值，再用 VSTACK + HSTACK 拼成一张摘要表。这是一个"纯公式报表生成器"的雏形。

---

## 应用注意点

| 要点 | 说明 |
|------|------|
| 返回值形状 | LAMBDA 必须返回**单个值**（标量），不能返回数组 |
| 结果方向 | BYCOL 返回**行向量**（1×M），BYROW 返回**列向量**（N×1） |
| 适用场景 | 列级汇总（求和、均值、极值）、数据质量检查（空值率、异常率） |
| 与 MAP 区别 | MAP 逐元素、BYCOL 逐列，当需要列级聚合时用 BYCOL |
| 性能 | 列数通常远少于行数，因此 BYCOL 的性能瓶颈一般不明显 |
