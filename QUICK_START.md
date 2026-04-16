# ⚡ 快速入门：5 分钟上手 Excel 函数式编程

> 本指南帮助你用最短时间理解 Excel 动态数组函数的核心思路，并动手写出第一个"纯公式流水线"。

---

## 前提条件

- **Microsoft 365** 或 **Excel 2021+**（支持动态数组与 LAMBDA）
- 根据你的区域设置，公式中的 `,` 可能需要替换为 `;`

---

## 第 1 步：认识动态数组

传统 Excel 公式返回单个值，动态数组函数可以**一次返回多个值**，自动溢出到相邻单元格。

在任意空白单元格中输入：

```excel
=SEQUENCE(5)
```

你会看到 5 个单元格自动填入 1, 2, 3, 4, 5 —— 这就是"动态溢出"。

---

## 第 2 步：用 LET 命名中间变量

`LET` 让你在公式中定义"变量"，避免重复计算，提升可读性：

```excel
=LET(
  price, 100,
  discount, 0.2,
  final, price * (1 - discount),
  final
)
```

→ 结果：`80`

**要点**：最后一个参数是返回值，之前的参数成对出现（名称, 值）。

---

## 第 3 步：SEQUENCE + INDEX = 万能选取

这是全库使用频率最高的组合：

```excel
-- 生成 1~10 的序列
=SEQUENCE(10)

-- 从数据表取第 3 列
=INDEX(A2:E100,, 3)

-- 取数据表前 5 行
=LET(
  data, A2:E100,
  INDEX(data, SEQUENCE(5), SEQUENCE(1, COLUMNS(data)))
)
```

→ [完整参考](formulas/02-sequence-index.md)

---

## 第 4 步：用 FILTER 筛选数据

告别手动筛选，用公式声明"我要什么"：

```excel
=FILTER(A2:E100, INDEX(A2:E100,,5) > 10000, "无结果")
```

→ 自动返回第 5 列 > 10000 的所有行。

组合 AND 条件（乘法）和 OR 条件（加法）：

```excel
-- 区域="华东" 且 金额>10000
=FILTER(A2:E100, (INDEX(A2:E100,,2)="华东") * (INDEX(A2:E100,,5)>10000))

-- 区域="华东" 或 区域="华南"
=FILTER(A2:E100, (INDEX(A2:E100,,2)="华东") + (INDEX(A2:E100,,2)="华南"))
```

→ [完整参考](formulas/07-filter-sort-unique.md)

---

## 第 5 步：你的第一个"纯公式流水线"

把前面学到的组合起来——不用 VBA、不用 Power Query，纯公式完成数据分析：

```excel
=LET(
  raw,    A2:E200,                                       -- ① 原始数据
  valid,  FILTER(raw, INDEX(raw,,5) > 0),                -- ② 清洗：去除无效行
  region, INDEX(valid,, 2),                              -- ③ 提取区域列
  amount, INDEX(valid,, 5),                              -- ④ 提取金额列
  u,      UNIQUE(region),                                -- ⑤ 去重：获取所有区域
  total,  MAP(u, LAMBDA(r, SUM(FILTER(amount, region=r)))),  -- ⑥ 聚合：按区域求和
  SORT(HSTACK(u, total), 2, -1)                          -- ⑦ 输出：降序排列
)
```

**这个公式做了什么？**

```
原始数据 → 过滤无效 → 提取字段 → 分组去重 → 按组求和 → 排序输出
```

一个单元格，零辅助列，完成了一个完整的数据分析管道。

---

## 下一步

| 想学什么 | 去哪里 |
|----------|--------|
| 系统学习每个函数 | 📘 [`formulas/`](formulas/) — 10 篇基础公式参考 |
| 针对单个函数深入练习 | 📗 [`scenarios/basic/`](scenarios/basic/) — 14 篇基础函数场景 |
| 挑战真实业务案例 | 📕 [`scenarios/advanced/`](scenarios/advanced/) — 6 篇综合实战 |
| 建立团队函数库 | 📒 [命名函数库](formulas/10-named-functions.md) |
| 查看推荐学习路径 | 🗺️ [README · 学习路径](README.md#️-学习路径) |

---

> 💡 **记住这个原则**：优先用 `LET` 命名 → `SEQUENCE` 生成索引 → `INDEX` 选取数据 → `FILTER` 筛选 → `MAP`/`REDUCE` 处理 → `SORT`/`HSTACK` 输出。这就是 Excel 函数式编程的核心模式。
