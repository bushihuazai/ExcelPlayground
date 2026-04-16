# 组合技：纯公式数据流水线

> 将多个函数串联，在一个公式中完成"清洗 → 转换 → 聚合 → 输出"全流程。
> 核心思路：**LET 搭骨架，SEQUENCE/INDEX 做索引，FILTER/SORT/UNIQUE 做声明式处理**。

---

## 1. 区域销售汇总表

目标：从原始明细生成"按区域汇总的销售排行"。

```excel
=LET(
  raw, A2:E200,
  valid, FILTER(raw, INDEX(raw,, 5) > 0),
  region, INDEX(valid,, 2),
  amount, INDEX(valid,, 5),
  u, UNIQUE(region),
  total, MAP(u, LAMBDA(r, SUM(FILTER(amount, region = r)))),
  SORT(HSTACK(u, total), 2, -1)
)
```

**解读：**
1. `INDEX(raw,, 5)` — 提取第 5 列用于筛选，避免硬编码 `E2:E200`
2. `FILTER` — 声明式清除无效数据
3. `UNIQUE` — 自动获取区域列表
4. `MAP + FILTER + SUM` — 对每个区域聚合
5. `HSTACK + SORT` — 拼成表并按金额降序

---

## 2. 分组 Top-N 报表

目标：每个区域取销售额前 3 名。

```excel
=LET(
  data, A2:E200,
  region_col, INDEX(data,, 2),
  name_col, INDEX(data,, 3),
  amount_col, INDEX(data,, 5),
  groups, SORT(UNIQUE(region_col)),
  n, 3,
  REDUCE("", SEQUENCE(ROWS(groups)), LAMBDA(acc, i,
    LET(
      g, INDEX(groups, i),
      subset, FILTER(HSTACK(region_col, name_col, amount_col), region_col = g),
      sorted_sub, SORT(subset, 3, -1),
      top, TAKE(sorted_sub, MIN(n, ROWS(sorted_sub))),
      IF(acc = "", top, VSTACK(acc, top))
    )
  ))
)
```

**关键技巧：**
- `SEQUENCE(ROWS(groups))` 生成分组索引
- `REDUCE + VSTACK` 逐组累积结果
- `TAKE` 截取前 N 行
- `MIN(n, ROWS(...))` 防止组内数据不足 N 条时报错

---

## 3. 交叉透视表（Pivot）

目标：行 = 区域，列 = 产品，值 = 销售额合计。

```excel
=LET(
  data, A2:E200,
  region, INDEX(data,, 2),
  product, INDEX(data,, 4),
  amount, INDEX(data,, 5),
  u_region, SORT(UNIQUE(region)),
  u_product, SORT(UNIQUE(product)),
  n_r, ROWS(u_region),
  n_p, ROWS(u_product),
  body, MAKEARRAY(n_r, n_p, LAMBDA(r, c,
    LET(
      rgn, INDEX(u_region, r),
      prd, INDEX(u_product, c),
      SUMPRODUCT((region = rgn) * (product = prd) * amount)
    )
  )),
  header, HSTACK("区域 \ 产品", TOROW(u_product)),
  rows, HSTACK(u_region, body),
  VSTACK(header, rows)
)
```

**关键技巧：**
- MAKEARRAY 生成交叉汇总的核心矩阵
- SUMPRODUCT 实现多条件求和
- VSTACK + HSTACK 拼接表头和数据

---

## 4. 同比环比计算

目标：给每月销售额添加环比增长率。

```excel
=LET(
  months, A2:A13,
  sales, B2:B13,
  n, ROWS(sales),
  prev, INDEX(sales, SEQUENCE(n - 1)),
  curr, INDEX(sales, SEQUENCE(n - 1, 1, 2)),
  growth, (curr - prev) / prev,
  result_months, INDEX(months, SEQUENCE(n - 1, 1, 2)),
  HSTACK(result_months, INDEX(sales, SEQUENCE(n - 1, 1, 2)), growth)
)
```

**关键技巧：**
- `SEQUENCE(n-1)` 和 `SEQUENCE(n-1, 1, 2)` 分别生成"上期"和"本期"的行号
- 两个 INDEX 取出错位数组后直接做运算

---

## 5. 日期区间统计

目标：统计每个自定义区间的订单数量。

```excel
=LET(
  order_dates, A2:A500,
  breaks, {0,7,30,90,365},
  n, COLUMNS(breaks) - 1,
  today, TODAY(),
  age, today - order_dates,
  labels, MAP(SEQUENCE(1, n), LAMBDA(i,
    INDEX(breaks,, i) & "~" & INDEX(breaks,, i + 1) & "天"
  )),
  counts, MAP(SEQUENCE(1, n), LAMBDA(i,
    LET(
      lo, INDEX(breaks,, i),
      hi, INDEX(breaks,, i + 1),
      SUM((age >= lo) * (age < hi))
    )
  )),
  VSTACK(labels, counts)
)
```

**关键技巧：**
- `{0,7,30,90,365}` 常量数组定义区间断点
- `SEQUENCE(1, n)` 遍历区间索引
- MAP + INDEX 动态生成标签和计数

---

## 6. 带条件的累计余额

目标：根据交易类型（收入/支出）和条件计算余额。

```excel
=LET(
  dates, A2:A50,
  types, B2:B50,
  amounts, C2:C50,
  init_balance, 10000,
  signed, MAP(types, amounts, LAMBDA(t, a,
    IF(t = "收入", a, -a)
  )),
  balance, SCAN(init_balance, signed, LAMBDA(acc, x, MAX(0, acc + x))),
  HSTACK(dates, types, amounts, balance)
)
```

> 📌 `MAX(0, acc + x)` 保证余额不为负。一个公式输出完整账单。

---

## 7. 数据验证报告

目标：检查数据质量，列出所有问题行。

```excel
=LET(
  data, A2:E100,
  n, ROWS(data),
  row_idx, SEQUENCE(n),
  name_empty, INDEX(data,, 1) = "",
  amount_neg, INDEX(data,, 5) < 0,
  date_invalid, NOT(ISNUMBER(INDEX(data,, 3))),
  has_error, name_empty + amount_neg + date_invalid,
  error_desc, MAP(row_idx, name_empty, amount_neg, date_invalid,
    LAMBDA(i, e1, e2, e3,
      LET(
        msgs, FILTER(
          HSTACK("姓名为空", "金额为负", "日期无效"),
          HSTACK(e1, e2, e3)
        ),
        IF(COLUMNS(msgs) > 0, "行" & i & ": " & TEXTJOIN("; ", TRUE, msgs), "")
      )
    )
  ),
  FILTER(error_desc, error_desc <> "", "数据验证通过 ✓")
)
```

> 📌 `SEQUENCE` 提供行号，多个条件布尔列组合判断，FILTER 输出非空错误描述。

---

[← 返回目录](../README.md)
