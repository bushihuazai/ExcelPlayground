# 02 — LET 局部变量：让公式清晰且高效

> **适用版本**：Microsoft 365 / Excel 2021+

## 简介

`LET` 允许在公式内部定义**局部变量**，将复杂表达式分步命名。好处有两个：

1. **可读性**：用名称代替重复的子表达式。
2. **性能**：同一子表达式只计算一次，Excel 自动缓存。

语法：`=LET(name1, value1, [name2, value2, ...], 计算表达式)`

---

## 示例 1 ⭐ — 基本用法：计算含税价

```excel
=LET(
  税率, 0.13,
  含税价, A2 * (1 + 税率),
  ROUND(含税价, 2)
)
```

**说明**：先定义 `税率`，再计算 `含税价`，最后四舍五入到两位小数。修改税率只需改一处。

---

## 示例 2 ⭐⭐ — 消除重复计算：BMI 评级

```excel
=LET(
  bmi, A2 / (B2/100)^2,
  IF(bmi<18.5, "偏瘦",
    IF(bmi<24, "正常",
      IF(bmi<28, "偏胖", "肥胖")
    )
  )
)
```

| 参数 | 说明 |
|------|------|
| `A2` | 体重（kg） |
| `B2` | 身高（cm） |

**说明**：`bmi` 只计算一次，后续多个 `IF` 分支直接引用变量名，避免重复计算。

---

## 示例 3 ⭐⭐⭐ — 多步数据处理：求部门最高薪与最低薪之差

```excel
=LET(
  dept,   C2:C100,
  salary, E2:E100,
  target, "技术部",
  filtered, FILTER(salary, dept=target),
  最大, MAX(filtered),
  最小, MIN(filtered),
  最大 - 最小
)
```

**说明**：`FILTER` 结果被缓存到 `filtered`，后续 `MAX` 和 `MIN` 共享同一筛选结果，避免两次 `FILTER` 调用。

---

## 示例 4 ⭐⭐⭐⭐ — 嵌套 LET 构建迷你流水线

```excel
=LET(
  raw,     A2:F200,
  noBlank, FILTER(raw, INDEX(raw,,1)<>""),
  sorted,  SORT(noBlank, 5, -1),
  top10,   INDEX(sorted, SEQUENCE(MIN(10,ROWS(sorted))), SEQUENCE(1, COLUMNS(sorted))),
  top10
)
```

**说明**：

1. `noBlank` — 去掉首列为空的行
2. `sorted`  — 按第 5 列降序排列
3. `top10`   — 取前 10 行（不足 10 行则取全部）

每一步都有清晰的变量名，相当于一个小型 ETL 管道。

---

## 示例 5 ⭐⭐⭐⭐⭐ — LET + LAMBDA 联合：分组统计

```excel
=LET(
  region,  B2:B200,
  amount,  E2:E200,
  u,       UNIQUE(FILTER(region, region<>"")),
  stats,   MAP(u, LAMBDA(r,
             LET(
               vals, FILTER(amount, region=r),
               avg,  AVERAGE(vals),
               cnt,  ROWS(vals),
               ROUND(avg, 2) & " (" & cnt & "笔)"
             )
           )),
  HSTACK(u, stats)
)
```

**说明**：外层 `LET` 定义数据源和去重列表，内层 `LET` 在 `MAP` 的 LAMBDA 中再做分步计算。嵌套使用让公式层次分明、维护方便。

---

## 应用注意点

| 要点 | 说明 |
|------|------|
| 变量作用域 | LET 中的变量只在当前公式内有效，不会污染其他单元格 |
| 变量顺序 | 后定义的变量可以引用前面的变量，但**不能**反向引用 |
| 性能优化 | 当同一子表达式出现 2 次以上，务必用 LET 缓存 |
| 命名规范 | 建议使用有意义的中英文名称，避免单字母（a, b, c） |
| 嵌套限制 | 理论上无嵌套层数限制，但过深的嵌套会降低可读性，建议不超过 5 层 |
