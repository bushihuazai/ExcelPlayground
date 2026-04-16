# ExcelPlayground

一个总结和示意 **Excel 基于纯公式的"可编程技巧库"**——用 LAMBDA、SEQUENCE、INDEX 等动态数组函数把 Excel 变成函数式编程环境。

> **适用版本**：Microsoft 365 / Excel 2021+（支持动态数组与 LAMBDA 系列函数）
>
> 所有示例可直接粘贴到单元格中（根据你的区域设置调整 `,` 或 `;` 分隔符）。

---

## 📚 目录

| # | 专题 | 说明 |
|---|------|------|
| 1 | [LAMBDA + LET](examples/01-lambda-let.md) | 定义可复用函数、局部变量、流水线骨架 |
| 2 | [SEQUENCE + INDEX](examples/02-sequence-index.md) | ⭐ 动态数组核心——生成序列、选行选列、Top-N、逆序、采样 |
| 3 | [MAP / BYROW / BYCOL](examples/03-map-byrow-bycol.md) | 逐元素、按行、按列的高阶映射 |
| 4 | [REDUCE / SCAN](examples/04-reduce-scan.md) | 折叠聚合与运行累计 |
| 5 | [MAKEARRAY](examples/05-makearray.md) | 规则生成数组（含 SEQUENCE 广播替代方案） |
| 6 | [拼接与重塑](examples/06-stack-reshape.md) | VSTACK/HSTACK/CHOOSECOLS/DROP/TAKE/TOCOL/WRAPCOLS |
| 7 | [声明式数据清洗](examples/07-filter-sort-unique.md) | FILTER/SORT/SORTBY/UNIQUE/XLOOKUP/XMATCH |
| 8 | [文本处理](examples/08-text-processing.md) | TEXTSPLIT/TEXTJOIN/TEXTBEFORE/TEXTAFTER |
| 9 | [组合技：实战流水线](examples/09-advanced-pipelines.md) | 多步纯公式数据管道——清洗→聚合→输出 |
| 10 | [命名函数库](examples/10-named-functions.md) | 名称管理器中可复用的 LAMBDA 函数集 |

> 📖 每篇文档包含**基础示例 + 进阶技巧 + 实用组合**，建议按顺序阅读或按需跳转。

---

## 1. 核心理念：把 Excel 当成函数式编程环境

### 核心函数速查

| 类别 | 函数 | 一句话说明 |
|------|------|-----------|
| **定义** | `LAMBDA` | 定义可复用函数 |
| | `LET` | 定义局部变量，提升可读性和性能 |
| **生成** | `SEQUENCE` | ⭐ 生成数字序列（行号、日期、索引…） |
| | `MAKEARRAY` | 按行列号规则生成二维数组 |
| **选取** | `INDEX` | ⭐ 按行号/列号选取（选行、选列、切片） |
| | `CHOOSECOLS` / `CHOOSEROWS` | 按位置选取列/行 |
| | `DROP` / `TAKE` | 截取/丢弃头尾行列 |
| **映射** | `MAP` | 逐元素映射 |
| | `BYROW` / `BYCOL` | 按行/按列映射 |
| **折叠** | `REDUCE` | 折叠聚合为单值 |
| | `SCAN` | 返回每步累计值（过程可视化） |
| **拼接** | `VSTACK` / `HSTACK` | 纵向/横向拼表 |
| | `TOCOL` / `TOROW` | 展平为列/行 |
| | `WRAPCOLS` / `WRAPROWS` | 将一维数组重塑为二维 |
| **清洗** | `FILTER` | 条件筛选 |
| | `SORT` / `SORTBY` | 排序 |
| | `UNIQUE` | 去重 |
| | `XLOOKUP` / `XMATCH` | 增强查找 |
| **文本** | `TEXTSPLIT` / `TEXTJOIN` | 拆分/合并文本 |
| | `TEXTBEFORE` / `TEXTAFTER` | 截取文本片段 |

---

## 2. 精选示例速览

> 完整示例及更多进阶技巧请点击上方目录中的对应链接。

### 2.1 LAMBDA + LET：封装可读函数

```excel
=LAMBDA(price, discount,
  LET(result, price * (1 - discount), MAX(0, result))
)(A2, B2)
```

→ [更多示例](examples/01-lambda-let.md)

---

### 2.2 ⭐ SEQUENCE + INDEX：简洁的核心技巧

**生成序列**

```excel
=SEQUENCE(10)                   -- 1~10 列向量
=SEQUENCE(1, 10)                -- 1~10 行向量
=SEQUENCE(10, 1, 0, 2)          -- 偶数：0,2,4,...,18
=SEQUENCE(7, 1, TODAY(), 1)     -- 连续 7 天日期
```

**INDEX 选列/选行**

```excel
=INDEX(A2:E100,, 3)             -- 取第 3 列
=INDEX(A2:E100, SEQUENCE(ROWS(A2:E100)), {1,3,5})  -- 取第 1/3/5 列
```

**SEQUENCE + INDEX 组合**

```excel
-- 取前 N 行
=LET(
  data, A2:E100, n, 10,
  INDEX(data, SEQUENCE(n), SEQUENCE(1, COLUMNS(data)))
)

-- 逆序排列
=LET(
  data, A2:A20, n, ROWS(data),
  INDEX(data, SEQUENCE(n, 1, n, -1))
)

-- 乘法表（SEQUENCE 广播 > MAKEARRAY）
=SEQUENCE(9) * SEQUENCE(1, 9)
```

→ [更多示例](examples/02-sequence-index.md)

---

### 2.3 MAP / BYROW / BYCOL：映射

```excel
=MAP(A2:A10, LAMBDA(x, UPPER(TRIM(x))))
=BYROW(B2:F10, LAMBDA(r, SUM(r)))
=BYCOL(B2:F10, LAMBDA(c, MAX(c)))
```

→ [更多示例](examples/03-map-byrow-bycol.md)

---

### 2.4 REDUCE / SCAN：折叠与累计

```excel
=REDUCE(0, A2:A10, LAMBDA(acc, x, acc + x))              -- 累加
=REDUCE(1, SEQUENCE(10), LAMBDA(acc, x, acc * x))         -- 10! 阶乘
=SCAN(1000, B2:B20, LAMBDA(bal, cf, bal + cf))             -- 余额轨迹
```

→ [更多示例](examples/04-reduce-scan.md)

---

### 2.5 MAKEARRAY vs SEQUENCE 广播

```excel
-- MAKEARRAY 写法
=MAKEARRAY(8, 8, LAMBDA(r, c, MOD(r + c, 2)))

-- SEQUENCE 简洁写法（推荐）
=MOD(SEQUENCE(8) + SEQUENCE(1, 8), 2)
```

→ [更多示例](examples/05-makearray.md)

---

### 2.6 拼接与重塑

```excel
=VSTACK(A2:C10, A15:C20)                        -- 上下合并
=HSTACK(SEQUENCE(ROWS(A2:E100)), A2:E100)        -- 添加序号列
=CHOOSECOLS(A2:F100, 1, 3, 5)                    -- 选取指定列
=TAKE(A2:F100, 10, 3)                            -- 取前 10 行前 3 列
=WRAPROWS(SEQUENCE(12), 4)                       -- 1~12 排为 3×4
```

→ [更多示例](examples/06-stack-reshape.md)

---

### 2.7 声明式数据清洗

```excel
-- 用 INDEX 替代硬编码列引用
=LET(
  data, A2:E100,
  region, INDEX(data,, 2),
  amount, INDEX(data,, 5),
  FILTER(data, (region = "华东") * (amount > 10000), "无结果")
)

-- 去重 + 计数 + 排序 = 频率表
=LET(
  col, C2:C100,
  u, UNIQUE(col),
  cnt, MAP(u, LAMBDA(v, SUM((col = v) * 1))),
  SORT(HSTACK(u, cnt), 2, -1)
)
```

→ [更多示例](examples/07-filter-sort-unique.md)

---

### 2.8 文本处理

```excel
=TEXTSPLIT(A2, "-")                              -- 拆分
=TEXTJOIN(",", TRUE, FILTER(B2:B100, B2:B100<>""))  -- 合并

-- SEQUENCE + MID 提取所有数字
=LET(
  text, A2, len, LEN(text),
  chars, MID(text, SEQUENCE(len), 1),
  digits, FILTER(chars, ISNUMBER(VALUE(chars)), ""),
  TEXTJOIN("", TRUE, digits)
)
```

→ [更多示例](examples/08-text-processing.md)

---

### 2.9 组合技：纯公式流水线

**区域销售汇总（完整管道）**

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

**同比环比计算**

```excel
=LET(
  sales, B2:B13,
  n, ROWS(sales),
  prev, INDEX(sales, SEQUENCE(n - 1)),
  curr, INDEX(sales, SEQUENCE(n - 1, 1, 2)),
  growth, (curr - prev) / prev,
  HSTACK(INDEX(A2:A13, SEQUENCE(n - 1, 1, 2)), curr, growth)
)
```

→ [更多示例](examples/09-advanced-pipelines.md)

---

## 3. 命名函数库

在 **公式 → 名称管理器** 中定义 LAMBDA，团队共享、无需 VBA：

| 函数名 | 用途 | 核心技巧 |
|--------|------|----------|
| `SAFE_DIV(a, b)` | 安全除法 | `IF(b=0, "", a/b)` |
| `RUNNING_TOTAL(arr)` | 运行累计 | `SCAN` |
| `GROUP_SUM(keys, vals)` | 分组求和 | `UNIQUE + MAP + FILTER` |
| `GROUP_COUNT(keys)` | 分组计数 | `UNIQUE + MAP` |
| `TOP_N(data, n, col)` | 前 N 行 | `SORT + TAKE` |
| `REVERSE(arr)` | 逆序 | `SEQUENCE(n,1,n,-1) + INDEX` |
| `PAGINATE(data, page, size)` | 分页取数 | `SEQUENCE(size,1,start) + INDEX` |
| `MOVING_AVG(arr, window)` | 移动平均 | `SEQUENCE` 滑动窗口 + `INDEX` |
| `UNPIVOT(rows, cols, vals)` | 逆透视 | `SEQUENCE + 整除/取余索引` |
| `PERCENTILE_RANK(arr)` | 百分位排名 | `MAP` + 向量化比较 |

→ [完整定义与用法](examples/10-named-functions.md)

---

## 4. 💡 SEQUENCE / INDEX 技巧速查

这两个函数是贯穿全库的"瑞士军刀"，以下是最常用的模式：

| 模式 | 公式 | 说明 |
|------|------|------|
| 行号序列 | `SEQUENCE(n)` | 1,2,…,n |
| 列号序列 | `SEQUENCE(1, n)` | 横向 1,2,…,n |
| 偶数序列 | `SEQUENCE(n, 1, 0, 2)` | 0,2,4,…,2(n-1) |
| 逆序序列 | `SEQUENCE(n, 1, n, -1)` | n,n-1,…,1 |
| 日期序列 | `SEQUENCE(7, 1, TODAY(), 1)` | 连续 7 天 |
| 选整列 | `INDEX(data,, k)` | 第 k 列 |
| 选多列 | `INDEX(data, SEQUENCE(ROWS(data)), {1,3,5})` | 第 1/3/5 列 |
| 取前 N 行 | `INDEX(data, SEQUENCE(n), SEQUENCE(1, COLUMNS(data)))` | Top-N 切片 |
| 取后 N 行 | `INDEX(data, SEQUENCE(n, 1, ROWS(data)-n+1), ...)` | 尾部切片 |
| 每隔 k 行 | `INDEX(data, SEQUENCE(INT(total/k), 1, 1, k), ...)` | 等间距采样 |
| 逆序排列 | `INDEX(data, SEQUENCE(n, 1, n, -1))` | 翻转数组 |
| 乘法广播 | `SEQUENCE(n) * SEQUENCE(1, m)` | n×m 乘法表 |
| 循环模式 | `MOD(SEQUENCE(n) - 1, k) + 1` | 1,2,…,k,1,2,…,k |
| 分组编号 | `INT((SEQUENCE(n) - 1) / k) + 1` | 1,1,1,2,2,2,… |
| 填充矩阵 | `SEQUENCE(rows, cols, 0, 0)` | 全零矩阵 |

---

## 5. 使用建议

1. **优先用 `LET`** 给中间结果命名——降低维护成本，便于调试
2. **优先用 `SEQUENCE` 生成索引**——替代辅助列和手写常量数组
3. **优先用 `INDEX(data,, col)` 引用列**——当数据结构变化时只需改 data 引用
4. **能用 SEQUENCE 广播就不用 MAKEARRAY**——更简洁（如乘法表、棋盘格）
5. **用 `TAKE`/`DROP` 替代复杂的 INDEX 切片**——语义更清晰
6. **用 `CHOOSECOLS`/`CHOOSEROWS` 选取指定位置**——比 INDEX 多列写法更直观
7. 先搭建"可读版本"，再考虑压缩为短公式
8. 对大型数据尽量减少重复计算（缓存到 `LET`）
9. 为关键 LAMBDA 建立命名函数，形成团队可复用模板

---

## 📜 License

本项目采用 [MIT License](LICENSE)。

欢迎持续补充更多基于 **LAMBDA / SEQUENCE / INDEX / REDUCE / SCAN** 的实战示例，把本仓库演进为完整的 Excel 纯公式编程手册！
