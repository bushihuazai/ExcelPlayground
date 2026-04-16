# 🧮 ExcelPlayground

> **把 Excel 变成函数式编程环境** —— 用 LAMBDA、SEQUENCE、INDEX 等动态数组函数，构建纯公式的"可编程技巧库"。

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
![Excel Version](https://img.shields.io/badge/Excel-Microsoft%20365%20%2F%202021%2B-217346?logo=microsoftexcel&logoColor=white)

**适用版本**：Microsoft 365 / Excel 2021+（支持动态数组与 LAMBDA 系列函数）  
所有示例可直接粘贴到单元格中（根据区域设置调整 `,` 或 `;` 分隔符）。

---

## 📁 项目结构

```
ExcelPlayground/
├── README.md                      ← 你正在阅读的主文档
├── QUICK_START.md                 ← 快速入门（5 分钟上手）
│
├── formulas/                      ← 📘 基础公式参考（10 篇专题）
│   ├── 01-lambda-let.md
│   ├── 02-sequence-index.md
│   ├── ...
│   └── 10-named-functions.md
│
├── scenarios/                     ← 📗 场景示例
│   ├── basic/                     ←   基础函数场景（14 篇·单函数深入）
│   │   ├── [01]LAMBDA基础.md
│   │   ├── ...
│   │   └── [14]XLOOKUP高级查找.md
│   └── advanced/                  ←   综合实战案例（6 篇·多函数组合）
│       ├── [15]综合实战流水线.md
│       ├── ...
│       └── [20]数据质量与异常检测.md
│
├── assets/                        ← 📊 PlantUML 图示源文件
│   ├── function-taxonomy.puml
│   ├── learning-path.puml
│   └── data-pipeline.puml
│
└── LICENSE
```

**三层内容体系**：

| 层级 | 目录 | 定位 | 阅读方式 |
|:----:|------|------|----------|
| 📘 | `formulas/` | **基础公式参考** — 按函数分类的语法、示例与对比 | 按需查阅 |
| 📗 | `scenarios/basic/` | **基础函数场景** — 每篇聚焦一个函数，3–5 个由浅入深的实战题 | 顺序学习 |
| 📕 | `scenarios/advanced/` | **综合实战案例** — 多函数组合的领域级解决方案 | 挑战进阶 |

---

## 🗺️ 学习路径

> 以下学习路径图帮助你按照由浅入深的顺序掌握 Excel 函数式编程，PlantUML 源文件见 [`assets/learning-path.puml`](assets/learning-path.puml)。

```
┌─────────────────────────────────────────────────────────────────┐
│                    第一阶段 · 基础概念                           │
│  📘 LAMBDA + LET        →  📘 SEQUENCE + INDEX ⭐              │
│  定义函数 & 局部变量         生成序列 & 动态选取                  │
└────────────────────────────────┬────────────────────────────────┘
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                    第二阶段 · 高阶映射                           │
│  📗 MAP / BYROW / BYCOL  →  📗 REDUCE / SCAN                  │
│  逐元素 / 按行 / 按列映射     折叠聚合 & 运行累计                │
└────────────────────────────────┬────────────────────────────────┘
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                    第三阶段 · 数据操作                           │
│  📙 FILTER / SORT / UNIQUE    拼接：VSTACK / HSTACK / TOCOL    │
│  📙 XLOOKUP / XMATCH          文本：TEXTSPLIT / TEXTJOIN       │
│  📙 MAKEARRAY                  重塑：WRAPCOLS / DROP / TAKE     │
└────────────────────────────────┬────────────────────────────────┘
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                    第四阶段 · 综合实战                           │
│  📕 组合技：纯公式流水线      📕 命名函数库                     │
│  清洗 → 转换 → 聚合 → 输出    沉淀可复用 LAMBDA                 │
└────────────────────────────────┬────────────────────────────────┘
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                    第五阶段 · 领域场景                           │
│  🏢 人力资源  💰 财务报表  📦 库存管理  📅 排班计划  🔍 数据质量 │
└─────────────────────────────────────────────────────────────────┘
```

---

## 🧩 函数分类速查

> PlantUML 源文件见 [`assets/function-taxonomy.puml`](assets/function-taxonomy.puml)。

### 定义类

| 函数 | 用途 | 核心特征 |
|------|------|----------|
| `LAMBDA` | 定义可复用函数 | 支持递归、可存入名称管理器 |
| `LET` | 定义局部变量 | 缓存中间结果、提升可读性与性能 |

### 生成类

| 函数 | 用途 | 核心特征 |
|------|------|----------|
| `SEQUENCE` ⭐ | 生成数字序列 | 行号、日期、索引…最常用的"瑞士军刀" |
| `MAKEARRAY` | 按规则生成二维数组 | `LAMBDA(r,c,...)` 控制每个单元格 |

### 选取类

| 函数 | 用途 | 核心特征 |
|------|------|----------|
| `INDEX` ⭐ | 按行号/列号选取 | 选行、选列、切片 |
| `CHOOSECOLS` / `CHOOSEROWS` | 按位置选取列/行 | 比 INDEX 多列写法更直观 |
| `DROP` / `TAKE` | 截取/丢弃头尾行列 | 语义比 INDEX 切片更清晰 |

### 映射类

| 函数 | 用途 | 核心特征 |
|------|------|----------|
| `MAP` | 逐元素映射 | 支持多数组同步映射 |
| `BYROW` | 按行计算 | 返回 N×1 列向量 |
| `BYCOL` | 按列计算 | 返回 1×M 行向量 |

### 折叠类

| 函数 | 用途 | 核心特征 |
|------|------|----------|
| `REDUCE` | 折叠聚合为单值 | `(acc, x) → acc'` 模式 |
| `SCAN` | 返回每步累计值 | 过程可视化，便于调试 |

### 拼接与重塑类

| 函数 | 用途 | 核心特征 |
|------|------|----------|
| `VSTACK` / `HSTACK` | 纵向/横向拼接 | 列/行数需匹配 |
| `TOCOL` / `TOROW` | 展平为列/行 | 可忽略空值或错误 |
| `WRAPCOLS` / `WRAPROWS` | 一维 → 二维重塑 | 指定每列/每行元素数 |

### 数据清洗类

| 函数 | 用途 | 核心特征 |
|------|------|----------|
| `FILTER` | 条件筛选 | 支持 AND（`*`）/ OR（`+`）逻辑 |
| `SORT` / `SORTBY` | 排序 | SORTBY 支持外部排序键 |
| `UNIQUE` | 去重 | 支持 `exactly_once` 参数 |
| `XLOOKUP` / `XMATCH` | 增强查找 | 替代 VLOOKUP，支持左查找 |

### 文本处理类

| 函数 | 用途 | 核心特征 |
|------|------|----------|
| `TEXTSPLIT` | 按分隔符拆分文本 | 支持行列双向拆分 |
| `TEXTJOIN` | 合并文本 | 可忽略空值 |
| `TEXTBEFORE` / `TEXTAFTER` | 截取文本片段 | 支持第 N 次出现 |

---

## 📚 内容导航

### 📘 基础公式参考（`formulas/`）

每篇包含 **语法说明 + 基础示例 + 进阶技巧 + 实用组合**，建议按顺序阅读或按需跳转。

| # | 专题 | 核心函数 | 说明 |
|:-:|------|----------|------|
| 01 | [LAMBDA + LET](formulas/01-lambda-let.md) | `LAMBDA` `LET` | 定义可复用函数、局部变量、流水线骨架 |
| 02 | [SEQUENCE + INDEX](formulas/02-sequence-index.md) | `SEQUENCE` `INDEX` | ⭐ 动态数组核心——生成序列、选行选列、Top-N |
| 03 | [MAP / BYROW / BYCOL](formulas/03-map-byrow-bycol.md) | `MAP` `BYROW` `BYCOL` | 逐元素、按行、按列的高阶映射 |
| 04 | [REDUCE / SCAN](formulas/04-reduce-scan.md) | `REDUCE` `SCAN` | 折叠聚合与运行累计 |
| 05 | [MAKEARRAY](formulas/05-makearray.md) | `MAKEARRAY` | 规则生成数组（含 SEQUENCE 广播对比） |
| 06 | [拼接与重塑](formulas/06-stack-reshape.md) | `VSTACK` `HSTACK` 等 | 纵横拼接 / 展平 / 重塑 / 截取 |
| 07 | [声明式数据清洗](formulas/07-filter-sort-unique.md) | `FILTER` `SORT` 等 | 条件筛选 / 排序 / 去重 / 增强查找 |
| 08 | [文本处理](formulas/08-text-processing.md) | `TEXTSPLIT` `TEXTJOIN` 等 | 拆分 / 合并 / 截取文本 |
| 09 | [组合技：实战流水线](formulas/09-advanced-pipelines.md) | 多函数组合 | 多步纯公式数据管道 |
| 10 | [命名函数库](formulas/10-named-functions.md) | `LAMBDA` 命名函数 | 名称管理器中可复用的 LAMBDA 函数集 |

---

### 📗 基础函数场景（`scenarios/basic/`）

每篇聚焦**一个核心函数**，包含 3–5 个由浅入深的实战示例，层层递进讲解。

| 编号 | 场景 | 核心函数 | 难度 |
|:----:|------|----------|:----:|
| [01](scenarios/basic/[01]LAMBDA基础.md) | LAMBDA 基础 | `LAMBDA` 递归 | ⭐ |
| [02](scenarios/basic/[02]LET局部变量.md) | LET 局部变量 | `LET` | ⭐ |
| [03](scenarios/basic/[03]MAP逐元素处理.md) | MAP 逐元素处理 | `MAP` | ⭐⭐ |
| [04](scenarios/basic/[04]BYROW按行计算.md) | BYROW 按行计算 | `BYROW` | ⭐⭐ |
| [05](scenarios/basic/[05]BYCOL按列计算.md) | BYCOL 按列计算 | `BYCOL` | ⭐⭐ |
| [06](scenarios/basic/[06]REDUCE累计折叠.md) | REDUCE 累计折叠 | `REDUCE` | ⭐⭐⭐ |
| [07](scenarios/basic/[07]SCAN过程可视化.md) | SCAN 过程可视化 | `SCAN` | ⭐⭐⭐ |
| [08](scenarios/basic/[08]MAKEARRAY规则生成.md) | MAKEARRAY 规则生成 | `MAKEARRAY` | ⭐⭐ |
| [09](scenarios/basic/[09]VSTACK纵向拼接.md) | VSTACK 纵向拼接 | `VSTACK` | ⭐ |
| [10](scenarios/basic/[10]HSTACK横向拼接.md) | HSTACK 横向拼接 | `HSTACK` | ⭐ |
| [11](scenarios/basic/[11]FILTER条件筛选.md) | FILTER 条件筛选 | `FILTER` | ⭐⭐ |
| [12](scenarios/basic/[12]SORT与SORTBY排序.md) | SORT 与 SORTBY 排序 | `SORT` `SORTBY` | ⭐⭐ |
| [13](scenarios/basic/[13]UNIQUE去重提取.md) | UNIQUE 去重提取 | `UNIQUE` | ⭐⭐ |
| [14](scenarios/basic/[14]XLOOKUP高级查找.md) | XLOOKUP 高级查找 | `XLOOKUP` | ⭐⭐⭐ |

---

### 📕 综合实战案例（`scenarios/advanced/`）

多函数组合解决真实业务问题，包含完整的 ETL 管道与领域分析。

| 编号 | 场景 | 核心函数组合 | 难度 |
|:----:|------|-------------|:----:|
| [15](scenarios/advanced/[15]综合实战流水线.md) | 🔗 综合实战流水线 | `FILTER` + `SORT` + `UNIQUE` + `MAP` + `REDUCE` | ⭐⭐⭐⭐ |
| [16](scenarios/advanced/[16]人力资源数据分析.md) | 🏢 人力资源数据分析 | `UNIQUE` + `MAP` + `FILTER` + `MAKEARRAY` | ⭐⭐⭐⭐ |
| [17](scenarios/advanced/[17]财务报表与预算分析.md) | 💰 财务报表与预算分析 | `SUMPRODUCT` + `XLOOKUP` + `SCAN` | ⭐⭐⭐⭐⭐ |
| [18](scenarios/advanced/[18]库存管理与预警.md) | 📦 库存管理与预警 | `SCAN` + `SORTBY` + `MAKEARRAY` + `MAP` | ⭐⭐⭐⭐⭐ |
| [19](scenarios/advanced/[19]日期时间与排班计划.md) | 📅 日期时间与排班计划 | `SEQUENCE` + `MAKEARRAY` + `WEEKDAY` | ⭐⭐⭐⭐ |
| [20](scenarios/advanced/[20]数据质量与异常检测.md) | 🔍 数据质量与异常检测 | `BYROW` + `BYCOL` + `MAP` + `PERCENTILE` | ⭐⭐⭐⭐⭐ |

---

## ⚡ 核心技巧速查

### SEQUENCE / INDEX 常用模式

这两个函数是贯穿全库的"瑞士军刀"：

| 模式 | 公式 | 说明 |
|------|------|------|
| 行号序列 | `SEQUENCE(n)` | 1, 2, …, n |
| 列号序列 | `SEQUENCE(1, n)` | 横向 1, 2, …, n |
| 偶数序列 | `SEQUENCE(n, 1, 0, 2)` | 0, 2, 4, …, 2(n-1) |
| 逆序序列 | `SEQUENCE(n, 1, n, -1)` | n, n-1, …, 1 |
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
| 全零矩阵 | `SEQUENCE(rows, cols, 0, 0)` | 填充常量矩阵 |

### 命名函数库速查

在 **公式 → 名称管理器** 中定义 LAMBDA，团队共享、无需 VBA：

| 函数名 | 用途 | 核心技巧 | 详情 |
|--------|------|----------|------|
| `SAFE_DIV(a, b)` | 安全除法 | `IF(b=0, "", a/b)` | [→](formulas/10-named-functions.md) |
| `RUNNING_TOTAL(arr)` | 运行累计 | `SCAN` | [→](formulas/10-named-functions.md) |
| `GROUP_SUM(keys, vals)` | 分组求和 | `UNIQUE` + `MAP` + `FILTER` | [→](formulas/10-named-functions.md) |
| `GROUP_COUNT(keys)` | 分组计数 | `UNIQUE` + `MAP` | [→](formulas/10-named-functions.md) |
| `TOP_N(data, n, col)` | 前 N 行 | `SORT` + `TAKE` | [→](formulas/10-named-functions.md) |
| `REVERSE(arr)` | 逆序 | `SEQUENCE` + `INDEX` | [→](formulas/10-named-functions.md) |
| `PAGINATE(data, page, size)` | 分页取数 | `SEQUENCE` + `INDEX` | [→](formulas/10-named-functions.md) |
| `MOVING_AVG(arr, window)` | 移动平均 | `SEQUENCE` 滑动窗口 | [→](formulas/10-named-functions.md) |
| `UNPIVOT(rows, cols, vals)` | 逆透视 | `SEQUENCE` + 整除/取余 | [→](formulas/10-named-functions.md) |
| `PERCENTILE_RANK(arr)` | 百分位排名 | `MAP` + 向量化比较 | [→](formulas/10-named-functions.md) |

---

## 🔧 纯公式数据流水线

> PlantUML 源文件见 [`assets/data-pipeline.puml`](assets/data-pipeline.puml)。

Excel 纯公式可以构建完整的 ETL（提取-转换-加载）管道，典型模式如下：

```
📥 数据输入          🧹 清洗              🔄 转换              📊 聚合              📋 输出
A2:E200    ──→    FILTER·TRIM     ──→    MAP·BYROW      ──→    UNIQUE·REDUCE  ──→    SORT·HSTACK
                  IFERROR              LET                    SUMPRODUCT           TAKE
                  INDEX(,,col)                                MAP+FILTER
```

**完整示例**（区域销售汇总）：

```excel
=LET(
  raw,    A2:E200,                                          -- ① 输入
  valid,  FILTER(raw, INDEX(raw,,5) > 0),                   -- ② 清洗：去除无效行
  region, INDEX(valid,, 2),                                 -- ③ 提取字段
  amount, INDEX(valid,, 5),
  u,      UNIQUE(region),                                   -- ④ 聚合：按区域分组
  total,  MAP(u, LAMBDA(r, SUM(FILTER(amount, region=r)))),
  SORT(HSTACK(u, total), 2, -1)                             -- ⑤ 输出：降序排列
)
```

→ [更多流水线示例](formulas/09-advanced-pipelines.md)

---

## 💡 最佳实践

| # | 建议 | 理由 |
|:-:|------|------|
| 1 | **优先用 `LET`** 给中间结果命名 | 降低维护成本，便于调试 |
| 2 | **优先用 `SEQUENCE` 生成索引** | 替代辅助列和手写常量数组 |
| 3 | **优先用 `INDEX(data,, col)` 引用列** | 数据结构变化时只需改 data 引用 |
| 4 | **能用 SEQUENCE 广播就不用 MAKEARRAY** | 更简洁（如乘法表、棋盘格） |
| 5 | **用 `TAKE`/`DROP` 替代复杂 INDEX 切片** | 语义更清晰 |
| 6 | **用 `CHOOSECOLS`/`CHOOSEROWS` 选取** | 比 INDEX 多列写法更直观 |
| 7 | 先搭建"可读版本"，再压缩为短公式 | 团队协作优先可读性 |
| 8 | 大型数据减少重复计算（缓存到 `LET`） | 避免 O(n²) 性能陷阱 |
| 9 | 为关键 LAMBDA 建立命名函数 | 形成团队可复用模板 |
| 10 | 超过 10,000 行时用 `SUMPRODUCT` 替代 `MAP+FILTER` | 性能优化 |

---

## 🔍 如何选择函数

遇到问题时，按以下决策流程选择合适的函数：

```
需要什么？
├── 生成数据？
│   ├── 数字/日期序列 → SEQUENCE
│   └── 按行列规则生成 → MAKEARRAY（不能用 SEQUENCE 广播时）
│
├── 选取数据？
│   ├── 按行号/列号 → INDEX
│   ├── 按位置选列/行 → CHOOSECOLS / CHOOSEROWS
│   └── 截取头尾 → TAKE / DROP
│
├── 逐个处理？
│   ├── 逐元素 → MAP
│   ├── 按行聚合 → BYROW
│   └── 按列聚合 → BYCOL
│
├── 累计/聚合？
│   ├── 只要最终结果 → REDUCE
│   └── 要每步中间值 → SCAN
│
├── 筛选/查找？
│   ├── 条件筛选 → FILTER
│   ├── 精确/模糊查找 → XLOOKUP
│   └── 返回位置 → XMATCH
│
├── 排序/去重？
│   ├── 直接排序 → SORT
│   ├── 外部键排序 → SORTBY
│   └── 去重 → UNIQUE
│
├── 拼接/重塑？
│   ├── 上下合并 → VSTACK
│   ├── 左右合并 → HSTACK
│   ├── 展平 → TOCOL / TOROW
│   └── 重塑 → WRAPCOLS / WRAPROWS
│
└── 文本处理？
    ├── 拆分 → TEXTSPLIT
    ├── 合并 → TEXTJOIN
    └── 截取 → TEXTBEFORE / TEXTAFTER
```

---

## 📊 PlantUML 图示

本项目在 `assets/` 目录提供 PlantUML 源文件，可在线渲染或本地生成图片：

| 图示 | 文件 | 说明 |
|------|------|------|
| 函数分类思维导图 | [`function-taxonomy.puml`](assets/function-taxonomy.puml) | 8 大类函数的分类体系 |
| 学习路径图 | [`learning-path.puml`](assets/learning-path.puml) | 5 阶段推荐学习路线 |
| 数据流水线图 | [`data-pipeline.puml`](assets/data-pipeline.puml) | 纯公式 ETL 典型模式 |

**在线渲染**：将 `.puml` 文件内容粘贴到 [PlantUML Online Server](https://www.plantuml.com/plantuml/uml/) 即可预览。

**本地渲染**（需安装 Java + PlantUML）：

```bash
java -jar plantuml.jar assets/*.puml
```

---

## 📜 License

本项目采用 [MIT License](LICENSE)。

---

## 🤝 贡献

欢迎持续补充更多基于 **LAMBDA / SEQUENCE / INDEX / REDUCE / SCAN** 的实战示例！

**贡献建议**：
- 新增场景文件放入 `scenarios/basic/`（单函数）或 `scenarios/advanced/`（多函数组合）
- 每个示例文件保持 3–5 个由浅入深的示例
- 公式中使用 `LET` 命名中间变量，确保可读性
- 在文件头部标注所需函数和难度等级

把本仓库演进为完整的 **Excel 纯公式编程手册**！
