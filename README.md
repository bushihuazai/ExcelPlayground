# ExcelPlayground

一个总结和示意 **Excel 基于纯公式的“可编程技巧库”**。

> 适用版本：Microsoft 365 / Excel 2021+（支持动态数组与 LAMBDA 系列函数）

---

## 1. 核心理念：把 Excel 当成函数式编程环境

- **LAMBDA**：定义可复用函数
- **LET**：定义局部变量，提升可读性和性能
- **MAP / BYROW / BYCOL**：逐元素或按行/列映射
- **REDUCE**：累计折叠（聚合）
- **SCAN**：返回每一步累计过程
- **MAKEARRAY**：按规则生成数组
- **VSTACK / HSTACK**：纵向/横向拼表
- **FILTER / SORT / UNIQUE / XLOOKUP**：声明式数据处理

---

## 2. 从入门到进阶：常用纯公式示例

> 下面示例可直接粘贴到单元格中（根据你的区域设置调整参数分隔符）。

### 2.1 LAMBDA + LET：封装可读函数

**示例：计算折后价（含最小为 0 保护）**

```excel
=LAMBDA(price, discount, LET(result, price*(1-discount), MAX(0, result)))(A2, B2)
```

**示例：定义命名函数 `SAFE_DIV`（名称管理器中）**

```excel
=LAMBDA(a,b,IF(b=0,"",a/b))
```

使用：

```excel
=SAFE_DIV(A2,B2)
```

---

### 2.2 MAP：对数组逐元素处理

**示例：批量转大写并去除前后空格**

```excel
=MAP(A2:A10, LAMBDA(x, UPPER(TRIM(x))))
```

**示例：两个数组对应相加**

```excel
=MAP(A2:A10, B2:B10, LAMBDA(x,y, x+y))
```

---

### 2.3 BYROW / BYCOL：按行列计算

**示例：按行求和（二维区域每行合计）**

```excel
=BYROW(B2:F10, LAMBDA(r, SUM(r)))
```

**示例：按列取最大值**

```excel
=BYCOL(B2:F10, LAMBDA(c, MAX(c)))
```

---

### 2.4 REDUCE：把数组折叠成结果

**示例：累计求和**

```excel
=REDUCE(0, A2:A10, LAMBDA(acc, x, acc+x))
```

**示例：累计连乘（几何增长）**

```excel
=REDUCE(1, A2:A10, LAMBDA(acc, x, acc*x))
```

**示例：拼接字符串清单**

```excel
=REDUCE("", A2:A10, LAMBDA(acc, x, IF(acc="", x, acc&"、"&x)))
```

---

### 2.5 SCAN：输出每一步中间值（过程可视化）

**示例：运行累计和（Running Total）**

```excel
=SCAN(0, A2:A10, LAMBDA(acc, x, acc+x))
```

**示例：账户余额轨迹（初始 1000）**

```excel
=SCAN(1000, B2:B20, LAMBDA(balance, cashflow, balance+cashflow))
```

---

### 2.6 MAKEARRAY：规则生成数据

**示例：生成 5x5 乘法表**

```excel
=MAKEARRAY(5,5, LAMBDA(r,c, r*c))
```

**示例：生成棋盘格 0/1**

```excel
=MAKEARRAY(8,8, LAMBDA(r,c, MOD(r+c,2)))
```

---

### 2.7 VSTACK / HSTACK：拼接多表

**示例：上下合并两块数据**

```excel
=VSTACK(A2:C10, A15:C20)
```

**示例：左右合并主表与维度表（演示）**

```excel
=HSTACK(A2:C10, E2:G10)
```

**示例：为不同来源补齐列后再纵向汇总**

```excel
=LET(
  t1, A2:C10,
  t2, H2:I8,
  t2_fix, HSTACK(t2, MAKEARRAY(ROWS(t2),1,LAMBDA(r,c,""))),
  VSTACK(t1, t2_fix)
)
```

---

### 2.8 FILTER / SORT / UNIQUE：声明式数据清洗

**示例：筛选华东区且销售额>10000**

```excel
=FILTER(A2:E100, (B2:B100="华东")*(E2:E100>10000), "无结果")
```

**示例：去重后排序客户列表**

```excel
=SORT(UNIQUE(C2:C100))
```

---

### 2.9 TEXTSPLIT / TEXTJOIN：文本结构化处理

**示例：把 “省-市-区” 拆分为三列**

```excel
=TEXTSPLIT(A2, "-")
```

**示例：把标签列合并为单字符串**

```excel
=TEXTJOIN(",", TRUE, FILTER(B2:B100, B2:B100<>""))
```

---

## 3. 组合技：一个“纯公式小流水线”示意

目标：把原始明细做成“区域销售汇总表”。

```excel
=LET(
  raw, A2:E200,
  valid, FILTER(raw, INDEX(raw,,5)>0),
  region, INDEX(valid,,2),
  amount, INDEX(valid,,5),
  u, UNIQUE(region),
  total, MAP(u, LAMBDA(r, SUM(FILTER(amount, region=r)))),
  SORT(HSTACK(u,total), 2, -1)
)
```

说明：
1. `FILTER` 先清掉无效数据
2. `UNIQUE` 找到区域列表
3. `MAP + FILTER + SUM` 对每个区域聚合
4. `HSTACK + SORT` 输出排序后的结果表

---

## 4. 命名函数建议（便于复用）

可在“公式 → 名称管理器”中沉淀函数库：

- `SAFE_DIV(a,b)`：安全除法
- `RUNNING_TOTAL(arr)`：运行累计（基于 `SCAN`）
- `GROUP_SUM(keys, vals)`：分组求和（基于 `UNIQUE + MAP`）
- `STACK_FIX(table, cols)`：补列后拼接（基于 `MAKEARRAY + HSTACK`）

---

## 5. 使用建议

- 优先用 `LET` 给中间结果命名，降低维护成本
- 先搭建“可读版本”，再考虑压缩为短公式
- 对大型数据尽量减少重复计算（把重复表达式缓存到 `LET`）
- 为关键 LAMBDA 建立命名函数，形成团队可复用模板

---

## 6. 场景示例集（scenarios/）

以下 15 个场景文件，每个包含 3–5 个从易到难的示例公式及详细说明：

| 编号 | 场景 | 核心函数 |
|------|------|----------|
| [01](scenarios/[01]LAMBDA基础.md) | LAMBDA 基础 | LAMBDA、递归 |
| [02](scenarios/[02]LET局部变量.md) | LET 局部变量 | LET |
| [03](scenarios/[03]MAP逐元素处理.md) | MAP 逐元素处理 | MAP |
| [04](scenarios/[04]BYROW按行计算.md) | BYROW 按行计算 | BYROW |
| [05](scenarios/[05]BYCOL按列计算.md) | BYCOL 按列计算 | BYCOL |
| [06](scenarios/[06]REDUCE累计折叠.md) | REDUCE 累计折叠 | REDUCE |
| [07](scenarios/[07]SCAN过程可视化.md) | SCAN 过程可视化 | SCAN |
| [08](scenarios/[08]MAKEARRAY规则生成.md) | MAKEARRAY 规则生成 | MAKEARRAY |
| [09](scenarios/[09]VSTACK纵向拼接.md) | VSTACK 纵向拼接 | VSTACK |
| [10](scenarios/[10]HSTACK横向拼接.md) | HSTACK 横向拼接 | HSTACK |
| [11](scenarios/[11]FILTER条件筛选.md) | FILTER 条件筛选 | FILTER |
| [12](scenarios/[12]SORT与SORTBY排序.md) | SORT 与 SORTBY 排序 | SORT、SORTBY |
| [13](scenarios/[13]UNIQUE去重提取.md) | UNIQUE 去重提取 | UNIQUE |
| [14](scenarios/[14]XLOOKUP高级查找.md) | XLOOKUP 高级查找 | XLOOKUP |
| [15](scenarios/[15]综合实战流水线.md) | 综合实战流水线 | 多函数组合 |

---

欢迎持续补充更多基于 **LAMBDA / REDUCE / SCAN / VSTACK** 的实战示例，把本仓库演进为完整 Excel 纯公式编程手册。
