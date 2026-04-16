# 12 — SORT 与 SORTBY 排序：灵活的声明式排序

> **适用版本**：Microsoft 365 / Excel 2021+

## 简介

- `SORT(array, [sort_index], [sort_order], [by_col])` — 按数组内某列排序
- `SORTBY(array, by_array1, [sort_order1], [by_array2, ...])` — 按**外部数组**排序

两者都返回动态数组，不改变原始数据。

---

## 示例 1 ⭐ — 按单列升序排序

```excel
=SORT(A2:D50, 3, 1)
```

**说明**：按第 3 列升序（`1`）排列。第三参数 `-1` 为降序。

---

## 示例 2 ⭐⭐ — 按多列排序

```excel
=SORT(SORT(A2:E100, 5, -1), 2, 1)
```

**说明**：先按第 5 列降序排，再按第 2 列升序排。由于 SORT 是稳定排序，内层排序的相对顺序在外层相同键值时保留。

> 💡 等价于 SQL：`ORDER BY col2 ASC, col5 DESC`

---

## 示例 3 ⭐⭐⭐ — SORTBY：按外部列排序

```excel
=SORTBY(A2:D50, E2:E50, -1)
```

**说明**：按 E 列的值降序排列 A:D 的数据。E 列不必在 A:D 范围内，这就是 SORTBY 比 SORT 灵活的地方。

---

## 示例 4 ⭐⭐⭐⭐ — SORTBY 多条件 + 自定义顺序

按部门自定义顺序排列（技术部 → 产品部 → 运营部），同部门内按工资降序：

```excel
=LET(
  data,  A2:E50,
  dept,  INDEX(data,,2),
  salary, INDEX(data,,5),
  order, MAP(dept, LAMBDA(d,
    SWITCH(d, "技术部",1, "产品部",2, "运营部",3, 99)
  )),
  SORTBY(data, order, 1, salary, -1)
)
```

**说明**：

1. `MAP + SWITCH` 将部门名称映射为自定义排序编号
2. `SORTBY` 先按 `order` 升序，再按 `salary` 降序
3. 未匹配的部门排在最后（编号 99）

---

## 示例 5 ⭐⭐⭐⭐⭐ — 随机打乱顺序（洗牌）

```excel
=SORTBY(A2:D50, RANDARRAY(ROWS(A2:D50)), 1)
```

**说明**：`RANDARRAY` 为每行生成一个随机数，`SORTBY` 按随机数排列，效果相当于"洗牌"。

> ⚠️ 每次工作表重算时顺序会变化。如需固定结果，可将结果**粘贴为值**。

---

## 应用注意点

| 要点 | 说明 |
|------|------|
| SORT vs SORTBY | SORT 按自身列排序；SORTBY 按任意外部列排序，更灵活 |
| 稳定排序 | Excel 的 SORT/SORTBY 是稳定排序，相同键值的行保持原始顺序 |
| 多列排序 | SORTBY 直接支持多键：`SORTBY(data, col1, 1, col2, -1)` |
| 与 FILTER 搭配 | `SORT(FILTER(...))` 是声明式数据处理的经典组合 |
| 自定义顺序 | 用 MAP + SWITCH 生成辅助排序列，再交给 SORTBY |
