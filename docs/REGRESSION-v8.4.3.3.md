# v8.4.3.3 回归测试报告

**结论：ALL_PASS**

## 语法

- 主 Userscript：PASS
- 预览 iframe JavaScript：PASS

## 新增“大表格安全拆页”

- 识别真实 `<tr>` 行边界：PASS
- 优先选择当前页最后可完整容纳的行：PASS
- `rowspan` 穿越边界检测：PASS
- `rowspan=0`（跨到行组结束）处理：PASS
- `colspan` 原样保留：通过结构设计，不参与纵向禁切判断
- 找不到安全行边界时回退整表换页：PASS
- 表格前/后正文分别保留：PASS（结构逻辑检查）
- 续页 `caption` 不重复：PASS
- 空 `thead/tbody/tfoot` 清理：PASS
- JavaScript 模拟 `rowspan` 场景：`TABLE_LOGIC_PASS`

## 8.4.3.2 冻结核心

以下函数 SHA-256 与 8.4.3.2 完全一致：

- `buildAnswerImageRowModel`
- `scanLocalSafeSourceCuts`
- `chooseSmartAnswerCut`
- `renderAnswerImageFlow`
- `renderPages`
- `updatePageNumbers`
- `findSectionNodeForQuestion`
- `getPrintableSourceNodes`

只有以下两个表格相关函数发生预期变化：

- `findOverflowBoundary`
- `splitQuestionAt`

## 历史功能

- 答案行级连续切页：PASS
- 复杂答案区域重叠兜底：PASS
- 重叠保护开关：PASS
- 原题号：PASS
- 默认分段页码：PASS
- 标准精排 14px：PASS
- 题型编辑文字清洗：PASS

## 浏览器端说明

尝试用当前容器的 headless Chromium 跑 DOM 端到端表格测试，但 Chromium 因当前环境 DBus/进程退出问题超时，因此没有把该项计为“通过”。本报告只声明实际完成的 Node 语法、函数哈希和 JavaScript 逻辑测试。

## 自动检查输出摘要

```text
ALL_PASS
PASS 主 Userscript 语法
PASS 找到预览 iframe 脚本
PASS 预览 iframe JavaScript 语法
PASS 版本号8.4.3.3
PASS 存在表格安全边界函数
PASS 存在rowspan穿越检测
PASS 只在tr边界拆表
PASS 从页底向上找最后安全行
PASS rowspan=0按行组到底处理
PASS 无安全边界回退整表换页
PASS 表格拆分保留前后内容
PASS 续页不重复caption
PASS 空rowgroup清理
PASS 表格拆分标记存在
PASS 答案行模型仍在
PASS 复杂答案块仍用重叠兜底
PASS 重叠保护开关仍在
PASS 原题号仍保留
PASS 默认分段页码仍在
PASS 标准精排14px
PASS 题型编辑文字清洗仍在
PASS 表格边界真实JavaScript逻辑测试 :: TABLE_LOGIC_PASS
```

完整函数哈希见 `docs/CORE-HASHES-v8.4.3.3.txt`。
