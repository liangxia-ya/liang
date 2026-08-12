# v8.4.3.4 回归测试

测试结果：**ALL_PASS**

## 新增功能检查

- PASS `metadata_version`
- PASS `preview_dot_css`
- PASS `export_dot_css`
- PASS `no_fake_dot_character_insertion`
- PASS `source_dot_tag_preserved_path`

## 语法检查

- PASS `node_check_main`
- PASS `node_check_iframe`

## 冻结核心函数

以下函数与 v8.4.3.3 完全一致：

- `findSectionNodeForQuestion`
- `getPrintableSourceNodes`
- `tableBoundaryCrossesRowspan`
- `findSafeTableSplitBoundary`
- `findOverflowBoundary`
- `scanLocalSafeSourceCuts`
- `chooseSmartAnswerCut`
- `renderAnswerImageFlow`
- `splitQuestionAtTableRow`
- `splitQuestionAt`
- `updatePageNumbers`
- `renderPages`

主 Userscript SHA-256：

`002eca39d203206ccc4ad2f321052019e9f6cb7c74db19100eeddee071b2b671`

说明：headless Chromium 在当前开发环境中仍因 DBus / 进程退出问题超时，因此浏览器端到端渲染测试不计为通过；语法、静态逻辑和核心函数冻结检查均已通过。
