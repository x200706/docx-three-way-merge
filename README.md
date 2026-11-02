# docx-three-way-merge

一個 Claude 技能：改錯 Word 檔時，把修改連同註解同步到正確的檔案，遇到衝突會先停下來問你。

A Claude skill for Word files: carry your edits and comments over to the right file, and stop to ask when both sides changed the same words.

## 什麼時候用

你在舊版或錯的檔案上改了內容，現在要把這些修改搬到正確的檔案。準備三個檔案：

| 檔案 | 意思 |
|---|---|
| `BASE` | 改之前的錯誤檔案 |
| `MODIFIED` | 改之後的錯誤檔案 |
| `TARGET` | 正確的檔案 |

只有 `BASE` 到 `MODIFIED` 的差異會被搬過去，`TARGET` 裡其他比較新的內容都會留著。

## 安裝

把技能資料夾複製到 Claude Code 的技能目錄，並安裝 `lxml`：

```bash
git clone https://github.com/x200706/docx-three-way-merge.git
cp -r docx-three-way-merge/skills/docx-three-way-merge ~/.claude/skills/
pip install lxml
```

之後直接跟 Claude 說「我改錯檔案了，幫我把差異同步到正確的母檔」就可以。

## 直接用腳本

```bash
# 先檢查，不會寫任何檔案；有衝突時結束代碼是 2
python scripts/docx_merge.py check BASE.docx MODIFIED.docx TARGET.docx --report report.json

# 決定好衝突後再合併
python scripts/docx_merge.py merge BASE.docx MODIFIED.docx TARGET.docx -o OUT.docx --resolutions resolutions.json
```

`resolutions.json` 用報告裡的衝突編號記錄每個決定：

```json
{
  "P12-1": "use_modified",
  "P12-2": "keep_target",
  "P40-1": {"text": "我要的最終文字"}
}
```

## 會同步什麼

| 內容 | 處理方式 |
|---|---|
| 內文文字的增刪修改 | 同步，並保留 `TARGET` 原本的字體 |
| 註解的新增、刪除與標記範圍 | 同步 |
| 新增或刪除整段文字 | 同步 |
| 同一處被兩邊改成不同內容 | 列為衝突，等你決定 |
| 新的圖片、公式 | 列為手動項目，不會寫出檔案 |
| 格式、表格結構、頁首頁尾、註腳 | 只回報，不同步 |

## 安全設計

- 有未解決的衝突或手動項目時，不會產生檔案。
- 不會覆蓋三個輸入檔；輸出檔已存在時要加 `--force` 才會取代。
- 寫出前會重新讀取結果；沒動到的段落若有變化，或註解標記對不上，就放棄輸出。

## 測試

```bash
python -m unittest discover tests
```

測試用程式產生的小型 Word 檔，涵蓋乾淨合併、三種衝突解法、兩邊改法相同、圖片、段落增刪和註解。

開發時也用一份約 2800 段的真實文件比對過手動合併的結果，文字和註解標記位置完全一致。那份文件沒有放進這個 repo。

## 限制

- 段落對應靠文字相似度。`TARGET` 改寫太多時，會回報找不到對應段落，由你處理。
- 只比對內文和註解的內容，合併後請自行檢查輸出結果。
- 尚未用真實的增刪段落和多段落註解跑過大型文件，這部分目前只有小型測試。

## License

MIT
