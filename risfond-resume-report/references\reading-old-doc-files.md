# 旧版 .doc 文件读取

## 问题说明
用户提供的候选人简历通常是旧版 .doc 格式（非 .docx），python-docx 无法直接读取。

## 读取方法

### 方法1：Word COM 转换法（推荐）
```powershell
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open("原始文件.doc")
$doc.SaveAs([ref]"目标文件.docx", [ref]16)  # 16 = wdFormatXMLDocument
$doc.Close()
$word.Quit()
```
转换后用 python-docx 读取。

### 方法2：olefile 原始提取（备用）
```python
import olefile
ole = olefile.OleFileIO(path)
data = ole.openstream('WordDocument').read()
# 提取ASCII字符
text = ''
for i in range(0, len(data)-1, 2):
    c = data[i]
    if 32 <= c <= 126:
        text += chr(c)
    elif c in (10, 13):
        text += chr(c)
ole.close()
```
注意：olefile 提取结果含大量乱码，只能捞出片段信息。

## 已知问题
- Word COM 转换有时成功但转换后的 .docx 全是空段落（text为空）。这种情况下只能依赖用户手动提供的信息。
- RPC错误 `0x800706BE` 或 `0x800706BA`：Word进程卡住。先杀进程再重试：
  ```powershell
  Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
  ```

## 当简历文件读不出时的应对策略
- **公司名称**：如果第一个公司的名称无法识别，使用"上海某XX公司"这种描述性名称
- **工作职责**：从用户提供的技能列表中反向推导
- **项目经历**：如果无法提取，直接省略
- **年龄**：用户没提供就不写，不要根据入学年份推算
