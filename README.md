# Word 文档信息修改器 v1.0

一个绿色小工具，用来修改 `.docx` 文件里的 **编辑总时间** 和 **作者名称**，
还可以直接把**另一个 Word 文档的作者**复制过来。

在另一台设备上使用：把整个文件夹拷过去（或拷 `DocxMetaEditor.zip` 解压），
双击 **`run.bat`** 即可。目标机器 **不需要安装 Python / Office / .NET**，
Windows 7 及以上自带的 PowerShell 就够了。

---

## 一、怎么用

1. 双击 `run.bat`，出现界面。
2. **① 选择文档** — 点「选择文档...」，选你要改的那个 `.docx`。
   下方会显示它当前的作者和编辑总时间，日志里会同时打印
   **文件系统创建时间**和**文档内容创建时间**两个值。
3. **② 编辑总时间** — 勾上「修改编辑总时间」，填小时 / 分钟。
   不勾就保持原值不动。Word 只精确到分钟（6 小时 = 360 分钟）。
4. **③ 作者名称** — 勾上「修改作者名称」，在输入框里填新作者。
   - 想改成和**另一个文档**一样：点「读取作者并填入...」，选那个文档，
     作者会自动填进输入框（同时自动勾选）。确认无误后点开始修改。
   - 「同时修改「上次保存者」」默认勾选。Word 的"作者"和"上次保存者"
     是两个独立字段，一般人看到的"作者"是前者；勾上就两个一起改。
5. **④ 创建时间** — 勾上「创建时间」，用日期选择器选一个时间点。
   - **默认只改文件系统时间戳** —— 也就是资源管理器「属性 → 详细信息」里
     显示的**「创建时间」**（以及「Date created」列）。这一项改完立刻能看到。
   - 想连文档内部的**「内容创建时间」**（`dcterms:created`）一起改，
     再勾上右边那个复选框。两者是**不同的字段**，详见第三节。
   - 选中文档时，选择器会自动填上该文件**当前的**创建时间，方便微调。
   - 时间按**本机时区**理解，内部按 UTC 存储 —— 和 Word 的做法一致。
6. **⑤ 保存方式** — 二选一：
   - **直接修改原文件**（默认）：改的就是你选的那个文件，同时会在同目录
     生成一个 `<原文件名>.backup.docx` 备份（同名备份已存在则不会覆盖）。
   - **另存为新文件**：原文件不动，另存一份。文件名留空则自动加 `_modified`。
7. 点 **「开始修改」**，会先弹一个确认框列出改什么、改到哪；确认后执行。
   完成后底部日志会打印**回读验证**结果（重新打开文件把值读出来核对）。

---

## 二、重要提醒：Word 是「累计」而不是「覆盖」

这条是**实测结论**，不是推测（环境：Microsoft 365 / Word 16.0，Office16 版）：

> **设 300 分钟 → 用 Word 打开、编辑一段时间 → 保存，回读值大于 300。**

也就是说，Word **不会**把你设定的数字抹掉重算，而是把它当作**基线**，
在这个基础上继续累加本次会话的编辑时间。编辑总时间的语义本来就是
「这份文档生命周期内的累计编辑时长」，而不是「最后一次打开用了多久」。

所以：

- 你设的 5 小时 + 之后编辑 10 分钟 → 保存后大约是 **5 小时 10 分**。
- **`app.xml` 里没有 `<TotalTime>` 字段时**，Word 会从 0 开始累加。
- **设一个比原值更小的数字**（先把 5 小时改成 10 分钟）可以当作「归零重置」
  的手段 —— 之后从 10 分钟重新往上加，这是给这份文档"洗"一个合理编辑时长的
  唯一可靠做法。
- Word 只统计「文档处于活动编辑状态」的时间，开着窗口不动不会等比例累加。

### 那么改完之后还要不要避免用 Word 保存？

**不用刻意避免了**（这一点和我早先的说法不同，早先说「会被覆盖」是不准确的）。
正常编辑、正常保存都不会毁掉你设的值，只是会在它之上继续累加。

唯一要留意的是：你设的是**基线**，不是**终值**。如果你要的是一个精确的最终数字
（比如交货时必须是 6 小时 00 分），那就**最后一步再改** —— 先改完所有内容，
再用这个工具把时间定死，之后别再让 Word 保存。

### 其它

- 修改前请**关闭 Word**（文件被占用时会提示失败，这是正常的）。
- 只改 zip 内部对应的那一两个 XML 条目，文档正文、样式、图片等一个字节都不动。
- 作者可以填中文，测试覆盖了中文往返。
- 作者字段的行为**和编辑总时间不同**：Word 保存时不会在旧作者上累加，
  它是照当前 Word 的用户名写入的（`dc:creator` 与 `cp:lastModifiedBy`），
  所以很可能被之后的一次 Word 保存覆盖掉。
  **这一条是推断，我没像编辑总时间那样实测过** —— 如果你在意，可以按上面
  同样的方法验一次（设一个假作者 → Word 打开保存 → 回读作者）。
  在验证之前，稳妥做法是一样的：作者也放在**最后一步**改。

---

## 三、「创建时间」有两个，别搞混

这一点很容易踩坑，所以单独说。资源管理器「属性 → 详细信息」里同时有
好几项时间，**来源完全不同**：

| 显示项 | 数据来源 | 要用什么改 |
|---|---|---|
| **创建时间** / Date created | **文件系统**时间戳 | 这个工具的「创建时间」（默认那项） |
| Date modified | 文件系统 LastWriteTime | 另存文件/写完自动更新 |
| Date accessed | 文件系统 LastAccessTime | 未开放（一般也用不到） |
| **内容创建时间** / Content created | 文档内 `dcterms:created` | 「一并修改文档内部的『内容创建时间』」 |
| **内容修改时间** / Date last saved | 文档内 `dcterms:modified` | 未开放（用 `Set-DocxMeta -ContentModified`） |

**所以：只改 docx 内部元数据，资源管理器里的「创建时间」是不会有任何变化的。**
这就是为什么这个工具的默认行为是改**文件系统**那一项 —— 用户看的就是它。

反过来，如果你把文件发给别人、或者拷到别的机器上，**文件系统时间戳可能不被保留**
（取决于拷贝方式），而**文档内部的 `dcterms:created` 会跟着文件走**。
要「看起来像是某天创建的」，通常两个都改比较稳妥。

### 时区

Word 用 **UTC** 存储这些时间（例如 `2026-09-22T02:30:00Z`），显示时再换算成
查看者所在时区。本机是 UTC+8 时，上面这个值显示为 `10:30`。
这个工具按**本机时区**理解你的输入，存储时转成 UTC，所以回读出来和你输入的
本地时间完全一致（测试里覆盖了这一点）。

---

## 四、文件夹里都是什么

| 文件 | 作用 |
|---|---|
| `run.bat` | **双击这个启动**（纯 ASCII + CRLF，怎么拷都不会坏） |
| `DocxMetaGui.ps1` | 界面（WinForms） |
| `DocxMeta.ps1` | 核心逻辑：读写 docx 的 `docProps/core.xml`、`docProps/app.xml`，以及文件系统时间戳 |
| `README.md` | 本文件 |
| `Make-TestDocx.ps1` | 现场生成一个合法的测试用 `.docx`（测试自给自足，不依赖任何本机文件） |
| `_test-core.ps1` | 逻辑层测试（45 项断言，含创建时间/时区/文件戳） |
| `_test-gui.ps1` | 界面层测试，进程内驱动真实按钮（65 项断言，含创建时间流程） |
| `_test-all.bat` | 双击跑全部测试（在那台设备上验证环境是否正常） |
| `_gui-smoke.ps1` | 真正启动一次界面、截图、再关掉（验证窗口能正常弹出） |

`_` 开头的东西和 `Make-TestDocx.ps1` 只用于自检，删掉不影响工具运行。
测试全部通过时输出末尾应该是 `=== FAILURES: 0 ===`。

### 如果那台设备上工具打不开

先双击 `_test-all.bat`：

- 两个套件都输出 `FAILURES: 0` → 环境没问题，说明是启动方式的问题
  （确认是双击 `run.bat`，而不是直接双击 `.ps1`；`.ps1` 默认会被
  执行策略拦住，`run.bat` 里已经带了 `-ExecutionPolicy Bypass`）。
- 报「无法加载文件，因为在此系统上禁止运行脚本」→ 说明有人直接运行了
  `.ps1`。改用 `run.bat` 即可。
- 中文变成乱码 → `.ps1` 的 UTF-8 BOM 在传输中丢了。
  `DocxMetaGui.ps1` 启动时会自动补回，正常情况下不用管。


---

## 五、命令行用法（可选）

不想开界面时，`DocxMeta.ps1` 也能直接在 PowerShell 里调：

```powershell
# 载入
. .\DocxMeta.ps1

# 读信息
Get-DocxMeta -Path 'D:\a.docx'

# 把编辑总时间改成 6 小时（原地改，自动备份）
Set-DocxMeta -Path 'D:\a.docx' -TotalMinutes 360

# 改作者 + 时间
Set-DocxMeta -Path 'D:\a.docx' -Author '张三' -ModifiedBy '张三' -TotalMinutes 200

# 只改作者，另存为新文件（原文件不动）
Set-DocxMeta -Path 'D:\a.docx' -Author '李四' -OutputPath 'D:\b.docx'

# 复制另一个文档的作者
$donor = Get-DocxMeta -Path 'D:\来源.docx'
Set-DocxMeta -Path 'D:\a.docx' -Author $donor.Author -ModifiedBy $donor.LastModifiedBy

# 清空作者
Set-DocxMeta -Path 'D:\a.docx' -Author ''

# 改文档内部的「内容创建时间」（dcterms:created，UTC 存储）
Set-DocxMeta -Path 'D:\a.docx' -Created '2015-06-07 08:09:00'

# 连「内容修改时间」（dcterms:modified）一起改
Set-DocxMeta -Path 'D:\a.docx' -Created '2015-06-07 08:09:00' -ContentModified '2015-06-08 10:00:00'

# 改文件系统时间戳（资源管理器里显示的「创建时间」就是这一项）
Set-FileTimestamps -Path 'D:\a.docx' -CreationTime '2015-06-07 08:09:00'

# 文件时间戳 + 文档内部时间一起改（界面上就是两个复选框都勾）
Set-DocxMeta -Path 'D:\a.docx' -Created '2015-06-07 08:09:00'
Set-FileTimestamps -Path 'D:\a.docx' -CreationTime '2015-06-07 08:09:00'
```

> 注意顺序：`Set-DocxMeta` 会重写文件，从而**重置文件系统时间戳**。
> 两个都改时，`Set-FileTimestamps` 必须放在**后面**（界面里已经这样处理了）。

## 六、技术上踩过的 7 个坑（改代码时注意）

1. **编码**：Windows PowerShell 5.1 只认 UTF-8 BOM，没有 BOM 的 `.ps1` 会被
   当成 GBK 读，中文变乱码甚至直接语法报错。所以含中文的 `.ps1` 必须保存为
   **UTF-8 with BOM**（`.bat` 则相反，必须纯 ASCII + CRLF）。
   `DocxMetaGui.ps1` 启动时会自检并补回 BOM。
2. **调用形式**：PowerShell 里调用函数必须写 `Fn 'a' 'b'`，
   **不能写 `Fn('a', 'b')`** —— 后者会被当成 .NET 方法调用，
   带类型约束的参数会报 `Cannot convert value to type System.String`。
3. **别用 `XmlDocument.Save(StringWriter)`**：它会把 `encoding="utf-16"`
   写进声明，而字节是 UTF-8，生成的文件严格解析器直接拒收。
   本项目的做法是全程按**字节**处理（`MemoryStream` + `XmlWriter`）。
4. **测试的弹窗拦截不能用"后定义同名函数"**：事件处理器解析函数时用的是
   **定义它的那个作用域**，在外面覆盖同名函数无效。所以界面里留了
   `$script:UiMessageHook` / `$script:UiConfirmHook` 两个显式钩子。
5. **XML 元素查找必须看命名空间，不能只看 local-name**（这个坑最深）：
   - `core.xml` 里 `dcterms:created` 和 `cp:created` **同名不同命名空间**。
     早期版本按 `local-name()` 查找，结果改错了元素、还新建了一个
     `<cp:created>`，同时 `xmlns:xsi` 被解析器丢掉。
   - `app.xml` 又是**第三个**命名空间。把查找限定在 core.xml 的命名空间里，
     `TotalTime` 就永远找不到，于是又追加了一个重复的 `<TotalTime>`。
   现在的做法：**先按 local-name 找到元素、用它自己的命名空间**；
   只有新建元素时才用优先表 + 根元素上已声明的前缀。
   读取时对会冲突的字段显式指定命名空间，并顺手清理历史遗留的重复元素。
6. **`SetAttribute("xsi:type", ...)` 会报"重复属性"**：它被解析成
   前缀 `xsi` + 本地名 `type`，和 `xmlns:xsi` 声明撞车。改用
   `CreateAttribute(name, uri)` + `SetAttributeNode`。
7. **`List.Add()` 的返回值会污染输出流**：`$list.Add(x)` 返回索引号，
   不接住的话函数会返回一个数组。本项目统一写 `$null = $list.Add(...)`。

---

## 七、传到 GitHub

直接传**可以**，但别用「拖拽上传网页」那种方式——它会把行尾规范化、
可能动 BOM，正好破坏这个项目最要紧的东西。用 git 命令行：

```bash
cd docx-meta-editor
git init -b main
git add -A
git commit -m "Word 文档信息修改器：改编辑总时间与作者，支持从另一文档复制作者"
git remote add origin https://github.com/<你的用户名>/<仓库名>.git
git push -u origin main
```

首次 push 前先在 `git status` 里确认没有多余文件被带进去。

### 仓库里的 `.gitattributes` 为什么写得那么怪

`*.ps1 binary` 看着像写错了，其实是刻意的：`.ps1` 里的 **UTF-8 BOM 是功能性的**
（见上面第 1 条坑）。任何行尾/编码自动转换都会把 BOM 弄丢，然后中文界面在
PowerShell 5.1 上就废了。标成 `binary` 让 git 原样存取字节，从根上避免
`core.autocrlf` 之类的意外。`.bat` 则固定 `eol=crlf`，因为 cmd.exe 按字节偏移
逐行解析，LF-only 的 `.bat` 在某些情况下会被错误解析。**这两条规则不要"顺手简化"。**

### 已经给你配好的东西

- `.gitignore` —— 排除测试临时目录、`_gui.png`、打包产物，以及
  `*.backup.docx` / `*_modified.docx`（万一在仓库目录里试跑工具，
  不会把真实文档连同里面的**作者姓名**一起提交上去）
- `.github/workflows/tests.yml` —— push / PR 时在 `windows-latest` 上：
  先校验 BOM 与 CRLF 有没有在 checkout 时丢掉，再跑两个测试套件。
  步骤刻意用 `shell: powershell`（5.1）而不是 `pwsh`（7.x），
  因为工具就是针对 5.1 发布的，测试套件验证的也是 5.1 的行为。
  两个测试脚本失败时会 `exit 1`，CI 能真正感知失败。
- `run.bat` / `_test-all.bat` 保持纯 ASCII —— 这样**即使**仓库被
  GitHub 网页编辑器改坏、或者别的工具弄乱了编码，启动器本身依然能用。

### 传之前建议补两样

1. **LICENSE**。没有许可证的仓库默认是「保留所有权利」，
   别人不能合法复用。要想让人随便用，加个 MIT：在 GitHub 上
   「Add file → Create new file」输入 `LICENSE`，它会提示选模板。
2. **如果想让别人免装环境直接下载用**，可以在 Releases 里附上
   `DocxMetaEditor.zip`。构建方式就是把这个文件夹里除 `.git*`、`_test*`、
   `_gui-smoke.ps1` 之外的文件压成一个以 `DocxMetaEditor/` 为顶层目录的 zip。

### 一个必须提醒的隐私点

**别把任何真实 `.docx`（或它的截图）提交到公开仓库。**
`docProps/core.xml` 里有作者名和「上次保存者」，`app.xml` 里有编辑总时间、
公司名，`word/document.xml` 里有全文。用这个工具改过的文档同样带着这些字段。
`.gitignore` 已经挡掉了常见的备份/改后文件名，但它挡不住你手动 `git add -f`。

