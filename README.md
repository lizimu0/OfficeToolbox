# Office 自动化工具箱

基于 Python + PySide6 的桌面办公自动化工具,批量处理 Word / Excel / PPT。

## 功能一览

| 标签页 | 功能 |
| --- | --- |
| 文本提取 | 从 docx/xlsx/pptx 批量提取文字(每文件一个 txt 或全部合并),或导出所有嵌入图片 |
| 批量替换 | 在多个 docx/xlsx/pptx 中按规则查找替换,支持正则表达式(分组引用),默认替换前自动备份原文件 |
| Excel 合并 | 多个 xlsx 合并为一个:每个文件一个工作表,或按行追加(表头相同) |
| 模板生成 | Word 模板 + Excel 数据源,每行数据生成一份文档(合同、通知书等),支持 `{{图片:字段}}` 插入图片 |
| 格式转换 | docx/xlsx/pptx → PDF(调用本机 Office,保真),docx → txt,csv → xlsx |
| 合并拆分 | PDF 合并、PDF 按每 N 页拆分、Word 多文档合并 |
| 批量重命名 | 文件名文本替换、前缀+序号、Excel 映射表三种规则 |
| 使用帮助 | 程序内置的各功能使用说明 |

所有文件列表都支持拖入文件或整个文件夹(自动递归扫描 Office 文件);常用路径会自动记忆,下次打开自动回填;任务完成后弹出系统通知;任务执行中可随时**取消**,单个文件出错不中断整批任务;菜单栏「视图」可切换**浅色/深色主题**并自动记住。

## 启动方式

**双击即可,无需任何命令:**

1. 推荐:双击 `dist\Office自动化工具箱\Office自动化工具箱.exe`(已打包,不依赖 Python 环境)
2. 备选:双击项目根目录的 `启动工具箱.bat`(需本机已装 Python 和依赖)

> 格式转换(PDF)需要本机安装 Microsoft Office;其余功能无需 Office。

## 使用说明

### 模板批量生成
1. 在 Word 模板中用 `{{字段名}}` 写占位符,例如:`兹聘用 {{姓名}} 担任 {{岗位}}`
2. 准备 Excel 数据源,首行为表头,字段名与占位符一致
3. 选择模板、数据源、输出目录,可选择某个字段作为生成文件的文件名
4. 需要插入图片(如员工照片、签章)时,在模板中写 `{{图片:字段名}}`,并在数据源该字段填本地图片路径,过宽图片自动等比缩小

### 批量替换
- 替换结果**直接保存到原文件**,默认先备份到指定目录(每次执行生成时间戳子文件夹)
- 支持多条规则同时执行,按表格行顺序依次应用
- 勾选「使用正则表达式」后,查找内容按正则解释,替换为中可用 `\1` 引用分组,例如把 `(\d+)年` 替换为 `\1年度`;无效正则会在执行时报错提示。正则始终按**整段文本**匹配(保证 `^` `$` `\b` 等锚点语义正确),命中的段落格式统一为该段首字符样式

### 格式转换
- 输出文件保存在源文件同目录,同名文件会被覆盖
- 不支持的转换组合会自动跳过并在日志中说明

## 项目结构

```
office_toolbox/
├── main.py              # 程序入口
├── resources/app.ico    # 程序图标(make_icon.py 生成)
├── core/                # 核心处理逻辑(与界面无关,可单独调用)
│   ├── extract.py       # 文本/图片提取
│   ├── replace.py       # 批量替换
│   ├── merge_excel.py   # Excel 合并
│   ├── template.py      # 模板填充生成
│   ├── convert.py       # 格式转换(COM 接口)
│   ├── pdf_tools.py     # PDF 合并拆分 / Word 合并
│   └── rename.py        # 批量重命名
└── gui/
    ├── widgets.py       # 公共组件(文件列表/日志/后台线程/通知)
    ├── styles.py        # 全局样式主题
    ├── tabs.py          # 各功能标签页
    └── main_window.py   # 主窗口(含设置记忆)
```

## 测试

```powershell
python test_core.py
```

生成样例 Office 文件并验证提取、替换、合并、模板生成、CSV 转换全流程。

## 重新打包 exe

代码有改动后,如需重新打包:

```powershell
python -m PyInstaller --noconsole --collect-all PySide6 --name OfficeToolbox --icon office_toolbox\resources\app.ico --add-data "office_toolbox\resources;resources" office_toolbox\main.py
```

产物在 `dist\OfficeToolbox\` 目录,可整体改名为中文并复制到其他电脑使用(转 PDF 功能仍需目标电脑装有 Office)。
