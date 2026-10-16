# Mac 第一次使用

程序在你的 Mac 上运行。预算 Excel 由你选择导入，计算和版本保存在本机，不需要购买服务器。

## 1. 下载程序

在 GitHub 的 `codex/bus-budget-mac` 分支页面，点击绿色的 **Code → Download ZIP**，解压。不要从 main 分支下载，main 仍是原始仓库。

此源码包不含你的预算 Excel、截图、云端方案数据或虚拟环境。

## 2. 安装两个软件（只需一次）

- [Python 官方 Mac 下载页面](https://www.python.org/downloads/macos/)：安装 Python 3.12 或更新版本，建议 3.13 的 macOS universal2 installer。它同时适用于 Apple 芯片和 Intel 芯片。
- [LibreOffice 官方下载页面](https://www.libreoffice.org/download/download-libreoffice/)：下载适合你的 Mac 芯片的版本，把 LibreOffice.app 拖进“应用程序”，并正常打开一次。芯片类型可在苹果菜单的“关于本机”查看。

LibreOffice 用于计算原表公式，不会替换或卸载 Microsoft Excel。

## 3. 启动工具

打开解压后的程序文件夹，打开 **启动预算工具.command**。首次会安装依赖，需要联网并等一会儿。成功后会自动打开浏览器。

如果双击没能运行，可以打开 Mac 自带“终端”，输入 `bash` 和一个空格，把“启动预算工具.command”文件拖入终端，再按回车。无需关闭系统安全检查。

保留启动时打开的终端窗口。关闭窗口或按 Control+C 会停止工具；已保存的预算方案仍留在程序文件夹的 `data` 内。

## 4. 导入自己的基础表

在页面里选择或拖入《2027年预算基础表.xlsx》。第一次导入会检查文件并进行真实公式重算，可能需要数十秒。之后即可选择线路和月份、减车试算、恢复、保存版本和导出。

导入另一个文件不删除之前的数据目录。重要方案可从网页导出 Excel 保存。暂未确认的公里数和最低标准仍会提示“待确认”。

## 常见情况

- 提示未找到 Python：完成第2步，再启动。Mac 系统自带的 Python 可能过旧。
- 提示未找到 LibreOffice：确认它已放在“应用程序”，并正常打开过一次。
- 提示安装包证书验证失败：运行 Python 安装目录自带的 `Install Certificates.command` 后重试，不要禁用 TLS 校验。
- 网页打不开：查看终端是否还在运行，以及程序文件夹里的 `data/mac-startup.log`。把错误文字提供给开发者即可，不需要发送预算文件。

自动测试目前在 Linux 云环境运行；Mac 原生界面及启动尚需在你的电脑上验证，不能将兼容性检查当作已经在真实 Mac 上运行通过。
