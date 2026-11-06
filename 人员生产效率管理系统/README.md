# 新人员管理系统

一个使用 Python、Tkinter 和 Excel 构建的桌面人员与每日产能数据管理工具，主要面向 Windows。

## 功能

- 按员工姓名、工号、班组、机种、物料、客户、工序、模具号等条件组合查询
- 按日期范围或指定日期查询
- 添加员工每日生产数据，自动计算达成率和合格率
- 查询结果汇总、Excel 导出、热力图和趋势图导出
- 使用 Excel 模板批量填写和导入数据
- 导入前校验内部选项、日期、数字格式和重复记录
- 维护下拉选项、查看最近历史记录、删除指定数据
- 将员工数据加密保存到本地
- 提供应用内操作帮助

## 数据与隐私

这是单机桌面应用，不包含任何真实员工数据或内部生产数据。首次运行时会自动创建空数据库和加密配置文件。

以下文件属于本机运行数据，已经加入 `.gitignore`，不要提交或分享：

- `员工数据.xlsx`
- `选项内容.dat`
- `程序授权.dat`
- `查询结果/`
- `表格模板/`
- `程序错误日志.txt`
- 任何解密的 Excel、密钥文件或管理工具

查询和模板导入后导出的 Excel 是明文文件。处理真实员工信息时，请遵守适用的隐私和个人信息保护要求。

## 环境要求

- Python 3.10 或更高版本
- Windows 10/11
- Tkinter，通常随 Windows Python 一起安装

## 安装与运行

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python main.py
```

如果 PowerShell 禁止执行激活脚本，也可以不激活虚拟环境：

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe main.py
```

## 配置密钥

源码中的 `EMP_DATA_KEY` 和 `EMP_RENEWAL_KEY` 回退值只用于公开演示，不具备保密作用。处理真实数据前，应在首次运行前设置自己的随机密钥：

```powershell
$env:EMP_DATA_KEY = "replace-with-a-long-random-data-key"
$env:EMP_RENEWAL_KEY = "replace-with-a-long-random-renewal-key"
$env:EMP_LICENSE_EXPIRE_DATE = "2099-12-31"
python main.py
```

请勿把环境变量值、真实密钥或包含真实数据的文件提交到 Git。

密钥在首次运行后不能随意更换，否则已有加密数据将无法读取。需要更换时，应先关闭程序并迁移已有数据。

## 测试

```powershell
python main.py --selftest
python main.py --smoke
```

`--selftest` 测试加密、数据读写和基础业务逻辑；`--smoke` 测试主要界面和下拉控件。

## 打包

安装 PyInstaller 后执行：

```powershell
python -m pip install pyinstaller
.\build.ps1
```

输出位于 `dist/新人员管理系统/`。分发时需要发送整个输出目录，不能只发送 EXE。

## 项目结构

```text
.
├─ main.py
├─ requirements.txt
├─ build.ps1
├─ README.md
├─ SECURITY.md
├─ CONTRIBUTING.md
├─ LICENSE
└─ .gitignore
```

## 安全边界

这是一个本地桌面程序。本地加密可以保护数据文件，但无法阻止拥有本机控制权和逆向能力的人修改程序或提取运行所需密钥。高敏感、多用户或需要强授权的场景应使用服务端存储和服务器授权验证。

## 许可证

本项目使用 MIT License，详见 [LICENSE](LICENSE)。
