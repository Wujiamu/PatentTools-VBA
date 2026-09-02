# Word 专利撰写宏

本目录保存从本机 Word 宏模板中提取的、宏名为中文的专利申请文件撰写辅助 VBA。

后续开发路线、样例要求、单入口工具面板方案和测试约定见 [AGENTS.md](AGENTS.md)。

当前已完成第一阶段合并：源码中新增唯一公开入口 `专利撰写工具箱`、工具分发器和第一版无模式面板；原六个业务入口已通过 `Option Private Module` 隐藏，改由面板调用。

## 已提取的中文入口宏

| 面板中的工具名 | 源模块 | 用途概览 |
| --- | --- | --- |
| `权利要求标号` | `modClaimFigureTag.bas` | 为选定的权利要求补充附图标记 |
| `权利要求转发明内容` | `modClaimToInvention.bas` | 将选定的权利要求转换为发明内容表述 |
| `具体实施方式标号` | `modEmbodimentFigureTag.bas` | 为具体实施方式补充附图标记 |
| `页眉双模式同步` | `modHeaderSync.bas` | 同步文档页眉中的案件信息 |
| `一键改格式_全文修改_不兼容修订` | `modCleanFormatFull.bas` | 对全文执行格式清理，不兼容修订模式 |
| `一键改格式_选中部分修改_兼容修订` | `modCleanFormatTrackLite.bas` | 对选区执行格式清理，兼容修订模式 |

`src/ThisDocument.cls` 也保留在导出结果中，用于完整还原 VBA 工程结构；它本身不包含入口宏。`modToolboxEntry.bas` 是唯一公开入口，`modToolDispatcher.bas` 和 `frmPatentToolbox.frm` 负责面板调用。

## 目录

- `src/`：可继续开发和审阅的 UTF-8 VBA 源码；导入 VBE 前需要按中文 Windows 代码页转换。
- `scripts/prepare_vbe_import.ps1`：把 `src/` 转为代码页 936 的临时导入副本，输出到被 Git 忽略的 `build/vbe-import/`。
- `dist/PatentTools.dotm`：从 Word 启动目录复制的原始宏模板，作为合并前的可运行基线；源码合并结果尚未写回这个二进制文件。

## 来源与筛选

- 主要来源：`%APPDATA%\\Microsoft\\Word\\STARTUP\\PatentTools.dotm`
- 桌面安装副本：`%USERPROFILE%\\Desktop\\VBA宏安装\\PatentTools.dotm`
- 两个来源文件在提取时内容一致；原始模板保存在 `dist/`。
- 模板内部模块名使用英文，但实际公开入口宏名均为中文，因此纳入本项目。
- `Normal.dotm` 以及启动目录中的 `PatentExtractor`、`AutoExport`、`DictModel`、`JsonWriter`、`Patterns` 等英文入口/模块按你的说明视为 CAD 相关宏，已排除，没有复制进本仓库。

导出的 `src/*.bas` 保持 VBA 源码内容，并统一保存为 UTF-8 便于 Git 审阅；`dist/PatentTools.dotm` 保留原始二进制。VBE 对 UTF-8 和中文标识符的支持不可靠，后续构建脚本应先把待导入副本转换为当前中文 Windows 的 ANSI/GBK 代码页，再将六个原业务模块、`modToolboxEntry.bas`、`modToolDispatcher.bas` 和 `frmPatentToolbox.frm` 导入测试模板。不要直接把仓库中的 UTF-8 文件导入正式模板。

准备导入副本：

```powershell
.\scripts\prepare_vbe_import.ps1
```

然后在测试用 `.dotm` 的 VBE 中导入 `build/vbe-import/` 下的模块和窗体，验证后再生成发布模板。

## 后续开发提示

这些宏依赖 Word 对象模型，并使用晚绑定的 `Scripting.Dictionary`、`VBScript.RegExp` 等组件。修改后可先在副本模板中验证，再替换 `dist/PatentTools.dotm`，避免直接影响 Word 的启动加载项。
