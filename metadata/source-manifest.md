# 来源清单

提取日期：2026-09-02（Asia/Shanghai）

## 目标文件

`%APPDATA%\\Microsoft\\Word\\STARTUP\\PatentTools.dotm`

同名桌面副本：`%USERPROFILE%\\Desktop\\VBA宏安装\\PatentTools.dotm`

目标模板的 SHA-256：`9FF8F22CDB86E3835DA1A51AD4FCF47E5F77DAC630826A8C388D0E662B2E2282`

## 导出映射

| 导出文件 | VBA 模块 | 中文入口宏 |
| --- | --- | --- |
| `src/modClaimFigureTag.bas` | `modClaimFigureTag` | `权利要求标号` |
| `src/modClaimToInvention.bas` | `modClaimToInvention` | `权利要求转发明内容` |
| `src/modEmbodimentFigureTag.bas` | `modEmbodimentFigureTag` | `具体实施方式标号` |
| `src/modHeaderSync.bas` | `modHeaderSync` | `页眉双模式同步` |
| `src/modCleanFormatFull.bas` | `modCleanFormatFull` | `一键改格式_全文修改_不兼容修订` |
| `src/modCleanFormatTrackLite.bas` | `modCleanFormatTrackLite` | `一键改格式_选中部分修改_兼容修订` |
| `src/ThisDocument.cls` | `ThisDocument` | 无 |

## 明确排除

按项目范围排除 `Normal.dotm` 和启动目录中的英文宏/测试模块（包括 `PatentExtractor`、`AutoExport`、`DictModel`、`JsonWriter`、`Patterns`、`TestRunner` 等）。

