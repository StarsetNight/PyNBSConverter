# PyNBSConverter

> **⚠️ This project has been archived and is no longer maintained.**
> 本项目已归档（Archived），不再进行功能更新或维护，仅作为历史项目与参考实现保留。

---

## About

PyNBSConverter 是一个用于将 **Minecraft Note Block Studio（NBS）工程** 自动转换为 **第四代编码格式** 的 Python 工具。

项目基于以下库开发：

* `pynbs`
* `xlwings`

转换结果输出为 `.xlsx` 文件，可使用 **Microsoft Excel** 或 **WPS 表格** 打开。

---

## Project Status

本项目已经停止开发，不再接受：

* 新功能请求（Feature Requests）
* Bug 修复
* Pull Request
* Issue 支持

仓库将作为历史代码保留，供有需要的开发者参考或 Fork。

---

## Usage Notes

输入的 NBS 文件建议进行以下预处理：

1. 在 Note Block Studio 中按 `Ctrl + A` 全选音符，并整体向后移动，避免开头音符无法编码。
2. 当前版本不会自动进行复杂的编码压缩，建议提前删除不必要的和弦轨道。
3. 推荐使用 **10 t/s** 的 NBS 工程（其他速度未经完整测试）。

输出为 `.xlsx` 文件，其中颜色含义如下：

* **蓝色**：普通音符
* **黄色**：无延迟执行编码
* **绿色**：带延迟执行编码（绿色数量代表延迟的 gt 数）

---

## Build

克隆仓库并安装依赖：

```bash
git clone <repository-url>
cd PyNBSConverter
pip install -r requirements.txt
```

随后即可运行程序。

---

## License

Unless otherwise specified, this project is released under the license contained in this repository.

---

## Author

**StarsetNight**

GitHub: https://github.com/StarsetNight

---

## Archive Notice

本项目完成了其历史使命，目前没有继续维护的计划。

如果未来有人希望继续开发，请自由 Fork 本项目，并根据自己的需求进行修改。

感谢所有曾经关注和使用过本项目的人。
