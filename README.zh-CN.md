# PixPoint 图片像素坐标拾取与标注工具

当前版本：`v1.2`

一个基于 PyQt5 的图片坐标拾取工具，适合查看大图、精确定位像素点并导出坐标。

## 功能说明

- 支持鼠标滚轮缩放，并且以鼠标位置为中心进行缩放。
- 支持右键拖拽平移，也支持按住 `Space` 后左键拖拽平移。
- 鼠标移动时，底部状态栏会实时显示原图坐标。
- 右侧提供放大镜，可查看当前像素附近的局部细节。
- 缩放到较高倍率时，会显示像素网格辅助线。
- 左键点击图片可添加点位。
- 右侧 `Points` 表格会记录 `Index / X / Y / Color`。
- 支持点位编辑、删除、点击表格自动定位。
- 支持将点位坐标保存为 `JSON` 文件。
- 支持将当前图片导出为 Excel 像素画 `.xlsx` 文件。
- 导出时可设置目标宽度，系统会自动按比例计算高度。
- Excel 像素画导出带有尺寸边界控制，当前支持常见 `2K/1440p` 级别导出。

## 启动方式

如果只是想直接使用 Windows 版本，可以从 GitHub Release 下载：

- [下载 PixPoint v1.2 Windows 版本](https://github.com/flowfish/pixpoint/releases/tag/v1.2)

如果想从源码运行，先克隆项目：

```bash
git clone https://github.com/flowfish/pixpoint.git
cd pixpoint
```

创建虚拟环境：

```bash
python -m venv .venv
```

在 Windows 上激活虚拟环境：

```bash
.venv\Scripts\activate
```

在 macOS / Linux 上激活虚拟环境：

```bash
source .venv/bin/activate
```

安装依赖：

```bash
pip install -r requirements.txt
```

启动程序：

```bash
python main.py
```

也可以直接带图片路径启动：

```bash
python main.py your_image.png
```

## 操作说明

- 打开图片：
  通过菜单栏 `File -> Open Image...` 选择图片。

- 缩放图片：
  鼠标滚轮向上放大，向下缩小。

- 平移图片：
  1. 按住鼠标右键拖拽。
  2. 按住 `Space`，再按左键拖拽。

- 查看当前坐标：
  鼠标移动时，底部状态栏会显示 `Current (Original): [X, Y]`。

- 添加点位：
  鼠标左键单击图片。

- 编辑点位：
  在右侧 `Points` 表格中双击某一行，或者先选中再点击 `Edit Point`。

- 删除点位：
  选中表格中的某一行后，点击 `Delete Point`，或直接按 `Delete` 键。

- 快速定位点位：
  点击右侧表格中的某一行，画布会自动定位到该点。

- 保存坐标：
  通过菜单栏 `File -> Save Coordinates...` 导出为 `JSON` 文件。

- 导出 Excel 像素画：
  通过菜单栏 `File -> Export Pixel To Excel...` 打开导出参数窗口。

- 设置导出宽度：
  默认宽度为 `100`，高度会按原图比例自动计算。

- 导出边界：
  为了兼顾性能和稳定性，当前限制最大导出尺寸为 `2560 x 2560`，总像素量不超过 `4194304`。

- 大图提示：
  当导出尺寸较大时，程序会弹出确认提示，并显示导出进度条。

## 导出格式

导出的 JSON 示例：

```json
{
  "image_path": "demo.png",
  "image_size": {
    "width": 1920,
    "height": 1080
  },
  "points": [
    {
      "index": 1,
      "x": 320,
      "y": 240,
      "color": "#ff6b6b"
    }
  ]
}
```

## 说明

- 推荐使用较新的 Python 3 环境。
- 当前实现基于 `PyQt5`。
- 如果只是日常取点和简单标注，这个版本已经可以直接使用。
