# DWG 批量转图片桌面工具

一个简洁的 Windows 桌面小工具（Tkinter），支持：

- 指定根目录后**递归查找多级子目录**中的 `.dwg`
- 批量输出为 `png/jpg`
- 输出模式支持“保留原目录结构”勾选（勾选时按源顶级目录镜像输出；取消勾选时统一输出到输出根目录）
- 布局渲染模式：`auto / model / layout(可指定布局名)`
- 颜色模式：`bw(黑白)` / `original(保留原色)`

## 原理

DWG 直接渲染在纯 Python 生态中不稳定，因此采用两步：

1. 使用 **ODA File Converter** 将 DWG 批量转 DXF
2. 使用 `ezdxf + matplotlib` 将 DXF 渲染为图片

## ODA File Converter 下载

- 官方下载页：https://www.opendesign.com/guestfiles/oda_file_converter

程序会尝试自动检测本机已安装的 ODA File Converter；若未检测到，会在界面中提示并给出下载地址。

## 源码运行（开发/调试）

1. 安装 Python 3.10+
2. 安装依赖：

```bash
pip install -r requirements.txt
```

3. 运行：

```bash
python app.py
```

## 生成 Windows EXE（给用户直接下载使用）

### 方案 A：本地 Windows 一键打包

在 Windows 命令行执行：

```bat
build_exe.bat
```

成功后会得到：

- `dist\\dwg2img.exe`

### 方案 B：GitHub Actions 自动打包（推荐）

仓库已提供工作流：

- `.github/workflows/build-windows-exe.yml`

使用方式：

1. 推送代码到远端仓库。
2. 在 GitHub 的 **Actions** 页面运行 **Build Windows EXE**（也支持 push 到 `work` 分支自动触发）。
3. 在该次 workflow 的 **Artifacts** 下载：`dwg2img-exe`。
4. 解压后得到 `dwg2img.exe`，可直接分发给用户。

### 方案 C：Release 自动上传 EXE（长期分享链接）

仓库已提供发布工作流：

- `.github/workflows/release-windows-exe.yml`

使用方式（推荐用于给外部用户稳定下载地址）：

1. 创建并推送版本标签（例如 `v1.0.0`）。
2. GitHub Actions 会自动构建并创建同名 Release。
3. `dwg2img.exe` 会作为 Release Asset 自动上传。
4. 你可直接把该 Release 页地址或 Asset 下载地址发给用户。

也支持在 Actions 页面手动触发该工作流，并指定 `tag_name`。

## 使用步骤

1. 选择 DWG 根目录（会递归扫描）
2. 选择输出根目录
3. 选择 ODA File Converter 可执行文件（可自动检测）
4. 选择输出模式：勾选“保留原目录结构”
5. 选择图片格式和 DPI（默认 `96`）
6. 设置并发参数（可选）：
   - 渲染并发：`0` 为自动（按 CPU 自适应）
   - ODA并发：`0` 为自动（默认保守并发）
   - ODA批大小：默认 `200`，大批量场景建议保持 100~300
7. 选择布局模式：
   - `auto`：优先非 Model 布局，找不到再回退 Model
   - `model`：强制渲染 Model
   - `layout`：优先指定布局名，未命中再选择首个有实体布局
8. 选择颜色模式：
   - `bw`：白底黑线，适合与客户标准黑白图对齐
   - `original`：尽量保留图层原色
9. 点击“开始批量转换”

## 输出一致性说明（尺寸 / 位深 / DPI）


## 大批量性能与内存策略

- 转换流程改为“**DWG->DXF 分批 + 图片渲染多进程并发**”，可显著提升吞吐。
- 渲染使用多进程池并启用 `max_tasks_per_child`，定期重建子进程，降低第三方库长期运行导致的内存膨胀风险。
- 每个 ODA 批次完成后会及时清理中间 DXF 目录，避免累计占用磁盘与内存缓存。
- 界面提供进度条与百分比进度，日志仅输出关键进度与结果摘要。

- 若图纸带有 Layout 纸张尺寸信息（paper width/height），程序会按该尺寸计算输出画布尺寸，尽量与源图纸页面尺寸一致。
- 若没有可用纸张尺寸，会按图元边界估算比例并固定画布，避免 `bbox_inches='tight'` 导致每张图像素尺寸漂移。
- 程序会在保存后统一转为 `RGB`，默认输出 **24-bit** 图像，并写入设定 DPI（默认 `96`）。
- 颜色模式选 `bw` 时按黑白出图（白底黑线），可减少 CAD 图层颜色差异带来的观感偏差；选 `original` 时尽量保留原色。
- 日志会显示每张图的布局名、像素尺寸、DPI 与 24-bit 标记，方便与客户样图核对。

## 说明

- 若图纸使用 SHX 字体、代理对象（Proxy Objects）或特殊 CAD 实体，`ezdxf + matplotlib` 可能仍会出现文字缺失；程序已优先使用代理图形渲染，但仍建议用 AutoCAD/专业引擎导出作为基准对照。
- 若某个 DWG 转 DXF 失败，会在日志中提示跳过。
- 未勾选“保留原目录结构”时，若存在同名文件会互相覆盖（可后续按需加入重名去重策略）。
