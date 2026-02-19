"""
一键打包脚本
生成单个 exe 文件，隐藏控制台窗口
"""

import PyInstaller.__main__
import os
import shutil

# 项目配置
APP_NAME = "宽面板转长面板工具"
MAIN_SCRIPT = "src/main.py"
ICON_FILE = "icon.ico"  # 可选图标

# 清理旧的构建文件
for folder in ['build', 'dist']:
    if os.path.exists(folder):
        shutil.rmtree(folder)
        print(f"✓ 清理 {folder} 目录")

# PyInstaller 参数
args = [
    MAIN_SCRIPT,
    '--onefile',                    # 单个 exe 文件
    '--windowed',                   # 隐藏控制台窗口
    '--name', APP_NAME,             # 应用名称
    '--noconfirm',                  # 覆盖时不询问
    '--clean',                      # 清理临时文件
    '--hidden-import=pandas',       # 显式导入
    '--hidden-import=openpyxl',
    '--hidden-import=customtkinter',
    '--hidden-import=tkinter',
    '--hidden-import=PIL',
    '--hidden-import=numpy',
]

# 添加图标（如果存在）
if os.path.exists(ICON_FILE):
    args.extend(['--icon', ICON_FILE])
    print(f"✓ 使用图标：{ICON_FILE}")
else:
    print("⚠️ 未找到 icon.ico，将使用默认图标")

# 执行打包
print("=" * 50)
print(f"开始打包：{APP_NAME}")
print("=" * 50)

PyInstaller.__main__.run(args)

# 检查输出
dist_path = os.path.join('dist', f'{APP_NAME}.exe')
if os.path.exists(dist_path):
    file_size = os.path.getsize(dist_path) / (1024 * 1024)
    print("=" * 50)
    print(f"✅ 打包成功!")
    print(f"📦 文件位置：{dist_path}")
    print(f"📊 文件大小：{file_size:.1f} MB")
    print("=" * 50)
else:
    print("❌ 打包失败，请检查错误信息")