"""
打包脚本 - 将PDF表格提取程序打包成EXE
"""
import subprocess
import sys
import os

def install_pyinstaller():
    """安装PyInstaller"""
    try:
        import PyInstaller
        print("PyInstaller已安装")
        return True
    except ImportError:
        print("正在安装PyInstaller...")
        result = subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"])
        return result.returncode == 0

def build_exe():
    """打包程序"""
    print("\n开始打包程序...")
    print("=" * 50)
    
    # 获取项目根目录（src的父目录）
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)
    os.chdir(project_root)  # 切换到项目根目录
    
    # 清理之前的文件
    import shutil
    for folder in ['build', 'dist']:
        folder_path = os.path.join(project_root, folder)
        if os.path.exists(folder_path):
            try:
                shutil.rmtree(folder_path)
                print(f"已清理: {folder}")
            except:
                pass
    
    spec_file = os.path.join(project_root, "PDF表格提取工具.spec")
    if os.path.exists(spec_file):
        try:
            os.remove(spec_file)
            print("已清理: PDF表格提取工具.spec")
        except:
            pass
    
    # 执行打包命令（从项目根目录执行，但指定src目录下的源文件）
    source_file = os.path.join(script_dir, "extract_pdf_tables.py")
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--console",
        "--name", "PDF表格提取工具",
        "--clean",
        source_file
    ]
    
    print("\n执行命令:", " ".join(cmd))
    print("=" * 50)
    
    result = subprocess.run(cmd, cwd=project_root)
    
    if result.returncode == 0:
        # 复制exe文件到release目录
        dist_exe = os.path.join(project_root, "dist", "PDF表格提取工具.exe")
        release_dir = os.path.join(project_root, "release")
        os.makedirs(release_dir, exist_ok=True)
        
        if os.path.exists(dist_exe):
            release_exe = os.path.join(release_dir, "PDF表格提取工具.exe")
            shutil.copy2(dist_exe, release_exe)
            print(f"\nEXE文件已复制到: release\\PDF表格提取工具.exe")
        
        # 复制使用说明到release目录
        docs_dir = os.path.join(project_root, "docs")
        usage_file = os.path.join(docs_dir, "使用说明.txt")
        if os.path.exists(usage_file):
            release_usage = os.path.join(release_dir, "使用说明.txt")
            shutil.copy2(usage_file, release_usage)
            print(f"使用说明已复制到: release\\使用说明.txt")
        
        print("\n" + "=" * 50)
        print("打包完成！")
        print("可分发文件位置: release\\")
        print("=" * 50)
    else:
        print("\n打包失败，请检查错误信息")
    
    return result.returncode == 0

if __name__ == "__main__":
    print("=" * 50)
    print("PDF表格提取程序 - 打包工具")
    print("=" * 50)
    
    if not install_pyinstaller():
        print("错误: 无法安装PyInstaller")
        input("按回车键退出...")
        sys.exit(1)
    
    if build_exe():
        print("\n打包成功！")
    else:
        print("\n打包失败！")
    
    input("\n按回车键退出...")
