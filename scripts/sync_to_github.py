import ctypes
import os
import sys
import subprocess
import shutil
import glob
from ctypes import wintypes
from export_all_properties_to_html import export_db_to_html

if sys.stdout:
    sys.stdout.reconfigure(encoding='utf-8')

advapi32 = ctypes.windll.advapi32
class CREDENTIAL(ctypes.Structure):
    _fields_ = [
        ('Flags', wintypes.DWORD),
        ('Type', wintypes.DWORD),
        ('TargetName', wintypes.LPWSTR),
        ('Comment', wintypes.LPWSTR),
        ('LastWritten', wintypes.FILETIME),
        ('CredentialBlobSize', wintypes.DWORD),
        ('CredentialBlob', ctypes.POINTER(ctypes.c_byte)),
        ('Persist', wintypes.DWORD),
        ('AttributeCount', wintypes.DWORD),
        ('Attributes', ctypes.c_void_p),
        ('TargetAlias', wintypes.LPWSTR),
        ('UserName', wintypes.LPWSTR),
    ]

PCREDENTIAL = ctypes.POINTER(CREDENTIAL)
CredReadW = advapi32.CredReadW
CredReadW.argtypes = [wintypes.LPWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.POINTER(PCREDENTIAL)]
CredReadW.restype = wintypes.BOOL

def main(commit_msg=None):
    print("==================================================")
    print("🚀 開始執行看屋平台程式碼與網站同步至 GitHub...")
    print("==================================================")

    # 1. Export latest PostgreSQL properties into local HTML
    print("\n[Step 1/4] 正在從 PostgreSQL 匯出最新物件資料至 HTML 儀表板...")
    try:
        export_db_to_html()
    except Exception as e:
        print(f"⚠️ 匯出 PostgreSQL 至 HTML 失敗 ({e})，繼續同步現有檔案...")

    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    
    # Preferred repository location (Documents/GitHub/taichung-house-hunting for GitHub Desktop integration)
    preferred_repo = os.path.join(os.path.expanduser('~'), 'Documents', 'GitHub', 'taichung-house-hunting')
    temp_repo = os.path.join(os.environ.get('TEMP', os.path.expanduser('~')), 'taichung-house-hunting')

    if os.path.exists(os.path.join(preferred_repo, '.git')):
        repo_dir = preferred_repo
    elif os.path.exists(os.path.join(temp_repo, '.git')):
        repo_dir = temp_repo
    else:
        repo_dir = preferred_repo
        print(f"正在 Clone GitHub 專案至 {repo_dir}...")
        os.makedirs(os.path.dirname(repo_dir), exist_ok=True)
        subprocess.run(['git', 'clone', 'https://github.com/derekhcw168/taichung-house-hunting.git', repo_dir], check=True)

    print(f"\n[Step 2/4] 本機 Git Repository 目錄：{repo_dir}")
    print("正在將所有程式碼、腳本、資料與說明文件同步至 Git 目錄...")

    # A. Static Web Frontends
    html_src = os.path.join(base_dir, "591看屋物件地圖與比較分析.html")
    csv_src = os.path.join(base_dir, "591看屋物件綜合比較表.csv")
    renov_src = os.path.join(base_dir, "大墩人文裝修估價單.html")

    if os.path.exists(html_src):
        shutil.copy2(html_src, os.path.join(repo_dir, "index.html"))
    if os.path.exists(csv_src):
        shutil.copy2(csv_src, os.path.join(repo_dir, "591看屋物件綜合比較表.csv"))
    if os.path.exists(renov_src):
        shutil.copy2(renov_src, os.path.join(repo_dir, "dadun-renovation.html"))

    # B. Scripts folder (Backend Python scripts, ETL, Crawler)
    scripts_src = os.path.join(base_dir, "scripts")
    scripts_dst = os.path.join(repo_dir, "scripts")
    if os.path.exists(scripts_src):
        if os.path.exists(scripts_dst):
            shutil.rmtree(scripts_dst, ignore_errors=True)
        shutil.copytree(scripts_src, scripts_dst, ignore=shutil.ignore_patterns('__pycache__', '*.pyc'))

    # C. Batch / VBS utilities
    for pattern in ["*.bat", "*.vbs"]:
        for file in glob.glob(os.path.join(base_dir, pattern)):
            shutil.copy2(file, os.path.join(repo_dir, os.path.basename(file)))

    # D. Markdown documentation & checklists
    for pattern in ["*.md", "*.txt"]:
        for file in glob.glob(os.path.join(base_dir, pattern)):
            # Don't overwrite README if already specialized, unless user modified root
            dest_file = os.path.join(repo_dir, os.path.basename(file))
            if os.path.basename(file).lower() != 'readme.md' or not os.path.exists(dest_file):
                shutil.copy2(file, dest_file)

    # Ensure .gitignore exists
    gitignore_path = os.path.join(repo_dir, ".gitignore")
    if not os.path.exists(gitignore_path):
        with open(gitignore_path, 'w', encoding='utf-8') as f:
            f.write("# Python\n__pycache__/\n*.py[cod]\n*.log\n\n# Binaries & Temp\n*.exe\n*.pdb\n*.tmp\n\n# Large Media & MHTML\n物件圖片/\n待看物件/\n看完列入考慮的物件/\n看完之後不考慮的物件/\n*.mhtml\n")

    # Ensure .nojekyll exists for GitHub Pages
    nojekyll_path = os.path.join(repo_dir, ".nojekyll")
    if not os.path.exists(nojekyll_path):
        open(nojekyll_path, 'w').close()

    # 3. Git commit
    print("\n[Step 3/4] 正在進行 Git 版本控管提交 (git commit)...")
    subprocess.run(['git', 'add', '-A'], cwd=repo_dir, check=True)
    
    if not commit_msg:
        commit_msg = 'Update PostgreSQL backend scripts, database ETL, and house hunting dashboard'
    
    # Check if there are changes to commit
    status_proc = subprocess.run(['git', 'status', '--porcelain'], cwd=repo_dir, capture_output=True, text=True)
    if status_proc.stdout.strip():
        subprocess.run(['git', 'commit', '-m', commit_msg], cwd=repo_dir, check=True)
        print(f"✅ 已成功建立 Commit: {commit_msg}")
    else:
        print("ℹ️ 本地檔案與 Git Repository 內容完全一致，無需新 Commit。")

    # 4. Push to GitHub
    print("\n[Step 4/4] 讀取 Windows 認證管理員並推送至 GitHub (main 與 gh-pages)...")
    pcred = PCREDENTIAL()
    target = 'GitHub - https://api.github.com/derekhcw168'
    if not CredReadW(target, 1, 0, ctypes.byref(pcred)):
        print("❌ 無法讀取 Windows 認證管理員中的 GitHub 憑證")
        return

    token = ctypes.string_at(pcred.contents.CredentialBlob, pcred.contents.CredentialBlobSize).decode('utf-8', errors='ignore')
    remote_url = f'https://derekhcw168:{token}@github.com/derekhcw168/taichung-house-hunting.git'

    # Push to main branch (Code version control)
    print("--> 正在推送程式碼至 [main] 分支 (全專案原始碼版本控管)...")
    subprocess.run(['git', 'push', remote_url, 'main'], cwd=repo_dir, check=True)

    # Push to gh-pages branch (Static website deployment)
    print("--> 正在發布最新網站至 [gh-pages] 分支 (GitHub Pages 雲端站台)...")
    subprocess.run(['git', 'push', remote_url, 'main:gh-pages', '--force'], cwd=repo_dir, check=True)

    print("\n" + "=" * 50)
    print("🎉 同步與部署全部完成！")
    print("📂 GitHub 原始碼倉庫：https://github.com/derekhcw168/taichung-house-hunting")
    print("🌐 GitHub Pages 線上網站：https://derekhcw168.github.io/taichung-house-hunting/")
    print("=" * 50)

if __name__ == '__main__':
    msg = sys.argv[1] if len(sys.argv) > 1 else None
    main(msg)
