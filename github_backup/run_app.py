"""
IQC 效率管理系統 - 啟動器
自動啟動 Streamlit 應用並開啟瀏覽器
"""

import subprocess
import sys
import os
import webbrowser
import time
import socket

def check_port_available(port):
    """檢查端口是否可用"""
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    result = sock.connect_ex(('localhost', port))
    sock.close()
    return result != 0

def find_available_port(start_port=8501, max_attempts=10):
    """找到可用的端口"""
    for port in range(start_port, start_port + max_attempts):
        if check_port_available(port):
            return port
    return start_port

def main():
    print("=" * 60)
    print("🚀 IQC 效率管理系統 - 啟動中...")
    print("=" * 60)
    
    # 取得程式所在目錄
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包後的環境
        base_path = sys._MEIPASS
        print(f"📦 執行模式: 打包模式")
    else:
        # 開發環境
        base_path = os.path.dirname(os.path.abspath(__file__))
        print(f"🔧 執行模式: 開發模式")
    
    app_path = os.path.join(base_path, 'iqc_monitor_V22.py')
    
    # 檢查主程式是否存在
    if not os.path.exists(app_path):
        print(f"❌ 錯誤: 找不到主程式檔案 {app_path}")
        input("\n按 Enter 鍵退出...")
        return
    
    print(f"📂 程式路徑: {app_path}")
    
    # 尋找可用端口
    port = find_available_port()
    print(f"🔌 使用端口: {port}")
    
    # 啟動 Streamlit
    print("\n⏳ 正在啟動 Streamlit 服務...")
    
    # 在打包環境中，使用絕對路徑啟動 streamlit
    if getattr(sys, 'frozen', False):
        # 打包環境：直接使用 streamlit 可執行檔
        streamlit_script = os.path.join(os.path.dirname(sys.executable), 'streamlit.exe')
        if not os.path.exists(streamlit_script):
            # 如果找不到，嘗試用模組方式
            streamlit_script = sys.executable
            cmd = [streamlit_script, '-m', 'streamlit', 'run', app_path]
        else:
            cmd = [streamlit_script, 'run', app_path]
    else:
        # 開發環境
        cmd = [sys.executable, '-m', 'streamlit', 'run', app_path]
    
    # 添加參數
    cmd.extend([
        f'--server.port={port}',
        '--server.headless=true',
        '--browser.gatherUsageStats=false',
        '--server.fileWatcherType=none',
        '--theme.base=light',
        '--server.address=localhost'
    ])
    
    print(f"📝 執行命令: {' '.join(cmd[:3])}...")
    
    try:
        # 不使用 PIPE，讓輸出直接顯示到控制台以便除錯
        process = subprocess.Popen(
            cmd,
            # stdout=subprocess.PIPE,
            # stderr=subprocess.PIPE,
            # creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
        )
        
        # 等待服務啟動
        print("⏳ 等待服務啟動 (約 5-15 秒)...")
        service_started = False
        for i in range(15):
            time.sleep(1)
            # 檢查進程是否還在運行
            if process.poll() is not None:
                print(f"\n❌ 錯誤: Streamlit 進程意外終止 (退出碼: {process.returncode})")
                print("💡 請檢查是否缺少相關模組或配置")
                input("\n按 Enter 鍵退出...")
                return
            
            if not check_port_available(port):
                print("✅ 服務啟動成功！")
                service_started = True
                break
            print(f"   等待中... ({i+1}/15)")
        
        if not service_started:
            print("\n⚠️  警告: 服務啟動超時")
            print("💡 可能的原因:")
            print("   1. Streamlit 模組未正確打包")
            print("   2. 端口被佔用")
            print("   3. 防火牆阻擋")
            
            # 給用戶選擇
            choice = input("\n是否仍要開啟瀏覽器? (y/n): ")
            if choice.lower() != 'y':
                print("正在終止...")
                process.terminate()
                return
        
        # 自動開啟瀏覽器
        url = f'http://localhost:{port}'
        print(f"\n🌐 正在開啟瀏覽器...")
        print(f"📌 網址: {url}")
        
        webbrowser.open(url)
        
        print("\n" + "=" * 60)
        print("✅ 系統已成功啟動！")
        print("=" * 60)
        print("\n💡 使用提示:")
        print("   • 如果瀏覽器沒有自動開啟，請手動訪問上述網址")
        print("   • 請勿關閉此視窗，否則系統將停止運行")
        print("   • 關閉瀏覽器分頁不會停止系統")
        print("   • 如需退出，請關閉此視窗或按 Ctrl+C")
        print("\n" + "=" * 60)
        
        # 保持運行
        try:
            process.wait()
        except KeyboardInterrupt:
            print("\n\n🛑 正在關閉系統...")
            process.terminate()
            time.sleep(2)
            print("✅ 系統已安全關閉")
    
    except FileNotFoundError:
        print("\n❌ 錯誤: 找不到 Streamlit")
        print("💡 請確保已安裝 Streamlit: pip install streamlit")
        input("\n按 Enter 鍵退出...")
    
    except Exception as e:
        print(f"\n❌ 啟動失敗: {e}")
        print("\n詳細錯誤資訊:")
        import traceback
        traceback.print_exc()
        input("\n按 Enter 鍵退出...")

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f"\n❌ 程式執行錯誤: {e}")
        import traceback
        traceback.print_exc()
        input("\n按 Enter 鍵退出...")
