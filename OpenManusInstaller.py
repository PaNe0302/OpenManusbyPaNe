import os
import subprocess
import requests
import shutil
import time
import tkinter as tk
from tkinter import messagebox, filedialog
import win32com.client  # Để tạo shortcut

# Thỏa thuận sử dụng dịch vụ (ToS)
TOS_TEXT = """
Thỏa thuận sử dụng dịch vụ OpenManusInstaller:
1. Phần mềm này miễn phí, không chịu trách nhiệm cho bất kỳ thiệt hại nào khi sử dụng.
2. Yêu cầu kết nối mạng để tải các thành phần cần thiết.
3. Người dùng cần máy tính đủ mạnh (ít nhất 8GB RAM, 15GB dung lượng trống).
4. Không sao chép hoặc phân phối lại mà không có sự cho phép.
Nhấn 'Đồng ý' để tiếp tục hoặc 'Hủy' để thoát.
"""

CONFIG_CONTENT = """
[llm]
model = "llama3"
base_url = "http://localhost:11434/v1"
api_key = "ollama"
"""

class OpenManusInstallerUI:
    def __init__(self, root):
        self.root = root
        self.root.title("OpenManusInstaller")
        self.root.geometry("400x300")
        
        # Biến lưu thư mục cài đặt
        self.install_dir = tk.StringVar(value=os.path.expanduser("~/OpenManusInstaller"))
        
        # Giao diện ToS
        self.show_tos()

    def show_tos(self):
        """Hiển thị thỏa thuận sử dụng dịch vụ."""
        self.clear_window()
        tk.Label(self.root, text="Thỏa thuận sử dụng dịch vụ", font=("Arial", 12, "bold")).pack(pady=10)
        tk.Label(self.root, text=TOS_TEXT, wraplength=350, justify="left").pack(pady=10)
        tk.Button(self.root, text="Đồng ý", command=self.show_install_options).pack(side="left", padx=20, pady=20)
        tk.Button(self.root, text="Hủy", command=self.root.quit).pack(side="right", padx=20, pady=20)

    def show_install_options(self):
        """Hiển thị tùy chọn cài đặt."""
        self.clear_window()
        tk.Label(self.root, text="Chọn thư mục cài đặt", font=("Arial", 12, "bold")).pack(pady=10)
        tk.Entry(self.root, textvariable=self.install_dir, width=40).pack(pady=5)
        tk.Button(self.root, text="Duyệt...", command=self.browse_folder).pack(pady=5)
        tk.Button(self.root, text="Bắt đầu cài đặt", command=self.start_installation).pack(pady=20)

    def clear_window(self):
        """Xóa các widget hiện tại trong cửa sổ."""
        for widget in self.root.winfo_children():
            widget.destroy()

    def browse_folder(self):
        """Mở hộp thoại chọn thư mục."""
        folder = filedialog.askdirectory(initialdir=self.install_dir.get())
        if folder:
            self.install_dir.set(folder)

    def run_command(self, command, shell=False):
        """Chạy lệnh và kiểm tra lỗi."""
        try:
            subprocess.run(command, check=True, shell=shell)
            return True
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Lỗi", f"Lỗi khi chạy {command}: {e}")
            return False

    def check_and_install_prerequisites(self):
        """Kiểm tra và cài đặt Python, Git."""
        if not shutil.which("python"):
            messagebox.showinfo("Thông báo", "Đang cài Python...")
            self.run_command('winget install -e --id Python.Python.3.11', shell=True)
        if not shutil.which("git"):
            messagebox.showinfo("Thông báo", "Đang cài Git...")
            self.run_command('winget install -e --id Git.Git', shell=True)

    def install_ollama(self):
        """Tải và cài đặt OLLAMA."""
        ollama_url = "https://ollama.com/download/OllamaSetup.exe"
        ollama_exe = os.path.join(self.install_dir.get(), "OllamaSetup.exe")
        
        if not os.path.exists(self.install_dir.get()):
            os.makedirs(self.install_dir.get())
        
        messagebox.showinfo("Thông báo", "Đang tải OLLAMA...")
        response = requests.get(ollama_url, stream=True)
        with open(ollama_exe, "wb") as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        
        messagebox.showinfo("Thông báo", "Đang cài đặt OLLAMA...")
        self.run_command(f'"{ollama_exe}" /S', shell=True)
        os.remove(ollama_exe)

    def pull_llama3(self):
        """Tải mô hình Llama 3."""
        messagebox.showinfo("Thông báo", "Đang khởi động OLLAMA...")
        subprocess.Popen("ollama serve", shell=True)
        time.sleep(5)
        messagebox.showinfo("Thông báo", "Đang tải mô hình Llama 3 (có thể mất vài phút)...")
        self.run_command("ollama pull llama3", shell=True)

    def install_openmanusweb(self):
        """Tải và cấu hình OpenManusWeb."""
        os.chdir(self.install_dir.get())
        messagebox.showinfo("Thông báo", "Đang tải OpenManusWeb...")
        self.run_command("git clone https://github.com/wuzufeng/openmanusweb.git", shell=True)
        
        os.chdir(os.path.join(self.install_dir.get(), "openmanusweb"))
        messagebox.showinfo("Thông báo", "Đang cài đặt phụ thuộc...")
        self.run_command("pip install -r requirements.txt", shell=True)
        
        config_path = os.path.join(self.install_dir.get(), "openmanusweb", "config", "config.toml")
        with open(config_path, "w") as f:
            f.write(CONFIG_CONTENT)

    def create_shortcut(self):
        """Tạo shortcut trên Desktop."""
        shell = win32com.client.Dispatch("WScript.Shell")
        desktop = shell.SpecialFolders("Desktop")
        shortcut_path = os.path.join(desktop, "OpenManus.lnk")
        target = os.path.join(self.install_dir.get(), "openmanusweb", "app", "web", "app.py")
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = "python"
        shortcut.Arguments = f'"{target}"'
        shortcut.WorkingDirectory = os.path.dirname(target)
        shortcut.save()

    def start_installation(self):
        """Bắt đầu quá trình cài đặt."""
        self.clear_window()
        tk.Label(self.root, text="Đang cài đặt... Vui lòng chờ!", font=("Arial", 12)).pack(pady=20)
        self.root.update()
        
        self.check_and_install_prerequisites()
        self.install_ollama()
        self.pull_llama3()
        self.install_openmanusweb()
        self.create_shortcut()
        
        messagebox.showinfo("Thành công", "Cài đặt hoàn tất! Nhấn OK để mở OpenManus.")
        os.chdir(os.path.join(self.install_dir.get(), "openmanusweb"))
        subprocess.Popen("python app/web/app.py", shell=True)
        self.root.quit()

if __name__ == "__main__":
    root = tk.Tk()
    app = OpenManusInstallerUI(root)
    root.mainloop()
