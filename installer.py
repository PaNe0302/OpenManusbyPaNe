import os
import subprocess
import requests
import shutil
import time
import tkinter as tk
from tkinter import messagebox, filedialog, Checkbutton
import win32com.client

# Thư mục cài đặt mặc định
DEFAULT_INSTALL_DIR = r"C:\Program Files (x86)\OpenManusbyPaNe"
CONFIG_CONTENT = """
[llm]
model = "llama3"
base_url = "http://localhost:11434/v1"
api_key = "ollama"
"""

class OpenManusInstallerUI:
    def __init__(self, root):
        self.root = root
        self.root.title("OpenManusbyPaNe Installer")
        self.root.geometry("500x400")
        
        # Biến lưu tùy chọn
        self.install_dir = tk.StringVar(value=DEFAULT_INSTALL_DIR)
        self.create_shortcut = tk.BooleanVar(value=True)
        
        # Đọc EULA từ tệp
        with open("EULA.txt", "r", encoding="utf-8") as f:
            self.eula_text = f.read()
        
        self.show_eula()

    def show_eula(self):
        self.clear_window()
        tk.Label(self.root, text="Thỏa thuận sử dụng phần mềm", font=("Arial", 12, "bold")).pack(pady=10)
        eula_frame = tk.Frame(self.root)
        eula_frame.pack(pady=10)
        eula_text_widget = tk.Text(eula_frame, height=10, width=60, wrap="word")
        eula_text_widget.insert("1.0", self.eula_text)
        eula_text_widget.config(state="disabled")
        eula_text_widget.pack()
        tk.Button(self.root, text="Đồng ý", command=self.show_install_options).pack(side="left", padx=20, pady=20)
        tk.Button(self.root, text="Hủy", command=self.root.quit).pack(side="right", padx=20, pady=20)

    def show_install_options(self):
        self.clear_window()
        tk.Label(self.root, text="Cài đặt OpenManusbyPaNe", font=("Arial", 12, "bold")).pack(pady=10)
        
        tk.Label(self.root, text="Thư mục cài đặt:").pack(pady=5)
        tk.Entry(self.root, textvariable=self.install_dir, width=50).pack(pady=5)
        tk.Button(self.root, text="Duyệt...", command=self.browse_folder).pack(pady=5)
        
        tk.Checkbutton(self.root, text="Tạo shortcut trên Desktop", variable=self.create_shortcut).pack(pady=10)
        
        tk.Button(self.root, text="Cài đặt", command=self.start_installation).pack(pady=20)

    def clear_window(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def browse_folder(self):
        folder = filedialog.askdirectory(initialdir=self.install_dir.get())
        if folder:
            self.install_dir.set(folder)

    def run_command(self, command, shell=False):
        try:
            subprocess.run(command, check=True, shell=shell)
            return True
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Lỗi", f"Lỗi: {e}")
            return False

    def check_and_install_prerequisites(self):
        if not shutil.which("python"):
            messagebox.showinfo("Thông báo", "Đang cài Python...")
            self.run_command('winget install -e --id Python.Python.3.11', shell=True)
        if not shutil.which("git"):
            messagebox.showinfo("Thông báo", "Đang cài Git...")
            self.run_command('winget install -e --id Git.Git', shell=True)

    def install_ollama(self):
        ollama_url = "https://ollama.com/download/OllamaSetup.exe"
        ollama_exe = os.path.join(self.install_dir.get(), "OllamaSetup.exe")
        
        if not os.path.exists(self.install_dir.get()):
            os.makedirs(self.install_dir.get())
        
        messagebox.showinfo("Thông báo", "Đang tải OLLAMA...")
        response = requests.get(ollama_url, stream=True)
        with open(ollama_exe, "wb") as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        
        messagebox.showinfo("Thông báo", "Đang cài OLLAMA...")
        self.run_command(f'"{ollama_exe}" /S', shell=True)
        os.remove(ollama_exe)

    def pull_llama3(self):
        messagebox.showinfo("Thông báo", "Đang khởi động OLLAMA...")
        subprocess.Popen("ollama serve", shell=True)
        time.sleep(5)
        messagebox.showinfo("Thông báo", "Đang tải Llama 3 (có thể mất vài phút)...")
        self.run_command("ollama pull llama3", shell=True)

    def install_openmanusweb(self):
        os.chdir(self.install_dir.get())
        messagebox.showinfo("Thông báo", "Đang tải OpenManusWeb...")
        self.run_command("git clone https://github.com/wuzufeng/openmanusweb.git", shell=True)
        
        os.chdir(os.path.join(self.install_dir.get(), "openmanusweb"))
        messagebox.showinfo("Thông báo", "Đang cài phụ thuộc...")
        self.run_command("pip install -r requirements.txt", shell=True)
        
        config_path = os.path.join(self.install_dir.get(), "openmanusweb", "config", "config.toml")
        with open(config_path, "w") as f:
            f.write(CONFIG_CONTENT)

    def create_shortcut(self):
        if self.create_shortcut.get():
            shell = win32com.client.Dispatch("WScript.Shell")
            desktop = shell.SpecialFolders("Desktop")
            shortcut_path = os.path.join(desktop, "OpenManusbyPaNe.lnk")
            target = os.path.join(self.install_dir.get(), "openmanusweb", "app", "web", "app.py")
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.Targetpath = "python"
            shortcut.Arguments = f'"{target}"'
            shortcut.WorkingDirectory = os.path.dirname(target)
            # Thêm icon nếu có (sẽ cập nhật sau khi bạn cung cấp)
            # shortcut.IconLocation = os.path.join(self.install_dir.get(), "resources", "icon.ico")
            shortcut.save()

    def start_installation(self):
        self.clear_window()
        tk.Label(self.root, text="Đang cài đặt... Vui lòng chờ!", font=("Arial", 12)).pack(pady=20)
        self.root.update()
        
        self.check_and_install_prerequisites()
        self.install_ollama()
        self.pull_llama3()
        self.install_openmanusweb()
        self.create_shortcut()
        
        messagebox.showinfo("Thành công", "Cài đặt xong! Nhấn OK để mở OpenManusbyPaNe.")
        os.chdir(os.path.join(self.install_dir.get(), "openmanusweb"))
        subprocess.Popen("python app/web/app.py", shell=True)
        self.root.quit()

if __name__ == "__main__":
    root = tk.Tk()
    app = OpenManusInstallerUI(root)
    root.mainloop()
