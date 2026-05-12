import os
import extract_msg
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from threading import Thread

class MsgExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MSG PDF Attachment Extractor")
        self.root.geometry("600x500")
        self.selected_path = tk.StringVar()

        # UI Layout
        tk.Label(root, text="MSG Attachment Extractor", font=("Arial", 14, "bold")).pack(pady=10)
        
        # Directory Selection Frame
        frame = tk.Frame(root)
        frame.pack(pady=10, padx=20, fill='x')
        
        tk.Entry(frame, textvariable=self.selected_path, state='readonly').pack(side='left', expand=True, fill='x', padx=(0, 5))
        tk.Button(frame, text="Browse Folder", command=self.browse_folder).pack(side='right')

        # Start Button
        self.start_btn = tk.Button(root, text="START EXTRACTION", bg="#2ecc71", fg="white", 
                                   font=("Arial", 10, "bold"), height=2, command=self.start_thread)
        self.start_btn.pack(pady=10, fill='x', padx=20)

        # Log Area
        self.log_area = scrolledtext.ScrolledText(root, width=70, height=15, font=("Consolas", 9))
        self.log_area.pack(pady=10, padx=20)

    def browse_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.selected_path.set(path)
            self.log(f"Selected: {path}")

    def log(self, message):
        self.log_area.insert(tk.END, f"{message}\n")
        self.log_area.see(tk.END)

    def start_thread(self):
        """Runs the extraction in a separate thread to keep UI from freezing."""
        path = self.selected_path.get()
        if not path:
            messagebox.showwarning("Warning", "Please select a directory first!")
            return
        
        self.start_btn.config(state='disabled')
        thread = Thread(target=self.process_folders, args=(path,))
        thread.start()

    def process_folders(self, root_path):
        folder_count = 0
        try:
            subdirs = [d for d in os.listdir(root_path) if os.path.isdir(os.path.join(root_path, d))]
            
            for subdir in subdirs:
                folder_path = os.path.join(root_path, subdir)
                for file in os.listdir(folder_path):
                    if file.lower().endswith(".msg"):
                        msg_file_path = os.path.join(folder_path, file)
                        try:
                            msg = extract_msg.Message(msg_file_path)
                            for attachment in msg.attachments:
                                att_name = attachment.longFilename or attachment.name
                                if att_name and att_name.lower().endswith(".pdf"):
                                    save_path = os.path.join(folder_path, att_name)
                                    with open(save_path, 'wb') as f:
                                        f.write(attachment.data)
                                    self.log(f"Success: {att_name} in [{subdir}]")
                            msg.close()
                        except Exception as e:
                            self.log(f"Error in {file}: {str(e)}")
                folder_count += 1

            self.log(f"\n--- Finished! Processed {folder_count} folders ---")
            messagebox.showinfo("Done", f"Successfully processed {folder_count} folders.")
        
        except Exception as e:
            messagebox.showerror("Critical Error", str(e))
        finally:
            self.start_btn.config(state='normal')

if __name__ == "__main__":
    root = tk.Tk()
    app = MsgExtractorApp(root)
    root.mainloop()