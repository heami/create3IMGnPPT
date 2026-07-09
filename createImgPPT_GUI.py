import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from PIL import Image, ImageTk
import threading
import sys
import os

from create3IMGnPPT import create_regular_images, create_ppts, dispatch_image_creation
import createImgPPT_config as cfg
import update_config as upcfg

# 콘솔 출력을 Text 위젯으로 리디렉션하는 클래스
class TextRedirector:
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, str_text):
        # Tkinter 위젯은 메인 스레드에서만 업데이트할 수 있습니다.
        # 따라서 after()를 사용하여 업데이트를 예약합니다.
        self.widget.after(0, self._insert_text, str_text)

    def _insert_text(self, str_text):
        self.widget.configure(state="normal")
        self.widget.insert(tk.END, str_text, (self.tag,))
        self.widget.see(tk.END)  # 스크롤을 항상 맨 아래로 이동
        self.widget.configure(state="disabled")

    def flush(self):
        # 파이썬 3에서 sys.stdout.flush() 호출을 위해 필요합니다.
        pass

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Image&PPT Maker")
        self.geometry("370x500")

        # 창 아이콘 설정
        try:
            # PyInstaller 환경인지 확인하고, 기본 경로를 설정합니다.
            base_path = sys._MEIPASS
        except AttributeError:
            base_path = os.path.abspath(".")

        ico_path = os.path.join(base_path, '3imgppt.ico')

        try:
            # icon.ico 파일을 아이콘으로 설정합니다.
            self.iconbitmap(ico_path)
        except tk.TclError:
            # 파일이 없거나 오류가 발생한 경우를 대비
            print(f"Error: Could not set icon from {ico_path}")
            pass

        # [수정] 실행창 전체의 배경색을 변경합니다.
        self.configure(bg='#DDE9EF') # 또는 self['bg'] = 'lightgray'

        # 원래의 stdout을 저장합니다.
        self.original_stdout = sys.stdout
        # 창이 닫힐 때 원본 stdout을 복원하도록 설정
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 위젯에 사용될 이미지를 생성
        try:
            # PyInstaller로 패키징된 경우
            base_path = sys._MEIPASS
        except AttributeError:
            # 소스 코드로 실행하는 경우
            base_path = os.path.abspath(".")

        image_path = os.path.join(base_path, 'folder_icon.png')

        try:
            original_image = Image.open(image_path)
            self.folder_icon = ctk.CTkImage(light_image=original_image,
                                             dark_image=original_image,
                                             size=(16, 16))
        except FileNotFoundError:
            # 이미지 파일을 찾지 못한 경우 오류를 처리합니다.
            print(f"Error: Could not find image file at {image_path}")
            self.folder_icon = None  # 이미지를 찾지 못하면 None으로 설정

        self.create_menubar()  # 메뉴바 생성 메서드 호출
        self.create_widgets()

        # stdout을 GUI의 출력창으로 리디렉션합니다.
        sys.stdout = TextRedirector(self.output_text, "stdout")

        # 앱 시작 시 업데이트 확인 로직을 추가합니다.
        self.check_for_updates_gui()

    def create_menubar(self):
        menubar = tk.Menu(self)
        self.config(menu=menubar)

        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Show Help", command=self.show_help)

    def show_help(self):
        help_text = """
[사용 방법]

1. 이미지 생성
   - 이미지 또는 Excel 파일이 있는 디렉토리 주소를 입력합니다.
   - "Only IMG" 버튼을 클릭합니다.
   - 입력 디렉토리 내의 이미지 또는 Excel 파일을 사용해 -l.jpg, -m.jpg 파일이 생성됩니다.

2. PPT 생성
   - "이미지가 있는 디렉토리 주소"와 "논문 또는 ToC URL"을 입력합니다.
   - "Only PPT" 버튼을 클릭합니다.
   - 입력 디렉토리 내의 이미지를 사용해 PPT가 생성됩니다.

3. 모두 생성
   - "이미지가 있는 디렉토리 주소"와 "논문 또는 ToC URL"을 입력합니다.
   - "ALL" 버튼을 클릭합니다.
   - 이미지 생성 후, Images 폴더의 이미지로 PPT가 생성됩니다.

* 작업 내용은 하단 출력창에서 확인 가능합니다.
        """
        messagebox.showinfo("Help", help_text)


    # def check_for_updates_gui(self):
    #
    #     def run_update_check():
    #         print("업데이트를 확인하는 중입니다...")
    #         try:
    #             # upcfg.check_for_updates() 함수가 (True/False, 최신버전, 다운로드URL) 튜플을 반환한다고 가정
    #             is_update_needed, latest_version, download_url, update_notes = upcfg.check_for_updates()
    #
    #             if is_update_needed:
    #                 # self.after에 함수와 전달할 인자들을 순서대로 나열합니다.
    #                 self.after(0, self.show_update_dialog, latest_version, download_url, update_notes)
    #             else:
    #                 # print 함수도 직접 전달할 수 있습니다.
    #                 self.after(0, print, "현재 최신 버전을 사용 중입니다.")
    #         except Exception as e:
    #             # f-string으로 생성된 문자열 자체를 인자로 전달합니다.
    #             self.after(0, print, f"업데이트 확인 중 오류가 발생했습니다: {e}")
    #
    #     thread = threading.Thread(target=run_update_check)
    #     thread.start()

    def check_for_updates_gui(self):
        """GUI가 멈추지 않도록 별도의 스레드에서 업데이트 확인을 시작하는 함수"""
        # 데몬 스레드로 만들어 메인 프로그램 종료 시 함께 종료되도록 함
        thread = threading.Thread(target=self._background_update_check, daemon=True)
        thread.start()

    def _background_update_check(self):
        """[백그라운드 스레드] 실제 네트워크 작업을 수행하는 별도의 메소드"""
        print("업데이트를 확인하는 중입니다...")
        try:
            # upcfg.py의 함수를 호출하여 결과를 받음
            result = upcfg.check_for_updates()
            # 작업이 끝나면, 결과를 가지고 GUI 스레드에서 실행될 함수를 예약
            self.after(0, self._process_update_result, result)
        except Exception as e:
            # 오류가 발생해도 GUI 스레드에서 처리하도록 예약
            self.after(0, self._process_update_error, e)

    def _process_update_result(self, result):
        """[GUI 스레드] 백그라운드 작업의 결과를 받아 안전하게 GUI를 업데이트"""
        is_update_needed, latest_version, download_url, update_notes = result
        if is_update_needed:
            # 이 함수는 self.after를 통해 호출되었으므로 lambda가 필요 없습니다.
            self.show_update_dialog(latest_version, download_url, update_notes)
        else:
            print("현재 최신 버전을 사용 중입니다.")

    def _process_update_error(self, error):
        """[GUI 스레드] 백그라운드 작업에서 발생한 오류를 처리"""
        print(f"업데이트 확인 중 오류가 발생했습니다: {error}")

    # [추가] 업데이트 진행 여부를 묻는 대화상자 함수
    def show_update_dialog(self, latest_version, download_url, update_notes):
        # 업데이트 노트가 비어있지 않다면, 제목과 함께 내용을 추가합니다.
        if update_notes and update_notes.strip():
            changelog = f"\n\n[업데이트 내용]\n{update_notes.strip()}"
        else:
            changelog = "" # 업데이트 노트가 없으면 아무것도 추가하지 않음

        message = (
            f"새로운 버전({latest_version})이 있습니다.\n"
            f"url: {download_url}\n"
            f"지금 업데이트하시겠습니까?"
            f"{changelog}"  # 구성된 업데이트 노트를 메시지 끝에 추가
        )

        response = messagebox.askyesno(
            "업데이트 알림",
            message
        )
        if response:
            # 사용자가 '예'를 선택한 경우, 업데이트 진행
            self.perform_update(latest_version, download_url)
        else:
            # 사용자가 '아니오'를 선택한 경우
            print("업데이트를 취소했습니다.")

    # [추가] 실제 업데이트를 수행하는 함수
    def perform_update(self, latest_version, download_url):
        print(f"새로운 버전({latest_version})을 다운로드하고 설치합니다... downloadURL: {download_url}")
        try:
            # upcfg.download_and_replace 함수를 호출하여 업데이트 실행
            upcfg.download_and_replace(download_url)
            print("업데이트가 완료되었습니다. 프로그램을 다시 시작해주세요.")
            messagebox.showinfo("업데이트 완료", "업데이트가 완료되었습니다. 프로그램을 다시 시작해주세요.")
            # 프로그램 재시작 또는 종료 로직 추가
        except Exception as e:
            print(f"업데이트 중 오류가 발생했습니다: {e}")
            messagebox.showerror("업데이트 오류", f"업데이트 중 오류가 발생했습니다: {e}")

    # 재활용 가능한 포커스 이벤트 핸들러 메서드
    def on_entry_focus_in(self, event):
        widget = event.widget
        if widget.get() == widget.placeholder_text:
            widget.delete(0, tk.END)
            widget.config(fg='black')

    def on_entry_focus_out(self, event):
        widget = event.widget
        if not widget.get():
            widget.insert(0, widget.placeholder_text)
            widget.config(fg='grey')

    def create_widgets(self):
        # 'self.frame'을 tkinter.Frame 객체로 생성합니다.
        self.frame = tk.Frame(self, padx=10, pady=10, bg='#DDE9EF')
        self.frame.pack(pady=(10, 20))

        # 라벨과 엔트리, 버튼을 담을 서브 프레임
        input_sub_frame = tk.Frame(self.frame, bg='#DDE9EF')
        input_sub_frame.pack(fill=tk.X)

        # 서브 프레임 내 왼쪽(W) 정렬
        tk.Label(input_sub_frame, text="** Folder Path", font=('Noto Sans', 10, 'bold'), bg='#DDE9EF').grid(row=0, column=0, sticky=tk.W)

        # 엔트리와 버튼을 한 줄에 배치
        self.dir_entry = tk.Entry(input_sub_frame, width=40)
        self.dir_entry.grid(row=1, column=0, padx=(0, 5), sticky=tk.EW)

        # 예제 문구를 삽입하고, 이벤트 핸들러를 바인딩합니다.
        self.dir_entry.placeholder_text = "ex) Y:/0063JKMS/2025/XMLinkPress/8월호/jkms-40-31-20250811/JKMS-40-e224-2025-0979/InDesign/image"
        self.dir_entry.insert(0, self.dir_entry.placeholder_text)
        self.dir_entry.config(fg='grey', font=('Arial', 10))

        self.dir_entry.bind("<FocusIn>", self.on_entry_focus_in)
        self.dir_entry.bind("<FocusOut>", self.on_entry_focus_out)

        ctk.CTkButton(
            master=input_sub_frame,
            text="Open",
            command=self.browse_dir,
            fg_color="#707070", # 배경색
            hover_color="#707070",
            text_color="#fff",  # 글자색
            font = ('Noto Sans', 12, 'bold'),
            image=self.folder_icon,
            compound=tk.LEFT,
            width=70,
            corner_radius=3,  # 모서리 둥글기 추가 (선택사항)
        ).grid(row=1, column=1, padx=5, sticky=tk.E)

        # 첫 번째 열이 나머지 공간을 모두 차지하도록 설정
        input_sub_frame.columnconfigure(0, weight=1)

        # 라벨과 엔트리를 담을 서브 프레임
        url_sub_frame = tk.Frame(self.frame, bg='#DDE9EF')
        url_sub_frame.pack(fill=tk.X, pady=(10, 0))

        # 라벨: 서브 프레임 내 왼쪽(W) 정렬
        tk.Label(url_sub_frame, text="** URL", font=('Noto Sans', 10, 'bold'), bg='#DDE9EF').grid(row=0, column=0, sticky=tk.W,
                                                                                      pady=5)
        # 엔트리: 서브 프레임의 가로 너비를 모두 사용
        self.url_entry = tk.Entry(url_sub_frame, width=50)
        self.url_entry.grid(row=1, column=0, sticky=tk.EW, columnspan=2)

        # 예제 문구를 삽입하고, 이벤트 핸들러를 바인딩합니다.
        self.url_entry.placeholder_text = "ex) https://jkms.org/DOIx.php?id=10.3346/jkms.2025.40.e224"
        self.url_entry.insert(0, self.url_entry.placeholder_text)
        self.url_entry.config(fg='grey', font=('Arial', 10))

        self.url_entry.bind("<FocusIn>", self.on_entry_focus_in)
        self.url_entry.bind("<FocusOut>", self.on_entry_focus_out)

        url_sub_frame.columnconfigure(0, weight=1)

        # 생성 버튼 프레임
        button_frame = tk.Frame(self, padx=10, pady=10, bg='#DDE9EF')
        button_frame.pack(fill=tk.X)

        # 버튼을 담는 프레임
        btn_container = tk.Frame(button_frame, bg='#DDE9EF')
        btn_container.pack(pady=(0, 10), anchor='center')

        ctk.CTkButton(master=btn_container, text="Image", command=self.run_create_images, width=90, corner_radius=8, fg_color="#CF90C1", hover_color="#CF90C1", text_color="#fff", font=('Noto Sans', 12, 'bold')).pack(side=tk.LEFT, padx=5, pady=5)
        ctk.CTkButton(master=btn_container, text="PPT", command=self.run_create_ppts, width=90, corner_radius=8, fg_color="#8D80B0", hover_color="#8D80B0", text_color="#fff", font=('Noto Sans', 12, 'bold')).pack(side=tk.LEFT, padx=5, pady=5)
        ctk.CTkButton(master=btn_container, text="ALL", command=self.run_all, width=90, corner_radius=8, fg_color="#8397C4", hover_color="#8397C4", text_color="#fff", font=('Noto Sans', 12, 'bold')).pack(side=tk.LEFT, padx=10, pady=5)

        self.status_label = tk.Label(self, text="Ready", fg="#030c8c", font=('Noto Sans', 10), bg='#DDE9EF')
        self.status_label.pack(pady=(10, 5))

        # 출력창 (Text 위젯) 프레임
        output_frame = tk.Frame(self, padx=10, pady=10, bg='#DDE9EF')
        output_frame.pack(fill=tk.BOTH, expand=True)

        self.output_text = tk.Text(
            output_frame,
            height=10,
            state='disabled',
            wrap='word',
            # [수정] font 옵션을 추가하여 폰트와 크기를 지정합니다.
            font=('Noto Sans KR', 10)
        )
        self.output_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(output_frame, command=self.output_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.output_text.config(yscrollcommand=scrollbar.set)


    def on_closing(self):
        # 창이 닫힐 때 원본 stdout을 복원합니다.
        sys.stdout = self.original_stdout
        self.destroy()

    def browse_dir(self):
        directory = filedialog.askdirectory()
        if directory:
            self.dir_entry.delete(0, tk.END)
            self.dir_entry.insert(0, directory)

    def run_task(self, task_function, *args):
        # UI가 멈추지 않도록 스레드를 사용하여 백그라운드에서 작업 실행
        self.status_label.config(text="Processing...", font=('Noto Sans', 10), fg="#ff4800")

        # 출력창 초기화
        self.output_text.configure(state="normal")
        self.output_text.delete(1.0, tk.END)
        self.output_text.configure(state="disabled")

        def run():
            try:
                task_function(*args)
                self.status_label.config(text="Success!", font=('Noto Sans', 10), fg="#005f24")
                # messagebox.showinfo("완료", "작업이 성공적으로 완료되었습니다.")
            except Exception as e:
                self.status_label.config(text="Error!", font=('Noto Sans', 10), fg="#aa0000")
                # messagebox.showerror("오류", f"작업 중 오류가 발생했습니다: {e}")

        thread = threading.Thread(target=run)
        thread.start()

    def run_create_images(self):
        directory = self.dir_entry.get()
        if not directory:
            messagebox.showerror("Error", "디렉토리 주소를 입력해주세요.")
            return
        self.run_task(dispatch_image_creation, directory)

    def run_create_ppts(self):
        directory = self.dir_entry.get()
        url = self.url_entry.get()
        if not directory or not url:
            messagebox.showerror("Error", "디렉토리와 URL을 모두 입력해주세요.")
            return
        # 기존 로직에 맞춰 articleURLs와 webURL을 생성
        webURL = url.split('/')[0] + '//' + url.split('/')[2] + '/'
        articleURLs = cfg.decide_URL(url, webURL)
        self.run_task(create_ppts, directory, articleURLs, webURL, False)

    def run_all(self):
        directory = self.dir_entry.get()
        url = self.url_entry.get()
        if not directory or not url:
            messagebox.showerror("Error", "디렉토리와 URL을 모두 입력해주세요.")
            return

        def all_tasks():
            self.status_label.config(text="Creating Images...", fg="#ff4800")
            if create_regular_images(directory):
                self.status_label.config(text="Creating PPT...", fg="#ff4800")
                webURL = url.split('/')[0] + '//' + url.split('/')[2] + '/'
                articleURLs = cfg.decide_URL(url, webURL)
                create_ppts(directory, articleURLs, webURL, True)
            else:
                raise Exception("Fail")

        self.run_task(all_tasks)


if __name__ == "__main__":
    app = App()
    app.mainloop()