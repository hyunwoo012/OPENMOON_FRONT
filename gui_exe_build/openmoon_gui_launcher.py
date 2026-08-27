from __future__ import annotations

import os
import sys
import threading
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import messagebox


HOST = "127.0.0.1"
PORT = 8000
WEB_URL = f"http://{HOST}:{PORT}"


def application_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def load_external_env() -> None:
    """Load an editable .env placed next to the EXE without bundling secrets."""
    env_path = application_dir() / ".env"
    if not env_path.exists():
        return
    for raw_line in env_path.read_text(encoding="utf-8-sig").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


class OpenMoonLauncher(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("YullinMoon AI 서버 실행기")
        self.geometry("430x270")
        self.resizable(False, False)
        self.configure(bg="#f4f1ea")
        self.protocol("WM_DELETE_WINDOW", self.close_application)

        self.server = None
        self.server_thread: threading.Thread | None = None
        self.closing = False

        tk.Label(
            self,
            text="YullinMoon AI",
            font=("Malgun Gothic", 22, "bold"),
            fg="#252a2e",
            bg="#f4f1ea",
        ).pack(pady=(30, 5))
        tk.Label(
            self,
            text="견적 업무 보조 서버",
            font=("Malgun Gothic", 11),
            fg="#77736d",
            bg="#f4f1ea",
        ).pack()

        self.status = tk.Label(
            self,
            text="서버가 꺼져 있습니다.",
            font=("Malgun Gothic", 10),
            fg="#8a4c32",
            bg="#f4f1ea",
        )
        self.status.pack(pady=(24, 15))

        self.button_frame = tk.Frame(self, bg="#f4f1ea")
        self.button_frame.pack()
        self.show_stopped_controls()

    def clear_buttons(self) -> None:
        for widget in self.button_frame.winfo_children():
            widget.destroy()

    def make_button(self, text: str, command, color: str, width: int = 14) -> tk.Button:
        return tk.Button(
            self.button_frame,
            text=text,
            command=command,
            width=width,
            height=2,
            relief="flat",
            cursor="hand2",
            font=("Malgun Gothic", 10, "bold"),
            fg="white",
            bg=color,
            activeforeground="white",
            activebackground=color,
        )

    def show_stopped_controls(self) -> None:
        self.clear_buttons()
        self.status.configure(text="서버가 꺼져 있습니다.", fg="#8a4c32")
        self.make_button("서버 실행", self.start_server, "#d26832", 18).pack()

    def show_starting_controls(self) -> None:
        self.clear_buttons()
        self.status.configure(text="서버를 시작하고 있습니다...", fg="#a66a2c")
        button = self.make_button("서버 시작 중...", lambda: None, "#aaa49b", 18)
        button.configure(state="disabled")
        button.pack()

    def show_running_controls(self) -> None:
        self.clear_buttons()
        self.status.configure(text="서버가 실행 중입니다.", fg="#34714a")
        self.make_button("웹 열기", self.open_web, "#34714a").pack(side="left", padx=6)
        self.make_button("서버 끄기", self.stop_server, "#a7433d").pack(side="left", padx=6)

    def start_server(self) -> None:
        if self.server_thread and self.server_thread.is_alive():
            return
        (application_dir() / "backend" / "data" / "quotation_files").mkdir(
            parents=True,
            exist_ok=True,
        )
        self.show_starting_controls()
        self.server_thread = threading.Thread(target=self._run_server, daemon=True)
        self.server_thread.start()
        self.after(150, self._poll_server_started)

    def _run_server(self) -> None:
        try:
            load_external_env()
            import uvicorn

            config = uvicorn.Config(
                "backend.app.main:app",
                host=HOST,
                port=PORT,
                reload=False,
                access_log=False,
                log_level="warning",
                log_config=None,
            )
            self.server = uvicorn.Server(config)
            self.server.run()
        except Exception as error:
            self.after(0, lambda: self._show_server_error(str(error)))

    def _poll_server_started(self) -> None:
        if self.closing:
            return
        if self.server and self.server.started:
            self.show_running_controls()
            return
        if self.server_thread and self.server_thread.is_alive():
            self.after(150, self._poll_server_started)
            return
        self.show_stopped_controls()

    def _show_server_error(self, error: str) -> None:
        self.show_stopped_controls()
        messagebox.showerror("서버 실행 실패", error)

    def open_web(self) -> None:
        webbrowser.open(WEB_URL)

    def stop_server(self) -> None:
        if not self.server or not self.server_thread or not self.server_thread.is_alive():
            self.show_stopped_controls()
            return
        self.clear_buttons()
        self.status.configure(text="서버를 종료하고 있습니다...", fg="#8a4c32")
        self.server.should_exit = True
        self.after(150, self._poll_server_stopped)

    def _poll_server_stopped(self) -> None:
        if self.server_thread and self.server_thread.is_alive():
            self.after(150, self._poll_server_stopped)
            return
        self.server = None
        self.server_thread = None
        if self.closing:
            self.destroy()
        else:
            self.show_stopped_controls()

    def close_application(self) -> None:
        self.closing = True
        if self.server and self.server_thread and self.server_thread.is_alive():
            self.server.should_exit = True
            self.status.configure(text="프로그램을 종료하고 있습니다...", fg="#8a4c32")
            self.clear_buttons()
            self.after(100, self._poll_server_stopped)
        else:
            self.destroy()


if __name__ == "__main__":
    OpenMoonLauncher().mainloop()
