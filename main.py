import customtkinter as ctk
from tkinter import messagebox

from OfficeOpeners import (
    open_word,
    open_excel,
    open_powerpoint,
    open_onenote,
    open_outlook,
)


APP_ACTIONS = {
    "Word": open_word,
    "Excel": open_excel,
    "PowerPoint": open_powerpoint,
    "OneNote": open_onenote,
    "Outlook": open_outlook,
}
APP_LABELS = list(APP_ACTIONS.keys())


class EssentialOpenerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Essential Opener")
        self.geometry("420x320")
        self.resizable(False, False)
        self.configure(fg_color="black")
        self._build_menu()
        self._build_body()

    def _build_menu(self):
        # customtkinter does not support native menu bars; use an option menu instead.
        pass

    def _build_body(self):
        container = ctk.CTkFrame(self, fg_color="black")
        container.pack(fill="both", expand=True, padx=16, pady=16)

        title = ctk.CTkLabel(
            container,
            text="Open Office Web Apps",
            font=("Segoe UI", 18, "bold"),
            text_color="white",
        )
        title.pack(pady=(0, 16))

        self.app_choice = ctk.StringVar(value=APP_LABELS[0])
        self.option_menu = ctk.CTkOptionMenu(
            container,
            values=APP_LABELS,
            variable=self.app_choice,
            fg_color="#111111",
            button_color="#1f1f1f",
            text_color="white",
        )
        self.option_menu.pack(pady=(0, 16))

        open_button = ctk.CTkButton(
            container,
            text="Open",
            fg_color="#222222",
            hover_color="#2f2f2f",
            command=self._open_selected,
        )
        open_button.pack()

    def _open_selected(self):
        selected = self.app_choice.get()
        action = APP_ACTIONS.get(selected)
        if action:
            self._open(action, selected)

    def _open(self, action, name):
        try:
            opened = action()
            if not opened:
                messagebox.showwarning("Open failed", f"Could not open {name} in your browser.")
        except Exception as exc:
            messagebox.showerror("Error", f"Failed to open {name}.\n\n{exc}")


def main():
    ctk.set_appearance_mode("dark")
    app = EssentialOpenerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
