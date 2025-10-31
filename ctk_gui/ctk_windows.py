import customtkinter as ctk
from customtkinter import filedialog
from tkcalendar import Calendar
from datetime import datetime
from textwrap import wrap

string_length = 60


class PopupError(ctk.CTk):
    def __init__(self, message):
        super().__init__()
        self.geometry("400x600")
        self.title("Error")
        self.message = "\n".join(wrap(message, string_length))

        self.label = ctk.CTkLabel(self, text=self.message)
        self.label.pack(pady=10)

        self.button = ctk.CTkButton(
            self,
            text="Error",
            command=self.ok_button_click,
            fg_color="red",
            text_color="white",
        )
        self.button.pack(pady=10)

        self.update_idletasks()
        required_width = max(self.label.winfo_reqwidth(), self.winfo_reqwidth()) + 20
        required_height = (
            self.label.winfo_reqheight() + self.button.winfo_reqheight() + 40
        )
        self.geometry(f"{required_width}x{required_height}")

        self.mainloop()

    def ok_button_click(self):
        self.destroy()


class PopupWarning(ctk.CTk):
    def __init__(self, message):
        super().__init__()
        self.geometry("400x150")
        self.title("Warning")
        self.message = "\n".join(wrap(message, string_length))

        self.label = ctk.CTkLabel(self, text=self.message)
        self.label.pack(pady=10)

        self.button = ctk.CTkButton(
            self,
            text="OK",
            command=self.ok_button_click,
            fg_color="yellow",
            text_color="black",
        )
        self.button.pack(pady=10)

        self.update_idletasks()
        required_width = max(self.label.winfo_reqwidth(), self.winfo_reqwidth()) + 100
        required_height = (
            self.label.winfo_reqheight() + self.button.winfo_reqheight() + 60
        )
        self.geometry(f"{required_width}x{required_height}")

        self.mainloop()

    def ok_button_click(self):
        self.destroy()


def PopupYesNo(message="Are you sure?", title="Popup", master=None):
    if master is None:
        master = ctk.CTk()
        master.withdraw()

    dialog = ctk.CTkToplevel(master)
    dialog.title(title)
    dialog.geometry("400x250")
    dialog.resizable(True, True)

    # Ensure the dialog is rendered before grabbing
    dialog.update()
    dialog.grab_set()

    label = ctk.CTkLabel(master=dialog, text="\n".join(wrap(message, string_length)))
    label.pack(pady=20)

    result = [None]

    def on_yes():
        result[0] = "Yes"
        dialog.destroy()

    def on_no():
        result[0] = "No"
        dialog.destroy()

    yes_button = ctk.CTkButton(master=dialog, text="Yes", command=on_yes)
    yes_button.pack(side="left", padx=20, pady=10)

    no_button = ctk.CTkButton(master=dialog, text="No", command=on_no)
    no_button.pack(side="right", padx=20, pady=10)

    dialog.protocol("WM_DELETE_WINDOW", on_no)
    dialog.wait_window()
    if master.master is None:
        master.destroy()
    return result[0]


class PopupGetText(ctk.CTk):
    def __init__(self, title="Text prompt", width=300, height=200):
        super().__init__()
        self.geometry(f"{width}x{height}")
        self.title(title)
        self.result = None

        self.input = ctk.CTkTextbox(self, width=width, height=height / 2)
        self.input.pack(pady=10)

        self.button = ctk.CTkButton(self, text="OK", command=self.ok_button_click)
        self.button.pack()

        self.update_idletasks()
        required_width = max(self.input.winfo_reqwidth(), self.winfo_reqwidth()) + 20
        required_height = (
            self.input.winfo_reqheight() + self.button.winfo_reqheight() + 40
        )
        self.geometry(f"{required_width}x{required_height}")

        self.mainloop()

    def ok_button_click(self):
        self.result = self.input.get(0.0, ctk.END).strip()
        self.destroy()

    def return_value(self):
        return self.result


class PopupGetDate:
    def __init__(self, title="Select Date"):
        self.root = ctk.CTk()
        self.root.withdraw()

        self.popup = ctk.CTkToplevel()
        self.popup.title(title)
        self.popup.geometry("300x350")
        self.popup.resizable(False, False)

        self.selected_date = None

        current_date = datetime.now()

        self.calendar = Calendar(
            self.popup,
            selectmode="day",
            year=current_date.year,
            month=current_date.month,
            day=current_date.day,
            font=("Arial", 12),
            background="#2b2b2b",
            foreground="white",
            selectbackground="#1a73e8",
            normalbackground="#2b2b2b",
            weekendbackground="#2b2b2b",
            othermonthbackground="#2b2b2b",
            othermonthwebackground="#2b2b2b",
            date_pattern="yyyy-mm-dd",  # Set the desired date format here
        )
        self.calendar.pack(pady=20, padx=20)

        self.confirm_button = ctk.CTkButton(
            self.popup, text="Confirm", command=self.confirm_date
        )
        self.confirm_button.pack(pady=10)

        # Make window modal
        self.popup.grab_set()

        # Center the window
        self.popup.update_idletasks()
        width = self.popup.winfo_width()
        height = self.popup.winfo_height()
        x = (self.popup.winfo_screenwidth() // 2) - (width // 2)
        y = (self.popup.winfo_screenheight() // 2) - (height // 2)
        self.popup.geometry(f"+{x}+{y}")

        # Start the event loop and get the result
        self.popup.wait_window()
        self.root.destroy()

    def confirm_date(self):
        """Store the selected date and close the window"""
        self.selected_date = self.calendar.get_date()
        self.popup.destroy()

    def get_date(self):
        return self.selected_date
