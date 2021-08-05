"""
Gui.py:
This gui module starts the CA application and is located in the root directory of the repository
(ca) to accommodate how Pyinstaller wants the directory structure to be ordered. All other gui modules are
located in HEM4/com/sca/ca/gui. Pyinstaller wants the initial module to be located in the root.
"""
import tkinter as tk
from com.sca.ca.gui.MainView import MainView


def on_closing(hem):

    if hem.running == True:

        hem.quit_app()
        if hem.aborted == True:
            root.destroy()

    else:
        root.destroy()


# infinite loop which is required to
# run tkinter program infinitely
# until an interrupt occurs
if __name__ == "__main__":
    root = tk.Tk()
    w, h = root.winfo_screenwidth(), root.winfo_screenheight()
    root.tk.call('wm', 'iconphoto', root._w, tk.PhotoImage(file='images/people-icon.png'))
    root.title("")

    main = MainView(root)
    main.pack(side="top", fill="both", expand=True)
    # root.protocol("WM_DELETE_WINDOW", lambda: on_closing(main.hem))

    root.wm_minsize(400, 300)
    root.mainloop()
