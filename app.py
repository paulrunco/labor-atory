from tkinter import Label, Entry, ttk, StringVar, Button, Tk, END, filedialog
from tkinter import messagebox as mb
from datetime import datetime

import functions.chart_order as chart_order
import functions.summarize_orders as summarize_orders


class App:
    def __init__(self, master):

        version = "Version 0.2"
        tagline = "PRunco"

        self.master = master
        master.title("Labor Analysis Utility")

        ## Entry and Processing Parameter Section ##
        # Excel Labor report entry heading
        self.report_entry_heading = Label(master, text="Browse for Labor Report")
        self.report_entry_heading.grid(
            row=1, column=0, sticky="e", padx=(40, 5), pady=(10, 0)
        )
        # Cutfile Entry Box
        self.report_name = Entry(master, width=80, textvariable=StringVar)
        self.report_name.grid(row=1, column=1, sticky="w", padx=5, pady=(20, 0))
        # Cutfile Browse Button
        self.report_browse_button = Button(
            master, text="Browse", command=lambda: self.browse_for("report")
        )
        self.report_browse_button.grid(
            row=1, column=2, padx=(5, 40), sticky="w", pady=(20, 0)
        )

        # separator
        self.separator = ttk.Separator(master, orient="horizontal")
        self.separator.grid(row=2, sticky="ew", columnspan=3, padx=30, pady=20)

        ## Parent order Entry
        # parent order heading
        self.order_entry_heading = Label(master, text="Parent Order to Chart")
        self.order_entry_heading.grid(row=3, column=0, sticky="e", padx=5, pady=(10, 0))
        # parent order entry
        self.order_entry = Entry(master, width=40, textvariable=StringVar)
        self.order_entry.grid(row=3, column=1, sticky="w", padx=5, pady=(10, 0))
        # Chart Button
        self.process_button = Button(
            master,
            width=20,
            text="Chart",
            fg="green",
            command=lambda: self.onClickChartOrder(),
        )
        self.process_button.grid(
            row=3, column=1, columnspan=2, sticky="e", padx=(5, 40), pady=(10, 0)
        )

        # separator
        self.separator = ttk.Separator(master, orient="horizontal")
        self.separator.grid(row=5, column=0, columnspan=3, sticky="ew", padx=30, pady=20)

        ## Parent Order Summary
        # summary button label
        self.summary_button_label = Label(master, text="Generate Order Summary")
        self.summary_button_label.grid(row=6, column=0, padx=5, pady=5, sticky="e")
        # summary button
        self.summary_button = Button(
            master,
            width=20,
            text="Summarize",
            fg="green",
            command=lambda: self.onClickSummarizeOrders(),
        )
        self.summary_button.grid(
            row=6, column=1, columnspan=2, sticky="ew", padx=(5, 40), pady=5
        )

        # separator
        self.separator = ttk.Separator(master, orient="horizontal")
        self.separator.grid(row=8, column=0, columnspan=3, sticky="ew", pady=(20, 0))

        ## Progress Bar
        self.pb = ttk.Progressbar(
            master, orient="horizontal", mode="indeterminate", length=200
        )

        ## Status Message
        self.status = Label(master, text="{} | {}".format(version, tagline), fg="black")
        self.status.grid(row=10, column=0, columnspan=3, pady=2)

    def show_progress_bar(self):
        self.pb.grid(row=9, column=1)

    def hide_progress_bar(self):
        self.pb.grid_forget()

    def browse_for(self, target):
        if target == "report":
            self.report_file_name = filedialog.askopenfilename(
                filetypes=(("Excel files", "*xlsx"), ("All files", "*"))
            )
            self.report_name.delete(0, END)
            self.report_name.insert(0, self.report_file_name)

        else:
            mb.showerror("Error", "No path or order specified")

    def ask_save_as(self):
        today = datetime.today()
        self.order_summary_file_name = filedialog.asksaveasfilename(
            title="Save as",
            filetypes=(("Excel files", "*xlsx"), ("All files", "*")),
            defaultextension="xlsx",
            initialfile="{}{}{} Parent Order Summary".format(today.year, today.month, today.day)  
        )
        if self.order_summary_file_name:
            return self.order_summary_file_name
        else:
            pass

    def onClickChartOrder(self):
        parent_order_number = self.order_entry.get()
        path_to_file = self.report_name.get()
        if parent_order_number and path_to_file:
            chart_order.chart_order(path_to_file, parent_order_number)

    def onClickSummarizeOrders(self):
        path_to_file = self.report_name.get()
        if path_to_file:
            save_as = self.ask_save_as()
            if save_as:
                summarize_orders.summarize_orders(path_to_file, save_as)


root = Tk()
root.resizable(False, False)
my_app = App(root)
root.mainloop()
