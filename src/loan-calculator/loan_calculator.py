import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from datetime import datetime
import csv
import matplotlib.pyplot as plt

class LoanCalculator:
    def __init__(self, master):
        self.master = master
        master.title("Loan Payoff Calculator")

        # Create input fields in the same row
        self.create_input_fields_in_row()

        # Create buttons (Calculate, Clear, Export, and Graph)
        self.create_buttons()

        # Apply Treeview style for column separators
        style = ttk.Style()
        style.configure("Treeview", rowheight=25, font=('Helvetica', 10))  # Adjust row height and font
        style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])  # Remove borders

        # Add a treeview with columns
        self.result_tree = ttk.Treeview(master, columns=("Date", "Opening Balance", "Payment", "Interest", "Principal", "Closing Balance", "Cumulative Interest", "Cumulative Principal"), show='headings', style="Treeview")
        self.result_tree.heading("Date", text="Date")
        self.result_tree.heading("Opening Balance", text="Opening Balance")
        self.result_tree.heading("Payment", text="Payment")
        self.result_tree.heading("Interest", text="Interest")
        self.result_tree.heading("Principal", text="Principal")
        self.result_tree.heading("Closing Balance", text="Closing Balance")
        self.result_tree.heading("Cumulative Interest", text="Cumulative Interest")
        self.result_tree.heading("Cumulative Principal", text="Cumulative Principal")

        # Configure column width and alignment (center align headers and numbers)
        self.result_tree.column("Date", anchor="center", width=100)
        self.result_tree.column("Opening Balance", anchor="center", width=100)  # Right-aligned
        self.result_tree.column("Payment", anchor="center", width=100)  # Right-aligned
        self.result_tree.column("Interest", anchor="center", width=100)  # Right-aligned
        self.result_tree.column("Principal", anchor="center", width=100)  # Right-aligned
        self.result_tree.column("Closing Balance", anchor="center", width=150)  # Right-aligned
        self.result_tree.column("Cumulative Interest", anchor="center", width=150)  # Right-aligned
        self.result_tree.column("Cumulative Principal", anchor="center", width=150)  # Right-aligned


        # Apply alternating row colors
        self.result_tree.tag_configure('evenrow', background='#F5F5F5')  # Light gray
        self.result_tree.tag_configure('oddrow', background='#FFFFFF')   # White
        self.result_tree.tag_configure('totals', background='#44e35c', font=('Helvetica', 12, 'bold'))

        self.result_tree.grid(row=7, column=0, columnspan=2, padx=5, pady=5, sticky='nsew')

        # Add scrollbars
        self.scrollbar = ttk.Scrollbar(master, orient="vertical", command=self.result_tree.yview)
        self.scrollbar.grid(row=7, column=2, sticky='ns')
        self.result_tree.configure(yscrollcommand=self.scrollbar.set)

        # Configure grid weights for resizing
        master.grid_rowconfigure(7, weight=1)
        master.grid_columnconfigure(1, weight=1)

        # Pre-fill current month and year
        now = datetime.now()
        self.entries[3].insert(0, str(now.month))  # Pre-fill start month
        self.entries[4].insert(0, str(now.year))   # Pre-fill start year

    def create_input_fields_in_row(self):
        labels = [
            ("Loan Balance ($):", "3494.83"), 
            ("Annual Interest Rate (%):", "10.00"), 
            ("Monthly Payment ($):", "155.16"), 
            ("Start Month (1-12):", ""),  # Default is current month
            ("Start Year:", "")  # Default is current year
        ]
        self.entries = []
        
        # Input fields all in one row
        row_frame = tk.Frame(self.master)
        row_frame.grid(row=0, column=0, columnspan=5, padx=5, pady=5, sticky='w')

        # Create each label and entry field on the same row
        for i, (label_text, default_value) in enumerate(labels):
            label = tk.Label(row_frame, text=label_text)
            label.grid(row=0, column=2*i, padx=5, pady=5, sticky='w')
            
            entry = tk.Entry(row_frame, width=12)
            entry.grid(row=0, column=2*i+1, padx=5, pady=5, sticky='w')
            if default_value:  # Only set default values if provided
                entry.insert(0, default_value)
            self.entries.append(entry)

    def create_buttons(self):
        # Create a frame to hold the buttons and time label for better alignment
        button_frame = tk.Frame(self.master)
        button_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky='w')

        # Create the calculate button inside the frame
        self.calculate_button = tk.Button(button_frame, text="Calculate Payoff Schedule", command=self.calculate_schedule, bg="#008000", fg="white", font=('Helvetica', 12, 'bold'), padx=10, pady=5)
        self.calculate_button.pack(side=tk.LEFT)

        # Label for displaying months and years, aligned next to the button inside the same frame
        self.time_label = tk.Label(button_frame, text="", bg="white", fg="black", font=('Helvetica', 12, 'bold'), padx=10, pady=5)
        self.time_label.pack(side=tk.LEFT, padx=10)

        # Add interest summary label (not button)
        self.show_total_interest = tk.Label(button_frame, text="", bg="white", fg="black", font=('Helvetica', 12, 'bold'), padx=10, pady=5)
        self.show_total_interest.pack(side=tk.LEFT)

        # Add clear button
        self.clear_button = tk.Button(button_frame, text="Clear", command=self.clear_inputs, bg="red", fg="white")
        self.clear_button.pack(side=tk.LEFT, padx=10)

        # Add export button
        self.export_button = tk.Button(button_frame, text="Export CSV", command=self.export_schedule_to_csv, bg="green", fg="white")
        self.export_button.pack(side=tk.LEFT, padx=10)

        # Add show graph button
        self.graph_button = tk.Button(button_frame, text="Show Graph", command=self.show_payoff_graph, bg="green", fg="white")
        self.graph_button.pack(side=tk.LEFT, padx=10)

    def calculate_schedule(self):
        # Input validation
        try:
            loan_balance = float(self.entries[0].get())
            annual_interest_rate = float(self.entries[1].get())
            monthly_payment = float(self.entries[2].get())
            start_month = int(self.entries[3].get())
            start_year = int(self.entries[4].get())

            if loan_balance <= 0 or annual_interest_rate < 0 or monthly_payment <= 0:
                raise ValueError
            if not (1 <= start_month <= 12):
                raise ValueError
        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid positive numbers for loan balance, interest rate, and payment. Ensure month is between 1 and 12.")
            return

        # Calculate the loan payoff schedule
        schedule, months = self.loan_payoff_schedule(loan_balance, annual_interest_rate, monthly_payment, start_month, start_year)

        # Clear existing table
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        # Add rows to the table with alternating colors
        for idx, row in enumerate(schedule[:-1]):  # Add rows except the totals row
            tags = ('evenrow',) if idx % 2 == 0 else ('oddrow',)
            self.result_tree.insert("", "end", values=row, tags=tags)

        # Insert the totals row with bold text and background color
        totals_row = schedule[-1]
        self.result_tree.insert("", "end", values=totals_row, tags=('totals',))

        # Update the time label with the total number of months and years
        years = months // 12
        remaining_months = months % 12
        self.time_label.config(text=f"Payoff Time:  {months} months  |  {round(months/12, 2)} years ")

        # Update total interest label
        total_interest = self.get_total_interest()
        self.show_total_interest.config(text=f"Total Interest: ${total_interest:,.2f}")

    def loan_payoff_schedule(self, loan_balance, annual_interest_rate, monthly_payment, start_month, start_year):
        monthly_interest_rate = round((annual_interest_rate / 100 / 365) * 30, 8) if annual_interest_rate > 0 else 0
        month = 0
        schedule_list = []
        
        total_interest = 0
        cumulative_interest = 0
        cumulative_principal = 0
        total_principal = 0
        total_payment = 0
        max_year = 9999  # Maximum allowed year in datetime

        # Append initial loan balance as the opening balance for the first entry
        opening_balance = loan_balance

        while loan_balance > 0:
            current_month = (start_month + month - 1) % 12 + 1
            current_year = start_year + (start_month + month - 1) // 12
            
            # Stop calculation if year exceeds maximum allowed year
            if current_year > max_year:
                break

            interest = loan_balance * monthly_interest_rate
            principal = min(monthly_payment - interest, loan_balance)
            loan_balance -= principal
            loan_balance = max(loan_balance, 0)

            # Update cumulative interest
            cumulative_interest += interest
            cumulative_principal += principal
            total_interest += interest
            total_principal += principal
            total_payment += monthly_payment
            
            # Add entry to schedule
            entry = (
                f"{current_month}/{current_year}",
                f"{opening_balance:,.2f}",  # Use original loan amount for opening balance
                f"{monthly_payment:,.2f}",
                f"{interest:,.2f}",
                f"{principal:,.2f}",
                f"{loan_balance:,.2f}",
                f"{cumulative_interest:,.2f}",
                f"{cumulative_principal:,.2f}"
            )
            schedule_list.append(entry)

            # Update opening balance for the next iteration
            opening_balance = loan_balance if loan_balance > 0 else 0  # Update only if loan balance is still positive
            month += 1

        # Append totals row
        total_row = (
            "Total",
            "$0.00", # Opening balance should go to zero at the end of the loan term
            f"{month * monthly_payment:,.2f}",
            f"{total_interest:,.2f}",
            f"{total_principal:,.2f}",
            "0.00",  # Closing balance should be 0
            f"{cumulative_interest:,.2f}",
            f"{cumulative_principal:,.2f}"
        )
        schedule_list.append(total_row)

        return schedule_list, month

    def get_total_interest(self):
        total_interest = 0
        for row in self.result_tree.get_children():
            values = self.result_tree.item(row, 'values')
            if values[0] == "Total":
                total_interest = float(values[3].replace(',', '').replace('$', ''))
        return total_interest

    def clear_inputs(self):
        for entry in self.entries:
            entry.delete(0, tk.END)
        self.result_tree.delete(*self.result_tree.get_children())
        self.time_label.config(text="")
        self.show_total_interest.config(text="")

    def export_schedule_to_csv(self):
        schedule = []
        for row in self.result_tree.get_children():
            schedule.append(self.result_tree.item(row, 'values'))

        with open("loan_payoff_schedule.csv", mode='w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(["Date", "Opening Balance", "Payment", "Interest", "Principal", "Closing Balance", "Cumulative Interest"])
            writer.writerows(schedule)
        messagebox.showinfo("Export Successful", "Loan payoff schedule exported successfully.")

    def extract_graph_data(self):
        data = []
        cumulative_principal = 0  # Initialize cumulative principal
        for row in self.result_tree.get_children():
            values = self.result_tree.item(row, 'values')
            if values[0] != "Total":  # Skip totals row
                date = values[0]
                cumulative_interest = float(values[6].replace(',', '').replace('$', ''))  # Cumulative interest
                closing_balance = float(values[5].replace(',', '').replace('$', ''))  # Closing balance
                
                # Calculate principal from payment and interest
                principal = float(values[4].replace(',', '').replace('$', ''))
                cumulative_principal += principal  # Update cumulative principal
                
                data.append((date, closing_balance, cumulative_interest, cumulative_principal))
        return data

    def show_payoff_graph(self):
        data = self.extract_graph_data()
        if data:
            dates, closing_balances, cumulative_interest, cumulative_principal = zip(*data)
            plt.figure(figsize=(12, 6))
            plt.plot(dates, cumulative_interest, marker='o', label='Cumulative Interest', color='blue')
            plt.plot(dates, cumulative_principal, marker='x', label='Cumulative Principal', color='orange')
            plt.plot(dates, closing_balances, marker='s', label='Closing Balance', color='green')
            plt.title('Loan Payoff Summary Over Time')
            plt.xlabel('Date')
            plt.ylabel('Amount ($)')
            plt.xticks(rotation=85)
            plt.grid()
            plt.legend()
            plt.tight_layout()
            plt.show()


if __name__ == "__main__":
    root = tk.Tk()
    loan_calculator = LoanCalculator(root)
    root.mainloop()
