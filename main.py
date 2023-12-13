import tkinter as tk
from tkinter import messagebox
from docxtpl import DocxTemplate
from datetime import datetime
import html
import csv


def restart_gui():
    root.destroy()
    start_gui()


def generate_forms():
    company_name = "H&N COMPANY"
    customers = []

    num_customers = int(entry_customers.get())

    for i in range(num_customers):
        ref_number = entry_refs[i].get()
        passenger_name = entry_names[i].get()
        ticket_number = entry_tickets[i].get()
        departure_date = entry_departures[i].get()
        cancelled_date = datetime.today().strftime("%d %b %Y")

        customers.append({
            'ref_number': ref_number,
            'passenger_name': passenger_name,
            'ticket_number': ticket_number,
            'departure_date': departure_date,
            'cancelled_date': cancelled_date
        })

    doc = DocxTemplate(
        "https://docs.google.com/document/d/1dOwjp0Is0Qh8C4inNaP0D3ThlGTDP0JI/edit")

    for customer in customers:
        today_date = datetime.today().strftime("%d %b %Y")

        context = {
            'company_name': html.escape(company_name),
            **customer,
            'today_date': today_date
        }

        doc.render(context)
        doc.save(f"refund_request_form_{customer['passenger_name']}.docx")

    save_to_csv = messagebox.askquestion(
        "Save to CSV", "Do you want to save the data to a CSV file?")

    if save_to_csv == 'yes':
        with open('data.csv', 'a', newline='') as csvfile:
            fieldnames = ['Reference', 'Passenger Name',
                          'Ticket Number', 'Departure Date', 'Cancelled Date']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

            csvfile.seek(0, 2)
            if csvfile.tell() == 0:
                writer.writeheader()

            for customer in customers:
                writer.writerow({
                    'Reference': customer['ref_number'],
                    'Passenger Name': customer['passenger_name'],
                    'Ticket Number': customer['ticket_number'],
                    'Departure Date': customer['departure_date'],
                    'Cancelled Date': customer['cancelled_date']
                })


def add_fields():
    num = int(entry_customers.get())

    # Destroy the current root window
    root.destroy()

    # Start a new GUI
    start_gui(num)


def start_gui(num_customers=0):
    global root, entry_customers, entry_refs, entry_names, entry_tickets, entry_departures

    root = tk.Tk()
    root.title("Refund Request Form Generator")

    tk.Label(root, text="Number of Customers:").grid(
        row=0, column=0, padx=10, pady=10)
    entry_customers = tk.Entry(root)
    entry_customers.grid(row=0, column=1, padx=10, pady=10)
    entry_customers.insert(0, num_customers)

    btn_add_fields = tk.Button(root, text="Submit", command=add_fields)
    btn_add_fields.grid(row=0, column=2, padx=10, pady=10)

    btn_generate = tk.Button(
        root, text="Generate Forms", command=generate_forms)
    btn_generate.grid(row=2, column=0, columnspan=8, padx=10, pady=10)

    entry_refs = []
    entry_names = []
    entry_tickets = []
    entry_departures = []

    for i in range(num_customers):
        tk.Label(root, text=f"Customer {i+1} Reference Number:").grid(
            row=i+3, column=0, padx=5, pady=5, sticky='w')
        entry_refs.append(tk.Entry(root))
        entry_refs[i].grid(row=i+3, column=1, padx=5, pady=5)

        tk.Label(root, text=f"Customer {i+1} Name:").grid(
            row=i+3, column=2, padx=5, pady=5, sticky='w')
        entry_names.append(tk.Entry(root))
        entry_names[i].grid(row=i+3, column=3, padx=5, pady=5)

        tk.Label(root, text=f"Customer {i+1} Ticket Number:").grid(
            row=i+3, column=4, padx=5, pady=5, sticky='w')
        entry_tickets.append(tk.Entry(root))
        entry_tickets[i].grid(row=i+3, column=5, padx=5, pady=5)

        tk.Label(root, text=f"Customer {i+1} Departure Date (DD/MM/YYYY):").grid(
            row=i+3, column=6, padx=5, pady=5, sticky='w')
        entry_departures.append(tk.Entry(root))
        entry_departures[i].grid(row=i+3, column=7, padx=5, pady=5)

    root.mainloop()


# Start the initial GUI
start_gui()
