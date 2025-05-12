# --- IMPORTS ---
import tkinter as tk
from tkinter import messagebox, ttk, simpledialog
import oracledb
import smtplib
from email.message import EmailMessage
from PIL import ImageTk, Image
from datetime import datetime

# --- GLOBALS ---
conn = None
cursor = None
selected_train_id = None

# --- DB INIT ---
def initialize_db():
    global conn, cursor
    conn = oracledb.connect(
        user="SYS",
        password="Pspk@123",
        dsn="localhost/XEPDB1",
        mode=oracledb.AUTH_MODE_SYSDBA
    )
    cursor = conn.cursor()

# --- SEARCH TRAINS ---
def search_trains():
    for row in train_tree.get_children():
        train_tree.delete(row)

    src = entry_source.get().strip().title()
    dst = entry_destination.get().strip().title()

    if not src or not dst:
        messagebox.showwarning("Missing Fields", "Enter both Source and Destination.")
        return

    query = """
        SELECT TRAIN_ID, TRAIN_NAME, SOURCE_STATION, DESTINATION_STATION, 
               DEPARTURE_TIME, ARRIVAL_TIME 
        FROM TRAINS 
        WHERE SOURCE_STATION = :1 AND DESTINATION_STATION = :2
    """
    cursor.execute(query, (src, dst))
    results = cursor.fetchall()

    travel_date = entry_date.get().strip()
    if not travel_date:
        messagebox.showwarning("Date Missing", "Enter Travel Date.")
        return

    if not results:
        messagebox.showinfo("No Trains", "No trains found for this route.")
    else:
        for row in results:
            seats = get_available_seats(row[0], travel_date)
            train_tree.insert("", tk.END, values=row + (seats,))

# --- GET SEAT COUNT ---
def get_available_seats(train_id, travel_date):
    cursor.execute("""
        SELECT COUNT(*) FROM TRAIN_TICKETS 
        WHERE TRAIN_ID = :1 AND TRAVEL_DATE = TO_DATE(:2, 'YYYY-MM-DD') AND STATUS = 'Confirmed'
    """, (train_id, travel_date))
    booked = cursor.fetchone()[0]
    cursor.execute("SELECT TOTAL_SEATS FROM TRAINS WHERE TRAIN_ID = :1", (train_id,))
    total = cursor.fetchone()[0]
    return total - booked

# --- TRAIN SELECT ---
def on_train_select(event):
    global selected_train_id

    item = train_tree.selection()
    if not item:
        return
    values = train_tree.item(item, "values")
    selected_train_id = values[0]

    selected_class = class_combobox.get()
    if not selected_class or selected_class == "Select Class":
        messagebox.showwarning("Class Required", "Please select a travel class")
        return

    amount = simpledialog.askinteger("Ticket Amount", "Enter ticket amount in ₹:")
    if not amount:
        return

    proceed_booking(values, selected_class, amount)

# --- BOOKING FUNCTION ---
def proceed_booking(train_data, travel_class, amount):
    name = entry_name.get().strip()
    age = entry_age.get().strip()
    travel_date = entry_date.get().strip()
    email = entry_email.get().strip()

    if not all([name, age, travel_date]):
        messagebox.showerror("Error", "All passenger fields are required.")
        return

    try:
        available_seats = get_available_seats(train_data[0], travel_date)
        status = "Confirmed" if available_seats > 0 else "Waiting"

        seat_number = get_next_seat_number(train_data[0], travel_date) if status == "Confirmed" else None

        cursor.execute("""
            INSERT INTO TRAIN_TICKETS 
              (PASSENGER_NAME, AGE, SOURCE_STATION, DESTINATION_STATION, TRAVEL_DATE, 
               TRAIN_ID, STATUS, CLASS_TYPE, AMOUNT, SEAT_NUMBER)
            VALUES 
              (:1, :2, :3, :4, TO_DATE(:5, 'YYYY-MM-DD'), :6, :7, :8, :9, :10)
        """, (name, int(age), train_data[2], train_data[3], travel_date, train_data[0],
              status, travel_class, amount, seat_number))
        conn.commit()

        cursor.execute("""
            SELECT TICKET_ID, BOOKING_DATE, BOOKING_TIMESTAMP 
            FROM TRAIN_TICKETS 
            WHERE TICKET_ID = (SELECT MAX(TICKET_ID) FROM TRAIN_TICKETS)
        """)
        ticket = cursor.fetchone()

        confirm_msg = (
            f"✅ Ticket Booked Successfully!\n\n"
            f"Ticket ID: {ticket[0]}\nName: {name}\nAge: {age}\n"
            f"Train: {train_data[1]} ({train_data[2]} → {train_data[3]})\n"
            f"Departure: {train_data[4]} | Arrival: {train_data[5]}\n"
            f"Class: {travel_class}\nFare: ₹{amount}\n"
            f"Seat Number: {seat_number if seat_number else 'N/A'}\n"
            f"Travel Date: {travel_date}\nStatus: {status}\n"
            f"Booking Date: {ticket[1]}\nTime: {ticket[2]}"
        )

        if email:
            send_confirmation_email(email, confirm_msg)
        else:
            messagebox.showinfo("Booking Confirmed", confirm_msg)

        clear_fields()

    except oracledb.DatabaseError as e:
        conn.rollback()
        messagebox.showerror("DB Error", str(e))

# --- GET NEXT SEAT NUMBER ---
def get_next_seat_number(train_id, travel_date):
    cursor.execute("""
        SELECT MAX(SEAT_NUMBER) FROM TRAIN_TICKETS
        WHERE TRAIN_ID = :1 AND TRAVEL_DATE = TO_DATE(:2, 'YYYY-MM-DD') AND STATUS = 'Confirmed'
    """, (train_id, travel_date))
    result = cursor.fetchone()[0]
    return 1 if result is None else result + 1

# --- TICKET STATUS CHECK ---
def check_ticket_status():
    ticket_id = simpledialog.askinteger("Check Ticket Status", "Enter your Ticket ID:")
    if not ticket_id:
        return
    try:
        cursor.execute("""
            SELECT TICKET_ID, PASSENGER_NAME, AGE, SOURCE_STATION, DESTINATION_STATION,
                   TO_CHAR(TRAVEL_DATE, 'YYYY-MM-DD'), CLASS_TYPE, AMOUNT, STATUS, SEAT_NUMBER
            FROM TRAIN_TICKETS WHERE TICKET_ID = :1
        """, (ticket_id,))
        result = cursor.fetchone()
        if result:
            info = (
                f"Ticket ID: {result[0]}\nName: {result[1]}\nAge: {result[2]}\n"
                f"From: {result[3]} To: {result[4]}\nTravel Date: {result[5]}\n"
                f"Class: {result[6]}\nFare: ₹{result[7]}\nStatus: {result[8]}\n"
                f"Seat Number: {result[9] if result[9] else 'N/A'}"
            )
            messagebox.showinfo("Ticket Status", info)
        else:
            messagebox.showwarning("Not Found", "No ticket found with that ID.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# --- EMAIL CONFIRMATION ---
def send_confirmation_email(to_email, body):
    msg = EmailMessage()
    msg['Subject'] = "Train Ticket Confirmation"
    msg['From'] = "irctc.railway.sbc@gmail.com"
    msg['To'] = to_email
    msg.set_content(body)

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login("irctc.railway.sbc@gmail.com", "rvpb pfmh nhlw icuk")
            smtp.send_message(msg)
        messagebox.showinfo("Email Sent", f"Confirmation sent to {to_email}")
    except Exception as e:
        messagebox.showerror("Email Error", f"Could not send email:\n{str(e)}")

# --- CLEAR INPUTS ---
def clear_fields():
    for e in [entry_name, entry_age, entry_source, entry_destination, entry_date, entry_email]:
        e.delete(0, tk.END)
    class_combobox.set("Select Class")
    for row in train_tree.get_children():
        train_tree.delete(row)

# --- EXIT ---
def on_closing():
    if cursor:
        cursor.close()
    if conn:
        conn.close()
    root.destroy()
# --- CANCEL TICKET ---
def cancel_ticket():
    ticket_id = simpledialog.askinteger("Cancel Ticket", "Enter Ticket ID to cancel:")
    if not ticket_id:
        return
    try:
        cursor.execute("""
            SELECT TRAIN_ID, TRAVEL_DATE, STATUS, SEAT_NUMBER FROM TRAIN_TICKETS
            WHERE TICKET_ID = :1
        """, (ticket_id,))
        ticket = cursor.fetchone()

        if not ticket:
            messagebox.showwarning("Invalid", "Ticket ID not found.")
            return

        train_id, travel_date, status, seat_no = ticket

        if status != 'Confirmed':
            messagebox.showinfo("Not Confirmed", "Only confirmed tickets can be cancelled.")
            return

        # Delete the confirmed ticket
        cursor.execute("DELETE FROM TRAIN_TICKETS WHERE TICKET_ID = :1", (ticket_id,))
        conn.commit()

        # Promote next waiting ticket (if any)
        cursor.execute("""
            SELECT TICKET_ID FROM TRAIN_TICKETS
            WHERE TRAIN_ID = :1 AND TRAVEL_DATE = :2 AND STATUS = 'Waiting'
            ORDER BY BOOKING_TIMESTAMP ASC FETCH FIRST 1 ROWS ONLY
        """, (train_id, travel_date))
        waiting = cursor.fetchone()

        if waiting:
            next_ticket_id = waiting[0]
            cursor.execute("""
                UPDATE TRAIN_TICKETS
                SET STATUS = 'Confirmed', SEAT_NUMBER = :1
                WHERE TICKET_ID = :2
            """, (seat_no, next_ticket_id))
            conn.commit()

        messagebox.showinfo("Cancelled", "Ticket cancelled successfully.")
    except Exception as e:
        conn.rollback()
        messagebox.showerror("Error", str(e))


# --- GUI SETUP ---
root = tk.Tk()
root.title("IRCTC Railway Booking")
root.geometry("850x780")

try:
    img = Image.open(r"C:\\Users\\DELL\\PycharmProjects\\sum.py\\IRCTC\\Train_pic.jpg")
    img = img.resize((400, 150), Image.LANCZOS)
    header_img = ImageTk.PhotoImage(img)
    img_label = tk.Label(root, image=header_img)
    img_label.image = header_img
    img_label.pack(pady=10)
except Exception as e:
    print(f"Error loading image: {str(e)}")

input_frame = tk.Frame(root)
input_frame.pack(pady=10, padx=20, fill=tk.X)

labels = ["Passenger Name:", "Age:", "Source Station:", "Destination Station:",
          "Travel Date (YYYY-MM-DD):", "Travel Class:", "Email (Optional):"]
tk.Label(input_frame, text=labels[0]).grid(row=0, column=0, sticky=tk.W)
tk.Label(input_frame, text=labels[1]).grid(row=0, column=2, sticky=tk.W)
tk.Label(input_frame, text=labels[2]).grid(row=1, column=0, sticky=tk.W)
tk.Label(input_frame, text=labels[3]).grid(row=1, column=2, sticky=tk.W)
tk.Label(input_frame, text=labels[4]).grid(row=2, column=0, sticky=tk.W)
tk.Label(input_frame, text=labels[5]).grid(row=2, column=2, sticky=tk.W)
tk.Label(input_frame, text=labels[6]).grid(row=3, column=0, sticky=tk.W)
#tk.Button(btn_frame, text="Cancel Ticket", command=cancel_ticket, bg="#A30000", fg="white", width=20).pack(side=tk.LEFT, padx=10)


entry_name = tk.Entry(input_frame, width=25); entry_name.grid(row=0, column=1)
entry_age = tk.Entry(input_frame, width=10); entry_age.grid(row=0, column=3)
entry_source = tk.Entry(input_frame, width=25); entry_source.grid(row=1, column=1)
entry_destination = tk.Entry(input_frame, width=25); entry_destination.grid(row=1, column=3)
entry_date = tk.Entry(input_frame, width=25); entry_date.grid(row=2, column=1)
class_combobox = ttk.Combobox(input_frame, values=["Sleeper", "AC", "General"], width=15, state="readonly")
class_combobox.set("Select Class")
class_combobox.grid(row=2, column=3)
entry_email = tk.Entry(input_frame, width=25); entry_email.grid(row=3, column=1)

btn_frame = tk.Frame(root)
btn_frame.pack(pady=10)
tk.Button(btn_frame, text="Search Trains", command=search_trains, bg="#0055A3", fg="white", width=20).pack(side=tk.LEFT, padx=10)
tk.Button(btn_frame, text="Check Ticket Status", command=check_ticket_status, bg="#007A3E", fg="white", width=20).pack(side=tk.LEFT, padx=10)
tk.Button(btn_frame, text="Cancel Ticket", command=cancel_ticket, bg="#A30000", fg="white", width=20).pack(side=tk.LEFT, padx=10)

cols = ("ID", "Train Name", "From", "To", "Departure", "Arrival", "Seats")
train_tree = ttk.Treeview(root, columns=cols, show='headings', height=8)
for col in cols:
    train_tree.heading(col, text=col)
    train_tree.column(col, anchor="center", width=100)
train_tree.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
train_tree.bind("<Double-1>", on_train_select)

try:
    initialize_db()
except Exception as e:
    messagebox.showerror("Connection Failed", str(e))
    exit()

root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()
