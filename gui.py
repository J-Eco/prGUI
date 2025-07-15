import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from style import *
import re


# Load Excel data
file_path = "Purchase Req Standard.xlsx"
items_df = pd.read_excel(file_path, sheet_name="Item List")

# Extract item names and descriptions
# items = items_df[['Item']].dropna().values.tolist()
# descriptions = items_df[['Description']].dropna().values.tolist()
item_desc_map = dict(zip(
    items_df['Item'].dropna(), 
    items_df['Description'].fillna('')
))
items = list(item_desc_map.keys())


class PurchaseReqApp:
    def __init__(self, root):
        self.finalCost = 0.0
        self.root = root
        root.configure(background=uibg, padx=30, pady=20)
        self.root.title("Purchase Requisition Form")
        root.tk.call('source', 'Forest-ttk-theme/forest-light.tcl')
        widgeStyle()

        root.rowconfigure(1, weight=1)  
        root.rowconfigure(2, weight=1)
        root.columnconfigure(0, weight=1)
        root.columnconfigure(1, weight=1)
        

        # Store selected items
        self.selected_items = []

        # Frame for form
        form_frame = ttk.LabelFrame(root, text="Requestor Information", style="Item.TLabelframe")
        form_frame.grid(row=0, column=1, padx=frameX, pady=frameY, sticky="nwe")


        vendor_frame = ttk.LabelFrame(root, text="Vendor Information", style="Item.TLabelframe")
        vendor_frame.grid(row=0, column=0, padx=frameX, pady=frameY, sticky="nwe")

        # Form fields
            # Vendor frames
        self.requestor_name = self._create_entry(vendor_frame, "Name:", 0)
        self.requestor_name = self._create_entry(vendor_frame, "Address:", 1)
        self.requestor_name = self._create_entry(vendor_frame, "City:", 2)
        self.requestor_name = self._create_entry(vendor_frame, "Phone:", 3)
        self.requestor_name = self._create_entry(vendor_frame, "SQ-", 4)


            # Request frames
        self.requestor_name = self._create_entry(form_frame, "Requestor Name:", 0)
        self.tech_name = self._create_entry(form_frame, "Technician:", 1)
        self.department = self._create_combobox(form_frame,  "Department:", 2, [
            "FA - FireA Alarm", 
            "FS - Fire Sprinkler", 
            "FI - Installation",
            "FES - Fire Extin/Supp", "FM - Fleet Maintenance"
        ])
        self.date = self._create_entry(form_frame, "Date:", 3)

        # Frame for item selection
        item_frame = ttk.LabelFrame(root, text="Item Selection", style="Item.TLabelframe",)
        item_frame.grid(row=1, column=0, columnspan=2, padx=frameX, pady=0, sticky="nsew")
        root.rowconfigure(1, weight=1)


        # Item / Part
        self.item_var = tk.StringVar()
        self.item_menu = ttk.Combobox(item_frame, textvariable=self.item_var, width=60)
        self.item_menu['values'] = items
        self.item_menu.grid(row=0, column=0, padx=5, pady=5)


        # Quantity
        ttk.Label(item_frame, text="Quantity:").grid(row=0, column=2, padx=5, pady=5, sticky="nsew")
        self.qty_var = tk.StringVar(value="1")
        self.qty_entry = ttk.Entry(item_frame, textvariable=self.qty_var, width=5)
        self.qty_entry.grid(row=0, column=3, padx=5, pady=5)

        # Cost
        ttk.Label(item_frame, text="Cost: $").grid(row=0, column=4, padx=5, pady=5, sticky="nsew")
        self.costVar= tk.StringVar(value="0")
        self.costEntry = ttk.Entry(item_frame, textvariable=self.costVar, width=25)
        self.total = self.costEntry
        self.costEntry.grid(row=0, column=5, padx=5, pady=5)
        self.lastCostVal = ""
        self.costVar.trace_add('write', self.check)
        

        add_btn = ttk.Button(item_frame, text="Add Item", style='Accent.TButton', command=self.add_item)
        add_btn.grid(row=0, column=6, padx=5, pady=5)

        remove_btn = ttk.Button(item_frame, text="Remove Item", command=self.removeItem)
        remove_btn.grid(row=0, column=7, padx=5, pady=5)



        # Frame for Requistions
        preview_frame = ttk.LabelFrame(root, text="Requisition Preview", style="Item.TLabelframe")
        preview_frame.grid(row=2, columnspan=2, padx=10, pady=10, sticky="nsew")

        self.tree = ttk.Treeview(preview_frame, columns=("Item", "Desc.", "Quant.", "Cost", "Total"), show="headings")
        self.tree.heading("Item", text="Item/Part Number")
        self.tree.heading("Desc.", text="Description")
        self.tree.heading("Quant.", text="Quantity")
        self.tree.heading("Cost", text="Cost Price")
        self.tree.heading("Total", text="Total")

        self.tree.column("Item", anchor="center", width=150)
        self.tree.column("Desc.", anchor="e", width=300) 
        self.tree.column("Quant.", anchor="center", width=50)
        self.tree.column("Cost", anchor="e", width=50)
        self.tree.column("Total", anchor="e", width=80)

        self.tree.tag_configure('evenrow', background="#ffffff")
        self.tree.tag_configure('oddrow', background="#dee2e2")


        self.tree.grid(row=0, column=0, sticky="nsew")

        # Resize treeview
        preview_frame.rowconfigure(0, weight=1)
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")




        # Final total section
        finalTotalFrame = ttk.Label(root, text="Final Total:", font=('Arial', 12, 'bold'))
        finalTotalFrame.grid(row=2, columnspan=2, padx=100, pady=20, sticky="se")
        self.summedTotal = tk.StringVar(value="0.00")
        totalValueLabel = ttk.Label(root, textvariable=self.summedTotal, font=('Arial', 12))
        totalValueLabel.grid(row=2, column=1, padx=50, pady=20, sticky="se")
        finalTotalFrame.grid(row=2, columnspan=2, padx=100, pady=20, sticky="se")




    # configures

        # ITem frames
        item_frame.columnconfigure(0, weight=3)  # Item Combobox
        item_frame.columnconfigure(1, weight=0)  # Label shares space with column 0
        item_frame.columnconfigure(2, weight=1)  # Quantity label
        item_frame.columnconfigure(3, weight=1)  # Quantity entry
        item_frame.columnconfigure(4, weight=1)  # Cost label
        item_frame.columnconfigure(5, weight=2)  # Cost entry
        item_frame.columnconfigure(6, weight=1)  # Add button
        item_frame.columnconfigure(7, weight=1)
        item_frame.grid(row=1, column=0, columnspan=2, padx=frameX, pady=0, sticky="ew")

        # self.item_menu.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        # self.qty_entry.grid(row=0, column=3, padx=5, pady=5, sticky="nsew")
        # self.costEntry.grid(row=0, column=5, padx=5, pady=5, sticky="nsew")
        # add_btn.grid(row=0, column=6, padx=5, pady=5, sticky="nsew")
        # remove_btn.grid(row=0, column=7, padx=5, pady=5, sticky="nsew")
        item_frame.rowconfigure(0, weight=1)




    # functions

    def _create_entry(self, parent, label, row):
        ttk.Label(parent, text=label).grid(row=row, column=0, padx=5, pady=5, sticky="e")
        var = tk.StringVar()
        entry = ttk.Entry(parent, textvariable=var, width=60)
        entry.grid(row=row, column=1, padx=10, pady=8, sticky="w")
        return var

    def _create_combobox(self, parent, label, row, options):
        ttk.Label(parent, text=label).grid(row=row, column=0, padx=5, pady=5, sticky="e")
        var = tk.StringVar()
        combo = ttk.Combobox(parent, textvariable=var, values=options, width=38)
        combo.grid(row=row, column=1, padx=5, pady=5, sticky="w")
        return var

    
    def add_item(self):
        item = self.item_var.get()
        qty_str = self.qty_var.get()
        cost_str = self.costVar.get()

        if not item:
            messagebox.showerror("Invalid Entry", "Please select a valid item, quantity, and cost.")
            return

        qty = int(qty_str)
        cost = float(cost_str)
        rowTotal = round(qty * cost, 2)

        desc = item_desc_map.get(item, "N/A")
        self.selected_items.append((item, desc, qty, (cost * qty), rowTotal))  
        total = sum(row[4] for row in self.selected_items)
        self.summedTotal.set(f"{total:.2f}")
        self.refreshTable()



        
    def removeItem(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showinfo("No selection", "Please select a row to remove.")
            return

        for item_id in selected_items:
            values = self.tree.item(item_id, 'values')
            item_name = values[0]
            qty = float(values[2])
            cost = float(values[3])

            for entry in self.selected_items:
                if entry[0] == item_name and entry[2] == qty and entry[3] == cost:
                    self.selected_items.remove(entry)
                    break

            self.tree.delete(item_id)

        self.refreshTable()


    
    def check(self, *args):
        val = self.costVar.get()
        if re.fullmatch(r'^\d+(\.\d{0,2})?$', val):
            self.lastCostVal = val
        else:
             self.costVar.set(self.lastCostVal)

    def refreshTable(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        self.finalCost = 0
        for index, (item, desc, qty, cost, lineTotal) in enumerate(self.selected_items):
            self.finalCost += lineTotal
            tag = 'evenrow' if index % 2 == 0 else 'oddrow'
            self.tree.insert(
                "", "end",
                values=(item, desc, qty, cost, self.finalCost),
                tags=(tag,)
            )



if __name__ == "__main__":
    root = tk.Tk()
    app = PurchaseReqApp(root)
    root.mainloop()
