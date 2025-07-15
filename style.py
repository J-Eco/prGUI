from tkinter import ttk


def widgeStyle():
        style = ttk.Style()
        ttk.Style().theme_use('forest-light')

        style.configure("Item.TLabelframe",
        )
        style.configure("Item.TLabelframe.Label",
                font=('Arial', 11, 'bold'),
                width=20,
                anchor = 'center',
                pady= 16   
        )



uibg = "#ffffff"

frameX = 14
frameY = 22