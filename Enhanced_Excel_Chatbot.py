# Enhanced_Excel_Chatbot.py

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from PIL import Image, ImageTk
import io
import re

class ExcelChatbot:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Chatbot with Enhanced Features")
        self.root.geometry("1000x700")
        
        self.df = None
        self.current_visualization = None

        self.style = ttk.Style()
        self.style.configure('TFrame', background='#f0f0f0')
        self.style.configure('TButton', font=('Arial', 10), padding=5)
        self.style.configure('TLabel', background='#f0f0f0', font=('Arial', 10))

        self.create_widgets()

    def create_widgets(self):
        top_frame = ttk.Frame(self.root, padding="10")
        top_frame.pack(fill=tk.X)
        
        ttk.Label(top_frame, text="Excel File:").pack(side=tk.LEFT)
        self.file_path = tk.StringVar()
        ttk.Entry(top_frame, textvariable=self.file_path, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="Browse", command=self.load_excel).pack(side=tk.LEFT)
        
        chat_frame = ttk.Frame(self.root, padding="10")
        chat_frame.pack(fill=tk.BOTH, expand=True)
        
        self.vis_frame = ttk.Frame(chat_frame)
        self.vis_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.chat_history = ScrolledText(chat_frame, wrap=tk.WORD, height=15, state='disabled')
        self.chat_history.pack(fill=tk.BOTH, expand=True)
        
        input_frame = ttk.Frame(self.root, padding="10")
        input_frame.pack(fill=tk.X)
        
        self.user_input = ttk.Entry(input_frame)
        self.user_input.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.user_input.bind("<Return>", self.process_query)
        ttk.Button(input_frame, text="Send", command=self.process_query).pack(side=tk.LEFT)
        
        self.status = ttk.Label(self.root, text="Please load an Excel file to begin", relief=tk.SUNKEN)
        self.status.pack(fill=tk.X)

    def load_excel(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filepath:
            try:
                self.df = pd.read_excel(filepath)
                self.file_path.set(filepath)
                self.update_chat("System", f"Loaded Excel file with {len(self.df)} rows and columns: {', '.join(self.df.columns)}")
                self.status.config(text=f"Loaded: {filepath}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file: {str(e)}")
                self.df = None

    def update_chat(self, sender, message):
        self.chat_history.config(state='normal')
        self.chat_history.insert(tk.END, f"{sender}: {message}\n\n")
        self.chat_history.config(state='disabled')
        self.chat_history.see(tk.END)

    def clear_visualization(self):
        for widget in self.vis_frame.winfo_children():
            widget.destroy()
        self.current_visualization = None

    def display_visualization(self, fig):
        self.clear_visualization()
        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=100)
        buf.seek(0)
        img = Image.open(buf)
        img = ImageTk.PhotoImage(img)
        label = ttk.Label(self.vis_frame, image=img)
        label.image = img
        label.pack(fill=tk.BOTH, expand=True)
        self.current_visualization = img
        plt.close(fig)

    def process_query(self, event=None):
        query = self.user_input.get()
        if not query:
            return

        if self.df is None or self.df.empty:
            self.update_chat("Bot", "Please load an Excel file first")
            return

        self.update_chat("You", query)
        self.user_input.delete(0, tk.END)

        try:
            chart_type = None
            condition_query = None

            if "pie" in query.lower():
                chart_type = "pie"
            elif "bar" in query.lower():
                chart_type = "bar"
            elif "line" in query.lower():
                chart_type = "line"
            elif "scatter" in query.lower():
                chart_type = "scatter"
            elif "hist" in query.lower():
                chart_type = "hist"

            col = self.find_column_in_query(query)
            if chart_type and col:
                response = self.generate_visualization(chart_type, [col], condition_query=query)
                self.update_chat("Bot", response)
            else:
                response = self.simple_query_processing(query)
                if not response:
                    response = self.simulate_deepseek_call(query)
                self.update_chat("Bot", response)

        except Exception as e:
            self.update_chat("Bot", f"Error processing your query: {str(e)}")

    def simple_query_processing(self, query):
        query = query.lower()
        if "columns" in query or "headers" in query:
            return f"The columns in the data are: {', '.join(self.df.columns)}"
        if "show first" in query:
            return f"First 5 rows:\n{self.df.head().to_string()}"
        if "show last" in query:
            return f"Last 5 rows:\n{self.df.tail().to_string()}"
        if "describe" in query or "statistics" in query:
            return f"Statistics:\n{self.df.describe().to_string()}"
        return None

    def simulate_deepseek_call(self, query):
        if "average" in query or "mean" in query:
            col = self.find_column_in_query(query)
            if col:
                avg = self.df[col].mean()
                return f"The average {col} is {avg:.2f}"
        if "sum" in query:
            col = self.find_column_in_query(query)
            if col:
                total = self.df[col].sum()
                return f"The total {col} is {total:.2f}"
        if "below" in query:
            match = re.search(r'(\w+)\D*(\d+)', query)
            if match:
                col, val = match.groups()
            for c in self.df.columns:
                if c.lower() == col.lower():
                    col = c
                    break

            if col in self.df.columns:
                subset = self.df[self.df[col] < int(val)]
                return f"Entries where {col} < {val}:\n{subset.to_string(index=False)}"
        return "I've processed your request but didn't find specific information. Try asking about specific columns or values."

    def find_column_in_query(self, query):
        query_lower = query.lower()
        for col in self.df.columns:
            if col.lower() in query_lower:
                return col
        return None


    def generate_visualization(self, chart_type, columns, condition_query=None):
        df_filtered = self.df.copy()
        if condition_query and "below" in condition_query:
            match = re.search(r'(\w+)\D*(\d+)', condition_query)
            if match:
                col, val = match.groups()
                col = col.capitalize()
                if col in df_filtered.columns:
                    df_filtered = df_filtered[df_filtered[col] < int(val)]

        try:
            fig, ax = plt.subplots(figsize=(8, 4))
            if chart_type == "bar":
                df_filtered[columns[0]].value_counts().plot(kind='bar', ax=ax)
                ax.set_title(f"Bar Chart of {columns[0]}")
            elif chart_type == "pie":
                df_filtered[columns[0]].value_counts().plot(kind='pie', ax=ax, autopct='%1.1f%%')
                ax.set_title(f"Pie Chart of {columns[0]}")
            elif chart_type == "line":
                df_filtered[columns].plot(kind='line', ax=ax)
                ax.set_title(f"Line Chart of {columns[0]}")
            elif chart_type == "scatter" and len(columns) >= 2:
                df_filtered.plot(kind='scatter', x=columns[0], y=columns[1], ax=ax)
                ax.set_title(f"Scatter Plot {columns[1]} vs {columns[0]}")
            elif chart_type == "hist":
                df_filtered[columns[0]].plot(kind='hist', ax=ax)
                ax.set_title(f"Histogram of {columns[0]}")
            else:
                return f"Unsupported chart type or column configuration."

            plt.tight_layout()
            self.display_visualization(fig)
            return f"Displaying {chart_type} chart for {', '.join(columns)}"

        except Exception as e:
            return f"Failed to generate visualization: {str(e)}"

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelChatbot(root)
    root.mainloop()
