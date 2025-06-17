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
        self.root.title("Excel Data Chatbot with DeepSeek")
        self.root.geometry("1000x700")
        
        # Initialize variables
        self.df = None
        self.current_visualization = None
        
        # Configure styles
        self.style = ttk.Style()
        self.style.configure('TFrame', background='#f0f0f0')
        self.style.configure('TButton', font=('Arial', 10), padding=5)
        self.style.configure('TLabel', background='#f0f0f0', font=('Arial', 10))
        
        # Create main frames
        self.create_widgets()
        
    def create_widgets(self):
        # Top frame for file selection
        top_frame = ttk.Frame(self.root, padding="10")
        top_frame.pack(fill=tk.X)
        
        ttk.Label(top_frame, text="Excel File:").pack(side=tk.LEFT)
        self.file_path = tk.StringVar()
        ttk.Entry(top_frame, textvariable=self.file_path, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="Browse", command=self.load_excel).pack(side=tk.LEFT)
        
        # Chat area
        chat_frame = ttk.Frame(self.root, padding="10")
        chat_frame.pack(fill=tk.BOTH, expand=True)
        
        # Visualization frame
        self.vis_frame = ttk.Frame(chat_frame)
        self.vis_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Chat history
        self.chat_history = ScrolledText(chat_frame, wrap=tk.WORD, height=15, state='disabled')
        self.chat_history.pack(fill=tk.BOTH, expand=True)
        
        # Input area
        input_frame = ttk.Frame(self.root, padding="10")
        input_frame.pack(fill=tk.X)
        
        self.user_input = ttk.Entry(input_frame)
        self.user_input.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.user_input.bind("<Return>", self.process_query)
        ttk.Button(input_frame, text="Send", command=self.process_query).pack(side=tk.LEFT)
        
        # Status bar
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
        
        # Save the figure to a temporary buffer
        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=100)
        buf.seek(0)
        
        # Open the image and display in Tkinter
        img = Image.open(buf)
        img = ImageTk.PhotoImage(img)
        
        label = ttk.Label(self.vis_frame, image=img)
        label.image = img  # Keep a reference
        label.pack(fill=tk.BOTH, expand=True)
        
        self.current_visualization = img
        plt.close(fig)  # Close the figure to free memory
    
    def process_query(self, event=None):
        query = self.user_input.get()
        if not query:
            return
        
        # Fix: Proper DataFrame existence check
        if self.df is None or self.df.empty:
            self.update_chat("Bot", "Please load an Excel file first")
            return
        
        self.update_chat("You", query)
        self.user_input.delete(0, tk.END)
        
        try:
            # First try to process with simple commands
            response = self.simple_query_processing(query)
            
            if not response:
                # If simple processing fails, use simulated DeepSeek processing
                response = self.simulate_deepseek_call(query)
            
            self.update_chat("Bot", response)
            
        except Exception as e:
            self.update_chat("Bot", f"Error processing your query: {str(e)}")
    
    def simple_query_processing(self, query):
        query = query.lower()
        
        # Show column names
        if "columns" in query or "headers" in query:
            return f"The columns in the data are: {', '.join(self.df.columns)}"
        
        # Show first/last rows
        if "show first" in query or "display first" in query:
            n = 5
            if "rows" in query:
                match = re.search(r'(\d+)\s*rows', query)
                if match:
                    n = int(match.group(1))
            return f"First {n} rows:\n{self.df.head(n).to_string()}"
        
        if "show last" in query or "display last" in query:
            n = 5
            if "rows" in query:
                match = re.search(r'(\d+)\s*rows', query)
                if match:
                    n = int(match.group(1))
            return f"Last {n} rows:\n{self.df.tail(n).to_string()}"
        
        # Basic statistics
        if "describe" in query or "statistics" in query:
            return f"Statistics:\n{self.df.describe().to_string()}"
        
        # Count values
        if "count" in query and "values" in query:
            col = self.find_column_in_query(query)
            if col:
                return f"Value counts for {col}:\n{self.df[col].value_counts().to_string()}"
        
        # Basic visualization requests
        if "plot" in query or "chart" in query or "graph" in query:
            return self.handle_visualization_request(query)
        
        return None
    
    def simulate_deepseek_call(self, query):
        # In a real implementation, you would call the DeepSeek API here
        # This is a simulation that handles some basic cases
        
        if "average" in query.lower() or "mean" in query.lower():
            col = self.find_column_in_query(query)
            if col:
                avg = self.df[col].mean()
                return f"The average {col} is {avg:.2f}"
        
        if "sum" in query.lower():
            col = self.find_column_in_query(query)
            if col:
                total = self.df[col].sum()
                return f"The total {col} is {total:.2f}"
        
        if "correlation" in query.lower() or "relationship" in query.lower():
            cols = [c for c in self.df.columns if c.lower() in query.lower()]
            if len(cols) >= 2:
                corr = self.df[cols[0]].corr(self.df[cols[1]])
                return f"The correlation between {cols[0]} and {cols[1]} is {corr:.2f}. [VISUALIZATION: scatter({cols[0]}, {cols[1]})]"
        
        return "I've processed your request but didn't find specific information. Try asking about specific columns or values."
    
    def find_column_in_query(self, query):
        query_lower = query.lower()
        for col in self.df.columns:
            if col.lower() in query_lower:
                return col
        return None
    
    def handle_visualization_request(self, query):
        # Try to extract columns from query
        cols = [c for c in self.df.columns if c.lower() in query.lower()]
        
        if not cols:
            return "I couldn't determine which columns to visualize. Please specify column names."
        
        chart_type = "bar"  # default
        if "line" in query.lower():
            chart_type = "line"
        elif "pie" in query.lower():
            chart_type = "pie"
        elif "scatter" in query.lower():
            chart_type = "scatter"
        elif "histogram" in query.lower():
            chart_type = "hist"
        
        return self.generate_visualization(chart_type, cols)
    
    def generate_visualization(self, chart_type, columns):
        if not columns:
            return "No columns specified for visualization."
        
        columns = [c for c in columns if c in self.df.columns]
        if not columns:
            return "Specified columns not found in data."
        
        try:
            fig, ax = plt.subplots(figsize=(8, 4))
            
            if chart_type == "bar":
                if len(columns) == 1:
                    self.df[columns[0]].value_counts().plot(kind='bar', ax=ax)
                    ax.set_title(f"Count of {columns[0]}")
                else:
                    self.df[columns].plot(kind='bar', x=columns[0], ax=ax)
                    ax.set_title(f"{columns[1]} by {columns[0]}")
            
            elif chart_type == "line":
                self.df[columns].plot(kind='line', x=columns[0], ax=ax)
                ax.set_title(f"{columns[1]} over {columns[0]}")
            
            elif chart_type == "pie":
                if len(columns) == 1:
                    self.df[columns[0]].value_counts().plot(kind='pie', ax=ax, autopct='%1.1f%%')
                    ax.set_title(f"Distribution of {columns[0]}")
                else:
                    return "Pie charts typically use one column for values and one for labels."
            
            elif chart_type == "scatter":
                if len(columns) >= 2:
                    self.df.plot(kind='scatter', x=columns[0], y=columns[1], ax=ax)
                    ax.set_title(f"{columns[1]} vs {columns[0]}")
                else:
                    return "Scatter plots require two columns."
            
            elif chart_type == "hist":
                self.df[columns[0]].plot(kind='hist', ax=ax)
                ax.set_title(f"Distribution of {columns[0]}")
            
            else:
                return f"Unsupported chart type: {chart_type}"
            
            plt.tight_layout()
            self.display_visualization(fig)
            return f"Displaying {chart_type} chart for {', '.join(columns)}"
        
        except Exception as e:
            return f"Failed to generate visualization: {str(e)}"

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelChatbot(root)
    root.mainloop()